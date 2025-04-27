/**
 * @fileoverview Handles core data operations like retrieving, saving,
 * updating, and propagating habit data between the main and history sheets.
 */

/** OnlyCurrentDoc */

const DataHandler = {
  // --- Helpers for Range Manipulation ---

  /**
   * Converts a Range object into an array of A1 notation cell strings.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object.
   * @returns {Array<string>} An array of A1 notation strings (e.g., ["A1", "B1", "A2", "B2"]).
   */
  getRangeCells: function (range) {
    if (!range) return [];
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    const rangeCells = [];

    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        rangeCells.push(this._cellToA1Notation(startRow + r, startCol + c));
      }
    }
    return rangeCells;
  },

  /**
   * Converts row and column numbers (1-based) to A1 notation.
   * @private
   * @param {number} row - The 1-based row number.
   * @param {number} col - The 1-based column number.
   * @returns {string} The A1 notation (e.g., "A1", "Z10", "AA1").
   */
  _cellToA1Notation: function (row, col) {
    let columnName = "";
    let current = col;
    while (current > 0) {
      const remainder = (current - 1) % 26;
      columnName = String.fromCharCode(65 + remainder) + columnName;
      current = Math.floor((current - 1) / 26);
    }
    return columnName + row;
  },

  /**
   * Retrieves an array of 1-based column indices within the given range.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object.
   * @returns {Array<number>} An array of column numbers.
   */
  getRangeColumns: function (range) {
    if (!range) return [];
    const startCol = range.getColumn();
    const numCols = range.getNumColumns();
    return Array.from({ length: numCols }, (_, i) => startCol + i);
  },

  // --- Main Sheet Data Access & Manipulation ---

  /**
   * Gets the list of row numbers corresponding to habits (identified by emojis).
   * Uses the stored emoji list from properties. Handles potential habit spread changes.
   * @returns {Array<number>} An array of 1-indexed row numbers. Returns empty array on error.
   */
  getRelevantRows: function () {
    // Check if user manually changed emojis in challenge mode
    const activitiesChanged = PropertyManager.getProperty(
      PropertyKeys.ACTIVITIES_COLUMN_UPDATED
    );
    const mode = PropertyManager.getProperty(PropertyKeys.MODE);
    let updateRan = false;

    if (
      activitiesChanged === BooleanTypes.TRUE &&
      mode === ModeTypes.CHALLENGE
    ) {
      LoggerManager.logDebug(
        "Activities column changed in challenge mode. Checking emoji spread..."
      );
      if (!HabitManager.checkEmojiSpread()) {
        // HabitManager now handles the check/revert
        // checkEmojiSpread already shows alert and reverts if needed.
        updateRan = true; // Flag that a check/revert might have happened
      }
      // Reset the flag regardless of whether changes were reverted
      PropertyManager.setProperty(
        PropertyKeys.ACTIVITIES_COLUMN_UPDATED,
        BooleanTypes.FALSE
      );
    }

    // Get the definitive list from properties
    let relevantRows = [];
    try {
      const storedEmojiSpread = JSON.parse(
        PropertyManager.getProperty(PropertyKeys.EMOJI_LIST) || "[]"
      );
      relevantRows = storedEmojiSpread.map((entry) => entry.row);
      LoggerManager.logDebug(
        `getRelevantRows: Rows from properties: ${relevantRows.join(", ")}`
      );

      // Validate the rows obtained from properties only if a check/revert happened
      if (updateRan) {
        if (!ValidationUtils._validateRowsArray(relevantRows, false)) {
          LoggerManager.handleError(
            "Emoji spread check/revert resulted in invalid relevant rows.",
            true
          );
          return []; // Return empty on validation failure after check
        }
      }
      // Basic validation even if no update ran
      if (!ValidationUtils._validateArray(relevantRows)) {
        LoggerManager.handleError(
          "Stored emoji list property is not a valid array.",
          false
        );
        return [];
      }
    } catch (e) {
      LoggerManager.handleError(
        `Failed to parse stored emoji list: ${e.message}`
      );
      return [];
    }

    return relevantRows;
  },

  /**
   * Gets the number of active habits.
   * @private
   * @returns {number} The count of relevant rows.
   */
  _getHabitCount: function () {
    return this.getRelevantRows().length;
  },

  /**
   * Creates a default list (array of arrays) for completion data (empty strings).
   * @returns {Array<Array<string>>} e.g., [[""], [""], [""]]
   */
  _getDefaultCompletionListForSheet: function () {
    const length = this._getHabitCount();
    return Array.from({ length }, () => [""]); // Structure matches setValues requirement
  },

  /**
   * Creates a default list (array of arrays) for buffer data (default buffer value).
   * @returns {Array<Array<number>>} e.g., [[1], [1], [1]]
   */
  _getDefaultBufferListForSheet: function () {
    const length = this._getHabitCount();
    return Array.from({ length }, () => [MainSheetConfig.defaultBuffer]); // Structure matches setValues
  },

  /**
   * Creates a 1D array of default completion values (empty strings).
   * Used for history storage.
   * @returns {Array<string>} e.g., ["", "", ""]
   */
  _getDefaultCompletionListForStorage: function () {
    const length = this._getHabitCount();
    return Array.from({ length }, () => "");
  },

  /**
   * Creates a 1D array of default buffer values.
   * Used for history storage and propagation logic.
   * @returns {Array<number>} e.g., [1, 1, 1]
   */
  getDefaultBufferList: function () {
    const length = this._getHabitCount();
    return Array.from({ length }, () => MainSheetConfig.defaultBuffer);
  },

  /**
   * Retrieves completion data only for the relevant habit rows.
   * Returns a 1D array. Returns empty array on error.
   * @returns {Array<*>} Completion values for habits (e.g., [true, false, true]).
   */
  getCompletionData: function () {
    const relevantRows = this.getRelevantRows();
    const completionRange = MainSheetConfig.getCompletionDataRange();
    if (!completionRange || relevantRows.length === 0) {
      LoggerManager.logDebug(
        "getCompletionData: No completion range or relevant rows found."
      );
      return [];
    }

    const allData = completionRange.getValues(); // Gets 2D array [[val1], [val2], ...]
    const firstDataSheetRow = MainSheetConfig.firstDataInputRow;

    // Filter data based on relevantRows indices relative to the range's start
    const filteredData = relevantRows
      .map((sheetRowIndex) => {
        const rangeRowIndex = sheetRowIndex - firstDataSheetRow; // 0-based index within the range
        if (rangeRowIndex >= 0 && rangeRowIndex < allData.length) {
          return allData[rangeRowIndex][0]; // Extract value from the inner array
        } else {
          LoggerManager.handleError(
            `Row index mismatch in getCompletionData. SheetRow: ${sheetRowIndex}, RangeIndex: ${rangeRowIndex}, DataLength: ${allData.length}`,
            false
          );
          return undefined; // Or some other indicator of error
        }
      })
      .filter((item) => item !== undefined); // Remove potential errors

    // Basic validation on the result
    if (!ValidationUtils._validateCompletionData(filteredData, false)) {
      LoggerManager.handleError(
        "Validation failed for retrieved completion data.",
        false
      );
      return []; // Return empty on validation failure
    }

    LoggerManager.logDebug(
      `getCompletionData filtered: ${JSON.stringify(filteredData)}`
    );
    return filteredData;
  },

  /**
   * Retrieves buffer data only for the relevant habit rows.
   * Returns a 1D array of numbers. Returns empty array on error.
   * @returns {Array<number>} Buffer values for habits (e.g., [1, 0, 2]).
   */
  getBufferData: function () {
    const relevantRows = this.getRelevantRows();
    const bufferRange = MainSheetConfig.getBufferDataRange();
    if (!bufferRange || relevantRows.length === 0) {
      LoggerManager.logDebug(
        "getBufferData: No buffer range or relevant rows found."
      );
      return [];
    }

    const allData = bufferRange.getValues(); // Gets 2D array [[val1], [val2], ...]
    const firstDataSheetRow = MainSheetConfig.firstDataInputRow;

    // Filter data
    const filteredData = relevantRows
      .map((sheetRowIndex) => {
        const rangeRowIndex = sheetRowIndex - firstDataSheetRow;
        if (rangeRowIndex >= 0 && rangeRowIndex < allData.length) {
          return allData[rangeRowIndex][0];
        } else {
          LoggerManager.handleError(
            `Row index mismatch in getBufferData. SheetRow: ${sheetRowIndex}, RangeIndex: ${rangeRowIndex}, DataLength: ${allData.length}`,
            false
          );
          return undefined;
        }
      })
      .filter((item) => item !== undefined);

    // Validate result contains only numbers
    if (!ValidationUtils._validateBufferData(filteredData, false)) {
      LoggerManager.handleError(
        "Validation failed for retrieved buffer data.",
        false
      );
      return []; // Return empty on validation failure
    }

    LoggerManager.logDebug(
      `getBufferData filtered: ${JSON.stringify(filteredData)}`
    );
    return filteredData;
  },

  /**
   * Retrieves the current streak value from the main sheet.
   * @returns {number} The current streak. Returns 0 on error.
   */
  getCurrentStreak: function () {
    const value = MainSheetConfig._getSheetValue(
      MainSheetConfig.currentStreakCell
    );
    return ValidationUtils._validateNonNegativeNumber(value) ? value : 0;
  },

  /**
   * Sets the current streak value on the main sheet.
   * @param {number} currentStreak - The value to set.
   */
  setCurrentStreak: function (currentStreak) {
    if (ValidationUtils._validateCurrentStreak(currentStreak)) {
      MainSheetConfig._setSheetValue(
        MainSheetConfig.currentStreakCell,
        currentStreak
      );
    }
  },

  /**
   * Retrieves the highest streak value from the main sheet.
   * @returns {number} The highest streak. Returns 0 on error.
   */
  getHighestStreak: function () {
    const value = MainSheetConfig._getSheetValue(
      MainSheetConfig.highestStreakCell
    );
    return ValidationUtils._validateNonNegativeNumber(value) ? value : 0;
  },

  /**
   * Sets the highest streak value on the main sheet.
   * @param {number} highestStreak - The value to set.
   */
  setHighestStreak: function (highestStreak) {
    if (ValidationUtils._validateHighestStreak(highestStreak)) {
      MainSheetConfig._setSheetValue(
        MainSheetConfig.highestStreakCell,
        highestStreak
      );
    }
  },

  /**
   * Retrieves the date value from the date cell as a Date object.
   * @returns {Date | null} The date object or null on error/invalid date.
   */
  getDate: function () {
    const rawValue = MainSheetConfig._getSheetValue(MainSheetConfig.dateCell);
    try {
      const date = DateManager.determineDate(rawValue);
      return date;
    } catch (e) {
      LoggerManager.handleError(
        `Invalid date value in cell ${MainSheetConfig.dateCell}: ${rawValue}. Error: ${e.message}`,
        false
      );
      return null;
    }
  },

  /**
   * Sets the date cell on the main sheet using a formatted date string.
   * @param {string} dateStr - The date string (YYYY-MM-DD) to set.
   */
  setDateStr: function (dateStr) {
    if (DateManager._validateDateStr(dateStr)) {
      // Use basic format validation
      MainSheetConfig._setSheetValue(MainSheetConfig.dateCell, dateStr);
    } else {
      LoggerManager.handleError(
        `Attempted to set invalid date string: ${dateStr}`,
        false
      );
    }
  },

  /**
   * Sets all core data (date, completion, buffer, streaks) on the main sheet.
   * Takes 1D arrays for completion/buffer and maps them to the correct rows.
   *
   * @param {string} dateStr - The date string (YYYY-MM-DD) to set.
   * @param {Array<*>} completionData - 1D array of completion values for relevant rows.
   * @param {Array<number>} bufferData - 1D array of buffer values for relevant rows.
   * @param {number} currentStreak - The current streak value.
   * @param {number} highestStreak - The highest streak value.
   */
  setAllMainSheetData: function (
    dateStr,
    completionData,
    bufferData,
    currentStreak,
    highestStreak
  ) {
    LoggerManager.logDebug(`setAllMainSheetData called for date: ${dateStr}`);
    // --- Input Validation ---
    if (!DateManager._validateDateStr(dateStr, false)) {
      // Allow setting future dates maybe? Validate format only.
      LoggerManager.handleError(
        `Invalid date string format for setAllMainSheetData: ${dateStr}`,
        true
      );
      return;
    }
    if (
      !ValidationUtils._validateCompletionAndBufferData(
        completionData,
        bufferData
      )
    ) {
      // Error handled in validation function
      return;
    }
    if (
      !ValidationUtils._validateCurrentAndHighestStreak(
        currentStreak,
        highestStreak
      )
    ) {
      // Error handled in validation function
      return;
    }

    const sheet = MainSheetConfig._getSheet();
    if (!sheet) return;

    const relevantRows = this.getRelevantRows();
    if (relevantRows.length !== completionData.length) {
      LoggerManager.handleError(
        `Mismatch between relevant rows (${relevantRows.length}) and provided data length (${completionData.length}) in setAllMainSheetData.`,
        true
      );
      return;
    }

    // --- Prepare Data for Batch Update ---
    // Get full ranges to overwrite existing data correctly
    const completionRange = MainSheetConfig.getCompletionDataRange();
    const bufferRange = MainSheetConfig.getBufferDataRange();

    // Initialize full-size arrays with empty/default values
    const numSheetRows = completionRange ? completionRange.getNumRows() : 0; // Use actual range size
    let paddedCompletionData = Array.from({ length: numSheetRows }, () => [""]); // Default empty
    let paddedBufferData = Array.from({ length: numSheetRows }, () => [0]); // Default 0 or null? Let's use 0.

    const firstDataSheetRow = MainSheetConfig.firstDataInputRow;

    // Populate padded arrays at the correct indices
    relevantRows.forEach((sheetRowIndex, i) => {
      const rangeRowIndex = sheetRowIndex - firstDataSheetRow; // 0-based index within the range
      if (rangeRowIndex >= 0 && rangeRowIndex < numSheetRows) {
        paddedCompletionData[rangeRowIndex] = [completionData[i]]; // Wrap in array for setValues
        paddedBufferData[rangeRowIndex] = [bufferData[i]]; // Wrap in array for setValues
      } else {
        LoggerManager.handleError(
          `Row index mismatch during data padding in setAllMainSheetData. SheetRow: ${sheetRowIndex}, RangeIndex: ${rangeRowIndex}`,
          false
        );
      }
    });

    // --- Perform Batch Update ---
    try {
      // Set Date and Streaks individually (simpler)
      this.setDateStr(dateStr);
      this.setCurrentStreak(currentStreak);
      this.setHighestStreak(highestStreak);

      // Set Completion and Buffer using ranges if they exist
      if (completionRange && paddedCompletionData.length > 0) {
        completionRange.setValues(paddedCompletionData);
        LoggerManager.logDebug(
          `Set completion data for range: ${completionRange.getA1Notation()}`
        );
      } else if (relevantRows.length > 0) {
        // Only log error if data was expected
        LoggerManager.logDebug(
          `Completion range not found or empty, skipping setValues.`
        );
      }

      if (bufferRange && paddedBufferData.length > 0) {
        bufferRange.setValues(paddedBufferData);
        LoggerManager.logDebug(
          `Set buffer data for range: ${bufferRange.getA1Notation()}`
        );
      } else if (relevantRows.length > 0) {
        LoggerManager.logDebug(
          `Buffer range not found or empty, skipping setValues.`
        );
      }

      SpreadsheetApp.flush(); // Flush all changes at once
      LoggerManager.logDebug(
        `setAllMainSheetData completed successfully for ${dateStr}.`
      );
    } catch (e) {
      LoggerManager.handleError(
        `Error during batch update in setAllMainSheetData: ${e.message}`
      );
    }
  },

  // --- History Sheet Interaction ---

  /**
   * Retrieves data for a specific date from the history sheet.
   * @param {string} dateStr - The date string (YYYY-MM-DD).
   * @returns {object | null} Object with { date, completionData, bufferData, currentStreak, highestStreak } or null if not found/error.
   */
  getHistoryDataAtDateStr: function (dateStr) {
    if (!DateManager._validateDateStr(dateStr, false)) {
      LoggerManager.logDebug(
        `getHistoryDataAtDateStr: Invalid date string format: ${dateStr}`
      );
      return null;
    }

    const rowIndex = this._getHistoryRowIndexForDate(dateStr);
    if (rowIndex === null) {
      LoggerManager.logDebug(
        `getHistoryDataAtDateStr: No history row found for date: ${dateStr}.`
      );
      return null;
    }

    return this._getHistoryDataAtRow(rowIndex); // Use 1-based index
  },

  /**
   * Retrieves data from a specific row (1-indexed) in the history sheet.
   * Handles JSON parsing with error checking.
   * @private
   * @param {number} rowIndex - The 1-indexed row number.
   * @returns {object | null} Object with data or null on error.
   */
  _getHistoryDataAtRow: function (rowIndex) {
    const sheet = HistorySheetConfig._getSheet();
    if (
      !sheet ||
      !ValidationUtils._validateHistorySheetRowIndex(rowIndex, false)
    ) {
      LoggerManager.logDebug(
        `_getHistoryDataAtRow: Invalid sheet or row index ${rowIndex}.`
      );
      return null;
    }

    const numCols = HistorySheetConfig._getNumberOfColumns();
    const startCol = HistorySheetConfig.dateColumn + 1; // 1-based start column

    try {
      const rowData = sheet
        .getRange(rowIndex, startCol, 1, numCols)
        .getValues()[0];

      const date = DateManager.determineDate(
        rowData[HistorySheetConfig.dateColumn]
      );
      if (!date) {
        LoggerManager.handleError(
          `Invalid date found in history row ${rowIndex}.`,
          false
        );
        // Decide whether to return null or continue with potentially bad data
        // return null;
      }

      let completionData, bufferData;
      // Safely parse JSON data
      try {
        completionData = JSON.parse(
          rowData[HistorySheetConfig.completionDataColumn] || "[]"
        );
        if (!ValidationUtils._validateCompletionData(completionData, false))
          throw new Error("Parsed completion data invalid");
      } catch (e) {
        LoggerManager.handleError(
          `Error parsing completionData in history row ${rowIndex}: ${
            e.message
          }. Raw: ${rowData[HistorySheetConfig.completionDataColumn]}`,
          false
        );
        completionData = this._getDefaultCompletionListForStorage(); // Use default on error
        Messages.showAlert(MessageTypes.DATA_PARSE_ERROR); // Inform user
      }
      try {
        bufferData = JSON.parse(
          rowData[HistorySheetConfig.bufferDataColumn] || "[]"
        );
        if (!ValidationUtils._validateBufferData(bufferData, false))
          throw new Error("Parsed buffer data invalid");
      } catch (e) {
        LoggerManager.handleError(
          `Error parsing bufferData in history row ${rowIndex}: ${
            e.message
          }. Raw: ${rowData[HistorySheetConfig.bufferDataColumn]}`,
          false
        );
        bufferData = this.getDefaultBufferList(); // Use default on error
        Messages.showAlert(MessageTypes.DATA_PARSE_ERROR); // Inform user
      }

      const currentStreak = rowData[HistorySheetConfig.currentStreakColumn];
      const highestStreak = rowData[HistorySheetConfig.highestStreakColumn];

      // Validate streak numbers
      if (
        !ValidationUtils._validateCurrentAndHighestStreak(
          currentStreak,
          highestStreak,
          false
        )
      ) {
        LoggerManager.handleError(
          `Invalid streak data in history row ${rowIndex}. Current: ${currentStreak}, Highest: ${highestStreak}`,
          false
        );
        // Decide how to handle - return null, or return potentially bad data?
        // Let's return the data but log the error.
      }

      const result = {
        date: date, // Return as Date object
        completionData: completionData,
        bufferData: bufferData,
        currentStreak: ValidationUtils._validateNonNegativeNumber(currentStreak)
          ? currentStreak
          : 0, // Default if invalid
        highestStreak: ValidationUtils._validateNonNegativeNumber(highestStreak)
          ? highestStreak
          : 0, // Default if invalid
      };
      LoggerManager.logDebug(
        `_getHistoryDataAtRow ${rowIndex}: ${JSON.stringify(result)}`
      );
      return result;
    } catch (e) {
      LoggerManager.handleError(
        `Failed to retrieve or process data for history row ${rowIndex}: ${e.message}`
      );
      return null;
    }
  },

  /**
   * Calculates the 1-indexed row number in the history sheet for a specific date.
   * Assumes dates are contiguous and sorted.
   * @private
   * @param {string} dateStr - The date string (YYYY-MM-DD).
   * @returns {number | null} The 1-indexed row number, or null if date is out of range or history is empty.
   */
  _getHistoryRowIndexForDate: function (dateStr) {
    const sheet = HistorySheetConfig._getSheet();
    if (!sheet) return null;

    const targetDate = DateManager.determineDate(dateStr);
    if (!targetDate) return null; // Invalid dateStr format

    const lastRowIndex = sheet.getLastRow();
    const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1; // 1-based

    if (lastRowIndex < firstDataSheetRow) {
      LoggerManager.logDebug(
        "_getHistoryRowIndexForDate: History sheet has no data rows."
      );
      return null; // No data rows exist
    }

    // Get the date from the last row to calculate offset
    const lastDate = this.getLastHistoryDate();
    if (!lastDate) {
      LoggerManager.logDebug(
        "_getHistoryRowIndexForDate: Could not retrieve last date from history."
      );
      // Fallback: could try searching? For now, assume failure.
      return null;
    }

    // Check if target date is within the range (first date needs to be checked too)
    const firstDate = this.getFirstHistoryDate();
    if (!firstDate || targetDate < firstDate || targetDate > lastDate) {
      LoggerManager.logDebug(
        `_getHistoryRowIndexForDate: Target date ${DateManager.determineFormattedDate(
          targetDate
        )} is outside history range (${DateManager.determineFormattedDate(
          firstDate
        )} - ${DateManager.determineFormattedDate(lastDate)}).`
      );
      return null;
    }

    // Calculate offset assuming contiguous dates
    const daysDifference = DateManager.daysBetween(targetDate, lastDate);
    const calculatedRowIndex = lastRowIndex - daysDifference;

    // Basic sanity check on calculated index
    if (calculatedRowIndex < firstDataSheetRow) {
      LoggerManager.handleError(
        `_getHistoryRowIndexForDate: Calculated row index ${calculatedRowIndex} is out of bounds for target date ${dateStr}. Last row: ${lastRowIndex}, Days diff: ${daysDifference}. This might indicate non-contiguous dates.`,
        false
      );
      // Consider implementing a search fallback here if needed.
      return null;
    }

    LoggerManager.logDebug(
      `_getHistoryRowIndexForDate: Date ${dateStr} maps to row ${calculatedRowIndex}.`
    );
    return calculatedRowIndex;
  },

  /**
   * Retrieves the first date recorded in the history sheet.
   * @returns {Date | null} The first date as a Date object, or null if history is empty/error.
   */
  getFirstHistoryDate: function () {
    const sheet = HistorySheetConfig._getSheet();
    if (!sheet) return null;

    const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1; // 1-based
    if (sheet.getLastRow() < firstDataSheetRow) {
      LoggerManager.logDebug(
        "getFirstHistoryDate: History sheet has no data rows."
      );
      return null; // No data
    }

    try {
      const dateVal = sheet
        .getRange(firstDataSheetRow, HistorySheetConfig.dateColumn + 1)
        .getValue();
      const date = DateManager.determineDate(dateVal);
      if (!date) {
        LoggerManager.handleError(
          `Invalid date format in first history data row (${firstDataSheetRow}). Value: ${dateVal}`,
          false
        );
        return null;
      }
      return date;
    } catch (e) {
      LoggerManager.handleError(
        `Error getting first history date: ${e.message}`
      );
      return null;
    }
  },

  /**
   * Retrieves the last date recorded in the history sheet.
   * @returns {Date | null} The last date as a Date object, or null if history is empty/error.
   */
  getLastHistoryDate: function () {
    const sheet = HistorySheetConfig._getSheet();
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1; // 1-based

    if (lastRow < firstDataSheetRow) {
      LoggerManager.logDebug(
        "getLastHistoryDate: History sheet has no data rows."
      );
      return null; // No data
    }

    try {
      const dateVal = sheet
        .getRange(lastRow, HistorySheetConfig.dateColumn + 1)
        .getValue();
      const date = DateManager.determineDate(dateVal);
      if (!date) {
        LoggerManager.handleError(
          `Invalid date format in last history data row (${lastRow}). Value: ${dateVal}`,
          false
        );
        return null;
      }
      return date;
    } catch (e) {
      LoggerManager.handleError(
        `Error getting last history date: ${e.message}`
      );
      return null;
    }
  },

  /**
   * Retrieves the first date of the current challenge from properties.
   * @returns {Date | null} The first challenge date as a Date object, or null if not set/error.
   */
  getFirstChallengeDate: function () {
    const dateStr = PropertyManager.getProperty(
      PropertyKeys.FIRST_CHALLENGE_DATE
    );
    if (!dateStr) return null;
    try {
      return DateManager.determineDate(dateStr);
    } catch (e) {
      LoggerManager.handleError(
        `Invalid first challenge date string in properties: ${dateStr}`,
        false
      );
      return null;
    }
  },

  /**
   * Ensures that a history entry exists for the given date. If not, creates a default entry.
   * Handles filling gaps between the last entry and the target date.
   *
   * @param {string} dateStr - The target date string (YYYY-MM-DD).
   * @returns {boolean} True if an entry exists or was successfully created, false on validation or creation error.
   */
  ensureHistoryDateEntry: function (dateStr) {
    LoggerManager.logDebug(`Ensuring history entry for date: ${dateStr}`);

    // 1. Validate the target date string format
    if (!DateManager._validateDateStr(dateStr, false)) {
      LoggerManager.handleError(
        `ensureHistoryDateEntry: Invalid date string format: ${dateStr}`,
        true
      );
      return false;
    }
    const targetDate = DateManager.determineDate(dateStr);

    // 2. Validate target date against challenge bounds (optional but good)
    const firstChallengeDate = this.getFirstChallengeDate();
    const today = DateManager.getToday();
    if (
      !firstChallengeDate ||
      targetDate < firstChallengeDate ||
      targetDate > today
    ) {
      LoggerManager.handleError(
        `ensureHistoryDateEntry: Target date ${dateStr} is outside the valid challenge range (${DateManager.determineFormattedDate(
          firstChallengeDate
        )} - ${DateManager.determineFormattedDate(today)}).`,
        false
      );
      // Allow proceeding maybe? Or return false? Let's return false for strictness.
      return false;
    }

    // 3. Check if the entry already exists
    const existingRowIndex = this._getHistoryRowIndexForDate(dateStr);
    if (existingRowIndex !== null) {
      LoggerManager.logDebug(
        `ensureHistoryDateEntry: Entry for ${dateStr} already exists at row ${existingRowIndex}.`
      );
      return true; // Entry exists
    }

    // 4. Entry does not exist. Need to create it (and potentially prior dates).
    LoggerManager.logDebug(
      `ensureHistoryDateEntry: No entry found for ${dateStr}. Creating required entries.`
    );
    const lastHistoryDate = this.getLastHistoryDate();

    let startDateToCreate;
    let previousDayData; // Data from the day *before* the first new entry

    if (!lastHistoryDate) {
      // History is empty, start from the very first challenge date
      startDateToCreate = firstChallengeDate;
      // No previous day data, use defaults (streak 0)
      previousDayData = {
        completionData: this._getDefaultCompletionListForStorage(),
        bufferData: this.getDefaultBufferList(),
        currentStreak: 0,
        highestStreak: 0,
        date: DateManager.getPreviousDate(startDateToCreate), // Dummy date for calculation start
      };
      LoggerManager.logDebug(
        `History empty, creating entries from first challenge date: ${DateManager.determineFormattedDate(
          startDateToCreate
        )}`
      );
    } else if (targetDate <= lastHistoryDate) {
      // This case *shouldn't* happen if _getHistoryRowIndexForDate worked correctly and returned null.
      // It implies a gap *before* the last date, which our current model doesn't handle well.
      // If this occurs, it might require recalculating from the date before the gap.
      LoggerManager.handleError(
        `ensureHistoryDateEntry: Target date ${dateStr} is before or same as last history date ${DateManager.determineFormattedDate(
          lastHistoryDate
        )}, but no row index was found. This indicates a potential data inconsistency or gap. Cannot reliably create entry.`,
        true
      );
      return false;
    } else {
      // Target date is after the last recorded date. Fill the gap.
      startDateToCreate = DateManager.getNextDate(lastHistoryDate); // Start creating from day after last entry
      previousDayData = this._getHistoryDataAtRow(
        this._getHistoryRowIndexForDate(
          DateManager.determineFormattedDate(lastHistoryDate)
        )
      );
      if (!previousDayData) {
        LoggerManager.handleError(
          `ensureHistoryDateEntry: Failed to retrieve data for the last history date ${DateManager.determineFormattedDate(
            lastHistoryDate
          )} needed to create new entries.`,
          true
        );
        return false;
      }
      LoggerManager.logDebug(
        `History exists, creating entries from ${DateManager.determineFormattedDate(
          startDateToCreate
        )} up to ${dateStr}`
      );
    }

    // 5. Loop from startDateToCreate up to targetDate, creating default entries
    let currentDate = new Date(startDateToCreate); // Clone start date
    let lastAppended = true; // Track success of appendRow

    while (currentDate <= targetDate && lastAppended) {
      const currentDateStr = DateManager.determineFormattedDate(currentDate);
      LoggerManager.logDebug(`Creating entry for: ${currentDateStr}`);

      // Calculate data for the current day based on previous day's data
      const calculatedData = this._calculateNextDayData(previousDayData); // Pass previous day's calculated/retrieved data

      // Format for appending
      const entryRow = [
        currentDateStr,
        JSON.stringify(calculatedData.completionData), // Store default empty completion
        JSON.stringify(calculatedData.bufferData),
        calculatedData.currentStreak,
        calculatedData.highestStreak,
      ];

      // Append the row
      lastAppended = HistorySheetConfig._appendRow(entryRow);

      if (lastAppended) {
        // Update previousDayData for the next iteration *using the newly calculated data*
        previousDayData = { ...calculatedData, date: new Date(currentDate) }; // Update date as well
      } else {
        LoggerManager.handleError(
          `ensureHistoryDateEntry: Failed to append row for date ${currentDateStr}. Aborting creation process.`,
          true
        );
        return false; // Stop if appending fails
      }

      // Move to the next day
      currentDate.setDate(currentDate.getDate() + 1);
    }

    return lastAppended; // Return true if all entries up to targetDate were created
  },

  /**
   * Saves the current state from the main sheet to the history sheet for a given date.
   * Assumes the history entry for this date already exists (e.g., via ensureHistoryDateEntry).
   * Then triggers propagation if needed.
   *
   * @param {string} dateStr - The date string (YYYY-MM-DD) to save data for.
   */
  saveMainSheetStateToHistory: function (dateStr) {
    LoggerManager.logDebug(
      `Saving main sheet state to history for date: ${dateStr}`
    );

    // 1. Get current state from Main Sheet UI
    const completionData = this.getCompletionData(); // Gets 1D array
    // Buffer and streaks are calculated during propagation, not saved directly from UI here.

    // 2. Find the row in the history sheet
    const rowIndex = this._getHistoryRowIndexForDate(dateStr);
    if (rowIndex === null) {
      LoggerManager.handleError(
        `Cannot save state: No history row found for date ${dateStr}. Ensure entry exists first.`,
        true
      );
      return;
    }

    // 3. Update only the completion data cell in the history sheet for that row
    const sheet = HistorySheetConfig._getSheet();
    if (!sheet) return;
    const completionCellCol = HistorySheetConfig.completionDataColumn + 1; // 1-based column

    try {
      const completionCell = sheet.getRange(rowIndex, completionCellCol);
      const jsonData = JSON.stringify(completionData);
      completionCell.setValue(jsonData);
      SpreadsheetApp.flush(); // Ensure completion data is written before propagation reads it
      LoggerManager.logDebug(
        `Updated completion data in history row ${rowIndex} for date ${dateStr}. Data: ${jsonData}`
      );

      // 4. Trigger propagation starting from the day *before* the saved date
      //    (or the saved date itself if it's the first day)
      const previousDate = DateManager.getPreviousDate(dateStr);
      const firstChallengeDate = this.getFirstChallengeDate();
      let propagationStartDate;

      if (!firstChallengeDate || previousDate < firstChallengeDate) {
        propagationStartDate = dateStr; // Start propagation from the current day if it's the first day
        LoggerManager.logDebug(
          `Propagation starting from the changed date itself: ${dateStr}`
        );
      } else {
        propagationStartDate = DateManager.determineFormattedDate(previousDate);
        LoggerManager.logDebug(
          `Propagation starting from the day before the change: ${propagationStartDate}`
        );
      }

      this._propagateHistoryChanges(propagationStartDate);
    } catch (e) {
      LoggerManager.handleError(
        `Failed to save completion data or propagate for date ${dateStr} in history row ${rowIndex}: ${e.message}`
      );
    }
  },

  // --- Data Propagation Logic ---

  /**
   * Calculates the buffer, streaks for the *next* day based on the *current* day's data.
   * @private
   * @param {object} currentDayData - Object with { date, completionData, bufferData, currentStreak, highestStreak }.
   * @returns {object} Object with calculated { completionData, bufferData, currentStreak, highestStreak } for the next day.
   */
  _calculateNextDayData: function (currentDayData) {
    LoggerManager.logDebug(
      `Calculating next day data based on: ${DateManager.determineFormattedDate(
        currentDayData.date
      )}`
    );
    const {
      completionData: currentCompletion,
      bufferData: currentBuffer,
      currentStreak,
      highestStreak,
    } = currentDayData;

    // Determine if challenge failed on the current day
    const challengeFailedCurrentDay = this._hasChallengeFailed(
      currentCompletion,
      currentBuffer
    );

    let nextStreak, nextHighest, nextBuffer;

    // Default next day's completion is empty until explicitly set
    const nextCompletion = this._getDefaultCompletionListForStorage();

    if (challengeFailedCurrentDay) {
      LoggerManager.logDebug(
        `Challenge failed on ${DateManager.determineFormattedDate(
          currentDayData.date
        )}. Resetting streak and buffer for next day.`
      );
      nextStreak = 0;
      nextHighest = highestStreak; // Highest streak persists through failure
      nextBuffer = this.getDefaultBufferList(); // Reset buffer to default
    } else {
      // Challenge succeeded, continue streak and calculate buffer
      nextStreak = currentStreak + 1;
      nextHighest = Math.max(nextStreak, highestStreak);

      // Check if buffer should increase based on the *newly achieved* next streak
      const shouldBoost = this._shouldBufferIncrease(nextStreak);
      if (shouldBoost) {
        LoggerManager.logDebug(
          `Buffer boost triggered for next day (streak ${nextStreak}).`
        );
      }

      // Calculate next buffer based on current completion and potential boost
      nextBuffer = currentBuffer.map((buffer, index) => {
        let newBuffer = buffer;
        if (shouldBoost) {
          newBuffer += 1; // Apply boost first
        }
        if (
          currentCompletion[index] === false ||
          currentCompletion[index] === ""
        ) {
          // If habit was missed today, decrement buffer for tomorrow
          newBuffer -= 1;
        }
        // Ensure buffer doesn't go below 0 (though failure check handles < 1)
        return Math.max(0, newBuffer);
      });
      LoggerManager.logDebug(
        `Calculated next buffer: ${JSON.stringify(
          nextBuffer
        )} (Boost: ${shouldBoost}, Current Completion: ${JSON.stringify(
          currentCompletion
        )})`
      );
    }

    return {
      completionData: nextCompletion,
      bufferData: nextBuffer,
      currentStreak: nextStreak,
      highestStreak: nextHighest,
    };
  },

  /**
   * Checks if the challenge failed on a given day based on completion and buffer.
   * Failure occurs if any habit is incomplete AND its buffer is less than 1.
   * @private
   * @param {Array<*>} completionData - Completion status array for the day.
   * @param {Array<number>} bufferData - Buffer status array for the day.
   * @returns {boolean} True if the challenge failed, false otherwise.
   */
  _hasChallengeFailed: function (completionData, bufferData) {
    if (
      !completionData ||
      !bufferData ||
      completionData.length !== bufferData.length
    ) {
      LoggerManager.handleError(
        "_hasChallengeFailed: Invalid input data.",
        false
      );
      return true; // Fail safe? Or return false? Let's assume failure on bad input.
    }
    for (let i = 0; i < completionData.length; i++) {
      // Check explicitly for false or empty string (common checkbox values)
      if (
        (completionData[i] === false || completionData[i] === "") &&
        bufferData[i] < 1
      ) {
        LoggerManager.logDebug(
          `Challenge failed on habit index ${i} (Completion: ${completionData[i]}, Buffer: ${bufferData[i]}).`
        );
        return true;
      }
    }
    return false;
  },

  /**
   * Determines if the buffer should increase based on the current streak and boost interval.
   * @private
   * @param {number} currentStreak - The current streak value being achieved.
   * @returns {boolean} True if the buffer should increase.
   */
  _shouldBufferIncrease: function (currentStreak) {
    const boostInterval = PropertyManager.getPropertyNumber(
      PropertyKeys.BOOST_INTERVAL
    );
    if (!boostInterval || boostInterval <= 0) {
      LoggerManager.logDebug(
        `_shouldBufferIncrease: Invalid boost interval (${boostInterval}), defaulting to no increase.`
      );
      return false;
    }
    // Increase happens *on* the day that is a multiple of the interval
    return currentStreak > 0 && currentStreak % boostInterval === 0;
  },

  /**
   * Propagates calculated changes (buffer, streaks) through the history sheet
   * starting from a given date up to the latest entry.
   *
   * @private
   * @param {string} startDateStr - The date string (YYYY-MM-DD) *from which* to start calculations (i.e., calculate data for the *next* day based on this date).
   */
  _propagateHistoryChanges: function (startDateStr) {
    LoggerManager.logDebug(
      `Starting history propagation from date: ${startDateStr}`
    );

    const sheet = HistorySheetConfig._getSheet();
    const lastHistoryDate = this.getLastHistoryDate();
    if (!sheet || !lastHistoryDate) {
      LoggerManager.logDebug(
        "Propagation stopped: History sheet not found or empty."
      );
      return;
    }
    const lastHistoryDateStr =
      DateManager.determineFormattedDate(lastHistoryDate);

    const startDate = DateManager.determineDate(startDateStr);
    if (!startDate || startDate >= lastHistoryDate) {
      LoggerManager.logDebug(
        `Propagation stopped: Start date ${startDateStr} is not before the last history date ${lastHistoryDateStr}.`
      );
      return; // Nothing to propagate
    }

    // Find the starting row index
    let startRowIndex = this._getHistoryRowIndexForDate(startDateStr); // 1-based
    if (startRowIndex === null) {
      LoggerManager.handleError(
        `Propagation error: Could not find history row for start date ${startDateStr}.`,
        true
      );
      return;
    }

    // Get all data from the start row onwards for efficient calculation
    const lastRow = sheet.getLastRow();
    const numRowsToProcess = lastRow - startRowIndex; // How many rows *after* the start row
    if (numRowsToProcess <= 0) {
      LoggerManager.logDebug(
        `Propagation stopped: No rows found after start date ${startDateStr}.`
      );
      return;
    }

    const numCols = HistorySheetConfig._getNumberOfColumns();
    const startCol = HistorySheetConfig.dateColumn + 1;
    const dataRange = sheet.getRange(
      startRowIndex,
      startCol,
      numRowsToProcess + 1,
      numCols
    );
    const historyData = dataRange.getValues();
    LoggerManager.logDebug(
      `Retrieved ${historyData.length} rows for propagation starting from row ${startRowIndex}.`
    );

    // Arrays to store calculated values for batch update
    const bufferUpdates = [];
    const currentStreakUpdates = [];
    const highestStreakUpdates = [];

    // Get the initial "previous day" data from the first row retrieved
    let previousDayData;
    try {
      previousDayData = {
        date: DateManager.determineDate(
          historyData[0][HistorySheetConfig.dateColumn]
        ),
        completionData: JSON.parse(
          historyData[0][HistorySheetConfig.completionDataColumn] || "[]"
        ),
        bufferData: JSON.parse(
          historyData[0][HistorySheetConfig.bufferDataColumn] || "[]"
        ),
        currentStreak: historyData[0][HistorySheetConfig.currentStreakColumn],
        highestStreak: historyData[0][HistorySheetConfig.highestStreakColumn],
      };
      // Validate the initial data
      if (
        !ValidationUtils._validateCompletionAndBufferData(
          previousDayData.completionData,
          previousDayData.bufferData,
          false
        ) ||
        !ValidationUtils._validateCurrentAndHighestStreak(
          previousDayData.currentStreak,
          previousDayData.highestStreak,
          false
        )
      ) {
        throw new Error("Initial propagation data invalid.");
      }
    } catch (e) {
      LoggerManager.handleError(
        `Propagation error: Failed to parse or validate data for starting date ${startDateStr} at row ${startRowIndex}. Error: ${e.message}`,
        true
      );
      return;
    }

    // Iterate through the *subsequent* rows in the retrieved data
    for (let i = 1; i < historyData.length; i++) {
      const currentRowIndexInSheet = startRowIndex + i; // 1-based index in the actual sheet
      const currentDayStoredCompletion = JSON.parse(
        historyData[i][HistorySheetConfig.completionDataColumn] || "[]"
      ); // Get stored completion for the day being calculated

      // Calculate buffer and streaks for row `i` based on `previousDayData` (from row i-1)
      const calculatedData = this._calculateNextDayData(previousDayData);

      // The completion data for row `i` isn't calculated, it's preserved from what was stored.
      // However, we need it to check for failure *on this day* to correctly calculate the *next* day.
      const currentDayActualCompletion = currentDayStoredCompletion; // Use the stored completion for failure check

      // Update the arrays for batch writing (Buffer, Current, Highest)
      // We are writing the *calculated* values for row `i`.
      bufferUpdates.push([JSON.stringify(calculatedData.bufferData)]);
      currentStreakUpdates.push([calculatedData.currentStreak]);
      highestStreakUpdates.push([calculatedData.highestStreak]);

      // Prepare `previousDayData` for the *next* iteration (i+1)
      // Use the calculated buffer/streaks, but the *stored* completion for this day (row i)
      previousDayData = {
        date: DateManager.determineDate(
          historyData[i][HistorySheetConfig.dateColumn]
        ),
        completionData: currentDayActualCompletion, // Use stored completion for next calculation base
        bufferData: calculatedData.bufferData, // Use calculated buffer
        currentStreak: calculatedData.currentStreak, // Use calculated streak
        highestStreak: calculatedData.highestStreak, // Use calculated highest
      };
      LoggerManager.logDebug(
        `Propagation: Row ${currentRowIndexInSheet} calculated: Buffer=${JSON.stringify(
          calculatedData.bufferData
        )}, CStreak=${calculatedData.currentStreak}, HStreak=${
          calculatedData.highestStreak
        }`
      );
    }

    // --- Perform Batch Updates ---
    // Write to rows starting from `startRowIndex + 1`
    const firstUpdateRow = startRowIndex + 1;
    try {
      if (bufferUpdates.length > 0) {
        sheet
          .getRange(
            firstUpdateRow,
            HistorySheetConfig.bufferDataColumn + 1,
            bufferUpdates.length,
            1
          )
          .setValues(bufferUpdates);
        sheet
          .getRange(
            firstUpdateRow,
            HistorySheetConfig.currentStreakColumn + 1,
            currentStreakUpdates.length,
            1
          )
          .setValues(currentStreakUpdates);
        sheet
          .getRange(
            firstUpdateRow,
            HistorySheetConfig.highestStreakColumn + 1,
            highestStreakUpdates.length,
            1
          )
          .setValues(highestStreakUpdates);
        SpreadsheetApp.flush(); // Flush propagation changes
        LoggerManager.logDebug(
          `Propagation completed. Updated ${bufferUpdates.length} rows starting from row ${firstUpdateRow}.`
        );
      } else {
        LoggerManager.logDebug(
          "Propagation finished: No rows needed updating."
        );
      }
    } catch (e) {
      LoggerManager.handleError(
        `Error during batch update in propagation: ${e.message}`
      );
    }
  },

  // --- High-Level Display Logic ---

  /**
   * Loads data for the specified date onto the main sheet UI.
   * Ensures history entry exists first.
   * @param {string} dateStr - The date string (YYYY-MM-DD) to display.
   */
  displayDate: function (dateStr) {
    LoggerManager.logDebug(`Attempting to display date: ${dateStr}`);

    // 1. Ensure the history entry exists, creating it if necessary.
    if (!this.ensureHistoryDateEntry(dateStr)) {
      // If ensure fails (e.g., invalid date range), show an error or default.
      Messages.showAlert(MessageTypes.INVALID_DATE); // Show generic invalid date
      // Attempt to display today's date instead as a fallback
      const todayStr = DateManager.getTodayStr();
      LoggerManager.logDebug(
        `ensureHistoryDateEntry failed for ${dateStr}. Attempting to display today: ${todayStr}`
      );
      if (dateStr !== todayStr) {
        // Avoid infinite loop if today itself fails
        this.displayDate(todayStr);
      } else {
        LoggerManager.handleError(
          `Failed to display even today's date (${todayStr}). Check history/properties.`,
          true
        );
      }
      return;
    }

    // 2. Retrieve the data for the date from history
    const data = this.getHistoryDataAtDateStr(dateStr);
    if (!data) {
      LoggerManager.handleError(
        `Failed to retrieve history data for ${dateStr} even after ensuring entry.`,
        true
      );
      // Could try loading default state?
      MainSheetConfig.resetChallengeDataUI(); // Reset UI as a fallback
      return;
    }

    // 3. Set the retrieved data onto the main sheet UI
    this.setAllMainSheetData(
      DateManager.determineFormattedDate(data.date), // Use formatted date from retrieved data
      data.completionData,
      data.bufferData,
      data.currentStreak,
      data.highestStreak
    );

    // 4. Ensure checkboxes are present (might have been cleared if data was defaulted)
    MainSheetConfig.insertCompletionCheckboxes();

    LoggerManager.logDebug(
      `Successfully displayed data for ${dateStr} on the main sheet.`
    );
  },

  /**
   * Handles the core logic after a cell edit event on the main sheet.
   * Determines the type of edit and triggers appropriate actions like saving,
   * propagating, or displaying data.
   *
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where the edit occurred.
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited range.
   * @param {*} oldValue - The value before the edit (can be undefined for multi-cell).
   */
  handleCellEdit: function (sheet, range, oldValue) {
    // --- Pre-checks ---
    if (!sheet || !range) {
      LoggerManager.logDebug(
        "handleCellEdit: Invalid sheet or range object received."
      );
      return;
    }
    if (sheet.getSheetId() !== MainSheetConfig._getSheet().getSheetId()) {
      // Handle edits on other sheets if necessary (e.g., lock history sheet)
      if (sheet.getSheetId() === HistorySheetConfig._getSheet().getSheetId()) {
        LoggerManager.logDebug(
          "Edit detected on history sheet. Reverting changes."
        );
        Messages.showAlert("The history sheet is read-only."); // Inform user
        // Attempt to revert based on oldValue - this might be unreliable for history
        try {
          range.setValue(oldValue);
          SpreadsheetApp.flush();
        } catch (e) {
          /* ignore */
        }
      }
      LoggerManager.logDebug(
        `Edit on non-main sheet (${sheet.getName()}), ignoring.`
      );
      return;
    }

    // Check if the edit is on a locked range/cell *before* processing mode logic
    if (MainSheetConfig.includesLockedRange(range)) {
      LoggerManager.logDebug(
        `Edit occurred on a locked range/column (${range.getA1Notation()}). Reverting.`
      );
      MainSheetConfig.maintainCellValue(range, oldValue);
      // Optionally show a message, but maintainCellValue might be sufficient
      return;
    }

    // --- Mode-Specific Handling ---
    const mode = PropertyManager.getProperty(PropertyKeys.MODE);
    const rangeCells = this.getRangeCells(range); // Get cell A1 notations
    const columns = this.getRangeColumns(range); // Get column numbers

    if (mode === ModeTypes.TERMINATED) {
      LoggerManager.logDebug(
        "Edit attempt in TERMINATED mode. Showing reminder and reverting."
      );
      Messages.showAlert(MessageTypes.TERMINATION_REMINDER);
      MainSheetConfig.maintainCellValue(range, oldValue); // Revert the change
      return; // Stop further processing
    } else if (mode === ModeTypes.HABIT_IDEATION) {
      // Check if the 'set habit' checkbox was ticked
      if (
        rangeCells.includes(MainSheetConfig.setterRanges.setHabit) &&
        range.getValue() === true
      ) {
        LoggerManager.logDebug(
          "'Set Habit' checkbox ticked in Habit Ideation mode."
        );
        HabitManager.processHabitSpreadConfirmation(range); // HabitManager handles the confirmation flow
      } else if (columns.includes(MainSheetConfig.activityDataColumn)) {
        LoggerManager.logDebug("Activity column edited during Habit Ideation.");
        // No specific action needed here, user is defining habits.
      } else if (
        rangeCells.includes(MainSheetConfig.setterRanges.resetHour) ||
        rangeCells.includes(MainSheetConfig.setterRanges.boostInterval)
      ) {
        LoggerManager.logDebug(
          "Reset Hour or Boost Interval edited during Habit Ideation."
        );
        // Validation happens when 'Set Habit' is clicked.
      } else {
        LoggerManager.logDebug(
          `Other edit (${range.getA1Notation()}) in Habit Ideation mode. Usually ignored or handled by includesLockedRange.`
        );
        // If it wasn't caught by includesLockedRange, maybe revert it?
        // MainSheetConfig.maintainCellValue(range, oldValue); // Optional: revert unexpected edits
      }
      PropertyManager.setDocumentProperties(); // Save any property changes potentially made
      return; // Don't proceed to challenge mode logic
    } else if (mode === ModeTypes.CHALLENGE) {
      let completionChanged = false;
      let dateChanged = false;

      // Check for activity column changes (for emoji spread check)
      if (columns.includes(MainSheetConfig.activityDataColumn)) {
        LoggerManager.logDebug(
          "Activity column edited during Challenge Mode. Flagging for check."
        );
        PropertyManager.setProperty(
          PropertyKeys.ACTIVITIES_COLUMN_UPDATED,
          BooleanTypes.TRUE
        );
        // Note: The actual check/revert happens in getRelevantRows() when data is next accessed.
      }

      // Check for completion data changes
      if (columns.includes(MainSheetConfig.completionDataColumn)) {
        LoggerManager.logDebug("Completion data edited.");
        completionChanged = true;
        PropertyManager.setProperty(
          PropertyKeys.LAST_COMPLETION_UPDATE,
          DateManager.getNow()
        );
        PropertyManager.setProperty(
          PropertyKeys.LAST_UPDATE,
          LastUpdateTypes.COMPLETION
        );
        // Ensure checkboxes are present, especially after paste/multi-edit
        if (
          oldValue === undefined ||
          range.getNumRows() > 1 ||
          range.getNumColumns() > 1
        ) {
          LoggerManager.logDebug(
            "Multi-cell edit or paste in completion column. Ensuring checkboxes."
          );
          // Delay insertion slightly to allow paste value to settle? Maybe not needed.
          SpreadsheetApp.flush(); // Ensure value is set before inserting checkbox might help
          MainSheetConfig.insertCompletionCheckboxes();
        }
      }

      // Check for date selector changes
      if (rangeCells.includes(MainSheetConfig.dateCell)) {
        LoggerManager.logDebug("Date selector edited.");
        dateChanged = true;
        PropertyManager.setProperty(
          PropertyKeys.LAST_DATE_SELECTOR_UPDATE,
          DateManager.getNow()
        );
        PropertyManager.setProperty(
          PropertyKeys.LAST_UPDATE,
          LastUpdateTypes.DATE_SELECTOR
        );
      }

      // --- Process Changes ---
      if (completionChanged || dateChanged) {
        let previousDateStr = null;
        if (dateChanged) {
          // Try to format the OLD value of the date cell
          try {
            previousDateStr = DateManager.determineFormattedDate(oldValue);
          } catch (e) {
            LoggerManager.logDebug(
              `Could not format previous date value: ${oldValue}. Will proceed without saving previous date.`
            );
            // Force lastUpdate to completion if date selector was changed *from* an invalid value
            PropertyManager.setProperty(
              PropertyKeys.LAST_UPDATE,
              LastUpdateTypes.COMPLETION
            );
          }
        } else {
          // If only completion changed, the "previous date" is the current date on the sheet
          const currentDateOnSheet = this.getDate();
          if (currentDateOnSheet) {
            previousDateStr =
              DateManager.determineFormattedDate(currentDateOnSheet);
          } else {
            LoggerManager.handleError(
              "Could not get current date from sheet while handling completion change.",
              true
            );
            return; // Cannot proceed without a valid date context
          }
        }

        // Get the new date selected (or the current date if only completion changed)
        const currentDateOnSheet = this.getDate(); // Re-get the current value
        let selectedDateStr;
        if (!currentDateOnSheet) {
          Messages.showAlert(MessageTypes.INVALID_DATE);
          LoggerManager.logDebug(
            "Invalid date found in date cell after edit. Defaulting to today."
          );
          selectedDateStr = DateManager.getTodayStr();
          this.setDateStr(selectedDateStr); // Update the sheet UI
        } else {
          selectedDateStr =
            DateManager.determineFormattedDate(currentDateOnSheet);
          // Validate range (ensure it's not before first challenge date or after today)
          if (!DateManager._validateDateStrRange(selectedDateStr, false)) {
            Messages.showAlert(MessageTypes.INVALID_DATE);
            LoggerManager.logDebug(
              `Selected date ${selectedDateStr} is out of range. Defaulting to today.`
            );
            selectedDateStr = DateManager.getTodayStr();
            this.setDateStr(selectedDateStr); // Update the sheet UI
          }
        }

        LoggerManager.logDebug(
          `Processing edit. Previous Date: ${previousDateStr}, Selected Date: ${selectedDateStr}, Completion Changed: ${completionChanged}, Date Changed: ${dateChanged}`
        );
        this.renewChecklist(previousDateStr, selectedDateStr);
      } else if (
        PropertyManager.getProperty(PropertyKeys.ACTIVITIES_COLUMN_UPDATED) ===
        BooleanTypes.TRUE
      ) {
        // If only activities changed, just save properties. Check happens later.
        LoggerManager.logDebug(
          "Only activities column changed, no immediate action besides flagging."
        );
      } else {
        LoggerManager.logDebug(
          `Edit in range ${range.getA1Notation()} did not trigger core logic update.`
        );
      }
    } else {
      LoggerManager.handleError(`Unhandled application mode: ${mode}`, true);
    }

    // Save properties at the end of handling
    PropertyManager.setDocumentProperties();
  },

  /**
   * Coordinates saving the previous date's state (if necessary) and loading the selected date's data.
   * Determines whether to save based on which was updated more recently: completion or date selector.
   *
   * @param {string | null} previousDateStr - The date string (YYYY-MM-DD) that was displayed *before* the edit, or null if invalid/not applicable.
   * @param {string} selectedDateStr - The new date string (YYYY-MM-DD) to display.
   */
  renewChecklist: function (previousDateStr, selectedDateStr) {
    const lastUpdateType = PropertyManager.getProperty(
      PropertyKeys.LAST_UPDATE
    );

    // Save the state of the *previous* date ONLY if completion was the last change AND previousDateStr is valid
    if (
      lastUpdateType === LastUpdateTypes.COMPLETION &&
      previousDateStr &&
      DateManager._validateDateStr(previousDateStr, false)
    ) {
      LoggerManager.logDebug(
        `Completion was the last update. Saving state for previous date: ${previousDateStr}`
      );
      this.saveMainSheetStateToHistory(previousDateStr); // This also triggers propagation
    } else {
      LoggerManager.logDebug(
        `Skipping save for previous date (${previousDateStr}). Last update was: ${lastUpdateType}.`
      );
    }

    // Always display the data for the (potentially new) selected date
    LoggerManager.logDebug(
      `Displaying data for selected date: ${selectedDateStr}`
    );
    this.displayDate(selectedDateStr); // This ensures history entry and loads data
  },
};

// Freeze the handler object
Object.freeze(DataHandler);
