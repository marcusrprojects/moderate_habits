/** OnlyCurrentDoc */

/**
 * Configuration constants for various sheets, implementing subtle changes in names, columns, rows, and cells.
 * These are JavaScript singleton objects, ensuring that only one instance exists and providing a global access point.
 * The properties within these objects are constant and immutable, reinforcing the integrity of the configuration.
 */

/**
 * Validation methods to ensure the integrity of data by throwing errors to prevent improper tampering.
 * These errors are designed for internal validation rather than user-level errors, so they should only produce generic user messages.
 * Methods prefixed with "__" are internal helpers, designed to return true/false without logging errors directly.
 */
const SheetConfig = {
  /**
   * @constant {string} mainColor - Primary background color, a very light cream shade.
   */
  mainColor: "#FFF9F5",

  /**
   * @constant {string} secondaryColor - Secondary color, a light peach hue.
   */
  secondaryColor: "#F2DFCE",

  /**
   * Validates an array of row numbers.
   * @param {Array<number>} rows - The rows to validate.
   * @returns {boolean} True if valid.
   */
  _validateRows: function (rows, throwError = true) {
    if (!UtilsManager.__validateArray(rows)) {
      LoggerManager.handleError(`Invalid rows: Must be an array`, throwError);
      return false;
    }
    if (rows.length === 0) {
      LoggerManager.handleError(
        `Invalid rows: Must be a non-empty array`,
        throwError
      );
      return false;
    }
    return true; // Return true if all rows are valid
  },
};

/** Configuration for managing common utilities within the application, non-specific to spreadsheet data. */
const UtilsManager = {
  /**
   * Regular expression pattern for detecting emojis.
   *
   * This pattern matches surrogate pairs in Unicode that represent emojis.
   * Specifically, it targets the Unicode range for high-surrogate characters
   * (D800–DBFF) followed by low-surrogate characters (DC00–DFFF).
   *
   * @constant {RegExp}
   * @default
   */
  emojiPattern: /(\p{Emoji_Presentation}|\p{Extended_Pictographic})+/gu,

  /**
   * Validates if the input is an array.
   * @param {any} data - The data to validate.
   * @param {boolean} [throwError=true] - Whether to throw an error on invalid input.
   * @returns {boolean} True if valid, False if invalid.
   */
  __validateArray: function (data) {
    return Array.isArray(data);
  },

  /**
   * Validates if the input is a valid date.
   * @param {any} data - The data to validate.
   * @returns {boolean} True if valid.
   */
  __validateDate: function (date) {
    return date instanceof Date && !isNaN(date);
  },

  /**
   * Validates if the input is a non-negative number.
   * @param {any} data - The data to validate.
   * @returns {boolean} True if valid.
   */
  __validateNonNegativeNumber: function (data) {
    if (typeof data !== "number" || data < 0) {
      return false;
    }
    return true;
  },

  /**
   * Validates if the input is a non-negative integer.
   * @param {any} data - The data to validate.
   * @returns {boolean} True if valid.
   */
  __validateNonNegativeInteger: function (data) {
    if (typeof data !== "number" || data < 0 || !Number.isInteger(data)) {
      return false;
    }
    return true;
  },
};

/** Configuration for managing activities within the application */
const ActivitiesManager = {
  /**
   * Validates the completion data array.
   * @param {Array<any>} completionData - The completion data to validate.
   * @returns {boolean} True if valid.
   */
  _validateCompletionData: function (completionData, throwError = true) {
    if (!UtilsManager.__validateArray(completionData)) {
      LoggerManager.handleError(
        `Invalid input: completion data must be an array.`,
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Validates the buffer data array.
   * @param {Array<any>} bufferData - The buffer data to validate.
   * @returns {boolean} True if valid.
   */
  _validateBufferData: function (bufferData, throwError = true) {
    if (!UtilsManager.__validateArray(bufferData)) {
      LoggerManager.handleError(
        `Invalid input: buffer data must be an array.`,
        throwError
      );
      return false;
    }
    // Ensure every value in the array is a number
    const allNumbers = bufferData.every(
      (item) => typeof item === "number" && !isNaN(item)
    );
    if (!allNumbers) {
      LoggerManager.handleError(
        `Invalid input: buffer data ${JSON.stringify(
          bufferData
        )} must contain only valid numbers.`,
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Validates that completion data and buffer data are arrays of the same length.
   * @param {Array<any>} completionData - The completion data to validate.
   * @param {Array<any>} bufferData - The buffer data to validate.
   * @returns {boolean} True if valid.
   */
  _validateCompletionAndBufferData: function (
    completionData,
    bufferData,
    throwError = true
  ) {
    this._validateCompletionData(completionData, throwError);
    this._validateBufferData(bufferData, throwError);
    if (completionData.length !== bufferData.length) {
      LoggerManager.handleError(
        `Invalid input: completionData and bufferData must be of the same length.`,
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Retrieves filtered completion data based on relevant rows.
   * @returns {Array<any>} The filtered completion data.
   */
  getCompletionData: function () {
    const relevantRows = MainSheetConfig.getRelevantRows();
    const allData = MainSheetConfig.getCompletionDataRange().getValues();

    // Filter data to only include relevant rows
    const filteredData = relevantRows
      .map((rowIndex) => allData[rowIndex - MainSheetConfig.firstDataInputRow])
      .flat(); // -1 to adjust for 1-indexed rows

    LoggerManager.logDebug(`getCompletionData filteredData: ${filteredData}.`);
    return filteredData;
  },

  /**
   * Retrieves filtered buffer data based on relevant rows.
   * @returns {Array<any>} The filtered buffer data.
   */
  getBufferData: function () {
    const relevantRows = MainSheetConfig.getRelevantRows();
    const allData = MainSheetConfig.getBufferDataRange().getValues();

    // Filter data to only include relevant rows
    const filteredData = relevantRows.map(
      (rowIndex) => allData[rowIndex - MainSheetConfig.firstDataInputRow]
    ); // -1 to adjust for 1-indexed rows

    LoggerManager.logDebug(`getBufferData filteredData: ${filteredData}`);

    return filteredData;
  },
};

/** Configuration for managing streaks within the application **/
const StreakManager = {
  /**
   * Validates the current streak value.
   * @param {number} currentStreak - The current streak to validate.
   * @returns {boolean} True if valid.
   */
  _validateCurrentStreak: function (currentStreak, throwError = true) {
    if (!UtilsManager.__validateNonNegativeNumber(currentStreak)) {
      LoggerManager.handleError(
        `Invalid input: currentStreak must be a non-negative number.`,
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Validates the highest streak value.
   * @param {number} highestStreak - The highest streak to validate.
   * @returns {boolean} True if valid.
   */
  _validateHighestStreak: function (highestStreak, throwError = true) {
    if (!UtilsManager.__validateNonNegativeNumber(highestStreak)) {
      LoggerManager.handleError(
        `Invalid input: highestStreak must be a non-negative number.`,
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Validates the current streak and highest streak values.
   * Ensures that the highest streak is always equal to or greater than the current streak.
   * @param {number} currentStreak - The current streak to validate.
   * @param {number} highestStreak - The highest streak to validate.
   * @returns {boolean} True if valid.
   */
  _validateCurrentAndHighestStreak: function (
    currentStreak,
    highestStreak,
    throwError = true
  ) {
    this._validateCurrentStreak(currentStreak, throwError);
    this._validateHighestStreak(highestStreak, throwError);

    if (currentStreak > highestStreak) {
      LoggerManager.handleError(
        `Invalid input: highestStreak must always be equal to or greater than currentStreak.`,
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Retrieves the current streak value.
   * @returns {number} The current streak.
   */
  getCurrentStreak: function () {
    return MainSheetConfig.getSheetValue(MainSheetConfig.currentStreakCell);
  },

  /**
   * Sets the current streak value.
   * @param {number} currentStreak - The current streak value to set.
   */
  setCurrentStreak: function (currentStreak) {
    MainSheetConfig.setSheetValue(
      MainSheetConfig.currentStreakCell,
      currentStreak
    );
  },

  /**
   * Retrieves the highest streak value.
   * @returns {number} The highest streak.
   */
  getHighestStreak: function () {
    return MainSheetConfig.getSheetValue(MainSheetConfig.highestStreakCell);
  },

  /**
   * Sets the highest streak value.
   * @param {number} highest - The highest streak value to set.
   */
  setHighestStreak: function (highestStreak) {
    MainSheetConfig.setSheetValue(
      MainSheetConfig.highestStreakCell,
      highestStreak
    );
  },
};

/**
 * MainSheetConfig manages the configuration and operations for the "main" sheet.
 * This includes methods to retrieve, update, and validate data related to activities, completion, buffer, and streaks.
 *
 * This object follows the singleton pattern, ensuring a single instance.
 */
const MainSheetConfig = {
  /**
   * @type {string} sheetName - The name of the main sheet.
   */
  sheetName: "main",

  /**
   * @type {number} firstDataInputRow - The 1-indexed row number for the first data input.
   */
  firstDataInputRow: 3,

  /**
   * @type {number} activityDataColumn - The 1-indexed column index for activity data.
   */
  activityDataColumn: 4,

  /**
   * @type {number} completionDataColumn - The 1-indexed column index for completion data.
   */
  completionDataColumn: 5,

  /**
   * @type {number} bufferDataColumn - The 1-indexed column index for buffer data.
   */
  bufferDataColumn: 6,

  /**
   * @type {number} streaksDataColumn - The column number for streaks data.
   */
  streaksDataColumn: 2,

  /**
   * @type {Object} setterRanges - Cell references for various configuration settings.
   * @property {string} setHabit - The cell used for setting habits.
   * @property {string} resetHour - The cell where the reset hour is defined.
   * @property {string} boostInterval - The cell where the boost interval is defined.
   */
  setterRanges: {
    setHabit: "H3",
    resetHour: "H6",
    boostInterval: "H8",
  },

  /**
   * @type {Object} setterLabelRanges - Label cell references for the corresponding settings.
   * @property {string} setHabit - The cell reference for the habit-setting label.
   * @property {string} resetHour - The cell reference for the reset hour label.
   * @property {string} boostInterval - The cell reference for the boost interval label.
   */
  setterLabelRanges: {
    setHabit: "H2",
    resetHour: "H5",
    boostInterval: "H7",
  },

  /**
   * @type {Object} setterLabels - The displayed labels for each setting.
   * @property {string} setHabit - Label for setting habits.
   * @property {string} resetHour - Label for defining the reset hour.
   * @property {string} boostInterval - Label for defining the boost interval.
   */
  setterLabels: {
    setHabit: "set habits",
    resetHour: "reset hour",
    boostInterval: "boost interval",
  },

  /**
   * @type {Object} setterNotes - Descriptions explaining the purpose of each setting.
   * @property {string} resetHour - Explanation for the reset hour.
   * @property {string} boostInterval - Explanation for the boost interval.
   * @property {string} setHabit - Explanation for habit-setting options.
   */
  setterNotes: {
    resetHour:
      "Define the hour at which the daily reset occurs. Default is 3, for a 3 A.M. reset.",
    boostInterval:
      "Define the interval (in days) before the next boost occurs. Default is 7, for +1 rest day across all categories weekly.",
    setHabit:
      "Check this box to set your habit spread. Data will be validated for correctness.",
  },

  /**
   * @type {string} currentStreakCell - The cell address for the current streak.
   */
  currentStreakCell: "B3",

  /**
   * @type {string} highestStreakCell - The cell address for the highest streak.
   */
  highestStreakCell: "B6",

  /**
   * @type {string} dateCell - The cell address for the date.
   */
  dateCell: "B9",

  /**
   * @constant {number} defaultBuffer - The buffer that each habit defaults to upon starting a challenge.
   */
  defaultBuffer: 1,

  /**
   * @type {Object} headerLabels - An object that stores the header labels for various columns.
   * @property {string} activities - Label for activities.
   * @property {string} completion - Label for completion data.
   * @property {string} buffer - Label for buffer data.
   * @property {string} currentStreak - Label for the current streak.
   * @property {string} highestStreak - Label for the highest streak.
   * @property {string} dateSelector - Label for the date selector.
   */
  headerLabels: {
    activities: "activities", // Default values (can be modified by user)
    completion: "completion",
    buffer: "buffer",
    currentStreak: "current streak",
    highestStreak: "highest streak",
    dateSelector: "date selector",
  },

  /**
   * An array of header range strings indicating which cells are used as headers on the sheet.
   * These header ranges are important for checking if a user is attempting to edit locked or protected data.
   *
   */
  headerLabelRanges: {
    activities: "D2", // The current range where the 'activities' label is located
    completion: "E2",
    buffer: "F2",
    currentStreak: "B2",
    highestStreak: "B5",
    dateSelector: "B8",
  },

  /**
   * @type {number} resetHourDefault - The hour that the main sheet resets at.
   */
  resetHourDefault: 3,

  /**
   * Retrieves the sheet object for the main sheet.
   * @returns {Sheet} The sheet object.
   */
  getSheet: function () {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
  },

  /**
   * Retrieves the value from the specified sheet cell.
   * @param {string} cell - The cell's A1 notation to get the value from.
   * @returns {*} - The value of the specified cell.
   */
  getSheetValue: function (cell) {
    return MainSheetConfig.getSheet().getRange(cell).getValue();
  },

  /**
   * Sets the value in the specified sheet cell.
   * @param {string} cell - The cell's A1 notation to set the value in.
   * @param {*} value - The value to set in the specified cell.
   */
  setSheetValue: function (cell, value) {
    try {
      MainSheetConfig.getSheet().getRange(cell).setValue(value);
      SpreadsheetApp.flush();
      LoggerManager.logDebug(`Value set successfully for cell ${cell}.`);
    } catch (e) {
      LoggerManager.handleError(
        `Failed to set value for cell ${cell}: ${e.message}`
      );
    }
  },

  /**
   * Determines the first row that contains a habit, identified by an emoji.
   * @returns {number} The row number of the first habit.
   */
  getFirstDataRow: function () {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();

    for (let i = 1; i <= lastRow; i++) {
      // 1-indexed, so start at 1.
      const emojiCell = sheet.getRange(i, this.activityDataColumn).getValue();
      if (emojiCell && UtilsManager.emojiPattern.test(emojiCell)) {
        // Regex to detect emoji
        return i;
      }
    }

    LoggerManager.logDebug(`Invalid Main Sheet: No habits...`);
  },

  /**
   * Checks if the current emoji spread matches the stored emoji spread.
   * If the spreads don't match, it alerts the user and resets the emoji cells to their original state.
   */
  checkEmojiSpread: function () {
    const currentEmojiSpread = this.getCurrentEmojiSpread(); // Get emoji spread from sheet
    const storedEmojiSpread = JSON.parse(
      PropertyManager.getProperty(PropertyKeys.EMOJI_LIST)
    );

    // Compare current emoji spread with stored spread
    if (!this.emojiSpreadMatches(currentEmojiSpread, storedEmojiSpread)) {
      const ui = SpreadsheetApp.getUi();
      Messages.showAlert(MessageTypes.HABIT_SPREAD_RESET);

      // Automatically revert the emoji spread to the stored one
      this.revertEmojiSpread(storedEmojiSpread);
    }
  },

  /**
   * Retrieves the current list of emoji and their corresponding row numbers from the main sheet.
   * @returns {Array<Object>} List of objects containing the emoji and their row numbers.
   */
  getCurrentEmojiSpread: function () {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    let emojiList = [];

    for (let row = 1; row <= lastRow; row++) {
      const emoji = sheet.getRange(row, this.activityDataColumn).getValue();
      if (emoji.match(UtilsManager.emojiPattern)) {
        emojiList.push({ emoji: emoji, row: row });
      }
    }

    return emojiList;
  },

  /**
   * Compares the current emoji spread with the stored emoji spread.
   * @param {Array<Object>} currentEmojiSpread - The current emoji spread from the sheet.
   * @param {Array<Object>} storedEmojiSpread - The stored emoji spread.
   * @returns {boolean} True if the emoji spreads match, false otherwise.
   */
  emojiSpreadMatches: function (currentEmojiSpread, storedEmojiSpread) {
    // Simple check: Do the emojis and row positions match?
    if (currentEmojiSpread.length !== storedEmojiSpread.length) {
      return false;
    }

    for (let i = 0; i < currentEmojiSpread.length; i++) {
      if (
        currentEmojiSpread[i].emoji !== storedEmojiSpread[i].emoji ||
        currentEmojiSpread[i].row !== storedEmojiSpread[i].row
      ) {
        return false;
      }
    }

    return true;
  },

  /**
   * Reverts the emoji spread in the main sheet to the stored emoji spread.
   * @param {Array<Object>} storedEmojiSpread - The stored emoji spread to restore.
   */
  revertEmojiSpread: function (storedEmojiSpread) {
    const sheet = this.getSheet();

    // Reset the emojis in the sheet based on the stored data
    storedEmojiSpread.forEach((item) => {
      sheet.getRange(item.row, this.activityDataColumn).setValue(item.emoji); // Reset emoji in the proper row
    });
  },

  /**
   * Returns rows that represent a habit, as indicated by an emoji in the activity column.
   * @returns {Array<number>} An array of relevant row numbers.
   */
  getRelevantRows: function () {
    const activitiesChanged = PropertyManager.getProperty(
      PropertyKeys.ACTIVITIES_COLUMN_UPDATED
    );
    const mode = PropertyManager.getProperty(PropertyKeys.MODE);
    let updateRan = false;

    if (activitiesChanged && mode === ModeTypes.CHALLENGE) {
      this.checkEmojiSpread(); // This will revert changes or show a message if necessary
      PropertyManager.setProperty(
        PropertyKeys.ACTIVITIES_COLUMN_UPDATED,
        BooleanTypes.FALSE
      ); // Reset the change flag
      updateRan = true;
    }

    const storedEmojiSpread = JSON.parse(
      PropertyManager.getProperty(PropertyKeys.EMOJI_LIST)
    );
    const relevantRows = storedEmojiSpread.map((entry) => entry.row);
    LoggerManager.logDebug(`getRelevantRows: relevantRows are ${relevantRows}`);

    // Validate the relevant rows
    if (updateRan) {
      SheetConfig._validateRows(relevantRows);
    }

    return relevantRows;
  },

  /**
   * Retrieves a dynamic range based on a column index.
   * @param {number} columnIndex - The column index for which to get the range.
   * @returns {Range} The range object.
   */
  _getDynamicRange: function (columnIndex) {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    const firstRow = this.firstDataInputRow;

    // Get the range from the first row to the last row in the specified column
    return sheet.getRange(firstRow, columnIndex, lastRow - firstRow + 1, 1);
  },

  /**
   * Retrieves the range of completion data.
   * @returns {Range} The range object for completion data.
   */
  getCompletionDataRange: function () {
    try {
      return this._getDynamicRange(this.completionDataColumn);
    } catch (e) {
      LoggerManager.handleError(
        `Failed to get completion data range: ${e.message}`
      );
      return false;
    }
  },

  /**
   * Retrieves the range of buffer data.
   * @returns {Range} The range object for buffer data.
   */
  getBufferDataRange: function () {
    try {
      return this._getDynamicRange(this.bufferDataColumn);
    } catch (e) {
      LoggerManager.handleError(
        `Failed to get buffer data range: ${e.message}`
      );
      return false;
    }
  },

  /**
   * Retrieves the number of relevant rows (habits).
   * @returns {number} The number of relevant rows.
   */
  _getLength: function () {
    const rows = this.getRelevantRows();
    return rows.length;
  },

  /**
   * Returns a list with empty strings for every relevant row in completion data.
   * @returns {Array<string>} An array of arrays with empty strings.
   */
  getDefaultCompletionList: function () {
    const length = this._getLength();
    return Array.from({ length }, () => ""); // Array of arrays with empty strings
  },

  /**
   * Returns a list with the default buffer set for every relevant row in completion data.
   * @returns {Array<string>} An array of arrays with the default buffer.
   */
  getDefaultBufferList: function () {
    const length = this._getLength();
    return Array.from({ length }, () => this.defaultBuffer);
  },

  /**
   * Consolidated getter for main sheet's user-mutable data.
   * @returns {Object} An object containing all main sheet data.
   */
  getMainSheetData: function () {
    return {
      completionData: ActivitiesManager.getCompletionData(),
    };
  },

  /**
   * Retrieves the date value from the date cell.
   * @returns {Date} The date.
   */
  getDate: function () {
    return this.getSheetValue(this.dateCell);
  },

  /**
   * Sets the date value from the date cell.
   * @param {Date} date - The date value to set. Must be a valid Date object.
   */
  setDate: function (date) {
    DateManager._validateDateRange(date);

    this.setSheetValue(
      this.dateCell,
      DateManager.determineFormattedDate(dateStr)
    );
  },

  /**
   * Sets the date in the main sheet.
   * @param {string} dateStr - The date string to set in the sheet.
   */
  setDateStr: function (dateStr) {
    DateManager._validateDateStrRange(dateStr);

    this.setSheetValue(this.dateCell, dateStr);
  },

  /**
   * Sets all relevant data (completion, buffer, streaks) on the main sheet.
   * @param {string} date - The date to set on the main sheet.
   * @param {Array<any>} completionData - The completion data to set.
   * @param {Array<any>} bufferData - The buffer data to set.
   * @param {number} currentStreak - The current streak value to set.
   * @param {number} highestStreak - The highest streak value to set.
   */
  setAllData: function (
    dateStr,
    completionData,
    bufferData,
    currentStreak,
    highestStreak
  ) {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();

    // Validate the data before setting it
    DateManager._validateDateStrRange(dateStr);
    ActivitiesManager._validateCompletionAndBufferData(
      completionData,
      bufferData
    );
    StreakManager._validateCurrentAndHighestStreak(
      currentStreak,
      highestStreak
    );
    const dataRowCount = lastRow - this.firstDataInputRow + 1;

    // Prepare padded arrays to match the total rows in the sheet
    let paddedCompletionData = Array(dataRowCount).fill([""]);
    let paddedBufferData = Array(dataRowCount).fill([""]);
    LoggerManager.logDebug(
      `setAllData: Initialized padded arrays for completion and buffer data.`
    );

    const relevantRows = this.getRelevantRows();
    LoggerManager.logDebug(
      `setAllData: Relevant rows identified: ${relevantRows.join(", ")}.`
    );

    relevantRows.forEach((rowIndex, i) => {
      // Ensure each item is wrapped once in an array
      paddedCompletionData[rowIndex - this.firstDataInputRow] = [
        completionData[i],
      ];
      paddedBufferData[rowIndex - this.firstDataInputRow] = [bufferData[i]];
    });
    LoggerManager.logDebug(
      `setAllData: Relevant rows: ${JSON.stringify(relevantRows)}.`
    );

    // Prepare the data to be set in the sheet
    const rangesAndValues = [
      { range: sheet.getRange(this.dateCell), values: [[dateStr]] },
      { range: this.getCompletionDataRange(), values: paddedCompletionData },
      { range: this.getBufferDataRange(), values: paddedBufferData },
      {
        range: sheet.getRange(this.currentStreakCell),
        values: [[currentStreak]],
      },
      {
        range: sheet.getRange(this.highestStreakCell),
        values: [[highestStreak]],
      },
    ];
    LoggerManager.logDebug(
      `setAllData: Prepared ranges and values for batch operation.`
    );

    rangesAndValues.forEach(({ range, values }) => {
      LoggerManager.logDebug(
        `setAllData: Range A1 Notation: ${range.getA1Notation()}, Values to be set: ${JSON.stringify(
          values
        )}`
      );
    });

    // Attempt to set all data in one batch operation
    try {
      rangesAndValues.forEach(({ range, values }) => range.setValues(values));
      SpreadsheetApp.flush(); // Ensure all changes are applied at once
      LoggerManager.logDebug(`setAllData: All data set successfully.`);
    } catch (e) {
      LoggerManager.handleError(
        `setAllData: Failed to set data: ${e.message}.`
      );
    }
  },

  /**
   * Sets main sheet data to the latest data entry row on the history sheet.
   */
  setLatestEntry: function () {
    const historySheet = HistorySheetConfig.getSheet();

    // Get the last row with data in the history sheet
    const lastRow = historySheet.getLastRow();

    if (lastRow < HistorySheetConfig.firstDataRow) {
      LoggerManager.logDebug(`No data available in the history sheet.`);
      return;
    }

    // Retrieve the data from the last entry in the history sheet
    const lastEntryData = historySheet
      .getRange(lastRow, 1, 1, historySheet.getLastColumn())
      .getValues()[0];

    // Assume history sheet columns map directly to main sheet's expected format
    const date = lastEntryData[HistorySheetConfig.dateColumn];
    const completionData = JSON.parse(
      lastEntryData[HistorySheetConfig.completionDataColumn]
    );
    const bufferData = JSON.parse(
      lastEntryData[HistorySheetConfig.bufferDataColumn]
    );
    const currentStreak = lastEntryData[HistorySheetConfig.currentStreakColumn];
    const highestStreak = lastEntryData[HistorySheetConfig.highestStreakColumn];

    // Set the completion, buffer, current streak, and highest streak data
    this.setAllData(
      DateManager.determineFormattedDate(date),
      completionData,
      bufferData,
      currentStreak,
      highestStreak
    );
    LoggerManager.logDebug(
      `Main sheet data set to the latest entry in the history sheet.`
    );
  },

  /**
   * Sets the checklist data for the provided date on the main sheet.
   * If no relevant data exists for the date, it will create a fresh checklist for that date.
   *
   * @param {string} dateStr - The date for which to set the checklist.
   */
  setChecklistWithDate: function (dateStr) {
    LoggerManager.logDebug(
      `in setChecklistWithDate, with date of: ${dateStr}.`
    );

    const data = HistorySheetConfig.getDataAtDateStr(dateStr);
    if (!data) {
      LoggerManager.logDebug(
        `setChecklistWithDate did not find ${dateStr}. Will have to load the most recent entry instead.`
      );
      this.setLatestEntry();
      return;
    }

    const { date, completionData, bufferData, currentStreak, highestStreak } =
      data;
    MainSheetConfig.setAllData(
      DateManager.determineFormattedDate(date),
      completionData,
      bufferData,
      currentStreak,
      highestStreak
    );
    LoggerManager.logDebug(`Loaded and applied checklist data for ${dateStr}.`);
  },

  /**
   * Displays data for the selected date on the main sheet. If the date is not already in the system, it defaults to today.
   *
   * @param {string} selectedDate - The date string (in YYYY-MM-DD format) to display data for.
   */
  displayDate: function (selectedDateStr) {
    // If the selectedDate is not valid.
    if (!HistorySheetConfig.ensureDateEntry(selectedDateStr)) {
      Messages.showAlert(MessageTypes.INVALID_DATE);
      this.displayDate(DateManager.getTodayStr());
      return;
    }
    this.setChecklistWithDate(selectedDateStr);
  },

  /**
   * Inserts checkboxes into the completion data column for all relevant rows.
   *
   * This method is used to ensure that all cells in the completion data column
   * have checkboxes, which are required for tracking completion status.
   * It retrieves the relevant rows from `MainSheetConfig`, then iterates over
   * each row and applies checkboxes to the specified completion data column.
   */
  insertCompletionCheckboxes: function () {
    const sheet = this.getSheet();
    const relevantRows = MainSheetConfig.getRelevantRows();
    relevantRows.forEach((row) => {
      const range = sheet.getRange(row, MainSheetConfig.completionDataColumn);
      range.insertCheckboxes();
    });
  },

  /**
   * Maintains the original cell value by setting the cell's value back to the old one.
   *
   * @param {Range} range - The range object representing the cell that was edited.
   * @param {any} oldValue - The previous value of the cell before the edit.
   */
  maintainCellValue: function (range, oldValue) {
    const rangeCells = DataHandler.getRangeCells(range);
    LoggerManager.logDebug(
      `maintainCellValue: rangeCells ${rangeCells} edited, with an oldValue of ${oldValue}`
    );
    const isHabitIdeationMode =
      PropertyManager.getProperty(PropertyKeys.MODE) ===
      ModeTypes.HABIT_IDEATION;
    const labelConfigs = [
      { ranges: this.headerLabelRanges, labels: this.headerLabels },
    ];

    // If the mode is habitIdeation, include setter labels
    if (isHabitIdeationMode) {
      labelConfigs.push({
        ranges: this.setterLabelRanges,
        labels: this.setterLabels,
      });
    }

    range.setBackground(SheetConfig.mainColor);

    const rangesAndValues = {};
    // Check if any labels are included in the range, and if so, add to rangesAndValues
    labelConfigs.forEach((config) => {
      Object.keys(config.ranges).forEach((key) => {
        const labelRange = config.ranges[key];
        // Check if labelRange is part of the rangeCells
        if (rangeCells.includes(labelRange)) {
          rangesAndValues[labelRange] = config.labels[key];
        }
      });
    });

    // If any labels were identified in the range, set them to bold font and skip setting oldValue
    if (Object.keys(rangesAndValues).length > 0) {
      for (const [cell, value] of Object.entries(rangesAndValues)) {
        const cellRange = MainSheetConfig.getSheet().getRange(cell);
        LoggerManager.logDebug(`setting cell ${cell} to the value ${value}.`);
        cellRange.setValue(value);
        cellRange.setFontWeight("bold");
        cellRange.setBackground(SheetConfig.secondaryColor);
      }
      LoggerManager.logDebug(`Label(s) have been reset after user changes.`);
    }

    if (
      oldValue === undefined &&
      Object.keys(rangesAndValues).length < rangeCells.length
    ) {
      Messages.showAlert(MessageTypes.UNDEFINED_CELL_CHANGES);
      LoggerManager.logDebug(`User attempted to input many non-label values.`);
    } else if (Object.keys(rangesAndValues).length === 0) {
      range.setValue(oldValue); // Set the old value back to the range
      SpreadsheetApp.flush();
      LoggerManager.logDebug(
        `Cell value reset to previous value: ${oldValue}.`
      );
    }
  },

  /**
   * Checks if the provided range is part of locked ranges (header, streak cells)
   * or locked columns (like buffer data).
   *
   * @param {Range} range - The range object representing the edited cell.
   * @returns {boolean} True if the range is locked, false otherwise.
   */
  includesLockedRange: function (range) {
    const columns = DataHandler.getRangeColumns(range);
    const rangeCells = DataHandler.getRangeCells(range);

    // Combine various locked ranges into one array using spread operator
    const lockedRanges = [
      this.currentStreakCell,
      this.highestStreakCell,
      ...Object.values(this.headerLabelRanges),
    ];
    const lockedColumns = [this.bufferDataColumn];

    // Check for locked property setters during Habit Ideation mode
    if (
      PropertyManager.getProperty(PropertyKeys.MODE) ===
      ModeTypes.HABIT_IDEATION
    ) {
      lockedRanges.push(...Object.values(this.setterLabelRanges));
      lockedColumns.push(
        this.completionDataColumn,
        this.bufferDataColumn,
        this.streaksDataColumn
      );
    }

    // Check if the range is in any of the locked ranges and columns
    const isInLockedColumns = columns.some((column) =>
      lockedColumns.includes(column)
    );
    const isInLockedRange = rangeCells.some((cell) =>
      lockedRanges.includes(cell)
    );

    // Return true if the range is in locked ranges or locked columns
    return isInLockedRange || isInLockedColumns;
  },

  /**
   * Resets the habit-related data including streaks and challenge date.
   *
   * Clears the completion and buffer columns, resets streak values, and updates the selected date to today.
   */
  resetData: function () {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();

    // Define the columns for completion and buffer
    const firstColumn = this.completionDataColumn;
    const lastColumn = this.bufferDataColumn;
    const firstDataInputRow = this.firstDataInputRow;
    const numberOfRows = lastRow - firstDataInputRow + 1;
    const numberOfColumns = lastColumn - firstColumn + 1;

    // Clear validations for the completion column
    const completionRange = sheet.getRange(
      firstDataInputRow,
      this.completionDataColumn,
      numberOfRows
    );
    completionRange.clearDataValidations();

    const fullRange = sheet.getRange(
      firstDataInputRow,
      firstColumn,
      numberOfRows,
      numberOfColumns
    );
    fullRange.clearContent();

    // Now batch the setting of new values
    const todayStr = DateManager.getTodayStr();
    const currentStreakCell = this.currentStreakCell;
    const highestStreakCell = this.highestStreakCell;
    const dateCell = this.dateCell;
    const defaultStreak = 0;

    const rangesAndValues = [
      { range: sheet.getRange(dateCell), values: [[todayStr]] },
      { range: sheet.getRange(currentStreakCell), values: [[defaultStreak]] },
      { range: sheet.getRange(highestStreakCell), values: [[defaultStreak]] },
    ];

    rangesAndValues.forEach(({ range, values }) => {
      range.setValues(values);
    });

    // Flush all changes at once
    SpreadsheetApp.flush();
    LoggerManager.logDebug(`Challenge data reset.`);
  },

  /**
   * Toggles the visibility of specific columns (completion, buffer, and streaks) based on the action.
   *
   * @param {string} action - 'show' to display the columns, 'hide' to hide the columns.
   */
  toggleColumns: function (action) {
    const sheet = MainSheetConfig.getSheet();

    // Define the first and last column in the range
    const streaksDataColumn = MainSheetConfig.streaksDataColumn;
    const firstColumn = MainSheetConfig.completionDataColumn;
    const lastColumn = MainSheetConfig.bufferDataColumn;
    const numberOfColumns = lastColumn - firstColumn + 1; // Calculate the number of columns

    if (action === ColumnAction.SHOW) {
      sheet.showColumns(firstColumn, numberOfColumns); // Show all columns at once
      sheet.showColumns(streaksDataColumn - 1, 2);
    } else if (action === ColumnAction.HIDE) {
      sheet.hideColumns(firstColumn, numberOfColumns); // Hide all columns at once
      sheet.hideColumns(streaksDataColumn - 1, 2);
    } else {
      LoggerManager.handleError(
        `toggleColumns called with unknown action: ${action}.`
      );
    }
  },
};

/**
 * HistorySheetConfig manages the configuration and operations for the "history" sheet.
 * This includes methods to retrieve, update, and validate data related to completion, buffer, and streaks.
 *
 * This object follows the singleton pattern, ensuring a single instance.
 */
const HistorySheetConfig = {
  /**
   * @type {string} sheetName - The name of the history sheet.
   */
  sheetName: "history",

  /**
   * @type {number} dateColumn - The 0-indexed column number for the date in the history sheet.
   */
  dateColumn: 0,

  /**
   * @type {number} completionDataColumn - The 0-indexed column number for the completion data in the history sheet.
   */
  completionDataColumn: 1,

  /**
   * @type {number} bufferDataColumn - The 0-indexed column number for the buffer data in the history sheet.
   */
  bufferDataColumn: 2,

  /**
   * @type {number} currentStreakColumn - The 0-indexed column number for the current streak in the history sheet.
   */
  currentStreakColumn: 3,

  /**
   * @type {number} highestStreakColumn - The 0-indexed column number for the highest streak in the history sheet.
   */
  highestStreakColumn: 4,

  /**
   * @type {number} firstDataRow - The first data row (0-indexed) in the history sheet, used to calculate data lengths.
   */
  firstDataRow: 1, // this is a static feature, across future iterations as well.

  /**
   * @constant {number} boostIntervalDefault - The default boost interval in days.
   */
  boostIntervalDefault: 7,

  /**
   * Retrieves the sheet object for the history sheet.
   * @returns {Sheet} The sheet object for the history sheet.
   */
  getSheet: function () {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
  },

  /**
   * Retrieves the total number of rows in the history sheet, including headers.
   * @returns {number} The total number of rows in the history sheet.
   */
  getTotalLength: function () {
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    return data.length;
  },

  /**
   * Validates if a row index is a non-negative number and within the valid range of the history sheet.
   * @param {number} row - The 1-indexed row index to validate.
   * @param {boolean} [throwError=true] - If true, throws an error when validation fails.
   * @returns {boolean} - Returns true if the row index is valid, false otherwise.
   */
  _validateRowIndex: function (row, throwError = true) {
    // Check if the row is a non-negative number
    if (!UtilsManager.__validateNonNegativeNumber(row)) {
      LoggerManager.handleError(
        `Invalid input: row must be a non-negative number.`,
        throwError
      );
      return false;
    }

    const firstRow =
      PropertyManager.getPropertyNumber(PropertyKeys.FIRST_CHALLENGE_ROW) + 1;
    LoggerManager.logDebug(`_validateRowIndex: firstRow is ${firstRow}`);

    // Check if the row is within the valid range of the sheet
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    if (row < firstRow || row > lastRow) {
      LoggerManager.handleError(
        `_validateRowIndex: Row index ${row} is out of range. Valid rows are between ${firstRow} and ${lastRow}.`,
        throwError
      );
      return false;
    }

    return true;
  },

  /**
   * Validates that the row array matches the number of columns between completion and highest streak columns.
   *
   * @param {Array} row - The array representing the row to be appended.
   * @returns {boolean} - True if the row length matches the expected number of columns, false otherwise.
   */
  _validateEntryRow: function (row, throwError) {
    // Calculate the number of columns from completion to highest streak
    const numColumns = this.highestStreakColumn - this.dateColumn + 1;

    // Validate the length of the row
    if (row.length !== numColumns) {
      LoggerManager.handleError(
        `Row length (${row.length}) does not match the expected number of columns (${numColumns}).`,
        throwError
      );
      return false;
    }

    return true;
  },

  /**
   * Calculates the row index for a specific date.
   * @param {string} dateStr - The date string (in YYYY-MM-DD format).
   * @returns {number|null} - The 1-indexed row index for the given date, or null if out of range.
   */
  getRowIndexForDate: function (dateStr) {
    const sheet = this.getSheet();

    if (!DateManager._validateDateStrRange(dateStr, false)) {
      return null;
    }

    // Get the last date and its corresponding row, and ensure that they exist.
    const lastRow = sheet.getLastRow();
    const lastDate = this.getLastDate();
    const targetDate = DateManager.determineDate(dateStr);
    if (!lastDate || !lastRow || targetDate > lastDate) {
      // Filters out valid, yet currently unentered dates (i.e. first today entry)
      LoggerManager.logDebug(
        `getRowIndexForDate: No valid/appropriate last date or last row found.`
      );
      return null;
    }

    // Calculate the difference in days between the last date and the target date
    const msPerDay = 1000 * 60 * 60 * 24;
    let daysDifference = Math.round((lastDate - targetDate) / msPerDay);

    // Calculate the target row based on the last row and the difference in days
    const rowIndex = lastRow - daysDifference;

    LoggerManager.logDebug(
      `getRowIndexForDate: Calculated rowIndex: ${rowIndex}, Passed dateStr and its targetDate: ${dateStr} and ${targetDate}, lastDate: ${lastDate}, daysDifference: ${daysDifference}`
    ); //, Actual date string: ${actualDateStr}`);
    return rowIndex;
  },

  /**
   * Retrieves data from a specific row in the history sheet.
   * @param {number} row - The 1-indexed row number to retrieve data from.
   * @returns {Object} An object containing the date, completionData, bufferData, currentStreak, and highestStreak.
   */
  getDataAtRow: function (row) {
    const sheet = this.getSheet();
    const data = sheet
      .getRange(row, this.dateColumn + 1, 1, this.highestStreakColumn + 1)
      .getValues()[0];

    const date = data[this.dateColumn]
      ? data[this.dateColumn]
      : (() => {
          LoggerManager.logDebug(`Invalid or missing date for row: ${row}.`);
          return null;
        })();

    let completionData, bufferData;

    try {
      completionData = JSON.parse(data[this.completionDataColumn]);
    } catch (e) {
      LoggerManager.handleError(`Error parsing completionData: ${e.message}.`);
      completionData = [];
    }

    try {
      bufferData = JSON.parse(data[this.bufferDataColumn]);
    } catch (e) {
      LoggerManager.handleError(`Error parsing bufferData: ${e.message}.`);
      bufferData = [];
    }

    const currentStreak = data[this.currentStreakColumn];
    const highestStreak = data[this.highestStreakColumn];

    LoggerManager.logDebug(
      `getDataAtRow at ${row}: returning date ${date}, completionData ${completionData}, bufferData ${bufferData}, currentStreak ${currentStreak}, highestStreak ${highestStreak}.`
    );

    return {
      date: date,
      completionData: completionData,
      bufferData: bufferData,
      currentStreak: currentStreak,
      highestStreak: highestStreak,
    };
  },

  /**
   * Saves the current state of the main sheet data to the history sheet for a specific date. This date should have already been loaded into the history sheet.
   * @param {string} dateStr - The date string (in YYYY-MM-DD format) to save data for.
   */
  saveCurrentState: function (dateStr) {
    const { completionData } = MainSheetConfig.getMainSheetData();

    // Save the current state. then, propagate
    LoggerManager.logDebug(
      `Trying saveCurrentState with: completionData = ${JSON.stringify(
        completionData
      )}`
    );
    this.setCompletionDataAtDate(dateStr, completionData); // automatically propagates
  },

  /**
   * Sets the completion data for a specified date into the history sheet.
   * Only changes the completion data cell for the given date.
   *
   * @param {string} dateStr - The date string (in YYYY-MM-DD format) to save completion data for.
   * @param {Array<any>} completionData - A 1-D array where each inner array represents a row of completion data.
   * @param {boolean} [propagate=true] - Optional. If true, propagates changes to subsequent rows.
   */
  setCompletionDataAtDate: function (
    dateStr,
    completionData,
    propagate = true
  ) {
    // Calculate the row index for the given date
    const rowIndex = this.getRowIndexForDate(dateStr);

    if (rowIndex === null) {
      LoggerManager.handleError(`No valid row found for date: ${dateStr}`);
      return;
    }

    const historySheet = this.getSheet();
    const completionCell = historySheet.getRange(
      rowIndex,
      this.completionDataColumn + 1
    ); // 1-indexed range
    completionCell.setValue(JSON.stringify(completionData));
    SpreadsheetApp.flush();

    LoggerManager.logDebug(
      `setCompletionDataAtDate: Completion data set for date: ${dateStr}. Completion Data: ${JSON.stringify(
        completionData
      )}`
    );

    // If propagation is enabled, propagate from the previous day
    if (propagate) {
      const previousDateStr = DateManager.getPreviousDateStr(dateStr); // Assume you have this method
      if (DateManager._validateDateStrRange(previousDateStr, false)) {
        this.propagateChanges(previousDateStr); // Propagate starting from the previous day
      } else {
        LoggerManager.logDebug(
          `No previous date to propagate from. Propagating from changed day instead.`
        );
        this.propagateChanges(dateStr);
      }
    }
  },

  /**
   * Sets a default entry in the history sheet for the specified date if no entry exists.
   * If an entry already exists, no changes are made.
   *
   * @param {string} dateStr - The date string (in YYYY-MM-DD format) to check or set as the default entry.
   * @returns {boolean} - Returns true if a default entry is created, or false if an entry already exists for the date.
   */
  setDefaultEntry: function (dateStr) {
    DateManager._validateDateStrRange(dateStr);

    // Check if the date already has an entry in the history sheet
    if (this.getRowIndexForDate(dateStr) !== null) {
      LoggerManager.logDebug(
        `setDefaultEntry: Date should already be loaded into the system: ${dateStr}.`
      );
      return false; // Entry already exists, no need to create a default entry
    }

    LoggerManager.logDebug(
      `setDefaultEntry: No entry found for ${dateStr}. Creating default entry.`
    );

    // If no entry exists, create a new one with default values
    const historySheet = this.getSheet(); // Ensure you get the history sheet first
    const defaultCompletionList = MainSheetConfig.getDefaultCompletionList(); // Default empty list for completion and buffer data
    const defaultBufferList = MainSheetConfig.getDefaultBufferList();
    const defaultStreak = 0; // Default streak value

    // Set the default entry for the provided date
    const entryRow = [
      dateStr,
      JSON.stringify(defaultCompletionList),
      JSON.stringify(defaultBufferList),
      defaultStreak,
      defaultStreak,
    ];
    this._validateEntryRow(entryRow);
    historySheet.appendRow(entryRow);

    LoggerManager.logDebug(
      `setDefaultEntry: Default entry set for ${dateStr}.`
    );
    return true; // Entry has been created
  },

  /**
   * Determines whether the buffer should increase.
   * This function checks if today marks a multiple of the boostInterval
   * since the first challenge date.
   *
   * @param {number} currentStreak - The current streak to consider how long the current iteration of the challenge has been going.
   * @returns {boolean} True if the buffer should increase, false otherwise.
   */
  shouldBufferIncrease: function (currentStreak) {
    const boostInterval = PropertyManager.getProperty(
      PropertyKeys.BOOST_INTERVAL
    );
    LoggerManager.logDebug(
      `shouldBufferIncrease: Current Streak: ${currentStreak}, Days to increment buffer: ${boostInterval}`
    );

    // Check if the row is a multiple of boostInterval relative to the currentStreak
    return currentStreak > 0 && currentStreak % boostInterval === 0;
  },

  /**
   * Propagates changes from the specified date row down to the bottommost row in the history sheet.
   * This function is primarily used to ensure that changes in one day's data affect subsequent days correctly.
   *
   * @param {string} changedDateStr - The date string for which changes have been made and need to be propagated.
   */
  propagateChanges: function (changedDateStr) {
    const todayStr = DateManager.getTodayStr();
    const lastDateStr = this.getLastDateStr();
    LoggerManager.logDebug(
      `In propagateChanges with a changedDateStr of ${changedDateStr}, whilst today's date is ${todayStr} and the last date is ${lastDateStr}`
    );

    // If the start row doesn't exist, throw an error.
    let startRowIndex = this.getRowIndexForDate(changedDateStr);
    if (startRowIndex === null || changedDateStr == lastDateStr) {
      LoggerManager.handleError(
        `No valid row needing propagation found for the changed date: ${changedDateStr}`,
        false
      );
      return;
    }

    startRowIndex -= 1; // make it 0-indexed
    const data = this.getDataRange();

    // Initialize arrays for buffer, current streaks, and highest streaks batch updates
    const bufferDataBatch = [];
    const currentStreakBatch = [];
    const highestStreakBatch = [];

    // Read the data for the starting row (i.e., the changedDate) ahead of the loop
    let previousData = {
      completionData: JSON.parse(
        data[startRowIndex][this.completionDataColumn]
      ),
      bufferData: JSON.parse(data[startRowIndex][this.bufferDataColumn]),
      currentStreak: data[startRowIndex][this.currentStreakColumn],
      highestStreak: data[startRowIndex][this.highestStreakColumn],
      challengeFailed: this.hasChallengeFailed(
        JSON.parse(data[startRowIndex][this.completionDataColumn]),
        JSON.parse(data[startRowIndex][this.bufferDataColumn])
      ),
    };
    LoggerManager.logDebug(
      `Data to propagate from date of ${changedDateStr}: ${JSON.stringify(
        previousData
      )}.`
    );

    // Start propagation from the determined row
    for (let i = startRowIndex + 1; i < data.length; i++) {
      // Use the saved values from the previous row to update the current row's data
      let currentData = {
        completionData: JSON.parse(data[i][this.completionDataColumn]),
        bufferData: JSON.parse(data[i][this.bufferDataColumn]),
        currentStreak: previousData.currentStreak + 1,
        highestStreak: Math.max(
          previousData.currentStreak + 1,
          previousData.highestStreak
        ),
        challengeFailed: false,
      };
      LoggerManager.logDebug(
        `Current data for row ${i + 1}: ${JSON.stringify(currentData)}.`
      );

      // If the previous day's challenge failed
      if (previousData.challengeFailed) {
        LoggerManager.logDebug(
          `Previous day's challenge failed, resetting streak.`
        );
        currentData.currentStreak = 0;
        currentData.highestStreak = previousData.highestStreak;
        currentData.bufferData = MainSheetConfig.getDefaultBufferList();
      } else {
        LoggerManager.logDebug(`Evaluating buffer for row ${i + 1}.`);
        const bufferIncrease = this.shouldBufferIncrease(
          currentData.currentStreak
        );

        currentData.bufferData = previousData.bufferData.map((buffer, j) => {
          if (bufferIncrease) {
            buffer += 1;
          }
          if (!previousData.completionData[j]) {
            buffer -= 1;
            if (buffer < 1 && !currentData.completionData[j]) {
              currentData.challengeFailed = true;
            }
          }
          return buffer;
        });
      }

      // Accumulate updates for batch processing
      bufferDataBatch.push([JSON.stringify(currentData.bufferData)]);
      currentStreakBatch.push([currentData.currentStreak]);
      highestStreakBatch.push([currentData.highestStreak]);
      LoggerManager.logDebug(
        `Final data for row ${i + 1}: ${JSON.stringify(currentData)}.`
      );

      // Save the newly updated data to be used in the next iteration
      previousData = currentData;
    }

    // Perform separate batch updates for flexibility.
    const sheet = this.getSheet();
    sheet
      .getRange(
        startRowIndex + 2,
        this.bufferDataColumn + 1,
        bufferDataBatch.length
      )
      .setValues(bufferDataBatch);
    sheet
      .getRange(
        startRowIndex + 2,
        this.currentStreakColumn + 1,
        currentStreakBatch.length
      )
      .setValues(currentStreakBatch);
    sheet
      .getRange(
        startRowIndex + 2,
        this.highestStreakColumn + 1,
        highestStreakBatch.length
      )
      .setValues(highestStreakBatch);
    SpreadsheetApp.flush();

    LoggerManager.logDebug("Finished propagating...");
  },

  /**
   * Retrieves the entire data range from the history sheet.
   * @returns {Array<Array<any>>} The data range as a 2D array.
   */
  getDataRange: function () {
    return this.getSheet().getDataRange().getValues();
  },

  /**
   * Sets the completion data for a specific row in the history sheet.
   * @param {number} row - The 0-indexed row number to set completion data for.
   * @param {Array} data - The completion data to be set (array of arrays).
   */
  setCompletionDataForRow: function (row, data) {
    this._validateRowIndex(row);
    try {
      this.getSheet()
        .getRange(row, this.completionDataColumn)
        .setValue(JSON.stringify(data));
      SpreadsheetApp.flush();
      LoggerManager.logDebug(
        `Completion data set successfully for row: ${row}.`
      );
    } catch (e) {
      LoggerManager.handleError(
        `Failed to set completion data for row ${row}: ${e.message}.`
      );
    }
  },

  /**
   * Sets the current streak value for a specific row in the history sheet.
   * @param {number} row - The 0-indexed row number to set the current streak for.
   * @param {number} value - The current streak value to be set.
   */
  setCurrentStreakForRow: function (row, value) {
    this._validateRowIndex(row);
    try {
      this.getSheet().getRange(row, this.currentStreakColumn).setValue(value);
      SpreadsheetApp.flush();
      LoggerManager.logDebug(
        `Current streak set successfully for row: ${row}.`
      );
    } catch (e) {
      LoggerManager.handleError(
        `Failed to set current streak for ${row}: ${e.message}.`
      );
    }
  },

  /**
   * Sets the highest streak value for a specific row in the history sheet.
   * @param {number} row - The 0-indexed row number to set the highest streak for.
   * @param {number} value - The highest streak value to be set.
   */
  setHighestStreakForRow: function (row, value) {
    this._validateRowIndex(row);
    try {
      this.getSheet().getRange(row, this.highestStreakColumn).setValue(value);
      SpreadsheetApp.flush();
      LoggerManager.logDebug(
        `Highest streak set successfully for row: ${row}.`
      );
    } catch (e) {
      LoggerManager.handleError(
        `Failed to set highest streak for row ${row}: ${e.message}.`
      );
    }
  },

  /**
   * Determines if the challenge failed based on the provided completion and buffer data (using data from the same day).
   *
   * @param {Array<any>} completionData - The completion data to check.
   * @param {Array<any>} bufferData - The buffer data to check.
   * @returns {boolean} - Returns true if the challenge failed, otherwise false.
   */
  hasChallengeFailed: function (completionData, bufferData) {
    LoggerManager.logDebug(
      `Checking for challenge failures on completion: ${JSON.stringify(
        completionData
      )} and buffer: ${JSON.stringify(bufferData)}.`
    );

    for (let i = 0; i < completionData.length; i++) {
      // If a completion box is unchecked and there is a corresponding 0 buffer, the challenge has failed
      if (
        (completionData[i] === false || completionData[i] === "") &&
        bufferData[i] < 1
      ) {
        LoggerManager.logDebug(
          `Challenge failed on Activity ${i + 1} with bufferData[i] of ${
            bufferData[i]
          }.`
        );
        return true;
      }
    }

    LoggerManager.logDebug(`No challenge failures detected.`);
    return false;
  },

  /**
   * Retrieves the first valid date from the properties.
   *
   * This function retrieves the earliest recorded start date from the stored properties.
   *
   * @returns {Date|null} - The first valid date as a Date object, or null if no valid date is found.
   */
  getFirstDate: function () {
    const firstDateStr = this.getFirstDateStr();
    return firstDateStr ? DateManager.determineDate(firstDateStr) : null; // Return formatted string from the Date object
  },

  /**
   * Retrieves the first valid date as a formatted string from the stored properties.
   *
   * @returns {string|null} The formatted date string or null if no date is found.
   */
  getFirstDateStr: function () {
    const firstDateStr = PropertyManager.getProperty(
      PropertyKeys.FIRST_CHALLENGE_DATE
    );

    if (firstDateStr) {
      LoggerManager.logDebug(
        `getFirstDateStr: Retrieved stored date: ${firstDateStr}`
      );
      return firstDateStr;
    }

    LoggerManager.logDebug(
      `getFirstDateStr: No valid start date found in properties.`
    );
    return null; // Return null if no valid date is found
  },

  /**
   * Retrieves the last valid date from the history sheet.
   * Assumes dates are stored in the `dateColumn`.
   *
   * @returns {string|null} The last valid date string found on the history sheet, or null if no valid date is found.
   */
  getLastDateStr: function () {
    const date = this.getLastDate();
    if (date == null) {
      return null;
    }
    return DateManager.determineFormattedDate(date); // Return null if no valid date is found
  },

  /**
   * Retrieves the last valid date from the history sheet as a Date object.
   *
   * This function first retrieves the latest recorded date string in the history sheet.
   * If a valid date string is found, it converts the string into a Date object and returns it.
   *
   * If no valid date is found, it returns null.
   *
   * @returns {Date|null} - The last valid date as a Date object, or null if no valid date is found.
   */
  getLastDate: function () {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow(); // Get the last row with data
    const dateColumn = this.dateColumn + 1; // Adjust for 1-based indexing

    // Get the date from the last row in the date column
    const lastDate = sheet.getRange(lastRow, dateColumn).getValue();

    if (UtilsManager.__validateDate(lastDate)) {
      // Validate if the date is valid
      LoggerManager.logDebug(
        `getLastDate: Valid date found at row ${lastRow}: ${lastDate}`
      );
      return lastDate; // Convert to Date object using DateManager
    }

    LoggerManager.logDebug(
      `getLastDate: No valid date found in the last row. Last date found is ${lastDate}`
    );
    return null; // No valid date found
  },

  /**
   * Retrieves data from a specific date in the history sheet.
   * @param {string} dateStr - The date string (in YYYY-MM-DD format) to retrieve data for.
   * @returns {Object|null} An object containing the date, completionData, bufferData, currentStreak, and highestStreak, or null if no data is found for the given date.
   */
  getDataAtDateStr: function (dateStr) {
    const rowIndex = this.getRowIndexForDate(dateStr);

    if (rowIndex === null) {
      LoggerManager.logDebug(
        `getDataAtDateStr: No data found for date: ${dateStr}.`
      );
      return null;
    }

    // Retrieve the data from the row
    return this.getDataAtRow(rowIndex);
  },

  /**
   * Ensures that data for a given date is available in the main sheet.
   * If data for the specified date is not found, it attempts to load data
   * from the previous day. If the previous day's data is not available either,
   * it initializes the main sheet with default values.
   *
   * @param {string} dateStr - The date string (in YYYY-MM-DD format) for which data needs to be ensured.
   * @returns {boolean} - Returns true if a valid date was requested (whether or not it already is in the database), otherwise false.
   */
  ensureDateEntry: function (dateStr) {
    const dateValid = DateManager._validateDateStrRange(dateStr, false);
    const todayStr = DateManager.getTodayStr();

    if (!dateValid) {
      // Date is not valid, so default to today.
      LoggerManager.logDebug(
        `ensureDateEntry: invalid Date. Will updated dateStr to todayStr: ${todayStr}`
      );
      return false;
    }

    LoggerManager.logDebug(
      `in ensureDateEntry as of ${todayStr}, with a dateStr of: ${dateStr}, and a valid date.`
    );

    // Set to default values.
    if (!this.setDefaultEntry(dateStr)) {
      return true;
    }

    // Attain a properly formatted date string for yesterday, and check its validity.
    const yesterdayStr = DateManager.getPreviousDateStr(todayStr);
    const yesterdayDataValid = DateManager._validateDateStrRange(
      yesterdayStr,
      false
    );

    if (yesterdayDataValid) {
      // Save and propagate the new data for today
      LoggerManager.logDebug(
        `Will propagate changes using data found for yesterday.`
      );
      this.propagateChanges(yesterdayStr);
    }

    return true;
  },
};

// Freeze the configuration objects to prevent modification
Object.seal(SheetConfig);
Object.seal(MainSheetConfig);
Object.seal(HistorySheetConfig);
