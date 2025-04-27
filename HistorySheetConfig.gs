/**
 * @fileoverview Configuration and utility methods specific to the 'history' sheet.
 * Contains static properties for sheet layout (name, columns) and methods
 * for basic sheet access. Data manipulation logic resides in DataHandler.
 */

/** OnlyCurrentDoc */

/**
 * HistorySheetConfig manages the configuration for the "history" sheet.
 * It acts as a singleton, holding static configuration data.
 * @namespace HistorySheetConfig
 */
const HistorySheetConfig = {
  /** @constant {string} sheetName - The exact name of the history sheet. */
  sheetName: "history",

  // --- Column Indices (0-based) ---
  // These define the structure of data arrays read from/written to the history sheet.
  /** @constant {number} dateColumn - 0-based index for the date column (expects YYYY-MM-DD string). */
  dateColumn: 0,
  /** @constant {number} completionDataColumn - 0-based index for completion data (expects JSON string array). */
  completionDataColumn: 1,
  /** @constant {number} bufferDataColumn - 0-based index for buffer data (expects JSON string array). */
  bufferDataColumn: 2,
  /** @constant {number} currentStreakColumn - 0-based index for the current streak number. */
  currentStreakColumn: 3,
  /** @constant {number} highestStreakColumn - 0-based index for the highest streak number. */
  highestStreakColumn: 4,

  // --- Row Indices ---
  /**
   * @constant {number} firstDataRow - The 0-based row index where history data *starts* (below the header row).
   * Assumes header is in row 1 (index 0), so data begins in row 2 (index 1).
   */
  firstDataRow: 1,

  // --- Default Values ---
  /** @constant {number} boostIntervalDefault - The default interval (days) for buffer boosts if not set by the user. */
  boostIntervalDefault: 7,

  // --- Private Cached Sheet Object ---
  /**
   * Cached Sheet object to avoid repeated calls to getSheetByName.
   * @private
   * @type {GoogleAppsScript.Spreadsheet.Sheet | null}
   */
  _sheet: null,

  // --- Basic Sheet Accessors ---

  /**
   * Retrieves the Sheet object for the relevant sheet (main or history) by iterating
   * through all sheets. This method was chosen for reliability over getSheetByName,
   * which exhibited issues in certain contexts after sheet copying.
   *
   * @private
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
   * @throws {Error} if the sheet cannot be found.
   */
  _getSheet: function () {
    const sheetNameToFind = this.sheetName; // e.g., "main" or "history"
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();

    // Critical check for active spreadsheet context
    if (!activeSS) {
      // If this happens, something is fundamentally wrong with the script environment
      LoggerManager.handleError(
        `FATAL: Cannot get Active Spreadsheet context in _getSheet for ${sheetNameToFind}.`,
        true
      );
      return null; // Should be unreachable due to throw
    }
    const activeSSName = activeSS.getName();
    LoggerManager.logDebug(
      `_getSheet: Locating sheet "${sheetNameToFind}" in SS: "${activeSSName}" via iteration.`
    );

    // Iterate through sheets to find the one with the matching name
    const allSheets = activeSS.getSheets();
    let foundSheet = null;
    for (let i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName() === sheetNameToFind) {
        foundSheet = allSheets[i];
        break; // Stop searching once found
      }
    }

    // Handle case where sheet is not found even after iterating
    if (!foundSheet) {
      const allSheetNames = allSheets.map((s) => `"${s.getName()}"`); // Get names for error message
      LoggerManager.handleError(
        `Sheet "${sheetNameToFind}" was not found via ITERATION in Spreadsheet "${activeSSName}". Available sheets: [${allSheetNames.join(
          ", "
        )}]. Application cannot function.`,
        true
      );
      return null; // Error thrown, unreachable
    }

    LoggerManager.logDebug(
      `_getSheet: Successfully found sheet "${sheetNameToFind}" by iterating.`
    );
    // Skip caching when using iteration; return the found sheet directly.
    // this._sheet = foundSheet; // Caching bypassed
    return foundSheet;
  },

  /**
   * Retrieves the full range containing data (including headers) from the history sheet.
   * Returns null if the sheet is not found or on error.
   * @private
   * @returns {GoogleAppsScript.Spreadsheet.Range | null} The full data range or null.
   */
  _getFullDataRange: function () {
    try {
      const sheet = this._getSheet();
      // getDataRange() correctly handles empty sheets (returns null or A1 range)
      // and sheets with only headers.
      return sheet.getDataRange();
    } catch (e) {
      LoggerManager.handleError(
        `Failed to get data range from history sheet: ${e.message}`,
        false
      );
      return null;
    }
  },

  /**
   * Appends a row of data to the history sheet after validating its structure.
   * @private
   * @param {Array<*>} rowData - An array containing the data for one row, matching the column structure
   *                              (e.g., [dateStr, completionJson, bufferJson, currentStreak, highestStreak]).
   * @returns {boolean} True if successful, false otherwise (e.g., validation failure, append error).
   */
  _appendRow: function (rowData) {
    // Validate the structure before appending
    if (!ValidationUtils._validateHistoryEntryRow(rowData, true)) {
      // Log error if invalid
      // Error already logged by validation function
      return false;
    }

    try {
      const sheet = this._getSheet();
      sheet.appendRow(rowData);
      // Avoid flush here; let higher-level operations manage flushing.
      LoggerManager.logDebug(
        `Successfully appended row to history sheet: ${rowData[0]}`
      ); // Log date part
      return true;
    } catch (e) {
      LoggerManager.handleError(
        `Failed to append row to history sheet for date ${rowData[0]}: ${e.message}`,
        false
      );
      return false;
    }
  },

  /**
   * Gets the number of columns expected in a history data row, based on configuration indices.
   * @private
   * @returns {number} The number of columns (e.g., 5).
   */
  _getNumberOfColumns: function () {
    // Calculate based on 0-based indices
    return this.highestStreakColumn - this.dateColumn + 1;
  },
};

// Freeze the configuration object to prevent modification at runtime.
Object.freeze(HistorySheetConfig);
