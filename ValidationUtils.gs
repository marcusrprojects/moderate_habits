/**
 * @fileoverview Provides utility functions for validating various data types and structures
 * used within the application. These methods are designed primarily for internal checks.
 */

/** OnlyCurrentDoc */

/**
 * Collection of validation utility methods.
 * Methods prefixed with "_" are generally intended for internal use within modules,
 * primarily returning boolean flags. Errors might be logged via LoggerManager.
 * @namespace ValidationUtils
 */
const ValidationUtils = {
  /**
   * Validates if the input is an array.
   * @private
   * @param {*} data - The data to validate.
   * @returns {boolean} True if `data` is an array, false otherwise.
   */
  _validateArray: function (data) {
    return Array.isArray(data);
  },

  /**
   * Validates an array of row numbers. Checks if it's a non-empty array.
   * @private
   * @param {Array<number>} rows - The array of row numbers to validate.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if `rows` is a non-empty array, false otherwise.
   */
  _validateRowsArray: function (rows, logError = true) {
    if (!this._validateArray(rows)) {
      if (logError)
        LoggerManager.handleError(`Invalid rows: Must be an array`, false);
      return false;
    }
    if (rows.length === 0) {
      if (logError)
        LoggerManager.handleError(
          `Invalid rows: Must be a non-empty array`,
          false
        );
      return false;
    }
    // TODO: Optionally add validation to ensure all elements are numbers > 0
    return true;
  },

  /**
   * Validates if the input is a valid Date object (and not NaN).
   * @private
   * @param {*} date - The data to validate.
   * @returns {boolean} True if `date` is a valid Date object.
   */
  _validateDate: function (date) {
    return date instanceof Date && !isNaN(date.getTime());
  },

  /**
   * Validates if the input is a non-negative number (integer or float, not NaN).
   * @private
   * @param {*} data - The data to validate.
   * @returns {boolean} True if `data` is a number >= 0.
   */
  _validateNonNegativeNumber: function (data) {
    return typeof data === "number" && !isNaN(data) && data >= 0;
  },

  /**
   * Validates if the input is a positive number (integer or float, > 0, not NaN).
   * @private
   * @param {*} data - The data to validate.
   * @returns {boolean} True if `data` is a number > 0.
   */
  _validatePositiveNumber: function (data) {
    return typeof data === "number" && !isNaN(data) && data > 0;
  },

  /**
   * Validates if the input is a non-negative integer.
   * @private
   * @param {*} data - The data to validate.
   * @returns {boolean} True if `data` is an integer >= 0.
   */
  _validateNonNegativeInteger: function (data) {
    return this._validateNonNegativeNumber(data) && Number.isInteger(data);
  },

  /**
   * Validates if the input is a positive integer (> 0).
   * @private
   * @param {*} data - The data to validate.
   * @returns {boolean} True if `data` is an integer > 0.
   */
  _validatePositiveInteger: function (data) {
    return this._validatePositiveNumber(data) && Number.isInteger(data);
  },

  /**
   * Validates the completion data array. Primarily checks if it's an array.
   * @private
   * @param {Array<*>} completionData - The completion data array to validate.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if `completionData` is an array.
   */
  _validateCompletionData: function (completionData, logError = true) {
    if (!this._validateArray(completionData)) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: completion data must be an array.`,
          false
        );
      return false;
    }
    // TODO: Could add checks for content type (e.g., boolean, string) if needed.
    return true;
  },

  /**
   * Validates the buffer data array. Checks if it's an array containing only valid numbers.
   * @private
   * @param {Array<*>} bufferData - The buffer data array to validate.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if `bufferData` is an array of numbers.
   */
  _validateBufferData: function (bufferData, logError = true) {
    if (!this._validateArray(bufferData)) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: buffer data must be an array.`,
          false
        );
      return false;
    }
    const allNumbers = bufferData.every(
      (item) => typeof item === "number" && !isNaN(item)
    );
    if (!allNumbers) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: buffer data ${JSON.stringify(
            bufferData
          )} must contain only valid numbers.`,
          false
        );
      return false;
    }
    return true;
  },

  /**
   * Validates that completion data and buffer data are arrays of the same length.
   * Also performs individual validation on each array structure and type.
   * @private
   * @param {Array<*>} completionData - The completion data array.
   * @param {Array<*>} bufferData - The buffer data array.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if both are valid arrays of the same length.
   */
  _validateCompletionAndBufferData: function (
    completionData,
    bufferData,
    logError = true
  ) {
    const isCompletionValid = this._validateCompletionData(
      completionData,
      logError
    );
    const isBufferValid = this._validateBufferData(bufferData, logError);

    if (!isCompletionValid || !isBufferValid) {
      return false; // Errors already logged by individual validation if logError=true
    }

    if (completionData.length !== bufferData.length) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: completionData and bufferData must be of the same length. Lengths are ${completionData.length} and ${bufferData.length}.`,
          false
        );
      return false;
    }
    return true;
  },

  /**
   * Validates the current streak value (must be a non-negative number).
   * @private
   * @param {number} currentStreak - The current streak value to validate.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if `currentStreak` is valid.
   */
  _validateCurrentStreak: function (currentStreak, logError = true) {
    if (!this._validateNonNegativeNumber(currentStreak)) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: currentStreak (${currentStreak}) must be a non-negative number.`,
          false
        );
      return false;
    }
    return true;
  },

  /**
   * Validates the highest streak value (must be a non-negative number).
   * @private
   * @param {number} highestStreak - The highest streak value to validate.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if `highestStreak` is valid.
   */
  _validateHighestStreak: function (highestStreak, logError = true) {
    if (!this._validateNonNegativeNumber(highestStreak)) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: highestStreak (${highestStreak}) must be a non-negative number.`,
          false
        );
      return false;
    }
    return true;
  },

  /**
   * Validates current and highest streak values, ensuring highest >= current.
   * Also performs individual validation on each value.
   * @private
   * @param {number} currentStreak - The current streak value.
   * @param {number} highestStreak - The highest streak value.
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if both values are valid and highest >= current.
   */
  _validateCurrentAndHighestStreak: function (
    currentStreak,
    highestStreak,
    logError = true
  ) {
    const isCurrentValid = this._validateCurrentStreak(currentStreak, logError);
    const isHighestValid = this._validateHighestStreak(highestStreak, logError);

    if (!isCurrentValid || !isHighestValid) {
      return false; // Errors already handled if logError=true
    }

    if (currentStreak > highestStreak) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: highestStreak (${highestStreak}) must be >= currentStreak (${currentStreak}).`,
          false
        );
      return false;
    }
    return true;
  },

  /**
   * Validates if a history sheet row index is a positive integer and within the valid data range of the sheet.
   * @private
   * @param {number} row - The 1-indexed row number to validate (matching sheet row numbers).
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} Returns true if the row index is valid for reading existing data.
   */
  _validateHistorySheetRowIndex: function (row, logError = true) {
    if (!this._validatePositiveInteger(row)) {
      if (logError)
        LoggerManager.handleError(
          `Invalid input: row index (${row}) must be a positive integer.`,
          false
        );
      return false;
    }

    const sheet = HistorySheetConfig._getSheet(); // Use internal getter
    if (!sheet) {
      if (logError)
        LoggerManager.handleError(
          `History sheet not found for row index validation.`,
          false
        );
      return false; // Cannot validate if sheet doesn't exist
    }

    const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1; // Convert 0-based config to 1-based for comparison
    const lastRow = sheet.getLastRow();

    if (row < firstDataSheetRow) {
      if (logError)
        LoggerManager.handleError(
          `_validateHistorySheetRowIndex: Row index ${row} is below the first possible data row ${firstDataSheetRow}.`,
          false
        );
      return false;
    }
    // Check if the row index exceeds the last row that actually contains data.
    if (lastRow >= firstDataSheetRow && row > lastRow) {
      if (logError)
        LoggerManager.handleError(
          `_validateHistorySheetRowIndex: Row index ${row} is out of range. Last data row is ${lastRow}.`,
          false
        );
      return false;
    }
    // If lastRow < firstDataRow, sheet is empty or header-only. Any row >= firstDataSheetRow is invalid for *reading*.
    if (lastRow < firstDataSheetRow && row >= firstDataSheetRow) {
      if (logError)
        LoggerManager.handleError(
          `_validateHistorySheetRowIndex: Attempted to validate row index ${row} but history sheet contains no data rows (lastRow: ${lastRow}).`,
          false
        );
      return false;
    }

    return true;
  },

  /**
   * Validates that a history entry row array (data to be appended/read) has the correct number of columns.
   * @private
   * @param {Array} row - The array representing the row data (e.g., [dateStr, completionJson, bufferJson, currentStreak, highestStreak]).
   * @param {boolean} [logError=true] - Whether to log an error message on failure.
   * @returns {boolean} True if the row length matches the expected number of columns.
   */
  _validateHistoryEntryRow: function (row, logError = true) {
    if (!this._validateArray(row)) {
      if (logError)
        LoggerManager.handleError(
          `Invalid history entry: Must be an array`,
          false
        );
      return false;
    }
    const numExpectedColumns = HistorySheetConfig._getNumberOfColumns();

    if (row.length !== numExpectedColumns) {
      if (logError)
        LoggerManager.handleError(
          `History entry row length (${
            row.length
          }) does not match the expected ${numExpectedColumns} columns. Row data: ${JSON.stringify(
            row
          )}`,
          false
        );
      return false;
    }
    // TODO: Optionally add type validation for each element (date string, string, string, number, number)
    return true;
  },
};

// Freeze the utility object to prevent modification at runtime.
Object.freeze(ValidationUtils);
