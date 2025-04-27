/**
 * @fileoverview General utility modules including version checking, trigger management,
 * logging, date handling, and property management.
 */

/** OnlyCurrentDoc */

/**
 * Manages checking for updates against a remote source.
 */
const LibraryManager = {
  /** @constant {string} LATEST_VERSION - The current hardcoded version of this script. */
  LATEST_VERSION: "6", // Update this when releasing a new version

  /** @constant {string} LATEST_VERSION_CSV_URL - URL to fetch the latest version number. */
  LATEST_VERSION_CSV_URL: `https://docs.google.com/spreadsheets/d/e/2PACX-1vTw2YxOfHTpUCcczl3G-rSUNhUe6OEMs1WhLypmZ4uMU_MBMbhqEeWfNvI7MdwK4ln-JRDhXPhhTCMF/pub?gid=0&single=true&output=csv`,

  /**
   * Fetches the latest version number from the remote CSV.
   * @returns {string | null} The latest version number as a string, or null on error.
   */
  fetchVersionInfo: function () {
    try {
      const response = UrlFetchApp.fetch(this.LATEST_VERSION_CSV_URL, {
        muteHttpExceptions: true,
      });
      const responseCode = response.getResponseCode();
      const content = response.getContentText();

      if (responseCode === 200) {
        const csvData = Utilities.parseCsv(content);
        if (csvData && csvData.length > 0 && csvData[0].length > 0) {
          const currentVersion = String(csvData[0][0]).trim();
          LoggerManager.logDebug(`Fetched latest version: ${currentVersion}`);
          return currentVersion;
        } else {
          LoggerManager.handleError(
            "Failed to parse version data from CSV.",
            false
          );
          return null;
        }
      } else {
        LoggerManager.handleError(
          `Failed to fetch version info. Response code: ${responseCode}, Content: ${content}`,
          false
        );
        return null;
      }
    } catch (e) {
      LoggerManager.handleError(
        `Error fetching version info: ${e.message}`,
        false
      );
      return null;
    }
  },
};
Object.freeze(LibraryManager);

/**
 * Manages time-driven triggers for the script.
 */
const TriggerManager = {
  _HANDLER_FUNCTION_NAME: "renewChecklistForToday",

  /**
   * Checks if the daily reset trigger already exists.
   * @returns {boolean} True if the trigger exists, false otherwise.
   */
  _checkTriggerExists: function () {
    try {
      const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
      return triggers.some(
        (trigger) =>
          trigger.getHandlerFunction() === this._HANDLER_FUNCTION_NAME
      );
    } catch (e) {
      LoggerManager.handleError(
        `Error checking for existing triggers: ${e.message}`,
        false
      );
      return false; // Assume it doesn't exist on error
    }
  },

  /**
   * Deletes all existing triggers for the daily reset function.
   * @private
   */
  _deleteExistingTriggers: function () {
    try {
      const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
      triggers.forEach((trigger) => {
        if (trigger.getHandlerFunction() === this._HANDLER_FUNCTION_NAME) {
          ScriptApp.deleteTrigger(trigger);
          LoggerManager.logDebug(
            `Deleted existing trigger with ID: ${trigger.getUniqueId()}`
          );
        }
      });
    } catch (e) {
      LoggerManager.handleError(
        `Error deleting existing triggers: ${e.message}`,
        false
      );
    }
  },

  /**
   * Creates the daily time-based trigger if it doesn't exist.
   * Uses the reset hour defined in properties. Deletes old triggers first.
   */
  createTrigger: function () {
    this._deleteExistingTriggers(); // Ensure only one trigger exists

    if (this._checkTriggerExists()) {
      LoggerManager.logDebug(
        `Trigger '${this._HANDLER_FUNCTION_NAME}' already exists.`
      );
      return;
    }

    const resetHour = PropertyManager.getPropertyNumber(
      PropertyKeys.RESET_HOUR
    );
    // Validate resetHour (0-23)
    if (resetHour === null || resetHour < 0 || resetHour > 23) {
      LoggerManager.handleError(
        `Invalid reset hour (${resetHour}) found in properties. Cannot create trigger. Using default ${MainSheetConfig.resetHourDefault}.`
      );
      // Optionally default here, or just fail. Let's try defaulting.
      // resetHour = MainSheetConfig.resetHourDefault;
      // For safety, let's just not create the trigger if the value is bad.
      return;
    }

    LoggerManager.logDebug(
      `Creating time-driven trigger for ${this._HANDLER_FUNCTION_NAME} at hour ${resetHour}.`
    );
    try {
      ScriptApp.newTrigger(this._HANDLER_FUNCTION_NAME)
        .timeBased()
        .atHour(resetHour)
        .everyDays(1)
        .create();
      LoggerManager.logDebug("Trigger created successfully.");
    } catch (e) {
      LoggerManager.handleError(`Failed to create trigger: ${e.message}`, true);
    }
  },
};
Object.freeze(TriggerManager);

/**
 * Manages logging and basic error reporting.
 */
const LoggerManager = {
  /** @constant {boolean} DEBUG_MODE - Set to true to enable debug logging. */
  DEBUG_MODE: false, // Default to false for production releases

  /**
   * Logs a message if DEBUG_MODE is true.
   * @param {string} message - The message to log.
   */
  logDebug: function (message) {
    if (this.DEBUG_MODE) {
      Logger.log(`DEBUG: ${message}`);
    }
  },

  /**
   * Handles errors. Logs the error message. If throwError is true,
   * shows a generic alert to the user and throws the error to stop execution.
   * @param {string} errorMessage - The detailed error message.
   * @param {boolean} [throwError=true] - Whether to throw the error and show UI alert.
   */
  handleError: function (errorMessage, throwError = true) {
    Logger.log(`ERROR: ${errorMessage}`); // Always log the error

    if (throwError) {
      try {
        // Show a generic message to the user
        SpreadsheetApp.getUi().alert(
          `An error occurred. Please check the logs (Extensions > Apps Script > Executions) for details or try again later.`
        );
      } catch (uiError) {
        // Ignore UI errors if running in a context without UI access
        Logger.log(`UI Alert failed: ${uiError.message}`);
      }
      // Throw the error to stop execution and provide stack trace in logs
      throw new Error(errorMessage);
    }
    // If throwError is false, execution continues after logging.
  },

  /**
   * Wraps a function call to measure and log its execution time.
   * Only logs if DEBUG_MODE is true.
   * @param {Function} func - The function to execute.
   * @param {string} [funcName='Anonymous Function'] - Optional name for logging.
   * @returns {*} The result of the function execution.
   */
  logExecutionTime: function (func, funcName = "Anonymous Function") {
    if (!this.DEBUG_MODE) {
      return func(); // Just execute if not debugging
    }

    const startTime = new Date().getTime();
    let result;
    let errorOccurred = false;
    try {
      result = func();
    } catch (e) {
      errorOccurred = true;
      LoggerManager.handleError(
        `Error during timed execution of ${funcName}: ${e.message}`,
        false
      ); // Log error but let caller handle re-throwing if needed
      throw e; // Re-throw the original error
    } finally {
      const endTime = new Date().getTime();
      const executionTime = endTime - startTime;
      this.logDebug(
        `${funcName} ${
          errorOccurred ? "failed" : "completed"
        } in ${executionTime} ms.`
      );
    }
    return result;
  },
};
Object.freeze(LoggerManager);

/**
 * Manages date operations, formatting, and validation.
 */
const DateManager = {
  /**
   * Validates if a string is in YYYY-MM-DD format and represents a real date.
   * @param {string} dateStr - The string to validate.
   * @param {boolean} [throwError=true] - Whether to log/throw error on failure.
   * @returns {boolean} True if valid.
   */
  _validateDateStr: function (dateStr, throwError = true) {
    if (typeof dateStr !== "string") {
      if (throwError)
        LoggerManager.handleError(
          `Invalid date string type: Expected string, got ${typeof dateStr}`,
          true
        );
      return false;
    }
    const regex = /^\d{4}-\d{2}-\d{2}$/;
    if (!regex.test(dateStr)) {
      if (throwError)
        LoggerManager.handleError(
          `Invalid date string format: ${dateStr}. Expected YYYY-MM-DD.`,
          true
        );
      return false;
    }
    // Check if it parses to a valid date
    const date = new Date(dateStr + "T00:00:00Z"); // Use UTC to avoid timezone issues in parsing check
    if (isNaN(date.getTime())) {
      if (throwError)
        LoggerManager.handleError(
          `Invalid date value: ${dateStr}. Does not represent a real date.`,
          true
        );
      return false;
    }
    // Optional: Check if the parsed date parts match the input string parts to catch month/day rollovers like '2023-02-30'
    const [year, month, day] = dateStr.split("-").map(Number);
    if (
      date.getUTCFullYear() !== year ||
      date.getUTCMonth() + 1 !== month ||
      date.getUTCDate() !== day
    ) {
      if (throwError)
        LoggerManager.handleError(
          `Invalid date value: ${dateStr}. Month/day out of range.`,
          true
        );
      return false;
    }

    return true;
  },

  /**
   * Validates if a Date object is within the allowed challenge range (first challenge date to today).
   * @param {Date} date - The Date object to validate.
   * @param {boolean} [throwError=true] - Whether to log/throw error on failure.
   * @returns {boolean} True if valid.
   */
  _validateDateRange: function (date, throwError = true) {
    if (!ValidationUtils._validateDate(date)) {
      LoggerManager.handleError(
        `Invalid date object passed to _validateDateRange.`,
        throwError
      );
      return false;
    }

    const today = this.getToday();
    const firstDate = DataHandler.getFirstChallengeDate(); // Get from DataHandler

    if (!firstDate) {
      LoggerManager.handleError(
        `Cannot validate date range: First challenge date is not set.`,
        throwError
      );
      // Allow validation if date is today or earlier? Or always fail? Let's fail.
      return false;
    }

    // Normalize dates to midnight UTC for comparison to avoid timezone issues
    const dateUTC = new Date(
      Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
    );
    const firstDateUTC = new Date(
      Date.UTC(
        firstDate.getFullYear(),
        firstDate.getMonth(),
        firstDate.getDate()
      )
    );
    const todayUTC = new Date(
      Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())
    );

    if (dateUTC < firstDateUTC || dateUTC > todayUTC) {
      const msg = `Invalid date range: Date ${this.determineFormattedDate(
        date
      )} must be between ${this.determineFormattedDate(
        firstDate
      )} and ${this.determineFormattedDate(today)}.`;
      LoggerManager.handleError(msg, throwError);
      return false;
    }
    return true;
  },

  /**
   * Validates if a date string (YYYY-MM-DD) is within the allowed challenge range.
   * @param {string} dateStr - The date string to validate.
   * @param {boolean} [throwError=true] - Whether to log/throw error on failure.
   * @returns {boolean} True if valid.
   */
  _validateDateStrRange: function (dateStr, throwError = true) {
    if (!this._validateDateStr(dateStr, throwError)) {
      return false; // Format/value validation failed
    }
    try {
      const date = this.determineDate(dateStr); // Convert string to Date object
      return this._validateDateRange(date, throwError); // Use Date object validation
    } catch (e) {
      LoggerManager.handleError(
        `Error validating date string range for ${dateStr}: ${e.message}`,
        throwError
      );
      return false;
    }
  },

  /**
   * Converts various date inputs (Date object, string, Sheets serial number) into a Date object normalized to midnight UTC.
   * @param {Date|string|number} dateInput - The input date.
   * @returns {Date} The normalized Date object.
   * @throws {Error} if the input cannot be parsed.
   */
  determineDate: function (dateInput) {
    let year, month, day;

    if (dateInput instanceof Date && !isNaN(dateInput)) {
      // Input is already a valid Date object
      year = dateInput.getFullYear();
      month = dateInput.getMonth(); // 0-indexed
      day = dateInput.getDate();
    } else if (
      typeof dateInput === "string" &&
      this._validateDateStr(dateInput, false)
    ) {
      // Input is a 'YYYY-MM-DD' string
      [year, month, day] = dateInput.split("-").map(Number);
      month -= 1; // Adjust month to 0-indexed
    } else if (typeof dateInput === "number" && dateInput > 0) {
      // Assume Sheets Serial Number (days since Dec 30, 1899)
      // Formula: (Serial Number - 25569) * 86400000 = Milliseconds since Unix Epoch UTC
      // Need to handle timezone carefully. Create date from components.
      const jsDate = new Date(Math.round((dateInput - 25569) * 86400 * 1000));
      // Extract components in UTC to avoid local timezone shifts affecting the date parts
      year = jsDate.getUTCFullYear();
      month = jsDate.getUTCMonth(); // 0-indexed
      day = jsDate.getUTCDate();
    } else {
      throw new Error(
        `Invalid date input type or value: ${dateInput} (Type: ${typeof dateInput})`
      );
    }

    // Return a new Date object set to midnight UTC for consistency
    return new Date(Date.UTC(year, month, day));
  },

  /**
   * Formats a Date object into a "YYYY-MM-DD" string.
   * @param {Date} date - The Date object.
   * @returns {string} The formatted date string.
   */
  determineFormattedDate: function (date) {
    if (!ValidationUtils._validateDate(date)) {
      throw new Error("Invalid Date object passed to determineFormattedDate.");
    }
    // Use UTC methods to avoid timezone shifting the date during formatting
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0"); // 0-indexed month
    const day = String(date.getUTCDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  },

  /**
   * Calculates the date for the day before the given date.
   * @param {Date|string} dateInput - The reference date (Date object or YYYY-MM-DD string).
   * @returns {Date} The Date object for the previous day.
   */
  getPreviousDate: function (dateInput) {
    const date = this.determineDate(dateInput); // Ensure we have a normalized Date object
    const previousDay = new Date(date); // Clone
    previousDay.setUTCDate(date.getUTCDate() - 1); // Subtract one day using UTC
    return previousDay;
  },

  /**
   * Calculates the date for the day after the given date.
   * @param {Date|string} dateInput - The reference date (Date object or YYYY-MM-DD string).
   * @returns {Date} The Date object for the next day.
   */
  getNextDate: function (dateInput) {
    const date = this.determineDate(dateInput); // Ensure we have a normalized Date object
    const nextDay = new Date(date); // Clone
    nextDay.setUTCDate(date.getUTCDate() + 1); // Add one day using UTC
    return nextDay;
  },

  /**
   * Calculates the previous date and returns it as a formatted string (YYYY-MM-DD).
   * @param {string} dateStr - The reference date string (YYYY-MM-DD).
   * @returns {string} The formatted previous date string.
   */
  getPreviousDateStr: function (dateStr) {
    const previousDate = this.getPreviousDate(dateStr);
    return this.determineFormattedDate(previousDate);
  },

  /**
   * Calculates the difference in days between two dates.
   * Result is positive if date2 is after date1.
   * @param {Date|string} date1Input - The earlier date.
   * @param {Date|string} date2Input - The later date.
   * @returns {number} The number of days between the dates.
   */
  daysBetween: function (date1Input, date2Input) {
    const date1 = this.determineDate(date1Input); // Normalized to midnight UTC
    const date2 = this.determineDate(date2Input); // Normalized to midnight UTC
    const msPerDay = 1000 * 60 * 60 * 24;
    // Calculate the difference in milliseconds and convert to days
    return Math.round((date2.getTime() - date1.getTime()) / msPerDay);
  },

  /**
   * Gets the current timestamp as an ISO 8601 string (UTC).
   * @returns {string} e.g., "2023-10-27T10:30:00.000Z"
   */
  getNow: function () {
    return new Date().toISOString();
  },

  /**
   * Gets today's date as a Date object, normalized to midnight UTC.
   * @returns {Date} Today's date.
   */
  getToday: function () {
    const today = new Date();
    // Normalize to midnight UTC of the current day
    return new Date(
      Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())
    );
  },

  /**
   * Gets today's date as a formatted string (YYYY-MM-DD).
   * @returns {string} Today's date string.
   */
  getTodayStr: function () {
    return this.determineFormattedDate(this.getToday());
  },
};
Object.freeze(DateManager);

/**
 * Manages script and document properties using PropertiesService.
 * Caches properties in memory for performance.
 */
const PropertyManager = {
  /** @private @type {Object<string, string> | null} _properties - In-memory cache. */
  _properties: null,
  /** @private @type {boolean} _hasChanged - Flag if properties need saving. */
  _hasChanged: false,

  /**
   * Loads properties from PropertiesService into the cache if not already loaded.
   * @private
   */
  _loadProperties: function () {
    if (this._properties === null) {
      LoggerManager.logDebug("Loading document properties into cache...");
      try {
        this._properties =
          PropertiesService.getDocumentProperties().getProperties();
        LoggerManager.logDebug(
          `Properties loaded: ${JSON.stringify(this._properties)}`
        );
      } catch (e) {
        LoggerManager.handleError(
          `Failed to load document properties: ${e.message}`,
          true
        );
        this._properties = {}; // Initialize empty on error to prevent repeated load attempts
      }
      this._hasChanged = false; // Reset changed flag after loading
    }
  },

  /**
   * Retrieves a property value from the cache, initializing to default if not found.
   * @param {PropertyKeys} key - The property key enum.
   * @returns {string} The property value (always a string).
   */
  getProperty: function (key) {
    this._loadProperties(); // Ensure properties are loaded
    if (!(key in this._properties)) {
      LoggerManager.logDebug(
        `Property '${key}' not found in cache. Initializing to default.`
      );
      this._properties[key] = this._initializeDefaultProperty(key);
      this._hasChanged = true; // Mark as changed since a default was set
    }
    return this._properties[key];
  },

  /**
   * Retrieves a property value and converts it to a number.
   * @param {PropertyKeys} key - The property key enum.
   * @returns {number | null} The numerical value, or null if not a valid number or not found.
   */
  getPropertyNumber: function (key) {
    const value = this.getProperty(key);
    const numValue = Number(value);
    if (isNaN(numValue)) {
      LoggerManager.handleError(
        `Property '${key}' value ('${value}') is not a valid number.`,
        false
      );
      return null;
    }
    return numValue;
  },

  /**
   * Checks if a property key exists in the cache.
   * @param {PropertyKeys} key - The property key enum.
   * @returns {boolean} True if the property exists.
   */
  hasProperty: function (key) {
    this._loadProperties(); // Ensure properties are loaded
    return key in this._properties;
  },

  /**
   * Sets a property value in the cache and marks for saving.
   * Ensures the key is valid according to PropertyKeys enum.
   * Converts non-string values to JSON strings.
   * @param {PropertyKeys} key - The property key enum.
   * @param {*} value - The value to set.
   * @param {boolean} [forceSave=false] - Whether to save properties immediately.
   */
  setProperty: function (key, value, forceSave = false) {
    // Validate the key
    if (!Object.values(PropertyKeys).includes(key)) {
      LoggerManager.handleError(
        `Invalid property key used in setProperty: ${key}`,
        true
      );
      return;
    }

    this._loadProperties(); // Ensure cache is initialized

    const stringValue =
      typeof value === "string" ? value : JSON.stringify(value);

    // Only mark as changed if the value is actually different
    if (this._properties[key] !== stringValue) {
      this._properties[key] = stringValue;
      this._hasChanged = true;
      LoggerManager.logDebug(
        `Property '${key}' set to '${stringValue}'. Marked for saving.`
      );
    } else {
      LoggerManager.logDebug(
        `Property '${key}' value is already '${stringValue}'. No change.`
      );
    }

    if (forceSave) {
      this.setDocumentProperties();
    }
  },

  /**
   * Saves cached properties to PropertiesService if changes have been made.
   */
  setDocumentProperties: function () {
    if (this._hasChanged && this._properties !== null) {
      LoggerManager.logDebug(
        "Saving changed properties to DocumentProperties..."
      );
      try {
        PropertiesService.getDocumentProperties().setProperties(
          this._properties,
          false
        ); // Set deleteAllOthers to false
        this._hasChanged = false; // Reset flag after successful save
        LoggerManager.logDebug(
          `Properties saved: ${JSON.stringify(this._properties)}`
        );
      } catch (e) {
        LoggerManager.handleError(
          `Failed to save document properties: ${e.message}`,
          true
        );
      }
    } else {
      LoggerManager.logDebug("No property changes to save.");
    }
  },

  /**
   * Provides default values for properties when they are first accessed.
   * @private
   * @param {PropertyKeys} key - The property key.
   * @returns {string} The default value as a string.
   */
  _initializeDefaultProperty: function (key) {
    LoggerManager.logDebug(`Initializing default value for property: ${key}`);
    switch (key) {
      case PropertyKeys.MODE:
        return ModeTypes.HABIT_IDEATION; // Start in setup mode

      case PropertyKeys.LAST_DATE_SELECTOR_UPDATE:
        // Initialize slightly in the past to ensure first completion update takes precedence
        return new Date(Date.now() - 60000).toISOString(); // 1 minute ago
      case PropertyKeys.LAST_COMPLETION_UPDATE:
        return DateManager.getNow(); // Initialize to current time
      case PropertyKeys.LAST_UPDATE:
        return this._recalculateLastUpdateProperty(); // Calculate based on other timestamps

      case PropertyKeys.ACTIVITIES_COLUMN_UPDATED:
        return BooleanTypes.FALSE;

      case PropertyKeys.FIRST_CHALLENGE_DATE:
      case PropertyKeys.FIRST_CHALLENGE_ROW:
        // This needs special handling as they depend on each other and sheet state
        this._updateFirstChallengeDateAndRow(); // Update both
        // Return the value *after* updating (check if it was set)
        return key in this._properties ? this._properties[key] : "";

      case PropertyKeys.EMOJI_LIST:
        // Default to empty list if called during first run setup before habits exist
        const mode =
          this._properties[PropertyKeys.MODE] || ModeTypes.HABIT_IDEATION; // Get current mode if loaded
        if (mode === ModeTypes.HABIT_IDEATION) {
          return "[]"; // Return empty array string during initial setup
        } else {
          // If called later and still missing, try to update from sheet (shouldn't happen ideally)
          HabitManager.updateEmojiSpreadProperty();
          return key in this._properties ? this._properties[key] : "[]";
        }

      case PropertyKeys.RESET_HOUR:
        return String(MainSheetConfig.resetHourDefault);
      case PropertyKeys.BOOST_INTERVAL:
        return String(HistorySheetConfig.boostIntervalDefault);

      default:
        LoggerManager.handleError(
          `Attempting to initialize unrecognized property key: ${key}`,
          false
        );
        return ""; // Return empty string for safety
    }
  },

  /**
   * Recalculates and returns the value for the LAST_UPDATE property based on other timestamps.
   * Does not set the property itself, only calculates the value.
   * @private
   * @returns {LastUpdateTypes}
   */
  _recalculateLastUpdateProperty: function () {
    // Ensure dependent properties are loaded/initialized first *without* causing infinite loops
    const lastDateUpdate = this.getProperty(
      PropertyKeys.LAST_DATE_SELECTOR_UPDATE
    );
    const lastCompUpdate = this.getProperty(
      PropertyKeys.LAST_COMPLETION_UPDATE
    );

    // Simple string comparison works for ISO dates
    return lastCompUpdate >= lastDateUpdate
      ? LastUpdateTypes.COMPLETION
      : LastUpdateTypes.DATE_SELECTOR;
  },

  /**
   * Updates the LAST_UPDATE property based on current timestamps.
   */
  updateLastUpdateProperty: function () {
    const newValue = this._recalculateLastUpdateProperty();
    this.setProperty(PropertyKeys.LAST_UPDATE, newValue);
  },

  /**
   * Updates the 'firstChallengeDate' and 'firstChallengeRow' properties.
   * Sets date to today, calculates row based on history sheet state.
   * This should only be called when a *new* challenge actually starts.
   * @private
   */
  _updateFirstChallengeDateAndRow: function () {
    const todayStr = DateManager.getTodayStr();
    this.setProperty(PropertyKeys.FIRST_CHALLENGE_DATE, todayStr); // Set date first
    LoggerManager.logDebug(`First challenge date property set to: ${todayStr}`);

    const historySheet = HistorySheetConfig._getSheet();
    let firstRowIndex = HistorySheetConfig.firstDataRow; // Default 0-based index

    if (historySheet) {
      const lastRow = historySheet.getLastRow();
      const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1; // 1-based

      if (lastRow >= firstDataSheetRow) {
        // If history has data, the new challenge starts *after* the last entry.
        // The 'first row' property likely refers to the row index in the history array/sheet.
        // A new entry will be appended at lastRow + 1.
        // What does FIRST_CHALLENGE_ROW actually represent? The index of the first entry *of this challenge*?
        // Let's assume it's the 0-based index where the *first* entry of this challenge *will be* or *is*.
        const lastDate = DataHandler.getLastHistoryDate(); // Use DataHandler method
        if (
          lastDate &&
          DateManager.determineFormattedDate(lastDate) === todayStr
        ) {
          // If the last entry is already today (e.g., reset on same day), reuse that row index.
          firstRowIndex = lastRow - 1; // Convert 1-based lastRow to 0-based index
        } else {
          // Otherwise, the first entry will be the next row index.
          firstRowIndex = lastRow; // lastRow is 1-based, which equals the 0-based index of the *next* row
        }
      } else {
        // History is empty or only has headers, first entry will be at the first data row index.
        firstRowIndex = HistorySheetConfig.firstDataRow; // 0-based index
      }
    } else {
      LoggerManager.handleError(
        "History sheet not found while updating first challenge row property.",
        false
      );
      // Keep default firstRowIndex
    }

    this.setProperty(PropertyKeys.FIRST_CHALLENGE_ROW, String(firstRowIndex)); // Store as string
    LoggerManager.logDebug(
      `First challenge row property (0-based index) set to: ${firstRowIndex}`
    );
  },
};
// No freeze here, properties need to be mutable internally
