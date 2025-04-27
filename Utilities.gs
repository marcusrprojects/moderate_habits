/**
 * Manages version control and checks for the latest version from an online CSV file.
 */
const LibraryManager = {
  /**
   * The latest version number of the script.
   *
   * @constant {string}
   * @default
   */
  LATEST_VERSION: "6",

  /**
   * The URL pointing to a CSV file that contains the latest version information.
   *
   * The CSV file is hosted on Google Sheets and the data can be fetched programmatically
   * to check if the user's version is up to date.
   *
   * @constant {string}
   * @default
   */
  LATEST_VERSION_CSV: `https://docs.google.com/spreadsheets/d/e/2PACX-1vTw2YxOfHTpUCcczl3G-rSUNhUe6OEMs1WhLypmZ4uMU_MBMbhqEeWfNvI7MdwK4ln-JRDhXPhhTCMF/pub?gid=0&single=true&output=csv`,

  /**
   * Fetches the current version information from the CSV file hosted online.
   * @returns {string} - The current version number.
   */
  fetchVersionInfo: function () {
    const response = UrlFetchApp.fetch(this.LATEST_VERSION_CSV);
    const csvData = Utilities.parseCsv(response.getContentText());

    // Assuming the version number is in the first cell of the CSV
    const currentVersion = csvData[0][0];
    Logger.log("Current Version: " + currentVersion);

    return currentVersion;
  },
};

/**
 * Manages the creation and validation of time-based triggers
 * for automatically resetting the checklist daily.
 */
const TriggerManager = {
  /**
   * Checks if the 'renewChecklistForToday' trigger already exists for the user.
   * @returns {boolean} - True if the trigger exists, false otherwise.
   */
  checkTriggerExists: function () {
    const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
    return triggers.some(
      (trigger) => trigger.getHandlerFunction() === "renewChecklistForToday"
    );
  },

  /**
   * Creates a time-based trigger that resets the checklist every day at a specified hour.
   * This ensures that the daily checklist is automatically updated at a consistent time.
   *
   * The hour is set using the resetHour, which allows customization of the reset time.
   * The trigger is set to fire daily at the specified hour.
   */
  createTrigger: function () {
    if (this.checkTriggerExists()) {
      LoggerManager.logDebug(`Trigger already exists.`);
      return;
    }

    LoggerManager.logDebug("Creating trigger...");

    // Create a new time-based trigger to call 'renewChecklistForToday' at a specified hour every day
    ScriptApp.newTrigger("renewChecklistForToday")
      .timeBased()
      .atHour(PropertyManager.getPropertyNumber(PropertyKeys.RESET_HOUR)) // Set the time you want the checklist to reset (e.g., 0 for midnight)
      .everyDays(1)
      .create();
  },
};

/**
 * LoggerManager object for handling logging and error management.
 * Includes methods for loading properties, retrieving property values, and managing errors.
 */
const LoggerManager = {
  /**
   * Switch this on or off for debugging.
   * When set to `true`, debugging messages will be logged via the Logger.
   * Set to `false` to turn off debugging logs.
   *
   * @constant {boolean}
   */
  DEBUG_MODE: true,

  /**
   * Logs a debug message if `DEBUG_MODE` is set to `true`.
   *
   * @param {string} message - The message to log to the Logger.
   */
  logDebug: function (message) {
    if (this.DEBUG_MODE) {
      Logger.log(message);
    }
  },

  /**
   * Handles errors by showing an alert to the user and throwing an error to stop execution.
   * @param {string} errorMessage - The error message to display and throw.
   * @param {boolean} throwError - Whether to throw an error or return false.
   * @returns {boolean} - Always returns false.
   */
  handleError: function (errorMessage, throwError = true) {
    if (throwError) {
      // Show the error message to the user
      this.logDebug(errorMessage);
      SpreadsheetApp.getUi().alert(
        `An error has occurred in processing this request.`
      );

      // Throw the error to stop execution
      throw new Error(errorMessage);
    } else {
      this.logDebug(`ERROR: ${errorMessage}`);
    }
  },

  /**
   * Wraps a function call to measure and log the execution time.
   * @param {Function} func - The function to be executed.
   * @returns {*} The result of the function execution.
   */
  logExecutionTime: function (funcCall) {
    const startTime = new Date();

    // Execute the function call
    const result = funcCall();

    const endTime = new Date();
    const executionTime = endTime - startTime;

    // Log the execution time
    const funcName =
      funcCall.name || String(funcCall).slice(6).trim() || "Anonymous Function";
    this.logDebug(`${funcName} took ${executionTime} ms to execute.`);

    return result; // Return the result of the function
  },
};

/**
 * The `DateManager` constant is responsible for handling all operations related to date formatting
 * and manipulation within the Google Sheets environment. This includes converting dates to
 * ISO strings, deserializing raw date values, and managing any date-related logic that
 * needs to be standardized across the application.
 *
 * The `DateManager` ensures that date handling is consistent, minimizing errors related to
 * date formatting or misinterpretation. It abstracts the complexity of date manipulation
 * and provides a clean interface for other parts of the code to interact with dates.
 */
const DateManager = {
  /**
   * Validates if the input is a valid date string in the format YYYY-MM-DD.
   * The input must be a formatted date string and cannot be a Date object.
   *
   * @param {any} date - The data to validate.
   * @param {boolean} [throwError=true] - Whether to throw an error on invalid input.
   * @returns {boolean} True if valid, False otherwise.
   */
  __validateDateStr: function (dateStr) {
    const regex = /^\d{4}-\d{2}-\d{2}$/;
    return (
      typeof dateStr === "string" &&
      regex.test(dateStr) &&
      !isNaN(Date.parse(dateStr))
    );
  },

  /**
   * Validates whether the given date is within the allowed range (firstDate <= date <= today).
   *
   * This function ensures that the input `date` (which must be a valid Date object) is within the
   * allowable range. The range is defined as the period between the first recorded date in the
   * history sheet and the current date (today). If the history sheet is empty, today's date
   * is used as the first date.
   *
   * @param {Date} date - The Date object to validate.
   * @returns {boolean} True if the date is valid and falls within the allowed range, False otherwise.
   */
  _validateDateRange: function (date, throwError = true) {
    // Ensure the input date is a valid Date object.
    if (!UtilsManager.__validateDate(date)) {
      LoggerManager.handleError(`Invalid date input.`, throwError);
      return false;
    }

    const today = this.getToday();
    const firstDate = HistorySheetConfig.getFirstDate() || today; // Default to today if no history

    LoggerManager.logDebug(
      `_validateDateRange: input date is ${date}, compared to the first date of ${firstDate} and today ${today}`
    );

    // Check if the date is within the valid range (firstDate <= date <= today)
    if (date < firstDate || date > today) {
      LoggerManager.handleError(
        `Invalid date: The date must be between ${firstDate} and ${today}, but is ${date} instead.`,
        throwError
      );
      return false;
    }

    return true; // Date is valid
  },

  /**
   * Validates whether the given date string is within the allowed range (firstDate <= dateStr <= today).
   *
   * This function converts the input `dateStr` (expected in `YYYY-MM-DD` format) into a Date object,
   * then checks whether it falls within the valid date range: from the first date recorded in the history
   * sheet up to today's date.
   *
   * @param {string} dateStr - The date string to validate, in `YYYY-MM-DD` format.
   * @returns {boolean} True if the date string is valid and falls within the allowed range, False otherwise.
   */
  _validateDateStrRange: function (dateStr, throwError = true) {
    if (!this.__validateDateStr(dateStr)) {
      LoggerManager.handleError(
        `_validateDateStrRange: Date string ${dateStr} is invalid.`,
        throwError
      );
      return false;
    }

    const today = this.getToday();
    const firstDate = HistorySheetConfig.getFirstDate() || today; // Default to today if no history

    // Parse the date string into a Date object, and validate it.
    const date = this.determineDate(dateStr);

    LoggerManager.logDebug(
      `_validateDateStrRange: input date is ${date}, compared to the first date of ${firstDate} and today ${today}`
    );

    // Check if the date is within the valid range (firstDate <= date <= today)
    if (date < firstDate || date > today) {
      LoggerManager.handleError(
        `_validateDateStrRange: Invalid date: The date must be between ${firstDate} and ${today}, but is ${date} instead.`,
        throwError
      );
      return false;
    }

    return true; // Date is valid
  },

  /**
   * Determines the appropriate method for formatting a given date input.
   * It will either deserialize the input or format it into an ISO string (YYYY-MM-DD).
   *
   * @param {Date|number|string} dateInput - The date input, which can be a Date object, a serialized date number, or a date string.
   * @returns {string} - The formatted date string in YYYY-MM-DD format.
   */
  determineFormattedDate: function (dateInput) {
    const date = this.determineDate(dateInput);
    return this._formatDateToISOString(date);
  },

  /**
   * Determines whether to deserialize the date input or use it as a Date object.
   * Converts serialized date numbers or date strings to a Date object if necessary.
   *
   * @param {Date|number|string} dateInput - The date input, which can be either a Date object, serialized date number, or date string.
   * @returns {Date} - The Date object.
   */
  determineDate: function (dateInput) {
    if (dateInput instanceof Date) {
      return dateInput;
    }

    if (
      typeof dateInput === "string" &&
      !isNaN(new Date(dateInput).getTime())
    ) {
      const parts = dateInput.split("-");
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Months are 0-indexed in JS
      const day = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }

    if (
      typeof dateInput === "number" ||
      (typeof dateInput === "string" && !isNaN(dateInput))
    ) {
      // If the input is a number, treat it as a serialized date and deserialize it.
      return this._deserializeDate(dateInput);
    }

    if (
      typeof dateInput === "object" &&
      !isNaN(new Date(dateInput).getTime())
    ) {
      const date = new Date(dateInput);
      return new Date(date.getFullYear(), date.getMonth(), date.getDate()); // Sets time to 00:00:00
    }

    LoggerManager.handleError(
      `Invalid date input type. Given ${dateInput} of type ${typeof dateInput}. Must be a Date object or a serialized date number.`
    );
  },

  /**
   * Converts a Date object into a string in the YYYY-MM-DD format (ISO 8601).
   *
   * @param {Date} date - The Date object to format.
   * @returns {string} - The formatted date string in YYYY-MM-DD format.
   */
  _formatDateToISOString: function (date) {
    return date.toISOString().split("T")[0];
  },

  /**
   * Deserializes a serialized date number into a Date object.
   *
   * @param {number} serializedDate - The serialized date value (Excel-style serial number).
   * @returns {Date} - The deserialized Date object.
   */
  _deserializeDate: function (serializedDate) {
    // Multiply by 86400000 (milliseconds per day) to convert serial number to milliseconds
    return new Date((serializedDate - 25569) * 86400 * 1000);
  },

  /**
   * Returns the previous date as a Date object.
   *
   * @param {string|Date} dateStr - The date string (YYYY-MM-DD) or Date object to calculate the previous date from.
   * @returns {Date} - The previous date as a Date object.
   */
  getPreviousDate: function (dateStr) {
    // Convert dateStr to a Date object
    const previousDay = this.determineDate(dateStr);
    // Create a Date object from the date string and subtract one day
    previousDay.setDate(previousDay.getDate() - 1); // Subtract 1 day
    LoggerManager.logDebug(
      `getPreviousDate: passed in date ${dateStr} v.s. determined previous date ${previousDay}`
    );
    return previousDay;
  },

  /**
   * Calculates and returns the previous date as a formatted string (YYYY-MM-DD).
   *
   * @param {string} dateStr - The date string in `YYYY-MM-DD` format.
   * @returns {string} - The formatted previous date string in `YYYY-MM-DD` format.
   */
  getPreviousDateStr: function (dateStr) {
    // Create a Date object from the date string
    const previousDay = this.getPreviousDate(dateStr);

    // Format the previous day
    return this.determineFormattedDate(previousDay);
  },

  /**
   * Returns the current timestamp in ISO 8601 format.
   *
   * This method captures the current date and time and formats it
   * to an ISO 8601 string, which includes the full timestamp with
   * the year, month, day, hours, minutes, seconds, and milliseconds
   * in the format `YYYY-MM-DDTHH:mm:ss.sssZ`.
   *
   * Example output: '2024-09-07T12:34:56.789Z'
   *
   * @returns {string} - The current timestamp in ISO 8601 format.
   */
  getNow: function () {
    return new Date().toISOString();
  },

  /**
   * Returns today's date as a Date object.
   *
   * @returns {Date} - The current date as a Date object.
   */
  getToday: function () {
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Ensure hours, minutes, seconds, and milliseconds are set to zero
    return today;
  },

  /**
   * Returns today's date as a formatted string (YYYY-MM-DD).
   *
   * @returns {string} - The current date formatted as YYYY-MM-DD.
   */
  getTodayStr: function () {
    return this.determineFormattedDate(this.getToday());
  },
};

/**
 * The `PropertyManager` constant is responsible for managing document properties within
 * Google Sheets. It handles the storage, retrieval, and updating of properties that are
 * essential for tracking the state of the document, such as timestamps of the last
 * updates or other important metadata.
 *
 * The `PropertyManager` abstracts the complexity of interacting with Google Sheets'
 * PropertiesService, providing an easy-to-use interface for storing and retrieving
 * key-value pairs. It plays a crucial role in maintaining the state and ensuring that
 * the application behaves consistently across sessions.
 */
const PropertyManager = {
  /**
   * Object to hold all loaded properties from the document.
   *
   * Once the properties are loaded from the PropertiesService, they are cached in this object
   * for quicker access. This prevents multiple calls to the PropertiesService, improving performance.
   *
   * @type {Object.<string, string>}
   */
  properties: {},

  /**
   * Boolean flag that tracks whether any properties have been modified.
   *
   * This flag is used to decide whether the properties need to be saved back to the
   * PropertiesService after being changed during execution.
   *
   * @type {boolean}
   * @default
   */
  hasChanged: false,

  /**
   * Boolean flag indicating whether the properties have been loaded into memory.
   *
   * This flag ensures that properties are only loaded once during the script's runtime.
   * It prevents unnecessary repeated calls to load the same properties.
   *
   * @type {boolean}
   * @default
   */
  loaded: false,

  /**
   * Loads all document properties into memory from PropertiesService.
   */
  loadDocumentProperties: function () {
    LoggerManager.logDebug(
      `Loading all document properties from PropertyService...`
    );
    const allProperties =
      PropertiesService.getDocumentProperties().getProperties(); // Get all properties at once
    LoggerManager.logDebug(`allProperties: ${allProperties}`);
    this.properties = allProperties || {}; // Store in-memory properties
    LoggerManager.logDebug(
      `Loaded properties: ${JSON.stringify(this.properties)}.`
    );
    this.loaded = true;
  },

  /**
   * Retrieves a property value, initializing it to a default value if not found.
   * @param {string} key - The property key.
   * @returns {string} - The property value, or a default if the key does not exist.
   */
  getProperty: function (key) {
    if (!this.loaded) {
      this.loadDocumentProperties();
    }
    if (!(key in this.properties)) {
      LoggerManager.logDebug(
        `Property ${key} not found. Initializing to default value...`
      );
      this.properties[key] = this.initializeDefaultProperty(key); // Initialize with default value
    }
    return this.properties[key];
  },

  /**
   * Checks if a given property exists in the loaded properties.
   * @param {string} key - The property key.
   * @returns {boolean} - True if the property exists, false otherwise.
   */
  hasProperty: function (key) {
    if (!this.loaded) {
      this.loadDocumentProperties();
    }
    return key in this.properties;
  },

  /**
   * Sets a property value and optionally forces saving the properties immediately.
   * @param {string} key - The property key.
   * @param {string} value - The value to set for the property.
   * @param {boolean} [forceSet=false] - Whether to immediately save the property changes.
   */
  setProperty: function (key, value, forceSet = false) {
    // Check if the key exists in PropertyKeys
    if (!Object.values(PropertyKeys).includes(key)) {
      this.handleError(
        `Invalid property key: ${key}. Allowed keys are: ${Object.values(
          PropertyKeys
        ).join(", ")}`
      );
      return;
    }

    // If the value is not a string, stringify it before storing
    this.properties[key] =
      typeof value === "string" ? value : JSON.stringify(value);
    this.hasChanged = true; // Mark that properties were changed

    if (forceSet) {
      this.setDocumentProperties();
    }
  },

  /**
   * Initializes a default value for a property if it does not exist.
   * @param {string} key - The property key to initialize.
   * @returns {string} - The default value for the property.
   */
  initializeDefaultProperty: function (key) {
    switch (key) {
      case PropertyKeys.LAST_DATE_SELECTOR_UPDATE:
        return new Date("2024-01-01T00:00:00Z").toISOString(); // Hardcoding so that LAST_DATE_SELECTOR_UPDATE always initializes before LAST_COMPLETION_UPDATE

      case PropertyKeys.LAST_COMPLETION_UPDATE:
        return DateManager.getNow();

      // Group cases that return true
      case PropertyKeys.ACTIVITIES_COLUMN_UPDATED:
      case PropertyKeys.MODE:
        return ModeTypes.HABIT_IDEATION;

      // Handle both 'firstChallengeDate' and 'firstChallengeRow' together
      case PropertyKeys.FIRST_CHALLENGE_DATE:
      case PropertyKeys.FIRST_CHALLENGE_ROW:
        this.updateFirstChallengeDateAndRow(); // Set both together if needed
        return this.properties[key]; // Return the newly set value

      // Handle emojiList separately
      case PropertyKeys.EMOJI_LIST:
        this.updateEmojiSpread(); // Initialize emoji list
        return this.properties[key];

      case PropertyKeys.LAST_UPDATE:
        return this.updateLastUpdateProperty();

      case PropertyKeys.BOOST_INTERVAL:
        return HistorySheetConfig.boostIntervalDefault;

      case PropertyKeys.RESET_HOUR:
        return MainSheetConfig.resetHourDefault;

      // Default case for unrecognized properties
      default:
        LoggerManager.handleError(
          `Attempting to get a property that does not exist: ${key}.`
        );
        return ""; // Return an empty string for unrecognized properties
    }
  },

  /**
   * Updates the stored emoji spread by fetching the current emoji distribution
   * from the main sheet and storing it as a JSON string.
   */
  updateEmojiSpread: function () {
    const currentEmojiSpread = MainSheetConfig.getCurrentEmojiSpread();
    this.setProperty(
      PropertyKeys.EMOJI_LIST,
      JSON.stringify(currentEmojiSpread)
    );
  },

  /**
   * Updates the 'firstChallengeDate' and 'firstChallengeRow' properties.
   * Sets the first challenge date to today and calculates the first row for the challenge.
   */
  updateFirstChallengeDateAndRow: function () {
    const todayStr = DateManager.getTodayStr();
    this.setProperty(PropertyKeys.FIRST_CHALLENGE_DATE, todayStr);
    LoggerManager.logDebug(`First challenge date updated to: ${todayStr}`);

    const dateColumn = HistorySheetConfig.dateColumn + 1;
    const lastRow = HistorySheetConfig.getSheet().getLastRow();
    const firstDataRow = HistorySheetConfig.firstDataRow + 1; // converting to 1-indexed value.
    let firstRow;

    if (lastRow < firstDataRow) {
      firstRow = firstDataRow;
    } else {
      const sheet = HistorySheetConfig.getSheet();
      const dateAtLastRow = sheet.getRange(lastRow, dateColumn).getValue(); // Get the Date object from the last row
      const formattedDate = DateManager.determineFormattedDate(dateAtLastRow);
      firstRow = formattedDate === todayStr ? lastRow : lastRow + 1;
    }

    firstRow -= 1;

    this.setProperty(PropertyKeys.FIRST_CHALLENGE_ROW, firstRow); // Store it as a 0-indexed number
    LoggerManager.logDebug(
      `First challenge row set to ${firstRow} for date ${todayStr} with lastRow of ${lastRow}`
    );
  },

  /**
   * Updates the `lastUpdate` property based on which of the tracked timestamps
   * (`lastCompletionUpdate` or `lastDateSelectorUpdate`) is more recent.
   *
   * @returns string - The value of the property.
   *
   * This method compares the timestamps of the last completion update and the last date selector
   * update, and sets the `lastUpdate` property to either completion or dateSelector
   * depending on which action occurred most recently.
   *
   * This ensures that subsequent logic can determine the most recent change and act accordingly.
   */
  updateLastUpdateProperty: function () {
    // Get the current values of the properties
    const lastDateSelectorUpdate = this.getProperty(
      PropertyKeys.LAST_DATE_SELECTOR_UPDATE
    );
    const lastCompletionUpdate = this.getProperty(
      PropertyKeys.LAST_COMPLETION_UPDATE
    );

    const newLastUpdate =
      lastCompletionUpdate >= lastDateSelectorUpdate
        ? LastUpdateTypes.COMPLETION
        : LastUpdateTypes.DATE_SELECTOR;
    this.setProperty(PropertyKeys.LAST_UPDATE, newLastUpdate);
    return newLastUpdate;
  },

  /**
   * Retrieves the property value for the given key, converts it to a number, and returns it.
   * If the property value is not a valid number, logs an error and returns null.
   *
   * @param {string} key - The key for the property to retrieve.
   * @returns {number|null} - The numerical value of the property or null if invalid.
   */
  getPropertyNumber: function (key) {
    const value = this.getProperty(key);
    const numValue = Number(value);

    if (isNaN(numValue)) {
      LoggerManager.handleError(
        `getPropertyNumber: Property ${key} is not a valid number`
      );
      return null; // Or handle it in another way, like throwing an error
    }

    return numValue;
  },

  /**
   * Saves the modified properties to the document if any changes have occurred.
   */
  setDocumentProperties: function () {
    LoggerManager.logDebug(
      `setDocumentProperties. current properties: ${JSON.stringify(
        this.properties
      )}`
    );
    if (this.hasChanged) {
      PropertiesService.getDocumentProperties().setProperties(this.properties);
      LoggerManager.logDebug(
        `Properties saved to document: ${JSON.stringify(this.properties)}.`
      );
      this.hasChanged = false; // Reset the flag after saving
    }
  },
};
