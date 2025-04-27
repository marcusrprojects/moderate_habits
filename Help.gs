/**
 * @fileoverview Manages the custom menu and help sidebar functionality.
 */

/** OnlyCurrentDoc */

/**
 * Manages UI elements like the help sidebar and message dialogs.
 */
const HelpManager = {
  /**
   * Displays the help sidebar, populating it with content relevant to the active sheet and mode.
   */
  showHelpSidebar: function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    const mode = PropertyManager.getProperty(PropertyKeys.MODE); // Ensure properties are loaded

    LoggerManager.logDebug(
      `Showing help sidebar for sheet: ${sheetName}, mode: ${mode}`
    );

    try {
      // Create the HTML template from the 'HelpContent' file
      const template = HtmlService.createTemplateFromFile("HelpContent");

      // Pass data to the template
      template.sheetName = sheetName;
      template.mode = mode;
      // Pass config objects directly to access their properties in HTML scriptlets
      template.MainSheetConfig = MainSheetConfig;
      template.HistorySheetConfig = HistorySheetConfig;
      template.GlobalConfig = GlobalConfig; // Pass global config too

      // Evaluate and display the sidebar
      const htmlOutput = template
        .evaluate()
        .setTitle("Moderate Habits Help")
        .setWidth(300); // Optional: Set a width
      SpreadsheetApp.getUi().showSidebar(htmlOutput);
    } catch (e) {
      LoggerManager.handleError(
        `Error creating or showing help sidebar: ${e.message}`,
        true
      );
      SpreadsheetApp.getUi().alert(
        "Could not display the help sidebar at this time."
      );
    }
  },
};
Object.freeze(HelpManager);

/**
 * Manages displaying alerts and confirmation dialogs to the user.
 */
const Messages = {
  /**
   * Predefined button types used for mapping responses.
   * @enum {string}
   */
  ButtonTypes: {
    OK: "OK",
    YES: "YES",
    NO: "NO",
    CLOSE: "CLOSE", // Used for dialog close 'X' button
  },

  /**
   * Returns the appropriate Apps Script ButtonSet enum based on a string key.
   * @private
   * @param {string} buttonsKey - e.g., "YES_NO", "OK".
   * @param {GoogleAppsScript.Base.Ui} ui - The UI instance.
   * @returns {GoogleAppsScript.Base.ButtonSet} The corresponding ButtonSet.
   */
  _getButtonSet: function (buttonsKey, ui) {
    switch (buttonsKey) {
      case "YES_NO":
        return ui.ButtonSet.YES_NO;
      case "OK_CANCEL": // Add if needed
        return ui.ButtonSet.OK_CANCEL;
      case "YES_NO_CANCEL": // Add if needed
        return ui.ButtonSet.YES_NO_CANCEL;
      case "OK":
      default:
        return ui.ButtonSet.OK;
    }
  },

  /**
   * Validates if a dialog response string matches one of the expected ButtonTypes.
   * @private
   * @param {string} response - The response string from ui.alert().
   * @returns {boolean} True if the response is valid.
   */
  _validateButtonResponse: function (response) {
    // Check if the response string is one of the values in our ButtonTypes enum
    return Object.values(this.ButtonTypes).includes(response);
  },

  /**
   * Displays an alert or confirmation dialog based on a message type key.
   * @param {MessageTypes} messageKey - The enum key for the message.
   * @returns {Messages.ButtonTypes} The user's response (e.g., Messages.ButtonTypes.YES).
   */
  showAlert: function (messageKey) {
    // Validate the key against the MessageTypes enum
    if (!Object.values(MessageTypes).includes(messageKey)) {
      LoggerManager.handleError(
        `Invalid message key passed to showAlert: ${messageKey}`,
        true
      );
      return this.ButtonTypes.CLOSE; // Return default on error
    }

    // Get the message configuration function associated with the key
    const messageConfigFn = this[`_get_${messageKey}_Config`];
    if (typeof messageConfigFn !== "function") {
      LoggerManager.handleError(
        `No config function found for message key: ${messageKey}`,
        true
      );
      return this.ButtonTypes.CLOSE;
    }

    const { title, body, buttons } = messageConfigFn.call(this); // Get title, body, buttonsKey
    const ui = SpreadsheetApp.getUi();
    const buttonSet = this._getButtonSet(buttons, ui);

    let response;
    try {
      response = ui.alert(title, body, buttonSet);
      // Convert Apps Script response (e.g., ui.Button.YES) to our standard string enum
      response = String(response).toUpperCase(); // Make it uppercase for consistency

      // Validate the response
      if (!this._validateButtonResponse(response)) {
        LoggerManager.logDebug(
          `Unrecognized dialog response: ${response}. Mapping to CLOSE.`
        );
        response = this.ButtonTypes.CLOSE;
      }
    } catch (e) {
      LoggerManager.handleError(
        `Error displaying alert for key ${messageKey}: ${e.message}`,
        false
      ); // Don't halt execution, return default
      response = this.ButtonTypes.CLOSE;
    }

    LoggerManager.logDebug(
      `Alert shown for ${messageKey}. User response: ${response}`
    );
    return response; // Return validated string from Messages.ButtonTypes
  },

  // --- Message Configuration Functions ---
  // Using a naming convention _get_KEY_Config

  _get_newVersionAvailable_Config: function () {
    return {
      title: "New Version Available",
      body:
        `There is now a new version of this library available.\n\n` +
        `To upgrade:\n` +
        `1. Go to Extensions > Apps Script.\n` +
        `2. In the editor, click 'Libraries' on the left sidebar.\n` +
        `3. Find 'ModerateHabits' (or similar name) in the list.\n` +
        `4. Select the highest version number from the dropdown.\n` +
        `5. Click 'Save'.`,
      buttons: "OK", // Corresponds to ui.ButtonSet.OK
    };
  },

  _get_noNewUpdates_Config: function () {
    return {
      title: "No New Updates",
      body: `You are using the latest version.\n\nThanks for checking!`,
      buttons: "OK",
    };
  },

  _get_terminationConfirmation_Config: function () {
    return {
      title: "Terminate Challenge?",
      body: `Are you sure you want to stop your current challenge?\n\nYour streaks and progress tracking will stop. Your history data will remain.\n\nClick "Yes" to confirm termination.`,
      buttons: "YES_NO", // Corresponds to ui.ButtonSet.YES_NO
    };
  },

  _get_terminated_Config: function () {
    return {
      title: "Challenge Terminated",
      body: "Your challenge tracking has been stopped. You can start a new challenge from the 'Moderate Habits Settings' menu.",
      buttons: "OK",
    };
  },

  _get_terminationCancelled_Config: function () {
    return {
      title: "Termination Canceled",
      body: "Challenge termination cancelled. Keep up the good work!",
      buttons: "OK",
    };
  },

  _get_terminationReminder_Config: function () {
    return {
      title: "Challenge Terminated",
      body: "Tracking is currently stopped. To resume, please start a new challenge using the 'Moderate Habits Settings' menu.",
      buttons: "OK",
    };
  },

  _get_invalidDate_Config: function () {
    return {
      title: "Invalid Date",
      body: `The selected date is invalid or outside the allowed range (between the challenge start date and today).\nPlease select a valid date. Defaulting to today's date.`,
      buttons: "OK",
    };
  },

  _get_invalidSetters_Config: function () {
    return {
      title: "Invalid Settings",
      body: `Please ensure the 'Reset Hour' (0-23) and 'Boost Interval' (1 or greater) have valid whole numbers before setting habits.`,
      buttons: "OK",
    };
  },

  _get_challengeReset_Config: function () {
    return {
      title: "Set Up Your Habits",
      body: "Challenge reset! Please define your habits by adding emojis in the 'activities' column.\n\nYou can also add non-emoji subheaders or notes.\n\nOnce ready, check the 'set habits' box on the right.",
      buttons: "OK",
    };
  },

  _get_challengeCancelled_Config: function () {
    return {
      title: "Challenge Reset Cancelled",
      body: "Starting a new challenge was cancelled.",
      buttons: "OK",
    };
  },

  _get_habitSpreadReset_Config: function () {
    return {
      title: "Habit Structure Changed",
      body: `It looks like the habits (emojis or their rows) in the 'activities' column were changed.\n\nChanges have been reverted to the original setup for consistency.\n\nTo change your habits, please use the "Start New Challenge" option in the menu.`,
      buttons: "OK",
    };
  },

  _get_confirmHabitSpread_Config: function () {
    return {
      title: "Confirm Habits?",
      body: "Finalize this habit setup and start the challenge?\n\nMake sure your emojis, reset hour, and boost interval are correct. This cannot be easily undone.",
      buttons: "YES_NO",
    };
  },

  _get_startNewChallenge_Config: function () {
    return {
      title: "Start New Challenge?",
      body: "Are you sure you want to start a new challenge?\n\nThis will reset your current streaks and buffers and require you to set up your habits again. Your past history data will be preserved.",
      buttons: "YES_NO",
    };
  },

  _get_welcomeMessage_Config: function () {
    return {
      title: "Welcome to Moderate Habits!",
      body:
        "Get ready to build consistent habits!\n\n" +
        "This tool helps you track daily activities using emojis, with integrated buffer days for flexibility.\n\n" +
        "You'll now be guided to set up your first challenge.\n\n" +
        "Tip: Use the help menu ('Moderate Habits Settings' > 'Show Help') for detailed instructions anytime.",
      buttons: "OK",
    };
  },

  _get_noHabitsSet_Config: function () {
    return {
      title: "No Habits Defined",
      body: "Please define at least one habit (using an emoji) in the 'activities' column before confirming.",
      buttons: "OK",
    };
  },

  _get_undefinedCellChanges_Config: function () {
    return {
      title: "Protected Cell Edit",
      body:
        "The cell(s) you tried to edit are protected or calculated automatically.\n\n" +
        "Your change has been reverted. Use Control/Cmd+Z if you need to undo previous actions.",
      buttons: "OK",
    };
  },

  _get_dataParseError_Config: function () {
    return {
      title: "Data Reading Error",
      body:
        "There was an issue reading some stored data from the history sheet.\n\n" +
        "Default values have been used instead. This might affect streak or buffer calculations temporarily.\n\n" +
        "Check the execution logs for details if the problem persists.",
      buttons: "OK",
    };
  },
};
Object.freeze(Messages); // Freeze the main object
// Individual config functions are properties and thus frozen.
