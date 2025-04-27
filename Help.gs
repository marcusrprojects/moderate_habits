/**
 * Adds a custom menu to the Google Sheets UI upon opening the spreadsheet.
 *
 * This function creates a new menu called "One Month Moderate Settings" with a single item
 * labeled "Show Help". Selecting this item will trigger the `showHelpSidebar` function,
 * which displays the help sidebar.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Moderate Habits Settings")
    .addItem("Show Help (Current Page)", "showHelpSidebar")
    .addItem("Start New Challenge", "startNewChallenge")
    .addItem("Terminate Challenge", "terminateChallenge")
    .addToUi();
}

/**
 * Displays a help sidebar in the Google Sheets UI.
 *
 * This function retrieves the name of the active sheet and passes it to an HTML template
 * (`HelpContent`). The sidebar is then displayed with the relevant content for the
 * current sheet.
 *
 * The help content is dynamically populated based on the current sheet, making it context-sensitive.
 */
function showHelpSidebar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  // Create the HTML template from the 'HelpContent' file
  const template = HtmlService.createTemplateFromFile("HelpContent");
  template.sheetName = sheetName;

  // Get the mode property
  const mode = PropertyManager.getProperty(PropertyKeys.MODE);
  template.mode = mode;

  // Evaluate and display the sidebar
  const htmlOutput = template.evaluate().setTitle("Help");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Handles the logic for starting a new challenge.
 *
 * This function presents a confirmation dialog to the user before resetting the entire challenge.
 * If the user confirms, it initializes the UI for setting new habits and alerts the user to define
 * their habit spread.
 */
function startNewChallenge() {
  const response = Messages.showAlert(MessageTypes.START_NEW_CHALLENGE);

  if (response == Messages.ButtonTypes.YES) {
    HabitManager.initializeSetHabitUI(); // Initialize the habit setting UI
    Messages.showAlert(MessageTypes.CHALLENGE_RESET);
    PropertyManager.setDocumentProperties();
  } else {
    Messages.showAlert(MessageTypes.CHALLENGE_CANCELLED);
  }
}

/**
 * Terminates the current challenge and updates the app state.
 *
 * This function handles the termination process of an ongoing challenge.
 * It asks the user for confirmation and, if confirmed, hides the tracking data,
 * resets the necessary sheet data, and changes the app mode to "terminated."
 * If the termination is canceled, a corresponding message is displayed.
 *
 * @function terminateChallenge
 */
function terminateChallenge() {
  const response = Messages.showAlert(MessageTypes.TERMINATION_CONFIRMATION);
  if (response === Messages.ButtonTypes.YES) {
    // Hide and reset tracking data
    MainSheetConfig.toggleColumns(ColumnAction.HIDE);
    MainSheetConfig.resetData();

    // Set the property to indicate we are in the terminated mode
    PropertyManager.setProperty(
      PropertyKeys.MODE,
      ModeTypes.TERMINATED,
      (forceSet = true)
    );
    Messages.showAlert(MessageTypes.TERMINATED);
  } else {
    Messages.showAlert(MessageTypes.TERMINATION_CANCELLED);
  }
}

/**
 * Configuration for displaying messages and alerts in the app.
 *
 * The `Messages` object defines the types of buttons used in alerts,
 * handles displaying alerts to the user, and provides custom messages
 * for various actions within the app. It ensures consistent user
 * interaction for important notifications and confirmations.
 *
 * @type {Object}
 */
const Messages = {
  /**
   * Predefined button types used for displaying alert dialogs in the application.
   *
   * The `ButtonTypes` object defines various button sets that can be used in
   * different types of alerts or prompts shown to the user. These types
   * represent common UI buttons that a user can interact with in confirmation
   * or information dialogs.
   *
   * @type {Object}
   */
  ButtonTypes: {
    OK: "OK",
    YES_NO: "YES_NO",
    YES: "YES",
    NO: "NO",
    CLOSE: "CLOSE",
  },

  /**
   * Displays an alert with a custom message based on the provided key.
   *
   * This method dynamically generates the alert's title, body, and buttons
   * based on the message key passed. It validates if the message key exists
   * and translates it to a UI alert for the user.
   *
   * @param {string} messageKey - The key that corresponds to a specific message.
   * @returns {string} - The user's response to the alert, such as "OK" or "YES."
   */
  showAlert: function (messageKey) {
    // Validate if the messageKey exists in MessageTypes
    if (!Object.values(MessageTypes).includes(messageKey)) {
      LoggerManager.handleError(
        `Invalid message key: ${messageKey}. It must be one of the following: ${Object.values(
          MessageTypes
        ).join(", ")}`
      );
    }

    const ui = SpreadsheetApp.getUi();
    const { title, body, buttons } = Messages[messageKey]();
    const buttonSet = Messages.getButtonSet(buttons, ui);

    let response = String(ui.alert(title, body, buttonSet));

    if (!this._validateButtonResponse(response, false)) {
      response = this.ButtonTypes.CLOSE;
      LoggerManager.logDebug(`Setting button ${response} to ButtonTypes.CLOSE`);
    }
    return response;
  },

  /**
   * Returns the appropriate button set based on the buttons type.
   *
   * @param {string} buttons - The button type (OK, YES_NO, etc.).
   * @param {Object} ui - The SpreadsheetApp UI instance.
   * @returns {Object} - The corresponding button set from the UI.
   */
  getButtonSet: function (buttons, ui) {
    if (buttons === this.ButtonTypes.YES_NO) {
      return ui.ButtonSet.YES_NO;
    } else {
      return ui.ButtonSet.OK; // Default to OK if no match
    }
  },

  /**
   * Validates the button response against the known button types.
   *
   * @param {string} response - The user's response.
   * @param {boolean} throwError - Whether to throw an error for invalid responses.
   * @returns {boolean} - True if the response is valid, false otherwise.
   */
  _validateButtonResponse: function (response, throwError) {
    // Check if the response is a recognized ButtonType
    if (!this.ButtonTypes[response]) {
      LoggerManager.handleError(
        `No matching button type for response: ${response}.`,
        throwError
      );
      return false;
    }

    return true;
  },

  [MessageTypes.NEW_VERSION_AVAILABLE]: function () {
    return {
      title: "New Version Available",
      body:
        `There is now a new version of this library. Here are the steps for upgrading:\n\n` +
        `1. In the menu, go to Extensions -> Apps Script.\n` +
        `2. Click on the Editor on the lefthand side (symbol: < >).\n` +
        `3. You should see "moderate habits" underneath "Libraries." Click that.\n` +
        `4. Select the newest version (the largest number).\n`,
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.NO_NEW_UPDATES]: function () {
    return {
      title: "No New Updates",
      body: `You're all caught up. \n\n` + `Thanks for checking!`,
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.TERMINATION_CONFIRMATION]: function () {
    return {
      title: "Stop Current Challenge",
      body: `Are you sure you want to stop your current challenge? Your streaks and progress will be reset. Click "Yes" if you will no longer be using this spreadsheet for the forseeable future.`,
      buttons: this.ButtonTypes.YES_NO,
    };
  },

  [MessageTypes.TERMINATED]: function () {
    return {
      title: "Challenge Terminated",
      body: "Your challenge has been stopped. You can start a new challenge at any time.",
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.TERMINATION_CANCELLED]: function () {
    return {
      title: "Challenge Termination Canceled",
      body: "Challenge will not terminate. Keep going!",
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.TERMINATION_REMINDER]: function () {
    return {
      title: "Challenge Termination Reminder",
      body: "Your previous challenge has been terminated. You must start a new challenge, using the settings menu, to proceed.",
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.INVALID_DATE]: function () {
    return {
      title: "Invalid Date Selected",
      body: `Must choose a valid date. Defaulting to today's date.`,
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.INVALID_SETTERS]: function () {
    return {
      title: "Invalid Setters",
      body: `Must choose a valid reset hour (integers >= 0) and boost interval (integers >= 1). Defaulting to today's date.`,
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.CHALLENGE_RESET]: function () {
    return {
      title: "Challenge Reset",
      body: "The challenge has been reset. Please define your habit spread by selecting your daily emoji spread. This is the time to make any subheaders and format the activities column as you would like.",
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.CHALLENGE_CANCELLED]: function () {
    return {
      title: "Challenge Canceled",
      body: "Challenge reset canceled.",
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.HABIT_SPREAD_RESET]: function () {
    return {
      title: "Habit Spread Reset",
      body: `The habit spread does not match the original setup. The emoji cells will be reset to their original state. To change your habit spread, please use the "Start New Challenge" option in the menu.`,
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.CONFIRM_HABIT_SPREAD]: function () {
    return {
      title: "Confirm Habit Spread",
      body: "Are you sure you want to proceed with this habit spread? This action is irreversible.",
      buttons: this.ButtonTypes.YES_NO,
    };
  },

  [MessageTypes.START_NEW_CHALLENGE]: function () {
    return {
      title: "Start New Challenge",
      body: "Are you sure you want to reset everything and start a new challenge? This action is irreversible.",
      buttons: this.ButtonTypes.YES_NO,
    };
  },

  [MessageTypes.WELCOME_MESSAGE]: function () {
    return {
      title: "Welcome to Moderate Habits!",
      body:
        "Thanks for using this tool!\n\n" +
        "It provides a simple way to structure your habits, with each habit represented by an emoji (repeated emojis are allowed!).\n" +
        "You will be redirected to start your first habit challenge after this dialog. Enjoy!\n\n" +
        'Tip: For guidance on how to get the most out of this tool, check out the help menu under the "Moderate Habits" dropdown for detailed instructions.',
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.NO_HABITS_SET]: function () {
    return {
      title: "No Habits Set",
      body: "Please set at least one habit (emoji) before proceeding.",
      buttons: this.ButtonTypes.OK,
    };
  },

  [MessageTypes.UNDEFINED_CELL_CHANGES]: function () {
    return {
      title: "Undefined Changes to Cells",
      body:
        "Woah there! Ensure that the change you just made to those cells was intentional. Otherwise, you may need to undo them via Control/Cmd + Z!\n" +
        "Attempting to recover the data...",
      buttons: this.ButtonTypes.OK,
    };
  },
};

/**
 * Manages the setup and configuration of habits within the application.
 *
 * The HabitManager handles various tasks related to habit ideation, validation, and finalization,
 * ensuring that the user experience is smooth and aligned with the overall application structure.
 * It provides methods to toggle UI elements, validate user inputs, and finalize habit configurations.
 */
const HabitManager = {
  /**
   * Initializes the UI for habit ideation mode by hiding unnecessary columns and adding a checkbox for habit setup.
   *
   * Sets the property 'mode' to 'habitIdeation' to indicate ideation mode.
   */
  initializeSetHabitUI: function () {
    // Hide the completion, buffer days, and streaks columns
    MainSheetConfig.toggleColumns(ColumnAction.HIDE);
    MainSheetConfig.resetData();
    this.toggleHabitFields(CellAction.SET);
    SpreadsheetApp.flush();

    // Set the property to indicate we are in habit ideation mode
    PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.HABIT_IDEATION);

    const historySheet = HistorySheetConfig.getSheet();
    const todayDateStr = DateManager.getTodayStr();
    const lastDateStr = HistorySheetConfig.getLastDateStr();
    LoggerManager.logDebug(
      `comparing today's date ${todayDateStr} and last date ${lastDateStr}.`
    );

    if (lastDateStr && lastDateStr === todayDateStr) {
      LoggerManager.logDebug(`Should be clearing last row...`);
      historySheet
        .getRange(historySheet.getLastRow(), 1, 1, historySheet.getLastColumn())
        .clearContent(); // Clears the entire row's content
    }
  },

  /**
   * Validates the reset hour value from the main sheet.
   *
   * This function checks whether the reset hour is a non-negative integer.
   * If the validation fails, it logs a debug message.
   *
   * @function __validateResetHour
   * @returns {boolean} - True if the reset hour is valid, false otherwise.
   */
  __validateResetHour: function () {
    const resetHourValue = MainSheetConfig.getSheet()
      .getRange(MainSheetConfig.setterRanges.resetHour)
      .getValue();

    // Check if it's a non-negative number
    if (UtilsManager.__validateNonNegativeInteger(resetHourValue)) {
      return true;
    } else {
      LoggerManager.logDebug(`Invalid reset hour: ${resetHourValue}.`);
      return false;
    }
  },

  /**
   * Validates the boost interval value from the main sheet.
   *
   * This function checks whether the boost interval is a non-negative integer
   * and greater than 0. If the validation fails, it logs a debug message.
   *
   * @function __validateBoostInterval
   * @returns {boolean} - True if the boost interval is valid, false otherwise.
   */
  __validateBoostInterval: function () {
    const boostIntervalValue = MainSheetConfig.getSheet()
      .getRange(MainSheetConfig.setterRanges.boostInterval)
      .getValue();

    // Check if it's a non-negative number
    if (
      UtilsManager.__validateNonNegativeInteger(boostIntervalValue) &&
      boostIntervalValue > 0
    ) {
      return true;
    } else {
      LoggerManager.logDebug(`Invalid boost interval: ${boostIntervalValue}.`);
      return false;
    }
  },

  /**
   * Validates both the reset hour and boost interval values.
   *
   * This method ensures that both the reset hour and boost interval are valid.
   * If either validation fails, it logs an error message and returns false.
   *
   * @function _validatingSetters
   * @param {boolean} throwError - Whether to throw an error if validation fails.
   * @returns {boolean} - True if both values are valid, false otherwise.
   */
  _validatingSetters: function (throwError) {
    if (!this.__validateResetHour() || !this.__validateBoostInterval()) {
      LoggerManager.handleError(
        "Invalid resetHour or boostInterval.",
        throwError
      );
      return false;
    }
    return true;
  },

  /**
   * Toggles and validates the display and value setting of the reset hour and boost interval fields.
   *
   * @param {string} action - Either 'CLEAR' to remove fields or 'SET' to set fields with validation.
   * @returns {boolean|null} - Returns true if successful, false if validation fails, or null if action is unknown.
   */
  togglePropertySetters: function (action) {
    if (action !== CellAction.CLEAR && action !== CellAction.SET) {
      LoggerManager.handleError(
        `togglePropertySetters called with unknown action: ${action}.`
      );
      return null;
    }

    const resetHourCell = MainSheetConfig.setterRanges.resetHour;
    const boostIntervalCell = MainSheetConfig.setterRanges.boostInterval;

    if (action === CellAction.CLEAR) {
      PropertyManager.setProperty(
        PropertyKeys.RESET_HOUR,
        MainSheetConfig.getSheetValue(resetHourCell)
      );
      PropertyManager.setProperty(
        PropertyKeys.BOOST_INTERVAL,
        MainSheetConfig.getSheetValue(boostIntervalCell)
      );
    } else {
      MainSheetConfig.setSheetValue(
        resetHourCell,
        PropertyManager.getProperty(PropertyKeys.RESET_HOUR)
      );
      MainSheetConfig.setSheetValue(
        boostIntervalCell,
        PropertyManager.getProperty(PropertyKeys.BOOST_INTERVAL)
      );
    }

    return true; // Return true if all validations passed and action completed
  },

  /**
   * Handles the logic when the habit spread checkbox is checked. It is called from the Handlers script.
   *
   * This function checks if the checkbox is checked, then confirms with the user if they want
   * to proceed with the current habit spread. If confirmed, it finalizes the habit spread;
   * otherwise, it alerts the user if no habits (emojis) have been set and resets the checkbox.
   */
  setHabitSpread: function () {
    const sheet = MainSheetConfig.getSheet();
    const checkboxCell = sheet.getRange(MainSheetConfig.setterRanges.setHabit);

    // Check if the checkbox was checked
    if (checkboxCell.getValue() === true) {
      const response = Messages.showAlert(MessageTypes.CONFIRM_HABIT_SPREAD);
      const setterValidation = this._validatingSetters(false);

      // Get the current emoji spread
      const currentEmojiSpread = MainSheetConfig.getCurrentEmojiSpread();
      if (
        response == Messages.ButtonTypes.YES &&
        currentEmojiSpread.length > 0 &&
        setterValidation
      ) {
        this.finalizeHabitSpread(); // Finalize the habit spread
        return;
      }

      if (currentEmojiSpread.length === 0) {
        Messages.showAlert(MessageTypes.NO_HABITS_SET);
      }

      if (!setterValidation) {
        Messages.showAlert(MessageTypes.INVALID_SETTERS);
      }

      checkboxCell.setValue(false); // Reset the checkbox to unchecked
      SpreadsheetApp.flush();
    }
  },

  /**
   * Toggles the display of the habit label and checkbox cells based on the provided action.
   *
   * @param {string} action - remove the habit label and checkbox, or add them.
   */
  toggleHabitFields: function (action) {
    if (action !== CellAction.CLEAR && action !== CellAction.SET) {
      LoggerManager.handleError(
        `toggleHabitFields called with unknown action: ${action}.`
      );
      return;
    }

    this.togglePropertySetters(action);

    // Define the ranges and labels for batch processing
    const setterFields = {
      labels: MainSheetConfig.setterLabelRanges,
      cells: MainSheetConfig.setterRanges,
      labelsText: MainSheetConfig.setterLabels,
      notes: MainSheetConfig.setterNotes,
    };

    const sheet = MainSheetConfig.getSheet();

    // Special case for 'setHabit' checkbox
    const setHabitCellRange = sheet.getRange(setterFields.cells.setHabit);
    action === CellAction.CLEAR
      ? setHabitCellRange.clearDataValidations()
      : setHabitCellRange.insertCheckboxes();

    // Loop over all fields and perform the appropriate action
    for (const key in setterFields.labels) {
      const labelRange = sheet.getRange(setterFields.labels[key]);
      const cellRange = sheet.getRange(setterFields.cells[key]);

      if (action === CellAction.CLEAR) {
        // Clear the label and cell
        labelRange.clearContent();
        labelRange.setBackground(SheetConfig.mainColor);
        cellRange.clearNote();
        cellRange.clearContent();
      } else if (action === CellAction.SET) {
        const labelText = setterFields.labelsText[key];
        const noteText = setterFields.notes[key];

        // Set the label, insert the checkbox for habits, and set the note
        labelRange.setBackground(SheetConfig.secondaryColor);
        labelRange.setFontWeight("bold");
        LoggerManager.logDebug(`setting font weight of ${labelRange} to bold`);
        labelRange.setValue(labelText);
        cellRange.setNote(noteText);
      }
    }
  },

  /**
   * Finalizes the habit spread after the user confirms their selection.
   *
   * Clears the habit label and checkbox, shows the hidden columns, and starts the challenge.
   */
  finalizeHabitSpread: function () {
    const sheet = MainSheetConfig.getSheet();

    // Remove habit ideation mode
    PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.CHALLENGE);

    PropertyManager.updateEmojiSpread();
    PropertyManager.updateFirstChallengeDateAndRow();

    this.toggleHabitFields(CellAction.CLEAR);
    MainSheetConfig.toggleColumns(ColumnAction.SHOW);
    SpreadsheetApp.flush();

    const startRow = MainSheetConfig.firstDataInputRow;
    const lastRow = sheet.getLastRow(); // Get the last row of data
    const activitiesRange = sheet.getRange(
      startRow,
      MainSheetConfig.activityDataColumn,
      lastRow - startRow + 1,
      1
    );
    activitiesRange.setBackground(SheetConfig.mainColor);

    // Re-add checkboxes in the completion column for relevant rows
    MainSheetConfig.insertCompletionCheckboxes();

    PropertyManager.setProperty(
      PropertyKeys.LAST_UPDATE,
      LastUpdateTypes.COMPLETION
    );
    MainSheetConfig.displayDate(DateManager.getTodayStr());

    LoggerManager.logDebug(`Habit spread finalized and challenge started.`);
  },
};
