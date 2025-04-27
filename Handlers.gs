/**
 * @fileoverview Contains high-level event handlers (onOpen, onEdit, time-driven)
 * and functions triggered directly by menu items. It orchestrates calls
 * to other modules like DataHandler, HabitManager, PropertyManager, etc.
 */

/** OnlyCurrentDoc */

/**
 * Runs when the spreadsheet is opened. Adds the custom menu.
 * Handles basic setup like property loading on subsequent opens.
 * Relies on user clicking 'Start New Challenge' for first-time setup and initial trigger creation.
 * @param {object} e The event object (unused but standard for onOpen).
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Moderate Habits Settings")
    .addItem("Show Help (Current Page)", "showHelpSidebar")
    .addSeparator()
    .addItem("Start New Challenge", "startNewChallenge")
    .addItem("Setup Daily Reset Trigger", "setupTriggerManually")
    .addItem("Terminate Challenge", "terminateChallenge")
    .addSeparator()
    .addItem("Check for Updates", "versionCheck")
    .addToUi();

  if (!isFirstRun()) {
    LoggerManager.logDebug("Not first run - loading properties.");
    PropertyManager.getProperty(PropertyKeys.MODE); // Load props lazily on access
  } else {
    LoggerManager.logDebug(
      "First run detected by onOpen. User should click 'Start New Challenge'."
    );
  }
}

/**
 * Checks if this is the first time the script is being run for this document
 * by checking for the existence of a specific property.
 * Uses PropertiesService directly to avoid initialization side effects during check.
 * @returns {boolean} True if it's the first run, false otherwise.
 */
function isFirstRun() {
  // Check directly to avoid initializing defaults just for this check
  const firstDate = PropertiesService.getDocumentProperties().getProperty(
    PropertyKeys.FIRST_CHALLENGE_DATE
  );
  LoggerManager.logDebug(
    `isFirstRun check: Property "${PropertyKeys.FIRST_CHALLENGE_DATE}" value is "${firstDate}"`
  );
  return !firstDate;
}

/**
 * Function mapped to the menu item "Start New Challenge".
 * Handles both the very first setup and resetting for subsequent new challenges.
 * Attempts initial trigger creation during the first run.
 */
function startNewChallenge() {
  if (isFirstRun()) {
    // --- First Time Setup Flow ---
    LoggerManager.logDebug(
      "Executing first-time setup via Start New Challenge menu item..."
    );
    try {
      // ... (Steps 1-3: Auth prompt simulation, Welcome msg, Init Props) ...
      LoggerManager.logDebug("Accessing spreadsheet service...");
      SpreadsheetApp.getActiveSpreadsheet().getName();
      Messages.showAlert(MessageTypes.WELCOME_MESSAGE);
      LoggerManager.logDebug("Initializing default properties...");
      PropertyManager.getProperty(PropertyKeys.MODE);
      PropertyManager.getProperty(PropertyKeys.FIRST_CHALLENGE_DATE);
      PropertyManager.getProperty(PropertyKeys.RESET_HOUR);
      PropertyManager.getProperty(PropertyKeys.BOOST_INTERVAL);
      PropertyManager.getProperty(PropertyKeys.EMOJI_LIST);
      PropertyManager.getProperty(PropertyKeys.LAST_DATE_SELECTOR_UPDATE);
      PropertyManager.getProperty(PropertyKeys.LAST_COMPLETION_UPDATE);
      PropertyManager.getProperty(PropertyKeys.LAST_UPDATE);
      PropertyManager.setDocumentProperties();

      // 4. Attempt to Create Trigger (Requires script.scriptapp Auth)
      LoggerManager.logDebug(
        "Attempting initial trigger creation (may require re-authorization)..."
      );
      const triggerCreated = TriggerManager.createTrigger(); // Call it here!
      if (!triggerCreated) {
        // Alert the user that trigger setup failed and they might need to do it manually
        SpreadsheetApp.getUi().alert(
          "Initial setup complete, but failed to create the daily reset trigger. Please ensure permissions were fully granted, and use 'Moderate Habits Settings > Setup Daily Reset Trigger' menu item to try again if needed."
        );
      } else {
        LoggerManager.logDebug("Initial trigger creation successful.");
      }

      // 5. Initialize UI for Habit Setup
      LoggerManager.logDebug("Initializing Habit Ideation UI...");
      HabitManager.initializeSetHabitUI(); // Sets mode to Ideation
      Messages.showAlert(MessageTypes.CHALLENGE_RESET); // Inform user

      LoggerManager.logDebug("First-time setup process finished.");
    } catch (error) {
      LoggerManager.handleError(
        `Error during first-time setup (startNewChallenge): ${error.message}\n${error.stack}`,
        true
      );
    }
  } else {
    // --- Start New Challenge (Reset) Flow ---
    LoggerManager.logDebug("Executing Start New Challenge (Reset)...");
    const response = Messages.showAlert(MessageTypes.START_NEW_CHALLENGE); // Confirmation dialog
    if (response === Messages.ButtonTypes.YES) {
      LoggerManager.logDebug("User confirmed starting new challenge (reset).");
      // Re-initialize UI for habit setup
      HabitManager.initializeSetHabitUI(); // Resets UI, sets mode to Ideation
      // Trigger is NOT automatically created/updated here. User uses manual option.
      Messages.showAlert(MessageTypes.CHALLENGE_RESET); // Inform user
      PropertyManager.setDocumentProperties(); // Save mode change etc.
    } else {
      LoggerManager.logDebug("User cancelled starting new challenge (reset).");
      Messages.showAlert(MessageTypes.CHALLENGE_CANCELLED);
    }
  }
}

/**
 * Handles the 'onEdit' event trigger fired by Google Sheets.
 * Delegates the processing to DataHandler.handleCellEdit.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object.
 */
function onEdit(e) {
  // Log entry point with details
  const rangeNotation = e && e.range ? e.range.getA1Notation() : "N/A";
  const sheetName = e && e.source ? e.source.getActiveSheet().getName() : "N/A";
  LoggerManager.logDebug(
    `onEdit Triggered: Sheet='${sheetName}', Range=${rangeNotation}, OldValue='${e.oldValue}', Value='${e.value}'`
  );

  if (e && e.range) {
    // Pass necessary info from the event object
    DataHandler.handleCellEdit(e.source.getActiveSheet(), e.range, e.oldValue);
  } else {
    LoggerManager.logDebug("onEdit event object or range was missing.");
  }
}

/**
 * Time-driven trigger function. Renews the checklist for the current day.
 * Simulates an edit on the date cell to trigger standard update logic.
 */
function renewChecklistForToday() {
  // Added try...catch around the whole function for robustness
  try {
    LoggerManager.logDebug(
      `Time-driven trigger: renewChecklistForToday running.`
    );
    const sheet = MainSheetConfig._getSheet(); // Use internal getter (throws if sheet missing)

    const range = sheet.getRange(MainSheetConfig.dateCell);
    const oldValue = range.getValue(); // Get the actual old value

    // Update the date cell visually first
    const todayStr = DateManager.getTodayStr();
    LoggerManager.logDebug(
      `Setting date cell ${MainSheetConfig.dateCell} to today: ${todayStr}`
    );
    range.setValue(todayStr); // Directly set the value

    // Call handleCellEdit to process the change
    LoggerManager.logDebug(
      `Simulating edit for date update. Old value was: ${oldValue}`
    );
    DataHandler.handleCellEdit(sheet, range, oldValue);

    LoggerManager.logDebug(`renewChecklistForToday finished successfully.`);
  } catch (error) {
    LoggerManager.handleError(
      `Error during time-driven trigger renewChecklistForToday: ${error.message}\n${error.stack}`,
      false
    ); // Log but don't stop trigger potentially
  }
}

/**
 * Menu item function: Checks for new library versions.
 */
function versionCheck() {
  try {
    const latestVersion = LibraryManager.fetchVersionInfo();
    if (latestVersion && latestVersion > LibraryManager.LATEST_VERSION) {
      // Simple comparison
      Messages.showAlert(MessageTypes.NEW_VERSION_AVAILABLE);
    } else if (latestVersion) {
      Messages.showAlert(MessageTypes.NO_NEW_UPDATES);
    } else {
      SpreadsheetApp.getUi().alert("Could not check for updates at this time.");
    }
  } catch (e) {
    LoggerManager.handleError(
      `Error during version check: ${e.message}`,
      false
    );
    SpreadsheetApp.getUi().alert(
      `An error occurred while checking for updates: ${e.message}`
    );
  }
}

/**
 * Menu item function: Sets up or re-creates the daily reset trigger manually.
 */
function setupTriggerManually() {
  try {
    LoggerManager.logDebug("Manual trigger setup requested.");
    // Attempt creation - this should force auth prompt if needed and not granted
    const success = TriggerManager.createTrigger(); // Will throw error if permissions missing
    if (success) {
      SpreadsheetApp.getUi().alert("Daily reset trigger setup successful.");
    }
    // If createTrigger failed due to non-permission error handled internally,
    // an alert might be needed here, but typically it throws on critical failure.
    // else { SpreadsheetApp.getUi().alert("Failed to setup trigger. Check logs."); }
  } catch (e) {
    // Catch potential permission error re-thrown from createTrigger
    LoggerManager.handleError(
      `Manual trigger setup failed: ${e.message}`,
      false
    ); // Log error from menu
    // Alert is handled by LoggerManager if throwError=true was used, otherwise:
    // SpreadsheetApp.getUi().alert("Failed to setup trigger. Please ensure permissions are granted and check logs.");
  }
}

/**
 * Menu item function: Terminates the current challenge.
 */
function terminateChallenge() {
  const response = Messages.showAlert(MessageTypes.TERMINATION_CONFIRMATION);
  if (response === Messages.ButtonTypes.YES) {
    LoggerManager.logDebug("User confirmed terminating challenge.");
    MainSheetConfig.toggleColumns(ColumnAction.HIDE);
    MainSheetConfig.resetChallengeDataUI();
    PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.TERMINATED, true); // forceSet = true
    Messages.showAlert(MessageTypes.TERMINATED);
  } else {
    LoggerManager.logDebug("User cancelled terminating challenge.");
    Messages.showAlert(MessageTypes.TERMINATION_CANCELLED);
  }
}

/**
 * Menu item function: Shows the help sidebar.
 */
function showHelpSidebar() {
  HelpManager.showHelpSidebar();
}
