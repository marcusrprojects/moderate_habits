/**
 * @fileoverview Contains high-level event handlers (onOpen, onEdit, time-driven)
 * and functions triggered directly by menu items. It orchestrates calls
 * to other modules like DataHandler, HabitManager, PropertyManager, etc.
 */

/** OnlyCurrentDoc */

/**
 * Runs when the spreadsheet is opened. Adds the custom menu.
 * Handles basic setup like property loading on subsequent opens.
 * Relies on user clicking 'Begin' for first-time setup and trigger creation.
 * @param {object} e The event object (unused but standard for onOpen).
 */
function onOpen(e) {
  // Add custom menu
  SpreadsheetApp.getUi()
    .createMenu("Moderate Habits Settings")
    .addItem("Show Help (Current Page)", "showHelpSidebar")
    .addSeparator()
    .addItem("Start New Challenge", "beginOrStartNewChallenge") // Combined menu item
    .addItem("Terminate Challenge", "terminateChallenge")
    .addSeparator()
    .addItem("Check for Updates", "versionCheck")
    .addToUi();

  // On subsequent opens (not first run), just load properties if needed.
  // Trigger creation/check is handled by 'Begin' or manually.
  if (!isFirstRun()) {
    LoggerManager.logDebug("Not first run - loading properties.");
    // Accessing a property ensures PropertyManager loads if not already loaded.
    PropertyManager.getProperty(PropertyKeys.MODE); // Example access
  } else {
    LoggerManager.logDebug(
      "First run detected by onOpen. User should click 'Begin / Start New Challenge'."
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
  const firstDate = PropertiesService.getDocumentProperties().getProperty(
    PropertyKeys.FIRST_CHALLENGE_DATE
  );
  LoggerManager.logDebug(
    `isFirstRun check: Property "${PropertyKeys.FIRST_CHALLENGE_DATE}" value is "${firstDate}"`
  );
  return !firstDate;
}

/**
 * Function mapped to the menu item "**Begin** / Start New Challenge".
 * Handles both the very first setup and resetting for a new challenge.
 */
function beginOrStartNewChallenge() {
  if (isFirstRun()) {
    // --- First Time Setup Flow ---
    LoggerManager.logDebug("Executing first-time setup via Begin menu item...");
    try {
      // 1. Prompt for Initial Authorization (Spreadsheet, UI, etc.)
      //    This happens implicitly when accessing services like SpreadsheetApp or PropertyManager
      //    if not already granted. We'll access one to ensure the prompt appears *before* trigger creation attempt.
      LoggerManager.logDebug(
        "Accessing spreadsheet service to potentially trigger initial auth..."
      );
      SpreadsheetApp.getActiveSpreadsheet().getName(); // Simple access

      // 2. Show Welcome Message (Requires UI Auth)
      Messages.showAlert(MessageTypes.WELCOME_MESSAGE);

      // 3. Initialize Properties (getProperty calls _initializeDefaultProperty internally)
      LoggerManager.logDebug("Initializing default properties...");
      PropertyManager.getProperty(PropertyKeys.MODE);
      PropertyManager.getProperty(PropertyKeys.FIRST_CHALLENGE_DATE); // Sets date/row
      PropertyManager.getProperty(PropertyKeys.RESET_HOUR);
      PropertyManager.getProperty(PropertyKeys.BOOST_INTERVAL);
      PropertyManager.getProperty(PropertyKeys.EMOJI_LIST);
      PropertyManager.getProperty(PropertyKeys.LAST_DATE_SELECTOR_UPDATE);
      PropertyManager.getProperty(PropertyKeys.LAST_COMPLETION_UPDATE);
      PropertyManager.getProperty(PropertyKeys.LAST_UPDATE);
      PropertyManager.setDocumentProperties(); // Save initialized properties

      // 4. Attempt to Create Trigger (Requires script.scriptapp Auth)
      //    THIS is where the re-auth prompt for the *new* scope should appear if manifest is correct.
      LoggerManager.logDebug(
        "Attempting to create trigger (may require re-authorization)..."
      );
      TriggerManager.createTrigger(); // User MUST grant script.scriptapp here

      // 5. Initialize UI for Habit Setup
      LoggerManager.logDebug("Initializing Habit Ideation UI...");
      HabitManager.initializeSetHabitUI();
      Messages.showAlert(MessageTypes.CHALLENGE_RESET); // Inform user

      LoggerManager.logDebug("First-time setup completed successfully.");
    } catch (error) {
      LoggerManager.handleError(
        `Error during first-time setup (beginOrStartNewChallenge): ${error.message}\n${error.stack}`,
        true
      );
    }
  } else {
    // --- Start New Challenge (Reset) Flow ---
    LoggerManager.logDebug("Executing Start New Challenge (Reset)...");
    const response = Messages.showAlert(MessageTypes.START_NEW_CHALLENGE);
    if (response === Messages.ButtonTypes.YES) {
      LoggerManager.logDebug("User confirmed starting new challenge.");
      // Re-initialize properties and UI for habit setup
      HabitManager.initializeSetHabitUI(); // This resets UI, sets mode to Ideation
      // Trigger should already exist, but let's ensure it's updated if resetHour changed during previous setup
      // initializeSetHabitUI doesn't create trigger, _finalizeHabitSpread does.
      Messages.showAlert(MessageTypes.CHALLENGE_RESET); // Inform user
      PropertyManager.setDocumentProperties(); // Save mode change etc.
    } else {
      LoggerManager.logDebug("User cancelled starting new challenge.");
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
