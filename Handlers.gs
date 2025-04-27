/**
 * @fileoverview Contains high-level event handlers (onOpen, onEdit, time-driven)
 * and functions triggered directly by menu items. It orchestrates calls
 * to other modules like DataHandler, HabitManager, PropertyManager, etc.
 */

/** OnlyCurrentDoc */

/**
 * Runs when the spreadsheet is opened. Adds the custom menu.
 * Handles first-run initialization.
 */
function onOpen(e) {
  // Add event object parameter
  // Add custom menu
  SpreadsheetApp.getUi()
    .createMenu("Moderate Habits Settings")
    .addItem("Show Help (Current Page)", "showHelpSidebar")
    .addSeparator()
    .addItem("Start New Challenge", "startNewChallenge")
    .addItem("Terminate Challenge", "terminateChallenge")
    .addSeparator()
    .addItem("Check for Updates", "versionCheck")
    .addToUi();

  // Check if this is the first time the script is run in this document
  if (isFirstRun()) {
    LoggerManager.logDebug("First run detected.");
    // Use PropertiesService to check to avoid infinite loop if begin fails midway
    if (
      !PropertiesService.getDocumentProperties().getProperty(
        PropertyKeys.FIRST_CHALLENGE_DATE
      )
    ) {
      // Delay slightly to allow menu to potentially render
      Utilities.sleep(1500);
      // Use invokeFunction to ensure it runs with proper context if needed,
      // though direct call should be fine here.
      beginFirstTime();
    }
  } else {
    LoggerManager.logDebug("Not first run.");
    // Ensure triggers are set up on subsequent opens if they were somehow deleted.
    TriggerManager.createTrigger();
    // Load properties on open to cache them
    PropertyManager.loadDocumentProperties();
  }
}

/**
 * Initial setup function called only on the very first run for a new sheet copy.
 * Ensures essential properties are initialized, creates triggers, shows welcome message,
 * and initiates the habit setup UI.
 */
function beginFirstTime() {
  LoggerManager.logDebug("Executing beginFirstTime...");
  try {
    // Ensure essential properties are initialized by simply accessing them.
    // If they don't exist, getProperty will call _initializeDefaultProperty internally.
    LoggerManager.logDebug("Initializing properties by access...");
    PropertyManager.getProperty(PropertyKeys.MODE); // Ensures MODE is set (defaults to HABIT_IDEATION)
    PropertyManager.getProperty(PropertyKeys.FIRST_CHALLENGE_DATE); // Ensures date/row are set
    PropertyManager.getProperty(PropertyKeys.RESET_HOUR); // Ensures reset hour is set
    PropertyManager.getProperty(PropertyKeys.BOOST_INTERVAL); // Ensures boost interval is set
    PropertyManager.getProperty(PropertyKeys.EMOJI_LIST); // Ensures emoji list is initialized (likely to '[]')
    PropertyManager.getProperty(PropertyKeys.LAST_DATE_SELECTOR_UPDATE); // Initialize timestamps
    PropertyManager.getProperty(PropertyKeys.LAST_COMPLETION_UPDATE);
    PropertyManager.getProperty(PropertyKeys.LAST_UPDATE);

    // Ensure properties are saved immediately after potential first initialization
    LoggerManager.logDebug("Saving potentially initialized properties...");
    PropertyManager.setDocumentProperties();

    // Create the daily trigger
    LoggerManager.logDebug("Creating trigger...");
    TriggerManager.createTrigger(); // <<< This now requires script.scriptapp scope

    // Show welcome message
    Messages.showAlert(MessageTypes.WELCOME_MESSAGE);

    // Initiate the habit setup UI
    LoggerManager.logDebug("Initializing Set Habit UI...");
    HabitManager.initializeSetHabitUI(); // This sets mode to HABIT_IDEATION again, which is fine
    Messages.showAlert(MessageTypes.CHALLENGE_RESET); // Inform user to set habits

    LoggerManager.logDebug("beginFirstTime completed successfully.");
  } catch (error) {
    // Use LoggerManager to handle the error display and potential throw
    LoggerManager.handleError(
      `Error during first-time initialization (beginFirstTime): ${error.message}\n${error.stack}`,
      true
    );
  }
}

/**
 * Checks if this is the first time the script is being run for this document
 * by checking for the existence of a specific property.
 * @returns {boolean} True if it's the first run, false otherwise.
 */
function isFirstRun() {
  // Check directly using PropertiesService to avoid initialization loops in onOpen
  const firstDate = PropertiesService.getDocumentProperties().getProperty(
    PropertyKeys.FIRST_CHALLENGE_DATE
  );
  LoggerManager.logDebug(
    `isFirstRun check: Property "${PropertyKeys.FIRST_CHALLENGE_DATE}" value is "${firstDate}"`
  );
  return !firstDate; // If the property doesn't exist, it's the first run
}

/**
 * Handles the 'onEdit' event trigger fired by Google Sheets.
 * Delegates the processing to DataHandler.handleCellEdit.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object.
 */
function onEdit(e) {
  LoggerManager.logDebug(
    `onEdit Triggered: Range ${e.range.getA1Notation()}, OldValue: ${
      e.oldValue
    }, Value: ${e.value}`
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
 * It simulates an edit on the date cell to trigger the standard update logic.
 */
function renewChecklistForToday() {
  LoggerManager.logDebug(
    `Time-driven trigger: renewChecklistForToday running.`
  );
  const sheet = MainSheetConfig._getSheet(); // Use internal getter
  if (!sheet) {
    LoggerManager.handleError(
      "renewChecklistForToday: Cannot find main sheet.",
      true
    );
    return;
  }
  const range = sheet.getRange(MainSheetConfig.dateCell);
  const oldValue = range.getValue(); // Get the actual old value from the sheet

  // Update the date cell visually first
  const todayStr = DateManager.getTodayStr();
  LoggerManager.logDebug(
    `Setting date cell ${MainSheetConfig.dateCell} to today: ${todayStr}`
  );
  range.setValue(todayStr); // Directly set the value

  // Now, call handleCellEdit to process the change as if the user did it
  // Ensure oldValue is passed correctly. If oldValue was already today,
  // the logic in handleCellEdit/renewChecklist should ideally handle this gracefully (e.g., no save needed).
  LoggerManager.logDebug(
    `Simulating edit for date update. Old value was: ${oldValue}`
  );
  DataHandler.handleCellEdit(sheet, range, oldValue); // Trigger the logic
  LoggerManager.logDebug(`renewChecklistForToday finished.`);
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
      // Handle case where version check failed
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
 * Menu item function: Starts the process for a new challenge.
 * Confirms with the user before resetting.
 */
function startNewChallenge() {
  const response = Messages.showAlert(MessageTypes.START_NEW_CHALLENGE);

  if (response === Messages.ButtonTypes.YES) {
    LoggerManager.logDebug("User confirmed starting new challenge.");
    HabitManager.initializeSetHabitUI(); // Resets data, sets up UI, sets mode
    Messages.showAlert(MessageTypes.CHALLENGE_RESET); // Inform user to set habits
    PropertyManager.setDocumentProperties(); // Save mode change etc.
  } else {
    LoggerManager.logDebug("User cancelled starting new challenge.");
    Messages.showAlert(MessageTypes.CHALLENGE_CANCELLED);
  }
}

/**
 * Menu item function: Terminates the current challenge.
 * Confirms with the user, hides columns, resets data, sets mode.
 */
function terminateChallenge() {
  const response = Messages.showAlert(MessageTypes.TERMINATION_CONFIRMATION);
  if (response === Messages.ButtonTypes.YES) {
    LoggerManager.logDebug("User confirmed terminating challenge.");
    // Hide tracking columns
    MainSheetConfig.toggleColumns(ColumnAction.HIDE);
    // Reset streaks, date, clear completion/buffer UI
    MainSheetConfig.resetChallengeDataUI();
    // Clear relevant history? Maybe not, keep history but stop tracking.

    // Set mode to terminated and save immediately
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
  // Delegate directly to the Help module function
  HelpManager.showHelpSidebar();
}
