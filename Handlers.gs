/**
 * Initializes the application by setting up necessary triggers, displaying a welcome message,
 * and starting a new challenge.
 */
function begin() {
  TriggerManager.createTrigger();
  Messages.showAlert(MessageTypes.WELCOME_MESSAGE);
  startNewChallenge();
}

/**
 * Checks whether a new version of the library is available.
 * If available, displays an alert to inform the user about the update.
 */
function versionCheck() {
  LibraryManager.fetchVersionInfo() > LibraryManager.LATEST_VERSION
    ? Messages.showAlert(MessageTypes.NEW_VERSION_AVAILABLE)
    : Messages.showAlert(MessageTypes.NO_NEW_UPDATES);
}

/**
 * Determines if this is the first time the application is being run.
 *
 * This function checks whether a specific property, `FIRST_CHALLENGE_DATE`,
 * has been set in the property manager. If the property does not exist,
 * it indicates that the application is being run for the first time.
 *
 * @returns {boolean} - Returns true if this is the first run, false otherwise.
 */
function isFirstRun() {
  return !PropertyManager.hasProperty(PropertyKeys.FIRST_CHALLENGE_DATE);
}

/**
 * Resets the checklist for today by updating the date cell to today's date and handling any necessary data updates.
 * This function is triggered automatically by a time-based trigger set to fire at a specific time each day.
 * It ensures that the checklist is up-to-date and handles the data for the new day.
 */
function renewChecklistForToday() {
  LoggerManager.logDebug(`renewChecklistForToday triggered.`);

  const sheet = MainSheetConfig.getSheet();
  const range = sheet.getRange(MainSheetConfig.dateCell);
  const oldValue = MainSheetConfig.getDate();

  // Update the date cell to today's date
  LoggerManager.logDebug("Updating date cell to today's date.");
  const todayStr = DateManager.getTodayStr();
  MainSheetConfig.setDateStr(todayStr);

  LoggerManager.logDebug(
    `renewChecklistForToday: Calling handleCellEdit for today's update with sheet: ${sheet.getName()}, range: ${range.getA1Notation()}, and oldValue: ${oldValue}`
  );
  DataHandler.handleCellEdit(sheet, range, oldValue);
}

/**
 * Handles interactions with the date selector, validates and formats dates, and manages the flow of saving, propagating,
 * and loading new data when a date is edited in the main sheet.
 *
 * This function is triggered whenever a cell in the sheet is edited. It checks if the edit is relevant (e.g.,
 * if the date selector or completion data is modified), and if so, it processes the update accordingly.
 *
 * @param {Object} e - The event object that triggered the onEdit function, containing details about the edit.
 */
function onEdit(e) {
  LoggerManager.logDebug(`onEdit triggered.`);

  // Check if the event object e exists
  if (e) {
    LoggerManager.logDebug(`Event object exists.`);

    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const oldValue = e.oldValue;
    LoggerManager.logDebug(
      `onEdit... Sheet: ${sheet.getName()}, Range: ${range.getA1Notation()}, Old Value: ${oldValue}`
    );

    // Call the handleEdit function with the appropriate parameters
    DataHandler.handleCellEdit(sheet, range, oldValue);
  } else {
    LoggerManager.logDebug(`Event object does not exist.`);
  }
}

/**
 * The `DataHandler` object manages and processes core data operations within the Google Sheets document.
 *
 * It is responsible for handling events triggered by user interactions, such as editing cells,
 * and for maintaining the consistency and integrity of the data across different dates.
 * The `DataHandler` ensures that updates to the main sheet and history sheet are correctly
 * applied and logged, managing state through properties stored in `PropertyManager`.
 *
 * This object interacts with other components like `MainSheetConfig`, `HistorySheetConfig`,
 * `PropertyManager`, and `DateManager` to coordinate the flow of data and maintain accurate records.
 */
const DataHandler = {
  /**
   * Converts a multi-cell range into an array of A1 notation cells.
   * @param {Range} range - The range object representing the edited cells.
   * @returns {Array<string>} - An array of A1 notation strings for each cell in the range.
   */
  getRangeCells: function (range) {
    const rangeA1Notation = range.getA1Notation();

    if (!rangeA1Notation.includes(":")) {
      // Single cell, add directly to rangeCells
      return [rangeA1Notation];
    }

    let rangeCells = [];

    // Extract start and end positions for rows and columns
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    // Loop through each cell within the range and construct A1 notation
    for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
      for (let colOffset = 0; colOffset < numCols; colOffset++) {
        const cellRow = startRow + rowOffset;
        const cellCol = startCol + colOffset;

        // Construct the A1 notation manually
        const columnLetter = this.columnToLetter(cellCol);
        const cellNotation = `${columnLetter}${cellRow}`;
        rangeCells.push(cellNotation);
      }
    }
    return rangeCells;
  },

  /**
   * Converts a column number (1-based) to a corresponding column letter in A1 notation.
   * For example, 1 -> "A", 2 -> "B", ..., 26 -> "Z", 27 -> "AA".
   * @param {number} column - The 1-based column number.
   * @returns {string} - The corresponding column letter.
   */
  columnToLetter: function (column) {
    let letter = "";
    while (column > 0) {
      let mod = (column - 1) % 26;
      letter = String.fromCharCode(mod + 65) + letter;
      column = Math.floor((column - 1) / 26);
    }
    return letter;
  },

  /**
   * Retrieves an array of column indices within the given range.
   *
   * @param {Range} range - The range object representing the edited cells.
   * @returns {Array<number>} - An array containing the column indices for each column within the range.
   *
   * This method is useful for cases where the edit may affect multiple columns,
   * and a quick check of which columns are involved is necessary.
   */
  getRangeColumns: function (range) {
    return Array.from(
      { length: range.getNumColumns() },
      (_, i) => range.getColumn() + i
    );
  },

  /**
   * Processes the cell edit event, handling various tasks such as updating properties,
   * saving data, and triggering necessary updates. This function is called internally
   * by the `onEdit` trigger or when manually updating the date selector.
   *
   * @param {Object} sheet - The active sheet where the edit event occurred.
   * @param {Object} range - The range object representing the edited cell(s).
   * @param {string} oldValue - The previous value of the edited cell before the change.
   */
  handleCellEdit: function (sheet, range, oldValue) {
    LoggerManager.logDebug(`handleCellEdit triggered`);
    const mainSheetID = MainSheetConfig.getSheet().getSheetId();
    const historySheetID = HistorySheetConfig.getSheet().getSheetId();

    // Ensure the edit is relevant to the main sheet
    if (sheet.getSheetId() !== mainSheetID) {
      if (sheet.getSheetId() === historySheetID) {
        MainSheetConfig.maintainCellValue(range, oldValue);
      }
      LoggerManager.logDebug(`Edit is not relevant to the main sheet, exiting`);
      return;
    }

    // Check for edits to locked data (buffer days, streaks)
    if (MainSheetConfig.includesLockedRange(range)) {
      MainSheetConfig.maintainCellValue(range, oldValue);
      LoggerManager.logDebug(
        `User attempted to edit the buffer column or streak cells, which can only be changed as product of the propagation of completion data. Exiting`
      );
    }

    const mode = PropertyManager.getProperty(PropertyKeys.MODE);
    if (mode === ModeTypes.TERMINATED) {
      this._handleTerminatedMode(range, oldValue);
    } else if (mode === ModeTypes.HABIT_IDEATION) {
      this._handleHabitIdeationMode(range);
    } else if (mode === ModeTypes.CHALLENGE) {
      this._handleChallengeMode(range, oldValue);
    }

    PropertyManager.setDocumentProperties();
  },

  /**
   * Handles cell edits when the application is in 'Terminated' mode.
   *
   * This method is triggered when the user attempts to make changes to the sheet while
   * in Terminated mode, which is when habit tracking has been intentionally stopped.
   * It alerts the user to start a new challenge to proceed and reverts the cell to its original value.
   *
   * @param {Range} range - The edited cell range.
   * @param {any} oldValue - The previous value of the edited cell.
   */
  _handleTerminatedMode: function (range, oldValue) {
    LoggerManager.logDebug(
      `Challenges have been terminated. You must start a new challenge to proceed.`
    );
    Messages.showAlert(MessageTypes.TERMINATION_REMINDER);
    MainSheetConfig.maintainCellValue(range, oldValue);
  },

  /**
   * Handles cell edits when the application is in 'Habit Ideation' mode.
   *
   * This method is triggered when the user is setting up new habits. It checks if the edited
   * cell corresponds to the range for setting habits, and if so, triggers the habit spread setup.
   *
   * @param {Range} range - The edited cell range.
   */
  _handleHabitIdeationMode: function (range) {
    const rangeCells = this.getRangeCells(range);
    if (rangeCells.includes(MainSheetConfig.setterRanges.setHabit)) {
      HabitManager.setHabitSpread();
    }
    LoggerManager.logDebug(
      `Habit Ideation Mode is on. You must set habit to proceed.`
    );
  },

  /**
   * Handles cell edits when the application is in 'Challenge' mode.
   *
   * This method processes edits based on the active columns. It updates properties and ensures
   * checkboxes are inserted if completion data cells were modified. It also checks for date
   * selection changes to trigger relevant updates.
   *
   * @param {Range} range - The edited cell range.
   * @param {any} oldValue - The previous value of the edited cell.
   */
  _handleChallengeMode: function (range, oldValue) {
    // Get an array of all column numbers in the range
    const columns = this.getRangeColumns(range);
    const rangeCells = this.getRangeCells(range);

    // Check if the edit is relevant to the activities column
    if (columns.includes(MainSheetConfig.activityDataColumn)) {
      // Mark that the activities column was edited
      PropertyManager.setProperty(
        PropertyKeys.ACTIVITIES_COLUMN_UPDATED,
        BooleanTypes.TRUE
      );
    }

    // Handle completion data updates
    if (columns.includes(MainSheetConfig.completionDataColumn)) {
      LoggerManager.logDebug(
        `Completion data updated, storing last completion update.`
      );
      PropertyManager.setProperty(
        PropertyKeys.LAST_COMPLETION_UPDATE,
        DateManager.getNow()
      );
      // If multiple cells were edited, ensure checkboxes have been inserted
      if (oldValue === undefined || range.getValue() === "") {
        LoggerManager.logDebug(`Ensuring checkboxes are inserted.`);
        MainSheetConfig.insertCompletionCheckboxes();
      }
    }

    // Ensure the edit is relevant to the target columns
    if (!rangeCells.includes(MainSheetConfig.dateCell)) {
      LoggerManager.logDebug(
        `Edit is not relevant to the main sheet or target columns, exiting.`
      );
      return; // Exit if the function is run manually from the editor
    }

    this._handleDateSelectorUpdate(range, oldValue);
  },

  /**
   * Updates the checklist based on changes to the date selector cell.
   *
   * This method handles edits to the date selector cell, triggering necessary updates
   * to the application’s state. It manages the 'last update' property and alerts the user
   * if the new date cannot be formatted correctly.
   *
   * @param {Range} range - The edited cell range.
   * @param {any} oldValue - The previous value of the edited cell.
   */
  _handleDateSelectorUpdate: function (range, oldValue) {
    // Set 'lastUpdate' property
    PropertyManager.updateLastUpdateProperty();

    // Store the current time when date selector is updated
    PropertyManager.setProperty(
      PropertyKeys.LAST_DATE_SELECTOR_UPDATE,
      DateManager.getNow()
    );

    let previousDateStr, selectedDateStr;

    try {
      previousDateStr = DateManager.determineFormattedDate(oldValue);
    } catch (e) {
      LoggerManager.logDebug(
        `handleCellEdit: previousDateStr can't be casted into a formatted date.`
      );
      previousDateStr = null;
      PropertyManager.updateLastUpdateProperty(); // Will not save the previous date.
    }

    try {
      selectedDateStr = DateManager.determineFormattedDate(range.getValue());
    } catch (e) {
      LoggerManager.logDebug(
        `handleCellEdit: selectedDate, ${range.getValue()}, can't be casted into a formatted date.`
      );
      Messages.showAlert(MessageTypes.INVALID_DATE);
      selectedDateStr = DateManager.getTodayStr();
    }

    LoggerManager.logDebug(
      `Calling renewChecklist with previousDateStr of ${previousDateStr} and selectedDate of ${selectedDateStr}.`
    );
    this.renewChecklist(previousDateStr, selectedDateStr);
  },

  /**
   * Determines whether to save and propagate changes or simply display the selected date
   * based on the most recent update. It checks the `lastUpdate` property to determine
   * whether the last significant change was related to completion data or the date selector.
   *
   * If the completion data was updated more recently, it proceeds with saving the current state
   * and propagating changes.
   *
   * Regardless, it then displays the data for the selected date.
   *
   * @param {string} previousDateStr - The previous date string (in YYYY-MM-DD format) that was selected before the change.
   * @param {string} selectedDateStr - The new date selected by the user as a string (in YYYY-MM-DD format).
   */
  renewChecklist: function (previousDateStr, selectedDateStr) {
    // Retrieve the property indicating the last significant update (completion or dateSelector)
    const lastUpdate = PropertyManager.getProperty(PropertyKeys.LAST_UPDATE);

    // If the last update was related to completion data, proceed with saving and propagating changes
    if (lastUpdate === LastUpdateTypes.COMPLETION) {
      LoggerManager.logDebug(
        `Completion data was updated after the date selector. Must first save current data.`
      );
      HistorySheetConfig.saveCurrentState(previousDateStr);
    } else {
      LoggerManager.logDebug(
        `No need to save and propagate. Date selector was updated more recently.`
      );
    }

    // Display the data for the newly selected date
    MainSheetConfig.displayDate(selectedDateStr);
  },
};
