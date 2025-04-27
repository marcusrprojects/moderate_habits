/**
 * @fileoverview Manages habit setup, validation, and finalization during
 * the Habit Ideation phase. Handles UI elements related to habit settings.
 */

/** OnlyCurrentDoc */

const HabitManager = {
  /**
   * Regular expression pattern for detecting emojis.
   * Matches Unicode emoji presentations and extended pictographics.
   * @constant {RegExp}
   */
  emojiPattern: /(\p{Emoji_Presentation}|\p{Extended_Pictographic})+/gu,

  /**
   * Initializes the UI for Habit Ideation mode.
   * - Hides tracking columns.
   * - Resets main sheet data/UI.
   * - Shows and sets up habit setting fields (emojis, setters).
   * - Sets application mode property.
   * - Clears today's history entry if it exists from a previous attempt.
   */
  initializeSetHabitUI: function () {
    LoggerManager.logDebug("Initializing Habit Ideation UI...");
    // 1. Hide core tracking columns
    MainSheetConfig.toggleColumns(ColumnAction.HIDE);

    // 2. Reset main sheet UI (streaks, date, clear completion/buffer)
    MainSheetConfig.resetChallengeDataUI();

    // 3. Show and configure the setter fields
    this._toggleSetterFields(CellAction.SET);

    // 4. Set the application mode
    PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.HABIT_IDEATION);

    // 5. Clear potentially stale history entry for today
    const historySheet = HistorySheetConfig._getSheet();
    const todayStr = DateManager.getTodayStr();
    const lastHistoryDate = DataHandler.getLastHistoryDate();
    if (
      historySheet &&
      lastHistoryDate &&
      DateManager.determineFormattedDate(lastHistoryDate) === todayStr
    ) {
      const lastRow = historySheet.getLastRow();
      const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1;
      if (lastRow >= firstDataSheetRow) {
        try {
          historySheet.deleteRow(lastRow);
          SpreadsheetApp.flush();
        } catch (e) {
          LoggerManager.handleError(
            `Failed to delete last history row (${lastRow}): ${e.message}`,
            false
          );
        }
      }
    }

    // Properties will be saved by the caller (startNewChallenge)
  },

  /**
   * Toggles the visibility and configuration of the habit/property setter fields.
   * @private
   * @param {CellAction} action - CellAction.SET or CellAction.CLEAR.
   */
  _toggleSetterFields: function (action) {
    if (action !== CellAction.CLEAR && action !== CellAction.SET) {
      LoggerManager.handleError(
        `_toggleSetterFields called with unknown action: ${action}.`,
        true
      );
      return;
    }
    LoggerManager.logDebug(`Toggling setter fields: ${action}`);

    const sheet = MainSheetConfig._getSheet();
    if (!sheet) return;

    const isSetting = action === CellAction.SET;

    // --- Handle Property Setters (Reset Hour, Boost Interval) ---
    const resetHourRange = sheet.getRange(
      MainSheetConfig.setterRanges.resetHour
    );
    const boostIntervalRange = sheet.getRange(
      MainSheetConfig.setterRanges.boostInterval
    );
    const resetHourLabelRange = sheet.getRange(
      MainSheetConfig.setterLabelRanges.resetHour
    );
    const boostIntervalLabelRange = sheet.getRange(
      MainSheetConfig.setterLabelRanges.boostInterval
    );

    if (isSetting) {
      // Set default values from properties into the cells
      resetHourRange.setValue(
        PropertyManager.getProperty(PropertyKeys.RESET_HOUR)
      );
      boostIntervalRange.setValue(
        PropertyManager.getProperty(PropertyKeys.BOOST_INTERVAL)
      );
      // Set labels and notes
      resetHourLabelRange
        .setValue(MainSheetConfig.setterLabels.resetHour)
        .setBackground(GlobalConfig.secondaryColor)
        .setFontWeight("bold");
      boostIntervalLabelRange
        .setValue(MainSheetConfig.setterLabels.boostInterval)
        .setBackground(GlobalConfig.secondaryColor)
        .setFontWeight("bold");
      resetHourRange.setNote(MainSheetConfig.setterNotes.resetHour);
      boostIntervalRange.setNote(MainSheetConfig.setterNotes.boostInterval);
    } else {
      // Clear values, notes, labels, formatting
      resetHourRange.clearContent().clearNote();
      boostIntervalRange.clearContent().clearNote();
      resetHourLabelRange
        .clearContent()
        .setBackground(GlobalConfig.mainColor)
        .setFontWeight("normal");
      boostIntervalLabelRange
        .clearContent()
        .setBackground(GlobalConfig.mainColor)
        .setFontWeight("normal");
    }

    // --- Handle "Set Habit" Checkbox ---
    const setHabitCellRange = sheet.getRange(
      MainSheetConfig.setterRanges.setHabit
    );
    const setHabitLabelRange = sheet.getRange(
      MainSheetConfig.setterLabelRanges.setHabit
    );

    if (isSetting) {
      setHabitCellRange.insertCheckboxes().setValue(false); // Insert unchecked checkbox
      setHabitCellRange.setNote(MainSheetConfig.setterNotes.setHabit);
      setHabitLabelRange
        .setValue(MainSheetConfig.setterLabels.setHabit)
        .setBackground(GlobalConfig.secondaryColor)
        .setFontWeight("bold");
    } else {
      setHabitCellRange.clearContent().clearDataValidations().clearNote();
      setHabitLabelRange
        .clearContent()
        .setBackground(GlobalConfig.mainColor)
        .setFontWeight("normal");
    }

    SpreadsheetApp.flush(); // Apply UI changes
  },

  /**
   * Validates the user-entered values for reset hour and boost interval.
   * @private
   * @param {boolean} [throwError=true] - Whether to throw error on failure.
   * @returns {boolean} True if both setters are valid.
   */
  _validateSetters: function (throwError = true) {
    const sheet = MainSheetConfig._getSheet();
    if (!sheet) return false;

    const resetHourValue = sheet
      .getRange(MainSheetConfig.setterRanges.resetHour)
      .getValue();
    const boostIntervalValue = sheet
      .getRange(MainSheetConfig.setterRanges.boostInterval)
      .getValue();

    let isValid = true;

    // Validate Reset Hour (0-23 integer)
    if (
      !ValidationUtils._validateNonNegativeInteger(resetHourValue) ||
      resetHourValue > 23
    ) {
      const msg = `Invalid Reset Hour: ${resetHourValue}. Must be a whole number between 0 and 23.`;
      LoggerManager.handleError(msg, throwError);
      isValid = false;
    }

    // Validate Boost Interval (>= 1 integer)
    if (!ValidationUtils._validatePositiveInteger(boostIntervalValue)) {
      const msg = `Invalid Boost Interval: ${boostIntervalValue}. Must be a whole number greater than or equal to 1.`;
      LoggerManager.handleError(msg, throwError);
      isValid = false;
    }

    if (!isValid && throwError) {
      // If throwing, the above handlers would have stopped execution.
      // If not throwing, show a single message.
      Messages.showAlert(MessageTypes.INVALID_SETTERS);
    } else if (!isValid) {
      LoggerManager.logDebug("Setter validation failed (throwError=false).");
    }

    return isValid;
  },

  /**
   * Retrieves the current list of habits (emojis and their rows) from the main sheet's activity column.
   * @returns {Array<{emoji: string, row: number}>} List of habit objects.
   */
  getCurrentEmojiSpread: function () {
    const activityRange = MainSheetConfig.getActivityDataRange();
    if (!activityRange) {
      LoggerManager.logDebug(
        "getCurrentEmojiSpread: Activity range not found."
      );
      return [];
    }

    const sheet = MainSheetConfig._getSheet();
    const firstRow = activityRange.getRow();
    const values = activityRange.getValues(); // [[val1], [val2], ...]
    const emojiList = [];

    values.forEach((cellValue, index) => {
      const emoji = cellValue[0];
      // Check if the cell content contains an emoji based on the regex
      if (typeof emoji === "string" && emoji.match(this.emojiPattern)) {
        const rowNum = firstRow + index; // Calculate 1-based row number
        emojiList.push({ emoji: emoji, row: rowNum });
      }
    });

    LoggerManager.logDebug(
      `getCurrentEmojiSpread found: ${JSON.stringify(emojiList)}`
    );
    return emojiList;
  },

  /**
   * Saves the current emoji spread from the sheet to document properties.
   */
  updateEmojiSpreadProperty: function () {
    const currentSpread = this.getCurrentEmojiSpread();
    PropertyManager.setProperty(
      PropertyKeys.EMOJI_LIST,
      JSON.stringify(currentSpread)
    );
    LoggerManager.logDebug("Updated emoji spread property.");
  },

  /**
   * Processes the confirmation step when the 'Set Habit' checkbox is ticked.
   * Validates habits and setters, asks for confirmation, then finalizes.
   * @param {GoogleAppsScript.Spreadsheet.Range} checkboxRange - The range of the 'Set Habit' checkbox.
   */
  processHabitSpreadConfirmation: function (checkboxRange) {
    LoggerManager.logDebug("Processing habit spread confirmation...");

    // 1. Validate Setters (Reset Hour, Boost Interval)
    if (!this._validateSetters(false)) {
      // Don't throw error yet, show specific message
      Messages.showAlert(MessageTypes.INVALID_SETTERS);
      checkboxRange.setValue(false); // Uncheck the box
      SpreadsheetApp.flush();
      return;
    }

    // 2. Validate Habits (at least one emoji must be present)
    const currentEmojiSpread = this.getCurrentEmojiSpread();
    if (currentEmojiSpread.length === 0) {
      Messages.showAlert(MessageTypes.NO_HABITS_SET);
      checkboxRange.setValue(false); // Uncheck the box
      SpreadsheetApp.flush();
      return;
    }

    // 3. Ask for User Confirmation
    const response = Messages.showAlert(MessageTypes.CONFIRM_HABIT_SPREAD);

    // 4. Finalize or Reset Checkbox
    if (response === Messages.ButtonTypes.YES) {
      LoggerManager.logDebug("User confirmed habit spread. Finalizing...");
      this._finalizeHabitSpread();
    } else {
      LoggerManager.logDebug("User cancelled habit spread confirmation.");
      checkboxRange.setValue(false); // Reset the checkbox
      SpreadsheetApp.flush();
    }
  },

  /**
   * Finalizes the habit setup:
   * - Saves valid setters to properties.
   * - Saves the emoji spread to properties.
   * - Updates first challenge date/row properties.
   * - Changes mode to CHALLENGE.
   * - Cleans up setter UI.
   * - Shows tracking columns.
   * - Sets background for activity column.
   * - Inserts completion checkboxes.
   * - Displays today's data.
   * @private
   */
  _finalizeHabitSpread: function () {
    const sheet = MainSheetConfig._getSheet();
    if (!sheet) return;

    LoggerManager.logDebug("Finalizing habit spread...");

    // 1. Save validated setters to properties (already validated)
    const resetHourValue = sheet
      .getRange(MainSheetConfig.setterRanges.resetHour)
      .getValue();
    const boostIntervalValue = sheet
      .getRange(MainSheetConfig.setterRanges.boostInterval)
      .getValue();
    PropertyManager.setProperty(
      PropertyKeys.RESET_HOUR,
      String(resetHourValue)
    );
    PropertyManager.setProperty(
      PropertyKeys.BOOST_INTERVAL,
      String(boostIntervalValue)
    );

    // 2. Save the validated emoji spread
    this.updateEmojiSpreadProperty();

    // 3. Set first challenge date/row (marks the official start)
    PropertyManager._updateFirstChallengeDateAndRow(); // Use internal method

    // 4. Change mode
    PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.CHALLENGE);

    // 5. Clean up UI
    this._toggleSetterFields(CellAction.CLEAR);
    MainSheetConfig.toggleColumns(ColumnAction.SHOW);

    // 6. Format activity column background (optional styling)
    const activitiesRange = MainSheetConfig.getActivityDataRange();
    if (activitiesRange) {
      activitiesRange.setBackground(GlobalConfig.mainColor);
    }

    // 7. Insert checkboxes for the newly defined relevant rows
    MainSheetConfig.insertCompletionCheckboxes();

    // 8. Display today's data (this will create the first history entry)
    PropertyManager.updateLastUpdateProperty(); // Ensure last update is set correctly before display
    DataHandler.displayDate(DateManager.getTodayStr());

    // 9. Save all property changes made during finalization
    PropertyManager.setDocumentProperties();

    LoggerManager.logDebug("Habit spread finalized and challenge started.");
  },

  /**
   * Checks if the current emoji spread on the sheet matches the stored one.
   * If not, reverts the sheet to the stored spread and alerts the user.
   * Called by DataHandler.getRelevantRows if changes are detected in challenge mode.
   * @returns {boolean} True if spread matches or was successfully reverted, False if mismatch occurred but revert failed.
   */
  checkEmojiSpread: function () {
    LoggerManager.logDebug(
      "Checking current emoji spread against stored property..."
    );
    const currentSpread = this.getCurrentEmojiSpread();
    let storedSpread = [];
    try {
      storedSpread = JSON.parse(
        PropertyManager.getProperty(PropertyKeys.EMOJI_LIST) || "[]"
      );
    } catch (e) {
      LoggerManager.handleError(
        `Failed to parse stored emoji list for comparison: ${e.message}`,
        true
      ); // Critical error
      return false;
    }

    // Compare spreads (simple length and element-wise check)
    let match = currentSpread.length === storedSpread.length;
    if (match) {
      for (let i = 0; i < currentSpread.length; i++) {
        if (
          currentSpread[i].emoji !== storedSpread[i].emoji ||
          currentSpread[i].row !== storedSpread[i].row
        ) {
          match = false;
          break;
        }
      }
    }

    if (!match) {
      LoggerManager.logDebug("Emoji spread mismatch detected! Reverting...");
      Messages.showAlert(MessageTypes.HABIT_SPREAD_RESET);
      if (!this._revertEmojiSpread(storedSpread)) {
        LoggerManager.handleError("Failed to revert emoji spread.", false); // Log failure but continue if possible
        return false; // Indicate failure
      }
      return true; // Reverted successfully
    } else {
      LoggerManager.logDebug("Emoji spread matches stored property.");
      return true; // Spread matches
    }
  },

  /**
   * Reverts the emoji cells in the main sheet to match the stored spread.
   * Clears cells that shouldn't have emojis according to the stored spread.
   * @private
   * @param {Array<{emoji: string, row: number}>} storedSpread - The correct spread data.
   * @returns {boolean} True if successful, False on error.
   */
  _revertEmojiSpread: function (storedSpread) {
    const sheet = MainSheetConfig._getSheet();
    if (!sheet) return false;

    const activityCol = MainSheetConfig.activityDataColumn;
    const activityRange = MainSheetConfig.getActivityDataRange();
    if (!activityRange) {
      LoggerManager.logDebug("_revertEmojiSpread: Cannot get activity range.");
      // Can't clear cells, but maybe can still set the correct ones?
      // Let's proceed but log.
    }

    const firstRow = activityRange
      ? activityRange.getRow()
      : MainSheetConfig.firstDataInputRow;
    const lastRow = activityRange
      ? firstRow + activityRange.getNumRows() - 1
      : sheet.getLastRow();

    LoggerManager.logDebug(
      `Reverting emoji spread. Stored: ${JSON.stringify(storedSpread)}`
    );

    try {
      // Create a map for quick lookup of what *should* be in each row
      const storedRowMap = new Map();
      storedSpread.forEach((item) => storedRowMap.set(item.row, item.emoji));

      // Iterate through all rows in the activity column range
      for (let r = firstRow; r <= lastRow; r++) {
        const cell = sheet.getRange(r, activityCol);
        const currentCellValue = cell.getValue();
        const expectedEmoji = storedRowMap.get(r);

        if (expectedEmoji) {
          // This row *should* have an emoji. Set it if it's wrong.
          if (currentCellValue !== expectedEmoji) {
            LoggerManager.logDebug(
              `Setting row ${r} to emoji: ${expectedEmoji}`
            );
            cell.setValue(expectedEmoji);
          }
        } else {
          // This row *should NOT* have an emoji (based on stored list). Clear it if it does.
          if (
            typeof currentCellValue === "string" &&
            currentCellValue.match(this.emojiPattern)
          ) {
            LoggerManager.logDebug(`Clearing unexpected emoji in row ${r}.`);
            cell.clearContent();
          }
          // Optionally reset background/formatting?
          cell.setBackground(GlobalConfig.mainColor);
        }
      }
      SpreadsheetApp.flush();
      LoggerManager.logDebug("Emoji spread reverted successfully.");
      return true;
    } catch (e) {
      LoggerManager.handleError(
        `Error during emoji spread revert: ${e.message}`,
        false
      );
      return false;
    }
  },
};

// Freeze the manager object
Object.freeze(HabitManager);
