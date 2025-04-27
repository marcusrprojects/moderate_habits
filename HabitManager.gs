/**
 * @fileoverview Manages the workflow for defining, validating, and finalizing user habits.
 * This includes handling the 'Habit Ideation' mode UI, processing user settings,
 * managing the emoji spread, and transitioning the application state to 'Challenge' mode.
 */

/** OnlyCurrentDoc */

/**
 * Manages the setup, validation, and finalization of habits within the application.
 * Handles the Habit Ideation UI state and transitions.
 * @namespace HabitManager
 */
const HabitManager = {
  /**
   * Regular expression pattern for detecting emojis based on Unicode properties.
   * Matches sequences of Emoji Presentation or Extended Pictographic characters.
   * @constant {RegExp}
   */
  emojiPattern: /(\p{Emoji_Presentation}|\p{Extended_Pictographic})+/gu,

  /**
   * Initializes the user interface for the 'Habit Ideation' mode.
   * This function prepares the main sheet for the user to define their habits and settings.
   * Actions include:
   * - Hiding columns related to active challenge tracking (Completion, Buffer, Streaks).
   * - Resetting streak and date values on the main sheet UI.
   * - Clearing previous completion and buffer data from the main sheet UI.
   * - Displaying and configuring the UI elements for setting habits, reset hour, and boost interval.
   * - Setting the application mode property to 'HABIT_IDEATION'.
   * - Deleting any existing history entry for the current day to prevent conflicts if resetting on the same day.
   * @throws {Error} If critical sheet operations fail (e.g., finding required sheets).
   */
  initializeSetHabitUI: function () {
    LoggerManager.logDebug(">>> Entering HabitManager.initializeSetHabitUI");
    try {
      // 1. Hide core tracking columns
      LoggerManager.logDebug("Hiding tracking columns.");
      MainSheetConfig.toggleColumns(ColumnAction.HIDE);

      // 2. Reset main sheet UI (streaks, date, clear completion/buffer)
      LoggerManager.logDebug("Resetting main sheet challenge UI.");
      MainSheetConfig.resetChallengeDataUI();

      // 3. Show and configure the setter fields in columns H/I
      LoggerManager.logDebug("Setting up setter fields UI.");
      this._toggleSetterFields(CellAction.SET);

      // 4. Set the application mode (property saved later by caller)
      LoggerManager.logDebug(
        "Setting application mode property to HABIT_IDEATION."
      );
      PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.HABIT_IDEATION);

      // 5. Clear potentially stale history entry for today (if resetting on same day)
      LoggerManager.logDebug(
        "Checking/Clearing potentially stale history entry for today..."
      );
      try {
        // Wrap history interaction in its own try/catch for resilience
        const historySheet = HistorySheetConfig._getSheet(); // Use internal getter safely
        const todayStr = DateManager.getTodayStr();
        const lastHistoryDate = DataHandler.getLastHistoryDate(); // Use DataHandler

        if (
          historySheet &&
          lastHistoryDate &&
          DateManager.determineFormattedDate(lastHistoryDate) === todayStr
        ) {
          const lastRow = historySheet.getLastRow();
          const firstDataSheetRow = HistorySheetConfig.firstDataRow + 1; // 1-based for comparison
          if (lastRow >= firstDataSheetRow) {
            LoggerManager.logDebug(
              `Attempting to delete last history row (${lastRow}) matching today.`
            );
            historySheet.deleteRow(lastRow); // Delete the potentially incomplete row
            SpreadsheetApp.flush(); // Ensure deletion completes
            LoggerManager.logDebug(`Cleared last history row (${lastRow}).`);
          }
        } else {
          LoggerManager.logDebug("No history entry for today found to clear.");
        }
      } catch (histError) {
        // Log error during history check/delete but don't necessarily stop the entire UI setup
        LoggerManager.handleError(
          `Non-fatal error during history cleanup in initializeSetHabitUI: ${histError.message}`,
          false
        );
      }

      LoggerManager.logDebug("<<< Exiting initializeSetHabitUI successfully.");
    } catch (error) {
      // Catch errors from critical sheet operations (e.g., _getSheet, toggleColumns, resetChallengeDataUI)
      LoggerManager.handleError(
        `Error during initializeSetHabitUI: ${error.message}\n${error.stack}`,
        true
      );
      // Error re-thrown by handleError
    }
  },

  /**
   * Toggles the visibility, content, and formatting of the habit/property setter UI elements
   * (Reset Hour, Boost Interval, Set Habit checkbox and their labels/notes).
   * @private
   * @param {CellAction} action - Whether to `CellAction.SET` (display/configure) or `CellAction.CLEAR` (hide/reset) the fields.
   */
  _toggleSetterFields: function (action) {
    if (action !== CellAction.CLEAR && action !== CellAction.SET) {
      LoggerManager.handleError(
        `_toggleSetterFields called with unknown action: ${action}.`,
        true
      );
      return;
    }
    LoggerManager.logDebug(`Toggling setter fields UI: ${action}`);

    try {
      const sheet = MainSheetConfig._getSheet(); // Throws if sheet missing
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
        resetHourRange
          .clearContent()
          .clearNote()
          .setBackground(GlobalConfig.mainColor);
        boostIntervalRange
          .clearContent()
          .clearNote()
          .setBackground(GlobalConfig.mainColor);
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
        setHabitCellRange
          .clearContent()
          .clearDataValidations()
          .clearNote()
          .setBackground(GlobalConfig.mainColor);
        setHabitLabelRange
          .clearContent()
          .setBackground(GlobalConfig.mainColor)
          .setFontWeight("normal");
      }

      SpreadsheetApp.flush(); // Apply all UI changes together
    } catch (error) {
      LoggerManager.handleError(
        `Error during _toggleSetterFields (Action: ${action}): ${error.message}\n${error.stack}`,
        true
      );
    }
  },

  /**
   * Validates the user-entered values for reset hour and boost interval from the setter cells.
   * @private
   * @param {boolean} [alertUser=true] - Whether to show an alert message to the user on failure.
   * @returns {boolean} True if both reset hour (0-23 integer) and boost interval (>=1 integer) are valid.
   */
  _validateSetters: function (alertUser = true) {
    let isValid = true;
    try {
      const sheet = MainSheetConfig._getSheet();
      const resetHourValue = sheet
        .getRange(MainSheetConfig.setterRanges.resetHour)
        .getValue();
      const boostIntervalValue = sheet
        .getRange(MainSheetConfig.setterRanges.boostInterval)
        .getValue();

      // Validate Reset Hour (0-23 integer)
      if (
        !ValidationUtils._validateNonNegativeInteger(resetHourValue) ||
        resetHourValue > 23
      ) {
        LoggerManager.logDebug(
          `Invalid Reset Hour: ${resetHourValue}. Must be 0-23.`
        );
        isValid = false;
      }

      // Validate Boost Interval (>= 1 integer)
      if (!ValidationUtils._validatePositiveInteger(boostIntervalValue)) {
        LoggerManager.logDebug(
          `Invalid Boost Interval: ${boostIntervalValue}. Must be >= 1.`
        );
        isValid = false;
      }

      if (!isValid && alertUser) {
        Messages.showAlert(MessageTypes.INVALID_SETTERS);
      }
    } catch (error) {
      LoggerManager.handleError(
        `Error validating setters: ${error.message}`,
        false
      );
      isValid = false; // Consider validation failed if sheet access fails
      if (alertUser) {
        Messages.showAlert(MessageTypes.INVALID_SETTERS);
      } // Show generic message on error too
    }
    return isValid;
  },

  /**
   * Retrieves the current list of habits defined by emojis in the 'activities' column
   * on the main sheet, along with their corresponding row numbers.
   * @returns {Array<{emoji: string, row: number}>} List of habit objects {emoji: string, row: 1-based number}. Returns empty array on error.
   */
  getCurrentEmojiSpread: function () {
    let emojiList = [];
    try {
      const activityRange = MainSheetConfig.getActivityDataRange();
      if (!activityRange) {
        LoggerManager.logDebug(
          "getCurrentEmojiSpread: Activity range not found (likely no data rows yet)."
        );
        return [];
      }

      const sheet = MainSheetConfig._getSheet(); // Should be safe after getting range
      const firstRow = activityRange.getRow();
      const values = activityRange.getValues(); // 2D array [[val1], [val2], ...]

      values.forEach((cellValue, index) => {
        const value = cellValue[0];
        // Check if the cell content is a string and contains an emoji
        if (typeof value === "string" && this.emojiPattern.test(value)) {
          // Reset regex lastIndex for global flag 'g'
          this.emojiPattern.lastIndex = 0;
          const rowNum = firstRow + index; // Calculate 1-based row number
          emojiList.push({ emoji: value, row: rowNum });
        }
      });
      LoggerManager.logDebug(
        `getCurrentEmojiSpread found: ${emojiList.length} habits.`
      );
    } catch (e) {
      LoggerManager.handleError(
        `Error in getCurrentEmojiSpread: ${e.message}`,
        false
      );
      emojiList = []; // Return empty on error
    }
    return emojiList;
  },

  /**
   * Saves the current emoji spread (read from the sheet) to document properties as a JSON string.
   */
  updateEmojiSpreadProperty: function () {
    try {
      const currentSpread = this.getCurrentEmojiSpread();
      PropertyManager.setProperty(
        PropertyKeys.EMOJI_LIST,
        JSON.stringify(currentSpread)
      );
      LoggerManager.logDebug(
        `Updated emoji spread property with ${currentSpread.length} habits.`
      );
    } catch (e) {
      LoggerManager.handleError(
        `Error updating emoji spread property: ${e.message}`,
        false
      );
    }
  },

  /**
   * Processes the action when the 'Set Habit' checkbox is ticked by the user.
   * Validates current settings and habit definitions, confirms with the user,
   * and then finalizes the setup if confirmed.
   * Called by DataHandler.handleCellEdit.
   * @param {GoogleAppsScript.Spreadsheet.Range} checkboxRange - The range object of the 'Set Habit' checkbox cell.
   */
  processHabitSpreadConfirmation: function (checkboxRange) {
    LoggerManager.logDebug(
      ">>> Entering HabitManager.processHabitSpreadConfirmation"
    );
    try {
      // Ensure checkbox is actually checked (redundant check, but safe)
      if (checkboxRange.getValue() !== true) {
        LoggerManager.logDebug("Set Habit checkbox is not checked. Exiting.");
        return;
      }

      // 1. Validate Setters (Reset Hour, Boost Interval) - Show alert on failure
      if (!this._validateSetters(true)) {
        checkboxRange.setValue(false);
        SpreadsheetApp.flush();
        return;
      }

      // 2. Validate Habits (at least one emoji must be present)
      const currentEmojiSpread = this.getCurrentEmojiSpread();
      if (currentEmojiSpread.length === 0) {
        Messages.showAlert(MessageTypes.NO_HABITS_SET);
        checkboxRange.setValue(false);
        SpreadsheetApp.flush();
        return;
      }

      // 3. Ask for User Confirmation
      const response = Messages.showAlert(MessageTypes.CONFIRM_HABIT_SPREAD);

      // 4. Finalize or Reset Checkbox
      if (response === Messages.ButtonTypes.YES) {
        LoggerManager.logDebug("User confirmed habit spread. Finalizing...");
        this._finalizeHabitSpread(); // Attempt finalization
      } else {
        LoggerManager.logDebug("User cancelled habit spread confirmation.");
        checkboxRange.setValue(false);
        SpreadsheetApp.flush();
      }
    } catch (error) {
      // Catch errors during validation or finalization attempt
      LoggerManager.handleError(
        `Error during processHabitSpreadConfirmation: ${error.message}\n${error.stack}`,
        false
      );
      try {
        checkboxRange.setValue(false);
        SpreadsheetApp.flush();
      } catch (e) {} // Attempt to reset checkbox on error
      SpreadsheetApp.getUi().alert(
        "An error occurred during habit setup finalization. Please check logs."
      );
    }
    LoggerManager.logDebug(
      "<<< Exiting HabitManager.processHabitSpreadConfirmation"
    );
  },

  /**
   * Finalizes the habit setup process after user confirmation.
   * Transitions the application state from 'Habit Ideation' to 'Challenge'.
   * Saves final settings, updates UI, and prepares for active tracking.
   * @private
   * @throws {Error} If critical operations fail (e.g., sheet access, property setting).
   */
  _finalizeHabitSpread: function () {
    LoggerManager.logDebug(">>> Entering HabitManager._finalizeHabitSpread");
    try {
      const sheet = MainSheetConfig._getSheet(); // Throws if missing

      // 1. Save validated setters to properties
      const resetHourValue = sheet
        .getRange(MainSheetConfig.setterRanges.resetHour)
        .getValue();
      const boostIntervalValue = sheet
        .getRange(MainSheetConfig.setterRanges.boostInterval)
        .getValue();
      // Re-validate just before saving (belt-and-suspenders)
      if (!this._validateSetters(false)) {
        // Don't alert again, just error out
        LoggerManager.handleError(
          "Setter validation failed immediately before finalization.",
          true
        );
        return; // Stop finalization
      }
      PropertyManager.setProperty(
        PropertyKeys.RESET_HOUR,
        String(resetHourValue)
      );
      PropertyManager.setProperty(
        PropertyKeys.BOOST_INTERVAL,
        String(boostIntervalValue)
      );
      // NOTE: Trigger creation is NOT done here. It's handled by 'Start New Challenge' or manually.

      // 2. Save the validated emoji spread
      this.updateEmojiSpreadProperty();

      // 3. Set/Update first challenge date/row properties
      PropertyManager._updateFirstChallengeDateAndRow(); // Marks the official start

      // 4. Change application mode
      PropertyManager.setProperty(PropertyKeys.MODE, ModeTypes.CHALLENGE);

      // 5. Clean up Habit Ideation UI elements
      this._toggleSetterFields(CellAction.CLEAR);
      MainSheetConfig.toggleColumns(ColumnAction.SHOW); // Show tracking columns

      // 6. Optional: Format activity column background
      const activitiesRange = MainSheetConfig.getActivityDataRange();
      if (activitiesRange) {
        activitiesRange.setBackground(GlobalConfig.mainColor);
      }

      // 7. Insert completion checkboxes for the new habits
      MainSheetConfig.insertCompletionCheckboxes();

      // 8. Display today's data (creates the first history entry)
      PropertyManager.updateLastUpdateProperty(); // Ensure LAST_UPDATE is correct
      DataHandler.displayDate(DateManager.getTodayStr()); // Load today's view

      // 9. Save all property changes made during finalization
      PropertyManager.setDocumentProperties();

      LoggerManager.logDebug(
        "<<< Exiting HabitManager._finalizeHabitSpread successfully."
      );
    } catch (error) {
      // Make finalization errors fatal to avoid inconsistent state
      LoggerManager.handleError(
        `Error during _finalizeHabitSpread: ${error.message}\n${error.stack}`,
        true
      );
    }
  },

  /**
   * Checks if the current emoji spread on the sheet matches the one stored in properties.
   * If a mismatch is detected (indicating manual user edits in Challenge mode),
   * it alerts the user and reverts the sheet cells to match the stored configuration.
   * Called by DataHandler.getRelevantRows.
   * @returns {boolean} True if spread matches or was successfully reverted, false otherwise (e.g., revert failed).
   */
  checkEmojiSpread: function () {
    LoggerManager.logDebug(
      "Checking current emoji spread against stored property..."
    );
    let match = true; // Assume match initially
    try {
      const currentSpread = this.getCurrentEmojiSpread();
      const storedSpread = JSON.parse(
        PropertyManager.getProperty(PropertyKeys.EMOJI_LIST) || "[]"
      );

      // Compare spreads (length and element-wise check for emoji and row)
      if (currentSpread.length !== storedSpread.length) {
        match = false;
      } else {
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
          LoggerManager.handleError("Failed to revert emoji spread.", false);
          return false; // Indicate that revert failed, though mismatch was detected
        }
        // Revert successful
      } else {
        LoggerManager.logDebug("Emoji spread matches stored property.");
      }
      return true; // Return true if matched or if revert succeeded
    } catch (e) {
      LoggerManager.handleError(
        `Error during checkEmojiSpread: ${e.message}`,
        true
      ); // Make parsing errors fatal here
      return false; // Indicate check failed
    }
  },

  /**
   * Reverts the content and formatting of cells in the activity column on the main sheet
   * to exactly match the provided `storedSpread` configuration.
   * Clears cells that shouldn't have emojis and sets those that should.
   * @private
   * @param {Array<{emoji: string, row: number}>} storedSpread - The definitive habit configuration.
   * @returns {boolean} True if the revert operation completed successfully, false on error.
   */
  _revertEmojiSpread: function (storedSpread) {
    LoggerManager.logDebug(
      `Reverting emoji spread to stored config: ${JSON.stringify(storedSpread)}`
    );
    try {
      const sheet = MainSheetConfig._getSheet(); // Throws if missing
      const activityCol = MainSheetConfig.activityDataColumn;
      const activityRange = MainSheetConfig.getActivityDataRange(); // Full potential range

      // Determine the range of rows to check/clear
      const firstRowToCheck = MainSheetConfig.firstDataInputRow;
      const lastRowToCheck = activityRange
        ? activityRange.getRow() + activityRange.getNumRows() - 1
        : sheet.getLastRow(); // Use range end if possible

      LoggerManager.logDebug(
        `Reverting/clearing rows ${firstRowToCheck} to ${lastRowToCheck} in column ${activityCol}.`
      );

      // Create a map for quick lookup of stored emojis by row number
      const storedRowMap = new Map();
      storedSpread.forEach((item) => storedRowMap.set(item.row, item.emoji));

      // Prepare batch updates for efficiency
      const cellsToUpdate = {}; // { "D5": "🚀", "D7": "📚", ... }
      const cellsToClear = []; // [ "D6", "D8", ... ]

      // Iterate through all potentially relevant rows on the sheet
      for (let r = firstRowToCheck; r <= lastRowToCheck; r++) {
        const cellA1 = DataHandler._cellToA1Notation(r, activityCol);
        const expectedEmoji = storedRowMap.get(r);

        if (expectedEmoji) {
          // This row SHOULD have an emoji. Prepare to set it.
          cellsToUpdate[cellA1] = expectedEmoji;
        } else {
          // This row should NOT have an emoji. Prepare to clear it.
          cellsToClear.push(cellA1);
        }
      }

      // Perform batch operations
      if (cellsToClear.length > 0) {
        LoggerManager.logDebug(
          `Clearing ${cellsToClear.length} cells: ${cellsToClear
            .slice(0, 5)
            .join(",")}...`
        );
        sheet
          .getRangeList(cellsToClear)
          .clearContent()
          .setBackground(GlobalConfig.mainColor);
      }
      if (Object.keys(cellsToUpdate).length > 0) {
        LoggerManager.logDebug(
          `Setting values for ${Object.keys(cellsToUpdate).length} cells.`
        );
        // Setting individual values might be necessary if RangeList setValue behaves unexpectedly
        Object.entries(cellsToUpdate).forEach(([cellA1, value]) => {
          sheet
            .getRange(cellA1)
            .setValue(value)
            .setBackground(GlobalConfig.mainColor); // Ensure background consistency
        });
      }

      SpreadsheetApp.flush(); // Apply all revert changes
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
// Freeze the manager object to prevent accidental modification.
Object.freeze(HabitManager);
