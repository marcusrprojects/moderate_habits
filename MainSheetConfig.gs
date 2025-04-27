/**
 * @fileoverview Configuration and utility methods specific to the 'main' sheet.
 * Contains static properties for sheet layout (name, columns, cells, labels)
 * and methods for basic sheet interactions and UI management related to the
 * main tracking interface (getting ranges, toggling columns, managing protected cells).
 */

/** OnlyCurrentDoc */

/**
 * MainSheetConfig manages the configuration and UI-related operations for the "main" sheet.
 * It acts as a singleton, holding static configuration data and providing helper methods
 * for interacting with specific parts of the main sheet's layout.
 * @namespace MainSheetConfig
 */
const MainSheetConfig = {
  /** @constant {string} sheetName - The exact name of the main tracking sheet. */
  sheetName: "main",

  /** @constant {number} firstDataInputRow - The 1-indexed row number where the first habit entry appears. */
  firstDataInputRow: 3,
  /** @constant {number} activityDataColumn - The 1-indexed column number for habit activity (emoji/description). */
  activityDataColumn: 4, // Column D
  /** @constant {number} completionDataColumn - The 1-indexed column number for completion checkboxes. */
  completionDataColumn: 5, // Column E
  /** @constant {number} bufferDataColumn - The 1-indexed column number for displaying buffer days. */
  bufferDataColumn: 6, // Column F
  /** @constant {number} streaksDataColumn - The 1-indexed column number where streak values are displayed. */
  streaksDataColumn: 2, // Column B
  /** @constant {string} currentStreakCell - The A1 notation for the cell displaying the current streak. */
  currentStreakCell: "B3",
  /** @constant {string} highestStreakCell - The A1 notation for the cell displaying the highest streak. */
  highestStreakCell: "B6",
  /** @constant {string} dateCell - The A1 notation for the cell used as the date selector. */
  dateCell: "B9",
  /** @constant {number} defaultBuffer - The initial buffer days each habit receives at the start of a challenge. */
  defaultBuffer: 1,
  /** @constant {number} resetHourDefault - The default hour (0-23) for the daily reset trigger if not set by the user. */
  resetHourDefault: 3, // 3 AM

  /**
   * Header labels for key data points shown on the main sheet.
   * @constant {Object<string, string>}
   */
  headerLabels: {
    activities: "activities",
    completion: "completion",
    buffer: "buffer",
    currentStreak: "current streak",
    highestStreak: "highest streak",
    dateSelector: "date selector",
  },

  /**
   * Cell addresses (A1 notation) for the header labels.
   * Used for identifying protected cells and applying styling.
   * @constant {Object<string, string>}
   */
  headerLabelRanges: {
    activities: "D2",
    completion: "E2",
    buffer: "F2",
    currentStreak: "B2",
    highestStreak: "B5",
    dateSelector: "B8",
  },

  /**
   * Cell addresses (A1 notation) for user settings shown during Habit Ideation mode.
   * @constant {Object<string, string>}
   */
  setterRanges: {
    setHabit: "H3",
    resetHour: "H6",
    boostInterval: "H8",
  },

  /**
   * Cell addresses (A1 notation) for the labels corresponding to user settings.
   * @constant {Object<string, string>}
   */
  setterLabelRanges: {
    setHabit: "H2",
    resetHour: "H5",
    boostInterval: "H7",
  },

  /**
   * Display text for the user setting labels shown during Habit Ideation mode.
   * @constant {Object<string, string>}
   */
  setterLabels: {
    setHabit: "set habits",
    resetHour: "reset hour",
    boostInterval: "boost interval",
  },

  /**
   * Explanatory notes added to the user setting cells during Habit Ideation mode.
   * @constant {Object<string, string>}
   */
  setterNotes: {
    resetHour:
      "Define the hour (0-23) at which the daily reset occurs. Default is 3 (3 A.M.). Whole numbers only.",
    boostInterval:
      "Define the interval (in days) for earning buffer day boosts. Default is 7 (+1 buffer day per habit weekly). Must be 1 or greater.",
    setHabit:
      "Check this box AFTER defining your habits (emojis) in the 'activities' column to finalize setup and start the challenge. Data will be validated.",
  },

  // --- Private Cached Sheet Object ---
  /**
   * Cached Sheet object to avoid repeated calls to getSheetByName.
   * @private
   * @type {GoogleAppsScript.Spreadsheet.Sheet | null}
   */
  _sheet: null,

  // --- Basic Sheet Accessors ---

  /**
   * Retrieves the Sheet object for the relevant sheet (main or history) by iterating
   * through all sheets. This method was chosen for reliability over getSheetByName,
   * which exhibited issues in certain contexts after sheet copying.
   *
   * @private
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
   * @throws {Error} if the sheet cannot be found.
   */
  _getSheet: function () {
    const sheetNameToFind = this.sheetName; // e.g., "main" or "history"
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();

    // Critical check for active spreadsheet context
    if (!activeSS) {
      // If this happens, something is fundamentally wrong with the script environment
      LoggerManager.handleError(
        `FATAL: Cannot get Active Spreadsheet context in _getSheet for ${sheetNameToFind}.`,
        true
      );
      return null; // Should be unreachable due to throw
    }
    const activeSSName = activeSS.getName();
    LoggerManager.logDebug(
      `_getSheet: Locating sheet "${sheetNameToFind}" in SS: "${activeSSName}" via iteration.`
    );

    // Iterate through sheets to find the one with the matching name
    const allSheets = activeSS.getSheets();
    let foundSheet = null;
    for (let i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getName() === sheetNameToFind) {
        foundSheet = allSheets[i];
        break; // Stop searching once found
      }
    }

    // Handle case where sheet is not found even after iterating
    if (!foundSheet) {
      const allSheetNames = allSheets.map((s) => `"${s.getName()}"`); // Get names for error message
      LoggerManager.handleError(
        `Sheet "${sheetNameToFind}" was not found via ITERATION in Spreadsheet "${activeSSName}". Available sheets: [${allSheetNames.join(
          ", "
        )}]. Application cannot function.`,
        true
      );
      return null; // Error thrown, unreachable
    }

    LoggerManager.logDebug(
      `_getSheet: Successfully found sheet "${sheetNameToFind}" by iterating.`
    );
    // Skip caching when using iteration; return the found sheet directly.
    // this._sheet = foundSheet; // Caching bypassed
    return foundSheet;
  },

  /**
   * Retrieves the value from a specific cell in the main sheet.
   * @private
   */
  _getSheetValue: function (cellA1Notation) {
    try {
      const sheet = this._getSheet();
      return sheet.getRange(cellA1Notation).getValue();
    } catch (e) {
      LoggerManager.handleError(
        `Failed to get value from cell ${cellA1Notation} on sheet ${this.sheetName}: ${e.message}`,
        false
      );
      return null;
    }
  },

  /**
   * Sets the value in a specific cell in the main sheet.
   * @private
   */
  _setSheetValue: function (cellA1Notation, value) {
    try {
      const sheet = this._getSheet();
      sheet.getRange(cellA1Notation).setValue(value);
      LoggerManager.logDebug(
        `Value set for cell ${cellA1Notation} on sheet ${this.sheetName}.`
      );
      return true;
    } catch (e) {
      LoggerManager.handleError(
        `Failed to set value for cell ${cellA1Notation} on sheet ${this.sheetName}: ${e.message}`,
        false
      );
      return false;
    }
  },

  /**
   * Retrieves a dynamic range for a given column, from the first data row down
   * to the last row containing any content in the sheet.
   * @private
   * @param {number} columnIndex - The 1-indexed column number.
   * @returns {GoogleAppsScript.Spreadsheet.Range | null} The range object, or null if no data rows exist or on error.
   */
  _getDynamicColumnRange: function (columnIndex) {
    try {
      const sheet = this._getSheet();
      const lastRow = sheet.getLastRow();
      const firstRow = this.firstDataInputRow;
      if (lastRow < firstRow) {
        return null;
      }
      const numRows = lastRow - firstRow + 1;
      return sheet.getRange(firstRow, columnIndex, numRows, 1);
    } catch (e) {
      LoggerManager.handleError(
        `Failed to get dynamic range for column ${columnIndex} on sheet ${this.sheetName}: ${e.message}`,
        false
      );
      return null;
    }
  },

  /**
   * Retrieves the Range object for the completion data column, covering all potential data rows.
   * @returns {GoogleAppsScript.Spreadsheet.Range | null} The range object or null on error / no data rows.
   */
  getCompletionDataRange: function () {
    return this._getDynamicColumnRange(this.completionDataColumn);
  },

  /**
   * Retrieves the Range object for the buffer data column, covering all potential data rows.
   * @returns {GoogleAppsScript.Spreadsheet.Range | null} The range object or null on error / no data rows.
   */
  getBufferDataRange: function () {
    return this._getDynamicColumnRange(this.bufferDataColumn);
  },

  /**
   * Retrieves the Range object for the activity data column, covering all potential data rows.
   * @returns {GoogleAppsScript.Spreadsheet.Range | null} The range object or null on error / no data rows.
   */
  getActivityDataRange: function () {
    return this._getDynamicColumnRange(this.activityDataColumn);
  },

  // --- UI and Edit Handling ---

  /**
   * Checks if the provided range intersects with any cells or columns considered "locked"
   * based on the current application mode. Locked areas cannot be directly edited by the user.
   *
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object representing the edited cell(s).
   * @returns {boolean} True if the range intersects with locked areas, false otherwise.
   */
  includesLockedRange: function (range) {
    // Basic safety check
    if (!range) return false;

    try {
      const sheet = this._getSheet(); // Ensure sheet context is correct
      if (range.getSheet().getSheetId() !== sheet.getSheetId()) {
        return false; // Edit is not on the main sheet
      }

      const columns = DataHandler.getRangeColumns(range); // Use helper
      const rangeCells = DataHandler.getRangeCells(range); // Use helper

      // Define base locked ranges/cells (always locked)
      let lockedRanges = [
        this.currentStreakCell,
        this.highestStreakCell,
        ...Object.values(this.headerLabelRanges),
      ];
      // Define base locked columns (always locked)
      let lockedColumns = [this.bufferDataColumn]; // Buffer display is calculated

      const mode = PropertyManager.getProperty(PropertyKeys.MODE);

      // Add mode-specific locking rules
      if (mode === ModeTypes.HABIT_IDEATION) {
        // Lock setter *labels*, completion/buffer/streak *display* columns during setup
        lockedRanges.push(...Object.values(this.setterLabelRanges));
        lockedColumns.push(
          this.completionDataColumn,
          this.bufferDataColumn,
          this.streaksDataColumn
        );
      } else if (mode === ModeTypes.TERMINATED) {
        // In terminated mode, effectively all grid edits are locked.
        // Menu actions are still allowed.
        LoggerManager.logDebug(
          "includesLockedRange: Mode is TERMINATED, locking edit."
        );
        return true;
      }
      // No additional rules needed for CHALLENGE mode currently (beyond base rules)

      // Remove duplicates for clarity
      lockedColumns = [...new Set(lockedColumns)];

      // Check for intersection
      const isInLockedColumns = columns.some((column) =>
        lockedColumns.includes(column)
      );
      const isInLockedRange = rangeCells.some((cell) =>
        lockedRanges.includes(cell)
      );

      const isLocked = isInLockedRange || isInLockedColumns;
      if (isLocked) {
        LoggerManager.logDebug(
          `includesLockedRange: Edit in locked area. Mode='${mode}', Columns=${columns}, Cells=${rangeCells}. Result=true`
        );
      }
      return isLocked;
    } catch (e) {
      LoggerManager.handleError(
        `Error in includesLockedRange: ${e.message}`,
        false
      );
      return true; // Fail safe: assume locked if error occurs during check
    }
  },

  /**
   * Resets edited cells within a given range to their previous state or to their
   * defined label configuration if they are protected label cells.
   * This prevents users from overriding calculated values or static labels.
   *
   * @param {GoogleAppsScript.Spreadsheet.Range} range - The range object representing the cell(s) that were edited.
   * @param {*} oldValue - The previous value of the edited cell(s). Can be unreliable for multi-cell edits (often `undefined`).
   */
  maintainCellValue: function (range, oldValue) {
    // Basic safety check
    if (!range) return;

    try {
      const sheet = this._getSheet();
      const rangeCells = DataHandler.getRangeCells(range); // Use helper
      LoggerManager.logDebug(
        `maintainCellValue: Restoring range ${range.getA1Notation()} (${
          rangeCells.length
        } cells). oldValue type: ${typeof oldValue}`
      );

      const isHabitIdeationMode =
        PropertyManager.getProperty(PropertyKeys.MODE) ===
        ModeTypes.HABIT_IDEATION;

      // Define label configurations to check against (add more if needed)
      const labelConfigs = [
        {
          ranges: this.headerLabelRanges,
          labels: this.headerLabels,
          color: GlobalConfig.secondaryColor,
          weight: "bold",
        },
      ];
      if (isHabitIdeationMode) {
        labelConfigs.push({
          ranges: this.setterLabelRanges,
          labels: this.setterLabels,
          color: GlobalConfig.secondaryColor,
          weight: "bold",
        });
      }

      const cellsToRestoreLabels = {}; // Store { a1Notation: { value: '...', color: '...', weight: '...' } }

      // Identify which edited cells are labels
      rangeCells.forEach((cellA1) => {
        labelConfigs.forEach((config) => {
          Object.entries(config.ranges).forEach(([key, labelRangeA1]) => {
            if (cellA1 === labelRangeA1) {
              cellsToRestoreLabels[cellA1] = {
                value: config.labels[key],
                color: config.color,
                weight: config.weight,
              };
            }
          });
        });
      });

      // Restore labels first
      Object.entries(cellsToRestoreLabels).forEach(([cellA1, properties]) => {
        try {
          const cellRange = sheet.getRange(cellA1);
          LoggerManager.logDebug(
            `Resetting label in cell ${cellA1} to value: '${properties.value}'`
          );
          cellRange
            .setValue(properties.value)
            .setFontWeight(properties.weight)
            .setBackground(properties.color);
        } catch (e) {
          LoggerManager.handleError(
            `Failed to reset label for cell ${cellA1}: ${e.message}`,
            false
          );
        }
      });

      // Handle non-label cells based on edit type
      const nonLabelCells = rangeCells.filter(
        (cellA1) => !cellsToRestoreLabels[cellA1]
      );

      if (nonLabelCells.length > 0) {
        if (rangeCells.length === 1 && oldValue !== undefined) {
          // Simple case: single non-label cell edit, restore known oldValue
          try {
            LoggerManager.logDebug(
              `Resetting single cell ${range.getA1Notation()} to previous value: ${oldValue}`
            );
            range.setValue(oldValue);
            range.setBackground(GlobalConfig.mainColor); // Ensure default background
          } catch (e) {
            LoggerManager.handleError(
              `Failed to reset single cell ${range.getA1Notation()} to old value: ${
                e.message
              }`,
              false
            );
          }
        } else {
          // Multi-cell edit OR single cell edit where oldValue is undefined.
          // Safest reliable action is to clear these unexpected/protected edits.
          Messages.showAlert(MessageTypes.UNDEFINED_CELL_CHANGES); // Inform user about revert
          LoggerManager.logDebug(
            `Multi-cell edit involving non-labels or unknown oldValue. Clearing affected non-label cells: ${nonLabelCells.join(
              ", "
            )}.`
          );
          nonLabelCells.forEach((cellA1) => {
            try {
              const cellRange = sheet.getRange(cellA1);
              cellRange.clearContent(); // Clear the content
              cellRange.setBackground(GlobalConfig.mainColor); // Reset background
            } catch (e) {
              LoggerManager.handleError(
                `Failed to clear potentially modified cell ${cellA1}: ${e.message}`,
                false
              );
            }
          });
        }
      } else if (Object.keys(cellsToRestoreLabels).length > 0) {
        LoggerManager.logDebug(
          `Label(s) reset successfully in range ${range.getA1Notation()}.`
        );
      }

      SpreadsheetApp.flush(); // Flush changes made by maintainCellValue
    } catch (e) {
      LoggerManager.handleError(
        `Error during maintainCellValue for range ${
          range ? range.getA1Notation() : "unknown"
        }: ${e.message}`,
        false
      );
    }
  },

  /**
   * Toggles the visibility of core tracking columns (completion, buffer) and the streak display area.
   * @param {ColumnAction} action - Whether to `ColumnAction.SHOW` or `ColumnAction.HIDE` the columns.
   */
  toggleColumns: function (action) {
    if (action !== ColumnAction.SHOW && action !== ColumnAction.HIDE) {
      LoggerManager.handleError(
        `toggleColumns called with unknown action: ${action}.`,
        true
      );
      return;
    }

    try {
      const sheet = this._getSheet();

      // Define the columns/groups to toggle
      const completionCol = this.completionDataColumn; // E.g., 5
      const bufferCol = this.bufferDataColumn; // E.g., 6
      const firstCoreCol = completionCol;
      const numCoreCols = bufferCol - firstCoreCol + 1; // Columns E, F -> count = 2

      // Streaks are in column B, labels might be adjacent. Let's toggle B and C?
      // Assuming streaks display is contained within columns B & C
      const streakColGroupStart = this.streaksDataColumn; // Column B (index 2)
      const streakColGroupCount = 2; // Toggle columns B and C

      if (action === ColumnAction.SHOW) {
        sheet.showColumns(firstCoreCol, numCoreCols);
        sheet.showColumns(streakColGroupStart, streakColGroupCount);
        LoggerManager.logDebug(
          `Shown columns ${firstCoreCol}-${
            firstCoreCol + numCoreCols - 1
          } and ${streakColGroupStart}-${
            streakColGroupStart + streakColGroupCount - 1
          }`
        );
      } else {
        // HIDE
        sheet.hideColumns(firstCoreCol, numCoreCols);
        sheet.hideColumns(streakColGroupStart, streakColGroupCount);
        LoggerManager.logDebug(
          `Hidden columns ${firstCoreCol}-${
            firstCoreCol + numCoreCols - 1
          } and ${streakColGroupStart}-${
            streakColGroupStart + streakColGroupCount - 1
          }`
        );
      }
    } catch (e) {
      LoggerManager.handleError(
        `Failed to toggle columns with action ${action}: ${e.message}`,
        false
      );
    }
  },

  /**
   * Inserts checkboxes into the completion data column for all currently relevant habit rows.
   * Retrieves relevant rows from DataHandler.
   */
  insertCompletionCheckboxes: function () {
    try {
      const sheet = this._getSheet();
      const relevantRows = DataHandler.getRelevantRows(); // Get current list of habit rows

      if (!relevantRows || relevantRows.length === 0) {
        LoggerManager.logDebug(
          "insertCompletionCheckboxes: No relevant rows found."
        );
        return;
      }

      LoggerManager.logDebug(
        `Inserting checkboxes for rows: ${relevantRows.join(", ")} in column ${
          this.completionDataColumn
        }`
      );
      relevantRows.forEach((row) => {
        try {
          // Get the specific cell range for the checkbox
          const range = sheet.getRange(row, this.completionDataColumn);
          // Using insertCheckboxes() handles existing checkboxes gracefully (it replaces/re-inserts).
          range.insertCheckboxes();
          // Optional: Ensure unchecked state if cell was previously empty/non-boolean
          // if (range.getValue() === '' || typeof range.getValue() !== 'boolean') {
          //    range.setValue(false);
          // }
        } catch (e) {
          // Log error for specific row but continue with others
          LoggerManager.handleError(
            `Failed to insert checkbox in row ${row}, column ${this.completionDataColumn}: ${e.message}`,
            false
          );
        }
      });
      // Consider SpreadsheetApp.flush() if immediate interaction is expected, but often not necessary here.
    } catch (e) {
      // Catch errors from _getSheet or getRelevantRows
      LoggerManager.handleError(
        `Error during insertCompletionCheckboxes setup: ${e.message}`,
        false
      );
    }
  },

  /**
   * Resets core challenge data display fields on the main sheet UI.
   * Clears completion/buffer columns, sets streaks to 0, and sets date to today.
   * Primarily used when starting a new challenge or terminating.
   */
  resetChallengeDataUI: function () {
    try {
      const sheet = this._getSheet();
      LoggerManager.logDebug(
        `Resetting challenge data UI on sheet ${this.sheetName}.`
      );

      const completionRange = this.getCompletionDataRange();
      const bufferRange = this.getBufferDataRange();

      // Prepare lists of ranges for batch operations
      const rangesToClearContent = [];
      const rangesToClearValidation = [];

      if (completionRange) {
        rangesToClearContent.push(completionRange.getA1Notation());
        rangesToClearValidation.push(completionRange.getA1Notation()); // Clear checkbox validation
      }
      if (bufferRange) {
        rangesToClearContent.push(bufferRange.getA1Notation());
      }

      // Perform clear operations
      if (rangesToClearValidation.length > 0) {
        sheet.getRangeList(rangesToClearValidation).clearDataValidations();
        LoggerManager.logDebug(
          `Cleared data validations for: ${rangesToClearValidation.join(", ")}`
        );
      }
      if (rangesToClearContent.length > 0) {
        sheet.getRangeList(rangesToClearContent).clearContent();
        LoggerManager.logDebug(
          `Cleared content for: ${rangesToClearContent.join(", ")}`
        );
      }

      // Set default values for streaks and date (individual calls are fine here)
      const todayStr = DateManager.getTodayStr();
      const defaultStreak = 0;
      this._setSheetValue(this.dateCell, todayStr);
      this._setSheetValue(this.currentStreakCell, defaultStreak);
      this._setSheetValue(this.highestStreakCell, defaultStreak);

      SpreadsheetApp.flush(); // Flush all reset changes together
      LoggerManager.logDebug(`Challenge data UI reset successfully.`);
    } catch (e) {
      LoggerManager.handleError(
        `Error during challenge data UI reset: ${e.message}`,
        true
      );
    }
  },
};

// Freeze the configuration object to prevent modification at runtime.
Object.freeze(MainSheetConfig);
