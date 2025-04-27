/**
 * @fileoverview Defines various enumerations (enums) used throughout the application.
 * These provide consistent keys and values for application modes, property keys,
 * message types, UI actions, etc., improving code readability and maintainability.
 */

/** OnlyCurrentDoc */

/**
 * Keys used for accessing script and document properties via PropertiesService.
 * @enum {string}
 * @readonly
 */
const PropertyKeys = {
  MODE: "mode",
  LAST_DATE_SELECTOR_UPDATE: "lastDateSelectorUpdate",
  LAST_COMPLETION_UPDATE: "lastCompletionUpdate",
  LAST_UPDATE: "lastUpdate",
  ACTIVITIES_COLUMN_UPDATED: "activitiesColumnUpdated",
  FIRST_CHALLENGE_DATE: "firstChallengeDate",
  FIRST_CHALLENGE_ROW: "firstChallengeRow",
  EMOJI_LIST: "emojiList",
  RESET_HOUR: "resetHour",
  BOOST_INTERVAL: "boostInterval",
};

/**
 * Standard boolean values represented as strings, primarily for storage in PropertiesService.
 * @enum {string}
 * @readonly
 */
const BooleanTypes = {
  TRUE: "true",
  FALSE: "false",
};

/**
 * Various operational modes of the application, dictating UI and behavior.
 * @enum {string}
 * @readonly
 */
const ModeTypes = {
  CHALLENGE: "challenge", // Normal habit tracking mode.
  HABIT_IDEATION: "habitIdeation", // Mode for setting up habits initially or resetting.
  TERMINATED: "terminated", // Mode when habit tracking is intentionally stopped.
};

/**
 * Types used to track the most recent significant update action (either completion or date change).
 * This helps determine whether to save data before loading a new date.
 * @enum {string}
 * @readonly
 */
const LastUpdateTypes = {
  COMPLETION: "completion",
  DATE_SELECTOR: "dateSelector",
};

/**
 * Keys identifying different types of standard messages or alerts shown to the user via the UI.
 * Used by the Messages manager.
 * @enum {string}
 * @readonly
 */
const MessageTypes = {
  TERMINATION_CONFIRMATION: "terminationConfirmation",
  TERMINATED: "terminated",
  TERMINATION_CANCELLED: "terminationCancelled",
  TERMINATION_REMINDER: "terminationReminder",
  INVALID_DATE: "invalidDate",
  CHALLENGE_RESET: "challengeReset",
  CHALLENGE_CANCELLED: "challengeCancelled",
  HABIT_SPREAD_RESET: "habitSpreadReset",
  CONFIRM_HABIT_SPREAD: "confirmHabitSpread",
  START_NEW_CHALLENGE: "startNewChallenge",
  WELCOME_MESSAGE: "welcomeMessage",
  NO_HABITS_SET: "noHabitsSet",
  INVALID_SETTERS: "invalidSetters",
  NEW_VERSION_AVAILABLE: "newVersionAvailable",
  NO_NEW_UPDATES: "noNewUpdates",
  UNDEFINED_CELL_CHANGES: "undefinedCellChanges",
  DATA_PARSE_ERROR: "dataParseError", // Added for JSON parsing errors in history
};

/**
 * Constants defining actions for showing or hiding sheet columns.
 * @enum {string}
 * @readonly
 */
const ColumnAction = {
  HIDE: "hide",
  SHOW: "show",
};

/**
 * Constants defining actions for clearing or setting cell content, validation, or notes.
 * @enum {string}
 * @readonly
 */
const CellAction = {
  CLEAR: "clear",
  SET: "set",
};

// Freeze enums to prevent modification at runtime.
Object.freeze(PropertyKeys);
Object.freeze(BooleanTypes);
Object.freeze(ModeTypes);
Object.freeze(LastUpdateTypes);
Object.freeze(MessageTypes);
Object.freeze(ColumnAction);
Object.freeze(CellAction);
