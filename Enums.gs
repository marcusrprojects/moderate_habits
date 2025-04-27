/**
 * @type {Object} PropertyKeys - Keys used for accessing various properties in the app.
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
 * @type {Object} BooleanTypes - Constants for true and false values.
 */
const BooleanTypes = {
  TRUE: "true",
  FALSE: "false",
};

/**
 * @type {Object} ModeTypes - Various operational modes of the app.
 */
const ModeTypes = {
  CHALLENGE: "challenge", // Normal habit tracking mode
  HABIT_IDEATION: "habitIdeation", // When you're setting up habits
  TERMINATED: "terminated", // When habits have been stopped
};

/**
 * @type {Object} LastUpdateTypes - Types of updates to track in the app.
 */
const LastUpdateTypes = {
  COMPLETION: "completion",
  DATE_SELECTOR: "dateSelector",
};

/**
 * @type {Object} MessageTypes - Constants for different types of messages displayed in the app.
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
};

/**
 * @type {Object} ColumnAction - Constants for actions taken on columns.
 */
const ColumnAction = {
  HIDE: "hide",
  SHOW: "show",
};

/**
 * @type {Object} CellAction - Constants for actions taken on cells.
 */
const CellAction = {
  CLEAR: "clear",
  SET: "set",
};
