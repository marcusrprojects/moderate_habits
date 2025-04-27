/**
 * @fileoverview Defines global configuration constants, like UI colors.
 * This ensures consistency in styling across the application.
 */

/** OnlyCurrentDoc */

/**
 * Global configuration constants, primarily related to UI styling.
 * This object acts as a singleton, frozen to prevent modification.
 * @namespace GlobalConfig
 */
const GlobalConfig = {
  /**
   * Primary background color, a very light cream shade.
   * Used for general sheet background.
   * @constant {string}
   */
  mainColor: "#FFF9F5",

  /**
   * Secondary color, a light peach hue.
   * Often used for headers, labels, or emphasis.
   * @constant {string}
   */
  secondaryColor: "#F2DFCE",
};

// Freeze the configuration object to prevent modification at runtime.
Object.freeze(GlobalConfig);
