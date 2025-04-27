# Moderate Habits - Google Sheet Habit Tracker

Moderate Habits is a habit-tracking application built entirely within Google Sheets, powered by Google Apps Script. It helps you build consistency by tracking daily habits, managing streaks, and providing flexible "buffer" days so missing a single day doesn't break your momentum.

It's designed for those who want a simple, customizable, and free way to monitor their progress directly in a familiar spreadsheet environment.

[![Moderate Habits Video Tutorial Thumbnail](./assets/moderate-habits-thumbnail.png)](https://tinyurl.com/moderate-habits-tutorial)
*(Click image to watch setup tutorial)*

## ✨ Features

* **Daily Habit Tracking:** Simple checkbox interface for marking daily completion.
* **Customizable Habits:** Define your habits using emojis in the `activities` column.
* **Streak Calculation:** Automatically tracks current and highest streaks based on completion.
* **Buffer Days:** Earn buffer days (rest days) for each habit, preventing streak resets on occasional misses.
* **Configurable Settings:**
  * Set the daily reset hour (e.g., 3 AM).
  * Configure the interval for earning buffer day boosts.
* **History Log:** Automatically logs daily completion, buffer status, and streaks to a separate 'history' sheet.
* **Date Navigation:** View and update data for past days within the current challenge using a date selector.
* **Dynamic Help:** Context-aware help sidebar explaining features based on the current sheet and mode.
* **First-Run Setup:** Guided setup process for new users via the "Start New Challenge" menu.
* **Update Checker:** Option to check if a newer version of the script is available.
* **Termination Mode:** Option to formally stop tracking while preserving history.

## ⚙️ How It Works

* **`main` Sheet:** Your primary interface for daily tracking and viewing current status.
* **`history` Sheet:** A read-only log storing the data for each day. All calculations are based on this sheet.
* **Google Apps Script:** The backend logic handling calculations, data persistence, UI, and automation.
* **Access Request:** The core script logic is contained in a private library. You need to request access so your copy of the sheet can utilize this library.

## 🚀 Setup Instructions

**IMPORTANT:** This tool uses a private Google Apps Script library. You must request access *first* before you can use the template sheet.

1. **Prerequisites:** You need a Google Account.
2. **Request Access:**
    * Fill out the access request form:
    * **➡️ [Google Form for Access Requests Here](https://tinyurl.com/moderate-habits-form) ⬅️**
    * Wait for the approval email titled **'Access Granted to "moderate habits"'** (check spam). This email contains links to the template and guide.
3. **Receive Links & Copy the Sheet:**
    * The approval email contains important links:
        * A link to the **Google Sheet Template**.
        * A link to the **User Guide**.
    * Click the **Google Sheet Template** link from the email to open it.
    * Go to **File > Make a copy**.
    * Save the copy to your own Google Drive. *(Rename your copied sheet if desired)*.
4. **Open Your Copy:** Open the Google Sheet *you just copied* to your Drive.
5. **Run Initial Setup / Start Challenge:**
    * Wait a few seconds for the custom menu "**Moderate Habits Settings**" to appear.
    * Click **Moderate Habits Settings > Start New Challenge**. *(This menu item handles both first-time setup and resetting).*
6. **Authorize the Script:**
    * You will likely be asked for authorization the first time you run a menu item. Click **Continue** / **OK**.
    * Choose the **correct Google Account** (the one granted access, if applicable).
    * You may see a "**Google hasn't verified this app**" screen. This is normal for custom scripts. Click "**Advanced**", then "**Go to [Your Sheet Name] (unsafe)**".
    * Review the permissions requested. It needs access to the current spreadsheet, ability to run triggers, display UI, and potentially connect to external services (for version check).
    * Click "**Allow**".
    * *Note: The script only runs within this Google Sheet document and only accesses its data.*
7. **Set Up Habits:**
    * Follow the prompts (Welcome message, then Challenge Reset message).
    * You are now in **Habit Ideation** mode (tracking columns hidden, settings visible).
    * Go to the `main` sheet.
    * In the `activities` column (Column D), enter your desired habits using **emojis**. Only emoji cells are tracked. Add text labels in other rows if desired.
    * *(Optional)* Configure `reset hour` (H6) and `boost interval` (H8).
    * Once ready, **check the `set habits` checkbox** (H3).
    * Confirm "Yes" when prompted.
8. **Start Tracking!** The sheet will switch to Challenge Mode.

## 💡 Usage Guide

* **Daily Tracking:** On the `main` sheet, check boxes in the `completion` column for the date shown in the `date selector` (B9).
* **Changing Dates:** Use the `date selector` (B9) to view/edit past data. Data saves automatically when changing dates *after* modifying completion status.
* **Streaks & Buffers:** Calculated automatically based on `history`. Missing a habit reduces its buffer the *next* day. Buffer = 0 + Miss = Streak Reset. Buffers increase based on `boost interval`.
* **Help:** Click **Moderate Habits Settings > Show Help (Current Page)** for the sidebar.
* **Detailed Guide:** Refer to the guide linked in the approval email or here:
  * **➡️ [Guide Link](http://tinyurl.com/moderate-habits) ⬅️**

## 🔧 Troubleshooting

* **Menu Not Appearing:** Wait up to 30s on first open, or reload.
* **Haven't Received Approval Email:** Please allow some time for manual processing of your access request. Remember to check your spam/junk folder for the email titled 'Access Granted to "moderate habits"'.
* **Authorization Errors/Loops:** Ensure you used the correct Google Account. Try removing permissions via Google Account settings (`Security > Third-party apps`) and re-authorizing using a menu item. Ensure `appsscript.json` scopes are correct if modifying the script.
* **Script Errors ("An error occurred..."):** Check **Extensions > Apps Script > Executions** for detailed error messages.
* **Performance:** Allow a few seconds for processing, especially with long history.

## 🙏 Contributing

Issues and feature requests are welcome! Please check the [GitHub repository issues page](https://github.com/marcusrprojects/moderate_habits/issues).

---

Happy habiting! 💪
