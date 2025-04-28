# Moderate Habits - Google Sheet Habit Tracker

![Version](https://img.shields.io/github/v/release/marcusrprojects/moderate_habits)

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
* **First-Run Setup:** Guided setup process for new users.
* **Update Checker:** Option to check if a newer version of the script is available.
* **Termination Mode:** Option to formally stop tracking while preserving history.

## ⚙️ How It Works

* **`main` Sheet:** Your primary interface for daily tracking and viewing current status.
* **`history` Sheet:** A read-only log storing the data for each day. All calculations are based on this sheet.
* **Google Apps Script:** The backend logic handling:
  * Daily resets (via time-driven trigger).
  * Calculations (streaks, buffers).
  * Data saving and propagation between `main` and `history`.
  * UI updates (help sidebar, dialogs).
  * Event handling (`onOpen`, `onEdit`).
* **Access Request:** The core script logic is contained in a private library. You need to request access so your copy of the sheet can utilize this library.

## 🚀 Setup Instructions

**IMPORTANT:** This tool uses a private Google Apps Script library. You must request access *first* before you can use the template sheet.

1. **Prerequisites:** You need a Google Account.
2. **Request Access:**
    * Fill out the access request form:
    * **➡️ [Link to your Google Form for Access Requests Here](https://tinyurl.com/moderate-habits-form) ⬅️**
    * This step is necessary to grant your Google Account permission to use the underlying script library.
3. **Wait for Approval:**
    * Your request needs to be processed manually. Please allow some time for approval.
    * You will receive an email titled **'Access Granted to "moderate habits"'** once your request is approved.
    * ***Check your spam folder*** if you don't see the email within a reasonable time.
4. **Receive Links & Copy the Sheet:**
    * The approval email contains important links:
        * A link to the **Google Sheet Template**.
        * A link to the **User Guide (PDF)**.
    * Click the **Google Sheet Template** link from the email to open it.
    * Go to **File > Make a copy**.
    * Save the copy to your own Google Drive. *(Rename your copied sheet if desired)*.
5. **Open Your Copy:** Open the Google Sheet *you just copied* to your Drive.
6. **Run Initial Setup:**
    * Wait a few seconds for the custom menu "**Moderate Habits Settings**" to appear at the top. *(It might take up to 30 seconds on the very first open).*
    * Click **Moderate Habits Settings > \*\*Begin\*\***.
7. **Authorize the Script:**
    * You will be asked for authorization for the script *within your copied sheet*. Click **Continue** / **OK**.
    * Choose the **same Google Account** you used to request access.
    * You'll likely see a "**Google hasn't verified this app**" screen. This is normal and expected for custom scripts you copy.
    * Click "**Advanced**" (usually a small link).
    * Click "**Go to [Your Sheet Name] (unsafe)**".
    * Review the permissions. The script needs access *to this specific spreadsheet* to function.
    * Click "**Allow**".
    * *Note: The script only runs within this Google Sheet document and only accesses its data.*
8. **Set Up Habits:**
    * Follow the prompts. A welcome message will appear, then a setup reminder.
    * You are now in **Habit Ideation** mode.
    * Go to the `main` sheet.
    * In the `activities` column (Column D), enter your desired habits. **Use one or more emojis** in the cells you want tracked. You can also add non-emoji text for labels or section headers (these won't be tracked).
    * *(Optional)* Add detailed descriptions to your emoji cells using Right-Click > `Insert Note`.
    * *(Optional)* Configure the `reset hour` (0-23) and `boost interval` (>=1) in the settings cells that appear on the right (Cells H6 and H8). Hover over them for explanatory notes.
    * Once you're happy with your habits (emojis) and settings, **check the `set habits` checkbox** (Cell H3).
    * A confirmation dialog will appear. Click **Yes** to start the challenge.
9. **Start Tracking!** Your sheet is now ready for tracking. The setup cells will disappear, the tracking columns will appear, and today's date will be loaded.

## 💡 Usage Guide

* **Daily Tracking:** On the `main` sheet, check the boxes in the `completion` column for habits you've completed *for the date shown in the `date selector` cell*.
* **Changing Dates:** Use the `date selector` cell (B9) to view or edit past days. Double-click for a calendar or type a date (YYYY-MM-DD). Data for the previously viewed date is saved automatically when you change dates *if* you modified completion status.
* **Streaks & Buffers:** These are calculated automatically based on the `history` sheet. Missing a habit reduces its buffer the *next* day. If the buffer is 0 and you miss it, the current streak resets to 0 the next day. Buffers increase automatically based on your `boost interval` setting.
* **Help:** Click **Moderate Habits Settings > Show Help (Current Page)** to open a sidebar with relevant information about the current sheet and mode.
* **Detailed Guide:** For a more in-depth explanation and examples, refer to the guide (also linked in your approval email):
  * **➡️ [Link to Guide Here](http://tinyurl.com/moderate-habits) ⬅️**

## 🔧 Troubleshooting

* **Menu Not Appearing:** The "Moderate Habits Settings" menu might take a few seconds (up to 30) to appear when you first open the sheet, especially the very first time. Try reloading the page if it takes longer.
* **Haven't Received Approval Email:** Please allow some time for manual processing of your access request. Remember to check your spam/junk folder for the email titled 'Access Granted to "moderate habits"'.
* **Authorization Loop / Errors:** If authorization keeps failing or you encounter permission errors (especially after copying the sheet), ensure you requested access with the *same Google account* you are using to authorize the script in your copy. If issues persist, try removing the script's permissions from your Google Account settings (`Manage your Google Account -> Security -> Third-party apps with account access` -> Find the script by name -> Remove Access) and then re-authorize by running a menu item (like `Show Help`) again.
* **Script Errors ("An error occurred..."):** If you see error messages, check the script's execution logs for details: **Extensions > Apps Script > Executions**. This often provides specific error messages useful for debugging.
* **Performance:** Depending on the length of your challenge history, some actions (like changing dates far back or the daily reset) might take a few seconds to process. This is normal for Apps Script.

## 🙏 Contributing

Contributions, issues, and feature requests are welcome! Please feel free to open an issue on the [GitHub repository issues page](https://github.com/marcusrprojects/moderate_habits/issues) to report bugs or suggest improvements.

---

Happy habiting! 💪
