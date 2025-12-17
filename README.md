# 🗓️ School-Term-At-A-Glance

**Latest Release:** v0.1.0

Welcome! **School-Term-At-A-Glance** is a custom Google Apps Script that generates a read-only *"Term At A Glance"* calendar style view in Google Sheets. The events used to populate the calendar are pulled from an existing Google Calendar.

This guide will help you set up **School-Term-At-A-Glance**. You do not need to be a programmer to set this up. There are 2 methods to choose from to get this working, just pick which ever you are more comfortable with and follow the instructions.

- [Easy Installation Guide](README.md#easy-installation-guide)
- [Manual Installation Guide](README.md#manual-installation-guide)

## Easy Installation Guide

- Click the following link to *"Make a copy"* of a pre-configured Google Sheet with the code already installed
  - [Latest Release Template](https://docs.google.com/spreadsheets/d/1PiKQ9XW1XuqSLhavRRS1R31L5SV4ifnHWxVxyS---AQ/copy)
- A Copy document screen will be shown with the message "*The attached Apps Script file and functionality will also be copied. Would you like to make a scopy of School-Term-At-A-Glance vx.x.x?*"
- Click "**Make a copy**"
- A copy of the file will be made in your "**My Drive**"
- Wait about 10-20 seconds.
- Look at the top menu bar (next to "Help"). You should see a new menu appear called "**Term Calendar**".
- Execute [Step 6](README.md#step-6-the-first-run-authorization) from the Manual Installation Guide.
- Run the [Configuration Guide](README.md#configuration-guide) steps.

## Manual Installation Guide

### Prerequisite: Get the Code

Before you start, ensure you have the code files from this repository open:

- `src/Code.js` (The main Google Apps script)
- `src/Index.html` (The settings page visual template)

### Step 1: Create the Google Sheet

- Go to Google Sheets and create a Blank Spreadsheet.
- Name the spreadsheet something clear, like `Term Planner 2025 - Term 4`.

### Step 2: Open the Script Editor

This is where the magic happens. We need to paste the code that generate the term planner.

- In your Google Sheet, look at the top menu bar.
- Click `Extensions` > `Apps Script`.
- A new browser tab will open. This is the "Script Editor."

### Step 3: Add the Main Code

- In the Script Editor, you will see a file on the left named `Code.gs`.
- In the main editing window, you will see some default code (`function myFunction...`). **Delete everything so the window is completely empty.**
- Copy **all** the text from the `src/Code.js` file provided in this repository.
- Paste it into the editor window.
- Click the 💾 Save icon (floppy disk) in the toolbar.

### Step 4: Add the Template File

Now we need to add the visual part of the tool.

- In the Script Editor, look at the **Files** list on the left side.
- Click the **+ (Plus)** icon next to "Files".
- Select **HTML** from the list.
- **Naming is important:** Name the new file `Index` (It will automatically become `Index.html`).
- Delete the default code inside this new file.
- Copy **all** the text from the `src/Index.html` file provided in this repository.
- Paste it into the editor window.
- Click the 💾 Save icon again.

### Step 5: Refresh and Verify

- Close the Apps Script tab and go back to your Google Sheet.
- **Refresh** the web page (press F5 or the circular arrow button on your browser).
- Wait about 5–10 seconds.
- Look at the top menu bar (next to "Help"). You should see a new menu appear called **Term Calendar**.

### Step 6: The First Run (Authorization)

The first time you try to run the tool, Google will ask for permission. This is normal!

- Click **Term Calendar** > **Setup Configuration**.
- A window will pop up stating "Authorization required" Click **OK**.
- Select your Google Account associated with the calendar the events should be retrieved from.
- You may see a scary screen that says "**Untitled project wants access to your Google Account**" - Don't Panic: This appears because you created the script, and Google doesn't know you yet. It is safe.
- Click **Advanced** (usually small text in the corner).
- Click **Go to Untitled Project (unsafe)** at the bottom.
- You will be taken to a screen that says "**Untitled project wants access to your Google Account**"
- Check the "**Select all**" check box to grant the required permissions
- Click the **Continue** button

The script should now run and you will be presented with the **Configure Term Calendar**. You won't have to do this authorization step again unless the code changes significantly.

## Configuration Guide

There are several items that are required to be configured so as to generate Term Calendars.

### Step 1: Open the configuration window

- In your Google Sheet, look at the top menu bar.
- Click `Term Calendar` > `Setup Configuration`.
- A popup window called "**Configure Term Calendar**" will be displayed.

### Step 2: Configure Global Settings

| Config Item | Description |
| :--- | :--- |
| **Select Google Calendar** | Select the Google Calendar that the events will be read from (applies to all calendars). |

### Step 3: Configure Individual Calendars

You can configure up to 9 calendars. Each calendar will create a separate sheet in your spreadsheet.

| Config Item | Description |
| :--- | :--- |
| **Term Name** | The text entered will be used to name the sheet and also displayed as the title. |
| **Start Date** | Select the start date for week 1 of the generated calendar. |
| **Number of Weeks** | Select the number of weeks for the school term (1-52). |
| **Week Header Color** | Choose the background color for week headers. |

**Calendar Management:**

- **Add Calendar:** Click the "+ Add Calendar" button to add a new calendar (up to 9 maximum).
- **Delete Calendar:** Click the "Delete" button on any calendar to remove it (minimum 1 required).
- **Reorder Calendars:** Drag and drop calendars using the ☰ handle to change their order. The order determines the sheet order in your spreadsheet.

### Step 4: Save Configuration

- Click *Save Configuration* and wait while the calendars are generated.

## Updating/Regenerating the calendar

Due to the method used to create the calendar, the script does not update individual entries in the existing sheet, it actually deletes the whole sheet and creates a calendar on a new sheet with all the updated information.

Based on this there are 2 options when it comes to re-generating the calendar: Manual and Scheduled

### Regenerate Calendar (Manual)

- In your Google Sheet, look at the top menu bar.
- Click `Term Calendar` > `Generate Calendar (Recreates Sheet)`.
- The existing sheet containing the calendar will be renamed (because we cannot zero sheets), the new one created and then the old renamed version will be deleted.

### Update Configuration Settings

- In your Google Sheet, look at the top menu bar.
- Click `Term Calendar` > `Setup Configuration`.
- A popup window called "**Configure Term Calendar**" will be displayed.
- Modify the configuration settings as required
- Click *Save Configuration* and wait while the calendar is generated.

### Regenerate Calendar (Scheduled)

To regenerate the calendar on a schedule, a trigger needs to be configured:

- In your Google Sheet, look at the top menu bar.
- Click **Extensions** > **Apps Script**.
- A new browser tab will open. This is the "Script Editor."
- Click on the **Triggers** icon in the menu on the left. It looks like an Alarm Clock.
- Click the "**Add Trigger**" button (normally located in the bottom right corner)
- Under "**Choose which function to run**" select `generateCalendar`
- Under "**Select event source**" select `Time-driven`
- Select the appropriate time-based trigger and interval required. Suggest to start with a `Day timer` and select a time of day that is appropriate. The sheet does NOT need to be opened for it to be regenerated.
- Click **Save**

## Troubleshooting

If the calendar is empty or looks wrong, check the **Config** sheet

## FAQ

Q: Finding my Google Calendar ID is difficult, is there an easier way?  
A: The latest versions (>= 0.0.33) will auto-populate the list of calendars that are available to you to select from.

Q: Can I change the colors of the week headers?  
A: Yes! Each calendar has its own week header color setting. You can customize the color individually for each calendar.

Q: Can I change the fonts and borders etc??  
A: Font and borders can currently only be modified by editing the script manually if you dare to do so. This ability may be added to the configuration UI in a future release.

Q: How do I know which version of the script my calendar is using?  
A: The version number is listed at the bottom of the calendar.

Q: I need to select a number of weeks which is not available in the configuration UI?  
A: The latest version (>= 0.0.33) allows you to enter values between 1 and 52 for the number of weeks. This was limited to 52 weeks due to performance issues, however if for some reason you need a higher number, you can modify it directly in the HTML code as that is where the validation is done.

Q: What will happen if I make changes to the generated calendar?  
A: Due to the fact we completely delete the sheet and create a new sheet and calendar, the changes will be lost. The script only pulls events from the Google Calendar as the source of truth. If you want to keep a copy of the modified calendar, you can rename the sheet (tab) to something else, and when the calendar is regenerated, the old renamed one will not be deleted.

Q: Are there plans to have updates without having to regenerate the whole calendar?  
A: No  

Q: Are there plans to add the ability to add/update events on the Google Calendar from this script?  
A: No, this was designed as an alternate read-only view of an existing Google Calendar calendar. There are no plans to change this.

Q: I have multiple school terms that I would like to use this for. Can I do this in one file?  
A: Yes! As of v0.1.0, you can configure up to 9 calendars in a single spreadsheet. Each calendar creates its own sheet. All calendars pull events from the same Google Calendar but can have different term names, start dates, week counts, and header colors.

Q: Can I pull events from different Google Calendars for each term?  
A: No, all calendars in a spreadsheet pull from the same Google Calendar source. If you need to use different Google Calendars, you'll need to create separate spreadsheet files.
