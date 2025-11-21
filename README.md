# 🗓️ School-Term-At-A-Glance

**Latest Release:** v0.0.31

## Installation Guide

Welcome! This guide will help you set up the School Term At A Glance tool. You do not need to be a programmer to set this up—just follow these steps one by one.

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

## ⚙️ Configuration Guide

There are several items that are required to be configured so as to generate a Term Calendar. 

### Step 1: Open the configuration window:

- In your Google Sheet, look at the top menu bar.
- Click `Term Calendar` > `Setup Configuration`.
- A popup window called "**Configure Term Calendar**" will be displayed.

### Step 2: Enter the minimum required details

| Config Item | Description |
| :--- | :--- |
| **Term Name** | This is the heading text that gets displayed at the top of the generated calendar. |
| **Start Date** | Select the start date for week 1 of the generated calendar. *The start date should always be the Monday of the week. If you choose any other day, the weeks will be out of whack. This will be address in a future update* |
| **Number of Weeks** | Select the number of weeks for the school term. Current options available are 9,10 & 11. |
| **Google Calendar ID** | This is the identifier of the Google Calendar to pull the events from. To find your Calendar ID: Open Google Calendar → Settings → Select your calendar → Scroll to "Integrate calendar" → Copy the Calendar ID |

- Click *Save Configuration*

### Step 3: Generate the calendar

- In your Google Sheet, look at the top menu bar.
- Click `Term Calendar` > `Generate Calendar (Recreates Sheet)`.
- A new sheet will be created which will be populated with events from the calendar specified in the configuration.

## Updating/Regenerating the calendar

Due to the method used to create the calendar, the script does not update individual entries in the existing sheet, it actually deletes the whole sheet and creates a calendar on a new sheet with all the updated information.

Based on this there are 2 options when it comes to re-generating the calendar: Manual and Scheduled

### Regenerate Calendar (Manual)

- In your Google Sheet, look at the top menu bar.
- Click `Term Calendar` > `Generate Calendar (Recreates Sheet)`.
- The existing sheet containing the calendar will be renamed (because we cannot zero sheets), the new one created and then the old renamed version will be deleted.

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
A: This is planned to be streamlined in a future release so that the configuration UI will present a list of available Google Calendars to choose from.

Q: Can I change the colors, fonts, borders etc??  
A: Yes it is possible to modify the formatting by editing the script manually. This ability will be added to the configuration UI in a future release

Q: How do I know which version of the script my calendar is using?  
A: The version number is listed at the bottom of the calendar.

Q: I need to select a number of weeks which is not available in the configuration UI?  
A: The options available in the configuraion UI were selected based on school terms in Sydney NSW. As a work around, just select one of the available options and generate the initial calendar. Once it has been generated, go into the Config sheet and update the `Week Count` value to whatever you need. The configuration UI will be updated in a future release and this will be fixed during that update.

Q: What will happen if I make changes to the generated calendar?  
A: Due to the fact we completely delete the sheet and create a new sheet and calendar, the changes will be lost. The script only pulls events from the Google Calendar as the source of truth. If you want to keep a copy of the modified calendar, you can rename the sheet (tab) to something else, and when the calendar is regenreted, the old renamed one will not be deleted.

Q: Are there plans to have updates without having to regenerate the whole calendar?  
A: No  

Q: Are there plans to add the ability to add/update events on the Google Calendar from this script?
A: No, this was designed as an alternate read-only view of an existing Google Calendar calendar. There are no plans to change this.

Q: I have multiple calendars that I would like to use this for. Can I do this in one file?  
A: No, a single Google Sheet is used for a single Google Calendar. Once you've got the script working in a Google Sheet, you can make a copy of the Google Sheet and the associated scripts will get copied to the new file and you can configure the new file to point to a different calendar. You will have to go through the Authorization process again though.
