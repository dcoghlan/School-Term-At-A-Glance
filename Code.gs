// School Term Calendar Generator for Google Sheets (v20)
// 
// SETUP INSTRUCTIONS:
// 1. Create a new Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire script
// 4. Save the project (give it a name like "Term Calendar")
// 5. Close the script editor and refresh your sheet
// 6. You'll see a new "Term Calendar" menu appear
// 7. Use "Term Calendar > Setup Configuration" to enter your settings
// 8. Use "Term Calendar > Generate Calendar" to create the complete calendar.

// Configuration sheet name
const CONFIG_SHEET = 'Config';
const CALENDAR_SHEET = 'Term Calendar';

// Other Variables
const RANGE_GENERATED = 'D2';
const TXT_GENERATED = 'Last Generated: ';
const DEFAULT_STALE_PERIOD_HRS = 6;

function listAllAccessibleCalendars() {
  // Use the CalendarApp service to retrieve all calendars the user can see.
  const calendars = CalendarApp.getAllCalendars();
  
  // Array to hold the list of calendar details
  const calendarList = [];
  
  Logger.log("--- All Accessible Calendars ---");
  
  // Loop through the array of Calendar objects
  calendars.forEach(calendar => {
    const calendarName = calendar.getName();
    const calendarId = calendar.getId();
    const isOwned = calendar.isOwnedByMe();
    
    // Log details for easy viewing
    Logger.log(`Name: ${calendarName} | ID: ${calendarId} | Owned: ${isOwned}`);
    
    // You can also store this data to display in a spreadsheet or HTML UI later
    calendarList.push({
      name: calendarName,
      id: calendarId,
      owned: isOwned
    });
  });
  
  // Return the full list (optional)
  return calendarList;
}

// Run when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Term Calendar')
    .addItem('Setup Configuration', 'showConfigDialog')
    .addItem('Generate Calendar (Recreates Sheet)', 'generateCalendar')
    .addToUi();

}

// Show configuration dialog
function showConfigDialog() {

  const config = getConfig();

  // 1. Load the HTML file as a template
  const template = HtmlService.createTemplateFromFile('Index');

  // 2. Assign variables to the template object
  template.termName = config.termName;

  // 3. Evaluate the template to create the final HtmlOutput object
  const html = template.evaluate()
  .setWidth(500)
  .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Term Calendar');
}

// Save configuration
function saveConfig(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG_SHEET);
  
  if (!configSheet) {
    // Insert Config sheet at the second position (index 1) for tidiness
    configSheet = ss.insertSheet(CONFIG_SHEET, 1);
  }
  
  configSheet.clear();
  configSheet.getRange('A1:B7').setValues([
    ['Setting', 'Value'],
    ['Term Name', config.termName],
    ['Start Date', config.startDate],
    ['Week Count', config.weekCount],
    ['Calendar ID', config.calendarId],
    ['Historical Date Shading', config.historicalShading],
    ['Stale Period (Hrs)', config.stalePeriod]

  ]);
  
  configSheet.getRange('A1:B1').setFontWeight('bold');
  configSheet.setColumnWidth(1, 150);
  configSheet.setColumnWidth(2, 300);
}

// Get configuration
function getConfig() {
  Logger.log("getConfig() function started executing.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  
  if (!configSheet) {
    Logger.log("getConfig() no config sheet found.");
    return {};
  }
  
  const data = configSheet.getRange('A2:B7').getValues();

  // Check if it's a Date object before formatting
  let startDateString = data[1][1];
  if (startDateString instanceof Date) {
    startDateString = Utilities.formatDate(startDateString, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    Logger.log("getConfig() Formatted start date to string: " + startDateString);
  }
  
  Logger.log("getConfig(): " +data)

  return {
    termName: data[0][1],
    startDate: startDateString,
    weekCount: data[2][1],
    calendarId: data[3][1],
    historicalShading: data[4][1],
    stalePeriod: data[5][1]
  }
}

// Test date object time for 00:00:00
function isMidnight(dateObject) {
  // Use getHours(), getMinutes(), getSeconds(), and getMilliseconds()
  // which all return values based on the *local* time zone.
  return dateObject.getHours() === 0 &&
         dateObject.getMinutes() === 0 &&
         dateObject.getSeconds() === 0 &&
         dateObject.getMilliseconds() === 0;
}

// Change font color of cell displaying the last time the sheet
// was updated/generated if its more than X number of hours ago.
function isStale() {
  const config = getConfig();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calSheet = ss.getSheetByName(CALENDAR_SHEET);

  if (!calSheet) {
    return {};
  }

  const dataRange = calSheet.getRange(RANGE_GENERATED);
  const dataValue = dataRange.getValue();
  const rawDate = dataValue.substring(TXT_GENERATED.length);

  const storedDate = new Date(rawDate);
  const now = new Date();
  const stalePeriodHrs = config.stalePeriod ? config.stalePeriod : DEFAULT_STALE_PERIOD_HRS;

  const hoursInMs = stalePeriodHrs * 60 * 60 * 1000

  const differenceMs = now.getTime() - storedDate.getTime()
  if (differenceMs > hoursInMs && hoursInMs > 0) {
    dataRange.setFontColor('#ff0000');
  }

}

// Generate the calendar
function generateCalendar() {
  const config = getConfig();
  Logger.log('Historical shading = ' + config.historicalShading)
  
  if (!config.termName || !config.startDate || !config.weekCount) {
    SpreadsheetApp.getUi().alert('Please configure the calendar first using "Term Calendar > Setup Configuration"');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startDate = new Date(config.startDate);
  const weekCount = parseInt(config.weekCount);
  const TIMEZONE = Session.getScriptTimeZone(); // Define timezone once

  // --- 1. Event Fetching and Dynamic Row Calculation ---
  const calendarId = config.calendarId || 'primary';
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + (weekCount * 7));
  
  let events = [];
  const eventsByDate = {};
  const eventsByWeek = [];
  const MIN_EVENT_ROWS = 5; 

  try {
    const calendar = CalendarApp.getCalendarById(calendarId) || CalendarApp.getDefaultCalendar();
    events = calendar.getEvents(startDate, endDate);
    
    // Initialize eventsByWeek with the minimum row count
    for (let i = 0; i < weekCount; i++) {
      eventsByWeek[i] = MIN_EVENT_ROWS; 
    }
    
    events.forEach(event => {
      const eventStart = event.getStartTime();
      const eventEnd = event.getEndTime();
      let isAllDay = event.isAllDayEvent();
      const eventTitle = event.getTitle();
      
      // Check to see if both the event start and end dates contains 00:00:00 AND the event
      // was NOT created using the "All Day" checkbox in the UI. In this case, set the isAllDay
      // flag manually so the start time does not get displayed in the term planner.
      if (isMidnight(eventStart) && isMidnight(eventEnd) && !isAllDay) {
        Logger.log('Manually set isAllDay to true for: ' + eventTitle);
        isAllDay = true;
      }
      // Determine the date range to check
      const dayStart = new Date(Utilities.formatDate(eventStart, TIMEZONE, 'yyyy/MM/dd'));
      
      // For all-day events, the end time is midnight of the *next* day, so we subtract one day.
      // For timed events, the end time can be later on the same day, so we check the date.
      const dayEnd = new Date(Utilities.formatDate(eventEnd, TIMEZONE, 'yyyy/MM/dd'));
      if (isAllDay) {
          dayEnd.setDate(dayEnd.getDate() - 1);
      }
      
      // Get the correct event display text (only show time for the start day)
      const eventTimeText = isAllDay 
          ? eventTitle 
          : `${eventTitle} [${Utilities.formatDate(eventStart, TIMEZONE, 'h:mma').toLowerCase()}]`;
          
      const eventTitleMultiDay = isAllDay ? eventTitle : eventTitle;


      // Loop through every day of the event
      let currentDateIterator = new Date(dayStart);
      while (currentDateIterator <= dayEnd) {
          const eventDateStr = Utilities.formatDate(currentDateIterator, TIMEZONE, 'yyyy-MM-dd');
          
          if (!eventsByDate[eventDateStr]) {
            eventsByDate[eventDateStr] = [];
          }

          // Use a different text format for subsequent days of a multi-day event
          let textToAdd = eventTimeText;
          if (Utilities.formatDate(currentDateIterator, TIMEZONE, 'yyyy/MM/dd') > Utilities.formatDate(dayStart, TIMEZONE, 'yyyy/MM/dd')) {
              // On days after the start day, just show the title to save space and remove repeated time
              textToAdd = `${eventTitleMultiDay} (cont.)`;
          }

          eventsByDate[eventDateStr].push(textToAdd);
          
          // Dynamic Row Calculation
          const eventDateObj = new Date(eventDateStr);
          const daysDiff = Math.floor((eventDateObj - startDate) / (1000 * 60 * 60 * 24));
          const weekIndex = Math.floor(daysDiff / 7);
          
          if (weekIndex >= 0 && weekIndex < weekCount) {
            eventsByWeek[weekIndex] = Math.max(eventsByWeek[weekIndex], eventsByDate[eventDateStr].length);
          }
          
          // Move to the next day
          currentDateIterator.setDate(currentDateIterator.getDate() + 1);
      }
    });
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error loading calendar events: ' + e.message, 'Error', 5);
    return;
  }
  // --- END Event Fetching ---


  // --- 2. Sheet Setup ---
  let calSheet = ss.getSheetByName(CALENDAR_SHEET);
  
  // Delete existing sheet for a completely clean slate
  if (calSheet) {
    if (ss.getSheets().length > 1) {
        ss.deleteSheet(calSheet);
    } else {
        // If it's the ONLY sheet, just clear ALL formatting and content
        calSheet.clear();
        calSheet.clearFormats();
        SpreadsheetApp.getUi().alert('The calendar sheet was the only sheet, so it was cleared instead of deleted.');
    }
  }
  
  // Insert the new sheet at index 0 (first position)
  calSheet = ss.insertSheet(CALENDAR_SHEET, 0);
  calSheet.setHiddenGridlines(true);
  calSheet.activate();
  

  // --- 3. Drawing the Calendar Structure ---
  
  // Set up title
  calSheet.getRange(1,3,1,3)
    .merge()
    .setValue(config.termName)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .activate();
  
  const now = new Date();
  // Format: 'dd MMM yyyy HH:mm:ss' (e.g., '14 Oct 2025 22:19:38')
  const refreshTime = Utilities.formatDate(now, TIMEZONE, 'dd MMM yyyy HH:mm:ss');
  
  calSheet.getRange(RANGE_GENERATED).setValue(TXT_GENERATED + refreshTime);
  calSheet.getRange(RANGE_GENERATED).setFontSize(9).setFontColor('#666666').setHorizontalAlignment('center');
  
  const dateMap = {};
  let currentRow = 3; // Starts at week header row
  
  // Generate each week
  for (let week = 0; week < weekCount; week++) {
    const eventRowsThisWeek = eventsByWeek[week]; // Use the dynamically calculated required rows
    
    // Week header
    calSheet.getRange(currentRow, 1, 1, 7)
      .merge()
      .setValue(`Week ${week + 1}`)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    currentRow++; // Moves to day header row
    
    // Day headers
    const dayHeaderRow = currentRow;
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    for (let day = 0; day < 7; day++) {
      const dayOffset = week * 7 + day;
      const currentDate = new Date(startDate);
      currentDate.setDate(startDate.getDate() + dayOffset);
      
      const cell = calSheet.getRange(dayHeaderRow, day + 1);
      cell.setValue(`${days[day].substring(0, 3)} ${Utilities.formatDate(currentDate, TIMEZONE, 'MMM d')}`);
      cell.setFontWeight('bold');
      cell.setVerticalAlignment('top');
      cell.setHorizontalAlignment('center');
      cell.setWrap(true);

      if (config.historicalShading) {
        if (Utilities.formatDate(currentDate, TIMEZONE, 'yyyy-MM-dd') < Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd')) {
          cell.setFontColor('#b7b7b7')
        }
      }

      // Activate the cell for the current week (if applicable))
      if (Utilities.formatDate(currentDate, TIMEZONE, 'yyyy-MM-dd') == Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd')) {
        const highlight = calSheet.getRange(dayHeaderRow - 1, day + 1);
        highlight.activate();
      }
      
      // Weekend styling
      if (day >= 5) {
        cell.setBackground('#f3f3f3');
      } else {
        cell.setBackground('#ffffff');
      }
      
      // Store date in hidden note for event population
      const dateStr = Utilities.formatDate(currentDate, TIMEZONE, 'yyyy-MM-dd');
      cell.setNote(dateStr);
      
      // Map for event population later
      dateMap[dateStr] = {
          col: day + 1,
          eventStartRow: dayHeaderRow + 1,
          maxRows: eventRowsThisWeek // Critical: use dynamic max rows
      };
    }
    
    currentRow++; // Moves to the first event row
    
    // Event rows (DYNAMIC rows per day for events)
    for (let eventRow = 0; eventRow < eventRowsThisWeek; eventRow++) {
      for (let day = 0; day < 7; day++) {
        const cell = calSheet.getRange(currentRow + eventRow, day + 1);
        cell.setVerticalAlignment('top');
        cell.setWrap(true);
        
        // Initial event cell styling (can be overwritten by event updates)
        if (day >= 5) {
          cell.setBackground('#f9f9f9');
        } else {
          cell.setBackground('#ffffff');
        }
      }
    }
    
    currentRow += eventRowsThisWeek; // Add the dynamic number of event rows
    currentRow++; // Gap between weeks
  }
  
  // --- 4. Event Population and Final Formatting ---

  // Batch update arrays for final cell population
  const updates = [];
    
  for (const dateStr in eventsByDate) {
    if (!dateMap[dateStr]) continue;
    
    const position = dateMap[dateStr];
    const dayEvents = eventsByDate[dateStr];
    
    dayEvents.forEach((eventText, index) => {
      if (index >= position.maxRows) return;
      
      const row = position.eventStartRow + index;
      const col = position.col;
    
      // Set event font color to black by default and override if the date is prior to todays date
      let eventFontColor = '#000000' 
      if (config.historicalShading) {
        const dateIter = Utilities.formatDate(new Date(dateStr), TIMEZONE, 'yyyy-MM-dd');
        const dateNow = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd')
        if (Utilities.formatDate(new Date(dateStr), TIMEZONE, 'yyyy-MM-dd') < Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd')) {
          eventFontColor = ('#b7b7b7');
        }
      }

      updates.push({
        range: calSheet.getRange(row, col),
        value: eventText,
        fontColor: eventFontColor,
        background: col >= 6 ? '#f9f9f9' : '#e3f2fd'
      });
    });
  }
  
  // Apply all event updates
  updates.forEach(update => {
    update.range.setValue(update.value);
    update.range.setFontSize(9);
    update.range.setFontColor(update.fontColor);
    update.range.setBackground(update.background);
    update.range.setBorder(true, true, true, true, false, false, '#90caf9', SpreadsheetApp.BorderStyle.SOLID);
  });


  // Final formatting
  for (let col = 1; col <= 7; col++) {
    calSheet.setColumnWidth(col, 160);
  }
  
  // Add borders to the whole calendar block
  const lastRow = currentRow - 1;
  calSheet.getRange(3, 1, lastRow - 2, 7).setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Calendar generated successfully! Loaded ${events.length} events from Google Calendar.`, 'Status', 7);
}