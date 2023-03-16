// Pulls data from target Google Sheet to add to target Google Calendar
function pullData() {
  // Get spreadsheet data
  const spreadsheetId = 'spreadhsheet';
  const dateRange = 'Main!B210:G211';
  const values = Sheets.Spreadsheets.Values.get(spreadsheetId, dateRange).values;
  // Get calendar data
  const calendarId = 'calendarId'
  const calendar = CalendarApp.getCalendarById(calendarId);

  try {  
    for (const row in values) {
      // Set variables baesd on Sheets data
      var date = new Date(values[row][0]);
      var [_, eventTitle, department, category, staff, privacy, eventType, notes, links] = values[row];

      // Check if event of same name already exists at date
      if (!calendar.getEventsForDay(date, {search: eventTitle}).length) {
        var eventDescription = `
          Department: ${department}
          Category: ${category}
          Staff Responsible: ${staff}
          Privacy: ${privacy}
          Type: ${eventType}
          Notes: ${notes}
          Link: ${links}
        `;

        // Add event to calendar
        var event = calendar.createAllDayEvent(eventTitle, date);
        event.setDescription(eventDescription);
      }
    }
  } 
  catch (err) {
    console.log(err.message);
  }
}

// Some blanks in description return undefined instead of blank.
// May use a seperate calendar for each department
