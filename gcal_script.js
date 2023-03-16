//**  https://developers.google.com/sheets/api/quickstart/apps-script */
//** Testing Google Scripts */

function myFunction() {
  const spreadsheetId = 'my spreadsheetId';
  const dateRange = 'Main!B210:G211';
  const calendarId = 'my calendarId'
  const calendar = CalendarApp.getCalendarById(calendarId);

  try {
    // Get the values from the spreadsheet using spreadsheetId and range.
    const values = Sheets.Spreadsheets.Values.get(spreadsheetId, dateRange).values;
    //  Print the values from spreadsheet if values are available.
    if (!values) {
      console.log('No data found.');
      return;
    }
    for (const row in values) {
      // Test to verify correct values are being used.

      // Extract Date
      var date = new Date(values[row][0]);
      var eventYear = date.getFullYear()
      var eventMonth = date.getMonth() + 1
      var eventDay = date.getDate()


      console.log(eventYear, eventMonth, eventDay);

      var eventTitle = values[row][1];

      console.log(eventTitle)

      calendar.createAllDayEvent(eventTitle, date);

      //Avoid duplicate events beings added
      //

      //console.log('Date: %s', values[row][0]);
      //console.log('Event: %s', values[row][1]);
      //console.log('Dept: %s', values[row][2]);
      //console.log('Cat: %s', values[row][3]);
      //console.log('Staff: %s', values[row][4]);
      //console.log('Priv: %s', values[row][5]); 


    }
  } 
  catch (err) {
    // TODO (developer) - Handle Values.get() exception from Sheet API
    console.log(err.message);
  }
}