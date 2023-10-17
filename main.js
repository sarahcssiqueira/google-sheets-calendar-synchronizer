/**
 * Creates or update an event in the user's default calendar from data inserted in a spreadsheet.
 * @see https://developers.google.com/calendar/api/v3/reference/events/insert
 * @see https://developers.google.com/calendar/api/v3/reference/events/update
 */
function createUpdateEvent() {
  /*
   * Open the Calendar
   */
  const calendarId = "yourcalendarID";
  const sheet = SpreadsheetApp.getActiveSheet();

  /*
   * Import events data from the spreadsheet
   */
  const events = sheet.getRange("A2:G1000").getValues();

  /*
   * Event details for creating event
   */
  var event; // Declare event variable outside the loop

  for (i = 0; i < events.length; i++) {
    var shift = events[i];
    var eventID = shift[0];
    var eventsubject = shift[1];
    var startTime = shift[2];
    var endTime = shift[3];
    var description = shift[4];
    var color = shift[5];
    var attendees = shift[7];

    // Check if the variables are defined
    if (
      eventID !== undefined &&
      eventsubject !== undefined &&
      description !== undefined &&
      // && attendees !== undefined
      color !== undefined &&
      startTime instanceof Date &&
      endTime instanceof Date
    ) {
      let event = {
        id: eventID,
        summary: eventsubject,
        description: description,
        start: {
          dateTime: startTime.toISOString(),
          timeZone: "America/Sao_Paulo",
        },
        end: {
          dateTime: endTime.toISOString(),
          timeZone: "America/Sao_Paulo",
        },
        colorId: color,
      };

      /**
       * Insert or update event
       **/
      try {
        if (event.id) {
          // Update an existing event
          update = Calendar.Events.update(event, calendarId, eventID);
        } else {
          // Insert a new event
          create = Calendar.Events.insert(event, calendarId);
        }
      } catch (e) {
        if (e.message && e.message.indexOf("Not Found") !== -1) {
          // Handle the "Not Found" error by creating a new event
          create = Calendar.Events.insert(event, calendarId);
        } else {
          // Handle other errors
          console.error("Error:", e);
        }
      }
    }
  }
}

/*
 * Creates or update an event in the user's spreadsheet from data inserted in a Google calendar
 * @see https://developers.google.com/calendar/api/v3/reference/events/insert
 * @see https://developers.google.com/calendar/api/v3/reference/events/update
 */
function calendarToSheet() {
  const calendar = CalendarApp.getCalendarById("yourcalendarID");
  const sheet = SpreadsheetApp.getActiveSheet();
  const events = calendar.getEvents(
    new Date("01/01/2023"),
    new Date("12/31/2023")
  );
  const data = [];

  if (events) {
    for (i = 0; i < events.length; i++) {
      var eventID = events[i].getId();
      var eventPart = eventID.split("@")[0];

      data.push([
        eventPart,
        events[i].getTitle(),
        events[i].getStartTime(),
        events[i].getEndTime(),
        events[i].getDescription(),
        events[i].getColor(),
      ]);
    }

    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  } else {
    console.log("No events exist for the specified range");
  }
}

/*
 * Menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Sync data with Calendar")
    .addItem("Calendar to Sheet", "calendarToSheet")
    .addItem("Sheet to Calendar", "createUpdateEvent")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("About").addItem("Documentation", "documentation")
    )
    .addToUi();
}

/*
 * Display link for info
 */
function documentation() {
  SpreadsheetApp.getUi().alert(
    "For more info go to https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer"
  );
}
