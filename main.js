/**
 * Creates or update an event in the user's default calendar from data inserted in a spreadsheet.
 * @see https://developers.google.com/calendar/api/v3/reference/events/insert
 * @see https://developers.google.com/calendar/api/v3/reference/events/update
 */
function createorUpdateEvents() {
  const calendarId = "YOUR CALENDAR ID";
  const sheet = SpreadsheetApp.getActiveSheet();
  const events = sheet.getRange("A2:G1000").getValues();
  let event;

  for (i = 0; i < events.length; i++) {
    const shift = events[i];
    const eventID = shift[0];
    const eventsubject = shift[1];
    const startTime = shift[2];
    const endTime = shift[3];
    const description = shift[4];
    const color = shift[5];

    if (
      eventID !== undefined &&
      eventsubject !== undefined &&
      description !== undefined &&
      color !== undefined &&
      startTime instanceof Date &&
      endTime instanceof Date
    ) {
      const event = {
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

      try {
        let createOrUpdate;
        if (event.id) {
          createOrUpdate = Calendar.Events.update(event, calendarId, eventID);
        } else {
          createOrUpdate = Calendar.Events.insert(event, calendarId);
        }
      } catch (e) {
        if (e.message && e.message.indexOf("Not Found") !== -1) {
          createOrUpdate = Calendar.Events.insert(event, calendarId);
        } else {
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
function exportCalendarEventsToSheet() {
  const calendarId = "YOUR CALENDAR ID";
  const startDate = new Date("2023-01-01");
  const endDate = new Date("2023-12-31");
  const calendar = CalendarApp.getCalendarById(calendarId);
  const sheet = SpreadsheetApp.getActiveSheet();
  const events = calendar.getEvents(startDate, endDate);
  const data = [];

  if (events.length > 0) {
    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      const eventID = event.getId().split("@")[0];
      const eventTitle = event.getTitle();
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      const description = event.getDescription();
      const color = event.getColor();

      data.push([eventID, eventTitle, startTime, endTime, description, color]);
    }

    const numRows = data.length;
    const numCols = data[0].length;
    sheet.getRange(2, 1, numRows, numCols).setValues(data);
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
    .addItem("Calendar to Sheet", "exportCalendarEventsToSheet")
    .addItem("Sheet to Calendar", "createorUpdateEvents")
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
