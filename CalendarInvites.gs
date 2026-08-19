function sendInitialRegistrationCalendarInvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  let detailsRange = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8);
  let detailsData = detailsRange.getValues();

  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const futureTime = new Date(now.getTime() + (30 * 24 * 60 * 60 * 1000)); 
  
  const events = calendar.getEvents(now, futureTime, { search: " Course" });
  if (events.length === 0) return;

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let internPreference = String(row[4]).toUpperCase(); // Column E: Intern WE/WD
    let calendarInviteStatus = String(row[5]); // Column F: Calendar Invite status
    let rowIndex = index + 2;

    if (email && (calendarInviteStatus === "Send" || calendarInviteStatus === "")) {
      for (let i = 0; i < events.length; i++) {
        let event = events[i];
        let eventDateObj = event.getStartTime();
        let dayOfWeek = eventDateObj.getDay(); 
        let sessionType = (dayOfWeek === 0 || dayOfWeek === 6) ? "WE" : "WD";

        if (sessionType === internPreference) {
          event.addGuest(email);
          detailsSheet.getRange(rowIndex, 6).setValue("Sent"); 
          break;
        }
      }
    }
  });
}