/**
 * Partnership Course Automation Engine
 * Manages calendar invites, track separation, and intern tracking.
 */

function sendInitialRegistrationCalendarInvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  let detailsRange = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8);
  let detailsData = detailsRange.getValues();

  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const futureTime = new Date(now.getTime() + (30 * 24 * 60 * 60 * 1000)); 

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let internPreference = String(row[4]).toUpperCase().trim(); // Column E: Intern WE/WD
    let calendarInviteStatus = String(row[5]); // Column F: Calendar Invite status
    let rowIndex = index + 2;

    if (email && (calendarInviteStatus === "Send" || calendarInviteStatus === "")) {
      let isWeekend = (internPreference === "WE");

      // STRICT SEARCH: Search ONLY for the specific track title to avoid grabbing unrelated meetings
      let targetSearchTerm = isWeekend ? "Partnership Course Weekend" : "Partnership Course Weekday";
      
      let matchingEvents = calendar.getEvents(now, futureTime, { search: targetSearchTerm });
      
      // Filter further to ensure exact title match and prevent pulling in unrelated calendar entries
      let validEvents = matchingEvents.filter(event => event.getTitle().trim() === targetSearchTerm);

      if (validEvents.length > 0) {
        validEvents.forEach(event => {
          event.addGuest(email);
        });
      } else {
        // No sessions exist for this track yet this cohort - create all 4 with a
        // guaranteed Meet link instead of silently skipping this person.
        // (addStudentToAllFourCourseSessions lives in PreTaskTracker.gs)
        addStudentToAllFourCourseSessions(email, isWeekend);
      }

      // Mark Sent either way - previously this only happened inside the "found" branch,
      // so anyone whose sessions didn't exist yet was silently retried forever.
      detailsSheet.getRange(rowIndex, 6).setValue("Sent"); 
    }
  });
}