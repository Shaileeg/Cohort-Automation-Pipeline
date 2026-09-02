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

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let internPreference = String(row[4]).toUpperCase().trim(); // Column E: Intern WE/WD
    let calendarInviteStatus = String(row[5]); // Column F: Calendar Invite status
    let rowIndex = index + 2;

    if (email && (calendarInviteStatus === "Send" || calendarInviteStatus === "")) {
      let isWeekend = (internPreference === "WE");

      // Finds the existing 4-week recurring series for this track (or creates it
      // if this is the first person this cohort) and adds them as a guest to the
      // whole series, sharing the same Meet link as everyone else.
      // (addStudentToAllFourCourseSessions lives in PreTaskTracker.gs)
      addStudentToAllFourCourseSessions(email, isWeekend);

      detailsSheet.getRange(rowIndex, 6).setValue("Sent"); 
    }
  });
}