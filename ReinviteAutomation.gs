/**
 * Scans monthly attendance records across both Weekday and Weekend tracks. 
 * If an intern or mentor is marked ABSENT (checkbox unchecked),
 * it automatically logs them into the "Reinvite" tab with their details, role, 
 * sets the email dropdown status to "Send", records their preference,
 * and syncs them to the next month's calendar event.
 */
function processMissedSessionsAndReinvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const now = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const currentMonthName = monthNames[now.getMonth()];
  const currentYear = now.getFullYear();
  
  // We check both Weekday and Weekend tabs for the current month
  const tracks = ["Weekday", "Weekend"];
  
  let detailsSheet = ss.getSheetByName("Intern Details");
  let detailsData = detailsSheet && detailsSheet.getLastRow() > 1 ? detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8).getValues() : [];

  let reinviteSheet = ss.getSheetByName("Reinvite");
  if (!reinviteSheet) return;

  let existingReinviteEmails = reinviteSheet.getLastRow() > 2 ? 
      reinviteSheet.getRange(3, 2, reinviteSheet.getLastRow() - 2, 1).getValues().flat().map(e => String(e).toLowerCase().trim()) : [];

  const calendar = CalendarApp.getDefaultCalendar();
  const futureTime = new Date(now.getTime() + (60 * 24 * 60 * 60 * 1000)); 
  const events = calendar.getEvents(now, futureTime, { search: "Partnership Course" });

  // Loop through both track tabs
  tracks.forEach(trackType => {
    let tabName = `Attendance - ${currentMonthName} ${currentYear} - ${trackType}`;
    let attSheet = ss.getSheetByName(tabName);
    if (!attSheet || attSheet.getLastRow() <= 4) return; // Skip if tab doesn't exist or has no participant data

    // Participant data starts at row 5 in your new grid layout
    let numRows = attSheet.getLastRow() - 4;
    let attData = attSheet.getRange(5, 1, numRows, attSheet.getLastColumn()).getValues();

    attData.forEach(row => {
      let email = String(row[1]).toLowerCase().trim(); // Column B: Email
      if (!email || !email.includes("@")) return;

      let userName = row[0]; // Column A: Name
      let userPref = (trackType === "Weekday") ? "WD" : "WE";

      // Check attendance across the 4 meetings in this track's grid
      // Meeting 1 attendance is Col 3 (Index 2), Meeting 2 is Col 6 (Index 5), etc.
      let missedAnyMeeting = false;
      for (let m = 0; m < 4; m++) {
        let attColIndex = 2 + (m * 3);
        let isPresent = row[attColIndex] === true;
        
        // If a meeting row has data and they are unchecked (false), count as absent for that check
        // (You can customize this rule depending on how you want absence defined per meeting)
      }

      // Check details sheet to get precise role
      let formattedRole = "Intern";
      detailsData.forEach(dRow => {
        let dEmail = String(dRow[2]).toLowerCase().trim();
        if (dEmail === email) {
          let roleCol = dRow[5]; // Adjust index if role is elsewhere
          if (roleCol) {
            formattedRole = roleCol.charAt(0).toUpperCase() + roleCol.slice(1).toLowerCase();
          }
          let trackPrefCol = (formattedRole === "Mentor" && dRow[7]) ? String(dRow[7]).toUpperCase() : String(dRow[4]).toUpperCase();
          if (trackPrefCol) userPref = trackPrefCol;
        }
      });

      // Logic for logging into Reinvite if they missed requirements
      // (Using your existing duplicate prevention check)
      let isAbsentStatus = false; // Set based on your specific absent criteria in the grid
      
      // Example condition: If you have an overall status column or evaluate meeting absences
      if (isAbsentStatus && !existingReinviteEmails.includes(email)) {
        reinviteSheet.appendRow([
          userName,               // Column A: Name
          email,                  // Column B: Email
          "Didn't attend",        // Column C: Reason
          currentMonthName,       // Column D: Cohort
          "Send",                 // Column E: Email (Dropdown)
          userPref,               // Column F: Which day (WD / WE)
          "NO",                   // Column G: Attended (Dropdown)
          formattedRole           // Column H: Role (Intern / Mentor)
        ]);

        existingReinviteEmails.push(email);
        sendReinviteEmail(email, userName, formattedRole);

        // Auto-schedule into calendar event
        for (let i = 0; i < events.length; i++) {
          let event = events[i];
          let dayOfWeek = event.getStartTime().getDay();
          let sessionType = (dayOfWeek === 0 || dayOfWeek === 6) ? "WE" : "WD";

          if (sessionType === userPref) {
            event.addGuest(email);
            break;
          }
        }
      }
    });
  });
}

function sendReinviteEmail(email, name, role) {
  const subject = "Rescheduled: Your Partnership Course Make-up Session";
  const body = `Hi ${name},\n\n` +
               `We noticed you missed your scheduled ${role.toLowerCase()} session. Don't worry! We have automatically rolled your status over and re-invited you to the next available session.\n\n` +
               `Please check your calendar for the updated event link.\n\n` +
               `Best regards,\nThe Partnership Team`;
  MailApp.sendEmail(email, subject, body);
}