/**
 * Scans monthly attendance records. If an intern or mentor is marked ABSENT,
 * it automatically logs them into the "Reinvite" tab with their details, role, 
 * sets the email dropdown status to "Send", records their preference,
 * and syncs them to the next month's calendar event.
 */
function processMissedSessionsAndReinvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const now = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const currentMonthName = monthNames[now.getMonth()];
  
  // 1. Current month attendance sheet
  const currentMonthTab = "Attendance - " + currentMonthName + " " + now.getFullYear();
  let attSheet = ss.getSheetByName(currentMonthTab);
  if (!attSheet || attSheet.getLastRow() <= 1) return;

  let attData = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, 6).getValues();
  
  let detailsSheet = ss.getSheetByName("Intern Details");
  let detailsData = detailsSheet.getLastRow() > 1 ? detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8).getValues() : [];

  let reinviteSheet = ss.getSheetByName("Reinvite");
  if (!reinviteSheet) return;

  let existingReinviteEmails = reinviteSheet.getLastRow() > 2 ? 
      reinviteSheet.getRange(3, 2, reinviteSheet.getLastRow() - 2, 1).getValues().flat().map(e => String(e).toLowerCase().trim()) : [];

  const calendar = CalendarApp.getDefaultCalendar();
  const futureTime = new Date(now.getTime() + (60 * 24 * 60 * 60 * 1000)); 
  const events = calendar.getEvents(now, futureTime, { search: "Partnership Course" });

  attData.forEach(row => {
    let email = String(row[1]).toLowerCase().trim();
    let role = row[4]; // "INTERN" or "MENTOR" from attendance log
    let status = row[5]; // "PRESENT" or "ABSENT"

    if (status === "ABSENT" && !existingReinviteEmails.includes(email)) {
      let userName = "Participant";
      let userPref = "WD";

      // Fetch name and preference from Intern Details
      detailsData.forEach(dRow => {
        let dEmail = String(dRow[2]).toLowerCase().trim();
        if (dEmail === email) {
          userName = dRow[1];
          // Use mentor preference if role is mentor and Column H exists, otherwise intern preference Column E
          userPref = (role === "MENTOR" && dRow[7]) ? String(dRow[7]).toUpperCase() : String(dRow[4]).toUpperCase();
        }
      });

      // Capitalize role cleanly for your sheet (e.g., "Intern" or "Mentor")
      let formattedRole = role.charAt(0) + role.slice(1).toLowerCase();

      // Append into your exact Reinvite tab structure (Columns A through H)
      reinviteSheet.appendRow([
        userName,                 // Column A: Name
        email,                    // Column B: Email
        "Didn't attend",          // Column C: Reason
        currentMonthName,         // Column D: Cohort
        "Send",                   // Column E: Email (Dropdown)
        userPref,                 // Column F: Which day (WD / WE)
        "NO",                     // Column G: Attended (Dropdown)
        formattedRole             // Column H: Role (Intern / Mentor)
      ]);

      existingReinviteEmails.push(email);

      // Send Re-invite email notification
      sendReinviteEmail(email, userName, formattedRole);

      // Auto-schedule into the next available shared calendar event matching their track
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
}

function sendReinviteEmail(email, name, role) {
  const subject = "Rescheduled: Your Partnership Course Make-up Session";
  const body = `Hi ${name},\n\n` +
               `We noticed you missed your scheduled ${role.toLowerCase()} session. Don't worry! We have automatically rolled your status over and re-invited you to the next available session.\n\n` +
               `Please check your calendar for the updated event link.\n\n` +
               `Best regards,\nOrganizing Committee`;
  MailApp.sendEmail(email, subject, body);
}