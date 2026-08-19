function trackCohortAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const now = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const currentMonthTab = "Attendance - " + monthNames[now.getMonth()] + " " + now.getFullYear();
  
  let attSheet = ss.getSheetByName(currentMonthTab);
  if (!attSheet) {
    attSheet = ss.insertSheet(currentMonthTab);
    attSheet.appendRow([
      "Timestamp", 
      "Email", 
      "Course Day", 
      "Course Date", 
      "Role", 
      "Attendance Status", 
      "Email Sent"
    ]);
  }

  let detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;
  
  let detailsRange = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8);
  let detailsData = detailsRange.getValues();
  
  let mentorEmails = [];
  
  detailsData.forEach((row, index) => {
    let email = row[2]; // Column C: Email
    let mentorStatus = String(row[6]).toLowerCase(); // Column G: Mentor
    
    if (email && (mentorStatus.startsWith("mento") || mentorStatus === "invited" || mentorStatus.includes("yes"))) {
      mentorEmails.push({
        email: email.toLowerCase().trim(),
        rowIndex: index + 2
      });
    }
  });

  const calendar = CalendarApp.getDefaultCalendar();
  const startTime = new Date(now.getTime() - (24 * 60 * 60 * 1000));
  const endTime = new Date(now.getTime());

  const events = calendar.getEvents(startTime, endTime, { search: "Partnership Course" });

  if (events.length === 0) {
    Logger.log("No recent Partnership Course sessions found to track.");
    return;
  }

  events.forEach(event => {
    let eventTitle = event.getTitle(); 
    let eventDate = event.getStartTime().toLocaleDateString(); 
    let guests = event.getGuestList();

    guests.forEach(guest => {
      let guestEmail = guest.getEmail().toLowerCase().trim();
      let status = guest.getGuestStatus();
      
      let foundMentor = mentorEmails.find(m => m.email === guestEmail);
      let role = foundMentor ? "MENTOR" : "INTERN";
      let attendanceRecord = (status === CalendarApp.GuestStatus.YES) ? "PRESENT" : "ABSENT";

      attSheet.appendRow([
        new Date(), 
        guestEmail, 
        eventTitle, 
        eventDate, 
        role, 
        attendanceRecord
      ]);
      
      let lastRow = attSheet.getLastRow();
      attSheet.getRange(lastRow, 7).insertCheckboxes();

      if (foundMentor && attendanceRecord === "PRESENT") {
        detailsSheet.getRange(foundMentor.rowIndex, 7).setValue("Mentor Done"); 
      }
    });
  });

  updateInternProgress(ss, detailsSheet, currentMonthTab);
}

function updateInternProgress(spreadsheetObj, detailsSheet, monthTabName) {
  if (!spreadsheetObj) return;
  let attSheet = spreadsheetObj.getSheetByName(monthTabName);
  if (!attSheet || attSheet.getLastRow() <= 1) return;

  let attData = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, 6).getValues();
  let detailsData = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8).getValues();

  let attendanceCounts = {};
  attData.forEach(row => {
    let email = String(row[1]).toLowerCase().trim();
    let status = row[5]; 
    let role = row[4];   

    if (role === "INTERN" && status === "PRESENT") {
      attendanceCounts[email] = (attendanceCounts[email] || 0) + 1;
    }
  });

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let name = row[1];                              // Column B: Intern Names
    let rowIndex = index + 2;
    let currentMentorCol = String(row[6]);          // Column G: Mentor

    if (attendanceCounts[email] === 4 && (currentMentorCol === "" || currentMentorCol.toLowerCase() === "no")) {
      detailsSheet.getRange(rowIndex, 7).setValue("Invited");
      sendMentorInviteEmail(email, name);
    }
  });
}

function sendMentorInviteEmail(email, name) {
  const subject = "Invitation to become a Mentor";
  const body = `Hi ${name || "there"},\n\n` +
               `Congratulations on completing your 4 classes! We would love to invite you to mentor our next cohort.\n` +
               `Please reply to let us know if you are interested and whether you prefer Weekday (WD) or Weekend (WE) sessions.\n\n` +
               `Best regards,\nOrganizing Committee`;
  MailApp.sendEmail(email, subject, body);
}

function createAttendanceTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "trackCohortAttendance") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("trackCohortAttendance")
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();
}