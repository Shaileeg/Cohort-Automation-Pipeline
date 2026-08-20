function getCurrentMonthSheetName(trackType) {
  const date = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  let currentMonthName = monthNames[date.getMonth()];
  let currentYear = date.getFullYear();
  
  return `Attendance - ${currentMonthName} ${currentYear} - ${trackType}`;
}

/**
 * -------------------------------------------------------------------
 * 1. AUTO-SYNC PARTICIPANTS (RUN THIS AT THE START OF EACH MONTH)
 * Creates the current month's tab dynamically and pulls data from "Intern Details"
 * -------------------------------------------------------------------
 */
function setupCurrentMonthAttendance(trackType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = getCurrentMonthSheetName(trackType);
  
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet) {
    SpreadsheetApp.getUi().alert("Could not find 'Intern Details' sheet!");
    return;
  }
  
  const detailsData = detailsSheet.getDataRange().getValues();
  
  let attSheet = ss.getSheetByName(sheetName);
  if (!attSheet) {
    attSheet = ss.insertSheet(sheetName);
    // Set up a simple header
    attSheet.getRange("A1:B4").merge().setValue(sheetName).setFontSize(14).setFontWeight("bold");
  }

  let rowToWrite = 5; // Start writing participant data from row 5 downward
  
  // Loop through details and copy rows over
  for (let i = 1; i < detailsData.length; i++) {
    let row = detailsData[i];
    let name = row[1];       // Assuming Name is in Column B (Index 1)
    let email = row[2];      // Assuming Email is in Column C (Index 2)

    if (email) {
      attSheet.getRange(rowToWrite, 1).setValue(name);  // Column A: Name
      attSheet.getRange(rowToWrite, 2).setValue(email); // Column B: Email/ID
      rowToWrite++;
    }
  }
  
  SpreadsheetApp.getUi().alert("Successfully created and synced: " + sheetName);
}

/**
 * -------------------------------------------------------------------
 * 2. MENTOR ACCESS & PERMISSION SENDER
 * Connect this to your Mentor Email Reply script. 
 * It grants access and emails them the live dynamic link.
 * -------------------------------------------------------------------
 */
function grantMentorTabAccess(mentorEmail, trackPreference) {
  let trackType = (trackPreference === "WD") ? "Weekday" : "Weekend";
  let sheetName = getCurrentMonthSheetName(trackType);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthlySheet = ss.getSheetByName(sheetName);
  if (!monthlySheet) return; // If the sheet hasn't been created yet, exit

  // Grant editor permission to the Google Sheet file
  let file = DriveApp.getFileById(ss.getId());
  file.addEditor(mentorEmail);

  // Automatically fetch the live URL of the specific tab
  let sheetUrl = monthlySheet.getUrl();

  let subject = `Your Mentor Access: ${trackType} Attendance Sheet`;
  let body = `
    <p>Hello,</p>
    <p>Thank you for accepting the invitation to become a mentor! You have been assigned to the <strong>${trackType}</strong> track for this month.</p>
    <p>You can now access your attendance and feedback sheet here: <a href="${sheetUrl}">Open Monthly Attendance Sheet</a></p>
    <p>Please use this sheet after each session to log attendance, tech checks, and participation feedback.</p>
    <p>Best regards,</p>
    <p>The Partnership Team</p>
  `;

  GmailApp.sendEmail(mentorEmail, subject, "", { htmlBody: body });
}

/**
 * -------------------------------------------------------------------
 * 3. ATTENDANCE & MASTERY EVALUATION ENGINE
 * Scans the sheet, sends unified post-session feedback, and evaluates mastery.
 * -------------------------------------------------------------------
 */
function processMonthlyAttendance(trackType, currentMeetingNum) {
  let sheetName = getCurrentMonthSheetName(trackType);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Could not find sheet: " + sheetName);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Loop through participants (starting at row 5, after headers)
  for (let i = 4; i < values.length; i++) {
    let row = values[i];
    let email = row[1]; // Column B: Email reference identifier
    if (!email || !email.includes("@")) continue;

    let meetingsAttendedCount = 0;
    let techGoodCount = 0;
    let activeContributionCount = 0;

    let targetTechStatus = "";
    let targetParticipationStatus = "";
    let isPresentForTargetMeeting = false;

    // Loop through the 4 meetings in the grid
    for (let m = 0; m < 4; m++) {
      let attColIndex = 2 + (m * 3);   
      let techColIndex = 3 + (m * 3);  
      let partColIndex = 4 + (m * 3);  

      let isPresent = row[attColIndex] === true;
      let techStatus = row[techColIndex];
      let participationStatus = row[partColIndex];

      if (isPresent) {
        meetingsAttendedCount++;
        if (techStatus === "Everything is good") techGoodCount++;
        if (participationStatus === "Active contributor") activeContributionCount++;
      }

      // Capture data specifically for the meeting that just wrapped up
      if ((m + 1) === currentMeetingNum) {
        isPresentForTargetMeeting = isPresent;
        targetTechStatus = techStatus;
        targetParticipationStatus = participationStatus;
      }
    }

    // Send consolidated per-session feedback email if they attended THIS specific meeting
    if (isPresentForTargetMeeting) {
      sendConsolidatedMeetingFeedback(email, currentMeetingNum, targetTechStatus, targetParticipationStatus);
    }

    // If this is Meeting 4, run final mastery evaluation and send the milestone report
    if (currentMeetingNum === 4) {
      let hasAttendedAll = (meetingsAttendedCount === 4);
      let hasGoodTechTwice = (techGoodCount >= 2);
      let hasBeenActiveTwice = (activeContributionCount >= 2);

      let finalMasteryAchieved = hasAttendedAll && hasGoodTechTwice && hasBeenActiveTwice;

      // Write final result to the Course Result column (Column Y / Index 24)
      let resultColIndex = 25; 
      sheet.getRange(i + 1, resultColIndex).setValue(finalMasteryAchieved ? "MASTERED" : "NOT MET");

      // Send the final milestone email report
      sendFinalMasteryEmail(email, finalMasteryAchieved, meetingsAttendedCount, techGoodCount, activeContributionCount);
    }
  }
}

/**
 * -------------------------------------------------------------------
 * 4. EMAIL DISPATCH HELPERS
 * -------------------------------------------------------------------
 */
function sendConsolidatedMeetingFeedback(email, meetingNum, techStatus, participationStatus) {
  let subject = `Attendance & Session Feedback - Meeting ${meetingNum}`;
  
  let techMessage = "";
  if (techStatus === "Everything is good") {
    techMessage = "<li><strong>Tech Check:</strong> Thank you for attending meeting! Your setup was fully functional.</li>";
  } else {
    techMessage = `<li><strong>Tech Check:</strong> You selected/experienced (${techStatus || "Not specified"}). Please fix issues for the next class as it is necessary for mastery.</li>`;
  }

  let participationMessage = "";
  if (participationStatus === "Active contributor") {
    participationMessage = "<li><strong>Participation:</strong> Good, your participation was active, keep it up!</li>";
  } else {
    participationMessage = `<li><strong>Participation:</strong> Your participation level was marked as (${participationStatus || "Not specified"}). Please do interact more in the meeting to enhance more and learn more.</li>`;
  }

  let body = `
    <p>Hello,</p>
    <p>Here is your automated performance and attendance feedback report for <strong>Meeting ${meetingNum}</strong>:</p>
    <ul>
      ${techMessage}
      ${participationMessage}
    </ul>
    <p>Keep reviewing your progress and working toward your course mastery goals!</p>
    <p>Best regards,</p>
    <p>The Partnership Team</p>
  `;

  GmailApp.sendEmail(email, subject, "", { htmlBody: body });
}

function sendFinalMasteryEmail(email, isMastered, attendanceCount, techCount, activeCount) {
  let subject = `Course Completion & Final Mastery Result`;
  let statusText = isMastered ? 
    "<span style='color: green; font-weight: bold;'>MASTERED 🎉</span>" : 
    "<span style='color: orange; font-weight: bold;'>NOT MET (Keep practicing!)</span>";

  let body = `
    <p>Hello,</p>
    <p>Your 4-class cohort journey has concluded! Here is your final evaluation summary:</p>
    <ul>
      <li><strong>Final Status:</strong> ${statusText}</li>
      <li><strong>Classes Attended:</strong> ${attendanceCount} / 4</li>
      <li><strong>Good Tech Check Days:</strong> ${techCount} (Required: at least 2)</li>
      <li><strong>Active Participation Days:</strong> ${activeCount} (Required: at least 2)</li>
    </ul>
    <p>Thank you for participating and working hard to improve your skills!</p>
    <p>Best regards,</p>
    <p>The Partnership Team</p>
  `;

  GmailApp.sendEmail(email, subject, "", { htmlBody: body });
}

/**
 * -------------------------------------------------------------------
 * 5. QUICK EXECUTION TRIGGERS (DYNAMIC)
 * Run these from the Apps Script editor menu. 
 * They automatically adapt to the current month/year!
 * -------------------------------------------------------------------
 */
// SETUP TRIGGERS
function runSetupCurrentMonthWeekday() { setupCurrentMonthAttendance("Weekday"); }
function runSetupCurrentMonthWeekend() { setupCurrentMonthAttendance("Weekend"); }

// WEEKDAY MEETING PROCESSORS (Run 5-6 hours after class)
function runWeekdayMeeting1() { processMonthlyAttendance("Weekday", 1); }
function runWeekdayMeeting2() { processMonthlyAttendance("Weekday", 2); }
function runWeekdayMeeting3() { processMonthlyAttendance("Weekday", 3); }
function runWeekdayMeeting4() { processMonthlyAttendance("Weekday", 4); }

// WEEKEND MEETING PROCESSORS (Run 5-6 hours after class)
function runWeekendMeeting1() { processMonthlyAttendance("Weekend", 1); }
function runWeekendMeeting2() { processMonthlyAttendance("Weekend", 2); }
function runWeekendMeeting3() { processMonthlyAttendance("Weekend", 3); }
function runWeekendMeeting4() { processMonthlyAttendance("Weekend", 4); }