// Attendance lives in its own spreadsheet now, separate from the master file
// (Intern Details, Registration Form, Reinvite, etc). Mentor access is granted
// only on this file, so mentors never see other participants' contact info.
const ATTENDANCE_SPREADSHEET_ID = "1OsVIi_rUivLyZfXABOZr8gb8Yc1Pyz0d5dOkPA_OJgc";
function getAttendanceSpreadsheet() {
  return SpreadsheetApp.openById(ATTENDANCE_SPREADSHEET_ID);
}

function getCurrentMonthSheetName(trackType) {
  // Intern Details extra columns used by the reinvite system (beyond the original A-H):
  // Column I: Mentor invite attempts (see MentorResponses.gs)
  // Column J: Course reinvite attempts (missed-meeting rollovers, see rollInternToNextCohort below)
  // Column K: Course reinvite status ("Reinvited - attempt N" or "Stopped - no response after N invites")
  //
  // Monthly attendance sheet layout (columns): A unused/decorative, B Name,
  // C Email, D Discord ID, then 4 meeting blocks of 5 columns each (Attendance,
  // Tech, Participation, Comment, spacer) starting at column E:
  //   Meeting 1: E-H   Meeting 2: J-M   Meeting 3: O-R   Meeting 4: T-W
  // New tracking columns added right after Meeting 4's Comment column (W):
  //   Column X (24): Eligibility status ("Not eligible - missed Meeting N")
  //   Column Y (25): Course Result ("MASTERED" / "NOT MET")
  // See getMeetingColumns() below for the exact per-meeting column math.
  //
  // Rows: section header rows ("Weekday Mentors" / "Weekday Interns" etc, Name
  // column filled but no email) mark which section the rows below belong to.
  // Mentor-section rows are informational only - no automated consequences.
  const date = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  let currentMonthName = monthNames[date.getMonth()];
  let currentYearShort = String(date.getFullYear()).slice(-2);
  let trackAbbrev = (trackType === "Weekend") ? "WE" : "WD";

  return `${currentMonthName} ${trackAbbrev} ${currentYearShort}`;
}

/**
 * Returns the 1-indexed column numbers for one meeting's 4 data columns.
 * Each meeting block is 5 columns wide (4 data + 1 spacer), starting at
 * column E (5) for Meeting 1: E-H, J-M, O-R, T-W for meetings 1-4.
 */
function getMeetingColumns(meetingNum) {
  let blockStart = 5 + (meetingNum - 1) * 5;
  return {
    attendance: blockStart,
    tech: blockStart + 1,
    participation: blockStart + 2,
    comment: blockStart + 3
  };
}

const ELIGIBILITY_COL = 24; // Column X - right after Meeting 4's Comment column (W/23)
const COURSE_RESULT_COL = 25; // Column Y

/**
 * Given a row's values array (0-indexed) from the attendance sheet, returns
 * "Mentor", "Intern", or null for a section-header/blank row that isn't
 * itself a person. Section headers are detected by having text in the Name
 * column (B) but no valid email in Email (C) - e.g. "Weekday Mentors".
 */
function getRowSectionType(row) {
  let nameCell = String(row[1] || "").trim(); // Column B
  let email = row[2]; // Column C
  let hasValidEmail = email && String(email).includes("@");

  if (hasValidEmail) return "DATA"; // an actual person row - caller checks section separately
  if (!nameCell) return null; // blank spacer row

  let lower = nameCell.toLowerCase();
  if (lower.includes("mentor")) return "Mentor";
  if (lower.includes("intern")) return "Intern";
  return null; // some other non-data row, e.g. a stray label
}

/**
 * -------------------------------------------------------------------
 * 1. SYNC INTERNS INTO AN ALREADY-CREATED MONTHLY TAB
 * This does NOT build the tab, the meeting grid, section headers, or
 * formatting - your monthly tabs are hand-built/duplicated from a template
 * (see the Mentor/Intern section structure). This just syncs names, emails,
 * and Discord IDs from "Intern Details" into the existing "... Interns"
 * section of an already-created tab, filtered to the correct track.
 * -------------------------------------------------------------------
 */
function setupCurrentMonthAttendance(trackType) {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet(); // Intern Details lives here
  const attSS = getAttendanceSpreadsheet();                 // monthly attendance tabs live here
  let sheetName = getCurrentMonthSheetName(trackType);
  
  const detailsSheet = masterSS.getSheetByName("Intern Details");
  if (!detailsSheet) {
    SpreadsheetApp.getUi().alert("Could not find 'Intern Details' sheet!");
    return;
  }
  
  const detailsData = detailsSheet.getDataRange().getValues();
  
  let attSheet = attSS.getSheetByName(sheetName);
  if (!attSheet) {
    SpreadsheetApp.getUi().alert(
      `Tab "${sheetName}" doesn't exist yet in the attendance spreadsheet. This function only syncs ` +
      `names/emails into an already-created tab (e.g. duplicated from your template) - it doesn't build ` +
      `the meeting grid, section headers, or formatting. Create the tab first, then re-run.`
    );
    return;
  }

  // Find the "... Interns" section header row, so we know where to start writing
  // and never accidentally write into the Mentors section above it.
  let lastRow = attSheet.getLastRow();
  let existingValues = lastRow > 0 ? attSheet.getRange(1, 1, lastRow, 2).getValues() : [];
  let internSectionRow = -1;
  for (let r = 0; r < existingValues.length; r++) {
    let cellText = String(existingValues[r][1] || "").toLowerCase(); // Column B
    if (cellText.includes("intern")) { internSectionRow = r + 1; break; }
  }

  if (internSectionRow === -1) {
    SpreadsheetApp.getUi().alert(`Could not find a "... Interns" section header row in "${sheetName}" - add one first, then re-run.`);
    return;
  }

  let rowToWrite = internSectionRow + 1; // first row right after the "Interns" header
  let expectedTrack = (trackType === "Weekend") ? "WE" : "WD";

  // Loop through details and copy rows over - only the matching track
  for (let i = 1; i < detailsData.length; i++) {
    let row = detailsData[i];
    let name = row[1];       // Column B: Name
    let email = row[2];      // Column C: Email
    let discord = row[3];    // Column D: Discord
    let track = String(row[4] || "").toUpperCase(); // Column E: Intern WE/WD

    if (email && track === expectedTrack) {
      attSheet.getRange(rowToWrite, 2).setValue(name);    // Column B: Name
      attSheet.getRange(rowToWrite, 3).setValue(email);   // Column C: Email
      attSheet.getRange(rowToWrite, 4).setValue(discord); // Column D: Discord ID
      rowToWrite++;
    }
  }
  
  SpreadsheetApp.getUi().alert("Successfully synced interns into: " + sheetName);
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
  
  const attSS = getAttendanceSpreadsheet();
  const monthlySheet = attSS.getSheetByName(sheetName);
  if (!monthlySheet) return; // If the sheet hasn't been created yet, exit

  // Grant editor permission to the ATTENDANCE spreadsheet only - never the master
  // file, so mentors can't see Intern Details, Registration Form, or other tracks.
  let file = DriveApp.getFileById(attSS.getId());
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
  const attSS = getAttendanceSpreadsheet();
  const sheet = attSS.getSheetByName(sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Could not find sheet: " + sheetName);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  let currentSection = null; // "Mentor" or "Intern" - set by section header rows as we scan down

  // Loop through participants (starting at row 5, after headers)
  for (let i = 4; i < values.length; i++) {
    let row = values[i];

    let rowType = getRowSectionType(row);
    if (rowType === "Mentor" || rowType === "Intern") {
      currentSection = rowType; // hit a section header row - update and move on
      continue;
    }
    if (rowType === null) continue; // blank spacer row, not a person

    // rowType === "DATA" - an actual person row
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let name = row[1]; // Column B: Name

    let meetingsAttendedCount = 0;
    let techGoodCount = 0;
    let activeContributionCount = 0;

    let targetTechStatus = "";
    let targetParticipationStatus = "";
    let isPresentForTargetMeeting = false;

    // Loop through the 4 meetings in the grid
    for (let m = 1; m <= 4; m++) {
      let cols = getMeetingColumns(m);
      let isPresent = row[cols.attendance - 1] === true;
      let techStatus = row[cols.tech - 1];
      let participationStatus = row[cols.participation - 1];

      if (isPresent) {
        meetingsAttendedCount++;
        if (techStatus === "Everything is good") techGoodCount++;
        if (participationStatus === "Active contributor") activeContributionCount++;
      }

      // Capture data specifically for the meeting that just wrapped up
      if (m === currentMeetingNum) {
        isPresentForTargetMeeting = isPresent;
        targetTechStatus = techStatus;
        targetParticipationStatus = participationStatus;
      }
    }

    // Mentor rows are a purely informational record - no automated consequences
    // (no feedback email, no eligibility tracking, no rollover, no mastery eval).
    if (currentSection === "Mentor") continue;

    // Everything below only applies to the Intern section.

    let currentEligibility = String(row[ELIGIBILITY_COL - 1] || "").trim();
    let alreadyIneligible = currentEligibility.length > 0;

    // Send consolidated per-session feedback email if they attended THIS specific meeting -
    // but only if they haven't already missed an earlier meeting this cohort. Once someone
    // is marked ineligible, they stop receiving per-meeting feedback (they can still attend,
    // it just no longer counts toward this month's mastery).
    if (!alreadyIneligible && isPresentForTargetMeeting) {
      sendConsolidatedMeetingFeedback(email, currentMeetingNum, targetTechStatus, targetParticipationStatus);
    } else if (!alreadyIneligible && !isPresentForTargetMeeting) {
      // First miss this cohort - whichever meeting it is. Since mastery requires all
      // 4, this one miss already disqualifies them for this month. Record it, remove
      // them from the rest of this month's sessions, and roll them into next month's
      // cohort right away (pre-task not required again - it's one-time).
      sheet.getRange(i + 1, ELIGIBILITY_COL).setValue(`Not eligible - missed Meeting ${currentMeetingNum}`);
      removeFromRemainingSessionsThisMonth(email, trackType);
      rollInternToNextCohort(email, name, trackType, currentMeetingNum);
    }

    // If this is Meeting 4, run final mastery evaluation and send the milestone report
    if (currentMeetingNum === 4) {
      let hasAttendedAll = (meetingsAttendedCount === 4);
      let hasGoodTechTwice = (techGoodCount >= 2);
      let hasBeenActiveTwice = (activeContributionCount >= 2);

      let finalMasteryAchieved = hasAttendedAll && hasGoodTechTwice && hasBeenActiveTwice;

      sheet.getRange(i + 1, COURSE_RESULT_COL).setValue(finalMasteryAchieved ? "MASTERED" : "NOT MET");

      // Send the final milestone email report
      sendFinalMasteryEmail(email, finalMasteryAchieved, meetingsAttendedCount, techGoodCount, activeContributionCount);

      // Bridge into Intern Details: mark them "Eligible" for a mentor invite so
      // sendMentorInvitationEmails() (MentorResponses.gs) will pick them up.
      if (finalMasteryAchieved) {
        markInternEligibleForMentor(email);
      }
    }
  }
}

/**
 * Removes someone from the rest of this month's course sessions (not past ones -
 * those already happened) once they've missed a meeting and been rolled to next
 * month. Searches the same "Partnership Course {trackType}" title PreTaskTracker.gs
 * uses when creating these events.
 */
function removeFromRemainingSessionsThisMonth(email, trackType) {
  const calendar = CalendarApp.getDefaultCalendar();
  const calendarId = calendar.getId();
  const now = new Date();
  const windowEnd = new Date(now.getTime() + 35 * 24 * 60 * 60 * 1000); // covers the rest of this cohort
  const eventTitle = `Partnership Course ${trackType}`;

  // Sessions are now one recurring series sharing a single Meet link (see
  // PreTaskTracker.gs). To remove someone from only the REMAINING occurrences -
  // without touching past ones or the series' shared link - we expand the series
  // into individual instances (singleEvents: true) and patch each future one's
  // attendee list separately. This creates a per-occurrence override rather than
  // editing the series itself.
  let response = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    timeMax: windowEnd.toISOString(),
    q: eventTitle,
    singleEvents: true,
    orderBy: 'startTime'
  });

  let items = response.items || [];
  items.forEach(item => {
    if (!item.summary || item.summary.trim() !== eventTitle) return;
    let attendees = item.attendees || [];
    let filtered = attendees.filter(a => a.email && a.email.toLowerCase() !== email.toLowerCase());
    if (filtered.length !== attendees.length) {
      Calendar.Events.patch({ attendees: filtered }, calendarId, item.id);
    }
  });
}

/**
 * Rolls an intern who just missed a meeting into next month's cohort: re-adds them
 * to the upcoming course sessions and tech check (skipping pre-task, since that's a
 * one-time requirement already satisfied), and sends a reinvite email. Tracks
 * attempts in Intern Details Column J, capped at 3 total invites - after that,
 * Column K is set to a clear final status and no further invites go out.
 */
function rollInternToNextCohort(email, name, trackType, missedMeetingNum) {
  const MAX_COURSE_REINVITE_ATTEMPTS = 3;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  const detailsData = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 11).getValues();
  const targetEmail = String(email).toLowerCase().trim();

  for (let i = 0; i < detailsData.length; i++) {
    let rowEmail = String(detailsData[i][2]).toLowerCase().trim(); // Column C
    if (rowEmail !== targetEmail) continue;

    let rowIndex = i + 2;
    let attempts = Number(detailsData[i][9]) || 0;              // Column J
    let existingStatus = String(detailsData[i][10] || "").trim(); // Column K

    if (existingStatus.indexOf("Stopped") === 0) return; // already capped out, leave alone

    if (attempts >= MAX_COURSE_REINVITE_ATTEMPTS) {
      detailsSheet.getRange(rowIndex, 11).setValue(`Stopped - no response after ${attempts} invites`); // Column K
      return;
    }

    let isWeekend = (trackType === "Weekend");
    let newAttemptNum = attempts + 1;

    // Re-add to next month's sessions and tech check (functions live in PreTaskTracker.gs).
    // No pre-task step here - it's a one-time requirement they've already met.
    addStudentToAllFourCourseSessions(email, isWeekend);
    addStudentToGroupTechCheck(email, isWeekend);

    const cohortDates = getCalculatedCohortDates(); // DateCalculator.gs
    sendCourseReinviteEmail(name, email, isWeekend ? cohortDates.sundayList : cohortDates.thursdayList, newAttemptNum, MAX_COURSE_REINVITE_ATTEMPTS, missedMeetingNum);

    detailsSheet.getRange(rowIndex, 10).setValue(newAttemptNum); // Column J
    detailsSheet.getRange(rowIndex, 11).setValue(`Reinvited - attempt ${newAttemptNum}`); // Column K
    return;
  }
}

/**
 * Finds this email in "Intern Details" and sets their Mentor Status (Column G)
 * to "Eligible", but only if they don't already have a status (so we don't
 * overwrite someone who already declined, accepted, or was already invited).
 */
function markInternEligibleForMentor(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  const detailsData = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8).getValues();
  const targetEmail = String(email).toLowerCase().trim();

  for (let i = 0; i < detailsData.length; i++) {
    let rowEmail = String(detailsData[i][2]).toLowerCase().trim(); // Column C: Email
    if (rowEmail === targetEmail) {
      let currentMentorStatus = String(detailsData[i][6] || "").trim(); // Column G
      if (currentMentorStatus === "") {
        detailsSheet.getRange(i + 2, 7).setValue("Eligible"); // Column G
      }
      break;
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

function sendCourseReinviteEmail(name, email, sessionDatesList, attemptNum, maxAttempts, missedMeetingNum) {
  const subject = "Missed a Session - You're Rescheduled for Next Month's Partnership Course";
  const body = `
    <p>Hello ${name},</p>
    <p>We noticed you weren't able to attend <strong>Meeting ${missedMeetingNum}</strong>. Since course mastery requires attending all 4 meetings, you're no longer eligible to continue this month's cohort, and you've been removed from the remaining sessions this month.</p>
    <p>The good news: you're already re-enrolled for next month's cohort, and your pre-task is a one-time requirement, so you do not need to submit it again.</p>
    <p><strong>Your Calendar Invites Have Been Sent:</strong></p>
    <ul>
      <li><strong>Tech Check:</strong> 1 hour before your first session. It is <strong>Mandatory</strong> to join tech check</li>
      <li><strong>Partnership Course Sessions:</strong> Scheduled across your new cohort dates (${sessionDatesList})</li>
    </ul>
    <p>This is invite ${attemptNum} of ${maxAttempts}.</p>
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

/**
 * -------------------------------------------------------------------
 * 6. DAY-AWARE SCHEDULING (Thursday for Weekday, Sunday for Weekend)
 * -------------------------------------------------------------------
 * The functions above (runWeekdayMeeting1..4) require YOU to know which
 * week of the cohort it is and call the right one manually. These two
 * functions instead figure that out from today's date automatically, so
 * you can put them on a single weekly Thursday / Sunday trigger instead
 * of remembering to swap which meeting number to run each week.
 */

/**
 * Returns which meeting number (1-4) today corresponds to for the current
 * cohort, or null if today isn't a scheduled meeting day this month (e.g.
 * it's a 5th Thursday/Sunday, or some other day entirely).
 */
function getCurrentCohortMeetingNumber(isWeekend) {
  const now = new Date();
  let firstOfMonth = new Date(Date.UTC(now.getFullYear(), now.getMonth(), 1));
  let targetDayOfWeek = isWeekend ? 0 : 4; // Sunday : Thursday
  let dayOffset = (targetDayOfWeek - firstOfMonth.getUTCDay() + 7) % 7;
  firstOfMonth.setUTCDate(firstOfMonth.getUTCDate() + dayOffset); // Meeting 1's date this month

  let today = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  let daysSinceMeeting1 = Math.round((today - firstOfMonth) / (24 * 60 * 60 * 1000));

  if (daysSinceMeeting1 < 0 || daysSinceMeeting1 % 7 !== 0) return null; // not a session day
  let meetingNum = (daysSinceMeeting1 / 7) + 1;
  return (meetingNum >= 1 && meetingNum <= 4) ? meetingNum : null;
}

function runScheduledWeekdayAttendance() {
  let meetingNum = getCurrentCohortMeetingNumber(false);
  if (!meetingNum) { Logger.log("Today isn't a scheduled Weekday (Thursday) meeting day."); return; }
  processMonthlyAttendance("Weekday", meetingNum);
}

function runScheduledWeekendAttendance() {
  let meetingNum = getCurrentCohortMeetingNumber(true);
  if (!meetingNum) { Logger.log("Today isn't a scheduled Weekend (Sunday) meeting day."); return; }
  processMonthlyAttendance("Weekend", meetingNum);
}

/**
 * Run this once from the Apps Script editor to install the weekly triggers.
 * Deletes any old daily/manual triggers on these handler names first, so
 * re-running this is safe and won't create duplicates.
 *
 * Times below assume the Apps Script project's timezone is set to Nepal
 * (Asia/Kathmandu, UTC+5:45) - check Project Settings > Time zone before
 * relying on this. Computed from actual class times:
 *   Weekday: 11:00-13:00 UTC = 4:45pm-6:45pm Nepal time (class ends 6:45pm)
 *   Weekend: 14:00-16:00 UTC = 7:45pm-9:45pm Nepal time (class ends 9:45pm)
 * Both stay on the same calendar day in Nepal time, no Thursday/Sunday shift needed.
 */
function createScheduledAttendanceTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    let fn = tr.getHandlerFunction();
    if (fn === "runScheduledWeekdayAttendance" || fn === "runScheduledWeekendAttendance") {
      ScriptApp.deleteTrigger(tr);
    }
  });

  // ~1 hour after class ends: Weekday ends 6:45pm -> 7:45pm; Weekend ends 9:45pm -> 10:45pm
  ScriptApp.newTrigger("runScheduledWeekdayAttendance").timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(19).nearMinute(45).create();
  ScriptApp.newTrigger("runScheduledWeekendAttendance").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(22).nearMinute(45).create();

  // Reinvite check (ReinviteAutomation.gs) runs ~1 hour after attendance is processed
  triggers.forEach(tr => {
    let fn = tr.getHandlerFunction();
    if (fn === "runScheduledWeekdayReinviteCheck" || fn === "runScheduledWeekendReinviteCheck") {
      ScriptApp.deleteTrigger(tr);
    }
  });
  ScriptApp.newTrigger("runScheduledWeekdayReinviteCheck").timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(20).nearMinute(45).create();
  ScriptApp.newTrigger("runScheduledWeekendReinviteCheck").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).nearMinute(45).create();
}