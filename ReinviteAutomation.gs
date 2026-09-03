/**
 * Scans ONE track's monthly attendance record (Weekday or Weekend - call this once
 * per track, same pattern as processMonthlyAttendance). If an intern or mentor is
 * marked ABSENT for the most recent meeting (checkbox unchecked for that meeting's
 * column), it automatically logs them into the "Reinvite" tab with their details,
 * role, sets the email dropdown status to "Send", records their preference, and
 * syncs them to the next upcoming calendar event.
 *
 * Anyone already sitting in the Reinvite tab as unresolved (Column G = "NO")
 * gets re-checked against this meeting too: if they showed up, they're marked
 * resolved; if they missed the make-up as well, they get re-invited again, up
 * to MAX_FOLLOWUP_ATTEMPTS times, after which it's flagged for manual follow-up
 * instead of emailing them indefinitely.
 *
 * Reinvite tab columns: A Name, B Email, C Reason, D Cohort, E Email status,
 * F Track (WD/WE), G Attended (YES/NO), H Role, I Follow-up attempts,
 * J Last meeting checked.
 *
 * @param {string} trackType - "Weekday" or "Weekend". Each track runs on a different
 *   day (Thursday / Sunday) and can be on a different meeting number in the same
 *   calendar week, so this must be checked per-track, not both at once.
 * @param {number} currentMeetingNum - Which of the 4 meetings just wrapped up (1-4)
 *   for THIS track. Call this the same way you call processMonthlyAttendance, right
 *   after a session, e.g. processMissedSessionsAndReinvites("Weekday", 2).
 */
function processMissedSessionsAndReinvites(trackType, currentMeetingNum) {
  if (trackType !== "Weekday" && trackType !== "Weekend") {
    Logger.log("processMissedSessionsAndReinvites requires trackType to be 'Weekday' or 'Weekend'.");
    return;
  }
  if (!currentMeetingNum || currentMeetingNum < 1 || currentMeetingNum > 4) {
    Logger.log("processMissedSessionsAndReinvites requires a valid currentMeetingNum (1-4).");
    return;
  }

  const MAX_FOLLOWUP_ATTEMPTS = 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const now = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const currentMonthName = monthNames[now.getMonth()];
  const currentYear = now.getFullYear();
  
  let detailsSheet = ss.getSheetByName("Intern Details");
  let detailsData = detailsSheet && detailsSheet.getLastRow() > 1 ? detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8).getValues() : [];

  let reinviteSheet = ss.getSheetByName("Reinvite");
  if (!reinviteSheet) return;

  let existingReinviteEmails = reinviteSheet.getLastRow() > 2 ? 
      reinviteSheet.getRange(3, 2, reinviteSheet.getLastRow() - 2, 1).getValues().flat().map(e => String(e).toLowerCase().trim()) : [];

  // Load full existing Reinvite records so we can resolve/escalate unresolved ones
  let reinviteRecords = {}; // email -> { rowNum, cohort, track, attended, attempts, lastChecked }
  let reinviteLastRow = reinviteSheet.getLastRow();
  if (reinviteLastRow > 2) {
    let allReinviteData = reinviteSheet.getRange(3, 1, reinviteLastRow - 2, 10).getValues();
    allReinviteData.forEach((r, idx) => {
      let rEmail = String(r[1]).toLowerCase().trim();
      if (!rEmail) return;
      reinviteRecords[rEmail] = {
        rowNum: idx + 3,
        cohort: r[3],
        track: r[5],
        attended: r[6],
        attempts: Number(r[8]) || 0,
        lastChecked: Number(r[9]) || 0
      };
    });
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const futureTime = new Date(now.getTime() + (60 * 24 * 60 * 60 * 1000)); 
  const events = calendar.getEvents(now, futureTime, { search: "Partnership Course" });

  function reAddToUpcomingEvent(email, userPref) {
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

  // This track's tab lives in the separate attendance spreadsheet;
  // Intern Details and Reinvite above stay in the master file.
  const attSS = getAttendanceSpreadsheet(); // defined in AttandanceTracker.gs
  let tabName = getCurrentMonthSheetName(trackType); // defined in AttandanceTracker.gs
  let attSheet = attSS.getSheetByName(tabName);
  if (!attSheet || attSheet.getLastRow() <= 4) return; // Skip if tab doesn't exist or has no participant data

  // Participant data starts at row 5 in your grid layout
  let numRows = attSheet.getLastRow() - 4;
  let attData = attSheet.getRange(5, 1, numRows, attSheet.getLastColumn()).getValues();

  let currentSection = null; // "Mentor" or "Intern" - set by section header rows as we scan down

  attData.forEach(row => {
      let rowType = getRowSectionType(row); // defined in AttandanceTracker.gs
      if (rowType === "Mentor" || rowType === "Intern") {
        currentSection = rowType;
        return;
      }
      if (rowType === null) return; // blank spacer row, not a person

      // Mentor attendance is informational only - no automated reinvite/make-up chasing.
      if (currentSection === "Mentor") return;

      let email = String(row[2]).toLowerCase().trim(); // Column C: Email
      let userName = row[1]; // Column B: Name

      // Skip entirely once they're already marked ineligible for this month's mastery
      // (set by AttandanceTracker.gs the moment someone misses any meeting). A make-up
      // session for one missed meeting can't restore eligibility once that's happened,
      // so there's no reason to chase them with a meeting-level reinvite.
      let eligibilityStatus = String(row[ELIGIBILITY_COL - 1] || "").trim(); // defined in AttandanceTracker.gs
      if (eligibilityStatus.length > 0) return;

      let userPref = (trackType === "Weekday") ? "WD" : "WE";

      // Attendance checkbox for the meeting that just wrapped up.
      let targetCols = getMeetingColumns(currentMeetingNum); // defined in AttandanceTracker.gs
      let wasPresentForTargetMeeting = row[targetCols.attendance - 1] === true;
      let isAbsentStatus = !wasPresentForTargetMeeting;

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

      // --- Case 1: already on the Reinvite tab, unresolved, and not yet checked for this meeting ---
      let existing = reinviteRecords[email];
      if (existing && existing.attended === "NO" && existing.cohort === currentMonthName &&
          existing.track === userPref && existing.lastChecked < currentMeetingNum) {

        reinviteSheet.getRange(existing.rowNum, 10).setValue(currentMeetingNum); // Column J: last checked

        if (wasPresentForTargetMeeting) {
          // They showed up to the make-up session: resolved
          reinviteSheet.getRange(existing.rowNum, 7).setValue("YES"); // Column G
        } else {
          // Missed again: escalate the attempt count
          let newAttempts = existing.attempts + 1;
          reinviteSheet.getRange(existing.rowNum, 9).setValue(newAttempts); // Column I

          if (newAttempts < MAX_FOLLOWUP_ATTEMPTS) {
            sendReinviteEmail(email, userName, formattedRole);
            reAddToUpcomingEvent(email, userPref);
          } else {
            reinviteSheet.getRange(existing.rowNum, 3).setValue(
              `Missed make-up ${newAttempts} times - needs manual follow-up`
            ); // Column C
          }
        }
        return; // handled as a follow-up, not a brand-new absence
      }

      // --- Case 2: brand-new absence, not already tracked ---
      if (isAbsentStatus && !existingReinviteEmails.includes(email)) {
        reinviteSheet.appendRow([
          userName,               // Column A: Name
          email,                  // Column B: Email
          "Didn't attend",        // Column C: Reason
          currentMonthName,       // Column D: Cohort
          "Send",                 // Column E: Email (Dropdown)
          userPref,               // Column F: Which day (WD / WE)
          "NO",                   // Column G: Attended (Dropdown)
          formattedRole,          // Column H: Role (Intern / Mentor)
          1,                      // Column I: Follow-up attempts
          currentMeetingNum       // Column J: Last meeting checked
        ]);

        existingReinviteEmails.push(email);
        sendReinviteEmail(email, userName, formattedRole);
        reAddToUpcomingEvent(email, userPref);
      }
  });
}

/**
 * -------------------------------------------------------------------
 * QUICK EXECUTION TRIGGERS
 * Run these (or hook them to time-based triggers) a few hours after
 * each meeting, same pattern as AttandanceTracker.gs's runWeekdayMeetingN().
 * -------------------------------------------------------------------
 */
function runWeekdayReinviteCheckMeeting1() { processMissedSessionsAndReinvites("Weekday", 1); }
function runWeekdayReinviteCheckMeeting2() { processMissedSessionsAndReinvites("Weekday", 2); }
function runWeekdayReinviteCheckMeeting3() { processMissedSessionsAndReinvites("Weekday", 3); }
function runWeekdayReinviteCheckMeeting4() { processMissedSessionsAndReinvites("Weekday", 4); }

function runWeekendReinviteCheckMeeting1() { processMissedSessionsAndReinvites("Weekend", 1); }
function runWeekendReinviteCheckMeeting2() { processMissedSessionsAndReinvites("Weekend", 2); }
function runWeekendReinviteCheckMeeting3() { processMissedSessionsAndReinvites("Weekend", 3); }
function runWeekendReinviteCheckMeeting4() { processMissedSessionsAndReinvites("Weekend", 4); }

/**
 * Day-aware version: figures out today's meeting number automatically
 * (getCurrentCohortMeetingNumber lives in AttandanceTracker.gs), so these can
 * go on a single weekly Thursday / Sunday trigger instead of manually picking
 * which meeting number to run.
 */
function runScheduledWeekdayReinviteCheck() {
  let meetingNum = getCurrentCohortMeetingNumber(false);
  if (!meetingNum) { Logger.log("Today isn't a scheduled Weekday (Thursday) meeting day."); return; }
  processMissedSessionsAndReinvites("Weekday", meetingNum);
}

function runScheduledWeekendReinviteCheck() {
  let meetingNum = getCurrentCohortMeetingNumber(true);
  if (!meetingNum) { Logger.log("Today isn't a scheduled Weekend (Sunday) meeting day."); return; }
  processMissedSessionsAndReinvites("Weekend", meetingNum);
}

function sendReinviteEmail(email, name, role) {
  const subject = "Rescheduled: Your Partnership Course Make-up Session";
  const body = `Hi ${name},\n\n` +
               `We noticed you missed your scheduled ${role.toLowerCase()} session. Don't worry! We have automatically rolled your status over and re-invited you to the next available session.\n\n` +
               `Please check your calendar for the updated event link.\n\n` +
               `Best regards,\nThe Partnership Team`;
  MailApp.sendEmail(email, subject, body);
}