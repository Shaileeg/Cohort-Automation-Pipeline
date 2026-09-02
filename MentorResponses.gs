/**
 * Intern Details extra columns used by the reinvite system (beyond the original A-H):
 * Column I: Mentor invite attempts (number, max 3)
 * Column J: Course reinvite attempts (number, max 3) - used by handleCohortNoShow in AttandanceTracker.gs
 * Column K: Course reinvite status (e.g. "Reinvited - attempt 2", "Stopped - No response after 3 invites")
 * Column L: Manual "Add as Mentor" checkbox - check it to immediately send that
 *   row an invite, going through the exact same logic/attempt-tracking as the
 *   automated path (see onEditInternDetails below).
 */
const MAX_MENTOR_INVITE_ATTEMPTS = 3;
const MANUAL_MENTOR_CHECKBOX_COL = 12; // Column L

/**
 * Sends the mentor invite email to one row and updates its tracking columns.
 * Shared by both sendMentorInvitationEmails() (automated) and
 * onEditInternDetails() (manual checkbox), so both paths behave identically.
 */
function sendMentorInvitationToRow(detailsSheet, rowIndex, attemptsSoFar) {
  let rowValues = detailsSheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
  let email = String(rowValues[2]).toLowerCase().trim(); // Column C
  let name = rowValues[1]; // Column B

  let subject = "Invitation to become a Mentor";
  let message = `Hi ${name},\n\n` +
                `Congratulations on reaching your milestones! We would love to invite you to become a mentor for our upcoming sessions.\n\n` +
                `Please reply to this email with "Yes" if you would like to mentor, or "No" if you cannot make it.\n` +
                `If you say Yes, you can also choose your track preference by including "Weekday" (WD) or "Weekend" (WE) in your reply.\n\n` +
                `Best regards,\nThe Partnership Team`;

  MailApp.sendEmail(email, subject, message);

  detailsSheet.getRange(rowIndex, 7).setValue("Invite Sent"); // Column G
  detailsSheet.getRange(rowIndex, 9).setValue(attemptsSoFar + 1); // Column I
}

function sendMentorInvitationEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  let detailsRange = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 9);
  let detailsData = detailsRange.getValues();

  detailsData.forEach((row, index) => {
    let rowIndex = index + 2;
    let currentStatus = String(row[6]);              // Column G: Mentor status
    let attempts = Number(row[8]) || 0;               // Column I: Mentor invite attempts

    // Only invite interns explicitly marked "Eligible" (set automatically once they
    // hit MASTERED in AttandanceTracker.gs, or reset to "Eligible" by
    // refreshDeclinedMentorInvites() for a monthly re-invite). A blank status just
    // means they haven't finished the cohort yet, so it must NOT trigger an invite.
    if (currentStatus === "Eligible") {
      if (attempts >= MAX_MENTOR_INVITE_ATTEMPTS) {
        // Safety net - shouldn't normally hit this since refreshDeclinedMentorInvites
        // stops resetting people to "Eligible" once they're at the cap.
        detailsSheet.getRange(rowIndex, 7).setValue(`Declined - stopped after ${attempts} invites`);
        return;
      }
      sendMentorInvitationToRow(detailsSheet, rowIndex, attempts);
    }
  });
}

/**
 * Installable onEdit trigger (needs to be installable, not the simple onEdit(e)
 * function, since simple triggers can't send email). Watches Column L on Intern
 * Details: when someone manually checks the box for a row, it immediately sends
 * that person a mentor invite - going through sendMentorInvitationToRow so it's
 * tracked with the exact same attempt-count/status columns as an automated
 * invite, and therefore follows the same 3-attempt decline/reinvite rules.
 */
function onEditInternDetails(e) {
  try {
    let sheet = e.range.getSheet();
    if (sheet.getName() !== "Intern Details") return;
    if (e.range.getRow() === 1 || e.range.getColumn() !== MANUAL_MENTOR_CHECKBOX_COL) return;
    if (e.range.getValue() !== true) return; // only fire when checked, not unchecked

    let rowIndex = e.range.getRow();
    let currentStatus = String(sheet.getRange(rowIndex, 7).getValue() || "").trim(); // Column G

    if (currentStatus === "Mentor" || currentStatus === "Mentor Done") {
      // Already a mentor - nothing to do, reset the checkbox to avoid confusion
      sheet.getRange(rowIndex, MANUAL_MENTOR_CHECKBOX_COL).setValue(false);
      Logger.log(`Row ${rowIndex} is already a mentor (status: ${currentStatus}) - manual checkbox ignored.`);
      return;
    }

    let attempts = Number(sheet.getRange(rowIndex, 9).getValue()) || 0; // Column I

    if (attempts >= MAX_MENTOR_INVITE_ATTEMPTS) {
      sheet.getRange(rowIndex, 7).setValue(`Declined - stopped after ${attempts} invites`);
      sheet.getRange(rowIndex, MANUAL_MENTOR_CHECKBOX_COL).setValue(false);
      return;
    }

    sendMentorInvitationToRow(sheet, rowIndex, attempts);

    // Reset the checkbox after sending, so it's ready to use again next time.
    sheet.getRange(rowIndex, MANUAL_MENTOR_CHECKBOX_COL).setValue(false);
  } catch (err) {
    Logger.log("onEditInternDetails error: " + err);
  }
}

/** Run this once from the editor to install the manual-checkbox trigger. */
function createInternDetailsEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => { if (tr.getHandlerFunction() === "onEditInternDetails") ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger("onEditInternDetails").forSpreadsheet(ss).onEdit().create();
}

/**
 * Run this once a month (e.g. alongside your monthly attendance setup). Anyone who
 * declined ("No") gets bumped back to "Eligible" so sendMentorInvitationEmails()
 * re-invites them next run - unless they're already at the 3-attempt cap, in which
 * case we stop and write a clear final status instead of inviting them again.
 */
function refreshDeclinedMentorInvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  let detailsData = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 9).getValues();

  detailsData.forEach((row, index) => {
    let rowIndex = index + 2;
    let mentorStatus = String(row[6] || "").trim(); // Column G
    let attempts = Number(row[8]) || 0;              // Column I

    if (mentorStatus !== "No") return; // only touch people who actively declined

    if (attempts >= MAX_MENTOR_INVITE_ATTEMPTS) {
      detailsSheet.getRange(rowIndex, 7).setValue(`Declined - stopped after ${attempts} invites`);
    } else {
      detailsSheet.getRange(rowIndex, 7).setValue("Eligible"); // picked up by next sendMentorInvitationEmails run
    }
  });
}

function createMonthlyMentorReinviteTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => { if (tr.getHandlerFunction() === "refreshDeclinedMentorInvites") ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger("refreshDeclinedMentorInvites").timeBased().onMonthDay(1).atHour(6).create();
}

/**
 * Run this once to install the monthly invite-sending trigger, 1 hour after
 * refreshDeclinedMentorInvites (above) - so anyone bumped back to "Eligible"
 * that same run gets swept into the same batch as newly-mastered interns.
 */
function createMonthlySendMentorInvitesTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => { if (tr.getHandlerFunction() === "sendMentorInvitationEmails") ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger("sendMentorInvitationEmails").timeBased().onMonthDay(1).atHour(7).create();
}

function processMentorEmailReplies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  let detailsRange = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8);
  let detailsData = detailsRange.getValues();

  let threads = GmailApp.search('subject:"Invitation to become a Mentor" newer_than:7d');
  let responses = {};

  threads.forEach(thread => {
    let messages = thread.getMessages();
    let lastMessage = messages[messages.length - 1];
    
    let senderEmail = extractEmailAddress(lastMessage.getFrom()).toLowerCase().trim();
    let messageBody = lastMessage.getPlainBody();
    
    // Fallback for HTML-only emails
    if (!messageBody) {
      messageBody = lastMessage.getBody().replace(/<[^>]*>/g, ""); 
    }
    messageBody = messageBody.toLowerCase();

    let decision = "";
    let preference = "";

    // Check decision
    if (messageBody.includes("no") || messageBody.includes("cant") || messageBody.includes("can't") || messageBody.includes("not interested")) {
      decision = "No";
    } else if (messageBody.includes("yes") || messageBody.includes("sure") || messageBody.includes("interested") || messageBody.includes("love to")) {
      decision = "Accepted";
      
      // Check track preference
      if (messageBody.includes("weekday") || messageBody.includes("wd")) {
        preference = "WD";
      } else if (messageBody.includes("weekend") || messageBody.includes("we")) {
        preference = "WE";
      }
    }

    if (decision && senderEmail) {
      responses[senderEmail] = { decision: decision, preference: preference };
    }
  });

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let internTrack = row[4];                          // Column E: Intern WE/WD
    let rowIndex = index + 2;
    let currentStatus = String(row[6]);                // Column G: Mentor Status

    // Process if we received a response and they aren't permanently locked yet
    if (responses[email] && currentStatus !== "Mentor Done") {
      let resp = responses[email];

      if (resp.decision === "No") {
        detailsSheet.getRange(rowIndex, 7).setValue("No"); 
      } else if (resp.decision === "Accepted") {
        detailsSheet.getRange(rowIndex, 7).setValue("Mentor"); 
        
        let finalPreference = resp.preference ? resp.preference : internTrack;
        detailsSheet.getRange(rowIndex, 8).setValue(finalPreference); // Column H: Mentor WD/WE
        
        // Adds both new and old participants into the exact same shared event matching their track
        sendSharedMentorCalendarInvite(email, finalPreference);

        // Auto-grant sheet editor access and email them their live attendance tab link
        // (grantMentorTabAccess lives in AttandanceTracker.gs)
        grantMentorTabAccess(email, finalPreference);
      }
    }
  });
}

function extractEmailAddress(fromField) {
  let emailMatch = fromField.match(/<([^>]+)>-?/);
  return emailMatch ? emailMatch[1] : fromField;
}

function sendSharedMentorCalendarInvite(email, preference) {
  // Finds the existing 4-week recurring series for this track (or creates it if
  // none exists yet) and adds the mentor as a guest to the whole series, so they
  // get the exact same shared Meet link everyone else has.
  // (addStudentToAllFourCourseSessions lives in PreTaskTracker.gs)
  addStudentToAllFourCourseSessions(email, preference === "WE");
}

function createMentorEmailTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => { if (tr.getHandlerFunction() === "processMentorEmailReplies") ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger("processMentorEmailReplies").timeBased().everyDays(1).atHour(21).create();
}