const FORM_LINKS = {
  PRETASK: "https://docs.google.com/forms/d/e/1FAIpQLScnXYgjSKsq0VKA_BsJWJpDK1HgGQ1IOo3KFuax31ohd_iOew/viewform?usp=send_form"
};

/**
 * Fires on a real form submission. Processes ONLY the row that just came in
 * (via e.range.getRow()) - NOT the whole sheet. This matters: if this scanned
 * every row on every submission, any backlog of unprocessed rows (e.g. from
 * before this trigger was turned on, or a past run that failed partway) would
 * all get emailed together the next time anyone submitted the form, instead of
 * just the new person.
 */
function onRegistrationFormSubmit(e) {
  if (!e || !e.range) {
    Logger.log("onRegistrationFormSubmit: no trigger event/range - not processing anything. Use processRegistrationBacklog() to catch up manually if needed.");
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Registration Form");
  if (!regSheet) return;

  processRegistrationRow(regSheet, e.range.getRow());
}

/**
 * Manual catch-up: scans the WHOLE Registration Form sheet and processes any
 * row that isn't marked done yet. Run this by hand (from the editor) if you
 * suspect some rows were missed - do not wire this to the form-submit trigger,
 * or you're back to the "processes the whole backlog on every submission" bug.
 */
function processRegistrationBacklog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Registration Form");
  if (!regSheet || regSheet.getLastRow() <= 1) return;

  let processedCount = 0;
  for (let rowIdx = 2; rowIdx <= regSheet.getLastRow(); rowIdx++) {
    if (processRegistrationRow(regSheet, rowIdx)) processedCount++;
  }
  Logger.log(`processRegistrationBacklog: processed ${processedCount} previously-unhandled row(s).`);
}

/**
 * Processes a single Registration Form row. Returns true if it actually did
 * something (i.e. wasn't already checked/empty/invalid), false otherwise.
 */
function processRegistrationRow(regSheet, rowIdx) {
  let row = regSheet.getRange(rowIdx, 1, 1, regSheet.getLastColumn()).getValues()[0];

  let isProcessedCheckbox = row[0];                   // Col A: Process Checkbox
  let email = String(row[2] || "").trim();           // Col C: Email Address
  let name = String(row[3] || "").trim();            // Col D: Full Name
  let coursesCompleted = String(row[5] || "").trim();// Col F: Courses Completed

  // Safety check: Skip if already checked or email is empty
  if (isProcessedCheckbox === true || isProcessedCheckbox === "TRUE" || !email || !email.includes("@")) {
    return false;
  }

  const cohortDates = getCalculatedCohortDates();
  const coursesLower = coursesCompleted.toLowerCase();
  const hasDesignThinking = coursesLower.includes("design thinking");
  const hasNetworkingLive = coursesLower.includes("networking live");
  const cleanCourses = coursesLower.replace("networking live", "");
  const hasNetworking = cleanCourses.includes("networking");

  if (hasDesignThinking && hasNetworkingLive && hasNetworking) {
    regSheet.getRange(rowIdx, 10).setValue("QUALIFIED");   // Col J
    regSheet.getRange(rowIdx, 12).setValue("PENDING");    // Col L (Pre-task status)
    
    // Send Pre-Task Email
    sendPreTaskEmail(name, email, cohortDates.preTaskDeadlineDate);

  } else {
    regSheet.getRange(rowIdx, 10).setValue("PREREQ_FAILED");
    sendPrereqFailedEmail(name, email);
  }

  // Mark row as processed so it never runs twice
  regSheet.getRange(rowIdx, 1).setValue(true);
  return true;
}

function sendPreTaskEmail(name, email, preTaskDeadline) {
  const subject = "Next Steps: Partnership Course Pre-Task Submission";
  const body = `
    <p>Hello ${name},</p>
    <p>Thank you for registering for the partnership course. We are excited to have you started on this course.</p>
    <p>This course is a vital step in mastering essential soft skills like active listening and empathy while you develop your professional network before moving into coaching.</p>
    <p>To ensure everyone is prepared and to help us design the best curriculum for your cohort, there is a mandatory pre-class task that must be completed to help us know your current level of understanding, what you are curious about, and ensures a high level of commitment from all participants.</p> 
    <p><strong>Important:</strong> Do not use AI tools to generate your answers. We need to know your thoughts and curiosity. Using AI at this point stops you from developing the skills needed to network effectively. 
</p>
    <p><strong>Your Next Step:</strong></p>
    <p>Please fill out and submit your pre-task by <strong>${preTaskDeadline}</strong> here: <br>
    <a href="${FORM_LINKS.PRETASK}">Pre-Task Form Link</a></p>
    <p><em>Note: Once your pre-task is submitted, you will receive your calendar invites for the course and the group tech check.</em></p>
    <p>Best regards,</p>
    <p>The Partnership Team</p>
  `;
  GmailApp.sendEmail(email, subject, "", { htmlBody: body });
}

function sendPrereqFailedEmail(name, email) {
  const subject = "Update Regarding Your Partnership Course Application";
  const body = `<p>Hi ${name}, you are missing required prerequisites. Please complete them and re-apply.</p>`;
  GmailApp.sendEmail(email, subject, "", { htmlBody: body });
}