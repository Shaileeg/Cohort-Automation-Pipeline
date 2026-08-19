const FORM_LINKS = {
  PRETASK: "[Link to Registration Form]" // Replace with your actual Pre-task form link
};

function onRegistrationFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Registration Form");
  if (!regSheet) return;

  const regData = regSheet.getDataRange().getValues();
  const cohortDates = getCalculatedCohortDates();

  for (let i = 1; i < regData.length; i++) {
    let row = regData[i];
    let rowIdx = i + 1;
    
    let isProcessedCheckbox = row[0];                   // Col A: Process Checkbox
    let email = String(row[2] || "").trim();           // Col C: Email Address
    let name = String(row[3] || "").trim();            // Col D: Full Name
    let coursesCompleted = String(row[5] || "").trim();// Col F: Courses Completed

    // Safety check: Skip if already checked or email is empty
    if (isProcessedCheckbox === true || isProcessedCheckbox === "TRUE" || !email || !email.includes("@")) {
      continue;
    }

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
  }
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