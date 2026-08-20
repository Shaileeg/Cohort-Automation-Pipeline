function sendMentorInvitationEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Intern Details");
  if (!detailsSheet || detailsSheet.getLastRow() <= 1) return;

  let detailsRange = detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 8);
  let detailsData = detailsRange.getValues();

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let name = row[1];                               // Column B: Name
    let rowIndex = index + 2;
    let currentStatus = String(row[6]);              // Column G: Mentor status

    // If they hit their milestone but haven't been emailed yet
    if (currentStatus === "Eligible" || currentStatus === "") {
      let subject = "Invitation to become a Mentor";
      let message = `Hi ${name},\n\n` +
                    `Congratulations on reaching your milestones! We would love to invite you to become a mentor for our upcoming sessions.\n\n` +
                    `Please reply to this email with "Yes" if you would like to mentor, or "No" if you cannot make it.\n` +
                    `If you say Yes, you can also choose your track preference by including "Weekday" (WD) or "Weekend" (WE) in your reply.\n\n` +
                    `Best regards,\nOrganizing Committee`;

      // Send the email
      MailApp.sendEmail(email, subject, message);

      // Update status to prevent spamming
      detailsSheet.getRange(rowIndex, 7).setValue("Invite Sent");
    }
  });
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
      }
    }
  });
}

function extractEmailAddress(fromField) {
  let emailMatch = fromField.match(/<([^>]+)>-?/);
  return emailMatch ? emailMatch[1] : fromField;
}

function sendSharedMentorCalendarInvite(email, preference) {
  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const futureTime = new Date(now.getTime() + (30 * 24 * 60 * 60 * 1000)); // 30 days ahead
  
  // Searches for your shared master session event
  const events = calendar.getEvents(now, futureTime, { search: "Partnership Course" });

  for (let i = 0; i < events.length; i++) {
    let event = events[i];
    let eventDateObj = event.getStartTime();
    let dayOfWeek = eventDateObj.getDay(); 
    // Automatically determines if the scheduled event is Weekend (Sat/Sun) or Weekday (Mon-Fri)
    let sessionType = (dayOfWeek === 0 || dayOfWeek === 6) ? "WE" : "WD";

    if (sessionType === preference) {
      // Adds them as a guest to the shared event, granting them the exact same Google Meet link
      event.addGuest(email);
      break; 
    }
  }
}

function createMentorEmailTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => { if (tr.getHandlerFunction() === "processMentorEmailReplies") ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger("processMentorEmailReplies").timeBased().everyDays(1).atHour(21).create();
}