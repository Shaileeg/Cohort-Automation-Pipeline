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
    let messageBody = lastMessage.getPlainBody().toLowerCase();

    let decision = "";
    let preference = "";

    if (messageBody.includes("no") || messageBody.includes("cant") || messageBody.includes("can't") || messageBody.includes("not interested")) {
      decision = "No";
    } else if (messageBody.includes("yes") || messageBody.includes("sure") || messageBody.includes("interested") || messageBody.includes("love to")) {
      decision = "Accepted";
      
      if (messageBody.includes("weekday") || messageBody.includes("wd")) {
        preference = "WD";
      } else if (messageBody.includes("weekend") || messageBody.includes("we")) {
        preference = "WE";
      }
    }

    if (decision) {
      responses[senderEmail] = { decision: decision, preference: preference };
    }
  });

  detailsData.forEach((row, index) => {
    let email = String(row[2]).toLowerCase().trim(); // Column C: Email
    let internTrack = row[4];                        // Column E: Intern WE/WD
    let rowIndex = index + 2;
    let currentStatus = String(row[6]);              // Column G: Mentor

    if (responses[email] && currentStatus !== "Mentor Done") {
      let resp = responses[email];

      if (resp.decision === "No") {
        detailsSheet.getRange(rowIndex, 7).setValue("No"); 
      } else if (resp.decision === "Accepted") {
        detailsSheet.getRange(rowIndex, 7).setValue("Mento..."); 
        
        if (resp.preference) {
          detailsSheet.getRange(rowIndex, 8).setValue(resp.preference); // Column H: Mentor WD/WE
          sendSharedMentorCalendarInvite(email, resp.preference);
        } else {
          detailsSheet.getRange(rowIndex, 8).setValue(internTrack);
          sendSharedMentorCalendarInvite(email, internTrack);
        }
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
  const futureTime = new Date(now.getTime() + (30 * 24 * 60 * 60 * 1000)); 
  const events = calendar.getEvents(now, futureTime, { search: "Partnership Course" });

  for (let i = 0; i < events.length; i++) {
    let event = events[i];
    let eventDateObj = event.getStartTime();
    let dayOfWeek = eventDateObj.getDay(); 
    let sessionType = (dayOfWeek === 0 || dayOfWeek === 6) ? "WE" : "WD";

    if (sessionType === preference) {
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