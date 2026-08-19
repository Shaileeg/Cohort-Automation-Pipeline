function onPreTaskFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preTaskSheet = ss.getSheetByName("Pre-task Form"); 
  const regSheet = ss.getSheetByName("Registration Form");
  
  if (!preTaskSheet || !regSheet) return;

  const preTaskData = preTaskSheet.getDataRange().getValues();
  const regData = regSheet.getDataRange().getValues();
  const cohortDates = getCalculatedCohortDates();

  let processedCount = 0;

  for (let i = 1; i < preTaskData.length; i++) {
    let row = preTaskData[i];
    let rowIdx = i + 1;
    
    let isProcessedCheckbox = row[0];             // Col A: Checkbox
    let studentEmail = String(row[2] || "").trim();     // Col C: Email Address

    // Skip if already processed, empty, or not an email
    if (isProcessedCheckbox === true || isProcessedCheckbox === "TRUE" || !studentEmail || !studentEmail.includes("@")) {
      continue;
    }

    let foundMatch = false;
    for (let j = 1; j < regData.length; j++) {
      let regRow = regData[j];
      let regEmail = String(regRow[2] || "").trim();     // Col C in Reg sheet
      let regName = String(regRow[3] || "").trim();      // Col D
      let chosenTiming = String(regRow[6] || "").toLowerCase(); // Col G (Index 6) - Weekday vs Weekend choice
      
      if (regEmail.toLowerCase() === studentEmail.toLowerCase()) {
        foundMatch = true;
        
        // Update Registration sheet Column L to COMPLETED
        regSheet.getRange(j + 1, 12).setValue("COMPLETED"); 

        // Determine track based on Registration Form Column G
        let isWeekend = chosenTiming.includes("weekend") || chosenTiming.includes("sunday");

        // 1. Add student to the correct (Weekday or Weekend) 4 weekly Course events
        addStudentToAllFourCourseSessions(regEmail, isWeekend);
        
        // 2. Add student to the correct Tech Check event
        addStudentToGroupTechCheck(regEmail, isWeekend);

        // 3. Send confirmation email with the correct dates (Sunday list if weekend, Thursday list if weekday)
        sendPreTaskCompletedEmail(regName, regEmail, isWeekend ? cohortDates.sundayList : cohortDates.thursdayList);
        break; 
      }
    }

    if (foundMatch) {
      preTaskSheet.getRange(rowIdx, 1).setValue(true);
      SpreadsheetApp.flush(); 
      processedCount++;
    }
  }
  
  Logger.log("Successfully processed " + processedCount + " new submission(s).");
}

function getNextCohortFirstDay(isWeekend) {
  const now = new Date();
  let nextYear = now.getFullYear();
  let nextMonth = now.getMonth() + 1;
  if (nextMonth > 11) {
    nextMonth = 0;
    nextYear++;
  }
  
  let targetDate = new Date(Date.UTC(nextYear, nextMonth, 1));
  let targetDayOfWeek = isWeekend ? 0 : 4; // 0 = Sunday (Weekend), 4 = Thursday (Weekday)
  
  let day = targetDate.getUTCDay();
  let diff = (targetDayOfWeek - day + 7) % 7;
  targetDate.setUTCDate(targetDate.getUTCDate() + diff);
  
  return targetDate; 
}

function addStudentToAllFourCourseSessions(email, isWeekend) {
  let calendar = CalendarApp.getDefaultCalendar();
  let firstSessionDate = getNextCohortFirstDay(isWeekend);
  let cohortType = isWeekend ? "Weekend" : "Weekday";

  for (let week = 0; week < 4; week++) {
    let eventDate = new Date(firstSessionDate);
    eventDate.setUTCDate(eventDate.getUTCDate() + (week * 7)); 
    
    let endDate = new Date(eventDate);
    
    if (isWeekend) {
      eventDate.setUTCHours(14, 0, 0, 0); 
      endDate.setUTCHours(16, 0, 0, 0);   
    } else {
      eventDate.setUTCHours(11, 0, 0, 0); 
      endDate.setUTCHours(13, 0, 0, 0);   
    }

    let eventTitle = `Partnership Course ${cohortType}`;
    
    let events = calendar.getEvents(
      new Date(eventDate.getTime() - 24 * 60 * 60 * 1000), 
      new Date(eventDate.getTime() + 24 * 60 * 60 * 1000), 
      { search: eventTitle }
    );

    let groupEvent;
    if (events.length > 0) {
      groupEvent = events[0];
    } else {
      groupEvent = calendar.createEvent(eventTitle, eventDate, endDate, {
        description: `Shared group course meeting link for Week ${week + 1}.`,
        timeZone: 'Etc/GMT'
      });
    }

    groupEvent.addGuest(email);
  }
}

function addStudentToGroupTechCheck(email, isWeekend) {
  let calendar = CalendarApp.getDefaultCalendar();
  let firstSessionDate = getNextCohortFirstDay(isWeekend);
  let cohortType = isWeekend ? "Weekend" : "Weekday";
  
  let techCheckStart = new Date(firstSessionDate);
  let techCheckEnd = new Date(techCheckStart);
  
  if (isWeekend) {
    techCheckStart.setUTCHours(13, 0, 0, 0); 
    techCheckEnd.setUTCHours(14, 0, 0, 0);   
  } else {
    techCheckStart.setUTCHours(10, 0, 0, 0); 
    techCheckEnd.setUTCHours(11, 0, 0, 0);   
  }

  let eventTitle = `Tech Check ${cohortType}`;
  
  let events = calendar.getEvents(
    new Date(techCheckStart.getTime() - 24 * 60 * 60 * 1000), 
    new Date(techCheckStart.getTime() + 24 * 60 * 60 * 1000), 
    { search: eventTitle }
  );

  let groupEvent;
  if (events.length > 0) {
    groupEvent = events[0];
  } else {
    groupEvent = calendar.createEvent(eventTitle, techCheckStart, techCheckEnd, {
      description: 'Shared group technical check meeting before the course begins.',
      timeZone: 'Etc/GMT'
    });
  }

  groupEvent.addGuest(email);
}

function sendPreTaskCompletedEmail(name, email, sessionDatesList) {
  const subject = "Pre-Task Received! Calendar Invites Confirmed - Partnership Course";
  const body = `
    <p>Hello ${name},</p>
    <p>We have successfully received and verified your pre-task submission!</p>
    <p>Congratulations on completing your pre-task! Thank you for putting in the time and effort to review the materials and submit your 20/20 activity. By doing this work yourself, you have already started the process of priming your brain for deep learning.</p>
    <p><strong>Your Shared Group Calendar Invites Have Been Sent:</strong></p>
    <ul>
      <li><strong>Tech Check:</strong> 1 hour before your first session. It is <strong>Mandatory</strong> to join tech check</li>
      <li><strong>Partnership Course Sessions:</strong> Scheduled across your cohort dates (${sessionDatesList})</li>
    </ul>
    <p><em>Check your Google Calendar to access all meeting links.</em></p>
    <p>We are excited to have such a committed group ready to dive into the practical application of partnership and networking.</p>
    <p>See you in class!</p>
    <p>Best regards,</p>
    <p>The Partnership Team</p>
  `;
  GmailApp.sendEmail(email, subject, "", { htmlBody: body });
}