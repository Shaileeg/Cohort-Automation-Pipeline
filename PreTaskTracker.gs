/**
 * Triggered when a new response is submitted to the Pre-task Form.
 * Matches the email against the Registration Form, marks status as COMPLETED,
 * auto-assigns calendar invites, and sends the confirmation email.
 */
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
    let studentEmail = String(row[2] || "").trim();    // Col C: Email Address

    // Skip if already processed, empty, or not an email
    if (isProcessedCheckbox === true || isProcessedCheckbox === "TRUE" || !studentEmail || !studentEmail.includes("@")) {
      continue;
    }

    let foundMatch = false;
    for (let j = 1; j < regData.length; j++) {
      let regRow = regData[j];
      let regEmail = String(regRow[2] || "").trim();    // Col C in Reg sheet
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

/**
 * Automatically checks if today is exactly 5 days before the upcoming cohort starts.
 * If it is, it scans the Registration Form and emails anyone whose status in Column L 
 * is NOT "COMPLETED".
 * 
 * SETUP INSTRUCTIONS: Set a daily time-driven trigger for this function.
 */
function sendPreTaskReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regSheet = ss.getSheetByName("Registration Form");
  
  if (!regSheet) {
    Logger.log("Could not find 'Registration Form' sheet!");
    return;
  }

  const now = new Date();
  
  // Check both tracks (Weekday and Weekend start dates)
  let weekdayFirstDay = getNextCohortFirstDay(false); // Thursday
  let weekendFirstDay = getNextCohortFirstDay(true);  // Sunday

  // Calculate target reminder dates (5 days before start date)
  let weekdayReminderDate = new Date(weekdayFirstDay.getTime() - (5 * 24 * 60 * 60 * 1000));
  let weekendReminderDate = new Date(weekendFirstDay.getTime() - (5 * 24 * 60 * 60 * 1000));

  let isWeekdayReminderDay = now.toDateString() === weekdayReminderDate.toDateString();
  let isWeekendReminderDay = now.toDateString() === weekendReminderDate.toDateString();

  // If today is NOT the 5-day mark for either track, stop here.
  if (!isWeekdayReminderDay && !isWeekendReminderDay) {
    Logger.log("Today is not the 5-day reminder threshold. No emails sent.");
    return;
  }

  const regData = regSheet.getDataRange().getValues();
  const preTaskSubmissionLink = "https://docs.google.com/forms/d/e/1FAIpQLScnXYgjSKsq0VKA_BsJWJpDK1HgGQ1IOo3KFuax31ohd_iOew/viewform?usp=send_form";
  
  let reminderCount = 0;

  for (let i = 1; i < regData.length; i++) {
    let regRow = regData[i];
    let regEmail = String(regRow[2] || "").trim(); // Column C: Email
    let regName = String(regRow[3] || "").trim();  // Column D: Name
    let chosenTiming = String(regRow[6] || "").toLowerCase(); // Column G: Timing choice
    let status = String(regRow[11] || "").trim();  // Column L: Status ("COMPLETED" or blank)

    let isWeekendUser = chosenTiming.includes("weekend") || chosenTiming.includes("sunday");

    // Only send if it matches the active reminder track day and they haven't completed it
    let matchesActiveTrack = (isWeekendUser && isWeekendReminderDay) || (!isWeekendUser && isWeekdayReminderDay);

    if (matchesActiveTrack && regEmail && regEmail.includes("@") && status.toUpperCase() !== "COMPLETED") {
      let subject = "Action Required: Complete Your Partnership Course Pre-Task (5 Days Left)";
      let body = `
        <p>Hello ${regName},</p>
        <p>Your cohort starts in 5 days, and we noticed you haven't completed your pre-task for the Partnership Course yet.</p>
        <p>Completing this is required before you can receive your session calendar invites and join the cohort.</p>
        <p>You can access and submit your pre-task here: <a href="${preTaskSubmissionLink}">Complete Pre-Task Now</a></p>
        <p>If you have already submitted it, please allow some time for verification.</p>
        <p>Best regards,</p>
        <p>The Partnership Team</p>
      `;

      GmailApp.sendEmail(regEmail, subject, "", { htmlBody: body });
      reminderCount++;
    }
  }

  Logger.log(`Successfully sent ${reminderCount} pre-task reminder emails for the 5-day mark.`);
}

/**
 * Creates a calendar event with a guaranteed Google Meet link, via the
 * Advanced Calendar Service - independent of whether the account has
 * "auto-add conferencing" turned on. Returns the event's ID so the caller
 * can fetch it back via CalendarApp.getEventById() to call addGuest() on it.
 *
 * REQUIRES: the "Calendar API" Advanced Service must be enabled for this
 * Apps Script project (Editor > Services > + > Google Calendar API), and the
 * underlying Google Calendar API enabled in the linked Cloud project.
 */
function createEventWithMeetLink(calendarId, title, description, startDate, endDate) {
  let event = {
    summary: title,
    description: description,
    start: { dateTime: startDate.toISOString(), timeZone: 'Etc/GMT' },
    end: { dateTime: endDate.toISOString(), timeZone: 'Etc/GMT' },
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: 'hangoutsMeet' }
      }
    }
  };

  let createdEvent = Calendar.Events.insert(event, calendarId, { conferenceDataVersion: 1 });
  return createdEvent.id;
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
  let calendarId = calendar.getId();
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
      // Create via the Advanced Calendar Service so a Meet link is guaranteed,
      // regardless of the account's own conferencing default.
      let createdEventId = createEventWithMeetLink(
        calendarId, eventTitle,
        `Shared group course meeting link for Week ${week + 1}.`,
        eventDate, endDate
      );
      groupEvent = calendar.getEventById(createdEventId);
    }

    groupEvent.addGuest(email);
  }
}

function addStudentToGroupTechCheck(email, isWeekend) {
  let calendar = CalendarApp.getDefaultCalendar();
  let calendarId = calendar.getId();
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
    // Create via the Advanced Calendar Service so a Meet link is guaranteed,
    // regardless of the account's own conferencing default.
    let createdEventId = createEventWithMeetLink(
      calendarId, eventTitle,
      'Shared group technical check meeting before the course begins.',
      techCheckStart, techCheckEnd
    );
    groupEvent = calendar.getEventById(createdEventId);
  }

  groupEvent.addGuest(email);
}

function sendPreTaskCompletedEmail(name, email, sessionDatesList) {
  const subject = "Pre-Task Received! Calendar Invites Confirmed - Partnership Course";
  const body = `
    <p>Hello ${name},</p>
    <p>We have successfully received and verified your pre-task submission!</p>
    <p>Congratulations on completing your pre-task! Thank you for putting in the time and effort to review the materials and submit your 20/20 activity. By doing this work yourself, you have already started the process of priming your brain for deep learning.</p>
    <p><strong>Your Shared Group Calendar Invites Have Sent:</strong></p>
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