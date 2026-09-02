/**
 * Fires on a real Pre-task Form submission. Processes ONLY the row that just
 * came in (via e.range.getRow()) - NOT the whole sheet. Same reasoning as
 * onRegistrationFormSubmit in RegistrationAutomation.gs: scanning every row on
 * every submission means any backlog of unprocessed rows gets emailed together
 * the next time anyone submits, not just the new person.
 */
function onPreTaskFormSubmit(e) {
  if (!e || !e.range) {
    Logger.log("onPreTaskFormSubmit: no trigger event/range - not processing anything. Use processPreTaskBacklog() to catch up manually if needed.");
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preTaskSheet = ss.getSheetByName("Pre-task Form");
  const regSheet = ss.getSheetByName("Registration Form");
  if (!preTaskSheet || !regSheet) return;

  let didProcess = processPreTaskRow(preTaskSheet, regSheet, e.range.getRow());
  Logger.log(didProcess ? "Processed 1 new pre-task submission." : "Row was already processed, empty, or had no matching registration.");
}

/**
 * Manual catch-up: scans the WHOLE Pre-task Form sheet and processes any row
 * that isn't marked done yet. Run this by hand if you suspect some rows were
 * missed - do not wire this to the form-submit trigger.
 */
function processPreTaskBacklog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preTaskSheet = ss.getSheetByName("Pre-task Form");
  const regSheet = ss.getSheetByName("Registration Form");
  if (!preTaskSheet || !regSheet || preTaskSheet.getLastRow() <= 1) return;

  let processedCount = 0;
  for (let rowIdx = 2; rowIdx <= preTaskSheet.getLastRow(); rowIdx++) {
    if (processPreTaskRow(preTaskSheet, regSheet, rowIdx)) processedCount++;
  }
  Logger.log(`processPreTaskBacklog: processed ${processedCount} previously-unhandled row(s).`);
}

/**
 * Processes a single Pre-task Form row: matches it against Registration Form,
 * marks COMPLETED, adds calendar invites, and sends the confirmation email.
 * Returns true if it actually did something, false if skipped (already
 * processed, empty, invalid, or no matching registration found).
 */
function processPreTaskRow(preTaskSheet, regSheet, rowIdx) {
  let row = preTaskSheet.getRange(rowIdx, 1, 1, preTaskSheet.getLastColumn()).getValues()[0];

  let isProcessedCheckbox = row[0];             // Col A: Checkbox
  let studentEmail = String(row[2] || "").trim();    // Col C: Email Address

  // Skip if already processed, empty, or not an email
  if (isProcessedCheckbox === true || isProcessedCheckbox === "TRUE" || !studentEmail || !studentEmail.includes("@")) {
    return false;
  }

  const regData = regSheet.getDataRange().getValues();
  const cohortDates = getCalculatedCohortDates();

  for (let j = 1; j < regData.length; j++) {
    let regRow = regData[j];
    let regEmail = String(regRow[2] || "").trim();    // Col C in Reg sheet
    let regName = String(regRow[3] || "").trim();      // Col D
    let chosenTiming = String(regRow[6] || "").toLowerCase(); // Col G (Index 6) - Weekday vs Weekend choice

    if (regEmail.toLowerCase() === studentEmail.toLowerCase()) {
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

      preTaskSheet.getRange(rowIdx, 1).setValue(true);
      SpreadsheetApp.flush();
      return true;
    }
  }

  return false; // no matching registration found - leave unchecked so it can be caught later
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

/**
 * Creates a WEEKLY RECURRING event (occurrenceCount occurrences) with ONE Meet
 * link shared across every occurrence - unlike createEventWithMeetLink, which
 * makes a standalone event with its own separate link each time it's called.
 * Returns the recurring series' master event ID.
 */
function createRecurringEventWithMeetLink(calendarId, title, description, startDate, endDate, occurrenceCount) {
  let event = {
    summary: title,
    description: description,
    start: { dateTime: startDate.toISOString(), timeZone: 'Etc/GMT' },
    end: { dateTime: endDate.toISOString(), timeZone: 'Etc/GMT' },
    recurrence: [`RRULE:FREQ=WEEKLY;COUNT=${occurrenceCount}`],
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

/**
 * Finds an existing recurring series (or standalone event) matching this exact
 * title within the given window, WITHOUT expanding it into individual occurrence
 * instances (singleEvents: false) - so we get the series' own master ID back,
 * suitable for adding a guest to the whole series at once.
 */
function findExistingSeriesId(calendarId, title, searchStart, searchEnd) {
  let response = Calendar.Events.list(calendarId, {
    timeMin: searchStart.toISOString(),
    timeMax: searchEnd.toISOString(),
    q: title,
    singleEvents: false
  });
  let items = response.items || [];
  let match = items.find(e => e.summary && e.summary.trim() === title);
  return match ? match.id : null;
}

/**
 * Adds a guest to an entire recurring series (or a standalone event) at once,
 * so they get the same shared Meet link across every occurrence. Safe to call
 * repeatedly - won't duplicate the guest if they're already on it.
 */
function addGuestToSeries(calendarId, seriesEventId, email) {
  let event = Calendar.Events.get(calendarId, seriesEventId);
  let attendees = event.attendees || [];
  let alreadyGuest = attendees.some(a => a.email && a.email.toLowerCase() === email.toLowerCase());
  if (!alreadyGuest) {
    attendees.push({ email: email });
    Calendar.Events.patch({ attendees: attendees }, calendarId, seriesEventId);
  }
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
  let eventTitle = `Partnership Course ${cohortType}`;

  let startDate = new Date(firstSessionDate);
  let endDate = new Date(firstSessionDate);

  if (isWeekend) {
    startDate.setUTCHours(14, 0, 0, 0);
    endDate.setUTCHours(16, 0, 0, 0);
  } else {
    startDate.setUTCHours(11, 0, 0, 0);
    endDate.setUTCHours(13, 0, 0, 0);
  }

  // Search a window covering the whole 4-week cohort for an already-existing series
  let searchStart = new Date(startDate.getTime() - 24 * 60 * 60 * 1000);
  let searchEnd = new Date(startDate.getTime() + 35 * 24 * 60 * 60 * 1000);

  let seriesId = findExistingSeriesId(calendarId, eventTitle, searchStart, searchEnd);

  if (!seriesId) {
    // Not found - create the whole 4-week series at once, one shared Meet link
    // for all 4 sessions instead of 4 separate links.
    seriesId = createRecurringEventWithMeetLink(
      calendarId, eventTitle,
      'Shared group course meeting link for all 4 weekly sessions.',
      startDate, endDate, 4
    );
  }

  addGuestToSeries(calendarId, seriesId, email);
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