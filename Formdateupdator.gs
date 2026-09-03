/**
 * -------------------------------------------------------------------
 * MONTHLY FORM DATE UPDATER
 * -------------------------------------------------------------------
 * The Registration Form's Weekday/Weekend timing question has the actual
 * session dates hardcoded into the choice text (e.g. "Weekday- Thursday
 * (September 3rd, 10th, 17th & 24th) 11am - 1pm GMT +0"). Nothing else in
 * this project ever touched the form itself - only form RESPONSES - so this
 * text had to be edited by hand every month or it would silently show stale
 * dates to new registrants (the actual calendar events would still be
 * correctly dated, since those are computed fresh, but the form itself
 * would be lying to whoever's filling it out).
 *
 * This function fixes that: it recomputes next month's Thursdays/Sundays
 * with the same ordinal-suffix style your form already uses, and pushes
 * the updated text directly into the form's choices via FormApp.
 *
 * NOTE: the first time this runs, Apps Script will prompt for additional
 * authorization (the Forms scope) since nothing in this project has used
 * FormApp before.
 */

const REGISTRATION_FORM_ID = "1d1TCei6mrMnxWqO4S7iGNBRH2IzZc_Jn3TLqc0azAaw";

function getOrdinalSuffix(day) {
  if (day % 10 === 1 && day !== 11) return "st";
  if (day % 10 === 2 && day !== 12) return "nd";
  if (day % 10 === 3 && day !== 13) return "rd";
  return "th";
}

/**
 * Returns the month name and the day-of-month numbers (up to 4) for the
 * given weekday (0=Sunday, 4=Thursday) in NEXT month - same "next cohort"
 * timing convention used elsewhere (DateCalculator.gs, PreTaskTracker.gs).
 */
function getNextCohortDayNumbers(targetDayOfWeek) {
  const now = new Date();
  let nextMonthYear = now.getFullYear();
  let nextMonthIndex = now.getMonth() + 1;
  if (nextMonthIndex > 11) {
    nextMonthIndex = 0;
    nextMonthYear++;
  }
  let firstDayOfNextMonth = new Date(nextMonthYear, nextMonthIndex, 1);

  let days = [];
  let current = new Date(firstDayOfNextMonth);
  let dayOfWeek = current.getDay();
  let daysUntilTarget = (targetDayOfWeek - dayOfWeek + 7) % 7;
  current.setDate(current.getDate() + daysUntilTarget);

  while (current.getMonth() === nextMonthIndex && days.length < 4) {
    days.push(current.getDate());
    current = new Date(current);
    current.setDate(current.getDate() + 7);
  }

  return {
    monthName: firstDayOfNextMonth.toLocaleDateString('en-US', { month: 'long' }),
    days: days
  };
}

/** Builds "September 3rd, 10th, 17th & 24th" style text from day numbers. */
function buildOrdinalDateList(monthName, days) {
  let withSuffix = days.map(d => `${d}${getOrdinalSuffix(d)}`);
  if (withSuffix.length === 0) return monthName;
  if (withSuffix.length === 1) return `${monthName} ${withSuffix[0]}`;
  if (withSuffix.length === 2) return `${monthName} ${withSuffix[0]} & ${withSuffix[1]}`;
  let allButLast = withSuffix.slice(0, -1).join(", ");
  let last = withSuffix[withSuffix.length - 1];
  return `${monthName} ${allButLast} & ${last}`;
}

/**
 * Run this once from the editor to test, or leave it on the monthly trigger
 * (see createFormDateUpdateTrigger below). Finds the Weekday/Weekend timing
 * question by looking for choices starting with "Weekday-" and "Weekend-"
 * (rather than matching on question title, which is more fragile), and
 * replaces just those two choice strings with the updated dates.
 */
function updateRegistrationFormTimingChoices() {
  const form = FormApp.openById(REGISTRATION_FORM_ID);

  let thursdayInfo = getNextCohortDayNumbers(4); // Thursday
  let sundayInfo = getNextCohortDayNumbers(0);   // Sunday

  let thursdayList = buildOrdinalDateList(thursdayInfo.monthName, thursdayInfo.days);
  let sundayList = buildOrdinalDateList(sundayInfo.monthName, sundayInfo.days);

  let weekdayChoiceText = `Weekday- Thursday (${thursdayList}) 11am - 1pm GMT +0`;
  let weekendChoiceText = `Weekend- Sunday (${sundayList}) 2pm - 4pm GMT +0`;

  let items = form.getItems();
  let targetItem = null;
  let targetType = null;

  items.forEach(item => {
    let type = item.getType();
    let choices = null;

    if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
      choices = item.asMultipleChoiceItem().getChoices();
    } else if (type === FormApp.ItemType.CHECKBOX) {
      choices = item.asCheckboxItem().getChoices();
    } else {
      return;
    }

    let hasWeekday = choices.some(c => c.getValue().trim().indexOf("Weekday-") === 0);
    let hasWeekend = choices.some(c => c.getValue().trim().indexOf("Weekend-") === 0);
    if (hasWeekday && hasWeekend) {
      targetItem = item;
      targetType = type;
    }
  });

  if (!targetItem) {
    Logger.log("updateRegistrationFormTimingChoices: could not find a Weekday/Weekend timing question on the form. Nothing was changed.");
    return;
  }

  // Safety check: only overwrite if there are exactly 2 choices, so we never
  // silently wipe out extra options someone may have added to this question.
  let currentChoiceCount = (targetType === FormApp.ItemType.MULTIPLE_CHOICE)
    ? targetItem.asMultipleChoiceItem().getChoices().length
    : targetItem.asCheckboxItem().getChoices().length;

  if (currentChoiceCount !== 2) {
    Logger.log(`updateRegistrationFormTimingChoices: expected exactly 2 choices on the timing question but found ${currentChoiceCount}. Nothing was changed - update manually or adjust this function.`);
    return;
  }

  if (targetType === FormApp.ItemType.MULTIPLE_CHOICE) {
    targetItem.asMultipleChoiceItem().setChoiceValues([weekdayChoiceText, weekendChoiceText]);
  } else {
    targetItem.asCheckboxItem().setChoiceValues([weekdayChoiceText, weekendChoiceText]);
  }

  Logger.log(`Updated form timing choices:\n${weekdayChoiceText}\n${weekendChoiceText}`);
}

/**
 * Run this once from the editor to install the monthly trigger. Runs on the
 * 1st of each month, early enough that the new choices are live before that
 * month's registrations start coming in.
 */
function createFormDateUpdateTrigger() {
  let triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    if (tr.getHandlerFunction() === "updateRegistrationFormTimingChoices") {
      ScriptApp.deleteTrigger(tr);
    }
  });
  ScriptApp.newTrigger("updateRegistrationFormTimingChoices").timeBased().onMonthDay(1).atHour(1).create();
}