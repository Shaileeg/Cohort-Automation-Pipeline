function getCalculatedCohortDates() {
  const now = new Date();
  
  let nextMonthYear = now.getFullYear();
  let nextMonthIndex = now.getMonth() + 1;
  if (nextMonthIndex > 11) {
    nextMonthIndex = 0;
    nextMonthYear++;
  }
  
  let firstDayOfNextMonth = new Date(nextMonthYear, nextMonthIndex, 1);
  const dateOptions = { month: 'long', day: 'numeric' };

  let thursdays = [];
  let currentThursday = new Date(firstDayOfNextMonth);
  let dayOfWeek = currentThursday.getDay(); 
  let daysUntilThursday = (4 - dayOfWeek + 7) % 7;
  currentThursday.setDate(currentThursday.getDate() + daysUntilThursday);

  while (currentThursday.getMonth() === nextMonthIndex && thursdays.length < 4) {
    thursdays.push(currentThursday.toLocaleDateString('en-US', dateOptions));
    currentThursday = new Date(currentThursday);
    currentThursday.setDate(currentThursday.getDate() + 7);
  }
  let formattedThursdayDates = formatNiceDateList(thursdays);

  let sundays = [];
  let currentSunday = new Date(firstDayOfNextMonth);
  dayOfWeek = currentSunday.getDay(); 
  let daysUntilSunday = (0 - dayOfWeek + 7) % 7;
  currentSunday.setDate(currentSunday.getDate() + daysUntilSunday);

  while (currentSunday.getMonth() === nextMonthIndex && sundays.length < 4) {
    sundays.push(currentSunday.toLocaleDateString('en-US', dateOptions));
    currentSunday = new Date(currentSunday);
    currentSunday.setDate(currentSunday.getDate() + 7);
  }
  let formattedSundayDates = formatNiceDateList(sundays);

  let firstThursdayDateObj = new Date(firstDayOfNextMonth);
  let dWeek = firstThursdayDateObj.getDay();
  let dUntilThu = (4 - dWeek + 7) % 7;
  firstThursdayDateObj.setDate(firstThursdayDateObj.getDate() + dUntilThu);
  
  let preTaskDeadline = new Date(firstThursdayDateObj);
  preTaskDeadline.setDate(preTaskDeadline.getDate() - 3);

  let firstSundayDateObj = new Date(firstDayOfNextMonth);
  let sWeek = firstSundayDateObj.getDay();
  let sUntilSun = (0 - sWeek + 7) % 7;
  firstSundayDateObj.setDate(firstSundayDateObj.getDate() + sUntilSun);

  const fullDateOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };

  return {
    thursdayList: formattedThursdayDates, 
    sundayList: formattedSundayDates,     
    firstThursdayFull: firstThursdayDateObj.toLocaleDateString('en-US', fullDateOptions),
    firstSundayFull: firstSundayDateObj.toLocaleDateString('en-US', fullDateOptions),
    preTaskDeadlineDate: preTaskDeadline.toLocaleDateString('en-US', fullDateOptions)
  };
}

function formatNiceDateList(dateStrings) {
  if (dateStrings.length === 0) return "";
  let parts = dateStrings[0].split(" ");
  let monthName = parts[0]; 
  let days = dateStrings.map(d => d.split(" ")[1]); 
  
  if (days.length === 1) return `${monthName} ${days[0]}`;
  if (days.length === 2) return `${monthName} ${days[0]} & ${days[1]}`;
  
  let allButLast = days.slice(0, days.length - 1).join(", ");
  let last = days[days.length - 1];
  
  return `${monthName} ${allButLast} & ${last}`;
}