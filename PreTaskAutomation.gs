function syncPreTaskToInternDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let regSheet = ss.getSheetByName("Registration Form");
  let preTaskSheet = ss.getSheetByName("Pre-task Form");
  let detailsSheet = ss.getSheetByName("Intern Details");
  
  if (!regSheet || !preTaskSheet || !detailsSheet) return;
  if (preTaskSheet.getLastRow() <= 1 || regSheet.getLastRow() <= 1) return;

  let regData = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, regSheet.getLastColumn()).getValues();
  let registrationsByEmail = {};
  
  regData.forEach(row => {
    let email = String(row[1]).toLowerCase().trim(); // Column B: Email in Registration Form
    if (email) {
      registrationsByEmail[email] = {
        name: row[2],          // Column C: Name
        discord: row[3] || "", // Column D: Discord
        preference: String(row[4] || "").toUpperCase() // Column E: WE/WD
      };
    }
  });

  let preTaskData = preTaskSheet.getRange(2, 1, preTaskSheet.getLastRow() - 1, preTaskSheet.getLastColumn()).getValues();
  
  let detailsLastRow = detailsSheet.getLastRow();
  let detailsData = detailsLastRow > 1 ? detailsSheet.getRange(2, 1, detailsLastRow - 1, 8).getValues() : [];
  let existingEmails = detailsData.map(row => String(row[2]).toLowerCase().trim());

  preTaskData.forEach(taskRow => {
    let taskEmail = String(taskRow[1]).toLowerCase().trim(); // Column B: Email in Pre-task Form

    if (taskEmail && registrationsByEmail[taskEmail]) {
      if (!existingEmails.includes(taskEmail)) {
        let regInfo = registrationsByEmail[taskEmail];
        let nowMonth = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][new Date().getMonth()];
        
        detailsSheet.appendRow([
          nowMonth,             // Column A: Cohort
          regInfo.name,         // Column B: Intern Names
          taskEmail,            // Column C: Email
          regInfo.discord,      // Column D: Discord
          regInfo.preference,   // Column E: Intern WE/WD
          "Send",               // Column F: Calendar Invite
          "",                   // Column G: Mentor Status
          ""                    // Column H: Mentor WD/WE
        ]);
        
        existingEmails.push(taskEmail);
      }
    }
  });
}