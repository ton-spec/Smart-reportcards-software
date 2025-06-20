function startBatchTrigger() {
  ScriptApp.newTrigger("generateReportCardsInBatches")
    .timeBased()
    .everyMinutes(1)
    .create();
}

function generateReportCardsInBatches() {
  const props = PropertiesService.getScriptProperties();
  const startIndex = Number(props.getProperty("lastIndex")) || 1;
  const batchSize = 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const templateFile = DriveApp.getFileById('1RBfCot_g6a5RwVEdVYCiQaa25vceu9vvUhsxfaJowME');
  const term = sheet.getRange("W1").getValue();
  const year = sheet.getRange("X1").getValue();
  const grade = sheet.getRange("Y1").getValue();

  let folderId = props.getProperty("folderId");
  let folder;
  if (!folderId) {
    folder = DriveApp.createFolder('Student Report Cards ' + new Date().toISOString());
    props.setProperty("folderId", folder.getId());
  } else {
    folder = DriveApp.getFolderById(folderId);
  }

  const logoUrl = "https://drive.google.com/uc?export=view&id=1afqTLwfAhb3oGVC8J4v-tavj_J8sfdwE";
  const logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();

  for (let i = startIndex; i < Math.min(startIndex + batchSize, data.length); i++) {
    try {
      const studentRow = data[i];
      generateOneReport(studentRow, i, templateFile, folder, term, year, grade, logoBlob);
    } catch (err) {
      Logger.log(`Error on row ${i + 1}: ${err}`);
    }
  }

  const nextIndex = startIndex + batchSize;
  if (nextIndex < data.length) {
    props.setProperty("lastIndex", nextIndex);
  } else {
    props.deleteAllProperties();
    generateRubricSummary(data, folder);
    stopAllTriggers();
    Logger.log("🎉 Finished processing all students.");
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), "✅ Report Card Generation Complete", 
    "All student report cards have been generated and saved successfully.");

    SpreadsheetApp.getUi().alert("✅ All report cards generated.");
  }
}

function stopAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "generateReportCardsInBatches") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
