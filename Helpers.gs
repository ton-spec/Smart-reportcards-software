function getRubric(score) {
  score = Number(score);
  if (score >= 90) return "E.E ~AL8";
  if (score >= 75) return "E.E ~AL7";
  if (score >= 60) return "M.E ~AL6";
  if (score >= 50) return "M.E ~AL5";
  if (score >= 35) return "A.E ~AL4";
  if (score >= 25) return "A.E ~AL3";
  if (score >= 15) return "B.E ~AL2";
  return "B.E ~AL1";
}

function getGeneralComment(rubric) {
  switch (rubric) {
    case "E.E ~AL8":
    case "E.E ~AL7":
      return "Excellent performance! Keep aiming high.";
    case "M.E ~AL6":
    case "M.E ~AL5":
      return "Good job! You're meeting expectations.";
    case "A.E ~AL4":
    case "A.E ~AL3":
      return "Fair performance. Put in more effort.";
    case "B.E ~AL2":
    case "B.E ~AL1":
      return "Needs improvement. Stay focused!";
    default:
      return "No comment available.";
  }
}

function insertStudentChart(name, cat1_scores, cat2_scores, folder, docId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName('TempChart');
  if (existing) ss.deleteSheet(existing);

  const chartSheet = ss.insertSheet('TempChart');
  const subjects = ["ENG", "KIS", "MATH", "SCI", "SST", "ART", "PRETECH", "AGRIC", "CRE"];
  chartSheet.getRange(1, 1, 1, 3).setValues([["Assessment", "CAT 1", "CAT 2"]]);

  for (let i = 0; i < subjects.length; i++) {
    chartSheet.getRange(i + 2, 1, 1, 3).setValues([[subjects[i], cat1_scores[i], cat2_scores[i]]]);
  }

  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartSheet.getRange(1, 1, subjects.length + 1, 3))
    .setPosition(1, 5, 0, 0)
    .build();

  chartSheet.insertChart(chart);
  const blob = chartSheet.getCharts()[0].getAs('image/png');
  const imgFile = folder.createFile(blob).setName(`${name}_Chart.png`);

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  body.appendParagraph("\nPerformance Chart:");
  body.appendImage(blob);
  doc.saveAndClose();

  ss.deleteSheet(chartSheet);
}

function generateRubricSummaryTotalAverage(data, folder) {
  let rubricCount = {
    'E.E ~AL8': 0, 'E.E ~AL7': 0, 'M.E ~AL6': 0, 'M.E ~AL5': 0,
    'A.E ~AL4': 0, 'A.E ~AL3': 0, 'B.E ~AL2': 0, 'B.E ~AL1': 0
  };

  for (let i = 1; i < data.length; i++) {
    let total = 0;
    let count = 0;

    for (let j = 2; j < data[i].length; j++) { // Start from column 2 (index 2)
      let score = Number(data[i][j]);
      if (!isNaN(score)) {
        total += score;
        count++;
      }
    }

    const avg = (total / count).toFixed(2);
    const rubric = getRubric(avg);

    if (rubric in rubricCount) rubricCount[rubric]++;
  }

  const doc = DocumentApp.create("Overall Rubric Summary");
  const body = doc.getBody();
  body.appendParagraph("Rubric Summary Based on Total Average per Student");

  for (let key in rubricCount) {
    body.appendParagraph(`${key}: ${rubricCount[key]}`);
  }

  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}

function generateRubricSummarySheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Rubric Summary';
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(sheetName);

  const subjects = ["ENG", "KIS", "MATH", "SCI", "SST", "ART", "PRETECH", "AGRIC", "CRE"];
  const rubricLevels = ['E.E ~AL8', 'E.E ~AL7', 'M.E ~AL6', 'M.E ~AL5', 'A.E ~AL4', 'A.E ~AL3', 'B.E ~AL2', 'B.E ~AL1'];

  // Initialize an object to hold rubric count per subject
  let rubricTable = {};
  for (let rubric of rubricLevels) {
    rubricTable[rubric] = Array(subjects.length).fill(0);
  }

  // Loop over each student
  for (let i = 1; i < data.length; i++) {
    for (let s = 0; s < subjects.length; s++) {
      let cat1 = Number(data[i][2 + s * 2]);
      let cat2 = Number(data[i][3 + s * 2]);

      if (!isNaN(cat1) && !isNaN(cat2)) {
        let avg = ((cat1 + cat2) / 2).toFixed(2);
        let rubric = getRubric(avg);
        if (rubric in rubricTable) {
          rubricTable[rubric][s]++;
        }
      }
    }
  }

  // Output to sheet
  const headerRow = ['Rubric'].concat(subjects);
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

  let output = rubricLevels.map(rubric => [rubric].concat(rubricTable[rubric]));
  sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📄 Report Cards')
    .addItem('📂 Open Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Report Card Generator')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
