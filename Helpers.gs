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

function generateRubricSummary(data, folder) {
  let rubricCount = {
    'E.E ~AL8': 0, 'E.E ~AL7': 0, 'M.E ~AL6': 0, 'M.E ~AL5': 0,
    'A.E ~AL4': 0, 'A.E ~AL3': 0, 'B.E ~AL2': 0, 'B.E ~AL1': 0
  };

  for (let i = 1; i < data.length; i++) {
    const avg = ((Number(data[i][2]) + Number(data[i][3])) / 2).toFixed(2); // Adjust column index if needed
    const rubric = getRubric(avg);
    if (rubric in rubricCount) rubricCount[rubric]++;
  }

  const doc = DocumentApp.create("Rubric Summary");
  const body = doc.getBody();
  body.appendParagraph("Rubric Summary Report");
  for (let key in rubricCount) {
    body.appendParagraph(`${key}: ${rubricCount[key]}`);
  }

  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
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
