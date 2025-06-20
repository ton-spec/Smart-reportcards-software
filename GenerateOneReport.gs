function generateOneReport(row, index, templateFile, folder, term, year, grade, logoBlob) {
  const [
    adm, name, eng_cat1, eng_cat2, kis_cat1, kis_cat2, math_cat1, math_cat2,
    sci_cat1, sci_cat2, sst_cat1, sst_cat2,
    art_cat1, art_cat2, pretech_cat1, pretech_cat2,
    agric_cat1, agric_cat2, cre_cat1, cre_cat2,
    email
  ] = row;

  const cat1_scores = [eng_cat1, kis_cat1, math_cat1, sci_cat1, sst_cat1, art_cat1, pretech_cat1, agric_cat1, cre_cat1].map(Number);
  const cat2_scores = [eng_cat2, kis_cat2, math_cat2, sci_cat2, sst_cat2, art_cat2, pretech_cat2, agric_cat2, cre_cat2].map(Number);
  const average_scores = cat1_scores.map((s1, idx) => ((s1 + cat2_scores[idx]) / 2).toFixed(2));

  const total1 = cat1_scores.reduce((sum, val) => sum + val, 0);
  const total2 = cat2_scores.reduce((sum, val) => sum + val, 0);
  const total3 = average_scores.reduce((sum, val) => sum + Number(val), 0);

  const mean1 = (total1 / cat1_scores.length).toFixed(2);
  const mean2 = (total2 / cat2_scores.length).toFixed(2);
  const mean3 = (total3 / average_scores.length).toFixed(2);

  const rubric1 = getRubric(mean1);
  const rubric2 = getRubric(mean2);
  const rubric3 = getRubric(mean3);

  const docCopy = templateFile.makeCopy(`${name} Report Card`, folder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  const placeholders = {
    '{{NAME}}': name,
    '{{ADM}}': adm,

    '{{ENG}}': average_scores[0],
    '{{KIS}}': average_scores[1],
    '{{MATH}}': average_scores[2],
    '{{SCI}}': average_scores[3],
    '{{SST}}': average_scores[4],
    '{{ART}}': average_scores[5],
    '{{PRETECH}}': average_scores[6],
    '{{AGRIC}}': average_scores[7],
    '{{CRE}}': average_scores[8],

    '{{ENG_CAT1}}': cat1_scores[0],
    '{{KIS_CAT1}}': cat1_scores[1],
    '{{MATH_CAT1}}': cat1_scores[2],
    '{{SCI_CAT1}}': cat1_scores[3],
    '{{SST_CAT1}}': cat1_scores[4],
    '{{ART_CAT1}}': cat1_scores[5],
    '{{PRETECH_CAT1}}': cat1_scores[6],
    '{{AGRIC_CAT1}}': cat1_scores[7],
    '{{CRE_CAT1}}': cat1_scores[8],

    '{{ENG_CAT2}}': cat2_scores[0],
    '{{KIS_CAT2}}': cat2_scores[1],
    '{{MATH_CAT2}}': cat2_scores[2],
    '{{SCI_CAT2}}': cat2_scores[3],
    '{{SST_CAT2}}': cat2_scores[4],
    '{{ART_CAT2}}': cat2_scores[5],
    '{{PRETECH_CAT2}}': cat2_scores[6],
    '{{AGRIC_CAT2}}': cat2_scores[7],
    '{{CRE_CAT2}}': cat2_scores[8],

    '{{ENG_CAT1_RUBRIC}}': getRubric(cat1_scores[0]),
    '{{KIS_CAT1_RUBRIC}}': getRubric(cat1_scores[1]),
    '{{MATH_CAT1_RUBRIC}}': getRubric(cat1_scores[2]),
    '{{SCI_CAT1_RUBRIC}}': getRubric(cat1_scores[3]),
    '{{SST_CAT1_RUBRIC}}': getRubric(cat1_scores[4]),
    '{{ART_CAT1_RUBRIC}}': getRubric(cat1_scores[5]),
    '{{PRETECH_CAT1_RUBRIC}}': getRubric(cat1_scores[6]),
    '{{AGRIC_CAT1_RUBRIC}}': getRubric(cat1_scores[7]),
    '{{CRE_CAT1_RUBRIC}}': getRubric(cat1_scores[8]),

    '{{ENG_CAT2_RUBRIC}}': getRubric(cat2_scores[0]),
    '{{KIS_CAT2_RUBRIC}}': getRubric(cat2_scores[1]),
    '{{MATH_CAT2_RUBRIC}}': getRubric(cat2_scores[2]),
    '{{SCI_CAT2_RUBRIC}}': getRubric(cat2_scores[3]),
    '{{SST_CAT2_RUBRIC}}': getRubric(cat2_scores[4]),
    '{{ART_CAT2_RUBRIC}}': getRubric(cat2_scores[5]),
    '{{PRETECH_CAT2_RUBRIC}}': getRubric(cat2_scores[6]),
    '{{AGRIC_CAT2_RUBRIC}}': getRubric(cat2_scores[7]),
    '{{CRE_CAT2_RUBRIC}}': getRubric(cat2_scores[8]),

    '{{ENG_RUBRIC}}': getRubric(average_scores[0]),
    '{{KIS_RUBRIC}}': getRubric(average_scores[1]),
    '{{MATH_RUBRIC}}': getRubric(average_scores[2]),
    '{{SCI_RUBRIC}}': getRubric(average_scores[3]),
    '{{SST_RUBRIC}}': getRubric(average_scores[4]),
    '{{ART_RUBRIC}}': getRubric(average_scores[5]),
    '{{PRETECH_RUBRIC}}': getRubric(average_scores[6]),
    '{{AGRIC_RUBRIC}}': getRubric(average_scores[7]),
    '{{CRE_RUBRIC}}': getRubric(average_scores[8]),

    '{{TOTAL_CAT1}}': total1.toString(),
    '{{TOTAL_CAT2}}': total2.toString(),
    '{{TOTAL_AVG}}': total3.toString(),

    '{{MEAN_CAT1}}': mean1,
    '{{MEAN_CAT2}}': mean2,
    '{{MEAN_AVG}}': mean3,
    '{{MEAN_CAT1_RUBRIC}}': rubric1,
    '{{MEAN_CAT2_RUBRIC}}': rubric2,
    '{{MEAN_AVG_RUBRIC}}': rubric3,
    '{{AVG_RUBRIC}}': rubric3,
    '{{AVG_COMMENT}}': getGeneralComment(rubric3),

    '{{TERM}}': term,
    '{{YEAR}}': year,
    '{{GRADE}}': grade
  };

  for (let key in placeholders) {
    body.replaceText(key, placeholders[key]);
  }

  const logoTag = body.findText('{{SCHOOL_LOGO}}');
  if (logoTag) {
    const element = logoTag.getElement().getParent();
    const index = body.getChildIndex(element);
    const headerTable = body.insertTable(0);
    const row = headerTable.appendTableRow();
    const logoCell = row.appendTableCell();
    const textCell = row.appendTableCell();
    const logo = logoCell.insertImage(0, logoBlob);
    logo.setWidth(110).setHeight(120);
    logoCell.setWidth(100);
    textCell.setWidth(400);
    textCell.setText(`🏫 Mungakha Junior School\n📍 Bungoma, Kenya\n📅 ${grade}, ${term} Year, ${year}`);
    textCell.setFontSize(20).setBold(true);
    body.removeChild(element);
  }

  doc.saveAndClose();
  insertStudentChart(name, cat1_scores, cat2_scores, folder, docCopy.getId());

  docCopy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Utilities.sleep(1000);

  const downloadLink = docCopy.getUrl();
  const reopenedDoc = DocumentApp.openById(docCopy.getId());
  const reopenedBody = reopenedDoc.getBody();
  reopenedBody.appendParagraph('\nDownload Your Report Card:');
  reopenedBody.appendParagraph(downloadLink).setLinkUrl(downloadLink);
  reopenedDoc.saveAndClose();

  const pdf = docCopy.getAs(MimeType.PDF);
  const pdfFile = folder.createFile(pdf);

  if (email) {
    GmailApp.sendEmail(email, 'Your Report Card',
      `Dear ${name},\n\nAttached is the report card for ${term} ${year}.\n\nDownload Link:\n${downloadLink}\n\nBest regards,\nMungakha Junior School`, {
        attachments: [pdfFile],
        name: 'School Reports'
      });
  }
}
