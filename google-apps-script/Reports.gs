/**
 * Reports.gs — printable one-page profiles, one per student.
 *
 * Builds a single Google Doc with a page break between students, so a teacher
 * can print once and hand out. Each page carries the student's totals, the
 * seven category results with mastery flags, growth statistics, and the two
 * areas to work on next — written in language a student or a caregiver can read
 * without a measurement glossary.
 *
 * Deliberately NOT included: emailing reports to families. That depends on
 * guardian-contact data this workbook does not hold and on district policy about
 * what may be sent outside the domain. Print or share the Doc instead.
 */

function generateStudentReports() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rep = ss.getSheetByName(SHEETS.report);
  if (!rep) { ui.alert('Run "Set up workbook" first.'); return; }

  var periodResp = ui.prompt('Student reports',
    'Class period to print (leave blank for all students):', ui.ButtonSet.OK_CANCEL);
  if (periodResp.getSelectedButton() !== ui.Button.OK) return;
  var wantPeriod = periodResp.getResponseText().trim();

  var data = rep.getRange(FIRST_ROW, 1, N_STUDENTS, R_FOCUS).getDisplayValues();
  var codes = CATEGORIES.map(function (c) { return c.code; });

  var rows = data.filter(function (r) {
    if (!r[R_NAME - 1]) return false;
    if (!r[R_PRE_T - 1] && !r[R_POST_T - 1]) return false;
    if (wantPeriod && String(r[R_PER - 1]).trim() !== wantPeriod) return false;
    return true;
  });
  if (!rows.length) { ui.alert('No students matched.', 'Check the period name and that scores are entered.', ui.ButtonSet.OK); return; }

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var doc = DocumentApp.create('Grade 7 ELA — Student Reports' +
    (wantPeriod ? ' (' + wantPeriod + ')' : '') + ' — ' + stamp);
  var body = doc.getBody();
  body.setMarginTop(48).setMarginBottom(48).setMarginLeft(54).setMarginRight(54);

  rows.forEach(function (r, idx) {
    if (idx > 0) body.appendPageBreak();
    writeOneReport_(body, r, codes);
  });

  doc.saveAndClose();
  ui.alert('Student reports created',
    rows.length + ' report' + (rows.length === 1 ? '' : 's') + ' written to:\n\n' +
    doc.getName() + '\n\n' + doc.getUrl(), ui.ButtonSet.OK);
}

function writeOneReport_(body, r, codes) {
  var name = r[R_NAME - 1];
  var period = r[R_PER - 1];
  var preT = r[R_PRE_T - 1], postT = r[R_POST_T - 1];
  var preL = r[R_PRE_L - 1], postL = r[R_POST_L - 1];
  var gain = r[R_GAIN - 1], reliable = r[R_REL - 1], metTarget = r[R_TGT - 1];
  var weak1 = r[R_WEAK1 - 1], weak2 = r[R_WEAK2 - 1];
  var near = r[R_NEAR - 1], essayFlag = r[R_ESSAY - 1];
  var focus = r[R_FOCUS - 1];

  var h = body.appendParagraph(name);
  h.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var sub = body.appendParagraph('Grade 7 English Language Arts   ·   Period ' + period +
    '   ·   ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy'));
  sub.setForegroundColor('#5B6470').setFontSize(9);

  // ---- overall -----------------------------------------------------------
  body.appendParagraph('Overall').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var overall = [['', 'Score (of 60)', 'Achievement level']];
  if (preT !== '') overall.push(['Beginning of year', preT, preL]);
  if (postT !== '') overall.push(['End of year', postT, postL]);
  styleTable_(body.appendTable(overall));

  if (near === 'NEAR CUT') {
    var n = body.appendParagraph(
      'Note: this score sits within one standard error of a level boundary. The level label is not a reliable ' +
      'description on its own — read the category results below alongside classroom work.');
    n.setFontSize(9).setItalic(true).setForegroundColor('#7F6000');
  }

  // ---- growth ------------------------------------------------------------
  if (preT !== '' && postT !== '') {
    body.appendParagraph('Growth').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var g = Number(gain);
    var sentence;
    if (reliable === 'Yes' && g > 0) {
      sentence = 'Gained ' + g + ' points from beginning to end of year. This gain is large enough to be ' +
                 'statistically reliable for one student — it is real growth, not measurement noise.';
    } else if (g > 0) {
      sentence = 'Gained ' + g + ' points from beginning to end of year. This is movement in the right ' +
                 'direction, but it is small enough that it may partly reflect measurement error rather than ' +
                 'growth alone.';
    } else if (g === 0) {
      sentence = 'Scored the same at the beginning and end of year.';
    } else {
      sentence = 'Scored ' + Math.abs(g) + ' points lower at the end of year. Before drawing a conclusion, ' +
                 'check testing conditions and effort on the second administration.';
    }
    body.appendParagraph(sentence).setFontSize(10);
    body.appendParagraph('Growth target met: ' + (metTarget || '—')).setFontSize(9)
      .setForegroundColor('#5B6470');
  }

  // ---- category profile --------------------------------------------------
  body.appendParagraph('What the results show, skill by skill')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var tbl = [['Skill area', 'Beginning', 'End', 'Where this stands']];
  codes.forEach(function (code, i) {
    var pre = r[R_PRECAT - 1 + i];
    var post = r[R_POSTCAT - 1 + i];
    var latest = post !== '' ? post : pre;
    tbl.push([CATEGORIES[i].name.replace(/^(Interpreting|Constructing) — /, ''),
              pre || '—', post || '—', masteryWord_(latest)]);
  });
  styleTable_(body.appendTable(tbl));
  body.appendParagraph(
    'Reading (Interpreting) skills are listed first, writing (Constructing) skills second. ' +
    'Each area is worth between 6 and 16 points, so a single percentage here is a starting point for a ' +
    'conversation, not a final verdict.').setFontSize(8.5).setItalic(true).setForegroundColor('#5B6470');

  // ---- next steps --------------------------------------------------------
  body.appendParagraph('What we are working on next').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (weak1) {
    var c1 = catByCode_(weak1);
    body.appendParagraph('Priority: ' + (c1 ? c1.name : weak1)).setFontSize(10).setBold(true);
    body.appendParagraph(focus || (c1 ? c1.focus : '')).setFontSize(10);
    if (weak2) {
      var c2 = catByCode_(weak2);
      body.appendParagraph('Also building: ' + (c2 ? c2.name : weak2)).setFontSize(9)
        .setForegroundColor('#5B6470');
    }
  }
  if (essayFlag) {
    body.appendParagraph(essayFlag + ' — the writing task could not be scored with the rubric, so the ' +
      'writing areas above are based on the other items only.')
      .setFontSize(9).setItalic(true).setForegroundColor('#7F6000');
  }

  body.appendParagraph(
    'About this assessment: a classroom pre/post measure built to the Georgia Milestones Grade 7 ELA blueprint. ' +
    'Achievement levels here are set by teacher content review, not by a state standard-setting panel, and are ' +
    'used to plan instruction — they do not predict a Georgia Milestones score.')
    .setFontSize(8).setItalic(true).setForegroundColor('#5B6470');
}

function masteryWord_(pctText) {
  if (!pctText) return '—';
  var v = parseFloat(String(pctText).replace('%', ''));
  if (isNaN(v)) return '—';
  if (v >= 80) return 'Strength — keep it sharp';
  if (v >= 60) return 'Almost there — small-group work';
  return 'Priority for reteaching';
}

function catByCode_(code) {
  for (var i = 0; i < CATEGORIES.length; i++) if (CATEGORIES[i].code === code) return CATEGORIES[i];
  return null;
}

function styleTable_(table) {
  table.setBorderColor('#C3C9D2');
  var head = table.getRow(0);
  for (var c = 0; c < head.getNumCells(); c++) {
    head.getCell(c).setBackgroundColor('#DCE6F1');
    head.getCell(c).editAsText().setBold(true).setFontSize(9);
  }
  for (var r = 1; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c2 = 0; c2 < row.getNumCells(); c2++) {
      row.getCell(c2).editAsText().setFontSize(9.5);
    }
  }
}
