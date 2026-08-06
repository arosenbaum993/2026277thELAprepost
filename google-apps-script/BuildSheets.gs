/**
 * BuildSheets.gs — constructs the whole reporting workbook inside Google Sheets.
 *
 * Every computed cell is a live formula, so the sheet recalculates when a teacher
 * edits a score, a cut point, or a threshold. Formula logic matches the Excel
 * build in reporting/build_reporting_workbook.py, which was verified cell by cell
 * against independently computed values.
 *
 * Run buildAll() once from the menu: ELA Reporting ▸ Set up workbook.
 */

// ---- Student Report column map -------------------------------------------
var R_ID = 1, R_NAME = 2, R_PER = 3;
var R_PRE_T = 4, R_PRE_P = 5, R_PRE_L = 6;
var R_POST_T = 7, R_POST_P = 8, R_POST_L = 9;
var R_GAIN = 10, R_PCTPOSS = 11, R_RCI = 12, R_REL = 13, R_TGT = 14;
var R_PRECAT = 15;                       // O..U
var R_POSTCAT = R_PRECAT + 7;            // V..AB
var R_CHG = R_POSTCAT + 7;               // AC..AI
var R_WEAK1 = R_CHG + 7;                 // AJ
var R_WEAK2 = R_WEAK1 + 1;               // AK
var R_NEAR = R_WEAK2 + 1;                // AL
var R_ESSAY = R_NEAR + 1;                // AM
var R_FOCUS = R_ESSAY + 1;               // AN
var R_SQ_PRE = R_FOCUS + 2;              // AP
var R_SQ_POST = R_SQ_PRE + 1;            // AQ

var PRE_CAT_COLS = [], POST_CAT_COLS = [], CHG_CAT_COLS = [];
for (var _i = 0; _i < 7; _i++) {
  PRE_CAT_COLS.push(colA1(R_PRECAT + _i));
  POST_CAT_COLS.push(colA1(R_POSTCAT + _i));
  CHG_CAT_COLS.push(colA1(R_CHG + _i));
}

// ---- helpers --------------------------------------------------------------

function sheetNamed_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); sh.clearConditionalFormatRules(); }
  else { sh = ss.insertSheet(name); }
  return sh;
}

function titleBar_(sh, text, width) {
  sh.getRange(1, 1, 1, width).merge()
    .setValue(text)
    .setBackground(COLORS.navy).setFontColor('#FFFFFF')
    .setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
}

function subtitle_(sh, row, text, width) {
  sh.getRange(row, 1, 1, width).merge()
    .setValue(text).setFontFamily('Arial').setFontSize(9)
    .setFontStyle('italic').setFontColor(COLORS.muted).setWrap(true);
}

function sectionBar_(sh, row, text, width) {
  sh.getRange(row, 1, 1, width).merge()
    .setValue(text).setBackground(COLORS.light).setFontColor(COLORS.navy)
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold');
}

function headerRow_(sh, row, col, values) {
  sh.getRange(row, col, 1, values.length).setValues([values])
    .setBackground(COLORS.blue).setFontColor('#FFFFFF')
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true).setBorder(true, true, true, true, true, true);
}

function addRule_(sh, ranges, formula, bg, fontColor) {
  var b = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula).setBackground(bg).setRanges(ranges);
  if (fontColor) b = b.setFontColor(fontColor);
  var rules = sh.getConditionalFormatRules();
  rules.push(b.build());
  sh.setConditionalFormatRules(rules);
}

// ---- main -----------------------------------------------------------------

function buildAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Set up reporting workbook',
    'This rebuilds all ten sheets. Any scores already entered on the entry sheets will be ERASED.\n\n' +
    'Continue?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  buildStart_(ss);
  buildSettings_(ss);
  buildRoster_(ss);
  buildEntry_(ss, SHEETS.pre, 'FORM A');
  buildEntry_(ss, SHEETS.post, 'FORM B');
  buildStudentReport_(ss);
  buildDashboard_(ss);
  buildItemAnalysis_(ss);
  buildGroups_(ss);
  buildCoverage_(ss);

  var order = [SHEETS.start, SHEETS.settings, SHEETS.roster, SHEETS.pre, SHEETS.post,
               SHEETS.report, SHEETS.dash, SHEETS.items, SHEETS.groups, SHEETS.coverage];
  order.forEach(function (n, i) {
    var sh = ss.getSheetByName(n);
    ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1);
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  ss.setActiveSheet(ss.getSheetByName(SHEETS.start));
  ui.alert('Done', 'Workbook built. Start on the "START HERE" sheet.', ui.ButtonSet.OK);
}

// ---- START HERE -----------------------------------------------------------

function buildStart_(ss) {
  var sh = sheetNamed_(ss, SHEETS.start);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Grade 7 ELA Pre/Post Assessment — Results Reporting', 4);
  sh.setColumnWidth(1, 30); sh.setColumnWidth(2, 240); sh.setColumnWidth(3, 700);

  var rows = [
    ['SECTION', 'Do this, in this order', ''],
    ['', '1.  Roster', 'Enter your students. Yellow cells are inputs; everything else is a formula. Row 6 holds an example — overwrite it.'],
    ['', '2.  Pre-Test Entry', 'After Form A, enter points earned per item. Enter 0 for a wrong answer — a blank means "not tested".'],
    ['', '3.  Student Report', 'Computed the moment you finish entering.'],
    ['', '4.  Intervention Groups', 'Group rosters build themselves. Read them and teach.'],
    ['', '5.  Post-Test Entry', 'After Form B in the spring, enter scores the same way.'],
    ['', '6.  Class Dashboard', 'Growth, effect size, and level movement, filterable by class period.'],
    ['SECTION', 'Auto-scoring from a Google Form (optional)', ''],
    ['', 'Create the answer sheet', 'ELA Reporting ▸ Create answer-sheet Form ▸ Form A. Students read the printed booklet and record answers in the Form.'],
    ['', 'Scoring', 'Responses score automatically on submit — partial credit included. Or run ELA Reporting ▸ Score Form responses now.'],
    ['', 'The essay is never auto-scored', 'Enter the three trait scores by hand in the Q39_T1 / T2 / T3 columns. A rubric score is a professional judgment.'],
    ['SECTION', 'Cell colors', ''],
    ['', 'Yellow', 'You type here. The only cells you should edit.'],
    ['', 'White / grey', 'Formulas. Editing one breaks that student’s report.'],
    ['', 'Blue text', 'Values you may deliberately change (cut scores, thresholds) on the Settings sheet.'],
    ['SECTION', 'Three things to know before reporting any of this', ''],
    ['', 'Category scores are small', 'A category rests on as few as 6 points. Treat a low category percent as a HYPOTHESIS to confirm against student work — never as a diagnosis on its own.'],
    ['', 'Not every gain is real', 'The Reliable Change Index says whether an individual gain exceeds measurement error. Report smaller gains as class-level movement, not student growth.'],
    ['', 'Cut scores are provisional', 'Set by content review, not a state standard-setting panel. Instructional triage — NOT a prediction of a student’s Georgia Milestones level.'],
    ['SECTION', 'Source of truth', ''],
    ['', '', 'Georgia’s K-12 ELA Standards, Grade 7 (May 2023) · Georgia Milestones Grade 7 ELA Assessment Blueprint (August 2025) · Draft Grade 7 ELA Achievement Level Descriptors (October 2025) · Georgia Milestones Grade 7 Eight-Point, Three-Trait Writing Rubric (August 2026) · Grade 7 Classroom Peer Revision Guidance (February 2026).']
  ];
  var r = 3;
  rows.forEach(function (row) {
    if (row[0] === 'SECTION') {
      sectionBar_(sh, r, row[1], 3);
    } else {
      if (row[1]) sh.getRange(r, 2).setValue(row[1]).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setWrap(true).setVerticalAlignment('top');
      sh.getRange(r, 3).setValue(row[2]).setFontFamily('Arial').setFontSize(10).setWrap(true).setVerticalAlignment('top');
    }
    r++;
  });
}

// ---- Settings -------------------------------------------------------------

function buildSettings_(ss) {
  var sh = sheetNamed_(ss, SHEETS.settings);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Settings — cut scores, thresholds, and the intervention lookup table', 5);
  subtitle_(sh, 2, 'Blue values are editable. Every formula in this workbook reads from these cells, so changing one here updates every report.', 5);
  [340, 100, 90, 430, 430].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  function setting(row, label, value, note, fmt, isFormula) {
    sh.getRange(row, 1).setValue(label).setFontFamily('Arial').setFontSize(10);
    var c = sh.getRange(row, 2).setValue(value)
      .setHorizontalAlignment('center').setBorder(true, true, true, true, false, false)
      .setFontFamily('Arial').setFontSize(10);
    if (isFormula) c.setBackground(COLORS.grey).setFontWeight('bold');
    else c.setBackground(COLORS.input).setFontColor(COLORS.inputText);
    if (fmt) c.setNumberFormat(fmt);
    sh.getRange(row, 4).setValue(note).setFontFamily('Arial').setFontSize(9)
      .setFontStyle('italic').setFontColor(COLORS.muted).setWrap(true);
  }

  sectionBar_(sh, 4, 'Achievement level cut scores  (raw points out of 60)', 5);
  setting(5, 'Developing Learner — minimum raw score', 30, 'Below this a student is a Beginning Learner. Default = 50% of 60.');
  setting(6, 'Proficient Learner — minimum raw score', 39, 'Default = 65% of 60.');
  setting(7, 'Distinguished Learner — minimum raw score', 49, 'Default = 82% of 60.');
  setting(8, 'Total points possible', 60, 'Change only if you shorten the instrument.');

  sectionBar_(sh, 10, 'Category mastery thresholds  (percent of category points)', 5);
  setting(11, 'Mastered — minimum percent', 0.80, 'At or above: maintain through spiral review.', '0%');
  setting(12, 'Approaching — minimum percent', 0.60, 'Between this and Mastered: small-group instruction. Below: explicit reteaching.', '0%');

  sectionBar_(sh, 14, 'Growth measurement', 5);
  setting(15, 'Individual growth target (raw points)', 9, 'About half a standard deviation. Set BELOW the reliable-change threshold so it works as a goal, not a statistical claim.');
  setting(16, 'Assumed reliability (Cronbach’s alpha)', 0.85, 'Replace with the value computed on the Item Analysis sheet once you have a full class of data.', '0.00');
  setting(17, 'Standard error of measurement (computed)',
    "=IFERROR(ROUND(STDEV('" + SHEETS.report + "'!$D$" + FIRST_ROW + ":$D$" + LAST_ROW + ")*SQRT(1-$B$16),2),\"\")",
    'SD of pre-test scores × SQRT(1 − reliability). Populates once pre-test data is entered.', '0.00', true);
  setting(18, 'Reliable change threshold (raw points)',
    '=IFERROR(ROUND($B$17*SQRT(2)*1.96,1),"")',
    'A gain at or above this is statistically reliable for an individual student (95% confidence). Smaller gains may be measurement noise.', '0.0', true);

  sectionBar_(sh, 20, 'Scoring quality (enter after double-scoring the essays)', 5);
  setting(21, 'Essay double-scoring exact-agreement rate', '', 'Below 70%, do not report individual essay growth — report class-level trait movement only.', '0%');
  setting(22, 'Number of essays double-scored', '', 'Target: at least 20% of papers.');

  sectionBar_(sh, 24, 'Reporting category reference and intervention lookup', 5);
  subtitle_(sh, 25, 'Column C max points replicate the Georgia Milestones Grade 7 ELA blueprint exactly. Columns D and E feed the Recommended Focus column on the Student Report.', 5);
  headerRow_(sh, 27, 1, ['Code', 'Category', 'Max pts', 'Primary instructional focus', 'Where to start']);

  var vals = CATEGORIES.map(function (c) { return [c.code, c.name, c.max, c.focus, c.start]; });
  sh.getRange(SET_CAT_FIRST, 1, vals.length, 5).setValues(vals)
    .setFontFamily('Arial').setFontSize(9).setWrap(true).setVerticalAlignment('top')
    .setBorder(true, true, true, true, true, true);
  sh.getRange(SET_CAT_FIRST, 1, vals.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(SET_CAT_FIRST, 3, vals.length, 1).setHorizontalAlignment('center');

  var totRow = SET_CAT_FIRST + CATEGORIES.length;
  sh.getRange(totRow, 2).setValue('TOTAL').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(totRow, 3).setFormula('=SUM(C' + SET_CAT_FIRST + ':C' + (totRow - 1) + ')')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // Class-period filter helper used by the Class Dashboard.
  sh.getRange('Y1').setValue('Helper: class-period filter used by the Class Dashboard')
    .setFontSize(9).setFontStyle('italic').setFontColor(COLORS.muted);
  sh.getRange('Z1').setFormula("=IF('" + SHEETS.dash + "'!$B$4=\"\",\"*\",'" + SHEETS.dash + "'!$B$4)");
}

// ---- Roster ---------------------------------------------------------------

function buildRoster_(ss) {
  var sh = sheetNamed_(ss, SHEETS.roster);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Roster', 9);
  subtitle_(sh, 2, 'Every yellow cell is an input. Row 6 is an example — overwrite it. Student order here fixes the row order on every other sheet.', 9);
  sh.getRange(3, 1).setFormula('="Students on roster: "&COUNTA(C' + FIRST_ROW + ':C' + LAST_ROW + ')').setFontWeight('bold');

  var cols = ['Student ID', 'Last Name', 'First Name', 'Class Period', 'Form Order',
              'Accommodations', 'EL', 'SWD', 'Gifted'];
  var widths = [90, 120, 110, 90, 95, 190, 50, 50, 60];
  headerRow_(sh, HDR_ROW, 1, cols);
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  sh.getRange(FIRST_ROW, 1, 1, 9)
    .setValues([['7001', 'Example', 'Student', '1st', 'A then B', 'None', 'N', 'N', 'N']]);
  sh.getRange(FIRST_ROW, 1, N_STUDENTS, 9)
    .setBackground(COLORS.input).setFontColor(COLORS.inputText)
    .setFontFamily('Arial').setFontSize(10)
    .setBorder(true, true, true, true, true, true);
  sh.getRange(FIRST_ROW, 4, N_STUDENTS, 6).setHorizontalAlignment('center');

  sh.getRange(FIRST_ROW, 5, N_STUDENTS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['A then B', 'B then A'], true).build());
  sh.getRange(FIRST_ROW, 7, N_STUDENTS, 3).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).build());

  sh.setFrozenRows(HDR_ROW);
  sh.setFrozenColumns(3);
}

// ---- Entry sheets ---------------------------------------------------------

function buildEntry_(ss, name, formLabel) {
  var sh = sheetNamed_(ss, name);
  sh.setHiddenGridlines(true);
  titleBar_(sh, name + ' — ' + formLabel, TOTAL_COL);
  subtitle_(sh, 2, 'Enter POINTS EARNED per item. Enter 0 for a wrong answer — a blank means "not tested" and removes the student from the statistics. ' +
    'Row 3 shows the reporting category; row 4 shows the maximum points.', TOTAL_COL);

  sh.getRange(CAT_ROW, 1).setValue('Category →').setFontSize(9).setFontStyle('italic').setHorizontalAlignment('right');
  sh.getRange(MAX_ROW, 1).setValue('Max points →').setFontSize(9).setFontStyle('italic').setHorizontalAlignment('right');

  headerRow_(sh, HDR_ROW, 1, ['Student ID', 'Student Name']);
  sh.setColumnWidth(1, 85); sh.setColumnWidth(2, 165);

  var cats = [], maxes = [], labels = [];
  ITEMS.forEach(function (it) { cats.push(it.cat); maxes.push(it.max); labels.push(it.label); });

  sh.getRange(CAT_ROW, FIRST_ITEM_COL, 1, N_ITEMS).setValues([cats])
    .setBackground(COLORS.light).setFontSize(8).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  sh.getRange(MAX_ROW, FIRST_ITEM_COL, 1, N_ITEMS).setValues([maxes])
    .setBackground(COLORS.grey).setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
  headerRow_(sh, HDR_ROW, FIRST_ITEM_COL, labels);
  for (var i = 0; i < N_ITEMS; i++) sh.setColumnWidth(FIRST_ITEM_COL + i, 42);

  headerRow_(sh, HDR_ROW, COND_COL, ['Essay Cond Code']);
  headerRow_(sh, HDR_ROW, TOTAL_COL, ['TOTAL']);
  sh.setColumnWidth(COND_COL, 80); sh.setColumnWidth(TOTAL_COL, 65);
  sh.getRange(MAX_ROW, TOTAL_COL).setFormula('=SUM(' + CL_FIRST + MAX_ROW + ':' + CL_LAST + MAX_ROW + ')')
    .setBackground(COLORS.grey).setFontWeight('bold').setHorizontalAlignment('center');

  var idF = [], nameF = [], totF = [];
  for (var r = FIRST_ROW; r <= LAST_ROW; r++) {
    idF.push(["=IF(" + SHEETS.roster + "!$A" + r + "=\"\",\"\"," + SHEETS.roster + "!$A" + r + ")"]);
    nameF.push(["=IF(" + SHEETS.roster + "!$B" + r + "=\"\",\"\"," + SHEETS.roster + "!$C" + r + "&\" \"&" + SHEETS.roster + "!$B" + r + ")"]);
    totF.push(["=IF(COUNT(" + CL_FIRST + r + ":" + CL_LAST + r + ")=0,\"\",SUM(" + CL_FIRST + r + ":" + CL_LAST + r + "))"]);
  }
  sh.getRange(FIRST_ROW, 1, N_STUDENTS, 1).setFormulas(idF).setHorizontalAlignment('center').setFontSize(9);
  sh.getRange(FIRST_ROW, 2, N_STUDENTS, 1).setFormulas(nameF).setFontSize(9);
  sh.getRange(FIRST_ROW, TOTAL_COL, N_STUDENTS, 1).setFormulas(totF)
    .setFontWeight('bold').setHorizontalAlignment('center');

  var entry = sh.getRange(FIRST_ROW, FIRST_ITEM_COL, N_STUDENTS, N_ITEMS + 1);
  entry.setBackground(COLORS.input).setFontColor(COLORS.inputText)
    .setHorizontalAlignment('center').setFontFamily('Arial').setFontSize(9)
    .setBorder(true, true, true, true, true, true);
  sh.getRange(FIRST_ROW, COND_COL, N_STUDENTS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['A', 'B', 'C', 'D', 'E'], true).build());

  // Flag any entry above the item maximum.
  addRule_(sh, [sh.getRange(FIRST_ROW, FIRST_ITEM_COL, N_STUDENTS, N_ITEMS)],
    '=AND(' + CL_FIRST + FIRST_ROW + '<>"",' + CL_FIRST + FIRST_ROW + '>' + CL_FIRST + '$' + MAX_ROW + ')',
    COLORS.red, '#9C0006');

  sh.setFrozenRows(HDR_ROW);
  sh.setFrozenColumns(2);
}

// ---- Student Report -------------------------------------------------------

function buildStudentReport_(ss) {
  var sh = sheetNamed_(ss, SHEETS.report);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Student Report — individual profiles, growth, and recommended focus', R_FOCUS);
  subtitle_(sh, 2, 'Every cell here is computed. "NEAR CUT" means the score is within one standard error of a level boundary — for those students trust the ' +
    'category profile and your own classroom evidence, not the level label.', R_FOCUS);

  function banner(c1, c2, text, color) {
    sh.getRange(3, c1, 1, c2 - c1 + 1).merge().setValue(text)
      .setBackground(color).setFontColor('#FFFFFF').setFontSize(9)
      .setFontWeight('bold').setHorizontalAlignment('center');
  }
  banner(R_PRE_T, R_PRE_L, 'PRE-TEST (Form A)', '#7B7B7B');
  banner(R_POST_T, R_POST_L, 'POST-TEST (Form B)', '#4472A8');
  banner(R_GAIN, R_TGT, 'GROWTH', '#548235');
  banner(R_PRECAT, R_PRECAT + 6, 'CATEGORY % — PRE', '#7B7B7B');
  banner(R_POSTCAT, R_POSTCAT + 6, 'CATEGORY % — POST', '#4472A8');
  banner(R_CHG, R_CHG + 6, 'CATEGORY CHANGE (percentage points)', '#548235');
  banner(R_WEAK1, R_FOCUS, 'ACTION', '#A6531C');

  // Row 4 carries the category codes; INDEX/MATCH reads it to name the weakest area.
  var codes = CATEGORIES.map(function (c) { return c.code; });
  [R_PRECAT, R_POSTCAT, R_CHG].forEach(function (base) {
    sh.getRange(4, base, 1, 7).setValues([codes])
      .setBackground(COLORS.light).setFontSize(8).setHorizontalAlignment('center');
  });

  headerRow_(sh, HDR_ROW, 1, ['Student ID', 'Student Name', 'Period',
    'Total', '%', 'Level', 'Total', '%', 'Level',
    'Raw gain', '% of possible gain', 'RCI', 'Reliable change?', 'Met target?']);
  [85, 155, 60, 55, 55, 135, 55, 55, 135, 60, 75, 55, 80, 70].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  [R_PRECAT, R_POSTCAT, R_CHG].forEach(function (base) {
    headerRow_(sh, HDR_ROW, base, codes);
    for (var i = 0; i < 7; i++) sh.setColumnWidth(base + i, 52);
  });
  headerRow_(sh, HDR_ROW, R_WEAK1, ['Weakest category', '2nd weakest', 'Near cut?', 'Essay flag', 'Recommended focus']);
  [95, 90, 75, 120, 460].forEach(function (w, i) { sh.setColumnWidth(R_WEAK1 + i, w); });

  var PRE = "'" + SHEETS.pre + "'", POST = "'" + SHEETS.post + "'", ROS = SHEETS.roster;
  var CAT_HDR_PRE = '$' + PRE_CAT_COLS[0] + '$4:$' + PRE_CAT_COLS[6] + '$4';
  var CAT_HDR_POST = '$' + POST_CAT_COLS[0] + '$4:$' + POST_CAT_COLS[6] + '$4';

  var block = [], sq = [];
  for (var r = FIRST_ROW; r <= LAST_ROW; r++) {
    var named = ROS + '!$B' + r + '<>""';
    var preHas = 'COUNT(' + PRE + '!$' + CL_FIRST + r + ':$' + CL_LAST + r + ')>0';
    var postHas = 'COUNT(' + POST + '!$' + CL_FIRST + r + ':$' + CL_LAST + r + ')>0';
    var row = [];

    row.push('=IF(' + ROS + '!$A' + r + '="","",' + ROS + '!$A' + r + ')');
    row.push('=IF(' + named + ',' + ROS + '!$C' + r + '&" "&' + ROS + '!$B' + r + ',"")');
    row.push('=IF(' + named + ',' + ROS + '!$D' + r + ',"")');

    row.push('=IF(AND(' + named + ',' + preHas + '),' + PRE + '!$' + CL_TOTAL + r + ',"")');
    row.push('=IFERROR(D' + r + '/' + SHEETS.settings + '!$B$8,"")');
    row.push(levelFormula_('D' + r));
    row.push('=IF(AND(' + named + ',' + postHas + '),' + POST + '!$' + CL_TOTAL + r + ',"")');
    row.push('=IFERROR(G' + r + '/' + SHEETS.settings + '!$B$8,"")');
    row.push(levelFormula_('G' + r));

    row.push('=IF(OR(D' + r + '="",G' + r + '=""),"",G' + r + '-D' + r + ')');
    // Percent-of-possible gain is undefined for a flat or declining score.
    row.push('=IFERROR(IF(OR(J' + r + '="",J' + r + '<=0),"",(G' + r + '-D' + r + ')/(' + SHEETS.settings + '!$B$8-D' + r + ')),"")');
    row.push('=IFERROR(IF(J' + r + '="","",(G' + r + '-D' + r + ')/(' + SHEETS.settings + '!$B$17*SQRT(2))),"")');
    row.push('=IF(L' + r + '="","",IF(ABS(L' + r + ')>=1.96,"Yes","No"))');
    row.push('=IF(J' + r + '="","",IF(J' + r + '>=' + SHEETS.settings + '!$B$15,"Yes","No"))');

    for (var i = 0; i < 7; i++) {
      var sr = SET_CAT_FIRST + i;
      row.push('=IFERROR(IF($D' + r + '="","",SUMIF(' + PRE + '!$' + CL_FIRST + '$' + CAT_ROW + ':$' + CL_LAST + '$' + CAT_ROW +
        ',' + SHEETS.settings + '!$A$' + sr + ',' + PRE + '!$' + CL_FIRST + r + ':$' + CL_LAST + r + ')/' + SHEETS.settings + '!$C$' + sr + '),"")');
    }
    for (var i2 = 0; i2 < 7; i2++) {
      var sr2 = SET_CAT_FIRST + i2;
      row.push('=IFERROR(IF($G' + r + '="","",SUMIF(' + POST + '!$' + CL_FIRST + '$' + CAT_ROW + ':$' + CL_LAST + '$' + CAT_ROW +
        ',' + SHEETS.settings + '!$A$' + sr2 + ',' + POST + '!$' + CL_FIRST + r + ':$' + CL_LAST + r + ')/' + SHEETS.settings + '!$C$' + sr2 + '),"")');
    }
    for (var i3 = 0; i3 < 7; i3++) {
      var pc = PRE_CAT_COLS[i3], oc = POST_CAT_COLS[i3];
      row.push('=IF(OR(' + pc + r + '="",' + oc + r + '=""),"",' + oc + r + '-' + pc + r + ')');
    }

    var pr = '$' + PRE_CAT_COLS[0] + r + ':$' + PRE_CAT_COLS[6] + r;
    var po = '$' + POST_CAT_COLS[0] + r + ':$' + POST_CAT_COLS[6] + r;
    row.push('=IFERROR(IF($G' + r + '<>"",INDEX(' + CAT_HDR_POST + ',MATCH(MIN(' + po + '),' + po + ',0)),' +
             'IF($D' + r + '<>"",INDEX(' + CAT_HDR_PRE + ',MATCH(MIN(' + pr + '),' + pr + ',0)),"")),"")');
    row.push('=IFERROR(IF($G' + r + '<>"",INDEX(' + CAT_HDR_POST + ',MATCH(SMALL(' + po + ',2),' + po + ',0)),' +
             'IF($D' + r + '<>"",INDEX(' + CAT_HDR_PRE + ',MATCH(SMALL(' + pr + ',2),' + pr + ',0)),"")),"")');
    row.push('=IF(OR(G' + r + '="",' + SHEETS.settings + '!$B$17=""),"",IF(OR(ABS(G' + r + '-' + SHEETS.settings + '!$B$5)<=' + SHEETS.settings + '!$B$17,' +
             'ABS(G' + r + '-' + SHEETS.settings + '!$B$6)<=' + SHEETS.settings + '!$B$17,ABS(G' + r + '-' + SHEETS.settings + '!$B$7)<=' + SHEETS.settings + '!$B$17),"NEAR CUT",""))');
    row.push('=IF(' + POST + '!$' + CL_COND + r + '<>"","Post essay: code "&' + POST + '!$' + CL_COND + r +
             ',IF(' + PRE + '!$' + CL_COND + r + '<>"","Pre essay: code "&' + PRE + '!$' + CL_COND + r + ',""))');
    row.push('=IFERROR(IF($' + colA1(R_WEAK1) + r + '="","",INDEX(' + SHEETS.settings + '!$D$' + SET_CAT_FIRST + ':$D$' + (SET_CAT_FIRST + 6) +
             ',MATCH($' + colA1(R_WEAK1) + r + ',' + SHEETS.settings + '!$A$' + SET_CAT_FIRST + ':$A$' + (SET_CAT_FIRST + 6) + ',0))),"")');

    block.push(row);
    sq.push(['=IF(D' + r + '="","",D' + r + '^2)', '=IF(G' + r + '="","",G' + r + '^2)']);
  }

  sh.getRange(FIRST_ROW, 1, N_STUDENTS, R_FOCUS).setFormulas(block)
    .setFontFamily('Arial').setFontSize(9).setBorder(true, true, true, true, true, true);
  sh.getRange(FIRST_ROW, R_SQ_PRE, N_STUDENTS, 2).setFormulas(sq);
  sh.getRange(HDR_ROW, R_SQ_PRE, 1, 2).setValues([['helper: pre^2', 'helper: post^2']])
    .setFontSize(8).setFontStyle('italic');
  sh.hideColumns(R_SQ_PRE, 2);

  sh.getRange(FIRST_ROW, R_PRE_P, N_STUDENTS, 1).setNumberFormat('0%').setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, R_POST_P, N_STUDENTS, 1).setNumberFormat('0%').setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, R_PCTPOSS, N_STUDENTS, 1).setNumberFormat('0%').setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, R_PRECAT, N_STUDENTS, 21).setNumberFormat('0%').setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, R_RCI, N_STUDENTS, 1).setNumberFormat('0.00').setHorizontalAlignment('center');
  [R_ID, R_PER, R_PRE_T, R_POST_T, R_GAIN, R_REL, R_TGT, R_NEAR, R_WEAK1, R_WEAK2]
    .forEach(function (c) { sh.getRange(FIRST_ROW, c, N_STUDENTS, 1).setHorizontalAlignment('center'); });
  [R_PRE_T, R_POST_T, R_GAIN].forEach(function (c) { sh.getRange(FIRST_ROW, c, N_STUDENTS, 1).setFontWeight('bold'); });
  sh.getRange(FIRST_ROW, R_FOCUS, N_STUDENTS, 1).setWrap(true).setVerticalAlignment('top');

  var catRanges = [sh.getRange(FIRST_ROW, R_PRECAT, N_STUDENTS, 7),
                   sh.getRange(FIRST_ROW, R_POSTCAT, N_STUDENTS, 7)];
  var first = colA1(R_PRECAT) + FIRST_ROW, firstPost = colA1(R_POSTCAT) + FIRST_ROW;
  addRule_(sh, [catRanges[0]], '=AND(ISNUMBER(' + first + '),' + first + '>=' + SHEETS.settings + '!$B$11)', COLORS.green);
  addRule_(sh, [catRanges[0]], '=AND(ISNUMBER(' + first + '),' + first + '>=' + SHEETS.settings + '!$B$12)', COLORS.amber);
  addRule_(sh, [catRanges[0]], '=AND(ISNUMBER(' + first + '),' + first + '<' + SHEETS.settings + '!$B$12)', COLORS.red);
  addRule_(sh, [catRanges[1]], '=AND(ISNUMBER(' + firstPost + '),' + firstPost + '>=' + SHEETS.settings + '!$B$11)', COLORS.green);
  addRule_(sh, [catRanges[1]], '=AND(ISNUMBER(' + firstPost + '),' + firstPost + '>=' + SHEETS.settings + '!$B$12)', COLORS.amber);
  addRule_(sh, [catRanges[1]], '=AND(ISNUMBER(' + firstPost + '),' + firstPost + '<' + SHEETS.settings + '!$B$12)', COLORS.red);

  var chgRange = sh.getRange(FIRST_ROW, R_CHG, N_STUDENTS, 7);
  var chgRules = sh.getConditionalFormatRules();
  chgRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#F8696B', SpreadsheetApp.InterpolationType.NUMBER, '-0.3')
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#63BE7B', SpreadsheetApp.InterpolationType.NUMBER, '0.3')
    .setRanges([chgRange]).build());
  sh.setConditionalFormatRules(chgRules);

  addRule_(sh, [sh.getRange(FIRST_ROW, R_NEAR, N_STUDENTS, 1)],
    '=' + colA1(R_NEAR) + FIRST_ROW + '="NEAR CUT"', COLORS.amber, '#7F6000');
  addRule_(sh, [sh.getRange(FIRST_ROW, R_REL, N_STUDENTS, 1)], '=' + colA1(R_REL) + FIRST_ROW + '="Yes"', COLORS.green);
  addRule_(sh, [sh.getRange(FIRST_ROW, R_TGT, N_STUDENTS, 1)], '=' + colA1(R_TGT) + FIRST_ROW + '="Yes"', COLORS.green);

  sh.setFrozenRows(HDR_ROW);
  sh.setFrozenColumns(3);
}

function levelFormula_(cell) {
  var S = SHEETS.settings;
  return '=IF(' + cell + '="","",IF(' + cell + '>=' + S + '!$B$7,"' + LEVELS[3] + '",' +
         'IF(' + cell + '>=' + S + '!$B$6,"' + LEVELS[2] + '",' +
         'IF(' + cell + '>=' + S + '!$B$5,"' + LEVELS[1] + '","' + LEVELS[0] + '"))))';
}
