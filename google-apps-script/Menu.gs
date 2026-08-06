/**
 * Menu.gs — the custom menu that appears when the spreadsheet opens.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ELA Reporting')
    .addItem('1. Set up workbook', 'buildAll')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('2. Answer-sheet Forms (optional)')
      .addItem('Create Form A (Pre-Test)', 'createFormA')
      .addItem('Create Form B (Post-Test)', 'createFormB')
      .addSeparator()
      .addItem('Show Form links', 'showFormLinks'))
    .addItem('3. Score Form responses now', 'scoreAllForms')
    .addSeparator()
    .addItem('4. Generate student reports (Google Doc)', 'generateStudentReports')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Checks')
      .addItem('Run scoring self-test', 'runScoringSelfTest')
      .addItem('About this workbook', 'showAbout'))
    .addToUi();
}

function runScoringSelfTest() {
  var res = testScoringRules();
  var ui = SpreadsheetApp.getUi();
  var failed = res.results.filter(function (l) { return l.indexOf('FAIL') === 0; });
  var msg = res.pass + ' checks passed, ' + res.fail + ' failed.\n\n';
  msg += failed.length ? failed.join('\n') :
    'Partial-credit rules, blueprint point totals, and both answer keys are consistent.';
  ui.alert('Scoring self-test', msg, ui.ButtonSet.OK);
}

function showAbout() {
  SpreadsheetApp.getUi().alert('Grade 7 ELA Pre/Post Reporting',
    'Built to the Georgia Milestones Grade 7 ELA blueprint: 60 points, 41 scored columns, ' +
    'seven reporting categories weighted exactly as the state weights them.\n\n' +
    'Three cautions worth repeating:\n\n' +
    '• A category rests on as few as 6 points. Confirm a low category against student work before acting.\n\n' +
    '• A gain below the reliable-change threshold (Settings B18) may be measurement noise, not growth.\n\n' +
    '• Cut scores were set by content review, not a state standard-setting panel. They are instructional ' +
    'triage, not a prediction of a Georgia Milestones level.\n\n' +
    'Full documentation lives with the assessment files in docs/.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
