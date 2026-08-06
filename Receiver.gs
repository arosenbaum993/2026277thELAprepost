/**
 * Receiver.gs — receives submissions from Grade-7-ELA-Pre-Post.html
 *
 * Paste this into the SAME spreadsheet you already set up with the ELA
 * Reporting menu (Extensions ▸ Apps Script ▸ + ▸ Script, name it "Receiver").
 * Then deploy it once as a Web App and paste the URL into the SHEET_URL line
 * near the top of the HTML file's <script> block.
 *
 * DEPLOY:  Deploy ▸ New deployment ▸ type "Web app"
 *          Execute as:      Me
 *          Who has access:  Anyone            <-- required; students are not
 *                                                 signed in to your account
 *          Copy the /exec URL it gives you.
 *
 * Re-deploy after ANY edit to this file: Deploy ▸ Manage deployments ▸
 * pencil icon ▸ Version: New version ▸ Deploy. The URL stays the same.
 *
 * WHAT IT DOES
 *   1. Logs every raw submission to "Submissions (raw)" FIRST, before anything
 *      else can fail. Nothing a student sent is ever lost.
 *   2. Finds the student on the Roster by Student ID, adding them if new.
 *   3. Writes the 38 auto-scored item points into "Pre-Test Entry" (Form A) or
 *      "Post-Test Entry" (Form B) — the same columns you would type into.
 *   4. Files the essay on an "Essays" tab for you to rubric-score.
 *
 * The three essay trait columns (Q39_T1/T2/T3) are deliberately left BLANK.
 * A rubric score is your judgment, not the computer's.
 */

/**
 * Leave this EMPTY if you pasted this script inside the spreadsheet itself
 * (Extensions ▸ Apps Script from the sheet). It then finds your spreadsheet
 * automatically and no ID is needed.
 *
 * Only fill it in if you made a STANDALONE script project instead, in which
 * case getActiveSpreadsheet() returns null and the script cannot find the
 * sheet. Paste the long id from your spreadsheet's address bar:
 *   docs.google.com/spreadsheets/d/THIS_PART/edit
 *
 * Note: this file lives in a public repository. Adding your id here and
 * committing it would publish which document holds your student data, so
 * prefer the container-bound route above and leave this blank.
 */
var SPREADSHEET_ID = '';

var ROSTER = 'Roster';
var PRE    = 'Pre-Test Entry';
var POST   = 'Post-Test Entry';
var RAW    = 'Submissions (raw)';
var ESSAYS = 'Essays';

/** The spreadsheet this script writes to, whichever way it was installed. */
function book_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  throw new Error(
    'No spreadsheet found. Paste this script into the spreadsheet itself ' +
    '(Extensions > Apps Script), or set SPREADSHEET_ID at the top of this file.');
}

var FIRST_ROW      = 6;   // first student row on every sheet
var FIRST_ITEM_COL = 3;   // column C holds item 1
var N_ITEMS        = 38;  // items 1-38 are auto-scored

function doPost(e) {
  var lock = LockService.getScriptLock();
  // A whole class submits within about a minute. Without this, two writes can
  // land on the same row and one silently overwrites the other.
  try { lock.waitLock(30000); } catch (err) {
    return reply_({ok: false, error: 'busy, please resubmit'});
  }
  try {
    var body = (e && e.postData && e.postData.contents) || '';
    var d = JSON.parse(body);
    var ss = book_();

    logRaw_(ss, body, d);                       // safety net first

    var row = findOrAddStudent_(ss, d);
    var target = (String(d.form).toUpperCase() === 'B') ? POST : PRE;
    writeScores_(ss, target, row, d);
    writeEssay_(ss, d);

    return reply_({ok: true, row: row, sheet: target});
  } catch (err) {
    try {
      book_()
        .getSheetByName(RAW)
        .appendRow([new Date(), 'ERROR', err.message, (e && e.postData) ? e.postData.contents : '']);
    } catch (ignored) {}
    return reply_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

/** Browsers post with mode:"no-cors", so the student never reads this. It is
 *  here for you when testing the URL directly in a browser tab. */
function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(
    'Grade 7 ELA receiver is deployed and listening. Submissions arrive by POST.');
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#2E5C8A').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function logRaw_(ss, body, d) {
  var sh = sheet_(ss, RAW,
    ['Received', 'Form', 'Student ID', 'Last', 'First', 'Period',
     'Minutes', 'Auto-scored total', 'Raw JSON']);
  sh.appendRow([new Date(), d.form, d.studentId, d.last, d.first, d.period,
                d.minutes, d.rawAuto, body]);
}

function findOrAddStudent_(ss, d) {
  var sh = ss.getSheetByName(ROSTER);
  if (!sh) throw new Error('No "' + ROSTER + '" sheet. Run ELA Reporting ▸ Set up workbook first.');

  var last = Math.max(sh.getLastRow(), FIRST_ROW);
  var n = last - FIRST_ROW + 1;
  var ids = n > 0 ? sh.getRange(FIRST_ROW, 1, n, 1).getValues() : [];
  var want = String(d.studentId).trim();

  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === want && want !== '') return FIRST_ROW + i;
  }
  // Not on the roster yet — add them rather than dropping the submission.
  var row = FIRST_ROW;
  while (row <= sh.getLastRow() && String(sh.getRange(row, 1).getValue()).trim() !== '') row++;
  sh.getRange(row, 1, 1, 4).setValues([[want, d.last || '', d.first || '', d.period || '']]);
  return row;
}

function writeScores_(ss, name, row, d) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('No "' + name + '" sheet. Run ELA Reporting ▸ Set up workbook first.');

  var pts = d.itemPoints || {};
  var out = [];
  for (var n = 1; n <= N_ITEMS; n++) {
    var v = pts[String(n)];
    out.push((v === undefined || v === null) ? '' : v);
  }
  sh.getRange(row, FIRST_ITEM_COL, 1, N_ITEMS).setValues([out]);
  // Q39_T1 / T2 / T3 are intentionally NOT written — you score the essay.
}

function writeEssay_(ss, d) {
  var sh = sheet_(ss, ESSAYS,
    ['Received', 'Form', 'Student ID', 'Last', 'First', 'Period',
     'Words', 'Essay', 'Trait 1 (0-3)', 'Trait 2 (0-3)', 'Trait 3 (0-2)']);
  var text = d.essay || '';
  var words = text.trim() ? text.trim().split(/\s+/).length : 0;
  sh.appendRow([new Date(), d.form, d.studentId, d.last, d.first, d.period, words, text]);
  sh.setColumnWidth(8, 520);
  sh.getRange(sh.getLastRow(), 8).setWrap(true).setVerticalAlignment('top');
}

/**
 * Run this once from the editor (Run ▸ testReceiver) BEFORE giving the link to
 * students. It posts a fake submission for student ID TEST-001 so you can see
 * a row appear on every tab. Delete that row afterward.
 */
function testReceiver() {
  var pts = {};
  for (var n = 1; n <= N_ITEMS; n++) pts[String(n)] = (n % 3 === 0) ? 0 : 1;
  var fake = {
    ts: new Date().toISOString(), form: 'A', studentId: 'TEST-001',
    last: 'Test', first: 'Student', period: '1st', minutes: 42, rawAuto: 25,
    essay: 'This is a test essay so you can see where essays land.',
    answers: {}, itemPoints: pts, categories: {}
  };
  var res = doPost({postData: {contents: JSON.stringify(fake)}});
  Logger.log(res.getContent());
}
