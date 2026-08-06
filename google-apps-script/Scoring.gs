/**
 * Scoring.gs — auto-scoring engine.
 *
 * Implements the exact partial-credit rules from the answer keys:
 *
 *   Selected response      1 if the key matches, else 0
 *   Two-part / EBSR        1 point per correct part  ->  2 / 1 / 0
 *   Select TWO             2 = both correct, no incorrect
 *                          1 = exactly one correct AND no more than one incorrect
 *                          0 = anything else, INCLUDING three or more selections
 *   Match (4 cells)        2 = all four, 1 = exactly three, 0 = two or fewer
 *
 * The essay is never auto-scored. A rubric score is a professional judgment and
 * the three trait columns stay under the teacher's hand.
 */

var PROP_FORM_A = 'formIdA';
var PROP_FORM_B = 'formIdB';

/** Canonical question titles. The Form builder and the scorer both call this,
 *  so they cannot drift apart. */
function questionTitle_(item, part) {
  var n = item.n;
  switch (item.type) {
    case 'SR':    return 'Item ' + n;
    case 'MP':    return 'Item ' + n + ' — Part ' + part;          // 'A' | 'B'
    case 'MS':    return 'Item ' + n + ' — select TWO';
    case 'MATCH': return 'Item ' + n + ' — row ' + part;           // 'a'..'d'
  }
  return 'Item ' + n;
}

/** Score one item. `resp` is a string, or an array for MS/MP/MATCH. */
function scoreItem_(item, key, resp) {
  if (resp === null || resp === undefined) return null;

  if (item.type === 'SR') {
    if (resp === '') return null;
    return resp === key ? 1 : 0;
  }

  if (item.type === 'MP') {
    if ((resp[0] === '' || resp[0] == null) && (resp[1] === '' || resp[1] == null)) return null;
    var pts = 0;
    if (resp[0] === key[0]) pts++;
    if (resp[1] === key[1]) pts++;
    return pts;
  }

  if (item.type === 'MS') {
    var sel = (resp || []).filter(function (x) { return x !== '' && x != null; });
    if (!sel.length) return null;
    if (sel.length >= 3) return 0;                    // over-selection scores zero
    var correct = 0, incorrect = 0;
    sel.forEach(function (s) { key.indexOf(s) >= 0 ? correct++ : incorrect++; });
    if (correct === 2 && incorrect === 0) return 2;
    if (correct === 1 && incorrect <= 1) return 1;
    return 0;
  }

  if (item.type === 'MATCH') {
    var given = resp || [];
    var any = given.some(function (x) { return x !== '' && x != null; });
    if (!any) return null;
    var hits = 0;
    for (var i = 0; i < 4; i++) if (given[i] === key[i]) hits++;
    if (hits === 4) return 2;
    if (hits === 3) return 1;
    return 0;
  }

  return null;
}

/**
 * Read every response from the stored Form for `formLabel` ('A' or 'B'),
 * score it, and write the item points into the matching entry sheet row.
 * Matching is by Student ID against the Roster.
 */
function scoreFormResponses_(formLabel) {
  var props = PropertiesService.getDocumentProperties();
  var formId = props.getProperty(formLabel === 'A' ? PROP_FORM_A : PROP_FORM_B);
  if (!formId) {
    return {scored: 0, unmatched: [], error:
      'No Form ' + formLabel + ' has been created yet. Use ELA Reporting ▸ Create answer-sheet Form.'};
  }

  var form;
  try { form = FormApp.openById(formId); }
  catch (e) { return {scored: 0, unmatched: [], error: 'Could not open the Form: ' + e.message}; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entry = ss.getSheetByName(formLabel === 'A' ? SHEETS.pre : SHEETS.post);
  var roster = ss.getSheetByName(SHEETS.roster);
  if (!entry || !roster) return {scored: 0, unmatched: [], error: 'Run "Set up workbook" first.'};

  // Student ID -> sheet row
  var ids = roster.getRange(FIRST_ROW, 1, N_STUDENTS, 1).getValues();
  var rowOf = {};
  ids.forEach(function (v, i) {
    var id = String(v[0]).trim();
    if (id) rowOf[id] = FIRST_ROW + i;
  });

  var key = KEYS[formLabel];
  var responses = form.getResponses();
  var scored = 0, unmatched = [];

  responses.forEach(function (fr) {
    var byTitle = {};
    fr.getItemResponses().forEach(function (ir) {
      byTitle[ir.getItem().getTitle()] = ir.getResponse();
    });

    var sid = String(byTitle['Student ID'] || '').trim();
    var row = rowOf[sid];
    if (!row) { if (sid) unmatched.push(sid); return; }

    var points = [];
    SCORED_ITEMS.forEach(function (it) {
      var resp;
      if (it.type === 'SR')        resp = byTitle[questionTitle_(it)] || '';
      else if (it.type === 'MP')   resp = [byTitle[questionTitle_(it, 'A')] || '', byTitle[questionTitle_(it, 'B')] || ''];
      else if (it.type === 'MS')   resp = byTitle[questionTitle_(it)] || [];
      else if (it.type === 'MATCH') resp = ['a', 'b', 'c', 'd'].map(function (p) { return byTitle[questionTitle_(it, p)] || ''; });
      points.push(scoreItem_(it, key[it.n], resp));
    });

    // Write items 1-38 only; the three essay trait columns stay untouched.
    entry.getRange(row, FIRST_ITEM_COL, 1, points.length)
         .setValues([points.map(function (p) { return p === null ? '' : p; })]);
    scored++;
  });

  return {scored: scored, unmatched: unmatched, total: responses.length};
}

/** Menu action: score both forms and report what happened. */
function scoreAllForms() {
  var ui = SpreadsheetApp.getUi();
  var out = [];
  ['A', 'B'].forEach(function (lbl) {
    var r = scoreFormResponses_(lbl);
    if (r.error) { out.push('Form ' + lbl + ': ' + r.error); return; }
    var line = 'Form ' + lbl + ': scored ' + r.scored + ' of ' + r.total + ' responses.';
    if (r.unmatched.length) {
      line += '\n   Unmatched Student IDs (not on the Roster): ' +
              r.unmatched.slice(0, 12).join(', ') + (r.unmatched.length > 12 ? ' …' : '');
    }
    out.push(line);
  });
  out.push('\nReminder: essay trait scores (Q39_T1 / T2 / T3) are entered by hand.');
  ui.alert('Auto-scoring complete', out.join('\n\n'), ui.ButtonSet.OK);
}

/** Installable trigger target — scores a single submission as it arrives. */
function onFormSubmitHandler(e) {
  try {
    var props = PropertiesService.getDocumentProperties();
    var srcId = e && e.source && e.source.getId ? e.source.getId() : null;
    if (srcId === props.getProperty(PROP_FORM_A)) scoreFormResponses_('A');
    else if (srcId === props.getProperty(PROP_FORM_B)) scoreFormResponses_('B');
    else { scoreFormResponses_('A'); scoreFormResponses_('B'); }
  } catch (err) {
    console.error('onFormSubmitHandler: ' + err.message);
  }
}

/**
 * Self-test for the partial-credit rules. Run from the editor and check the log.
 * These assertions encode the scoring table in the answer keys.
 */
function testScoringRules() {
  var results = [], pass = 0, fail = 0;
  function t(label, got, want) {
    var ok = JSON.stringify(got) === JSON.stringify(want);
    ok ? pass++ : fail++;
    results.push((ok ? 'PASS  ' : 'FAIL  ') + label + '  got=' + JSON.stringify(got) + ' want=' + JSON.stringify(want));
  }
  var SR = {n: 1, type: 'SR'}, MP = {n: 2, type: 'MP'}, MS = {n: 8, type: 'MS'}, MT = {n: 22, type: 'MATCH'};

  t('SR correct',        scoreItem_(SR, 'B', 'B'), 1);
  t('SR incorrect',      scoreItem_(SR, 'B', 'A'), 0);
  t('SR blank',          scoreItem_(SR, 'B', ''), null);

  t('MP both',           scoreItem_(MP, ['B', 'C'], ['B', 'C']), 2);
  t('MP part A only',    scoreItem_(MP, ['B', 'C'], ['B', 'D']), 1);
  t('MP part B only',    scoreItem_(MP, ['B', 'C'], ['A', 'C']), 1);
  t('MP neither',        scoreItem_(MP, ['B', 'C'], ['A', 'D']), 0);
  t('MP blank',          scoreItem_(MP, ['B', 'C'], ['', '']), null);

  t('MS both correct',   scoreItem_(MS, ['A', 'C'], ['A', 'C']), 2);
  t('MS one correct only',        scoreItem_(MS, ['A', 'C'], ['A']), 1);
  t('MS one right one wrong',     scoreItem_(MS, ['A', 'C'], ['A', 'B']), 1);
  t('MS two wrong',               scoreItem_(MS, ['A', 'C'], ['B', 'D']), 0);
  t('MS three selected -> 0',     scoreItem_(MS, ['A', 'C'], ['A', 'C', 'B']), 0);
  t('MS four selected -> 0',      scoreItem_(MS, ['A', 'C'], ['A', 'B', 'C', 'D']), 0);
  t('MS blank',                   scoreItem_(MS, ['A', 'C'], []), null);

  var K = ['S1', 'S2', 'B', 'B'];
  t('MATCH all four',    scoreItem_(MT, K, ['S1', 'S2', 'B', 'B']), 2);
  t('MATCH three',       scoreItem_(MT, K, ['S1', 'S2', 'B', 'S1']), 1);
  t('MATCH two',         scoreItem_(MT, K, ['S1', 'S2', 'S1', 'S1']), 0);
  t('MATCH blank',       scoreItem_(MT, K, ['', '', '', '']), null);

  // Blueprint integrity: category point totals must match the state blueprint.
  var totals = {};
  ITEMS.forEach(function (it) { totals[it.cat] = (totals[it.cat] || 0) + it.max; });
  CATEGORIES.forEach(function (c) { t('category ' + c.code + ' points', totals[c.code], c.max); });
  var grand = ITEMS.reduce(function (s, it) { return s + it.max; }, 0);
  t('total points', grand, 60);
  t('item columns', ITEMS.length, 41);

  // Every scored item must have a key on both forms.
  ['A', 'B'].forEach(function (f) {
    var missing = SCORED_ITEMS.filter(function (it) { return KEYS[f][it.n] === undefined; })
                              .map(function (it) { return it.n; });
    t('Form ' + f + ' keys complete', missing, []);
  });
  // Key shapes must match item types.
  ['A', 'B'].forEach(function (f) {
    var bad = SCORED_ITEMS.filter(function (it) {
      var k = KEYS[f][it.n];
      if (it.type === 'SR') return typeof k !== 'string';
      if (it.type === 'MP' || it.type === 'MS') return !(k instanceof Array) || k.length !== 2;
      if (it.type === 'MATCH') return !(k instanceof Array) || k.length !== 4;
      return false;
    }).map(function (it) { return it.n; });
    t('Form ' + f + ' key shapes', bad, []);
  });
  // No item may share a key letter across forms (form-security check).
  var shared = SCORED_ITEMS.filter(function (it) {
    return it.type === 'SR' && KEYS.A[it.n] === KEYS.B[it.n];
  }).map(function (it) { return it.n; });
  t('no shared SR keys across forms', shared, []);

  Logger.log(results.join('\n'));
  Logger.log('\n' + pass + ' passed, ' + fail + ' failed.');
  return {pass: pass, fail: fail, results: results};
}
