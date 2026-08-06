/**
 * BuildSheets2.gs — Class Dashboard, Item Analysis, Intervention Groups,
 * Standards Coverage.
 *
 * The Intervention Groups sheet uses TEXTJOIN + FILTER to build group rosters
 * automatically. That is a genuine Google Sheets advantage: the Excel version
 * asks the teacher to sort a column by hand, while here the group lists write
 * themselves and update as scores change.
 */

function buildDashboard_(ss) {
  var sh = sheetNamed_(ss, SHEETS.dash);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Class Dashboard', 7);
  subtitle_(sh, 2, 'Type a class period in cell B4 to filter every statistic on this sheet to that period. Leave B4 blank for all students.', 7);
  [330, 110, 110, 110, 120, 130, 200].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  sh.getRange(4, 1).setValue('Filter to class period (blank = all):')
    .setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(4, 2).setBackground(COLORS.input).setFontColor(COLORS.inputText)
    .setHorizontalAlignment('center').setBorder(true, true, true, true, false, false);

  var REP = "'" + SHEETS.report + "'";
  var S = SHEETS.settings;
  var PER = S + '!$Z$1';
  var PRE_T = REP + '!$D$' + FIRST_ROW + ':$D$' + LAST_ROW;
  var POST_T = REP + '!$G$' + FIRST_ROW + ':$G$' + LAST_ROW;
  var PER_R = REP + '!$C$' + FIRST_ROW + ':$C$' + LAST_ROW;
  var SQ_PRE = REP + '!$' + colA1(R_SQ_PRE) + '$' + FIRST_ROW + ':$' + colA1(R_SQ_PRE) + '$' + LAST_ROW;
  var SQ_POST = REP + '!$' + colA1(R_SQ_POST) + '$' + FIRST_ROW + ':$' + colA1(R_SQ_POST) + '$' + LAST_ROW;

  sectionBar_(sh, 6, 'Overall performance', 7);
  headerRow_(sh, 7, 1, ['', 'Pre-Test', 'Post-Test', 'Change']);

  var stats = [
    ['Students with scores',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + PRE_T + ',">=0")',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + POST_T + ',">=0")', '', '0'],
    ['Mean raw score (of 60)',
     '=IFERROR(AVERAGEIFS(' + PRE_T + ',' + PER_R + ',' + PER + '),"")',
     '=IFERROR(AVERAGEIFS(' + POST_T + ',' + PER_R + ',' + PER + '),"")',
     '=IFERROR(C9-B9,"")', '0.0'],
    ['Mean percent', '=IFERROR(B9/' + S + '!$B$8,"")', '=IFERROR(C9/' + S + '!$B$8,"")',
     '=IFERROR(C10-B10,"")', '0%'],
    // Conditional SD without an array formula: SQRT((Sum(x^2) - n*mean^2)/(n-1))
    ['Standard deviation',
     '=IFERROR(SQRT((SUMIFS(' + SQ_PRE + ',' + PER_R + ',' + PER + ')-B8*B9^2)/(B8-1)),"")',
     '=IFERROR(SQRT((SUMIFS(' + SQ_POST + ',' + PER_R + ',' + PER + ')-C8*C9^2)/(C8-1)),"")', '', '0.0'],
    ['Lowest score',
     '=IFERROR(MINIFS(' + PRE_T + ',' + PER_R + ',' + PER + '),"")',
     '=IFERROR(MINIFS(' + POST_T + ',' + PER_R + ',' + PER + '),"")', '', '0'],
    ['Highest score',
     '=IFERROR(MAXIFS(' + PRE_T + ',' + PER_R + ',' + PER + '),"")',
     '=IFERROR(MAXIFS(' + POST_T + ',' + PER_R + ',' + PER + '),"")', '', '0']
  ];
  stats.forEach(function (row, i) {
    var r = 8 + i;
    sh.getRange(r, 1).setValue(row[0]).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2, 1, 3).setValues([[row[1], row[2], row[3]]])
      .setNumberFormat(row[4]).setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true).setFontFamily('Arial').setFontSize(10);
  });

  sh.getRange(15, 1).setValue("Class effect size (Cohen's d)").setFontWeight('bold');
  sh.getRange(15, 2).setFormula('=IFERROR((C9-B9)/SQRT((B11^2+C11^2)/2),"")')
    .setNumberFormat('0.00').setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false);
  sh.getRange(15, 3).setFormula('=IFERROR(IF(B15="","",IF(B15>=0.7,"Strong year",' +
    'IF(B15>=0.4,"Solid year",IF(B15>=0.2,"Modest","Below expectation")))),"")')
    .setFontStyle('italic').setFontSize(9);
  sh.getRange(15, 4).setValue('Middle-grades ELA benchmarks: 0.40 solid, 0.70 strong')
    .setFontStyle('italic').setFontSize(9).setFontColor(COLORS.muted);

  sectionBar_(sh, 17, 'Achievement level distribution', 7);
  headerRow_(sh, 18, 1, ['Achievement level', 'Pre — n', 'Pre — %', 'Post — n', 'Post — %', 'Change in n']);
  LEVELS.forEach(function (lv, i) {
    var r = 19 + i;
    sh.getRange(r, 1).setValue(lv).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2, 1, 5).setValues([[
      '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$F$' + FIRST_ROW + ':$F$' + LAST_ROW + ',"' + lv + '")',
      '=IFERROR(B' + r + '/SUM($B$19:$B$22),"")',
      '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$I$' + FIRST_ROW + ':$I$' + LAST_ROW + ',"' + lv + '")',
      '=IFERROR(D' + r + '/SUM($D$19:$D$22),"")',
      '=D' + r + '-B' + r
    ]]).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
    sh.getRange(r, 3).setNumberFormat('0%');
    sh.getRange(r, 5).setNumberFormat('0%');
  });

  sectionBar_(sh, 24, 'Reporting category performance', 7);
  headerRow_(sh, 25, 1, ['Reporting category', 'Max pts', 'Pre — mean %', 'Post — mean %',
                         'Change', "Students 'Not Yet' (post)", 'Whole-class reteach?']);
  CATEGORIES.forEach(function (c, i) {
    var r = 26 + i;
    var pc = REP + '!$' + PRE_CAT_COLS[i] + '$' + FIRST_ROW + ':$' + PRE_CAT_COLS[i] + '$' + LAST_ROW;
    var oc = REP + '!$' + POST_CAT_COLS[i] + '$' + FIRST_ROW + ':$' + POST_CAT_COLS[i] + '$' + LAST_ROW;
    sh.getRange(r, 1).setValue(c.code + ' — ' + c.name).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2, 1, 6).setValues([[
      c.max,
      '=IFERROR(AVERAGEIFS(' + pc + ',' + PER_R + ',' + PER + ',' + pc + ',">=0"),"")',
      '=IFERROR(AVERAGEIFS(' + oc + ',' + PER_R + ',' + PER + ',' + oc + ',">=0"),"")',
      '=IFERROR(D' + r + '-C' + r + ',"")',
      '=COUNTIFS(' + PER_R + ',' + PER + ',' + oc + ',"<"&' + S + '!$B$12,' + oc + ',">=0")',
      '=IF(C' + r + '="","",IF(F' + r + '>=0.4*MAX(1,COUNTIFS(' + PER_R + ',' + PER + ',' + POST_T + ',">=0")),' +
        '"YES — reteach to whole class","No — small group"))'
    ]]).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true)
      .setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 3, 1, 3).setNumberFormat('0%');
    sh.getRange(r, 7).setWrap(true);
  });
  addRule_(sh, [sh.getRange(26, 7, CATEGORIES.length, 1)], '=LEFT($G26,3)="YES"', COLORS.red, '#9C0006');

  sectionBar_(sh, 34, 'Growth summary', 7);
  var growth = [
    ['Students with both pre and post scores',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$J$' + FIRST_ROW + ':$J$' + LAST_ROW + ',">=-999")', '0'],
    ['Students meeting the growth target',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$N$' + FIRST_ROW + ':$N$' + LAST_ROW + ',"Yes")', '0'],
    ['Percent meeting the growth target', '=IFERROR(B36/B35,"")', '0%'],
    ['Students with statistically reliable change',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$M$' + FIRST_ROW + ':$M$' + LAST_ROW + ',"Yes")', '0'],
    ['Students who declined (negative gain)',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$J$' + FIRST_ROW + ':$J$' + LAST_ROW + ',"<0")', '0'],
    ['Students within one SEM of a cut score (post)',
     '=COUNTIFS(' + PER_R + ',' + PER + ',' + REP + '!$' + colA1(R_NEAR) + '$' + FIRST_ROW + ':$' + colA1(R_NEAR) + '$' + LAST_ROW + ',"NEAR CUT")', '0']
  ];
  growth.forEach(function (g, i) {
    var r = 35 + i;
    sh.getRange(r, 1).setValue(g[0]).setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2).setFormula(g[1]).setNumberFormat(g[2]).setFontWeight('bold')
      .setHorizontalAlignment('center').setBorder(true, true, true, true, false, false);
  });

  sh.getRange(42, 1, 2, 7).merge().setValue(
    'Read the growth summary in this order: (1) effect size for the class, (2) percent meeting the growth target, ' +
    '(3) students with reliable change. A student below the reliable-change threshold has not demonstrated individually ' +
    'measurable growth — say so plainly rather than reporting the raw gain as if it were.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COLORS.muted).setWrap(true).setVerticalAlignment('top');
}

// ---- Item Analysis --------------------------------------------------------

function buildItemAnalysis_(ss) {
  var sh = sheetNamed_(ss, SHEETS.items);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Item Analysis — difficulty, discrimination, and reliability', 14);
  subtitle_(sh, 2, 'p-value = mean score ÷ maximum points. Discrimination is the correlation between the item score and the student’s total. ' +
    'An item below 0.20 is not distinguishing stronger from weaker readers and should be reviewed before the next administration.', 14);

  headerRow_(sh, HDR_ROW, 1, ['Item', 'Max', 'Category', 'Standard', 'Expectation', 'DOK', 'Trait',
    'Pre mean', 'Pre p-value', 'Post mean', 'Post p-value', 'Change in p', 'Discrimination (pre)', 'Flag']);
  [70, 45, 70, 105, 175, 45, 45, 70, 80, 70, 80, 75, 95, 230]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var PRE = "'" + SHEETS.pre + "'", POST = "'" + SHEETS.post + "'";
  var preTot = PRE + '!$' + CL_TOTAL + '$' + FIRST_ROW + ':$' + CL_TOTAL + '$' + LAST_ROW;
  var rows = [], helper = [];

  ITEMS.forEach(function (it, i) {
    var r = FIRST_ROW + i;
    var L = colA1(FIRST_ITEM_COL + i);
    var preRng = PRE + '!$' + L + '$' + FIRST_ROW + ':$' + L + '$' + LAST_ROW;
    var postRng = POST + '!$' + L + '$' + FIRST_ROW + ':$' + L + '$' + LAST_ROW;
    rows.push([it.label, it.max, it.cat, it.std, it.exp, it.dok, it.trait,
      '=IFERROR(AVERAGE(' + preRng + '),"")',
      '=IFERROR(H' + r + '/$B' + r + ',"")',
      '=IFERROR(AVERAGE(' + postRng + '),"")',
      '=IFERROR(J' + r + '/$B' + r + ',"")',
      '=IFERROR(K' + r + '-I' + r + ',"")',
      '=IFERROR(CORREL(' + preRng + ',' + preTot + '),"")',
      '=IF(I' + r + '="","",IF(I' + r + '<0.25,"Very difficult — check key and instruction",' +
        'IF(I' + r + '>0.9,"Very easy — little information",' +
        'IF(AND(M' + r + '<>"",M' + r + '<0.2),"LOW DISCRIMINATION — review item","OK"))))']);
    helper.push(['=IFERROR(VARP(' + preRng + '),0)']);
  });

  sh.getRange(FIRST_ROW, 1, rows.length, 14).setValues(rows)
    .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  sh.getRange(FIRST_ROW, 5, rows.length, 1).setHorizontalAlignment('left').setWrap(true);
  sh.getRange(FIRST_ROW, 14, rows.length, 1).setHorizontalAlignment('left').setWrap(true);
  sh.getRange(FIRST_ROW, 8, rows.length, 1).setNumberFormat('0.00');
  sh.getRange(FIRST_ROW, 9, rows.length, 1).setNumberFormat('0%');
  sh.getRange(FIRST_ROW, 10, rows.length, 1).setNumberFormat('0.00');
  sh.getRange(FIRST_ROW, 11, rows.length, 2).setNumberFormat('0%');
  sh.getRange(FIRST_ROW, 13, rows.length, 1).setNumberFormat('0.00');

  sh.getRange(HDR_ROW, 16).setValue('Item variance (helper)').setFontSize(8).setFontStyle('italic');
  sh.getRange(FIRST_ROW, 16, helper.length, 1).setFormulas(helper).setNumberFormat('0.000');
  sh.hideColumns(16);

  addRule_(sh, [sh.getRange(FIRST_ROW, 14, rows.length, 1)],
    '=LEFT($N' + FIRST_ROW + ',18)="LOW DISCRIMINATION"', COLORS.red, '#9C0006');

  var relRow = FIRST_ROW + N_ITEMS + 2;   // 49
  sectionBar_(sh, relRow, "Reliability (Cronbach's alpha)", 14);
  var lastItemRow = FIRST_ROW + N_ITEMS - 1;
  var alpha = [
    ['Number of items (k)', '=' + N_ITEMS, '0'],
    ['Sum of item variances', '=IFERROR(SUM($P$' + FIRST_ROW + ':$P$' + lastItemRow + '),"")', '0.000'],
    ['Variance of total scores', '=IFERROR(VARP(' + preTot + '),"")', '0.000'],
    ["Cronbach's alpha (pre-test)",
     '=IFERROR((' + N_ITEMS + '/(' + N_ITEMS + '-1))*(1-B' + (relRow + 2) + '/B' + (relRow + 3) + '),"")', '0.000']
  ];
  alpha.forEach(function (a, i) {
    var r = relRow + 1 + i;
    sh.getRange(r, 1).setValue(a[0]).setFontWeight(i === 3 ? 'bold' : 'normal').setFontFamily('Arial').setFontSize(10);
    sh.getRange(r, 2).setFormula(a[1]).setNumberFormat(a[2]).setHorizontalAlignment('center')
      .setFontWeight(i === 3 ? 'bold' : 'normal').setBorder(true, true, true, true, false, false);
  });
  sh.getRange(relRow + 6, 1, 2, 14).merge().setValue(
    'Expect alpha between 0.82 and 0.88 for a 41-item mixed-format instrument. Below 0.70, category-level reporting is not yet ' +
    'trustworthy for individual students — use it as a class signal only, and copy the computed alpha into Settings cell B16 so ' +
    'the growth statistics use your real value instead of the 0.85 default.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COLORS.muted).setWrap(true).setVerticalAlignment('top');

  sh.setFrozenRows(HDR_ROW);
}

// ---- Intervention Groups --------------------------------------------------

function buildGroups_(ss) {
  var sh = sheetNamed_(ss, SHEETS.groups);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Intervention Groups', 16);
  subtitle_(sh, 2, 'Group rosters build themselves from the Student Report and update as scores change. ' +
    'Green = Mastered (80%+), amber = Approaching (60–79%), red = Not Yet (below 60%).', 16);

  var REP = "'" + SHEETS.report + "'", S = SHEETS.settings;
  var NAMES = REP + '!$B$' + FIRST_ROW + ':$B$' + LAST_ROW;
  var WEAK = REP + '!$' + colA1(R_WEAK1) + '$' + FIRST_ROW + ':$' + colA1(R_WEAK1) + '$' + LAST_ROW;

  sectionBar_(sh, 4, 'Whole-class picture — how many students are "Not Yet" in each category?', 16);
  headerRow_(sh, 5, 1, ['Category', 'Not Yet', 'Approaching', 'Mastered', 'Recommendation']);
  [80, 75, 90, 75, 620].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  CATEGORIES.forEach(function (c, i) {
    var r = 6 + i;
    var oc = REP + '!$' + POST_CAT_COLS[i] + '$' + FIRST_ROW + ':$' + POST_CAT_COLS[i] + '$' + LAST_ROW;
    var pc = REP + '!$' + PRE_CAT_COLS[i] + '$' + FIRST_ROW + ':$' + PRE_CAT_COLS[i] + '$' + LAST_ROW;
    sh.getRange(r, 1, 1, 5).setValues([[
      c.code,
      '=COUNTIFS(' + oc + ',"<"&' + S + '!$B$12,' + oc + ',">=0")+COUNTIFS(' + oc + ',"",' + pc + ',"<"&' + S + '!$B$12,' + pc + ',">=0")',
      '=COUNTIFS(' + oc + ',">="&' + S + '!$B$12,' + oc + ',"<"&' + S + '!$B$11)+COUNTIFS(' + oc + ',"",' + pc + ',">="&' + S + '!$B$12,' + pc + ',"<"&' + S + '!$B$11)',
      '=COUNTIFS(' + oc + ',">="&' + S + '!$B$11)+COUNTIFS(' + oc + ',"",' + pc + ',">="&' + S + '!$B$11)',
      '=IFERROR(IF(B' + r + '+C' + r + '+D' + r + '=0,"",IF(B' + r + '/(B' + r + '+C' + r + '+D' + r + ')>=0.4,' +
        '"WHOLE CLASS: "&' + S + '!$E$' + (SET_CAT_FIRST + i) + ',' +
        'IF(B' + r + '=0,"Maintain through spiral review",' +
        '"SMALL GROUP ("&B' + r + '&" students): "&' + S + '!$E$' + (SET_CAT_FIRST + i) + '))),"")'
    ]]).setFontFamily('Arial').setFontSize(9).setBorder(true, true, true, true, true, true);
    sh.getRange(r, 1, 1, 4).setHorizontalAlignment('center');
    sh.getRange(r, 1).setFontWeight('bold');
    sh.getRange(r, 5).setWrap(true).setVerticalAlignment('top');
  });
  addRule_(sh, [sh.getRange(6, 5, CATEGORIES.length, 1)], '=LEFT($E6,11)="WHOLE CLASS"', COLORS.red, '#9C0006');

  sectionBar_(sh, 14, 'Group rosters — built automatically from each student’s weakest category', 16);
  headerRow_(sh, 15, 1, ['Group', 'n', 'Students']);
  CATEGORIES.forEach(function (c, i) {
    var r = 16 + i;
    sh.getRange(r, 1, 1, 3).setValues([[
      c.code,
      '=COUNTIF(' + WEAK + ',$A' + r + ')',
      '=IFERROR(IF(COUNTIF(' + WEAK + ',$A' + r + ')=0,"—",TEXTJOIN(", ",TRUE,FILTER(' + NAMES + ',' + WEAK + '=$A' + r + '))),"—")'
    ]]).setFontFamily('Arial').setFontSize(9).setBorder(true, true, true, true, true, true);
    sh.getRange(r, 1, 1, 2).setHorizontalAlignment('center');
    sh.getRange(r, 1).setFontWeight('bold');
    sh.getRange(r, 3).setWrap(true).setVerticalAlignment('top');
  });
  sh.getRange(23, 1, 1, 16).merge().setValue(
    'Do not run more than three groups at once. A student appears in the group for their WEAKEST category; ' +
    'the student table below shows the full profile, including a second group if you have capacity.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COLORS.muted).setWrap(true);

  sectionBar_(sh, 25, 'Student-level mastery heat map', 16);
  var codes = CATEGORIES.map(function (c) { return c.code; });
  headerRow_(sh, 26, 1, ['Student ID', 'Student Name', 'Period', 'Post total', 'Achievement level',
                         'PRIMARY GROUP', 'Secondary group', 'Not Yet count'].concat(codes));

  var first = 27, block = [];
  for (var k = 0; k < N_STUDENTS; k++) {
    var r = first + k, sr = FIRST_ROW + k, row = [];
    row.push('=' + REP + '!$A' + sr, '=' + REP + '!$B' + sr, '=' + REP + '!$C' + sr,
             '=' + REP + '!$G' + sr, '=' + REP + '!$I' + sr,
             '=' + REP + '!$' + colA1(R_WEAK1) + sr, '=' + REP + '!$' + colA1(R_WEAK2) + sr,
             '=IF($B' + r + '="","",COUNTIF(I' + r + ':O' + r + ',"Not Yet"))');
    for (var i = 0; i < 7; i++) {
      var oc = REP + '!$' + POST_CAT_COLS[i] + sr, pc = REP + '!$' + PRE_CAT_COLS[i] + sr;
      row.push('=IFERROR(IF(' + oc + '<>"",IF(' + oc + '>=' + SHEETS.settings + '!$B$11,"Mastered",' +
        'IF(' + oc + '>=' + SHEETS.settings + '!$B$12,"Approaching","Not Yet")),' +
        'IF(' + pc + '<>"",IF(' + pc + '>=' + SHEETS.settings + '!$B$11,"Mastered",' +
        'IF(' + pc + '>=' + SHEETS.settings + '!$B$12,"Approaching","Not Yet")),"")),"")');
    }
    block.push(row);
  }
  sh.getRange(first, 1, N_STUDENTS, 15).setFormulas(block)
    .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  sh.getRange(first, 6, N_STUDENTS, 1).setFontWeight('bold');
  [85, 155, 60, 75, 135, 100, 95, 75].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  for (var c2 = 9; c2 <= 15; c2++) sh.setColumnWidth(c2, 85);

  var heat = sh.getRange(first, 9, N_STUDENTS, 7);
  addRule_(sh, [heat], '=I' + first + '="Mastered"', COLORS.green);
  addRule_(sh, [heat], '=I' + first + '="Approaching"', COLORS.amber);
  addRule_(sh, [heat], '=I' + first + '="Not Yet"', COLORS.red);
  addRule_(sh, [sh.getRange(first, 8, N_STUDENTS, 1)], '=AND(ISNUMBER($H' + first + '),$H' + first + '>=4)',
    COLORS.red, '#9C0006');

  sh.setFrozenRows(26);
  sh.setFrozenColumns(2);
}

// ---- Standards Coverage ---------------------------------------------------

function buildCoverage_(ss) {
  var sh = sheetNamed_(ss, SHEETS.coverage);
  sh.setHiddenGridlines(true);
  titleBar_(sh, 'Standards Coverage — what this instrument measures, and what it does not', 4);
  subtitle_(sh, 2, 'Read the second table before planning from this data. Roughly a quarter of the Grade 7 expectations cannot be ' +
    'measured in a written booklet and must be assessed through classroom performance tasks.', 4);
  [110, 470, 130, 90].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var assessed = [
    ['7.L.GC.1', 'Grammar, usage & mechanics (semicolons in a complex series)', '35', 'C-LNG'],
    ['7.L.GC.2.a', 'Apply syntax understanding to comprehend and analyze texts', '21', 'I-LNG'],
    ['7.L.GC.2.b', 'Sentence variety; consistent verb tense', '32, essay T3', 'C-LNG'],
    ['7.L.GC.2.c', 'Distinguish active and passive voice; revise for active voice', '33, essay T3', 'C-LNG'],
    ['7.L.GC.2.d', 'Avoid misplaced and dangling modifiers', '34, essay T3', 'C-LNG'],
    ['7.L.V.2.a', 'Greek and Latin roots and affixes', '14', 'I-LNG'],
    ['7.L.V.2.b', 'Parts of speech as a clue to meaning', '15', 'I-LNG'],
    ['7.L.V.2.d', 'Use parts of speech to choose precise words when writing', '36', 'C-LNG'],
    ['7.L.V.3.b', 'Word relationships and context clues beyond the sentence', '3', 'I-LNG'],
    ['7.L.V.3.c', 'Connotations of words sharing a denotation', '19', 'I-LNG'],
    ['7.L.V.1.b', 'Use grade-level academic vocabulary to communicate precisely', 'essay T3', 'C-LNG'],
    ['7.T.C.1.a', 'Multiple purposes within a text and the audiences they target', '16', 'I-CSS'],
    ['7.T.C.1.c', 'Construct texts for a specific purpose and audience', 'essay T1', 'C-CSS'],
    ['7.T.C.2.a', 'Determine the prevailing perspective in a text', '20A', 'I-CSS'],
    ['7.T.C.2.b', 'How evidence and tone reveal perspective and affect credibility', '20B', 'I-CSS'],
    ['7.T.C.2.d', 'Use credible sources to research answers to questions', 'essay T2', 'C-TEC'],
    ['7.T.SS.1.a', 'How authors modify organizational structures to achieve purposes', '11', 'I-CSS'],
    ['7.T.SS.1.b', 'Design texts employing structures and features', 'essay T1', 'C-CSS'],
    ['7.T.SS.1.c', 'Varied transitions connecting ideas, sentences, and paragraphs', '30, essay T1', 'C-CSS'],
    ['7.T.SS.1.d', 'Multi-paragraph texts: introduction, support, conclusion', '25, 38, essay T1', 'C-CSS'],
    ['7.T.SS.2.a', 'How figurative and connotative language shape meaning, mood, tone', '4', 'I-CSS'],
    ['7.T.SS.2.c', 'Situational use of formal or informal style', '31, essay T1', 'C-CSS'],
    ['7.T.T.1.a', 'Narrative techniques developing plot, character, setting', '1', 'I-TEC'],
    ['7.T.T.1.b', 'Plot structure, conflict, and narrative devices', '2', 'I-TEC'],
    ['7.T.T.1.c', 'How themes are developed and expressed across texts', '5, 8', 'I-TEC'],
    ['7.T.T.1.d', 'Compare a fictional portrayal with an account of the same material', '10', 'I-PRA'],
    ['7.T.T.2.a', 'Expository techniques: main ideas, facts, statistics, text features', '12, 13', 'I-TEC'],
    ['7.T.T.2.b', 'How two authors on one topic emphasize different evidence', '22', 'I-TEC'],
    ['7.T.T.2.d', 'Apply expository techniques; elaborate on reasons', '28, essay T2', 'C-TEC'],
    ['7.T.T.3.a', 'Argumentative techniques: claim, evidence, counterclaim, conclusion', '17, 18', 'I-TEC'],
    ['7.T.T.3.c', 'Apply argumentative techniques when writing', '26, 27, 29, essay T2', 'C-TEC'],
    ['7.T.T.4.a', 'Poetic techniques: stanzas, rhyme, imagery, sound devices', '6, 7', 'I-TEC'],
    ['7.T.PM.1.a', 'Analyze a myth a modern writer adapted; features of style and theme', '9, 10', 'I-PRA'],
    ['7.T.RA.2.a', 'Locate evidence; record standard bibliographic information', '24', 'I-PRA'],
    ['7.T.RA.2.b', 'Analyze sources for credibility and relevance', '23', 'I-PRA'],
    ['7.T.RA.2.c', 'Follow a standard citation format when integrating evidence', '27, 37, essay T2', 'C-TEC']
  ];
  sectionBar_(sh, 4, 'Expectations assessed by this instrument', 4);
  headerRow_(sh, 5, 1, ['Expectation', 'Skill', 'Items', 'Category']);
  sh.getRange(6, 1, assessed.length, 4).setValues(assessed)
    .setFontFamily('Arial').setFontSize(9).setBorder(true, true, true, true, true, true);
  sh.getRange(6, 2, assessed.length, 1).setWrap(true);
  sh.getRange(6, 1, assessed.length, 1).setHorizontalAlignment('center');
  sh.getRange(6, 3, assessed.length, 2).setHorizontalAlignment('center');

  var gapRow = 6 + assessed.length + 2;
  var gaps = [
    ['7.T.C.1.b', 'Use text mode features across disciplinary texts',
     'Partially covered by item 13 (data table). Full coverage needs multimodal and digital texts — assess in a science or social studies reading task.'],
    ['7.T.SS.2.b', 'Use figurative language for intentional effect when writing',
     'Narrative or poetic writing task. The essay on this instrument is argumentative.'],
    ['7.T.T.1.e', 'Apply narrative techniques to enhance writing', 'Narrative writing performance task.'],
    ['7.T.T.4.b', 'Apply poetic techniques to produce poetry', 'Poetry writing performance task.'],
    ['7.T.RA.1.a-c', 'Generate research questions; conduct research; analyze from research',
     'Multi-day research unit with a process rubric — cannot be captured in a single sitting.'],
    ['7.T.PM.1.b', 'Analyze a genre of literature from a particular time period',
     'Extended text study with a culminating analysis.'],
    ['7.L.V.1.a', 'Acquire vocabulary through print, digital, and multimodal texts',
     'Ongoing vocabulary assessment across the year.'],
    ['7.L.V.3.d / e', 'Use print and digital reference tools to clarify meaning',
     'Requires tool access, which conflicts with secure testing. Assess during writing conferences.'],
    ['7.L.V.2.c', 'Construct words from Greek and Latin roots and affixes',
     'Assessed only indirectly through essay word choice. Use a morphology quiz.'],
    ['K-12.P.CP.1', 'Collaboration', 'Structured discussion protocol with an observation rubric.'],
    ['K-12.P.CP.2', 'Presentation', 'Oral presentation with a presentation rubric.']
  ];
  sectionBar_(sh, gapRow, 'Expectations NOT assessed — and how to assess them instead', 4);
  headerRow_(sh, gapRow + 1, 1, ['Expectation', 'Skill', 'How to assess it instead']);
  sh.getRange(gapRow + 2, 1, gaps.length, 3).setValues(gaps)
    .setFontFamily('Arial').setFontSize(9).setBorder(true, true, true, true, true, true);
  sh.getRange(gapRow + 2, 2, gaps.length, 2).setWrap(true).setVerticalAlignment('top');
  sh.getRange(gapRow + 2, 1, gaps.length, 1).setHorizontalAlignment('center');
  sh.setColumnWidth(3, 470);
}
