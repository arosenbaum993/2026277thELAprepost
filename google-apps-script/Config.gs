/**
 * Config.gs — single source of truth for the Grade 7 ELA pre/post reporting system.
 *
 * Everything downstream (sheet construction, auto-scoring, student reports) reads
 * from this file. If you change an item's category or point value here, change it
 * in docs/02-Blueprint-Alignment-Matrix.md too, or the two will drift apart.
 */

/** Roster capacity. Raise or lower freely; every formula scales with it. */
var N_STUDENTS = 150;

/** Row geometry — identical on every sheet so student N is always the same row. */
var CAT_ROW = 3;      // reporting-category code for each item column
var MAX_ROW = 4;      // maximum points for each item column
var HDR_ROW = 5;      // column headers
var FIRST_ROW = 6;    // first student row
var LAST_ROW = FIRST_ROW + N_STUDENTS - 1;

var FIRST_ITEM_COL = 3;   // column C

/**
 * Item specification. Order defines column order on the entry sheets.
 *   type: SR    selected response, one key
 *         MP    two parts (includes evidence-based selected response) — 1 pt per part
 *         MS    multiple-select "choose TWO"
 *         MATCH four cells matched to S1 / S2 / B
 *   opts: number of answer choices (used by the Form builder)
 */
var ITEMS = [
  {n: 1,  label: 'Q1',  max: 1, cat: 'I-TEC', std: '7.T.T.1',  exp: '7.T.T.1.a',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 2,  label: 'Q2',  max: 2, cat: 'I-TEC', std: '7.T.T.1',  exp: '7.T.T.1.b',            dok: 3, trait: '',  type: 'MP',    opts: 4},
  {n: 3,  label: 'Q3',  max: 1, cat: 'I-LNG', std: '7.L.V.3',  exp: '7.L.V.3.b',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 4,  label: 'Q4',  max: 1, cat: 'I-CSS', std: '7.T.SS.2', exp: '7.T.SS.2.a',           dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 5,  label: 'Q5',  max: 1, cat: 'I-TEC', std: '7.T.T.1',  exp: '7.T.T.1.c',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 6,  label: 'Q6',  max: 1, cat: 'I-TEC', std: '7.T.T.4',  exp: '7.T.T.4.a',            dok: 1, trait: '',  type: 'SR',    opts: 4},
  {n: 7,  label: 'Q7',  max: 2, cat: 'I-TEC', std: '7.T.T.4',  exp: '7.T.T.4.a',            dok: 3, trait: '',  type: 'MP',    opts: 4},
  {n: 8,  label: 'Q8',  max: 2, cat: 'I-TEC', std: '7.T.T.1',  exp: '7.T.T.1.c',            dok: 3, trait: '',  type: 'MS',    opts: 5},
  {n: 9,  label: 'Q9',  max: 1, cat: 'I-PRA', std: '7.T.PM.1', exp: '7.T.PM.1.a',           dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 10, label: 'Q10', max: 2, cat: 'I-PRA', std: '7.T.T.1',  exp: '7.T.T.1.d/7.T.PM.1.a', dok: 3, trait: '',  type: 'MP',    opts: 4},
  {n: 11, label: 'Q11', max: 1, cat: 'I-CSS', std: '7.T.SS.1', exp: '7.T.SS.1.a',           dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 12, label: 'Q12', max: 1, cat: 'I-TEC', std: '7.T.T.2',  exp: '7.T.T.2.a',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 13, label: 'Q13', max: 1, cat: 'I-TEC', std: '7.T.T.2',  exp: '7.T.T.2.a',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 14, label: 'Q14', max: 1, cat: 'I-LNG', std: '7.L.V.2',  exp: '7.L.V.2.a',            dok: 1, trait: '',  type: 'SR',    opts: 4},
  {n: 15, label: 'Q15', max: 1, cat: 'I-LNG', std: '7.L.V.2',  exp: '7.L.V.2.b',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 16, label: 'Q16', max: 2, cat: 'I-CSS', std: '7.T.C.1',  exp: '7.T.C.1.a',            dok: 3, trait: '',  type: 'MP',    opts: 4},
  {n: 17, label: 'Q17', max: 1, cat: 'I-TEC', std: '7.T.T.3',  exp: '7.T.T.3.a',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 18, label: 'Q18', max: 2, cat: 'I-TEC', std: '7.T.T.3',  exp: '7.T.T.3.a',            dok: 3, trait: '',  type: 'MP',    opts: 4},
  {n: 19, label: 'Q19', max: 1, cat: 'I-LNG', std: '7.L.V.3',  exp: '7.L.V.3.c',            dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 20, label: 'Q20', max: 2, cat: 'I-CSS', std: '7.T.C.2',  exp: '7.T.C.2.a/b',          dok: 3, trait: '',  type: 'MP',    opts: 4},
  {n: 21, label: 'Q21', max: 2, cat: 'I-LNG', std: '7.L.GC.2', exp: '7.L.GC.2.a',           dok: 2, trait: '',  type: 'MP',    opts: 4},
  {n: 22, label: 'Q22', max: 2, cat: 'I-TEC', std: '7.T.T.2',  exp: '7.T.T.2.b',            dok: 3, trait: '',  type: 'MATCH', opts: 3},
  {n: 23, label: 'Q23', max: 1, cat: 'I-PRA', std: '7.T.RA.2', exp: '7.T.RA.2.b',           dok: 2, trait: '',  type: 'SR',    opts: 4},
  {n: 24, label: 'Q24', max: 2, cat: 'I-PRA', std: '7.T.RA.2', exp: '7.T.RA.2.a',           dok: 2, trait: '',  type: 'MP',    opts: 4},
  {n: 25, label: 'Q25', max: 1, cat: 'C-CSS', std: '7.T.SS.1', exp: '7.T.SS.1.d',           dok: 2, trait: '1', type: 'SR',    opts: 4},
  {n: 26, label: 'Q26', max: 1, cat: 'C-TEC', std: '7.T.T.3',  exp: '7.T.T.3.c',            dok: 2, trait: '2', type: 'SR',    opts: 4},
  {n: 27, label: 'Q27', max: 2, cat: 'C-TEC', std: '7.T.T.3',  exp: '7.T.T.3.c/7.T.RA.2.c', dok: 3, trait: '2', type: 'MS',    opts: 4},
  {n: 28, label: 'Q28', max: 2, cat: 'C-TEC', std: '7.T.T.2',  exp: '7.T.T.2.d',            dok: 2, trait: '2', type: 'MP',    opts: 4},
  {n: 29, label: 'Q29', max: 1, cat: 'C-TEC', std: '7.T.T.3',  exp: '7.T.T.3.c',            dok: 2, trait: '2', type: 'SR',    opts: 4},
  {n: 30, label: 'Q30', max: 2, cat: 'C-CSS', std: '7.T.SS.1', exp: '7.T.SS.1.c',           dok: 3, trait: '1', type: 'MP',    opts: 4},
  {n: 31, label: 'Q31', max: 1, cat: 'C-CSS', std: '7.T.SS.2', exp: '7.T.SS.2.c',           dok: 2, trait: '1', type: 'SR',    opts: 4},
  {n: 32, label: 'Q32', max: 1, cat: 'C-LNG', std: '7.L.GC.2', exp: '7.L.GC.2.b',           dok: 2, trait: '3', type: 'SR',    opts: 4},
  {n: 33, label: 'Q33', max: 2, cat: 'C-LNG', std: '7.L.GC.2', exp: '7.L.GC.2.c',           dok: 2, trait: '3', type: 'MS',    opts: 5},
  {n: 34, label: 'Q34', max: 1, cat: 'C-LNG', std: '7.L.GC.2', exp: '7.L.GC.2.d',           dok: 2, trait: '3', type: 'SR',    opts: 4},
  {n: 35, label: 'Q35', max: 1, cat: 'C-LNG', std: '7.L.GC.1', exp: 'GUM (Master Gr 6)',    dok: 1, trait: '3', type: 'SR',    opts: 4},
  {n: 36, label: 'Q36', max: 1, cat: 'C-LNG', std: '7.L.V.2',  exp: '7.L.V.2.d',            dok: 2, trait: '3', type: 'SR',    opts: 4},
  {n: 37, label: 'Q37', max: 1, cat: 'C-TEC', std: '7.T.RA.2', exp: '7.T.RA.2.c',           dok: 2, trait: '2', type: 'SR',    opts: 4},
  {n: 38, label: 'Q38', max: 1, cat: 'C-CSS', std: '7.T.SS.1', exp: '7.T.SS.1.d',           dok: 2, trait: '1', type: 'SR',    opts: 4},
  {n: 39, label: 'Q39_T1', max: 3, cat: 'C-CSS', std: 'Essay Trait 1', exp: 'Purpose & Organization',       dok: 3, trait: '1', type: 'ESSAY'},
  {n: 39, label: 'Q39_T2', max: 3, cat: 'C-TEC', std: 'Essay Trait 2', exp: 'Evidence & Elaboration',       dok: 3, trait: '2', type: 'ESSAY'},
  {n: 39, label: 'Q39_T3', max: 2, cat: 'C-LNG', std: 'Essay Trait 3', exp: 'Language Usage & Conventions', dok: 3, trait: '3', type: 'ESSAY'}
];

var N_ITEMS = ITEMS.length;                               // 41 columns
var LAST_ITEM_COL = FIRST_ITEM_COL + N_ITEMS - 1;         // AQ
var COND_COL = LAST_ITEM_COL + 1;                         // AR — essay condition code
var TOTAL_COL = LAST_ITEM_COL + 2;                        // AS — total

/** Scored items only (excludes the three essay traits). Used by the Form builder. */
var SCORED_ITEMS = ITEMS.filter(function (it) { return it.type !== 'ESSAY'; });

/**
 * Answer keys. Verified against assessments/Form-A-Answer-Key.md and
 * assessments/Form-B-Answer-Key.md.
 *   SR    -> 'B'
 *   MP    -> ['PartA', 'PartB']
 *   MS    -> ['A','C']  (the two correct choices)
 *   MATCH -> ['S1','S2','B','B']  (cells a, b, c, d)
 */
var KEYS = {
  A: {
    1: 'B',  2: ['B', 'C'],  3: 'D',  4: 'C',  5: 'A',  6: 'D',  7: ['A', 'B'],
    8: ['A', 'C'],  9: 'A', 10: ['C', 'B'], 11: 'C', 12: 'B', 13: 'D', 14: 'A',
    15: 'C', 16: ['B', 'B'], 17: 'B', 18: ['B', 'B'], 19: 'C', 20: ['B', 'A'],
    21: ['C', 'B'], 22: ['S1', 'S2', 'B', 'B'], 23: 'A', 24: ['B', 'C'],
    25: 'D', 26: 'B', 27: ['A', 'D'], 28: ['C', 'B'], 29: 'A', 30: ['D', 'A'],
    31: 'B', 32: 'C', 33: ['A', 'D'], 34: 'D', 35: 'A', 36: 'B', 37: 'C', 38: 'D'
  },
  B: {
    1: 'A',  2: ['C', 'D'],  3: 'C',  4: 'A',  5: 'B',  6: 'A',  7: ['B', 'C'],
    8: ['A', 'C'],  9: 'C', 10: ['C', 'C'], 11: 'D', 12: 'A', 13: 'B', 14: 'D',
    15: 'B', 16: ['C', 'B'], 17: 'D', 18: ['B', 'C'], 19: 'A', 20: ['C', 'C'],
    21: ['B', 'B'], 22: ['S1', 'S2', 'B', 'B'], 23: 'B', 24: ['C', 'C'],
    25: 'A', 26: 'D', 27: ['A', 'D'], 28: ['D', 'C'], 29: 'C', 30: ['B', 'C'],
    31: 'D', 32: 'B', 33: ['B', 'D'], 34: 'C', 35: 'D', 36: 'C', 37: 'B', 38: 'C'
  }
};

/** Reporting categories — max points replicate the Georgia Milestones blueprint exactly. */
var CATEGORIES = [
  {code: 'I-CSS', name: 'Interpreting — Context, Structure & Style', max: 6,
   focus: "Author's purpose and audience; perspective, tone, and credibility; how structure and figurative language shape meaning",
   start: "Model 'why did the author build it this way?' on short texts. Standards 7.T.C.1.a, 7.T.C.2.a-b, 7.T.SS.1.a, 7.T.SS.2.a."},
  {code: 'I-TEC', name: 'Interpreting — Techniques', max: 16,
   focus: 'Narrative, expository, argumentative, and poetic technique analysis; theme; claim and counterclaim; comparing two authors on one topic',
   start: 'Largest category on the test (16 pts). Split by technique type before grouping — narrative gaps and argumentative gaps need different lessons.'},
  {code: 'I-PRA', name: 'Interpreting — Periods & Movements + Research', max: 6,
   focus: 'Myth and adaptation analysis; comparing fiction to an account of the same material; source credibility, relevance, and bibliographic information',
   start: 'Pair every literary text with its source myth or historical account. Teach credibility as a two-part judgment: credible AND relevant to the question.'},
  {code: 'I-LNG', name: 'Interpreting — Language (Grammar & Vocabulary)', max: 6,
   focus: 'Context clues beyond the sentence; Greek and Latin roots; part of speech as a meaning clue; connotation; reading syntax',
   start: 'Daily morphology routine plus "read the next sentence" modeling for multiple-meaning words. Standards 7.L.V.2, 7.L.V.3, 7.L.GC.2.a.'},
  {code: 'C-CSS', name: 'Constructing — Context, Structure & Style', max: 8,
   focus: 'Introductions that establish a claim; cohesive structure; transitions that name relationships; style for a target audience; conclusions that finalize',
   start: 'Peer Revision Task routine, Trait 1 guided questions. Start with claim-writing and conclusion-writing; transitions after.'},
  {code: 'C-TEC', name: 'Constructing — Techniques', max: 10,
   focus: 'Reasons that support the claim; evidence from MULTIPLE sources; elaboration in the student’s own words; counterclaims; crediting sources',
   start: 'Explicit evidence-to-elaboration instruction. Sentence stems: "This matters because…" Reuse test items 27 and 28 as warm-ups.'},
  {code: 'C-LNG', name: 'Constructing — Language (Grammar & Vocabulary)', max: 8,
   focus: 'Sentence variety; consistent verb tense; active voice; misplaced and dangling modifiers; punctuation; precise word choice',
   start: 'Sentence combining plus targeted GUM work on rows marked M at Grades 6-7. Standards 7.L.GC.2.b-d, 7.L.GC.1.'}
];

var SET_CAT_FIRST = 28;   // Settings row of the first category in the lookup table

var LEVELS = ['Beginning Learner', 'Developing Learner', 'Proficient Learner', 'Distinguished Learner'];

var SHEETS = {
  start: 'START HERE',
  settings: 'Settings',
  roster: 'Roster',
  pre: 'Pre-Test Entry',
  post: 'Post-Test Entry',
  report: 'Student Report',
  dash: 'Class Dashboard',
  items: 'Item Analysis',
  groups: 'Intervention Groups',
  coverage: 'Standards Coverage'
};

var COLORS = {
  navy: '#1F3864', blue: '#2E5C8A', light: '#DCE6F1', band: '#F2F5FA',
  input: '#FFF2CC', grey: '#E8E8E8',
  green: '#C6EFCE', amber: '#FFEB9C', red: '#FFC7CE',
  inputText: '#0000FF', muted: '#595959'
};

/** Convert a 1-based column index to an A1 column letter. */
function colA1(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

var CL_FIRST = colA1(FIRST_ITEM_COL);   // C
var CL_LAST = colA1(LAST_ITEM_COL);     // AQ
var CL_COND = colA1(COND_COL);          // AR
var CL_TOTAL = colA1(TOTAL_COL);        // AS
