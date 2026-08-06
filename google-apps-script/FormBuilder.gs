/**
 * FormBuilder.gs — generates a Google Form answer sheet for Form A or Form B.
 *
 * Design decision worth knowing: the Form does NOT contain the passages.
 * Students read the printed or PDF booklet and use the Form the way they would
 * use a bubble sheet. Two reasons:
 *
 *   1. Scrolling 1,500 words of passage inside a Google Form, then scrolling
 *      back and forth to each item, is a genuinely worse reading experience
 *      than a booklet — and it would depress reading scores for reasons that
 *      have nothing to do with reading ability.
 *   2. Keeping passages out of the Form keeps the item content out of a
 *      student-accessible document, which matters because Form A must stay
 *      secure until Form B has been administered.
 *
 * The Form collects raw responses only. All scoring happens in Scoring.gs, so
 * the partial-credit rules match the paper answer keys exactly. (Google Forms'
 * built-in quiz grading is all-or-nothing on checkbox items, which would score
 * every "select TWO" item wrong for a student who got one of the two right.)
 */

function createFormA() { createAnswerSheetForm_('A'); }
function createFormB() { createAnswerSheetForm_('B'); }

function createAnswerSheetForm_(formLabel) {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();
  var prop = formLabel === 'A' ? PROP_FORM_A : PROP_FORM_B;

  if (props.getProperty(prop)) {
    var again = ui.alert('Form ' + formLabel + ' already exists',
      'A Form ' + formLabel + ' answer sheet has already been created for this spreadsheet.\n\n' +
      'Create a NEW one? The old Form and its responses will be left untouched, but this ' +
      'spreadsheet will stop scoring from it.', ui.ButtonSet.YES_NO);
    if (again !== ui.Button.YES) return;
  }

  var testName = formLabel === 'A' ? 'Pre-Test' : 'Post-Test';
  var form = FormApp.create('Grade 7 ELA — Form ' + formLabel + ' (' + testName + ') Answer Sheet');
  form.setDescription(
    'Record your answers here. Read the passages and the full answer choices in your test booklet.\n\n' +
    '• Items worth 2 points have a Part A and a Part B, or ask you to select TWO answers.\n' +
    '• Answer every item. A blank earns nothing; a good guess sometimes earns a point.\n' +
    '• Your essay is written on paper and is not part of this form.');
  form.setProgressBar(true);
  form.setAllowResponseEdits(false);
  form.setShuffleQuestions(false);

  var id = form.addTextItem().setTitle('Student ID').setRequired(true);
  id.setHelpText('Type your Student ID exactly as it appears on your booklet. Scoring matches on this.');

  var LETTERS = ['A', 'B', 'C', 'D', 'E'];
  var partCounts = {1: 10, 2: 24, 3: 38};   // last item number in each part
  var partTitles = {
    1: 'Part 1 — Literary and Poetic Texts (items 1–10)',
    2: 'Part 2 — Paired Informational Texts (items 11–24)',
    3: 'Part 3 — Peer Revision Task (items 25–38)'
  };
  var currentPart = 0;

  SCORED_ITEMS.forEach(function (it) {
    var part = it.n <= partCounts[1] ? 1 : (it.n <= partCounts[2] ? 2 : 3);
    if (part !== currentPart) {
      currentPart = part;
      form.addPageBreakItem().setTitle(partTitles[part]);
    }

    var opts = LETTERS.slice(0, it.opts);

    if (it.type === 'SR') {
      form.addMultipleChoiceItem()
        .setTitle(questionTitle_(it))
        .setChoiceValues(opts)
        .setRequired(false);

    } else if (it.type === 'MP') {
      ['A', 'B'].forEach(function (p) {
        form.addMultipleChoiceItem()
          .setTitle(questionTitle_(it, p))
          .setChoiceValues(opts)
          .setRequired(false);
      });

    } else if (it.type === 'MS') {
      form.addCheckboxItem()
        .setTitle(questionTitle_(it))
        .setHelpText('Select exactly TWO. Selecting three or more scores zero.')
        .setChoiceValues(opts)
        .setRequired(false);

    } else if (it.type === 'MATCH') {
      var help = 'Write S1 (Source 1 only), S2 (Source 2 only), or B (both sources) for each statement.';
      ['a', 'b', 'c', 'd'].forEach(function (p, i) {
        var mc = form.addMultipleChoiceItem()
          .setTitle(questionTitle_(it, p))
          .setChoiceValues(['S1', 'S2', 'B'])
          .setRequired(false);
        if (i === 0) mc.setHelpText(help);
      });
    }
  });

  form.addPageBreakItem().setTitle('Part 4 — Extended Writing Task');
  form.addSectionHeaderItem()
    .setTitle('Write your essay on paper')
    .setHelpText('Your essay is scored on the three-trait rubric by your teacher and is not submitted here. ' +
                 'Make sure your name is on your lined pages.');

  props.setProperty(prop, form.getId());
  installFormTrigger_();

  var msg = 'Created: ' + form.getTitle() +
    '\n\nShare this link with students:\n' + form.getPublishedUrl() +
    '\n\nEdit the Form here:\n' + form.getEditUrl() +
    '\n\nResponses will score automatically on submit. You can also run ' +
    'ELA Reporting ▸ Score Form responses now at any time.' +
    '\n\nStudents still need the printed booklet — this Form holds no passages or answer choices.';
  ui.alert('Form ' + formLabel + ' answer sheet created', msg, ui.ButtonSet.OK);
}

/** Install the on-submit trigger once. */
function installFormTrigger_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler';
  });
  if (exists) return;
  ScriptApp.newTrigger('onFormSubmitHandler').forSpreadsheet(ss).onFormSubmit().create();
}

/** Show the links for whichever Forms exist. */
function showFormLinks() {
  var props = PropertiesService.getDocumentProperties();
  var lines = [];
  [['A', PROP_FORM_A], ['B', PROP_FORM_B]].forEach(function (pair) {
    var fid = props.getProperty(pair[1]);
    if (!fid) { lines.push('Form ' + pair[0] + ': not created yet.'); return; }
    try {
      var f = FormApp.openById(fid);
      lines.push('Form ' + pair[0] + ' — ' + f.getTitle() +
        '\n  Student link: ' + f.getPublishedUrl() +
        '\n  Edit link:    ' + f.getEditUrl() +
        '\n  Responses:    ' + f.getResponses().length);
    } catch (e) {
      lines.push('Form ' + pair[0] + ': could not be opened (' + e.message + ')');
    }
  });
  SpreadsheetApp.getUi().alert('Answer-sheet Forms', lines.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
