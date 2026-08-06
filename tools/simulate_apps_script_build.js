/**
 * simulate_apps_script_build.js
 *
 * Runs the Apps Script bundle's buildAll() against a mock of the Google Sheets
 * API, recording every merge, freeze, formula and range operation, then asserts
 * the invariants Google enforces at runtime.
 *
 * This exists because of a real failure: Google rejects
 *
 *     "you can't freeze columns which contain only part of a merged cell"
 *
 * and nothing in a syntax check or a scoring unit test can catch it — the
 * conflict only appears when a full-width merged title bar meets a frozen
 * column. This harness reproduces that class of error offline.
 *
 * Usage:  node tools/simulate_apps_script_build.js [path/to/bundle.gs]
 */

const fs = require('fs');
const path = require('path');

const BUNDLE = process.argv[2] ||
  path.join(__dirname, '..', 'google-apps-script', 'ELA-Reporting-ALL-IN-ONE.gs');

// ---------------------------------------------------------------- mock sheets

const sheets = new Map();
const problems = [];
let opCount = 0;

function chainable(obj) {
  return new Proxy(obj, {
    get(target, prop) {
      if (prop in target) return target[prop];
      // Any unimplemented setter is a no-op that keeps the chain alive.
      return () => { opCount++; return chainable(target); };
    }
  });
}

function makeRange(sheet, row, col, numRows, numCols) {
  const r = {
    _sheet: sheet, _row: row, _col: col, _rows: numRows, _cols: numCols,
    merge() {
      opCount++;
      sheet.merges.push({ r1: row, r2: row + numRows - 1, c1: col, c2: col + numCols - 1 });
      return chainable(r);
    },
    setValue(v) { opCount++; return chainable(r); },
    setValues(v) {
      opCount++;
      if (!Array.isArray(v) || !Array.isArray(v[0])) {
        problems.push(`${sheet.name}: setValues at R${row}C${col} was not a 2-D array`);
      } else if (v.length !== numRows || v[0].length !== numCols) {
        problems.push(`${sheet.name}: setValues at R${row}C${col} shape ${v.length}x${v[0].length} ` +
                      `but range is ${numRows}x${numCols}`);
      }
      return chainable(r);
    },
    setFormula(f) {
      opCount++;
      if (typeof f !== 'string' || f[0] !== '=') {
        problems.push(`${sheet.name}: setFormula at R${row}C${col} does not start with "=": ${String(f).slice(0, 40)}`);
      }
      return chainable(r);
    },
    setFormulas(f) {
      opCount++;
      if (!Array.isArray(f) || !Array.isArray(f[0])) {
        problems.push(`${sheet.name}: setFormulas at R${row}C${col} was not a 2-D array`);
        return chainable(r);
      }
      if (f.length !== numRows || f[0].length !== numCols) {
        problems.push(`${sheet.name}: setFormulas at R${row}C${col} shape ${f.length}x${f[0].length} ` +
                      `but range is ${numRows}x${numCols}`);
      }
      f.forEach((rowArr, i) => rowArr.forEach((cell, j) => {
        if (typeof cell !== 'string' || cell[0] !== '=') {
          problems.push(`${sheet.name}: non-formula in setFormulas at ` +
                        `R${row + i}C${col + j}: ${JSON.stringify(String(cell).slice(0, 40))}`);
        }
      }));
      return chainable(r);
    },
    getNumCells: () => numRows * numCols,
  };
  return chainable(r);
}

function makeSheet(name) {
  const s = {
    name,
    merges: [],
    frozenRows: 0,
    frozenColumns: 0,
    cfRules: [],
    getRange(a, b, c, d) {
      opCount++;
      if (typeof a === 'string') {
        // A1 notation used in a couple of places (e.g. 'Z1')
        const m = /^([A-Z]+)(\d+)$/.exec(a);
        if (!m) { problems.push(`${name}: unparsed A1 range ${a}`); return makeRange(s, 1, 1, 1, 1); }
        let col = 0;
        for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
        return makeRange(s, Number(m[2]), col, 1, 1);
      }
      return makeRange(s, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
    },
    clear() { return s; },
    clearConditionalFormatRules() { s.cfRules = []; return s; },
    getConditionalFormatRules() { return s.cfRules; },
    setConditionalFormatRules(rules) { s.cfRules = rules; return s; },
    setFrozenRows(n) { opCount++; s.frozenRows = n; return s; },
    setFrozenColumns(n) { opCount++; s.frozenColumns = n; return s; },
    setRowHeight() { return s; },
    setColumnWidth() { return s; },
    setHiddenGridlines() { return s; },
    hideColumns() { return s; },
    getName: () => name,
  };
  return chainable(s);
}

const spreadsheet = chainable({
  getSheetByName: (n) => sheets.get(n) || null,
  insertSheet(n) { const sh = makeSheet(n); sheets.set(n, sh); return sh; },
  getSheets: () => Array.from(sheets.values()),
  setActiveSheet: (sh) => sh,
  moveActiveSheet: () => {},
  deleteSheet(sh) { sheets.delete(sh.getName()); },
  getId: () => 'mock-id',
});

const ui = chainable({
  alert: () => 'YES',
  prompt: () => chainable({ getSelectedButton: () => 'OK', getResponseText: () => '' }),
  createMenu: () => chainable({ addItem: () => ui._menu, addSeparator: () => ui._menu, addSubMenu: () => ui._menu, addToUi: () => {} }),
  ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
  Button: { YES: 'YES', OK: 'OK' },
});
ui._menu = ui;

global.SpreadsheetApp = chainable({
  getActiveSpreadsheet: () => spreadsheet,
  getUi: () => ui,
  newConditionalFormatRule: () => {
    const rule = chainable({
      whenFormulaSatisfied: () => rule, setBackground: () => rule, setFontColor: () => rule,
      setRanges: () => rule, setGradientMinpointWithValue: () => rule,
      setGradientMidpointWithValue: () => rule, setGradientMaxpointWithValue: () => rule,
      build: () => ({}),
    });
    return rule;
  },
  newDataValidation: () => {
    const dv = chainable({ requireValueInList: () => dv, build: () => ({}) });
    return dv;
  },
  InterpolationType: { NUMBER: 'NUMBER' },
});
global.PropertiesService = chainable({ getDocumentProperties: () => chainable({ getProperty: () => null, setProperty: () => {} }) });
global.FormApp = chainable({});
global.ScriptApp = chainable({ getProjectTriggers: () => [] });
global.DocumentApp = chainable({ ParagraphHeading: { HEADING1: 1, HEADING2: 2 } });
global.Utilities = chainable({ formatDate: () => '2026-01-01' });
global.Session = chainable({ getScriptTimeZone: () => 'America/New_York' });
global.Logger = chainable({ log: () => {} });

// ---------------------------------------------------------------- run build

const src = fs.readFileSync(BUNDLE, 'utf8');
eval(src);

console.log('Running buildAll() against the mock Sheets API…\n');
try {
  buildAll();
} catch (e) {
  problems.push(`buildAll() threw: ${e.message}`);
}

// ---------------------------------------------------------------- assertions

console.log(`Sheets created: ${sheets.size}`);
console.log(`Range operations recorded: ${opCount}\n`);

let frozenReport = [];
for (const [name, sh] of sheets) {
  const merges = sh.merges.length;
  frozenReport.push(
    `  ${name.padEnd(22)} merges=${String(merges).padStart(3)}  ` +
    `frozenRows=${sh.frozenRows}  frozenCols=${sh.frozenColumns}`);

  // Google's rule: a frozen region must not contain only PART of a merged cell.
  for (const m of sh.merges) {
    if (sh.frozenColumns > 0 && m.c1 <= sh.frozenColumns && m.c2 > sh.frozenColumns) {
      problems.push(
        `${name}: merge R${m.r1}C${m.c1}:R${m.r2}C${m.c2} is split by frozenColumns=${sh.frozenColumns} ` +
        `-> "you can't freeze columns which contain only part of a merged cell"`);
    }
    if (sh.frozenRows > 0 && m.r1 <= sh.frozenRows && m.r2 > sh.frozenRows) {
      problems.push(
        `${name}: merge R${m.r1}C${m.c1}:R${m.r2}C${m.c2} is split by frozenRows=${sh.frozenRows} ` +
        `-> "you can't freeze rows which contain only part of a merged cell"`);
    }
  }

  // Overlapping merges are also a runtime error in Sheets.
  for (let i = 0; i < sh.merges.length; i++) {
    for (let j = i + 1; j < sh.merges.length; j++) {
      const a = sh.merges[i], b = sh.merges[j];
      const overlap = a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
      if (overlap) {
        problems.push(`${name}: overlapping merges R${a.r1}C${a.c1}:R${a.r2}C${a.c2} and ` +
                      `R${b.r1}C${b.c1}:R${b.r2}C${b.c2}`);
      }
    }
  }
}
console.log(frozenReport.join('\n'));

console.log('\n' + '='.repeat(72));
if (problems.length) {
  console.log(`FAILED — ${problems.length} problem(s):\n`);
  problems.forEach(p => console.log('  • ' + p));
  process.exit(1);
}
console.log('PASSED — no merge/freeze conflicts, no malformed formula writes.');
