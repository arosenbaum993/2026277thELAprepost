# Google Sheets / Apps Script Reporting

A Google-native build of the reporting system. Same seven reporting categories, same cut scores, same growth statistics as the Excel workbook — plus three things the spreadsheet alone cannot do:

- **Auto-scoring** from a Google Form answer sheet, with the exact partial-credit rules from the paper answer keys
- **Group rosters that build themselves** and update as scores change
- **Printable one-page student profiles** generated as a Google Doc, written in language a student or caregiver can read

---

## Which path should you take?

| | Upload the `.xlsx` to Drive | Apps Script build |
|---|---|---|
| Setup time | 2 minutes | ~10 minutes, once |
| Students test on paper | ✅ | ✅ |
| Students answer on Chromebooks | ❌ | ✅ auto-scored on submit |
| You hand-enter 38 item scores per student | ✅ (~25 min/class) | Only the 3 essay trait scores |
| Intervention groups | Sort a column yourself | Rosters write themselves |
| Student handouts | Build your own | One click → Google Doc |
| Shared with a co-teacher | ✅ | ✅ |

**Take the upload path if** you're testing on paper this year and just want the reports. It genuinely works — see the note at the bottom about the four cells to check.

**Take the Apps Script path if** you have Chromebooks and want the item entry to disappear. Hand-entering 38 scores × 130 students is about two hours per administration; the Form removes almost all of it.

You can start with the upload and move to Apps Script later without losing anything.

---

## Setup (Apps Script path)

### 👉 The easy way — one file

**Use [`ELA-Reporting-ALL-IN-ONE.gs`](ELA-Reporting-ALL-IN-ONE.gs).** It is all seven script files concatenated into one, so there is a single thing to paste.

**1. Create the spreadsheet.** In Drive: **New ▸ Google Sheets**. Name it something like `Grade 7 ELA Pre-Post Reporting`.

**2. Open the script editor.** **Extensions ▸ Apps Script**.

**3. Paste the script.** Select everything in the default `Code.gs` and delete it, then paste the entire contents of `ELA-Reporting-ALL-IN-ONE.gs` in its place. Save (💾).

**4. Add the manifest.** Click ⚙️ **Project Settings** and tick *Show "appsscript.json" manifest file in editor*. Go back to the editor, open **appsscript.json**, and replace its contents with the file of that name in this folder.

> `appsscript.json` is JSON, not script — it just declares which permissions the project needs. It cannot go in the same file as the code.

**5. Save, then reload the spreadsheet tab.** An **ELA Reporting** menu appears next to Help.

**6. Run `ELA Reporting ▸ 1. Set up workbook`.**

Google asks you to authorize on first run. If you see an **unverified app** screen, click *Advanced ▸ Go to (project name)* — it is your own script, and the permissions are limited to this spreadsheet, Google Forms, and Google Docs.

Ten sheets appear. Start on **START HERE**.

**7. Optional sanity check:** `ELA Reporting ▸ Checks ▸ Run scoring self-test`. It should report 33 passed, 0 failed.

### The modular way

If you'd rather keep the code in separate files — easier to edit later — click **+ ▸ Script** for each one and name it exactly as shown, then delete the default `Code.gs`:

| Create a file named | Paste from |
|---|---|
| `Config` | `Config.gs` |
| `BuildSheets` | `BuildSheets.gs` |
| `BuildSheets2` | `BuildSheets2.gs` |
| `Scoring` | `Scoring.gs` |
| `FormBuilder` | `FormBuilder.gs` |
| `Reports` | `Reports.gs` |
| `Menu` | `Menu.gs` |

**File order matters.** `Config` must sit first in the file list — Apps Script evaluates files in order, and the others read values it defines. Drag it to the top if it isn't. The all-in-one file has this order baked in, which is the main reason to prefer it.

Then continue from step 4 above.

> If you edit any individual `.gs` file, regenerate the bundle with `python tools/build_apps_script_bundle.py` so the two never drift apart.

---

## Using it

### Paper testing

1. Fill in the **Roster**.
2. Administer Form A from the printed booklet.
3. Enter item scores on **Pre-Test Entry** — 0 for a wrong answer, never a blank.
4. Enter the three essay trait scores in `Q39_T1`, `Q39_T2`, `Q39_T3`.
5. Read **Student Report**, **Class Dashboard**, **Intervention Groups**.

### Chromebook testing (auto-scored)

1. Fill in the **Roster** — **Student ID is what matching runs on**, so it must be exactly what students will type.
2. `ELA Reporting ▸ 2. Answer-sheet Forms ▸ Create Form A`. You get a student link.
3. **Students still need the printed booklet.** The Form holds no passages and no answer choices — it's a bubble sheet. See "Why no passages in the Form?" below.
4. Students read the booklet and record answers in the Form.
5. Responses score automatically on submit. If a trigger misses, run `ELA Reporting ▸ 3. Score Form responses now`.
6. **Enter the three essay trait scores by hand.** The essay is never auto-scored.

### Student handouts

`ELA Reporting ▸ 4. Generate student reports` → a Google Doc with one page per student: overall level, growth in plain language, a skill-by-skill table, and the two areas being worked on next. Print it and hand it out, or share it for conferences.

You'll be asked for a class period; leave it blank for everyone.

---

## Why no passages in the Form?

Two reasons, both about validity rather than convenience:

1. **Scrolling 1,500 words of passage inside a Google Form, then scrolling back to each item, is a materially worse reading experience than a booklet.** It would depress reading scores for reasons that have nothing to do with reading ability — and if it depressed them at pre-test only, it would manufacture false growth.
2. **Form A has to stay secure until Form B is administered.** Putting the items in a student-accessible document works against that.

The Form also collects raw responses rather than using Google Forms' built-in quiz grading. Forms scores checkbox items all-or-nothing, which would give zero to every student who got one of the two right on a "select TWO" item. All scoring happens in `Scoring.gs`, where the partial-credit rules match the paper answer keys exactly.

---

## What's been verified

Run `ELA Reporting ▸ Checks ▸ Run scoring self-test` any time. It checks 33 assertions:

- every partial-credit rule (selected response, two-part, select TWO including the three-or-more-selections case, and the four-cell match)
- that the seven category point totals still equal the blueprint values (6 / 16 / 6 / 6 / 8 / 10 / 8) and sum to 60
- that both answer keys are complete, correctly shaped for each item type, and share no key letter on any item

Before this was committed, these further checks were run outside Apps Script:

- **The whole build is simulated offline.** `tools/simulate_apps_script_build.js` runs `buildAll()` against a mock of the Sheets API, recording every merge, freeze, and formula write, then asserts the invariants Google enforces at runtime — no merged band split by a frozen row or column, no overlapping merges, no malformed formula arrays. Run it with `node tools/simulate_apps_script_build.js` after any edit to a build function.
- All seven `.gs` files, and the all-in-one bundle, parse as valid JavaScript.
- The bundle was loaded and executed end to end: all globals resolve in concatenation order, `onOpen()` builds its menu, and the self-test passes **from the bundle** — which is what proves the file order is right.
- All 37 top-level functions are declared exactly once in the bundle; concatenation introduced no duplicate or shadowed declarations.
- Every menu item points at a function that exists.
- Every key and point value in `Config.gs` was compared against the published `Form-A-Answer-Key.md` and `Form-B-Answer-Key.md` — all 76 keys match.
- The item spec in `Config.gs` was compared against `reporting/build_reporting_workbook.py` — label, max points, category, and DOK agree on all 41 columns, so the Excel and Sheets paths cannot drift apart.

The formula logic itself is the same logic verified cell by cell against independently computed values in the Excel build (`reporting/verify_workbook.py`).

**Not verified:** the script has not been executed inside Google's real runtime from here — no Google account is available in this environment. Expect to click through the authorization screen on first run, and if a `setFormulas` call ever errors on a locale that uses semicolons as argument separators, that's the first thing to look at.

### Known issue, fixed

An earlier version failed on `buildAll()` with:

> Sorry, you can't freeze columns which contain only part of a merged cell.

Four sheets freeze columns so student names stay visible while you scroll (Roster, both entry sheets, Student Report, Intervention Groups), and each had a full-width merged title bar running straight through the freeze boundary — 14 conflicts in total. Banners are now split at that boundary by `bandMerge_()`, which looks identical because both blocks carry the same fill. The simulation harness above was written to reproduce the failure offline and now guards against it.

---

## Editing it later

**Changing cut scores or thresholds:** don't touch the script. Edit the yellow cells on the **Settings** sheet.

**Changing an item's category or points:** edit `ITEMS` in `Config.gs`, and make the same change in `docs/02-Blueprint-Alignment-Matrix.md` and `reporting/build_reporting_workbook.py`, or the three will disagree. Then re-run **Set up workbook** and re-run the self-test.

**Roster larger than 150:** change `N_STUDENTS` at the top of `Config.gs` and re-run **Set up workbook**.

> ⚠️ **Set up workbook rebuilds all ten sheets and erases entered scores.** It asks first. If you have live data, duplicate the spreadsheet before running it.

---

## If you take the upload path instead

Upload `reporting/GA7-ELA-PrePost-Reporting-Workbook.xlsx` to Drive and open it with Google Sheets. Almost everything converts cleanly — `SUMIF`, `COUNTIFS`, `AVERAGEIFS`, `INDEX`/`MATCH`, `VARP`, and `CORREL` all behave identically.

**Check these four cells afterward.** They use `MINIFS` and `MAXIFS`, which are written with an Excel-internal `_xlfn.` prefix that Sheets may not strip on import:

- `Class Dashboard!B12` and `B13` (lowest / highest pre-test score)
- `Class Dashboard!C12` and `C13` (lowest / highest post-test score)

If any shows `#NAME?`, delete the `_xlfn.` from the formula — `=IFERROR(MINIFS(...))` works natively in Sheets. Nothing else on the sheet depends on those four cells, so the rest of your reports are correct either way.
