# Grade 7 ELA Pre/Post Assessment

A blueprint-faithful pre-test and post-test for 7th grade English Language Arts, with parallel forms, full answer keys, the Georgia Milestones three-trait writing rubric, and a results reporting workbook that turns item scores into instructional groups.

Built against five source documents, treated as the sole authority for content, weighting, and language:

- Georgia's K–12 English Language Arts Standards, **Grade 7** (Approved May 2023)
- Georgia Milestones **Grade 7 ELA Assessment Blueprint** (August 2025)
- **Draft Grade 7 ELA Achievement Level Descriptors** (October 2025, effective 2025–2026)
- Georgia Milestones **Grade 7 Eight-Point, Three-Trait Writing Rubric** (August 2026)
- **Grade 7 Classroom Peer Revision Guidance** (February 2026)

---

## What's here

| | |
|---|---|
| **Two 60-point forms** | Form A (pre) and Form B (post), item-for-item parallel |
| **39 items per form** | 24 one-point · 14 two-point technology-enhanced · 1 extended writing task |
| **Blueprint match** | Every reporting category carries the exact point weight the operational Milestones uses |
| **Peer Revision Task** | Part 3 reproduces the Milestones PRT, built from the state's own sample item stems |
| **Extended writing** | Scored on the state's 8-point, 3-trait rubric, with anchors and a double-scoring protocol |
| **Reporting workbook** | Item-level entry → category profiles, achievement levels, growth statistics, item analysis, and auto-generated intervention groups |

---

## Website

The project is published as a static site for GitHub Pages, built to the repository root so Pages can serve it from **main / (root)**:

```bash
pip install markdown
python tools/build_site.py     # regenerates index.html, guides/, forms/, assets/
```

`index.html` is a teacher-facing hub — start-here steps, both booklets, test structure, blueprint weights, and every guide. `guides/` holds the rendered documentation and `forms/` the printable booklets. `.nojekyll` tells Pages to serve the files as they are.

**Answer keys are deliberately excluded from the site.** They stay in `assessments/` in the repository, because the pre/post design depends on Form A's keys staying unseen until Form B is administered. To publish them anyway, add them to `PAGES` in `tools/build_site.py`.

Re-run the builder after editing any Markdown, or the site and the sources drift apart.

---

## Quick start

**1. Print the booklets.**

```bash
pip install markdown
python tools/build_print_booklets.py
```

Open `print/form-a-pre-test.html` in a browser and print to PDF — Letter, default margins, background graphics ON, headers/footers OFF. (Or print the Markdown in `assessments/` directly from any editor.)

**2. Administer Form A** in the first three weeks of school. Two 60-minute sessions. Scripts, timing, and accommodations are in `docs/03-Administration-Guide.md`.

**3. Score it** with `assessments/Form-A-Answer-Key.md` and `rubrics/Writing-Rubric-and-Scoring-Guide.md`.

**4. Enter item-level scores** in `reporting/GA7-ELA-PrePost-Reporting-Workbook.xlsx`. Item-level, not totals — the totals tell you nothing you can teach from.

**5. Read `Intervention Groups`,** sort by column F, and teach.

**6. Repeat with Form B** three to four weeks before the Milestones window.

To rebuild the workbook from source:

```bash
pip install openpyxl
python reporting/build_reporting_workbook.py                    # 150-row roster
python reporting/build_reporting_workbook.py my-class.xlsx 30   # smaller roster
```

### Verifying the workbook

Every computed cell is a formula, so the numbers are only as good as the formulas. `reporting/verify_workbook.py` builds a small copy, fills it with synthetic scores whose correct answers are known in advance, evaluates every formula with a pure-Python Excel engine, and compares cell by cell:

```bash
pip install openpyxl formulas
python reporting/verify_workbook.py
```

It checks totals, achievement-level placement, all seven category percentages against blueprint point values, weakest-category identification, growth statistics and their divide-by-zero and decline guards, class means and standard deviations, Cohen's *d*, level distributions, period filtering, item p-values, item variances, discrimination correlations, Cronbach's alpha, blank-row handling, and that no cell anywhere evaluates to an Excel error. Run it after any edit to the builder.

> **Note on cached values.** The workbook ships without cached formula results — Excel, LibreOffice, and Google Sheets all compute them on open, so this is invisible in normal use. It only matters if you read the file programmatically (e.g. with `pandas`), where formula cells will read as empty until the file has been opened and saved by a spreadsheet application once.

### Google Sheets version

`google-apps-script/` holds a Google-native build with the same categories, cut scores, and statistics — plus three things the spreadsheet alone can't do:

- **Auto-scoring from a Google Form answer sheet**, applying the exact partial-credit rules from the paper keys (Forms' built-in quiz grading can't — it scores checkbox items all-or-nothing, which would zero every student who got one of two right on a "select TWO" item)
- **Intervention group rosters that build themselves** and update as scores change
- **One-click printable student profiles** as a Google Doc, in language a student or caregiver can read

Setup is one paste: `google-apps-script/ELA-Reporting-ALL-IN-ONE.gs` bundles all seven script files into one, so there's a single thing to copy into Apps Script. Steps are in `google-apps-script/README.md`. If you're testing on paper and just want the reports, uploading the `.xlsx` to Drive works too — that README names the four cells to check after import.

The Form deliberately contains **no passages** — students read the printed booklet and use the Form as a bubble sheet. Scrolling passages inside a Form would depress reading scores for reasons unrelated to reading, and Form A needs to stay secure until Form B is given.

---

## Test structure

| Part | Content | Items | Points | Time |
|---|---|---|---:|---|
| **1** | Modern short story adapting a classical myth, a poem on the same myth, and a boxed summary of the source myth | 1–10 | 14 | 30 min |
| **2** | Paired informational texts — an expository article with a data table and an editorial on the same topic, plus a Works Cited list | 11–24 | 20 | 30 min |
| **3** | **Peer Revision Task** — a student draft with numbered sentences and standards-targeted planted flaws | 25–38 | 18 | 30 min |
| **4** | **Extended Writing Task** — argumentative essay using both Part 2 sources | 39 | 8 | 30 min |

| Form | Myth adapted | Informational topic |
|---|---|---|
| **A (Pre)** | Icarus — *Wax Wings* / *Flight Lessons* | Light pollution and a dark-sky ordinance |
| **B (Post)** | Arachne — *Thread Count* / *The Loom* | Electronic waste and repairable devices |

All passages are original works written for this assessment.

---

## Blueprint fidelity

Point weights replicate the Georgia Milestones Grade 7 ELA blueprint exactly.

| Reporting category | Points | % of test |
|---|---:|---:|
| Interpreting — Context, Structure & Style | 6 | 10.0% |
| Interpreting — Techniques | 16 | 26.7% |
| Interpreting — Periods & Movements + Research | 6 | 10.0% |
| Interpreting — Language (Grammar & Vocabulary) | 6 | 10.0% |
| **Interpreting Texts** | **34** | **56.7%** |
| Constructing — Context, Structure & Style | 8 | 13.3% |
| Constructing — Techniques | 10 | 16.7% |
| Constructing — Language (Grammar & Vocabulary) | 8 | 13.3% |
| **Constructing Texts** | **26** | **43.3%** |
| **Total** | **60** | |

Depth of Knowledge falls inside every blueprint band: **DOK 1** 3 pts (5%, band 3–10%) · **DOK 2** 29 pts (48%, band 45–60%) · **DOK 3** 28 pts (47%, band 35–50%).

**One deliberate departure.** The operational Milestones assesses Constructing Texts entirely through selected-response and technology-enhanced items, with no extended writing. This instrument converts 8 of those 26 points into an essay scored on the state's three-trait rubric — because the rubric is marked *"intended for classroom use,"* and because a pre/post instrument that never asks a student to write cannot measure sustained composition. The three traits map onto the three Constructing Texts categories, so the blueprint weighting survives the substitution to the point. Full reasoning in `docs/01-Assessment-Design-and-Technical-Notes.md` §2.

---

## The reporting workbook

`reporting/GA7-ELA-PrePost-Reporting-Workbook.xlsx` — ten sheets, every computed cell a live formula.

| Sheet | What it gives you |
|---|---|
| **START HERE** | Instructions, cell-color legend, and the three cautions to read before reporting anything |
| **Settings** | Cut scores, mastery thresholds, growth target, reliability — all editable, everything downstream reads from here |
| **Roster** | Students, class period, form order, accommodations, subgroup flags |
| **Pre-Test Entry / Post-Test Entry** | Item-level score entry with category and max-point rows, entry validation, and out-of-range flagging |
| **Student Report** | Per student: totals, percents, achievement levels, raw gain, percent-of-possible gain, Reliable Change Index, seven category percentages pre and post, category change, two weakest categories, near-cut flag, and a recommended instructional focus |
| **Class Dashboard** | Filterable by class period: means, SDs, level distributions, category performance, Cohen's *d*, and a whole-class-reteach flag per category |
| **Item Analysis** | p-values pre and post, discrimination, change in difficulty, Cronbach's alpha, and flags for items that aren't working |
| **Intervention Groups** | Mastery heat map plus each student's primary and secondary group — sort column F and teach |
| **Standards Coverage** | Every Grade 7 expectation the instrument measures, and every one it does not, with how to assess those instead |

### Growth is reported four ways, on purpose

1. **Raw gain** — post minus pre.
2. **Percent-of-possible gain** — corrects for ceiling effects on high scorers.
3. **Reliable Change Index** — whether an individual student's gain exceeds measurement error. The workbook computes the threshold from your own data. **This is the number most classroom pre/post reports omit, and it is the one that decides whether a gain is real.**
4. **Cohen's *d*** — class-level effect size, per section and grade-wide.

---

## What this instrument does not do

Stated up front, because the alternative is that someone over-reads the data.

- **It does not predict Georgia Milestones scores.** Cut scores were set by content review, not by a state standard-setting panel. Achievement levels here are instructional triage.
- **Category scores rest on as few as 6 points.** Use them to form groups; confirm them against student work before acting on an individual.
- **It measures no listening, viewing, speaking, or collaboration.** Standards `K–12.P.CP.1` and `K–12.P.CP.2` require performance tasks. So do `7.T.RA.1`, `7.T.PM.1.b`, `7.T.T.1.e`, `7.T.T.4.b`, and `7.T.SS.2.b`. Roughly a quarter of the Grade 7 expectations are outside what a written booklet can reach — the `Standards Coverage` sheet names every one.
- **Form equivalence is by construction, not yet by data.** The forms are content-matched item for item. Establishing statistical equivalence takes one year of local data; `docs/01-Assessment-Design-and-Technical-Notes.md` §5 gives the three checks and the workbook computes them.

---

## Files

```
docs/
  01-Assessment-Design-and-Technical-Notes.md   design rationale, psychometrics, cut scores, limitations
  02-Blueprint-Alignment-Matrix.md              every item → standard, expectation, ALD, DOK, type, Form A↔B pair
  03-Administration-Guide.md                    timing, scripts, accommodations, scoring workflow, first-week use
assessments/
  Form-A-Pre-Test.md            Form-A-Answer-Key.md
  Form-B-Post-Test.md           Form-B-Answer-Key.md
rubrics/
  Writing-Rubric-and-Scoring-Guide.md            3-trait rubric, anchors, condition codes, double-scoring protocol
  Student-Checklist-and-Peer-Revision-Guide.md   student-facing revision tools
reporting/
  build_reporting_workbook.py                    regenerates the workbook
  GA7-ELA-PrePost-Reporting-Workbook.xlsx        the teacher data tool
google-apps-script/
  README.md                                      setup, and which path to take
  ELA-Reporting-ALL-IN-ONE.gs                    ← paste THIS one file into Apps Script
  Config.gs                                      item spec, answer keys, categories, cut scores
  BuildSheets.gs · BuildSheets2.gs               builds all ten sheets natively in Google Sheets
  Scoring.gs                                     auto-scoring engine + 33-assertion self-test
  FormBuilder.gs                                 generates the Google Form answer sheet
  Reports.gs                                     one-page student profiles as a Google Doc
  Menu.gs · appsscript.json                      custom menu and manifest
tools/
  build_print_booklets.py                        Markdown → print-ready HTML booklets
  build_apps_script_bundle.py                    regenerates the all-in-one .gs bundle
print/
  form-a-pre-test.html · form-b-post-test.html   generated; print to PDF
```

**Answer keys are secure documents.** Do not distribute Form A keys or item content to students before Form B has been administered — reviewing Form A items contaminates the growth measure.
