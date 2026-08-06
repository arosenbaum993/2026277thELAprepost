#!/usr/bin/env python3
"""
Build the GitHub Pages site for the Grade 7 ELA pre/post assessment.

Renders the project's Markdown into a styled, responsive, print-friendly static
site at the repository root, so GitHub Pages can serve it with the simple
branch source: main / (root).

    python tools/build_site.py

Outputs:
    index.html          hub / landing page
    assets/site.css     shared stylesheet
    guides/*.html       rendered documentation
    forms/*.html        the student booklets (self-contained, print to PDF)
    .nojekyll           serve files as-is, no Jekyll processing

Deliberately NOT published: the two answer-key files. They stay in the
repository for teachers and are excluded from the site, because the pre/post
design depends on Form A's keys staying unseen until Form B is administered.
Add them to PAGES below if you decide otherwise.
"""

import html
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import markdown
except ImportError:
    sys.exit("Missing dependency. Run:  pip install markdown")

ROOT = Path(__file__).resolve().parent.parent

# (source markdown, output path, card label, blurb, short nav label)
PAGES = [
    ("docs/01-Assessment-Design-and-Technical-Notes.md", "guides/design-notes.html",
     "Design & technical notes", "Design",
     "Blueprint math, DOK and ALD calibration, cut scores, parallel-form equivalence, and the limitations stated plainly."),
    ("docs/02-Blueprint-Alignment-Matrix.md", "guides/blueprint-matrix.html",
     "Blueprint alignment matrix", "Blueprint",
     "Every item mapped to standard, expectation, achievement level, DOK, item type, and its Form A/B partner."),
    ("docs/03-Administration-Guide.md", "guides/administration.html",
     "Administration guide", "Administering",
     "Timing, read-aloud scripts, accommodations, the scoring workflow, and how to use the results in week one."),
    ("rubrics/Writing-Rubric-and-Scoring-Guide.md", "guides/writing-rubric.html",
     "Writing rubric & scoring guide", "Rubric",
     "The 8-point, 3-trait rubric with anchors, condition codes, and a double-scoring protocol."),
    ("rubrics/Student-Checklist-and-Peer-Revision-Guide.md", "guides/student-checklist.html",
     "Student checklist & peer revision", "Student tools",
     "Student-facing revision tools built from the state's peer revision guidance. Print double-sided."),
    ("google-apps-script/README.md", "guides/reporting-setup.html",
     "Reporting setup", "Reporting",
     "Google Sheets and Excel reporting: which to choose, how to set it up, and what has been verified."),
]

FORMS = [
    ("print/form-a-pre-test.html", "forms/form-a-pre-test.html", "Form A — Pre-Test",
     "Icarus adaptation · light pollution · 60 points"),
    ("print/form-b-post-test.html", "forms/form-b-post-test.html", "Form B — Post-Test",
     "Arachne adaptation · electronic waste · 60 points"),
]

CSS = """
:root{
  --bg:#ffffff; --ink:#16191f; --muted:#5b6470; --rule:#e2e6eb; --soft:#f6f8fa;
  --accent:#1f3864; --accent-2:#2e5c8a; --chip:#eef3f9; --shadow:0 1px 2px rgba(16,24,40,.05),0 6px 20px rgba(16,24,40,.06);
}
@media (prefers-color-scheme:dark){
  :root{ --bg:#0f1319; --ink:#e6e9ee; --muted:#98a2b3; --rule:#252c37; --soft:#151b23;
         --accent:#8fb3e0; --accent-2:#a9c6ea; --chip:#1a2230; --shadow:none; }
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{
  margin:0; background:var(--bg); color:var(--ink);
  font:16px/1.65 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
}
.wrap{max-width:860px;margin:0 auto;padding:0 20px}
a{color:var(--accent-2);text-decoration:none}
a:hover{text-decoration:underline}

/* ---- masthead ---- */
.mast{border-bottom:1px solid var(--rule);background:var(--soft)}
.mast .wrap{padding-top:34px;padding-bottom:30px}
.eyebrow{font-size:12px;letter-spacing:.13em;text-transform:uppercase;color:var(--muted);margin:0 0 8px}
.mast h1{margin:0 0 10px;font-size:clamp(26px,4.6vw,40px);line-height:1.15;letter-spacing:-.02em}
.mast p{margin:0;color:var(--muted);max-width:62ch;font-size:16.5px}
.chips{margin-top:16px;display:flex;flex-wrap:wrap;gap:7px}
.chip{background:var(--chip);border:1px solid var(--rule);border-radius:999px;
      padding:4px 11px;font-size:12.5px;color:var(--muted);white-space:nowrap}

/* ---- nav ---- */
.nav{position:sticky;top:0;z-index:5;background:var(--bg);border-bottom:1px solid var(--rule)}
.nav .wrap{display:flex;gap:18px;overflow-x:auto;padding-top:11px;padding-bottom:11px;
           font-size:14px;scrollbar-width:none}
.nav .wrap::-webkit-scrollbar{display:none}
.nav a{color:var(--muted);white-space:nowrap}
.nav a:hover,.nav a.on{color:var(--ink);text-decoration:none}

main{padding:36px 0 72px}
section{margin:0 0 40px}
h2{font-size:21px;letter-spacing:-.01em;margin:0 0 6px}
h2+.lede{margin:0 0 16px;color:var(--muted);font-size:15px}

/* ---- cards ---- */
.cards{display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(255px,1fr))}
.cards.steps{grid-template-columns:repeat(auto-fit,minmax(178px,1fr))}
.cards.steps .card h3{font-size:14.5px}
.cards.steps .card p{font-size:13.2px}
.card{display:block;border:1px solid var(--rule);border-radius:11px;padding:16px 17px;
      background:var(--bg);box-shadow:var(--shadow);transition:border-color .15s,transform .15s}
.card:hover{border-color:var(--accent-2);text-decoration:none;transform:translateY(-1px)}
.card h3{margin:0 0 5px;font-size:15.5px;color:var(--ink)}
.card p{margin:0;font-size:13.7px;color:var(--muted);line-height:1.55}
.card .meta{margin-top:9px;font-size:12px;color:var(--accent-2);letter-spacing:.02em}

/* ---- callout ---- */
.note{border:1px solid var(--rule);border-left:3px solid var(--accent-2);border-radius:8px;
      background:var(--soft);padding:14px 16px;font-size:14.3px;color:var(--muted)}
.note strong{color:var(--ink)}

/* ---- data table on the hub ---- */
.stat{width:100%;border-collapse:collapse;font-size:14px;margin-top:4px}
.stat th,.stat td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--rule)}
.stat th{font-weight:600;color:var(--muted);font-size:12.5px;text-transform:uppercase;letter-spacing:.05em}
.stat td:last-child,.stat th:last-child{text-align:right;font-variant-numeric:tabular-nums}

/* ---- rendered markdown ---- */
.doc h1{font-size:clamp(24px,4vw,34px);line-height:1.2;letter-spacing:-.02em;margin:0 0 18px;
        padding-bottom:12px;border-bottom:1px solid var(--rule)}
.doc h2{font-size:20px;margin:34px 0 10px;padding-bottom:6px;border-bottom:1px solid var(--rule)}
.doc h3{font-size:16.5px;margin:24px 0 8px}
.doc p,.doc li{font-size:15.6px}
.doc ul,.doc ol{padding-left:22px}
.doc li{margin:4px 0}
.doc blockquote{margin:16px 0;padding:12px 16px;background:var(--soft);
                border-left:3px solid var(--accent-2);border-radius:0 8px 8px 0}
.doc blockquote p:last-child{margin-bottom:0}
.doc code{background:var(--soft);border:1px solid var(--rule);border-radius:4px;
          padding:1px 5px;font-size:13.4px;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}
.doc pre{background:var(--soft);border:1px solid var(--rule);border-radius:8px;
         padding:13px 15px;overflow-x:auto}
.doc pre code{background:none;border:none;padding:0;font-size:13.2px;line-height:1.55}
.doc hr{border:none;border-top:1px solid var(--rule);margin:28px 0}
.table-scroll{overflow-x:auto;margin:14px 0;border:1px solid var(--rule);border-radius:8px}
.doc table{border-collapse:collapse;width:100%;font-size:13.8px;min-width:520px}
.doc th,.doc td{border-bottom:1px solid var(--rule);padding:8px 11px;text-align:left;vertical-align:top}
.doc thead th{background:var(--soft);font-weight:600;white-space:nowrap}
.doc tbody tr:last-child td{border-bottom:none}

.backlink{display:inline-block;margin-bottom:20px;font-size:14px;color:var(--muted)}
.backlink:hover{color:var(--ink);text-decoration:none}

footer{border-top:1px solid var(--rule);padding:26px 0 44px;font-size:13.2px;color:var(--muted)}
footer p{margin:0 0 6px}

@media print{
  .nav,.backlink,footer,.mast{display:none}
  body{font-size:11pt} .wrap{max-width:none;padding:0}
  .doc a{color:inherit;text-decoration:none}
  .table-scroll{overflow:visible;border:none}
}
"""

NAV = [("index.html", "Overview")] + [(p[1], p[3]) for p in PAGES]


def rel(from_path: str, to_path: str) -> str:
    """Relative href from one output page to another."""
    depth = from_path.count("/")
    return ("../" * depth) + to_path


def nav_html(current: str) -> str:
    items = []
    for href, label in NAV:
        cls = ' class="on"' if href == current else ""
        items.append(f'<a href="{rel(current, href)}"{cls}>{html.escape(label)}</a>')
    return f'<nav class="nav"><div class="wrap">{"".join(items)}</div></nav>'


def shell(current: str, title: str, body: str, mast: str = "") -> str:
    css = rel(current, "assets/site.css")
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title>
<meta name="description" content="A blueprint-faithful Grade 7 ELA pre/post assessment aligned to Georgia's K-12 ELA Standards and the Georgia Milestones Grade 7 ELA blueprint.">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="{css}">
</head>
<body>
{mast}
{nav_html(current)}
<main><div class="wrap">
{body}
</div></main>
<footer><div class="wrap">
<p><strong>Grade 7 English Language Arts — Pre/Post Assessment.</strong> Aligned to Georgia's K–12 ELA Standards (Grade 7, May 2023), the Georgia Milestones Grade 7 ELA Assessment Blueprint (August 2025), the draft Grade 7 Achievement Level Descriptors (October 2025), the Georgia Milestones Grade 7 Eight-Point Three-Trait Writing Rubric (August 2026), and the Grade 7 Classroom Peer Revision Guidance (February 2026).</p>
<p>All passages are original works written for this assessment. Achievement levels are set by content review, not by a state standard-setting panel, and do not predict a Georgia Milestones score.</p>
</div></footer>
</body>
</html>
"""


def render_markdown(md_text: str) -> str:
    md = markdown.Markdown(extensions=["tables", "attr_list", "sane_lists", "fenced_code"])
    out = md.convert(md_text)
    # Make wide tables scroll rather than blow out the layout on phones.
    out = re.sub(r"<table>", '<div class="table-scroll"><table>', out)
    out = re.sub(r"</table>", "</table></div>", out)
    return out


def build() -> None:
    (ROOT / "assets").mkdir(exist_ok=True)
    (ROOT / "guides").mkdir(exist_ok=True)
    (ROOT / "forms").mkdir(exist_ok=True)
    (ROOT / "assets" / "site.css").write_text(CSS.strip() + "\n", encoding="utf-8")
    (ROOT / ".nojekyll").write_text("", encoding="utf-8")

    made = []

    # --- documentation pages ---
    for src, out, label, short, _blurb in PAGES:
        src_path = ROOT / src
        if not src_path.exists():
            print(f"  ! missing {src}")
            continue
        body = render_markdown(src_path.read_text(encoding="utf-8"))
        page = shell(out, f"{label} · Grade 7 ELA Pre/Post",
                     f'<a class="backlink" href="{rel(out, "index.html")}">← Back to overview</a>'
                     f'<article class="doc">{body}</article>')
        (ROOT / out).write_text(page, encoding="utf-8")
        made.append(out)

    # --- student booklets (already self-contained HTML) ---
    for src, out, *_ in FORMS:
        src_path = ROOT / src
        if not src_path.exists():
            print(f"  ! missing {src} — run tools/build_print_booklets.py first")
            continue
        shutil.copyfile(src_path, ROOT / out)
        made.append(out)

    # --- hub ---
    mast = """<header class="mast"><div class="wrap">
<p class="eyebrow">Georgia Milestones aligned · Grade 7</p>
<h1>English Language Arts<br>Pre/Post Assessment</h1>
<p>Two parallel 60-point forms, full answer keys, the state's three-trait writing rubric, and a reporting system that turns item scores into instructional groups.</p>
<div class="chips">
<span class="chip">Blueprint-exact weighting</span>
<span class="chip">Item-for-item parallel forms</span>
<span class="chip">Peer Revision Task</span>
<span class="chip">3-trait scored writing</span>
<span class="chip">Auto-scoring &amp; growth reporting</span>
</div>
</div></header>"""

    form_cards = "".join(
        f'<a class="card" href="{out}"><h3>{html.escape(label)}</h3>'
        f'<p>{html.escape(blurb)}</p>'
        f'<p class="meta">Open booklet → print to PDF</p></a>'
        for _src, out, label, blurb in FORMS
    )

    guide_cards = "".join(
        f'<a class="card" href="{out}"><h3>{html.escape(label)}</h3><p>{html.escape(blurb)}</p></a>'
        for _src, out, label, _short, blurb in PAGES
    )

    body = f"""
<section>
  <h2>Start here</h2>
  <p class="lede">Four steps from booklet to instructional groups.</p>
  <div class="cards steps">
    <div class="card"><h3>1 · Print the booklets</h3><p>Open a form below and print to PDF — Letter, default margins, background graphics on.</p></div>
    <div class="card"><h3>2 · Administer Form A</h3><p>First three weeks of school. Two 60-minute sessions. Scripts and accommodations are in the administration guide.</p></div>
    <div class="card"><h3>3 · Score and enter</h3><p>Item-level, not totals. Enter 0 for a wrong answer — a blank means "not tested".</p></div>
    <div class="card"><h3>4 · Read the groups</h3><p>The reporting workbook names each student's weakest reporting category and builds the groups for you.</p></div>
  </div>
</section>

<section>
  <h2>The forms</h2>
  <p class="lede">Item <em>n</em> on both forms shares standard, expectation, achievement level, DOK, item type, and point value. Only the passages and answer positions differ — that is what makes a pre&nbsp;→&nbsp;post difference mean growth rather than a difference in difficulty.</p>
  <div class="cards">{form_cards}</div>
  <p style="margin-top:14px"><span class="note" style="display:block"><strong>Answer keys are not published here.</strong> They live in the repository under <code>assessments/</code>, because the pre/post design depends on Form&nbsp;A's keys staying unseen until Form&nbsp;B has been administered.</span></p>
</section>

<section>
  <h2>Test structure</h2>
  <p class="lede">60 points per form, 39 items.</p>
  <table class="stat">
    <thead><tr><th>Part</th><th>Content</th><th>Points</th></tr></thead>
    <tbody>
      <tr><td>1</td><td>Short story adapting a classical myth, a poem, and the source myth</td><td>14</td></tr>
      <tr><td>2</td><td>Expository article with a data table, an editorial, and a Works Cited list</td><td>20</td></tr>
      <tr><td>3</td><td>Peer Revision Task — a student draft with targeted planted flaws</td><td>18</td></tr>
      <tr><td>4</td><td>Extended writing — argumentative essay using both sources</td><td>8</td></tr>
    </tbody>
  </table>
</section>

<section>
  <h2>Blueprint fidelity</h2>
  <p class="lede">Reporting-category weights replicate the operational Georgia Milestones blueprint exactly.</p>
  <table class="stat">
    <thead><tr><th>Reporting category</th><th>Points</th></tr></thead>
    <tbody>
      <tr><td>Interpreting — Context, Structure &amp; Style</td><td>6</td></tr>
      <tr><td>Interpreting — Techniques</td><td>16</td></tr>
      <tr><td>Interpreting — Periods &amp; Movements + Research</td><td>6</td></tr>
      <tr><td>Interpreting — Language</td><td>6</td></tr>
      <tr><td>Constructing — Context, Structure &amp; Style</td><td>8</td></tr>
      <tr><td>Constructing — Techniques</td><td>10</td></tr>
      <tr><td>Constructing — Language</td><td>8</td></tr>
      <tr><td><strong>Total</strong></td><td><strong>60</strong></td></tr>
    </tbody>
  </table>
  <p class="lede" style="margin-top:12px">Depth of Knowledge falls inside every blueprint band: DOK&nbsp;1 at 5% (band 3–10%), DOK&nbsp;2 at 48% (band 45–60%), DOK&nbsp;3 at 47% (band 35–50%).</p>
</section>

<section>
  <h2>Documentation</h2>
  <p class="lede">Everything needed to administer, score, and defend the instrument.</p>
  <div class="cards">{guide_cards}</div>
</section>

<section>
  <h2>What this instrument does not do</h2>
  <div class="note">
    <p style="margin-top:0"><strong>It does not predict Georgia Milestones scores.</strong> Cut scores were set by content review, not by a state standard-setting panel. Achievement levels here are instructional triage.</p>
    <p><strong>Category scores rest on as few as 6 points.</strong> Use them to form groups; confirm against student work before acting on an individual.</p>
    <p><strong>It measures no listening, speaking, collaboration, or presentation.</strong> Roughly a quarter of the Grade 7 expectations need classroom performance tasks. The reporting workbook names every one.</p>
    <p style="margin-bottom:0"><strong>Form equivalence is by construction, not yet by data.</strong> Three local checks establish it statistically; the workbook computes them.</p>
  </div>
</section>
"""
    (ROOT / "index.html").write_text(
        shell("index.html", "Grade 7 ELA Pre/Post Assessment", body, mast), encoding="utf-8")
    made.insert(0, "index.html")

    # --- custom 404 -------------------------------------------------------
    nf_body = """
<section style="padding-top:8px">
  <h2>Page not found</h2>
  <p class="lede">That address doesn't match anything on this site.</p>
  <div class="cards" style="margin-top:18px">
    <a class="card" href="index.html"><h3>Overview</h3><p>Start here — the hub with every form and guide.</p></a>
    <a class="card" href="forms/form-a-pre-test.html"><h3>Form A — Pre-Test</h3><p>Open the booklet and print to PDF.</p></a>
    <a class="card" href="forms/form-b-post-test.html"><h3>Form B — Post-Test</h3><p>Open the booklet and print to PDF.</p></a>
    <a class="card" href="guides/administration.html"><h3>Administration guide</h3><p>Timing, scripts, accommodations, scoring.</p></a>
  </div>
  <p class="lede" style="margin-top:22px">If you followed a link from somewhere else, note that addresses here are
  case-sensitive.</p>
</section>
"""
    nf_mast = """<header class="mast"><div class="wrap">
<p class="eyebrow">Grade 7 English Language Arts</p>
<h1>404</h1>
<p>The page you asked for isn't here.</p>
</div></header>"""
    (ROOT / "404.html").write_text(
        shell("index.html", "Page not found · Grade 7 ELA Pre/Post", nf_body, nf_mast),
        encoding="utf-8")
    made.append("404.html")

    print(f"Built {len(made)} pages:")
    for m in made:
        size = (ROOT / m).stat().st_size
        print(f"  {m:<38} {size // 1024:>4} KB")
    print("  assets/site.css")
    print("  .nojekyll")


if __name__ == "__main__":
    # Make sure the booklets exist before copying them.
    subprocess.run([sys.executable, str(ROOT / "tools" / "build_print_booklets.py")],
                   check=True, stdout=subprocess.DEVNULL)
    build()
