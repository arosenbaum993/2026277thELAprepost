#!/usr/bin/env python3
"""
Render the Markdown test booklets into print-ready HTML.

Produces clean, single-column, letter-sized pages with:
  - a running header and page numbers
  - page breaks between parts
  - passages and item blocks kept off page boundaries where possible
  - answer bubbles that print legibly in black and white
  - lined pages appended for the extended writing task

Usage:
    python tools/build_print_booklets.py                 # both forms
    python tools/build_print_booklets.py assessments/Form-A-Pre-Test.md

Open the resulting .html in a browser and print to PDF (Letter, margins
"Default", background graphics ON).

Requires: markdown  (pip install markdown)
"""

import re
import sys
from pathlib import Path

try:
    import markdown
except ImportError:
    sys.exit("Missing dependency. Run:  pip install markdown")

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SOURCES = [
    ROOT / "assessments" / "Form-A-Pre-Test.md",
    ROOT / "assessments" / "Form-B-Post-Test.md",
]
OUTDIR = ROOT / "print"

LINED_PAGES = 3
LINES_PER_PAGE = 26

CSS = """
@page { size: letter; margin: 0.7in 0.75in 0.8in 0.75in; }

:root {
  --ink: #14181f;
  --muted: #5b6470;
  --rule: #c3c9d2;
  --tint: #f2f5f9;
}

* { box-sizing: border-box; }

body {
  font-family: "Georgia", "Cambria", serif;
  font-size: 11.5pt;
  line-height: 1.5;
  color: var(--ink);
  max-width: 7in;
  margin: 0 auto;
  padding: 0.4in 0;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}

h1 {
  font-family: "Helvetica Neue", Arial, sans-serif;
  font-size: 19pt;
  letter-spacing: -0.01em;
  margin: 0 0 2pt;
  padding-bottom: 6pt;
  border-bottom: 2.5pt solid var(--ink);
}
h1 + h1 { border-bottom: none; padding-bottom: 0; margin-top: -4pt; }

h1.pagebreak, h1#part-1, h1#part-2, h1#part-3, h1#part-4 { page-break-before: always; }

h2 {
  font-family: "Helvetica Neue", Arial, sans-serif;
  font-size: 14pt;
  margin: 20pt 0 8pt;
  padding-bottom: 3pt;
  border-bottom: 1pt solid var(--rule);
  page-break-after: avoid;
}

h3 {
  font-family: "Helvetica Neue", Arial, sans-serif;
  font-size: 11.5pt;
  margin: 14pt 0 5pt;
  page-break-after: avoid;
}

p { margin: 0 0 8pt; orphans: 3; widows: 3; }

strong { font-weight: 700; }
em { font-style: italic; }

hr {
  border: none;
  border-top: 1pt solid var(--rule);
  margin: 14pt 0;
}

/* Item blocks: a numbered item and its options stay together */
.item {
  page-break-inside: avoid;
  margin: 0 0 13pt;
  padding-left: 2pt;
}
.item > p:first-child { font-weight: 600; }

ul { margin: 4pt 0 8pt; padding-left: 22pt; list-style: none; }
li { margin: 0 0 3pt; text-indent: -22pt; padding-left: 22pt; }

table {
  border-collapse: collapse;
  width: 100%;
  margin: 8pt 0 12pt;
  font-family: "Helvetica Neue", Arial, sans-serif;
  font-size: 9.5pt;
  page-break-inside: avoid;
}
th, td {
  border: 0.75pt solid var(--rule);
  padding: 4pt 6pt;
  text-align: left;
  vertical-align: top;
}
th { background: var(--tint); font-weight: 700; }

blockquote {
  margin: 10pt 0;
  padding: 9pt 12pt;
  background: var(--tint);
  border-left: 3pt solid var(--ink);
  page-break-inside: avoid;
}
blockquote h3 {
  margin-top: 0;
  font-size: 10.5pt;
  letter-spacing: 0.06em;
  text-transform: uppercase;
}

pre {
  font-family: "Courier New", monospace;
  font-size: 10pt;
  line-height: 1.55;
  background: var(--tint);
  padding: 10pt 12pt;
  border: 0.75pt solid var(--rule);
  white-space: pre-wrap;
  page-break-inside: avoid;
}

code { font-family: "Courier New", monospace; font-size: 10pt; }

/* Passage bodies get slightly tighter leading so they fit a page */
.passage p { margin-bottom: 6pt; }

.lined-page { page-break-before: always; }
.lined-page h3 { margin-bottom: 10pt; }
.rule-line {
  border-bottom: 0.75pt solid var(--rule);
  height: 26pt;
}

.footer-note {
  margin-top: 18pt;
  padding-top: 6pt;
  border-top: 1pt solid var(--rule);
  font-family: "Helvetica Neue", Arial, sans-serif;
  font-size: 8.5pt;
  color: var(--muted);
}

@media screen {
  body { background: #fff; box-shadow: 0 0 0 1px #e3e6ea; padding: 0.5in 0.6in; margin: 20px auto; }
}
"""


def wrap_items(html: str) -> str:
    """Wrap each numbered item and the options that follow it in .item divs."""
    parts = re.split(r'(?=<p><strong>\d+\.</strong>)', html)
    if len(parts) == 1:
        return html
    out = [parts[0]]
    for chunk in parts[1:]:
        # an item ends at the next <hr /> that separates items
        m = re.search(r'<hr\s*/?>', chunk)
        if m:
            body, rest = chunk[:m.start()], chunk[m.start():]
        else:
            body, rest = chunk, ""
        out.append(f'<div class="item">{body}</div>{rest}')
    return "".join(out)


def mark_passages(html: str) -> str:
    """Tag passage sections so they can be styled more tightly."""
    return re.sub(
        r'(<h2[^>]*>(?:Passage|Source)[^<]*</h2>)',
        r'<div class="passage-start"></div>\1',
        html,
    )


def lined_pages(n_pages: int, lines: int, label: str) -> str:
    blocks = []
    for p in range(1, n_pages + 1):
        rules = "".join(f'<div class="rule-line"></div>' for _ in range(lines))
        blocks.append(
            f'<div class="lined-page"><h3>{label} — page {p} of {n_pages}</h3>{rules}</div>'
        )
    return "".join(blocks)


def build(src: Path) -> Path:
    text = src.read_text(encoding="utf-8")

    # Replace the lined-paper placeholder with real ruled lines.
    text = text.replace(
        "*[Insert 3 pages of lined paper for the student response.]*",
        "<!--LINED-->",
    )

    html = markdown.markdown(
        text,
        extensions=["tables", "attr_list", "sane_lists", "md_in_html"],
    )
    html = wrap_items(html)
    html = mark_passages(html)
    html = html.replace(
        "<p><!--LINED--></p>",
        lined_pages(LINED_PAGES, LINES_PER_PAGE, "Extended Writing Task"),
    ).replace("<!--LINED-->", lined_pages(LINED_PAGES, LINES_PER_PAGE, "Extended Writing Task"))

    # Force a page break before each PART heading.
    html = re.sub(r'<h1>(PART \d)', r'<h1 class="pagebreak">\1', html)

    form = "Form A — Pre-Test" if "Form-A" in src.name else "Form B — Post-Test"
    title = f"Grade 7 ELA · {form}"

    page = (
        f'<meta charset="utf-8">\n<title>{title}</title>\n'
        f"<style>{CSS}</style>\n"
        f"{html}\n"
        f'<div class="footer-note">Grade 7 English Language Arts · {form} · '
        f"Aligned to Georgia's K-12 ELA Standards (Grade 7) and the Georgia Milestones "
        f"Grade 7 ELA Assessment Blueprint. Passages are original works written for this "
        f"assessment.</div>"
    )

    OUTDIR.mkdir(exist_ok=True)
    out = OUTDIR / (src.stem.lower() + ".html")
    out.write_text(page, encoding="utf-8")
    return out


def main() -> None:
    sources = [Path(a) for a in sys.argv[1:]] or DEFAULT_SOURCES
    for src in sources:
        if not src.exists():
            print(f"  skipped (not found): {src}")
            continue
        out = build(src)
        print(f"  wrote {out.relative_to(ROOT)}  ({out.stat().st_size // 1024} KB)")
    print("\nOpen each file in a browser and print to PDF:")
    print("  Paper: Letter · Margins: Default · Background graphics: ON · Headers/footers: OFF")


if __name__ == "__main__":
    main()
