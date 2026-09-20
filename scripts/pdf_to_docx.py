#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_to_docx.py — fallback when a paper questionnaire exists only as a PDF: convert it to a Word
document a person can review and correct, then feed that docx to DIFF / BUILD like any other paper.

Usage
-----
  python pdf_to_docx.py <paper.pdf> [--out <paper_FROMPDF.docx>] [--no-tables] [--min-chars 40]

  --out         default: same folder, same stem + "_FROMPDF.docx"
  --no-tables   skip table detection (text blocks only), for PDFs whose ruling lines confuse it
  --min-chars   a page with fewer extractable characters than this is treated as scanned (image)

What it does
------------
  - one pass per page with PyMuPDF: detected tables become Word tables (cell text preserved, one
    paragraph per line); remaining text blocks become paragraphs, in reading order (top to bottom,
    interleaved with the tables by position);
  - a banner paragraph at the top states the source file, date, page count and the review caveats;
  - a sidecar report `<out>.conversion_report.txt` lists, per page, tables found, text blocks kept,
    and flags: pages that look scanned (no text layer -> needs OCR, which this script does not do),
    tables with ragged row lengths, and pages where no table was found although the questionnaire
    layout suggests one (many short lines starting with a question code).

What it cannot do
-----------------
  A PDF has no comments and no tracked changes, so the docx it yields carries none: DIFF must treat
  the paper as "final as printed" and say so in SOURCES (`PDF-derived`). Merged cells, skip arrows
  drawn as graphics, check-boxes and two-column layouts can come out wrong; that is why the docx is
  for review, not for direct use. Scanned PDFs need OCR first.

Requires PyMuPDF (`pip install pymupdf`) and python-docx.
"""
import argparse, os, re, sys
from datetime import date

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    raise SystemExit("PyMuPDF is required: pip install pymupdf")
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT

CODE_LINE = re.compile(r"^\s*[A-Z]{1,2}\d{1,3}[a-z]?(?:\.\d{1,2})?[a-z]?\.?\s")


def _clean(s):
    s = (s or "").replace("­", "").replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return "\n".join(l.strip() for l in s.split("\n")).strip()


def page_items(page, use_tables=True):
    """Return [(y0, kind, payload)] in reading order: kind 'table' -> list of rows; 'text' -> str."""
    items, covered = [], []
    if use_tables and hasattr(page, "find_tables"):
        try:
            tabs = page.find_tables()
        except Exception as e:  # pragma: no cover
            tabs = None; print(f"WARN  page {page.number + 1}: table detection failed ({e})")
        if tabs:
            for t in tabs.tables:
                rows = [[_clean(c) if c is not None else "" for c in r] for r in t.extract()]
                rows = [r for r in rows if any(r)]
                if rows:
                    items.append((t.bbox[1], "table", rows)); covered.append(fitz.Rect(t.bbox))
    for b in page.get_text("blocks"):
        x0, y0, x1, y1, txt, _bno, btype = b[:7]
        if btype != 0: continue                                   # image block
        r = fitz.Rect(x0, y0, x1, y1)
        if any(r.intersects(c) and abs(r & c) > 0.5 * abs(r) for c in covered): continue  # inside a table
        txt = _clean(txt)
        if txt: items.append((y0, "text", txt))
    items.sort(key=lambda t: t[0])
    return items


def convert(pdf_path, out_path, use_tables=True, min_chars=40):
    doc_pdf = fitz.open(pdf_path)
    doc = Document()
    sec = doc.sections[0]
    if doc_pdf.page_count and doc_pdf[0].rect.width > doc_pdf[0].rect.height:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = Inches(11), Inches(8.5)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"): setattr(sec, m, Inches(0.7))
    doc.styles["Normal"].font.size = Pt(10)

    banner = (f"CONVERTED FROM PDF — review before use. Source: {os.path.basename(pdf_path)} "
              f"({doc_pdf.page_count} pages), converted {date.today().isoformat()} by pdf_to_docx.py. "
              "Table boundaries, merged cells, skip arrows and option codes may be wrong. A PDF carries no "
              "comments or tracked changes; treat this paper as final as printed.")
    p = doc.add_paragraph(); r = p.add_run(banner); r.bold = True; r.font.size = Pt(9)

    report = [f"# pdf_to_docx conversion report — {os.path.basename(pdf_path)} -> {os.path.basename(out_path)}",
              f"pages: {doc_pdf.page_count}   tables: {'on' if use_tables else 'off'}", ""]
    flags, n_tables, n_text = [], 0, 0
    for page in doc_pdf:
        chars = len(page.get_text().strip())
        if chars < min_chars:
            flags.append(f"page {page.number + 1}: only {chars} characters of text -> probably scanned; needs OCR")
            doc.add_paragraph(f"[page {page.number + 1}: no text layer — scanned page, not converted]").runs[0].italic = True
            report.append(f"page {page.number + 1}: SCANNED? chars={chars}")
            continue
        items = page_items(page, use_tables)
        pt = sum(1 for i in items if i[1] == "table"); px = sum(1 for i in items if i[1] == "text")
        n_tables += pt; n_text += px
        code_lines = sum(1 for i in items if i[1] == "text" for l in i[2].split("\n") if CODE_LINE.match(l))
        if use_tables and pt == 0 and code_lines >= 5:
            flags.append(f"page {page.number + 1}: no table detected but {code_lines} question-code lines -> check layout")
        for _y, kind, payload in items:
            if kind == "table":
                width = max(len(r) for r in payload)
                if any(len(r) != width for r in payload):
                    flags.append(f"page {page.number + 1}: table with ragged rows (widths {sorted(set(len(r) for r in payload))})")
                t = doc.add_table(rows=0, cols=width); t.style = "Table Grid"
                for row in payload:
                    cells = t.add_row().cells
                    for j, val in enumerate(row + [""] * (width - len(row))):
                        cells[j].text = val
                        for par in cells[j].paragraphs:
                            for run in par.runs: run.font.size = Pt(9)
                doc.add_paragraph()
            else:
                for line in payload.split("\n"):
                    if line.strip(): doc.add_paragraph(line)
        report.append(f"page {page.number + 1}: tables={pt} text_blocks={px} code_lines={code_lines}")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    report += ["", f"totals: tables={n_tables} text_blocks={n_text}", ""]
    report += (["FLAGS:"] + [f"  - {f}" for f in flags]) if flags else ["FLAGS: none"]
    report += ["", "Review checklist: every question code present once; option codes and skip arrows intact;",
               "no table split across a page break lost its header; section headings kept; nothing merged across columns."]
    with open(out_path + ".conversion_report.txt", "w", encoding="utf-8") as f: f.write("\n".join(report))
    print(f"wrote {out_path}\n  tables={n_tables} text_blocks={n_text} flags={len(flags)}  report: {out_path}.conversion_report.txt")
    return out_path, flags


def main():
    sys.stdout = __import__("io").TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")   # only as a script
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf"); ap.add_argument("--out"); ap.add_argument("--no-tables", action="store_true")
    ap.add_argument("--min-chars", type=int, default=40)
    a = ap.parse_args()
    out = a.out or os.path.splitext(a.pdf)[0] + "_FROMPDF.docx"
    convert(a.pdf, out, use_tables=not a.no_tables, min_chars=a.min_chars)


if __name__ == "__main__":
    main()
