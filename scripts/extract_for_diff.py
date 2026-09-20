#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_for_diff.py — render Word questionnaires and XLSForms as plain-text dumps for DIFF mode.

Usage
-----
  python extract_for_diff.py --out <dumps_dir> [--lang-tag ENG] <file> [<file> ...]
  python extract_for_diff.py --out <dumps_dir> --name capi_03 <form.xlsx>
  python extract_for_diff.py --out <dumps_dir> --compare <a.xlsx> <b.xlsx> [--strict]

  <file> may be .docx (paper questionnaire) or .xlsx (XLSForm). Output name defaults to the
  input stem with a .txt suffix; --name overrides it (one input only).

What it does
------------
DOCX  - reads word/document.xml with lxml; text is gathered from every <w:t>, so tracked
        DELETIONS (<w:delText>) are excluded and INSERTIONS included => "accept all" reading.
      - tables rendered as pipe tables in document order (| inside cells -> /).
      - comments dumped as an appendix with author, date, text and THREAD PARENT
        (word/commentsExtended.xml, w15:paraIdParent); inline [[C<id>]] markers show anchors.
XLSX  - survey rows as `r<row>: [type] name | <TAG>: label | L2: label || extras`, choices per
        list as `value | <TAG> | L2`, settings as key: value. Label columns detected flexibly:
        the design language via --lang-tag (default ENG; matches label:ENG / label::English (en)),
        the field language = the bare `label` column.
--compare  cell-level diff of two XLSForms after folding SurveyCTO round-trip artifacts
        (numeric<->text, whitespace, CRLF, `publishable`). --strict keeps them.

OneDrive/Word file locks: files are copied to a temp path with retries and read from there.
No input file is ever modified.
"""
import argparse, os, sys, shutil, tempfile, zipfile, time, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    etree = None
import openpyxl

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
def w(tag): return f"{{{W}}}{tag}"


# ----------------------------------------------------------------------------- locks
def _tmp_copy(path):
    tmp = os.path.join(tempfile.gettempdir(), f"_capi_diff_{os.getpid()}_{os.path.basename(path)}")
    last = None
    for _ in range(6):
        try:
            shutil.copyfile(path, tmp)
            return tmp
        except (PermissionError, OSError) as e:
            last = e; time.sleep(1.0)
    raise last

def _zip_member(path, member):
    for _ in range(2):
        try:
            with zipfile.ZipFile(path) as z:
                return z.read(member) if member in z.namelist() else None
        except (PermissionError, OSError):
            path = _tmp_copy(path)
    return None


# ----------------------------------------------------------------------------- docx
def _text(el):
    parts = []
    for node in el.iter():
        t = node.tag
        if t == w("t"): parts.append(node.text or "")
        elif t == w("tab"): parts.append(" ")
        elif t in (w("br"), w("cr")): parts.append("\n")
        elif t == w("commentReference"):
            cid = node.get(w("id"))
            if cid is not None: parts.append(f" [[C{cid}]] ")
    return "".join(parts)

def _table(tbl):
    out = []
    for tr in tbl.findall(w("tr")):
        cells = [_text(tc).strip().replace("\n", " / ").replace("|", "/") for tc in tr.findall(w("tc"))]
        out.append("| " + " | ".join(cells) + " |")
    return out

def _blocks(parent, out):
    for child in parent:
        t = child.tag
        if t == w("p"):
            s = _text(child).strip()
            if s: out.append(s)
        elif t == w("tbl"):
            out += ["", "--- TABLE ---"] + _table(child) + ["--- /TABLE ---", ""]
        elif t == w("sdt"):
            c = child.find(w("sdtContent"))
            if c is not None: _blocks(c, out)

def _comments(path):
    data = _zip_member(path, "word/comments.xml")
    if not data: return {}
    root = etree.fromstring(data)
    comments, para2cid = {}, {}
    for c in root.findall(w("comment")):
        cid = c.get(w("id"))
        comments[cid] = {"author": c.get(w("author")) or "", "date": (c.get(w("date")) or "")[:10],
                         "text": _text(c).strip(), "parent": None}
        for p in c.iter(w("p")):
            pid = p.get(f"{{{W14}}}paraId")
            if pid: para2cid[pid.upper()] = cid
    ext = _zip_member(path, "word/commentsExtended.xml")
    if ext:
        xr = etree.fromstring(ext)
        for ce in xr.iter(f"{{{W15}}}commentEx"):
            pid = (ce.get(f"{{{W15}}}paraId") or "").upper()
            par = (ce.get(f"{{{W15}}}paraIdParent") or "").upper()
            if pid in para2cid and par in para2cid:
                comments[para2cid[pid]]["parent"] = para2cid[par]
    return comments

def extract_docx(path, out_path):
    if etree is None:
        raise SystemExit("lxml is required for .docx input: pip install lxml")
    data = _zip_member(path, "word/document.xml")
    root = etree.fromstring(data)
    lines = [f"# DOCX DUMP (tracked changes accepted): {os.path.basename(path)}", ""]
    _blocks(root.find(w("body")), lines)
    cm = _comments(path)
    if cm:
        lines += ["", "=== COMMENTS (anchored inline as [[C<id>]]; 'reply to' = thread parent) ==="]
        for cid, c in cm.items():
            rep = f" (reply to C{c['parent']})" if c["parent"] else ""
            lines.append(f"[C{cid}] {c['author']} {c['date']}{rep}: {c['text']}")
    _write(out_path, lines)
    return len(lines)


# ----------------------------------------------------------------------------- xlsx
def _hdr_index(headers, prefix, tag):
    """Index of the design-language column for `prefix` (label/hint/...) given a language tag."""
    tag = tag.lower()
    for i, h in enumerate(headers):
        hl = (h or "").strip().lower()
        if hl.startswith(prefix + ":") and (tag in hl or tag[:2] in hl.split(":")[-1]):
            return i
    return None

def _bare_index(headers, name):
    for i, h in enumerate(headers):
        if (h or "").strip().lower() == name: return i
    return None

def _g(row, i):
    if i is None or i >= len(row) or row[i] is None: return ""
    return str(row[i]).strip()

def extract_xlsx(path, out_path, tag="ENG"):
    for _ in range(2):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=False); break
        except (PermissionError, OSError):
            path = _tmp_copy(path)
    lines = [f"# XLSX DUMP (XLSForm): {os.path.basename(path)}", ""]
    if "settings" in wb.sheetnames:
        rows = list(wb["settings"].iter_rows(values_only=True))
        if rows:
            lines.append("=== SETTINGS ===")
            for h, v in zip(rows[0], rows[1] if len(rows) > 1 else ()):
                if h is not None and v is not None and str(v).strip(): lines.append(f"{h}: {v}")
            lines.append("")
    rows = list(wb["survey"].iter_rows(values_only=True))
    hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
    i_type, i_name = _bare_index(hdr, "type"), _bare_index(hdr, "name")
    i_l1, i_l2 = _hdr_index(hdr, "label", tag), _bare_index(hdr, "label")
    extras = [("rel", _bare_index(hdr, "relevance") if _bare_index(hdr, "relevance") is not None else _bare_index(hdr, "relevant")),
              ("constraint", _bare_index(hdr, "constraint")), ("cfilter", _bare_index(hdr, "choice_filter")),
              ("appear", _bare_index(hdr, "appearance")), ("calc", _bare_index(hdr, "calculation")),
              ("req", _bare_index(hdr, "required")), ("repeat", _bare_index(hdr, "repeat_count")),
              ("default", _bare_index(hdr, "default")),
              (f"hint{tag}", _hdr_index(hdr, "hint", tag)), ("hint", _bare_index(hdr, "hint")),
              (f"cmsg{tag}", _hdr_index(hdr, "constraint message", tag)), ("cmsg", _bare_index(hdr, "constraint message")),
              ("media", _bare_index(hdr, "media:image")), (f"media{tag}", _hdr_index(hdr, "media:image", tag))]
    lines.append(f"=== SURVEY (row: [type] name | {tag}: label | L2: label || extras) ===")
    for ridx, r in enumerate(rows[1:], start=2):
        t, n, l1, l2 = _g(r, i_type), _g(r, i_name), _g(r, i_l1), _g(r, i_l2)
        if not any([t, n, l1, l2]): continue
        line = f"r{ridx}: [{t}] {n}"
        if l1: line += f" | {tag}: {l1}"
        if l2: line += f" | L2: {l2}"
        ex = [f"{k}={_g(r, i)}" for k, i in extras if _g(r, i)]
        if ex: line += "  || " + " ; ".join(ex)
        lines.append(line)
    if "choices" in wb.sheetnames:
        rows = list(wb["choices"].iter_rows(values_only=True))
        hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
        i_list = _bare_index(hdr, "list_name") if _bare_index(hdr, "list_name") is not None else _bare_index(hdr, "list name")
        i_val = _bare_index(hdr, "value") if _bare_index(hdr, "value") is not None else _bare_index(hdr, "name")
        i_c1, i_c2 = _hdr_index(hdr, "label", tag), _bare_index(hdr, "label")
        i_f = _bare_index(hdr, "filter")
        lines += ["", f"=== CHOICES (list | value | {tag} | L2 [| filter]) ==="]
        cur = None
        for r in rows[1:]:
            ln = _g(r, i_list)
            if not ln and not _g(r, i_val): continue
            if ln and ln != cur:
                cur = ln; lines.append(f"-- list: {ln} --")
            f = _g(r, i_f)
            lines.append(f"   {_g(r, i_val)} | {tag}: {_g(r, i_c1)} | L2: {_g(r, i_c2)}" + (f" | filter={f}" if f else ""))
    wb.close()
    _write(out_path, lines)
    return len(lines)


# ----------------------------------------------------------------------------- compare
def _norm(v, strict):
    if v is None: return ""
    if isinstance(v, str):
        s = v.replace("\r\n", "\n")
        if strict: return s.rstrip()
        s = s.strip()
        try:
            f = float(s); return int(f) if f.is_integer() else f
        except ValueError:
            return s
    if isinstance(v, float) and v.is_integer(): return int(v)
    return v

def _load_named(path, strict):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    out = {}
    for ws in wb.worksheets:
        rows = [[_norm(c, strict) for c in r] for r in ws.iter_rows(values_only=True)]
        if not rows: continue
        hdr = [str(h) if h != "" else f"col{i}" for i, h in enumerate(rows[0])]
        keycol = "name" if "name" in hdr else ("value" if "value" in hdr else None)
        recs, order = {}, []
        for r in rows[1:]:
            r = (r + [""] * (len(hdr) - len(r)))[:len(hdr)]
            d = dict(zip(hdr, r))
            if ws.title == "choices":
                k = f"{d.get('list_name','')}::{d.get('value', d.get('name',''))}"
            elif keycol:
                k = d.get(keycol, "")
            else:
                k = str(len(order))
            if not k or k in recs: k = f"{k}#{len(order)}"
            recs[k] = d; order.append(k)
        out[ws.title] = (hdr, recs, order)
    wb.close()
    return out

def compare(a, b, out_path, strict=False):
    A, B = _load_named(a, strict), _load_named(b, strict)
    ignore = set() if strict else {"publishable"}
    lines = [f"# XLSFORM COMPARE: {os.path.basename(a)}  ->  {os.path.basename(b)}",
             f"# artifact folding: {'OFF (strict)' if strict else 'ON (numeric/text, whitespace, CRLF, publishable)'}", ""]
    n = 0
    for sheet in sorted(set(A) | set(B)):
        if sheet not in A or sheet not in B:
            lines.append(f"## sheet {sheet}: only in {'A' if sheet in A else 'B'}"); n += 1; continue
        ha, ra, oa = A[sheet]; hb, rb, ob = B[sheet]
        lines.append(f"## sheet {sheet}")
        for k in oa:
            if k not in rb: lines.append(f"  DEL  {k}"); n += 1
        for k in ob:
            if k not in ra:
                d = rb[k]; lines.append(f"  ADD  {k}  type={d.get('type','')}  label={str(d.get('label',''))[:80]}"); n += 1
        for k in oa:
            if k in rb:
                for col in set(ha) | set(hb):
                    if col in ignore: continue
                    va, vb = ra[k].get(col, ""), rb[k].get(col, "")
                    if va != vb:
                        lines.append(f"  CHG  {k}  [{col}]  {str(va)[:70]!r}  ->  {str(vb)[:70]!r}"); n += 1
    lines.insert(2, f"# differences: {n}")
    _write(out_path, lines)
    return n


# ----------------------------------------------------------------------------- io
def _write(path, lines):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out", required=True, help="dumps directory")
    ap.add_argument("--name", help="output stem (single input only)")
    ap.add_argument("--lang-tag", default="ENG", help="design-language label tag, e.g. ENG")
    ap.add_argument("--compare", action="store_true", help="cell-diff two XLSForms")
    ap.add_argument("--strict", action="store_true", help="keep round-trip artifacts in --compare")
    a = ap.parse_args()
    if a.compare:
        if len(a.files) != 2: raise SystemExit("--compare needs exactly two .xlsx files")
        stem = a.name or f"{os.path.splitext(os.path.basename(a.files[0]))[0]}_vs_{os.path.splitext(os.path.basename(a.files[1]))[0]}"
        out = os.path.join(a.out, stem + ".txt")
        n = compare(a.files[0], a.files[1], out, a.strict)
        print(f"{out}  differences={n}"); return
    if a.name and len(a.files) != 1: raise SystemExit("--name works with a single input")
    for f in a.files:
        if not os.path.exists(f): print(f"MISSING  {f}"); continue
        stem = a.name or os.path.splitext(os.path.basename(f))[0]
        out = os.path.join(a.out, stem + ".txt")
        try:
            if f.lower().endswith(".pdf"):
                # fallback: paper exists only as PDF -> convert to a reviewable docx next to the dump,
                # then dump that. The dump header says PDF-DERIVED (no comments / tracked changes).
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                from pdf_to_docx import convert
                docx_path = os.path.join(a.out, stem + "_FROMPDF.docx")
                convert(f, docx_path)
                n = extract_docx(docx_path, out)
                with open(out, encoding="utf-8") as fh: body = fh.read()
                _write(out, [f"# PDF-DERIVED: converted from {os.path.basename(f)} by pdf_to_docx.py; no comments or tracked changes exist; review {os.path.basename(docx_path)}"] + body.split("\n"))
                print(f"{out}  {n} lines  <- {os.path.basename(f)} (via {os.path.basename(docx_path)})")
                continue
            n = extract_docx(f, out) if f.lower().endswith(".docx") else extract_xlsx(f, out, a.lang_tag)
            print(f"{out}  {n} lines  <- {os.path.basename(f)}")
        except Exception as e:
            print(f"ERROR  {f}: {e}")

if __name__ == "__main__":
    main()
