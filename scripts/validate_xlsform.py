#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate_xlsform.py — PASS/FAIL checks that BUILD and PATCH must run before reporting success.

Usage
-----
  python validate_xlsform.py <form.xlsx> [--base <prior.xlsx>] [--attachments <dir>]
                             [--lang-tag ENG] [--server-version 2607100001]
                             [--planned name1,name2,...] [--dk -88] [--other -77]

Checks (each prints PASS / FAIL / WARN; exit code 1 if any FAIL):
  1  group/repeat balance (begin/end pairs, nesting, names match)
  2  ${refs} in relevance/constraint/calculation/repeat_count/labels/hints/choice_filter resolve
  3  every select_one/select_multiple list exists; no duplicate values in a list; no blank values
  4  bilingual completeness on visible rows (design-language label and bare label both present;
     hints and constraint messages present in both languages when present in one)
  5  question-code prefix present in both languages when present in one (WARN)
  6  pulldata('<file>','<col>','<key>',...) -> file exists in --attachments, col + key in header
     (header row only is read); search() appearance file exists
  7  expression traps: regex 'contains' written without .*; '.>0' constraints (WARN);
     decimal comma in hints (WARN); translate() anywhere; jr:choice-name on a pulldata calc (WARN)
  8  settings: version numeric, > --server-version if given; form_id present; default_language set
  9  --base: intended-diff report — changed/added/deleted names vs the base; if --planned given,
     any changed name outside the planned set is a FAIL (unplanned edit)
 10  stale-cell check: no values beyond the header width; no orphan rows without type/name
 11  repeats without a label (WARN); 'other' companion present for lists containing --other (WARN)
"""
import argparse, csv, io, os, re, sys
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
FAILS, WARNS = [], []

def out(status, check, msg):
    print(f"[{status}] {check}: {msg}")
    if status == "FAIL": FAILS.append((check, msg))
    elif status == "WARN": WARNS.append((check, msg))

def load(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    sheets = {}
    for name in ("survey", "choices", "settings"):
        if name in wb.sheetnames:
            rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
            hdr = [str(h).strip() if h is not None else "" for h in (rows[0] if rows else [])]
            body = [(r + [None] * (len(hdr) - len(r))) for r in rows[1:]]
            sheets[name] = (hdr, body)
    wb.close()
    return sheets

def col(hdr, name):
    name = name.lower()
    for i, h in enumerate(hdr):
        if h.lower() == name: return i
    return None

def col_tag(hdr, prefix, tag):
    tag = tag.lower()
    for i, h in enumerate(hdr):
        hl = h.lower()
        if hl.startswith(prefix.lower() + ":") and tag in hl: return i
    return None

def g(row, i):
    if i is None or i >= len(row) or row[i] is None: return ""
    return str(row[i]).strip()

REF = re.compile(r"\$\{([A-Za-z0-9_.\-]+)\}")
PULL = re.compile(r"pulldata\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'", re.I)
SEARCH = re.compile(r"search\(\s*'([^']+)'", re.I)
VISIBLE_SKIP = {"calculate", "calculate_here", "start", "end", "deviceid", "subscriberid", "simserial",
                "phonenumber", "username", "duration", "caseid", "hidden", "end group", "end repeat",
                "audio audit", "text audit", "speed violations count", "speed violations list", "speed violations audit"}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("form"); ap.add_argument("--base"); ap.add_argument("--attachments")
    ap.add_argument("--lang-tag", default="ENG"); ap.add_argument("--server-version")
    ap.add_argument("--planned", help="comma-separated field names the plan allows to change")
    ap.add_argument("--dk", default="-88"); ap.add_argument("--other", default="-77")
    a = ap.parse_args()
    S = load(a.form)
    if "survey" not in S: out("FAIL", "load", "no survey sheet"); return finish()
    hdr, rows = S["survey"]
    ci = {k: col(hdr, k) for k in ("type", "name", "label", "relevance", "constraint", "calculation",
                                   "repeat_count", "choice_filter", "appearance", "hint", "constraint message",
                                   "required", "default", "media:image")}
    if ci["relevance"] is None: ci["relevance"] = col(hdr, "relevant")
    l1 = col_tag(hdr, "label", a.lang_tag); h1 = col_tag(hdr, "hint", a.lang_tag)
    cm1 = col_tag(hdr, "constraint message", a.lang_tag)
    rows = [r for r in rows if any(g(r, i) for i in (ci["type"], ci["name"]))]
    names = [g(r, ci["name"]) for r in rows]
    nameset = set(n for n in names if n)

    # 1 balance
    stack = []; bad = 0
    for r in rows:
        t, n = g(r, ci["type"]).lower(), g(r, ci["name"])
        if t in ("begin group", "begin repeat", "begin_group", "begin_repeat"): stack.append((t.split()[0] + " " + t.split("_")[-1] if "_" in t else t, n))
        elif t in ("end group", "end repeat", "end_group", "end_repeat"):
            kind = t.replace("_", " ").split()[1]
            if not stack: bad += 1; out("FAIL", "balance", f"'{t}' {n} with nothing open"); continue
            k, on = stack.pop()
            if k.split()[1] != kind or (n and on and n.lower() != on.lower()): bad += 1; out("FAIL", "balance", f"'{t}' {n} closes '{k}' {on}")
            elif n and on and n != on: out("WARN", "balance", f"'{t}' {n} closes '{k}' {on} (case differs; SurveyCTO tolerates it)")
    if stack: bad += 1; out("FAIL", "balance", f"unclosed: {stack}")
    if not bad: out("PASS", "balance", "groups/repeats balanced")

    # 2 refs
    ref_cols = [ci[k] for k in ("relevance", "constraint", "calculation", "repeat_count", "choice_filter", "label", "hint", "default", "media:image", "constraint message") if ci[k] is not None] + [i for i in (l1, h1, cm1) if i is not None]
    unresolved = set()
    for r in rows:
        for i in ref_cols:
            for m in REF.findall(g(r, i)):
                if m not in nameset: unresolved.add((g(r, ci["name"]), m))
    out("FAIL" if unresolved else "PASS", "refs", f"{len(unresolved)} unresolved ${{}} refs" + (f": {sorted(unresolved)[:15]}" if unresolved else ""))

    # 3 lists
    lists, dups, blanks = {}, [], []
    if "choices" in S:
        chdr, crows = S["choices"]
        cl, cv = col(chdr, "list_name") if col(chdr, "list_name") is not None else col(chdr, "list name"), col(chdr, "value") if col(chdr, "value") is not None else col(chdr, "name")
        ccl1, ccl2 = col_tag(chdr, "label", a.lang_tag), col(chdr, "label")
        for r in crows:
            ln, v = g(r, cl), g(r, cv)
            if not ln: continue
            if not v: blanks.append(ln)
            if v in lists.setdefault(ln, {}): dups.append((ln, v))
            lists[ln][v] = (g(r, ccl1), g(r, ccl2))
    missing_lists, other_no_companion = [], []
    for r in rows:
        t = g(r, ci["type"])
        m = re.match(r"select_(one|multiple)(?:_from_file)?\s+(\S+)", t)
        if m and not t.endswith(".csv") and "from_file" not in t:
            ln, n = m.group(2), g(r, ci["name"])
            if ln not in lists: missing_lists.append((n, ln))
            elif a.other in lists[ln] and (n + "_o") not in nameset and not any(x.startswith(n) and x.endswith("_o") for x in nameset):
                other_no_companion.append(n)
    out("FAIL" if missing_lists else "PASS", "lists", f"{len(missing_lists)} select lists missing" + (f": {missing_lists[:10]}" if missing_lists else ""))
    out("FAIL" if dups else "PASS", "list-dups", f"{len(dups)} duplicate values" + (f": {dups[:10]}" if dups else ""))
    if blanks: out("FAIL", "list-blank", f"blank values in lists {sorted(set(blanks))[:10]}")
    if other_no_companion: out("WARN", "other-specify", f"{len(other_no_companion)} selects with {a.other} but no *_o companion: {other_no_companion[:10]}")

    # 4/5 bilingual + code prefix
    miss, prefix_mismatch = [], []
    code = re.compile(r"^\s*(?:<[^>]+>\s*)*[A-Z]{1,2}\d{1,3}[a-z]?(?:[._]\d+)?[a-z]?\.")
    for r in rows:
        t = g(r, ci["type"]).lower(); n = g(r, ci["name"])
        if not t or t in VISIBLE_SKIP or t.startswith("begin") or t.startswith("calculate"): continue
        a1, a2 = g(r, l1), g(r, ci["label"])
        if bool(a1) != bool(a2): miss.append((n, "label"))
        for (x, y, what) in ((g(r, h1), g(r, ci["hint"]), "hint"), (g(r, cm1), g(r, ci["constraint message"]), "constraint message")):
            if bool(x) != bool(y): miss.append((n, what))
        if a1 and a2 and (bool(code.match(a1)) != bool(code.match(a2))): prefix_mismatch.append(n)
    out("FAIL" if miss else "PASS", "bilingual", f"{len(miss)} visible rows with one language missing" + (f": {miss[:12]}" if miss else ""))
    if prefix_mismatch: out("WARN", "code-prefix", f"{len(prefix_mismatch)} rows with a code prefix in one language only: {prefix_mismatch[:12]}")
    if "choices" in S:
        cmiss = [(ln, v) for ln, d in lists.items() for v, (x, y) in d.items() if bool(x) != bool(y)]
        out("FAIL" if cmiss else "PASS", "bilingual-choices", f"{len(cmiss)} choices with one language missing" + (f": {cmiss[:10]}" if cmiss else ""))

    # 6 pulldata / search (header only)
    if a.attachments:
        headers = {}
        for f in os.listdir(a.attachments):
            if f.lower().endswith(".csv"):
                try:
                    with io.open(os.path.join(a.attachments, f), encoding="utf-8-sig", errors="replace") as fh:
                        headers[os.path.splitext(f)[0]] = [c.strip() for c in next(csv.reader(fh))]
                except StopIteration:
                    headers[os.path.splitext(f)[0]] = []
        probs = []
        for r in rows:
            for i in ref_cols + [ci["appearance"]]:
                s = g(r, i)
                for fn, c, k in PULL.findall(s):
                    if fn not in headers: probs.append((g(r, ci["name"]), fn, "file missing"))
                    else:
                        if c not in headers[fn]: probs.append((g(r, ci["name"]), fn, f"column {c} missing"))
                        if k not in headers[fn]: probs.append((g(r, ci["name"]), fn, f"key {k} missing"))
                for fn in SEARCH.findall(g(r, ci["appearance"])):
                    if fn not in headers: probs.append((g(r, ci["name"]), fn, "search file missing"))
        media = {f for f in os.listdir(a.attachments)}
        for r in rows:
            m = g(r, ci["media:image"])
            if m and not m.startswith("${") and m not in media and not any(z.lower().endswith(".zip") for z in media):
                probs.append((g(r, ci["name"]), m, "media file missing (no zip present)"))
        out("FAIL" if probs else "PASS", "attachments", f"{len(probs)} pulldata/search/media problems" + (f": {probs[:12]}" if probs else ""))
    else:
        out("WARN", "attachments", "no --attachments dir given; pulldata columns not checked")

    # 7 traps
    traps = []
    pulled = {g(r, ci["name"]) for r in rows if "pulldata(" in g(r, ci["calculation"]).lower()}
    for r in rows:
        n = g(r, ci["name"]); c = g(r, ci["constraint"]); calc = g(r, ci["calculation"])
        for m in re.finditer(r"regex\(\s*\.\s*,\s*(['\"])(.*?)\1", c):
            pat = m.group(2)
            if not (pat.startswith(".*") or pat.startswith("^") or pat.endswith("+") or pat.endswith("*") or pat.endswith("$") or pat.endswith("}")):
                traps.append(("FAIL", n, f"regex not whole-string safe: {pat}"))
        if re.search(r"(^|[^.\d])\.\s*>\s*0(?![.\d])", c): traps.append(("WARN", n, "constraint '.>0' excludes 0"))
        for i in (h1, ci["hint"], cm1, ci["constraint message"]):
            if re.search(r"\d,\d", g(r, i)) and re.search(r"\b0,5\b|\b\d,\d\b(?!\d)", g(r, i)): traps.append(("WARN", n, "decimal comma in hint/message")); break
        if "translate(" in calc.lower() or "translate(" in g(r, ci["media:image"]).lower(): traps.append(("FAIL", n, "translate() is not supported"))
        m = re.search(r"jr:choice-name\(\s*\$\{([^}]+)\}", calc)
        if m and m.group(1) in pulled: traps.append(("WARN", n, f"jr:choice-name on pulldata field {m.group(1)} renders blank"))
    for st, n, msg in traps: out(st, "traps", f"{n}: {msg}")
    if not traps: out("PASS", "traps", "no regex/zero/decimal/translate/choice-name traps")

    # 8 settings
    if "settings" in S:
        shdr, srows = S["settings"]; sv = srows[0] if srows else []
        ver, fid, dl = g(sv, col(shdr, "version")), g(sv, col(shdr, "form_id")), g(sv, col(shdr, "default_language"))
        if not ver.isdigit(): out("FAIL", "settings", f"version not a literal number: {ver!r}")
        elif a.server_version and int(ver) <= int(a.server_version): out("FAIL", "settings", f"version {ver} not above server {a.server_version}")
        else: out("PASS", "settings", f"version {ver}" + (f" > server {a.server_version}" if a.server_version else ""))
        out("PASS" if fid else "FAIL", "form_id", fid or "missing")
        out("PASS" if dl else "WARN", "default_language", dl or "blank (server will warn 'unknown language')")
    else:
        out("FAIL", "settings", "no settings sheet")

    # 10 stale cells / orphan rows
    wide = [i for i, r in enumerate(rows, start=2) if any(v not in (None, "") for v in r[len(hdr):])]
    orphan = [i for i, r in enumerate(rows, start=2) if not g(r, ci["type"]) and g(r, ci["name"])]
    out("FAIL" if wide else "PASS", "stale-cells", f"{len(wide)} rows with values beyond the header" + (f" at {wide[:10]}" if wide else ""))
    if orphan: out("FAIL", "orphan-rows", f"rows with a name but no type at {orphan[:10]}")

    # 11 repeat labels
    nolabel = [g(r, ci["name"]) for r in rows if g(r, ci["type"]).lower().replace("_", " ") == "begin repeat" and not (g(r, l1) or g(r, ci["label"]))]
    if nolabel: out("WARN", "repeat-label", f"repeats without a label: {nolabel}")

    # 9 base diff
    if a.base:
        B = load(a.base)
        if "survey" in B:
            bh, br = B["survey"]; bn = col(bh, "name")
            bmap = {g(r, bn): dict(zip(bh, r)) for r in br if g(r, bn)}
            nmap = {g(r, ci["name"]): dict(zip(hdr, r)) for r in rows if g(r, ci["name"])}
            added = [n for n in nmap if n not in bmap]; deleted = [n for n in bmap if n not in nmap]
            changed = []
            for n in nmap:
                if n in bmap:
                    for k in set(hdr) | set(bh):
                        if k in ("publishable",): continue
                        va, vb = str(bmap[n].get(k, "") or "").strip(), str(nmap[n].get(k, "") or "").strip()
                        if va != vb: changed.append(n); break
            print(f"[INFO] diff vs base: {len(added)} added {added[:15]} | {len(deleted)} deleted {deleted[:15]} | {len(changed)} changed {changed[:25]}")
            if a.planned:
                planned = {p.strip() for p in a.planned.split(",") if p.strip()}
                unplanned = [n for n in added + deleted + changed if n not in planned]
                out("FAIL" if unplanned else "PASS", "intended-diff", f"{len(unplanned)} changed fields outside the plan" + (f": {unplanned[:20]}" if unplanned else ""))
            else:
                out("WARN", "intended-diff", "no --planned list given; every change above must map to a plan row")
        # 12 formatting carried through? (WARN only: server downloads have none, authored workbooks do)
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            from xlsform_io import summary
            sb, sn = summary(a.base).get("survey", {}), summary(a.form).get("survey", {})
            lost = [k for k in ("fills", "colored_font", "widths", "hidden_cols", "cf_rules", "dv") if sb.get(k, 0) > sn.get(k, 0)]
            if sb.get("freeze") and not sn.get("freeze"): lost.append("freeze")
            if lost: out("WARN", "formatting", f"base had formatting the new file lost: {lost} (use xlsform_io.XLSFormBook or `xlsform_io.py carry`)")
            else: out("PASS", "formatting", "no formatting lost vs base")
        except Exception as e:  # pragma: no cover
            out("WARN", "formatting", f"could not compare formatting: {e}")
    return finish()

def finish():
    print(f"\n{'FAILED' if FAILS else 'PASSED'}: {len(FAILS)} FAIL, {len(WARNS)} WARN")
    sys.exit(1 if FAILS else 0)

if __name__ == "__main__":
    main()
