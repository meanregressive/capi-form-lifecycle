# Mode: DIFF — where the CAPI and the paper (or two form versions) disagree

Invocation: `/capi-form-lifecycle diff <instrument>`; `<instrument>` = number, short name or full name
from the brief's register (§2). Findings files use the register number, zero-padded (`findings_NN.md`).

Read-only mode. Produces a findings table per instrument and, optionally, one consolidated Word
report across instruments. **Writes nothing into any form, paper or attachment.** Its output is the
input to BUILD (when the paper is ahead of the form) or PATCH (when the form is ahead of the paper,
or when the team wants to know what changed between versions).

Two comparisons are supported:

| Comparison | Sides | Typical trigger |
|---|---|---|
| **CAPI vs paper** | deployed or draft XLSForm; EN paper docx; other-language paper docx | PIs revised the paper after the form was built; before a training; before regenerating paper from CAPI |
| **Version vs version** | two XLSForms of the same instrument | confirming a server download matches the local mirror; documenting what a patch batch changed; endline vs midline form |

Read the project brief first. The brief decides which side is authoritative for content, for codes,
and for each language's text. DIFF reports disagreements and **recommends** per the brief's
authority order. It never picks a winner where the brief says to escalate.

---

## Inputs

- The XLSForm(s). For CAPI vs paper, use the form as deployed (download), not the last build
  output, unless the user says otherwise.
- The paper docx per language. Note the date of each; if one language's paper is older than the
  other, the diff is one-sided for that language and the report must say so.
- The brief: authority order, coding conventions, plumbing conventions (which field names are
  scaffolding), instrument list.
- The extractor script `scripts/extract_for_diff.py` (or the project's own copy). It renders docx
  and xlsx into plain-text dumps so the comparison is deterministic and reviewable.

---

## Phase D0 — Extract

Run the extractor for the instrument. It writes to `05_python_scripts/_diff_dumps/`:

| Dump | Content |
|---|---|
| `capi_NN.txt` | survey rows as `r<row>: [type] name / EN label / other-language label / extras`; choices as `value / EN / other`; settings |
| `en_NN.txt` | EN paper as text plus pipe tables, **tracked changes accepted** (insertions kept, deletions dropped); comments appended with author, date, thread parent, and `[[C<id>]]` anchors in the text |
| `<lang>_NN.txt` | same for each other language |
| `NN_vs_NN.txt` | version-vs-version: cell-level diff after benign-artifact normalization |

The extractor handles OneDrive file locks by copying to temp with retries. If a docx is missing
for a language, write nothing for that side and record `MISSING-DOC` in the findings.

**PDF fallback.** If a paper exists only as PDF, pass the `.pdf` to the extractor: it converts it
with `scripts/pdf_to_docx.py` to `<name>_FROMPDF.docx` in the dumps folder (reviewable, with a
conversion report) and dumps that. The dump's first line reads `PDF-DERIVED`. Write `(PDF-derived)`
after that side in SOURCES, treat the paper as final as printed (no comment threads, no tracked
changes to read), and lower confidence on `RESPONSE-OPTIONS` and `SKIP-LOGIC` rows where the
conversion report flags ragged tables on that page.

## Phase D1 — Read the rubric, then compare

The rubric (`references/diff_rubric.md`) has three parts
that must be applied in this order:

1. **Authority order** (copied from the brief). Default for a paper-first project: the EN paper is
   final on question content, order, options and skip logic; the CAPI is final on response codes,
   list names and ID conventions; each language's paper is final on that language's text.
   Disagreement between languages on *content* is an escalation, not a fix.
2. **Do-not-flag list.** Known conventions and noise: yes/no coded 1/0 while paper shows 1/2;
   other-specify -77 while paper shows 99; don't-know -88 or -99; plumbing rows (start, end,
   deviceid, calculates, pulldata preloads, ID concat fields, `_o` companions, notes that are pure
   enumerator instructions); HTML tags and whitespace in labels; skip arrows versus relevance
   expressions when the logic is the same; one paper question split into several CAPI fields for
   programming; presence or absence of a code prefix in the label. The brief may extend this list.
3. **Flag list** with fixed diff types:

| Diff type | Meaning |
|---|---|
| `MISSING-IN-CAPI` | content question or option in the paper, absent from the form |
| `EXTRA-IN-CAPI` | content question in the form the paper no longer has (not plumbing) |
| `RESPONSE-OPTIONS` | the set of options differs in content (added, removed, merged, reworded so meaning changes) |
| `SKIP-LOGIC` | relevance or skip target differs substantively |
| `WORDING` | EN paper vs `label:ENG` differ in meaning, scope, reference period or respondent |
| `LANG-MISALIGN` | languages disagree on content, or one language names the wrong entity or role |
| `RENUMBERED / ORDER` | code changed or order changed in a way that affects flow |
| `TYPE` | response type differs materially (single vs multi, numeric vs text) |
| `MISSING-DOC` | a whole side was not provided; one row per missing side |

Alignment method: walk the EN paper top to bottom by **question code** as the spine; then sweep
the CAPI for content fields the paper lacks; then check each other language against EN; then, for
version-vs-version, list cell diffs grouped by field.

Severity: **High** = changes what data is collected or produces wrong or blank data (missing
question, wrong options, broken skip, wrong-role translation). **Med** = meaning-changing wording or
order unlikely to break data. **Low** = worth noting, no data consequence.

## Phase D2 — Write `findings_NN.md`

Exactly this structure. No `|` characters inside cells. Quote the differing fragment, not whole
paragraphs.

```
# Instrument NN — <name>
SOURCES: CAPI=<file, version, form_id> ; EN=<file or NONE (date)> ; <LANG>=<file or NONE (date)>

## DIFFERENCES
| Q code | Section / topic | Diff type | Paper says (EN; note other language if it differs) | CAPI has | Severity | Recommendation |
|---|---|---|---|---|---|---|

## SUMMARY
- 3 to 8 bullets: the most consequential differences; whole-section issues; which side is wrong
  when a language misalignment is a paper-side error.
- Counts of flagged rows: High / Med / Low.
- If essentially aligned, say so plainly.
```

Recommendations use the brief's authority order and one of four verbs: **update CAPI**, **update
paper** (translator or PI), **escalate to PI** (design decision), **no action** (with the reason).
Never recommend changing a CAPI code to match the paper's display numbering.

## Phase D3 — Report to the user

Short: instrument, counts by severity, the three to five findings that matter most, and where the
file is. Offer the two follow-ups explicitly:

- **hand to BUILD or PATCH**: the findings table becomes a fix-list source, severity scope decided
  per instrument (High only / High+Med / all) at the start of that pass, not as a blanket rule;
- **consolidated report**: when several instruments are diffed, assemble every `findings_NN.md`
  into one landscape docx with `scripts/build_diff_docx.py --dumps <dumps_dir> --out <docx>
  --project <name> --round <label>`, color-coded by severity, with an executive-summary table
  (sources per instrument, counts by severity). Name it
  `<Project>_<Round>_CAPI_vs_Paper_Differences_<YYYYMMDD>.docx` at the project root and log it in
  the CHANGELOG as a deliverable.

---

## Version-vs-version specifics

- Normalize before comparing: numeric text to number, strip whitespace, CRLF to LF, ignore
  `publishable` unless strict mode. Report drift only after normalization, and say which artifacts
  were folded away.
- Group cell diffs by field `name`, then by column; report whole-row additions and deletions
  separately from edits.
- For a server-download vs mirror check, the expected result is "identical after normalization";
  anything else is a plan row for PATCH (console edit to keep or override).
- For a round-over-round check (endline vs midline), also run the cross-round consistency checks
  from BUILD: identifier fields unchanged, shared choice lists identical, ID suffixes stable, no
  field-name semantic drift (same `name`, different meaning is a hard flag).

---

## Rules specific to DIFF

| # | Rule |
|---|---|
| D1 | Read-only. No form, paper or attachment is modified. |
| D2 | Extract first; compare dumps, not live files. The dumps are the audit trail. |
| D3 | Apply the do-not-flag list before the flag list; conventions are not differences. |
| D4 | Recommendations follow the brief's authority order; language content disagreements are escalated, never resolved by the skill. |
| D5 | A missing or out-of-date side is stated in SOURCES and as a `MISSING-DOC` row; never infer what the missing side says. |
| D6 | Findings files are append-only history: a re-run writes `findings_NN_<YYYYMMDD>.md`, it does not overwrite the previous one. |

## Failure modes

- **Paper docx older than the form** for one language: report one-sided; do not treat the older
  side's silence as a deletion.
- **Tracked changes not accepted** in a dump: the extractor reads `w:t` only, so `w:delText` is
  excluded by construction. If the paper carries tracked changes the PIs have not decided on, say so;
  the dump shows the "accept all" reading.
- **Paper codes and CAPI names diverge** (paper `CD11`, form `D11`): align by content and record the
  rename as `RENUMBERED`, not as missing plus extra.
- **Instrument with no paper at all** (built from scratch, paper regenerated from CAPI later): the
  comparison is CAPI-to-regenerated-paper, which is circular; say so and skip unless a PI-edited paper
  exists.
