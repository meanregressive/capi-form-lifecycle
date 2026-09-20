# Mode: BUILD — new round of an instrument from the paper questionnaire and the prior-round form

Use this mode once per instrument at the start of a round: the PIs have finalized the paper
questionnaire (all languages) and no form for this round is on the server yet. A prior-round XLSForm
is an input when one exists; a first round, or a new instrument, builds from the paper alone (see
"From-scratch mode"). Everything after the first upload is PATCH, not BUILD.

BUILD keeps three things consistent at once:

- **paper ↔ form** (the questionnaire the PIs signed off vs what the tablet asks),
- **language A ↔ language B** (every visible string present in every language column),
- **round N ↔ round N−1** (identifiers, shared lists, ID scheme, field-name meaning unchanged).

Read the project brief first (`00_reference/capi_project_brief.md`). It supplies: the instrument
register (§2), where the inputs live and the data-governance boundary (§3), the authority order (§4),
the translation policy (§5), the coding conventions (§6), the identifier scheme and registry (§7),
versioning and deployment (§8), the logs (§9), the folder layout (§10). This mode file does not
restate them; where it says "per the brief" it means look it up there.

Invocation: `/capi-form-lifecycle build <instrument>` where `<instrument>` is a number, short name or full
name from the brief's register (§2) · add `--dry-run` for Phases A–B only (nothing written) ·
add `--from-scratch` when the instrument has no prior-round form (implied when the register's
"Built from" column says `scratch`, or the brief's prior-round folder is empty).

---

## Inputs

1. Paper questionnaire in the source language, latest dated subfolder (brief §3). PI comments in the
   docx are read as decisions.
2. Paper questionnaire in each other language, latest dated subfolder. Note its date relative to the
   source language; a lag is reported, not silently absorbed.
3. Prior-round XLSForm for this instrument (brief §3), or `--from-scratch`.
4. Prior-round preload CSVs: **header rows only** (brief §3 governance rule), or a `describe` log.
5. The brief, the ID-suffix registry, the glossary, the code patterns (`references/xlsform_patterns.md`).

If 1, 2 or 3 is missing, stop and say which. Do not build from an older dated subfolder without
saying so.

**Paper available only as PDF.** Run `scripts/pdf_to_docx.py <paper.pdf>`; it writes
`<stem>_FROMPDF.docx` plus a conversion report. Hand both to the user for review (table boundaries,
merged cells, skip arrows, option codes) **before** Phase A reads the docx, and record in the
CHANGELOG "Built from" line that the paper was PDF-derived. A PDF carries no comments and no
tracked changes, so PI decisions must come from the conversation instead. A scanned PDF (no text
layer) is reported by the script and needs OCR first; the skill does not do OCR.

---

## Phase A — Discovery (silent)

1. Resolve `<instrument>` to its register row; identify the three files by the row's `file stem` (or the
   leading number in the filename when the stem is blank). Record filenames and dates.
2. Parse the source-language docx: section headings; question tables (code | text | options);
   skip arrows (`>>`, `→`); comments from `word/comments.xml` **grouped into threads** via
   `w15:paraIdParent` in `word/commentsExtended.xml`; tracked changes read as accepted.
3. Parse each other-language docx the same way and align rows to the source language by question
   code.
4. Load the prior-round form (skip under `--from-scratch`): `survey`, `choices`, `settings`. Capture
   preload calc patterns, ID-suffix use, choice-list values, groups and repeats, the phone/name/age
   constraint conventions.
5. Detect: field-name reuse with a different meaning across rounds; `[NAME]`/`[ROLE]` placeholders
   (signal preload columns); skip arrows needing `relevance`; unresolved comment threads
   (resolution judged by whether the thread's proposal is reflected in the paper, not by the
   `w15:done` flag).
6. **External dependency scan**: every `pulldata(...)`, `search(...)`, `select_*_from_file`,
   `type: enumerator`, and `media:*` reference. Separate **form attachments** (files uploaded to
   the form) from **server datasets** (enumerator dataset, attached on the console). Check each
   attachment reference against the attachments folder by **header only**.

## Phase B — Diff summary (ALWAYS shown, ALWAYS requires sign-off)

A terse table per section:

| Field | Status | Prior round → New | Notes |
|---|---|---|---|
| A05 | NEW | — → calculate pulldata | needs `village` column in upload_school |
| B06 | DELETED | text → — | replaced by deputy block |
| C10a | LABEL-CHANGED | "leader" → "president" | EN only; PT already says Presidente |
| C12 block | NEW SECTION | — → 5 fields | source: PI comment thread 2026-05-21 |

Statuses: `NEW`, `DELETED`, `RENAMED`, `RESTRUCTURED`, `LABEL-CHANGED`, `SKIP-CHANGED`,
`CODE-CHANGED`, `TYPE-CHANGED`. Under `--from-scratch` every row is `NEW` and the "Prior round"
side reads `—`; the table is still shown and still needs sign-off, because it is the agent's reading
of the paper.

Then report:

1. **Language alignment**: rows present in one language's paper and missing in another. Handle per
   the brief's translation policy (§5); never fill silently.
2. **Prior-round fields the paper is silent on**: default keep; list them for the user to drop.
3. **Comment threads**: one entry per thread, leaf comments verbatim with author and date, one
   resolution status per thread.
4. **Cross-instrument and cross-round flags**: ID-suffix collisions (registry), shared-list value
   drift, identifier fields changed, field-name semantic drift (hard flag).
5. **External dependency inventory**: reference · mechanism (attachment / server dataset) · expected
   location · present (header check) · columns required. Missing attachments are flags, not blockers;
   server datasets are always "configure on console".
6. **form_id and bindings**: the new form_id per brief §8 (`<prior_form_id>_<round>_v1`) and the
   reminder that every binding must be attached to it on first upload.
7. **Data-team handoff preview**: new columns needed in existing CSVs; new files needed.

Stop. Ask: *"Diff above. Sign off to proceed to clarifying questions, or redirect?"*
`--dry-run` ends here.

## Phase C — Clarifying questions (only when needed)

Structural decisions only, never label wording. Cap: two rounds of at most four questions. If more
is needed the instrument should be split into passes; say so.

Typical questions: keep or drop a restructured prior-round block; repeat group vs chained prompts
(default: keep `begin repeat`); ambiguous skip target; which of two dated papers to trust; a new
constraint the brief has no convention for (constraints with a brief convention are copied without
asking); a missing translation under a "halt" policy.

## Phase D — Build

1. **Copy first.** `shutil.copyfile(prior_round_xlsx, output_xlsx)`; the prior-round file is never
   edited. Output name per brief §8, default `<prior_stem>_<round>_v1.xlsx`.
2. **Transformation script** at `05_python_scripts/_build_<round>_<short_name>.py` (folder per brief
   §10). Self-contained, re-runnable, docstring lists the Phase B rows it implements. Helpers: find
   row by `name`; find begin/end pair; insert rows at a named anchor; delete a named range. Read and
   write through `scripts/xlsform_io.py` (`XLSFormBook`): sheets are rebuilt from row lists (openpyxl
   `value=None` does not clear a cell) and the prior-round workbook's formatting is carried into the
   new form by field name, so an authored workbook's color coding and widths are not lost.
3. Apply the brief's conventions to every new field: codes, other-specify companion, don't-know and
   exclusive-option constraint, question-code prefix in every language column, required flag,
   phone/name/age constraints copied from the nearest prior-round sibling, ID suffix from the
   registry (add a row to the registry if a new role appears and tell the user).
4. Translation per brief §5. Every string the skill writes in a non-source language is marked.
5. `settings`: `form_id` per brief §8; `version` = `YYMMDDNNNN` literal (never a formula);
   `form_title` with the round appended; `default_language` set to the field language's tag.
6. **Round-trip verify** with `scripts/validate_xlsform.py` or equivalent: every NEW field present in
   every language column; every DELETED name absent; `${refs}` resolve; braces balanced; groups and
   repeats balanced; every `select_*` list exists with no duplicate values; pulldata files and
   columns exist (header check); regex anchored; no `.>0` where the paper allows 0; no stale cells.
   Print PASS/FAIL per check. Any FAIL: fix the script and re-run before reporting.

## Phase E — Logs and report

Append (never overwrite) to the logs named in brief §9:

1. **CHANGELOG.md**: `## Instrument <number> — <full name> — <round> v1 — <date>` with Changes, Data-team
   handoff (columns to add, files needed, attachments confirmed), Open questions, Built from (paper
   files with dates, prior-round file, script).
2. **`<round>_flags_and_decisions.md`**: language misalignments (observed / decision with date /
   risk if reversed / translator action), prior-round quirks preserved, paper-omitted-but-retained
   fields, resolved comment threads (thread structure shown), design decisions.
3. **`<round>_todos.md`**: open actions by owner (data team, user/PI, translator, cleanup).
4. **Preload spec** (`CAPI_PRELOAD_UPDATES_NEEDED.md`): the columns and files this form needs, test
   vs production values.
5. **Translation-flags file** if the policy is draft-and-flag.

Report to the user, terse: output path; changelog excerpt; data-team handoff; pointers to the flags
and todos files; script path; the first-upload checklist (`references/deployment_checklist.md`):
attach the enumerator dataset, upload the attachments subfolder, publish, download the deployed xlsx
into the mirror folder.

---

## From-scratch mode

Used when there is no prior-round form for this instrument. Two cases:

**A sibling form of this round already exists** (a new instrument in an ongoing project):

1. Start from the sibling's header block (metadata calculates, enumerator field, identification
   group, consent), copied with openpyxl, not retyped.
2. Reuse shared choice lists (`yesno`, personnel, languages) from the sibling verbatim.
3. `settings`: `form_id = <short_name>_<round>_v1`, title per the paper, version literal.
4. Build the instrument's sections from the paper with the patterns file. New roles get a registry
   row. The script records which sibling was the stub so the build is reproducible.

**No form exists yet in the project** (first round, first instrument):

1. Header block from the patterns file: metadata rows (`start`, `end`, `deviceid`, and whichever the
   brief §6 lists), `type: enumerator` if the brief §8 names an enumerator dataset, the
   identification group built on the brief §7 primary key, a consent block from the paper.
2. Shared choice lists are created now from the brief §6 conventions (`yesno` with the brief's
   values, don't-know / refused / other codes) and become the project's canonical lists: every
   later sibling copies them from this form. Say so in the CHANGELOG.
3. `settings` as above. `default_language` = the tag of the un-tagged `label` column (brief §1).
4. The ID-suffix registry (brief §7) starts here with the roles this instrument introduces.
5. Phase B still runs: the table lists every paper question as `NEW` and is signed off before
   building. Cross-round checks do not apply; cross-instrument checks apply from the second form on.

## Cross-round consistency checks (run when a prior-round form exists)

1. Identifier fields (cluster ID, names of the identification block): same `name`, `type`, labels.
2. Shared choice lists: identical values.
3. ID suffixes: same role, same suffix.
4. Field-name semantic drift: same `name`, different meaning → hard flag, do not build until
   resolved.

## Cross-instrument consistency (same round)

For blocks shared across instruments (enumerator block, school identification, consent): identical
`name`, `type`, labels, `choice_filter`, `appearance`, list values. Flag, never silently reconcile.

---

## Rules specific to BUILD

| # | Rule |
|---|---|
| B1 | Never edit the prior-round form. Copy first. |
| B2 | Phase B sign-off before any clarifying question or build step. |
| B3 | One build script per instrument per round; all Phase B rows reproducible from it. |
| B4 | Conventions come from the brief; a paper value that contradicts a convention is paper formatting unless the user says otherwise. |
| B5 | Translation per the brief's policy; every skill-written string is marked. |
| B6 | Constraints: copy the brief's convention for the field type without asking; anything else, ask. |
| B7 | Prior-round quirks (misnamed groups, typos) are flagged, not silently fixed. |
| B8 | Skip arrows become `relevance` on the following fields, never literal jumps. |
| B9 | Round-trip verification is mandatory before reporting success. |
| B10 | Every external dependency is inventoried; attachments verified by header only. |

## Failure modes

- **Locked input** (Word or OneDrive): copy to temp with retries and parse the copy.
- **Locked output** (open in Excel): ask the user to close it; never force-overwrite.
- **Sync-conflict filename** (`-DESKTOP-XXX`): stop, let the user resolve.
- **Preload column referenced but absent from the CSV header**: flag; never invent the column.
- **Choice value defined twice in one list**: hard fail with the location.
- **Paper code differs from the prior-round name** (paper `CD11`, form `D11`): keep the form name,
  note the rename in the flags log.
- **Two dated papers for one language**: use the latest unless told otherwise; record which.
- **Phase C would exceed its cap**: stop and propose splitting the instrument into passes.
