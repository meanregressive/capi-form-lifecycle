# Mode: PATCH — fix a deployed CAPI form and ship the next version

Use this mode when a form is **already on the SurveyCTO server** (testing, training, pilot or live
collection) and a list of fixes has to be applied. This is the mode used most often in a survey
round: BUILD runs once per instrument, PATCH runs every time the team finds something.

PATCH is deliberately different from BUILD:

| | BUILD | PATCH |
|---|---|---|
| Base file | prior-round XLSForm + paper in every language | **the XLSForm downloaded from the server today** |
| Driver | paper questionnaire | a **fix list** (testing feedback, diff findings, deferred backlog) |
| Output | `<stem>_<round>_v1.xlsx`, new form_id | same form, **version bumped**, form_id usually unchanged |
| Audit script | one full build script from the prior round | one **patch script per patch**, chained |

Invocation: `/capi-form-lifecycle patch <instrument>`; `<instrument>` = number, short name or full name
from the brief's register (§2).

Read the project brief (`00_reference/capi_project_brief.md`) before starting. It holds the
conventions (codes, ID scheme, translation rules, authority order) every patch must respect.

---

## Inputs

1. **The deployed form, downloaded from SurveyCTO now.** Never patch the file the last script wrote;
   the team may have edited the form on the console since. If a local mirror exists
   (`04_output_<round>_capi/v<n>_deployed/`), compare it cell-for-cell with the download and report
   drift. Ignore the benign round-trip artifacts SurveyCTO introduces: numeric to text in choices
   `value`/`filter`, `publishable` populated, leading/trailing whitespace trimmed, CRLF to LF.
2. **The fix list.** One or more of:
   - the team testing/feedback document (layout in `references/testing_feedback_TEMPLATE.md`; only
     rows whose Status is `Open`; re-read statuses every time, the team edits them; after the patch
     write `Fixed in vN`, `By design`, `Deferred` or `Paper-side` plus a one-line Response, and never
     touch the tester's own columns);
   - the deferred-patches backlog (`00_reference/<round>_deferred_patches.md`), applied only when
     the user says the batch is open;
   - a DIFF-mode findings table (`findings_NN.md`);
   - ad-hoc user requests in the conversation.
3. **The form's attachments folder** (`07_capi_attachments/<n>_<name>/`) so pulldata columns can be
   checked.
4. **The currently deployed `version` value.** Ask the user or read it from the download. The new
   version must be strictly greater than what the server has, not than what the local file has.

If (1) is missing, stop. Patching a stale base silently reverts the team's console edits.

---

## Phase P0 — Intake (silent)

- Resolve `<instrument>` to its register row; identify the download path and the deployed form_id and version.
- **Back up first**: copy the download to `99_archive/<instr>_backup_<tag>_<YYYYMMDD>/` before
  touching it.
- Load `survey`, `choices`, `settings`. Index rows by `name`, never by row number.
- Compare with the local mirror if one exists (benign-artifact normalization on). Report drift as a
  plan item, not as an error; the console edit may be intentional.
- Collect the fix list into one table (source, item id, field(s), what the reporter saw, what they
  asked for).

## Phase P1 — Plan table (ALWAYS shown, ALWAYS requires sign-off)

One row per fix-list item:

| # | Source | Field(s) | Change type | Action | Translation | Preload impact | Notes |
|---|---|---|---|---|---|---|---|
| 1 | Testing doc R7 | E01c | CONSTRAINT | apply: `.>0` to `.>=0 or .=-99` | hint PT = draft, flag (per brief policy) | none | paper says 0 if under 1 yr |
| 2 | Backlog | lang, lang_2 | CHOICES | apply: append 9 Chimakonde, 10 Kimwani | proper nouns, same in both | none | 8 already = English, do not reuse |
| 3 | Testing doc R9 | C04.2 | — | **by design** | — | — | count-driven roster, explain to team |
| 4 | Diff findings | B07 | NEW FIELD | **flag to PI** | — | — | changes eligibility rule; needs decision |
| 5 | Console drift | note_b1 | LABEL | keep console version | — | — | team edited on server 06-20 |

Change types: `LABEL`, `HINT`, `CONSTRAINT`, `RELEVANCE`, `CHOICES`, `NEW FIELD`, `DELETE FIELD`,
`MOVE`, `TYPE`, `CALC`, `PRELOAD`, `SETTINGS`.
Actions: `apply`, `by design` (recorded, not changed), `defer` (goes to the backlog), `flag to PI`
(needs a design decision), `keep console version`.

Below the table state:
- **Versioning decision** (see P2): proposed `version` stamp and whether form_id changes.
- **Cross-form check**: if a fix is a pattern (a decimal/comma trap, a `.>0` that should allow 0, a
  missing code prefix, a list that exists in several forms), say which other forms you scanned and
  whether they have the same problem. Patterns get fixed in every form in one batch.
- **Attachment consequences**: any pulldata column added or renamed, any media added.

Stop. Ask: *"Plan above. Approve, edit, or drop rows before I build?"* Do not fold in unrelated
backlog items the user did not open; list them as a reminder instead.

## Phase P2 — Versioning rule

- **`version` always increases**, even for a one-character label fix, even for an attachment-only
  redeploy. SurveyCTO will not push a form or its replaced attachments to devices otherwise. Stamp
  format `YYMMDDNNNN` (date plus 4-digit sequence); it must exceed the value **on the server**.
- **`form_id` stays the same by default.** Uploading with the same form_id updates the deployed form
  in place and keeps every server binding (enumerator dataset, attachments, publishing). Bump the
  form_id only when the user wants a **new data stream**: a new round, a structurally incompatible
  change, or a deliberate split of test versus production data. A form_id bump means every binding
  must be re-attached and mid-collection data lands in two datasets. Say so in the plan.
- Batch patches across several forms use **one shared version stamp**.
- Record the decision and the before/after values in the CHANGELOG entry.

## Phase P3 — Build (patch script)

1. Script path: `05_python_scripts/_patch_<short_name>_<tag>_<YYYYMMDD>.py`. One script per patch; the
   docstring lists the plan rows it implements. Scripts chain: each reads the previous deployed
   state, so the history is the sequence of patch scripts plus the CHANGELOG, not one master build.
2. Edit by field **name**, never by row index; insert relative to a named anchor.
3. **Read, edit and write through `scripts/xlsform_io.py`** (`XLSFormBook`): it rebuilds every sheet
   from row lists (openpyxl's `cell(row, col, value=None)` is a no-op and leaves stale values behind)
   and then reapplies the base workbook's formatting by field name, so fills, colored fonts, widths,
   hidden columns, frozen panes, conditional-formatting rules and validation dropdowns survive; an
   inserted row inherits the style of the row above it. A server download carries no formatting, so
   this matters when the brief §8 names a local formatted master as the authoring copy.
4. Extending a choice list: append new values, never renumber, check the value is not already used
   in that list (a collision is a hard fail), put proper nouns identically in `label:ENG` and
   `label`.
5. New or reworded text: keep the question-code prefix in every language column. Source-language
   text comes from the user or the paper. For the other language(s), **follow the translation policy
   the brief chose at project start** (section "Translation policy" of
   `capi_project_brief.md`). The brief picks one of: (a) *halt* — never write a translation, ask the
   user; (b) *draft and flag* — write a rendering, mark it in the plan's Translation column and in
   the translation-flags file for the professional reviewer; (c) *reuse only* — copy from the paper
   or a sibling form, otherwise halt. Whatever the policy, never write a translation without
   recording that you did.
6. New fields copy the sibling convention for their type (phone constraint, name regex, don't-know
   code and exclusive-DK constraint, required flag) from the brief.
7. Settings: bump `version`; form_id per P2; do not touch `form_title` or `default_language`.
8. Do not edit the paper docx. Paper-side errors go to the translator/PI list.

## Phase P4 — Validate (mandatory, print PASS/FAIL per check)

Run `scripts/validate_xlsform.py <patched.xlsx> --base <download.xlsx> --attachments <folder>`
(or equivalent inline checks):

1. **Intended-diff only**: every changed cell maps to an approved plan row; list any that do not.
2. Group and repeat balance (begin/end pairs, nesting).
3. Every `${name}` in relevance, constraint, calculation, label and repeat_count resolves to an
   existing field.
4. Every `select_one` / `select_multiple <list>` resolves to a choices list; no duplicate values in
   a list.
5. Bilingual completeness on visible rows (`label:ENG` and `label` both present, including new
   hints and constraint messages).
6. No stale cells (compare row lengths and trailing columns with the base).
7. `pulldata('<file>','<col>',...)`: file exists in the attachments folder, column exists in its
   header, key column exists.
8. Traps: `regex()` patterns are whole-string anchored; `.>0` on a field whose paper allows 0;
   decimal hints written with a comma; `translate()` or `jr:choice-name()` used on pulldata output.
9. `version` is greater than the server value; form_id matches the P2 decision.

A FAIL means fix the script and rerun. Never report success with a FAIL on the list.

## Phase P5 — Log and hand off

1. **CHANGELOG.md** entry (newest on top):

```
## Instrument <N> — <name> — <round> patch <tag> — <YYYY-MM-DD>
form_id <id> (unchanged | bumped from <old>); version <old> -> <new>
### Applied
- <field>: <what changed> (plan row #, source)
### Recorded as by-design / deferred / flagged to PI
- ...
### Deploy
- upload as update to existing form | new form
- attachments: none changed | re-upload <files> (version bump covers the push)
- form_id bumped -> re-attach enumerator dataset + all attachments
### Translation review
- <fields with drafted non-source-language text, per the brief's policy>
### Built from
- Base: <download path> . Backup: 99_archive/<...> . Script: 05_python_scripts/_patch_...py
```

2. Preload spec (`CAPI_PRELOAD_UPDATES_NEEDED.md` or the brief's equivalent) if any pulldata
   column changed.
3. Tick applied items in the backlog / feedback doc; add deferred ones.
4. Translation-flags file (if the brief's policy is "draft and flag"): append the fields carrying
   drafted text in the non-source language.
5. Copy the patched xlsx into the deploy mirror folder (`v<n>_deployed/`) and update that folder's
   README version table. **Rule for the team: every console edit or upload means the xlsx comes back
   into this folder the same day**, otherwise the mirror lies.

Report to the user: output path, version before and after, form_id decision, the deploy checklist,
what needs translator review, and the reminder list of backlog items you did *not* touch.

---

## Deploy checklist (hand this to whoever uploads)

- [ ] Upload the xlsx **as an update to the existing form** (same form_id), or as a new form if
      form_id was bumped.
- [ ] If any attachment changed: re-upload it. The version bump in the form is what pushes it to
      devices.
- [ ] If form_id changed: re-attach the enumerator dataset (must be an ENUMERATORS-type dataset, not
      a plain server dataset) and every attachment; re-create publishing and exports.
- [ ] Shared CSVs stay byte-identical across every form's attachment subfolder.
- [ ] On devices: delete the form and *Get Blank Form* to force a clean re-sync; confirm large
      media zips finished downloading.
- [ ] Download the deployed xlsx back into the deploy mirror folder.

---

## Batch patches (same fix, many forms)

When one fix applies to several forms (a shared choice list, a consent text, a language list, a
version-only bump to push refreshed attachments):

- one script loops over the forms, same plan table with a `Forms` column, one shared version stamp;
- validate each form independently;
- one CHANGELOG entry listing every form and its before/after version;
- forms that were scanned and found clean are listed as "verified, no change".

---

## Rules specific to PATCH

| # | Rule |
|---|---|
| P1 | The base is today's server download. Never the previous script's output. |
| P2 | Back up before editing. One backup folder per patch. |
| P3 | Plan table sign-off before any edit. |
| P4 | `version` always increases and must exceed the server value. form_id changes only by explicit decision. |
| P5 | Intended-diff-only validation: no unplanned cell changes. |
| P6 | Never renumber or reuse choice values. Append only. |
| P7 | "By design" and "deferred" outcomes are written down, not dropped. |
| P8 | Paper docx are never edited from this mode. |
| P9 | The deploy mirror is updated the same day as the upload. |

## Failure modes

- **Server ahead of mirror** (console edits): treat as plan rows, decide keep or override with the
  user.
- **Version rejected on upload**: the stamp was not above the server's; ask for the server value and
  restamp. Do not guess by adding 1 to the local file.
- **Attachment replaced but devices still show old data**: version was not bumped.
- **"No enumerator dataset attached"** after a form_id bump: re-attach; check it is an
  ENUMERATORS-type dataset.
- **Choice value already in use** (wanting `8` for a new language when `8` = English): stop, pick
  the next free value, record it.
- **Locked file** (open in Excel, OneDrive sync): copy to temp with retries for reading; for
  writing, ask the user to close the file.
