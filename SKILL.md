---
name: capi-form-lifecycle
description: Build, patch, audit and paper-regenerate bilingual SurveyCTO CAPI forms for a survey round, first or later (paper questionnaire <-> XLSForm <-> prior-round form when one exists). Four modes: build (new round from paper + prior XLSForm), patch (edit a deployed form, ship the next version), diff (CAPI vs paper gap report), paper (regenerate the paper questionnaire from the form for training or circulating for feedback). Project specifics live in the project's 00_reference/capi_project_brief.md.
---

# /capi-form-lifecycle — SurveyCTO CAPI forms across survey rounds

One skill, four modes, one project brief. The skill describes the **method** for taking a survey
instrument through a data-collection round: build the electronic form from the paper questionnaire
(and the prior round's form, when there is one), patch it while it is deployed, audit it against the
paper, and regenerate the paper from it. A first round with no prior form is a supported case, not
an exception. The **brief** in each project describes the project: instruments, conventions, authority
order, translation policy, deployment rules. The method is generic; everything project-specific is read
from the brief.

Target platform: SurveyCTO (XLSForm). Most of the method works on any ODK-based platform; the deployment checklist and a few formula pitfalls are specific to SurveyCTO. Those pitfalls are checked as part of the full workflow.

## Step 0 — read the project brief

Every project using this skill has `<project>/00_reference/capi_project_brief.md`, filled from
`references/project_brief_TEMPLATE.md`. It fixes, with dates: the instrument register; where raw
materials live and the data-governance boundary (what data the skill is given access to read); which source is
authoritative for content, for codes and for each language's text; the translation policy; the
coding conventions; the identifier scheme; versioning and deployment rules; the logs the project
keeps; the folder layout; any known project-specific traps.

**Nothing in this skill overrides the brief.** If the brief is missing or a section is blank, stop
and offer to fill it from the template before doing anything else. If the brief and a mode file
disagree, follow the brief and tell the user the mode file needs updating.

## Step 1 — pick the mode

| Invocation | Mode file | Use when |
|---|---|---|
| `/capi-form-lifecycle build <instrument> [--dry-run] [--from-scratch]` | `modes/build.md` | Start of a round: paper questionnaires + prior-round XLSForm (if any) → this round's v1 form. Once per instrument per round. Nothing is on the server yet. |
| `/capi-form-lifecycle patch <instrument>` | `modes/patch.md` | The form is on the server; apply a fix list (testing feedback, deferred backlog, diff findings); ship the next version. The everyday mode. |
| `/capi-form-lifecycle diff <instrument>` | `modes/diff.md` | Report where the form and the paper (or two versions of the form) disagree. Read-only. Its findings feed BUILD or PATCH. |
| `/capi-form-lifecycle paper <instrument or all>` | `modes/paper.md` | Regenerate the paper questionnaire (every language) from the deployed form, for training and for the record. Never edits the form. |

**`<instrument>` is resolved against the brief's instrument register (§2)**, which gives every
instrument a number, a short name and a full name. Any of the three is accepted, case-insensitive:
`patch 9`, `patch hh`, `patch household`, `patch "Household survey"`. If the argument matches nothing,
or matches more than one row, show the register and ask. With no mode: ask which. With no instrument:
show the register as a menu. Numbers are a shortcut for projects that number their instruments and
files consistently; they are not required. Input files are then located by the register's
`file stem` column (or, if blank, by the leading number in the filename).

## Round lifecycle (where each mode sits)

```
PI paper (all languages)
   └─ DIFF (optional: paper vs prior-round form, to size the work)
   └─ BUILD → v1 upload → attach bindings → team testing
         └─ PATCH ×n (testing feedback)
   PAPER (for enumerator training)
         └─ PATCH ×n (training, pilot)
   PAPER (before final deploy) → field period
         └─ PATCH ×n (field fixes, attachment refreshes)
   PAPER (end of round: archive what was actually asked) → next round's "prior-round form"
```

## Rules that hold in every mode

| # | Rule |
|---|---|
| G1 | The brief wins. Conventions, authority order and translation policy are read there, never assumed. |
| G2 | Never edit a prior-round form or a paper docx. Copy first; paper errors go to the translator/PI list. |
| G3 | Every change is made by a script kept in the project's scripts folder and is described in `CHANGELOG.md` the same day. |
| G4 | Personal data: read only de-identified inputs or header rows / variable descriptions, unless the brief says otherwise. Ask for a `describe` log rather than opening rows. |
| G5 | Sign-off before building: BUILD Phase B, PATCH plan table. No exceptions for "small" edits. |
| G6 | Validate before reporting; a FAIL is never reported as success. |
| G7 | Log every file touched in the project's session change log (for syncing to the team share). |
| G8 | Edit forms through `scripts/xlsform_io.py` (`XLSFormBook`): sheets are rebuilt from row lists (openpyxl's `cell(..., value=None)` does not clear a cell) and the workbook's formatting is carried back by field name, so fills, widths, frozen panes, rules and validations survive the edit. |

## Shared references

- `references/project_brief_TEMPLATE.md` — what a project must decide up front (12 sections).
- `references/deployment_checklist.md` — SurveyCTO server mechanics: enumerator dataset type, bindings on form_id change, version bump to push attachments, media limits, device re-sync, mirror rule.
- `references/xlsform_patterns.md` — copyable coding patterns: yes/no with skip, other-specify, preload calc, role-based and cascaded IDs, group wrapper, repeat, per-item repeat, exclusive option, confirm-preload-with-fallback, settings block.
- `references/tool_pitfalls.md` — openpyxl, OneDrive locks, anchored regex, `translate()` / `jr:choice-name()` on pulldata, Stata CRLF, Git Bash vs PowerShell for Stata.
- `references/diff_rubric.md` — authority order, do-not-flag list, flag types, severity.
- `references/what_the_project_can_change.md` — which settings a project adjusts through the brief or per-run flags, and which rules are the fixed method (with rule IDs); read when a user asks whether something can be changed.
- `scripts/extract_for_diff.py` — docx/xlsx → text dumps (tracked changes accepted, comments threaded, OneDrive-safe); `--compare a.xlsx b.xlsx` = version-vs-version cell diff with artifact folding.
- `scripts/build_diff_docx.py` — assembles the `findings_NN.md` files into one landscape Word report (executive summary + per-instrument tables, severity color-coded).
- `scripts/paper_from_xlsform.py` — PAPER-mode engine: XLSForm → landscape docx per language (code | question | response tables, grids/matrices for repeats, `[SHOW IF]`/`[CHECK]` markers); per-instrument knobs in a JSON config, `--batch` for all instruments.
- `scripts/pdf_to_docx.py` — fallback when the paper exists only as PDF: PDF → reviewable docx (tables kept) + conversion report; `extract_for_diff.py` invokes it for `.pdf` inputs and marks the dump `PDF-DERIVED`. No OCR.
- `scripts/xlsform_io.py` — the read-edit-write helper every build and patch script uses: `XLSFormBook(path)` → edit `rows('survey')` → `save(out)`; rebuild for correctness, formatting snapshot reapplied by field name (new rows inherit the row above); `carry` and `inspect` CLI.
- `scripts/validate_xlsform.py` — the PASS/FAIL checks BUILD and PATCH require (balance, refs, lists, bilingual, code prefix, pulldata header check, expression traps, settings/version, stale cells, intended-diff vs `--base` with `--planned`).

## Provenance

Developed by Deboleena Rakshit (IFPRI) from the CAPI work of a multi-round field survey, 2026. The
skill text records the method; the lessons that shaped it are folded into the references. Version
history is in `CHANGELOG.md`.
