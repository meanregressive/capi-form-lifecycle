# CAPI project brief — DEMO Primary School Survey (fictional starter example)

Filled from `references/project_brief_TEMPLATE.md`. Every value below is invented for the
capi-form-lifecycle starter example; nothing refers to a real school, person or project.

## 1. Project identity

| Item | Value |
|---|---|
| Project name and short code | Demo Primary School Survey — `DEMO` |
| Study type and design | Fictional two-arm school-level RCT, 60 schools, randomized by school |
| Rounds and dates | Round 1 (2025) done; **this brief serves Round 2 (2026)**. A Round-1 XLSForm exists for instrument 1 only. |
| Sample geography | Two invented districts: Demo North, Demo South |
| Principal investigators | "PI (fictional)" |
| Implementing / funding partners | none (example) |
| Field data-collection firm | "Demo Field Team": deploys, tests and enumerates |
| SurveyCTO server | `demo.surveycto.com` (placeholder; nothing is uploaded from the example) |
| Languages | Design language English (`label:ENG`); field language Portuguese (un-tagged `label`; `default_language = POR`) |

**Ownership.** Form content: PI. Translation: field team's translator (role). Preload data: data
manager (role). SurveyCTO deployment: field team. Data cleaning: data manager. Skill operator: you.

## 2. Instrument register

| # | Short name | Instrument (full name) | Respondent | Prior-round form_id (blank if none) | Built from | Current form_id | File stem | Status | Notes |
|---|---|---|---|---|---|---|---|---|---|
| 1 | head | School head questionnaire | School head | `school_head_r1_v1` | prior round | | `1_school_head` | to build (Round 2) | the worked example |
| 2 | teacher | Teacher questionnaire | Grade-3 teacher | | scratch | | `2_teacher` | not included in the example | listed to show a from-scratch row |

## 3. Raw materials and where they live

| Material | Location | Naming / dating convention | Notes |
|---|---|---|---|
| Paper questionnaires, EN | `01_input_paper_EN/20260105/` | leading instrument number; date subfolder | EN docx carries a PI comment thread and one tracked change |
| Paper questionnaires, PT | `02_input_paper_PT/20260105/` | same | PT text is a machine draft for demo purposes |
| Prior-round XLSForms (if any) | `03_input_prior_capi/` | `<stem>_r1_v1.xlsx` | never edited |
| Current-round XLSForms | `04_output_r2_capi/v<n>_draft/` and `v<n>_deployed/` | | mirror updated same day as any upload |
| Preload CSVs and media | `07_capi_attachments/1_school_head/` | one subfolder per instrument | `upload_school.csv`: header + 3 test rows |
| Server-side datasets | `08_SurveyCTO_datasets/` | | `enumerators_demo.csv` (id, name, users) |
| Prior-round data for preloads | none in the example | | |
| Scripts | `05_python_scripts/` | build / patch scripts land here; `paper_configs/` holds PAPER-mode JSON | |

**Data-governance boundary.** The example contains no personal data. The rule the example
demonstrates: the skill reads preload CSVs by header row only, never opens rows of survey data, and
asks for a variable list rather than a data extract. No folder in the example is off-limits.

## 4. Authority order (which source wins)

| Disagreement about | Winner | Reasoning | Agreed |
|---|---|---|---|
| Question content, order, options, skip logic | EN paper | PIs sign the EN paper | 2026-01-02, PI |
| Response codes, list names, ID conventions, appearance | Round-1 XLSForm | comparability with Round-1 data | 2026-01-02, PI |
| Text of Portuguese | PT paper | translator's text | 2026-01-02, PI |
| PT change with no matching EN change | flag, never adopt | content is decided in EN | 2026-01-02, PI |
| Form vs paper after console edits | the form; paper regenerated from it | | 2026-01-02, PI |
| Team feedback vs paper | paper unless PI agrees | | 2026-01-02, PI |

**Round lifecycle.** Paper → BUILD → deploy → PATCH cycles → PAPER regenerated for training → PATCH →
PAPER regenerated before final deploy. Regeneration tool: `scripts/paper_from_xlsform.py` with
`05_python_scripts/paper_configs/`. Skip-logic prose is added by the field team's trainer.

## 5. Translation policy

**Reuse first, then draft and flag** (chosen 2026-01-02, PI). Look for existing PT text (PT paper,
sibling form, Round-1 form, glossary) first; only a genuinely new item may be drafted, marked in the
plan and appended to the flags file.

| Item | Value |
|---|---|
| Policy chosen | reuse first, then draft and flag |
| Professional reviewer (role) | field team's translator |
| Translation-flags file | `00_reference/r2_translation_flags.md` |
| Proper nouns and codes | written identically in both columns, no flag |
| Glossary | `00_reference/demo_glossary_EN_PT.md` |

## 6. Coding conventions

| Convention | Value | Origin |
|---|---|---|
| Yes/No list values | `yesno`: 1 = Yes, 0 = No (paper shows 1/2: display only) | Round-1 form |
| "Other, specify" value | -77 (paper shows 99: display only) | Round-1 form |
| "Don't know" value | -88 | Round-1 form |
| "Refused" value | -66 | Round-1 form |
| Numeric sentinel (phones, ages, counts) | -99 | Round-1 form |
| Exclusive option in select_multiple | `not(selected(., '<code>')) or count-selected(.)=1`, bilingual message | Round-1 E03 |
| Question-code prefix in labels | yes, both languages (`A01. …`) | Round-1 form |
| Label columns | `label:ENG` (EN), `label` (PT) | Round-1 form |
| `default_language` | `POR` | Round-1 form |
| Phone constraint | `(.>=100000000 and .<=999999999) or .=-99` (fictional 9-digit range) | Round-1 B04 |
| Person-name constraint | `regex(., "^[A-Za-zÀ-ÿ '’.-]+$")` | Round-1 A06 |
| Age / date rules | age allows -99; years-in-post ≤ age unless age = -99 | Round-1 B02/B03 |
| Decimal separator | period | |
| Required by default? | yes for content fields | |
| Repeats vs chained prompts | repeat groups | Round-1 D02 |
| Paper display codes that are NOT data codes | 1/2 for Yes/No, 99 for Other | |

## 7. Identifier scheme

| Item | Value |
|---|---|
| Primary cluster key | `A01` (school ID), preload key column `school_id_key` |
| Person / role IDs | role-based: `concat(${A01}, '<suffix>')` |
| Secondary keys | teachers: repeat index within the school |
| Two roles, one person | keep both IDs |
| Registry file | `00_reference/demo_id_suffix_registry.md` |

## 8. Versioning and deployment

| Item | Value |
|---|---|
| `version` stamp format | `YYMMDDNNNN`; always increases; must exceed the server value |
| `form_id` at the start of a round | `<prior_form_id>` with `_r1_` → `_r2_`: instrument 1 becomes `school_head_r2_v1` |
| `form_id` on patches within a round | unchanged; bump only for a new data stream |
| Deploy mirror folder | `04_output_r2_capi/v<n>_deployed/` with a README version table |
| Attachment layout | one subfolder per instrument; shared CSVs byte-identical |
| Enumerator dataset | `enumerators_demo` (ENUMERATORS-type); columns id, name, users |
| Test IDs | school IDs 9997–9999; phone 100000000 |
| Media limits | < 100 MB per file, 300 MB per form |
| Device re-sync rule | delete form + Get Blank Form after attachment changes |
| Same-day mirror rule | every console edit or upload → xlsx downloaded into the mirror |

## 9. Logs the project keeps

| Log | Path | What goes in | Written by |
|---|---|---|---|
| Change log | `CHANGELOG.md` | every build and patch, newest on top | skill |
| Flags and decisions | `00_reference/r2_flags_and_decisions.md` | design and translation decisions | skill |
| To-dos | `00_reference/r2_todos.md` | open actions by owner | skill |
| Deferred patches | `00_reference/r2_deferred_patches.md` | parked fixes | skill |
| Preload spec | `CAPI_PRELOAD_UPDATES_NEEDED.md` | columns each form needs | skill |
| Translation flags | `00_reference/r2_translation_flags.md` | drafted PT text | skill |
| Session file changes | `SESSION_FILE_CHANGES.md` | every file touched | skill |
| Team testing feedback | not in the example | | team |

## 10. Folder layout (as it actually is)

```
example/
  00_reference/            this brief, ID registry, glossary, (logs the skill creates)
  01_input_paper_EN/       Round-2 EN paper, dated subfolder
  02_input_paper_PT/       Round-2 PT paper, dated subfolder
  03_input_prior_capi/     Round-1 XLSForm (never edited)
  04_output_r2_capi/       BUILD and PATCH outputs, v<n>_draft / v<n>_deployed
  05_python_scripts/       build/patch scripts the skill writes; paper_configs/ for PAPER mode
  06_output_paper_FROMCAPI/ regenerated paper, one subfolder per language tag
  07_capi_attachments/     preload CSVs per instrument
  08_SurveyCTO_datasets/   enumerator dataset definition
  CHANGELOG.md, SESSION_FILE_CHANGES.md
```

## 11. Known traps for this project

- `.>0` rejects 0 where the paper allows "0 if less than a year" (B03).
- select_multiple lists with a "none" option need the exclusive constraint (E03).
- Years-in-post constraint must tolerate age = -99 (B03).
- A `village` column is referenced by the Round-2 paper (A04) but is not in `upload_school.csv`:
  the build must flag it for the data manager, never invent it.

## 12. History of this brief

| Date | Change | By |
|---|---|---|
| 2026-01-02 | created from template for the starter example | skill author |
