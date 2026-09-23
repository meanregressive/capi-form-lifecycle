# CAPI project brief — TEMPLATE

Copy this file to `<project root>/00_reference/capi_project_brief.md` and fill every section before
the first `/capi-form-lifecycle` run. The skill reads the brief at the start of every mode and treats
it as the authority for anything project-specific. **If a section is blank, the skill stops and
asks.** Keep decisions dated and attributed; a future round will ask "why did we do it this way?"

## Filling the brief with your agent (the default way)

You do not have to fill this template by hand. Open your agent in the project folder and say
"set up the CAPI project brief" (or run any `/capi-form-lifecycle` mode; the skill offers this step
itself when the brief is missing or has blank sections). The agent then works section by section:

1. **Infers first, asks second.** From the files already in the project it proposes: the instrument
   register (paper and prior-round form filenames), the coding conventions, label columns,
   `default_language`, constraint patterns and form_id pattern (from a prior-round XLSForm), the folder
   layout (from the real tree), test IDs (from a preload CSV header row). You confirm or correct each
   guess instead of typing it.
2. **Asks only what files cannot tell it**, one section at a time, a few questions per turn: who
   decides content, translation, preloads and deployment (§1); the authority order (§4); the
   translation policy (§5); the data-governance boundary and folders it must never open (§3); the
   server and dataset names (§8); who fills the testing-feedback sheet (§9).
3. **Writes every answer as a dated, attributed decision** ("agreed 2026-01-02, PI"), so a later round
   can see why.
4. **Marks what you cannot answer yet as `[CONFIRM: <question>]`** instead of leaving the cell blank. A
   blank section stops the skill; a `[CONFIRM]` marker lets work proceed and shows up in every report
   until it is resolved.
5. **Reads the result back to you** as a short summary before the first BUILD or PATCH, and edits the
   brief in place, with a row in §12, whenever a decision changes later.

Filling it manually works too: copy the file, fill every section, delete the options you did not
choose. The rest of this template is the same either way.

Conventions for filling it in:
- Write decisions as facts with a date: "Yes/No coded 1/0 (agreed 2026-05-21, PI)".
- Where the template offers options, delete the ones not chosen; leave the reasoning.
- Anything that changes mid-project is edited in place with the new date, and the old value is kept
  in a "History" line under it.

---

## 1. Project identity

| Item | Value |
|---|---|
| Project name and short code | |
| Study type and design | e.g. cluster RCT, arms, unit of randomization |
| Rounds and dates | e.g. baseline YYYY / midline YYYY / endline YYYY, or a single round; which round this brief serves; whether a prior-round form exists at all |
| Sample geography | |
| Principal investigators | |
| Implementing / funding partners | |
| Field data-collection firm | who deploys, who tests, who enumerates |
| SurveyCTO server | `<subdomain>.surveycto.com`; who holds admin |
| Languages | source language of design; field language(s); which is the un-tagged `label` column |

**Ownership (who decides what).** One line each: form design / questionnaire content; translation;
preload data production; SurveyCTO deployment; data cleaning; the CAPI skill operator.

## 2. Instrument register

| # | Short name | Instrument (full name) | Respondent | Prior-round form_id (blank if none) | Built from | Current form_id | File stem | Status | Notes |
|---|---|---|---|---|---|---|---|---|---|
| | | | | | prior round / scratch / team-built | | | | |

The skill accepts **any of number, short name or full name** as the `<instrument>` argument, so give
every row a unique short name (`hh`, `teacher`, `council`). `File stem` is the leading part shared by
the instrument's paper docx, prior-round xlsx and output files (e.g. `9_household_survey`); leave it
blank if the files simply start with the number. Numbering is optional but, if used, state gaps and
dropped instruments explicitly ("#7 and #8 do not exist; #11 dropped at midline"). Instruments the
team builds outside the skill still get a row.

## 3. Raw materials and where they live

| Material | Location | Naming / dating convention | Notes |
|---|---|---|---|
| Paper questionnaires, source language | `01_input_paper_<LANG>/<YYYYMMDD>/` | leading instrument number; date subfolder per PI revision | docx preferred (comments and tracked changes are read); a PDF-only paper goes through `scripts/pdf_to_docx.py` and human review first |
| Paper questionnaires, other language(s) | `02_input_paper_<LANG>/<YYYYMMDD>/` | | may lag the source language; the brief says how to treat lag |
| Prior-round XLSForms (if any) | `03_input_prior_capi/` | | never edited; empty for a first round, and BUILD then runs from scratch |
| Current-round XLSForms | `04_output_<round>_capi/v<n>_deployed/` (mirror of server) and `v<n>_draft/` | | mirror updated same day as any upload |
| Preload CSVs and media | `07_capi_attachments/<n>_<name>/` | one subfolder per instrument | column lists in the preload spec file |
| Server-side datasets | `08_SurveyCTO_datasets/` | | enumerator dataset definition |
| Prior-round data for building preloads | | | see data-governance boundary below |
| Scripts | `05_python_scripts/` (or as laid out in §10) | build / patch / diff / validators / preload builders / `paper_configs/*.json` | |

**Data-governance boundary.** Which folders hold personal data (names, phones, GPS, household
rosters) and what the skill operator may do with them. Name the folders the skill must never read.
Recommended default: the skill reads only de-identified inputs, or only header rows / variable names,
labels and descriptions; to validate a dataset it asks for a Stata `describe` log or a header-only
extract rather than opening rows. Say who runs any step that must touch identifying values.

## 4. Authority order (which source wins)

Fill the table for this round. Date every line.

| Disagreement about | Winner | Reasoning | Agreed |
|---|---|---|---|
| Question content, order, options, skip logic | e.g. source-language paper | | |
| Response codes, list names, ID conventions, appearance flags | e.g. prior-round XLSForm | | |
| Text of language X | e.g. the language-X paper | | |
| Other-language change with no matching source-language change | e.g. flag, never adopt | | |
| Form vs paper after the form is deployed and edited on the console | e.g. the form; paper regenerated from it | | |
| Team feedback vs paper | e.g. paper unless PI agrees | | |

Also record the **round lifecycle**: paper → build → deploy → patch cycles → **paper regenerated from
the form for training** → patches → paper regenerated before final deploy. State the tool that
regenerates the paper and who adds skip-logic prose afterwards.

## 5. Translation policy

Choose one and delete the others. Record who chose it and when.

- **Halt.** The skill never writes text in a non-source language. Missing text stops the build until
  a human supplies it.
- **Draft and flag.** The skill may write a rendering when no paper or sibling text exists, must mark
  it in the plan and append the field to `<translation-flags file>`, and a named professional
  reviewer clears the flags before production.
- **Reuse only.** The skill copies text only from the paper or from a sibling form; otherwise halts.
- **Reuse first, then draft and flag.** The skill always looks for an existing translation first (paper
  of that language, sibling form, prior-round form, glossary); only a genuinely new item with no
  translation offered may be drafted, and every drafted item is marked in the plan and in the flags
  file.

| Item | Value |
|---|---|
| Policy chosen | |
| Professional reviewer (role, not name) | |
| Translation-flags file | |
| Do translations of proper nouns and codes count? | e.g. list values that are proper nouns are written identically in both columns without a flag |
| Glossary | `00_reference/<project>_glossary_<L1>_<L2>.md` |

## 6. Coding conventions

| Convention | Value | Origin |
|---|---|---|
| Yes/No list values | | |
| "Other, specify" value | | |
| "Don't know" value | | |
| "Refused" value | | |
| Numeric sentinel for unknown / not applicable (phones, ages, counts) | | |
| Exclusive-option rule for select_multiple (none / DK / other) | constraint pattern + bilingual message | |
| Question-code prefix in displayed labels | yes / no; both languages? | |
| Label columns | `label:<TAG>` for language A, `label` for language B | |
| `default_language` in settings | | |
| Phone constraint | pattern + sentinel | |
| Person-name constraint | pattern; whole-string anchored | |
| Age / date rules | allow sentinel in age ranges; DOB default; age-DOB consistency | |
| Decimal separator in hints and examples | | |
| Required by default? | | |
| Repeat groups vs chained prompts | | |
| Paper display codes that are NOT data codes | e.g. paper 1/2 and 99 are formatting | |

## 7. Identifier scheme

| Item | Value |
|---|---|
| Primary cluster key (e.g. school ID) | field name, preload key column |
| Person / role IDs | role-based (`concat(key, suffix)`) or person-based; registry file |
| Secondary keys (households, pupils, members) | field names and key columns |
| Rule for a person holding two roles | |
| Registry file | `00_reference/<project>_id_suffix_registry.md` |

## 8. Versioning and deployment

| Item | Value |
|---|---|
| `version` stamp format | e.g. `YYMMDDNNNN`; always increases; must exceed the server value |
| `form_id` at the start of a round | e.g. `<prior_form_id>_<round>_v1`; first round or new instrument: `<short_name>_<round>_v1` |
| `form_id` on patches within a round | same form_id (in-place update) by default; bump only for a new data stream |
| Authoring copy | which xlsx is edited: the server download (default; carries no formatting) or a local formatted master whose fills/widths/rules must survive every edit (`xlsform_io` carries them) |
| Deploy mirror folder and README table | |
| Attachment layout | one subfolder per instrument; shared CSVs byte-identical across subfolders |
| Enumerator dataset | name(s) for test and production; must be an ENUMERATORS-type dataset; columns |
| Test IDs and test data filter | e.g. cluster IDs 9997–9999; phone `840000000` |
| Media limits | per-file and per-form size; bundle many files into one zip |
| Device re-sync rule | delete + Get Blank Form after attachment changes |
| Same-day mirror rule | every console edit or upload → xlsx downloaded back into the mirror |

## 9. Logs the project keeps

| Log | Path | What goes in | Written by |
|---|---|---|---|
| Change log | `CHANGELOG.md` | every build and patch, newest on top | skill |
| Flags and decisions | `00_reference/<round>_flags_and_decisions.md` | design and translation decisions with "risk if reversed" | skill |
| To-dos | `00_reference/<round>_todos.md` | open actions by owner | skill, ticked by humans |
| Deferred patches | `00_reference/<round>_deferred_patches.md` | fixes parked for a later version | skill |
| Preload spec | `CAPI_PRELOAD_UPDATES_NEEDED.md` | columns each form needs, test vs production values | skill |
| Translation flags | | fields carrying drafted text | skill |
| Attachment audit | | per-form dataset/column/encoding audit | skill |
| Session file changes | `SESSION_FILE_CHANGES.md` | every file touched, for syncing to the team share | skill |
| Team testing feedback | `00_reference/<round>_testing_feedback.md` (or a shared Google Doc / Word file), from `references/testing_feedback_TEMPLATE.md` | one row per issue with Status; only `Open` rows are actioned; the skill writes Response and Status | team |

## 10. Folder layout (as it actually is)

Paste the real top-level tree with one line of purpose each. The skill refuses to guess paths.

## 11. Known traps for this project

Seed from the previous round's deferred backlog and the skill's `tool_pitfalls.md`. Examples of
the kind of line that belongs here:

- required fields inside a `field-list` group whose relevance depends on a same-screen field do not
  enforce reliably; put them on their own screen;
- `.>0` rejects 0 where the paper allows "0 if less than a year";
- select_multiple lists with a "none" option need the exclusive constraint;
- years-worked constraints must reference the age field;
- `jr:choice-name()` and `translate()` on pulldata output break or cache wrongly.

## 12. History of this brief

| Date | Change | By |
|---|---|---|
| | created from template | |
