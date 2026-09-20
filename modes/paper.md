# Mode: PAPER — regenerate the paper questionnaire from the deployed form

Use this mode whenever the form has moved ahead of the paper and people need to read it on paper:
before enumerator training, before piloting, before the final deploy, and at the end of a round to
archive what was actually asked. The brief's lifecycle rule (§4) says every round does this at least
once. The form is the source; the paper is the derived document. **This mode never edits a form.**

Invocation: `/capi-form-lifecycle paper <instrument>` or `/capi-form-lifecycle paper all`; `<instrument>` =
number, short name or full name from the brief's register (§2).

---

## Inputs

1. The deployed form(s): the XLSForm as downloaded from the server today (or the deploy mirror
   folder if the user confirms it is current). Not a draft, unless the user says the paper is for a
   draft review.
2. The engine `scripts/paper_from_xlsform.py` (shipped with the skill) and the project's
   per-instrument JSON configs (`05_python_scripts/paper_configs/<n>_<short_name>.json`, or the
   folder the brief §10 names): banner titles per language, `var_labels` (roster-name variables
   collapsed to a generic word), `note_markers`, `matrices` (repeat rendered as a matrix). A
   `batch.json` lists every instrument for `paper all`. A project that keeps its own engine
   records that in the brief §3 instead.
3. The output folder (brief §10; e.g. `06_output_paper_FROMCAPI/`). The engine writes one
   subfolder per language tag.

## Steps

1. **Confirm scope**: which instruments, which xlsx (path, form_id, version), which output folder,
   which version suffix the filenames should carry.
2. **Do not rewrite the engine.** Run it:
   `python scripts/paper_from_xlsform.py <form.xlsx> --out <dir> --base <stem> --langs "EN=label:ENG,PT=label" --config <instrument.json> --suffix FROMCAPI_v<n>`
   or `--batch <batch.json>` for every instrument. `--langs` maps each language tag to its label
   column (the brief §1 says which column is which language). For a new instrument with no
   config, copy the closest sibling's JSON and change only its titles and knobs; say that you did.
   A construct the engine cannot lay out is reported to the user, not patched into the engine
   for one instrument.
3. **Output**: one docx per language per instrument, `<out>/<TAG>/<base>_<TAG>_<suffix>.docx`.
   Never overwrite a previous version's file; the suffix changes. Add `--pdf` (or `"pdf": true` in
   `batch.json`) to also write a PDF beside each docx; this uses Microsoft Word and is skipped, with a
   message, where Word is not installed.
4. **Sanity checks** before handing over:
   - no `${variable}` leaked into question text (allowed only inside the `[SHOW IF: …]` /
     `[CHECK: …]` markers the engine writes for skip and constraint logic). Check this **cell by
     cell with python-docx**, not on the raw document XML: a marker and its variable can sit in
     different XML runs, and a raw-XML scan reports false leaks;
   - code column: paper code (bold), or the form's field name in gray italics when the label has
     no code but the name looks like one (`A07`, `C12c2`), or the `[CAPI CODE]` flag for rows that
     have neither (enumerator picker, preload calculates, comment boxes). Count the flags; a jump
     between runs means labels lost their code prefixes;
   - every visible field appears once, in form order, with its code prefix;
   - every language file has the same question count;
   - very wide tables, new question types, and repeats rendered as grids are listed for the
     user's eye.
5. **Human finish** (state it in the report, do not attempt it): the `[SHOW IF]` / `[CHECK]`
   markers are turned into skip-logic prose by a person; banners and titles live in the scripts, not
   in the form.
6. **Log**: CHANGELOG entry "Paper regenerated from <form_id> version <v> — <date>" listing the
   files written; session change log.

## Rules specific to PAPER

| # | Rule |
|---|---|
| R1 | The form is the source. Never edit the paper by hand to fix a question; fix the form and re-run. |
| R2 | Never modify the engine to suit one instrument; use the per-instrument JSON knobs. |
| R3 | Regenerate from the deployed version, and say which version the paper reflects (in the docx header and the CHANGELOG). |
| R4 | Every language present in the form gets a paper file; a language column that is blank for some fields is reported, not padded. |

## Failure modes

- **Glob picks the wrong file** (two files for one instrument in the version folder, e.g. a stale
  plain-named copy next to the deployed URL-encoded one): stop and ask which is of record.
- **Engine raises on a new construct** (a field-list matrix, a new appearance): report the field,
  do not patch the engine silently; the per-instrument JSON has knobs for layouts.
- **A language with no built-in UI strings** (the engine ships EN, PT, ES, FR): the engine warns and
  falls back to English for words like "Prefilled"; supply them under `strings.<TAG>` in the config.
- **Paper differs from what the team remembers**: run DIFF mode on the two form versions rather than
  arguing from the docx.
