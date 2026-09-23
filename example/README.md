# Starter example — a fictional two-round school survey

A complete, tiny project laid out the way the skill expects, so you can run all four modes in a few
minutes without touching real data. Everything here is invented: no real school, person, phone
number or study. The Portuguese text is a machine draft written for the demo, not a professional
translation.

The project is "DEMO Primary School Survey", entering **Round 2**. Instrument 1 (school head
questionnaire) has a Round-1 XLSForm and a revised Round-2 paper questionnaire in English and
Portuguese. The paper differs from the form in six planted ways (listed at the bottom), and the EN
paper carries a PI comment thread and a tracked change, so DIFF and BUILD have real work to do.

## Layout

| Path | What it is |
|---|---|
| `00_reference/capi_project_brief.md` | the filled-in brief: every section the skill needs, with dated decisions |
| `00_reference/demo_id_suffix_registry.md`, `demo_glossary_EN_PT.md` | the two small companions the brief names |
| `00_reference/r2_testing_feedback.md` | a filled testing-feedback sheet (two `Open` rows, one `By design`, one fixed) for the PATCH walkthrough |
| `01_input_paper_EN/20260105/1_school_head_EN.docx` | Round-2 paper, EN, with a comment thread on B06 and a tracked change on C01 |
| `01_input_paper_EN/20260105/1_school_head_EN.pdf` | the same paper printed to PDF (Word, tracked changes shown as final), to try the PDF fallback |
| `02_input_paper_PT/20260105/1_school_head_PT.docx` | Round-2 paper, PT, with one planted content misalignment (C04) |
| `03_input_prior_capi/1_school_head_r1_v1.xlsx` | the Round-1 XLSForm (never edited), built on SurveyCTO's basic template: its sheets, columns, widths, frozen header and 31 type-coloring rules, plus the marks a team adds while authoring (yellow 'translation pending' fills, a red-font hint, a validation dropdown, a hidden column, tab color, one extra rule) |
| `_template/empty_instrument.xlsx` | SurveyCTO's basic template form (`empty_instrument.xlsx`), the starting workbook the DIME Wiki's [Questionnaire Programming](https://dimewiki.worldbank.org/Questionnaire_Programming) page points to ([Google Sheets copy](https://docs.google.com/spreadsheets/d/1JroFCu0HqnPsZwt0jAYa7o311VeOZFc6FRYfET4wAsA)). Copyright Dobility, Inc.; included unchanged as the base the generator builds on |
| `07_capi_attachments/1_school_head/upload_school.csv` | preload: header + 3 test schools |
| `08_SurveyCTO_datasets/enumerators_demo.csv` | enumerator dataset rows |
| `05_python_scripts/paper_configs/` | PAPER-mode JSON config for instrument 1 and a `batch.json` |
| `04_output_r2_capi/`, `CHANGELOG.md`, `SESSION_FILE_CHANGES.md` | empty; the modes write here |
| `expected/` | what the scripts and DIFF should produce (dumps, paper docx, `findings_01.md`) |
| `_generate_example.py` | writes every file above except this README, `expected/findings_01.md` and `00_reference/r2_testing_feedback.md`; re-run after edits |
| `run_checks.py` | runs the skill's scripts against the example and fails loudly if anything regressed |

## Try it

Open your agent (for example Claude Code) with `example/` as the working folder, so the brief is at
`00_reference/capi_project_brief.md`. Then:

1. `/capi-form-lifecycle diff head` — should reproduce the six rows in `expected/findings_01.md`.
2. `/capi-form-lifecycle build head --dry-run` — Phase B table only; compare with the same six rows.
3. `/capi-form-lifecycle build head` — signs off, then writes `04_output_r2_capi/v1_draft/1_school_head_r2_v1.xlsx`
   and a build script under `05_python_scripts/`. Expect the A04 village column to be flagged, not invented.
4. `/capi-form-lifecycle patch head` — pretend the built form is deployed (copy it to `v1_deployed/`).
   The fix list is `00_reference/r2_testing_feedback.md`: expect a plan table with its two `Open` rows
   (B03 hint in PT, E03 exclusive option) and the `By design` row left alone.
5. `/capi-form-lifecycle paper head` — regenerates the paper from a form into `06_output_paper_FROMCAPI/EN` and `/PT`.

Scripts on their own, from the repository root:

```bash
python scripts/validate_xlsform.py example/03_input_prior_capi/1_school_head_r1_v1.xlsx --attachments example/07_capi_attachments/1_school_head
python scripts/extract_for_diff.py --out example/expected/dumps --name en_01 example/01_input_paper_EN/20260105/1_school_head_EN.docx
python scripts/paper_from_xlsform.py --batch example/05_python_scripts/paper_configs/batch.json
python scripts/pdf_to_docx.py example/01_input_paper_EN/20260105/1_school_head_EN.pdf --out /tmp/1_school_head_EN_FROMPDF.docx
python example/run_checks.py
```

## The six planted differences

| Q | Type | Detail |
|---|---|---|
| A04 | NEW | village, prefilled; needs a `village` column that `upload_school.csv` does not have |
| B06 | NEW | phone has internet access; the PI comment thread asks for it and a reply agrees |
| C01 | WORDING | "enrolled" becomes "enrolled this school year", as a tracked change |
| C03 | RESPONSE-OPTIONS | meal provider list gains "Parents' association" |
| C04 | LANG-MISALIGN | PT paper says "por mês" where EN says "per week"; form is right |
| E02 | DELETED | council meetings last term is gone from the paper |

The Round-1 form is built on SurveyCTO's basic template form (`empty_instrument.xlsx`), the starting workbook the DIME Wiki's [Questionnaire Programming](https://dimewiki.worldbank.org/Questionnaire_Programming) page points to ([Google Sheets copy](https://docs.google.com/spreadsheets/d/1JroFCu0HqnPsZwt0jAYa7o311VeOZFc6FRYfET4wAsA)), so it carries the color coding by field type that most SurveyCTO teams work with; `run_checks.py` verifies that a patch through `xlsform_io` keeps every one of those rules.

Deliberately present and **not** to be flagged: paper shows Yes=1/No=2 and Other=99 while the form
uses 1/0 and -77; skip arrows on C02 and E01 match the form's relevance; plumbing rows.
