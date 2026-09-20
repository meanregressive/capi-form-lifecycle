# Instrument 01 — School head questionnaire
SOURCES: CAPI=03_input_prior_capi/1_school_head_r1_v1.xlsx, version 2501150001, form_id school_head_r1_v1 ; EN=01_input_paper_EN/20260105/1_school_head_EN.docx (2026-01-05) ; PT=02_input_paper_PT/20260105/1_school_head_PT.docx (2026-01-05)

This is the findings table DIFF mode should produce for the starter example (Round-1 form vs the
Round-2 paper). Tracked changes in the EN paper are read as accepted. Written by hand as the
reference answer; your run may word things differently but should flag the same six rows.

## DIFFERENCES
| Q code | Section / topic | Diff type | Paper says (EN; note PT if differs) | CAPI has | Severity | Recommendation |
|---|---|---|---|---|---|---|
| A04 | A. Identification | MISSING-IN-CAPI | "A04. Village" prefilled | no A04; upload_school.csv has no village column | High | update CAPI: add pulldata calc; data manager must add the column (flag, do not invent) |
| B06 | B. Respondent | MISSING-IN-CAPI | "Does your phone have internet access?" Yes/No; PI comment thread says skip if no phone | no B06 | High | update CAPI: select_one yesno, relevance B04 not -99 (per comment thread, resolved by reply) |
| C01 | C. School | WORDING | "enrolled this school year" (tracked insertion) | "enrolled" (no reference period) | Med | update CAPI label in both languages |
| C03 | C. School feeding | RESPONSE-OPTIONS | adds "Parents' association" | list meal_provider has 1, 2, 3, -77 | High | update CAPI: append value 4, before -77; never renumber |
| C04 | C. School feeding | LANG-MISALIGN | EN "days per week"; PT "dias por mês" (per month) | form: per week in both languages | High | escalate to PI / translator: PT paper error, form is right; do not change the form |
| E02 | E. School council | EXTRA-IN-CAPI | question absent from the paper | "How many times did the council meet last term?" | Med | update CAPI: delete E02 after PI confirms the drop; prior-round data unaffected |

## SUMMARY
- Two new content questions in the paper (A04 prefilled village, B06 phone internet); A04 needs a new preload column the data manager must supply.
- One response-option addition (meal_provider value 4) and one deletion (E02).
- One meaning-changing wording edit (C01 reference period), delivered as a tracked change.
- One language misalignment on content (C04 week vs month): paper-side error in PT, escalate, form stays.
- Not flagged, by the rubric: paper Yes=1/No=2 vs form 1/0; paper Other=99 vs form -77; skip arrows vs relevance on C02/E01; the transparent `survey` wrapper group; plumbing rows (start, end, deviceid, t_index).
- Counts: High 4 / Med 2 / Low 0.
