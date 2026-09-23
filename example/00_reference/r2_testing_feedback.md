# DEMO Primary School Survey — Round 2 testing feedback

Filled from `references/testing_feedback_TEMPLATE.md`. Fictional rows for the starter example: PATCH mode reads
the `Open` rows below as its fix list (see `example/README.md`, step 4).

## Project

| Item | Value |
|---|---|
| Project and round | DEMO Primary School Survey — Round 2 (2026) |
| Server / test link (Design tab) | `demo.surveycto.com` (placeholder; nothing is deployed from the example) |
| Test enumerator code(s) | 1, 2 (`enumerators_demo.csv`) |
| Test cluster / school code(s) | 9997, 9998, 9999 |
| Test respondent / household ID(s) | n/a |
| Who fills Response and Status | the skill operator |

---

## Instrument 1 — School head questionnaire

Form ID: `school_head_r2_v1`  ·  Test link: (placeholder)  ·  Test IDs specific to this form: school 9998 has no deputy preloaded

### Testing round: team testing  ·  version tested: `2601200001`  ·  date: 2026-01-22

| # | Question | What I saw | What I expected / suggest | Reviewer | Date | Status | Response |
|---|---|---|---|---|---|---|---|
| 1 | B03 | Hint says "0 if less than one year" in English only | Same hint in Portuguese | Tester A (fictional) | 2026-01-22 | Open | |
| 2 | E03 | "None of these" can be ticked together with other options | Should be exclusive | Tester A (fictional) | 2026-01-22 | Open | |
| 3 | C04 | Asked even after answering No to C02 | Should be skipped | Tester B (fictional) | 2026-01-22 | By design | C04 has relevance on C02 = Yes; with school 9997 the preload does not set C02, retest and answer C02 = No explicitly |
| 4 | A06 | Accepts digits in the respondent name | Letters only | Tester B (fictional) | 2026-01-22 | Fixed in v2601200001 | name regex added in Round-2 v1 |
