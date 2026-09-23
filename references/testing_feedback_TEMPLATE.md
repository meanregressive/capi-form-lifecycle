# CAPI testing feedback — TEMPLATE

Copy this file into the project (the brief §9 names the path, e.g. `00_reference/<round>_testing_feedback.md`)
or paste it into a shared Google Doc / Word file (a `.docx` version sits next to this file). One section per
instrument. Testers add rows; the CAPI programmer fills **Response** and **Status**.

## How to fill it in

1. **One issue per row.** Two problems on the same question are two rows.
2. **Quote the question code** (`B06`, `C03_o`) as shown on the tablet or the paper. If there is none, describe where you were.
3. **Say what you saw, then what you expected.** "Skips to Section D when I answer No" / "should ask C03 first."
4. **Write the form version you tested** in the round header (it is on the form's opening screen or in the server's Design tab).
5. **Do not change another person's row.** Add a new row that refers to it (`see #4`).

## Status values (use exactly these)

| Status | Meaning |
|---|---|
| `Open` | not yet looked at, or looked at and still needs a change |
| `Fixed in vN` | changed in form version N (the programmer writes the version) |
| `By design` | the form is behaving as intended; the Response says why |
| `Deferred` | agreed change, parked for a later version (goes to the deferred-patches list) |
| `Paper-side` | the form is right; the paper questionnaire or translation needs the change |

**How the agent uses this file.** PATCH mode reads the table, takes every row whose Status is `Open` as a
fix-list item, and proposes a plan table for sign-off. After the patch it writes `Fixed in vN`, `By design`,
`Deferred` or `Paper-side` in Status and a one-line reason in Response. It never edits your other columns.

---

## Project

| Item | Value |
|---|---|
| Project and round | |
| Server / test link (Design tab) | |
| Test enumerator code(s) | |
| Test cluster / school code(s) | |
| Test respondent / household ID(s) | |
| Who fills Response and Status | |

---

## Instrument <N> — <full name>

Form ID: `<form_id>`  ·  Test link: <url>  ·  Test IDs specific to this form: <…>

### Testing round: <purpose, e.g. "team testing" / "training" / "pilot">  ·  version tested: `<version>`  ·  date: <YYYY-MM-DD>

| # | Question | What I saw | What I expected / suggest | Reviewer | Date | Status | Response |
|---|---|---|---|---|---|---|---|
| 1 | | | | | | Open | |
| 2 | | | | | | Open | |
| 3 | | | | | | Open | |

### Testing round: <next round>  ·  version tested: `<version>`  ·  date: <YYYY-MM-DD>

| # | Question | What I saw | What I expected / suggest | Reviewer | Date | Status | Response |
|---|---|---|---|---|---|---|---|
| 4 | | | | | | Open | |

---

## Instrument <N> — <full name>

Form ID: `<form_id>`  ·  Test link: <url>

### Testing round: <…>  ·  version tested: `<version>`  ·  date: <YYYY-MM-DD>

| # | Question | What I saw | What I expected / suggest | Reviewer | Date | Status | Response |
|---|---|---|---|---|---|---|---|
| 1 | | | | | | Open | |

---

*Worked example of three filled rows (delete when using):*

| # | Question | What I saw | What I expected / suggest | Reviewer | Date | Status | Response |
|---|---|---|---|---|---|---|---|
| 1 | B03 | Rejects 0 for years as head | Should accept 0 ("less than a year") | A. Tester | 2026-02-03 | Fixed in v2 | constraint now `.>=0` |
| 2 | C04 | Asked even when C02 = No | Should be skipped | A. Tester | 2026-02-03 | By design | C04 has relevance on C02 = Yes; retest with school 9998, the preload for 9997 sets C02 |
| 3 | E03 | "None of these" can be ticked with other options | Should be exclusive | B. Tester | 2026-02-04 | Open | |
