# DIFF rubric — comparing a CAPI form with the paper questionnaire

Read fully before comparing. Applies to `modes/diff.md`. Sources are the text dumps produced by
`scripts/extract_for_diff.py`: the form dump (`capi_NN.txt`), the source-language paper dump
(`en_NN.txt` or the project's tag) and one dump per other language. Tracked changes in the paper
dumps are already accepted; treat them as final.

The examples use one project's values (EN source language, PT field language, codes -77/-88/-99). Replace
with the brief's values; the brief §4 authority order and §6 conventions override anything here.

## 1. Authority order (drives the Recommendation column)

Default for a paper-first project, unless the brief says otherwise:

- **Source-language paper = final on question content and wording**: which questions exist, their
  order, meaning, response options, skip logic. Form disagrees → "update CAPI to match paper".
- **Form = final on response codes, list names, ID conventions, appearance**. Never recommend
  changing a code to match the paper's display numbering.
- **Each other-language paper = final on that language's text only.** If it disagrees with the
  source language on *content* (options, skip, meaning) that is a **language misalignment** →
  "translator/PI follow-up"; do not pick a winner.
- After deployment, if the brief says the form is the working version, the paper dump is compared
  as "what the PIs last approved", and EXTRA-IN-CAPI rows are expected rather than defects.

## 2. Do NOT flag (conventions and noise)

1. Yes/No coding: paper shows 1/2, form uses the project's values (1/0).
2. Other-specify and sentinel codes: paper 99 vs form -77; -88 / -99 conventions.
3. Plumbing rows with no paper counterpart: `start`, `end`, `deviceid`, `subscriberid`,
   `simserial`, `phonenumber`, `username`, `duration`, `calculate`/`calculate_here` rows,
   `once()`, `index()`, `position()`, `indexed-repeat()`, `concat()` ID fields, `pulldata()`
   preloads, `*_id`, `start_*/end_*/duration_*`, notes that are pure enumerator instructions,
   `_o` companions when the paper has "specify" inline.
4. HTML/formatting in labels (`<p>`, `<b>`, `&nbsp;`), punctuation, accents, capitalization,
   whitespace.
5. Skip notation: paper `>> Cxx` vs form `relevance` when the logic is the same.
6. One paper question split into several form fields for programming (a `_nopre` fallback, a
   phone split the paper also splits) unless it changes what is asked.
7. Presence or absence of the question-code prefix in a label.
8. Confirmation notes and availability gates the brief lists as standard blocks.

The brief may extend this list (project §11 traps that are by design).

## 3. DO flag — diff types

| Type | Meaning |
|---|---|
| `MISSING-IN-CAPI` | content question or option present in the source paper, absent from the form |
| `EXTRA-IN-CAPI` | content question in the form that the paper no longer/never has (not plumbing) |
| `RESPONSE-OPTIONS` | the set of options differs in content: added, removed, merged, reworded so meaning changes |
| `SKIP-LOGIC` | relevance/skip differs substantively: target, trigger value, a branch on one side only |
| `WORDING` | source-language label vs form label differ in meaning, scope, reference period or respondent |
| `LANG-MISALIGN` | languages disagree on content, or one language names the wrong entity/role; also form text of a language vs that language's paper differing substantively |
| `RENUMBERED / ORDER` | code changed or order changed in a way that affects flow or skips |
| `TYPE` | response type differs materially (single vs multi, numeric vs text, integer vs decimal) |
| `MISSING-DOC` | a whole side not provided; one row per missing side; never infer the missing side |

## 4. Alignment method

Align by **question code** (A01, B05a, C10a). The paper's left column is the code; the form uses it
as `name` and usually as a label prefix. Walk the source-language paper top to bottom as the spine;
then sweep the form for content fields the paper lacks; then check each other language against the
source; for version-vs-version, list cell diffs grouped by field. A code on one side only is
MISSING/EXTRA; a renamed code (paper `CD11`, form `D11`) is `RENUMBERED`, not both.

## 5. Severity

- **High** — changes the data collected or yields wrong/blank data: missing question, wrong option
  set, broken skip, wrong-role translation, type change that loses information.
- **Med** — meaning-changing wording or order unlikely to break data.
- **Low** — worth noting, no data consequence.

## 6. Output

`findings_NN.md` in the dumps folder, structure fixed by `modes/diff.md` Phase D2. No `|` inside
cells. Quote only the differing fragment. Summary: 3–8 bullets, counts by severity, and a plain
"essentially aligned" when true. Recommendations use four verbs: update CAPI · update paper
(translator/PI) · escalate to PI · no action (with reason).
