# capi-form-lifecycle

An [Agent Skill](https://agentskills.io/specification) for taking Computer-Assisted Personal Interviewing (CAPI) survey instruments through a
data-collection round on SurveyCTO: **build** the XLSForm from the paper questionnaire and the prior
round's form, **patch** it while it is deployed, **diff** it against the paper, and regenerate the
**paper** questionnaire from the form for training.

Developed by Deboleena Rakshit from July to September, 2026, at the International Food Policy Research Institute (IFPRI), assisted by Claude Code (Fable 5.1).

## Why

A CAPI survey round has to maintain version control and consistency across three streams of work: the paper questionnaire that investigators, research, or program teams sign off on; the electronic form(s) the tablets run; and, in the case of multiple rounds of data collection, any prior round's forms that data must stay comparable with (often across multiple languages in the context of most International Development work). 

A first (or "baseline") round has no prior form; the **build** mode then works from the paper questionnaire alone, and the CAPI Excel form it produces becomes the next round's (e.g., "midline" or "endline") reference. 

In practice, most of the work in a round is not in the initial CAPI form build, but the **patch** build of deployed forms. This refers to the process of implementing survey or item-level changes that can be minor (like tweaks to question wording and skip logic) or structural (like adding survey modules or plugins), based on testing or piloting performed by the research or program teams. 

This **patch** mode is complemented by a **diff** mode where any drifts or changes from the paper questionnaire are checked by the AI agent and verified by the user as intentional before redeploying the updated form. The **diff** mode is also invoked when checking differences between iteratively updated versions of the CAPI forms.

The final **paper** mode is an optional step where the AI agent uses this skill to generate a more reader-friendly PDF/text version of the CAPI Excel form to use during training for survey firms, or for circulating among research and program teams, doing away with the need for manually updating the paper version of the questionnaire after each iterative change to the form.

This skill encodes this multi-step workflow so an AI agent can run it with the same discipline every time: a plan table signed off by the user/CAPI programmer before any edit, a reproducible script for every change, validation before any success report, and a changelog entry the same day.

## What is in the box

| Path | Purpose |
|---|---|
| `SKILL.md` | Entry point: read the project brief, pick a mode, global rules |
| `modes/build.md` | New round: prior-round XLSForm + paper -> v1 form (five phases, sign-off gate) |
| `modes/patch.md` | Deployed form + fix list -> next version (plan table, versioning rule, validation) |
| `modes/diff.md` | Form vs paper, or version vs version, read-only findings table |
| `modes/paper.md` | Regenerate the paper questionnaire from the deployed form |
| `references/project_brief_TEMPLATE.md` | The 12 decisions a project fixes before the first run |
| `references/testing_feedback_TEMPLATE.md` / `.docx` | Feedback sheet for testing and piloting rounds: one row per issue, fixed status values; PATCH reads the `Open` rows |
| `references/deployment_checklist.md` | SurveyCTO server mechanics that break forms on tablets |
| `references/xlsform_patterns.md` | 16 copyable coding patterns |
| `references/tool_pitfalls.md` | openpyxl, OneDrive, expression and Stata traps |
| `references/diff_rubric.md` | Authority order, do-not-flag list, diff types, severity |
| `references/what_the_project_can_change.md` | Which settings a project adjusts through the brief or flags, and which rules are fixed |
| `scripts/extract_for_diff.py` | docx/xlsx -> text dumps; `--compare` for version diffs |
| `scripts/build_diff_docx.py` | findings_NN.md files -> one color-coded Word report |
| `scripts/xlsform_io.py` | read-edit-write helper for XLSForms: rebuilds sheets from rows and carries the workbook's formatting (fills, widths, frozen panes, rules) through every edit |
| `scripts/validate_xlsform.py` | PASS/FAIL checks every build and patch must pass |
| `scripts/paper_from_xlsform.py` | XLSForm -> paper questionnaire docx, one per language (PAPER mode engine); `--pdf` also exports PDFs through Word |
| `scripts/pdf_to_docx.py` | fallback for a PDF-only paper: PDF -> reviewable docx + conversion report; the extractor calls it for `.pdf` inputs |
| `requirements.txt` | Python packages the scripts need (openpyxl, lxml, python-docx, PyMuPDF) |
| `example/` | fictional starter project: brief, Round-1 form, Round-2 papers, configs, expected outputs, `run_checks.py` |

## Install

Claude Code (user-level, available in every project):

```bash
git clone https://github.com/meanregressive/capi-form-lifecycle ~/.claude/skills/capi-form-lifecycle
```

Project-level instead: clone into `<project>/.claude/skills/capi-form-lifecycle`. Other Agent-Skills
hosts: place the folder in the host's skills directory; `SKILL.md` must sit at the folder root.

Python 3.10+ for the six scripts; install their dependencies with `pip install -r requirements.txt`.

## First use in a project

1. Copy `references/project_brief_TEMPLATE.md` to `<project>/00_reference/capi_project_brief.md` and
   fill it in. The skill refuses to guess anything the brief leaves blank.
2. Lay the project out as the brief's section 10 describes (numbered input, output, attachments and
   scripts folders), or record your own layout there.
3. Invoke a mode: `/capi-form-lifecycle build <instrument>`, `patch <instrument>`,
   `diff <instrument>`, `paper <instrument|all>`. `<instrument>` is a number, short name or full name
   from the brief's register.

## Try it on the starter example

`example/` is a complete, fictional two-round school-survey project laid out the way the brief
expects: a filled-in brief, a Round-1 XLSForm, Round-2 paper questionnaires in two languages (with a
PI comment thread and a tracked change), preload CSVs and PAPER-mode configs. Six differences between
form and paper are planted and documented, with the expected DIFF findings in `example/expected/`.
Open your agent with `example/` as the working folder and run `/capi-form-lifecycle diff head`.
`python example/run_checks.py` regenerates the example and tests every script against it. See
`example/README.md`.

## Data handling

The skill is designed to work without reading personal data: preload files are checked by header row
only, and the brief names the folders the agent may never open. Nothing in this repository contains
project data. Review your organization's data-governance and AI-use policies before pointing an agent
at survey material.

## How this differs from other SurveyCTO skills

Two other agent skills cover SurveyCTO forms:

- [SurveyCTO's official skill](https://github.com/surveycto/surveycto-agent-skill) is a domain-knowledge
  skill: fifteen reference primers (XLSForm columns and field types, the expression language and its
  divergences from ODK, dataset XML, Data Explorer workbooks, field plug-ins, translation, conversion from
  Kobo, ODK, CommCare and Qualtrics), an XLSForm template, and workflows for creating, translating and
  converting a single form. Its optional MCP server adds live documentation search and row-level form
  editing. It makes an agent fluent in SurveyCTO.
- [dmbwebb/surveycto-questionnaire](https://github.com/dmbwebb/surveycto-questionnaire) is a Claude Code
  skill plus CLI toolkit: a 20-check XLSForm validator, a text dump for review and diffing, and a script
  that uploads a form definition to a SurveyCTO server.

This skill is a process skill. It makes an agent fluent in *your survey round* rather than in the
platform, and covers ground the other two do not:

- **Multi-round continuity.** BUILD takes the prior round's form and the new paper as joint inputs; PAPER
  at the end of a round produces the next round's reference form. The other skills work on one form
  instance with no concept of a round.
- **Form-to-paper regeneration** as a training instrument (landscape Word document per language, grids
  for repeats, skip and constraint logic shown as markers), not just a text dump.
- **Paper-versus-form audit** as a first-class mode, with an authority order, a do-not-flag list and
  severity grading that reflect how PI revisions, translations and console edits actually collide.
- **The brief as configuration.** Method and project are separated, with "the brief wins" as a hard
  rule, so the skill travels between projects instead of recording one survey.
- **Data-governance rules** written for research organizations: de-identified inputs only, header rows
  rather than data rows, folders the agent may never open.

What the other skills cover that this doesn't (yet): reference depth on the expression language (we keep copyable patterns, not
a full operator reference); scripted upload to the server (our deployment step is a checklist a person
follows because we want to keep the server control in the user's hands); datasets, Data Explorer workbooks and field plug-ins, which are outside this skill's scope. 

Our validator overlaps these skills to some extent on syntax checks. This skill adds bilingual coverage, question-code conventions, and the
intended-diff check that keeps a patch to its approved plan.

Recommended composition: keep this skill as the orchestrator and let the agent defer to the official surveycto
skill or its MCP server for platform lookups. Form edits should stay on this skill's scripted, logged
path, since editing through a hosted endpoint would bypass the per-change script and changelog that the
audit trail depends on.

## Resources

- [DIME Wiki: Questionnaire Programming](https://dimewiki.worldbank.org/Questionnaire_Programming),
  the World Bank DIME Analytics guide to CAPI programming, with linked pages on SurveyCTO coding
  practices, form testing and the paper-to-electronic workflow this skill automates.

- [SurveyCTO Agent Skill](https://github.com/surveycto/surveycto-agent-skill) (Dobility, Apache-2.0):
  SurveyCTO's own skill for designing, editing, debugging and converting XLSForms, server datasets,
  Data Explorer workbooks and field plug-ins, with an optional MCP server for live documentation
  search. The two skills are complementary: this one supplies the round workflow (brief, sign-off
  tables, versioning, validation, paper regeneration); the SurveyCTO skill supplies platform
  expertise. Install both and, when a BUILD or PATCH step needs a SurveyCTO-specific answer (an
  expression, an appearance, a plug-in bundle, a dataset XML), the agent can invoke the SurveyCTO
  skill for that step and return to this skill's plan table. Nothing in this repository depends on it.

## Status

Working version, used in production on one multi-round survey. Expect the mode files to tighten with
each new project; contributions and issue reports are welcome.

## License

MIT (see `LICENSE`). 