# capi-form-lifecycle

An [Agent Skill](https://agentskills.io/specification) for building, editing, and finalizing Computer-Assisted Personal Interviewing (CAPI) survey instruments for data collection on SurveyCTO (or similar platform). The skill runs **build** to generate the XLSForm from the paper questionnaire and any prior round's form; **patch** while it is deployed; **diff** against the paper questionnaire, and **paper** to regenerate the paper questionnaire from the final form.

Developed by Deboleena Rakshit from July to September, 2026, at the International Food Policy Research Institute (IFPRI), assisted by Claude Code (Fable 5.1).

## Why

A CAPI survey round has to maintain version control and consistency across three streams of work: the paper questionnaire that investigators, research, or program teams sign off on; the electronic form(s) the tablets run; and, in the case of multiple rounds of data collection, any prior round's forms that data structures must stay comparable with. This can also often across multiple languages for most international development projects. 

A first (or "baseline") round has no prior form; the **build** mode then works from the paper questionnaire alone, and the CAPI Excel form it produces becomes the next round's (e.g., "midline" or "endline") input for **build**. 

In practice, most of the work happens during the **patch** build of deployed forms. This refers to the process of implementing survey-level or item-level changes that can be *minor* (like tweaks to question wording, skip logic, response codes and labels) or *structural* (like adding survey modules or plugins), based on testing and piloting performed by the research or program teams. 

This **patch** mode is complemented by a **diff** mode where any deviations from the paper questionnaire are checked by the AI agent and verified by the user as intentional before redeploying the updated form. The **diff** mode is also invoked when checking differences between iteratively updated versions of the CAPI forms.

The final **paper** mode is an optional step where the AI agent uses this skill to generate a reader-friendly .pdf or .docx version of the CAPI Excel form to use during training for survey firms, or for circulating among research and program teams, doing away with the need for manually updating the paper version of the questionnaire after each iterative change to the form.

This skill encodes this multi-step workflow so an AI agent can run it with the same discipline each time: a plan table signed off by the user (CAPI programmer) before any edits, a reproducible script for every change, validation by the user after the changes, and a changelog entry updated after each session.

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
| `references/xlsform_patterns.md` | 16 copyable SurveyCTO coding patterns |
| `references/tool_pitfalls.md` | openpyxl, file syncing, expression and Stata traps |
| `references/diff_rubric.md` | Authority order, do-not-flag list, diff types, severity of issues |
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

Prerequisites: Git (for the clone) and Python 3.10+ for the six scripts; install their dependencies
with `pip install -r requirements.txt`.

## First use in a project

Run the starter example first (next section) to confirm the install; restart Claude Code once after
cloning so the skill appears under `/`.

1. Make a project folder and put the paper questionnaires in it, one subfolder per language, plus the
   prior round's XLSForm if there is one. No raw survey data.
2. Open the folder in your agent and say "set up the CAPI project brief". It infers what it can from
   your files, asks the rest a few questions at a time, and writes `00_reference/capi_project_brief.md`
   with dated decisions; anything you cannot answer yet is marked `[CONFIRM]` so work can continue.
   You can also fill `references/project_brief_TEMPLATE.md` manually before loading up your Claude Code session.
3. Invoke a mode: `/capi-form-lifecycle diff <instrument>` to see where paper and form disagree, then
   `build`, `patch` or `paper`. `<instrument>` is a number, short name or full name from the brief's
   register. Every mode shows a plan table and waits for your approval before editing anything, and
   none of them uploads to SurveyCTO; you do that from the console with the checklist it prints.
4. For testing and piloting rounds, give testers `references/testing_feedback_TEMPLATE.docx` (or paste
   it into a shared Google Doc); `patch` reads its `Open` rows as the fix list.

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

What the other skills cover that this doesn't (yet): unlike the official surveycto skill our skill does not store
a full reference of SurveyCTO operators); unlike dmbwebb's skill we do not have a scripted upload to the server (our deployment step is a checklist a person follows because we want to keep the server control in the user's hands); finally SurveyCTO datasets, Data Explorer workbooks and field plug-ins are currently outside this skill's scope. 

Our validator overlaps these skills to some extent on syntax checks. This skill adds bilingual coverage, question-code conventions, and the
intended-diff check that keeps a patch to its approved plan.

Recommended composition: keep this skill as the orchestrator and let the agent defer to the official surveycto
skill or its MCP server for platform lookups. Form edits should stay on this skill's scripted, logged
path, since that ensures there is a changelog and audit trail that is maintained on your system to track changes over versions and time.

## Optional suggested resources

- [DIME Wiki: Questionnaire Programming](https://dimewiki.worldbank.org/Questionnaire_Programming),
  the World Bank DIME Analytics guide to CAPI programming, with linked pages on SurveyCTO coding
  practices, form testing and the paper-to-electronic workflow this skill automates.

- [SurveyCTO Agent Skill](https://github.com/surveycto/surveycto-agent-skill) (Dobility, Apache-2.0):
  SurveyCTO's own skill for designing, editing, debugging and converting XLSForms, server datasets,
  Data Explorer workbooks and field plug-ins, with an optional MCP server for live documentation
  search. 

## Status

Working version, used in production on one multi-round survey. Expect the mode files to tighten with
each new project; contributions and issue reports are welcome.

## License

MIT (see `LICENSE`). 