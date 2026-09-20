# What a project can change, and what is fixed

The skill has one design principle: **the brief wins**. Anything the project brief has a slot for is
the project's decision and is read at run time. Anything without a slot is the skill's *method*, and a
project changes it only by editing the skill's own files (in practice, by forking the repository).
This page lists which is which, so a team knows before its first run where its choices go.

## 1. Adjustable through the brief (no skill edit needed)

| What | Brief section | Where it takes effect |
|---|---|---|
| Instruments: numbers, short names, full names, file stems, which have a prior-round form | §2 | every mode resolves `<instrument>` against this register |
| Folder layout and file naming; the numbered folders are only defaults | §10 (and §3) | every mode locates inputs and writes outputs here |
| Languages: design language, field language, which label column is which | §1 | BUILD, PATCH, PAPER (`--langs`), validator (`--lang-tag`) |
| Authority order: who wins on content, codes, each language's text, console edits, team feedback | §4 | DIFF recommendation column; BUILD Phase B; PATCH plan actions |
| Translation policy: halt / draft-and-flag / reuse-only / reuse-first | §5 | BUILD Phase D, PATCH P3; the one setting that changes agent behavior materially |
| Coding conventions: yes/no values, other / don't-know / refused codes, sentinel, code prefix in labels, constraint patterns, required-by-default, repeats vs chained prompts | §6 | every new field the skill writes; validator flags `--dk`, `--other` |
| Identifier scheme: primary key, role suffixes, registry file | §7 | BUILD, PATCH |
| Versioning: stamp format, form_id pattern at round start, when form_id may change, mirror folder, test IDs, enumerator dataset name, media limits | §8 | PATCH P2, deployment checklist, validator `--server-version` |
| Authoring copy: server download or a local formatted master | §8 | which file PATCH starts from; formatting carried either way |
| Which logs exist and where | §9 | BUILD Phase E, PATCH P5, PAPER step 6 |
| Extra do-not-flag items for DIFF; extra known traps | §11 | DIFF D1 (extends `diff_rubric.md` §2); BUILD/PATCH trap checks |
| Data-governance boundary: folders never opened, header-rows-only rule | §3 | every mode (rule G4) |
| Who owns what: content, translation, preloads, deployment, cleaning | §1 | who the skill addresses in to-dos and hand-offs |

## 2. Adjustable per run (flags and configs)

| What | How |
|---|---|
| BUILD without writing anything | `--dry-run` (stops after Phase B) |
| BUILD with no prior-round form | `--from-scratch` (implied when the register says `scratch`) |
| Which fix-list sources PATCH takes in | named in the conversation; backlog items only when the user opens the batch |
| Severity scope handed from DIFF to PATCH (High only / High+Med / all) | decided per instrument at the start of the PATCH pass |
| PAPER layout per instrument: banner titles, roster-variable substitutions, matrix rendering of a repeat, note markers, UI strings for a new language | the instrument's JSON config (`paper_configs/<instrument>.json`) |
| PAPER output suffix, language map, PDF export | `--suffix`, `--langs`, `--pdf` (or the same keys in `batch.json`) |
| Validator inputs | `--base`, `--planned`, `--attachments`, `--server-version` |
| PDF conversion behavior | `--no-tables`, `--min-chars` on `pdf_to_docx.py` |

## 3. Fixed: the method (mode files, rule tables)

These do not have a brief slot. They are what makes the skill run "with the same discipline every time".

| Rule | Where |
|---|---|
| Sign-off before any edit: BUILD Phase B table, PATCH plan table; no exception for small edits | G5, B2, P3 |
| Copy first; never edit a prior-round form or a paper docx; paper errors go to the translator/PI list | G2, B1, P8 |
| Every change is a script in the project's scripts folder; CHANGELOG entry the same day; session file log | G3, G7, B3 |
| Validate before reporting; a FAIL is never reported as success | G6, B9, P5 |
| PATCH base is today's server download, never the previous script's output; back up before editing | P1, P2 |
| `version` always increases; `form_id` changes only by explicit decision | P4 (format adjustable in brief §8; the rule is not) |
| Choice values are appended, never renumbered or reused | P6 |
| "By design" and "deferred" outcomes are written down, not dropped | P7 |
| DIFF is read-only; compares dumps, not live files; findings files are append-only history | D1, D2, D6 |
| A missing side is stated, never inferred | D5 |
| PAPER never edits the form; the engine is never patched for one instrument | R1, R2 |
| Clarifying questions capped at two rounds of four; beyond that the instrument is split into passes | BUILD Phase C |
| Fixed vocabularies: Phase B statuses; plan-table change types and actions; the nine diff types; three severities; four recommendation verbs; the findings-file structure | build.md, patch.md, diff.md, diff_rubric.md |
| Prior-round quirks are flagged, not silently fixed; skip arrows become `relevance`, never literal jumps | B7, B8 |

## 4. Fixed in the references, but stated as examples

- `xlsform_patterns.md` and its settings block use one project's values and say so; the brief's §6
  values replace them when the agent writes a form.
- `diff_rubric.md`: the default authority order (§1) and do-not-flag list (§2) are overridden or
  extended by the brief's §4 and §11; the diff types (§3), alignment method (§4) and severity
  definitions (§5) are fixed.
- `deployment_checklist.md` is SurveyCTO server mechanics, not a project choice; names come from §8.
- `tool_pitfalls.md` is a list of known traps; a project adds its own in brief §11 rather than editing it.

## 5. Fixed in the scripts

- `validate_xlsform.py`: the set of checks is fixed. Flags supply inputs (language tag, codes,
  server version, planned fields, attachments folder); no flag turns a check off.
- `paper_from_xlsform.py`: the rendering rules (three-column table, grid or matrix for repeats,
  `[SHOW IF]`/`[CHECK]` markers, plumbing dropped) are fixed; the per-instrument JSON covers layout
  knobs and UI strings, not rules.
- `extract_for_diff.py`, `build_diff_docx.py`, `pdf_to_docx.py`: dump format, report layout and
  conversion behavior are fixed beyond the flags listed above.
- `xlsform_io.py`: what is carried (styles, widths, hidden columns, freeze panes, rules, validations,
  tab color) is fixed; merged cells are reported and dropped.

## 6. What would need a fork

Dropping a sign-off gate; editing prior-round forms or paper in place; writing non-source-language
text without recording it, whatever the policy; changing the diff vocabulary or severity definitions;
turning validator failures into warnings. If a project needs one of these, the recommended route is
not a silent local edit but a dated "Rules this project relaxes" note in its brief, so the exception is
visible and attributable, plus an issue on the repository if the rule seems wrong in general.
