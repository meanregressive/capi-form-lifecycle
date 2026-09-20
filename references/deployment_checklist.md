# SurveyCTO deployment checklist

Server mechanics that are outside the XLSForm but decide whether a form works on a tablet. Every
item here has cost a real field team at least a day. Project-specific names (server, dataset names,
test IDs, mirror folder) come from the brief §8.

## Three layers, three lifecycles

| Layer | What it is | Changes when | Identified by |
|---|---|---|---|
| Form definition | the xlsx (`survey`, `choices`, `settings`) | content edits | `form_id` + `version` |
| Form attachments | CSVs, images, ZIPs, field plugins uploaded **to one form** | data refresh, media swap | filename, per form_id |
| Server datasets | tables living on the server, bound to many forms (enumerator dataset, published-to datasets) | rows edited on the console or via publishing | dataset id |

Rules that follow:
- **A content change needs a `version` bump** (strictly greater than the server's value). Same
  `form_id` = in-place update, bindings kept.
- **An attachment swap needs a `version` bump too**, or devices keep the old file. `form_id` stays.
- **Attaching a dataset, swapping a CSV or re-publishing the same xlsx never bumps anything by
  itself**; you bump `version` in the xlsx.
- **A `form_id` change creates a new form**: every attachment and dataset must be re-attached,
  publishing re-created, and submissions land in a new dataset. Do it only by decision (brief §8).

## First upload of a form (BUILD)

- [ ] Upload the xlsx; note the server's accepted `version`.
- [ ] Attach the **enumerator dataset** (see below) if the form has a `type: enumerator` field.
- [ ] Upload every file in the form's attachments subfolder (`07_capi_attachments/<n>_<name>/`).
- [ ] Deploy from draft to live; open the web-collect test link; pull the form in Collect.
- [ ] Test with the project's test IDs; discard test data or filter it by test ID later.
- [ ] Download the deployed xlsx into the mirror folder; update its README version table.

## Re-publish of a form (PATCH)

- [ ] Same `form_id` unless the plan says otherwise; `version` above the server's value.
- [ ] Changed attachments re-uploaded; unchanged ones left alone.
- [ ] If `form_id` did change: re-attach enumerator dataset + every attachment, re-create publishing.
- [ ] Devices: delete form + *Get Blank Form*; confirm large ZIPs finish downloading.
- [ ] Mirror folder updated the same day; CHANGELOG entry written.

## Enumerator dataset

- Must be a true **enumerator-type dataset**: created via *Design → Add → Dataset* with the
  "this dataset manages enumerators" box ticked; its definition XML contains
  `<discriminator>ENUMERATORS</discriminator>`. A plain server dataset with the right columns
  fails at runtime with "No enumerator dataset is attached to this form".
- Columns `id, name, users` (`users` = SurveyCTO usernames, comma-separated; the dropdown filters
  to the logged-in user). Keep a test dataset (test rows only) and a production one.
- Re-attach on every `form_id` change. Attaching never bumps the form.

## Attachments

- One subfolder per form; **shared CSVs byte-identical** across every subfolder that uses them
  (edit one canonical copy, copy verbatim, check md5). Different callers referencing the same file
  name with different columns is how key-column bugs enter.
- Key column names: match the form's `pulldata()` exactly (`school_id_key` vs `school_id`).
- Encoding: UTF-8 without BOM; check accented names survive (double-encoding shows as `LÃºcia`).
- Limits: < 100 MB per file, 300 MB per form. Thousands of loose media files crash the console's
  attachments view (jQuery call-stack error) → bundle media into **one flat ZIP** (files at zip
  root); `media:image=<name>.png` still resolves. CSV datasets stay as separate attachments.
- Field plugins (`*.fieldplugin.zip`) are attachments too; attach at first upload.
- Test rows (9999-series) are removed from production attachments before fieldwork.

## Console edits by the team

- Anyone with console rights can edit a form or bump a version. Rule for the team: **every console
  edit or upload → download the xlsx into the mirror folder the same day** and note the version.
- Before any PATCH, compare the server download with the mirror (DIFF mode, version-vs-version).
- Server-side version bumps you did not make are normal (attachment refreshes); record them.

## Versions in the wild

- Submissions carry `formdef_version`; use it to confirm a patch is live (expect a lag of a day or
  two for device sync) and to filter data collected on a superseded version.
- Two forms with different `form_id`s for the same instrument = two data streams to merge
  (it happens when a pre-visit team keeps an older form live after the main forms move on). Avoid unless intended.

## Decimals, dates, clocks

- SurveyCTO decimals use a period; a hint showing `0,5` teaches enumerators an input that fails.
- `today()`-based age checks depend on the tablet clock; a wrong device date produces false
  constraint failures.
