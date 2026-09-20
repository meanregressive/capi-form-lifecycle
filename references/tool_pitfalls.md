# Tool pitfalls

Things that silently produce wrong output. Each one was hit at least once in a real field round.

## openpyxl

- **`ws.cell(row, col, value=None)` is a no-op.** It does not clear the cell; stale values survive
  and end up in the deployed form. Rebuild a sheet from row lists: `del wb[name]`,
  `create_sheet(name, index)`, `append(row)` for every row. Or assign `cell.value = None`.
- **`insert_rows` / `delete_rows` do not move what is anchored to rows.** Conditional-formatting
  ranges, data validations and merged cells keep their old row numbers, so after a structural edit
  they point at the wrong fields. `xlsform_io.py` avoids this by rebuilding and re-anchoring by field
  name; do not edit an XLSForm in place with insert/delete rows.
- **A plain rebuild drops every fill, font, width and rule.** Teams mark translation status with
  color; use `xlsform_io.XLSFormBook` (or `xlsform_io.py carry base rebuilt --out`) rather than a
  bare remove/create/append.
- **Trailing `None` columns.** Rows read with `iter_rows(values_only=True)` have ragged lengths;
  pad to the header length before indexing by column.
- **Numbers vs strings.** SurveyCTO downloads turn choice `value` `1` into `'1'`; compare with a
  normalizer (`'1' == 1`) unless you want the artifact reported.
- **`read_only=True`** is fine for reading; never for a workbook you will save.
- **Settings `version` as a formula** (baseline forms had `=TEXT(NOW()...)`): openpyxl returns the
  formula string, not a number. Always write a literal stamp.

## OneDrive / SharePoint folders

- Files open in Word/Excel or mid-sync raise `PermissionError` / "resource busy". Read via a temp
  copy with a short retry loop; for writing, ask the user to close the file. Never delete or force.
- Sync conflicts appear as `name-DESKTOP-XXXX.xlsx`; stop and let the user resolve.
- Two files for one instrument in a version folder (a stale plain-named copy beside the deployed
  URL-encoded `#9+-+…` one) → globs pick the wrong base. Confirm which is of record.
- OneDrive "files on demand" placeholders read as empty; make sure the file is downloaded.

## SurveyCTO / ODK expression traps

- `regex()` is **whole-string**. "Contains a letter" = `.*[A-Za-zÀ-ÿ].*`; a bare class rejects every
  multi-character input.
- `translate()` and other string functions inside a `media:image` filename calc break the form
  parser; precompute filenames into a CSV column.
- `jr:choice-name(value, 'list')` on a **pulldata** result renders blank; use an `if()` map.
  Inside a repeat over a field answered in the form it works.
- `${calc}` substitutions in labels are HTML-escaped: `<b>` inside a calculated string shows
  literally. Put formatting in the static label text, the value in the calc.
- `<img src>` in HTML labels does not render offline in Collect; use `media:image`.
- `.>0` rejects a legitimate 0; `.>=0` unless the paper requires at least 1.
- Decimal hints with a comma teach an input that fails; use a period.
- A `choice_filter` that reads a repeat (`indexed-repeat(..., ${hh_list}, filter)`) returns nothing
  when the repeat has zero instances → required field with no options.
- `required` on a field revealed dynamically inside a `field-list` group is not reliably enforced.
- `select_multiple` relevance must quote the code: `selected(${x}, '-77')`.
- `pulldata()` inside a repeat: literal column name + dynamic key works; dynamic column name does not.
- `today()` and duration calcs depend on the device clock.

## SurveyCTO console

- Enumerator field needs an ENUMERATORS-type dataset (see deployment checklist).
- Attachment swap without a `version` bump does nothing on devices.
- Thousands of attachments crash the attachments view; use one flat ZIP for media.
- Downloaded XLSForms carry benign artifacts: `publishable` populated, numeric→text in choices,
  whitespace trimmed, CRLF→LF. Fold them away before calling it drift.

## Word documents (python / lxml)

- Tracked changes: text lives in `w:t` (insertions included), deletions in `w:delText`. Reading
  `w:t` only = "accept all". Say so in any dump.
- Comment threads: `word/commentsExtended.xml` links replies to parents by `w15:paraIdParent`; the
  `w15:done` flag is unreliable for resolution. Judge resolution by whether the proposal is in the
  paper or form.
- Cells with `|` break pipe tables in dumps; replace with `/`.
- `python-docx` cannot read comments; use lxml on the zip members directly.

## Stata (when preload builders or checks run in Stata)

- `.do` files must be **CRLF**. Git-bash `sed -i` strips the CRs silently and Stata then executes
  nothing (every line shows as a `>` continuation). Edit with python in bytes; verify
  `b.count(b'\r\n') == line count`.
- Block comments **nest**: `/*` inside a path like `folder/*.do` reopens a comment; a stray `*/`
  closes one early.
- Launching `StataMP-64.exe /e do file.do` from **Git Bash** mangles `/e` into a drive letter and
  opens an interactive window with no log. Launch from **PowerShell** with `Start-Process … -Wait`.
  Check the log's mtime before believing a run happened.

## Encodings

- Windows consoles default to cp1252; wrap `sys.stdout` in a UTF-8 writer or set
  `PYTHONIOENCODING=utf-8` when printing Portuguese.
- Double-encoded names (`LÃºcia`) come from reading UTF-8 as latin-1 once; repair with
  `s.encode('latin-1').decode('utf-8')` inside a loop that stops on failure and never introduces
  U+FFFD.
- Write CSV attachments UTF-8 **without** BOM.
- Stata exports with value labels instead of codes ("label exports") recur; normalize before HFC or
  merges.

## Git Bash on Windows

- `/e`, `/c`, `/d` at the start of an argument are rewritten as drive paths (MSYS path conversion).
- Heredocs with long Markdown bodies can fail on quote balancing; write the body with a file tool,
  then process it with a short script.
