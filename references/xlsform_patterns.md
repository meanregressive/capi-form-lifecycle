# XLSForm coding patterns

Copyable SurveyCTO/XLSForm patterns the modes draw on. Values shown (`-77`, `-99`, `yesno`, `ENG`,
`POR`) are one project's conventions, used as examples; **take the actual values from the project
brief §6–§7**. Column names follow the bilingual layout `label:<TAG>` (design language) + `label`
(field language, set as `default_language`).

Rows are written as `column: value` lists; each block is one or more survey rows.

---

## 1. Yes/No with a paper skip arrow

Paper: `B05a  1 – Yes >> consent   2 – No`. Skips become `relevance` on the **following** fields,
never a literal jump.

```yaml
type: select_one yesno
name: B05a
label:ENG: B05a. Is ${deputy_name_baseline} still the deputy school director?
label:     B05a. ${deputy_name_baseline} continua a ser o(a) Director(a)-Adjunto(a) da Escola?
required: yes
# fields between B05a and consent:
relevance: ${B05a}=0
```

Paper "1 = Yes / 2 = No" is display formatting; the list keeps the project's values (1/0).

## 2. Select with "Other, specify" companion

```yaml
type: select_one school_personnel
name: B02
label:ENG: B02. With whom did you meet?
label:     B02. Com quem se reuniu?

type: text
name: B02_o
label:ENG: B02.1. Other person, specify:
label:     B02.1. Outra pessoa, especifique:
relevance: ${B02}=-77                 # select_one
# relevance: selected(${B02}, '-77')  # select_multiple: quote the code
```

Choices: append new values at the end; never renumber or reuse a value; proper nouns identical in
both label columns.

## 3. Exclusive option in a select_multiple ("none", "don't know", "I don't like them")

```yaml
type: select_multiple food_groups
name: D13
constraint: not(selected(., '11')) or count-selected(.) = 1
constraint message:ENG: This option cannot be combined with others.
constraint message:     Esta opção não pode ser combinada com outras.
```

Add an enumerator note before a multi-select the paper used to show as single-select ("more than
one option can be selected" / "pode seleccionar mais do que uma opção").

## 4. Preload from a CSV attachment

```yaml
type: calculate
name: village
calculation: pulldata('upload_school', 'village', 'school_id_key', ${A01})
```

- Key column name must match the CSV exactly (`school_id_key` vs `school_id` has bitten twice).
- Column and file are validated by **header only** (governance rule).
- Every new pulldata column goes into the preload spec file.
- Pulled **codes** shown in a note need an explicit map (pattern 12); pulled text/numbers display as is.

## 5. Confirm-preload with empty-preload fallback

When a question substitutes a preloaded value that may be blank for some records:

```yaml
type: calculate
name: deputy_name_baseline
calculation: pulldata('upload_council', 'adj_name', 'school_id_key', ${A01})

type: select_one yesno
name: B02c
label:ENG: B02c. Is ${deputy_name_baseline} the deputy school director you are interviewing today?
relevance: ${deputy_name_baseline}!=''
required: yes

type: select_one yesno
name: B02c_nopre
label:ENG: B02c. We have no record of the deputy school director for this school. Is there currently a deputy?
relevance: ${deputy_name_baseline}=''
required: yes

# downstream capture of the current incumbent:
relevance: ${B02c}=0 or ${B02c_nopre}=1
```

Wording rule learned the hard way: ask "is X the person you are interviewing today?", not "is X still
the person from <prior round>?", and **do not gate demographics on the confirm** — a confirmed name
can still be a new person whose details were never collected.

## 6. Role-based and cascaded IDs

```yaml
type: calculate
name: deputy_id
calculation: concat(${A01}, '21')          # suffix from the ID registry

type: calculate
name: meals_lead_id
calculation: if(${C12b}=1, ${C07b}, if(${C12c}=1, concat(${A01},'21'), concat(${A01},'40')))
```

IDs are role-based (stable when the incumbent changes). A new role = a new registry row first.

## 7. Group wrapper and section timing

```yaml
type: begin group
name: section_c
label:ENG: SECTION C. ...
label:     SECÇÃO C. ...
  type: calculate_here   name: start_c     calculation: once(duration())
  # ... content ...
  type: calculate_here   name: end_c       calculation: once(duration())
  type: calculate        name: duration_c  calculation: ${end_c} - ${start_c}
type: end group
name: section_c
```

Relevance on the `begin group` row applies to all children. **Avoid `field-list`** for a group that
contains a required field whose relevance depends on another field on the same screen: the
required flag is not reliably enforced. Put such follow-ups on their own screen.

## 8. Variable-length list (repeat)

```yaml
type: integer
name: C03
label:ENG: C03. How many teachers teach grade-3 reading?

type: begin repeat
name: C04_list
label:ENG: List of teachers          # repeats need a label or the console warns
label:     Lista de professores
repeat_count: ${C03}
  type: calculate_here  name: C04_index  calculation: index()
  type: text            name: C04        label:ENG: C04.1. Name of teacher ${C04_index}
type: end repeat
name: C04_list
```

The paper may unroll this as chained C04.1 / C04.2 / C04.2.1 questions; keep the repeat in the form.

## 9. Per-item repeat ("for each option selected in X, ask Y")

```yaml
type: begin repeat
name: d02_items
repeat_count: count-selected(${D02})
  type: calculate  name: D02_idx  calculation: position(..)
  type: calculate  name: D02_val  calculation: selected-at(${D02}, ${D02_idx}-1)
  type: calculate  name: D02_lbl  calculation: jr:choice-name(${D02_val}, '${D02}')
  type: integer    name: D03      label:ENG: D03. How much ${D02_lbl} was delivered?
type: end repeat
name: d02_items
```

`position(..)` not `index()` inside the repeat; `selected-at` is zero-based. `jr:choice-name` on a
**field answered in this form** works and is language-aware; on a **pulled code** it does not (see 12).

## 10. Two-phase roster: review prior-round members, then add new ones

```yaml
type: begin repeat   name: hh_list      repeat_count: ${n_base}      # baseline members
  type: calculate    name: mkey         calculation: concat(${A0}, '_', position(..))
  type: calculate    name: name_pre     calculation: pulldata('baseline_roster','name','member_key',${mkey})
  type: select_one yesno  name: hh00    label:ENG: Is ${name_pre} still a member of this household?
  type: select_one missing_reason name: reason_left   relevance: ${hh00}=0   required: yes
type: end repeat     name: hh_list

type: select_one yesno  name: hh00a    label:ENG: Are there any other members not listed?
type: integer           name: hh00a_n  relevance: ${hh00a}=1

type: begin repeat   name: hh_list_new  repeat_count: ${n_new}       # new members
  # name / sex / age / relationship
type: end repeat     name: hh_list_new

# consumers over the combined roster (slot i):
calculation: if(i <= ${n_base}, indexed-repeat(${x}, ${hh_list}, i), indexed-repeat(${x_new}, ${hh_list_new}, i - ${n_base}))
```

pulldata inside a repeat works with a **literal column name and a dynamic key**. A choice_filter that
reads the baseline repeat must cope with `n_base = 0` (new records) or every option disappears.

## 11. Constraint conventions (copy from the brief, not from memory)

```yaml
# phone: national mobile range plus the unknown sentinel (values from the brief)
constraint: (.>=<min> and .<=<max>) or .=-99
hint:ENG: If the person has no phone, write -99.

# person name: letters, accents, space, apostrophe, hyphen, period; no digits
constraint: regex(., "[A-Za-zÀ-ÿ '’.-]+")

# digits only
constraint: regex(., '^[0-9]+$')

# age with unknown sentinel
constraint: .=-99 or (.>=18 and .<=99)

# date of birth shown only when age unknown, must agree with age
relevance: ${age}=-99
constraint: .=today() or (int((today()-.) div 365.25) >= ${age}-1 and int((today()-.) div 365.25) <= ${age}+1)

# years in a job cannot exceed age; attended days cannot exceed total days (constraint on the LATER field)
constraint: .>=0 and .<=${age}
constraint: .>=0 and .>=${days_attended}

# counts that may legitimately be 0
constraint: .>=0        # not .>0
```

`regex()` is whole-string: "contains a letter" must be written `.*[A-Za-zÀ-ÿ].*`. Decimal examples
in hints use a period.

## 12. Displaying a pulled code as a label

`jr:choice-name(${code}, 'list')` renders blank for pulldata output. Use an explicit map:

```yaml
calculation: if(${sex_pre}='1','Masculino', if(${sex_pre}='2','Feminino',''))
```

Not language-aware: hardcode the field language, or build one calc per language.

## 13. Dynamic media from a preload

```yaml
type: calculate  name: t1_card_img  calculation: pulldata('school_choice_tasks','task1_card','version_id',${version_id})
type: note       name: x_t1_table   media:image: ${t1_card_img}
```

`translate()` and string functions in the filename calc break the parser: precompute the filename
as a CSV column. Many media files → one flat ZIP attachment (see deployment checklist). `<img>` in
HTML labels does not render offline; `media:image` does.

## 14. Confirmation note for preloaded identification

```yaml
type: select_one yesno
name: A07
label:ENG: |
  Enumerator, confirm the following before continuing:
  School ID: ${A01}  ·  School: <b>${A02}</b>  ·  District: ${A03a}  ·  Village: ${A05}
  <b>Is the information above correct?</b>
required: yes
```

When a preload field is added, update this note in every language. Give any "confirm and update"
note the same relevance as the fields it introduces, or it shows with nothing to confirm.

## 15. Availability / existence gate at the top of an interview

```yaml
type: select_one yesno   name: A07a  label:ENG: A07a. Does this school have a farmer group?
type: select_one fa_absent_reasons  name: A07a_reason  relevance: ${A07a}=0
type: note               name: A07a_end  relevance: ${A07a}=0   label:ENG: ENUMERATOR: the survey ends here...
type: select_one yesno   name: A07b  relevance: ${A07a}=1  label:ENG: A07b. Is a representative available today?
# every downstream section: relevance ... and ${A07b}=1
```

Prefer this over letting the enumerator swipe to the end; it records *why* no interview happened.

## 16. Settings block

Start a brand-new form from SurveyCTO's basic template (`empty_instrument.xlsx`, linked from the DIME Wiki's
Questionnaire Programming page; a copy ships in `example/_template/`): it carries the full column set, the
metadata rows and the conditional formatting that colors rows by field type. Replace its `version` formula
with a literal stamp.

```
form_title:        #<N> - <field-language short title> — <Round>
form_id:           <prior_form_id>_<round>_v1        # at round start; first round / new instrument: <short_name>_<round>_v1
                                                      # then in-place version bumps (brief §8)
version:           2609180001                         # literal YYMMDDNNNN, never a formula
default_language:  POR                                # tag of the un-tagged `label` column
```
