#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
_generate_example.py — (re)build the fictional starter project under example/.

Everything in example/ except this script, README.md and expected/findings_01.md is written by
this file, so the content is provably synthetic: no real school, person, phone number or project
appears anywhere. Re-run after editing the definitions below.

    python example/_generate_example.py

Writes:
  00_reference/capi_project_brief.md, demo_id_suffix_registry.md, demo_glossary_EN_PT.md
  01_input_paper_EN/20260105/1_school_head_EN.docx      (Round-2 paper, with a PI comment thread
                                                          and a tracked change)
  02_input_paper_PT/20260105/1_school_head_PT.docx      (Round-2 paper, PT; one planted
                                                          content misalignment)
  03_input_prior_capi/1_school_head_r1_v1.xlsx          (the Round-1 XLSForm, built on SurveyCTO's
                                                          basic template in _template/, see below)
  07_capi_attachments/1_school_head/upload_school.csv   (3 test schools, header + rows)
  08_SurveyCTO_datasets/enumerators_demo.csv
  05_python_scripts/paper_configs/1_school_head.json, batch.json
  04_output_r2_capi/README.md, CHANGELOG.md, SESSION_FILE_CHANGES.md (empty logs)

NOT written here: 01_input_paper_EN/20260105/1_school_head_EN.pdf is the EN docx printed to PDF once
with Word (tracked changes accepted for print) and committed as a fixture for the PDF fallback;
00_reference/r2_testing_feedback.md is a hand-written filled copy of references/testing_feedback_TEMPLATE.md.

Planted Round-1 form vs Round-2 paper differences (what BUILD / DIFF should find):
  NEW            A04 village (prefilled; needs a `village` column upload_school.csv does not have)
  NEW            B06 phone has internet access
  LABEL-CHANGED  C01 "enrolled" -> "enrolled this school year" (as a tracked change in the EN docx)
  RESPONSE-OPT   meal_provider gains "4 Parents' association"
  DELETED        E02 council meetings last term
  LANG-MISALIGN  PT paper C04 says "por mês" (per month) where EN says "per week"
  do-not-flag    paper shows Yes=1/No=2 and Other=99; form uses 1/0 and -77; skip arrows vs relevance
"""
import csv, json, os, zipfile, shutil, tempfile
import openpyxl
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
def P(*parts):
    p = os.path.join(HERE, *parts); os.makedirs(os.path.dirname(p), exist_ok=True); return p

FROZEN_STAMP = "2026-01-05T00:00:00Z"          # fixed save time so regenerated fixtures are byte-identical
FROZEN_ZIP_DT = (2026, 1, 5, 0, 0, 0)

def freeze_timestamps(path):
    """Office files (docx/xlsx) are zips that embed a save time in docProps/core.xml and in every zip
    entry. Fix both, so re-running the generator does not make git see a change without content."""
    import re
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(created|modified)[^>]*>)[^<]*(</dcterms:\2>)", rb"\g<1>" + FROZEN_STAMP.encode() + rb"\g<3>", data)
            zi = zipfile.ZipInfo(item.filename, date_time=FROZEN_ZIP_DT); zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)
    os.replace(tmp, path); return path

PAPER_DATE = "20260105"
R1_VERSION = 2501150001

# ============================================================================ choice lists
CHOICES = {
    "yesno":         [(1, "Yes", "Sim"), (0, "No", "Não")],
    "sex":           [(1, "Male", "Masculino"), (2, "Female", "Feminino")],
    "education":     [(1, "Primary", "Primário"), (2, "Secondary", "Secundário"),
                      (3, "Teacher training college", "Instituto de formação de professores"),
                      (4, "University", "Universidade"), (-77, "Other, specify", "Outro, especifique")],
    "meal_provider": [(1, "Government program", "Programa do governo"), (2, "NGO", "ONG"),
                      (3, "Community", "Comunidade"), (-77, "Other, specify", "Outro, especifique")],
    "council_issues": [(1, "Pupil attendance", "Assiduidade dos alunos"), (2, "School meals", "Refeições escolares"),
                       (3, "Infrastructure", "Infra-estruturas"), (4, "Teacher performance", "Desempenho dos professores"),
                       (5, "None of these", "Nenhum destes")],
}
# Round-2 paper adds an option (RESPONSE-OPTIONS finding):
PAPER_EXTRA_OPTION = ("meal_provider", 4, "Parents' association", "Associação de pais")

# ============================================================================ questions
# Each row: dict with keys used by both the XLSForm writer and the paper writer.
#   r1=True  -> in the Round-1 form ; paper=True -> in the Round-2 paper
NAME_RE = "^[A-Za-zÀ-ÿ '’.-]+$"
Q = [
 # ---- header / plumbing (form only)
 dict(type="start", name="starttime", r1=True, paper=False),          # metadata rows as in the SurveyCTO template
 dict(type="end", name="endtime", r1=True, paper=False),
 dict(type="deviceid", name="deviceid", r1=True, paper=False),
 dict(type="subscriberid", name="subscriberid", r1=True, paper=False),
 dict(type="simserial", name="simid", r1=True, paper=False),
 dict(type="phonenumber", name="devicephonenum", r1=True, paper=False),
 # ---- Section A
 dict(type="begin group", name="section_a", en="SECTION A. Identification", pt="SECÇÃO A. Identificação", r1=True, paper=True),
 dict(type="integer", name="A01", en="A01. School ID", pt="A01. Código da escola", required="yes", r1=True, paper=True,
      hint_en="Enter the 4-digit school ID from the assignment sheet.", hint_pt="Introduza o código de 4 dígitos da escola.",
      constraint=".>=1000 and .<=9999", cmsg_en="School ID has 4 digits.", cmsg_pt="O código da escola tem 4 dígitos."),
 dict(type="calculate", name="A02", en="A02. School name", pt="A02. Nome da escola", r1=True, paper=True,
      calc="pulldata('upload_school', 'school_name', 'school_id_key', ${A01})"),
 dict(type="calculate", name="A03", en="A03. District", pt="A03. Distrito", r1=True, paper=True,
      calc="pulldata('upload_school', 'district', 'school_id_key', ${A01})"),
 # NEW in Round-2 paper: village (needs a column the CSV does not have)
 dict(type="calculate", name="A04", en="A04. Village", pt="A04. Aldeia", r1=False, paper=True,
      calc="pulldata('upload_school', 'village', 'school_id_key', ${A01})"),
 dict(type="note", name="A07", r1=True, paper=True,
      en="Enumerator, confirm before continuing: School ID ${A01} · School ${A02} · District ${A03}.",
      pt="Inquiridor(a), confirme antes de continuar: Código ${A01} · Escola ${A02} · Distrito ${A03}."),
 dict(type="enumerator", name="A05", en="A05. Enumerator", pt="A05. Inquiridor(a)", required="yes", r1=True, paper=True),
 dict(type="text", name="A06", en="A06. Name of the respondent (school head)", pt="A06. Nome do(a) respondente (director(a) da escola)",
      required="yes", r1=True, paper=True, constraint=f'regex(., "{NAME_RE}")',
      cmsg_en="Letters and spaces only.", cmsg_pt="Apenas letras e espaços."),
 dict(type="note", name="consent_note", r1=True, paper=True,
      en="My name is ${A05}. I work for a research team studying primary schools. The interview takes about 20 minutes. Your answers are confidential and you may stop at any time.",
      pt="O meu nome é ${A05}. Trabalho para uma equipa de pesquisa sobre escolas primárias. A entrevista dura cerca de 20 minutos. As suas respostas são confidenciais e pode parar a qualquer momento."),
 dict(type="select_one yesno", name="consent", en="Do you agree to participate?", pt="Concorda em participar?", required="yes", r1=True, paper=True),
 dict(type="end group", name="section_a", r1=True, paper=True),
 dict(type="begin group", name="survey", relevance="${consent}=1", r1=True, paper=False),   # transparent wrapper
 # ---- Section B
 dict(type="begin group", name="section_b", en="SECTION B. Respondent", pt="SECÇÃO B. Respondente", r1=True, paper=True),
 dict(type="select_one sex", name="B01", en="B01. Sex of the respondent", pt="B01. Sexo do(a) respondente", required="yes", r1=True, paper=True),
 dict(type="integer", name="B02", en="B02. Age (years)", pt="B02. Idade (anos)", required="yes", r1=True, paper=True,
      hint_en="If unknown, write -99.", hint_pt="Se desconhecida, escreva -99.",
      constraint=".=-99 or (.>=18 and .<=99)", cmsg_en="Age 18–99, or -99.", cmsg_pt="Idade 18–99, ou -99."),
 dict(type="integer", name="B03", en="B03. Years as head of this school", pt="B03. Anos como director(a) desta escola", required="yes", r1=True, paper=True,
      hint_en="0 if less than one year.", hint_pt="0 se menos de um ano.",
      constraint=".>=0 and (.<=${B02} or ${B02}=-99)", cmsg_en="Cannot exceed age.", cmsg_pt="Não pode exceder a idade."),
 dict(type="integer", name="B04", en="B04. Mobile phone number", pt="B04. Número de telemóvel", required="yes", r1=True, paper=True,
      hint_en="9 digits. If no phone, write -99.", hint_pt="9 dígitos. Se não tem telefone, escreva -99.",
      constraint="(.>=100000000 and .<=999999999) or .=-99", cmsg_en="9 digits or -99.", cmsg_pt="9 dígitos ou -99."),
 dict(type="select_one education", name="B05", en="B05. Highest level of education completed", pt="B05. Nível de escolaridade mais elevado concluído",
      required="yes", r1=True, paper=True),
 dict(type="text", name="B05_o", en="B05.1. Other, specify:", pt="B05.1. Outro, especifique:", relevance="${B05}=-77", required="yes", r1=True, paper=True),
 # NEW in Round-2 paper
 dict(type="select_one yesno", name="B06", en="B06. Does your phone have internet access?", pt="B06. O seu telefone tem acesso à internet?",
      relevance="${B04}!=-99", required="yes", r1=False, paper=True),
 dict(type="end group", name="section_b", r1=True, paper=True),
 # ---- Section C
 dict(type="begin group", name="section_c", en="SECTION C. School and school feeding", pt="SECÇÃO C. Escola e alimentação escolar", r1=True, paper=True),
 dict(type="integer", name="C01", en="C01. How many pupils are enrolled?", pt="C01. Quantos alunos estão matriculados?", required="yes", r1=True, paper=True,
      en_paper="C01. How many pupils are enrolled this school year?", pt_paper="C01. Quantos alunos estão matriculados neste ano lectivo?",
      constraint=".>=0 and .<=5000", cmsg_en="0–5000.", cmsg_pt="0–5000."),
 dict(type="select_one yesno", name="C02", en="C02. Does the school have a school feeding program?", pt="C02. A escola tem um programa de alimentação escolar?",
      required="yes", r1=True, paper=True, skip_en="If No >> Section D", skip_pt="Se Não >> Secção D"),
 dict(type="select_one meal_provider", name="C03", en="C03. Who provides the meals?", pt="C03. Quem fornece as refeições?",
      relevance="${C02}=1", required="yes", r1=True, paper=True),
 dict(type="text", name="C03_o", en="C03.1. Other, specify:", pt="C03.1. Outro, especifique:", relevance="${C02}=1 and ${C03}=-77", required="yes", r1=True, paper=True),
 dict(type="integer", name="C04", en="C04. On how many days per week are meals served?", pt="C04. Em quantos dias por semana são servidas refeições?",
      pt_paper="C04. Em quantos dias por mês são servidas refeições?",       # planted LANG-MISALIGN
      relevance="${C02}=1", required="yes", r1=True, paper=True, constraint=".>=0 and .<=7", cmsg_en="0–7.", cmsg_pt="0–7."),
 dict(type="end group", name="section_c", r1=True, paper=True),
 # ---- Section D (repeat)
 dict(type="begin group", name="section_d", en="SECTION D. Teachers", pt="SECÇÃO D. Professores", r1=True, paper=True),
 dict(type="integer", name="D01", en="D01. How many teachers teach at this school?", pt="D01. Quantos professores ensinam nesta escola?",
      required="yes", r1=True, paper=True, constraint=".>=0 and .<=50", cmsg_en="0–50.", cmsg_pt="0–50."),
 dict(type="begin repeat", name="teacher_list", en="Teacher list", pt="Lista de professores", repeat_count="${D01}", r1=True, paper=True),
 dict(type="calculate", name="t_index", calc="index()", r1=True, paper=False),
 dict(type="text", name="D02", en="D02. Name of teacher ${t_index}", pt="D02. Nome do(a) professor(a) ${t_index}", required="yes", r1=True, paper=True,
      constraint=f'regex(., "{NAME_RE}")', cmsg_en="Letters and spaces only.", cmsg_pt="Apenas letras e espaços."),
 dict(type="select_one sex", name="D03", en="D03. Sex", pt="D03. Sexo", required="yes", r1=True, paper=True),
 dict(type="select_one yesno", name="D04", en="D04. Teaches reading in grade 3?", pt="D04. Ensina leitura na 3.ª classe?", required="yes", r1=True, paper=True),
 dict(type="end repeat", name="teacher_list", r1=True, paper=True),
 dict(type="end group", name="section_d", r1=True, paper=True),
 # ---- Section E
 dict(type="begin group", name="section_e", en="SECTION E. School council", pt="SECÇÃO E. Conselho de escola", r1=True, paper=True),
 dict(type="select_one yesno", name="E01", en="E01. Does the school have a school council?", pt="E01. A escola tem um conselho de escola?",
      required="yes", r1=True, paper=True, skip_en="If No >> End", skip_pt="Se Não >> Fim"),
 # DELETED in Round-2 paper
 dict(type="integer", name="E02", en="E02. How many times did the council meet last term?", pt="E02. Quantas vezes o conselho se reuniu no último trimestre?",
      relevance="${E01}=1", required="yes", r1=True, paper=False, constraint=".>=0 and .<=20", cmsg_en="0–20.", cmsg_pt="0–20."),
 dict(type="select_multiple council_issues", name="E03", en="E03. Which issues did the council discuss this year?", pt="E03. Que assuntos o conselho discutiu este ano?",
      relevance="${E01}=1", required="yes", r1=True, paper=True,
      hint_en="More than one option can be selected.", hint_pt="Pode seleccionar mais do que uma opção.",
      constraint="not(selected(., '5')) or count-selected(.)=1",
      cmsg_en="'None of these' cannot be combined with other options.", cmsg_pt="'Nenhum destes' não pode ser combinado com outras opções."),
 dict(type="end group", name="section_e", r1=True, paper=True),
 dict(type="end group", name="survey", r1=True, paper=False),
]

# ============================================================================ XLSForm (Round 1)
# The Round-1 form is built ON SurveyCTO's basic template ("empty_instrument.xlsx", the "Basic
# template form" the DIME Wiki's Questionnaire Programming page links to:
# https://docs.google.com/spreadsheets/d/1JroFCu0HqnPsZwt0jAYa7o311VeOZFc6FRYfET4wAsA). A copy sits in
# example/_template/. Its sheets, column order, widths, frozen header and the 31 conditional-formatting
# rules that color each row by field type are kept; the bilingual columns (label:ENG, hint:ENG,
# constraint message:ENG) are inserted beside their counterparts and inherit the header style; the
# help-* sheets ride along untouched. xlsform_io does the carrying.
TEMPLATE = "empty_instrument.xlsx"
COLMAP = {"type": "type", "name": "name", "label:ENG": "en", "label": "pt", "hint:ENG": "hint_en", "hint": "hint_pt",
          "required": "required", "relevance": "relevance", "constraint": "constraint",
          "constraint message:ENG": "cmsg_en", "constraint message": "cmsg_pt", "calculation": "calc",
          "appearance": "appearance", "repeat_count": "repeat_count"}

def _bilingual(header):
    out = []
    for c in header:
        if c == "label": out += ["label:ENG", "label"]
        elif c == "hint": out += ["hint:ENG", "hint"]
        elif c == "constraint message": out += ["constraint message:ENG", "constraint message"]
        else: out.append(c)
    return out

def write_r1_form(path):
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))
    from xlsform_io import XLSFormBook
    book = XLSFormBook(P("_template", TEMPLATE))
    shdr = _bilingual(book.header("survey"))
    rows = [[q.get(COLMAP.get(c, "__none__")) for c in shdr] for q in Q if q.get("r1")]
    book.set_rows("survey", shdr, rows)
    chdr = _bilingual(book.header("choices"))
    crows = [[{"list_name": ln, "value": v, "label:ENG": en, "label": pt}.get(c) for c in chdr]
             for ln, opts in CHOICES.items() for v, en, pt in opts]
    book.set_rows("choices", chdr, crows)
    sthdr = book.header("settings")
    svals = {"form_title": "#1 - Questionário do(a) Director(a) da Escola — Ronda 1", "form_id": "school_head_r1_v1",
             "version": R1_VERSION, "default_language": "POR"}
    book.set_rows("settings", sthdr, [[svals.get(c) for c in sthdr]])
    book.save(path)
    _mark_r1(path); return path

def _mark_r1(path):
    """On top of the template's own formatting, the marks an authoring team adds while working:
    yellow 'translation pending' fills on two PT labels, a red italic hint under review, a
    validation dropdown on `required`, a hidden rarely-used column, a tab color, and one extra rule
    (PT label blank while EN filled). run_checks.py verifies xlsform_io carries all of it."""
    from openpyxl.styles import Font, PatternFill
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
    wb = openpyxl.load_workbook(path); ws = wb["survey"]
    hdr = [c.value for c in ws[1]]
    pending = PatternFill("solid", fgColor="FFF2CC")
    i_name, i_en, i_pt, i_hint_pt = hdr.index("name") + 1, hdr.index("label:ENG") + 1, hdr.index("label") + 1, hdr.index("hint") + 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, i_name).value in ("B03", "E03"): ws.cell(r, i_pt).fill = pending
        if ws.cell(r, i_name).value == "B04": ws.cell(r, i_hint_pt).font = Font(color="C00000", italic=True)
    col_en, col_pt, last = get_column_letter(i_en), get_column_letter(i_pt), ws.max_row
    ws.conditional_formatting.add(f"{col_pt}2:{col_pt}{last}",
        FormulaRule(formula=[f'AND(${col_pt}2="",${col_en}2<>"")'], fill=PatternFill("solid", fgColor="F4CCCC")))
    dv = DataValidation(type="list", formula1='"yes,no"', allow_blank=True)
    col_req = get_column_letter(hdr.index("required") + 1); dv.add(f"{col_req}2:{col_req}{last}"); ws.add_data_validation(dv)
    ws.column_dimensions[get_column_letter(hdr.index("read only") + 1)].hidden = True
    ws.sheet_properties.tabColor = "1F3864"
    wb.save(path)

# ============================================================================ paper docx (Round 2)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

def _opts_text(list_name, lang, extra=True, skip=None):
    opts = list(CHOICES[list_name])
    if extra and PAPER_EXTRA_OPTION[0] == list_name:
        _, v, en, pt = PAPER_EXTRA_OPTION
        opts = [o for o in opts if o[0] != -77] + [(v, en, pt)] + [o for o in opts if o[0] == -77]
    lines = []
    for v, en, pt in opts:
        lab = en if lang == "EN" else pt
        # paper display codes: Yes/No as 1/2, Other as 99 (the do-not-flag exercise)
        if list_name == "yesno": v = 1 if v == 1 else 2
        if v == -77: v = 99
        lines.append(f"{v} = {lab}")
    if skip: lines.append(skip)
    return "\n".join(lines)

def _resp_text(q, lang):
    t = q["type"]
    if t.startswith("select_"):
        skip = q.get("skip_en") if lang == "EN" else q.get("skip_pt")
        return _opts_text(t.split()[1], lang, skip=skip)
    if t == "calculate": return "Prefilled" if lang == "EN" else "Pré-preenchido"
    if t == "enumerator": return "List" if lang == "EN" else "Lista"
    hint = q.get("hint_en") if lang == "EN" else q.get("hint_pt")
    return hint or ("|__|__|__|" if t == "integer" else "")

def _strip_code(label):
    return label.split(". ", 1)[1] if ". " in label[:8] else label

def write_paper(path, lang):
    """Round-2 paper. EN carries a tracked change on C01 and a two-comment thread on B06."""
    doc = Document()
    sec = doc.sections[0]; sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11), Inches(8.5)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"): setattr(sec, m, Inches(0.7))
    title = ("DEMO PRIMARY SCHOOL SURVEY — ROUND 2 (2026)\nINSTRUMENT 1: SCHOOL HEAD QUESTIONNAIRE" if lang == "EN" else
             "INQUÉRITO DEMO ÀS ESCOLAS PRIMÁRIAS — RONDA 2 (2026)\nINSTRUMENTO 1: QUESTIONÁRIO DO(A) DIRECTOR(A) DA ESCOLA")
    for line in title.split("\n"):
        p = doc.add_paragraph(); r = p.add_run(line); r.bold = True; r.font.size = Pt(13)
    note = ("Fictional instrument for the capi-form-lifecycle starter example. All names and places are invented. "
            "Version " + PAPER_DATE + "." if lang == "EN" else
            "Instrumento fictício para o exemplo inicial do capi-form-lifecycle. Todos os nomes e lugares são inventados. "
            "Texto português redigido automaticamente para fins de demonstração. Versão " + PAPER_DATE + ".")
    p = doc.add_paragraph(); r = p.add_run(note); r.italic = True; r.font.size = Pt(9)

    marks = {}   # name -> question-cell paragraph (for tracked change / comments)
    table = None
    for q in Q:
        if not q.get("paper"): continue
        t = q["type"]
        if t == "begin group":
            p = doc.add_paragraph(); r = p.add_run(q["en"] if lang == "EN" else q["pt"]); r.bold = True; r.font.size = Pt(12)
            table = None; continue
        if t == "begin repeat":
            p = doc.add_paragraph(); r = p.add_run((q["en"] if lang == "EN" else q["pt"]) + (" (repeat for each teacher)" if lang == "EN" else " (repetir para cada professor(a))"))
            r.italic = True; table = None; continue
        if t in ("end group", "end repeat"): table = None; continue
        if t == "note":
            p = doc.add_paragraph(); r = p.add_run(q["en"] if lang == "EN" else q["pt"]); r.font.size = Pt(10); table = None; continue
        if table is None:
            table = doc.add_table(rows=1, cols=3); table.style = "Table Grid"
            for i, h in enumerate(["Code", "Question", "Response"] if lang == "EN" else ["Código", "Pergunta", "Resposta"]):
                table.rows[0].cells[i].text = h; table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        label = (q.get("en_paper") or q["en"]) if lang == "EN" else (q.get("pt_paper") or q["pt"])
        code = q["name"].replace("_o", ".1")
        cells = table.add_row().cells
        cells[0].text = code; cells[1].text = _strip_code(label); cells[2].text = _resp_text(q, lang)
        marks[q["name"]] = cells[1].paragraphs[0]
        for c in cells:
            for par in c.paragraphs:
                for run in par.runs: run.font.size = Pt(10)

    if lang == "EN":
        _tracked_change(marks["C01"], _strip_code(Q_by("C01")["en"]), _strip_code(Q_by("C01")["en_paper"]))
        _comment_anchor(marks["B06"], [0, 1])
    doc.save(path)
    if lang == "EN": _add_comment_parts(path, [
        dict(id=0, author="PI (fictional)", date="2026-01-03T10:00:00Z", para="11111111",
             text="New question for Round 2: we need to know whether heads can receive the monthly bulletin by WhatsApp. Skip if no phone."),
        dict(id=1, author="Programmer (fictional)", date="2026-01-04T09:30:00Z", para="22222222", parent="11111111",
             text="Agreed. Will add as B06 with relevance on B04 not equal to -99."),
    ])
    return path

def Q_by(name): return next(q for q in Q if q["name"] == name)

def _run(text, deleted=False):
    r = OxmlElement("w:r"); t = OxmlElement("w:delText" if deleted else "w:t")
    t.set(qn("xml:space"), "preserve"); t.text = text; r.append(t); return r

def _tracked_change(paragraph, old, new):
    p = paragraph._p
    for r in list(p.findall(qn("w:r"))): p.remove(r)
    d = OxmlElement("w:del"); d.set(qn("w:id"), "901"); d.set(qn("w:author"), "PI (fictional)"); d.set(qn("w:date"), "2026-01-03T10:05:00Z")
    d.append(_run(old, deleted=True)); p.append(d)
    i = OxmlElement("w:ins"); i.set(qn("w:id"), "902"); i.set(qn("w:author"), "PI (fictional)"); i.set(qn("w:date"), "2026-01-03T10:05:00Z")
    i.append(_run(new)); p.append(i)

def _comment_anchor(paragraph, ids):
    p = paragraph._p; runs = p.findall(qn("w:r"))
    first = runs[0] if runs else None
    for cid in ids:
        s = OxmlElement("w:commentRangeStart"); s.set(qn("w:id"), str(cid))
        (first.addprevious(s) if first is not None else p.append(s))
    for cid in ids:
        e = OxmlElement("w:commentRangeEnd"); e.set(qn("w:id"), str(cid)); p.append(e)
        r = OxmlElement("w:r"); ref = OxmlElement("w:commentReference"); ref.set(qn("w:id"), str(cid)); r.append(ref); p.append(r)

def _add_comment_parts(path, comments):
    # Verified to open in Word: the commentsExtended content type MUST be the
    # openxmlformats one below (the older vnd.ms-word.* type makes Word report the file as corrupted).
    cx = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          f'<w:comments xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="{W15}">']
    ex = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', f'<w15:commentsEx xmlns:w="{W}" xmlns:w15="{W15}">']
    for c in comments:
        cx.append(f'<w:comment w:id="{c["id"]}" w:author="{c["author"]}" w:date="{c["date"]}" w:initials="X">'
                  f'<w:p w14:paraId="{c["para"]}" w14:textId="{c["para"]}"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
                  f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>'
                  f'<w:r><w:t xml:space="preserve">{c["text"]}</w:t></w:r></w:p></w:comment>')
        par = f' w15:paraIdParent="{c["parent"]}"' if c.get("parent") else ""
        ex.append(f'<w15:commentEx w15:paraId="{c["para"]}"{par} w15:done="0"/>')
    cx.append("</w:comments>"); ex.append("</w15:commentsEx>")
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.decode("utf-8").replace("</Types>",
                    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
                    '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>').encode("utf-8")
            elif item.filename == "word/_rels/document.xml.rels":
                data = data.decode("utf-8").replace("</Relationships>",
                    '<Relationship Id="rIdCmt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
                    '<Relationship Id="rIdCmt2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>').encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("word/comments.xml", "\n".join(cx)); zout.writestr("word/commentsExtended.xml", "\n".join(ex))
    os.replace(tmp, path)

# ============================================================================ attachments, configs, logs
def write_csvs():
    with open(P("07_capi_attachments", "1_school_head", "upload_school.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["school_id_key", "school_name", "district"])
        w.writerow([9997, "EP Test One", "Demo North"]); w.writerow([9998, "EP Test Two", "Demo North"]); w.writerow([9999, "EP Test Three", "Demo South"])
    with open(P("08_SurveyCTO_datasets", "enumerators_demo.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["id", "name", "users"])
        w.writerow([1, "Enumerator One (test)", "demo_user1"]); w.writerow([2, "Enumerator Two (test)", "demo_user2"])

def write_configs():
    cfg = {"titles": {"EN": "DEMO PRIMARY SCHOOL SURVEY — SCHOOL HEAD QUESTIONNAIRE", "PT": "INQUÉRITO DEMO ÀS ESCOLAS PRIMÁRIAS — QUESTIONÁRIO DO(A) DIRECTOR(A)"},
           "var_labels": {"A05": {"EN": "the enumerator", "PT": "o(a) inquiridor(a)"}}}
    json.dump(cfg, open(P("05_python_scripts", "paper_configs", "1_school_head.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    batch = {"out": "../../06_output_paper_FROMCAPI", "langs": "EN=label:ENG,PT=label", "suffix": "FROMCAPI_r1",
             "forms": [{"form": "../../03_input_prior_capi/1_school_head_r1_v1.xlsx", "base": "1_school_head", "config": "1_school_head.json"}]}
    json.dump(batch, open(P("05_python_scripts", "paper_configs", "batch.json"), "w", encoding="utf-8"), indent=2)

def write_text(rel, text):
    with open(P(*rel.split("/")), "w", encoding="utf-8") as f: f.write(text.lstrip("\n"))

# ============================================================================ the brief
BRIEF = f"""
# CAPI project brief — DEMO Primary School Survey (fictional starter example)

Filled from `references/project_brief_TEMPLATE.md`. Every value below is invented for the
capi-form-lifecycle starter example; nothing refers to a real school, person or project.

## 1. Project identity

| Item | Value |
|---|---|
| Project name and short code | Demo Primary School Survey — `DEMO` |
| Study type and design | Fictional two-arm school-level RCT, 60 schools, randomized by school |
| Rounds and dates | Round 1 (2025) done; **this brief serves Round 2 (2026)**. A Round-1 XLSForm exists for instrument 1 only. |
| Sample geography | Two invented districts: Demo North, Demo South |
| Principal investigators | "PI (fictional)" |
| Implementing / funding partners | none (example) |
| Field data-collection firm | "Demo Field Team": deploys, tests and enumerates |
| SurveyCTO server | `demo.surveycto.com` (placeholder; nothing is uploaded from the example) |
| Languages | Design language English (`label:ENG`); field language Portuguese (un-tagged `label`; `default_language = POR`) |

**Ownership.** Form content: PI. Translation: field team's translator (role). Preload data: data
manager (role). SurveyCTO deployment: field team. Data cleaning: data manager. Skill operator: you.

## 2. Instrument register

| # | Short name | Instrument (full name) | Respondent | Prior-round form_id (blank if none) | Built from | Current form_id | File stem | Status | Notes |
|---|---|---|---|---|---|---|---|---|---|
| 1 | head | School head questionnaire | School head | `school_head_r1_v1` | prior round | | `1_school_head` | to build (Round 2) | the worked example |
| 2 | teacher | Teacher questionnaire | Grade-3 teacher | | scratch | | `2_teacher` | not included in the example | listed to show a from-scratch row |

## 3. Raw materials and where they live

| Material | Location | Naming / dating convention | Notes |
|---|---|---|---|
| Paper questionnaires, EN | `01_input_paper_EN/{PAPER_DATE}/` | leading instrument number; date subfolder | EN docx carries a PI comment thread and one tracked change |
| Paper questionnaires, PT | `02_input_paper_PT/{PAPER_DATE}/` | same | PT text is a machine draft for demo purposes |
| Prior-round XLSForms (if any) | `03_input_prior_capi/` | `<stem>_r1_v1.xlsx` | never edited |
| Current-round XLSForms | `04_output_r2_capi/v<n>_draft/` and `v<n>_deployed/` | | mirror updated same day as any upload |
| Preload CSVs and media | `07_capi_attachments/1_school_head/` | one subfolder per instrument | `upload_school.csv`: header + 3 test rows |
| Server-side datasets | `08_SurveyCTO_datasets/` | | `enumerators_demo.csv` (id, name, users) |
| Prior-round data for preloads | none in the example | | |
| Scripts | `05_python_scripts/` | build / patch scripts land here; `paper_configs/` holds PAPER-mode JSON | |

**Data-governance boundary.** The example contains no personal data. The rule the example
demonstrates: the skill reads preload CSVs by header row only, never opens rows of survey data, and
asks for a variable list rather than a data extract. No folder in the example is off-limits.

## 4. Authority order (which source wins)

| Disagreement about | Winner | Reasoning | Agreed |
|---|---|---|---|
| Question content, order, options, skip logic | EN paper | PIs sign the EN paper | 2026-01-02, PI |
| Response codes, list names, ID conventions, appearance | Round-1 XLSForm | comparability with Round-1 data | 2026-01-02, PI |
| Text of Portuguese | PT paper | translator's text | 2026-01-02, PI |
| PT change with no matching EN change | flag, never adopt | content is decided in EN | 2026-01-02, PI |
| Form vs paper after console edits | the form; paper regenerated from it | | 2026-01-02, PI |
| Team feedback vs paper | paper unless PI agrees | | 2026-01-02, PI |

**Round lifecycle.** Paper → BUILD → deploy → PATCH cycles → PAPER regenerated for training → PATCH →
PAPER regenerated before final deploy. Regeneration tool: `scripts/paper_from_xlsform.py` with
`05_python_scripts/paper_configs/`. Skip-logic prose is added by the field team's trainer.

## 5. Translation policy

**Reuse first, then draft and flag** (chosen 2026-01-02, PI). Look for existing PT text (PT paper,
sibling form, Round-1 form, glossary) first; only a genuinely new item may be drafted, marked in the
plan and appended to the flags file.

| Item | Value |
|---|---|
| Policy chosen | reuse first, then draft and flag |
| Professional reviewer (role) | field team's translator |
| Translation-flags file | `00_reference/r2_translation_flags.md` |
| Proper nouns and codes | written identically in both columns, no flag |
| Glossary | `00_reference/demo_glossary_EN_PT.md` |

## 6. Coding conventions

| Convention | Value | Origin |
|---|---|---|
| Yes/No list values | `yesno`: 1 = Yes, 0 = No (paper shows 1/2: display only) | Round-1 form |
| "Other, specify" value | -77 (paper shows 99: display only) | Round-1 form |
| "Don't know" value | -88 | Round-1 form |
| "Refused" value | -66 | Round-1 form |
| Numeric sentinel (phones, ages, counts) | -99 | Round-1 form |
| Exclusive option in select_multiple | `not(selected(., '<code>')) or count-selected(.)=1`, bilingual message | Round-1 E03 |
| Question-code prefix in labels | yes, both languages (`A01. …`) | Round-1 form |
| Label columns | `label:ENG` (EN), `label` (PT) | Round-1 form |
| `default_language` | `POR` | Round-1 form |
| Phone constraint | `(.>=100000000 and .<=999999999) or .=-99` (fictional 9-digit range) | Round-1 B04 |
| Person-name constraint | `regex(., "^[A-Za-zÀ-ÿ '’.-]+$")` | Round-1 A06 |
| Age / date rules | age allows -99; years-in-post ≤ age unless age = -99 | Round-1 B02/B03 |
| Decimal separator | period | |
| Required by default? | yes for content fields | |
| Repeats vs chained prompts | repeat groups | Round-1 D02 |
| Paper display codes that are NOT data codes | 1/2 for Yes/No, 99 for Other | |

## 7. Identifier scheme

| Item | Value |
|---|---|
| Primary cluster key | `A01` (school ID), preload key column `school_id_key` |
| Person / role IDs | role-based: `concat(${{A01}}, '<suffix>')` |
| Secondary keys | teachers: repeat index within the school |
| Two roles, one person | keep both IDs |
| Registry file | `00_reference/demo_id_suffix_registry.md` |

## 8. Versioning and deployment

| Item | Value |
|---|---|
| `version` stamp format | `YYMMDDNNNN`; always increases; must exceed the server value |
| `form_id` at the start of a round | `<prior_form_id>` with `_r1_` → `_r2_`: instrument 1 becomes `school_head_r2_v1` |
| `form_id` on patches within a round | unchanged; bump only for a new data stream |
| Deploy mirror folder | `04_output_r2_capi/v<n>_deployed/` with a README version table |
| Attachment layout | one subfolder per instrument; shared CSVs byte-identical |
| Enumerator dataset | `enumerators_demo` (ENUMERATORS-type); columns id, name, users |
| Test IDs | school IDs 9997–9999; phone 100000000 |
| Media limits | < 100 MB per file, 300 MB per form |
| Device re-sync rule | delete form + Get Blank Form after attachment changes |
| Same-day mirror rule | every console edit or upload → xlsx downloaded into the mirror |

## 9. Logs the project keeps

| Log | Path | What goes in | Written by |
|---|---|---|---|
| Change log | `CHANGELOG.md` | every build and patch, newest on top | skill |
| Flags and decisions | `00_reference/r2_flags_and_decisions.md` | design and translation decisions | skill |
| To-dos | `00_reference/r2_todos.md` | open actions by owner | skill |
| Deferred patches | `00_reference/r2_deferred_patches.md` | parked fixes | skill |
| Preload spec | `CAPI_PRELOAD_UPDATES_NEEDED.md` | columns each form needs | skill |
| Translation flags | `00_reference/r2_translation_flags.md` | drafted PT text | skill |
| Session file changes | `SESSION_FILE_CHANGES.md` | every file touched | skill |
| Team testing feedback | not in the example | | team |

## 10. Folder layout (as it actually is)

```
example/
  00_reference/            this brief, ID registry, glossary, (logs the skill creates)
  01_input_paper_EN/       Round-2 EN paper, dated subfolder
  02_input_paper_PT/       Round-2 PT paper, dated subfolder
  03_input_prior_capi/     Round-1 XLSForm (never edited)
  04_output_r2_capi/       BUILD and PATCH outputs, v<n>_draft / v<n>_deployed
  05_python_scripts/       build/patch scripts the skill writes; paper_configs/ for PAPER mode
  06_output_paper_FROMCAPI/ regenerated paper, one subfolder per language tag
  07_capi_attachments/     preload CSVs per instrument
  08_SurveyCTO_datasets/   enumerator dataset definition
  CHANGELOG.md, SESSION_FILE_CHANGES.md
```

## 11. Known traps for this project

- `.>0` rejects 0 where the paper allows "0 if less than a year" (B03).
- select_multiple lists with a "none" option need the exclusive constraint (E03).
- Years-in-post constraint must tolerate age = -99 (B03).
- A `village` column is referenced by the Round-2 paper (A04) but is not in `upload_school.csv`:
  the build must flag it for the data manager, never invent it.

## 12. History of this brief

| Date | Change | By |
|---|---|---|
| 2026-01-02 | created from template for the starter example | skill author |
"""

REGISTRY = """
# DEMO — ID-suffix registry

Role-based person IDs are `concat(${A01}, '<suffix>')`. Add a row before using a new suffix.

| Suffix | Role | Instrument(s) | Added |
|---|---|---|---|
| 10 | School head | 1 | 2025-01-15 (Round 1) |
| 20 | Grade-3 teacher | 2 | 2025-01-15 (Round 1) |
"""

GLOSSARY = """
# DEMO — glossary EN ↔ PT

Fixed renderings used in every instrument. PT here is a machine draft for demo purposes.

| EN | PT |
|---|---|
| School head | Director(a) da escola |
| School council | Conselho de escola |
| School feeding program | Programa de alimentação escolar |
| Enumerator | Inquiridor(a) |
| Don't know | Não sabe |
| Other, specify | Outro, especifique |
"""

def main():
    write_text("00_reference/capi_project_brief.md", BRIEF)
    write_text("00_reference/demo_id_suffix_registry.md", REGISTRY)
    write_text("00_reference/demo_glossary_EN_PT.md", GLOSSARY)
    for f in (write_r1_form(P("03_input_prior_capi", "1_school_head_r1_v1.xlsx")),
              write_paper(P("01_input_paper_EN", PAPER_DATE, "1_school_head_EN.docx"), "EN"),
              write_paper(P("02_input_paper_PT", PAPER_DATE, "1_school_head_PT.docx"), "PT")):
        freeze_timestamps(f)
    write_csvs(); write_configs()
    write_text("04_output_r2_capi/README.md", "# Round-2 CAPI outputs\n\nBUILD writes `v1_draft/`; after upload the server download goes in `v1_deployed/`.\n\n| Version | form_id | Uploaded | Notes |\n|---|---|---|---|\n")
    write_text("CHANGELOG.md", "# Changelog — DEMO Primary School Survey\n\n(empty; BUILD and PATCH append here, newest on top)\n")
    write_text("SESSION_FILE_CHANGES.md", "# Session file/folder change log\n\n(empty; every mode appends the files it touched)\n")
    print("example written under", HERE)

if __name__ == "__main__":
    main()
