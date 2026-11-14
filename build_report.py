# -*- coding: utf-8 -*-
"""Build the Data Mining Lab report (.docx) from the executed notebook outputs."""

import os
import re
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ----------------------------------------------------------------------------
# Palette, fonts, and global counters
# ----------------------------------------------------------------------------
NAVY = "1F3864"
BLUE = "2E5A88"
TEAL = "2E7D7B"
LIGHT_TEAL = "EAF4F3"
PALE_BLUE = "F2F5FA"
GRAY = "595959"
MID_GRAY = "7F7F7F"
BORDER = "D6DCE4"
WHITE = "FFFFFF"
BODY_FONT = "Cambria"
HEAD_FONT = "Calibri"
MONO_FONT = "Consolas"

TABLE_NO = 0
FIGURE_NO = 0

ASSETS = "report_assets"


def rgb(hexstr):
    return RGBColor.from_string(hexstr)


# ----------------------------------------------------------------------------
# Low-level XML helpers
# ----------------------------------------------------------------------------
def _remove_child(parent, tag):
    for child in parent.findall(qn(tag)):
        parent.remove(child)


def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    _remove_child(tcPr, "w:shd")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_margins(table, top=55, bottom=55, left=100, right=100):
    tblPr = table._tbl.tblPr
    _remove_child(tblPr, "w:tblCellMar")
    mar = OxmlElement("w:tblCellMar")
    for side, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def set_table_layout(table, widths):
    tblPr = table._tbl.tblPr
    _remove_child(tblPr, "w:tblW")
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(sum(widths) * 1440)))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)
    _remove_child(tblPr, "w:tblLayout")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for i, width in enumerate(widths):
        table.columns[i].width = Inches(width)
    for row in table.rows:
        for i, width in enumerate(widths):
            row.cells[i].width = Inches(width)


def set_table_borders(table):
    tblPr = table._tbl.tblPr
    _remove_child(tblPr, "w:tblBorders")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "8")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), NAVY)
        elif edge == "insideH":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), BORDER)
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trPr.append(header)


def cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def paragraph_border(paragraph, color=BORDER, size=6):
    pPr = paragraph._p.get_or_add_pPr()
    _remove_child(pPr, "w:pBdr")
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "6")
        el.set(qn("w:color"), color)
        pBdr.append(el)
    pPr.append(pBdr)


def add_field(paragraph, code, placeholder="1"):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {code} "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(text)
    run._r.append(fld_end)


# ----------------------------------------------------------------------------
# Text helpers
# ----------------------------------------------------------------------------
def add_rich(paragraph, text, size=None, color=None, bold=False, italic=False, font=None):
    """Add text supporting **bold** and `monospace` inline markers."""
    tokens = re.split(r"(\*\*.*?\*\*|`.*?`)", text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) > 4:
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("`") and token.endswith("`") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run.font.name = MONO_FONT
            run.font.size = Pt((size or 10.5) - 1.0)
            run.font.color.rgb = rgb(BLUE)
        else:
            run = paragraph.add_run(token)
        if size and not (token.startswith("`") and token.endswith("`")):
            run.font.size = Pt(size)
        if color:
            run.font.color.rgb = rgb(color)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if font:
            run.font.name = font
    return paragraph


def body(text, space_after=6, keep_with_next=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.keep_with_next = keep_with_next
    add_rich(p, text, size=10.5, italic=italic)
    return p


def bullets(items, size=10.5):
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.30)
        p.paragraph_format.first_line_indent = Inches(-0.18)
        p.paragraph_format.space_after = Pt(3)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_rich(p, "•  " + item, size=size)


def numbered(items, size=10.5):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.38)
        p.paragraph_format.first_line_indent = Inches(-0.28)
        p.paragraph_format.space_after = Pt(4)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_rich(p, f"{i}.  " + item, size=size)


def h1(text, page_break=True):
    if page_break:
        doc.add_page_break()
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.page_break_before = False
    p.add_run(text)
    return p


def h2(text, page_break=False):
    p = doc.add_paragraph(style="Heading 2")
    if page_break:
        p.paragraph_format.page_break_before = True
    p.add_run(text)
    return p


def spacer(points=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(points)


def caption_table(text):
    global TABLE_NO
    TABLE_NO += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(f"Table {TABLE_NO}. ")
    r.bold = True
    r.font.name = HEAD_FONT
    r.font.size = Pt(8.5)
    r.font.color.rgb = rgb(NAVY)
    r2 = p.add_run(text)
    r2.font.name = HEAD_FONT
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = rgb(GRAY)
    r2.italic = True


def caption_figure(text):
    global FIGURE_NO
    FIGURE_NO += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(f"Figure {FIGURE_NO}. ")
    r.bold = True
    r.font.name = HEAD_FONT
    r.font.size = Pt(8.5)
    r.font.color.rgb = rgb(NAVY)
    r2 = p.add_run(text)
    r2.font.name = HEAD_FONT
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = rgb(GRAY)
    r2.italic = True


def add_table(headers, rows, widths, caption, font_size=8.5, aligns=None, first_col_bold=False):
    caption_table(caption)
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    set_cell_margins(table)
    set_table_layout(table, widths)

    if aligns is None:
        aligns = [WD_ALIGN_PARAGRAPH.CENTER] * len(headers)

    header_row = table.rows[0]
    repeat_header(header_row)
    cant_split(header_row)
    for j, header in enumerate(headers):
        cell = header_row.cells[j]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(cell, NAVY)
        p = cell.paragraphs[0]
        p.alignment = aligns[j] if len(aligns) == len(headers) else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(header)
        r.bold = True
        r.font.name = HEAD_FONT
        r.font.size = Pt(font_size)
        r.font.color.rgb = rgb(WHITE)

    for i, row in enumerate(rows):
        cells = table.add_row().cells
        cant_split(table.rows[-1])
        for j, value in enumerate(row):
            cell = cells[j]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i % 2 == 1:
                shade_cell(cell, PALE_BLUE)
            p = cell.paragraphs[0]
            p.alignment = aligns[j]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.0
            add_rich(p, str(value), size=font_size)
            for run in p.runs:
                run.font.name = HEAD_FONT
                if first_col_bold and j == 0:
                    run.bold = True
    spacer(2)


def add_figure(filename, caption, width=6.2):
    path = os.path.join(ASSETS, filename)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    run.add_picture(path, width=Inches(width))
    paragraph_border(p, color=BORDER, size=4)
    caption_figure(caption)


def callout(label, text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_layout(table, [6.27])
    set_cell_margins(table, top=110, bottom=110, left=160, right=160)
    cell = table.rows[0].cells[0]
    shade_cell(cell, LIGHT_TEAL)
    tcPr = cell._tc.get_or_add_tcPr()
    _remove_child(tcPr, "w:tcBorders")
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        if edge == "left":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "24")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), TEAL)
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(label.upper())
    r.bold = True
    r.font.name = HEAD_FONT
    r.font.size = Pt(8)
    r.font.color.rgb = rgb(TEAL)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    add_rich(p2, text, size=9.8)
    for run in p2.runs:
        run.font.name = BODY_FONT if run.font.name != MONO_FONT else MONO_FONT
    spacer(6)


# ----------------------------------------------------------------------------
# Document setup
# ----------------------------------------------------------------------------
doc = Document()

section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.3)
section.bottom_margin = Cm(2.2)
section.left_margin = Cm(2.4)
section.right_margin = Cm(2.4)
section.header_distance = Cm(1.1)
section.footer_distance = Cm(1.0)
section.different_first_page_header_footer = True

styles = doc.styles
normal = styles["Normal"]
normal.font.name = BODY_FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = rgb("1A1A1A")
normal.paragraph_format.line_spacing = 1.15
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.widow_control = True
rpr = normal.element.get_or_add_rPr()
rfonts = rpr.get_or_add_rFonts()
for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
    rfonts.set(qn(attr), BODY_FONT)

heading_specs = {
    "Heading 1": (18, NAVY, 16, 7),
    "Heading 2": (12.5, BLUE, 12, 4),
    "Heading 3": (11, TEAL, 10, 3),
}
for name, (size, color, before, after) in heading_specs.items():
    st = styles[name]
    st.font.name = HEAD_FONT
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = rgb(color)
    st.paragraph_format.space_before = Pt(before)
    st.paragraph_format.space_after = Pt(after)
    st.paragraph_format.keep_with_next = True
    st.paragraph_format.line_spacing = 1.0
    r = st.element.get_or_add_rPr()
    rf = r.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(attr), HEAD_FONT)
    # Remove theme color so the explicit navy/blue is preserved.
    _remove_child(r, "w:color")
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), color)
    r.append(color_el)

# Header and footer (not on the title page)
header = section.header
hp = header.paragraphs[0]
hp.paragraph_format.tab_stops.add_tab_stop(Cm(15.9), WD_TAB_ALIGNMENT.RIGHT)
hp.paragraph_format.space_after = Pt(0)
r = hp.add_run("Data Mining Lab Report  |  Airline Passenger Satisfaction")
r.font.name = HEAD_FONT
r.font.size = Pt(8)
r.font.color.rgb = rgb(MID_GRAY)
r = hp.add_run("\tFeni University")
r.font.name = HEAD_FONT
r.font.size = Pt(8)
r.font.color.rgb = rgb(MID_GRAY)
paragraph_border(hp, color=BORDER, size=4)

footer = section.footer
fp = footer.paragraphs[0]
fp.paragraph_format.tab_stops.add_tab_stop(Cm(15.9), WD_TAB_ALIGNMENT.RIGHT)
fp.paragraph_format.space_before = Pt(2)
r = fp.add_run("Taisin Nigar Tisha  |  ID 232031033")
r.font.name = HEAD_FONT
r.font.size = Pt(8)
r.font.color.rgb = rgb(MID_GRAY)
r = fp.add_run("\tPage ")
r.font.name = HEAD_FONT
r.font.size = Pt(8)
r.font.color.rgb = rgb(MID_GRAY)
add_field(fp, "PAGE")
r = fp.add_run(" of ")
r.font.name = HEAD_FONT
r.font.size = Pt(8)
r.font.color.rgb = rgb(MID_GRAY)
add_field(fp, "NUMPAGES")
for run in fp.runs:
    run.font.name = HEAD_FONT
    run.font.size = Pt(8)
    run.font.color.rgb = rgb(MID_GRAY)

# PAGE and NUMPAGES are updated by Word/Office when the document is rendered.

# ----------------------------------------------------------------------------
# Title page
# ----------------------------------------------------------------------------
band = doc.add_table(rows=2, cols=1)
band.autofit = False
set_table_layout(band, [6.27])
set_cell_margins(band, top=0, bottom=0, left=0, right=0)
band.rows[0].height = Cm(0.20)
band.rows[1].height = Cm(0.06)
shade_cell(band.rows[0].cells[0], NAVY)
shade_cell(band.rows[1].cells[0], TEAL)

for _ in range(3):
    spacer(15)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_after = Pt(2)
r = p.add_run("FENI UNIVERSITY")
r.font.name = HEAD_FONT
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = rgb(NAVY)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Data Mining Lab  ·  Project Report")
r.font.name = HEAD_FONT
r.font.size = Pt(10)
r.font.color.rgb = rgb(TEAL)

for _ in range(5):
    spacer(18)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run("Service Optimization Using Airline Passenger Satisfaction Data")
r.font.name = "Georgia"
r.font.size = Pt(28)
r.font.bold = True
r.font.color.rgb = rgb(NAVY)
p.paragraph_format.line_spacing = 1.05

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(8)
r = p.add_run("A Complete Data Mining Lab Report")
r.font.name = HEAD_FONT
r.font.size = Pt(15)
r.font.color.rgb = rgb(TEAL)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
r = p.add_run(
    "Preprocessing, exploratory analysis, classification, clustering, and an "
    "end-to-end service optimization study of 129,880 airline passengers."
)
r.font.name = BODY_FONT
r.font.size = Pt(10.5)
r.font.italic = True
r.font.color.rgb = rgb(GRAY)

for _ in range(4):
    spacer(18)

info = [
    ("Course", "Data Mining Lab"),
    ("Submitted by", "Taisin Nigar Tisha"),
    ("Student ID", "232031033"),
    ("Batch", "31st (UG), Feni University"),
    ("Instructor", "Argina Akter, Lecturer"),
    ("Date", "November 2026"),
    ("Source notebook", "notebook/service_optimization.ipynb"),
]
info_table = doc.add_table(rows=0, cols=2)
info_table.autofit = False
set_table_layout(info_table, [1.55, 4.35])
for label, value in info:
    cells = info_table.add_row().cells
    p = cells[0].paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(label)
    r.font.name = HEAD_FONT
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = rgb(NAVY)
    p = cells[1].paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(value)
    r.font.name = BODY_FONT
    r.font.size = Pt(10)
    r.font.color.rgb = rgb("1A1A1A")

for _ in range(3):
    spacer(14)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(0)
r = p.add_run(
    "All metrics, tables, and figures in this report are taken from the saved outputs of the "
    "executed notebook. No results have been simulated or re-estimated."
)
r.font.name = HEAD_FONT
r.font.size = Pt(8)
r.font.italic = True
r.font.color.rgb = rgb(MID_GRAY)

# ----------------------------------------------------------------------------
# Contents
# ----------------------------------------------------------------------------
doc.add_page_break()
p = doc.add_paragraph(style="Heading 1")
p.paragraph_format.page_break_before = False
p.add_run("Contents")

contents = [
    ("0", "Introduction and Project Overview", "Problem, dataset, environment, and scope of the report."),
    ("1", "Lab 1 — Introduction to Python for Data Mining", "Python environment, core language constructs, functions and loops, and CSV handling."),
    ("2", "Lab 2 — Data Preprocessing and Exploration", "Missing values, duplicates, encoding, scaling, and a leakage-safe split."),
    ("3", "Lab 3 — Exploratory Data Analysis", "Target balance, passenger groups, correlations, and service ratings."),
    ("6", "Lab 6 — Decision Tree Classification", "An interpretable nonlinear baseline for satisfaction prediction."),
    ("7", "Lab 7 — Naive Bayes Classification", "A fast probabilistic baseline and the cost of its independence assumption."),
    ("8", "Lab 8 — Support Vector Machine Classification", "A kernel SVM trained under a documented computational adaptation."),
    ("9", "Lab 9 — K-Means Clustering", "Passenger segmentation by service experience, with k chosen by silhouette score."),
    ("10", "Lab 10 — Hierarchical Clustering", "Agglomerative clustering, a dendrogram, and a comparison with K-Means."),
    ("11", "Lab 11 — Outlier Detection and Ensemble Methods", "IQR screening, Random Forest, and permutation feature importance."),
    ("12", "Lab 12 — End-to-End Mini Project", "Feature engineering, model selection, tuning, service recommendations, and the final model."),
    ("13", "References", "Dataset, software, and methodological sources."),
]
for num, title, desc in contents:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.left_indent = Inches(0.35)
    p.paragraph_format.first_line_indent = Inches(-0.35)
    r = p.add_run(f"{num}   ")
    r.font.name = HEAD_FONT
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = rgb(TEAL)
    r = p.add_run(title)
    r.font.name = HEAD_FONT
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = rgb(NAVY)
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.35)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(desc)
    r.font.name = BODY_FONT
    r.font.size = Pt(8.8)
    r.font.italic = True
    r.font.color.rgb = rgb(GRAY)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(16)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("A note on scope")
r.font.name = HEAD_FONT
r.font.size = Pt(9)
r.font.bold = True
r.font.color.rgb = rgb(NAVY)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.space_after = Pt(0)
add_rich(
    p,
    "The notebook moves directly into applied work: imports and CSV loading, then Lab 2 preprocessing. "
    "The Python fundamentals required by Lab 1 are documented in Section 1 from the code the project "
    "actually uses. "
    "Labs 4 and 5 are not implemented, and Labs 10 and 11 are implemented only in part. Those boundaries "
    "are stated explicitly rather than filled with invented results; Section 0.6 maps every reported "
    "section to the executed notebook.",
    size=9,
    color=GRAY,
)

# ============================================================================
# 0. INTRODUCTION AND PROJECT OVERVIEW
# ============================================================================
doc.add_page_break()
p = doc.add_paragraph(style="Heading 1")
p.paragraph_format.page_break_before = False
p.add_run("0. Introduction and Project Overview")

h2("0.1 Problem Statement")
body(
    "Airlines collect a large amount of passenger feedback at every stage of the journey: booking, "
    "check-in, boarding, the in-flight experience, and baggage collection. In practice, this feedback "
    "is often stored as raw survey data and is rarely converted into concrete service improvement "
    "decisions. The result is a gap between measurement and action. An airline may know that its "
    "overall satisfaction score is mediocre, but not which specific service touchpoints are most "
    "strongly associated with satisfaction, which passenger groups are most at risk of dissatisfaction, "
    "or where a limited service budget would produce the greatest improvement."
)
body(
    "This project addresses that gap using the Airline Passenger Satisfaction dataset. The task is "
    "framed as a supervised binary classification problem: predict whether a passenger is Satisfied "
    "or Neutral or Dissatisfied from demographic, trip, and service-rating information. Prediction "
    "alone, however, is not the final objective. The project also uses exploratory analysis, feature "
    "importance, and clustering to explain why the model makes its predictions, so that the results "
    "can be translated into service optimization priorities rather than left as an abstract accuracy "
    "score."
)

h2("0.2 Motivation")
body(
    "Passenger satisfaction affects airline reputation, customer retention, and revenue. A dissatisfied "
    "passenger is less likely to return and less likely to recommend the airline, while a satisfied "
    "passenger represents repeat business and positive word of mouth. Simple averages and manual review "
    "cannot reliably separate the effect of one service from another, because service ratings are "
    "correlated, passenger groups differ systematically, and some effects are nonlinear. For example, "
    "the difference between an Online Boarding rating of 2 and 3 may matter more than the difference "
    "between 4 and 5."
)
body(
    "Data mining provides the tools to handle this complexity. Classification models can combine many "
    "weak and correlated signals into a usable prediction; permutation importance can rank the features "
    "that the fitted model actually relies on; clustering can reveal passenger segments defined by "
    "shared service experience; and a leakage-safe evaluation can establish whether the results are "
    "likely to generalize beyond the training sample. The dataset is large enough (129,880 passengers) "
    "to support these analyses and rich enough (14 separate service ratings) to make the service "
    "optimization question meaningful."
)

h2("0.3 Objectives")
body(
    "The project follows the objectives stated in the notebook. Each objective maps to a concrete part "
    "of the analysis reported in the following sections."
)
numbered([
    "Explore and understand the structure of the Airline Passenger Satisfaction dataset, including its "
    "feature groups, target balance, missing values, and data types.",
    "Preprocess the data by handling missing values, removing irrelevant identifiers, encoding "
    "categorical variables, and scaling features without leaking test information into training.",
    "Perform exploratory data analysis to uncover patterns between service quality, passenger "
    "characteristics, and overall satisfaction.",
    "Engineer and select features that are relevant to predicting passenger satisfaction and useful "
    "for communicating the results.",
    "Train and compare multiple classification models — Decision Tree, Naive Bayes, Support Vector "
    "Machine, and Random Forest — using accuracy, precision, recall, F1-score, and ROC-AUC.",
    "Apply K-Means clustering to segment passengers by service experience, and compare the result with "
    "hierarchical clustering.",
    "Translate the analysis into concrete service optimization recommendations grounded in the measured "
    "results.",
    "Package the best-performing model as a complete, reusable pipeline suitable for the project's "
    "FastAPI and ReactJS prediction application.",
])

h2("0.4 Dataset Description")
body(
    "The analysis uses the Airline Passenger Satisfaction dataset distributed through Kaggle as "
    "`airline_passenger_satisfaction.csv`. The file contains 129,880 passenger records and 24 columns: "
    "23 predictor columns and one target column. Each row represents one passenger, and there is no "
    "time dimension, so the study is cross-sectional rather than longitudinal."
)
body(
    "The target variable is `Satisfaction`, with two classes: `Satisfied` and `Neutral or Dissatisfied`. "
    "The classes are mildly imbalanced: 73,452 passengers (56.55%) are Neutral or Dissatisfied and "
    "56,428 (43.45%) are Satisfied. Only `Arrival Delay` contains missing values — 393 of 129,880 rows, "
    "or approximately 0.30%. No duplicate rows were found. The `ID` column is a unique passenger "
    "identifier with 129,880 distinct values and carries no predictive information, so it is removed "
    "before modeling."
)
add_table(
    ["Feature group", "Columns", "Count", "Role in the analysis"],
    [
        ["Passenger profile", "Gender, Age, Customer Type", "3", "Demographic predictors"],
        ["Trip profile", "Type of Travel, Class, Flight Distance, Departure Delay, Arrival Delay", "5", "Travel and operational predictors"],
        ["Service ratings", "Departure and Arrival Time Convenience, Ease of Online Booking, Check-in Service, Online Boarding, Gate Location, On-board Service, Seat Comfort, Leg Room Service, Cleanliness, Food and Drink, In-flight Service, In-flight Wifi Service, In-flight Entertainment, Baggage Handling", "14", "Ordinal ratings on a 0–5 scale (0 = not applicable)"],
        ["Identifier", "ID", "1", "Unique row identifier; dropped before modeling"],
        ["Target", "Satisfaction", "1", "Binary label: Satisfied or Neutral or Dissatisfied"],
    ],
    [1.15, 3.15, 0.55, 1.42],
    "Feature groups in the airline passenger satisfaction dataset.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "The 14 service-rating columns are the analytical core of the project. According to the dataset's "
    "data dictionary, each is measured on a 0–5 scale where 0 means “not applicable” and 1 to 5 runs "
    "from lowest to highest satisfaction. This convention matters for interpretation: the ratings are "
    "ordinal, not continuous, and the zero category is qualitatively different from a low rating. The "
    "notebook treats the ratings as numeric and scales them, which is a pragmatic choice for distance- "
    "and tree-based models but is revisited as a limitation in Section 12.8."
)
add_table(
    ["Column", "Data type", "Missing", "Treatment and role"],
    [
        ["ID", "int64", "0", "Unique identifier; dropped"],
        ["Gender", "object (str)", "0", "Nominal predictor; one-hot encoded"],
        ["Age", "int64", "0", "Continuous predictor; standardised"],
        ["Customer Type", "object (str)", "0", "Nominal predictor; one-hot encoded"],
        ["Type of Travel", "object (str)", "0", "Nominal predictor; one-hot encoded"],
        ["Class", "object (str)", "0", "Nominal predictor; one-hot encoded"],
        ["Flight Distance", "int64", "0", "Continuous predictor; standardised"],
        ["Departure Delay", "int64", "0", "Continuous predictor; standardised"],
        ["Arrival Delay", "float64", "393", "Continuous predictor; median-imputed, then standardised"],
        ["14 service ratings", "int64", "0", "Ordinal predictors (0–5); used as numeric and standardised"],
        ["Satisfaction", "object (str)", "0", "Target; label-encoded to 0/1"],
    ],
    [1.45, 1.05, 0.65, 3.12],
    "Schema, data types, missing values, and treatment of the 24 columns.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)

h2("0.5 Tools and Environment")
body(
    "The notebook was executed in a Python 3.14.2 virtual environment using Jupyter. Data handling and "
    "numerical work used pandas and NumPy; visualization used Matplotlib and Seaborn; preprocessing, "
    "modeling, and evaluation used scikit-learn; hierarchical clustering and the dendrogram used SciPy; "
    "and model persistence used joblib. A fixed `random_state` of 42 was used throughout, which makes "
    "the train/test split, model initialization, cross-validation folds, and sampling reproducible."
)
body(
    "Two computational adaptations were necessary because the notebook was run on a standard single-CPU "
    "machine. First, the kernel SVM was trained on a stratified 8,000-row subsample of the training set "
    "rather than the full 103,904 rows. Second, the Random Forest grid search was run on a stratified "
    "20,000-row subsample of the training set, after which the selected hyperparameters were refitted on "
    "the full training set. In both cases, evaluation was performed on the full held-out test set, so "
    "the reported test metrics remain comparable across models. These adaptations are documented rather "
    "than hidden because they affect the interpretation of the results."
)

h2("0.6 Scope and Mapping to the Lab Manual")
body(
    "The report follows the task order of the Data Mining Lab Manual, but it documents only what the "
    "notebook actually executes. Lab 1 is a Python-foundations lab: the notebook contains no separate "
    "syntax-exercise block, but its import, data-loading, preprocessing, evaluation, and clustering "
    "cells use the Lab 1 constructs that the applied workflow needs. Section 1 therefore documents the "
    "Lab 1 foundation from that executed code and maps each manual topic to its actual use. Labs 4 "
    "(Apriori) and 5 "
    "(FP-Growth) are not implemented, so they are not represented here. Lab 10 implements agglomerative "
    "clustering and a dendrogram comparison but not DBSCAN, and Lab 11 implements IQR screening and a "
    "Random Forest with permutation importance but not the Z-score method or Isolation Forest. Rather "
    "than inventing results for those missing components, the report states the boundary explicitly and "
    "evaluates the work that was actually performed."
)
add_table(
    ["Lab manual task", "Coverage in the executed notebook", "Report section"],
    [
        ["Lab 1 — Introduction to Python", "Sections 2–3 and throughout (imports, CSV loading, core Python constructs)", "Section 1"],
        ["Lab 2 — Data preprocessing and exploration", "Sections 3–5", "Section 2"],
        ["Lab 3 — Data visualization and EDA", "Section 6", "Section 3"],
        ["Lab 4 — Apriori", "Not implemented", "—"],
        ["Lab 5 — FP-Growth", "Not implemented", "—"],
        ["Lab 6 — Decision Trees", "Sections 9–12 (Decision Tree model)", "Section 6"],
        ["Lab 7 — Naive Bayes", "Sections 9–12 (GaussianNB model)", "Section 7"],
        ["Lab 8 — Support Vector Machines", "Sections 9–12 (RBF SVM model)", "Section 8"],
        ["Lab 9 — K-Means", "Section 14", "Section 9"],
        ["Lab 10 — Hierarchical clustering and DBSCAN", "Section 15 (agglomerative only; no DBSCAN)", "Section 10"],
        ["Lab 11 — Outliers and ensemble methods", "Sections 5.4 and 11–13 (IQR, Random Forest, permutation importance; no Z-score or Isolation Forest)", "Section 11"],
        ["Lab 12 — End-to-end mini project", "Sections 1, 7–8, and 12–18", "Section 12"],
    ],
    [1.75, 3.25, 1.27],
    "Mapping between the lab manual and the executed notebook.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "This mapping also explains the gaps in the section numbering: Sections 4 and 5 are absent because "
    "the corresponding labs were not part of the notebook. The remaining sections follow the manual's "
    "order, with Lab 12 treated as the integrating mini project that draws on the earlier labs."
)

# ============================================================================
# 1. LAB 1 — INTRODUCTION TO PYTHON FOR DATA MINING
# ============================================================================
h1("1. Lab 1 — Introduction to Python for Data Mining")
body(
    "Lab 1 establishes the Python foundation on which every later lab depends: variables and data "
    "types, operators, lists, tuples and dictionaries, functions and loops, and reading and writing "
    "CSV data with pandas. In a conventional lab sequence these topics are practised in small isolated "
    "exercises. In this project they are applied directly to the airline dataset, so the evidence for "
    "Lab 1 is the working code that imports the libraries, loads the CSV file, prepares the feature "
    "lists, evaluates the models, and drives the clustering experiments."
)
body(
    "The executed notebook does not contain a separate syntax-exercise block; its first executable "
    "cells move directly into imports, CSV loading, and applied preprocessing. This section therefore "
    "reports Lab 1 from the code that the notebook actually runs rather than inventing basic exercises "
    "that were "
    "never executed. Every construct described below appears in the saved notebook, and the dataset "
    "values quoted are the outputs that the notebook recorded. The section is a foundation chapter, "
    "not a substitute for the applied labs that follow."
)

h2("1.1 Python Environment and Libraries")
body(
    "The notebook was executed with Python 3.14.2 in a Jupyter environment. The import cell is "
    "organised into functional groups so that the dependencies of each stage are explicit: pandas and "
    "NumPy for data handling, Matplotlib and Seaborn for visualization, scikit-learn for preprocessing, "
    "modeling, and evaluation, SciPy for hierarchical clustering, and joblib for model persistence. "
    "The cell also fixes `RANDOM_STATE = 42`, suppresses non-actionable warnings, and sets a common "
    "plot style and figure size. Fixing the random state at the top of the notebook is what makes the "
    "train/test split, model initialisation, cross-validation folds, and sampling reproducible in every "
    "later section."
)
add_table(
    ["Import group", "Libraries", "Role in the project"],
    [
        ["Data handling", "pandas, NumPy", "Loading the CSV file, tabular operations, numeric arrays"],
        ["Visualization", "Matplotlib, Seaborn", "Histograms, bar charts, box plots, heatmaps, dendrogram"],
        ["Preprocessing and models", "scikit-learn", "Encoding, scaling, pipelines, classifiers, clustering"],
        ["Evaluation", "scikit-learn metrics", "Accuracy, precision, recall, F1, ROC-AUC, confusion matrix"],
        ["Hierarchical clustering", "SciPy", "Ward linkage and dendrogram construction"],
        ["Persistence and utilities", "joblib, warnings", "Saving the fitted pipeline and keeping output clean"],
    ],
    [1.35, 1.95, 3.05],
    "Python libraries imported in the executed notebook and their roles.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "The import block ends with a short verification message, `Libraries imported successfully.`, and "
    "the next cell loads the dataset. This ordering is deliberate: a failure in an import is detected "
    "before any data is read, so a missing or incompatible library cannot be mistaken for a data "
    "problem later in the workflow."
)

h2("1.2 Variables, Data Types, and Operators")
body(
    "Python variables are created by assignment and are not declared with an explicit type. The "
    "project uses this flexibility to keep configuration values, column names, thresholds, and model "
    "objects in named variables. `RANDOM_STATE = 42` is an integer used throughout the notebook; "
    "`target_col = 'Satisfaction'` is a string that names the prediction target; "
    "`median_arrival_delay` is a floating-point value returned by a pandas method; and `df` is a "
    "two-dimensional DataFrame that holds the loaded passenger records. The expression `df.shape` "
    "returns a tuple, `(129880, 24)`, which is the first saved output that confirms the file was read "
    "correctly."
)
body(
    "Operators do most of the analytical work. Arithmetic produces the interquartile range and its "
    "bounds (`iqr = q3 - q1`, `lb = q1 - 1.5 * iqr`, `ub = q3 + 1.5 * iqr`); comparison produces "
    "Boolean Series such as `df['Satisfaction'] == 'Satisfied'`; and the elementwise logical operator "
    "`|` combines two outlier conditions into one mask. The distinction between `|` and the Python "
    "keyword `or` matters: `|` compares every row of two pandas Series, whereas `or` would attempt to "
    "reduce the whole Series to a single truth value. Type conversion is used where a model or a "
    "display requires it, for example `astype(int)` when converting the Boolean satisfaction indicator "
    "to 0 and 1, and `round(value, 4)` when reporting metrics to four decimal places."
)
add_table(
    ["Construct", "Actual example from the notebook", "Role"],
    [
        ["Integer variable", "`RANDOM_STATE = 42`", "Reproducible splits, models, and sampling"],
        ["String variable", "`target_col = 'Satisfaction'`", "Names the target column"],
        ["Floating-point result", "`median_arrival_delay = df['Arrival Delay'].median()`", "Median used to impute 393 missing delays"],
        ["Tuple", "`df.shape` returns `(129880, 24)`", "Rows and columns of the loaded data"],
        ["Arithmetic", "`iqr = q3 - q1`; `lb = q1 - 1.5 * iqr`", "IQR lower and upper bounds"],
        ["Comparison and logical OR", "`(df[col] < lb) | (df[col] > ub)`", "Boolean mask of IQR outliers"],
        ["Type conversion", "`astype(int)`; `round(value, 4)`", "Binary target and metric formatting"],
    ],
    [1.45, 3.05, 1.85],
    "Core Python variables, data types, and operators used in the project.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)

h2("1.3 Lists, Tuples, and Dictionaries")
body(
    "Lists hold the ordered feature groups that drive preprocessing. `numeric_features` contains the "
    "four continuous columns (`Age`, `Flight Distance`, `Departure Delay`, and `Arrival Delay`); "
    "`categorical_features` contains the four nominal columns (`Gender`, `Customer Type`, "
    "`Type of Travel`, and `Class`); and `ordinal_rating_features` contains the 14 service-rating "
    "columns. Because these lists are passed to the `ColumnTransformer`, the same feature definitions "
    "are reused for scaling, one-hot encoding, clustering, and permutation importance. A change to a "
    "list therefore propagates through the pipeline instead of requiring edits in several places."
)
body(
    "Tuples appear where a function naturally returns more than one value. The quartiles are assigned "
    "with `q1, q3 = df[col].quantile(0.25), df[col].quantile(0.75)`, which packs two values and "
    "unpacks them in one statement. `df.shape` is a tuple of two integers, and the plotting loop uses "
    "`for ax, col in zip(axes, numeric_features)` to unpack each axis-and-column pair. Tuple unpacking "
    "keeps the code compact without hiding the order of the returned values."
)
body(
    "Dictionaries map a label to a value and are used for every summarised result. `outlier_summary = {}` "
    "is filled with one entry per numeric column, keyed by the column name; `inertias, sil_scores = {}, {}` "
    "stores K-Means diagnostics keyed by the number of clusters; `metrics` stores the named evaluation "
    "scores for one model; and `param_grid` stores the candidate values for each Random Forest "
    "hyperparameter. In each case the dictionary preserves the meaning of the value, so a result can be "
    "retrieved by name rather than by a fragile positional index. Python 3.7 and later preserve "
    "insertion order, which is why the notebook can print or plot these dictionaries directly."
)

h2("1.4 Functions, Loops, and Conditional Logic", page_break=True)
body(
    "Two user-defined functions remove repetition from the analysis. `age_group(age)` applies an "
    "`if`/`elif`/`else` chain: passengers younger than 25 are labelled `Youth`, passengers younger than "
    "60 are labelled `Adult`, and the remainder are labelled `Senior`. The function is applied to the "
    "`Age` column with `df['Age'].apply(age_group)`, producing the `Age Group` feature used in the "
    "mini-project analysis. `evaluate_pipeline(name, pipe, X_te, y_te)` centralises prediction and "
    "scoring: it computes accuracy, precision, recall, F1, and ROC-AUC, builds the confusion matrix, "
    "and returns the pair `(metrics, cm)`. The same function is reused for the Decision Tree, Naive "
    "Bayes, SVM, and Random Forest, which is why their results are directly comparable."
)
body(
    "Loops handle repeated operations over columns, axes, and cluster counts. `for i, col in "
    "enumerate(df.columns, 1)` prints a numbered column inventory; `for ax, col in zip(axes, "
    "numeric_features)` draws one box plot per continuous feature; and `for k in range(2, 9)` fits "
    "K-Means for every candidate cluster count from 2 to 8 and records inertia and silhouette score. "
    "Conditional logic is also used inside `evaluate_pipeline` to decide whether a model exposes "
    "`predict_proba`; if it does, ROC-AUC is computed, and if it does not, the ROC-AUC entry is left "
    "empty rather than failing the whole evaluation."
)
add_table(
    ["Control construct", "Actual use in the project", "Effect"],
    [
        ["Function with conditionals", "`age_group(age)` with thresholds 25 and 60", "Creates the Youth, Adult, and Senior groups"],
        ["Function with a return tuple", "`evaluate_pipeline(...)` returns `(metrics, cm)`", "Reuses one evaluation routine for four models"],
        ["Loop over columns", "`for i, col in enumerate(df.columns, 1)`", "Prints the numbered column inventory"],
        ["Loop over plot axes", "`for ax, col in zip(axes, numeric_features)`", "Draws four IQR box plots in one figure"],
        ["Loop over cluster counts", "`for k in range(2, 9)`", "Fits K-Means for k = 2 through 8"],
        ["Conditional capability check", "`if hasattr(pipe, 'predict_proba')`", "Computes ROC-AUC only when probabilities exist"],
    ],
    [1.55, 3.00, 1.80],
    "Functions, loops, and conditional statements used in the executed notebook.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)

h2("1.5 Reading and Writing CSV Data with pandas")
body(
    "The dataset is read with a single pandas call: `df = pd.read_csv(\"../dataset/"
    "airline_passenger_satisfaction.csv\")`. pandas infers the header, the column names, and the data "
    "types, and the notebook immediately prints `df.shape`, which returns `(129880, 24)`. This one "
    "operation is the bridge between Lab 1 and the rest of the project: after it, the passenger "
    "records are available as a labelled DataFrame rather than as raw text. The initial inspection "
    "uses `df.head()`, `df.tail()`, `df.info()`, `df.describe()`, `df.columns`, and "
    "`value_counts()` on the categorical columns. Those calls confirm the 24-column schema, the "
    "integer and string types, the single floating-point column `Arrival Delay`, and the 393 missing "
    "arrival-delay values that are imputed in Lab 2."
)
body(
    "The project reads CSV data but does not write a processed CSV file. Its outputs are figures, "
    "model metrics, and the fitted pipeline, and the pipeline is persisted with `joblib.dump` rather "
    "than exported as a table. The pandas counterpart for tabular export would be "
    "`df.to_csv(path, index=False)`, which writes the DataFrame without its row index; it is not "
    "called in the notebook because no downstream stage consumes a processed CSV. The distinction is "
    "worth stating explicitly: the reading half of the Lab 1 requirement is exercised directly, while "
    "the writing half is not needed by this project's output format. No CSV output is claimed where "
    "none was produced."
)

h2("1.6 Lab 1 Outcomes and Manual Coverage")
body(
    "The Lab 1 outcomes are met operationally rather than through a separate quiz or syntax "
    "worksheet. The project creates and uses variables of several types, applies arithmetic, "
    "comparison, and logical operators, stores ordered feature groups in lists, unpacks tuples, "
    "summarises results in dictionaries, defines reusable functions, iterates over columns and cluster "
    "counts, and reads the source data with pandas. The table below maps each topic in the lab manual "
    "to the evidence in the executed project."
)
add_table(
    ["Lab 1 topic", "Evidence in the executed project", "Report subsection"],
    [
        ["Python basics", "Import cell, configuration variables, and the CSV load", "1.1, 1.5"],
        ["Variables, data types, operators", "`RANDOM_STATE`, `target_col`, Boolean masks, IQR arithmetic", "1.2"],
        ["Lists", "`numeric_features`, `categorical_features`, `ordinal_rating_features`", "1.3"],
        ["Tuples", "Quantile unpacking, `df.shape`, `zip` in plotting loops", "1.3"],
        ["Dictionaries", "`outlier_summary`, `inertias`, `sil_scores`, `metrics`, `param_grid`", "1.3"],
        ["Functions and loops", "`age_group`, `evaluate_pipeline`, column and cluster loops", "1.4"],
        ["Reading CSV with pandas", "`pd.read_csv(...)` followed by shape and schema inspection", "1.5"],
        ["Writing CSV with pandas", "Not required by the project; the fitted model is saved with joblib", "1.5"],
    ],
    [1.60, 3.25, 1.50],
    "Lab 1 manual topics mapped to evidence in the executed project.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "The verification for Lab 1 is therefore the successful execution of the notebook's first cells: "
    "the libraries import, the CSV file loads, and the shape prints as 129,880 rows by 24 columns. No "
    "standalone Lab 1 figure or metric is reported because the notebook did not produce one, and none "
    "is fabricated here. The section nevertheless covers every topic listed in the manual and shows "
    "where each one is used in the applied work that follows."
)

# ============================================================================
# 2. LAB 2 — DATA PREPROCESSING AND EXPLORATION
# ============================================================================
h1("2. Lab 2 — Data Preprocessing and Exploration")
body(
    "The preprocessing stage prepares a modeling-ready dataset without allowing test-set information "
    "to influence training. It covers loading and inspection, descriptive statistics, missing-value "
    "treatment, duplicate and identifier handling, categorical encoding, feature scaling, and a "
    "stratified train/test split. Outlier screening is also performed in the notebook at this stage, "
    "but because the manual places outlier detection in Lab 11, the detailed discussion is deferred to "
    "Section 11.1."
)

h2("2.1 Loading and Initial Inspection")
body(
    "The dataset was loaded with `pandas.read_csv` from the project's `dataset/` folder. The shape "
    "printed as 129,880 rows by 24 columns. Inspecting the first and last five rows confirmed that the "
    "file contains the expected passenger, trip, service-rating, and target fields, and that the "
    "categorical columns were read as strings rather than accidentally coerced to numbers. The "
    "`DataFrame.info()` output showed 18 integer columns, one floating-point column (`Arrival Delay`), "
    "and five string columns, with a memory footprint of 29.6 MB."
)
body(
    "Only `Arrival Delay` had missing values: 129,487 non-null entries out of 129,880. All other "
    "columns were complete. This first inspection also confirmed plausible value ranges — for example, "
    "Age runs from 7 to 85 and Flight Distance from 31 to 4,983 miles — so there was no evidence of "
    "corrupted or mis-scaled fields that would require row-level repair. The inspection therefore "
    "established a clean starting point and narrowed preprocessing to a small number of deliberate "
    "decisions."
)

h2("2.2 Descriptive Statistics and Feature Roles")
body(
    "Descriptive statistics were computed for the four continuous numeric features. The results are "
    "summarized below. `Arrival Delay` is shown with its original non-null count because the table "
    "describes the data before imputation."
)
add_table(
    ["Statistic", "Age", "Flight Distance", "Departure Delay", "Arrival Delay"],
    [
        ["Count", "129,880", "129,880", "129,880", "129,487"],
        ["Mean", "39.43", "1,190.32", "14.71", "15.09"],
        ["Standard deviation", "15.12", "997.45", "38.07", "38.47"],
        ["Minimum", "7", "31", "0", "0"],
        ["25th percentile", "27", "414", "0", "0"],
        ["Median", "40", "844", "0", "0"],
        ["75th percentile", "51", "1,744", "12", "13"],
        ["Maximum", "85", "4,983", "1,592", "1,584"],
    ],
    [1.45, 1.15, 1.35, 1.15, 1.17],
    "Descriptive statistics for the four continuous numeric features.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 4,
    first_col_bold=True,
)
body(
    "Age is roughly symmetric: the mean (39.43) and median (40) are almost identical, and the range "
    "from 7 to 85 describes a broad, adult-dominated passenger base. Flight Distance is right-skewed. "
    "Its mean (1,190.32 miles) is well above its median (844 miles), and the maximum of 4,983 miles "
    "pulls the mean upward. The two delay variables are even more strongly skewed: their medians are "
    "zero, their means are about 15 minutes, and their standard deviations (38.07 and 38.47) are more "
    "than twice the means. This is the signature of a distribution in which most flights are on time "
    "but a small number are very late. It directly justifies median rather than mean imputation for "
    "`Arrival Delay` and argues against treating delay outliers as data errors."
)
body(
    "Categorical frequencies were also examined. The table below reports the counts and shares for the "
    "four nominal predictors and the target."
)
add_table(
    ["Feature", "Category", "Count", "Share"],
    [
        ["Gender", "Female", "65,899", "50.74%"],
        ["Gender", "Male", "63,981", "49.26%"],
        ["Customer Type", "Returning", "106,100", "81.69%"],
        ["Customer Type", "First-time", "23,780", "18.31%"],
        ["Type of Travel", "Business", "89,693", "69.06%"],
        ["Type of Travel", "Personal", "40,187", "30.94%"],
        ["Class", "Business", "62,160", "47.86%"],
        ["Class", "Economy", "58,309", "44.89%"],
        ["Class", "Economy Plus", "9,411", "7.25%"],
        ["Satisfaction", "Neutral or Dissatisfied", "73,452", "56.55%"],
        ["Satisfaction", "Satisfied", "56,428", "43.45%"],
    ],
    [1.55, 2.45, 1.10, 1.05],
    "Frequency distribution of the nominal predictors and the target.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "Gender is nearly balanced, so it is unlikely to introduce a strong sampling bias on its own. The "
    "other three predictors are imbalanced. Returning customers make up 81.69% of the sample, business "
    "travelers 69.06%, and Business-class passengers 47.86% while Economy Plus accounts for only 7.25%. "
    "This composition is important for interpretation: the model is trained predominantly on "
    "experienced business travelers, and the small Economy Plus group will produce less stable "
    "group-level estimates. The target imbalance is milder (56.55% versus 43.45%) but still relevant "
    "for metric choice, because a trivial majority-class classifier would already achieve 56.55% "
    "accuracy."
)

h2("2.3 Missing Value Treatment")
body(
    "`Arrival Delay` was the only column with missing data: 393 of 129,880 rows, or 0.30%. The missing "
    "values were filled with the median of the column, which is 0.0 minutes. Median imputation was "
    "chosen because the delay distribution is strongly right-skewed; the mean would be pulled upward "
    "by a relatively small number of very long delays and would assign an unrealistically high arrival "
    "delay to passengers whose delay was simply unrecorded. With only 0.30% of rows affected, the choice "
    "of imputation method cannot materially change the distribution, but the median is the more "
    "defensible central estimate for skewed data."
)
body(
    "The alternative of dropping the 393 incomplete rows was rejected because it would discard valid "
    "information about every other feature for those passengers, and because the missingness is small "
    "enough that a simple, transparent imputation is preferable to a more complex model-based method. "
    "After imputation, the dataset contained zero missing values. One methodological detail is worth "
    "recording: the median was computed on the full dataset before the train/test split. Because it is "
    "a single robust statistic and affects only 0.30% of one column, the practical leakage risk is "
    "negligible; a stricter implementation would place the imputer inside the modeling pipeline. The "
    "scaling and encoding steps, which have a much larger leakage potential, were correctly fitted only "
    "on the training data."
)

h2("2.4 Duplicate and Identifier Handling")
body(
    "A duplicate-row check returned zero duplicates, so no rows were removed. The `ID` column was then "
    "dropped. It contains 129,880 unique values for 129,880 rows and therefore identifies the row rather "
    "than describing the passenger. Keeping it would allow a model to memorize individual records and "
    "would add a meaningless feature with no generalization value. After the drop, 23 columns remained: "
    "22 predictors and the target."
)

h2("2.5 Categorical Encoding")
body(
    "The four nominal predictors — `Gender`, `Customer Type`, `Type of Travel`, and `Class` — were "
    "one-hot encoded with `OneHotEncoder(drop='first', handle_unknown='ignore')`. One-hot encoding is "
    "appropriate because these variables have no inherent order; assigning integer codes such as "
    "Female = 0 and Male = 1 would imply a ranking that does not exist. Dropping the first category of "
    "each variable avoids perfect multicollinearity between the dummy columns, which is especially "
    "relevant for models that estimate coefficients or distances. The `handle_unknown='ignore'` setting "
    "makes the transformer robust if a future passenger record contains a category that was not present "
    "during training: the unknown category is encoded as all zeros rather than causing an error."
)
body(
    "The 14 service-rating columns were not re-encoded. They are already numeric and genuinely ordinal, "
    "so they were used as-is and subsequently standardised. This preserves their ordering while putting "
    "them on a comparable scale. The target `Satisfaction` was label-encoded to 0 and 1, with the "
    "classes ordered as `Neutral or Dissatisfied` = 0 and `Satisfied` = 1. Label encoding is acceptable "
    "here because the target is binary and is used only as the model's output; it does not create a "
    "false ordinal relationship among the input features."
)
body(
    "One caveat deserves emphasis. The data dictionary defines 0 on the service-rating scale as “not "
    "applicable,” not as the lowest possible satisfaction rating. Treating the ratings as a numeric "
    "0–5 scale therefore merges “not applicable” with “very dissatisfied.” The notebook does this "
    "consistently across all 14 ratings, and the resulting models perform well, but the interpretation "
    "of a rating of 0 should remain cautious. A stricter treatment would add a separate indicator for "
    "“not applicable” or impute those entries, which is left as a possible extension rather than "
    "retrofitted into the reported results."
)

h2("2.6 Feature Scaling and Stratified Train/Test Split")
body(
    "The data was split into 80% training and 20% test sets using `train_test_split` with "
    "`stratify=y_enc` and `random_state=42`. Stratification preserves the target ratio in both sets, "
    "which matters because the classes are mildly imbalanced: without it, a random split could "
    "accidentally over- or under-represent the Satisfied class in the test set and produce misleading "
    "metrics. The resulting split is shown below."
)
add_table(
    ["Subset", "Rows", "Neutral or Dissatisfied", "Satisfied", "Class balance"],
    [
        ["Training", "103,904", "58,762 (56.55%)", "45,142 (43.45%)", "56.55% / 43.45%"],
        ["Test", "25,976", "14,690 (56.55%)", "11,286 (43.45%)", "56.55% / 43.45%"],
    ],
    [1.20, 0.90, 1.60, 1.40, 1.17],
    "Stratified 80/20 train/test split and class balance.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "Preprocessing was assembled as a `ColumnTransformer` inside a scikit-learn `Pipeline`. A "
    "`StandardScaler` was applied to the 18 numeric and ordinal columns (the four continuous features "
    "plus the 14 service ratings), and the `OneHotEncoder` was applied to the four nominal columns. "
    "Because the pipeline is fitted with `pipeline.fit(X_train, ...)`, the scaler's means and standard "
    "deviations and the encoder's category list are estimated only from the training rows. The test set "
    "is transformed with those training statistics and never contributes to them. This is the central "
    "leakage control in the workflow: it prevents the model from indirectly seeing the distribution of "
    "the test data during training."
)
body(
    "Scaling is essential for the RBF-kernel SVM, which measures distances between observations and is "
    "sensitive to the units of its inputs. It is harmless for the tree-based models, which split on "
    "thresholds and do not depend on monotone rescaling, and it is not strictly required by Gaussian "
    "Naive Bayes, which estimates a separate mean and variance for each feature. Using one shared "
    "preprocessor keeps the comparison fair and ensures that every model sees exactly the same "
    "information in exactly the same representation. The one-hot step expands the four nominal columns "
    "into five dummy columns, so the transformed model matrix contains 23 inputs."
)

h2("2.7 Preprocessing Outcome")
body(
    "After preprocessing, the training matrix has 103,904 rows and 23 transformed features, and the "
    "test matrix has 25,976 rows and the same 23 features. There are no missing values, no duplicate "
    "rows, no identifier column, and no raw categorical strings. The target is encoded as a binary "
    "integer. Delay outliers were deliberately retained and are examined in Section 11.1, where the "
    "IQR results are reported. The data is therefore ready for exploratory analysis and modeling, with "
    "the important caveat that the service-rating scale and the pre-split median imputation have been "
    "documented rather than hidden."
)

# ============================================================================
# 3. LAB 3 — EXPLORATORY DATA ANALYSIS
# ============================================================================
h1("3. Lab 3 — Exploratory Data Analysis")
body(
    "The exploratory analysis translates the raw dataset into service-level insight. Each visualization "
    "was chosen to answer a question relevant to service optimization: How balanced is the target? Who "
    "are the passengers? Which groups are less satisfied? Which individual services correlate most "
    "strongly with overall satisfaction? Which services receive the lowest ratings? And do the "
    "relationships look strong enough to support a multivariate model?"
)

h2("3.1 Target Distribution")
body(
    "The target distribution is shown below. Neutral or Dissatisfied passengers outnumber Satisfied "
    "passengers by 73,452 to 56,428, a 56.55% to 43.45% split. The imbalance is mild enough that "
    "resampling is not necessary, but it is large enough that accuracy alone would be a misleading "
    "metric: a model that always predicts Neutral or Dissatisfied would already score 56.55%. For this "
    "reason, the classification evaluation in Sections 6–8 and 12 relies on precision, recall, "
    "F1-score, and ROC-AUC in addition to accuracy."
)
add_figure("fig_32_2.png", "Distribution of the target variable. Neutral or Dissatisfied is the majority class at 56.55%.", width=4.3)

h2("3.2 Passenger Demographics and Travel Profiles")
body(
    "The age distribution is broad and adult-dominated. Ages range from 7 to 85, with a mean of 39.43 "
    "and a median of 40. Most passengers are between the late twenties and early fifties, and the "
    "distribution tapers at both ends rather than forming a narrow peak. Gender is almost balanced: "
    "65,899 female passengers (50.74%) and 63,981 male passengers (49.26%)."
)
add_figure("fig_34_3.png", "Passenger age distribution (left) and gender distribution (right).", width=6.25)
body(
    "The travel profile is more strongly skewed. Returning customers account for 106,100 passengers "
    "(81.69%), while first-time customers account for 23,780 (18.31%). Business travel accounts for "
    "89,693 passengers (69.06%) versus 40,187 personal travelers (30.94%). Cabin class is split between "
    "Business (62,160; 47.86%) and Economy (58,309; 44.89%), with Economy Plus a distant third at 9,411 "
    "passengers (7.25%)."
)
add_figure("fig_36_4.png", "Customer type distribution (left) and type-of-travel distribution (right).", width=6.25)
add_figure("fig_38_5.png", "Cabin class distribution. Economy Plus is a small minority category.", width=3.9)
body(
    "These proportions describe a dataset dominated by returning business travelers flying in Business "
    "or Economy class. That composition is useful for an airline that wants to understand its core "
    "customer base, but it also means that conclusions about first-time, personal, and Economy Plus "
    "passengers are based on smaller groups. The group-level satisfaction rates in the next subsection "
    "should be read with that caveat in mind."
)

h2("3.3 Satisfaction Gaps Across Passenger Groups")
body(
    "Satisfaction rates were computed for each category of the nominal predictors. The differences are "
    "substantial and, in several cases, much larger than the differences between individual service "
    "ratings."
)
add_figure("fig_40_6.png", "Satisfaction rate by gender, customer type, and type of travel.", width=6.25)
body(
    "Gender shows almost no gap: 42.90% of female passengers are satisfied versus 44.01% of male "
    "passengers, a difference of about one percentage point. Customer type shows a large gap: 47.81% of "
    "returning customers are satisfied, compared with only 23.97% of first-time customers. Type of "
    "travel shows the largest gap in the dataset: 58.37% of business travelers are satisfied, compared "
    "with 10.13% of personal travelers. Personal travelers are therefore overwhelmingly dissatisfied, "
    "not merely slightly less satisfied."
)
add_figure("fig_42_7.png", "Satisfaction rate by cabin class. Business class is far ahead of Economy and Economy Plus.", width=3.9)
body(
    "Cabin class follows the same pattern. Business-class passengers have a 69.44% satisfaction rate, "
    "Economy Plus passengers 24.64%, and Economy passengers 18.77%. The Business-versus-Economy gap is "
    "roughly 50 percentage points. These are strong bivariate associations, but they should not be read "
    "as causal effects. Type of Travel and Class are correlated: business travelers are more likely to "
    "fly Business class, and the two variables may partly proxy for each other. Disentangling them "
    "requires the multivariate model and the feature-importance analysis that follow."
)

h2("3.4 Correlation Between Service Ratings and Satisfaction")
body(
    "A correlation matrix was computed for the 14 service ratings, the four continuous features, and a "
    "binary version of the target (Satisfied = 1, Neutral or Dissatisfied = 0). The heatmap below shows "
    "both the target correlations and the blocks of correlation among the service ratings themselves."
)
add_figure("fig_44_8.png", "Correlation heatmap for service ratings, numeric features, and satisfaction.", width=5.25)
body(
    "The target correlations are reported in the table below, sorted from strongest positive to "
    "strongest negative. Pearson correlation is used here as a descriptive screen; it captures linear "
    "association only and does not prove that a service causes satisfaction."
)
add_table(
    ["Feature", "Correlation with satisfaction", "Direction"],
    [
        ["Online Boarding", "0.502", "Positive, strongest"],
        ["In-flight Entertainment", "0.398", "Positive"],
        ["Seat Comfort", "0.349", "Positive"],
        ["On-board Service", "0.322", "Positive"],
        ["Leg Room Service", "0.312", "Positive"],
        ["Cleanliness", "0.307", "Positive"],
        ["Flight Distance", "0.298", "Positive"],
        ["In-flight Wifi Service", "0.283", "Positive"],
        ["Baggage Handling", "0.249", "Positive"],
        ["In-flight Service", "0.245", "Positive"],
        ["Check-in Service", "0.237", "Positive"],
        ["Food and Drink", "0.211", "Positive"],
        ["Ease of Online Booking", "0.169", "Positive, weak"],
        ["Age", "0.134", "Positive, weak"],
        ["Gate Location", "−0.003", "Negligible"],
        ["Departure Delay", "−0.051", "Negative, weak"],
        ["Departure and Arrival Time Convenience", "−0.054", "Negative, weak"],
        ["Arrival Delay", "−0.058", "Negative, weak"],
    ],
    [2.75, 1.70, 1.75],
    "Pearson correlation of each feature with the binary satisfaction target.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "Online Boarding is the strongest single correlate (r = 0.502), followed by In-flight Entertainment "
    "(0.398), Seat Comfort (0.349), On-board Service (0.322), Leg Room Service (0.312), and Cleanliness "
    "(0.307). These are all service-experience variables, and they form a coherent cluster: passengers "
    "who rate one aspect of the cabin experience highly tend to rate the others highly as well. The "
    "positive correlation for Flight Distance (0.298) is likely indirect; longer flights are more often "
    "Business-class or international journeys, so distance may be picking up part of the class effect."
)
body(
    "The delay variables behave differently. Departure Delay (−0.051), Arrival Delay (−0.058), and "
    "Departure and Arrival Time Convenience (−0.054) are all close to zero, and Gate Location is "
    "essentially uncorrelated with satisfaction (−0.003). This does not mean that delays are irrelevant "
    "to passengers; it means that their linear association with the overall label is weak in this "
    "dataset, possibly because delays are concentrated in a minority of flights or because their effect "
    "is mediated by other service ratings. The multivariate model is better suited to detecting "
    "nonlinear or interaction effects than a single correlation coefficient."
)

h2("3.5 Service Rating Levels")
body(
    "Average ratings were computed for all 14 services and sorted from lowest to highest. This answers a "
    "different question from the correlation analysis: not “which service predicts satisfaction?” but "
    "“where is the airline already performing poorly?”"
)
add_figure("fig_46_9.png", "Average rating for each service on the 0–5 scale, sorted from lowest to highest.", width=5.05)
body(
    "The four lowest-rated services are In-flight Wifi Service (2.73), Ease of Online Booking (2.76), "
    "Gate Location (2.98), and Departure and Arrival Time Convenience (3.06). The highest-rated services "
    "are In-flight Service (3.64) and Baggage Handling (3.63), followed by Seat Comfort (3.44), On-board "
    "Service (3.38), In-flight Entertainment (3.36), and Leg Room Service (3.35). The remaining services "
    "cluster in the middle: Check-in Service (3.31), Cleanliness (3.29), Online Boarding (3.25), and "
    "Food and Drink (3.20). The overall mean across the 14 ratings is 3.24."
)
body(
    "The combination of rating level and correlation strength is more informative than either measure "
    "alone. In-flight Wifi is both the lowest-rated service and a meaningful positive correlate "
    "(r = 0.283), which makes it a strong candidate for improvement. Ease of Online Booking is also "
    "low-rated, but its correlation is weaker (0.169). Gate Location is low-rated yet essentially "
    "uncorrelated with satisfaction, suggesting that passengers may dislike the gate but not hold it "
    "against the airline enough to affect overall satisfaction. This distinction becomes the basis of "
    "the priority synthesis in Section 12.5."
)

h2("3.6 Distribution and Relationship Checks")
body(
    "Two additional plots examined the shape of the relationship rather than a single summary number. "
    "The first compares the Online Boarding rating across the two satisfaction classes. The satisfied "
    "group is concentrated at higher ratings, while the dissatisfied group is concentrated at lower "
    "ratings, which is consistent with Online Boarding having the strongest correlation in the dataset. "
    "The distributions overlap, however, so Online Boarding is informative rather than deterministic."
)
add_figure("fig_48_10.png", "Online Boarding rating by satisfaction class.", width=4.25)
body(
    "The second plot shows Flight Distance against Age for a random sample of 3,000 passengers, colored "
    "by satisfaction. The colors overlap heavily across the plane, with no clean geometric separation. "
    "This visual impression matches the modest correlation for Flight Distance and the low permutation "
    "importance reported later: neither variable is a strong standalone discriminator. It also "
    "illustrates why a multivariate model is necessary — the useful signal is distributed across many "
    "features rather than concentrated in a two-dimensional projection."
)
add_figure("fig_48_11.png", "Flight Distance versus Age, colored by satisfaction (3,000-row sample for readability).", width=4.25)

h2("3.7 EDA Synthesis", page_break=True)
body(
    "The exploratory analysis supports five conclusions that shape the rest of the project. First, the "
    "target is mildly imbalanced, so F1-score and ROC-AUC should accompany accuracy. Second, service "
    "experience dominates the raw correlations: Online Boarding, In-flight Entertainment, Seat Comfort, "
    "On-board Service, and Leg Room Service are the strongest individual correlates. Third, In-flight "
    "Wifi is the clearest service-quality weakness because it is both the lowest-rated service and a "
    "meaningful correlate. Fourth, personal travelers, first-time customers, and Economy passengers are "
    "much less satisfied than their counterparts, which points to structural segment differences rather "
    "than a single failing service. Fifth, the delay variables have weak linear correlations, so their "
    "role, if any, is likely nonlinear or indirect."
)
callout(
    "Key finding",
    "The strongest raw signals are service-experience variables, not delay variables. Online Boarding "
    "is the single strongest correlate, while In-flight Wifi combines the lowest average rating with a "
    "meaningful positive correlation — a combination that makes it a high-leverage improvement target.",
)
body(
    "The ratings are also intercorrelated, which has two consequences. It supports the concern that "
    "Gaussian Naive Bayes may struggle because its conditional-independence assumption is violated. It "
    "also means that individual correlation coefficients understate the joint importance of the service "
    "ratings: a model that can combine them may extract more signal than any single coefficient "
    "suggests. The classification and feature-importance sections test that possibility directly."
)

# ============================================================================
# 6. LAB 6 — DECISION TREE
# ============================================================================
h1("6. Lab 6 — Decision Tree Classification")
body(
    "The first classification model is a decision tree. It is included because it is interpretable, "
    "handles nonlinear threshold effects naturally, and provides a transparent baseline against which "
    "the ensemble and kernel models can be judged. The same training and test sets, and the same "
    "preprocessing pipeline, are used here as in every other classification section."
)

h2("6.1 Objective and Model Rationale")
body(
    "A decision tree recursively partitions the feature space by choosing splits that reduce impurity. "
    "At each node, the algorithm evaluates candidate thresholds and selects the split that produces the "
    "largest improvement in the chosen criterion. The notebook uses scikit-learn's default Gini "
    "impurity, which measures the probability that a randomly chosen observation would be misclassified "
    "if it were labeled according to the class distribution in the node. Entropy is the information- "
    "theoretic alternative; it was not tuned in this project, but it would produce a similar family of "
    "axis-aligned splits."
)
body(
    "The tree is a natural fit for this dataset because the service ratings are ordinal and their "
    "effects are unlikely to be strictly linear. A split such as “Online Boarding ≤ 2” can isolate a "
    "group with a much lower satisfaction rate without assuming that the difference between 2 and 3 is "
    "the same as the difference between 4 and 5. Trees also handle interactions automatically: a split "
    "on Type of Travel can be followed by a different split on Class or Seat Comfort in each branch."
)

h2("6.2 Training Configuration")
body(
    "The classifier was wrapped in a pipeline with the shared preprocessor and trained on the full "
    "103,904-row training set. The tree used `max_depth=12`, `random_state=42`, and scikit-learn's "
    "default splitter and minimum-samples settings. The depth limit is a regularization choice: an "
    "unrestricted tree can grow until it memorizes the training data, producing low training error but "
    "poor test performance. A maximum depth of 12 constrains the model to a more generalizable set of "
    "rules while still allowing a substantial number of interactions."
)
body(
    "The fitted tree was not exported as a graphical diagram in the notebook. The depth limit is the "
    "primary interpretability control, and feature-level interpretation is supplied by the permutation "
    "importance analysis in Section 11.3, which is more informative at this feature count than a "
    "full-depth tree plot. The model was evaluated on the held-out 25,976-row test set using accuracy, "
    "precision, recall, F1-score, and ROC-AUC, with the positive class defined as Satisfied."
)

h2("6.3 Results")
body(
    "The decision tree performed strongly on the unseen test set. Its results are summarized below."
)
add_table(
    ["Metric", "Test-set value"],
    [
        ["Accuracy", "0.9501"],
        ["Precision", "0.9581"],
        ["Recall", "0.9257"],
        ["F1-score", "0.9416"],
        ["ROC-AUC", "0.9829"],
    ],
    [3.10, 3.10],
    "Decision Tree performance on the held-out test set.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "Accuracy is 95.01%, meaning the tree classifies roughly 19 out of every 20 passengers correctly. "
    "Precision (95.81%) is higher than recall (92.57%): when the tree predicts Satisfied, it is usually "
    "right, but it misses about 7.4% of the passengers who are actually satisfied. The F1-score of "
    "0.9416 balances these two errors, and the ROC-AUC of 0.9829 indicates excellent separation across "
    "all possible decision thresholds. The confusion matrix for this model appears in the comparative "
    "figure in Section 12.3, where it can be read alongside the other classifiers."
)

h2("6.4 Interpretation")
body(
    "The decision tree is a strong baseline and clearly beats the 56.55% majority-class rate. Its "
    "performance is consistent with the EDA: the service ratings contain enough signal to separate most "
    "satisfied from dissatisfied passengers, and the tree can exploit threshold effects that a linear "
    "model would miss. The gap between precision and recall also has a practical interpretation. If the "
    "airline uses the model to identify passengers who are likely to be satisfied, a high-precision "
    "rule is useful for targeted positive outreach; if it uses the model to flag passengers at risk, a "
    "higher-recall rule would be more appropriate."
)
body(
    "The main limitation is variance. A single tree is sensitive to the particular training sample: "
    "small changes in the data can change the early splits and therefore the entire downstream "
    "structure. This is exactly the weakness that the Random Forest ensemble addresses in Section 11.2. "
    "The tree's result is therefore best read as a transparent, competitive benchmark rather than as "
    "the final model."
)

# ============================================================================
# 7. LAB 7 — NAIVE BAYES
# ============================================================================
h1("7. Lab 7 — Naive Bayes Classification")
body(
    "The second classifier is Gaussian Naive Bayes. It is included as a fast probabilistic baseline and "
    "as a test of how much performance is possible under a much stronger modeling assumption than the "
    "one made by the tree or the ensemble."
)

h2("7.1 Objective and Model Rationale")
body(
    "Naive Bayes applies Bayes' theorem to classification. For a passenger with feature vector x, the "
    "model estimates the posterior probability of each class as the class prior multiplied by the "
    "product of the per-feature conditional probabilities, and then assigns the class with the larger "
    "posterior. The “naive” part is the assumption that features are conditionally independent given "
    "the class. Gaussian Naive Bayes further assumes that, within each class, each continuous feature "
    "follows a normal distribution, so it estimates a mean and variance per feature per class."
)
body(
    "Neither assumption is a perfect description of this dataset. The service ratings are bounded "
    "ordinal variables, not unbounded Gaussian variables, and they are substantially correlated with "
    "one another. A passenger who rates Seat Comfort low is also likely to rate Leg Room Service low, "
    "so the conditional-independence assumption is violated. The model may therefore underweight or "
    "double-count information. It is still worth fitting because it is extremely fast, provides "
    "probability estimates, and establishes how much accuracy is available from a simple generative "
    "model."
)
body(
    "Multinomial Naive Bayes was not used. It is designed for count data such as word frequencies, "
    "whereas the inputs here are ordinal ratings and continuous trip variables. Gaussian Naive Bayes is "
    "the appropriate variant for this feature representation."
)

h2("7.2 Training Configuration")
body(
    "GaussianNB was placed in the same pipeline as the other classifiers and trained on the full "
    "training set with scikit-learn's default smoothing and variance settings. No feature selection or "
    "special transformation was applied beyond the shared standardization and one-hot encoding. The "
    "model was evaluated on the same 25,976-row test set with the same five metrics and the same "
    "positive class."
)

h2("7.3 Results")
body(
    "The Naive Bayes results are shown below. They are noticeably weaker than the decision tree's, "
    "although still well above the majority-class baseline."
)
add_table(
    ["Metric", "Test-set value"],
    [
        ["Accuracy", "0.8626"],
        ["Precision", "0.8592"],
        ["Recall", "0.8178"],
        ["F1-score", "0.8380"],
        ["ROC-AUC", "0.9228"],
    ],
    [3.10, 3.10],
    "Gaussian Naive Bayes performance on the held-out test set.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "Accuracy is 86.26%, about 8.75 percentage points below the decision tree. Precision is 85.92% and "
    "recall 81.78%, so the model misses roughly 18% of satisfied passengers and produces a moderate "
    "number of false positives. The F1-score is 0.8380 and the ROC-AUC is 0.9228. The ROC-AUC is still "
    "high in absolute terms, which means the model ranks passengers reasonably well even when its "
    "thresholded labels are less accurate."
)

h2("7.4 Interpretation")
body(
    "The performance gap is consistent with the violated independence assumption. The 14 service "
    "ratings carry overlapping information, and Gaussian Naive Bayes treats each one as independent "
    "evidence. When several ratings move together, the model can become overconfident in one direction "
    "or effectively count the same signal multiple times. The Gaussian likelihood is also an "
    "approximation for bounded 0–5 ratings, especially near the boundaries, where a normal distribution "
    "assigns probability mass outside the valid range."
)
body(
    "Despite these limitations, Naive Bayes is not a failure. It recovers most of the signal with a "
    "very simple model and no hyperparameter tuning, and its probability outputs are informative enough "
    "to rank passengers. Its role in the project is therefore as a lower bound on model complexity: any "
    "model that cannot comfortably beat 86.26% accuracy is not adding much beyond a fast probabilistic "
    "baseline. The decision tree, SVM, and Random Forest all clear that bar."
)

# ============================================================================
# 8. LAB 8 — SUPPORT VECTOR MACHINE
# ============================================================================
h1("8. Lab 8 — Support Vector Machine Classification")
body(
    "The third classifier is a Support Vector Machine with a radial basis function kernel. It is "
    "included because it can represent smooth nonlinear decision boundaries that neither a single tree "
    "nor a Naive Bayes model expresses in the same way, and because it tests whether a margin-based "
    "method can improve on the tree's axis-aligned splits."
)

h2("8.1 Objective and Model Rationale")
body(
    "A support vector machine finds the decision boundary that maximizes the margin between classes. "
    "With a nonlinear kernel, the model implicitly maps the inputs into a higher-dimensional space "
    "where a linear separator corresponds to a nonlinear boundary in the original space. The RBF "
    "kernel measures similarity as a decreasing function of squared distance between observations, "
    "which allows the classifier to capture local, curved structure. The notebook used "
    "`SVC(kernel='rbf', probability=True, random_state=42)` with scikit-learn's default `C=1.0` and "
    "`gamma='scale'`."
)
body(
    "The SVM is a plausible fit for this problem because the relationship between service ratings and "
    "satisfaction may involve smooth interactions rather than sharp thresholds. Its main drawback is "
    "computational: kernel SVM training scales roughly quadratically to cubically with the number of "
    "rows, and the model does not produce a simple set of interpretable rules."
)

h2("8.2 Computational Adaptation")
body(
    "Training the RBF SVM on the full 103,904-row training set was not feasible on the available "
    "single-CPU environment; the training time grows steeply with row count, and the kernel matrix "
    "would be very large. The model was therefore trained on a stratified random subsample of 8,000 "
    "training rows, preserving the class ratio, and evaluated on the full 25,976-row test set. Because "
    "the test set is unchanged, the reported metrics are measured on the same unseen passengers as the "
    "other models."
)
body(
    "The entire pipeline, including the scaler and encoder, was fitted on the 8,000-row subsample. That "
    "keeps the experiment leakage-free, but it means the SVM's preprocessing statistics are estimated "
    "from fewer rows than the other models'. With 8,000 observations and only 23 transformed features, "
    "the estimates should be stable. The notebook fixes the RBF kernel and does not perform a search "
    "over kernels, `C`, or `gamma`; the result is therefore a reasonable default SVM configuration, not "
    "an exhaustively tuned one."
)

h2("8.3 Results")
body(
    "The SVM results are summarized below. The model is competitive but does not overtake the decision "
    "tree or the Random Forest."
)
add_table(
    ["Metric", "Test-set value"],
    [
        ["Accuracy", "0.9379"],
        ["Precision", "0.9365"],
        ["Recall", "0.9193"],
        ["F1-score", "0.9278"],
        ["ROC-AUC", "0.9811"],
    ],
    [3.10, 3.10],
    "RBF-kernel SVM performance on the held-out test set (trained on an 8,000-row stratified subsample).",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "Accuracy is 93.79%, precision 93.65%, recall 91.93%, F1-score 0.9278, and ROC-AUC 0.9811. The "
    "precision-recall balance is similar to the decision tree's, but both are lower. The ROC-AUC is "
    "almost identical to the tree's 0.9829, which suggests that the SVM ranks passengers about as well "
    "as the tree even though its thresholded labels are slightly less accurate."
)

h2("8.4 Interpretation")
body(
    "The SVM lands in third place among the four untuned classifiers, ahead of Naive Bayes and behind "
    "the decision tree and Random Forest. That ordering is plausible. The SVM's smooth RBF boundary is "
    "a good fit for continuous geometry, but the strongest signal in this dataset appears to be "
    "threshold-like: a rating of 2 versus 3 on Online Boarding carries a large change in satisfaction "
    "probability, and an axis-aligned tree can isolate that change with a single split. The SVM must "
    "approximate the same effect with a smooth surface."
)
body(
    "The computational adaptation is also part of the interpretation. Training on 8,000 rows rather "
    "than 103,904 necessarily leaves some information unused, so the SVM's result should be read as a "
    "lower bound on what a fully trained kernel SVM could achieve. Even with that caveat, the model adds "
    "useful evidence: a nonlinear margin-based classifier can reach 93.79% accuracy on this problem, "
    "which confirms that the signal is not an artifact of tree-specific splitting. The cost-performance "
    "tradeoff nevertheless favors the tree-based models for this dataset and deployment context."
)

# ============================================================================
# 9. LAB 9 — K-MEANS
# ============================================================================
h1("9. Lab 9 — K-Means Clustering")
body(
    "The classification models answer the question “will this passenger be satisfied?” The clustering "
    "analysis answers a different and complementary question: “what natural groups of passengers exist "
    "when they are described only by their service experience?” The goal is segmentation for service "
    "strategy, not prediction of the target."
)

h2("9.1 Objective and Feature Space")
body(
    "K-Means partitions observations into k groups by minimizing the within-cluster sum of squared "
    "distances to the cluster centroids. The algorithm alternates between assigning each passenger to "
    "the nearest centroid and recomputing the centroids, and it is sensitive to the scale of the "
    "features because those distances are Euclidean. For that reason, the 14 service ratings were "
    "standardised with `StandardScaler` before clustering."
)
body(
    "Only the 14 service ratings were used as clustering features. Including demographics or the target "
    "would answer a different question, and including the target would create a circular analysis in "
    "which the clusters merely reproduce the label. Restricting the feature space to service experience "
    "means that the resulting segments describe how passengers rated the airline, not who they are. "
    "That is the appropriate basis for a service improvement strategy."
)
body(
    "The clustering matrix therefore has 129,880 rows and 14 standardised columns. K-Means was run for "
    "k = 2 through 8 with `n_init=10` and `random_state=42`. Inertia (the within-cluster sum of squares) "
    "was recorded for the elbow curve, and the silhouette score was computed on a random sample of "
    "15,000 passengers to keep the repeated scoring computationally manageable."
)

h2("9.2 Choosing k with the Elbow Method and Silhouette Score")
body(
    "The elbow method plots inertia against k and looks for the point where additional clusters stop "
    "producing a large reduction in within-cluster variance. The silhouette score measures, for each "
    "observation, how much closer it is to its own cluster than to the nearest other cluster; it ranges "
    "from −1 to 1, with higher values indicating better-separated clusters. Using both is more robust "
    "than relying on the elbow alone, especially for survey-style data where the inertia curve is often "
    "smooth."
)
add_figure("fig_76_15.png", "Elbow curve (left) and silhouette score (right) for k = 2 to 8.", width=6.2)
body(
    "The silhouette scores are reported below. The highest value is at k = 2 (0.1821), after which the "
    "score drops to 0.1458 at k = 3 and then fluctuates slightly before declining to 0.1283 at k = 8."
)
add_table(
    ["Number of clusters (k)", "Silhouette score", "Comment"],
    [
        ["2", "0.1821", "Highest separation; selected"],
        ["3", "0.1458", "Sharp drop from k = 2"],
        ["4", "0.1473", "Small uptick, still below k = 2"],
        ["5", "0.1396", "Declining"],
        ["6", "0.1346", "Declining"],
        ["7", "0.1295", "Declining"],
        ["8", "0.1283", "Lowest of the tested values"],
    ],
    [1.55, 1.25, 3.40],
    "Silhouette scores for K-Means solutions with k = 2 to 8.",
    aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
)
body(
    "The elbow curve declines smoothly without a sharp corner. That is common for ordinal survey data: "
    "passengers do not form crisp, tightly separated clusters, and each additional cluster produces a "
    "gradual reduction in inertia rather than a sudden improvement. The silhouette score provides a "
    "clearer decision criterion, and it favors k = 2. The absolute silhouette value of 0.1821 is modest "
    "— values above 0.5 would indicate strong structure — so the result should be described as the best "
    "available two-group separation rather than as proof that there are exactly two true clusters. "
    "Given the business need for an interpretable segmentation, k = 2 is the defensible choice."
)

h2("9.3 Cluster Profiles")
body(
    "The final K-Means model with k = 2 assigned 72,159 passengers (55.56%) to Cluster 0 and 57,721 "
    "passengers (44.44%) to Cluster 1. The average rating for each service within each cluster is shown "
    "below and visualized in the accompanying figure."
)
add_table(
    ["Service", "Cluster 0 mean", "Cluster 1 mean", "Difference"],
    [
        ["Departure and Arrival Time Convenience", "3.17", "2.92", "0.25"],
        ["Ease of Online Booking", "2.96", "2.50", "0.46"],
        ["Check-in Service", "3.62", "2.91", "0.71"],
        ["Online Boarding", "3.77", "2.61", "1.16"],
        ["Gate Location", "3.01", "2.94", "0.07"],
        ["On-board Service", "3.88", "2.76", "1.12"],
        ["Seat Comfort", "4.16", "2.54", "1.62"],
        ["Leg Room Service", "3.75", "2.85", "0.90"],
        ["Cleanliness", "4.04", "2.34", "1.70"],
        ["Food and Drink", "3.86", "2.39", "1.47"],
        ["In-flight Service", "4.08", "3.10", "0.98"],
        ["In-flight Wifi Service", "3.08", "2.29", "0.79"],
        ["In-flight Entertainment", "4.27", "2.22", "2.05"],
        ["Baggage Handling", "4.06", "3.10", "0.96"],
        ["Mean across services", "3.69", "2.68", "1.02"],
    ],
    [2.75, 1.18, 1.18, 1.05],
    "Average service rating per K-Means cluster (0–5 scale).",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
add_figure("fig_80_16.png", "Average service rating per cluster. Cluster 0 is higher on every service dimension.", width=5.5)
body(
    "Cluster 0 is consistently higher than Cluster 1 on all 14 services, with a mean rating of 3.69 "
    "versus 2.68. The largest gaps are In-flight Entertainment (2.05 points), Cleanliness (1.70), Seat "
    "Comfort (1.62), Food and Drink (1.47), Online Boarding (1.16), and On-board Service (1.12). The "
    "smallest gap is Gate Location (0.07), which barely distinguishes the two groups. That pattern is "
    "informative: the clusters differ most on the aspects of the experience that passengers actively "
    "evaluate as part of the cabin product, and least on an airport attribute that may be outside the "
    "airline's control."
)

h2("9.4 Cluster Interpretation")
body(
    "The clusters were compared with the held-out target label, not used to construct it. The result is "
    "a strong validation of the service-rating space: Cluster 0 contains 64.2% satisfied passengers, "
    "while Cluster 1 contains 82.5% Neutral or Dissatisfied passengers."
)
add_table(
    ["Cluster", "Passengers", "Share of dataset", "Neutral or Dissatisfied", "Satisfied", "Interpretation"],
    [
        ["0", "72,159", "55.56%", "35.8%", "64.2%", "Positive service experience"],
        ["1", "57,721", "44.44%", "82.5%", "17.5%", "Poor service experience"],
    ],
    [0.70, 0.95, 1.00, 1.20, 0.90, 1.35],
    "Cluster sizes and satisfaction distribution. The target was not used to form the clusters.",
    aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
)
body(
    "The alignment is not perfect, which is expected: service ratings are informative but not "
    "deterministic, and some passengers with modest ratings may still report satisfaction for reasons "
    "not captured in the survey. The important point is that a segmentation built without the target "
    "still tracks the target closely. That supports the validity of the 14 ratings as measures of the "
    "passenger experience and gives the airline a practical way to identify a large at-risk group: "
    "Cluster 1 contains 44.44% of all passengers and is overwhelmingly dissatisfied."
)
callout(
    "Key finding",
    "K-Means with k = 2 splits the passenger base into a positive-experience group (55.56% of "
    "passengers, 64.2% satisfied) and a poor-experience group (44.44% of passengers, 82.5% "
    "dissatisfied). The poor-experience group is the natural priority cohort for service recovery.",
)

# ============================================================================
# 10. LAB 10 — HIERARCHICAL CLUSTERING
# ============================================================================
h1("10. Lab 10 — Hierarchical Clustering")
body(
    "Hierarchical clustering provides a second view of the same service-rating space. Instead of "
    "specifying k in advance and iterating centroids, agglomerative clustering starts with each "
    "passenger as its own cluster and repeatedly merges the closest pair until a stopping criterion is "
    "reached. The sequence of merges can be visualized as a dendrogram, which shows how observations "
    "group together at different distance thresholds."
)

h2("10.1 Objective and Sampling Strategy")
body(
    "Agglomerative clustering has roughly quadratic time and memory complexity because it must consider "
    "pairwise distances and maintain a proximity matrix. Running it on all 129,880 passengers was "
    "therefore not feasible in the available environment. The comparison was performed on a random "
    "sample of 2,000 passengers, using the same 14 standardised service ratings and the same k = 2 "
    "solution selected for K-Means."
)
body(
    "Ward linkage was used. Ward's criterion chooses merges that minimize the increase in within-cluster "
    "variance, which tends to produce compact, roughly equal-sized clusters and is a sensible match for "
    "standardised continuous features. The K-Means comparison was performed on exactly the same 2,000 "
    "rows: the full-data K-Means model was used to predict labels for those rows, and the silhouette "
    "score was computed against the agglomerative labels on the same feature matrix. This keeps the "
    "comparison as fair as the different training strategies allow."
)

h2("10.2 Dendrogram Analysis")
body(
    "The dendrogram below was built with Ward linkage on the first 200 rows of the 2,000-row sample, "
    "truncated to the last 20 merges so that the structure remains readable. The truncation is a "
    "visualization choice; it does not change the clustering result."
)
add_figure("fig_83_17.png", "Dendrogram for Ward linkage on a 200-passenger sample, showing the last 20 merges.", width=6.2)
body(
    "The dendrogram shows two dominant branches that remain separate until a relatively high merge "
    "distance. Within each branch, observations combine at much lower distances. This is consistent "
    "with the K-Means result: the data supports a broad two-group split better than a large number of "
    "finely separated groups. The dendrogram should be treated as an exploratory sanity check rather "
    "than as definitive evidence of two clusters, because it is based on a small sample and because the "
    "merge distances are not uniformly large."
)

h2("10.3 Comparison of K-Means and Agglomerative Clustering")
body(
    "The silhouette scores on the common 2,000-row sample are reported below. K-Means achieves 0.1897 "
    "and agglomerative clustering 0.1475, a difference of about 0.042 in favor of K-Means."
)
add_table(
    ["Method", "Silhouette score on the 2,000-row sample", "Notes"],
    [
        ["K-Means (k = 2)", "0.1897", "Labels predicted from the full-data model"],
        ["Agglomerative, Ward (k = 2)", "0.1475", "Fitted on the 2,000-row sample"],
    ],
    [2.15, 2.30, 1.75],
    "Silhouette comparison between K-Means and agglomerative clustering on the same sample.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "The comparison has a built-in asymmetry that should be acknowledged: K-Means was fitted on the "
    "full 129,880-row dataset and then asked to label the sample, whereas agglomerative clustering was "
    "fitted only on the 2,000 rows. K-Means therefore benefits from more information, and part of its "
    "silhouette advantage may reflect that difference rather than the algorithm alone. Even with this "
    "caveat, the practical conclusion is unchanged. K-Means produces better-separated groups on the "
    "sample, scales to the full dataset in seconds, and can assign new passengers to an existing "
    "cluster without refitting."
)

h2("10.4 Discussion")
body(
    "For this project, K-Means is the more useful segmentation method. It handles the full dataset, "
    "produces an interpretable two-cluster solution, and supports out-of-sample assignment, which is "
    "important if the airline wants to score new passengers as they complete a survey. Agglomerative "
    "clustering remains valuable as an exploratory tool: the dendrogram confirms that a two-branch "
    "structure is plausible and shows how tightly the observations combine within each branch. It is "
    "not the right choice for production-scale repeated segmentation here."
)
body(
    "The notebook does not implement DBSCAN. Density-based clustering would answer a different question "
    "— whether there are dense regions separated by sparse noise — and would require careful tuning of "
    "the neighborhood radius and minimum-points parameters. No DBSCAN result is reported because none "
    "was produced; the comparison in this section is strictly between K-Means and agglomerative "
    "clustering."
)

# ============================================================================
# 11. LAB 11 — OUTLIERS AND ENSEMBLE METHODS
# ============================================================================
h1("11. Lab 11 — Outlier Detection and Ensemble Methods")
body(
    "This lab combines two related activities. The first is outlier screening, which asks whether "
    "extreme values in the continuous features are data errors or genuine operational events. The "
    "second is ensemble learning, which asks whether a collection of decision trees can improve on the "
    "single tree from Lab 6. The two are connected by a practical question: should the long-delay "
    "observations be removed before modeling, or are they part of the signal the model needs to learn?"
)

h2("11.1 Outlier Screening with the IQR Method")
body(
    "The IQR rule flags an observation as an outlier if it lies below Q1 − 1.5 × IQR or above "
    "Q3 + 1.5 × IQR, where Q1 and Q3 are the first and third quartiles and IQR = Q3 − Q1. The rule is "
    "simple, does not assume normality, and is therefore appropriate for the skewed delay variables. "
    "The four continuous features were screened, and the results are shown below and in the accompanying "
    "box plots."
)
add_table(
    ["Feature", "Q1", "Q3", "IQR", "Lower bound", "Upper bound", "Flagged rows", "% flagged"],
    [
        ["Age", "27", "51", "24", "−9", "87", "0", "0.00%"],
        ["Flight Distance", "414", "1,744", "1,330", "−1,581", "3,739", "2,855", "2.20%"],
        ["Departure Delay", "0", "12", "12", "−18", "30", "18,098", "13.93%"],
        ["Arrival Delay", "0", "13", "13", "−19.5", "32.5", "17,492", "13.47%"],
    ],
    [1.25, 0.58, 0.58, 0.58, 0.82, 0.82, 0.82, 0.78],
    "IQR outlier screening for the four continuous features.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 7,
    first_col_bold=True,
)
add_figure("fig_23_1.png", "Box plots of the four continuous features used in the IQR screening.", width=6.2)
body(
    "Age has no IQR outliers: its lower bound is negative and its upper bound of 87 exceeds the observed "
    "maximum of 85. Flight Distance has a small flagged share of 2.20% (2,855 rows). The two delay "
    "variables have much larger flagged shares: 13.93% for Departure Delay (18,098 rows) and 13.47% for "
    "Arrival Delay (17,492 rows). These measured shares are higher than the notebook's earlier rough "
    "estimate of 8–9%, because the delay distribution is concentrated at zero, which makes the IQR "
    "narrow and the upper bound relatively low. The measured values are the ones used in this report."
)
body(
    "The decision was to retain the delay outliers. A departure delay of several hours is not a "
    "measurement error; it is a real operational event that may affect a passenger's satisfaction. "
    "Deleting those rows would remove precisely the observations that make the service problem visible "
    "and would bias the analysis toward on-time flights. Tree-based models are also robust to monotone "
    "transformations and do not require the outliers to be removed. The screening therefore serves a "
    "descriptive purpose — it documents how much of the sample sits in the tail — rather than a "
    "filtering purpose."
)
body(
    "The notebook implements the IQR method only. It does not run the Z-score method or Isolation "
    "Forest, so no results from those techniques are reported. This is a scope limitation of the "
    "executed work, not a claim that the IQR rule is the only appropriate method."
)

h2("11.2 Random Forest Classification")
body(
    "The Random Forest is an ensemble of decision trees. Each tree is trained on a bootstrap sample of "
    "the training data, and at each split only a random subset of features is considered. The "
    "predictions are combined by majority vote, which reduces the variance that makes a single tree "
    "unstable. The notebook used 150 trees, `max_depth=15`, `random_state=42`, and all available cores "
    "for training."
)
body(
    "The Random Forest was fitted on the full 103,904-row training set through the same preprocessing "
    "pipeline as the other classifiers and evaluated on the full test set. It achieved accuracy 0.9594, "
    "precision 0.9647, recall 0.9410, F1-score 0.9527, and ROC-AUC 0.9931. Those values are better than "
    "the single decision tree on every metric: accuracy rises by about 0.93 percentage points, F1 by "
    "1.11 points, and ROC-AUC by 1.02 points. The improvement is consistent with the variance-reduction "
    "argument. A single tree may split early on a feature that happens to fit the training sample; the "
    "ensemble averages over many such choices, so idiosyncratic splits have less influence."
)
body(
    "The ensemble does not lose the tree's ability to represent threshold effects. Each individual tree "
    "still makes axis-aligned splits, and the forest combines them. The result is a model that is both "
    "nonlinear and more stable than a single tree, at the cost of interpretability: the forest is no "
    "longer a single diagram that a stakeholder can read. That is why the next subsection uses "
    "permutation importance to recover a feature-level explanation."
)

h2("11.3 Permutation Feature Importance")
body(
    "The importance values in this subsection are computed on the tuned Random Forest, whose "
    "hyperparameter search is described in Section 12.4. The tuning is reported later to keep the Lab 11 "
    "discussion focused on interpretation; the model itself is the same final pipeline used in Section "
    "12.6."
)
body(
    "Permutation importance measures how much a fitted model's performance drops when a feature is "
    "randomly shuffled. If shuffling a feature destroys a lot of predictive signal, the model was "
    "relying on it; if shuffling changes little, the feature was not contributing much. The notebook "
    "computed permutation importance on a 4,000-row sample of the test set, with five repeats and "
    "F1-score as the scoring metric. The sample is used because full-test-set permutation importance "
    "requires re-scoring the model once per feature per repeat and is expensive to run repeatedly."
)
add_figure("fig_71_14.png", "Top 15 features by permutation importance on the tuned Random Forest (F1-based).", width=4.65)
body(
    "The ten most important features are listed below. Importance is the mean drop in F1-score when the "
    "feature is shuffled, averaged over five repeats."
)
add_table(
    ["Rank", "Feature", "Importance (mean F1 drop)"],
    [
        ["1", "Type of Travel", "0.1634"],
        ["2", "In-flight Wifi Service", "0.1436"],
        ["3", "Customer Type", "0.0693"],
        ["4", "Online Boarding", "0.0390"],
        ["5", "Seat Comfort", "0.0287"],
        ["6", "Check-in Service", "0.0285"],
        ["7", "In-flight Service", "0.0236"],
        ["8", "Cleanliness", "0.0227"],
        ["9", "Baggage Handling", "0.0221"],
        ["10", "Class", "0.0136"],
    ],
    [0.65, 3.20, 2.35],
    "Top ten features by permutation importance on the tuned Random Forest.",
    aligns=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
)
body(
    "Type of Travel is the dominant feature by a wide margin (0.1634), followed by In-flight Wifi "
    "Service (0.1436) and Customer Type (0.0693). Online Boarding (0.0390), Seat Comfort (0.0287), and "
    "Check-in Service (0.0285) form a second tier. The remaining features in the top ten — In-flight "
    "Service, Cleanliness, Baggage Handling, and Class — contribute smaller but non-zero amounts. "
    "Gender and Flight Distance are not in the top ten and contribute effectively nothing, which is "
    "consistent with the weak bivariate correlations for those variables."
)
body(
    "Two cautions apply. First, permutation importance is not causal. A feature can be important because "
    "it is a proxy for another variable that is not in the model or because it captures a segment "
    "difference rather than a service effect. Type of Travel is the clearest example: it may summarize "
    "a bundle of expectations and fare conditions rather than a single actionable service. Second, "
    "correlated features share importance. When several service ratings move together, shuffling one "
    "may leave the others to compensate, so each individual importance can understate the joint "
    "contribution of the group. The rankings should therefore be read as a description of the fitted "
    "model's reliance, not as a decomposition of satisfaction into independent causes."
)

h2("11.4 Interpretation")
body(
    "The Random Forest is the strongest untuned classifier in the project, and its importance profile "
    "adds a service-level insight that the correlation analysis could not provide. In-flight Wifi "
    "Service is the second most important feature overall and the most important individual service, "
    "while also being the lowest-rated service in the dataset. That combination — low performance and "
    "high model reliance — is exactly the profile of a high-leverage improvement target. Online "
    "Boarding is the strongest single correlate and remains important in the forest, but its "
    "permutation importance is smaller than Wifi's, suggesting that part of its bivariate signal is "
    "shared with other service ratings."
)
body(
    "Type of Travel's dominance also deserves attention. It is a passenger/trip characteristic, not a "
    "service that the airline can directly improve. Its importance indicates that the model separates "
    "business and personal travelers very effectively, which is consistent with the 58.37% versus "
    "10.13% satisfaction gap in the EDA. The practical implication is that a single service strategy is "
    "unlikely to work across both segments: personal travelers may have different expectations, "
    "constraints, and tolerance for ancillary fees. This is a segmentation insight that leads naturally "
    "into the mini-project synthesis."
)

# ============================================================================
# 12. LAB 12 — END-TO-END MINI PROJECT
# ============================================================================
h1("12. Lab 12 — End-to-End Mini Project")
body(
    "The final lab is the integrating mini project. It brings together the preprocessing, exploratory "
    "analysis, classification, clustering, and feature-importance results into a single workflow, adds "
    "the feature engineering and model-selection steps that were not part of the individual labs, and "
    "translates the technical results into service optimization recommendations. It also documents the "
    "final model artifact and verifies that the saved pipeline can be reloaded and used for prediction."
)

h2("12.1 Feature Engineering for Interpretation")
body(
    "Two derived features were created after the exploratory analysis. The first, `Overall Service "
    "Score`, is the mean of the 14 individual service ratings for each passenger. It summarizes the "
    "overall service experience in a single number and makes it easy to compare passengers who rated "
    "many services. Its distribution has a mean of 3.241, a standard deviation of 0.662, and a range "
    "from 1.071 to 5.000. The second, `Age Group`, bins Age into Youth (under 25), Adult (25–59), and "
    "Senior (60 and over). It converts a continuous demographic variable into a form that is easier to "
    "communicate to airline stakeholders."
)
add_figure("fig_51_12.png", "Overall Service Score by satisfaction class.", width=4.25)
body(
    "The engineered features produce two interpretable findings. First, the Overall Service Score is "
    "clearly higher for satisfied passengers, which confirms that the 14 ratings carry a coherent "
    "overall signal rather than unrelated noise. Second, satisfaction varies by age group: 49.76% of "
    "adult passengers are satisfied, compared with 27.74% of youth and 27.07% of seniors. Adults make up "
    "93,003 passengers (71.61%), youth 24,662 (18.99%), and seniors 12,215 (9.40%). The similar rates "
    "for youth and seniors suggest that the age effect is not a simple linear trend; both ends of the "
    "age range are less satisfied than the middle."
)
body(
    "Neither engineered feature was fed into the classification pipeline. The model already receives "
    "the 14 individual ratings directly, so adding their mean would be redundant and highly correlated "
    "with existing inputs. The Age Group variable would also discard information relative to the raw "
    "Age feature. Both derived features are used for interpretation and reporting, not for prediction, "
    "which avoids unnecessary complexity and matches the project's ReactJS interface: that interface "
    "collects individual, human-readable service ratings rather than a pre-computed score. The "
    "engineered features are also safe from target leakage because they are computed only from input "
    "columns, never from `Satisfaction`."
)

h2("12.2 Feature Selection Strategy")
body(
    "No predictor was dropped before modeling. The 22 remaining input features were retained for three "
    "reasons. First, tree-based and kernel models can tolerate weak or redundant features better than "
    "linear models; they do not require a sparse feature set to perform well. Second, keeping all "
    "fields allows the prediction interface to collect a complete, realistic passenger profile rather "
    "than a reduced form that might confuse users or omit relevant information. Third, feature "
    "selection performed before seeing the model's behavior can inadvertently remove a feature that is "
    "useful only in interaction with another variable."
)
body(
    "Feature selection was therefore handled as interpretation rather than elimination. The correlation "
    "analysis in Section 3.4 provided an initial ranking, and the permutation importance in Section "
    "11.3 provided a model-based ranking after training. The two views agree on the central service "
    "variables — Online Boarding, In-flight Entertainment, Seat Comfort, and On-board Service are "
    "prominent in both — while the permutation analysis adds Type of Travel, In-flight Wifi Service, "
    "and Customer Type as the features the fitted forest relies on most. This combined evidence feeds "
    "the service optimization synthesis in Section 12.5."
)

h2("12.3 Comparative Model Evaluation")
body(
    "The four untuned classifiers and the tuned Random Forest are compared below on the same held-out "
    "test set. The tuned model is included here because it is the final selected model; its tuning "
    "procedure is described in Section 12.4."
)
add_table(
    ["Model", "Accuracy", "Precision", "Recall", "F1-score", "ROC-AUC", "Configuration note"],
    [
        ["Decision Tree", "0.9501", "0.9581", "0.9257", "0.9416", "0.9829", "Single tree, depth 12"],
        ["Naive Bayes", "0.8626", "0.8592", "0.8178", "0.8380", "0.9228", "GaussianNB"],
        ["SVM", "0.9379", "0.9365", "0.9193", "0.9278", "0.9811", "RBF, trained on 8,000 rows"],
        ["Random Forest", "0.9594", "0.9647", "0.9410", "0.9527", "0.9931", "150 trees, depth 15"],
        ["Random Forest (tuned)", "0.9632", "0.9731", "0.9414", "0.9570", "0.9941", "Final model"],
    ],
    [1.45, 0.68, 0.68, 0.68, 0.68, 0.68, 1.35],
    "Test-set performance of all classifiers. The SVM was trained on a stratified 8,000-row subsample but evaluated on the full test set.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5 + [WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
add_figure("fig_confusion_2x2.png", "Confusion matrices for the four untuned classifiers on the test set.", width=5.9)
body(
    "The ranking is stable and matches the progression of the labs: Random Forest (tuned) > Random "
    "Forest (untuned) > Decision Tree > SVM > Naive Bayes. The tuned forest has the best score on every "
    "metric, with 96.32% accuracy, 97.31% precision, 94.14% recall, 0.9570 F1, and 0.9941 ROC-AUC. The "
    "untuned forest is close behind, and the single tree is a competitive third. The SVM is fourth, and "
    "Naive Bayes is clearly last. The confusion matrices show the same ordering visually: the Random "
    "Forest has the smallest off-diagonal counts, while Naive Bayes has the largest."
)
body(
    "Accuracy alone would not justify this ranking. With a 56.55% majority class, a trivial classifier "
    "would already score 56.55%, so every model must be judged against that baseline. All five models "
    "clear it by a wide margin, but the differences among the top three are small enough that F1 and "
    "ROC-AUC are more informative than accuracy. F1 rewards a balance between precision and recall, "
    "which matters if the airline wants to identify satisfied and dissatisfied passengers with similar "
    "reliability. ROC-AUC summarizes ranking quality across all thresholds, which matters if the "
    "airline later wants to adjust the decision threshold to prioritize recall — for example, to flag "
    "more at-risk passengers for service recovery."
)
body(
    "The choice of Random Forest as the final model is therefore based on its best F1 and ROC-AUC, not "
    "on a marginal accuracy gain. The ensemble also has a practical advantage: it is robust to the "
    "correlated service ratings and does not require the strong distributional assumptions of Naive "
    "Bayes or the quadratic training cost of the kernel SVM. Its main cost is interpretability, which is "
    "addressed by the permutation importance analysis rather than by abandoning the model."
)

h2("12.4 Hyperparameter Tuning")
body(
    "The Random Forest was tuned with `GridSearchCV` using 3-fold stratified cross-validation and "
    "F1-score as the optimization metric. To keep the search computationally reasonable, the grid "
    "search was run on a stratified 20,000-row subsample of the training set. The best hyperparameters "
    "were then refitted on the full 103,904-row training set, and the refitted model was evaluated on "
    "the untouched test set. This two-stage procedure is a common compromise for large tabular "
    "datasets: the search explores the hyperparameter space on a representative subset, while the final "
    "model uses all available training data."
)
add_table(
    ["Component", "Values tested or used", "Selected"],
    [
        ["Number of trees (`n_estimators`)", "100, 150", "150"],
        ["Maximum depth (`max_depth`)", "15, None (unlimited)", "None"],
        ["Cross-validation", "3-fold stratified", "—"],
        ["Search sample", "20,000 training rows, stratified", "—"],
        ["Scoring metric", "F1-score", "—"],
        ["Best cross-validation F1", "—", "0.9409"],
    ],
    [2.25, 2.55, 1.40],
    "Random Forest hyperparameter grid, search configuration, and selected values.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "The grid contained four candidate combinations (two values of `n_estimators` times two values of "
    "`max_depth`), and 3-fold cross-validation produced 12 fits in total. The best combination was 150 "
    "trees with unlimited depth, with a cross-validation F1 of 0.9409 on the search sample. The "
    "unlimited depth result is plausible here: the forest averages many deep trees, and the bootstrap "
    "sampling plus feature subsampling provide enough regularization to prevent the catastrophic "
    "overfitting that an unlimited single tree would show."
)
add_table(
    ["Model", "Accuracy", "Precision", "Recall", "F1-score", "ROC-AUC"],
    [
        ["Random Forest (untuned)", "0.9594", "0.9647", "0.9410", "0.9527", "0.9931"],
        ["Random Forest (tuned)", "0.9632", "0.9731", "0.9414", "0.9570", "0.9941"],
        ["Change", "+0.0038", "+0.0084", "+0.0004", "+0.0043", "+0.0010"],
    ],
    [2.00, 0.90, 0.90, 0.80, 0.80, 0.87],
    "Effect of hyperparameter tuning on Random Forest test performance.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5,
    first_col_bold=True,
)
body(
    "Tuning produces a modest but consistent improvement. Accuracy rises by 0.38 percentage points, "
    "precision by 0.84 points, F1 by 0.43 points, and ROC-AUC by 0.10 points; recall is essentially "
    "unchanged. The gain is small because the untuned forest was already well configured for this "
    "problem. The more important benefit is methodological: cross-validation was used to choose the "
    "hyperparameters, rather than a single train/validation split, which reduces the risk of selecting "
    "a configuration that happens to fit one particular split. This protects the integrity of model "
    "selection in the same spirit as the pipeline protects the integrity of preprocessing."
)

h2("12.5 Service Optimization Analysis")
body(
    "The technical results are only useful if they change a decision. This section synthesizes the "
    "correlation analysis, the service-rating averages, the passenger-group satisfaction rates, the "
    "permutation importance, and the clustering results into seven findings and a set of practical "
    "recommendations."
)
body(
    "**1. Which services have the strongest relationship with satisfaction?** Among the service "
    "features, Online Boarding has the strongest bivariate correlation (r = 0.502), followed by "
    "In-flight Entertainment (0.398) and Seat Comfort (0.349). The permutation importance on the tuned "
    "forest adds In-flight Wifi Service as the second most important feature overall (0.1436), after "
    "Type of Travel. The two views are complementary: correlation describes the raw association, while "
    "permutation importance describes the fitted model's reliance after controlling for the other "
    "features."
)
body(
    "**2. Which services have the lowest ratings?** In-flight Wifi Service is lowest at 2.73 out of 5, "
    "followed by Ease of Online Booking at 2.76, Gate Location at 2.98, and Departure and Arrival Time "
    "Convenience at 3.06. At the other end, In-flight Service (3.64) and Baggage Handling (3.63) are the "
    "strongest-rated services. The airline therefore has a clear performance gap in connectivity and "
    "digital booking, and a relative strength in onboard service and baggage handling."
)
body(
    "**3. Which passenger groups are most dissatisfied?** Personal travelers have a 10.13% satisfaction "
    "rate, compared with 58.37% for business travelers. First-time customers have a 23.97% rate, "
    "compared with 47.81% for returning customers. Economy passengers have an 18.77% rate, compared with "
    "69.44% for Business class. These gaps are much larger than the differences between individual "
    "service ratings, which suggests that segment-level expectations and fare conditions matter as much "
    "as any single service touchpoint."
)
body(
    "**4. Which travel classes have poorer service experiences?** Economy (18.77% satisfied) and "
    "Economy Plus (24.64%) lag far behind Business class (69.44%). The similarity between Economy and "
    "Economy Plus is notable: the small premium cabin does not close the satisfaction gap with Business "
    "class, even though it is positioned as an intermediate product. This may reflect unmet expectations "
    "about what Economy Plus should include."
)
body(
    "**5. Which features are most important for prediction?** The tuned forest relies most on Type of "
    "Travel (0.1634), In-flight Wifi Service (0.1436), Customer Type (0.0693), Online Boarding (0.0390), "
    "and Seat Comfort (0.0287). Gender and Flight Distance contribute effectively nothing. The "
    "importance of Type of Travel and Customer Type is a reminder that the model is partly learning "
    "who the passenger is, not only what the airline did."
)
body(
    "**6. Which passenger clusters require service improvement?** Cluster 1, the poor-service "
    "experience group, contains 57,721 passengers — 44.44% of the dataset — and 82.5% of its members "
    "are Neutral or Dissatisfied. That is the clearest priority cohort in the analysis. Because the "
    "cluster is defined only by service ratings, it can be monitored without waiting for the satisfaction "
    "label: a passenger whose ratings begin to resemble the Cluster 1 profile can be flagged for "
    "proactive service recovery."
)
body(
    "**7. What should the airline do?** The recommendations below follow directly from the evidence and "
    "are ordered by leverage rather than by ease of implementation."
)
bullets([
    "**Prioritize In-flight Wifi Service.** It is simultaneously the lowest-rated service (2.73) and "
    "the second most important feature in the tuned model. Improving connectivity is therefore the "
    "single clearest high-leverage investment in this dataset.",
    "**Treat Ease of Online Booking as a secondary digital priority.** It is the second-lowest-rated "
    "service (2.76), but its correlation with satisfaction is weaker (0.169) and it does not appear in "
    "the top ten permutation importances. It is worth improving, but it is not the first lever.",
    "**Investigate the Personal-travel and Economy-class experience as a structural issue.** The gap "
    "between Business (58–69% satisfied) and Personal/Economy (10–19%) is far larger than any single "
    "service-rating gap. This points to fare-class amenities, seat and legroom conditions, and "
    "expectation management rather than an isolated service failure.",
    "**Build a first-time-customer onboarding and service-recovery focus.** First-time customers are "
    "satisfied at roughly half the rate of returning customers. Targeted communication, realistic "
    "expectation setting, and faster recovery when something goes wrong could improve both satisfaction "
    "and retention.",
    "**Use Cluster 1 as a monitoring cohort.** The 44.44% of passengers with consistently low service "
    "ratings can be flagged before they churn. The model's predicted probability of satisfaction is a "
    "useful continuous score for prioritizing outreach, even when the thresholded label is uncertain.",
    "**Do not over-invest in Gate Location on satisfaction grounds.** It is low-rated (2.98) but almost "
    "uncorrelated with satisfaction (−0.003) and does not appear in the top permutation importances. "
    "Passengers may dislike it without holding it against the airline's overall service.",
])
body(
    "The table below combines the three quantitative lenses — average rating, correlation with "
    "satisfaction, and permutation importance — into a single priority view. It is the most compact "
    "summary of the service optimization argument in this report."
)
add_table(
    ["Service", "Mean rating", "Correlation", "Permutation importance", "Priority"],
    [
        ["In-flight Wifi Service", "2.73", "0.283", "0.1436 (rank 2)", "Highest"],
        ["Online Boarding", "3.25", "0.502", "0.0390 (rank 4)", "High"],
        ["Seat Comfort", "3.44", "0.349", "0.0287 (rank 5)", "High"],
        ["In-flight Entertainment", "3.36", "0.398", "Not in top 10", "High"],
        ["Ease of Online Booking", "2.76", "0.169", "Not in top 10", "Medium"],
        ["Gate Location", "2.98", "−0.003", "Not in top 10", "Low"],
    ],
    [1.85, 0.95, 0.95, 1.65, 0.87],
    "Service priority synthesis combining rating level, correlation, and permutation importance.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER],
    first_col_bold=True,
)
body(
    "The synthesis resolves an ambiguity that no single measure could resolve. A low average rating "
    "identifies a pain point, but not necessarily a satisfaction lever: Gate Location is disliked yet "
    "uncorrelated with the overall label. A high correlation identifies a lever, but not necessarily a "
    "weakness: Online Boarding is strongly correlated but rated in the middle of the pack. Only the "
    "combination identifies the best targets. In-flight Wifi is the clearest case because it is low on "
    "performance and high on model reliance. Online Boarding, Seat Comfort, and In-flight Entertainment "
    "are strong levers that are already performing reasonably well, so the goal there is to protect "
    "performance rather than rescue it."
)

h2("12.6 Final Model, Persistence and Verification")
body(
    "The final model is the tuned Random Forest pipeline. It contains the `ColumnTransformer` "
    "preprocessing step and the fitted Random Forest classifier as a single scikit-learn `Pipeline`. "
    "Saving the complete pipeline, rather than only the classifier, guarantees that new passenger data "
    "passes through exactly the same scaling and one-hot encoding that were used during training. This "
    "eliminates training-serving skew, the common failure mode in which a model is trained on one "
    "representation and served another."
)
body(
    "The pipeline was saved with `joblib.dump` to `model/service_optimization_model.pkl`. The file is "
    "139.1 MB, which is consistent with a 150-tree forest stored together with its preprocessing "
    "objects. The notebook then reloaded the file with `joblib.load` and used it to predict a sample "
    "passenger: a 35-year-old male business traveler, returning customer, Business class, 1,200 miles, "
    "10 minutes departure delay, 5 minutes arrival delay, with mostly 4–5 ratings across the 14 service "
    "dimensions. The reloaded model predicted Satisfied with a probability of 81.33%."
)
add_table(
    ["Verification item", "Result"],
    [
        ["Predicted class", "Satisfied"],
        ["Probability of Neutral or Dissatisfied", "0.1867 (18.67%)"],
        ["Probability of Satisfied", "0.8133 (81.33%)"],
        ["Model source", "Reloaded from model/service_optimization_model.pkl"],
    ],
    [2.55, 3.55],
    "Saved-model verification on a sample business traveler.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)
body(
    "The verification confirms that serialization round-trips correctly and that the saved artifact "
    "reproduces the trained model's behavior outside the training session. It also demonstrates the "
    "intended inference pattern: load the pipeline once, pass a single passenger record as a "
    "dictionary or one-row DataFrame, and read both the predicted label and the class probabilities. "
    "The notebook stops at this verification. The project's FastAPI backend and ReactJS interface are "
    "outside the notebook's executed scope, but the saved pipeline is the artifact those components "
    "consume, and the reload test is the handoff point between the modeling work and the application."
)
body(
    "Reporting both probabilities, rather than only the label, is important for the service use case. A "
    "prediction of Satisfied at 81.33% is informative; the same label at 51% would be much less "
    "actionable. The probability gives the airline a way to rank passengers by risk and to tune the "
    "decision threshold to its operational priorities."
)

h2("12.7 Mini-Project Requirement Coverage")
body(
    "The manual specifies seven requirements for the end-to-end mini project. The table below maps each "
    "requirement to the section where it is addressed and to the concrete evidence produced."
)
add_table(
    ["Manual requirement", "Where addressed", "Evidence"],
    [
        ["1. Select a real dataset", "Section 0.4", "Airline Passenger Satisfaction, 129,880 rows and 24 columns"],
        ["2. Perform preprocessing", "Section 2", "Median imputation, duplicate check, one-hot/label encoding, scaling, stratified split"],
        ["3. Conduct exploratory analysis", "Section 3", "Target balance, group satisfaction rates, correlations, service-rating levels"],
        ["4. Apply at least one classification algorithm", "Sections 6–8 and 11", "Decision Tree, GaussianNB, RBF SVM, Random Forest; tuned forest F1 = 0.9570"],
        ["5. Apply at least one clustering or association rule method", "Sections 9–10", "K-Means with k = 2; agglomerative comparison and dendrogram"],
        ["6. Evaluate the results", "Sections 6–12", "Accuracy, precision, recall, F1, ROC-AUC, confusion matrices, silhouette, permutation importance"],
        ["7. Present findings in a short report and presentation", "This report", "Written report completed; the notebook does not contain presentation slides"],
    ],
    [1.55, 1.30, 3.30],
    "Coverage of the seven mini-project requirements stated in the lab manual.",
    aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
    first_col_bold=True,
)

h2("12.8 Conclusion")
body(
    "This project set out to turn airline passenger feedback into service optimization decisions. The "
    "dataset contained 129,880 passengers, 14 service ratings, four demographic/trip predictors, and a "
    "binary satisfaction label. After median imputation of 393 missing arrival-delay values, removal of "
    "the identifier column, one-hot encoding of the nominal predictors, standardisation of the numeric "
    "features, and a stratified 80/20 split, the data supported a complete modeling workflow without "
    "test-set leakage in the scaling or encoding steps."
)
body(
    "The exploratory analysis found that service experience, not delay, dominates the raw signal. "
    "Online Boarding was the strongest single correlate with satisfaction (r = 0.502), followed by "
    "In-flight Entertainment (0.398) and Seat Comfort (0.349). In-flight Wifi was the lowest-rated "
    "service (2.73) and also a meaningful correlate (0.283), making it the clearest high-leverage "
    "improvement target. Personal travelers (10.13% satisfied), first-time customers (23.97%), and "
    "Economy passengers (18.77%) were far less satisfied than business travelers (58.37%), returning "
    "customers (47.81%), and Business-class passengers (69.44%)."
)
body(
    "Four classifiers were compared. The Decision Tree reached 95.01% accuracy and 0.9416 F1; Gaussian "
    "Naive Bayes reached 86.26% and 0.8380; the RBF SVM reached 93.79% and 0.9278; and the Random Forest "
    "reached 95.94% and 0.9527. Hyperparameter tuning selected 150 trees with unlimited depth, and the "
    "refitted model improved to 96.32% accuracy, 0.9570 F1, and 0.9941 ROC-AUC. The final model was "
    "saved as a complete pipeline and reloaded successfully, predicting a high-rating business traveler "
    "as Satisfied with 81.33% probability."
)
body(
    "K-Means with k = 2 produced an interpretable service segmentation: a positive-experience cluster "
    "(55.56% of passengers, 64.2% satisfied) and a poor-experience cluster (44.44%, 82.5% dissatisfied). "
    "The silhouette score was highest at k = 2 (0.1821), and agglomerative clustering on a 2,000-row "
    "sample produced a comparable but lower separation (0.1475). The poor-experience cluster is the "
    "natural priority cohort for service recovery, and the dendrogram supports the plausibility of a "
    "two-branch structure."
)
body(
    "The project has several limitations. The data is observational and cross-sectional, so the "
    "associations and feature importances should not be interpreted as causal effects. The service "
    "ratings are ordinal and include a “not applicable” zero that the notebook treats as a low numeric "
    "rating. The evaluation uses a single train/test split rather than repeated nested validation, and "
    "the SVM and grid search use subsamples for computational reasons. The silhouette value of 0.1821 "
    "indicates modest cluster separation, and the dataset is dominated by returning business travelers, "
    "which may limit generalization to other passenger populations. Finally, the notebook does not "
    "implement the association-rule, DBSCAN, Z-score, or Isolation Forest components of the manual, so "
    "those techniques remain outside the reported evidence."
)
body(
    "Within those limits, the project achieves its practical objective. It identifies a small set of "
    "high-leverage service priorities — led by In-flight Wifi, Online Boarding, Seat Comfort, and "
    "In-flight Entertainment — distinguishes them from low-rated but low-impact attributes such as Gate "
    "Location, and shows that a large poor-experience segment can be monitored using service ratings "
    "alone. The tuned Random Forest provides a reusable prediction artifact, and the combination of "
    "predictive modeling, feature importance, and clustering gives the airline a defensible basis for "
    "where to invest next."
)

# ============================================================================
# 13. REFERENCES
# ============================================================================
h1("13. References")
body(
    "Airline Passenger Satisfaction dataset. Kaggle. The dataset contains 129,880 passenger records and "
    "24 columns covering demographics, trip details, 14 service ratings, and the satisfaction label."
)
body(
    "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., "
    "Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., "
    "Perrot, M., & Duchesnay, É. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine "
    "Learning Research, 12, 2825–2830."
)
body(
    "McKinney, W. (2010). Data Structures for Statistical Computing in Python. Proceedings of the 9th "
    "Python in Science Conference, 56–61."
)
body(
    "Hunter, J. D. (2007). Matplotlib: A 2D Graphics Environment. Computing in Science & Engineering, "
    "9(3), 90–95."
)
body(
    "Waskom, M. L. (2021). Seaborn: Statistical Data Visualization. Journal of Open Source Software, "
    "6(60), 3021."
)
body(
    "Virtanen, P., Gommers, R., Oliphant, T. E., et al. (2020). SciPy 1.0: Fundamental Algorithms for "
    "Scientific Computing in Python. Nature Methods, 17, 261–272."
)
body(
    "Han, J., Kamber, M., & Pei, J. (2011). Data Mining: Concepts and Techniques (3rd ed.). Morgan "
    "Kaufmann."
)

# ----------------------------------------------------------------------------
# Metadata and save
# ----------------------------------------------------------------------------
cp = doc.core_properties
cp.title = "Service Optimization Using Airline Passenger Satisfaction Data"
cp.subject = "Data Mining Lab Project Report"
cp.author = "Taisin Nigar Tisha"
cp.comments = "Generated from notebook/service_optimization.ipynb executed outputs."
cp.keywords = "data mining, airline satisfaction, classification, clustering, Random Forest, K-Means"
cp.category = "Lab Report"

out_path = os.path.join("reports", "Service_Optimization_Lab_Report.docx")
doc.save(out_path)
print(f"Saved {out_path}")
print(f"Tables: {TABLE_NO}, Figures: {FIGURE_NO}")
