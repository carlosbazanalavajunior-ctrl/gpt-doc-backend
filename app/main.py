from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Literal, Union
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import requests


app = FastAPI(
    title="GPT Doc Backend",
    version="0.5.0",
    description="Backend para generación de cartas e informes profesionales en formato Word (.docx)."
)


# =========================
# MODELOS
# =========================

class FigureItem(BaseModel):
    title: Optional[str] = None
    url: str
    caption: Optional[str] = None
    width_inches: Optional[float] = 5.8


class ChartItem(BaseModel):
    title: Optional[str] = None
    chart_config: dict
    caption: Optional[str] = None
    width_inches: Optional[float] = 5.8


class TableItem(BaseModel):
    title: Optional[str] = None
    headers: List[str]
    rows: List[List[Union[str, int, float]]]
    note: Optional[str] = None


class ReportSection(BaseModel):
    heading: str
    level: int = Field(default=1, ge=1, le=4)
    paragraphs: Optional[List[str]] = None
    bullets: Optional[List[str]] = None
    numbered: Optional[List[str]] = None
    tables: Optional[List[TableItem]] = None
    figures: Optional[List[FigureItem]] = None
    charts: Optional[List[ChartItem]] = None


class LetterPayload(BaseModel):
    filename: str = "carta"
    logo_url: Optional[str] = None
    organization: Optional[str] = None
    city_date: Optional[str] = None
    recipient_name: Optional[str] = None
    recipient_title: Optional[str] = None
    recipient_organization: Optional[str] = None
    subject: Optional[str] = None
    greeting: Optional[str] = "De mi consideración:"
    body: List[str]
    closing: Optional[str] = None
    signature_name: Optional[str] = None
    signature_title: Optional[str] = None
    annexes: Optional[List[str]] = None


class ReportPayload(BaseModel):
    filename: str = "informe"
    report_kind: Optional[str] = "Informe"
    logo_url: Optional[str] = None
    institution: Optional[str] = None
    faculty: Optional[str] = None
    department: Optional[str] = None
    title: str
    subtitle: Optional[str] = None
    author: Optional[str] = None
    reviewer: Optional[str] = None
    city: Optional[str] = None
    date: Optional[str] = None
    header_text: Optional[str] = None
    footer_text: Optional[str] = None
    executive_summary: Optional[List[str]] = None
    sections: Optional[List[ReportSection]] = None
    conclusions: Optional[List[str]] = None
    recommendations: Optional[List[str]] = None
    references: Optional[List[str]] = None


class GenerateDocumentRequest(BaseModel):
    document_type: Literal["carta", "informe"]
    letter: Optional[LetterPayload] = None
    report: Optional[ReportPayload] = None


# =========================
# HELPERS XML / DOCX
# =========================

def set_run_font(run, name="Times New Roman", size=12, bold=False, italic=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def fetch_binary(url: str) -> BytesIO:
    response = requests.get(
        url,
        timeout=25,
        headers={"User-Agent": "Mozilla/5.0"}
    )
    response.raise_for_status()
    return BytesIO(response.content)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("left", "top", "right", "bottom", "insideH", "insideV"):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            tag = f"w:{edge}"

            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            for key in ["val", "sz", "space", "color"]:
                if key in edge_data:
                    element.set(qn(f"w:{key}"), str(edge_data[key]))


def shade_cell(cell, fill="EDEDED"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def remove_cell_borders(cell):
    set_cell_border(
        cell,
        top={"val": "nil"},
        bottom={"val": "nil"},
        left={"val": "nil"},
        right={"val": "nil"},
    )


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")

    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def add_horizontal_rule(paragraph, color="A6A6A6", size="6"):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pbdr = pPr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        pPr.append(pbdr)

    bottom = pbdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pbdr.append(bottom)

    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)


# =========================
# ESTILO GENERAL
# =========================

def configure_document(doc: Document):
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.1)
    section.right_margin = Inches(1.0)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.first_line_indent = Inches(0.28)

    styles = [
        ("Heading 1", 14, False, Pt(12), Pt(6)),
        ("Heading 2", 13, False, Pt(10), Pt(5)),
        ("Heading 3", 12, False, Pt(8), Pt(4)),
        ("Heading 4", 12, True, Pt(6), Pt(4)),
    ]

    for style_name, size, italic, before, after in styles:
        style = doc.styles[style_name]
        style.font.name = "Times New Roman"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.italic = italic
        style.paragraph_format.space_before = before
        style.paragraph_format.space_after = after
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.first_line_indent = Inches(0)


def add_logo(doc: Document, logo_url: Optional[str], width=1.0):
    if not logo_url:
        return
    try:
        stream = fetch_binary(logo_url)
        doc.add_picture(stream, width=Inches(width))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass


def add_rich_text_paragraph(doc: Document, text: str, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=True):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0.28) if first_indent else Inches(0)

    parts = text.split("**")
    for i, part in enumerate(parts):
        run = p.add_run(part)
        set_run_font(run, size=12, bold=(i % 2 == 1))
    return p


def add_bullet_list(doc: Document, items: List[str]):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(0)

        parts = item.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_run_font(run, size=11, bold=(i % 2 == 1))


def add_numbered_list(doc: Document, items: List[str]):
    for item in items:
        p = doc.add_paragraph(style="List Number")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(0)

        parts = item.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_run_font(run, size=11, bold=(i % 2 == 1))


# =========================
# HEADER / FOOTER
# =========================

def add_header_footer(section, header_text: Optional[str], footer_text: Optional[str]):
    if header_text:
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.text = ""
        run = hp.add_run(header_text)
        set_run_font(run, size=9, italic=True, color="666666")

    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.text = ""

    if footer_text:
        run = fp.add_run(footer_text + " | Página ")
        set_run_font(run, size=9, color="666666")
    else:
        run = fp.add_run("Página ")
        set_run_font(run, size=9, color="666666")

    add_page_number(fp)


# =========================
# PORTADA INFORME
# =========================

def add_report_cover(doc: Document, report: dict):
    add_logo(doc, report.get("logo_url"), width=1.1)

    institution = (report.get("institution") or "").strip()
    faculty = (report.get("faculty") or "").strip()
    department = (report.get("department") or "").strip()
    report_kind = (report.get("report_kind") or "Informe").strip()
    title = (report.get("title") or "").strip()
    subtitle = (report.get("subtitle") or "").strip()
    author = (report.get("author") or "").strip()
    reviewer = (report.get("reviewer") or "").strip()
    city = (report.get("city") or "").strip()
    date = (report.get("date") or "").strip()

    if institution:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(institution.upper())
        set_run_font(r, size=12, bold=True)

    if faculty:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(faculty)
        set_run_font(r, size=11, bold=True)

    if department:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run(department)
        set_run_font(r, size=10, color="555555")

    sep = doc.add_paragraph()
    sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sep.paragraph_format.space_after = Pt(10)
    add_horizontal_rule(sep, color="BFBFBF", size="6")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(report_kind.upper())
    set_run_font(r, size=12, bold=True, color="404040")

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(title)
        set_run_font(r, size=16, bold=True)

    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(18)
        r = p.add_run(subtitle)
        set_run_font(r, size=11, italic=True, color="555555")

    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        r1 = p.add_run("Autor: ")
        set_run_font(r1, size=11, bold=True)
        r2 = p.add_run(author)
        set_run_font(r2, size=11)

    if reviewer:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        r1 = p.add_run("Revisor / Asesor: ")
        set_run_font(r1, size=11, bold=True)
        r2 = p.add_run(reviewer)
        set_run_font(r2, size=11)

    if city or date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(10)
        text = f"{city}, {date}" if city and date else city or date
        r = p.add_run(text)
        set_run_font(r, size=11, color="555555")

    doc.add_page_break()


# =========================
# TABLAS / FIGURAS / GRÁFICOS
# =========================

def add_apa_table(doc: Document, table_data: dict):
    title = (table_data.get("title") or "").strip()
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])
    note = (table_data.get("note") or "").strip()

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(title)
        set_run_font(r, size=11, italic=True)

    ncols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if ncols == 0:
        return

    table = doc.add_table(rows=1, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        shade_cell(cell, "F2F2F2")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)

        r = p.add_run(str(header))
        set_run_font(r, size=10.5, bold=True)

        set_cell_border(
            cell,
            top={"sz": 10, "val": "single", "color": "808080"},
            bottom={"sz": 8, "val": "single", "color": "808080"},
            left={"val": "nil"},
            right={"val": "nil"},
        )

    # Data rows
    for row in rows:
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            cell = row_cells[i]
            cell.text = ""
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            r = p.add_run(str(value))
            set_run_font(r, size=10.5)

            set_cell_border(
                cell,
                top={"val": "nil"},
                bottom={"sz": 4, "val": "single", "color": "D9D9D9"},
                left={"val": "nil"},
                right={"val": "nil"},
            )

    doc.add_paragraph()

    if note:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.first_line_indent = Inches(0)
        note_text = note if note.lower().startswith("nota.") else f"Nota. {note}"
        r = p.add_run(note_text)
        set_run_font(r, size=9.5, italic=True, color="555555")


def add_figure_from_url(doc: Document, figure: dict):
    try:
        title = figure.get("title")
        caption = figure.get("caption")
        width_inches = figure.get("width_inches", 5.8)

        if title:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(title)
            set_run_font(r, size=11, bold=True)

        image_stream = fetch_binary(figure["url"])
        doc.add_picture(image_stream, width=Inches(width_inches))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        if caption:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(caption)
            set_run_font(r, size=9.5, italic=True, color="555555")

    except Exception as e:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(f"[No se pudo insertar la figura desde URL: {str(e)}]")
        set_run_font(r, size=9.5, italic=True, color="AA0000")


def add_quickchart(doc: Document, chart: dict):
    try:
        title = chart.get("title")
        caption = chart.get("caption")
        width_inches = chart.get("width_inches", 5.8)
        chart_config = chart["chart_config"]

        if title:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(title)
            set_run_font(r, size=11, bold=True)

        response = requests.post(
            "https://quickchart.io/chart",
            json={
                "chart": chart_config,
                "format": "png",
                "width": 1000,
                "height": 550,
                "backgroundColor": "white"
            },
            timeout=30,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        response.raise_for_status()

        image_stream = BytesIO(response.content)
        doc.add_picture(image_stream, width=Inches(width_inches))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        if caption:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(caption)
            set_run_font(r, size=9.5, italic=True, color="555555")

    except Exception as e:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(f"[No se pudo insertar el gráfico: {str(e)}]")
        set_run_font(r, size=9.5, italic=True, color="AA0000")


# =========================
# CARTA
# =========================

def build_letter_doc(letter: LetterPayload) -> BytesIO:
    doc = Document()
    configure_document(doc)

    add_logo(doc, letter.logo_url, width=0.95)

    if letter.organization:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.organization.upper())
        set_run_font(r, size=12.5, bold=True)

    if letter.city_date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.city_date)
        set_run_font(r, size=12)

    for line in [letter.recipient_name, letter.recipient_title, letter.recipient_organization]:
        if line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(line)
            set_run_font(r, size=12)

    if any([letter.recipient_name, letter.recipient_title, letter.recipient_organization]):
        doc.add_paragraph()

    if letter.subject:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r1 = p.add_run("Asunto: ")
        set_run_font(r1, size=12, bold=True)
        r2 = p.add_run(letter.subject)
        set_run_font(r2, size=12)

    if letter.greeting:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.greeting)
        set_run_font(r, size=12)

    for paragraph in letter.body:
        add_rich_text_paragraph(doc, paragraph)

    if letter.closing:
        add_rich_text_paragraph(doc, letter.closing)

    if letter.signature_name:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(20)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.signature_name)
        set_run_font(r, size=12, bold=True)

    if letter.signature_title:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.signature_title)
        set_run_font(r, size=12)

    if letter.annexes:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("Anexos:")
        set_run_font(r, size=12, bold=True)
        add_bullet_list(doc, letter.annexes)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =========================
# INFORME
# =========================

def build_report_doc(report: ReportPayload) -> BytesIO:
    doc = Document()
    configure_document(doc)

    # Portada sin header/footer
    add_report_cover(doc, report.model_dump())

    # Desde aquí, nueva sección para contenido con header/footer
    section = doc.sections[-1]
    add_header_footer(section, report.header_text, report.footer_text)

    if report.executive_summary:
        doc.add_heading("Resumen Ejecutivo", level=1)
        for paragraph in report.executive_summary:
            add_rich_text_paragraph(doc, paragraph)

    if report.sections:
        for section_item in report.sections:
            doc.add_heading(section_item.heading, level=section_item.level)

            if section_item.paragraphs:
                for paragraph in section_item.paragraphs:
                    add_rich_text_paragraph(doc, paragraph)

            if section_item.bullets:
                add_bullet_list(doc, section_item.bullets)

            if section_item.numbered:
                add_numbered_list(doc, section_item.numbered)

            if section_item.tables:
                for table in section_item.tables:
                    add_apa_table(doc, table.model_dump())

            if section_item.figures:
                for fig in section_item.figures:
                    add_figure_from_url(doc, fig.model_dump())

            if section_item.charts:
                for chart in section_item.charts:
                    add_quickchart(doc, chart.model_dump())

    if report.conclusions:
        doc.add_heading("Conclusiones", level=1)
        add_numbered_list(doc, report.conclusions)

    if report.recommendations:
        doc.add_heading("Recomendaciones", level=1)
        add_numbered_list(doc, report.recommendations)

    if report.references:
        doc.add_heading("Referencias", level=1)
        for ref in report.references:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.left_indent = Inches(0.3)
            p.paragraph_format.first_line_indent = Inches(-0.3)
            r = p.add_run(ref)
            set_run_font(r, size=12)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =========================
# ENDPOINTS
# =========================

@app.get("/")
def root():
    return {
        "message": "GPT Doc Backend activo",
        "docs": "/docs",
        "health": "/health",
        "main_endpoint": "/generate-document",
        "allowed_document_types": ["carta", "informe"]
    }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate-document")
def generate_document(payload: GenerateDocumentRequest):
    try:
        if payload.document_type == "carta":
            if not payload.letter:
                raise HTTPException(status_code=400, detail="Falta el objeto 'letter' para generar la carta.")
            file_stream = build_letter_doc(payload.letter)
            filename = f"{payload.letter.filename}.docx"

        elif payload.document_type == "informe":
            if not payload.report:
                raise HTTPException(status_code=400, detail="Falta el objeto 'report' para generar el informe.")
            file_stream = build_report_doc(payload.report)
            filename = f"{payload.report.filename}.docx"

        else:
            raise HTTPException(status_code=400, detail="Tipo de documento no permitido.")

        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": "No se pudo generar el documento.", "detail": str(e)}
        )
