from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Literal, Union
from io import BytesIO
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import requests


app = FastAPI(
    title="GPT Doc Backend",
    version="0.8.0",
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
# HELPERS BASE
# =========================

def set_run_font(run, name="Times New Roman", size=11, bold=False, italic=False, color=None):
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
        timeout=30,
        headers={"User-Agent": "Mozilla/5.0"}
    )
    response.raise_for_status()
    return BytesIO(response.content)


def normalize_image_stream(stream: BytesIO, max_width=2200) -> BytesIO:
    stream.seek(0)
    img = Image.open(stream)
    img = img.convert("RGB")

    if img.width > max_width:
        ratio = max_width / float(img.width)
        new_size = (int(img.width * ratio), int(img.height * ratio))
        img = img.resize(new_size, Image.LANCZOS)

    out = BytesIO()
    img.save(out, format="PNG", optimize=True)
    out.seek(0)
    return out


def safe_fetch_image(url: str, max_width=2200) -> BytesIO:
    stream = fetch_binary(url)
    return normalize_image_stream(stream, max_width=max_width)


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


def shade_cell(cell, fill="F5F5F5"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


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


def add_horizontal_rule(paragraph, color="C8C8C8", size="5"):
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
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(0.9)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(11)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.first_line_indent = Inches(0.22)

    styles = [
        ("Heading 1", 13, False, Pt(10), Pt(4)),
        ("Heading 2", 12, False, Pt(8), Pt(4)),
        ("Heading 3", 11.5, False, Pt(7), Pt(3)),
        ("Heading 4", 11, True, Pt(6), Pt(3)),
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
        style.paragraph_format.line_spacing = 1.05
        style.paragraph_format.first_line_indent = Inches(0)


def add_logo(doc: Document, logo_url: Optional[str], width=0.85):
    if not logo_url:
        return
    try:
        stream = safe_fetch_image(logo_url)
        doc.add_picture(stream, width=Inches(width))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass


def add_rich_text_paragraph(doc: Document, text: str, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=True):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.first_line_indent = Inches(0.22) if first_indent else Inches(0)

    parts = text.split("**")
    for i, part in enumerate(parts):
        run = p.add_run(part)
        set_run_font(run, size=11, bold=(i % 2 == 1))
    return p


def add_bullet_list(doc: Document, items: List[str]):
    for item in items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.1
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.16)

        bullet_run = p.add_run("• ")
        set_run_font(bullet_run, size=10.5, bold=True, color="404040")

        parts = item.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_run_font(run, size=10.5, bold=(i % 2 == 1))


def add_numbered_list(doc: Document, items: List[str]):
    for idx, item in enumerate(items, start=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.1
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.left_indent = Inches(0.27)
        p.paragraph_format.first_line_indent = Inches(-0.19)

        prefix = p.add_run(f"{idx}. ")
        set_run_font(prefix, size=10.5, bold=True, color="404040")

        parts = item.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_run_font(run, size=10.5, bold=(i % 2 == 1))


# =========================
# HEADER / FOOTER
# =========================

def add_header_footer(section, header_text: Optional[str], footer_text: Optional[str]):
    if header_text:
        hp = section.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.text = ""
        run = hp.add_run(header_text)
        set_run_font(run, size=8, italic=True, color="6F6F6F")

    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.text = ""

    if footer_text:
        left = fp.add_run(footer_text)
        set_run_font(left, size=8, color="6F6F6F")
        sep = fp.add_run("  |  ")
        set_run_font(sep, size=8, color="A5A5A5")

    page_label = fp.add_run("Página ")
    set_run_font(page_label, size=8, color="6F6F6F")
    add_page_number(fp)


# =========================
# PORTADA Y PRELIMINARES
# =========================

def add_report_cover(doc: Document, report: dict):
    add_logo(doc, report.get("logo_url"), width=0.82)

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
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(institution.upper())
        set_run_font(r, size=11, bold=True)

    if faculty:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(faculty)
        set_run_font(r, size=10.5, bold=True)

    if department:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(5)
        r = p.add_run(department)
        set_run_font(r, size=9.5, color="555555")

    sep = doc.add_paragraph()
    sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sep.paragraph_format.space_after = Pt(4)
    add_horizontal_rule(sep, color="C6C6C6", size="5")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(report_kind.upper())
    set_run_font(r, size=11, bold=True, color="404040")

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(title)
        set_run_font(r, size=14, bold=True)

    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(7)
        r = p.add_run(subtitle)
        set_run_font(r, size=10, italic=True, color="555555")

    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r1 = p.add_run("Autor: ")
        set_run_font(r1, size=10, bold=True)
        r2 = p.add_run(author)
        set_run_font(r2, size=10)

    if reviewer:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r1 = p.add_run("Revisor / Asesor: ")
        set_run_font(r1, size=10, bold=True)
        r2 = p.add_run(reviewer)
        set_run_font(r2, size=10)

    if city or date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        text = f"{city}, {date}" if city and date else city or date
        r = p.add_run(text)
        set_run_font(r, size=9.5, color="555555")

    doc.add_page_break()


def add_generated_index_page(doc: Document, title: str, entries: List[str]):
    doc.add_heading(title, level=1)

    if not entries:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("No se registraron elementos en esta sección.")
        set_run_font(r, size=10, italic=True, color="666666")
    else:
        for entry in entries:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(entry)
            set_run_font(r, size=10.2)

    doc.add_page_break()


# =========================
# CAPTIONS Y REGISTRO
# =========================

def add_caption(doc: Document, label: str, number: int, caption_text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0)

    r1 = p.add_run(f"{label} {number}. ")
    set_run_font(r1, size=9.5, bold=True)

    r2 = p.add_run(caption_text)
    set_run_font(r2, size=9.5, italic=True, color="555555")

    return p


# =========================
# TABLAS / FIGURAS / GRÁFICOS
# =========================

def add_apa_table(doc: Document, table_data: dict, table_number: int):
    title = (table_data.get("title") or "").strip()
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])
    note = (table_data.get("note") or "").strip()

    if title:
        add_caption(doc, "Tabla", table_number, title)

    ncols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if ncols == 0:
        return

    table = doc.add_table(rows=1, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        shade_cell(cell, "F0F0F0")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)

        r = p.add_run(str(header))
        set_run_font(r, size=10, bold=True)

        set_cell_border(
            cell,
            top={"sz": 8, "val": "single", "color": "808080"},
            bottom={"sz": 6, "val": "single", "color": "A8A8A8"},
            left={"val": "nil"},
            right={"val": "nil"},
        )

    for row_idx, row in enumerate(rows):
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            cell = row_cells[i]
            cell.text = ""
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            if row_idx % 2 == 1:
                shade_cell(cell, "FAFAFA")

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            r = p.add_run(str(value))
            set_run_font(r, size=10)

            set_cell_border(
                cell,
                top={"val": "nil"},
                bottom={"sz": 3, "val": "single", "color": "DDDDDD"},
                left={"val": "nil"},
                right={"val": "nil"},
            )

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(1)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.first_line_indent = Inches(0)

    if note:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0)
        note_text = note if note.lower().startswith("nota.") else f"Nota. {note}"
        r = p.add_run(note_text)
        set_run_font(r, size=9, italic=True, color="555555")


def add_figure_from_url(doc: Document, figure: dict, figure_number: int):
    try:
        title = figure.get("title")
        caption = figure.get("caption")
        width_inches = figure.get("width_inches", 5.8)

        if title:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(title)
            set_run_font(r, size=10.5, bold=True)

        image_stream = safe_fetch_image(figure["url"], max_width=2600)
        doc.add_picture(image_stream, width=Inches(width_inches))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        if caption:
            add_caption(doc, "Figura", figure_number, caption)

    except Exception:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("[Imagen no disponible en esta prueba]")
        set_run_font(r, size=9.5, italic=True, color="777777")


def add_quickchart(doc: Document, chart: dict, figure_number: int):
    try:
        title = chart.get("title")
        caption = chart.get("caption")
        width_inches = chart.get("width_inches", 5.8)
        chart_config = chart["chart_config"]

        if title:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(title)
            set_run_font(r, size=10.5, bold=True)

        response = requests.post(
            "https://quickchart.io/chart",
            json={
                "chart": chart_config,
                "format": "png",
                "width": 1600,
                "height": 900,
                "devicePixelRatio": 2,
                "backgroundColor": "white"
            },
            timeout=30,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        response.raise_for_status()

        image_stream = BytesIO(response.content)
        image_stream = normalize_image_stream(image_stream, max_width=2600)

        doc.add_picture(image_stream, width=Inches(width_inches))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        if caption:
            add_caption(doc, "Figura", figure_number, caption)

    except Exception:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("[Gráfico no disponible en esta prueba]")
        set_run_font(r, size=9.5, italic=True, color="777777")


# =========================
# CARTA
# =========================

def build_letter_doc(letter: LetterPayload) -> BytesIO:
    doc = Document()
    configure_document(doc)

    add_logo(doc, letter.logo_url, width=0.82)

    if letter.organization:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.organization.upper())
        set_run_font(r, size=11.5, bold=True)

    if letter.city_date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.city_date)
        set_run_font(r, size=11)

    for line in [letter.recipient_name, letter.recipient_title, letter.recipient_organization]:
        if line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(line)
            set_run_font(r, size=11)

    if any([letter.recipient_name, letter.recipient_title, letter.recipient_organization]):
        doc.add_paragraph()

    if letter.subject:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.first_line_indent = Inches(0)
        r1 = p.add_run("Asunto: ")
        set_run_font(r1, size=11, bold=True)
        r2 = p.add_run(letter.subject)
        set_run_font(r2, size=11)

    if letter.greeting:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.greeting)
        set_run_font(r, size=11)

    for paragraph in letter.body:
        add_rich_text_paragraph(doc, paragraph)

    if letter.closing:
        add_rich_text_paragraph(doc, letter.closing)

    if letter.signature_name:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.signature_name)
        set_run_font(r, size=11, bold=True)

    if letter.signature_title:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.signature_title)
        set_run_font(r, size=11)

    if letter.annexes:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("Anexos:")
        set_run_font(r, size=11, bold=True)
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

    # Recolectar índices previos de backend
    toc_entries: List[str] = []
    table_index_entries: List[str] = []
    figure_index_entries: List[str] = []

    if report.executive_summary:
        toc_entries.append("Resumen Ejecutivo")

    figure_counter = 0
    table_counter = 0

    if report.sections:
        for section_item in report.sections:
            indent = "    " * max(0, section_item.level - 1)
            toc_entries.append(f"{indent}{section_item.heading}")

            if section_item.tables:
                for table in section_item.tables:
                    table_counter += 1
                    table_title = (table.title or f"Tabla sin título {table_counter}").strip()
                    table_index_entries.append(f"Tabla {table_counter}. {table_title}")

            if section_item.figures:
                for fig in section_item.figures:
                    figure_counter += 1
                    caption = (fig.caption or fig.title or f"Figura sin título {figure_counter}").strip()
                    figure_index_entries.append(f"Figura {figure_counter}. {caption}")

            if section_item.charts:
                for chart in section_item.charts:
                    figure_counter += 1
                    caption = (chart.caption or chart.title or f"Gráfico sin título {figure_counter}").strip()
                    figure_index_entries.append(f"Figura {figure_counter}. {caption}")

    if report.conclusions:
        toc_entries.append("Conclusiones")
    if report.recommendations:
        toc_entries.append("Recomendaciones")
    if report.references:
        toc_entries.append("Referencias")

    # Portada
    add_report_cover(doc, report.model_dump())

    # Índices generados por backend
    add_generated_index_page(doc, "Índice general", toc_entries)
    add_generated_index_page(doc, "Índice de tablas", table_index_entries)
    add_generated_index_page(doc, "Índice de figuras", figure_index_entries)

    section = doc.sections[-1]
    add_header_footer(section, report.header_text, report.footer_text)

    # Reset contadores para render final
    figure_counter = 0
    table_counter = 0

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
                    table_counter += 1
                    add_apa_table(doc, table.model_dump(), table_counter)

            if section_item.figures:
                for fig in section_item.figures:
                    figure_counter += 1
                    add_figure_from_url(doc, fig.model_dump(), figure_counter)

            if section_item.charts:
                for chart in section_item.charts:
                    figure_counter += 1
                    add_quickchart(doc, chart.model_dump(), figure_counter)

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
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.left_indent = Inches(0.34)
            p.paragraph_format.first_line_indent = Inches(-0.34)

            r = p.add_run(ref)
            set_run_font(r, size=10.5)

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
            headers={"Content-Disposition": f'attachment; filename=\"{filename}\"'}
        )

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": "No se pudo generar el documento.", "detail": str(e)}
        )
