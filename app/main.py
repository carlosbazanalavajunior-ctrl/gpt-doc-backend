from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Literal, Union
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
import requests


app = FastAPI(
    title="GPT Doc Backend",
    version="0.3.0",
    description="Backend para generación de cartas e informes profesionales en formato Word (.docx)."
)


# =========================
# MODELOS Pydantic
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
# HELPERS VISUALES DOCX
# =========================

def set_cell_border(cell, **kwargs):
    """
    Permite controlar bordes de una celda.
    Ejemplo:
    set_cell_border(cell, top={"sz": 8, "val": "single", "color": "000000"})
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        from docx.oxml import OxmlElement
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("left", "top", "right", "bottom", "insideH", "insideV"):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            tag = f"w:{edge}"

            element = tcBorders.find(qn(tag))
            if element is None:
                from docx.oxml import OxmlElement
                element = OxmlElement(tag)
                tcBorders.append(element)

            for key in ["val", "sz", "space", "color"]:
                if key in edge_data:
                    element.set(qn(f"w:{key}"), str(edge_data[key]))


def set_paragraph_font(run, name="Times New Roman", size=12, bold=False, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def configure_document(doc: Document):
    section = doc.sections[0]

    # Márgenes refinados
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.1)
    section.right_margin = Inches(1.0)

    # Estilo base
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal_style.font.size = Pt(12)

    pf = normal_style.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(8)
    pf.space_before = Pt(0)
    pf.first_line_indent = Inches(0.3)

    # Heading 1
    h1 = doc.styles["Heading 1"]
    h1.font.name = "Times New Roman"
    h1._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.line_spacing = 1.15
    h1.paragraph_format.first_line_indent = Inches(0)

    # Heading 2
    h2 = doc.styles["Heading 2"]
    h2.font.name = "Times New Roman"
    h2._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.line_spacing = 1.15
    h2.paragraph_format.first_line_indent = Inches(0)

    # Heading 3
    h3 = doc.styles["Heading 3"]
    h3.font.name = "Times New Roman"
    h3._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    h3.font.size = Pt(12)
    h3.font.bold = True
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(4)
    h3.paragraph_format.line_spacing = 1.15
    h3.paragraph_format.first_line_indent = Inches(0)

    # Heading 4
    h4 = doc.styles["Heading 4"]
    h4.font.name = "Times New Roman"
    h4._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    h4.font.size = Pt(12)
    h4.font.bold = True
    h4.font.italic = True
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    h4.paragraph_format.line_spacing = 1.15
    h4.paragraph_format.first_line_indent = Inches(0)


def add_header_footer(doc: Document, header_text: Optional[str], footer_text: Optional[str]):
    for section in doc.sections:
        if header_text:
            header = section.header
            p = header.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.text = ""
            run = p.add_run(header_text)
            set_paragraph_font(run, size=10)

        if footer_text:
            footer = section.footer
            p = footer.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.text = ""
            run = p.add_run(footer_text)
            set_paragraph_font(run, size=9)


def add_paragraph_block(doc: Document, text: str, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align

    fmt = p.paragraph_format
    fmt.line_spacing = 1.5
    fmt.space_after = Pt(8)
    fmt.space_before = Pt(0)
    fmt.first_line_indent = Inches(0.3)

    if not text:
        return p

    parts = text.split("**")

    for i, part in enumerate(parts):
        run = p.add_run(part)
        set_paragraph_font(run, size=12, bold=(i % 2 == 1))

    return p


def add_bullet_list(doc: Document, items: List[str]):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.left_indent = Inches(0.3)

        parts = item.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_paragraph_font(run, size=12, bold=(i % 2 == 1))


def add_numbered_list(doc: Document, items: List[str]):
    for item in items:
        p = doc.add_paragraph(style="List Number")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.left_indent = Inches(0.3)

        parts = item.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_paragraph_font(run, size=12, bold=(i % 2 == 1))


def add_report_cover(doc: Document, report: dict):
    doc.add_paragraph()

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
        r = p.add_run(institution.upper())
        set_paragraph_font(r, size=13, bold=True)
        p.paragraph_format.space_after = Pt(6)

    if faculty:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(faculty)
        set_paragraph_font(r, size=12, bold=True)
        p.paragraph_format.space_after = Pt(4)

    if department:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(department)
        set_paragraph_font(r, size=11)
        p.paragraph_format.space_after = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(report_kind.upper())
    set_paragraph_font(r, size=14, bold=True)
    p.paragraph_format.space_after = Pt(20)

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        set_paragraph_font(r, size=16, bold=True)
        p.paragraph_format.space_after = Pt(10)

    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(subtitle)
        set_paragraph_font(r, size=12, italic=True)
        p.paragraph_format.space_after = Pt(28)

    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p.add_run("Autor: ")
        set_paragraph_font(r1, size=12, bold=True)
        r2 = p.add_run(author)
        set_paragraph_font(r2, size=12)
        p.paragraph_format.space_after = Pt(6)

    if reviewer:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p.add_run("Revisor / Asesor: ")
        set_paragraph_font(r1, size=12, bold=True)
        r2 = p.add_run(reviewer)
        set_paragraph_font(r2, size=12)
        p.paragraph_format.space_after = Pt(6)

    if city or date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        text = f"{city}, {date}" if city and date else city or date
        r = p.add_run(text)
        set_paragraph_font(r, size=12)

    doc.add_page_break()


def add_apa_table(doc: Document, table_data: dict):
    title = (table_data.get("title") or "").strip()
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])
    note = (table_data.get("note") or "").strip()

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0)

        r = p.add_run(title)
        set_paragraph_font(r, size=12, italic=True)

    ncols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if ncols == 0:
        return

    table = doc.add_table(rows=1, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    # Encabezado
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = hdr_cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.first_line_indent = Inches(0)

        r = p.add_run(str(header))
        set_paragraph_font(r, size=11, bold=True)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        set_cell_border(
            cell,
            top={"sz": 10, "val": "single", "color": "000000"},
            bottom={"sz": 10, "val": "single", "color": "000000"},
            left={"sz": 6, "val": "single", "color": "000000"},
            right={"sz": 6, "val": "single", "color": "000000"},
        )

    # Filas
    for row in rows:
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            cell = row_cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)

            r = p.add_run(str(value))
            set_paragraph_font(r, size=11)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            set_cell_border(
                cell,
                top={"sz": 6, "val": "single", "color": "000000"},
                bottom={"sz": 6, "val": "single", "color": "000000"},
                left={"sz": 6, "val": "single", "color": "000000"},
                right={"sz": 6, "val": "single", "color": "000000"},
            )

    doc.add_paragraph()

    if note:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.first_line_indent = Inches(0)

        note_text = note if note.lower().startswith("nota.") else f"Nota. {note}"
        r = p.add_run(note_text)
        set_paragraph_font(r, size=10)


def download_image(url: str) -> BytesIO:
    response = requests.get(url, timeout=20)
    response.raise_for_status()
    return BytesIO(response.content)


def add_figure_from_url(doc: Document, figure: dict):
    try:
        title = figure.get("title")
        caption = figure.get("caption")
        width_inches = figure.get("width_inches", 5.8)

        if title:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(title)
            set_paragraph_font(r, size=11, bold=True)

        image_stream = download_image(figure["url"])
        doc.add_picture(image_stream, width=Inches(width_inches))

        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if caption:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(caption)
            set_paragraph_font(r, size=10, italic=True)

    except Exception as e:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(f"[No se pudo insertar la figura desde URL: {str(e)}]")
        set_paragraph_font(r, size=10, italic=True)


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
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(title)
            set_paragraph_font(r, size=11, bold=True)

        response = requests.post(
            "https://quickchart.io/chart",
            json={"chart": chart_config, "format": "png", "width": 900, "height": 500},
            timeout=25
        )
        response.raise_for_status()

        image_stream = BytesIO(response.content)
        doc.add_picture(image_stream, width=Inches(width_inches))

        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if caption:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(caption)
            set_paragraph_font(r, size=10, italic=True)

    except Exception as e:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(f"[No se pudo insertar el gráfico: {str(e)}]")
        set_paragraph_font(r, size=10, italic=True)


# =========================
# GENERACIÓN DE CARTA
# =========================

def build_letter_doc(letter: LetterPayload) -> BytesIO:
    doc = Document()
    configure_document(doc)

    if letter.organization:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.organization.upper())
        set_paragraph_font(r, size=13, bold=True)

    if letter.city_date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(14)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.city_date)
        set_paragraph_font(r, size=12)

    recipient_lines = [
        letter.recipient_name,
        letter.recipient_title,
        letter.recipient_organization
    ]
    for line in recipient_lines:
        if line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(line)
            set_paragraph_font(r, size=12, bold=False)

    if any(recipient_lines):
        doc.add_paragraph()

    if letter.subject:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)

        r1 = p.add_run("Asunto: ")
        set_paragraph_font(r1, size=12, bold=True)

        r2 = p.add_run(letter.subject)
        set_paragraph_font(r2, size=12)

    if letter.greeting:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.greeting)
        set_paragraph_font(r, size=12)

    for paragraph in letter.body:
        add_paragraph_block(doc, paragraph)

    if letter.closing:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(18)
        p.paragraph_format.first_line_indent = Inches(0.3)
        r = p.add_run(letter.closing)
        set_paragraph_font(r, size=12)

    if letter.signature_name:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(24)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.signature_name)
        set_paragraph_font(r, size=12, bold=True)

    if letter.signature_title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(letter.signature_title)
        set_paragraph_font(r, size=12)

    if letter.annexes:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("Anexos:")
        set_paragraph_font(r, size=12, bold=True)

        for annex in letter.annexes:
            p = doc.add_paragraph(style="List Bullet")
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(annex)
            set_paragraph_font(r, size=12)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =========================
# GENERACIÓN DE INFORME
# =========================

def build_report_doc(report: ReportPayload) -> BytesIO:
    doc = Document()
    configure_document(doc)

    add_header_footer(doc, report.header_text, report.footer_text)
    add_report_cover(doc, report.model_dump())

    if report.executive_summary:
        doc.add_heading("Resumen Ejecutivo", level=1)
        for paragraph in report.executive_summary:
            add_paragraph_block(doc, paragraph)

    if report.sections:
        for section in report.sections:
            doc.add_heading(section.heading, level=section.level)

            if section.paragraphs:
                for paragraph in section.paragraphs:
                    add_paragraph_block(doc, paragraph)

            if section.bullets:
                add_bullet_list(doc, section.bullets)

            if section.numbered:
                add_numbered_list(doc, section.numbered)

            if section.tables:
                for table in section.tables:
                    add_apa_table(doc, table.model_dump())

            if section.figures:
                for fig in section.figures:
                    add_figure_from_url(doc, fig.model_dump())

            if section.charts:
                for chart in section.charts:
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
            p.paragraph_format.first_line_indent = Inches(-0.3)
            p.paragraph_format.left_indent = Inches(0.3)
            r = p.add_run(ref)
            set_paragraph_font(r, size=12)

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

        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"'
        }

        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers=headers
        )

    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": "No se pudo generar el documento.", "detail": str(e)}
        )
