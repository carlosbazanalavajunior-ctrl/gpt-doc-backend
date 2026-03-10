from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional, Literal, Union
from io import BytesIO
from pathlib import Path
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
    version="0.9.0",
    description="Backend para generación de cartas e informes profesionales usando plantilla DOCX."
)

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = BASE_DIR / "templates"
REPORT_TEMPLATE_PATH = TEMPLATE_DIR / "professional_report_template.docx"


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


# =========================
# HELPERS DE PLANTILLA
# =========================

def replace_placeholder_in_paragraph(paragraph, placeholder: str, value: str):
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder not in full_text:
        return False

    new_text = full_text.replace(placeholder, value)

    for run in paragraph.runs:
        run.text = ""

    if paragraph.runs:
        paragraph.runs[0].text = new_text
    else:
        paragraph.add_run(new_text)

    return True


def replace_placeholder_everywhere(doc: Document, placeholder: str, value: str):
    for paragraph in doc.paragraphs:
        replace_placeholder_in_paragraph(paragraph, placeholder, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_placeholder_in_paragraph(paragraph, placeholder, value)


def find_paragraph_with_placeholder(doc: Document, placeholder: str):
    for paragraph in doc.paragraphs:
        text = "".join(run.text for run in paragraph.runs)
        if placeholder in text:
            return paragraph
    return None


def clear_paragraph(paragraph):
    for run in paragraph.runs:
        run.text = ""


def insert_paragraph_after(paragraph, text="", style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = paragraph._parent.add_paragraph()
    new_para._p = new_p
    if style:
        new_para.style = style
    if text:
        new_para.add_run(text)
    return new_para


# =========================
# ESTILO DE CONTENIDO
# =========================

def add_rich_text_after(anchor_paragraph, text: str, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=True):
    p = insert_paragraph_after(anchor_paragraph)
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


def add_heading_after(anchor_paragraph, text: str, level: int):
    style_name = f"Heading {min(level, 4)}"
    p = insert_paragraph_after(anchor_paragraph, style=style_name)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0)
    run = p.add_run(text)
    return p


def add_bullet_after(anchor_paragraph, text: str):
    p = insert_paragraph_after(anchor_paragraph)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.1
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.16)

    bullet_run = p.add_run("• ")
    set_run_font(bullet_run, size=10.5, bold=True, color="404040")

    parts = text.split("**")
    for i, part in enumerate(parts):
        run = p.add_run(part)
        set_run_font(run, size=10.5, bold=(i % 2 == 1))
    return p


def add_numbered_after(anchor_paragraph, idx: int, text: str):
    p = insert_paragraph_after(anchor_paragraph)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.1
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Inches(0.27)
    p.paragraph_format.first_line_indent = Inches(-0.19)

    prefix = p.add_run(f"{idx}. ")
    set_run_font(prefix, size=10.5, bold=True, color="404040")

    parts = text.split("**")
    for i, part in enumerate(parts):
        run = p.add_run(part)
        set_run_font(run, size=10.5, bold=(i % 2 == 1))
    return p


def add_caption_after(anchor_paragraph, label: str, number: int, caption_text: str):
    p = insert_paragraph_after(anchor_paragraph)
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

def add_table_after(anchor_paragraph, table_data: dict, table_number: int, doc: Document):
    title = (table_data.get("title") or "").strip()
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])
    note = (table_data.get("note") or "").strip()

    current = anchor_paragraph

    if title:
        current = add_caption_after(current, "Tabla", table_number, title)

    ncols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if ncols == 0:
        return current

    tbl_p = OxmlElement("w:p")
    current._p.addnext(tbl_p)

    table = doc.add_table(rows=1, cols=ncols)
    tbl_xml = table._tbl
    current._p.addnext(tbl_xml)

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

    after_para = insert_paragraph_after(current)
    after_para.paragraph_format.first_line_indent = Inches(0)
    after_para.paragraph_format.space_before = Pt(1)
    after_para.paragraph_format.space_after = Pt(0)
    current = after_para

    if note:
        current = add_rich_text_after(current, note if note.lower().startswith("Nota.") else f"Nota. {note}", align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=False)
        if current.runs:
            for run in current.runs:
                set_run_font(run, size=9, italic=True, color="555555")

    return current


def add_figure_after(anchor_paragraph, figure: dict, figure_number: int, doc: Document):
    title = figure.get("title")
    caption = figure.get("caption")
    width_inches = figure.get("width_inches", 5.8)
    current = anchor_paragraph

    if title:
        p = insert_paragraph_after(current)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(title)
        set_run_font(r, size=10.5, bold=True)
        current = p

    try:
        image_stream = safe_fetch_image(figure["url"], max_width=2600)

        img_p = insert_paragraph_after(current)
        img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img_p.paragraph_format.first_line_indent = Inches(0)
        img_p.add_run().add_picture(image_stream, width=Inches(width_inches))
        current = img_p
    except Exception:
        p = insert_paragraph_after(current)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("[Imagen no disponible en esta prueba]")
        set_run_font(r, size=9.5, italic=True, color="777777")
        current = p

    if caption:
        current = add_caption_after(current, "Figura", figure_number, caption)

    return current


def add_chart_after(anchor_paragraph, chart: dict, figure_number: int):
    title = chart.get("title")
    caption = chart.get("caption")
    width_inches = chart.get("width_inches", 5.8)
    chart_config = chart["chart_config"]
    current = anchor_paragraph

    if title:
        p = insert_paragraph_after(current)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(title)
        set_run_font(r, size=10.5, bold=True)
        current = p

    try:
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

        img_p = insert_paragraph_after(current)
        img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img_p.paragraph_format.first_line_indent = Inches(0)
        img_p.add_run().add_picture(image_stream, width=Inches(width_inches))
        current = img_p
    except Exception:
        p = insert_paragraph_after(current)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run("[Gráfico no disponible en esta prueba]")
        set_run_font(r, size=9.5, italic=True, color="777777")
        current = p

    if caption:
        current = add_caption_after(current, "Figura", figure_number, caption)

    return current


# =========================
# ÍNDICES GENERADOS
# =========================

def build_indexes(report: ReportPayload):
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
                    title = (table.title or f"Tabla {table_counter}").strip()
                    table_index_entries.append(f"Tabla {table_counter}. {title}")

            if section_item.figures:
                for fig in section_item.figures:
                    figure_counter += 1
                    caption = (fig.caption or fig.title or f"Figura {figure_counter}").strip()
                    figure_index_entries.append(f"Figura {figure_counter}. {caption}")

            if section_item.charts:
                for chart in section_item.charts:
                    figure_counter += 1
                    caption = (chart.caption or chart.title or f"Figura {figure_counter}").strip()
                    figure_index_entries.append(f"Figura {figure_counter}. {caption}")

    if report.conclusions:
        toc_entries.append("Conclusiones")
    if report.recommendations:
        toc_entries.append("Recomendaciones")
    if report.references:
        toc_entries.append("Referencias")

    return toc_entries, table_index_entries, figure_index_entries


# =========================
# INFORME CON PLANTILLA
# =========================

def build_report_doc(report: ReportPayload) -> BytesIO:
    if not REPORT_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"No se encontró la plantilla: {REPORT_TEMPLATE_PATH}. "
            "Cree el archivo templates/professional_report_template.docx"
        )

    doc = Document(str(REPORT_TEMPLATE_PATH))

    # Reemplazos directos de portada
    replace_placeholder_everywhere(doc, "{{INSTITUTION}}", report.institution or "")
    replace_placeholder_everywhere(doc, "{{FACULTY}}", report.faculty or "")
    replace_placeholder_everywhere(doc, "{{DEPARTMENT}}", report.department or "")
    replace_placeholder_everywhere(doc, "{{REPORT_KIND}}", report.report_kind or "")
    replace_placeholder_everywhere(doc, "{{TITLE}}", report.title or "")
    replace_placeholder_everywhere(doc, "{{SUBTITLE}}", report.subtitle or "")
    replace_placeholder_everywhere(doc, "{{AUTHOR}}", report.author or "")
    replace_placeholder_everywhere(doc, "{{REVIEWER}}", report.reviewer or "")
    city_date = f"{report.city}, {report.date}" if report.city and report.date else (report.city or report.date or "")
    replace_placeholder_everywhere(doc, "{{CITY_DATE}}", city_date)

    # Índices
    toc_entries, table_index_entries, figure_index_entries = build_indexes(report)

    replace_placeholder_everywhere(doc, "{{TOC}}", "\n".join(toc_entries) if toc_entries else "No se registraron secciones.")
    replace_placeholder_everywhere(doc, "{{TABLE_INDEX}}", "\n".join(table_index_entries) if table_index_entries else "No se registraron tablas.")
    replace_placeholder_everywhere(doc, "{{FIGURE_INDEX}}", "\n".join(figure_index_entries) if figure_index_entries else "No se registraron figuras.")

    # Header / Footer
    for section in doc.sections:
        if report.header_text:
            hp = section.header.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hp.text = ""
            run = hp.add_run(report.header_text)
            set_run_font(run, size=8, italic=True, color="6F6F6F")

        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.text = ""
        if report.footer_text:
            left = fp.add_run(report.footer_text)
            set_run_font(left, size=8, color="6F6F6F")
            sep = fp.add_run("  |  ")
            set_run_font(sep, size=8, color="A5A5A5")
        page_label = fp.add_run("Página ")
        set_run_font(page_label, size=8, color="6F6F6F")
        add_page_number(fp)

    body_anchor = find_paragraph_with_placeholder(doc, "{{BODY_CONTENT}}")
    if body_anchor is None:
        raise ValueError("La plantilla no contiene el marcador {{BODY_CONTENT}}")

    clear_paragraph(body_anchor)
    current = body_anchor

    # Resumen ejecutivo
    if report.executive_summary:
        current = add_heading_after(current, "Resumen Ejecutivo", 1)
        for paragraph in report.executive_summary:
            current = add_rich_text_after(current, paragraph)

    # Contenido principal
    table_counter = 0
    figure_counter = 0

    if report.sections:
        for section_item in report.sections:
            current = add_heading_after(current, section_item.heading, section_item.level)

            if section_item.paragraphs:
                for paragraph in section_item.paragraphs:
                    current = add_rich_text_after(current, paragraph)

            if section_item.bullets:
                for item in section_item.bullets:
                    current = add_bullet_after(current, item)

            if section_item.numbered:
                for idx, item in enumerate(section_item.numbered, start=1):
                    current = add_numbered_after(current, idx, item)

            if section_item.tables:
                for table in section_item.tables:
                    table_counter += 1
                    current = add_table_after(current, table.model_dump(), table_counter, doc)

            if section_item.figures:
                for fig in section_item.figures:
                    figure_counter += 1
                    current = add_figure_after(current, fig.model_dump(), figure_counter, doc)

            if section_item.charts:
                for chart in section_item.charts:
                    figure_counter += 1
                    current = add_chart_after(current, chart.model_dump(), figure_counter)

    if report.conclusions:
        current = add_heading_after(current, "Conclusiones", 1)
        for idx, item in enumerate(report.conclusions, start=1):
            current = add_numbered_after(current, idx, item)

    if report.recommendations:
        current = add_heading_after(current, "Recomendaciones", 1)
        for idx, item in enumerate(report.recommendations, start=1):
            current = add_numbered_after(current, idx, item)

    if report.references:
        current = add_heading_after(current, "Referencias", 1)
        for ref in report.references:
            p = insert_paragraph_after(current)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.left_indent = Inches(0.34)
            p.paragraph_format.first_line_indent = Inches(-0.34)
            r = p.add_run(ref)
            set_run_font(r, size=10.5)
            current = p

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# =========================
# CARTA SIMPLE
# =========================

def build_letter_doc(letter: LetterPayload) -> BytesIO:
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(0.9)

    if letter.organization:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(letter.organization.upper())
        set_run_font(r, size=11.5, bold=True)

    if letter.city_date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(letter.city_date)
        set_run_font(r, size=11)

    for line in [letter.recipient_name, letter.recipient_title, letter.recipient_organization]:
        if line:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(line)
            set_run_font(r, size=11)

    if letter.subject:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p.add_run("Asunto: ")
        set_run_font(r1, size=11, bold=True)
        r2 = p.add_run(letter.subject)
        set_run_font(r2, size=11)

    if letter.greeting:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(letter.greeting)
        set_run_font(r, size=11)

    for paragraph in letter.body:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.first_line_indent = Inches(0.22)
        parts = paragraph.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            set_run_font(run, size=11, bold=(i % 2 == 1))

    if letter.closing:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.first_line_indent = Inches(0.22)
        r = p.add_run(letter.closing)
        set_run_font(r, size=11)

    if letter.signature_name:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        r = p.add_run(letter.signature_name)
        set_run_font(r, size=11, bold=True)

    if letter.signature_title:
        p = doc.add_paragraph()
        r = p.add_run(letter.signature_title)
        set_run_font(r, size=11)

    if letter.annexes:
        p = doc.add_paragraph()
        r = p.add_run("Anexos:")
        set_run_font(r, size=11, bold=True)

        for annex in letter.annexes:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.16)
            b = p.add_run("• ")
            set_run_font(b, size=10.5, bold=True)
            r = p.add_run(annex)
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
