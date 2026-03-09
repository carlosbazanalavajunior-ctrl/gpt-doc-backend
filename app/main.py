from io import BytesIO
from typing import Optional, List, Literal
import json
import re

import requests
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


app = FastAPI(
    title="GPT Doc Backend",
    description="Motor documental para generar cartas e informes profesionales en Word",
    version="0.3.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
)


# =========================
# MODELOS
# =========================

class TableBlock(BaseModel):
    title: str
    headers: List[str]
    rows: List[List[str]] = []
    note: Optional[str] = None
    source: Optional[str] = None


class FigureBlock(BaseModel):
    title: str
    url: str
    note: Optional[str] = None
    source: Optional[str] = None
    width_inches: float = 5.5


class ChartDataset(BaseModel):
    label: str
    data: List[float]


class ChartBlock(BaseModel):
    title: str
    chart_type: str = Field(default="bar", alias="type")
    labels: List[str]
    datasets: List[ChartDataset]
    note: Optional[str] = None
    source: Optional[str] = None
    width: int = 900
    height: int = 450


class SectionBlock(BaseModel):
    heading: str
    level: int = 1
    paragraphs: List[str] = []
    bullets: List[str] = []
    numbered: List[str] = []
    tables: List[TableBlock] = []
    figures: List[FigureBlock] = []
    charts: List[ChartBlock] = []


class LetterRequest(BaseModel):
    filename: str = "carta_formal"
    logo_url: Optional[str] = None
    organization: Optional[str] = None
    city_date: str
    recipient_name: str
    recipient_title: Optional[str] = None
    recipient_organization: Optional[str] = None
    subject: str
    greeting: str = "De mi consideración:"
    body: List[str]
    closing: str = "Atentamente,"
    signature_name: str
    signature_title: Optional[str] = None
    annexes: List[str] = []


class ReportRequest(BaseModel):
    filename: str = "informe_profesional"
    report_kind: str = "Informe"
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
    executive_summary: List[str] = []
    sections: List[SectionBlock] = []
    conclusions: List[str] = []
    recommendations: List[str] = []
    references: List[str] = []


class DocumentRequest(BaseModel):
    document_type: Literal["carta", "informe"]
    letter: Optional[LetterRequest] = None
    report: Optional[ReportRequest] = None


# =========================
# UTILIDADES GENERALES
# =========================

def sanitize_filename(name: str) -> str:
    safe = name.strip().lower().replace(" ", "_")
    safe = re.sub(r"[^a-zA-Z0-9_\-]", "", safe)
    return safe or "documento"


def configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    styles = doc.styles
    normal_style = styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)


def set_header_footer(doc: Document, header_text: Optional[str], footer_text: Optional[str]) -> None:
    section = doc.sections[0]

    if header_text:
        header = section.header
        p = header.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(header_text)
        run.font.size = Pt(9)

    if footer_text:
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(footer_text)
        run.font.size = Pt(9)


def add_text_with_bold(paragraph, text: str) -> None:
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)


def add_paragraph_block(doc: Document, text: str, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY) -> None:
    p = doc.add_paragraph()
    p.alignment = alignment
    add_text_with_bold(p, text)
    p.paragraph_format.space_after = Pt(6)


def download_image_to_buffer(url: str) -> BytesIO:
    response = requests.get(
        url,
        timeout=40,
        headers={
            "User-Agent": "Mozilla/5.0 GPTDocBackend"
        }
    )
    response.raise_for_status()
    return BytesIO(response.content)


def build_quickchart_url(chart: ChartBlock) -> str:
    chart_config = {
        "type": chart.chart_type,
        "data": {
            "labels": chart.labels,
            "datasets": [
                {
                    "label": dataset.label,
                    "data": dataset.data
                }
                for dataset in chart.datasets
            ],
        },
        "options": {
            "plugins": {
                "title": {
                    "display": True,
                    "text": chart.title
                },
                "legend": {
                    "display": True
                }
            },
            "responsive": True
        }
    }

    config_json = json.dumps(chart_config, separators=(",", ":"))
    return (
        "https://quickchart.io/chart"
        f"?c={requests.utils.quote(config_json)}"
        f"&width={chart.width}"
        f"&height={chart.height}"
        "&format=png"
        "&backgroundColor=white"
    )


# =========================
# APA 7: TABLAS Y FIGURAS
# =========================

def set_cell_border(cell, top=None, bottom=None):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)

    for edge_name, edge_data in {"top": top, "bottom": bottom}.items():
        if edge_data:
            tag = f"w:{edge_name}"
            element = tc_borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tc_borders.append(element)
            for key, value in edge_data.items():
                element.set(qn(f"w:{key}"), str(value))


def apply_apa_table_borders(table):
    total_rows = len(table.rows)
    if total_rows == 0:
        return

    for i, row in enumerate(table.rows):
        for cell in row.cells:
            if i == 0:
                set_cell_border(
                    cell,
                    top={"val": "single", "sz": 12, "color": "000000"},
                    bottom={"val": "single", "sz": 8, "color": "000000"},
                )
            elif i == total_rows - 1:
                set_cell_border(
                    cell,
                    bottom={"val": "single", "sz": 12, "color": "000000"},
                )


def add_apa_table(doc: Document, table_data: TableBlock, table_number: int) -> None:
    p_num = doc.add_paragraph()
    p_num.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    run_num = p_num.add_run(f"Tabla {table_number}")
    run_num.bold = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    run_title = p_title.add_run(table_data.title)
    run_title.italic = True

    table = doc.add_table(rows=1, cols=len(table_data.headers))

    for col_idx, header in enumerate(table_data.headers):
        cell = table.rows[0].cells[col_idx]
        p = cell.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(str(header))
        run.bold = True

    for row in table_data.rows:
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            if i < len(row_cells):
                row_cells[i].text = str(value)

    apply_apa_table_borders(table)

    if table_data.note or table_data.source:
        p_note = doc.add_paragraph()
        p_note.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        note_label = p_note.add_run("Nota. ")
        note_label.italic = True

        note_text = table_data.note or ""
        if table_data.source:
            if note_text:
                note_text += f" Fuente: {table_data.source}"
            else:
                note_text = f"Fuente: {table_data.source}"

        p_note.add_run(note_text)

    doc.add_paragraph("")


def add_apa_figure_from_url(doc: Document, figure_data: FigureBlock, figure_number: int) -> None:
    p_num = doc.add_paragraph()
    run_num = p_num.add_run(f"Figura {figure_number}")
    run_num.bold = True

    p_title = doc.add_paragraph()
    run_title = p_title.add_run(figure_data.title)
    run_title.italic = True

    try:
        img_buffer = download_image_to_buffer(figure_data.url)
        doc.add_picture(img_buffer, width=Inches(figure_data.width_inches))
        doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    except Exception as e:
        p_err = doc.add_paragraph()
        run = p_err.add_run(f"[No se pudo cargar la figura: {e}]")
        run.italic = True

    if figure_data.note or figure_data.source:
        p_note = doc.add_paragraph()
        note_label = p_note.add_run("Nota. ")
        note_label.italic = True

        note_text = figure_data.note or ""
        if figure_data.source:
            if note_text:
                note_text += f" Fuente: {figure_data.source}"
            else:
                note_text = f"Fuente: {figure_data.source}"

        p_note.add_run(note_text)

    doc.add_paragraph("")


def add_apa_chart(doc: Document, chart_data: ChartBlock, figure_number: int) -> None:
    p_num = doc.add_paragraph()
    run_num = p_num.add_run(f"Figura {figure_number}")
    run_num.bold = True

    p_title = doc.add_paragraph()
    run_title = p_title.add_run(chart_data.title)
    run_title.italic = True

    try:
        chart_url = build_quickchart_url(chart_data)
        img_buffer = download_image_to_buffer(chart_url)
        doc.add_picture(img_buffer, width=Inches(6.2))
        doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    except Exception as e:
        p_err = doc.add_paragraph()
        run = p_err.add_run(f"[No se pudo generar el gráfico: {e}]")
        run.italic = True

    if chart_data.note or chart_data.source:
        p_note = doc.add_paragraph()
        note_label = p_note.add_run("Nota. ")
        note_label.italic = True

        note_text = chart_data.note or ""
        if chart_data.source:
            if note_text:
                note_text += f" Fuente: {chart_data.source}"
            else:
                note_text = f"Fuente: {chart_data.source}"

        p_note.add_run(note_text)

    doc.add_paragraph("")


# =========================
# INFORME
# =========================

def add_report_cover(doc: Document, report: ReportRequest) -> None:
    if report.logo_url:
        try:
            logo = download_image_to_buffer(report.logo_url)
            doc.add_picture(logo, width=Inches(1.3))
            doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except Exception:
            pass

    if report.institution:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(report.institution)
        run.bold = True
        run.font.size = Pt(14)

    if report.faculty:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run(report.faculty)

    if report.department:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run(report.department)

    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("")

    p_kind = doc.add_paragraph()
    p_kind.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run_kind = p_kind.add_run(report.report_kind.upper())
    run_kind.bold = True
    run_kind.font.size = Pt(16)

    doc.add_paragraph("")

    p_title = doc.add_paragraph()
    p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run_title = p_title.add_run(report.title)
    run_title.bold = True
    run_title.font.size = Pt(20)

    if report.subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run_sub = p_sub.add_run(report.subtitle)
        run_sub.italic = True
        run_sub.font.size = Pt(12)

    doc.add_paragraph("")
    doc.add_paragraph("")

    if report.author:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run_lbl = p.add_run("Autor: ")
        run_lbl.bold = True
        p.add_run(report.author)

    if report.reviewer:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run_lbl = p.add_run("Revisor/Asesor: ")
        run_lbl.bold = True
        p.add_run(report.reviewer)

    doc.add_paragraph("")
    doc.add_paragraph("")

    if report.city or report.date:
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        text = " - ".join([x for x in [report.city, report.date] if x])
        p.add_run(text)

    doc.add_page_break()


def build_numbered_heading(text: str, level: int, counters: List[int]) -> str:
    level = max(1, min(level, 4))
    counters[level - 1] += 1

    for i in range(level, len(counters)):
        counters[i] = 0

    prefix = ".".join(str(counters[i]) for i in range(level) if counters[i] > 0)
    return f"{prefix}. {text}"


def add_report_section(
    doc: Document,
    section: SectionBlock,
    heading_counters: List[int],
    table_counter_ref: dict,
    figure_counter_ref: dict
) -> None:
    heading_text = build_numbered_heading(section.heading, section.level, heading_counters)
    doc.add_heading(heading_text, level=max(1, min(section.level, 4)))

    for paragraph_text in section.paragraphs:
        add_paragraph_block(doc, paragraph_text)

    for bullet in section.bullets:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        add_text_with_bold(p, bullet)

    for item in section.numbered:
        p = doc.add_paragraph(style="List Number")
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        add_text_with_bold(p, item)

    for table_data in section.tables:
        table_counter_ref["value"] += 1
        add_apa_table(doc, table_data, table_counter_ref["value"])

    for figure_data in section.figures:
        figure_counter_ref["value"] += 1
        add_apa_figure_from_url(doc, figure_data, figure_counter_ref["value"])

    for chart_data in section.charts:
        figure_counter_ref["value"] += 1
        add_apa_chart(doc, chart_data, figure_counter_ref["value"])


def build_report_docx(report: ReportRequest) -> BytesIO:
    doc = Document()
    configure_document(doc)
    set_header_footer(doc, report.header_text, report.footer_text)

    add_report_cover(doc, report)

    if report.executive_summary:
        doc.add_heading("Resumen ejecutivo", level=1)
        for item in report.executive_summary:
            add_paragraph_block(doc, item)

    heading_counters = [0, 0, 0, 0]
    table_counter_ref = {"value": 0}
    figure_counter_ref = {"value": 0}

    for section in report.sections:
        add_report_section(
            doc,
            section,
            heading_counters,
            table_counter_ref,
            figure_counter_ref
        )

    if report.conclusions:
        doc.add_heading("Conclusiones", level=1)
        for item in report.conclusions:
            p = doc.add_paragraph(style="List Bullet")
            p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            add_text_with_bold(p, item)

    if report.recommendations:
        doc.add_heading("Recomendaciones", level=1)
        for item in report.recommendations:
            p = doc.add_paragraph(style="List Number")
            p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            add_text_with_bold(p, item)

    if report.references:
        doc.add_heading("Referencias", level=1)
        for ref in report.references:
            add_paragraph_block(doc, ref, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =========================
# CARTA
# =========================

def build_letter_docx(letter: LetterRequest) -> BytesIO:
    doc = Document()
    configure_document(doc)

    if letter.logo_url:
        try:
            logo = download_image_to_buffer(letter.logo_url)
            doc.add_picture(logo, width=Inches(1.1))
            doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except Exception:
            pass

    if letter.organization:
        p_org = doc.add_paragraph()
        p_org.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run_org = p_org.add_run(letter.organization)
        run_org.bold = True
        run_org.font.size = Pt(14)

    doc.add_paragraph("")

    p_date = doc.add_paragraph()
    p_date.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p_date.add_run(letter.city_date)

    doc.add_paragraph("")

    p_rec = doc.add_paragraph()
    p_rec.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    run_name = p_rec.add_run(letter.recipient_name)
    run_name.bold = True

    if letter.recipient_title:
        p = doc.add_paragraph()
        p.add_run(letter.recipient_title)

    if letter.recipient_organization:
        p = doc.add_paragraph()
        p.add_run(letter.recipient_organization)

    doc.add_paragraph("")

    p_sub = doc.add_paragraph()
    run_lbl = p_sub.add_run("Asunto: ")
    run_lbl.bold = True
    p_sub.add_run(letter.subject)

    doc.add_paragraph("")
    doc.add_paragraph(letter.greeting)

    for paragraph_text in letter.body:
        add_paragraph_block(doc, paragraph_text)

    doc.add_paragraph("")
    doc.add_paragraph(letter.closing)

    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("")

    p_sign = doc.add_paragraph()
    p_sign.add_run(letter.signature_name)

    if letter.signature_title:
        p_title = doc.add_paragraph()
        p_title.add_run(letter.signature_title)

    if letter.annexes:
        doc.add_paragraph("")
        p_ann = doc.add_paragraph()
        run_ann = p_ann.add_run("Anexos:")
        run_ann.bold = True

        for item in letter.annexes:
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(item)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =========================
# ENDPOINTS
# =========================

@app.get("/")
def root():
    return {
        "message": "Backend activo",
        "service": "gpt-doc-backend"
    }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate-document")
def generate_document(payload: DocumentRequest):
    if payload.document_type == "carta":
        if not payload.letter:
            raise HTTPException(status_code=400, detail="Falta el bloque 'letter' para generar una carta.")

        buffer = build_letter_docx(payload.letter)
        safe_name = sanitize_filename(payload.letter.filename)

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{safe_name}.docx"'
            }
        )

    if payload.document_type == "informe":
        if not payload.report:
            raise HTTPException(status_code=400, detail="Falta el bloque 'report' para generar un informe.")

        buffer = build_report_docx(payload.report)
        safe_name = sanitize_filename(payload.report.filename)

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{safe_name}.docx"'
            }
        )

    raise HTTPException(status_code=400, detail="Tipo de documento no válido.")
