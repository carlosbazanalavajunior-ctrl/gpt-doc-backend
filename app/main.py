from io import BytesIO
from typing import Optional, List
import json
import re

import requests
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


app = FastAPI(
    title="GPT Doc Backend",
    description="Backend profesional para generar documentos Word desde FastAPI",
    version="0.2.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
)


class ReportRequest(BaseModel):
    title: str
    paragraphs: List[str] = []


class MetadataItem(BaseModel):
    label: str
    value: str


class TableBlock(BaseModel):
    title: Optional[str] = None
    headers: List[str]
    rows: List[List[str]] = []


class ImageBlock(BaseModel):
    url: str
    caption: Optional[str] = None
    width_inches: float = 5.5


class ChartDataset(BaseModel):
    label: str
    data: List[float]


class ChartBlock(BaseModel):
    title: Optional[str] = None
    chart_type: str = Field(default="bar", alias="type")
    labels: List[str]
    datasets: List[ChartDataset]
    width: int = 800
    height: int = 450


class SectionBlock(BaseModel):
    heading: str
    level: int = 1
    paragraphs: List[str] = []
    bullets: List[str] = []
    numbered: List[str] = []
    tables: List[TableBlock] = []
    images: List[ImageBlock] = []
    charts: List[ChartBlock] = []


class RichReportRequest(BaseModel):
    filename: str = "informe_profesional"
    title: str
    subtitle: Optional[str] = None
    metadata: List[MetadataItem] = []
    executive_summary: List[str] = []
    sections: List[SectionBlock] = []
    conclusions: List[str] = []
    recommendations: List[str] = []


def sanitize_filename(name: str) -> str:
    safe = name.strip().lower().replace(" ", "_")
    safe = re.sub(r"[^a-zA-Z0-9_\-]", "", safe)
    return safe or "documento"


def configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    styles = doc.styles
    normal_style = styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)


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


def add_paragraph_block(doc: Document, text: str, italic: bool = False) -> None:
    p = doc.add_paragraph()
    add_text_with_bold(p, text)
    if italic:
        for run in p.runs:
            run.italic = True
    p.paragraph_format.space_after = Pt(6)


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    level = max(1, min(level, 4))
    doc.add_heading(text, level=level)


def add_metadata(doc: Document, metadata: List[MetadataItem]) -> None:
    if not metadata:
        return

    add_heading(doc, "Datos generales", level=1)
    for item in metadata:
        p = doc.add_paragraph()
        label_run = p.add_run(f"{item.label}: ")
        label_run.bold = True
        p.add_run(item.value)
        p.paragraph_format.space_after = Pt(2)


def add_table_block(doc: Document, table_data: TableBlock) -> None:
    if table_data.title:
        p_title = doc.add_paragraph()
        run = p_title.add_run(table_data.title)
        run.bold = True
        p_title.paragraph_format.space_after = Pt(4)

    table = doc.add_table(rows=1, cols=len(table_data.headers))
    table.style = "Table Grid"

    header_cells = table.rows[0].cells
    for i, header in enumerate(table_data.headers):
        header_paragraph = header_cells[i].paragraphs[0]
        run = header_paragraph.add_run(str(header))
        run.bold = True

    for row in table_data.rows:
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            if i < len(row_cells):
                row_cells[i].text = str(value)

    doc.add_paragraph("")


def download_image_to_buffer(url: str) -> BytesIO:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return BytesIO(response.content)


def add_image_block(doc: Document, image_data: ImageBlock) -> None:
    try:
        img_buffer = download_image_to_buffer(image_data.url)
        doc.add_picture(img_buffer, width=Inches(image_data.width_inches))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        if image_data.caption:
            cap = doc.add_paragraph()
            cap.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = cap.add_run(image_data.caption)
            run.italic = True
            run.font.size = Pt(10)

        doc.add_paragraph("")
    except Exception as e:
        p = doc.add_paragraph()
        run = p.add_run(f"[No se pudo cargar la imagen: {e}]")
        run.italic = True


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
                    "display": bool(chart.title),
                    "text": chart.title or ""
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


def add_chart_block(doc: Document, chart: ChartBlock) -> None:
    try:
        chart_url = build_quickchart_url(chart)
        img_buffer = download_image_to_buffer(chart_url)
        doc.add_picture(img_buffer, width=Inches(6.2))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        doc.add_paragraph("")
    except Exception as e:
        p = doc.add_paragraph()
        run = p.add_run(f"[No se pudo generar el gráfico: {e}]")
        run.italic = True


def add_section_block(doc: Document, section: SectionBlock) -> None:
    add_heading(doc, section.heading, level=section.level)

    for paragraph_text in section.paragraphs:
        add_paragraph_block(doc, paragraph_text)

    for bullet in section.bullets:
        p = doc.add_paragraph(style="List Bullet")
        add_text_with_bold(p, bullet)

    for item in section.numbered:
        p = doc.add_paragraph(style="List Number")
        add_text_with_bold(p, item)

    for table_data in section.tables:
        add_table_block(doc, table_data)

    for image_data in section.images:
        add_image_block(doc, image_data)

    for chart_data in section.charts:
        add_chart_block(doc, chart_data)


def build_basic_docx(payload: ReportRequest) -> BytesIO:
    doc = Document()
    configure_document(doc)

    title = doc.add_paragraph()
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = title.add_run(payload.title)
    run.bold = True
    run.font.size = Pt(18)

    doc.add_paragraph("")

    for paragraph in payload.paragraphs:
        add_paragraph_block(doc, paragraph)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def build_rich_docx(payload: RichReportRequest) -> BytesIO:
    doc = Document()
    configure_document(doc)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run_title = p_title.add_run(payload.title)
    run_title.bold = True
    run_title.font.size = Pt(20)

    if payload.subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run_sub = p_sub.add_run(payload.subtitle)
        run_sub.italic = True
        run_sub.font.size = Pt(12)

    doc.add_paragraph("")

    add_metadata(doc, payload.metadata)
    if payload.metadata:
        doc.add_paragraph("")

    if payload.executive_summary:
        add_heading(doc, "Resumen ejecutivo", level=1)
        for item in payload.executive_summary:
            add_paragraph_block(doc, item)

    for section in payload.sections:
        add_section_block(doc, section)

    if payload.conclusions:
        add_heading(doc, "Conclusiones", level=1)
        for item in payload.conclusions:
            p = doc.add_paragraph(style="List Bullet")
            add_text_with_bold(p, item)

    if payload.recommendations:
        add_heading(doc, "Recomendaciones", level=1)
        for item in payload.recommendations:
            p = doc.add_paragraph(style="List Number")
            add_text_with_bold(p, item)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


@app.get("/")
def root():
    return {
        "message": "Backend activo",
        "service": "gpt-doc-backend"
    }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate-docx")
def generate_docx(payload: ReportRequest):
    buffer = build_basic_docx(payload)
    safe_name = sanitize_filename(payload.title)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_name}.docx"'
        }
    )


@app.post("/generate-rich-docx")
def generate_rich_docx(payload: RichReportRequest):
    buffer = build_rich_docx(payload)
    safe_name = sanitize_filename(payload.filename)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_name}.docx"'
        }
    )
