from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Any, List, Literal, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph


app = FastAPI(
    title="GPT DOC Backend",
    version="0.9.2 template-based robust tables + clean toc",
    description="Generador de documentos DOCX profesionales basado en plantilla."
)

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "templates" / "professional_report_template.docx"


# =========================
# MODELOS
# =========================

class TableBlock(BaseModel):
    title: str
    headers: List[str] = Field(default_factory=list)
    rows: List[List[Any]] = Field(default_factory=list)
    note: Optional[str] = None


class FigureBlock(BaseModel):
    title: str
    caption: Optional[str] = None


class SectionBlock(BaseModel):
    heading: str
    paragraphs: List[str] = Field(default_factory=list)
    bullet_points: List[str] = Field(default_factory=list)
    numbered_points: List[str] = Field(default_factory=list)
    tables: List[TableBlock] = Field(default_factory=list)
    figures: List[FigureBlock] = Field(default_factory=list)


class GenerateDocumentRequest(BaseModel):
    document_type: Literal["carta", "informe"]
    filename: Optional[str] = None

    institution: Optional[str] = ""
    faculty: Optional[str] = ""
    department: Optional[str] = ""
    report_kind: Optional[str] = ""
    title: Optional[str] = ""
    subtitle: Optional[str] = ""
    author: Optional[str] = ""
    reviewer: Optional[str] = ""
    city_date: Optional[str] = ""

    intro_paragraphs: List[str] = Field(default_factory=list)
    sections: List[SectionBlock] = Field(default_factory=list)

    # Campos específicos para carta
    recipient_name: Optional[str] = ""
    recipient_position: Optional[str] = ""
    recipient_institution: Optional[str] = ""
    subject: Optional[str] = ""
    greeting: Optional[str] = "De mi consideración:"
    body_paragraphs: List[str] = Field(default_factory=list)
    closing: Optional[str] = "Atentamente,"
    signature_name: Optional[str] = ""
    signature_position: Optional[str] = ""


# =========================
# UTILIDADES GENERALES
# =========================

def safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def clean_heading_for_index(text: str) -> str:
    text = safe_text(text).strip()
    text = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", text)
    return text.strip()


def sanitize_filename(filename: str) -> str:
    filename = filename.strip() or "documento.docx"
    filename = re.sub(r'[\\/*?:"<>|]+', "_", filename)
    if not filename.lower().endswith(".docx"):
        filename += ".docx"
    return filename


def set_run_font(run, size=11, bold=False, italic=False):
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def format_paragraph(
    paragraph: Paragraph,
    *,
    size: int = 11,
    bold: bool = False,
    italic: bool = False,
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
    space_after: int = 6,
    line_spacing: float = 1.15
):
    paragraph.alignment = alignment
    pf = paragraph.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    for run in paragraph.runs:
        set_run_font(run, size=size, bold=bold, italic=italic)


def set_paragraph_text(
    paragraph: Paragraph,
    text: str,
    *,
    size: int = 11,
    bold: bool = False,
    italic: bool = False,
    alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
    space_after: int = 6
):
    paragraph.text = text
    format_paragraph(
        paragraph,
        size=size,
        bold=bold,
        italic=italic,
        alignment=alignment,
        space_after=space_after
    )


def find_paragraph_with_placeholder(document: Document, placeholder: str) -> Optional[Paragraph]:
    for paragraph in document.paragraphs:
        if placeholder in paragraph.text:
            return paragraph
    return None


def replace_placeholder_in_paragraph(paragraph: Paragraph, placeholder: str, value: str):
    if placeholder not in paragraph.text:
        return

    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            return

    paragraph.text = paragraph.text.replace(placeholder, value)


def replace_placeholder_everywhere(document: Document, mapping: dict[str, str]):
    for paragraph in document.paragraphs:
        for placeholder, value in mapping.items():
            if placeholder in paragraph.text:
                replace_placeholder_in_paragraph(paragraph, placeholder, value)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for placeholder, value in mapping.items():
                        if placeholder in paragraph.text:
                            replace_placeholder_in_paragraph(paragraph, placeholder, value)


def add_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def replace_placeholder_with_lines(document: Document, placeholder: str, lines: List[str]):
    target = find_paragraph_with_placeholder(document, placeholder)
    lines = lines or ["Sin elementos registrados."]

    if not target:
        replace_placeholder_everywhere(document, {placeholder: "\n".join(lines)})
        return

    first_line = lines[0]
    set_paragraph_text(
        target,
        target.text.replace(placeholder, first_line),
        size=11,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_after=2
    )

    anchor = target
    for line in lines[1:]:
        anchor = add_paragraph_after(anchor, line)
        format_paragraph(
            anchor,
            size=11,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=2,
            line_spacing=1.0
        )


def get_existing_table_style(document: Document) -> Optional[str]:
    preferred_candidates = [
        "Table Grid",
        "Tabla con cuadrícula",
        "Normal Table",
        "Tabla normal",
        "Light Grid",
        "Cuadrícula clara",
    ]

    existing_styles = {style.name for style in document.styles if getattr(style, "name", None)}

    for style_name in preferred_candidates:
        if style_name in existing_styles:
            return style_name

    return None


# =========================
# UTILIDADES DE TABLAS
# =========================

def set_cell_shading(cell, fill: str = "D9EAF7"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, color: str = "808080", size: str = "6"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()

    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)

    for edge in ("top", "left", "bottom", "right"):
        edge_tag = qn(f"w:{edge}")
        element = tc_borders.find(edge_tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tc_borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_cell_text(
    cell,
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    size: int = 10,
    alignment=WD_ALIGN_PARAGRAPH.LEFT,
    shaded: bool = False
):
    cell.text = safe_text(text)

    for paragraph in cell.paragraphs:
        format_paragraph(
            paragraph,
            size=size,
            bold=bold,
            italic=italic,
            alignment=alignment,
            space_after=0,
            line_spacing=1.0
        )

    set_cell_border(cell)

    if shaded:
        set_cell_shading(cell)


def normalize_rows(headers: List[str], rows: List[List[Any]]) -> List[List[str]]:
    normalized = []
    total_cols = max(1, len(headers))

    for row in rows:
        row = [safe_text(v) for v in row]
        if len(row) < total_cols:
            row = row + [""] * (total_cols - len(row))
        elif len(row) > total_cols:
            row = row[:total_cols]
        normalized.append(row)

    return normalized


def insert_table_after(
    document: Document,
    anchor_paragraph: Paragraph,
    headers: List[str],
    rows: List[List[Any]]
) -> Paragraph:
    headers = headers or ["Columna 1"]
    rows = normalize_rows(headers, rows)

    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    table_style = get_existing_table_style(document)
    if table_style:
        table.style = table_style

    header_cells = table.rows[0].cells
    for idx, header in enumerate(headers):
        set_cell_text(
            header_cells[idx],
            safe_text(header),
            bold=True,
            size=10,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            shaded=True
        )

    for row in rows:
        row_cells = table.add_row().cells
        for idx, value in enumerate(row):
            set_cell_text(
                row_cells[idx],
                safe_text(value),
                size=10,
                alignment=WD_ALIGN_PARAGRAPH.LEFT
            )

    tbl = table._tbl
    anchor_paragraph._p.addnext(tbl)

    new_p = OxmlElement("w:p")
    tbl.addnext(new_p)
    new_anchor = Paragraph(new_p, anchor_paragraph._parent)

    return new_anchor


# =========================
# ÍNDICES
# =========================

def build_toc_lines(payload: GenerateDocumentRequest) -> List[str]:
    if payload.document_type == "carta":
        return [
            "1. Encabezado institucional",
            "2. Datos del destinatario",
            "3. Asunto",
            "4. Cuerpo de la carta",
            "5. Despedida y firma",
        ]

    lines = []
    current = 1

    if payload.intro_paragraphs:
        lines.append(f"{current}. Introducción")
        current += 1

    for section in payload.sections:
        clean_heading = clean_heading_for_index(section.heading)
        lines.append(f"{current}. {clean_heading}")
        current += 1

    return lines or ["1. Contenido principal"]


def build_table_index_lines(payload: GenerateDocumentRequest) -> List[str]:
    lines = []
    counter = 1

    for section in payload.sections:
        for table in section.tables:
            lines.append(f"Tabla {counter}. {table.title}")
            counter += 1

    return lines or ["Sin tablas registradas."]


def build_figure_index_lines(payload: GenerateDocumentRequest) -> List[str]:
    lines = []
    counter = 1

    for section in payload.sections:
        for figure in section.figures:
            lines.append(f"Figura {counter}. {figure.title}")
            counter += 1

    return lines or ["Sin figuras registradas."]


# =========================
# RENDER DEL CUERPO
# =========================

def heading_font_size(heading: str) -> int:
    token = heading.strip().split(" ")[0]
    dot_count = token.count(".")
    if dot_count >= 2:
        return 11
    if dot_count == 1:
        return 12
    return 14


def render_report_body(document: Document, anchor: Paragraph, payload: GenerateDocumentRequest):
    table_counter = 1
    figure_counter = 1

    if payload.intro_paragraphs:
        anchor = add_paragraph_after(anchor, "Introducción")
        format_paragraph(
            anchor,
            size=14,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=6
        )

        for paragraph_text in payload.intro_paragraphs:
            anchor = add_paragraph_after(anchor, paragraph_text)
            format_paragraph(anchor, size=11, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6)

    for section in payload.sections:
        anchor = add_paragraph_after(anchor, section.heading)
        format_paragraph(
            anchor,
            size=heading_font_size(section.heading),
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=6
        )

        for paragraph_text in section.paragraphs:
            anchor = add_paragraph_after(anchor, paragraph_text)
            format_paragraph(anchor, size=11, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6)

        for bullet in section.bullet_points:
            anchor = add_paragraph_after(anchor, f"• {bullet}")
            format_paragraph(anchor, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=3)

        for idx, item in enumerate(section.numbered_points, start=1):
            anchor = add_paragraph_after(anchor, f"{idx}. {item}")
            format_paragraph(anchor, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=3)

        for table in section.tables:
            anchor = add_paragraph_after(anchor, f"Tabla {table_counter}. {table.title}")
            format_paragraph(
                anchor,
                size=11,
                bold=True,
                italic=False,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                space_after=3
            )

            anchor = insert_table_after(document, anchor, table.headers, table.rows)

            if table.note:
                anchor = add_paragraph_after(anchor, f"Nota. {table.note}")
                format_paragraph(
                    anchor,
                    size=10,
                    italic=True,
                    alignment=WD_ALIGN_PARAGRAPH.LEFT,
                    space_after=6
                )

            table_counter += 1

        for figure in section.figures:
            anchor = add_paragraph_after(anchor, f"Figura {figure_counter}. {figure.title}")
            format_paragraph(
                anchor,
                size=11,
                bold=True,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                space_after=3
            )

            if figure.caption:
                anchor = add_paragraph_after(anchor, figure.caption)
                format_paragraph(
                    anchor,
                    size=10,
                    italic=True,
                    alignment=WD_ALIGN_PARAGRAPH.LEFT,
                    space_after=6
                )

            figure_counter += 1


def render_letter_body(document: Document, anchor: Paragraph, payload: GenerateDocumentRequest):
    if payload.city_date:
        anchor = add_paragraph_after(anchor, payload.city_date)
        format_paragraph(
            anchor,
            size=11,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            space_after=12
        )

    recipient_lines = [
        safe_text(payload.recipient_name),
        safe_text(payload.recipient_position),
        safe_text(payload.recipient_institution),
    ]
    recipient_lines = [line for line in recipient_lines if line.strip()]

    for line in recipient_lines:
        anchor = add_paragraph_after(anchor, line)
        format_paragraph(
            anchor,
            size=11,
            bold=True if line == payload.recipient_name else False,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=1
        )

    if recipient_lines:
        anchor = add_paragraph_after(anchor, "")
        format_paragraph(anchor, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=6)

    if payload.subject:
        anchor = add_paragraph_after(anchor, f"Asunto: {payload.subject}")
        format_paragraph(
            anchor,
            size=11,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=12
        )

    if payload.greeting:
        anchor = add_paragraph_after(anchor, payload.greeting)
        format_paragraph(
            anchor,
            size=11,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=12
        )

    body_paragraphs = payload.body_paragraphs or [
        "Se deja constancia del contenido principal de la presente comunicación."
    ]

    for paragraph_text in body_paragraphs:
        anchor = add_paragraph_after(anchor, paragraph_text)
        format_paragraph(
            anchor,
            size=11,
            alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            space_after=8
        )

    if payload.closing:
        anchor = add_paragraph_after(anchor, "")
        format_paragraph(anchor, size=11, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=6)

        anchor = add_paragraph_after(anchor, payload.closing)
        format_paragraph(
            anchor,
            size=11,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=18
        )

    if payload.signature_name:
        anchor = add_paragraph_after(anchor, payload.signature_name)
        format_paragraph(
            anchor,
            size=11,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=1
        )

    if payload.signature_position:
        anchor = add_paragraph_after(anchor, payload.signature_position)
        format_paragraph(
            anchor,
            size=11,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            space_after=1
        )


# =========================
# GENERACIÓN PRINCIPAL
# =========================

def build_document(payload: GenerateDocumentRequest) -> io.BytesIO:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"No se encontró la plantilla en: {TEMPLATE_PATH}")

    document = Document(str(TEMPLATE_PATH))

    report_kind = payload.report_kind.strip() if payload.report_kind else ""
    if not report_kind:
        report_kind = "Carta formal" if payload.document_type == "carta" else "Informe técnico"

    toc_lines = build_toc_lines(payload)
    table_index_lines = build_table_index_lines(payload)
    figure_index_lines = build_figure_index_lines(payload)

    simple_mapping = {
        "{{INSTITUTION}}": safe_text(payload.institution),
        "{{FACULTY}}": safe_text(payload.faculty),
        "{{DEPARTMENT}}": safe_text(payload.department),
        "{{REPORT_KIND}}": safe_text(report_kind),
        "{{TITLE}}": safe_text(payload.title),
        "{{SUBTITLE}}": safe_text(payload.subtitle),
        "{{AUTHOR}}": safe_text(payload.author),
        "{{REVIEWER}}": safe_text(payload.reviewer),
        "{{CITY_DATE}}": safe_text(payload.city_date),
    }
    replace_placeholder_everywhere(document, simple_mapping)

    replace_placeholder_with_lines(document, "{{TOC}}", toc_lines)
    replace_placeholder_with_lines(document, "{{TABLE_INDEX}}", table_index_lines)
    replace_placeholder_with_lines(document, "{{FIGURE_INDEX}}", figure_index_lines)

    body_anchor = find_paragraph_with_placeholder(document, "{{BODY_CONTENT}}")
    if body_anchor:
        replace_placeholder_in_paragraph(body_anchor, "{{BODY_CONTENT}}", "")
    else:
        if document.paragraphs:
            body_anchor = document.paragraphs[-1]
        else:
            document.add_paragraph("")
            body_anchor = document.paragraphs[-1]

    if payload.document_type == "carta":
        render_letter_body(document, body_anchor, payload)
    else:
        render_report_body(document, body_anchor, payload)

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output


# =========================
# ENDPOINTS
# =========================

@app.get("/")
def root():
    return {
        "message": "GPT DOC Backend activo",
        "version": "0.9.2 template-based robust tables + clean toc",
        "template_path": str(TEMPLATE_PATH),
        "template_exists": TEMPLATE_PATH.exists(),
        "allowed_document_types": ["carta", "informe"],
    }


@app.get("/health")
def health():
    return {
        "status": "ok",
        "template_exists": TEMPLATE_PATH.exists(),
        "template_path": str(TEMPLATE_PATH),
    }


@app.post("/generate-document")
def generate_document(payload: GenerateDocumentRequest):
    try:
        output = build_document(payload)

        default_name = "carta_profesional.docx" if payload.document_type == "carta" else "informe_profesional.docx"
        filename = sanitize_filename(payload.filename or default_name)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"'
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo generar el documento. Detalle técnico: {str(e)}"
        )