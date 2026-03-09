from io import BytesIO

from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from docx import Document


app = FastAPI(
    title="GPT Doc Backend",
    description="Backend básico para generar documentos Word desde FastAPI",
    version="0.1.0"
)


class ReportRequest(BaseModel):
    title: str
    paragraphs: list[str] = []


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
    doc = Document()
    doc.add_heading(payload.title, 0)

    for paragraph in payload.paragraphs:
        doc.add_paragraph(paragraph)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    safe_name = payload.title.strip().replace(" ", "_").lower()
    if not safe_name:
        safe_name = "documento"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_name}.docx"'
        }
    )
