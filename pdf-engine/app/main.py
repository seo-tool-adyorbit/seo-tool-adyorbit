"""
FastAPI Server — AI SEO PDF Automation Suite
LibreOffice + python-docx + PyMuPDF Pipeline
"""
import io
import re
import zipfile
from fastapi import FastAPI, HTTPException, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional

from app.tfn_engine import SmartTfnEngine
from app.pdf_service import LibreOfficePdfService, VectorPdfService

app = FastAPI(
    title="AI SEO PDF Automation Suite",
    version="2.0.0",
    description="LibreOffice-powered Google Docs quality PDF engine"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Request Models ─────────────────────────────────────────────────────────────

class FormatTfnRequest(BaseModel):
    competitor_str: str
    target_input: str

class ReplaceTextRequest(BaseModel):
    html_content: str
    target_lines: List[str]

class GeneratePdfRequest(BaseModel):
    html_content: str
    title: Optional[str] = "SEO PDF Document"

class CsvRow(BaseModel):
    airline_name: str
    headline: str
    pdf_name: str

class BulkZipRequest(BaseModel):
    source_airline: str
    template_html: str
    target_tfn_lines: List[str]
    csv_rows: List[CsvRow]


# ── Health ─────────────────────────────────────────────────────────────────────

@app.get("/")
def read_root():
    lo_available = LibreOfficePdfService.libreoffice_available()
    return {
        "service": "AI SEO PDF Automation Suite Engine",
        "version": "2.0.0",
        "status": "online",
        "libreoffice_available": lo_available,
        "pdf_engine": "LibreOffice + python-docx + PyMuPDF" if lo_available else "ReportLab (fallback)"
    }

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "libreoffice": LibreOfficePdfService.libreoffice_available()
    }


# ── TFN Endpoints ──────────────────────────────────────────────────────────────

@app.post("/api/v1/tfn/format")
def format_tfn(req: FormatTfnRequest):
    formatted = SmartTfnEngine.format_to_competitor_pattern(req.competitor_str, req.target_input)
    return {"original": req.competitor_str, "formatted": formatted}

@app.post("/api/v1/tfn/replace-text")
def replace_text(req: ReplaceTextRequest):
    new_html, count = SmartTfnEngine.replace_text_alternating(req.html_content, req.target_lines)
    return {"new_html": new_html, "replaced_count": count}


# ── PDF Endpoint ───────────────────────────────────────────────────────────────

@app.post("/api/v1/pdf/generate")
def generate_pdf(req: GeneratePdfRequest):
    """
    Generate Google Docs quality PDF via LibreOffice pipeline.
    HTML → python-docx DOCX → LibreOffice headless → PyMuPDF optimize
    Returns: application/pdf binary response
    """
    try:
        pdf_bytes = LibreOfficePdfService.generate_pdf(req.html_content, req.title)
        safe_title = re.sub(r'[^\w\s-]', '', req.title or 'SEO_Document').strip().replace(' ', '_')
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{safe_title}.pdf"',
                "X-PDF-Engine": "LibreOffice",
                "X-PDF-Searchable": "true",
            }
        )
    except RuntimeError as e:
        # LibreOffice not installed or conversion failed
        raise HTTPException(status_code=503, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")


# ── Bulk ZIP Endpoint ──────────────────────────────────────────────────────────

@app.post("/api/v1/pdf/bulk-zip")
def generate_bulk_zip(req: BulkZipRequest):
    """
    Generate bulk PDFs and return as ZIP archive.
    Each row: airline replacement + headline + TFN replacement → individual PDF.
    """
    try:
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for idx, row in enumerate(req.csv_rows):
                doc_html = req.template_html

                # 1. Replace Airline Name
                if req.source_airline:
                    doc_html = re.sub(
                        re.escape(req.source_airline),
                        row.airline_name,
                        doc_html,
                        flags=re.IGNORECASE
                    )

                # 2. Replace Headline (h1 tag)
                if row.headline:
                    if re.search(r'<h1[^>]*>.*?</h1>', doc_html, flags=re.IGNORECASE | re.DOTALL):
                        doc_html = re.sub(
                            r'<h1[^>]*>.*?</h1>',
                            f'<h1>{row.headline}</h1>',
                            doc_html,
                            flags=re.IGNORECASE | re.DOTALL
                        )
                    else:
                        doc_html = f'<h1>{row.headline}</h1>' + doc_html

                # 3. Replace TFNs
                if req.target_tfn_lines:
                    doc_html, _ = SmartTfnEngine.replace_text_alternating(doc_html, req.target_tfn_lines)

                # 4. Generate PDF
                pdf_bytes = LibreOfficePdfService.generate_pdf(doc_html, row.headline or row.airline_name)
                clean_name = re.sub(r'[^a-zA-Z0-9\s\-_]', '', row.pdf_name).strip() or f"Doc_{idx+1}"
                if not clean_name.lower().endswith('.pdf'):
                    clean_name += '.pdf'

                zip_file.writestr(clean_name, pdf_bytes)

        zip_bytes = zip_buffer.getvalue()
        zip_buffer.close()

        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers={"Content-Disposition": 'attachment; filename="Bulk_SEO_PDFs.zip"'}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Bulk ZIP failed: {str(e)}")
