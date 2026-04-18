"""
Motore OCR ibrido — docling (preciso) con fallback Tesseract (leggero).

Strategia:
  1. Prova docling (EasyOCR) se installato → risultato di alta qualità
  2. Se docling non disponibile o fallisce → Tesseract via subprocess
  3. Se anche Tesseract manca → PyMuPDF grezzo come ultima risorsa

Nessun crash: restituisce sempre ExtractionResult, con engine="fallback"
se nessun motore OCR è disponibile.
"""
from __future__ import annotations
from pathlib import Path
import logging
import io
import tempfile
import shutil

from src.extractor_pymupdf import ExtractionResult, Element, BBox

logger = logging.getLogger(__name__)


def _try_docling(pdf_path: Path) -> ExtractionResult | None:
    """Prova a usare docling. Ritorna None se non disponibile."""
    try:
        from docling.document_converter import DocumentConverter
        from docling.datamodel.base_models import InputFormat
    except ImportError:
        logger.info("docling non installato — salto")
        return None

    try:
        logger.info("Avvio docling su %s", pdf_path.name)
        converter = DocumentConverter()
        doc_result = converter.convert(str(pdf_path))
        doc = doc_result.document

        md = doc.export_to_markdown()
        elements: list[Element] = []

        # Estrai elementi con bounding box dove disponibili
        for item, _ in doc.iterate_items():
            text = getattr(item, "text", "") or ""
            if not text.strip():
                continue
            item_type = type(item).__name__.lower()
            elem_type = "heading" if "section" in item_type or "heading" in item_type else "paragraph"
            prov = getattr(item, "prov", None)
            if prov and len(prov) > 0:
                p = prov[0]
                bbox_raw = getattr(p, "bbox", None)
                page_no  = getattr(p, "page_no", 0) or 0
                if bbox_raw:
                    bbox = BBox(
                        x0=float(getattr(bbox_raw, "l", 0)),
                        y0=float(getattr(bbox_raw, "t", 0)),
                        x1=float(getattr(bbox_raw, "r", 0)),
                        y1=float(getattr(bbox_raw, "b", 0)),
                    )
                else:
                    bbox = BBox(0, 0, 0, 0)
                page_idx = max(0, page_no - 1)
            else:
                bbox    = BBox(0, 0, 0, 0)
                page_idx = 0

            elements.append(Element(
                type=elem_type, text=text.strip(),
                page=page_idx, bbox=bbox,
            ))

        return ExtractionResult(
            markdown=md,
            elements=elements,
            engine="docling",
        )

    except Exception as exc:
        logger.warning("docling fallito: %s", exc)
        return None


def _find_tess() -> str | None:
    """Cerca Tesseract nel PATH e nei percorsi comuni Windows."""
    import os
    t = shutil.which("tesseract")
    if t:
        return t
    for p in [
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
    ]:
        if os.path.isfile(p):
            return p
    return None


def _try_tesseract(pdf_path: Path) -> ExtractionResult | None:
    """Prova OCR con Tesseract. Cerca anche nei percorsi comuni Windows."""
    tess = _find_tess()
    if not tess:
        logger.info("Tesseract non trovato nel PATH ne nei percorsi comuni — salto")
        return None

    try:
        import fitz
        import subprocess
    except ImportError:
        return None

    try:
        logger.info("OCR Tesseract su %s", pdf_path.name)
        md_pages: list[str] = []
        elements: list[Element] = []

        from PIL import Image, ImageEnhance
        with fitz.open(str(pdf_path)) as doc:
            page_count = doc.page_count
            for page_idx, page in enumerate(doc):
                with tempfile.TemporaryDirectory() as tmp:
                    img_path = Path(tmp) / "page.png"
                    # 300 DPI + contrast per scansioni
                    pix = page.get_pixmap(dpi=300)
                    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
                    img = ImageEnhance.Contrast(img).enhance(2.0)
                    img.save(str(img_path))

                    out_base = Path(tmp) / "out"
                    subprocess.run(
                        [tess, str(img_path), str(out_base),
                         "-l", "ita+eng", "--psm", "3", "--oem", "3"],
                        check=True, capture_output=True
                    )
                    txt_path = out_base.with_suffix(".txt")
                    text = txt_path.read_text(encoding="utf-8", errors="replace").strip()

                md_pages.append(f"\n\n---\n*Pagina {page_idx + 1}*\n\n{text}")
                if text:
                    elements.append(Element(
                        type="paragraph", text=text,
                        page=page_idx, bbox=BBox(0, 0, page.rect.width, page.rect.height),
                    ))

        return ExtractionResult(
            markdown="\n".join(md_pages).strip(),
            elements=elements,
            page_count=page_count,
            engine="tesseract",
        )

    except Exception as exc:
        logger.warning("Tesseract fallito: %s", exc)
        return None


def extract_ocr(pdf_path: Path) -> ExtractionResult:
    """
    Estrazione OCR con fallback automatico:
      1. docling (se installato)
      2. Tesseract (se nel PATH)
      3. PyMuPDF grezzo (sempre disponibile)
    """
    # Tentativo 1: docling
    result = _try_docling(pdf_path)
    if result and not result.error:
        return result

    # Tentativo 2: Tesseract
    result = _try_tesseract(pdf_path)
    if result and not result.error:
        return result

    # Fallback finale: PyMuPDF grezzo (restituisce quello che riesce a leggere)
    logger.warning(
        "Nessun motore OCR disponibile per %s — uso PyMuPDF grezzo",
        pdf_path.name
    )
    from src.extractor_pymupdf import extract as pymupdf_extract
    result = pymupdf_extract(pdf_path)
    result.engine = "pymupdf_fallback"
    return result
