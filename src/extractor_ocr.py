"""
Motore OCR ibrido v3 — integra i miglioramenti del Convertitore File AI v15.

Miglioramenti rispetto alla versione precedente:
  - PSM 4 (colonna singola testo variabile) invece di PSM 6
  - 400 DPI per documenti VETT-RASTER, 200 DPI per scansioni normali
  - Pre-processing: grayscale + contrasto x2 + nitidezza x2
  - Fallback robusto: docling -> Tesseract -> PyMuPDF
"""
from __future__ import annotations

import io
import logging
import shutil
import subprocess
import tempfile
from pathlib import Path

from src.extractor_pymupdf import ExtractionResult, Element, BBox

logger = logging.getLogger(__name__)


def _find_tess() -> str | None:
    """Cerca Tesseract nel PATH e nei percorsi comuni Windows e Linux."""
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


def _preprocess_image(img_path: Path) -> Path | None:
    """Pre-processing v15: grayscale + contrasto x2 + nitidezza x2."""
    try:
        from PIL import Image, ImageEnhance
        img = Image.open(str(img_path)).convert("L")
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        out = img_path.parent / f"prep_{img_path.stem}.png"
        img.save(str(out))
        return out
    except Exception as exc:
        logger.debug("Pre-processing non riuscito: %s", exc)
        return None


def _try_tesseract(
    pdf_path: Path,
    page_from: int = 0,
    page_to: int | None = None,
    dpi: int = 200,
    psm: str = "4",
) -> ExtractionResult | None:
    """OCR con Tesseract — motore v15.

    page_from: indice 0-based prima pagina
    page_to:   indice 1-based ultima pagina (None = tutte)
    dpi:       200 per scansioni normali, 400 per VETT-RASTER
    psm:       4 (colonna singola) per documenti strutturati
    """
    tess = _find_tess()
    if not tess:
        logger.info("Tesseract non trovato — salto OCR")
        return None

    try:
        import fitz
    except ImportError:
        return None

    try:
        logger.info("OCR Tesseract su %s (dpi=%d psm=%s pag %s->%s)",
                    pdf_path.name, dpi, psm, page_from + 1, page_to or "fine")
        md_pages: list[str] = []
        elements: list[Element] = []

        with fitz.open(str(pdf_path)) as doc:
            page_count = doc.page_count
            end = min(page_to, page_count) if page_to else page_count
            pages = list(enumerate(doc))[page_from:end]
            logger.info("Elaboro %d pagine (da %d a %d di %d totali)",
                        len(pages), page_from + 1, end, page_count)

            for page_idx, page in pages:
                with tempfile.TemporaryDirectory() as tmp:
                    img_path = Path(tmp) / "page.png"
                    pix = page.get_pixmap(dpi=dpi)
                    pix.save(str(img_path))

                    # Pre-processing v15
                    processed = _preprocess_image(img_path)
                    ocr_input = processed if processed else img_path

                    out_base = Path(tmp) / "out"
                    subprocess.run(
                        [tess, str(ocr_input), str(out_base),
                         "-l", "ita+eng", "--psm", psm, "--oem", "3"],
                        check=True, capture_output=True
                    )
                    txt_path = out_base.with_suffix(".txt")
                    text = txt_path.read_text(encoding="utf-8", errors="replace").strip()

                md_pages.append(f"\n\n---\n*Pagina {page_idx + 1}*\n\n{text}")
                if text:
                    elements.append(Element(
                        type="paragraph", text=text,
                        page=page_idx,
                        bbox=BBox(0, 0, page.rect.width, page.rect.height),
                    ))

        return ExtractionResult(
            markdown="\n".join(md_pages).strip(),
            elements=elements,
            page_count=len(pages),
            engine="tesseract_v15",
        )

    except Exception as exc:
        logger.warning("Tesseract fallito: %s", exc)
        return None


def _try_docling(pdf_path: Path) -> ExtractionResult | None:
    """Prova docling (alta qualita). Ritorna None se non disponibile."""
    try:
        from docling.document_converter import DocumentConverter
    except ImportError:
        return None

    try:
        logger.info("Avvio docling su %s", pdf_path.name)
        converter = DocumentConverter()
        doc_result = converter.convert(str(pdf_path))
        doc = doc_result.document
        md = doc.export_to_markdown()
        elements: list[Element] = []

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
                page_no = getattr(p, "page_no", 0) or 0
                bbox = BBox(
                    x0=float(getattr(bbox_raw, "l", 0)) if bbox_raw else 0,
                    y0=float(getattr(bbox_raw, "t", 0)) if bbox_raw else 0,
                    x1=float(getattr(bbox_raw, "r", 0)) if bbox_raw else 0,
                    y1=float(getattr(bbox_raw, "b", 0)) if bbox_raw else 0,
                )
                page_idx = max(0, page_no - 1)
            else:
                bbox = BBox(0, 0, 0, 0)
                page_idx = 0

            elements.append(Element(
                type=elem_type, text=text.strip(),
                page=page_idx, bbox=bbox,
            ))

        return ExtractionResult(markdown=md, elements=elements, engine="docling")

    except Exception as exc:
        logger.warning("docling fallito: %s", exc)
        return None


def extract_ocr(
    pdf_path: Path,
    page_from: int = 0,
    page_to: int | None = None,
    dpi: int = 200,
    psm: str = "4",
) -> ExtractionResult:
    """
    Estrazione OCR con fallback automatico v15:
      1. docling (solo se nessun intervallo specificato)
      2. Tesseract con pre-processing v15 (PSM 4, DPI configurabile)
      3. PyMuPDF grezzo come ultima risorsa

    page_from: indice 0-based prima pagina (default 0)
    page_to:   indice 1-based ultima pagina (default None = tutte)
    dpi:       DPI rasterizzazione (200 normale, 400 per VETT-RASTER)
    psm:       Page Segmentation Mode Tesseract (default 4)
    """
    # docling solo se elabora tutto il documento
    if page_from == 0 and page_to is None:
        result = _try_docling(pdf_path)
        if result and not result.error:
            return result

    # Tesseract v15
    result = _try_tesseract(pdf_path, page_from=page_from, page_to=page_to, dpi=dpi, psm=psm)
    if result and not result.error:
        return result

    # Fallback PyMuPDF
    logger.warning("Nessun OCR disponibile per %s — uso PyMuPDF grezzo", pdf_path.name)
    from src.extractor_pymupdf import extract as pymupdf_extract
    result = pymupdf_extract(pdf_path)
    result.engine = "pymupdf_fallback"
    return result
