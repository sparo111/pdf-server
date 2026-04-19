"""
PDF Detector ibrido — usa il motore v15 del Convertitore File AI.
Classificazione precisa: TESTO, SCANSIONE, VETTORIALE_RASTER, MISTO, SCONOSCIUTO.
"""
from __future__ import annotations
import logging
from enum import Enum
from pathlib import Path

logger = logging.getLogger(__name__)

DRAWINGS_THRESHOLD = 500
TEXT_THRESHOLD = 0


class PdfType(str, Enum):
    TESTO = "TESTO"
    SCANSIONE = "SCANSIONE"
    VETT_RASTER = "VETT_RASTER"
    MISTO = "MISTO"
    SCONOSCIUTO = "SCONOSCIUTO"


def detect(pdf_path: Path) -> PdfType:
    """Rileva il tipo di PDF usando la logica v15."""
    try:
        import fitz
    except ImportError:
        logger.error("PyMuPDF non installato")
        return PdfType.SCONOSCIUTO

    try:
        with fitz.open(str(pdf_path)) as doc:
            page_types = [_classify_page(page) for page in doc]

        if not page_types:
            return PdfType.SCONOSCIUTO

        unique = set(page_types)
        if len(unique) == 1:
            return unique.pop()

        real = unique - {PdfType.SCONOSCIUTO}
        if len(real) == 1:
            return real.pop()

        return PdfType.MISTO

    except Exception as exc:
        logger.exception("Errore rilevamento tipo PDF: %s", exc)
        return PdfType.SCONOSCIUTO


def _classify_page(page) -> PdfType:
    text_blocks = len(page.get_text("blocks"))
    images = len(page.get_images(full=False))
    drawings = len(page.get_drawings())

    logger.debug("Pagina %d: text=%d img=%d draw=%d", page.number, text_blocks, images, drawings)

    # VETT-RASTER: molti drawings, nessun testo, nessuna immagine
    if drawings >= DRAWINGS_THRESHOLD and text_blocks == 0 and images == 0:
        return PdfType.VETT_RASTER

    # SCANSIONE: immagini raster, nessun testo
    if images > 0 and text_blocks == 0:
        return PdfType.SCANSIONE

    # TESTO: testo selezionabile
    if text_blocks > TEXT_THRESHOLD:
        return PdfType.TESTO

    return PdfType.MISTO
