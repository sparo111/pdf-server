"""
Rilevatore tipo PDF — PDF Server ibrido.
Classifica un PDF in: TESTO, SCANSIONE, VETT_RASTER, MISTO, SCONOSCIUTO.
Usa PyMuPDF (fitz) — nessuna dipendenza AI.
"""
from __future__ import annotations
from enum import Enum
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

DRAWINGS_THRESHOLD = 500
TEXT_THRESHOLD = 0


class PdfType(str, Enum):
    TESTO       = "TESTO"
    SCANSIONE   = "SCANSIONE"
    VETT_RASTER = "VETT_RASTER"
    MISTO       = "MISTO"
    SCONOSCIUTO = "SCONOSCIUTO"


def detect(pdf_path: Path) -> PdfType:
    """Classifica il PDF analizzando testo, immagini e drawings per pagina."""
    try:
        import fitz
    except ImportError:
        logger.warning("PyMuPDF non installato — tipo sconosciuto")
        return PdfType.SCONOSCIUTO

    try:
        with fitz.open(str(pdf_path)) as doc:
            page_types: list[PdfType] = []
            for page in doc:
                tb  = len(page.get_text("blocks"))
                img = len(page.get_images(full=False))
                drw = len(page.get_drawings())

                if drw >= DRAWINGS_THRESHOLD and tb == 0 and img == 0:
                    page_types.append(PdfType.VETT_RASTER)
                elif img > 0 and tb == 0:
                    page_types.append(PdfType.SCANSIONE)
                elif tb > TEXT_THRESHOLD:
                    page_types.append(PdfType.TESTO)
                else:
                    page_types.append(PdfType.MISTO)

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
