"""
Estrazione testo con PyMuPDF — motore veloce per PDF con testo selezionabile.
Produce Markdown e JSON strutturato con bounding box per ogni elemento.
Nessuna dipendenza AI, nessun modello da scaricare.
"""
from __future__ import annotations
from dataclasses import dataclass, field, asdict
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


@dataclass
class BBox:
    x0: float; y0: float; x1: float; y1: float


@dataclass
class Element:
    type: str          # "heading", "paragraph", "table_row", "image"
    text: str
    page: int
    bbox: BBox
    level: int = 0     # per heading: 1=H1, 2=H2 ecc.


@dataclass
class ExtractionResult:
    markdown: str = ""
    elements: list[Element] = field(default_factory=list)
    page_count: int = 0
    pdf_type: str = ""
    engine: str = "pymupdf"
    error: str = ""

    def to_dict(self) -> dict:
        d = asdict(self)
        return d


def extract(pdf_path: Path) -> ExtractionResult:
    """Estrae testo, struttura e bounding box con PyMuPDF."""
    try:
        import fitz
    except ImportError:
        return ExtractionResult(error="PyMuPDF non installato. Esegui: pip install pymupdf")

    result = ExtractionResult()
    md_lines: list[str] = []

    try:
        with fitz.open(str(pdf_path)) as doc:
            result.page_count = doc.page_count

            for page_idx, page in enumerate(doc):
                blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
                md_lines.append(f"\n\n---\n*Pagina {page_idx + 1}*\n")

                for block in blocks:
                    if block.get("type") == 1:
                        # Blocco immagine
                        bbox = block.get("bbox", (0,0,0,0))
                        elem = Element(
                            type="image",
                            text="[immagine]",
                            page=page_idx,
                            bbox=BBox(*bbox),
                        )
                        result.elements.append(elem)
                        md_lines.append("\n![immagine]\n")
                        continue

                    # Blocco testo
                    for line_group in block.get("lines", []):
                        spans = line_group.get("spans", [])
                        if not spans:
                            continue

                        text = " ".join(s["text"].strip() for s in spans if s["text"].strip())
                        if not text:
                            continue

                        bbox = line_group.get("bbox", (0,0,0,0))
                        # Stima tipo elemento dal font size
                        max_size = max((s.get("size", 12) for s in spans), default=12)
                        flags    = spans[0].get("flags", 0) if spans else 0
                        is_bold  = bool(flags & 2**4)

                        if max_size >= 18 or (is_bold and max_size >= 14):
                            level = 1 if max_size >= 20 else 2
                            elem_type = "heading"
                            prefix = "#" * level + " "
                        else:
                            level = 0
                            elem_type = "paragraph"
                            prefix = ""

                        elem = Element(
                            type=elem_type,
                            text=text,
                            page=page_idx,
                            bbox=BBox(*bbox),
                            level=level,
                        )
                        result.elements.append(elem)
                        md_lines.append(f"{prefix}{text}\n")

    except Exception as exc:
        logger.exception("Errore estrazione PyMuPDF: %s", exc)
        result.error = str(exc)

    result.markdown = "".join(md_lines).strip()
    return result
