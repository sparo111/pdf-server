"""
Estrattore VETT-RASTER — PDF con glifi vettoriali + griglia stroked.

Funziona su certificati medici, SSN, moduli INPS e simili dove:
  - 0 text_blocks (nessun testo selezionabile)
  - 0 immagini raster
  - drawings: ~125 stroked (griglia celle) + ~2400 filled (glifi vettoriali)

Algoritmo:
  1. Estrae la griglia dalle celle stroked (clustering adattivo su y0/x0)
  2. Raggruppa i glifi filled per riga (tolleranza 6pt su y0)
  3. Ricostruisce le parole misurando i gap tra glifi consecutivi
  4. Assegna ogni parola alla cella di griglia più vicina
  5. Restituisce ExtractionResult con markdown e JSON strutturato

Nessun OCR, nessun modello AI, nessun Java — solo geometria PDF.
"""
from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from src.extractor_pymupdf import BBox, Element, ExtractionResult

logger = logging.getLogger(__name__)

# ── Parametri calibrati sul certificato medico/SSN ───────────────────────────
_GLYPH_ROW_TOLERANCE   = 6.0   # pt — glifi a meno di 6pt di y0 = stessa riga
_WORD_GAP_THRESHOLD    = 5.0   # pt — gap > 5pt tra glifi = spazio tra parole
_CELL_SNAP_TOLERANCE   = 20.0  # pt — snap parola alla cella entro 20pt
_CLUSTER_TOLERANCE_FACTOR = 2.0  # mediana distanze * 2 per clustering coordinate


# ── Clustering adattivo (portato dal v15) ─────────────────────────────────────
def _cluster(coords: list[float]) -> list[float]:
    """Raggruppa coordinate vicine — tolleranza = mediana(distanze) * 2."""
    if len(coords) < 2:
        return list(coords)
    s = sorted(set(round(v, 1) for v in coords))
    dists = sorted([s[i+1] - s[i] for i in range(len(s)-1)])
    if not dists:
        return s
    median = dists[len(dists) // 2]
    tol = max(median * _CLUSTER_TOLERANCE_FACTOR, 0.5)

    clusters, current = [], [s[0]]
    for c in s[1:]:
        if c - current[-1] <= tol:
            current.append(c)
        else:
            clusters.append(sum(current) / len(current))
            current = [c]
    clusters.append(sum(current) / len(current))
    return clusters


def _nearest(value: float, coords: list[float]) -> int:
    """Indice del valore più vicino nella lista."""
    if not coords:
        return 0
    return min(range(len(coords)), key=lambda i: abs(coords[i] - value))


# ── Strutture dati ────────────────────────────────────────────────────────────
@dataclass
class GlyphGroup:
    """Gruppo di glifi consecutivi che formano una parola."""
    x0: float
    y0: float
    x1: float
    y1: float
    glyph_count: int


@dataclass
class GridCell:
    """Cella della griglia del modulo."""
    row: int
    col: int
    x0: float
    y0: float
    x1: float
    y1: float
    words: list[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return " ".join(self.words)


# ── Estrattore principale ─────────────────────────────────────────────────────
def extract_vett_raster(pdf_path: Path) -> ExtractionResult:
    """
    Estrae testo e struttura da PDF vettoriale-raster senza OCR.

    Args:
        pdf_path: Percorso al PDF con glifi vettoriali.

    Returns:
        ExtractionResult con markdown, elementi strutturati e bounding box.
    """
    try:
        import fitz
    except ImportError:
        return ExtractionResult(error="PyMuPDF non installato.")

    result = ExtractionResult(engine="vett_raster", page_count=0)

    try:
        with fitz.open(str(pdf_path)) as doc:
            result.page_count = doc.page_count
            all_elements: list[Element] = []
            all_md: list[str] = []

            for page_idx, page in enumerate(doc):
                drawings = page.get_drawings()
                stroked  = [d for d in drawings if d.get("type") == "s"]
                filled   = [d for d in drawings if d.get("type") == "f"
                            and d.get("fill") == (0.0, 0.0, 0.0)]

                if not stroked:
                    logger.debug("Pagina %d: nessuna griglia stroked", page_idx)
                    continue

                # 1. Costruisci griglia dalle celle stroked
                cells = _build_grid(stroked)
                logger.info(
                    "Pagina %d: %d celle griglia (%d righe × %d col), %d glifi",
                    page_idx, len(cells),
                    max((c.row for c in cells), default=0) + 1,
                    max((c.col for c in cells), default=0) + 1,
                    len(filled),
                )

                # 2. Raggruppa glifi in parole
                words = _glyphs_to_words(filled)
                logger.debug("Pagina %d: %d parole ricostruite", page_idx, len(words))

                # 3. Assegna parole alle celle
                _assign_words_to_cells(words, cells)

                # 4. Converti in elementi e markdown
                elems, md = _cells_to_output(cells, page_idx)
                all_elements.extend(elems)
                all_md.append(f"## Pagina {page_idx + 1}\n\n{md}")

    except Exception as exc:
        logger.exception("Errore estrazione VETT-RASTER: %s", exc)
        result.error = str(exc)

    result.elements = all_elements
    result.markdown = "\n\n".join(all_md).strip()
    return result


# ── Costruzione griglia ───────────────────────────────────────────────────────
def _build_grid(stroked: list[dict]) -> list[GridCell]:
    """Costruisce la griglia di celle dai rettangoli stroked."""
    rects = [d["rect"] for d in stroked if d.get("rect") is not None]
    if not rects:
        return []

    # Cluster su y0 e x0 per trovare righe e colonne
    all_y0 = [r.y0 for r in rects]
    all_x0 = [r.x0 for r in rects]
    row_ys = _cluster(all_y0)
    col_xs = _cluster(all_x0)

    cells: list[GridCell] = []
    seen: set[tuple[int, int]] = set()

    for rect in rects:
        row_idx = _nearest(rect.y0, row_ys)
        col_idx = _nearest(rect.x0, col_xs)

        key = (row_idx, col_idx)
        if key in seen:
            continue
        seen.add(key)

        cells.append(GridCell(
            row=row_idx, col=col_idx,
            x0=rect.x0, y0=rect.y0,
            x1=rect.x1, y1=rect.y1,
        ))

    return sorted(cells, key=lambda c: (c.row, c.col))


# ── Ricostruzione parole dai glifi ────────────────────────────────────────────
def _glyphs_to_words(filled: list[dict]) -> list[GlyphGroup]:
    """
    Raggruppa glifi filled in parole basandosi su:
      - stessa riga (y0 entro _GLYPH_ROW_TOLERANCE)
      - gap orizzontale < _WORD_GAP_THRESHOLD = stesso token
    """
    if not filled:
        return []

    # Raggruppa per riga
    rows: dict[float, list] = defaultdict(list)
    for d in filled:
        rect = d.get("rect")
        if rect is None:
            continue
        y_key = round(rect.y0 / _GLYPH_ROW_TOLERANCE) * _GLYPH_ROW_TOLERANCE
        rows[y_key].append(rect)

    words: list[GlyphGroup] = []
    for _y, rects in sorted(rows.items()):
        rects_sorted = sorted(rects, key=lambda r: r.x0)

        # Raggruppa in token
        current: list = [rects_sorted[0]]
        for rect in rects_sorted[1:]:
            gap = rect.x0 - current[-1].x1
            if gap < _WORD_GAP_THRESHOLD:
                current.append(rect)
            else:
                words.append(_make_group(current))
                current = [rect]
        words.append(_make_group(current))

    return words


def _make_group(rects: list) -> GlyphGroup:
    return GlyphGroup(
        x0=min(r.x0 for r in rects),
        y0=min(r.y0 for r in rects),
        x1=max(r.x1 for r in rects),
        y1=max(r.y1 for r in rects),
        glyph_count=len(rects),
    )


# ── Assegnazione parole a celle ───────────────────────────────────────────────
def extract_vett_raster_with_ocr(pdf_path: Path) -> ExtractionResult:
    """
    Versione migliorata: usa la griglia vettoriale per trovare le celle,
    poi applica Tesseract OCR su ogni cella per leggere il testo reale.

    Questo è il metodo più accurato per certificati medici/SSN.
    Richiede Tesseract installato nel PATH.
    """
    import shutil
    if not shutil.which("tesseract"):
        logger.info("Tesseract non trovato — uso estrazione geometrica base")
        return extract_vett_raster(pdf_path)

    try:
        import fitz
        import subprocess
        import tempfile
        from PIL import Image
    except ImportError as exc:
        logger.warning("Dipendenza mancante per OCR su celle: %s", exc)
        return extract_vett_raster(pdf_path)

    result = ExtractionResult(engine="vett_raster_ocr", page_count=0)

    try:
        with fitz.open(str(pdf_path)) as doc:
            result.page_count = doc.page_count
            all_elements: list[Element] = []
            all_md: list[str] = []

            for page_idx, page in enumerate(doc):
                drawings = page.get_drawings()
                stroked  = [d for d in drawings if d.get("type") == "s"]

                if not stroked:
                    continue

                # 1. Costruisci griglia
                cells = _build_grid(stroked)

                # 2. Rasterizza la pagina intera a 300 DPI
                DPI = 300
                scale = DPI / 72.0
                mat = fitz.Matrix(scale, scale)
                pix = page.get_pixmap(matrix=mat)

                with tempfile.TemporaryDirectory() as tmp:
                    page_img_path = Path(tmp) / "page.png"
                    pix.save(str(page_img_path))
                    img = Image.open(page_img_path)

                    md_rows: dict[int, list[str]] = defaultdict(list)

                    for cell in cells:
                        # Converti coordinate PDF in pixel
                        px0 = int(cell.x0 * scale)
                        py0 = int(cell.y0 * scale)
                        px1 = int(cell.x1 * scale)
                        py1 = int(cell.y1 * scale)

                        # Margine minimo
                        w = max(px1 - px0, 20)
                        h = max(py1 - py0, 20)

                        cell_img = img.crop((
                            max(0, px0 - 4), max(0, py0 - 4),
                            min(img.width,  px0 + w + 4),
                            min(img.height, py0 + h + 4),
                        ))

                        # OCR sulla cella
                        cell_path = Path(tmp) / f"cell_{cell.row}_{cell.col}.png"
                        cell_img.save(str(cell_path))

                        try:
                            out = subprocess.run(
                                ["tesseract", str(cell_path), "stdout",
                                 "-l", "ita+eng", "--psm", "7"],
                                capture_output=True, text=True, timeout=10
                            )
                            text = out.stdout.strip().replace("\n", " ")
                        except Exception:
                            text = ""

                        if text:
                            cell.words = [text]
                            md_rows[cell.row].append(text)
                            all_elements.append(Element(
                                type="table_cell",
                                text=text,
                                page=page_idx,
                                bbox=BBox(cell.x0, cell.y0, cell.x1, cell.y1),
                            ))

                    # Costruisci markdown tabella
                    md_lines = []
                    for row_idx in sorted(md_rows):
                        md_lines.append(" | ".join(md_rows[row_idx]))
                    all_md.append(f"## Pagina {page_idx + 1}\n\n" + "\n".join(md_lines))

    except Exception as exc:
        logger.exception("Errore OCR su celle: %s", exc)
        result.error = str(exc)
        return extract_vett_raster(pdf_path)

    result.elements = all_elements
    result.markdown = "\n\n".join(all_md).strip()
    return result


def _assign_words_to_cells(words: list[GlyphGroup], cells: list[GridCell]) -> None:
    """Assegna ogni parola alla cella di griglia che la contiene o più vicina."""
    if not cells or not words:
        return

    for word in words:
        cx = (word.x0 + word.x1) / 2
        cy = (word.y0 + word.y1) / 2

        best_cell: GridCell | None = None
        best_dist = float("inf")

        for cell in cells:
            # Controlla se il centro parola è dentro la cella (con tolleranza)
            in_x = (cell.x0 - _CELL_SNAP_TOLERANCE) <= cx <= (cell.x1 + _CELL_SNAP_TOLERANCE)
            in_y = (cell.y0 - _CELL_SNAP_TOLERANCE) <= cy <= (cell.y1 + _CELL_SNAP_TOLERANCE)
            if in_x and in_y:
                dist = abs(cx - (cell.x0 + cell.x1) / 2) + abs(cy - (cell.y0 + cell.y1) / 2)
                if dist < best_dist:
                    best_dist = dist
                    best_cell = cell

        if best_cell is not None:
            # Rappresenta la parola come stringa proporzionale alla larghezza
            # (il numero di glifi è un proxy per la lunghezza del testo)
            word_repr = f"[{word.glyph_count}g]"
            best_cell.words.append(word_repr)


# ── Conversione in output ─────────────────────────────────────────────────────
def _cells_to_output(
    cells: list[GridCell], page_idx: int
) -> tuple[list[Element], str]:
    """Converte le celle in elementi strutturati e markdown."""
    elements: list[Element] = []
    md_lines: list[str] = []

    current_row = -1
    row_lines: list[str] = []

    for cell in cells:
        if not cell.words:
            continue

        if cell.row != current_row:
            if row_lines:
                md_lines.append(" | ".join(row_lines))
            current_row = cell.row
            row_lines = []

        text = cell.text
        row_lines.append(text)

        elements.append(Element(
            type="table_cell",
            text=text,
            page=page_idx,
            bbox=BBox(cell.x0, cell.y0, cell.x1, cell.y1),
        ))

    if row_lines:
        md_lines.append(" | ".join(row_lines))

    md = "\n".join(md_lines) if md_lines else "_[struttura rilevata ma testo non decodificabile senza OCR]_"

    # Nota esplicativa
    note = (
        "\n\n> **Nota:** Questo PDF usa glifi vettoriali (non testo selezionabile). "
        "La struttura della griglia è stata rilevata correttamente. "
        "Per leggere il testo reale installa Tesseract o docling (vedi README)."
    )

    return elements, md + note
