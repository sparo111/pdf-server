"""
Rilevatore di griglia vettoriale con clustering adattivo — Convertitore File AI v14. ★

Funzione chiave:
    cluster_coords_adaptive(coords) -> list[float]
        Raggruppa coordinate usando tolleranza = median(distanze) * 0.5
        invece di una soglia fissa (3pt), rendendola robusta a PDF di scale diverse.

Classe pubblica:
    VectorGridDetector
        extract_grid(page) -> GridResult
            Estrae colonne e righe dalla griglia vettoriale di una pagina.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Clustering adattivo ★
# ---------------------------------------------------------------------------

def cluster_coords_adaptive(coords: list[float]) -> list[float]:
    """Raggruppa coordinate vicine usando una tolleranza basata sulla mediana.

    Tolleranza = median(distanze_consecutive) * 0.5
    Vantaggio rispetto a soglia fissa (3pt):
    - Robusto su PDF di scale diverse (A4, A3, DPI variabili)
    - Si adatta automaticamente alla densità della griglia
    - Gestisce variazioni di arrotondamento da software diversi

    Args:
        coords: Lista di coordinate float (es. valori x di linee verticali).

    Returns:
        Lista di coordinate rappresentative dei cluster, ordinate.
        Se l'input ha meno di 2 elementi, lo restituisce invariato.
    """
    if len(coords) < 2:
        return list(coords)

    # Deduplicazione e ordinamento con arrotondamento a 1 decimale.
    sorted_c = sorted(set(round(v, 1) for v in coords))

    distances = [
        sorted_c[i + 1] - sorted_c[i]
        for i in range(len(sorted_c) - 1)
    ]

    if not distances:
        return sorted_c

    # Tolleranza adattiva v15: mediana delle distanze * 2.0
    # Fattore 2.0 (invece di 0.5 della v14) perché serve coprire gap interni
    # ai cluster (es. [1,2] ha gap=1, mediana=1, tol=2 → li unisce correttamente)
    # mantenendo separati cluster distanti (gap >> tol).
    distances_sorted = sorted(distances)
    mid = len(distances_sorted) // 2
    median_dist = distances_sorted[mid]
    tolerance = max(median_dist * 2.0, 0.5)  # minimo 0.5pt

    logger.debug(
        "Clustering adattivo v15: median_dist=%.2f tolerance=%.2f su %d coords",
        median_dist,
        tolerance,
        len(sorted_c),
    )

    clusters: list[float] = []
    current: list[float] = [sorted_c[0]]

    for c in sorted_c[1:]:
        if c - current[-1] <= tolerance:
            current.append(c)
        else:
            clusters.append(sum(current) / len(current))
            current = [c]
    clusters.append(sum(current) / len(current))

    return clusters


# ---------------------------------------------------------------------------
# Risultato griglia
# ---------------------------------------------------------------------------

@dataclass
class GridResult:
    """Risultato dell'estrazione della griglia vettoriale.

    Args:
        col_xs: Coordinate x delle colonne (cluster adattivo).
        row_ys: Coordinate y delle righe (cluster adattivo).
        cells: Dizionario (row_idx, col_idx) → testo estratto (vuoto inizialmente).
        raw_drawings_count: Numero di drawings totali rilevati nella pagina.
    """

    col_xs: list[float] = field(default_factory=list)
    row_ys: list[float] = field(default_factory=list)
    cells: dict[tuple[int, int], str] = field(default_factory=dict)
    raw_drawings_count: int = 0


# ---------------------------------------------------------------------------
# VectorGridDetector
# ---------------------------------------------------------------------------

class VectorGridDetector:
    """Estrae la struttura a griglia da una pagina PDF con linee vettoriali.

    Analizza i rettangoli stroked (tipo 's') per ricostruire le coordinate
    di colonne e righe tramite clustering adattivo.

    Usato dalla modalità VETTORIALE del router per PDF con testo+linee puliti.
    """

    def extract_grid(self, page: Any) -> GridResult:
        """Estrae la griglia vettoriale da una pagina PyMuPDF.

        Args:
            page: Oggetto fitz.Page da analizzare.

        Returns:
            GridResult con col_xs, row_ys e raw_drawings_count.
        """
        drawings = page.get_drawings()  # type: ignore[attr-defined]
        result = GridResult(raw_drawings_count=len(drawings))

        if not drawings:
            return result

        # Raccoglie le coordinate x0, x1, y0, y1 dei rettangoli stroked.
        x_coords: list[float] = []
        y_coords: list[float] = []

        for d in drawings:
            rect = d.get("rect")
            if rect is None:
                continue
            draw_type = d.get("type", "")
            # Considera solo rettangoli stroked (griglia) — esclude filled (testo vettoriale)
            if draw_type == "s":
                x0, y0, x1, y1 = rect
                x_coords.extend([x0, x1])
                y_coords.extend([y0, y1])

        result.col_xs = cluster_coords_adaptive(x_coords)
        result.row_ys = cluster_coords_adaptive(y_coords)

        logger.debug(
            "VectorGrid: %d col, %d righe da %d drawings stroked",
            len(result.col_xs),
            len(result.row_ys),
            len(drawings),
        )

        return result

    def assign_text_to_cells(
        self,
        grid: GridResult,
        page: Any,
        scale: float = 1.0,
    ) -> None:
        """Associa il testo estratto dalla pagina alle celle della griglia.

        Args:
            grid: GridResult con col_xs e row_ys popolati.
            page: Oggetto fitz.Page da cui estrarre il testo.
            scale: Fattore di scala pixel/punto (DPI/72). Default 1.0.
        """
        if not grid.col_xs or not grid.row_ys:
            return

        blocks = page.get_text("blocks")  # type: ignore[attr-defined]

        for block in blocks:
            x0, y0, x1, y1, text, *_ = block
            if not text.strip():
                continue

            # Centro del blocco testo scalato
            cx = (x0 + x1) / 2 * scale
            cy = (y0 + y1) / 2 * scale

            col_idx = self._nearest_index(cx, grid.col_xs)
            row_idx = self._nearest_index(cy, grid.row_ys)

            key = (row_idx, col_idx)
            existing = grid.cells.get(key, "")
            grid.cells[key] = (existing + " " + text.strip()).strip()

    @staticmethod
    def _nearest_index(value: float, coords: list[float]) -> int:
        """Restituisce l'indice del coordinata più vicina a value.

        Args:
            value: Valore da cercare.
            coords: Lista di coordinate ordinate.

        Returns:
            Indice del valore più prossimo.
        """
        if not coords:
            return 0
        return min(range(len(coords)), key=lambda i: abs(coords[i] - value))
