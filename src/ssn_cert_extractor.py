"""
Estrattore dedicato per certificati di malattia SSN/INPS — v2.

Coordinate calibrate su modulo INPS ufficiale (A4, 595x842pt).
Strategia: OCR su righe intere → estrazione con regex robusti.
Preprocessing: contrast 3x + upscale 4x + binarizzazione.
"""
from __future__ import annotations

import io
import logging
import os
import re
import shutil
import subprocess
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path

logger = logging.getLogger(__name__)

_DPI       = 300
_UPSCALE   = 4
_BIN_THR   = 140   # soglia binarizzazione (0-255)


# ── Ricerca Tesseract ─────────────────────────────────────────────────────────
def _find_tesseract() -> str | None:
    tess = shutil.which("tesseract")
    if tess:
        return tess
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


# ── Struttura risultato ───────────────────────────────────────────────────────
@dataclass
class CertificatoMalattia:
    protocollo:       str = ""
    data_rilascio:    str = ""
    medico:           str = ""
    cod_reg:          str = ""
    asl:              str = ""
    dal:              str = ""
    al:               str = ""
    diagnosi:         str = ""
    cognome:          str = ""
    nome:             str = ""
    codice_fiscale:   str = ""
    data_nascita:     str = ""
    comune_residenza: str = ""
    provincia:        str = ""
    indirizzo:        str = ""

    def to_dict(self) -> dict:
        return {k: v for k, v in self.__dict__.items() if v}

    def is_valid(self) -> bool:
        return bool(self.cognome or self.protocollo or self.dal)


# ── Estrattore ────────────────────────────────────────────────────────────────
class SsnCertExtractor:

    def __init__(self, dpi: int = _DPI) -> None:
        self._dpi  = dpi
        self._tess = _find_tesseract()
        if not self._tess:
            logger.warning("Tesseract non trovato — OCR non disponibile")

    # ── API pubblica ──────────────────────────────────────────────────────────
    def extract(
        self,
        pdf_path: Path,
        cancel_event: threading.Event | None = None,
    ) -> list[CertificatoMalattia]:
        if not self._tess:
            return []
        try:
            import fitz
            from PIL import Image, ImageEnhance
        except ImportError as exc:
            logger.error("Dipendenza mancante: %s", exc)
            return []

        results: list[CertificatoMalattia] = []
        with fitz.open(str(pdf_path)) as doc:
            for page_idx in range(doc.page_count):
                if cancel_event and cancel_event.is_set():
                    break
                page  = doc.load_page(page_idx)
                w, h  = page.rect.width, page.rect.height
                pix   = page.get_pixmap(dpi=self._dpi)
                img   = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
                img   = ImageEnhance.Contrast(img).enhance(3.0)
                scale = img.width / w

                n_forms = _count_forms(page)
                offsets = [0.0] if n_forms == 1 else [0.0, h / 2]
                for off in offsets:
                    cert = self._extract_form(img, scale, off)
                    if cert.is_valid():
                        results.append(cert)
        return results

    # ── Estrazione singolo modulo ─────────────────────────────────────────────
    def _extract_form(self, img, scale: float, off: float) -> CertificatoMalattia:
        cert = CertificatoMalattia()

        def ocr(x0, y0, x1, y1, psm=6):
            return _ocr(img, x0, y0, x1, y1, scale, off, psm, self._tess)

        # ── Protocollo e data rilascio (riga y=47-65) ─────────────────────────
        # Usa x=100-280 per evitare la label "Protocollo (*)" a sinistra
        r2 = ocr(100, 47, 280, 65, psm=7)
        cert.protocollo    = _first_number(r2, min_digits=7)
        # Data rilascio nella stessa riga ma più a destra
        r2b = ocr(300, 47, 566, 65, psm=7)
        cert.data_rilascio = _first_date(r2b) or _first_date(r2)

        # ── Medico (riga y=83-101) ────────────────────────────────────────────
        r4 = ocr(30, 83, 566, 101, psm=6)
        # Il nome medico precede "Opera nel ruolo"
        r4c = re.sub(r"(?i)(cognome\s+e\s+nome|nome|cognome)\s*[|:_\s]\s*", "", r4).strip()
        # Prende solo le prime 2-3 parole (cognome + nome medico)
        # si ferma a "Opera", "Cod", "cod", "ASL", cifre
        m_med = re.match(
            r"([A-ZÀÈÉÌÒÙ][a-zA-ZÀ-ÿ\-']+(?:\s+[A-ZÀÈÉÌÒÙ][a-zA-ZÀ-ÿ\-']+){0,2})"
            r"(?=\s+(?:Opera|[Cc]od|ASL|asl|\d)|\s*[|]|$)",
            r4c.strip("'\"| "),
        )
        if m_med:
            cert.medico = m_med.group(1).strip()
        m_cr = re.search(r"[Cc]od\.?\s*[Rr]eg\.?\s*[|:_]?\s*'?(\d+)", r4)
        if m_cr:
            cert.cod_reg = m_cr.group(1)
        m_asl = re.search(r"[Aa][Ss][Ll]\s*[|:_]?\s*(\d+)", r4)
        if m_asl:
            cert.asl = m_asl.group(1)

        # ── Date malattia (riga y=115-133) ────────────────────────────────────
        r6 = ocr(30, 115, 566, 133, psm=6)
        dates = _all_dates(r6)
        if len(dates) >= 1:
            cert.dal = dates[0]
        if len(dates) >= 2:
            cert.al = dates[1]
        if not cert.al:
            cert.al = _first_date(ocr(398, 115, 566, 133, psm=7))

        # ── Diagnosi (righe y=169-228) ────────────────────────────────────────
        r9 = ocr(30, 169, 566, 228, psm=6)
        r9c = re.sub(r"(?i)diagnosi\s*[|:]?\s*", "", r9)
        r9c = re.sub(r"(?i)cod\.?\s*nosologico.*?(?=[A-Z]{4})", "", r9c)
        r9c = re.sub(r"(?i)la\s+malattia.*?traumatico\s*\S*\s*", "", r9c)
        r9c = re.sub(r"(?i)visita.*?(?:domiciliare|ambulatori\w*)\s*\S*\s*", "", r9c)
        r9c = re.sub(r"(?i)patologia\s+grave.*", "", r9c)
        cert.diagnosi = _clean(r9c)

        # ── Dati lavoratore (riga y=258-276) ──────────────────────────────────
        r_lav = ocr(30, 258, 566, 276, psm=6)

        # Codice fiscale: OCR sbaglia O→0, l/I→1 — correggiamo
        cf_raw = re.search(
            r"\b([A-Z]{6}[0-9OlI]{2}[A-Z][0-9OlI]{2}[A-Z][0-9OlI]{3}[A-Z])\b",
            r_lav.upper(),
        )
        if cf_raw:
            cf = cf_raw.group(1)
            # Correzione OCR nelle posizioni numeriche del CF (pos 6-7, 9-10, 12-14)
            cf = _fix_cf_ocr(cf)
            cert.codice_fiscale = cf

        m_cn = re.search(
            r"[Cc]ognome\s+([A-ZÀÈÉÌÒÙ][A-ZÀ-Ÿa-zà-ÿ]+)\s+[Nn]ome\s+([A-ZÀÈÉÌÒÙ][A-ZÀ-Ÿa-zà-ÿ]+)",
            r_lav,
        )
        if m_cn:
            cert.cognome = m_cn.group(1).upper()
            cert.nome    = m_cn.group(2).upper()
        else:
            mc = re.search(r"[Cc]ognome\s+([A-Z]{3,})", r_lav)
            mn = re.search(r"[Nn]ome\s+([A-Z]{3,})", r_lav)
            if mc:
                cert.cognome = mc.group(1)
            if mn:
                cert.nome = mn.group(1)

        # ── Nascita e comune residenza (riga y=276-294) ───────────────────────
        r_res1 = ocr(30, 276, 566, 294, psm=6)
        cert.data_nascita = _first_date(r_res1)
        # Comune: compare dopo "Stato estero" o "estero"
        m_com = re.search(r"[Ss]tato\s+estero\s+([A-ZÀÈÉÌÒÙ][A-ZÀ-Ÿa-zà-ÿ]+)", r_res1)
        if m_com:
            cert.comune_residenza = m_com.group(1)
        m_prov = re.search(r"[Pp]rovincia\s+([A-Z]{2})\b", r_res1)
        if m_prov:
            cert.provincia = m_prov.group(1)

        # ── Indirizzo (riga y=294-312) ────────────────────────────────────────
        r_res2 = ocr(30, 294, 566, 312, psm=6)
        m_via = re.search(r"[Ii]n\s+via/piazza\s*[_|]?\s*(.+?)(?:\s*\||\s*$)", r_res2)
        if m_via:
            cert.indirizzo = _clean(m_via.group(1))
        m_cap = re.search(r"\b(\d{5})\b", r_res2)
        if m_cap:
            cert.cap = m_cap.group(1)  # type: ignore[attr-defined]

        return cert


# ── Funzioni di supporto ──────────────────────────────────────────────────────

def _count_forms(page) -> int:
    drawings = page.get_drawings()
    h = page.rect.height
    mid = h / 2
    top = sum(1 for d in drawings if d.get("type") == "s"
              and d.get("rect") is not None and d["rect"].y0 < mid)
    bot = sum(1 for d in drawings if d.get("type") == "s"
              and d.get("rect") is not None and d["rect"].y0 >= mid)
    if top >= 3 and bot >= 3 and abs(top - bot) / max(top, bot) < 0.4:
        return 2
    return 1


def _ocr(img, x0, y0, x1, y1, scale, y_offset, psm, tess_path) -> str:
    from PIL import Image
    ay0 = y0 + y_offset
    ay1 = y1 + y_offset
    px0 = max(0, int(x0 * scale) - 4)
    py0 = max(0, int(ay0 * scale) - 4)
    px1 = int(x1 * scale) + 4
    py1 = int(ay1 * scale) + 4
    crop = img.crop((px0, py0, px1, py1))
    nw = max(crop.width * _UPSCALE, 400)
    nh = max(crop.height * _UPSCALE, 80)
    crop = crop.resize((nw, nh), Image.LANCZOS)
    crop = crop.point(lambda px: 0 if px < _BIN_THR else 255)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
        crop.save(f.name)
        fname = f.name
    try:
        out = subprocess.run(
            [tess_path, fname, "stdout", "-l", "ita+eng",
             "--psm", str(psm), "--oem", "3"],
            capture_output=True, text=True, timeout=15,
        )
        return " ".join(out.stdout.strip().split())
    except Exception as exc:
        logger.warning("OCR timeout/errore: %s", exc)
        return ""
    finally:
        try:
            os.unlink(fname)
        except OSError:
            pass


def _fix_cf_ocr(cf: str) -> str:
    """Corregge errori OCR tipici nel codice fiscale.
    Posizioni numeriche nel CF: 6,7 (anno), 9,10 (giorno), 12,13,14 (comune).
    """
    CF_NUM_POS = {6, 7, 9, 10, 12, 13, 14}
    result = []
    for i, c in enumerate(cf):
        if i in CF_NUM_POS:
            c = c.replace("O", "0").replace("I", "1").replace("L", "1").replace("Z", "2").replace("S", "5")
        result.append(c)
    return "".join(result)


def _first_number(text: str, min_digits: int = 6) -> str:
    nums = re.findall(r"\d{" + str(min_digits) + r",}", text)
    return nums[0] if nums else ""


def _first_date(text: str) -> str:
    m = re.search(r"\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4}", text)
    return m.group().replace("-", "/").replace(".", "/") if m else ""


def _all_dates(text: str) -> list[str]:
    return [d.replace("-", "/").replace(".", "/")
            for d in re.findall(r"\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4}", text)]


def _clean(text: str) -> str:
    text = re.sub(r"[^\w\s/\-\.,àèéìòùÀÈÉÌÒÙ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return " ".join(t for t in text.split() if len(t) > 1 or t.isalpha())
