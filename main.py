"""
PDF Server ibrido — FastAPI
Converte PDF in Markdown, JSON strutturato e DOCX scaricabile.

Endpoint:
  GET  /           → pagina HTML con form di upload
  GET  /health     → {"status": "ok", "version": "1.0"}
  POST /convert    → carica PDF, ricevi risultato
  GET  /docs       → documentazione interattiva automatica FastAPI
"""
from __future__ import annotations

import json
import logging
import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)
logger = logging.getLogger(__name__)

# ── App ───────────────────────────────────────────────────────────────────────
app = FastAPI(
    title="PDF Server ibrido",
    description=(
        "Converte PDF in Markdown, JSON strutturato e DOCX. "
        "Usa PyMuPDF per PDF con testo, docling/Tesseract per scansioni."
    ),
    version="1.0",
)

# Directory temporanea per i file generati
TEMP_DIR = Path(tempfile.gettempdir()) / "pdf_server"
TEMP_DIR.mkdir(parents=True, exist_ok=True)


# ── Pagina HTML con form upload ───────────────────────────────────────────────
_HTML = """<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PDF Server ibrido</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #f0f4f8; min-height: 100vh;
           display: flex; align-items: center; justify-content: center; padding: 20px; }
    .card { background: white; border-radius: 12px; padding: 40px;
            max-width: 600px; width: 100%; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
    h1 { color: #1F497D; margin-bottom: 8px; font-size: 1.6rem; }
    .subtitle { color: #666; margin-bottom: 32px; font-size: 0.95rem; }
    label { display: block; font-weight: bold; color: #333; margin-bottom: 6px; }
    .drop-zone { border: 2px dashed #1F497D; border-radius: 8px; padding: 40px;
                 text-align: center; cursor: pointer; color: #1F497D;
                 transition: background 0.2s; margin-bottom: 20px; }
    .drop-zone:hover { background: #e8f0fe; }
    .drop-zone input { display: none; }
    .drop-zone .icon { font-size: 2.5rem; margin-bottom: 8px; }
    select, button { width: 100%; padding: 12px; border-radius: 8px;
                     font-size: 1rem; margin-bottom: 12px; }
    select { border: 1px solid #ccc; color: #333; }
    button { background: #1F497D; color: white; border: none;
             cursor: pointer; font-weight: bold; transition: background 0.2s; }
    button:hover { background: #163a63; }
    button:disabled { background: #aaa; cursor: not-allowed; }
    #status { margin-top: 16px; padding: 12px; border-radius: 8px;
              display: none; font-size: 0.95rem; }
    .info  { background: #e3f2fd; color: #1565c0; }
    .ok    { background: #e8f5e9; color: #2e7d32; }
    .error { background: #ffebee; color: #c62828; }
    #result-links { margin-top: 16px; }
    #result-links a { display: inline-block; margin: 4px 8px 4px 0;
                      padding: 8px 16px; background: #1F497D; color: white;
                      border-radius: 6px; text-decoration: none; font-size: 0.9rem; }
    #result-links a:hover { background: #163a63; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 12px;
             font-size: 0.8rem; font-weight: bold; margin-left: 8px; }
    .badge-testo      { background: #e8f5e9; color: #2e7d32; }
    .badge-scan       { background: #fff3e0; color: #e65100; }
    .badge-vett       { background: #e3f2fd; color: #1565c0; }
  </style>
</head>
<body>
  <div class="card">
    <h1>📄 PDF Server ibrido</h1>
    <p class="subtitle">Carica un PDF — ricevi Markdown, JSON e DOCX</p>

    <form id="form">
      <label>File PDF</label>
      <div class="drop-zone" id="drop" onclick="document.getElementById('file').click()">
        <div class="icon">📁</div>
        <div id="drop-label">Clicca o trascina un PDF qui</div>
        <input type="file" id="file" accept=".pdf" onchange="updateLabel(this)">
      </div>

      <label for="output">Formato output</label>
      <select id="output" name="output">
        <option value="all">Tutto (Markdown + JSON + DOCX)</option>
        <option value="markdown">Solo Markdown</option>
        <option value="json">Solo JSON</option>
        <option value="docx">Solo DOCX</option>
      </select>

      <label for="max_pages">Limite pagine OCR (0 = tutte)</label>
      <select id="max_pages" name="max_pages">
        <option value="0">Tutte le pagine</option>
        <option value="1">Max 1 pagina</option>
        <option value="3" selected>Max 3 pagine (consigliato su Render)</option>
        <option value="5">Max 5 pagine</option>
        <option value="10">Max 10 pagine</option>
      </select>

      <button type="submit" id="btn">Converti PDF</button>
    </form>

    <div id="status"></div>
    <div id="result-links"></div>
  </div>

  <script>
    function updateLabel(input) {
      const label = document.getElementById('drop-label');
      label.textContent = input.files[0] ? '✅ ' + input.files[0].name : 'Clicca o trascina un PDF qui';
    }

    // Drag & drop
    const drop = document.getElementById('drop');
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.style.background = '#e8f0fe'; });
    drop.addEventListener('dragleave', () => { drop.style.background = ''; });
    drop.addEventListener('drop', e => {
      e.preventDefault(); drop.style.background = '';
      const file = e.dataTransfer.files[0];
      if (file && file.name.endsWith('.pdf')) {
        document.getElementById('file').files = e.dataTransfer.files;
        updateLabel(document.getElementById('file'));
      }
    });

    document.getElementById('form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const file = document.getElementById('file').files[0];
      const output = document.getElementById('output').value;
      const status = document.getElementById('status');
      const links  = document.getElementById('result-links');
      const btn    = document.getElementById('btn');

      if (!file) { showStatus('Seleziona un file PDF', 'error'); return; }

      btn.disabled = true;
      btn.textContent = 'Elaborazione in corso...';
      links.innerHTML = '';
      showStatus('Caricamento e analisi PDF...', 'info');

      const fd = new FormData();
      fd.append('file', file);
      fd.append('output', output);
      fd.append('max_pages', document.getElementById('max_pages').value);

      try {
        const res  = await fetch('/convert', { method: 'POST', body: fd });
        const data = await res.json();

        if (!res.ok) {
          showStatus('Errore: ' + (data.detail || 'problema sconosciuto'), 'error');
          return;
        }

        const typeLabels = {
          TESTO: ['testo','badge-testo'],
          SCANSIONE: ['scansione','badge-scan'],
          VETT_RASTER: ['vettoriale-raster (certificato)','badge-vett'],
          MISTO: ['misto','badge-testo'],
        };
        const [tl, tc] = typeLabels[data.pdf_type] || ['sconosciuto','badge-testo'];
        showStatus(
          `✅ Conversione completata — motore: <b>${data.engine}</b> ` +
          `<span class="badge ${tc}">${tl}</span> — ${data.page_count} pagine`,
          'ok'
        );

        // Bottoni download
        let html = '';
        if (data.download_markdown) html += `<a href="${data.download_markdown}" download>⬇ Markdown</a>`;
        if (data.download_json)     html += `<a href="${data.download_json}" download>⬇ JSON</a>`;
        if (data.download_docx)     html += `<a href="${data.download_docx}" download>⬇ DOCX</a>`;
        links.innerHTML = html;

      } catch (err) {
        showStatus('Errore di rete: ' + err.message, 'error');
      } finally {
        btn.disabled = false;
        btn.textContent = 'Converti PDF';
      }
    });

    function showStatus(msg, cls) {
      const el = document.getElementById('status');
      el.innerHTML = msg;
      el.className = cls;
      el.style.display = 'block';
    }
  </script>
</body>
</html>"""


# ── Endpoint root — pagina HTML ───────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def root():
    return HTMLResponse(_HTML)


# ── Endpoint health ───────────────────────────────────────────────────────────
@app.get("/health")
async def health():
    return {"status": "ok", "version": "1.0"}


# ── Endpoint convert ──────────────────────────────────────────────────────────
@app.post("/convert")
async def convert(
    file: UploadFile = File(..., description="File PDF da convertire"),
    output: str = Form("all", description="Formato: all | markdown | json | docx"),
    max_pages: int = Form(0, description="Limite pagine OCR. 0 = tutte le pagine."),
):
    """
    Carica un PDF e ricevi Markdown, JSON strutturato e/o DOCX.

    - **file**: PDF da convertire (multipart/form-data)
    - **output**: `all` | `markdown` | `json` | `docx`

    Risponde con un JSON contenente i link per scaricare i file generati.
    """
    # Validazione
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Il file deve essere un PDF (.pdf)")

    if output not in ("all", "markdown", "json", "docx"):
        raise HTTPException(status_code=400, detail="output deve essere: all, markdown, json, docx")

    # Salva il PDF in una directory temporanea dedicata a questa richiesta
    req_id  = uuid.uuid4().hex[:8]
    req_dir = TEMP_DIR / req_id
    req_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = req_dir / "input.pdf"
    try:
        content = await file.read()
        pdf_path.write_bytes(content)
        logger.info("PDF ricevuto: %s (%d bytes) — req %s", file.filename, len(content), req_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Errore salvataggio file: {exc}")

    # Rilevamento tipo
    from src.pdf_detector import detect, PdfType
    pdf_type = detect(pdf_path)
    logger.info("Tipo PDF rilevato: %s — req %s", pdf_type, req_id)

    # Estrazione in base al tipo
    from src.extractor_pymupdf import extract as pymupdf_extract
    from src.extractor_ocr import extract_ocr

    if pdf_type in (PdfType.TESTO, PdfType.MISTO):
        extraction = pymupdf_extract(pdf_path)
        extraction.pdf_type = pdf_type.value
    elif pdf_type == PdfType.VETT_RASTER:
        # PDF con glifi vettoriali + griglia — certificati medici, SSN, moduli INPS
        # Usa SsnCertExtractor che conosce le coordinate precise del modulo
        try:
            from src.ssn_cert_extractor import SsnCertExtractor
            import threading, json as _json
            extractor = SsnCertExtractor(dpi=300)
            certs = extractor.extract(pdf_path, cancel_event=threading.Event())

            if certs:
                from src.extractor_pymupdf import ExtractionResult, Element, BBox
                extraction = ExtractionResult(
                    engine="ssn_cert_extractor",
                    page_count=1,
                    pdf_type=pdf_type.value,
                )
                # Etichette leggibili per ogni campo
                LABELS = {
                    "protocollo":       "Protocollo",
                    "data_rilascio":    "Data rilascio",
                    "medico":           "Medico",
                    "cod_reg":          "Cod. Reg.",
                    "asl":              "ASL",
                    "dal":              "Dal",
                    "al":               "Al",
                    "diagnosi":         "Diagnosi",
                    "cognome":          "Cognome",
                    "nome":             "Nome",
                    "codice_fiscale":   "Codice Fiscale",
                    "data_nascita":     "Data di nascita",
                    "comune_residenza": "Comune",
                    "provincia":        "Provincia",
                    "indirizzo":        "Indirizzo",
                }
                md_lines = []
                for i, cert in enumerate(certs):
                    tipo = "Copia Lavoratore" if i == 0 else "Copia Medico"
                    md_lines.append(f"## Certificato {i+1} — {tipo}\n")
                    md_lines.append("| Campo | Valore |")
                    md_lines.append("|---|---|")
                    for fname, fval in vars(cert).items():
                        if fval:
                            label = LABELS.get(fname, fname.replace("_"," ").title())
                            md_lines.append(f"| **{label}** | {fval} |")
                            extraction.elements.append(Element(
                                type="field", text=f"{label}: {fval}",
                                page=0, bbox=BBox(0,0,0,0),
                            ))
                    md_lines.append("")
                    if i < len(certs) - 1:
                        md_lines.append("---\n")
                extraction.markdown = "\n".join(md_lines).strip()
            else:
                # Fallback: estrattore geometrico con OCR
                from src.extractor_vett_raster import extract_vett_raster_with_ocr
                extraction = extract_vett_raster_with_ocr(pdf_path)
                extraction.pdf_type = pdf_type.value

        except Exception as exc:
            logger.warning("SsnCertExtractor fallito (%s) — uso fallback", exc)
            from src.extractor_vett_raster import extract_vett_raster_with_ocr
            extraction = extract_vett_raster_with_ocr(pdf_path)
            extraction.pdf_type = pdf_type.value
    else:
        # SCANSIONE → OCR
        _max = max_pages if max_pages and max_pages > 0 else None
        extraction = extract_ocr(pdf_path, max_pages=_max)
        extraction.pdf_type = pdf_type.value

    extraction.page_count = extraction.page_count or 1

    if extraction.error:
        logger.warning("Estrazione con errore: %s", extraction.error)

    # Costruzione risposta
    response: dict = {
        "pdf_type":   pdf_type.value,
        "engine":     extraction.engine,
        "page_count": extraction.page_count,
        "filename":   file.filename,
        "error":      extraction.error or None,
    }

    # ── Markdown ──────────────────────────────────────────────────────────────
    if output in ("all", "markdown"):
        md_path = req_dir / "output.md"
        md_path.write_text(extraction.markdown or "", encoding="utf-8")
        response["download_markdown"] = f"/download/{req_id}/output.md"

    # ── JSON ──────────────────────────────────────────────────────────────────
    if output in ("all", "json"):
        json_path = req_dir / "output.json"
        payload = {
            "pdf_type":   pdf_type.value,
            "engine":     extraction.engine,
            "page_count": extraction.page_count,
            "filename":   file.filename,
            "elements":   [e.__dict__ if hasattr(e, '__dict__') else e
                           for e in (extraction.to_dict().get("elements") or [])],
        }
        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )
        response["download_json"] = f"/download/{req_id}/output.json"

    # ── DOCX ──────────────────────────────────────────────────────────────────
    if output in ("all", "docx"):
        try:
            from src.docx_writer import to_docx
            docx_path = req_dir / "output.docx"
            to_docx(extraction, docx_path)
            response["download_docx"] = f"/download/{req_id}/output.docx"
        except ImportError as exc:
            response["docx_error"] = str(exc)
            logger.warning("DOCX non generato: %s", exc)

    return JSONResponse(response)


# ── Endpoint download file ────────────────────────────────────────────────────
@app.get("/download/{req_id}/{filename}")
async def download(req_id: str, filename: str):
    """Scarica un file generato dalla conversione."""
    # Sicurezza: blocca path traversal
    if ".." in req_id or ".." in filename or "/" in req_id:
        raise HTTPException(status_code=400, detail="Percorso non valido")

    file_path = TEMP_DIR / req_id / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File non trovato o scaduto")

    media_types = {
        ".md":   "text/markdown",
        ".json": "application/json",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    media_type = media_types.get(file_path.suffix, "application/octet-stream")
    return FileResponse(str(file_path), media_type=media_type, filename=filename)
