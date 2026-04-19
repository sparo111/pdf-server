"""
PDF Server ibrido — FastAPI
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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="PDF Server ibrido",
    description="Converte PDF in Markdown, JSON strutturato e DOCX.",
    version="1.1",
)

TEMP_DIR = Path(tempfile.gettempdir()) / "pdf_server"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

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
    label.block { display: block; font-weight: bold; color: #333; margin-bottom: 6px; }
    .drop-zone { border: 2px dashed #1F497D; border-radius: 8px; padding: 40px;
                 text-align: center; cursor: pointer; color: #1F497D;
                 transition: background 0.2s; margin-bottom: 20px; display: block; }
    .drop-zone:hover { background: #e8f0fe; }
    .drop-zone .icon { font-size: 2.5rem; margin-bottom: 8px; }
    select, button, input[type=number] { width: 100%; padding: 12px; border-radius: 8px;
                     font-size: 1rem; margin-bottom: 12px; }
    select, input[type=number] { border: 1px solid #ccc; color: #333; }
    .row { display: flex; gap: 16px; margin-bottom: 6px; }
    .row > div { flex: 1; }
    .row label { display: block; font-size: 0.85rem; color: #555; margin-bottom: 4px; }
    .row input { margin-bottom: 0; }
    .hint { font-size: 0.82rem; color: #888; margin-bottom: 14px; }
    button { background: #1F497D; color: white; border: none;
             cursor: pointer; font-weight: bold; transition: background 0.2s; margin-top: 8px; }
    button:hover { background: #163a63; }
    button:disabled { background: #aaa; cursor: not-allowed; }
    #status { margin-top: 16px; padding: 14px; border-radius: 8px;
              display: none; font-size: 0.95rem; line-height: 1.6; }
    .info    { background: #e3f2fd; color: #1565c0; }
    .ok      { background: #e8f5e9; color: #2e7d32; }
    .error   { background: #ffebee; color: #c62828; }
    .warning { background: #fff8e1; color: #e65100; }
    #timer { font-size: 0.85rem; margin-top: 6px; color: #555; }
    #progress-bar-wrap { background: #dde6f0; border-radius: 4px; height: 6px; margin-top: 10px; overflow: hidden; display: none; }
    #progress-bar { height: 6px; background: #1F497D; border-radius: 4px; width: 0%; transition: width 0.5s; }
    #result-links { margin-top: 16px; }
    #result-links a { display: inline-block; margin: 4px 8px 4px 0;
                      padding: 8px 16px; background: #1F497D; color: white;
                      border-radius: 6px; text-decoration: none; font-size: 0.9rem; }
    #result-links a:hover { background: #163a63; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 12px;
             font-size: 0.8rem; font-weight: bold; margin-left: 8px; }
    .badge-testo  { background: #e8f5e9; color: #2e7d32; }
    .badge-scan   { background: #fff3e0; color: #e65100; }
    .badge-vett   { background: #e3f2fd; color: #1565c0; }
  </style>
</head>
<body>
  <div class="card">
    <h1>&#128196; PDF Server ibrido</h1>
    <p class="subtitle">Carica un PDF — ricevi Markdown, JSON e DOCX</p>

    <label class="block">File PDF</label>
    <label for="file" class="drop-zone" id="drop">
      <div class="icon">&#128193;</div>
      <div id="drop-label">Clicca o trascina un PDF qui</div>
    </label>
    <input type="file" id="file" accept=".pdf" style="display:none">

    <form id="form">
      <label class="block" for="output">Formato output</label>
      <select id="output" name="output">
        <option value="all">Tutto (Markdown + JSON + DOCX)</option>
        <option value="markdown">Solo Markdown</option>
        <option value="json">Solo JSON</option>
        <option value="docx">Solo DOCX</option>
      </select>

      <label class="block">Intervallo pagine da elaborare (OCR)</label>
      <div class="row">
        <div>
          <label for="page_from">Da pagina</label>
          <input type="number" id="page_from" value="1" min="1" max="999">
        </div>
        <div>
          <label for="page_to">A pagina (0 = fino alla fine)</label>
          <input type="number" id="page_to" value="3" min="0" max="999">
        </div>
      </div>
      <p class="hint">Es: da 1 a 3, poi da 4 a 6, ecc. &nbsp;|&nbsp; Consigliato su Render: blocchi di 3 pagine</p>

      <button type="submit" id="btn">Converti PDF</button>
    </form>

    <div id="status"></div>
    <div id="timer"></div>
    <div id="progress-bar-wrap"><div id="progress-bar"></div></div>
    <div id="result-links"></div>
  </div>

  <script>
    // Aggiorna etichetta quando si sceglie un file
    document.getElementById('file').addEventListener('change', function() {
      const label = document.getElementById('drop-label');
      label.textContent = this.files[0] ? '\u2705 ' + this.files[0].name : 'Clicca o trascina un PDF qui';
    });

    // Drag & drop
    const drop = document.getElementById('drop');
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.style.background = '#e8f0fe'; });
    drop.addEventListener('dragleave', () => { drop.style.background = ''; });
    drop.addEventListener('drop', e => {
      e.preventDefault(); drop.style.background = '';
      const f = e.dataTransfer.files[0];
      if (f && f.name.toLowerCase().endsWith('.pdf')) {
        const dt = new DataTransfer();
        dt.items.add(f);
        document.getElementById('file').files = dt.files;
        document.getElementById('drop-label').textContent = '\u2705 ' + f.name;
      } else {
        showStatus('\u26a0\ufe0f Trascina un file PDF valido', 'warning');
      }
    });

    let timerInterval = null;

    function startTimer(nPages) {
      const timerEl = document.getElementById('timer');
      const bar = document.getElementById('progress-bar');
      const barWrap = document.getElementById('progress-bar-wrap');
      const start = Date.now();
      const estimatedMs = (nPages > 0 ? nPages : 5) * 90000;
      barWrap.style.display = 'block';
      bar.style.width = '0%';
      const msgs = [
        '\u23f3 Caricamento PDF sul server...',
        '\ud83d\udd0d Analisi tipo documento...',
        '\ud83d\uddbc\ufe0f Conversione pagine in immagini...',
        '\ud83d\udd24 OCR in corso — lettura testo...',
        '\ud83d\udcdd Quasi pronto, finalizzazione...',
        '\u26a0\ufe0f Ci vuole un po\' con file grandi, attendi...',
      ];
      timerInterval = setInterval(() => {
        const elapsed = Math.floor((Date.now() - start) / 1000);
        const mins = Math.floor(elapsed / 60);
        const secs = elapsed % 60;
        const timeStr = mins > 0 ? mins + 'm ' + secs + 's' : secs + 's';
        const msgIdx = Math.min(Math.floor(elapsed / 25), msgs.length - 1);
        timerEl.textContent = msgs[msgIdx] + ' — Tempo trascorso: ' + timeStr;
        bar.style.width = Math.min((Date.now() - start) / estimatedMs * 95, 95) + '%';
        if (elapsed === 300) showStatus('\u26a0\ufe0f Il server sta impiegando molto. Riprova con meno pagine se non risponde.', 'warning');
      }, 1000);
    }

    function stopTimer(ok) {
      if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
      const bar = document.getElementById('progress-bar');
      const barWrap = document.getElementById('progress-bar-wrap');
      bar.style.width = '100%';
      bar.style.background = ok ? '#2e7d32' : '#c62828';
      setTimeout(() => { barWrap.style.display = 'none'; document.getElementById('timer').textContent = ''; }, 2000);
    }

    function showStatus(msg, cls) {
      const el = document.getElementById('status');
      el.innerHTML = msg;
      el.className = cls;
      el.style.display = 'block';
    }

    document.getElementById('form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const file = document.getElementById('file').files[0];
      const output = document.getElementById('output').value;
      const pageFrom = Math.max(1, parseInt(document.getElementById('page_from').value) || 1);
      const pageTo = parseInt(document.getElementById('page_to').value) || 0;
      const links = document.getElementById('result-links');
      const btn = document.getElementById('btn');

      if (!file) { showStatus('\u26a0\ufe0f Seleziona un file PDF', 'error'); return; }
      if (pageTo > 0 && pageTo < pageFrom) { showStatus('\u26a0\ufe0f "A pagina" deve essere >= "Da pagina"', 'error'); return; }

      btn.disabled = true;
      btn.textContent = 'Elaborazione in corso...';
      links.innerHTML = '';
      showStatus('Caricamento PDF...', 'info');
      startTimer(pageTo > 0 ? pageTo - pageFrom + 1 : 5);

      const fd = new FormData();
      fd.append('file', file);
      fd.append('output', output);
      fd.append('page_from', pageFrom);
      fd.append('page_to', pageTo);

      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 480000);

      try {
        const res = await fetch('/convert', { method: 'POST', body: fd, signal: controller.signal });
        clearTimeout(timeout);
        const data = await res.json();
        if (!res.ok) { stopTimer(false); showStatus('\u274c Errore: ' + (data.detail || 'problema sconosciuto'), 'error'); return; }
        stopTimer(true);
        const typeLabels = {
          TESTO: ['testo','badge-testo'], SCANSIONE: ['scansione','badge-scan'],
          VETT_RASTER: ['vettoriale-raster','badge-vett'], MISTO: ['misto','badge-testo'],
        };
        const [tl, tc] = typeLabels[data.pdf_type] || ['sconosciuto','badge-testo'];
        const pagesLabel = data.pages_processed
          ? data.pages_processed + ' di ' + data.total_pages + ' pagine elaborate'
          : data.page_count + ' pagine';
        showStatus('\u2705 Conversione completata — motore: <b>' + data.engine + '</b> <span class="badge ' + tc + '">' + tl + '</span> — ' + pagesLabel, 'ok');
        let html = '';
        if (data.download_markdown) html += '<a href="' + data.download_markdown + '" download>\u2b07 Markdown</a>';
        if (data.download_json)     html += '<a href="' + data.download_json + '" download>\u2b07 JSON</a>';
        if (data.download_docx)     html += '<a href="' + data.download_docx + '" download>\u2b07 DOCX</a>';
        links.innerHTML = html;
      } catch (err) {
        stopTimer(false);
        clearTimeout(timeout);
        if (err.name === 'AbortError') {
          showStatus('\u23f0 Timeout: oltre 8 minuti. Riprova con meno pagine.', 'error');
        } else {
          showStatus('\u274c Errore di rete: ' + err.message, 'error');
        }
      } finally {
        btn.disabled = false;
        btn.textContent = 'Converti PDF';
      }
    });
  </script>
</body>
</html>"""


@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def root():
    return HTMLResponse(content=_HTML.encode("utf-8", errors="replace").decode("utf-8"), media_type="text/html; charset=utf-8")


@app.get("/health")
async def health():
    return {"status": "ok", "version": "1.0"}


@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    output: str = Form("all"),
    page_from: int = Form(1),
    page_to: int = Form(0),
):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Il file deve essere un PDF (.pdf)")
    if output not in ("all", "markdown", "json", "docx"):
        raise HTTPException(status_code=400, detail="output deve essere: all, markdown, json, docx")

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

    from src.pdf_detector import detect, PdfType
    pdf_type = detect(pdf_path)
    logger.info("Tipo PDF: %s — req %s", pdf_type, req_id)

    from src.extractor_pymupdf import extract as pymupdf_extract
    from src.extractor_ocr import extract_ocr

    if pdf_type in (PdfType.TESTO, PdfType.MISTO):
        extraction = pymupdf_extract(pdf_path)
        extraction.pdf_type = pdf_type.value
    elif pdf_type == PdfType.VETT_RASTER:
        try:
            from src.ssn_cert_extractor import SsnCertExtractor
            import threading
            extractor = SsnCertExtractor(dpi=300)
            certs = extractor.extract(pdf_path, cancel_event=threading.Event())
            if certs:
                from src.extractor_pymupdf import ExtractionResult, Element, BBox
                extraction = ExtractionResult(engine="ssn_cert_extractor", page_count=1, pdf_type=pdf_type.value)
                LABELS = {
                    "protocollo": "Protocollo", "data_rilascio": "Data rilascio",
                    "medico": "Medico", "cod_reg": "Cod. Reg.", "asl": "ASL",
                    "dal": "Dal", "al": "Al", "diagnosi": "Diagnosi",
                    "cognome": "Cognome", "nome": "Nome", "codice_fiscale": "Codice Fiscale",
                    "data_nascita": "Data di nascita", "comune_residenza": "Comune",
                    "provincia": "Provincia", "indirizzo": "Indirizzo",
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
                            extraction.elements.append(Element(type="field", text=f"{label}: {fval}", page=0, bbox=BBox(0,0,0,0)))
                    md_lines.append("")
                    if i < len(certs) - 1:
                        md_lines.append("---\n")
                extraction.markdown = "\n".join(md_lines).strip()
            else:
                from src.extractor_vett_raster import extract_vett_raster_with_ocr
                extraction = extract_vett_raster_with_ocr(pdf_path)
                extraction.pdf_type = pdf_type.value
        except Exception as exc:
            logger.warning("SsnCertExtractor fallito (%s) — uso fallback", exc)
            from src.extractor_vett_raster import extract_vett_raster_with_ocr
            extraction = extract_vett_raster_with_ocr(pdf_path)
            extraction.pdf_type = pdf_type.value
    else:
        _from = max(0, page_from - 1)
        _to = page_to if page_to > 0 else None
        extraction = extract_ocr(pdf_path, page_from=_from, page_to=_to)
        extraction.pdf_type = pdf_type.value

    extraction.page_count = extraction.page_count or 1

    if extraction.error:
        logger.warning("Estrazione con errore: %s", extraction.error)

    try:
        import fitz as _fitz
        with _fitz.open(str(pdf_path)) as _doc:
            total_pages = _doc.page_count
    except Exception:
        total_pages = extraction.page_count

    pages_processed = min(extraction.page_count, total_pages)

    response: dict = {
        "pdf_type": pdf_type.value,
        "engine": extraction.engine,
        "page_count": pages_processed,
        "pages_processed": pages_processed,
        "total_pages": total_pages,
        "filename": file.filename,
        "error": extraction.error or None,
    }

    if output in ("all", "markdown"):
        md_path = req_dir / "output.md"
        md_path.write_text(extraction.markdown or "", encoding="utf-8")
        response["download_markdown"] = f"/download/{req_id}/output.md"

    if output in ("all", "json"):
        json_path = req_dir / "output.json"
        payload = {
            "pdf_type": pdf_type.value, "engine": extraction.engine,
            "page_count": extraction.page_count, "filename": file.filename,
            "elements": [e.__dict__ if hasattr(e, '__dict__') else e
                         for e in (extraction.to_dict().get("elements") or [])],
        }
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        response["download_json"] = f"/download/{req_id}/output.json"

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


@app.get("/download/{req_id}/{filename}")
async def download(req_id: str, filename: str):
    if ".." in req_id or ".." in filename or "/" in req_id:
        raise HTTPException(status_code=400, detail="Percorso non valido")
    file_path = TEMP_DIR / req_id / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File non trovato o scaduto")
    media_types = {
        ".md": "text/markdown",
        ".json": "application/json",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    media_type = media_types.get(file_path.suffix, "application/octet-stream")
    return FileResponse(str(file_path), media_type=media_type, filename=filename)
