"""PDF Server ibrido - FastAPI v3.0"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import shutil
import tempfile
import threading
import time
import uuid
from pathlib import Path

import fitz
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"),
                    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s")
logger = logging.getLogger(__name__)

TEMP_DIR = Path(os.getenv("PDF_SERVER_TEMP_DIR", tempfile.gettempdir())) / "pdf_server"
TEMP_DIR.mkdir(parents=True, exist_ok=True)
MAX_FILE_SIZE = int(os.getenv("MAX_FILE_SIZE", str(50 * 1024 * 1024)))
OCR_TIMEOUT = float(os.getenv("OCR_TIMEOUT", "300"))

app = FastAPI(title="PDF Server ibrido", version="3.0")

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
    .lbl { display: block; font-weight: bold; color: #333; margin-bottom: 6px; margin-top: 12px; }
    .drop-zone { border: 2px dashed #1F497D; border-radius: 8px; padding: 40px;
                 text-align: center; cursor: pointer; color: #1F497D;
                 transition: background 0.2s; margin-bottom: 8px; display: block; }
    .drop-zone:hover { background: #e8f0fe; }
    .icon { font-size: 2.5rem; margin-bottom: 8px; }
    select, button, input[type=number] { width: 100%; padding: 12px; border-radius: 8px;
                     font-size: 1rem; margin-bottom: 12px; }
    select, input[type=number] { border: 1px solid #ccc; color: #333; }
    .row { display: flex; gap: 16px; margin-bottom: 6px; }
    .row > div { flex: 1; }
    .row label { display: block; font-size: 0.85rem; color: #555; margin-bottom: 4px; }
    .row input { margin-bottom: 0; }
    .hint { font-size: 0.82rem; color: #888; margin-bottom: 14px; }
    button { background: #1F497D; color: white; border: none; cursor: pointer;
             font-weight: bold; transition: background 0.2s; margin-top: 8px; }
    button:hover { background: #163a63; }
    button:disabled { background: #aaa; cursor: not-allowed; }
    #status { margin-top: 16px; padding: 14px; border-radius: 8px;
              display: none; font-size: 0.95rem; line-height: 1.6; }
    .info    { background: #e3f2fd; color: #1565c0; }
    .ok      { background: #e8f5e9; color: #2e7d32; }
    .error   { background: #ffebee; color: #c62828; }
    .warning { background: #fff8e1; color: #e65100; }
    #timer { font-size: 0.85rem; margin-top: 6px; color: #555; }
    #pbwrap { background: #dde6f0; border-radius: 4px; height: 6px;
              margin-top: 10px; overflow: hidden; display: none; }
    #pb { height: 6px; background: #1F497D; border-radius: 4px; width: 0%; transition: width 0.5s; }
    #links { margin-top: 16px; }
    #links a { display: inline-block; margin: 4px 8px 4px 0; padding: 8px 16px;
               background: #1F497D; color: white; border-radius: 6px;
               text-decoration: none; font-size: 0.9rem; }
    #links a:hover { background: #163a63; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 12px;
             font-size: 0.8rem; font-weight: bold; margin-left: 8px; }
    .bt { background: #e8f5e9; color: #2e7d32; }
    .bs { background: #fff3e0; color: #e65100; }
    .bv { background: #e3f2fd; color: #1565c0; }
  </style>
</head>
<body>
  <div class="card">
    <h1>PDF Server ibrido</h1>
    <p class="subtitle">Carica un PDF - ricevi Markdown, JSON e DOCX</p>

    <span class="lbl">File PDF</span>
    <label for="fileinput" class="drop-zone" id="dropzone">
      <div class="icon">&#128193;</div>
      <div id="filelabel">Clicca o trascina un PDF qui</div>
    </label>
    <input type="file" id="fileinput" accept=".pdf" style="display:none">

    <form id="mainform">
      <span class="lbl">Formato output</span>
      <select id="output">
        <option value="all">Tutto (Markdown + JSON + DOCX)</option>
        <option value="markdown">Solo Markdown</option>
        <option value="json">Solo JSON</option>
        <option value="docx">Solo DOCX</option>
      </select>

      <span class="lbl">Intervallo pagine (OCR)</span>
      <div class="row">
        <div>
          <label for="pfrom">Da pagina</label>
          <input type="number" id="pfrom" value="1" min="1" max="999">
        </div>
        <div>
          <label for="pto">A pagina (0 = tutte)</label>
          <input type="number" id="pto" value="3" min="0" max="999">
        </div>
      </div>
      <p class="hint">Es: da 1 a 3, poi da 4 a 6 | Consigliato su Render: blocchi di 3</p>
      <button type="submit" id="btn">Converti PDF</button>
    </form>

    <div id="status"></div>
    <div id="timer"></div>
    <div id="pbwrap"><div id="pb"></div></div>
    <div id="links"></div>
  </div>

  <script>
(function() {
  var fi = document.getElementById("fileinput");
  var fl = document.getElementById("filelabel");
  var dz = document.getElementById("dropzone");
  var frm = document.getElementById("mainform");
  var btn = document.getElementById("btn");
  var st = document.getElementById("status");
  var tm = document.getElementById("timer");
  var pbw = document.getElementById("pbwrap");
  var pb = document.getElementById("pb");
  var lk = document.getElementById("links");
  var ti = null;

  fi.addEventListener("change", function() {
    fl.textContent = (this.files && this.files[0]) ? "Selezionato: " + this.files[0].name : "Clicca o trascina un PDF qui";
  });

  dz.addEventListener("dragover", function(e) { e.preventDefault(); dz.style.background = "#e8f0fe"; });
  dz.addEventListener("dragleave", function() { dz.style.background = ""; });
  dz.addEventListener("drop", function(e) {
    e.preventDefault();
    dz.style.background = "";
    var f = e.dataTransfer.files[0];
    if (f && f.name.toLowerCase().indexOf(".pdf") !== -1) {
      var dt = new DataTransfer();
      dt.items.add(f);
      fi.files = dt.files;
      fl.textContent = "Selezionato: " + f.name;
    } else {
      showSt("Trascina un file PDF valido", "warning");
    }
  });

  function startTm(n) {
    var s = Date.now();
    var est = (n > 0 ? n : 5) * 90000;
    pbw.style.display = "block";
    pb.style.width = "0%";
    pb.style.background = "#1F497D";
    var msgs = [
      "Caricamento PDF sul server...",
      "Analisi tipo documento...",
      "Conversione pagine in immagini...",
      "OCR in corso - lettura testo...",
      "Quasi pronto...",
      "Elaborazione in corso, attendere..."
    ];
    ti = setInterval(function() {
      var el = Math.floor((Date.now() - s) / 1000);
      var mn = Math.floor(el / 60);
      var sc = el % 60;
      var ts = mn > 0 ? mn + "m " + sc + "s" : sc + "s";
      var mi = Math.min(Math.floor(el / 25), msgs.length - 1);
      tm.textContent = msgs[mi] + " - Tempo: " + ts;
      pb.style.width = Math.min((Date.now() - s) / est * 95, 95) + "%";
      if (el === 300) { showSt("Il server sta impiegando molto. Riprova con meno pagine.", "warning"); }
    }, 1000);
  }

  function stopTm(ok) {
    if (ti) { clearInterval(ti); ti = null; }
    pb.style.width = "100%";
    pb.style.background = ok ? "#2e7d32" : "#c62828";
    setTimeout(function() { pbw.style.display = "none"; tm.textContent = ""; }, 2000);
  }

  function showSt(msg, cls) {
    st.innerHTML = msg;
    st.className = cls;
    st.style.display = "block";
  }

  frm.addEventListener("submit", function(e) {
    e.preventDefault();
    var file = fi.files[0];
    var out = document.getElementById("output").value;
    var pf = Math.max(1, parseInt(document.getElementById("pfrom").value) || 1);
    var pt = parseInt(document.getElementById("pto").value) || 0;

    if (!file) { showSt("Seleziona un file PDF prima di procedere", "error"); return; }
    if (pt > 0 && pt < pf) { showSt("A pagina deve essere maggiore o uguale a Da pagina", "error"); return; }

    btn.disabled = true;
    btn.textContent = "Elaborazione in corso...";
    lk.innerHTML = "";
    showSt("Caricamento PDF...", "info");
    startTm(pt > 0 ? pt - pf + 1 : 5);

    var fd = new FormData();
    fd.append("file", file);
    fd.append("output", out);
    fd.append("page_from", pf);
    fd.append("page_to", pt);

    var ctrl = new AbortController();
    var tmo = setTimeout(function() { ctrl.abort(); }, 480000);

    fetch("/convert", { method: "POST", body: fd, signal: ctrl.signal })
      .then(function(res) {
        clearTimeout(tmo);
        return res.json().then(function(d) { return { ok: res.ok, data: d }; });
      })
      .then(function(r) {
        if (!r.ok) { stopTm(false); showSt("Errore: " + (r.data.detail || "problema sconosciuto"), "error"); return; }
        stopTm(true);
        var d = r.data;
        var tl = { TESTO: ["testo","bt"], SCANSIONE: ["scansione","bs"], VETT_RASTER: ["vett-raster","bv"], MISTO: ["misto","bt"] };
        var tp = tl[d.pdf_type] || ["sconosciuto","bt"];
        var pl = d.pages_processed ? d.pages_processed + " di " + d.total_pages + " pag." : d.page_count + " pag.";
        showSt("Conversione completata - motore: <b>" + d.engine + "</b> <span class='badge " + tp[1] + "'>" + tp[0] + "</span> - " + pl, "ok");
        var h = "";
        if (d.download_markdown) h += '<a href="' + d.download_markdown + '" download>Scarica Markdown</a>';
        if (d.download_json)     h += '<a href="' + d.download_json + '" download>Scarica JSON</a>';
        if (d.download_docx)     h += '<a href="' + d.download_docx + '" download>Scarica DOCX</a>';
        lk.innerHTML = h;
      })
      .catch(function(err) {
        stopTm(false);
        clearTimeout(tmo);
        if (err.name === "AbortError") {
          showSt("Timeout: oltre 8 minuti. Riprova con meno pagine.", "error");
        } else {
          showSt("Errore di rete: " + err.message, "error");
        }
      })
      .finally(function() {
        btn.disabled = false;
        btn.textContent = "Converti PDF";
      });
  });
})();
  </script>
</body>
</html>"""


_last_cleanup = 0.0
_cleanup_lock = asyncio.Lock()


def _extract_page_range(pdf_path: Path, page_from: int, page_to: int, extract_fn):
    if page_from > page_to:
        raise ValueError(f"Range non valido: {page_from} > {page_to}")
    temp_pdf = pdf_path.parent / f"range_{uuid.uuid4().hex[:8]}.pdf"
    try:
        with fitz.open(str(pdf_path)) as src:
            with fitz.open() as dst:
                for pn in range(page_from - 1, page_to):
                    dst.insert_pdf(src, from_page=pn, to_page=pn)
                dst.save(str(temp_pdf))
        result = extract_fn(temp_pdf)
        if not getattr(result, "page_count", 0):
            result.page_count = page_to - page_from + 1
        return result
    finally:
        if temp_pdf.exists():
            temp_pdf.unlink()


async def _cleanup_old_dirs():
    global _last_cleanup
    async with _cleanup_lock:
        now = time.time()
        if now - _last_cleanup < 60:
            return
        _last_cleanup = now
    try:
        dirs = list(TEMP_DIR.iterdir())
    except Exception:
        return
    for d in dirs:
        if d.is_dir() and (time.time() - d.stat().st_mtime) > 3600:
            try:
                await asyncio.to_thread(shutil.rmtree, d)
                logger.info("Rimossa dir: %s", d.name)
            except Exception:
                pass


@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def root():
    return HTMLResponse(content=_HTML, media_type="text/html; charset=utf-8")


@app.get("/health")
async def health():
    return {"status": "ok", "version": "3.0"}


@app.post("/convert")
async def convert(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    output: str = Form("all"),
    page_from: int = Form(1),
    page_to: int = Form(0),
):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Il file deve essere un PDF (.pdf)")
    if output not in ("all", "markdown", "json", "docx"):
        raise HTTPException(status_code=400, detail="output deve essere: all, markdown, json, docx")

    req_id = uuid.uuid4().hex[:8]
    req_dir = TEMP_DIR / req_id
    req_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = req_dir / "input.pdf"
    file_size = 0
    try:
        with open(pdf_path, "wb") as out:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                file_size += len(chunk)
                if file_size > MAX_FILE_SIZE:
                    raise HTTPException(status_code=413, detail=f"File supera il limite di {MAX_FILE_SIZE // (1024*1024)} MB")
                out.write(chunk)
        logger.info("PDF ricevuto: %s (%d bytes) - req %s", file.filename, file_size, req_id)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Errore salvataggio: {exc}")
    finally:
        await file.close()

    try:
        with fitz.open(str(pdf_path)) as doc:
            total_pages = doc.page_count
        if total_pages == 0:
            raise HTTPException(status_code=400, detail="Il PDF non contiene pagine")
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"PDF non valido o corrotto: {exc}")

    page_from = max(1, page_from)
    if page_to <= 0 or page_to > total_pages:
        page_to = total_pages
    if page_from > total_pages:
        page_from = total_pages
    if page_from > page_to:
        page_to = page_from

    from src.pdf_detector import detect, PdfType
    pdf_type = detect(pdf_path)
    logger.info("Tipo PDF: %s - req %s", pdf_type, req_id)

    try:
        from src.extractor_pymupdf import extract as pymupdf_extract, ExtractionResult, Element, BBox
    except ImportError as exc:
        raise HTTPException(status_code=500, detail=f"Modulo extractor_pymupdf mancante: {exc}")

    try:
        from src.extractor_ocr import extract_ocr
    except ImportError as exc:
        raise HTTPException(status_code=500, detail=f"Modulo extractor_ocr mancante: {exc}")

    extraction = None
    try:
        if pdf_type in (PdfType.TESTO, PdfType.MISTO):
            extraction = await asyncio.wait_for(
                asyncio.to_thread(_extract_page_range, pdf_path, page_from, page_to, pymupdf_extract),
                timeout=120.0
            )
            extraction.pdf_type = pdf_type.value

        elif pdf_type == PdfType.VETT_RASTER:
            try:
                from src.ssn_cert_extractor import SsnCertExtractor
                extractor = SsnCertExtractor(dpi=300)
                certs = await asyncio.wait_for(
                    asyncio.to_thread(extractor.extract, pdf_path, threading.Event()),
                    timeout=OCR_TIMEOUT
                )
                if certs:
                    extraction = ExtractionResult(
                        engine="ssn_cert_extractor",
                        page_count=total_pages,
                        pdf_type=pdf_type.value,
                    )
                    extraction.elements = []
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
                        md_lines.append(f"## Certificato {i+1} - {tipo}\n")
                        md_lines.append("| Campo | Valore |")
                        md_lines.append("|---|---|")
                        for fname, fval in vars(cert).items():
                            if fval:
                                label = LABELS.get(fname, fname.replace("_", " ").title())
                                md_lines.append(f"| **{label}** | {fval} |")
                                extraction.elements.append(Element(
                                    type="field", text=f"{label}: {fval}",
                                    page=0, bbox=BBox(0, 0, 0, 0)
                                ))
                        md_lines.append("")
                        if i < len(certs) - 1:
                            md_lines.append("---\n")
                    extraction.markdown = "\n".join(md_lines).strip()
            except Exception as exc:
                logger.exception("SsnCertExtractor fallito - uso fallback")
                extraction = None

            if extraction is None:
                try:
                    from src.extractor_vett_raster import extract_vett_raster_with_ocr
                    extraction = await asyncio.wait_for(
                        asyncio.to_thread(_extract_page_range, pdf_path, page_from, page_to, extract_vett_raster_with_ocr),
                        timeout=OCR_TIMEOUT
                    )
                    extraction.pdf_type = pdf_type.value
                except Exception as exc:
                    logger.exception("extract_vett_raster fallito - uso pymupdf")
                    extraction = await asyncio.wait_for(
                        asyncio.to_thread(_extract_page_range, pdf_path, page_from, page_to, pymupdf_extract),
                        timeout=120.0
                    )
                    extraction.pdf_type = pdf_type.value

        else:
            _from = max(0, page_from - 1)
            _to = page_to
            try:
                extraction = await asyncio.wait_for(
                    asyncio.to_thread(extract_ocr, pdf_path, _from, _to),
                    timeout=OCR_TIMEOUT
                )
            except asyncio.TimeoutError:
                raise HTTPException(status_code=504, detail=f"OCR troppo lento (oltre {int(OCR_TIMEOUT)}s). Riduci il numero di pagine.")
            extraction.pdf_type = pdf_type.value

    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Errore durante estrazione")
        raise HTTPException(status_code=500, detail=f"Estrazione fallita: {str(exc)}")

    if extraction is None:
        extraction = await asyncio.wait_for(
            asyncio.to_thread(_extract_page_range, pdf_path, page_from, page_to, pymupdf_extract),
            timeout=120.0
        )
        extraction.pdf_type = pdf_type.value

    page_count = getattr(extraction, "page_count", 0) or (page_to - page_from + 1)
    extraction.page_count = page_count
    error = getattr(extraction, "error", None)
    engine = getattr(extraction, "engine", "sconosciuto")
    pages_processed = getattr(extraction, "page_count", 0)

    if error:
        logger.warning("Estrazione con errore: %s", error)

    response: dict = {
        "pdf_type": pdf_type.value,
        "engine": engine,
        "page_count": pages_processed,
        "pages_processed": pages_processed,
        "total_pages": total_pages,
        "filename": file.filename,
        "error": error,
    }

    if output in ("all", "markdown"):
        md_path = req_dir / "output.md"
        md_path.write_text(getattr(extraction, "markdown", "") or "", encoding="utf-8")
        response["download_markdown"] = f"/download/{req_id}/output.md"

    if output in ("all", "json"):
        elements = []
        for e in getattr(extraction, "elements", []) or []:
            if hasattr(e, "to_dict"):
                elements.append(e.to_dict())
            elif hasattr(e, "__dict__"):
                elements.append(e.__dict__)
            else:
                elements.append(e)
        json_path = req_dir / "output.json"
        payload = {
            "pdf_type": pdf_type.value, "engine": engine,
            "page_count": page_count, "filename": file.filename,
            "elements": elements,
        }
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        response["download_json"] = f"/download/{req_id}/output.json"

    if output in ("all", "docx"):
        try:
            from src.docx_writer import to_docx
            docx_path = req_dir / "output.docx"
            to_docx(extraction, docx_path)
            response["download_docx"] = f"/download/{req_id}/output.docx"
        except Exception as exc:
            response["docx_error"] = str(exc)
            logger.warning("DOCX non generato: %s", exc)

    background_tasks.add_task(_cleanup_old_dirs)
    return JSONResponse(response)


@app.get("/download/{req_id}/{filename}")
async def download(req_id: str, filename: str):
    safe_req = Path(req_id).name
    safe_file = Path(filename).name
    if not safe_req or not safe_file:
        raise HTTPException(status_code=400, detail="Percorso non valido")
    file_path = (TEMP_DIR / safe_req / safe_file).resolve()
    if not str(file_path).startswith(str(TEMP_DIR.resolve())):
        raise HTTPException(status_code=400, detail="Percorso non valido")
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File non trovato o scaduto")
    media_types = {
        ".md": "text/markdown",
        ".json": "application/json",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    return FileResponse(str(file_path), media_type=media_types.get(file_path.suffix, "application/octet-stream"), filename=safe_file)
