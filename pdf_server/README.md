# PDF Server ibrido

Converte PDF in **Markdown**, **JSON strutturato** e **DOCX** tramite una semplice interfaccia web.

## Come funziona

1. Carica un PDF tramite il browser
2. Il server rileva automaticamente il tipo (testo / scansione / vettoriale-raster)
3. Sceglie il motore giusto: **PyMuPDF** per PDF con testo, **docling/Tesseract** per scansioni
4. Restituisce i file da scaricare

## Avvio in locale (Windows)

```
# 1. Apri il Prompt dei comandi nella cartella pdf_server
# 2. Installa le dipendenze
pip install -r requirements.txt

# 3. Avvia il server
uvicorn main:app --reload

# 4. Apri il browser su:
http://localhost:8000
```

## Aggiungere OCR avanzato (opzionale)

```
pip install -r requirements-ocr.txt
```
Scarica circa 2 GB di modelli AI. Al primo avvio dopo l'installazione
il server sarà più lento (scarica i modelli), poi va in cache.

## Endpoint API

| Metodo | URL | Descrizione |
|--------|-----|-------------|
| GET | `/` | Interfaccia web |
| GET | `/health` | Stato del server |
| POST | `/convert` | Converti PDF |
| GET | `/docs` | Documentazione interattiva |
| GET | `/download/{id}/{file}` | Scarica file generato |

## Deploy su Render

Vedi la guida passo passo nel documento allegato.
