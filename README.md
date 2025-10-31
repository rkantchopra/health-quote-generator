# Health Quote Generator

Generate your Health Quote DOCX from an Excel file (with **Client Details** and **Premiums** sheets), including **logos above plan names** in the same cells.

## 1) VS Code Local Script
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
python -m app.cli /path/to/input.xlsx -o output/Health_Quote.docx -l logos
```

## 2) Web Server (hide backend code)
```bash
uvicorn backend.main:app --reload --port 8000
```
Serve `frontend/index.html` from the same domain (e.g., Nginx) and proxy `/generate` to `http://localhost:8000/generate`.

## 3) Docker
```bash
docker build -t health-quote .
docker run -p 8000:8000 health-quote
```

## Logos
Put your insurer logos in `./logos/` with filenames matching:
- icici_lombard.png
- niva_reassure3.png
- niva_aspire.png
- tata_aig.png
- hdfc_ergo.png
- care_health.png

(Any of `.png/.jpg/.jpeg/.webp` will work.)

## Excel
The input workbook must include sheets: **Client Details** and **Premiums**.
