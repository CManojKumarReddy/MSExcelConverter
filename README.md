# DocToExcel — Document to Excel Converter

A full-stack web app with a chat UI that converts documents to Excel spreadsheets.

## Stack
- **Frontend**: React 18 + Vite (port 5173)
- **Backend**: Python FastAPI (port 8000)

## Supported Input Formats
| Format | Extraction Method |
|--------|------------------|
| PDF    | pdfplumber (tables + text) |
| DOCX   | python-docx (tables + paragraphs) |
| PNG/JPG | pytesseract OCR |
| CSV    | Direct → xlsx |
| TXT    | Line parsing → xlsx |

---

## Quick Start

### 1. Backend

```bash
cd backend

# Create and activate a virtual environment (recommended)
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Start the server
uvicorn main:app --reload --port 8000
```

The API will be available at `http://localhost:8000`.

#### Optional: Tesseract OCR (for image files)
For PNG/JPG conversion, install the Tesseract OCR engine:
- **Windows**: Download installer from https://github.com/UB-Mannheim/tesseract/wiki
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt install tesseract-ocr`

---

### 2. Frontend

```bash
cd frontend

# Install Node dependencies
npm install

# Start the development server
npm run dev
```

Open `http://localhost:5173` in your browser.

---

### 3. Tests (backend)

Image→Excel conversion is covered by an end-to-end pytest suite (requires Tesseract installed):

```bash
cd backend
pip install -r requirements-dev.txt   # first time only (installs pytest)
python -m pytest                       # run all tests
python -m pytest -v                    # verbose
```

Fixtures live in `backend/tests/fixtures/`. The suite verifies the best-of-N strategy
selector on known-tricky images (dark-background two-table screenshot, coloured-header bank
statement) — see `backend/CONVERSION_NOTES.md` for the design it guards. Tests auto-skip if
Tesseract isn't installed.

---

## Admin mode + cloud OCR (optional)

Image conversion can optionally use a **cloud AI engine** instead of the built-in Tesseract OCR,
for much higher accuracy and native table structure on hard inputs (low-contrast photos,
faint-ruled statements). Two engines are supported; if both are configured, **Gemini is preferred**.

### Option A — Google Gemini (recommended free option, **no credit card**)

1. Get a free API key at **https://aistudio.google.com/apikey** (just a Google account — no card).
2. In `backend/`, copy `.env.example` to `.env` and set `GEMINI_API_KEY=...` (the `.env` is
   gitignored). Restart the backend.
   - Privacy note: on Gemini's **free** tier, Google may use your inputs to improve its products.
     Keep that in mind for bank statements / personal data.

### Option B — Azure AI Document Intelligence (requires an Azure account + card)

1. Create a **Document Intelligence** resource in the Azure portal; copy its endpoint + key.
2. Set `AZURE_DI_ENDPOINT` / `AZURE_DI_KEY` in `backend/.env`. Restart the backend.

### Using it

Press **Ctrl + M + S** in the web UI to toggle **admin mode** — an `ADMIN` badge appears and an
**"AI OCR (cloud)"** checkbox is revealed in the options bar, with a note showing which engine is
live (`✓ Gemini ready` / `✓ Azure ready` / `⚠ Not configured`). Tick it, then upload a PNG/JPG.

If no cloud engine is configured or the call fails, conversion **automatically falls back to
Tesseract** and the result message says so. Admin mode is a UI convenience only — not a security
boundary; cloud OCR runs solely when its server-side credentials are configured.

---

## Project Structure

```
MS Excel Converter/
├── backend/
│   ├── main.py            # FastAPI app — all conversion logic
│   ├── requirements.txt   # Python dependencies
│   └── outputs/           # Generated .xlsx files (auto-created)
├── frontend/
│   ├── index.html
│   ├── vite.config.js     # Vite config with API proxy
│   ├── package.json
│   └── src/
│       ├── main.jsx
│       ├── App.jsx        # Chat UI component
│       └── App.css        # All styles (dark theme)
└── README.md
```

---

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `POST` | `/api/convert` | Upload a file, receive `{ output_filename, message }` |
| `GET`  | `/api/download/{filename}` | Download a converted Excel file |
| `GET`  | `/api/files` | List all converted files |

---

## Features
- Chat-style UI — user messages on the right, bot responses on the left
- Drag & drop file upload onto anywhere in the window
- File type validation with friendly error messages
- Loading spinner during conversion
- One-click Excel download button
- Multi-sheet Excel output (one sheet per table/page for PDFs)
- Auto-styled headers (teal) with borders and auto-width columns
- Dark theme with teal (#00c896) accents
- Fully mobile responsive
