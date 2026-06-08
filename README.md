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
