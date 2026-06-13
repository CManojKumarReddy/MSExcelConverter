# ── Stage 1: build the React frontend ────────────────────────────────────────
FROM node:20-slim AS frontend
WORKDIR /build/frontend

# Install deps from the lockfile first (better layer caching).
COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci

# Build the static site → /build/frontend/dist
COPY frontend/ ./
RUN npm run build


# ── Stage 2: Python backend (serves the API + the built frontend) ─────────────
FROM python:3.12-slim AS app

# System deps:
#   tesseract-ocr            → OCR for image/PDF-scan conversion (pytesseract)
#   libgl1, libglib2.0-0     → shared libs required by opencv-python at import
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        tesseract-ocr \
        libgl1 \
        libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    MAX_CONCURRENT_CONVERSIONS=2

WORKDIR /app/backend

# Install Python deps first (cached unless requirements.txt changes).
COPY backend/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Backend source.
COPY backend/ ./

# Built frontend from stage 1 → /app/frontend/dist
# (main.py resolves this as parent.parent/frontend/dist and serves it.)
COPY --from=frontend /build/frontend/dist /app/frontend/dist

# Render injects $PORT; default to 8000 for local `docker run`.
ENV PORT=8000
EXPOSE 8000
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
