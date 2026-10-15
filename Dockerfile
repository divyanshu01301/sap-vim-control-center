FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    ghostscript libreoffice-writer libreoffice-calc fonts-dejavu-core \
    tesseract-ocr poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

CMD gunicorn app:app --bind 0.0.0.0:${PORT:-10000} --workers 2 --threads 4 --timeout 300
