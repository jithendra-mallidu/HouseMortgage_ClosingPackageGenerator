FROM python:3.11-slim

# Install LibreOffice headless for DOCX → PDF conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-common \
        fonts-liberation \
        fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and template files
COPY app.py .
COPY templates/ ./templates/

# Create cache directory
RUN mkdir -p /app/cache

# Cloud Run uses PORT env var (defaults to 8080)
ENV PORT=8080

# Run with gunicorn (production WSGI server)
CMD exec gunicorn --bind :$PORT --workers 2 --threads 4 --timeout 120 app:app
