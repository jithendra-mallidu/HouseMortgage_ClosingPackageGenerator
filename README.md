# House Mortgage Closing Package Generator

A web application to generate state-specific mortgage closing packages that bundle the **Security Instrument**, **Promissory Note**, and **Notice of Right to Cancel** into a single PDF.

## Architecture

| Component | Technology | Hosting |
|-----------|-----------|---------|
| **Backend API** | Python / Flask | Google Cloud Run |
| **Frontend** | HTML / CSS / JS | Vercel |
| **Doc Conversion** | LibreOffice (headless) | Bundled in Docker |

## Features

- **All 50 states + DC + territories** — automatically selects the right Fannie Mae template
- **Empty or Pre-filled packages** — generate blank templates or fill with loan data
- **Auto-fill sample data** — one-click demo with realistic test data
- **Deed of Trust / Mortgage** — shows trustee fields only for DOT states
- **DOCX placeholder filling** — fills underscore blanks contextually in the legal templates

## Project Structure

```
├── backend/
│   ├── app.py                  # Flask REST API
│   ├── Dockerfile              # Cloud Run container
│   ├── requirements.txt        # Python dependencies
│   └── templates/              # Fannie Mae legal document templates
│       ├── security_instruments/
│       ├── notes/
│       ├── Notice_Of_Right_To_Cancel.pdf
│       └── Notary_Acknowledgment_California.pdf
├── frontend/
│   ├── index.html              # Static HTML
│   ├── style.css               # Styles
│   ├── app.js                  # API client logic
│   └── vercel.json             # Vercel config
└── README.md
```

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/health` | Health check |
| GET | `/api/states` | List all states + DOT states |
| GET | `/api/sample-data?state=California` | Generate sample loan data |
| POST | `/api/generate` | Generate closing package PDF |

## Local Development

### Backend
```bash
cd backend
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py
# Runs on http://localhost:8080
```

Requires **LibreOffice** for DOCX→PDF conversion:
- macOS: `brew install --cask libreoffice`
- Linux: `apt install libreoffice-writer`

### Frontend
```bash
cd frontend
# Update window.CLOSING_API_URL in index.html to 'http://localhost:8080'
python3 -m http.server 3000
# Opens on http://localhost:3000
```

## Deployment

### Backend → Cloud Run
```bash
cd backend
gcloud run deploy closing-package-api \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --memory 1Gi \
  --timeout 120
```

### Frontend → Vercel
```bash
cd frontend
vercel --prod
```

After deploying the backend, update `window.CLOSING_API_URL` in `frontend/index.html` with the Cloud Run service URL.
