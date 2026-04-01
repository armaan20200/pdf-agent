# PDF Agent — Render.com Deployment

## Deploy in 3 steps (free, no credit card needed)

1. Go to https://render.com and sign up / log in.
2. Click  New → Web Service → "Deploy an existing image or upload a repo".
   - OR: connect your GitHub repo if you push this folder there.
   - Select "Python" as the runtime.
3. Fill in:
   - Build Command:  pip install -r requirements.txt
   - Start Command:  python main.py
   - Environment variable (optional for AI features):
       OPENAI_API_KEY = your-openai-key

Render automatically assigns a PORT and the app serves everything
(frontend + PDF API) on that single URL.

## What's in this zip

- main.py / agent.py / pdf_tools.py / cleanup.py  — Python backend
- requirements.txt  — all Python dependencies
- render.yaml       — Render Blueprint (optional auto-config)
- Procfile          — fallback for Railway / Heroku
- static/           — built React frontend (served by FastAPI)

## Local testing

  pip install -r requirements.txt
  python main.py

Then open http://localhost:10000
