# Atlas GEO PDF Generator — Deploy Guide

## What this app does
Paste any Atlas report URL → get a print-ready PDF deck in ~45 seconds.

---

## Deploy to Railway (recommended — free, 5 minutes)

### Step 1 — Push to GitHub
```bash
cd atlas-app
git init
git add .
git commit -m "Atlas PDF generator"
```
Create a new repo on github.com, then:
```bash
git remote add origin https://github.com/YOUR_ORG/atlas-pdf-generator.git
git push -u origin main
```

### Step 2 — Deploy on Railway
1. Go to **railway.app** and sign in with GitHub
2. Click **"New Project" → "Deploy from GitHub repo"**
3. Select your `atlas-pdf-generator` repo
4. Railway auto-detects the Dockerfile — click **Deploy**
5. Wait ~3-4 minutes for the build (LibreOffice is large)

### Step 3 — Get your URL
1. In Railway, go to your service → **Settings → Networking**
2. Click **"Generate Domain"**
3. Your app is live at `https://atlas-pdf-generator.up.railway.app`

That's it. Share that URL with your team.

---

## Deploy to Render (alternative free option)

1. Go to **render.com** → New → Web Service
2. Connect your GitHub repo
3. Set:
   - **Environment**: Docker
   - **Build Command**: *(leave empty, uses Dockerfile)*
   - **Start Command**: `node server.js`
4. Click **Deploy**

---

## Run locally

```bash
cd atlas-app
npm install
npx playwright install chromium
node server.js
```
Then open http://localhost:3000

---

## Environment variables (optional)

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT`   | `3000`  | Server port |

---

## How it works

```
User pastes URL → POST /generate
  ↓
Playwright (headless Chrome) opens Atlas report URL
  ↓
Waits for JS to render, extracts all data from DOM
  ↓
pptxgenjs builds a 5-slide branded deck
  ↓
LibreOffice converts PPTX → PDF
  ↓
PDF streams back to browser → auto-download
```

Total time: ~30–60 seconds depending on server.

---

## Adding more report sections

The data schema in `server.js` (the `normalizeData` function) maps scraped DOM 
elements to the slide builder. As Atlas adds new sections, update the selectors 
in `scrapeAtlasReport()` and add new slide functions in `buildPPTX()`.

The existing slides are:
1. Cover (brand name, key stats)
2. Brand Leaderboard (bar chart)
3. Competitor Mentions (horizontal bars)
4. AI Platform Breakdown (table with coverage bars)
5. Key Insights (2×2 card grid)

---

## Troubleshooting

**"PDF conversion failed"** — LibreOffice not found. Make sure you're using the Docker deployment, not bare Node.

**"Scraping returned empty data"** — Atlas may have updated its CSS class names. Open the URL in Chrome DevTools, inspect the elements, and update the selectors in `scrapeAtlasReport()`.

**Timeout errors** — Increase the `timeout: 60000` value in the Playwright `goto()` call.
