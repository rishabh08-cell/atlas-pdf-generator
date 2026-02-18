# Atlas PDF Generator — Complete Setup Guide

## What You'll Have in 15 Minutes

A live web app at `https://your-app.up.railway.app` where you paste any Atlas report URL and download a print-ready PDF deck in ~60 seconds.

---

## Prerequisites

- GitHub account (free)
- Railway account (free tier — no credit card needed initially)
- 15 minutes

---

## Step 1 — Download & Unzip

1. Download `atlas-app-v3.zip` (from this conversation)
2. Unzip it anywhere on your computer
3. You should see a folder called `atlas-app/` with these files inside:
   ```
   atlas-app/
   ├── server.js          ← The backend (Playwright + PPTX + PDF)
   ├── public/
   │   └── index.html     ← The frontend UI
   ├── Dockerfile         ← Everything pre-configured
   ├── package.json
   ├── .gitignore
   └── DEPLOY.md
   ```

---

## Step 2 — Push to GitHub

### Option A — Using GitHub Desktop (easiest)
1. Open **GitHub Desktop** (download from desktop.github.com if you don't have it)
2. File → Add Local Repository → Choose the `atlas-app` folder
3. Click "Create a repository" when prompted
4. Name it: `atlas-pdf-generator`
5. Description: `Atlas GEO report URL → PDF converter`
6. Click **Publish repository** → Uncheck "Keep this code private" → Publish

### Option B — Using Terminal
```bash
cd atlas-app
git init
git add .
git commit -m "Initial commit: Atlas PDF generator"
```

Then go to github.com → Click "+" (top right) → New repository:
- Name: `atlas-pdf-generator`
- Public
- Don't initialize with README (you already have files)
- Create repository

Copy the commands GitHub shows you (they look like this):
```bash
git remote add origin https://github.com/YOUR_USERNAME/atlas-pdf-generator.git
git branch -M main
git push -u origin main
```

---

## Step 3 — Deploy on Railway

1. Go to **railway.app**
2. Click **"Login"** → Sign in with GitHub
3. Click **"New Project"**
4. Select **"Deploy from GitHub repo"**
5. **Authorize Railway** if prompted
6. Select your `atlas-pdf-generator` repo
7. Railway auto-detects the Dockerfile → Click **"Deploy Now"**

**What happens next:**
- Railway starts building (you'll see logs streaming)
- Docker builds the image (~3-4 minutes — LibreOffice is big)
- Status changes to "Active" when done

---

## Step 4 — Get Your Live URL

1. In Railway, click on your deployed service
2. Go to **Settings** tab
3. Scroll to **Networking** section
4. Click **"Generate Domain"**
5. Railway gives you a URL like: `https://atlas-pdf-generator-production-abc123.up.railway.app`

**Copy that URL — that's your live app.**

---

## Step 5 — Test It

1. Open your Railway URL in a browser
2. You should see the Atlas PDF Generator interface
3. Paste this test URL:
   ```
   https://atlas.pepper.inc/public/reports/019c666d-058d-71f8-91c3-ede064ee461f/overview
   ```
4. Click **"Generate PDF"**
5. Watch the progress (it takes ~45-60 seconds)
6. PDF downloads automatically

---

## Troubleshooting

### "Generation failed" Error

**Check the Railway logs:**
1. Railway dashboard → Your service → **Deployments** tab
2. Click the latest deployment
3. Scroll through the logs

**Common issues:**

**Issue 1: "Scraping returned empty data"**
- **Cause:** CSS selectors in `server.js` don't match Atlas's actual DOM structure
- **Fix:** 
  1. Open the Atlas URL in Chrome
  2. Right-click on the brand name → Inspect
  3. Look at the class name (e.g., `class="brand-header-title"`)
  4. Update `server.js` line ~25 to match:
     ```javascript
     '[class*="brand-header"]', '[class*="brand-title"]'
     ```
  5. Commit and push → Railway auto-redeploys

**Issue 2: "PDF conversion failed"**
- **Cause:** LibreOffice not found
- **Fix:** This shouldn't happen if using the Dockerfile. Check Railway logs for "soffice" errors.

**Issue 3: Timeout**
- **Cause:** Page loads slowly
- **Fix:** In `server.js` line ~15, increase timeout:
   ```javascript
   await page.goto(url, { waitUntil: "networkidle", timeout: 120000 }); // 2 min
   ```

---

## Updating the App

Made changes to the code? Railway auto-deploys when you push to GitHub:

```bash
cd atlas-app
# Make your changes to server.js or index.html
git add .
git commit -m "Updated scraper selectors"
git push
```

Railway detects the push and redeploys automatically (~3-4 min).

---

## How to Verify Scraping is Working

Add this to `server.js` right after line 95 (after scraping completes):

```javascript
console.log("DEBUG — Scraped data:", JSON.stringify(data, null, 2));
```

Then check Railway logs after a generation to see exactly what got extracted.

---

## Sharing with Your Team

Just give them the Railway URL. No login needed — anyone with the link can use it.

**Pro tip:** Bookmark it or add it to Slack with:
```
🔗 Generate Atlas PDFs: https://your-app.up.railway.app
```

---

## Cost

**Railway free tier includes:**
- 500 hours/month of usage
- 100GB egress
- $5 credit

This should handle 200-300 report generations per month for free. After that, Railway charges ~$5-10/month depending on usage.

---

## Optional: Custom Domain

Want `reports.pepper.inc` instead of the Railway URL?

1. In Railway → Settings → Networking
2. Click **"Custom Domain"**
3. Enter: `reports.pepper.inc`
4. Railway gives you a CNAME record
5. Add that CNAME to your DNS (Cloudflare, etc.)
6. Wait 5-10 min for DNS to propagate

---

## Getting Help

**If scraping fails on first deploy:**

1. Check Railway logs for the "Extraction results" section
2. You'll see lines like:
   ```
   ✓ Brand name found via: [class*="brand-name"]
   ⚠ No leaderboard data found
   ```
3. For any `⚠` warnings, inspect the actual Atlas page in Chrome DevTools
4. Update the selectors in `server.js` to match
5. Push changes → Railway redeploys

**Still stuck?**
- DM me the Railway logs
- Or: Open an issue on the GitHub repo

---

## What's Next

Once deployed and working:
1. Test with 2-3 different Atlas reports to verify scraping is robust
2. If certain sections aren't extracting, tune the selectors in `server.js`
3. Share the URL with your team
4. Add more slide sections as Atlas adds new data (update `buildPPTX` function)

---

## Summary

```
1. Unzip atlas-app-v3.zip
2. Push to GitHub (via GitHub Desktop or terminal)
3. Deploy on Railway (connect GitHub repo)
4. Generate domain
5. Test with an Atlas URL
6. Share with team
```

**Total time:** 10-15 minutes

**Result:** Live app that converts any Atlas URL → PDF
