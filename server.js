const express = require("express");
const { chromium } = require("playwright");
const pptxgen = require("pptxgenjs");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const https = require("https");
const http = require("http");

const app = express();
app.use(express.json());
app.use(express.static("public"));

const TMP = path.join(__dirname, "tmp");
if (!fs.existsSync(TMP)) fs.mkdirSync(TMP);

const LOGOS_DIR = path.join(__dirname, "public", "logos");
if (!fs.existsSync(LOGOS_DIR)) fs.mkdirSync(LOGOS_DIR, { recursive: true });

// ─── Pepper brand colour palette ────────────────────────────────────────────
const C = {
  navy:      "0D007D",   // primary deep blue
  purple:    "3D35B5",   // primary mid blue
  violet:    "7B7FD4",   // primary lavender
  lilac:     "A8ABEA",   // primary light lavender
  orange:    "DA5012",   // secondary orange
  teal:      "0B7251",   // secondary green
  green:     "0E9468",   // secondary green mid
  yellow:    "F9B02A",   // secondary yellow
  white:     "FFFFFF",
  offwhite:  "F5F5F8",
  slate:     "64748B",
  lightgray: "E2E1F0",
  darkgray:  "1A1650",
};

function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 };
}

// ─── Logo helpers ────────────────────────────────────────────────────────────
// Returns a local file path to use as a logo image, or null if unavailable.
// Priority: 1) pre-committed file in public/logos/<slug>.png
//           2) auto-scraped favicon from brand domain
//           3) null (fall back to text initials)

async function fetchBrandLogo(brandName, domain) {
  const slug = brandName.toLowerCase().replace(/[^a-z0-9]/g, "");
  const exts = [".png", ".jpg", ".jpeg", ".svg"];

  // 1) Check pre-committed logo
  for (const ext of exts) {
    const p = path.join(LOGOS_DIR, slug + ext);
    if (fs.existsSync(p)) {
      console.log("  🖼  Found pre-committed logo:", p);
      return p;
    }
  }

  // 2) Try to scrape a decent logo from the brand domain
  try {
    const logoPath = await scrapeBrandLogo(domain, slug);
    if (logoPath) return logoPath;
  } catch (e) {
    console.warn("  ⚠️  Logo scrape failed:", e.message);
  }

  console.warn("  ⚠️  No logo found for", brandName, "— will use text initials");
  return null;
}

async function scrapeBrandLogo(domain, slug) {
  // Try common logo URLs then fall back to favicon
  const cleanDomain = domain.replace(/^https?:\/\//, "").replace(/\/.*$/, "");
  const candidates = [
    `https://${cleanDomain}/favicon.ico`,
    `https://logo.clearbit.com/${cleanDomain}`,
  ];

  for (const logoUrl of candidates) {
    try {
      const buf = await downloadUrl(logoUrl, 5000);
      if (buf && buf.length > 500) {
        const ext = logoUrl.endsWith(".ico") ? ".ico" : ".png";
        const destPath = path.join(TMP, slug + "-logo" + ext);
        fs.writeFileSync(destPath, buf);
        console.log("  🌐 Downloaded logo from:", logoUrl, "(", buf.length, "bytes )");
        return destPath;
      }
    } catch {}
  }
  return null;
}

function downloadUrl(url, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith("https") ? https : http;
    const req = proto.get(url, { timeout: timeoutMs, headers: { "User-Agent": "Mozilla/5.0" } }, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        return resolve(downloadUrl(res.headers.location, timeoutMs));
      }
      if (res.statusCode !== 200) return resolve(null);
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => resolve(Buffer.concat(chunks)));
    });
    req.on("error", reject);
    req.on("timeout", () => { req.destroy(); reject(new Error("timeout")); });
  });
}

// ─── Atlas tab URL resolver ──────────────────────────────────────────────────
function getAtlasTabUrls(baseUrl) {
  const root = baseUrl.replace(/\/(overview|competitors-comparison|platforms|prompts-themes).*$/, "");
  return {
    overview:    root + "/overview",
    competitors: root + "/competitors-comparison",
    platforms:   root + "/platforms",
    prompts:     root + "/prompts-themes",
  };
}

// ─── STEP 1: Scrape ─────────────────────────────────────────────────────────
async function scrapeAtlasReport(url) {
  console.log("\n🔍 Scraping Atlas report:", url);
  const browser = await chromium.launch({ args: ["--no-sandbox", "--disable-setuid-sandbox"], headless: true });
  const tabs = getAtlasTabUrls(url);
  const screenshots = {};
  try {
    for (const [name, tabUrl] of Object.entries(tabs)) {
      console.log("  📸 Capturing:", name, "→", tabUrl);
      const page = await browser.newPage();
      await page.setViewportSize({ width: 1440, height: 900 });
      try {
        await page.goto(tabUrl, { waitUntil: "networkidle", timeout: 60000 });
        await page.waitForTimeout(5000);
        const buf = await page.screenshot({ fullPage: true });
        screenshots[name] = buf.toString("base64");
        console.log("  ✅", name, "—", Math.round(buf.length / 1024), "KB");
      } catch (e) {
        console.warn("  ⚠️ Could not capture", name, ":", e.message);
      } finally {
        await page.close();
      }
    }
  } finally {
    await browser.close();
  }
  if (!screenshots.overview) throw new Error("Overview screenshot failed — aborting.");
  return screenshots;
}

// ─── STEP 2: Extract ─────────────────────────────────────────────────────────
async function extractData(screenshots) {
  console.log("  🤖 Sending screenshots to Claude Vision API...");
  const imageBlocks = Object.entries(screenshots).flatMap(([name, b64]) => [
    { type: "text", text: `## Tab: ${name}` },
    { type: "image", source: { type: "base64", media_type: "image/png", data: b64 } },
  ]);

  const MAX_RETRIES = 4;
  const RETRY_DELAYS = [3000, 8000, 20000, 40000];
  let lastErr, data;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    if (attempt > 0) {
      const delay = RETRY_DELAYS[attempt - 1];
      console.log(`  ⏳ Claude overloaded — retrying in ${delay / 1000}s (attempt ${attempt}/${MAX_RETRIES})...`);
      await new Promise(r => setTimeout(r, delay));
    }
    let res;
    try {
      res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": process.env.ANTHROPIC_API_KEY || "",
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: "claude-opus-4-5",
          max_tokens: 8192,
          system: "You are a precise data extraction expert. Extract every number, label, and data point visible in the Atlas GEO audit report screenshots. Return ONLY valid JSON, no markdown, no explanation.",
          messages: [{ role: "user", content: [
            ...imageBlocks,
            { type: "text", text: `Extract ALL data from these Atlas report tab screenshots into this exact JSON structure. Read numbers directly from the UI — never guess.
{
  "brandName": "NESTLE",
  "domain": "nestleprofessional.in",
  "totalMentions": 28,
  "totalCitations": 33,
  "avgBrandCoverage": "9.3%",
  "avgDomainCoverage": "9.0%",
  "leaderboard": [
    { "rank": 1, "name": "WMF", "mentions": 33 },
    { "rank": 2, "name": "NESTLE", "mentions": 28 },
    { "rank": 3, "name": "Jura", "mentions": 21 }
  ],
  "brandLeaderboardRank": 2,
  "competitorMentions": [
    { "name": "WMF", "percentage": 11, "mentions": 33 },
    { "name": "NESTLE", "percentage": 9, "mentions": 28 }
  ],
  "platforms": [
    { "name": "ChatGPT", "mentions": 10, "citations": 5, "brandVisibility": 3, "domainCoverage": 1 },
    { "name": "Google AI Overview", "mentions": 12, "citations": 21, "brandVisibility": 4, "domainCoverage": 6 },
    { "name": "Perplexity", "mentions": 6, "citations": 7, "brandVisibility": 2, "domainCoverage": 2 }
  ],
  "promptThemes": [
    { "theme": "24/7 Operations Efficiency & Compliance Management", "promptCount": 7, "prompts": ["best grab and go beverage solutions"] },
    { "theme": "Beverage Quality Consistency & Menu Diversification", "promptCount": 8, "prompts": ["80 beverage options single machine"] }
  ],
  "competitorVisibilityMatrix": [
    { "theme": "24/7 Operations Efficiency & Compliance Management", "brandVisibility": 7, "competitors": { "WMF": 0, "Jura": 0, "Kaapi Machines": 0, "Franke": 3 } }
  ],
  "brandVisibilityByPlatform": [
    { "theme": "24/7 Operations Efficiency & Compliance Management", "ChatGPT": 10, "Google AI Overview": 10, "Perplexity": 0 },
    { "theme": "Beverage Quality Consistency & Menu Diversification", "ChatGPT": 0, "Google AI Overview": 10, "Perplexity": 0 }
  ],
  "domainCitations": [
    { "domain": "www.reddit.com", "domainCoverage": "14%", "uniquePagesCited": 72, "domainShare": "3%" }
  ],
  "brandPages": [
    { "name": "Page title or URL", "prompts": 5 }
  ]
}
Rules:
- Extract ALL rows from every table visible.
- brandLeaderboardRank: read the brand rank number directly from leaderboard text (e.g. "#5 Netskope" = 5). Do NOT infer from array position.
- competitorVisibilityMatrix: from "Brand Visibility x Competitors" on competitors tab. Extract ALL rows and ALL competitor columns.
- brandVisibilityByPlatform: from the "[BrandName] brand visibility" table on platforms tab. Extract ALL rows visible.
- Return ONLY the JSON.` }
          ]}],
        }),
      });
    } catch (networkErr) {
      lastErr = networkErr;
      console.warn(`  ⚠️ Network error on attempt ${attempt}: ${networkErr.message}`);
      continue;
    }

    if (res.status === 529 || res.status === 503) {
      const body = await res.text();
      lastErr = new Error(`Claude API error: ${res.status} — ${body}`);
      console.warn(`  ⚠️ Claude ${res.status} on attempt ${attempt}`);
      continue;
    }
    if (!res.ok) throw new Error(`Claude API error: ${res.status} — ${await res.text()}`);

    const json = await res.json();
    const text = json.content[0].text.trim();
    try { data = JSON.parse(text); }
    catch { const m = text.match(/\{[\s\S]*\}/); if (m) data = JSON.parse(m[0]); else throw new Error("Could not parse Claude JSON response"); }
    lastErr = null;
    break;
  }

  if (!data) throw lastErr || new Error("Claude API failed after all retries");
  console.log("  Brand:", data.brandName);
  console.log("  Leaderboard:", data.leaderboard?.length, "entries");
  console.log("  Platforms:", data.platforms?.length, "entries");
  console.log("  Themes:", data.promptThemes?.length, "themes");
  return data;
}

// ─── STEP 3: Normalize ───────────────────────────────────────────────────────
function normalizeData(raw) {
  console.log("\n📊 Normalizing extracted data...");
  const tm = raw.totalMentions || 0;
  const tc = raw.totalCitations || 0;
  const platforms = raw.platforms?.length > 0 ? raw.platforms : [
    { name: "ChatGPT",           mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
    { name: "Google AI Overview",mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
    { name: "Perplexity",        mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
  ];
  const leaderboard = raw.leaderboard?.length > 0 ? raw.leaderboard : [{ rank: 1, name: raw.brandName, mentions: tm }];
  const competitorMentions = raw.competitorMentions?.length > 0 ? raw.competitorMentions
    : leaderboard.map((b, i) => ({ name: b.name, percentage: Math.max(30 - i * 8, 2), mentions: b.mentions }));
  const isLeader = leaderboard[0]?.name === raw.brandName;
  const leaderboardRank = raw.brandLeaderboardRank ? "#" + raw.brandLeaderboardRank
    : isLeader ? "#1"
    : "#" + (leaderboard.findIndex(b => b.name === raw.brandName) + 1 || leaderboard.length);
  return {
    brandName: raw.brandName || "Brand",
    domain: raw.domain || raw.brandName?.toLowerCase().replace(/\s+/g, "") + ".com",
    totalMentions: tm,
    totalCitations: tc,
    avgBrandCoverage: raw.avgBrandCoverage || "0%",
    avgDomainCoverage: raw.avgDomainCoverage || "0%",
    leaderboardRank,
    platformCount: platforms.length,
    leaderboard: leaderboard.slice(0, 3),
    competitorMentions: competitorMentions.slice(0, 10),
    platforms,
    promptThemes: raw.promptThemes?.length > 0 ? raw.promptThemes : [],
    domainCitations: (raw.domainCitations || []).slice(0, 10),
    brandPages: (raw.brandPages || []).slice(0, 8),
    competitorVisibilityMatrix: raw.competitorVisibilityMatrix?.length > 0 ? raw.competitorVisibilityMatrix : [],
    brandVisibilityByPlatform: raw.brandVisibilityByPlatform?.length > 0 ? raw.brandVisibilityByPlatform : [],
  };
}

// ─── PPTX layout helpers ─────────────────────────────────────────────────────
// logoPill: draws "BrandLogo | pepper" pill in top-right corner of any slide
// brandLogoPath: local file path or null (falls back to text initials)
function logoPill(s, pres, brandName, brandLogoPath) {
  // Pill background
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.6, y: 0.06, w: 2.25, h: 0.38,
    fill: { color: C.white }, line: { color: C.lightgray, pt: 1 }, rectRadius: 0.06,
    shadow: makeShadow(),
  });

  // Divider line between brand and pepper sections
  s.addShape(pres.shapes.RECTANGLE, {
    x: 8.72, y: 0.1, w: 0.015, h: 0.28,
    fill: { color: C.lightgray }, line: { color: C.lightgray },
  });

  // LEFT: brand logo or initials
  if (brandLogoPath && fs.existsSync(brandLogoPath)) {
    try {
      s.addImage({ path: brandLogoPath, x: 7.65, y: 0.09, w: 0.9, h: 0.28, sizing: { type: "contain", w: 0.9, h: 0.28 } });
    } catch {
      s.addText(brandName.substring(0, 4).toUpperCase(), {
        x: 7.65, y: 0.09, w: 0.9, h: 0.28, fontSize: 7, bold: true,
        color: C.navy, align: "center", valign: "middle", fontFace: "Calibri",
      });
    }
  } else {
    s.addText(brandName.substring(0, 6).toUpperCase(), {
      x: 7.65, y: 0.09, w: 0.9, h: 0.28, fontSize: 7, bold: true,
      color: C.navy, align: "center", valign: "middle", fontFace: "Calibri",
    });
  }

  // RIGHT: Pepper logo (text-based — clean & consistent)
  s.addText("🌶 pepper", {
    x: 8.74, y: 0.09, w: 1.08, h: 0.28, fontSize: 7.5, bold: true,
    color: C.orange, align: "center", valign: "middle", fontFace: "Calibri",
  });
}

function hdr(s, pres, title, brandName, brandLogoPath) {
  // Top accent bar
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.teal}, line:{color:C.teal} });
  // atlas wordmark
  s.addText("atlas", { x:0.3, y:0.12, w:0.9, h:0.3, fontSize:11, bold:true, color:C.navy, fontFace:"Calibri" });
  s.addText("by pepper.inc", { x:1.2, y:0.17, w:1.3, h:0.22, fontSize:7, color:C.slate, fontFace:"Calibri" });
  // Brand | Pepper pill top-right
  logoPill(s, pres, brandName, brandLogoPath);
  // Slide title
  s.addText(title, { x:0.3, y:0.5, w:7.2, h:0.48, fontSize:18, bold:true, color:C.navy, fontFace:"Calibri" });
  // Underline
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:0.95, w:1.8, h:0.03, fill:{color:C.teal}, line:{color:C.teal} });
}

function ftr(s, pres, brand, domain) {
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.2, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText(brand + " · " + domain + " · GEO Audit by Atlas / Pepper.inc", {
    x:0.3, y:5.43, w:9.4, h:0.18, fontSize:6.5, color:"AAAACC", fontFace:"Calibri"
  });
}

// Shared static header for slides 12-18 (white bg, brand|pepper pill)
function staticHdr(s, pres, title, brandName, brandLogoPath) {
  logoPill(s, pres, brandName, brandLogoPath);
  s.addText(title, { x:0.35, y:0.12, w:7.15, h:0.42, fontSize:18, bold:true, color:C.navy, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:0.55, w:3.5, h:0.025, fill:{color:C.navy}, line:{color:C.navy} });
}

// ─── SLIDE 1: Cover ──────────────────────────────────────────────────────────
function buildSlide1(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  // Decorative circles
  s.addShape(pres.shapes.OVAL, { x:7.5, y:-0.5, w:3.5, h:3.5, fill:{color:"150050"}, line:{color:"150050"} });
  s.addShape(pres.shapes.OVAL, { x:8.2, y:0.2, w:2.0, h:2.0, fill:{color:C.purple}, line:{color:C.purple} });

  // Top-left: atlas wordmark
  s.addText("atlas", { x:0.5, y:0.38, w:1.1, h:0.38, fontSize:15, bold:true, color:C.orange, fontFace:"Calibri" });
  s.addText("by pepper.inc", { x:1.63, y:0.44, w:1.5, h:0.26, fontSize:8, color:C.lilac, fontFace:"Calibri" });

  // Top-right: Brand | Pepper co-logo pill (white pill on dark bg)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x:7.4, y:0.28, w:2.35, h:0.46,
    fill:{color:"FFFFFF"}, line:{color:"FFFFFF"}, rectRadius:0.08,
  });
  s.addShape(pres.shapes.RECTANGLE, { x:8.56, y:0.32, w:0.015, h:0.36, fill:{color:C.lightgray}, line:{color:C.lightgray} });
  if (brandLogoPath && fs.existsSync(brandLogoPath)) {
    try {
      s.addImage({ path: brandLogoPath, x:7.44, y:0.3, w:0.95, h:0.38, sizing:{type:"contain",w:0.95,h:0.38} });
    } catch {
      s.addText(d.brandName.substring(0,5).toUpperCase(), { x:7.44, y:0.3, w:0.95, h:0.38, fontSize:7.5, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
    }
  } else {
    s.addText(d.brandName.substring(0,5).toUpperCase(), { x:7.44, y:0.3, w:0.95, h:0.38, fontSize:7.5, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
  }
  s.addText("🌶 pepper", { x:8.58, y:0.3, w:1.14, h:0.38, fontSize:8, bold:true, color:C.orange, align:"center", valign:"middle", fontFace:"Calibri" });

  // Main content
  s.addText("GEO AUDIT REPORT", { x:0.5, y:1.08, w:6, h:0.28, fontSize:9, color:C.orange, bold:true, charSpacing:4, fontFace:"Calibri" });
  s.addText(d.brandName, { x:0.5, y:1.38, w:7, h:1.18, fontSize:50, bold:true, color:C.white, fontFace:"Calibri" });
  s.addText(d.domain, { x:0.5, y:2.6, w:5, h:0.38, fontSize:13, color:C.lilac, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.02, w:1.2, h:0.04, fill:{color:C.orange}, line:{color:C.orange} });

  // KPI strip
  const kpis = [
    { v: String(d.totalMentions), l: "Total Mentions" },
    { v: d.avgBrandCoverage,      l: "Brand Coverage" },
    { v: String(d.platformCount), l: "AI Platforms" },
    { v: d.leaderboardRank,       l: "Leaderboard" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.5 + i * 2.3;
    s.addText(k.v, { x, y:3.16, w:2.1, h:0.52, fontSize:24, bold:true, color:C.orange, fontFace:"Calibri" });
    s.addText(k.l, { x, y:3.66, w:2.1, h:0.22, fontSize:8, color:C.lilac, fontFace:"Calibri" });
  });
  s.addText("Powered by atlas · pepper.inc", { x:0.5, y:5.15, w:9, h:0.22, fontSize:7.5, color:C.slate, fontFace:"Calibri" });
}

// ─── SLIDE 2: Prompts & Themes ───────────────────────────────────────────────
function buildSlide2(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  const totalPrompts = d.promptThemes.reduce((a,t)=>a+(t.promptCount||t.prompts?.length||0),0);
  hdr(s, pres, `We Have Mapped ${totalPrompts} Prompts Across ${d.promptThemes.length} Themes`, d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const themes = d.promptThemes.slice(0, 9);
  const cols = 3, colW = 3.0, colGap = 0.15, rowH = 0.72, rowGap = 0.1;
  const startX = 0.28, startY = 1.1;
  themes.forEach((t, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (colW + colGap), y = startY + row * (rowH + rowGap);
    const count = t.promptCount || t.prompts?.length || 0;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:colW, h:rowH, fill:{color:C.offwhite}, line:{color:C.lightgray}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.04, h:rowH, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText(t.theme, { x:x+0.1, y:y+0.08, w:colW-0.15, h:0.34, fontSize:9.5, bold:true, color:C.navy, fontFace:"Calibri", wrap:true });
    s.addText(count+" prompts", { x:x+0.1, y:y+0.46, w:colW-0.15, h:0.2, fontSize:8.5, color:C.slate, fontFace:"Calibri" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:2.8, y:4.78, w:4.4, h:0.36, fill:{color:C.yellow}, line:{color:C.yellow}, rectRadius:0.05 });
  s.addText("Link to the entire list of prompts", { x:2.8, y:4.78, w:4.4, h:0.36, fontSize:10, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
  s.addText("This is a relevant set of non-branded prompts for your brand.", { x:0.3, y:5.16, w:9.4, h:0.2, fontSize:7.5, color:C.slate, italic:true, align:"center", fontFace:"Calibri" });
}

// ─── SLIDE 3: Brand Leaderboard + Competitor Mentions ───────────────────────
function buildSlide3(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "Brand Leaderboard & Competitor Mentions", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const brands = d.leaderboard;
  if (brands.length > 0) {
    const maxM = Math.max(...brands.map(b => b.mentions), 1);
    const barW = 1.0, gap = 0.55, startX = 0.3, chartBottom = 4.65, chartH = 2.6;
    brands.forEach((brand, i) => {
      const x = startX + i * (barW + gap);
      const barH = Math.max((brand.mentions / maxM) * chartH, 0.15);
      const barY = chartBottom - barH;
      const isB = brand.name === d.brandName;
      const col = isB ? C.orange : (i === 0 ? C.teal : "BDBDCD");
      s.addShape(pres.shapes.OVAL, { x:x+barW/2-0.22, y:barY-0.96, w:0.44, h:0.44, fill:{color:C.white}, line:{color:C.lightgray} });
      s.addText("#"+brand.rank, { x:x+barW/2-0.22, y:barY-0.96, w:0.44, h:0.44, fontSize:11, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
      s.addText(brand.name, { x:x-0.1, y:barY-0.5, w:barW+0.2, h:0.34, fontSize:8, bold:isB, color:isB?C.orange:C.navy, align:"center", fontFace:"Calibri", wrap:true });
      s.addShape(pres.shapes.RECTANGLE, { x, y:barY, w:barW, h:barH, fill:{color:col}, line:{color:col}, shadow:makeShadow() });
      s.addText(brand.mentions+" mentions", { x:x-0.1, y:chartBottom+0.06, w:barW+0.2, h:0.2, fontSize:7.5, color:C.slate, align:"center", fontFace:"Calibri" });
      if (isB) s.addText("👑", { x:x+barW/2-0.24, y:barY+0.06, w:0.48, h:0.35, fontSize:18, align:"center" });
    });
    s.addText("Brand Leaderboard", { x:0.3, y:1.1, w:4.5, h:0.25, fontSize:9, bold:true, color:C.slate, fontFace:"Calibri" });
  }
  const comps = d.competitorMentions.slice(0, 10);
  const maxPct = Math.max(...comps.map(c => c.percentage), 1);
  const rowH = 0.33, startY = 1.1, barMaxW = 3.5;
  s.addText("Competitor Mentions vs. "+d.brandName, { x:5.0, y:1.1, w:4.7, h:0.25, fontSize:9, bold:true, color:C.slate, fontFace:"Calibri" });
  comps.forEach((comp, i) => {
    const y = startY + 0.32 + i * rowH;
    const isB = comp.name === d.brandName;
    s.addShape(pres.shapes.OVAL, { x:5.0, y:y+0.04, w:0.22, h:0.22, fill:{color:isB?C.purple:C.lightgray}, line:{color:isB?C.purple:C.lightgray} });
    s.addText(comp.name[0].toUpperCase(), { x:5.0, y:y+0.04, w:0.22, h:0.22, fontSize:7, bold:true, color:isB?C.white:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
    s.addText(comp.name, { x:5.26, y:y+0.05, w:1.4, h:0.2, fontSize:8, bold:isB, color:isB?C.purple:C.navy, fontFace:"Calibri" });
    const bw = (comp.percentage / maxPct) * barMaxW;
    s.addShape(pres.shapes.RECTANGLE, { x:6.7, y:y+0.06, w:Math.max(bw,0.05), h:0.18, fill:{color:isB?C.navy:C.lightgray}, line:{color:isB?C.navy:C.lightgray} });
    s.addText(comp.percentage+"% · "+comp.mentions+" mentions", { x:6.72+bw, y:y+0.05, w:2.5, h:0.2, fontSize:7.5, color:C.slate, fontFace:"Calibri" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:4.75, y:1.05, w:0.03, h:3.8, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 4: Top Cited Sources ──────────────────────────────────────────────
function buildSlide4(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "Top Cited Sources (Category vs Us)", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const domains = d.domainCitations.slice(0, 9);
  s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:1.08, w:4.4, h:0.28, fill:{color:C.navy}, line:{color:C.navy} });
  ["Domain","Pages","Responses"].forEach((h, i) => {
    const xs = [0.35, 2.8, 3.65];
    s.addText(h, { x:xs[i], y:1.1, w:1.1, h:0.24, fontSize:8, bold:true, color:C.white, fontFace:"Calibri" });
  });
  domains.forEach((row, i) => {
    const y = 1.38 + i * 0.34;
    const bg = i % 2 === 0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.25, y, w:4.4, h:0.32, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(row.domain, { x:0.35, y:y+0.06, w:2.4, h:0.22, fontSize:8, color:C.navy, fontFace:"Calibri" });
    s.addText(String(row.uniquePagesCited||row.pages||""), { x:2.82, y:y+0.06, w:0.6, h:0.22, fontSize:8, color:C.slate, fontFace:"Calibri" });
    s.addText(String(row.domainShare||row.responses||""), { x:3.67, y:y+0.06, w:0.9, h:0.22, fontSize:8, color:C.slate, fontFace:"Calibri" });
  });
  const pages = d.brandPages.slice(0, 6);
  s.addText("Sources from "+d.brandName+" Domain", { x:5.0, y:1.08, w:4.7, h:0.28, fontSize:9, bold:true, color:C.navy, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:1.36, w:4.7, h:0.02, fill:{color:C.lightgray}, line:{color:C.lightgray} });
  pages.forEach((pg, i) => {
    const y = 1.42 + i * 0.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y, w:4.6, h:0.44, fill:{color:C.offwhite}, line:{color:C.lightgray}, rectRadius:0.04, shadow:makeShadow() });
    s.addText(pg.name, { x:5.15, y:y+0.04, w:3.6, h:0.2, fontSize:8.5, bold:true, color:C.navy, fontFace:"Calibri" });
    s.addText((pg.prompts||0)+" Response"+(pg.prompts!==1?"s":""), { x:9.1, y:y+0.04, w:0.5, h:0.2, fontSize:7.5, bold:true, color:C.teal, align:"right", fontFace:"Calibri" });
    s.addText(pg.url||d.domain, { x:5.15, y:y+0.24, w:4.4, h:0.16, fontSize:7, color:C.slate, fontFace:"Calibri" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.0, y:4.75, w:4.7, h:0.34, fill:{color:C.yellow}, line:{color:C.yellow}, rectRadius:0.05 });
  s.addText("Top Cited Sources from our website ↑", { x:5.0, y:4.75, w:4.7, h:0.34, fontSize:9, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:4.75, y:1.05, w:0.03, h:3.9, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 5: Competitor Visibility Matrix ───────────────────────────────────
function buildSlide5(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "Theme Benchmarks (% Visibility across Competitors)", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const matrix = d.competitorVisibilityMatrix;
  if (!matrix || matrix.length === 0) {
    s.addText("No competitor visibility matrix data available.", { x:0.5, y:2.5, w:9, h:0.5, fontSize:12, color:C.slate, align:"center", fontFace:"Calibri" });
    return;
  }
  const compNames = [];
  matrix.forEach(row => { if (row.competitors) Object.keys(row.competitors).forEach(k => { if (!compNames.includes(k)) compNames.push(k); }); });
  const allCols = [d.brandName, ...compNames].slice(0, 10);
  const themeColW = 1.9, dataColW = (9.5 - themeColW) / allCols.length;
  const startX = 0.25, headerY = 1.08, rowH = 0.32;
  s.addShape(pres.shapes.RECTANGLE, { x:startX, y:headerY, w:9.5, h:rowH, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("Topic", { x:startX+0.05, y:headerY+0.05, w:themeColW-0.08, h:rowH-0.08, fontSize:7, bold:true, color:C.white, fontFace:"Calibri" });
  allCols.forEach((col, ci) => {
    const x = startX + themeColW + ci * dataColW;
    const isBrand = col === d.brandName;
    s.addText(col, { x:x+0.02, y:headerY+0.04, w:dataColW-0.04, h:rowH-0.08, fontSize:6, bold:isBrand, color:isBrand?C.orange:C.white, align:"center", fontFace:"Calibri", wrap:true });
  });
  matrix.slice(0, 11).forEach((row, ri) => {
    const y = headerY + rowH + ri * rowH;
    const bg = ri % 2 === 0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:startX, y, w:9.5, h:rowH, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(row.theme||"", { x:startX+0.05, y:y+0.05, w:themeColW-0.08, h:rowH-0.08, fontSize:6, color:C.navy, fontFace:"Calibri", wrap:true });
    allCols.forEach((col, ci) => {
      const x = startX + themeColW + ci * dataColW;
      const isBrand = col === d.brandName;
      const pct = isBrand ? (typeof row.brandVisibility==='number'?row.brandVisibility:0) : (typeof row.competitors?.[col]==='number'?row.competitors[col]:0);
      let cellCol = bg;
      if (isBrand && pct>0) cellCol = C.purple;
      else if (pct>=15) cellCol = "8B85D4";
      else if (pct>=5)  cellCol = "C4C0EA";
      if (cellCol !== bg) s.addShape(pres.shapes.RECTANGLE, { x:x+0.02, y:y+0.04, w:dataColW-0.04, h:rowH-0.08, fill:{color:cellCol}, line:{color:cellCol} });
      s.addText(pct>0?pct+"%":"0%", { x:x+0.02, y:y+0.05, w:dataColW-0.04, h:rowH-0.1, fontSize:7, bold:isBrand, color:(isBrand&&pct>0)?C.white:(pct>=8?C.white:C.slate), align:"center", fontFace:"Calibri" });
    });
  });
  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.", { x:0.3, y:5.22, w:9.4, h:0.16, fontSize:7, italic:true, color:C.slate, align:"center", fontFace:"Calibri" });
}

// ─── SLIDE 6: Metric Definitions (static) ───────────────────────────────────
function buildSlide6(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "What does each of these mean?", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const defs = [
    { term: "Brand Mentions",  body: "Number of times your brand appeared in AI answers out of total tracked prompts" },
    { term: "Share of Voice",  body: "Percentage of your brand mentions compared to all total brand mentions" },
    { term: "Brand Position",  body: "Average position of your brand in AI answers" },
    { term: "Domain Citation", body: "Number of times your website was cited on AI Search Engines" },
    { term: "Brand Coverage",  body: "Percentage of prompts that mention your brand" },
    { term: "Domain Coverage", body: "Percentage of prompts that cited your website" },
  ];
  defs.forEach((def, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.28 + col * 3.22, y = 1.2 + row * 1.62;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.06, h:1.52, fill:{color:C.yellow}, line:{color:"D4AA30"}, shadow:makeShadow() });
    s.addText(def.term, { x:x+0.14, y:y+0.14, w:2.78, h:0.3, fontSize:12, bold:true, color:C.purple, fontFace:"Calibri" });
    s.addText("— — — — — — — — — — — —", { x:x+0.14, y:y+0.44, w:2.78, h:0.16, fontSize:7, color:C.purple, fontFace:"Calibri" });
    s.addText(def.body, { x:x+0.14, y:y+0.6, w:2.78, h:0.78, fontSize:9, color:C.navy, italic:true, bold:true, fontFace:"Calibri", wrap:true });
  });
  s.addText("Source: Otterly.ai", { x:0.3, y:5.22, w:3, h:0.16, fontSize:7.5, bold:true, color:C.navy, fontFace:"Calibri" });
}

// ─── SLIDE 7: Platform mentions table ───────────────────────────────────────
function buildSlide7(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, d.brandName+" Mentions by AI Platform", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const kpis = [
    { v: String(d.totalMentions),  l: "Total Brand Mentions" },
    { v: String(d.totalCitations), l: "Total Domain Citations" },
    { v: d.avgBrandCoverage,       l: "Avg Brand Coverage" },
    { v: d.avgDomainCoverage,      l: "Avg Domain Coverage" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.25 + i * 2.42;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.08, w:2.3, h:0.72, fill:{color:C.white}, line:{color:C.lightgray}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.08, w:2.3, h:0.06, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText(k.v, { x, y:1.16, w:2.3, h:0.38, fontSize:20, bold:true, color:C.navy, align:"center", fontFace:"Calibri" });
    s.addText(k.l, { x, y:1.52, w:2.3, h:0.24, fontSize:7.5, color:C.slate, align:"center", fontFace:"Calibri" });
  });
  const cols = ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"];
  const colX = [0.25, 2.15, 3.2, 4.3, 7.1];
  const colW = [1.85, 1.0, 1.05, 2.75, 2.6];
  s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:1.9, w:9.5, h:0.28, fill:{color:C.lightgray}, line:{color:C.lightgray} });
  cols.forEach((h, i) => s.addText(h, { x:colX[i], y:1.93, w:colW[i], h:0.22, fontSize:7.5, bold:true, color:C.slate, fontFace:"Calibri" }));
  d.platforms.forEach((p, i) => {
    const y = 2.22 + i * 0.5;
    const bg = i % 2 === 0 ? "F4F3FD" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:y-0.04, w:9.5, h:0.48, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(p.name, { x:colX[0], y, w:colW[0], h:0.28, fontSize:9.5, bold:true, color:C.navy, fontFace:"Calibri" });
    s.addText(String(p.mentions), { x:colX[1], y, w:colW[1], h:0.28, fontSize:9.5, color:C.navy, fontFace:"Calibri" });
    s.addText(String(p.citations), { x:colX[2], y, w:colW[2], h:0.28, fontSize:9.5, color:C.navy, fontFace:"Calibri" });
    const bvPct = Math.min(p.brandVisibility||0,100), bvW = (bvPct/100)*2.5;
    s.addShape(pres.shapes.RECTANGLE, { x:colX[3], y:y+0.08, w:2.5, h:0.14, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    if (bvW>0) s.addShape(pres.shapes.RECTANGLE, { x:colX[3], y:y+0.08, w:bvW, h:0.14, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(bvPct+"%", { x:colX[3], y, w:0.45, h:0.28, fontSize:8, bold:true, color:C.navy, fontFace:"Calibri" });
    const dcPct = Math.min(p.domainCoverage||0,100), dcW = (dcPct/100)*2.5;
    s.addShape(pres.shapes.RECTANGLE, { x:colX[4], y:y+0.08, w:2.5, h:0.14, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    if (dcW>0) s.addShape(pres.shapes.RECTANGLE, { x:colX[4], y:y+0.08, w:dcW, h:0.14, fill:{color:C.violet}, line:{color:C.violet} });
    s.addText(dcPct+"%", { x:colX[4], y, w:0.45, h:0.28, fontSize:8, bold:true, color:C.violet, fontFace:"Calibri" });
  });
}

// ─── SLIDE 8: Brand Visibility by Platform × Theme ──────────────────────────
function buildSlide8(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, d.brandName+" Brand Visibility by Platform & Theme", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);
  const rows = d.brandVisibilityByPlatform;
  if (!rows || rows.length === 0) {
    s.addText("No platform visibility data available.", { x:0.5, y:2.8, w:9, h:0.5, fontSize:12, color:C.slate, align:"center", fontFace:"Calibri" });
    return;
  }
  const platNames = Object.keys(rows[0]).filter(k => k !== 'theme');
  const themeColW = 3.2, dataColW = (9.5-themeColW)/platNames.length;
  const startX = 0.25, headerY = 1.08, rowH = 0.35;
  s.addShape(pres.shapes.RECTANGLE, { x:startX, y:headerY, w:9.5, h:rowH, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("Themes", { x:startX+0.08, y:headerY+0.07, w:themeColW-0.12, h:rowH-0.1, fontSize:8, bold:true, color:C.white, fontFace:"Calibri" });
  platNames.forEach((pn, pi) => {
    const x = startX + themeColW + pi * dataColW;
    s.addText(pn, { x:x+0.04, y:headerY+0.05, w:dataColW-0.08, h:rowH-0.1, fontSize:8, bold:true, color:C.white, align:"center", fontFace:"Calibri", wrap:true });
  });
  rows.slice(0,11).forEach((row, ri) => {
    const y = headerY + rowH + ri * rowH;
    const bg = ri%2===0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:startX, y, w:9.5, h:rowH, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(row.theme||"", { x:startX+0.08, y:y+0.07, w:themeColW-0.14, h:rowH-0.1, fontSize:7, color:C.navy, fontFace:"Calibri", wrap:true });
    platNames.forEach((pn, pi) => {
      const x = startX + themeColW + pi * dataColW;
      const pct = typeof row[pn]==='number' ? row[pn] : 0;
      let cellCol = bg;
      if (pct>=15) cellCol = C.purple;
      else if (pct>=5) cellCol = "C4C0EA";
      if (cellCol!==bg) s.addShape(pres.shapes.RECTANGLE, { x:x+0.04, y:y+0.05, w:dataColW-0.08, h:rowH-0.1, fill:{color:cellCol}, line:{color:cellCol} });
      s.addText(pct>0?pct+"%":"0%", { x:x+0.04, y:y+0.07, w:dataColW-0.08, h:rowH-0.12, fontSize:8, color:pct>=5?C.white:C.slate, align:"center", fontFace:"Calibri" });
    });
  });
  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.", { x:0.3, y:5.22, w:9.4, h:0.16, fontSize:7, italic:true, color:C.slate, align:"center", fontFace:"Calibri" });
}

// ─── SLIDE 7: Platform Mentions Table ───────────────────────────────────────
function buildSlide7(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, d.brandName + " Mentions by AI Platform", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);

  const kpis = [
    { v: String(d.totalMentions),  l: "Total Brand Mentions" },
    { v: String(d.totalCitations), l: "Total Domain Citations" },
    { v: d.avgBrandCoverage,       l: "Avg Brand Coverage" },
    { v: d.avgDomainCoverage,      l: "Avg Domain Coverage" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.25 + i * 2.42;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.08, w:2.3, h:0.72, fill:{color:C.white}, line:{color:C.lightgray}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.08, w:2.3, h:0.06, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText(k.v, { x, y:1.16, w:2.3, h:0.38, fontSize:20, bold:true, color:C.navy, align:"center", fontFace:"Calibri" });
    s.addText(k.l, { x, y:1.52, w:2.3, h:0.24, fontSize:7.5, color:C.slate, align:"center", fontFace:"Calibri" });
  });

  const colLabels = ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"];
  const colX = [0.25, 2.15, 3.2, 4.3, 7.1];
  const colW = [1.85, 1.0, 1.05, 2.75, 2.6];
  s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:1.9, w:9.5, h:0.28, fill:{color:C.lightgray}, line:{color:C.lightgray} });
  colLabels.forEach((h, i) => s.addText(h, { x:colX[i], y:1.93, w:colW[i], h:0.22, fontSize:7.5, bold:true, color:C.slate, fontFace:"Calibri" }));
  d.platforms.forEach((p, i) => {
    const y = 2.22 + i * 0.5;
    const bg = i % 2 === 0 ? "F4F3FD" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:y-0.04, w:9.5, h:0.48, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(p.name, { x:colX[0], y, w:colW[0], h:0.28, fontSize:9.5, bold:true, color:C.navy, fontFace:"Calibri" });
    s.addText(String(p.mentions), { x:colX[1], y, w:colW[1], h:0.28, fontSize:9.5, color:C.navy, fontFace:"Calibri" });
    s.addText(String(p.citations), { x:colX[2], y, w:colW[2], h:0.28, fontSize:9.5, color:C.navy, fontFace:"Calibri" });
    const bvPct = Math.min(p.brandVisibility || 0, 100);
    const bvW = (bvPct / 100) * 2.5;
    s.addShape(pres.shapes.RECTANGLE, { x:colX[3], y:y+0.08, w:2.5, h:0.14, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    if (bvW > 0) s.addShape(pres.shapes.RECTANGLE, { x:colX[3], y:y+0.08, w:bvW, h:0.14, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(bvPct+"%", { x:colX[3], y, w:0.45, h:0.28, fontSize:8, bold:true, color:C.navy, fontFace:"Calibri" });
    const dcPct = Math.min(p.domainCoverage || 0, 100);
    const dcW = (dcPct / 100) * 2.5;
    s.addShape(pres.shapes.RECTANGLE, { x:colX[4], y:y+0.08, w:2.5, h:0.14, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    if (dcW > 0) s.addShape(pres.shapes.RECTANGLE, { x:colX[4], y:y+0.08, w:dcW, h:0.14, fill:{color:C.violet}, line:{color:C.violet} });
    s.addText(dcPct+"%", { x:colX[4], y, w:0.45, h:0.28, fontSize:8, bold:true, color:C.violet, fontFace:"Calibri" });
  });
}

// ─── SLIDE 8: Brand Visibility by Platform & Theme ───────────────────────────
function buildSlide8(pres, d, brandLogoPath) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, d.brandName + " Brand Visibility by Platform & Theme", d.brandName, brandLogoPath);
  ftr(s, pres, d.brandName, d.domain);

  const rows = d.brandVisibilityByPlatform;
  if (!rows || rows.length === 0) {
    s.addText("No platform visibility data available.", { x:0.5, y:2.8, w:9, h:0.5, fontSize:12, color:C.slate, align:"center", fontFace:"Calibri" });
    return;
  }
  const platNames = Object.keys(rows[0]).filter(k => k !== "theme");
  const themeColW = 3.2, dataColW = (9.5 - themeColW) / platNames.length;
  const startX = 0.25, headerY = 1.08, rowH = 0.35;

  s.addShape(pres.shapes.RECTANGLE, { x:startX, y:headerY, w:9.5, h:rowH, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText("Themes", { x:startX+0.08, y:headerY+0.07, w:themeColW-0.12, h:rowH-0.1, fontSize:8, bold:true, color:C.white, fontFace:"Calibri" });
  platNames.forEach((pn, pi) => {
    const x = startX + themeColW + pi * dataColW;
    s.addText(pn, { x:x+0.04, y:headerY+0.05, w:dataColW-0.08, h:rowH-0.1, fontSize:8, bold:true, color:C.white, align:"center", fontFace:"Calibri", wrap:true });
  });
  rows.slice(0, 11).forEach((row, ri) => {
    const y = headerY + rowH + ri * rowH;
    const bg = ri % 2 === 0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:startX, y, w:9.5, h:rowH, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(row.theme || "", { x:startX+0.08, y:y+0.07, w:themeColW-0.14, h:rowH-0.1, fontSize:7, color:C.navy, fontFace:"Calibri", wrap:true });
    platNames.forEach((pn, pi) => {
      const x = startX + themeColW + pi * dataColW;
      const pct = typeof row[pn] === "number" ? row[pn] : 0;
      let cellCol = bg;
      if (pct >= 15) cellCol = C.purple;
      else if (pct >= 5) cellCol = "C4C0EA";
      if (cellCol !== bg) s.addShape(pres.shapes.RECTANGLE, { x:x+0.04, y:y+0.05, w:dataColW-0.08, h:rowH-0.1, fill:{color:cellCol}, line:{color:cellCol} });
      s.addText(pct > 0 ? pct+"%" : "0%", { x:x+0.04, y:y+0.07, w:dataColW-0.08, h:rowH-0.12, fontSize:8, color:pct>=5?C.white:C.slate, align:"center", fontFace:"Calibri" });
    });
  });
  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.", { x:0.3, y:5.22, w:9.4, h:0.16, fontSize:7, italic:true, color:C.slate, align:"center", fontFace:"Calibri" });
}

// ═══════════════════════════════════════════════════════════════════════════
// STATIC SLIDES 12–18  (appended to every report, only logo pill changes)
// ═══════════════════════════════════════════════════════════════════════════

// ─── SLIDE 12: The Approach For Solving GEO ─────────────────────────────────
function buildSlide12(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  staticHdr(s, pres, "The Approach For Solving GEO", brandLogoPath, brandName);

  // Centre circle diagram (representational — 3 arcs as thick ovals)
  const cx = 4.6, cy = 3.1, r = 1.55;
  // Outer arc ring (navy)
  s.addShape(pres.shapes.OVAL, { x:cx-r, y:cy-r, w:r*2, h:r*2, fill:{color:"E8E6F5"}, line:{color:C.navy, pt:18} });
  // Inner fill
  s.addShape(pres.shapes.OVAL, { x:cx-0.9, y:cy-0.9, w:1.8, h:1.8, fill:{color:C.yellow}, line:{color:C.yellow} });
  s.addText("Pepper's GEO\nApproach", { x:cx-0.85, y:cy-0.38, w:1.7, h:0.76, fontSize:9, bold:true, color:C.navy, align:"center", fontFace:"Calibri" });

  // LEFT: Visibility
  s.addText("Visibility", { x:0.3, y:1.5, w:2.8, h:0.36, fontSize:16, bold:true, color:C.purple, fontFace:"Calibri" });
  s.addText("Can LLMs see your content?", { x:0.3, y:1.84, w:2.8, h:0.24, fontSize:9, italic:true, color:C.purple, fontFace:"Calibri" });
  const visPoints = ["We audit if you're being cited across AI Search (ChatGPT, Perplexity, SGE)", "We identify which competitors are winning those spots - and why", "We check if your URLs are indexable, link-worthy, and retrievable"];
  visPoints.forEach((pt, i) => {
    s.addText("• " + pt, { x:0.3, y:2.14+i*0.42, w:2.8, h:0.38, fontSize:8, color:"333333", fontFace:"Calibri", wrap:true });
  });

  // RIGHT: Citability
  s.addText("Citability", { x:6.8, y:1.5, w:3.0, h:0.36, fontSize:16, bold:true, color:C.purple, fontFace:"Calibri" });
  s.addText("Can LLMs trust your content?", { x:6.8, y:1.84, w:3.0, h:0.24, fontSize:9, italic:true, color:C.purple, fontFace:"Calibri" });
  const citPoints = ["We rewrite content to include expert quotes, references, structured data", "We improve source credibility through media presence, high-authority citations, and entity recognition", "We restructure pages to be chunkable and retrievable"];
  citPoints.forEach((pt, i) => {
    s.addText("• " + pt, { x:6.8, y:2.14+i*0.42, w:3.0, h:0.38, fontSize:8, color:"333333", fontFace:"Calibri", wrap:true });
  });

  // BOTTOM: Retrievability
  s.addText("Retrievability:", { x:6.4, y:3.85, w:3.4, h:0.36, fontSize:16, bold:true, color:C.purple, fontFace:"Calibri" });
  s.addText("Can LLMs use your content to answer future questions?", { x:6.4, y:4.2, w:3.4, h:0.36, fontSize:9, italic:true, color:C.purple, fontFace:"Calibri", wrap:true });
  const retPoints = ["We chunk and tag your content to feed RAG systems better", "We add LLM-readable markup and context layering (FAQs, comparisons, summaries)", "We monitor which prompts lead to brand visibility and close the loop"];
  retPoints.forEach((pt, i) => {
    s.addText("• " + pt, { x:6.4, y:4.58+i*0.28, w:3.4, h:0.26, fontSize:7.5, color:"333333", fontFace:"Calibri", wrap:true });
  });

  // Footer line
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 13: Reverse Engineering How LLMs Index Content ───────────────────
function buildSlide13(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  staticHdr(s, pres, "Reverse Engineering How LLMs Index Content", brandLogoPath, brandName);

  // Formula band
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:0.82, w:9.4, h:0.62, fill:{color:"EEF0FF"}, line:{color:"EEF0FF"}, rectRadius:0.08 });
  s.addText("LLM Retrieval Score", { x:0.6, y:0.9, w:2.8, h:0.42, fontSize:13, bold:true, italic:true, color:C.navy, fontFace:"Calibri" });
  s.addText("∝", { x:3.4, y:0.9, w:0.4, h:0.42, fontSize:18, color:C.navy, align:"center", fontFace:"Calibri" });
  s.addText("( Chunking × Structure × Schema × Source Weight × Trust Signals )", { x:3.8, y:0.9, w:5.7, h:0.42, fontSize:12, bold:true, italic:true, color:C.navy, fontFace:"Calibri", wrap:true });

  // Table
  const headers = ["Variable", "What It Means", "How Pepper Optimizes It"];
  const colX = [0.3, 1.85, 5.5];
  const colW = [1.5, 3.6, 4.1];
  const tableY = 1.62;
  const tableH = 0.32;

  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:tableY, w:9.4, h:tableH, fill:{color:C.navy}, line:{color:C.navy} });
  headers.forEach((h, i) => s.addText(h, { x:colX[i]+0.08, y:tableY+0.06, w:colW[i]-0.1, h:tableH-0.1, fontSize:8, bold:true, italic:true, color:C.white, fontFace:"Calibri" }));

  const rows = [
    ["Chunking",       "Atomic 2–4 sentence blocks ideal for embedding + summarization",               "We rewrite long-form into discrete semantic units"],
    ["Structure",      "Use of TL;DRs, bullets, lists, Q&A formatting",                               "Content is formatted with high semantic clarity"],
    ["Schema",         "Machine-readable metadata (FAQPage, HowTo, Product)",                         "Implemented across product pages, glossaries, and help docs"],
    ["Source Weight",  "LLM preference hierarchy (Wikipedia > PDF > Help Docs > Blogs > Social)",     "Content is distributed into high-weight surfaces LLMs trust"],
    ["Trust Signals",  "Presence of citations, statistics, interlinking, and cross-source agreement", "We embed outbound and inbound credibility into every content artifact"],
  ];
  rows.forEach((row, ri) => {
    const y = tableY + tableH + ri * 0.52;
    const bg = ri % 2 === 0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:9.4, h:0.5, fill:{color:bg}, line:{color:C.lightgray} });
    row.forEach((cell, ci) => {
      s.addText(cell, { x:colX[ci]+0.08, y:y+0.08, w:colW[ci]-0.12, h:0.36, fontSize:7.5, italic:ci>0, color:C.navy, fontFace:"Calibri", wrap:true });
    });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 14: The Content Strategy — Source Weightages by LLMs ─────────────
function buildSlide14(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  staticHdr(s, pres, "The Content Strategy : Source Weightages by LLMs", brandLogoPath, brandName);

  const platforms = [
    { name: "ChatGPT",    bg: "EDE9FF", line: C.purple },
    { name: "Gemini",     bg: "E8F4FF", line: "4285F4" },
    { name: "perplexity", bg: "F0F0F0", line: "666666" },
    { name: "Claude",     bg: "FFF0EC", line: "D97659" },
  ];
  const colX = [0.28, 2.28, 4.28, 6.28, 8.28];
  const colW = [1.96, 1.96, 1.96, 1.96, 1.44];

  // Header row
  s.addShape(pres.shapes.RECTANGLE, { x:0.28, y:0.82, w:9.44, h:0.44, fill:{color:C.offwhite}, line:{color:C.lightgray} });
  s.addText("FACTORS", { x:colX[0]+0.06, y:0.9, w:colW[0]-0.1, h:0.28, fontSize:8, bold:true, color:C.slate, fontFace:"Calibri" });
  platforms.forEach((pl, pi) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:colX[pi+1]+0.04, y:0.84, w:colW[pi+1]-0.08, h:0.38, fill:{color:pl.bg}, line:{color:pl.line, pt:1.5}, rectRadius:0.1 });
    s.addText(pl.name, { x:colX[pi+1]+0.04, y:0.84, w:colW[pi+1]-0.08, h:0.38, fontSize:9, bold:true, color:pl.line, align:"center", valign:"middle", fontFace:"Calibri" });
  });

  const factors = [
    { label:"Authoritative List Mentions
(e.g. "Top CRMs on G2", "Best VPNs on TechRadar")", vals:["✅ (41%)", "✅ (49%)", "✅ (64%)", "✗"] },
    { label:"Awards / Accreditations
(e.g. "Gartner MQ", "Inc. 5000 Badge")",               vals:["✅ (18%)", "✅ (15%)", "✅ (5%)",  "✅ (19%)"] },
    { label:"Online Reviews
(e.g. Google Reviews, TrustPilot, Capterra ratings)",             vals:["✅ (16%)", "✅ (13%)", "✅ (31%)", "✗"] },
    { label:"Customer Examples / Usage Data
(e.g. "Used by IBM", "Powered by Salesforce")",  vals:["✅ (14%)", "✗",        "✗",        "✅ (13%)"] },
    { label:"Social Sentiment
(e.g. Reddit threads, Quora/Twitter buzz)",                     vals:["✅ (11%)", "✗",        "✗",        "✗"] },
    { label:"Local Reviews
(e.g. Google Business Profile, Yelp)",                             vals:["✗",        "✅ (Local)","✅ (Local)","✗"] },
    { label:"Traditional Directories
(e.g. NY Times, Bloomberg, Hoovers)",                   vals:["✗",        "✗",        "✗",        "✅ (68%)"] },
  ];

  factors.forEach((row, ri) => {
    const y = 1.32 + ri * 0.56;
    const bg = ri % 2 === 0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.28, y, w:9.44, h:0.54, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(row.label, { x:colX[0]+0.06, y:y+0.06, w:colW[0]-0.1, h:0.44, fontSize:6.5, color:C.navy, fontFace:"Calibri", wrap:true });
    row.vals.forEach((val, vi) => {
      const isCheck = val.startsWith("✅");
      const isCross = val === "✗";
      const color = isCheck ? C.teal : isCross ? C.orange : C.navy;
      s.addText(val, { x:colX[vi+1]+0.04, y:y+0.1, w:colW[vi+1]-0.08, h:0.36, fontSize:8, color, align:"center", fontFace:"Calibri" });
    });
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 15: The Content Strategy — Topical Authority ─────────────────────
function buildSlide15(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  staticHdr(s, pres, "The Content Strategy : Topical Authority", brandLogoPath, brandName);

  // Formula band
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:0.82, w:9.4, h:0.52, fill:{color:"EEF0FF"}, line:{color:"EEF0FF"}, rectRadius:0.08 });
  s.addText("LLM Recommendations", { x:0.6, y:0.88, w:3.0, h:0.36, fontSize:13, bold:true, italic:true, color:C.navy, fontFace:"Calibri" });
  s.addText("∝", { x:3.62, y:0.88, w:0.4, h:0.36, fontSize:18, color:C.navy, align:"center", fontFace:"Calibri" });
  s.addText("RRF Score = Σ [1 / (60 + SERP position)]", { x:4.05, y:0.88, w:5.4, h:0.36, fontSize:13, bold:true, italic:true, color:C.navy, fontFace:"Calibri" });

  // Illustration label
  s.addText("Illustration", { x:0.3, y:1.48, w:2, h:0.28, fontSize:10, bold:true, color:C.navy, fontFace:"Calibri" });

  // Brand A box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:1.78, w:3.8, h:0.28, fill:{color:"E8E6F5"}, line:{color:C.purple, pt:1.5}, rectRadius:0.06 });
  s.addText("Brand A", { x:0.45, y:1.8, w:1.5, h:0.24, fontSize:9, bold:true, color:C.purple, fontFace:"Calibri" });

  // Brand A table (simple left table)
  const aRows = [["Query","Rank","RRF Score"],["Best Credit Card for Travellers","#1","0.0164"]];
  aRows.forEach((row, ri) => {
    const y = 2.1 + ri * 0.36;
    const bg = ri === 0 ? C.navy : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.3, y, w:3.8, h:0.34, fill:{color:bg}, line:{color:C.lightgray} });
    const textColor = ri === 0 ? C.white : C.navy;
    const cellW = [2.0, 0.85, 0.85];
    const cellX = [0.38, 2.38, 3.23];
    row.forEach((cell, ci) => s.addText(cell, { x:cellX[ci], y:y+0.06, w:cellW[ci], h:0.22, fontSize:7.5, bold:ri===0, italic:ri>0, color:textColor, fontFace:"Calibri" }));
  });

  // Chase/BofA/PNC logos placeholder
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.3, y:2.86, w:1.85, h:1.2, fill:{color:"F0EEFF"}, line:{color:C.purple, pt:1}, rectRadius:0.08 });
  ["CHASE", "BANK OF AMERICA", "PNC"].forEach((b, i) => s.addText(b, { x:0.38, y:2.96+i*0.34, w:1.68, h:0.28, fontSize:8, bold:true, color:C.navy, fontFace:"Calibri" }));

  // nerdwallet/credit karma placeholder
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:2.26, y:3.28, w:1.85, h:0.88, fill:{color:"E8F5EE"}, line:{color:C.teal, pt:1}, rectRadius:0.08 });
  ["nerdwallet", "credit karma"].forEach((b, i) => s.addText(b, { x:2.34, y:3.36+i*0.32, w:1.68, h:0.28, fontSize:8, color:C.teal, fontFace:"Calibri" }));

  // Brand B table (right)
  const bHeaders = ["Query","Rank","RRF Score"];
  const bRows = [
    ["Best Credit Card for Travellers","#4","0.0156"],
    ["Travel On A Budget","#5","0.0154"],
    ["Use Reward Points for Flights","#6","0.0152"],
    ["Cheapest Hotel Tricks on Your Credit Card","#4","0.0156"],
    ["Eligibility for Credit Card","#7","0.0149"],
    ["Total","","0.0767"],
  ];
  s.addShape(pres.shapes.RECTANGLE, { x:4.3, y:1.78, w:5.4, h:0.3, fill:{color:C.navy}, line:{color:C.navy} });
  bHeaders.forEach((h, i) => {
    const xs2 = [4.38, 7.1, 8.5]; const ws2 = [2.68, 1.36, 1.12];
    s.addText(h, { x:xs2[i], y:1.82, w:ws2[i], h:0.22, fontSize:7.5, bold:true, italic:true, color:C.white, fontFace:"Calibri" });
  });
  bRows.forEach((row, ri) => {
    const y = 2.12 + ri * 0.4;
    const bg = ri % 2 === 0 ? C.offwhite : C.white;
    const isTotal = ri === bRows.length - 1;
    s.addShape(pres.shapes.RECTANGLE, { x:4.3, y, w:5.4, h:0.38, fill:{color:isTotal?C.lightgray:bg}, line:{color:C.lightgray} });
    const xs2 = [4.38, 7.1, 8.5]; const ws2 = [2.68, 1.36, 1.12];
    row.forEach((cell, ci) => s.addText(cell, { x:xs2[ci], y:y+0.06, w:ws2[ci], h:0.26, fontSize:7.5, italic:true, bold:isTotal, color:C.navy, fontFace:"Calibri" }));
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:4.3, y:4.56, w:2.0, h:0.28, fill:{color:"E8F5EE"}, line:{color:C.teal, pt:1}, rectRadius:0.06 });
  s.addText("Brand B", { x:4.38, y:4.58, w:1.84, h:0.24, fontSize:9, bold:true, color:C.teal, fontFace:"Calibri" });

  // Right bullets
  const bullets = [
    { bold:"Building Topical Authority", rest:" still tops the content priority, with the relevant technical optimisation for each URL" },
    { bold:"", rest:"The key reason of smaller publishers/brands doing well on LLM queries is their " },
  ];
  s.addText("Building Topical Authority", { x:6.6, y:1.82, w:3.1, h:0.28, fontSize:9, bold:true, color:C.purple, fontFace:"Calibri" });
  s.addText("still tops the content priority, with the relevant technical optimisation for each URL", { x:6.6, y:2.1, w:3.1, h:0.52, fontSize:8.5, color:"333333", fontFace:"Calibri", wrap:true });
  s.addText("The key reason of smaller publishers/brands doing well on LLM queries is their ", { x:6.6, y:2.7, w:3.1, h:0.52, fontSize:8.5, color:"333333", fontFace:"Calibri", wrap:true });
  s.addText("trust-signalling coverage", { x:6.6, y:3.2, w:3.1, h:0.28, fontSize:9, bold:true, color:C.purple, fontFace:"Calibri" });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 16: Strategy to Dominate Generative Search ───────────────────────
function buildSlide16(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  staticHdr(s, pres, "Strategy to Dominate Generative Search (GEO)", brandLogoPath, brandName);
  s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:0.62, w:5.5, h:0.025, fill:{color:C.teal}, line:{color:C.teal} });

  const steps = [
    { num:"1.", title:"Curate Prompt List",             body:"Curate a list of relevant prompts",                             icon:"📋" },
    { num:"2.", title:"Multi LLM Analysis",             body:"Analyze the responses from different LLMs",                    icon:"🖥" },
    { num:"3.", title:"Citation & Brand Mention Audit", body:"Review of cited URLs & brand mentions",                        icon:"🔍" },
    { num:"4.", title:"Page Creation & Optimization",   body:"Identifying pages to be created & optimized",                  icon:"🎯" },
    { num:"5.", title:"Community Visibility",           body:"Access your presence of reddit, Quora forums",                 icon:"👥" },
  ];

  const boxW = 1.68, boxH = 1.8, startX = 0.28, boxY = 1.85, gap = 0.12;
  steps.forEach((step, i) => {
    const x = startX + i * (boxW + gap);
    const isHighlight = i === 1; // Step 2 highlighted in navy border
    s.addText(step.icon, { x:x+boxW/2-0.28, y:boxY-0.5, w:0.56, h:0.44, fontSize:22, align:"center" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y:boxY, w:boxW, h:boxH,
      fill:{color:isHighlight?"E8E6F5":C.offwhite},
      line:{color:isHighlight?C.navy:C.teal, pt:isHighlight?2:1.5},
      rectRadius:0.1,
    });
    s.addText(step.num, { x:x+0.1, y:boxY+0.12, w:0.4, h:0.28, fontSize:10, bold:true, color:C.navy, fontFace:"Calibri" });
    s.addText(step.title, { x:x+0.1, y:boxY+0.4, w:boxW-0.2, h:0.52, fontSize:9, bold:true, color:C.navy, fontFace:"Calibri", wrap:true });
    s.addText(step.body,  { x:x+0.1, y:boxY+0.96, w:boxW-0.2, h:0.72, fontSize:8, italic:true, color:C.slate, fontFace:"Calibri", wrap:true });

    // Arrow between boxes
    if (i < steps.length - 1) {
      s.addText("→", { x:x+boxW+0.01, y:boxY+0.72, w:0.12, h:0.36, fontSize:12, bold:true, color:C.navy, align:"center", fontFace:"Calibri" });
    }
  });

  // Bottom banner: Step 6
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.28, y:3.88, w:9.44, h:0.72, fill:{color:"E8E6F5"}, line:{color:C.navy, pt:1.5}, rectRadius:0.08 });
  s.addText("6. Implement & Iterate", { x:0.48, y:3.96, w:2.8, h:0.28, fontSize:10, bold:true, color:C.navy, fontFace:"Calibri" });
  s.addText("– Create new pages, update existing pages, implement schemas, and community replies and re-run the prompt set monthly to gauge lift and uncover new topical gaps.", { x:3.28, y:3.96, w:6.3, h:0.52, fontSize:8.5, color:C.navy, fontFace:"Calibri", wrap:true });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 17: Customised Content Strategy Table ─────────────────────────────
function buildSlide17(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  staticHdr(s, pres, "Here, is the customised content strategy table for you!", brandLogoPath, brandName);

  const headers = ["Source Type", "Estimated Weight", "Notes", "% Weightage (Approx.)", "Relevant Examples"];
  const colX = [0.28, 1.52, 2.76, 6.44, 7.58];
  const colW = [1.2,  1.2,  3.64, 1.1,  2.14];
  const tableY = 0.82;

  s.addShape(pres.shapes.RECTANGLE, { x:0.28, y:tableY, w:9.44, h:0.34, fill:{color:C.purple}, line:{color:C.purple} });
  headers.forEach((h, i) => {
    s.addText(h, { x:colX[i]+0.05, y:tableY+0.05, w:colW[i]-0.08, h:0.26, fontSize:7, bold:true, color:C.white, fontFace:"Calibri", wrap:true, align:"center" });
  });

  const rows = [
    ["Product & Platform Pages",     "Very High", 'Core retrieval: prompts like "ServiceNow ITSM," "workflow automation," "Now Platform AI."',                          "25%", "ITSM, ITOM, CSM, HRSD, Creator Workflows"],
    ["Industry Solutions Pages",     "High",      'For prompts like "for banking," "workflow automation for healthcare," "government digital services."',              "14%", "Financial services, healthcare, manufacturing"],
    ["AI & Technology Innovation",   "High",      'For prompts like "GenAI," "Now Assist generative AI," "platform intelligence."',                                   "13%", "GenAI copilots, predictive intelligence releases"],
    ["Customer Stories / Case Studies","High",    'Retrieval for "who uses [brand]," "ROI," "success stories."',                                                      "12%", "Case studies with Citi, DHL, Novartis"],
    ["Pricing / Demo Pages",         "Medium-High",'For prompts like "[brand] demo," "pricing," "enterprise license."',                                               "10%", "Request a demo, ROI tools"],
    ["Documentation & Knowledge Base","Medium",  'Retrieval for "[brand] API," "developer docs," "workflow integration."',                                            "8%",  "Developer portal, integration hub, API references"],
    ["Events & Webinars",            "Medium",   'For prompts like "[brand] Knowledge conference," "future of IT workflows."',                                        "6%",  "Knowledge conference sessions, leadership keynotes"],
    ["Press & Newsroom",             "Medium-Low",'For prompts like "[brand] earnings," "acquisitions," "expansion."',                                                "6%",  "Acquisitions, quarterly earnings"],
    ["Support / Help Center",        "Low-Medium",'For prompts like "[brand] login," "ticket help," "support portal."',                                               "4%",  "Support portal, instance upgrades, troubleshooting"],
    ["Careers & Corporate Pages",    "Low",      'For prompts like "jobs at [brand]," "[brand] culture," "DEI."',                                                    "2%",  "Careers site, employee stories, ESG reports"],
  ];

  rows.forEach((row, ri) => {
    const y = tableY + 0.34 + ri * 0.42;
    const bg = ri % 2 === 0 ? C.offwhite : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.28, y, w:9.44, h:0.4, fill:{color:bg}, line:{color:C.lightgray} });
    row.forEach((cell, ci) => {
      const isWeight = ci === 1;
      const isPct = ci === 3;
      const color = isPct ? C.purple : isWeight ? C.teal : C.navy;
      s.addText(cell, { x:colX[ci]+0.05, y:y+0.06, w:colW[ci]-0.08, h:0.3, fontSize:6.5, bold:isPct||isWeight, color, fontFace:"Calibri", wrap:true, align:ci===3?"center":"left" });
    });
  });

  s.addText("Note: The weightage percentage is a relative effort guide. If you're putting X effort on a source with 2% weight, then a source with 8% weight deserves 4X effort.", {
    x:0.28, y:5.1, w:9.44, h:0.28, fontSize:6.5, italic:true, color:C.slate, fontFace:"Calibri", wrap:true
  });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 18: Do you think this was helpful? ────────────────────────────────
function buildSlide18(pres, brandLogoPath, brandName) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addLogoPill(s, brandLogoPath, brandName);

  s.addText("Do you think this was helpful and want\nto dive deeper?", {
    x:0.5, y:0.9, w:8.5, h:1.4, fontSize:28, bold:true, color:C.purple, fontFace:"Calibri", wrap:true
  });
  s.addText("Reach us at +14157545133", { x:0.5, y:2.55, w:6, h:0.38, fontSize:14, color:C.navy, fontFace:"Calibri" });
  s.addText("or write to us at kishan@peppercontent.io", { x:0.5, y:2.95, w:6, h:0.38, fontSize:14, color:C.navy, fontFace:"Calibri" });

  // Yellow CTA band
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:3.9, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:3.98, w:10, h:1.1, fill:{color:C.yellow}, line:{color:C.yellow} });
  s.addText("We are hosting some free GEO workshops for corporates and if you think you'd be interested in the same, please write to us", {
    x:0.5, y:4.1, w:9, h:0.84, fontSize:12, color:C.purple, align:"center", fontFace:"Calibri", wrap:true, italic:true
  });

  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.03, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── STEP 4: Build the full PPTX ─────────────────────────────────────────────
function buildPPTX(data, brandLogoPath, outputPath) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = data.brandName + " GEO Audit — Atlas / Pepper";
  pres.author = "Pepper.inc Atlas";

  // Data slides (1–8)
  buildSlide1(pres, data, brandLogoPath);
  buildSlide2(pres, data, brandLogoPath);
  buildSlide3(pres, data, brandLogoPath);
  buildSlide4(pres, data, brandLogoPath);
  buildSlide5(pres, data, brandLogoPath);
  buildSlide6(pres, data, brandLogoPath);
  buildSlide7(pres, data, brandLogoPath);
  buildSlide8(pres, data, brandLogoPath);

  // Static appended slides (12–18 from template)
  buildSlide12(pres, data, brandLogoPath);
  buildSlide13(pres, data, brandLogoPath);
  buildSlide14(pres, data, brandLogoPath);
  buildSlide15(pres, data, brandLogoPath);
  buildSlide16(pres, data, brandLogoPath);
  buildSlide17(pres, data, brandLogoPath);
  buildSlide18(pres, data, brandLogoPath);

  pres.writeFile({ fileName: outputPath });
  console.log("✅ PPTX written:", outputPath);
}

// ─── API: POST /generate ─────────────────────────────────────────────────────
app.post("/generate", async (req, res) => {
  const { url } = req.body;
  if (!url || !url.includes("atlas.pepper.inc")) {
    return res.status(400).json({ error: "Please provide a valid Atlas report URL." });
  }
  const id = uuidv4();
  const pptxOut = path.join(TMP, id + ".pptx");
  const pdfOut  = path.join(TMP, id + ".pdf");
  try {
    // Step 1 — Scrape
    const screenshots = await scrapeAtlasReport(url);
    // Step 2 — Extract
    const rawData = await extractData(screenshots);
    // Step 3 — Normalize
    const data = normalizeData(rawData);
    // Step 3b — Fetch brand logo (pre-committed or auto-scraped)
    const brandLogoPath = await fetchBrandLogo(data.brandName, data.domain);
    // Step 4 — Build PPTX
    buildPPTX(data, brandLogoPath, pptxOut);
    // Step 5 — Convert to PDF
    await new Promise(r => setTimeout(r, 1000));
    console.log("📄 Converting to PDF...");
    execSync(`soffice --headless --convert-to pdf --outdir ${TMP} ${pptxOut}`, { timeout: 60000 });
    if (!fs.existsSync(pdfOut)) throw new Error("PDF conversion failed");
    // Step 6 — Send with correct filename
    const fileName = `${data.brandName} x Pepper - GEO report.pdf`;
    res.download(pdfOut, fileName, (err) => {
      try { fs.unlinkSync(pptxOut); } catch {}
      try { fs.unlinkSync(pdfOut); } catch {}
      // Clean up any tmp logo fetched at runtime
      if (brandLogoPath && brandLogoPath.startsWith(TMP)) {
        try { fs.unlinkSync(brandLogoPath); } catch {}
      }
      if (err && !res.headersSent) res.status(500).json({ error: "Download failed." });
    });
  } catch (err) {
    console.error("❌ Error:", err.message);
    res.status(500).json({ error: err.message || "Generation failed." });
  }
});

// ─── Health check ─────────────────────────────────────────────────────────────
app.get("/health", (_, res) => res.json({ status: "ok" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Atlas PDF Generator on port ${PORT}`));
