const express = require("express");
const { chromium } = require("playwright");
const pptxgen = require("pptxgenjs");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");

const app = express();
app.use(express.json());
app.use(express.static("public"));

const TMP = path.join(__dirname, "tmp");
if (!fs.existsSync(TMP)) fs.mkdirSync(TMP);

// ─── Color palette ────────────────────────────────────────────────────────────
const C = {
  navy:      "1A1A3E",
  purple:    "3D3A8C",
  violet:    "5B4FBE",
  lilac:     "9B93E3",
  orange:    "F4A419",
  yellow:    "F5C842",
  white:     "FFFFFF",
  offwhite:  "F8F7FF",
  slate:     "64748B",
  lightgray: "E8E6F5",
  darkgray:  "2D2B55",
  teal:      "0891B2",
  green:     "16A34A",
};

function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 };
}

// ─── Atlas tab URL resolver ────────────────────────────────────────────────────
function getAtlasTabUrls(baseUrl) {
  const root = baseUrl.replace(/\/(overview|competitors-comparison|platforms|prompts-themes).*$/, "");
  return {
    overview:    root + "/overview",
    competitors: root + "/competitors-comparison",
    platforms:   root + "/platforms",
    prompts:     root + "/prompts-themes",
  };
}

// ─── STEP 1: Scrape — visit each tab, screenshot, send to Claude Vision ───────
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
        console.log("    ✅", name, "—", Math.round(buf.length / 1024), "KB");
      } catch (e) {
        console.warn("    ⚠️  Could not capture", name, ":", e.message);
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

// ─── STEP 2: Extract — send screenshots to Claude, get structured JSON data ───
async function extractData(screenshots) {
  console.log("  🤖 Sending screenshots to Claude Vision API...");

  const imageBlocks = Object.entries(screenshots).flatMap(([name, b64]) => [
    { type: "text", text: `## Tab: ${name}` },
    { type: "image", source: { type: "base64", media_type: "image/png", data: b64 } },
  ]);

  const res = await fetch("https://api.anthropic.com/v1/messages", {
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
      messages: [{
        role: "user",
        content: [
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
    { "theme": "24/7 Operations Efficiency & Compliance Management", "promptCount": 7, "prompts": ["best grab and go beverage solutions", "FSSAI compliant beverage vending solutions"] },
    { "theme": "Beverage Quality Consistency & Menu Diversification", "promptCount": 8, "prompts": ["80 beverage options single machine"] }
  ],

  "competitorVisibilityMatrix": [
    {
      "theme": "24/7 Operations Efficiency & Compliance Management",
      "brandVisibility": 7,
      "competitors": { "WMF": 0, "Jura": 0, "Kaapi Machines": 0, "Franke": 3, "De'Longhi": 0, "La Cimbali": 0, "La Marzocco": 0, "Vendekin": 10, "Atlantis": 3 }
    }
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
- Extract ALL rows from every table visible (competitors, platforms, domains, brand pages, themes).
- promptThemes: extract theme name, how many prompts it has, and list all prompt text visible.
- competitorVisibilityMatrix: sourced from "Brand Visibility x Competitors" on the competitors tab. Each row = one theme, "brandVisibility" = the brand's own column integer %, "competitors" = {CompetitorName: integerPct, ...} for every other visible column. Extract ALL rows and ALL competitor columns.
- brandVisibilityByPlatform: sourced from the "[BrandName] brand visibility" table on the platforms tab. Each row = one theme, keys are platform names (ChatGPT / Google AI Overview / Perplexity) with integer % values. Extract ALL rows visible.
- Return ONLY the JSON.` }
        ],
      }],
    }),
  });

  if (!res.ok) throw new Error(`Claude API error: ${res.status} — ${await res.text()}`);
  const json = await res.json();
  const text = json.content[0].text.trim();

  let data;
  try { data = JSON.parse(text); }
  catch {
    const m = text.match(/\{[\s\S]*\}/);
    if (m) data = JSON.parse(m[0]);
    else throw new Error("Could not parse Claude JSON response");
  }

  console.log("  Brand:", data.brandName);
  console.log("  Leaderboard:", data.leaderboard?.length, "entries");
  console.log("  Competitors:", data.competitorMentions?.length, "entries");
  console.log("  Platforms:", data.platforms?.length, "entries");
  console.log("  Themes:", data.promptThemes?.length, "themes");
  console.log("  Total mentions:", data.totalMentions);
  return data;
}

// ─── STEP 3: Normalize — fill defaults so PPTX builder never crashes ──────────
function normalizeData(raw) {
  console.log("\n📊 Normalizing extracted data...");
  const tm = raw.totalMentions || 0;
  const tc = raw.totalCitations || 0;

  const platforms = raw.platforms?.length > 0 ? raw.platforms : [
    { name: "ChatGPT",           mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
    { name: "Google AI Overview",mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
    { name: "Perplexity",        mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
  ];

  const leaderboard = raw.leaderboard?.length > 0 ? raw.leaderboard
    : [{ rank: 1, name: raw.brandName, mentions: tm }];

  const competitorMentions = raw.competitorMentions?.length > 0 ? raw.competitorMentions
    : leaderboard.map((b, i) => ({ name: b.name, percentage: Math.max(30 - i * 8, 2), mentions: b.mentions }));

  const promptThemes = raw.promptThemes?.length > 0 ? raw.promptThemes : [];
  const domainCitations = raw.domainCitations?.length > 0 ? raw.domainCitations : [];
  const brandPages = raw.brandPages?.length > 0 ? raw.brandPages : [];
  const competitorVisibilityMatrix = raw.competitorVisibilityMatrix?.length > 0 ? raw.competitorVisibilityMatrix : [];

  const isLeader = leaderboard[0]?.name === raw.brandName;
  const leaderboardRank = isLeader ? "#1"
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
    promptThemes,
    domainCitations: domainCitations.slice(0, 10),
    brandPages: brandPages.slice(0, 8),
    competitorVisibilityMatrix,
    brandVisibilityByPlatform: raw.brandVisibilityByPlatform?.length > 0 ? raw.brandVisibilityByPlatform : [],
  };
}


// ─── PPTX helpers ─────────────────────────────────────────────────────────────
function hdr(s, pres, title, brand) {
  // Cyan top accent bar
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.08, fill:{color:C.teal}, line:{color:C.teal} });
  // Logo area
  s.addText("atlas", { x:0.3, y:0.12, w:0.9, h:0.3, fontSize:11, bold:true, color:C.navy, fontFace:"Calibri" });
  s.addText("by pepper.inc", { x:1.2, y:0.17, w:1.3, h:0.22, fontSize:7, color:C.slate, fontFace:"Calibri" });
  // Brand pill top-right
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:8.2, y:0.1, w:1.6, h:0.28, fill:{color:C.lightgray}, line:{color:C.lightgray}, rectRadius:0.05 });
  s.addText(brand.toUpperCase(), { x:8.2, y:0.1, w:1.6, h:0.28, fontSize:7, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
  // Slide title
  s.addText(title, { x:0.3, y:0.5, w:9.4, h:0.48, fontSize:18, bold:true, color:C.navy, fontFace:"Calibri" });
  // Thin underline
  s.addShape(pres.shapes.RECTANGLE, { x:0.3, y:0.95, w:1.8, h:0.03, fill:{color:C.teal}, line:{color:C.teal} });
}

function ftr(s, pres, brand, domain) {
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.42, w:10, h:0.2, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText(brand + " · " + domain + " · GEO Audit by Atlas / Pepper.inc", {
    x:0.3, y:5.43, w:9.4, h:0.18, fontSize:6.5, color:"AAAACC", fontFace:"Calibri"
  });
}

// ─── SLIDE 1: Cover ───────────────────────────────────────────────────────────
function buildSlide1(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  // Decorative circles
  s.addShape(pres.shapes.OVAL, { x:7.5, y:-0.5, w:3.5, h:3.5, fill:{color:"2D2B55"}, line:{color:"2D2B55"} });
  s.addShape(pres.shapes.OVAL, { x:8.2, y:0.2, w:2.0, h:2.0, fill:{color:C.purple}, line:{color:C.purple} });
  // Logo
  s.addText("atlas", { x:0.5, y:0.4, w:1.0, h:0.38, fontSize:15, bold:true, color:C.orange, fontFace:"Calibri" });
  s.addText("by pepper.inc", { x:1.52, y:0.46, w:1.4, h:0.26, fontSize:8, color:C.lilac, fontFace:"Calibri" });
  // Tag
  s.addText("GEO AUDIT REPORT", { x:0.5, y:1.05, w:6, h:0.28, fontSize:9, color:C.orange, bold:true, charSpacing:4, fontFace:"Calibri" });
  // Brand name
  s.addText(d.brandName, { x:0.5, y:1.35, w:7, h:1.2, fontSize:50, bold:true, color:C.white, fontFace:"Calibri" });
  s.addText(d.domain, { x:0.5, y:2.58, w:5, h:0.38, fontSize:13, color:C.lilac, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.0, w:1.2, h:0.04, fill:{color:C.orange}, line:{color:C.orange} });
  // KPI strip
  const kpis = [
    { v: String(d.totalMentions), l: "Total Mentions" },
    { v: d.avgBrandCoverage, l: "Brand Coverage" },
    { v: String(d.platformCount), l: "AI Platforms" },
    { v: d.leaderboardRank, l: "Leaderboard" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.5 + i * 2.3;
    s.addText(k.v, { x, y:3.15, w:2.1, h:0.52, fontSize:24, bold:true, color:C.orange, fontFace:"Calibri" });
    s.addText(k.l, { x, y:3.65, w:2.1, h:0.22, fontSize:8, color:C.lilac, fontFace:"Calibri" });
  });
  s.addText("Powered by atlas · pepper.inc", { x:0.5, y:5.15, w:9, h:0.22, fontSize:7.5, color:C.slate, fontFace:"Calibri" });
}

// ─── SLIDE 2: Prompts & Themes overview — 3-col grid ─────────────────────────
function buildSlide2(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, `We Have Mapped ${d.promptThemes.reduce((a,t)=>a+(t.promptCount||t.prompts?.length||0),0)} Prompts Across ${d.promptThemes.length} Themes`, d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const themes = d.promptThemes.slice(0, 9);
  const cols = 3;
  const colW = 3.0, colGap = 0.15, rowH = 0.72, rowGap = 0.1;
  const startX = 0.28, startY = 1.1;

  themes.forEach((t, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (colW + colGap);
    const y = startY + row * (rowH + rowGap);
    const count = t.promptCount || t.prompts?.length || 0;

    s.addShape(pres.shapes.RECTANGLE, { x, y, w:colW, h:rowH, fill:{color:C.offwhite}, line:{color:C.lightgray}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.04, h:rowH, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText(t.theme, { x:x+0.1, y:y+0.08, w:colW-0.15, h:0.34, fontSize:9.5, bold:true, color:C.navy, fontFace:"Calibri", wrap:true });
    s.addText(count + " prompts", { x:x+0.1, y:y+0.46, w:colW-0.15, h:0.2, fontSize:8.5, color:C.slate, fontFace:"Calibri" });
  });

  // Yellow CTA bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:2.8, y:4.78, w:4.4, h:0.36, fill:{color:C.yellow}, line:{color:C.yellow}, rectRadius:0.05 });
  s.addText("Link to the entire list of prompts", { x:2.8, y:4.78, w:4.4, h:0.36, fontSize:10, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
  s.addText("This is a relevant set of non-branded prompts for your brand.", {
    x:0.3, y:5.16, w:9.4, h:0.2, fontSize:7.5, color:C.slate, italic:true, align:"center", fontFace:"Calibri"
  });
}


// ─── SLIDE 3: Brand Leaderboard + Competitor Mentions ────────────────────────
function buildSlide3(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "Brand Leaderboard & Competitor Mentions", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  // LEFT: Leaderboard podium
  const brands = d.leaderboard;
  if (brands.length > 0) {
    const maxM = Math.max(...brands.map(b => b.mentions), 1);
    const barW = 1.0, gap = 0.55;
    const startX = 0.3, chartBottom = 4.65, chartH = 2.6;

    brands.forEach((brand, i) => {
      const x = startX + i * (barW + gap);
      const barH = Math.max((brand.mentions / maxM) * chartH, 0.15);
      const barY = chartBottom - barH;
      const isB = brand.name === d.brandName;
      const col = isB ? C.orange : (i === 0 ? C.teal : "BDBDCD");

      s.addText(brand.name, { x:x-0.1, y:barY-0.38, w:barW+0.2, h:0.3, fontSize:8, bold:isB, color:isB?C.orange:C.navy, align:"center", fontFace:"Calibri" });
      s.addShape(pres.shapes.OVAL, { x:x+barW/2-0.22, y:barY-0.62, w:0.44, h:0.44, fill:{color:C.white}, line:{color:C.lightgray} });
      s.addText("#"+brand.rank, { x:x+barW/2-0.22, y:barY-0.62, w:0.44, h:0.44, fontSize:11, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });
      s.addShape(pres.shapes.RECTANGLE, { x, y:barY, w:barW, h:barH, fill:{color:col}, line:{color:col}, shadow:makeShadow() });
      s.addText(brand.mentions+" mentions", { x:x-0.1, y:chartBottom+0.06, w:barW+0.2, h:0.2, fontSize:7.5, color:C.slate, align:"center", fontFace:"Calibri" });
      if (isB) s.addText("👑", { x:x+barW/2-0.24, y:barY+0.06, w:0.48, h:0.35, fontSize:18, align:"center" });
    });
    s.addText("Brand Leaderboard", { x:0.3, y:1.1, w:4.5, h:0.25, fontSize:9, bold:true, color:C.slate, fontFace:"Calibri" });
  }

  // RIGHT: Competitor mentions bar chart
  const comps = d.competitorMentions.slice(0, 10);
  const maxPct = Math.max(...comps.map(c => c.percentage), 1);
  const rowH = 0.33, startY = 1.1, barMaxW = 3.5;
  s.addText("Competitor Mentions vs. " + d.brandName, { x:5.0, y:1.1, w:4.7, h:0.25, fontSize:9, bold:true, color:C.slate, fontFace:"Calibri" });

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

  // Divider
  s.addShape(pres.shapes.RECTANGLE, { x:4.75, y:1.05, w:0.03, h:3.8, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}

// ─── SLIDE 4: Domains vs Brand + Brand Pages ──────────────────────────────────
function buildSlide4(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "Top Cited Sources (Category vs Us)", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  // LEFT: Domain citations table
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
    const pages = row.uniquePagesCited || row.pages || "";
    const responses = row.domainShare || row.responses || "";
    s.addText(String(pages), { x:2.82, y:y+0.06, w:0.6, h:0.22, fontSize:8, color:C.slate, fontFace:"Calibri" });
    s.addText(String(responses), { x:3.67, y:y+0.06, w:0.9, h:0.22, fontSize:8, color:C.slate, fontFace:"Calibri" });
  });

  // RIGHT: Brand pages
  const pages = d.brandPages.slice(0, 6);
  s.addText("Sources from " + d.brandName + " Domain", { x:5.0, y:1.08, w:4.7, h:0.28, fontSize:9, bold:true, color:C.navy, fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:5.0, y:1.36, w:4.7, h:0.02, fill:{color:C.lightgray}, line:{color:C.lightgray} });

  pages.forEach((pg, i) => {
    const y = 1.42 + i * 0.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.05, y, w:4.6, h:0.44, fill:{color:C.offwhite}, line:{color:C.lightgray}, rectRadius:0.04, shadow:makeShadow() });
    s.addText(pg.name, { x:5.15, y:y+0.04, w:3.6, h:0.2, fontSize:8.5, bold:true, color:C.navy, fontFace:"Calibri" });
    s.addText((pg.prompts||0)+" Response"+(pg.prompts!==1?"s":""), { x:9.1, y:y+0.04, w:0.5, h:0.2, fontSize:7.5, bold:true, color:C.teal, align:"right", fontFace:"Calibri" });
    s.addText(pg.url || d.domain, { x:5.15, y:y+0.24, w:4.4, h:0.16, fontSize:7, color:C.slate, fontFace:"Calibri" });
  });

  // Yellow CTA
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:5.0, y:4.75, w:4.7, h:0.34, fill:{color:C.yellow}, line:{color:C.yellow}, rectRadius:0.05 });
  s.addText("Top Cited Sources from our website ↑", { x:5.0, y:4.75, w:4.7, h:0.34, fontSize:9, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri" });

  // Divider
  s.addShape(pres.shapes.RECTANGLE, { x:4.75, y:1.05, w:0.03, h:3.9, fill:{color:C.lightgray}, line:{color:C.lightgray} });
}


// ─── SLIDE 5: Competitor Visibility Matrix (Theme × Brand heatmap) ────────────
function buildSlide5(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "Theme Benchmarks (% Visibility across Competitors)", d.brandName);
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
    s.addText(row.theme || "", { x:startX+0.05, y:y+0.05, w:themeColW-0.08, h:rowH-0.08, fontSize:6, color:C.navy, fontFace:"Calibri", wrap:true });
    allCols.forEach((col, ci) => {
      const x = startX + themeColW + ci * dataColW;
      const isBrand = col === d.brandName;
      const pct = isBrand
        ? (typeof row.brandVisibility === 'number' ? row.brandVisibility : 0)
        : (typeof row.competitors?.[col] === 'number' ? row.competitors[col] : 0);
      let cellCol = bg;
      if (isBrand && pct > 0) cellCol = C.purple;
      else if (pct >= 15) cellCol = "8B85D4";
      else if (pct >= 5) cellCol = "C4C0EA";
      if (cellCol !== bg) s.addShape(pres.shapes.RECTANGLE, { x:x+0.02, y:y+0.04, w:dataColW-0.04, h:rowH-0.08, fill:{color:cellCol}, line:{color:cellCol} });
      s.addText(pct > 0 ? pct+"%" : "0%", { x:x+0.02, y:y+0.05, w:dataColW-0.04, h:rowH-0.1, fontSize:7, bold:isBrand, color:(isBrand&&pct>0)?C.white:(pct>=8?C.white:C.slate), align:"center", fontFace:"Calibri" });
    });
  });

  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.", { x:0.3, y:5.22, w:9.4, h:0.16, fontSize:7, italic:true, color:C.slate, align:"center", fontFace:"Calibri" });
}

// ─── SLIDE 6: "What does each metric mean?" — static definitions ──────────────
function buildSlide6(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, "What does each of these mean?", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const defs = [
    { term: "Brand Mentions", body: "Number of times your brand appeared in AI answers out of total tracked prompts" },
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


// ─── SLIDE 7: Platform mentions table ────────────────────────────────────────
function buildSlide7(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, d.brandName + " Mentions by AI Platform", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  // Top 4 KPI boxes
  const kpis = [
    { v: String(d.totalMentions), l: "Total Brand Mentions" },
    { v: String(d.totalCitations), l: "Total Domain Citations" },
    { v: d.avgBrandCoverage, l: "Avg Brand Coverage" },
    { v: d.avgDomainCoverage, l: "Avg Domain Coverage" },
  ];
  kpis.forEach((k, i) => {
    const x = 0.25 + i * 2.42;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.08, w:2.3, h:0.72, fill:{color:C.white}, line:{color:C.lightgray}, shadow:makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.08, w:2.3, h:0.06, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText(k.v, { x, y:1.16, w:2.3, h:0.38, fontSize:20, bold:true, color:C.navy, align:"center", fontFace:"Calibri" });
    s.addText(k.l, { x, y:1.52, w:2.3, h:0.24, fontSize:7.5, color:C.slate, align:"center", fontFace:"Calibri" });
  });

  // Platform table
  const cols = ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"];
  const colX = [0.25, 2.15, 3.2, 4.3, 7.1];
  const colW = [1.85, 1.0, 1.05, 2.75, 2.6];

  // Header
  s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:1.9, w:9.5, h:0.28, fill:{color:C.lightgray}, line:{color:C.lightgray} });
  cols.forEach((h, i) => s.addText(h, { x:colX[i], y:1.93, w:colW[i], h:0.22, fontSize:7.5, bold:true, color:C.slate, fontFace:"Calibri" }));

  d.platforms.forEach((p, i) => {
    const y = 2.22 + i * 0.5;
    const bg = i % 2 === 0 ? "F4F3FD" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x:0.25, y:y-0.04, w:9.5, h:0.48, fill:{color:bg}, line:{color:C.lightgray} });
    s.addText(p.name, { x:colX[0], y, w:colW[0], h:0.28, fontSize:9.5, bold:true, color:C.navy, fontFace:"Calibri" });
    s.addText(String(p.mentions), { x:colX[1], y, w:colW[1], h:0.28, fontSize:9.5, color:C.navy, fontFace:"Calibri" });
    s.addText(String(p.citations), { x:colX[2], y, w:colW[2], h:0.28, fontSize:9.5, color:C.navy, fontFace:"Calibri" });

    // Brand visibility progress bar
    const bvPct = Math.min(p.brandVisibility || 0, 100);
    const bvW = (bvPct / 100) * 2.5;
    s.addShape(pres.shapes.RECTANGLE, { x:colX[3], y:y+0.08, w:2.5, h:0.14, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    if (bvW > 0) s.addShape(pres.shapes.RECTANGLE, { x:colX[3], y:y+0.08, w:bvW, h:0.14, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(bvPct+"%", { x:colX[3], y, w:0.45, h:0.28, fontSize:8, bold:true, color:C.navy, fontFace:"Calibri" });

    // Domain coverage progress bar
    const dcPct = Math.min(p.domainCoverage || 0, 100);
    const dcW = (dcPct / 100) * 2.5;
    s.addShape(pres.shapes.RECTANGLE, { x:colX[4], y:y+0.08, w:2.5, h:0.14, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    if (dcW > 0) s.addShape(pres.shapes.RECTANGLE, { x:colX[4], y:y+0.08, w:dcW, h:0.14, fill:{color:C.violet}, line:{color:C.violet} });
    s.addText(dcPct+"%", { x:colX[4], y, w:0.45, h:0.28, fontSize:8, bold:true, color:C.violet, fontFace:"Calibri" });
  });
}

// ─── SLIDE 8: Brand Visibility by Platform × Theme heatmap ─────────────────
function buildSlide8(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  hdr(s, pres, d.brandName + " Brand Visibility by Platform & Theme", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  // Data: brandVisibilityByPlatform — from "[Brand] brand visibility" table on /platforms
  // Each row: { theme: "...", "ChatGPT": 10, "Google AI Overview": 10, "Perplexity": 0 }
  const rows = d.brandVisibilityByPlatform;
  if (!rows || rows.length === 0) {
    s.addText("No platform visibility data available.", { x:0.5, y:2.8, w:9, h:0.5, fontSize:12, color:C.slate, align:"center", fontFace:"Calibri" });
    return;
  }

  const platNames = Object.keys(rows[0]).filter(k => k !== 'theme');
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
      const pct = typeof row[pn] === 'number' ? row[pn] : 0;
      let cellCol = bg;
      if (pct >= 15) cellCol = C.purple;
      else if (pct >= 5) cellCol = "C4C0EA";
      if (cellCol !== bg) s.addShape(pres.shapes.RECTANGLE, { x:x+0.04, y:y+0.05, w:dataColW-0.08, h:rowH-0.1, fill:{color:cellCol}, line:{color:cellCol} });
      s.addText(pct > 0 ? pct+"%" : "0%", { x:x+0.04, y:y+0.07, w:dataColW-0.08, h:rowH-0.12, fontSize:8, color:pct>=5?C.white:C.slate, align:"center", fontFace:"Calibri" });
    });
  });

  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.", { x:0.3, y:5.22, w:9.4, h:0.16, fontSize:7, italic:true, color:C.slate, align:"center", fontFace:"Calibri" });
}

// ─── STEP 4: Build the PPTX ───────────────────────────────────────────────────
function buildPPTX(data, outputPath) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = data.brandName + " GEO Audit — Atlas";
  pres.author = "Pepper.inc Atlas";

  buildSlide1(pres, data);   // Cover
  buildSlide2(pres, data);   // Prompts & Themes grid
  buildSlide3(pres, data);   // Leaderboard + Competitor mentions
  buildSlide4(pres, data);   // Domains vs brand + brand pages
  buildSlide5(pres, data);   // Competitor visibility matrix heatmap
  buildSlide6(pres, data);   // Metric definitions (static)
  buildSlide7(pres, data);   // Platform mentions table
  buildSlide8(pres, data);   // Brand visibility by platform heatmap

  pres.writeFile({ fileName: outputPath });
  console.log("✅ PPTX written:", outputPath);
}

// ─── API: POST /generate ──────────────────────────────────────────────────────
app.post("/generate", async (req, res) => {
  const { url } = req.body;
  if (!url || !url.includes("atlas.pepper.inc")) {
    return res.status(400).json({ error: "Please provide a valid Atlas report URL." });
  }

  const id = uuidv4();
  const pptxOut = path.join(TMP, id + ".pptx");
  const pdfOut  = path.join(TMP, id + ".pdf");

  try {
    // Step 1 — Scrape: visit all 4 tabs, take full-page screenshots
    const screenshots = await scrapeAtlasReport(url);

    // Step 2 — Extract: send screenshots to Claude Vision, get structured JSON
    const rawData = await extractData(screenshots);

    // Step 3 — Normalize: fill defaults so nothing crashes
    const data = normalizeData(rawData);

    // Step 4 — Build: generate PPTX from structured data
    buildPPTX(data, pptxOut);

    // Step 5 — Convert to PDF
    await new Promise(r => setTimeout(r, 1000));
    console.log("📄 Converting to PDF...");
    execSync(`soffice --headless --convert-to pdf --outdir ${TMP} ${pptxOut}`, { timeout: 60000 });

    if (!fs.existsSync(pdfOut)) throw new Error("PDF conversion failed");

    const brandSlug = data.brandName.toLowerCase().replace(/\s+/g, "-");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${brandSlug}-geo-audit.pdf"`);
    fs.createReadStream(pdfOut).pipe(res);
    res.on("finish", () => {
      try { fs.unlinkSync(pptxOut); fs.unlinkSync(pdfOut); } catch {}
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
