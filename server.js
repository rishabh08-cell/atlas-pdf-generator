const express = require("express");
const pptxgen = require("pptxgenjs");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const multer = require("multer")
const XLSX = require("xlsx");
const JSZip = require("jszip");

const app = express()
app.use(express.json());
app.use(express.static("public"));

const TMP = path.join(__dirname, "tmp");
if (!fs.existsSync(TMP)) fs.mkdirSync(TMP);

const UPLOADS = path.join(TMP, "uploads");
if (!fs.existsSync(UPLOADS)) fs.mkdirSync(UPLOADS);

const upload = multer({ dest: UPLOADS, limits: { fileSize: 5 * 1024 * 1024 } });

const MAX_BULK_REPORTS = 25;

// ─── Pepper brand colour palette ──────────────────────────────────────────────
const C = {
  navy: "0D007D",
  purple: "3D35B5",
  violet: "7B7FD4",
  lilac: "A8ABEA",
  orange: "EE7F51",
  teal: "0B7251",
  green: "0E9468",
  yellow: "F9B02A",
  white: "FFFFFF",
  offwhite: "F5F5F8",
  slate: "64748B",
  lightgray: "E2E1F0",
  darkgray: "1A1650",
};

function makeShadow() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 };
}

// ─── STEP 1: Fetch data from Atlas API ─────────────────────────────────────
const API_BASE = "https://hub.peppercontent.io/atlas-service/api/public/reports";

function extractReportId(input) {
  const uuidPattern = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
  const match = input.match(uuidPattern);
  if (match) return match[0];
  return input.trim();
}

async function fetchReportData(reportId) {
  console.log("\n📡 Fetching report data for:", reportId);
  const endpoints = {
    overview: "overview",
    competitors: "competitors-comparison",
    platforms: "platforms",
    prompts: "prompts-themes",
  };
  const results = {};

  for (const [key, ep] of Object.entries(endpoints)) {
    const url = `${API_BASE}/${reportId}/${ep}`;
    console.log(`  ⬇️ ${key}: ${url}`);
    const res = await fetch(url);
    if (!res.ok) { const err = new Error(`API error for ${ep}: ${res.status} ${res.statusText}`); err.statusCode = res.status; throw err; }
    results[key] = await res.json();
    console.log(`  ✅ ${key} fetched`);
  }

  return results;
}

// ─── STEP 2: Normalize API data into slide-ready shape ──────────────────────
function normalizeData(api) {
  const { overview, competitors, platforms, prompts } = api;
  if (overview.status === "failed") throw new Error("This report failed to generate. Please re-run the audit in Pepper.");
  const brand = overview.brand || { name: "Unknown", domain: "" };

  const totalMentions = parseInt(platforms.stats.find(s => s.label === "Total Brand Mentions")?.value) || 0;
  const totalCitations = parseInt(platforms.stats.find(s => s.label === "Total Domain Citations")?.value) || 0;
  const avgBrandCoverage = platforms.stats.find(s => s.label === "Brand Visibility")?.value || "0%";
  const avgDomainCoverage = platforms.stats.find(s => s.label === "Domain Prompt Presence")?.value || "0%";

  const sortedCompetitors = [...(overview.competitors || [])].sort((a, b) => a.rank - b.rank);
  const userCompetitor = (overview.competitors || []).find(c => c.is_user);
  const leaderboardRank = userCompetitor ? "#" + userCompetitor.rank : "#N/A";

  const leaderboard = sortedCompetitors.slice(0, 3).map(c => ({
    rank: c.rank,
    name: c.name,
    mentions: c.mentions,
  }));

  const competitorMentions = sortedCompetitors.slice(0, 10).map(c => ({
    name: c.name,
    percentage: Math.round(c.mention_rate * 10) / 10,
    mentions: c.mentions,
  }));

  const platformRows = (platforms.rows || []).map(r => ({
    name: r.platform,
    mentions: r.mentions,
    citations: r.citations,
    brandVisibility: Math.round((r.brand_visibility || 0) * 10) / 10,
    domainCoverage: Math.round((r.domain_coverage || 0) * 10) / 10,
  }));

  const promptThemes = (prompts.themes || []).map(t => ({
    theme: t.name,
    promptCount: t.prompts.length,
    prompts: t.prompts.map(p => p.text),
  }));

  const domainCitations = (overview.domains || []).slice(0, 10).map(d => ({
    domain: d.domain,
    uniquePagesCited: d.total_pages_cited,
    domainCoverage: Math.round(d.mention_rate * 10) / 10 + "%",
    domainShare: Math.round(d.share_of_voice * 10) / 10 + "%",
  }));

  const brandPages = (overview.pages || []).slice(0, 8).map(p => ({
    name: p.title || p.url,
    url: p.url,
    prompts: p.prompts,
  }));

  const compVis = competitors.visibility || {};
  const compVisThemes = compVis.themes || [];
  const compVisCompetitors = compVis.competitors || [];
  const compVisValues = compVis.values || [];
  const brandCompIdx = compVisCompetitors.findIndex(c => c.name === brand.name);

  const competitorVisibilityMatrix = compVisThemes.map((theme, ti) => {
    const row = {
      theme,
      brandVisibility: (compVisValues[ti] && brandCompIdx >= 0) ? (compVisValues[ti][brandCompIdx] || 0) : 0,
      competitors: {},
    };
    compVisCompetitors.forEach((comp, ci) => {
      if (comp.name !== brand.name) {
        row.competitors[comp.name] = (compVisValues[ti] && compVisValues[ti][ci]) || 0;
      }
    });
    return row;
  });

  const platVis = platforms.visibility || {};
  const platVisThemes = platVis.themes || [];
  const platVisPlatforms = platVis.platforms || [];
  const platVisValues = platVis.values || [];

  const brandVisibilityByPlatform = platVisThemes.map((theme, ti) => {
    const row = { theme };
    platVisPlatforms.forEach((pn, pi) => {
      row[pn] = (platVisValues[ti] && platVisValues[ti][pi]) || 0;
    });
    return row;
  });

  const cleanDomain = brand.domain.replace(/^https?:\/\//, "").replace(/\/$/, "");

  console.log("  Brand:", brand.name, "| Platforms:", platformRows.length, "| Themes:", promptThemes.length, "| Matrix:", competitorVisibilityMatrix.length, "| VisRows:", brandVisibilityByPlatform.length);

  return {
    brandName: brand.name || "Brand",
    domain: cleanDomain,
    totalMentions,
    totalCitations,
    avgBrandCoverage,
    avgDomainCoverage,
    leaderboardRank,
    platformCount: platformRows.length,
    leaderboard,
    competitorMentions,
    platforms: platformRows.length > 0 ? platformRows : [
      { name: "ChatGPT", mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
      { name: "Google AI Overview", mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
      { name: "Perplexity", mentions: 0, citations: 0, brandVisibility: 0, domainCoverage: 0 },
    ],
    promptThemes,
    domainCitations,
    brandPages,
    competitorVisibilityMatrix,
    brandVisibilityByPlatform,
  };
}

// ─── PPTX layout helpers ───────────────────────────────────────────────────
function logoPill(s, pres) {
  const pepperLogoPath = path.join(__dirname, "public", "logos", "pepper-logo.png");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:7.551,y:0.128,w:0.595,h:0.256, fill:{color:C.white}, line:{color:C.lightgray,pt:1}, rectRadius:0.05, shadow:makeShadow() });
  if (fs.existsSync(pepperLogoPath)) {
    s.addImage({ path:pepperLogoPath, x:7.592,y:0.16,w:0.513,h:0.192, sizing:{type:"contain",w:0.513,h:0.192} });
  } else {
    s.addText("pepper", { x:7.592,y:0.16,w:0.513,h:0.192, fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri" });
  }
}

function hdr(s, pres, title, brandName) {
  s.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:8.27,h:0.085, fill:{color:C.teal}, line:{color:C.teal} });
  s.addText("pepper", { x:0.248,y:0.128,w:0.744,h:0.32, fontSize:11,bold:true,color:C.navy,fontFace:"Calibri" });
  s.addText("by pepper.inc", { x:0.992,y:0.181,w:1.075,h:0.235, fontSize:7,color:C.slate,fontFace:"Calibri" });
  logoPill(s, pres);
  s.addText(title, { x:0.248,y:0.533,w:5.954,h:0.512, fontSize:18,bold:true,color:C.navy,fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.248,y:1.013,w:1.489,h:0.032, fill:{color:C.teal}, line:{color:C.teal} });
}

function ftr(s, pres, brand, domain) {
  s.addShape(pres.shapes.RECTANGLE, { x:0,y:5.781,w:8.27,h:0.213, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText(brand+" · "+domain+" · GEO Audit by Pepper", { x:0.248,y:5.792,w:7.774,h:0.192, fontSize:6.5,color:"AAAACC",fontFace:"Calibri" });
}

function staticHdr(s, pres, title, brandName) {
  logoPill(s, pres);
  s.addText(title, { x:0.289,y:0.128,w:5.913,h:0.448, fontSize:18,bold:true,color:C.navy,fontFace:"Calibri" });
  s.addShape(pres.shapes.RECTANGLE, { x:0.289,y:0.597,w:2.895,h:0.027, fill:{color:C.navy}, line:{color:C.navy} });
}

// ─── SLIDE 1: Cover ────────────────────────────────────────────────────────
function buildSlide1(pres, d) {
  const s=pres.addSlide(); s.background={color:C.navy};
  s.addShape(pres.shapes.OVAL,{x:6.203,y:-0.5,w:2.895,h:3.733,fill:{color:"150050"},line:{color:"150050"}});
  s.addShape(pres.shapes.OVAL,{x:6.781,y:0.213,w:1.654,h:2.133,fill:{color:C.purple},line:{color:C.purple}});
  s.addText("pepper",{x:0.414,y:0.405,w:0.91,h:0.405,fontSize:15,bold:true,color:C.orange,fontFace:"Calibri"});
  s.addText("by pepper.inc",{x:1.348,y:0.469,w:1.241,h:0.277,fontSize:8,color:C.lilac,fontFace:"Calibri"});
  const pepperLogoPath = path.join(__dirname, "public", "logos", "pepper-logo.png");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:7.551,y:0.373,w:0.595,h:0.256,fill:{color:"FFFFFF"},line:{color:"FFFFFF"},rectRadius:0.066});
  if (fs.existsSync(pepperLogoPath)) {
    s.addImage({path:pepperLogoPath,x:7.592,y:0.405,w:0.513,h:0.192,sizing:{type:"contain",w:0.513,h:0.192}});
  } else {
    s.addText("pepper",{x:7.592,y:0.405,w:0.513,h:0.192,fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
  }
  s.addText("GEO AUDIT REPORT",{x:0.414,y:1.152,w:4.962,h:0.299,fontSize:9,color:C.orange,bold:true,charSpacing:4,fontFace:"Calibri"});
  s.addText(d.brandName,{x:0.414,y:1.472,w:5.789,h:1.259,fontSize:50,bold:true,color:C.white,fontFace:"Calibri"});
  s.addText(d.domain,{x:0.414,y:2.773,w:4.135,h:0.405,fontSize:13,color:C.lilac,fontFace:"Calibri"});
  s.addShape(pres.shapes.RECTANGLE,{x:0.414,y:3.221,w:0.992,h:0.043,fill:{color:C.orange},line:{color:C.orange}});
  [{v:String(d.totalMentions),l:"Total Mentions"},{v:d.avgBrandCoverage,l:"Brand Visibility"},{v:String(d.platformCount),l:"AI Platforms"},{v:d.leaderboardRank,l:"Leaderboard"}].forEach((k,i)=>{
    const x=0.5+i*2.3;
    s.addText(k.v,{x,y:3.371,w:1.737,h:0.555,fontSize:24,bold:true,color:C.orange,fontFace:"Calibri"});
    s.addText(k.l,{x,y:3.904,w:1.737,h:0.235,fontSize:8,color:C.lilac,fontFace:"Calibri"});
  });
  s.addText("Powered by pepper",{x:0.414,y:5.493,w:7.443,h:0.235,fontSize:7.5,color:C.slate,fontFace:"Calibri"});
}

// ─── SLIDE 2: Prompts & Themes ─────────────────────────────────────────────
function buildSlide2(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  const tp=d.promptThemes.reduce((a,t)=>a+(t.promptCount||t.prompts?.length||0),0);
  hdr(s,pres,`We Have Mapped ${tp} Prompts Across ${d.promptThemes.length} Themes`,d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  d.promptThemes.slice(0,9).forEach((t,i)=>{
    const col=i%3,row=Math.floor(i/3),x=0.25+col*2.685,y=1.1+row*0.82;
    s.addShape(pres.shapes.RECTANGLE,{x,y,w:2.4,h:0.768,fill:{color:C.offwhite},line:{color:C.lightgray},shadow:makeShadow()});
    s.addShape(pres.shapes.RECTANGLE,{x,y,w:0.033,h:0.768,fill:{color:C.teal},line:{color:C.teal}});
    s.addText(t.theme,{x:x+0.1,y:y+0.08,w:2.276,h:0.363,fontSize:9.5,bold:true,color:C.navy,fontFace:"Calibri",wrap:true});
    s.addText((t.promptCount||t.prompts?.length||0)+" prompts",{x:x+0.1,y:y+0.46,w:2.276,h:0.213,fontSize:8.5,color:C.slate,fontFace:"Calibri"});
  });
}

// ─── SLIDE 3: Leaderboard + Competitors ───────────────────────────────────
function buildSlide3(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  hdr(s,pres,"Brand Leaderboard & Competitor Mentions",d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  const brands=d.leaderboard;
  if (brands.length>0) {
    const maxM=Math.max(...brands.map(b=>b.mentions),1),barW=0.85,gap=0.3,startX=0.25,cb=4.55,ch=2.4;
    brands.forEach((brand,i)=>{
      const x=startX+i*(barW+gap),barH=Math.max((brand.mentions/maxM)*ch,0.15),barY=Math.max(cb-barH,1.55),isB=brand.name===d.brandName;
      const col=isB?C.orange:(i===0?C.teal:"BDBDCD");
      s.addShape(pres.shapes.OVAL,{x:x+barW/2-0.22,y:barY-0.96,w:0.364,h:0.469,fill:{color:C.white},line:{color:C.lightgray}});
      s.addText("#"+brand.rank,{x:x+barW/2-0.22,y:barY-0.96,w:0.364,h:0.469,fontSize:11,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
      s.addText(brand.name,{x:x-0.1,y:barY-0.5,w:barW+0.2,h:0.363,fontSize:8,bold:isB,color:isB?C.orange:C.navy,align:"center",fontFace:"Calibri",wrap:true});
      s.addShape(pres.shapes.RECTANGLE,{x,y:barY,w:barW,h:barH,fill:{color:col},line:{color:col},shadow:makeShadow()});
      s.addText(brand.mentions+" mentions",{x:x-0.1,y:cb+0.06,w:barW+0.2,h:0.213,fontSize:7.5,color:C.slate,align:"center",fontFace:"Calibri"});
    });
  }
  const comps=d.competitorMentions.slice(0,10),maxPct=Math.max(...comps.map(c=>c.percentage),1);
  s.addText("Competitor Mentions vs. "+d.brandName,{x:3.7,y:1.0,w:4.35,h:0.267,fontSize:9,bold:true,color:C.slate,fontFace:"Calibri"});
  comps.forEach((comp,i)=>{
    const y=1.42+i*0.33,isB=comp.name===d.brandName;
    s.addShape(pres.shapes.OVAL,{x:3.7,y:y+0.04,w:0.182,h:0.235,fill:{color:isB?C.purple:C.lightgray},line:{color:isB?C.purple:C.lightgray}});
    s.addText(comp.name[0].toUpperCase(),{x:3.7,y:y+0.04,w:0.182,h:0.235,fontSize:7,bold:true,color:isB?C.white:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
    s.addText(comp.name,{x:3.92,y:y+0.05,w:1.158,h:0.213,fontSize:8,bold:isB,color:isB?C.purple:C.navy,fontFace:"Calibri"});
    const bw=Math.min((comp.percentage/maxPct)*2.0,2.0);
    s.addShape(pres.shapes.RECTANGLE,{x:5.1,y:y+0.06,w:Math.max(bw,0.05),h:0.192,fill:{color:isB?C.navy:C.lightgray},line:{color:isB?C.navy:C.lightgray}});
    s.addText(comp.percentage+"% · "+comp.mentions+" mentions",{x:5.2+bw,y:y+0.05,w:Math.max(8.05-(5.2+bw),0.85),h:0.213,fontSize:7.5,color:C.slate,fontFace:"Calibri"});
  });
  s.addShape(pres.shapes.RECTANGLE,{x:3.55,y:1.12,w:0.02,h:3.9,fill:{color:C.lightgray},line:{color:C.lightgray}});
}

// ─── SLIDE 4: Top Cited Sources ────────────────────────────────────────────
function buildSlide4(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  hdr(s,pres,"Top Cited Sources (Category vs Us)",d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  const domains=d.domainCitations.slice(0,9);
  s.addShape(pres.shapes.RECTANGLE,{x:0.207,y:1.152,w:3.639,h:0.299,fill:{color:C.navy},line:{color:C.navy}});
  [["Domain",0.289,1.985],["Pages",2.332,0.496],["Responses",3.035,0.744]].forEach(([h,x,w])=>s.addText(h,{x,y:1.173,w,h:0.256,fontSize:8,bold:true,color:C.white,fontFace:"Calibri"}));
  domains.forEach((row,i)=>{
    const y=1.38+i*0.34,bg=i%2===0?C.offwhite:C.white;
    s.addShape(pres.shapes.RECTANGLE,{x:0.207,y,w:3.639,h:0.341,fill:{color:bg},line:{color:C.lightgray}});
    s.addText(row.domain,{x:0.289,y:y+0.06,w:1.985,h:0.235,fontSize:8,color:C.navy,fontFace:"Calibri"});
    s.addText(String(row.uniquePagesCited||row.pages||""),{x:2.332,y:y+0.06,w:0.496,h:0.235,fontSize:8,color:C.slate,fontFace:"Calibri"});
    s.addText(String(row.domainShare||row.responses||""),{x:3.035,y:y+0.06,w:0.744,h:0.235,fontSize:8,color:C.slate,fontFace:"Calibri"});
  });
  const pages=d.brandPages.slice(0,6);
  s.addText("Sources from "+d.brandName+" Domain",{x:4.135,y:1.152,w:3.887,h:0.299,fontSize:9,bold:true,color:C.navy,fontFace:"Calibri"});
  s.addShape(pres.shapes.RECTANGLE,{x:4.135,y:1.451,w:3.887,h:0.021,fill:{color:C.lightgray},line:{color:C.lightgray}});
  if(!pages||pages.length===0){s.addText("No pages from "+d.brandName+" domain were cited in AI responses.",{x:4.176,y:2.667,w:3.804,h:1.067,fontSize:8.5,italic:true,color:C.slate,align:"center",valign:"middle",fontFace:"Calibri"});}else{ pages.forEach((pg,i)=>{
    const y=1.42+i*0.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:4.176,y,w:3.804,h:0.469,fill:{color:C.offwhite},line:{color:C.lightgray},rectRadius:0.033,shadow:makeShadow()});
    s.addText(pg.name,{x:4.259,y:y+0.04,w:3.639,h:0.213,fontSize:8.5,bold:true,color:C.navy,fontFace:"Calibri"});
    s.addText(pg.url||d.domain,{x:4.259,y:y+0.24,w:3.639,h:0.171,fontSize:7,color:C.slate,fontFace:"Calibri"});
  });} s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:4.135,y:5.067,w:3.887,h:0.363,fill:{color:C.yellow},line:{color:C.yellow},rectRadius:0.041});
  s.addText("Top Cited Sources from our website ↑",{x:4.135,y:5.067,w:3.887,h:0.363,fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
  s.addShape(pres.shapes.RECTANGLE,{x:3.928,y:1.12,w:0.025,h:4.16,fill:{color:C.lightgray},line:{color:C.lightgray}});
}

// ─── SLIDE 5: Competitor Visibility Matrix ─────────────────────────────────
function buildSlide5(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  hdr(s,pres,"Theme Benchmarks (% Visibility across Competitors)",d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  const matrix=d.competitorVisibilityMatrix;
  if (!matrix||matrix.length===0){s.addText("No competitor visibility matrix data available.",{x:0.414,y:2.667,w:7.443,h:0.533,fontSize:12,color:C.slate,align:"center",fontFace:"Calibri"});return;}
  const compNames=[];
  matrix.forEach(row=>{if(row.competitors)Object.keys(row.competitors).forEach(k=>{if(!compNames.includes(k))compNames.push(k);});});
  const allCols=[d.brandName,...compNames].slice(0,10),themeColW=1.9,dataColW=(7.857-themeColW)/allCols.length,startX=0.25,headerY=1.08,rowH=0.32;
  s.addShape(pres.shapes.RECTANGLE,{x:startX,y:headerY,w:7.857,h:rowH,fill:{color:C.navy},line:{color:C.navy}});
  s.addText("Topic",{x:startX+0.05,y:headerY+0.05,w:themeColW-0.08,h:rowH-0.08,fontSize:7,bold:true,color:C.white,fontFace:"Calibri"});
  allCols.forEach((col,ci)=>{
    const x=startX+themeColW+ci*dataColW,isB=col===d.brandName;
    s.addText(col,{x:x+0.02,y:headerY+0.04,w:dataColW-0.04,h:rowH-0.08,fontSize:6,bold:isB,color:isB?C.orange:C.white,align:"center",fontFace:"Calibri",wrap:true});
  });
  matrix.slice(0,11).forEach((row,ri)=>{
    const y=headerY+rowH+ri*rowH,bg=ri%2===0?C.offwhite:C.white;
    s.addShape(pres.shapes.RECTANGLE,{x:startX,y,w:7.857,h:rowH,fill:{color:bg},line:{color:C.lightgray}});
    s.addText(row.theme||"",{x:startX+0.05,y:y+0.05,w:themeColW-0.08,h:rowH-0.08,fontSize:6,color:C.navy,fontFace:"Calibri",wrap:true});
    allCols.forEach((col,ci)=>{
      const x=startX+themeColW+ci*dataColW,isB=col===d.brandName;
      const pct=isB?(typeof row.brandVisibility==='number'?row.brandVisibility:0):(typeof row.competitors?.[col]==='number'?row.competitors[col]:0);
      let cc=bg;
      if(isB&&pct>0)cc=C.purple; else if(pct>=15)cc="8B85D4"; else if(pct>=5)cc="C4C0EA";
      if(cc!==bg)s.addShape(pres.shapes.RECTANGLE,{x:x+0.02,y:y+0.04,w:dataColW-0.04,h:rowH-0.08,fill:{color:cc},line:{color:cc}});
      s.addText(pct>0?pct+"%":"0%",{x:x+0.02,y:y+0.05,w:dataColW-0.04,h:rowH-0.1,fontSize:7,bold:isB,color:(isB&&pct>0)?C.white:(pct>=8?C.white:C.slate),align:"center",fontFace:"Calibri"});
    });
  });
  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.",{x:0.248,y:5.568,w:7.774,h:0.171,fontSize:7,italic:true,color:C.slate,align:"center",fontFace:"Calibri"});
}

// ─── SLIDE 6: Metric Definitions ───────────────────────────────────────────
function buildSlide6(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  hdr(s,pres,"What does each of these mean?",d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  [{term:"Brand Mentions",body:"Number of times your brand appeared in AI answers out of total tracked prompts"},
   {term:"Share of Voice",body:"Percentage of your brand mentions compared to all total brand mentions"},
   {term:"Brand Position",body:"Average position of your brand in AI answers"},
   {term:"Domain Citation",body:"Number of times your website was cited on AI Search Engines"},
   {term:"Brand Visibility",body:"Percentage of prompts that mention your brand"},
   {term:"Domain Coverage",body:"Percentage of prompts that cited your website"},
  ].forEach((def,i)=>{
    const col=i%3,row=Math.floor(i/3),x=0.28+col*3.22,y=1.2+row*1.62;
    s.addShape(pres.shapes.RECTANGLE,{x,y,w:2.531,h:1.621,fill:{color:C.yellow},line:{color:"D4AA30"},shadow:makeShadow()});
    s.addText(def.term,{x:x+0.14,y:y+0.14,w:2.299,h:0.32,fontSize:12,bold:true,color:C.purple,fontFace:"Calibri"});
    s.addText("— — — — — — — — — — — —",{x:x+0.14,y:y+0.44,w:2.299,h:0.171,fontSize:7,color:C.purple,fontFace:"Calibri"});
    s.addText(def.body,{x:x+0.14,y:y+0.6,w:2.299,h:0.832,fontSize:9,color:C.navy,italic:true,bold:true,fontFace:"Calibri",wrap:true});
  });
  s.addText("Source: Otterly.ai",{x:0.248,y:5.568,w:2.481,h:0.171,fontSize:7.5,bold:true,color:C.navy,fontFace:"Calibri"});
}

// ─── SLIDE 7: Platform mentions table ──────────────────────────────────────
function buildSlide7(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  hdr(s,pres,d.brandName+" Mentions by AI Platform",d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  [{v:String(d.totalMentions),l:"Total Brand Mentions"},{v:String(d.totalCitations),l:"Total Domain Citations"},{v:d.avgBrandCoverage,l:"Brand Visibility"},{v:d.avgDomainCoverage,l:"Domain Prompt Presence"}].forEach((k,i)=>{
    const x=0.25+i*2.42;
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.152,w:1.902,h:0.768,fill:{color:C.white},line:{color:C.lightgray},shadow:makeShadow()});
    s.addShape(pres.shapes.RECTANGLE,{x,y:1.152,w:1.902,h:0.064,fill:{color:C.teal},line:{color:C.teal}});
    s.addText(k.v,{x,y:1.237,w:1.902,h:0.405,fontSize:20,bold:true,color:C.navy,align:"center",fontFace:"Calibri"});
    s.addText(k.l,{x,y:1.621,w:1.902,h:0.256,fontSize:7.5,color:C.slate,align:"center",fontFace:"Calibri"});
  });
  const colX=[0.25,2.15,3.2,4.3,7.1],colW=[1.85,1.0,1.05,2.75,2.6];
  s.addShape(pres.shapes.RECTANGLE,{x:0.207,y:2.027,w:7.857,h:0.299,fill:{color:C.lightgray},line:{color:C.lightgray}});
  ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"].forEach((h,i)=>s.addText(h,{x:colX[i],y:2.059,w:colW[i],h:0.235,fontSize:7.5,bold:true,color:C.slate,fontFace:"Calibri"}));
  d.platforms.forEach((p,i)=>{
    const y=2.22+i*0.5,bg=i%2===0?"F4F3FD":C.white;
    s.addShape(pres.shapes.RECTANGLE,{x:0.207,y:y-0.04,w:7.857,h:0.512,fill:{color:bg},line:{color:C.lightgray}});
    s.addText(p.name,{x:colX[0],y,w:colW[0],h:0.299,fontSize:9.5,bold:true,color:C.navy,fontFace:"Calibri"});
    s.addText(String(p.mentions),{x:colX[1],y,w:colW[1],h:0.299,fontSize:9.5,color:C.navy,fontFace:"Calibri"});
    s.addText(String(p.citations),{x:colX[2],y,w:colW[2],h:0.299,fontSize:9.5,color:C.navy,fontFace:"Calibri"});
    const bv=Math.min(parseFloat(String(p.brandVisibility||0))||0,100),bvW=(bv/100)*2.0;
    s.addShape(pres.shapes.RECTANGLE,{x:colX[3]+0.48,y:y+0.08,w:1.654,h:0.149,fill:{color:C.lightgray},line:{color:C.lightgray}});
    if(bvW>0)s.addShape(pres.shapes.RECTANGLE,{x:colX[3]+0.48,y:y+0.08,w:bvW,h:0.149,fill:{color:C.navy},line:{color:C.navy}});
    s.addText(bv+"%",{x:colX[3],y,w:0.372,h:0.299,fontSize:8,bold:true,color:C.navy,fontFace:"Calibri"});
    const dc=Math.min(parseFloat(String(p.domainCoverage||0))||0,100),dcW=(dc/100)*2.0;
    s.addShape(pres.shapes.RECTANGLE,{x:colX[4]+0.48,y:y+0.08,w:1.654,h:0.149,fill:{color:C.lightgray},line:{color:C.lightgray}});
    if(dcW>0)s.addShape(pres.shapes.RECTANGLE,{x:colX[4]+0.48,y:y+0.08,w:dcW,h:0.149,fill:{color:C.violet},line:{color:C.violet}});
    s.addText(dc+"%",{x:colX[4],y,w:0.372,h:0.299,fontSize:8,bold:true,color:C.violet,fontFace:"Calibri"});
  });
}

// ─── SLIDE 8: Brand Visibility by Platform ─────────────────────────────────
function buildSlide8(pres, d) {
  const _rows = d.brandVisibilityByPlatform;
  if (!_rows || _rows.length === 0 || _rows.every(r => Object.keys(r).length <= 1)) return;
  const s=pres.addSlide(); s.background={color:C.white};
  hdr(s,pres,d.brandName+" Brand Visibility by Platform & Theme",d.brandName);
  ftr(s,pres,d.brandName,d.domain);
  const rows=d.brandVisibilityByPlatform;
  if(!rows||rows.length===0){s.addText("No platform visibility data available.",{x:0.414,y:2.987,w:7.443,h:0.533,fontSize:12,color:C.slate,align:"center",fontFace:"Calibri"});return;}
  const platNames=Object.keys(rows[0]).filter(k=>k!=='theme'),themeColW=3.2,dataColW=(9.5-themeColW)/platNames.length,startX=0.25,headerY=1.08,rowH=0.35;
  s.addShape(pres.shapes.RECTANGLE,{x:startX,y:headerY,w:7.857,h:rowH,fill:{color:C.navy},line:{color:C.navy}});
  s.addText("Themes",{x:startX+0.08,y:headerY+0.07,w:themeColW-0.12,h:rowH-0.1,fontSize:8,bold:true,color:C.white,fontFace:"Calibri"});
  platNames.forEach((pn,pi)=>{
    const x=startX+themeColW+pi*dataColW;
    s.addText(pn,{x:x+0.04,y:headerY+0.05,w:dataColW-0.08,h:rowH-0.1,fontSize:8,bold:true,color:C.white,align:"center",fontFace:"Calibri",wrap:true});
  });
  rows.slice(0,11).forEach((row,ri)=>{
    const y=headerY+rowH+ri*rowH,bg=ri%2===0?C.offwhite:C.white;
    s.addShape(pres.shapes.RECTANGLE,{x:startX,y,w:7.857,h:rowH,fill:{color:bg},line:{color:C.lightgray}});
    s.addText(row.theme||"",{x:startX+0.08,y:y+0.07,w:themeColW-0.14,h:rowH-0.1,fontSize:7,color:C.navy,fontFace:"Calibri",wrap:true});
    platNames.forEach((pn,pi)=>{
      const x=startX+themeColW+pi*dataColW,pct=parseFloat(String(row[pn]||0))||0;
      let cc=bg;
      if(pct>=15)cc=C.purple; else if(pct>=5)cc="C4C0EA";
      if(cc!==bg)s.addShape(pres.shapes.RECTANGLE,{x:x+0.04,y:y+0.05,w:dataColW-0.08,h:rowH-0.1,fill:{color:cc},line:{color:cc}});
      s.addText(pct>0?pct+"%":"0%",{x:x+0.04,y:y+0.07,w:dataColW-0.08,h:rowH-0.12,fontSize:8,color:pct>=5?C.white:C.slate,align:"center",fontFace:"Calibri"});
    });
  });
  s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.",{x:0.248,y:5.568,w:7.774,h:0.171,fontSize:7,italic:true,color:C.slate,align:"center",fontFace:"Calibri"});
}

// ═══════════════════════════════════════════════════════════════════════════
// STATIC SLIDES 12–18
// ═══════════════════════════════════════════════════════════════════════════
function buildSlide12(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  staticHdr(s,pres,"The Approach For Solving GEO",d.brandName);
  s.addShape(pres.shapes.OVAL,{x:2.564,y:0.96,w:2.977,h:3.84,fill:{color:C.lightgray},line:{color:"CCCCDD",pt:1}});
  s.addShape(pres.shapes.OVAL,{x:2.895,y:1.387,w:2.316,h:2.987,fill:{color:C.white},line:{color:C.white}});
  s.addShape(pres.shapes.OVAL,{x:3.076,y:1.621,w:1.952,h:2.517,fill:{color:"E8E4FF"},line:{color:C.violet,pt:1}});
  s.addShape(pres.shapes.OVAL,{x:3.308,y:1.92,w:1.489,h:1.92,fill:{color:C.white},line:{color:C.white}});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:3.374,y:2.347,w:1.356,h:1.067,fill:{color:C.yellow},line:{color:C.yellow},rectRadius:0.066});
  s.addText("Pepper's GEO\nApproach",{x:3.374,y:2.347,w:1.356,h:1.067,fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
  s.addText("Visibility",{x:0.232,y:1.12,w:2.274,h:0.384,fontSize:16,bold:true,color:C.navy,fontFace:"Calibri"});
  s.addText("Can LLMs see your content?",{x:0.232,y:1.493,w:2.274,h:0.299,fontSize:10,italic:true,color:C.purple,fontFace:"Calibri"});
  ["We audit if you're being cited across AI Search (ChatGPT, Perplexity, SGE)","We identify which competitors are winning those spots and why","We check if your URLs are indexable, link-worthy, and retrievable"].forEach((b,i)=>{
    s.addShape(pres.shapes.OVAL,{x:0.232,y:1.877+i*0.46,w:0.083,h:0.107,fill:{color:C.navy},line:{color:C.navy}});
    s.addText(b,{x:0.364,y:1.835+i*0.46,w:2.134,h:0.469,fontSize:7.5,color:C.darkgray,fontFace:"Calibri",wrap:true});
  });
  s.addText("Citability",{x:5.706,y:1.067,w:2.357,h:0.384,fontSize:16,bold:true,color:C.navy,fontFace:"Calibri"});
  s.addText("Can LLMs trust your content?",{x:5.706,y:1.429,w:2.357,h:0.299,fontSize:10,italic:true,color:C.purple,fontFace:"Calibri"});
  ["We rewrite content to include expert quotes, references, structured data","We improve source credibility through media presence, high-authority citations, and entity recognition","We restructure pages to be chunkable and retrievable"].forEach((b,i)=>{
    s.addShape(pres.shapes.OVAL,{x:5.706,y:1.813+i*0.42,w:0.083,h:0.107,fill:{color:C.navy},line:{color:C.navy}});
    s.addText(b,{x:5.839,y:1.771+i*0.42,w:2.192,h:0.427,fontSize:7.5,color:C.darkgray,fontFace:"Calibri",wrap:true});
  });
  s.addText("Retrievability:",{x:5.706,y:3.52,w:2.276,h:0.363,fontSize:16,bold:true,color:C.navy,fontFace:"Calibri"});
  s.addText("Can LLMs use your content to answer future questions?",{x:5.706,y:3.861,w:2.357,h:0.469,fontSize:9,italic:true,color:C.purple,fontFace:"Calibri",wrap:true});
  ["We chunk and tag your content to feed RAG systems better","We add LLM-readable markup and context layering (FAQs, comparisons, summaries)","We monitor which prompts lead to brand visibility and close the loop"].forEach((b,i)=>{
    s.addShape(pres.shapes.OVAL,{x:5.706,y:4.373+i*0.38,w:0.083,h:0.107,fill:{color:C.navy},line:{color:C.navy}});
    s.addText(b,{x:5.839,y:4.331+i*0.38,w:2.192,h:0.384,fontSize:7.5,color:C.darkgray,fontFace:"Calibri",wrap:true});
  });
}

function buildSlide13(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  staticHdr(s,pres,"Reverse Engineering How LLMs Index Content",d.brandName);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.248,y:0.747,w:7.774,h:0.704,fill:{color:"EEF0FF"},line:{color:C.lightgray},rectRadius:0.066});
  s.addText("LLM Retrieval Score",{x:0.414,y:0.811,w:2.398,h:0.555,fontSize:13,bold:true,italic:true,color:C.navy,fontFace:"Calibri"});
  s.addText("\u221D",{x:2.812,y:0.811,w:0.414,h:0.555,fontSize:18,color:C.navy,align:"center",fontFace:"Calibri"});
  s.addText("(Chunking \u00D7 Structure \u00D7 Schema \u00D7 Source Weight \u00D7 Trust Signals)",{x:3.225,y:0.811,w:4.631,h:0.555,fontSize:12,bold:true,italic:true,color:C.navy,fontFace:"Calibri",wrap:true});
  const hX13=[0.3,1.72,5.74],hW13=[1.4,4.0,3.82];
  s.addShape(pres.shapes.RECTANGLE,{x:0.248,y:1.6,w:7.774,h:0.341,fill:{color:C.navy},line:{color:C.navy}});
  [["Variable","What It Means","How Pepper Optimizes It"]].flat().forEach((h,i)=>s.addText(h,{x:hX13[i]+0.06,y:1.632,w:hW13[i]-0.1,h:0.277,fontSize:8.5,bold:true,italic:true,color:C.white,fontFace:"Calibri"}));
  [["Chunking","Atomic 2-4 sentence blocks ideal for embedding + summarization","We rewrite long-form into discrete semantic units"],
   ["Structure","Use of TL;DRs, bullets, lists, Q&A formatting","Content is formatted with high semantic clarity"],
   ["Schema","Machine-readable metadata (FAQPage, HowTo, Product)","Implemented across product pages, glossaries, and help docs"],
   ["Source Weight","LLM preference hierarchy (Wikipedia > PDF > Help Docs > Blogs > Social)","Content is distributed into high-weight surfaces LLMs trust"],
   ["Trust Signals","Presence of citations, statistics, interlinking, and cross-source agreement","We embed outbound and inbound credibility into every content artifact"],
  ].forEach((row,ri)=>{
    const y=1.84+ri*0.58,bg=ri%2===0?C.white:C.offwhite;
    s.addShape(pres.shapes.RECTANGLE,{x:0.248,y,w:7.774,h:0.597,fill:{color:bg},line:{color:C.lightgray}});
    [1.72,5.74].forEach(dx=>s.addShape(pres.shapes.RECTANGLE,{x:dx,y,w:0.012,h:0.597,fill:{color:C.lightgray},line:{color:C.lightgray}}));
    row.forEach((cell,ci)=>s.addText(cell,{x:hX13[ci]+0.06,y:y+0.08,w:hW13[ci]-0.12,h:0.448,fontSize:7.5,italic:true,color:C.darkgray,align:"center",fontFace:"Calibri",wrap:true}));
  });
}

function buildSlide14(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  staticHdr(s,pres,"The Content Strategy : Source Weightages by LLMs",d.brandName);
  const plats=[{name:"ChatGPT",color:"F0EEFF",hc:C.purple},{name:"Gemini",color:"EAF4FF",hc:"4285F4"},{name:"perplexity",color:"F5F5F5",hc:"6C6C6C"},{name:"Claude",color:"FFF3EE",hc:C.orange}];
  const colW14=1.9,sx14=2.1,hy14=0.72;
  s.addShape(pres.shapes.RECTANGLE,{x:0.232,y:hy14,w:1.472,h:0.448,fill:{color:"D8F0F0"},line:{color:C.teal}});
  s.addText("FACTORS",{x:0.232,y:hy14,w:1.472,h:0.448,fontSize:9,bold:true,color:C.darkgray,align:"center",valign:"middle",fontFace:"Calibri"});
  plats.forEach((p,i)=>{
    const x=sx14+i*colW14;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:x+0.04,y:hy14-0.04,w:colW14-0.08,h:0.533,fill:{color:p.color},line:{color:p.hc},rectRadius:0.099});
    s.addText(p.name,{x:x+0.04,y:hy14-0.04,w:colW14-0.08,h:0.533,fontSize:10,bold:true,color:p.hc,align:"center",valign:"middle",fontFace:"Calibri"});
  });
  [["Authoritative List Mentions (e.g. Top CRMs on G2, Best VPNs on TechRadar)",["\u2713 (41%)","\u2713 (49%)","\u2713 (64%)","\u2717"]],
   ["Awards / Accreditations (e.g. Gartner MQ, Inc. 5000 Badge)",["\u2713 (18%)","\u2713 (15%)","\u2713 (5%)","\u2713 (19%)"]],
   ["Online Reviews (e.g. Google Reviews, TrustPilot, Capterra ratings)",["\u2713 (16%)","\u2713 (13%)","\u2713 (31%)","\u2717"]],
   ["Customer Examples / Usage Data (e.g. Used by IBM, Powered by Salesforce)",["\u2713 (14%)","\u2717","\u2717","\u2713 (13%)"]],
   ["Social Sentiment (e.g. Reddit threads, Quora/Twitter buzz)",["\u2713 (11%)","\u2717","\u2717","\u2717"]],
   ["Local Reviews (e.g. Google Business Profile, Yelp)",["\u2717","\u2713 (Local)","\u2713 (Local)","\u2717"]],
   ["Traditional Directories (e.g. NY Times, Bloomberg, Hoovers)",["\u2717","\u2717","\u2717","\u2713 (68%)"]],
  ].forEach(([factor,vals],ri)=>{
    const y=hy14+0.46+ri*0.46,bg=ri%2===0?C.white:C.offwhite;
    s.addShape(pres.shapes.RECTANGLE,{x:0.232,y,w:7.807,h:0.469,fill:{color:bg},line:{color:C.lightgray}});
    s.addText(factor,{x:0.281,y:y+0.04,w:1.406,h:0.405,fontSize:6.5,color:C.darkgray,fontFace:"Calibri",wrap:true,align:"center"});
    vals.forEach((v,vi)=>{
      const x=sx14+vi*colW14;
      s.addText(v,{x:x+0.04,y:y+0.1,w:colW14-0.08,h:0.277,fontSize:9,bold:true,color:v==="\u2717"?"E03030":"1A7A40",align:"center",fontFace:"Calibri"});
    });
  });
}

function buildSlide15(pres, d) {
  const s=pres.addSlide();
  s.background={color:C.white};
  staticHdr(s,pres,"The Content Strategy : Topical Authority",d.brandName);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.248,y:0.725,w:7.774,h:0.661,fill:{color:"EEF0FF"},line:{color:C.lightgray},rectRadius:0.066});
  s.addText("LLM Recommendations",{x:0.414,y:0.789,w:2.646,h:0.512,fontSize:12,bold:true,italic:true,color:C.navy,fontFace:"Calibri"});
  s.addText("\u221D",{x:3.06,y:0.789,w:0.414,h:0.512,fontSize:16,color:C.navy,align:"center",fontFace:"Calibri"});
  s.addText("RRF Score = \u03A3 [1 / (60 + SERP position)]",{x:3.473,y:0.789,w:4.383,h:0.512,fontSize:12,bold:true,italic:true,color:C.navy,fontFace:"Calibri",wrap:true});

  s.addText("Illustration",{x:0.248,y:1.493,w:1.654,h:0.299,fontSize:10,bold:true,color:C.navy,fontFace:"Calibri"});

  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.248,y:1.835,w:0.662,h:0.299,fill:{color:C.navy},line:{color:C.navy},rectRadius:0.041});
  s.addText("Brand A",{x:0.248,y:1.835,w:0.662,h:0.299,fontSize:7.5,bold:true,color:C.white,align:"center",valign:"middle",fontFace:"Calibri"});

  const hX15=[0.3,2.05,2.7],hW15=[1.72,0.62,0.9];
  s.addShape(pres.shapes.RECTANGLE,{x:0.248,y:2.176,w:2.713,h:0.32,fill:{color:C.navy},line:{color:C.navy}});
  [["Query","Rank","RRF Score"]].flat().forEach((h,i)=>s.addText(h,{x:hX15[i]+0.04,y:2.208,w:hW15[i]-0.06,h:0.256,fontSize:7.5,bold:true,color:C.white,fontFace:"Calibri",align:"center"}));

  s.addShape(pres.shapes.RECTANGLE,{x:0.248,y:2.517,w:2.713,h:0.32,fill:{color:C.offwhite},line:{color:C.lightgray}});
  s.addText("Best Credit Card for Travellers",{x:0.298,y:2.549,w:1.373,h:0.256,fontSize:7.5,color:C.darkgray,fontFace:"Calibri"});
  s.addText("#1",{x:1.728,y:2.549,w:0.447,h:0.256,fontSize:7.5,color:C.darkgray,align:"center",fontFace:"Calibri"});
  s.addText("0.0164",{x:2.266,y:2.549,w:0.662,h:0.256,fontSize:7.5,color:C.darkgray,align:"center",fontFace:"Calibri"});

  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:3.143,y:1.835,w:0.662,h:0.299,fill:{color:C.navy},line:{color:C.navy},rectRadius:0.041});
  s.addText("Brand B",{x:3.143,y:1.835,w:0.662,h:0.299,fontSize:7.5,bold:true,color:C.white,align:"center",valign:"middle",fontFace:"Calibri"});

  const bTableY=2.04;
  const hX15b=[3.8,5.6,6.3],hW15b=[1.78,0.68,0.88];
  s.addShape(pres.shapes.RECTANGLE,{x:3.143,y:bTableY,w:2.795,h:0.32,fill:{color:C.navy},line:{color:C.navy}});
  [["Query","Rank","RRF Score"]].flat().forEach((h,i)=>s.addText(h,{x:hX15b[i]+0.04,y:bTableY+0.03,w:hW15b[i]-0.06,h:0.256,fontSize:7.5,bold:true,color:C.white,fontFace:"Calibri",align:"center"}));

  [["Best Credit Card for Travellers","#4","0.0156"],
   ["Travel On A Budget","#5","0.0154"],
   ["USe Reward Points for Flights","#6","0.0152"],
   ["Cheapest Hotel Tricks on Your Credit Card","#4","0.0156"],
   ["Eligibility for Credit Card","#7","0.0149"],
   ["Total","","0.0767"]].forEach((r,ri)=>{
    const y=bTableY+0.32+ri*0.28,bg=ri%2===0?C.offwhite:C.white;
    s.addShape(pres.shapes.RECTANGLE,{x:3.143,y,w:2.795,h:0.299,fill:{color:bg},line:{color:C.lightgray}});
    s.addText(r[0],{x:3.192,y:y+0.04,w:1.406,h:0.235,fontSize:6.5,color:C.darkgray,fontFace:"Calibri",wrap:true});
    s.addText(r[1],{x:4.664,y:y+0.04,w:0.496,h:0.235,fontSize:6.5,color:C.darkgray,align:"center",fontFace:"Calibri"});
    s.addText(r[2],{x:5.243,y:y+0.04,w:0.662,h:0.235,fontSize:6.5,bold:ri===5,color:ri===5?C.navy:C.darkgray,align:"center",fontFace:"Calibri"});
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.248,y:4.821,w:5.69,h:0.555,fill:{color:C.yellow},line:{color:C.yellow},rectRadius:0.066});
  s.addText("For AI visibility, Brand B >>>> Brand A",{x:0.248,y:4.821,w:5.69,h:0.555,fontSize:13,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});

  s.addShape(pres.shapes.OVAL,{x:6.244,y:2.176,w:0.099,h:0.128,fill:{color:C.navy},line:{color:C.navy}});
  s.addText([
    {text:"Building Topical Authority ",options:{fontSize:9,bold:true,color:C.navy,fontFace:"Calibri"}},
    {text:"still tops the content priority, with the relevant technical optimisation for each URL",options:{fontSize:8,color:C.darkgray,fontFace:"Calibri"}}
  ],{x:6.384,y:2.112,w:1.654,h:0.768,wrap:true});

  s.addShape(pres.shapes.OVAL,{x:6.244,y:3.072,w:0.099,h:0.128,fill:{color:C.navy},line:{color:C.navy}});
  s.addText([
    {text:"The key reason of smaller publishers/brands doing well on LLM queries is their ",options:{fontSize:8,color:C.darkgray,fontFace:"Calibri"}},
    {text:"trust-signalling coverage",options:{fontSize:9,bold:true,color:"CC0000",fontFace:"Calibri"}}
  ],{x:6.384,y:3.008,w:1.654,h:0.832,wrap:true});
}
function buildSlide16(pres, d) {
  const s=pres.addSlide();
  s.background={color:C.white};
  staticHdr(s,pres,"Strategy to Dominate Generative Search (GEO)",d.brandName);

  const steps=[
    {n:"1",title:"Curate Prompt List",body:"Curate a list of relevant prompts"},
    {n:"2",title:"Multi LLM Analysis",body:"Analyze the responses from different LLMs"},
    {n:"3",title:"Citation & Brand Mention Audit",body:"Review of cited URLs & brand mentions"},
    {n:"4",title:"Page Creation & Optimization",body:"Identifying pages to be created & optimized"},
    {n:"5",title:"Community Visibility",body:"Access your presence of reddit, Quora forums"}
  ];
  const bW=1.58,bH=1.52,bY=1.32,gap=0.18,sx16=0.28;

  steps.forEach((step,i)=>{
    const x=sx16+i*(bW+gap),isH=i===1;
    s.addShape(pres.shapes.OVAL,{x:x+bW/2-0.18,y:bY-0.44,w:0.298,h:0.384,fill:{color:C.navy},line:{color:C.navy}});
    s.addText(step.n,{x:x+bW/2-0.18,y:bY-0.44,w:0.298,h:0.384,fontSize:12,bold:true,color:C.white,align:"center",valign:"middle",fontFace:"Calibri"});
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:bY,w:bW,h:bH,fill:{color:C.white},line:{color:isH?C.navy:C.teal,pt:isH?2:1.5},rectRadius:0.099});
    s.addText(step.n+".",{x:x+0.1,y:bY+0.1,w:bW-0.2,h:0.277,fontSize:9.5,bold:true,color:C.navy,align:"center",fontFace:"Calibri"});
    s.addText(step.title,{x:x+0.08,y:bY+0.32,w:bW-0.16,h:0.448,fontSize:9,bold:true,color:C.navy,align:"center",fontFace:"Calibri",wrap:true});
    s.addText(step.body,{x:x+0.08,y:bY+0.78,w:bW-0.16,h:0.704,fontSize:7.5,italic:true,color:C.slate,align:"center",fontFace:"Calibri",wrap:true});
    if(i<4)s.addShape(pres.shapes.RECTANGLE,{x:x+bW,y:bY+bH/2-0.04,w:gap,h:0.085,fill:{color:C.navy},line:{color:C.navy}});
  });

  const boxY=3.1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.232,y:boxY,w:7.807,h:0.832,fill:{color:C.yellow},line:{color:C.yellow},rectRadius:0.083});
  s.addShape(pres.shapes.OVAL,{x:0.397,y:boxY+0.21,w:0.298,h:0.384,fill:{color:C.navy},line:{color:C.navy}});
  s.addText("6",{x:0.397,y:boxY+0.21,w:0.298,h:0.384,fontSize:12,bold:true,color:C.white,align:"center",valign:"middle",fontFace:"Calibri"});
  s.addText("Implement & Iterate",{x:0.81,y:boxY+0.21,w:2.068,h:0.384,fontSize:10,bold:true,color:C.navy,valign:"middle",fontFace:"Calibri"});
  s.addText("Create new pages, update existing pages, implement schemas, and community replies and re-run the prompt set monthly to gauge lift and uncover new topical gaps.",{x:2.398,y:boxY+0.1,w:5.541,h:0.597,fontSize:9,bold:true,color:C.navy,fontFace:"Calibri",wrap:true});
}
function buildSlide17(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  staticHdr(s,pres,"Here, is the customised content strategy table for you!",d.brandName);
  const hX17=[0.25,1.48,2.52,7.1,8.12],hW17=[1.2,1.02,4.55,0.98,1.6];
  s.addShape(pres.shapes.RECTANGLE,{x:0.207,y:0.693,w:7.857,h:0.363,fill:{color:C.purple},line:{color:C.purple}});
  [["Source Type",0],["Estimated Weight",1],["Notes",2],["% Weightage",3],["Relevant Examples",4]].forEach(([h,i])=>s.addText(h,{x:hX17[i]+0.04,y:0.725,w:hW17[i]-0.06,h:0.299,fontSize:7,bold:true,color:C.white,fontFace:"Calibri",align:"center",wrap:true}));
  [["Product & Platform Pages","Very High","Core retrieval: prompts like platform AI, ITSM","25%","ITSM, ITOM, CSM, HRSD, Now Assist (GenAI)"],
   ["Industry Solutions Pages","High","For prompts like workflow automation for banking, government digital services","14%","Financial services, healthcare, manufacturing, government, telecom"],
   ["AI & Technology Innovation","High","For prompts like Now Assist generative AI, platform intelligence","13%","GenAI copilots, predictive intelligence, Vancouver releases"],
   ["Customer Stories / Case Studies","High","Retrieval for ServiceNow success stories, ServiceNow ROI","12%","Case studies with Citi, DHL, Novartis, government agencies"],
   ["Pricing / Demo Pages","Medium-High","For prompts like ServiceNow demo, ServiceNow pricing","10%","Request a demo, Now Assist ROI tools"],
   ["Documentation & Knowledge Base","Medium","Retrieval for prompts like ServiceNow API, developer docs","8%","Developer portal, integration hub docs, API references"],
   ["Events & Webinars","Medium","For prompts like ServiceNow Knowledge conference","6%","Knowledge conference sessions, leadership keynotes"],
   ["Press & Newsroom","Medium-Low","For prompts like ServiceNow earnings, acquisitions","6%","Acquisitions, quarterly earnings"],
   ["Support / Help Center","Low-Medium","For prompts like ServiceNow login, support portal","4%","Support portal, instance upgrades, troubleshooting"],
   ["Careers & Corporate Pages","Low","For prompts like jobs at ServiceNow, company culture","2%","Careers site, employee stories, ESG reports"],
  ].forEach((row,ri)=>{
    const y=1.01+ri*0.38,bg=ri%2===0?C.offwhite:C.white;
    s.addShape(pres.shapes.RECTANGLE,{x:0.207,y,w:7.857,h:0.384,fill:{color:bg},line:{color:C.lightgray}});
    [1.48,2.52,7.1,8.12].forEach(dx=>s.addShape(pres.shapes.RECTANGLE,{x:dx,y,w:0.012,h:0.384,fill:{color:C.lightgray},line:{color:C.lightgray}}));
    row.forEach((cell,ci)=>s.addText(cell,{x:hX17[ci]+0.04,y:y+0.06,w:hW17[ci]-0.08,h:0.277,fontSize:6.5,color:C.darkgray,fontFace:"Calibri",align:"center",wrap:true}));
  });
  s.addText("Note: The weightage percentage is a relative effort guide. If you're putting X effort on a source with 2% weight, then a source with 8% weight deserves 4X effort. Prioritize accordingly.",{x:0.207,y:5.173,w:7.857,h:0.341,fontSize:7,italic:true,color:C.slate,fontFace:"Calibri",wrap:true});
}

function buildSlide18(pres, d) {
  const s=pres.addSlide(); s.background={color:C.white};
  logoPill(s, pres);
  s.addText("Do you think this was helpful and want to dive deeper?",{x:0.289,y:0.587,w:7.691,h:1.6,fontSize:32,bold:true,color:C.navy,fontFace:"Calibri",wrap:true});
  s.addText("Reach us at +14157545133",{x:0.289,y:2.453,w:4.962,h:0.405,fontSize:14,color:C.darkgray,fontFace:"Calibri"});
  s.addText("or write to us at kishan@peppercontent.io",{x:0.289,y:2.859,w:4.962,h:0.405,fontSize:14,color:C.darkgray,fontFace:"Calibri"});
  s.addShape(pres.shapes.RECTANGLE,{x:0,y:3.861,w:8.27,h:0.064,fill:{color:C.yellow},line:{color:C.yellow}});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.289,y:4.075,w:7.691,h:1.024,fill:{color:"FFF3C4"},line:{color:C.yellow},rectRadius:0.066});
  s.addText("We are hosting some free GEO workshops for corporates and if you think you'd be interested in the same, please write to us",{x:0.455,y:4.181,w:7.443,h:0.811,fontSize:12,color:C.purple,align:"center",valign:"middle",fontFace:"Calibri",wrap:true});
}

// ─── STEP 3: Build PPTX ─────────────────────────────────────────────────────
function buildPPTX(data, outputPath) {
  const pres=new pptxgen();
  pres.defineLayout({ name: "PRINT_8x6", width: 8.27, height: 6.0 });
  pres.layout = "PRINT_8x6";
  pres.title=data.brandName+" GEO Audit \u2014 Pepper";
  pres.author="Pepper.inc";
  buildSlide1(pres,data);
  buildSlide2(pres,data);
  buildSlide3(pres,data);
  buildSlide4(pres,data);
  buildSlide5(pres,data);
  buildSlide6(pres,data);
  buildSlide7(pres,data);
  buildSlide8(pres,data);
  buildSlide12(pres,data);
  buildSlide13(pres,data);
  buildSlide14(pres,data);
  buildSlide15(pres,data);
  buildSlide16(pres,data);
  buildSlide17(pres,data);
  buildSlide18(pres,data);
  pres.writeFile({ fileName:outputPath });
  console.log("\u2705 PPTX written:", outputPath, "(15 slides)");
}

// ─── Helper: generate a single report file (used by both endpoints) ─────────
async function generateSingleReport(reportId, format) {
  const apiData = await fetchReportData(reportId);
  const data = normalizeData(apiData);
  const id = uuidv4();
  const pptxOut = path.join(TMP, id + ".pptx");

  console.log("\n\u{1F5BC} Building PPTX for:", data.brandName);
  buildPPTX(data, pptxOut);
  await new Promise(r => setTimeout(r, 1000));

  if (format === "pptx") {
    return { filePath: pptxOut, brandName: data.brandName, ext: "pptx" };
  }

  const pdfOut = path.join(TMP, id + ".pdf");
  console.log("\u{1F4C4} Converting to PDF...");
  { let pdfOk = false; for (let attempt = 1; attempt <= 3; attempt++) { try { try { execSync("killall -9 soffice.bin 2>/dev/null || true", { timeout: 5000 }); } catch {} execSync(`soffice --headless --norestore --convert-to pdf --outdir ${TMP} ${pptxOut}`, { timeout: 120000 }); pdfOk = true; break; } catch (e) { console.log(`  ⚠️ soffice attempt ${attempt}/3 failed: ${e.message}`); if (attempt < 3) await new Promise(r => setTimeout(r, 2000)); } } if (!pdfOk) throw new Error("PDF conversion failed after 3 attempts — soffice crashed"); }
  try { fs.unlinkSync(pptxOut); } catch {}
  if (!fs.existsSync(pdfOut)) throw new Error("PDF conversion failed");
  return { filePath: pdfOut, brandName: data.brandName, ext: "pdf" };
}

// ─── API: POST /generate (single report) ──────────────────────────────────
app.post("/generate", async (req, res) => {
  let reportId = req.body.reportId;
  if (!reportId && req.body.url) {
    reportId = extractReportId(req.body.url);
  }
  if (!reportId) {
    return res.status(400).json({ error: "Please provide a reportId or a valid Pepper report URL." });
  }

  const format = (req.body.format || "pdf").toLowerCase();
  if (!["pdf", "pptx"].includes(format)) {
    return res.status(400).json({ error: "Invalid format. Use 'pdf' or 'pptx'." });
  }

  try {
    const result = await generateSingleReport(reportId, format);
    const fileName = `${result.brandName} x Pepper - GEO report.${result.ext}`;
    res.download(result.filePath, fileName, (err) => {
      try { fs.unlinkSync(result.filePath); } catch {}
      if (err && !res.headersSent) res.status(500).json({ error: "Download failed." });
    });
  } catch (err) {
    console.error("\u274C Error:", err.message);
    res.status(500).json({ error: err.message || "Generation failed." });
  }
});

// ─── API: POST /generate-bulk (multiple reports from Excel/CSV) ───────────────
app.post("/generate-bulk", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "Please upload an Excel (.xlsx) or CSV (.csv) file." });
  }

  const format = (req.body.format || "pdf").toLowerCase();
  if (!["pdf", "pptx"].includes(format)) {
    return res.status(400).json({ error: "Invalid format. Use 'pdf' or 'pptx'." });
  }

  let reportIds = [];
  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Extract all non-empty string cells, try to parse report IDs
    const seen = new Set();
    for (const row of rows) {
      for (const cell of row) {
        if (typeof cell === "string" && cell.trim()) {
          const id = extractReportId(cell.trim());
          if (id && !seen.has(id)) {
            seen.add(id);
            reportIds.push({ raw: cell.trim(), id });
          }
        }
      }
    }
  } catch (parseErr) {
    return res.status(400).json({ error: "Could not parse the uploaded file. Please upload a valid .xlsx or .csv file." });
  } finally {
    try { fs.unlinkSync(req.file.path); } catch {}
  }

  if (reportIds.length === 0) {
    return res.status(400).json({ error: "No valid Pepper report URLs or IDs found in the uploaded file." });
  }

  if (reportIds.length > MAX_BULK_REPORTS) {
    return res.status(400).json({ error: `Too many reports. Maximum is ${MAX_BULK_REPORTS}, but found ${reportIds.length}.` });
  }

  console.log(`\n\u{1F4E6} Bulk generation: ${reportIds.length} reports, format=${format}`);

  const zip = new JSZip();
  const results = [];

  for (let i = 0; i < reportIds.length; i++) {
    const { raw, id } = reportIds[i];
    console.log(`  [${i + 1}/${reportIds.length}] Processing: ${id}`);
    try {
      const result = await generateSingleReport(id, format);
      const fileName = `${result.brandName} x Pepper - GEO report.${result.ext}`;
      zip.file(fileName, fs.readFileSync(result.filePath));
      try { fs.unlinkSync(result.filePath); } catch {}
      results.push({ index: i + 1, reportId: id, brand: result.brandName, status: "success" });
      console.log(`  \u2705 [${i + 1}] ${result.brandName} — done`);
    } catch (err) {
      const isNotFound = err.statusCode === 404 || err.message.includes("404"); results.push({ index: i + 1, reportId: id, status: "failed", error: err.message, permanent: isNotFound }); if (isNotFound) console.log(`  ⛔ [${i + 1}] ${id} — permanently failed (404 Not Found, report does not exist)`); else
      console.log(`  \u274C [${i + 1}] ${id} — failed: ${err.message}`);
    }
  }

  const succeeded = results.filter(r => r.status === "success").length;
  const failed = results.filter(r => r.status === "failed").length;

  if (succeeded === 0) {
    return res.status(500).json({
      error: "All reports failed to generate.",
      results,
    });
  }

  // Add a summary text file inside the zip
  const summary = results.map(r => {
    if (r.status === "success") return `\u2705 [${r.index}] ${r.brand} (${r.reportId})`;
    return `\u274C [${r.index}] ${r.reportId} — ${r.error}`;
  }).join("\n");
  zip.file("_bulk-summary.txt", `Pepper GEO Bulk Report Summary\n${"=".repeat(40)}\nTotal: ${reportIds.length} | Success: ${succeeded} | Failed: ${failed}\n\n${summary}\n`);

  const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

  res.set("Content-Type", "application/zip");
  res.set("Content-Disposition", `attachment; filename="Pepper-GEO-Reports-Bulk.${format}.zip"`);
  res.send(zipBuffer);
  console.log(`\u{1F4E6} Bulk ZIP sent: ${succeeded} succeeded, ${failed} failed`);
});

// ─── API: GET /bulk-progress (SSE for real-time updates — future use) ──────
// Placeholder for Server-Sent Events based progress tracking


// ─── API: GET /templates/bulk-template.xlsx (dynamically generated) ─────────
app.get("/templates/bulk-template.xlsx", (req, res) => {
  const wb = XLSX.utils.book_new();
  const data = [
    ["Pepper Report URL"],
    ["https://atlas.pepper.inc/public/reports/YOUR-REPORT-ID-HERE/overview"],
    ["https://atlas.pepper.inc/public/reports/ANOTHER-REPORT-ID/overview"],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{ wch: 70 }];
  XLSX.utils.book_append_sheet(wb, ws, "Reports");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.set("Content-Disposition", 'attachment; filename="bulk-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ─── API: GET /templates/bulk-template.csv (dynamically generated) ─────────
app.get("/templates/bulk-template.csv", (req, res) => {
  const csv = "Atlas Report URL\nhttps://atlas.pepper.inc/public/reports/YOUR-REPORT-ID-HERE/overview\nhttps://atlas.pepper.inc/public/reports/ANOTHER-REPORT-ID/overview\n";
  res.set("Content-Type", "text/csv");
  res.set("Content-Disposition", 'attachment; filename="bulk-template.csv"');
  res.send(csv);
});
app.get("/health", (_, res) => res.json({ status: "ok" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`\u{1F680} Atlas PDF Generator on port ${PORT}`));
