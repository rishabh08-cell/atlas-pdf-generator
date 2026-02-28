const express = require("express");
const pptxgen = require("pptxgenjs");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const JSZip = require("jszip");
const TEMPLATE_PATH = path.join(__dirname, "public", "templates", "Template Slides.pptx");

const app = express();
app.use(express.json());
app.use(express.static("public"));

const TMP = path.join(__dirname, "tmp");
if (!fs.existsSync(TMP)) fs.mkdirSync(TMP);

// ─── Pepper brand colour palette ──────────────────────────────────────────────
const C = {
    navy: "0D007D",
    purple: "3D35B5",
    violet: "7B7FD4",
    lilac: "A8ABEA",
    orange: "DA5012",
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
    // Accept a raw report ID or a full Atlas URL
  // URL format: https://atlas.pepper.inc/reports/{reportId}/overview
  // or hub.peppercontent.io/atlas-service/api/public/reports/{reportId}/...
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
        console.log(`  ⬇️  ${key}: ${url}`);
        const res = await fetch(url);
        if (!res.ok) throw new Error(`API error for ${ep}: ${res.status} ${res.statusText}`);
        results[key] = await res.json();
        console.log(`  ✅ ${key} fetched`);
  }

  return results;
}

// ─── STEP 2: Normalize API data into slide-ready shape ──────────────────────
function normalizeData(api) {
    const { overview, competitors, platforms, prompts } = api;
    const brand = overview.brand;

  // ── Summary stats from /platforms ──
  const totalMentions = parseInt(platforms.stats.find(s => s.label === "Total Brand Mentions")?.value) || 0;
    const totalCitations = parseInt(platforms.stats.find(s => s.label === "Total Domain Citations")?.value) || 0;
    const avgBrandCoverage = platforms.stats.find(s => s.label === "Avg Brand Coverage")?.value || "0%";
    const avgDomainCoverage = platforms.stats.find(s => s.label === "Avg Domain Coverage")?.value || "0%";

  // ── Leaderboard & competitor mentions from /overview ──
  const sortedCompetitors = [...overview.competitors].sort((a, b) => a.rank - b.rank);
    const userCompetitor = overview.competitors.find(c => c.is_user);
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

  // ── Platforms table from /platforms ──
  const platformRows = (platforms.rows || []).map(r => ({
        name: r.platform,
        mentions: r.mentions,
        citations: r.citations,
        brandVisibility: r.brand_visibility,
        domainCoverage: r.domain_coverage,
  }));

  // ── Prompt themes from /prompts-themes ──
  const promptThemes = (prompts.themes || []).map(t => ({
        theme: t.name,
        promptCount: t.prompts.length,
        prompts: t.prompts.map(p => p.text),
  }));

  // ── Domain citations from /overview ──
  const domainCitations = (overview.domains || []).slice(0, 10).map(d => ({
        domain: d.domain,
        uniquePagesCited: d.total_pages_cited,
      domainCoverage: Math.round(d.mention_rate * 10) / 10 + "%",
        domainShare: Math.round(d.share_of_voice * 10) / 10 + "%",
  }));

  // ── Brand pages from /overview ──
  const brandPages = (overview.pages || []).slice(0, 8).map(p => ({
        name: p.title || p.url,
        url: p.url,
        prompts: p.prompts,
  }));

  // ── Competitor visibility matrix from /competitors-comparison ──
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

  // ── Brand visibility by platform from /platforms ──
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

  // ── Clean domain string ──
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
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:9.0,y:0.12,w:0.85,h:0.26, fill:{color:C.white}, line:{color:C.lightgray,pt:1}, rectRadius:0.06, shadow:makeShadow() });
    if (fs.existsSync(pepperLogoPath)) {
          s.addImage({ path:pepperLogoPath, x:9.03,y:0.14,w:0.75,h:0.20, sizing:{type:"contain",w:0.75,h:0.20} });
    } else {
          s.addText("pepper", { x:9.03,y:0.14,w:0.75,h:0.20, fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri" });
    }
}

function hdr(s, pres, title, brandName) {
    s.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:10,h:0.08, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText("atlas", { x:0.3,y:0.12,w:0.9,h:0.3, fontSize:11,bold:true,color:C.navy,fontFace:"Calibri" });
    s.addText("by pepper.inc", { x:1.2,y:0.17,w:1.3,h:0.22, fontSize:7,color:C.slate,fontFace:"Calibri" });
    logoPill(s, pres);
    s.addText(title, { x:0.3,y:0.5,w:7.2,h:0.48, fontSize:18,bold:true,color:C.navy,fontFace:"Calibri" });
    s.addShape(pres.shapes.RECTANGLE, { x:0.3,y:0.95,w:1.8,h:0.03, fill:{color:C.teal}, line:{color:C.teal} });
}

function ftr(s, pres, brand, domain) {
    s.addShape(pres.shapes.RECTANGLE, { x:0,y:5.42,w:10,h:0.2, fill:{color:C.navy}, line:{color:C.navy} });
    s.addText(brand+" · "+domain+" · GEO Audit by Atlas / Pepper.inc", { x:0.3,y:5.43,w:9.4,h:0.18, fontSize:6.5,color:"AAAACC",fontFace:"Calibri" });
}



// ─── SLIDE 1: Cover ────────────────────────────────────────────────────────
function buildSlide1(pres, d) {
    const s=pres.addSlide(); s.background={color:C.navy};
    s.addShape(pres.shapes.OVAL,{x:7.5,y:-0.5,w:3.5,h:3.5,fill:{color:"150050"},line:{color:"150050"}});
    s.addShape(pres.shapes.OVAL,{x:8.2,y:0.2,w:2.0,h:2.0,fill:{color:C.purple},line:{color:C.purple}});
    s.addText("atlas",{x:0.5,y:0.38,w:1.1,h:0.38,fontSize:15,bold:true,color:C.orange,fontFace:"Calibri"});
    s.addText("by pepper.inc",{x:1.63,y:0.44,w:1.5,h:0.26,fontSize:8,color:C.lilac,fontFace:"Calibri"});
    const pepperLogoPath = path.join(__dirname, "public", "logos", "pepper-logo.png");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:9.0,y:0.35,w:0.85,h:0.26,fill:{color:"FFFFFF"},line:{color:"FFFFFF"},rectRadius:0.08});
    if (fs.existsSync(pepperLogoPath)) {
          s.addImage({path:pepperLogoPath,x:9.03,y:0.37,w:0.75,h:0.20,sizing:{type:"contain",w:0.75,h:0.20}});
    } else {
          s.addText("pepper",{x:9.03,y:0.37,w:0.75,h:0.20,fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
    }
    s.addText("GEO AUDIT REPORT",{x:0.5,y:1.08,w:6,h:0.28,fontSize:9,color:C.orange,bold:true,charSpacing:4,fontFace:"Calibri"});
    s.addText(d.brandName,{x:0.5,y:1.38,w:7,h:1.18,fontSize:50,bold:true,color:C.white,fontFace:"Calibri"});
    s.addText(d.domain,{x:0.5,y:2.6,w:5,h:0.38,fontSize:13,color:C.lilac,fontFace:"Calibri"});
    s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:3.02,w:1.2,h:0.04,fill:{color:C.orange},line:{color:C.orange}});
    [{v:String(d.totalMentions),l:"Total Mentions"},{v:d.avgBrandCoverage,l:"Brand Coverage"},{v:String(d.platformCount),l:"AI Platforms"},{v:d.leaderboardRank,l:"Leaderboard"}].forEach((k,i)=>{
          const x=0.5+i*2.3;
          s.addText(k.v,{x,y:3.16,w:2.1,h:0.52,fontSize:24,bold:true,color:C.orange,fontFace:"Calibri"});
          s.addText(k.l,{x,y:3.66,w:2.1,h:0.22,fontSize:8,color:C.lilac,fontFace:"Calibri"});
    });
    s.addText("Powered by atlas · pepper.inc",{x:0.5,y:5.15,w:9,h:0.22,fontSize:7.5,color:C.slate,fontFace:"Calibri"});
}

// ─── SLIDE 2: Prompts & Themes ─────────────────────────────────────────────
function buildSlide2(pres, d) {
    const s=pres.addSlide(); s.background={color:C.white};
    const tp=d.promptThemes.reduce((a,t)=>a+(t.promptCount||t.prompts?.length||0),0);
    hdr(s,pres,`We Have Mapped ${tp} Prompts Across ${d.promptThemes.length} Themes`,d.brandName);
    ftr(s,pres,d.brandName,d.domain);
    d.promptThemes.slice(0,9).forEach((t,i)=>{
          const col=i%3,row=Math.floor(i/3),x=0.28+col*3.15,y=1.1+row*0.82;
          s.addShape(pres.shapes.RECTANGLE,{x,y,w:3.0,h:0.72,fill:{color:C.offwhite},line:{color:C.lightgray},shadow:makeShadow()});
          s.addShape(pres.shapes.RECTANGLE,{x,y,w:0.04,h:0.72,fill:{color:C.teal},line:{color:C.teal}});
          s.addText(t.theme,{x:x+0.1,y:y+0.08,w:2.85,h:0.34,fontSize:9.5,bold:true,color:C.navy,fontFace:"Calibri",wrap:true});
          s.addText((t.promptCount||t.prompts?.length||0)+" prompts",{x:x+0.1,y:y+0.46,w:2.85,h:0.2,fontSize:8.5,color:C.slate,fontFace:"Calibri"});
    });
}

// ─── SLIDE 3: Leaderboard + Competitors ───────────────────────────────────
function buildSlide3(pres, d) {
    const s=pres.addSlide(); s.background={color:C.white};
    hdr(s,pres,"Brand Leaderboard & Competitor Mentions",d.brandName);
    ftr(s,pres,d.brandName,d.domain);
    const brands=d.leaderboard;
    if (brands.length>0) {
          const maxM=Math.max(...brands.map(b=>b.mentions),1),barW=1.0,gap=0.55,startX=0.3,cb=4.65,ch=2.6;
          brands.forEach((brand,i)=>{
                  const x=startX+i*(barW+gap),barH=Math.max((brand.mentions/maxM)*ch,0.15),barY=Math.max(cb-barH,1.55),isB=brand.name===d.brandName;
                  const col=isB?C.orange:(i===0?C.teal:"BDBDCD");
                  s.addShape(pres.shapes.OVAL,{x:x+barW/2-0.22,y:barY-0.96,w:0.44,h:0.44,fill:{color:C.white},line:{color:C.lightgray}});
                  s.addText("#"+brand.rank,{x:x+barW/2-0.22,y:barY-0.96,w:0.44,h:0.44,fontSize:11,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
                  s.addText(brand.name,{x:x-0.1,y:barY-0.5,w:barW+0.2,h:0.34,fontSize:8,bold:isB,color:isB?C.orange:C.navy,align:"center",fontFace:"Calibri",wrap:true});
                  s.addShape(pres.shapes.RECTANGLE,{x,y:barY,w:barW,h:barH,fill:{color:col},line:{color:col},shadow:makeShadow()});
                  s.addText(brand.mentions+" mentions",{x:x-0.1,y:cb+0.06,w:barW+0.2,h:0.2,fontSize:7.5,color:C.slate,align:"center",fontFace:"Calibri"});
                  if (isB) s.addText("\u{1F451}",{x:x+barW/2-0.24,y:barY+0.06,w:0.48,h:0.35,fontSize:18,align:"center"});
          });
    }
    const comps=d.competitorMentions.slice(0,10),maxPct=Math.max(...comps.map(c=>c.percentage),1);
    s.addText("Competitor Mentions vs. "+d.brandName,{x:5.0,y:1.1,w:4.7,h:0.25,fontSize:9,bold:true,color:C.slate,fontFace:"Calibri"});
    comps.forEach((comp,i)=>{
          const y=1.42+i*0.33,isB=comp.name===d.brandName;
          s.addShape(pres.shapes.OVAL,{x:5.0,y:y+0.04,w:0.22,h:0.22,fill:{color:isB?C.purple:C.lightgray},line:{color:isB?C.purple:C.lightgray}});
          s.addText(comp.name[0].toUpperCase(),{x:5.0,y:y+0.04,w:0.22,h:0.22,fontSize:7,bold:true,color:isB?C.white:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
          s.addText(comp.name,{x:5.26,y:y+0.05,w:1.4,h:0.2,fontSize:8,bold:isB,color:isB?C.purple:C.navy,fontFace:"Calibri"});
          const bw=Math.min((comp.percentage/maxPct)*2.6,2.6);
          s.addShape(pres.shapes.RECTANGLE,{x:6.7,y:y+0.06,w:Math.max(bw,0.05),h:0.18,fill:{color:isB?C.navy:C.lightgray},line:{color:isB?C.navy:C.lightgray}});
          s.addText(comp.percentage+"% · "+comp.mentions+" mentions",{x:6.72+bw,y:y+0.05,w:Math.max(9.7-(6.72+bw),0.8),h:0.2,fontSize:7.5,color:C.slate,fontFace:"Calibri"});
    });
    s.addShape(pres.shapes.RECTANGLE,{x:4.75,y:1.05,w:0.03,h:3.8,fill:{color:C.lightgray},line:{color:C.lightgray}});
}

// ─── SLIDE 4: Top Cited Sources ────────────────────────────────────────────
function buildSlide4(pres, d) {
    const s=pres.addSlide(); s.background={color:C.white};
    hdr(s,pres,"Top Cited Sources (Category vs Us)",d.brandName);
    ftr(s,pres,d.brandName,d.domain);
    const domains=d.domainCitations.slice(0,9);
    s.addShape(pres.shapes.RECTANGLE,{x:0.25,y:1.08,w:4.4,h:0.28,fill:{color:C.navy},line:{color:C.navy}});
    [["Domain",0.35],["Pages",2.8],["Responses",3.65]].forEach(([h,x])=>s.addText(h,{x,y:1.1,w:1.1,h:0.24,fontSize:8,bold:true,color:C.white,fontFace:"Calibri"}));
    domains.forEach((row,i)=>{
          const y=1.38+i*0.34,bg=i%2===0?C.offwhite:C.white;
          s.addShape(pres.shapes.RECTANGLE,{x:0.25,y,w:4.4,h:0.32,fill:{color:bg},line:{color:C.lightgray}});
          s.addText(row.domain,{x:0.35,y:y+0.06,w:2.4,h:0.22,fontSize:8,color:C.navy,fontFace:"Calibri"});
          s.addText(String(row.uniquePagesCited||row.pages||""),{x:2.82,y:y+0.06,w:0.6,h:0.22,fontSize:8,color:C.slate,fontFace:"Calibri"});
          s.addText(String(row.domainShare||row.responses||""),{x:3.67,y:y+0.06,w:0.9,h:0.22,fontSize:8,color:C.slate,fontFace:"Calibri"});
    });
    const pages=d.brandPages.slice(0,6);
    s.addText("Sources from "+d.brandName+" Domain",{x:5.0,y:1.08,w:4.7,h:0.28,fontSize:9,bold:true,color:C.navy,fontFace:"Calibri"});
    s.addShape(pres.shapes.RECTANGLE,{x:5.0,y:1.36,w:4.7,h:0.02,fill:{color:C.lightgray},line:{color:C.lightgray}});
    pages.forEach((pg,i)=>{
          const y=1.42+i*0.5;
          s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:5.05,y,w:4.6,h:0.44,fill:{color:C.offwhite},line:{color:C.lightgray},rectRadius:0.04,shadow:makeShadow()});
          s.addText(pg.name,{x:5.15,y:y+0.04,w:3.6,h:0.2,fontSize:8.5,bold:true,color:C.navy,fontFace:"Calibri"});
          s.addText((pg.prompts||0)+" Response"+(pg.prompts!==1?"s":""),{x:9.1,y:y+0.04,w:0.5,h:0.2,fontSize:7.5,bold:true,color:C.teal,align:"right",fontFace:"Calibri"});
          s.addText(pg.url||d.domain,{x:5.15,y:y+0.24,w:4.4,h:0.16,fontSize:7,color:C.slate,fontFace:"Calibri"});
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:5.0,y:4.75,w:4.7,h:0.34,fill:{color:C.yellow},line:{color:C.yellow},rectRadius:0.05});
    s.addText("Top Cited Sources from our website \u2191",{x:5.0,y:4.75,w:4.7,h:0.34,fontSize:9,bold:true,color:C.navy,align:"center",valign:"middle",fontFace:"Calibri"});
    s.addShape(pres.shapes.RECTANGLE,{x:4.75,y:1.05,w:0.03,h:3.9,fill:{color:C.lightgray},line:{color:C.lightgray}});
}

// ─── SLIDE 5: Competitor Visibility Matrix ─────────────────────────────────
function buildSlide5(pres, d) {
    const s=pres.addSlide(); s.background={color:C.white};
    hdr(s,pres,"Theme Benchmarks (% Visibility across Competitors)",d.brandName);
    ftr(s,pres,d.brandName,d.domain);
    const matrix=d.competitorVisibilityMatrix;
    if (!matrix||matrix.length===0){s.addText("No competitor visibility matrix data available.",{x:0.5,y:2.5,w:9,h:0.5,fontSize:12,color:C.slate,align:"center",fontFace:"Calibri"});return;}
    const compNames=[];
    matrix.forEach(row=>{if(row.competitors)Object.keys(row.competitors).forEach(k=>{if(!compNames.includes(k))compNames.push(k);});});
    const allCols=[d.brandName,...compNames].slice(0,10),themeColW=1.9,dataColW=(9.5-themeColW)/allCols.length,startX=0.25,headerY=1.08,rowH=0.32;
    s.addShape(pres.shapes.RECTANGLE,{x:startX,y:headerY,w:9.5,h:rowH,fill:{color:C.navy},line:{color:C.navy}});
    s.addText("Topic",{x:startX+0.05,y:headerY+0.05,w:themeColW-0.08,h:rowH-0.08,fontSize:7,bold:true,color:C.white,fontFace:"Calibri"});
    allCols.forEach((col,ci)=>{
          const x=startX+themeColW+ci*dataColW,isB=col===d.brandName;
          s.addText(col,{x:x+0.02,y:headerY+0.04,w:dataColW-0.04,h:rowH-0.08,fontSize:6,bold:isB,color:isB?C.orange:C.white,align:"center",fontFace:"Calibri",wrap:true});
    });
    matrix.slice(0,11).forEach((row,ri)=>{
          const y=headerY+rowH+ri*rowH,bg=ri%2===0?C.offwhite:C.white;
          s.addShape(pres.shapes.RECTANGLE,{x:startX,y,w:9.5,h:rowH,fill:{color:bg},line:{color:C.lightgray}});
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
    s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.",{x:0.3,y:5.22,w:9.4,h:0.16,fontSize:7,italic:true,color:C.slate,align:"center",fontFace:"Calibri"});
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
     {term:"Brand Coverage",body:"Percentage of prompts that mention your brand"},
     {term:"Domain Coverage",body:"Percentage of prompts that cited your website"},
       ].forEach((def,i)=>{
             const col=i%3,row=Math.floor(i/3),x=0.28+col*3.22,y=1.2+row*1.62;
             s.addShape(pres.shapes.RECTANGLE,{x,y,w:3.06,h:1.52,fill:{color:C.yellow},line:{color:"D4AA30"},shadow:makeShadow()});
             s.addText(def.term,{x:x+0.14,y:y+0.14,w:2.78,h:0.3,fontSize:12,bold:true,color:C.purple,fontFace:"Calibri"});
             s.addText("\u2014 \u2014 \u2014 \u2014 \u2014 \u2014 \u2014 \u2014 \u2014 \u2014 \u2014 \u2014",{x:x+0.14,y:y+0.44,w:2.78,h:0.16,fontSize:7,color:C.purple,fontFace:"Calibri"});
             s.addText(def.body,{x:x+0.14,y:y+0.6,w:2.78,h:0.78,fontSize:9,color:C.navy,italic:true,bold:true,fontFace:"Calibri",wrap:true});
       });
    s.addText("Source: Otterly.ai",{x:0.3,y:5.22,w:3,h:0.16,fontSize:7.5,bold:true,color:C.navy,fontFace:"Calibri"});
}

// ─── SLIDE 7: Platform mentions table ──────────────────────────────────────
function buildSlide7(pres, d) {
    const s=pres.addSlide(); s.background={color:C.white};
    hdr(s,pres,d.brandName+" Mentions by AI Platform",d.brandName);
    ftr(s,pres,d.brandName,d.domain);
    [{v:String(d.totalMentions),l:"Total Brand Mentions"},{v:String(d.totalCitations),l:"Total Domain Citations"},{v:d.avgBrandCoverage,l:"Avg Brand Coverage"},{v:d.avgDomainCoverage,l:"Avg Domain Coverage"}].forEach((k,i)=>{
          const x=0.25+i*2.42;
          s.addShape(pres.shapes.RECTANGLE,{x,y:1.08,w:2.3,h:0.72,fill:{color:C.white},line:{color:C.lightgray},shadow:makeShadow()});
          s.addShape(pres.shapes.RECTANGLE,{x,y:1.08,w:2.3,h:0.06,fill:{color:C.teal},line:{color:C.teal}});
          s.addText(k.v,{x,y:1.16,w:2.3,h:0.38,fontSize:20,bold:true,color:C.navy,align:"center",fontFace:"Calibri"});
          s.addText(k.l,{x,y:1.52,w:2.3,h:0.24,fontSize:7.5,color:C.slate,align:"center",fontFace:"Calibri"});
    });
    const colX=[0.25,2.15,3.2,4.3,7.1],colW=[1.85,1.0,1.05,2.75,2.6];
    s.addShape(pres.shapes.RECTANGLE,{x:0.25,y:1.9,w:9.5,h:0.28,fill:{color:C.lightgray},line:{color:C.lightgray}});
    ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"].forEach((h,i)=>s.addText(h,{x:colX[i],y:1.93,w:colW[i],h:0.22,fontSize:7.5,bold:true,color:C.slate,fontFace:"Calibri"}));
    d.platforms.forEach((p,i)=>{
          const y=2.22+i*0.5,bg=i%2===0?"F4F3FD":C.white;
          s.addShape(pres.shapes.RECTANGLE,{x:0.25,y:y-0.04,w:9.5,h:0.48,fill:{color:bg},line:{color:C.lightgray}});
          s.addText(p.name,{x:colX[0],y,w:colW[0],h:0.28,fontSize:9.5,bold:true,color:C.navy,fontFace:"Calibri"});
          s.addText(String(p.mentions),{x:colX[1],y,w:colW[1],h:0.28,fontSize:9.5,color:C.navy,fontFace:"Calibri"});
          s.addText(String(p.citations),{x:colX[2],y,w:colW[2],h:0.28,fontSize:9.5,color:C.navy,fontFace:"Calibri"});
          const bv=Math.min(parseFloat(String(p.brandVisibility||0))||0,100),bvW=(bv/100)*2.0;
          s.addShape(pres.shapes.RECTANGLE,{x:colX[3]+0.48,y:y+0.08,w:2.0,h:0.14,fill:{color:C.lightgray},line:{color:C.lightgray}});
          if(bvW>0)s.addShape(pres.shapes.RECTANGLE,{x:colX[3]+0.48,y:y+0.08,w:bvW,h:0.14,fill:{color:C.navy},line:{color:C.navy}});
          s.addText(bv+"%",{x:colX[3],y,w:0.45,h:0.28,fontSize:8,bold:true,color:C.navy,fontFace:"Calibri"});
          const dc=Math.min(parseFloat(String(p.domainCoverage||0))||0,100),dcW=(dc/100)*2.0;
          s.addShape(pres.shapes.RECTANGLE,{x:colX[4]+0.48,y:y+0.08,w:2.0,h:0.14,fill:{color:C.lightgray},line:{color:C.lightgray}});
          if(dcW>0)s.addShape(pres.shapes.RECTANGLE,{x:colX[4]+0.48,y:y+0.08,w:dcW,h:0.14,fill:{color:C.violet},line:{color:C.violet}});
          s.addText(dc+"%",{x:colX[4],y,w:0.45,h:0.28,fontSize:8,bold:true,color:C.violet,fontFace:"Calibri"});
    });
}

// ─── SLIDE 8: Brand Visibility by Platform ─────────────────────────────────
function buildSlide8(pres, d) {
    const s=pres.addSlide(); s.background={color:C.white};
    hdr(s,pres,d.brandName+" Brand Visibility by Platform & Theme",d.brandName);
    ftr(s,pres,d.brandName,d.domain);
    const rows=d.brandVisibilityByPlatform;
    if(!rows||rows.length===0){s.addText("No platform visibility data available.",{x:0.5,y:2.8,w:9,h:0.5,fontSize:12,color:C.slate,align:"center",fontFace:"Calibri"});return;}
    const platNames=Object.keys(rows[0]).filter(k=>k!=='theme'),themeColW=3.2,dataColW=(9.5-themeColW)/platNames.length,startX=0.25,headerY=1.08,rowH=0.35;
    s.addShape(pres.shapes.RECTANGLE,{x:startX,y:headerY,w:9.5,h:rowH,fill:{color:C.navy},line:{color:C.navy}});
    s.addText("Themes",{x:startX+0.08,y:headerY+0.07,w:themeColW-0.12,h:rowH-0.1,fontSize:8,bold:true,color:C.white,fontFace:"Calibri"});
    platNames.forEach((pn,pi)=>{
          const x=startX+themeColW+pi*dataColW;
          s.addText(pn,{x:x+0.04,y:headerY+0.05,w:dataColW-0.08,h:rowH-0.1,fontSize:8,bold:true,color:C.white,align:"center",fontFace:"Calibri",wrap:true});
    });
    rows.slice(0,11).forEach((row,ri)=>{
          const y=headerY+rowH+ri*rowH,bg=ri%2===0?C.offwhite:C.white;
          s.addShape(pres.shapes.RECTANGLE,{x:startX,y,w:9.5,h:rowH,fill:{color:bg},line:{color:C.lightgray}});
          s.addText(row.theme||"",{x:startX+0.08,y:y+0.07,w:themeColW-0.14,h:rowH-0.1,fontSize:7,color:C.navy,fontFace:"Calibri",wrap:true});
          platNames.forEach((pn,pi)=>{
                  const x=startX+themeColW+pi*dataColW,pct=parseFloat(String(row[pn]||0))||0;
                  let cc=bg;
                  if(pct>=15)cc=C.purple; else if(pct>=5)cc="C4C0EA";
                  if(cc!==bg)s.addShape(pres.shapes.RECTANGLE,{x:x+0.04,y:y+0.05,w:dataColW-0.08,h:rowH-0.1,fill:{color:cc},line:{color:cc}});
                  s.addText(pct>0?pct+"%":"0%",{x:x+0.04,y:y+0.07,w:dataColW-0.08,h:rowH-0.12,fontSize:8,color:pct>=5?C.white:C.slate,align:"center",fontFace:"Calibri"});
          });
    });
    s.addText("The above is a combination of all results from ChatGPT, AI Overviews, Claude and Perplexity.",{x:0.3,y:5.22,w:9.4,h:0.16,fontSize:7,italic:true,color:C.slate,align:"center",fontFace:"Calibri"});
}

// ─── STEP 3: Build PPTX ─────────────────────────────────────────────────────
function buildPPTX(data, outputPath) {
  const pres=new pptxgen();
  pres.layout="LAYOUT_16x9";
  pres.title=data.brandName+" GEO Audit \u2014 Atlas";
  pres.author="Pepper.inc Atlas";
  buildSlide1(pres,data);
  buildSlide2(pres,data);
  buildSlide3(pres,data);
  buildSlide4(pres,data);
  buildSlide5(pres,data);
  buildSlide6(pres,data);
  buildSlide7(pres,data);
  buildSlide8(pres,data);
  pres.writeFile({ fileName:outputPath });
  console.log("\u2705 PPTX written:", outputPath, "(8 dynamic slides)");
}

// ─── STEP 3b: Merge dynamic PPTX with template slides + inject Pepper logo ──
async function mergeWithTemplate(dynamicPptxPath) {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    console.log("\u26A0\uFE0F Template not found at", TEMPLATE_PATH, "- skipping merge");
    return;
  }
  console.log("\u{1F4CE} Merging template slides from:", TEMPLATE_PATH);

  const dynBuf = fs.readFileSync(dynamicPptxPath);
  const tplBuf = fs.readFileSync(TEMPLATE_PATH);

  const dynZip = await JSZip.loadAsync(dynBuf);
  const tplZip = await JSZip.loadAsync(tplBuf);

  /* ── count existing dynamic slides ── */
  let dynSlideCount = 0;
  while (dynZip.file("ppt/slides/slide" + (dynSlideCount + 1) + ".xml")) dynSlideCount++;
  console.log("  Dynamic slides found:", dynSlideCount);

  /* ── count template slides ── */
  let tplSlideCount = 0;
  while (tplZip.file("ppt/slides/slide" + (tplSlideCount + 1) + ".xml")) tplSlideCount++;
  console.log("  Template slides found:", tplSlideCount);

  /* ── Copy images / media from template that don't already exist ── */
  const tplMedia = Object.keys(tplZip.files).filter(f => f.startsWith("ppt/media/"));
  const dynMedia = new Set(Object.keys(dynZip.files).filter(f => f.startsWith("ppt/media/")));
  const mediaMap = {};
  let nextMediaIdx = 1;
  // find highest existing media index in dynamic
  dynMedia.forEach(f => {
    const m = f.match(/image(\d+)/);
    if (m) nextMediaIdx = Math.max(nextMediaIdx, parseInt(m[1]) + 1);
  });
  for (const mf of tplMedia) {
    const ext = mf.split(".").pop();
    const newName = "ppt/media/image" + nextMediaIdx + "." + ext;
    mediaMap[mf.replace("ppt/media/", "")] = newName.replace("ppt/media/", "");
    dynZip.file(newName, await tplZip.file(mf).async("uint8array"));
    nextMediaIdx++;
  }

  /* ── Prepare pepper logo for injection ── */
  const pepperLogoPath = path.join(__dirname, "public", "logos", "pepper-logo.png");
  let pepperMediaName = null;
  if (fs.existsSync(pepperLogoPath)) {
    const logoData = fs.readFileSync(pepperLogoPath);
    pepperMediaName = "image" + nextMediaIdx + ".png";
    dynZip.file("ppt/media/" + pepperMediaName, logoData);
    nextMediaIdx++;
    console.log("  Pepper logo added as:", pepperMediaName);
  }

  /* ── Copy slide layouts and masters from template ── */
  const tplLayouts = Object.keys(tplZip.files).filter(f => f.startsWith("ppt/slideLayouts/"));
  const dynLayouts = new Set(Object.keys(dynZip.files).filter(f => f.startsWith("ppt/slideLayouts/")));
  let nextLayoutIdx = 1;
  dynLayouts.forEach(f => { const m = f.match(/slideLayout(\d+)/); if (m) nextLayoutIdx = Math.max(nextLayoutIdx, parseInt(m[1]) + 1); });
  const layoutMap = {};
  for (const lf of tplLayouts) {
    if (lf.endsWith(".rels")) continue;
    const num = lf.match(/slideLayout(\d+)/);
    if (!num) continue;
    const newLayoutName = "ppt/slideLayouts/slideLayout" + nextLayoutIdx + ".xml";
    layoutMap["slideLayout" + num[1] + ".xml"] = "slideLayout" + nextLayoutIdx + ".xml";
    let layoutXml = await tplZip.file(lf).async("string");
    // remap media references in layout
    for (const [oldM, newM] of Object.entries(mediaMap)) {
      layoutXml = layoutXml.split(oldM).join(newM);
    }
    dynZip.file(newLayoutName, layoutXml);
    // Copy layout rels if exist
    const relsPath = lf.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels";
    if (tplZip.file(relsPath)) {
      let relsXml = await tplZip.file(relsPath).async("string");
      for (const [oldM, newM] of Object.entries(mediaMap)) {
        relsXml = relsXml.split(oldM).join(newM);
      }
      const newRelsPath = newLayoutName.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels";
      dynZip.file(newRelsPath, relsXml);
    }
    nextLayoutIdx++;
  }

  /* ── Copy each template slide ── */
  for (let ti = 1; ti <= tplSlideCount; ti++) {
    const newIdx = dynSlideCount + ti;
    console.log("  Copying template slide", ti, "-> slide" + newIdx);

    // Copy slide XML
    let slideXml = await tplZip.file("ppt/slides/slide" + ti + ".xml").async("string");
    // remap media references
    for (const [oldM, newM] of Object.entries(mediaMap)) {
      slideXml = slideXml.split(oldM).join(newM);
    }
    dynZip.file("ppt/slides/slide" + newIdx + ".xml", slideXml);

    // Copy slide rels
    const relsFile = tplZip.file("ppt/slides/_rels/slide" + ti + ".xml.rels");
    let relsXml = relsFile ? await relsFile.async("string") : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    // remap media references in rels
    for (const [oldM, newM] of Object.entries(mediaMap)) {
      relsXml = relsXml.split(oldM).join(newM);
    }
    // remap layout references
    for (const [oldL, newL] of Object.entries(layoutMap)) {
      relsXml = relsXml.split(oldL).join(newL);
    }

    // Inject pepper logo relationship if needed
    if (pepperMediaName) {
      const logoRelId = "rIdPepperLogo";
      if (!relsXml.includes(logoRelId)) {
        relsXml = relsXml.replace(
          "</Relationships>",
          '<Relationship Id="' + logoRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + pepperMediaName + '"/></Relationships>'
        );
      }

      // Inject the logo pill shape into slide XML
      // Position: x=9.0in, y=0.12in — converted to EMUs (1 inch = 914400 EMU)
      const logoShapeXml = '<p:sp><p:nvSpPr><p:cNvPr id="9999" name="PepperLogoPill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="8229600" y="109728"/><a:ext cx="777240" cy="237744"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="E2E1F0"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>';
      const logoImageXml = '<p:pic><p:nvPicPr><p:cNvPr id="9998" name="PepperLogoImg"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="' + logoRelId + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="8256240" y="128016"/><a:ext cx="685800" cy="182880"/></a:xfrm></p:spPr></p:pic>';

      // Check if logo pill already exists
      if (!slideXml.includes("PepperLogoPill")) {
        // Insert before </p:spTree>
        slideXml = slideXml.replace("</p:spTree>", logoShapeXml + logoImageXml + "</p:spTree>");
        dynZip.file("ppt/slides/slide" + newIdx + ".xml", slideXml);
      }
    }

    dynZip.file("ppt/slides/_rels/slide" + newIdx + ".xml.rels", relsXml);
  }

  /* ── Update presentation.xml to include new slides ── */
  let presXml = await dynZip.file("ppt/presentation.xml").async("string");
  // Find the highest rId in presentation.xml
  let maxRId = 0;
  const rIdMatches = presXml.match(/rId(\d+)/g) || [];
  rIdMatches.forEach(r => { const n = parseInt(r.replace("rId", "")); if (n > maxRId) maxRId = n; });

  for (let ti = 1; ti <= tplSlideCount; ti++) {
    const newIdx = dynSlideCount + ti;
    maxRId++;
    const rId = "rId" + maxRId;
    // Add sldIdLst entry
    const sldId = 256 + newIdx;
    presXml = presXml.replace("</p:sldIdLst>", '<p:sldId id="' + sldId + '" r:id="' + rId + '"/></p:sldIdLst>');
  }
  dynZip.file("ppt/presentation.xml", presXml);

  /* ── Update ppt/_rels/presentation.xml.rels ── */
  let presRels = await dynZip.file("ppt/_rels/presentation.xml.rels").async("string");
  let rIdCounter = maxRId - tplSlideCount; // reset to where we started adding
  for (let ti = 1; ti <= tplSlideCount; ti++) {
    const newIdx = dynSlideCount + ti;
    rIdCounter++;
    const rId = "rId" + rIdCounter;
    presRels = presRels.replace(
      "</Relationships>",
      '<Relationship Id="' + rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + newIdx + '.xml"/></Relationships>'
    );
  }
  dynZip.file("ppt/_rels/presentation.xml.rels", presRels);

  /* ── Update [Content_Types].xml ── */
  let contentTypes = await dynZip.file("[Content_Types].xml").async("string");
  for (let ti = 1; ti <= tplSlideCount; ti++) {
    const newIdx = dynSlideCount + ti;
    const override = '<Override PartName="/ppt/slides/slide' + newIdx + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
    if (!contentTypes.includes("slide" + newIdx + ".xml")) {
      contentTypes = contentTypes.replace("</Types>", override + "</Types>");
    }
  }
  // Also add layout content types if needed
  for (const [oldL, newL] of Object.entries(layoutMap)) {
    const override = '<Override PartName="/ppt/slideLayouts/' + newL + '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
    if (!contentTypes.includes(newL)) {
      contentTypes = contentTypes.replace("</Types>", override + "</Types>");
    }
  }
  dynZip.file("[Content_Types].xml", contentTypes);

  /* ── Write merged PPTX back ── */
  const merged = await dynZip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(dynamicPptxPath, merged);
  console.log("\u2705 Merged PPTX written:", dynamicPptxPath, "(" + (dynSlideCount + tplSlideCount) + " total slides)");
}

// ─── API: POST /generate ───────────────────────────────────────────────────
app.post("/generate", async (req, res) => {
    // Accept either { reportId } or { url } for backward compatibility
           let reportId = req.body.reportId;
    if (!reportId && req.body.url) {
          reportId = extractReportId(req.body.url);
    }
    if (!reportId) {
          return res.status(400).json({ error: "Please provide a reportId or a valid Atlas report URL." });
    }

           const id = uuidv4(), pptxOut = path.join(TMP, id + ".pptx"), pdfOut = path.join(TMP, id + ".pdf");

           try {
                 const apiData = await fetchReportData(reportId);
                 const data = normalizeData(apiData);

      console.log("\n\u{1F5BC} Building PPTX for:", data.brandName);
                 buildPPTX(data, pptxOut);
    await new Promise(r => setTimeout(r, 500));

    console.log("\u{1F4CE} Merging with template slides...");
    await mergeWithTemplate(pptxOut);
    await new Promise(r => setTimeout(r, 500));

      console.log("\u{1F4C4} Converting to PDF...");
                 execSync(`soffice --headless --convert-to pdf --outdir ${TMP} ${pptxOut}`, { timeout: 60000 });
                 if (!fs.existsSync(pdfOut)) throw new Error("PDF conversion failed");

      const fileName = `${data.brandName} x Pepper - GEO report.pdf`;
                 res.download(pdfOut, fileName, (err) => {
                         try { fs.unlinkSync(pptxOut); } catch {}
                         try { fs.unlinkSync(pdfOut); } catch {}
                         if (err && !res.headersSent) res.status(500).json({ error: "Download failed." });
                 });
           } catch (err) {
                 console.error("\u274C Error:", err.message);
                 res.status(500).json({ error: err.message || "Generation failed." });
           
});

app.get("/health", (_, res) => res.json({ status: "ok" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`\u{1F680} Atlas PDF Generator on port ${PORT}`));
