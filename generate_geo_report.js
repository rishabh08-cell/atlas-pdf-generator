/**
 * Atlas GEO Report → PPTX Generator v2
 *
 * Slides:
 *  1. Cover
 *  2. Brand Leaderboard
 *  3. Competitor Mentions
 *  4. AI Platform Breakdown
 *  5. Competitor Visibility Matrix (max 6 brands, no overflow)
 *  6. Brand Pages Cited
 *  7. Prompt Themes Sample (3 themes, 10 prompts each)
 *  8. Key Insights
 */

const pptxgen = require("pptxgenjs");
const fs      = require("fs");

const args    = process.argv.slice(2);
const dataFile = args[args.indexOf("--data")   + 1] || "report_data.json";
const outFile  = args[args.indexOf("--output") + 1] || "geo_report.pptx";
const data     = JSON.parse(fs.readFileSync(dataFile, "utf8"));

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  navy:     "1A1A3E",
  purple:   "3D3A8C",
  violet:   "5B4FBE",
  lilac:    "9B93E3",
  orange:   "F4A419",
  white:    "FFFFFF",
  offwhite: "F8F7FF",
  slate:    "64748B",
  gray:     "E8E6F5",
  darkgray: "2D2B55",
  rowalt:   "F2F1FB",
};

const mkS = () => ({ type:"outer", blur:8, offset:2, angle:135, color:"000000", opacity:0.1 });

function heatBg(val, isClient) {
  if (isClient) return val >= 50 ? C.purple : val >= 25 ? C.violet : C.lilac;
  if (val >= 70) return C.navy;
  if (val >= 40) return C.violet;
  if (val >= 20) return C.lilac;
  if (val >  0)  return "C8C3EE";   // faint purple for low values
  return "E4E2F5";                   // consistent light tint for 0% — never pure white
}
function heatFg(val) { return val >= 20 ? C.white : C.navy; }

// ─── Shared helpers ───────────────────────────────────────────────────────────

function hdr(s, pres, title, sub) {
  // Solid bar only — NO decorative shapes on content slides
  s.addShape(pres.shapes.RECTANGLE, { x:0,   y:0,    w:8.27,  h:0.587, fill:{color:C.navy},  line:{color:C.navy}  });
  s.addText("pepper",         { x:0.248,  y:0.096, w:0.91,  h:0.395, fontSize:14, bold:true, color:C.orange, fontFace:"Calibri", margin:0 });
  s.addText("by pepper.inc", { x:1.183, y:0.149, w:1.158,  h:0.288, fontSize:8,  color:C.lilac,  fontFace:"Calibri", margin:0 });
  if (sub) s.addText(sub.toUpperCase(), { x:0, y:0.107, w:8.022, h:0.363, fontSize:8, color:C.lilac, align:"right", charSpacing:2, fontFace:"Calibri", margin:0 });
  s.addText(title, { x:0.248, y:0.715, w:7.774, h:0.533, fontSize:20, bold:true, color:C.navy, fontFace:"Calibri", margin:0 });
}

function ftr(s, pres, brand, domain) {
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:5.76, w:8.27, h:0.24, fill:{color:C.gray}, line:{color:C.gray} });
  s.addText(`${brand}  ·  ${domain}  ·  GEO Audit by Pepper`, { x:0.248, y:5.771, w:7.774, h:0.213, fontSize:7, color:C.slate, fontFace:"Calibri", margin:0 });
}

// ─── SLIDE 1: Cover ───────────────────────────────────────────────────────────
function slideCover(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Decorative orbs — only on cover, pushed to right so they NEVER touch content
  s.addShape(pres.shapes.OVAL, { x:5.954, y:-0.9, w:3.308, h:4.267, fill:{color:C.darkgray}, line:{color:C.darkgray} });
  s.addShape(pres.shapes.OVAL, { x:6.533, y:-0.3, w:2.15, h:2.773, fill:{color:C.purple},   line:{color:C.purple}   });

  s.addText("GEO AUDIT REPORT", { x:0.414, y:0.96,  w:5.376, h:0.373, fontSize:10, color:C.orange, bold:true, charSpacing:4, fontFace:"Calibri", margin:0 });
  s.addText(d.brandName,         { x:0.414, y:1.387,  w:5.376, h:1.387,  fontSize:52, bold:true, color:C.white,  fontFace:"Calibri", margin:0 });
  s.addText(d.domain,            { x:0.414, y:2.827, w:4.135, h:0.405, fontSize:14, color:C.lilac, fontFace:"Calibri", margin:0 });
  s.addShape(pres.shapes.RECTANGLE, { x:0.414, y:3.36, w:1.075, h:0.043, fill:{color:C.orange}, line:{color:C.orange} });

  [
    { label:"Total Mentions",   value: d.overview.totalMentions.toLocaleString() },
    { label:"Brand Visibility",   value: d.overview.avgBrandCoverage },
    { label:"AI Platforms",     value: String(d.overview.platforms) },
    { label:"Leaderboard Rank", value: d.overview.leaderboardRank },
  ].forEach((st, i) => {
    const x = 0.5 + i * 2.3;
    s.addText(st.value, { x, y:3.573, w:1.737, h:0.619, fontSize:26, bold:true, color:C.orange, fontFace:"Calibri", margin:0 });
    s.addText(st.label, { x, y:4.171, w:1.737, h:0.299, fontSize:9,  color:C.lilac,  fontFace:"Calibri", margin:0 });
  });

  s.addText("Powered by pepper", { x:0.414, y:5.493, w:7.443, h:0.299, fontSize:8, color:C.slate, fontFace:"Calibri", margin:0 });
}

// ─── SLIDE 2: Brand Leaderboard ───────────────────────────────────────────────
function slideLeaderboard(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  hdr(s, pres, "Brand Leaderboard", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const brands = d.leaderboard.slice(0, 3);
  const maxM   = Math.max(...brands.map(b => b.mentions));
  const barW = 1.6, gap = 1.2;
  const startX = (10 - (brands.length * barW + (brands.length-1) * gap)) / 2;
  const bottom = 4.8, chartH = 2.7;

  brands.forEach((brand, i) => {
    const x    = startX + i * (barW + gap);
    const barH = Math.max((brand.mentions / maxM) * chartH, 0.12);
    const barY = bottom - barH;
    const isC  = brand.name === d.brandName;
    const col  = isC ? C.orange : (i === 1 ? "BDBDCD" : "C0824A");

    s.addText(brand.name, { x:x-0.2, y:barY-0.44, w:barW+0.4, h:0.341, fontSize:13, bold:isC, color:isC?C.orange:C.navy, align:"center", fontFace:"Calibri", margin:0 });
    s.addShape(pres.shapes.OVAL, { x:x+barW/2-0.28, y:barY-0.82, w:0.463, h:0.597, fill:{color:C.white}, line:{color:C.gray} });
    s.addText(`#${brand.rank}`, { x:x+barW/2-0.28, y:barY-0.82, w:0.463, h:0.597, fontSize:14, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x, y:barY, w:barW, h:barH, fill:{color:col}, line:{color:col}, shadow:mkS() });
    if (isC) s.addText("👑", { x:x+barW/2-0.28, y:barY+0.1, w:0.463, h:0.427, fontSize:22, align:"center", margin:0 });
    s.addText(`${brand.mentions.toLocaleString()} mentions`, { x:x-0.2, y:bottom+0.12, w:barW+0.4, h:0.267, fontSize:9, color:C.slate, align:"center", fontFace:"Calibri", margin:0 });
  });
}

// ─── SLIDE 3: Competitor Mentions ─────────────────────────────────────────────
function slideCompetitorMentions(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  hdr(s, pres, `Competitor Mentions vs. ${d.brandName}`, d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const comps  = d.competitorMentions.slice(0, 10);
  const maxPct = Math.max(...comps.map(c => c.percentage), 1);

  comps.forEach((comp, i) => {
    const y   = 1.38 + i * 0.375;
    const isC = comp.name === d.brandName;

    s.addShape(pres.shapes.OVAL, { x:0.207, y:y+0.05, w:0.215, h:0.277, fill:{color:isC?C.purple:C.gray}, line:{color:isC?C.purple:C.gray} });
    s.addText(comp.name[0].toUpperCase(), { x:0.207, y:y+0.05, w:0.215, h:0.277, fontSize:9, bold:true, color:isC?C.white:C.navy, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
    s.addText(comp.name, { x:0.48, y:y+0.07, w:1.241, h:0.245, fontSize:10, bold:isC, color:isC?C.purple:C.navy, fontFace:"Calibri", margin:0 });

    const barW = (comp.percentage / maxPct) * 6.0;
    s.addShape(pres.shapes.RECTANGLE, { x:1.819, y:y+0.07, w:Math.max(barW, 0.06), h:0.235, fill:{color:isC?C.navy:C.gray}, line:{color:isC?C.navy:C.gray} });
    s.addText(`${comp.percentage}%  ${comp.mentions.toLocaleString()} mentions`, { x:1.861+barW, y:y+0.07, w:2.316, h:0.235, fontSize:9, bold:isC, color:isC?C.navy:C.slate, fontFace:"Calibri", margin:0 });
  });
}

// ─── SLIDE 4: AI Platform Breakdown ──────────────────────────────────────────
function slidePlatforms(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  hdr(s, pres, `${d.brandName} Mentions by AI Platform`, d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  // Stat cards
  [
    { label:"Total Brand Mentions",   value: d.platformData.totalMentions.toLocaleString() },
    { label:"Total Domain Citations", value: d.platformData.totalCitations.toLocaleString() },
    { label:"Brand Visibility",     value: d.platformData.avgBrandCoverage },
    { label:"Domain Prompt Presence",    value: d.platformData.avgDomainCoverage },
  ].forEach((st, i) => {
    const x = 0.25 + i * 2.42;
    s.addShape(pres.shapes.RECTANGLE, { x, y:1.429, w:1.886, h:0.811, fill:{color:C.white}, line:{color:C.gray}, shadow:mkS() });
    s.addText(st.value, { x, y:1.472, w:1.886, h:0.405, fontSize:18, bold:true, color:C.navy,  align:"center", fontFace:"Calibri", margin:0 });
    s.addText(st.label, { x, y:1.856, w:1.886, h:0.299, fontSize:8,  color:C.slate, align:"center", fontFace:"Calibri", margin:0 });
  });

  // Header row
  const cx = [0.25, 2.15, 3.4, 4.65, 7.2];
  const cw = [1.85, 1.2,  1.2, 2.45, 2.5];
  s.addShape(pres.shapes.RECTANGLE, { x:0.207, y:2.507, w:7.857, h:0.309, fill:{color:C.gray}, line:{color:C.gray} });
  ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"]
    .forEach((h, i) => s.addText(h, { x:cx[i], y:2.528, w:cw[i], h:0.267, fontSize:8, bold:true, color:C.slate, fontFace:"Calibri", margin:0 }));

  d.platformData.platforms.forEach((p, i) => {
    const y = 2.7 + i * 0.54;
    if (i % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x:0.207, y:y-0.04, w:7.857, h:0.555, fill:{color:C.rowalt}, line:{color:C.rowalt} });

    s.addText(p.name,                       { x:cx[0], y, w:cw[0], h:0.32, fontSize:10, bold:true, color:C.navy, fontFace:"Calibri", margin:0 });
    s.addText(p.mentions.toLocaleString(),  { x:cx[1], y, w:cw[1], h:0.32, fontSize:10, color:C.navy, fontFace:"Calibri", margin:0 });
    s.addText(p.citations.toLocaleString(), { x:cx[2], y, w:cw[2], h:0.32, fontSize:10, color:C.navy, fontFace:"Calibri", margin:0 });

    const bvW = Math.max((p.brandVisibility / 100) * (cw[3]-0.52), 0.05);
    s.addText(`${p.brandVisibility}%`, { x:cx[3], y, w:0.364, h:0.32, fontSize:9, bold:true, color:C.navy,  fontFace:"Calibri", margin:0 });
    s.addShape(pres.shapes.RECTANGLE,  { x:cx[3]+0.46, y:y+0.07, w:bvW,                  h:0.181, fill:{color:C.navy},  line:{color:C.navy}  });
    s.addShape(pres.shapes.RECTANGLE,  { x:cx[3]+0.46+bvW, y:y+0.07, w:cw[3]-0.52-bvW,  h:0.181, fill:{color:C.gray},  line:{color:C.gray}  });

    const dcW = Math.max((p.domainCoverage / 100) * (cw[4]-0.44), 0.05);
    s.addText(`${p.domainCoverage}%`, { x:cx[4], y, w:0.298, h:0.32, fontSize:9, bold:true, color:C.violet, fontFace:"Calibri", margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x:cx[4]+0.38, y:y+0.07, w:dcW,                   h:0.181, fill:{color:C.violet}, line:{color:C.violet} });
    s.addShape(pres.shapes.RECTANGLE, { x:cx[4]+0.38+dcW, y:y+0.07, w:cw[4]-0.44-dcW,   h:0.181, fill:{color:C.gray},   line:{color:C.gray}  });
  });
}

// ─── SLIDE 5: Competitor Visibility Matrix ────────────────────────────────────
// Hard-cap at 6 brands so columns never overflow the 10" slide width
function slideCompetitorMatrix(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  hdr(s, pres, "Brand Visibility × Competitors", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const mx     = d.competitorVisibilityMatrix;
  const brands = mx.brands.slice(0, 6);           // MAX 6
  const rows   = mx.rows;

  const labelW  = 2.75;
  const tableW  = 9.5 - labelW - 0.1;
  const cellW   = tableW / brands.length;
  const startX  = 0.25 + labelW + 0.1;
  const startY  = 1.38;
  const cellH   = 0.375;

  // Column headers
  brands.forEach((b, i) => {
    const isC = b === d.brandName;
    const x   = startX + i * cellW;
    s.addShape(pres.shapes.RECTANGLE, { x, y:startY, w:cellW-0.03, h:0.32, fill:{color:isC?C.purple:C.navy}, line:{color:isC?C.purple:C.navy} });
    s.addText(b, { x, y:startY, w:cellW-0.03, h:0.32, fontSize:8, bold:isC, color:C.white, align:"center", fontFace:"Calibri", margin:0 });
  });

  // Data rows
  rows.forEach((row, ti) => {
    const y = startY + 0.34 + ti * cellH;
    if (ti % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x:0.207, y, w:7.857, h:cellH-0.02, fill:{color:C.rowalt}, line:{color:C.rowalt} });

    s.addText(row.theme, { x:0.232, y:y+0.04, w:labelW-0.08, h:cellH-0.1, fontSize:8.5, color:C.navy, fontFace:"Calibri", margin:0, wrap:true });

    brands.forEach((b, bi) => {
      const val = row.values[b] ?? 0;
      const isC = b === d.brandName;
      const x   = startX + bi * cellW;
      const bg  = heatBg(val, isC);
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.02, y:y+0.03, w:cellW-0.07, h:cellH-0.08, fill:{color:bg}, line:{color:bg} });
      s.addText(`${val}%`, { x:x+0.02, y:y+0.03, w:cellW-0.07, h:cellH-0.08, fontSize:9, bold:isC, color:heatFg(val), align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
    });
  });

  // Inline legend — top right corner of header bar
  const legend = [{label:"70%+",color:C.navy},{label:"40–69%",color:C.violet},{label:"20–39%",color:C.lilac},{label:"<20%",color:"DDDAF5"}];
  s.addText("Scale:", { x:4.714, y:0.779, w:0.538, h:0.192, fontSize:7, color:C.lilac, fontFace:"Calibri", margin:0 });
  legend.forEach((l, i) => {
    const lx = 6.4 + i * 0.88;
    s.addShape(pres.shapes.RECTANGLE, { x:lx, y:0.811, w:0.124, h:0.139, fill:{color:l.color}, line:{color:l.color==="DDDAF5"?C.gray:l.color} });
    s.addText(l.label, { x:lx+0.18, y:0.789, w:0.562, h:0.192, fontSize:7, color:C.lilac, fontFace:"Calibri", margin:0 });
  });
}

// ─── SLIDE 6: Brand Pages Cited ───────────────────────────────────────────────
function slideBrandPages(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  hdr(s, pres, `${d.brandName} Pages Cited by AI`, d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const pages   = d.brandPages.slice(0, 10);
  const maxProm = Math.max(...pages.map(p => p.prompts), 1);

  // Column headers
  s.addShape(pres.shapes.RECTANGLE, { x:0.207, y:1.451, w:7.857, h:0.299, fill:{color:C.gray}, line:{color:C.gray} });
  s.addText("Page",             { x:0.43, y:1.472, w:4.218,  h:0.256, fontSize:8, bold:true, color:C.slate, fontFace:"Calibri", margin:0 });
  s.addText("Citation Frequency", { x:4.755, y:1.472, w:2.398,  h:0.256, fontSize:8, bold:true, color:C.slate, fontFace:"Calibri", margin:0 });
  s.addText("Prompts",          { x:7.319, y:1.472, w:0.703, h:0.256, fontSize:8, bold:true, color:C.slate, align:"right", fontFace:"Calibri", margin:0 });

  pages.forEach((page, i) => {
    const y    = 1.7 + i * 0.365;
    const isTop = i === 0;
    if (i % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x:0.207, y:y-0.03, w:7.857, h:0.368, fill:{color:C.rowalt}, line:{color:C.rowalt} });

    // Rank pill
    s.addShape(pres.shapes.OVAL, { x:0.215, y:y+0.04, w:0.182, h:0.235, fill:{color:isTop?C.orange:C.gray}, line:{color:isTop?C.orange:C.gray} });
    s.addText(String(i+1), { x:0.215, y:y+0.04, w:0.182, h:0.235, fontSize:7, bold:true, color:isTop?C.navy:C.slate, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });

    const name = page.name.length > 74 ? page.name.slice(0, 71) + "…" : page.name;
    s.addText(name, { x:0.438, y:y+0.04, w:4.259, h:0.256, fontSize:9, bold:isTop, color:isTop?C.purple:C.navy, fontFace:"Calibri", margin:0 });

    const barW = (page.prompts / maxProm) * 2.75;
    s.addShape(pres.shapes.RECTANGLE, { x:4.755, y:y+0.08, w:Math.max(barW, 0.05), h:0.16, fill:{color:isTop?C.orange:C.violet}, line:{color:isTop?C.orange:C.violet} });

    s.addText(String(page.prompts), { x:7.319, y:y+0.04, w:0.703, h:0.256, fontSize:10, bold:true, color:C.navy, align:"right", fontFace:"Calibri", margin:0 });
  });
}

// ─── SLIDE 7: Prompt Themes Sample ───────────────────────────────────────────
function slidePromptThemes(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  hdr(s, pres, "Sample Prompt Themes", d.brandName);
  ftr(s, pres, d.brandName, d.domain);

  const themes  = d.promptThemes.slice(0, 3);
  const colW    = 2.95, colGap = 0.17, startX = 0.25, startY = 1.37;

  themes.forEach((theme, ti) => {
    const x = startX + ti * (colW + colGap);

    // Theme header
    s.addShape(pres.shapes.RECTANGLE, { x, y:startY, w:colW, h:0.555, fill:{color:C.navy}, line:{color:C.navy}, shadow:mkS() });
    s.addShape(pres.shapes.RECTANGLE, { x, y:startY, w:0.041, h:0.555, fill:{color:C.orange}, line:{color:C.orange} });
    s.addText(theme.theme, { x:x+0.1, y:startY+0.05, w:colW-0.75, h:0.448, fontSize:9, bold:true, color:C.white, fontFace:"Calibri", margin:0, wrap:true });

    // Count badge
    s.addShape(pres.shapes.RECTANGLE, { x:x+colW-0.62, y:startY+0.11, w:0.455, h:0.299, fill:{color:C.orange}, line:{color:C.orange} });
    s.addText(`${theme.prompts.length}`, { x:x+colW-0.62, y:startY+0.11, w:0.455, h:0.299, fontSize:9, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });

    // Prompts list
    theme.prompts.slice(0, 10).forEach((prompt, pi) => {
      const py = startY + 0.62 + pi * 0.36;
      if (pi % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x, y:py-0.03, w:colW, h:0.363, fill:{color:C.rowalt}, line:{color:C.rowalt} });

      s.addText(String(pi + 1), { x:x+0.06, y:py+0.04, w:0.165, h:0.245, fontSize:8, color:C.lilac, align:"right", fontFace:"Calibri", margin:0 });
      const label = prompt.length > 38 ? prompt.slice(0, 36) + "…" : prompt;
      s.addText(label, { x:x+0.3, y:py+0.04, w:colW-0.34, h:0.245, fontSize:8.5, color:C.navy, fontFace:"Calibri", margin:0 });
    });
  });
}

// ─── SLIDE 8: Key Insights ────────────────────────────────────────────────────
function slideInsights(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  ftr(s, pres, d.brandName, d.domain);

  s.addText("KEY INSIGHTS", { x:0.414, y:0.341, w:7.443, h:0.352, fontSize:10, color:C.orange, bold:true, charSpacing:4, fontFace:"Calibri", margin:0 });
  s.addText(`${d.brandName} — GEO Audit Summary`, { x:0.414, y:0.693, w:7.443, h:0.512, fontSize:22, bold:true, color:C.white, fontFace:"Calibri", margin:0 });

  const pos = [
    { x:0.207, y:1.365 }, { x:4.176, y:1.365 },
    { x:0.207, y:3.072 }, { x:4.176, y:3.072 },
  ];
  d.keyInsights.slice(0, 4).forEach((ins, i) => {
    const { x, y } = pos[i];
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:3.763, h:1.547, fill:{color:C.darkgray}, line:{color:C.darkgray}, shadow:mkS() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w:0.05, h:1.547, fill:{color:C.orange},   line:{color:C.orange}   });
    s.addText(ins.label.toUpperCase(), { x:x+0.14, y:y+0.11, w:3.597, h:0.235, fontSize:7,  bold:true, color:C.orange, charSpacing:2, fontFace:"Calibri", margin:0 });
    s.addText(ins.stat,                { x:x+0.14, y:y+0.3,  w:3.597, h:0.533,  fontSize:24, bold:true, color:C.white,  fontFace:"Calibri", margin:0 });
    s.addText(ins.description,         { x:x+0.14, y:y+0.8,  w:3.597, h:0.587, fontSize:9,  color:C.lilac, fontFace:"Calibri", margin:0, wrap:true });
  });
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
async function main() {
  console.log(`\n📊 Atlas GEO Report Generator v2`);
  console.log(`   Brand: ${data.brandName} · ${data.domain}\n`);

  const pres = new pptxgen();
  pres.defineLayout({ name: "PRINT_8x6", width: 8.27, height: 6.0 });
  pres.layout = "PRINT_8x6";
  pres.title  = `${data.brandName} GEO Audit — Pepper`;
  pres.author = "Pepper.inc";

  slideCover(pres, data);              console.log("  ✓ Cover");
  slideLeaderboard(pres, data);        console.log("  ✓ Brand Leaderboard");
  slideCompetitorMentions(pres, data); console.log("  ✓ Competitor Mentions");
  slidePlatforms(pres, data);          console.log("  ✓ AI Platform Breakdown");
  slideCompetitorMatrix(pres, data);   console.log("  ✓ Competitor Visibility Matrix");
  slideBrandPages(pres, data);         console.log("  ✓ Brand Pages Cited");
  slidePromptThemes(pres, data);       console.log("  ✓ Prompt Themes Sample");
  slideInsights(pres, data);           console.log("  ✓ Key Insights");

  await pres.writeFile({ fileName: outFile });
  console.log(`\n✅  → ${outFile}\n`);
}

main().catch(e => { console.error("❌", e.message); process.exit(1); });
