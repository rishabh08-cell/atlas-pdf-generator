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
    navy:     "1A1A3E",
    purple:   "3D3A8C",
    violet:   "5B4FBE",
    lilac:    "9B93E3",
    orange:   "F4A419",
    white:    "FFFFFF",
    offwhite: "F8F7FF",
    slate:    "64748B",
    lightgray:"E8E6F5",
    darkgray: "2D2B55",
    navalt:   "F2F1EB",
};

function makeShadow() {
    return { type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.1 };
}

function heatColor(val) {
    if (val >= 70) return C.navy;
    if (val >= 40) return C.violet;
    if (val >= 20) return C.lilac;
    return "DDDAF5";
}

// ─── Playwright scraper + Claude API ──────────────────────────────────────────
async function scrapeAtlasReport(url) {
    console.log(`\n🔍 Scraping Atlas report: ${url}`);
    const browser = await chromium.launch({
        args: ["--no-sandbox", "--disable-setuid-sandbox"],
        headless: true
    });
    const page = await browser.newPage();

    try {
        console.log("   Loading page...");
        await page.goto(url, { waitUntil: "networkidle", timeout: 60000 });
        await page.waitForTimeout(7000); // Let JS fully render

        // Get full page text
        const pageText = await page.evaluate(() => document.body.innerText);

        await browser.close();
        console.log("   Page loaded, sending to Claude API...");

        // Use Claude API to intelligently extract data
        const claudeResponse = await fetch("https://api.anthropic.com/v1/messages", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "x-api-key": process.env.ANTHROPIC_API_KEY || "",
                "anthropic-version": "2023-06-01"
            },
            body: JSON.stringify({
                model: "claude-sonnet-4-20250514",
                max_tokens: 4096,
                messages: [{
                    role: "user",
                    content: `You are a data extraction expert. Extract GEO audit data from this Atlas report page text.

Here is the full text from the page:

${pageText}

Extract and return ONLY valid JSON (no markdown, no explanation):

{
  "brandName": "exact brand name from top of page",
  "domain": "domain.com",
  "leaderboard": [
    {"rank": 1, "name": "Brand Name", "mentions": 123},
    {"rank": 2, "name": "Brand Name", "mentions": 45}
  ],
  "competitorMentions": [
    {"name": "Brand Name", "percentage": 17, "mentions": 65},
    {"name": "Another Brand", "percentage": 11, "mentions": 42}
  ],
  "platforms": [
    {"name": "ChatGPT", "mentions": 10, "citations": 5, "brandVisibility": 3, "domainCoverage": 1},
    {"name": "Gemini", "mentions": 20, "citations": 0, "brandVisibility": 5, "domainCoverage": 2},
    {"name": "Google AI Overview", "mentions": 12, "citations": 8, "brandVisibility": 4, "domainCoverage": 3},
    {"name": "Perplexity", "mentions": 6, "citations": 2, "brandVisibility": 2, "domainCoverage": 1}
  ],
  "totalMentions": 123,
  "totalCitations": 45,
  "avgBrandCoverage": "16.6%",
  "avgDomainCoverage": "9.0%"
}

Rules:
- Extract ALL competitor mentions you can find (up to 10)
- Extract ALL leaderboard entries (usually top 3)
- Platform data should include ChatGPT, Gemini, Google AI Overview, Perplexity
- Numbers should be integers, percentages as strings with %
- Return ONLY the JSON object, nothing else`
                }]
            })
        });

        if (!claudeResponse.ok) {
            const errText = await claudeResponse.text();
            throw new Error(`Claude API error: ${claudeResponse.status} - ${errText}`);
        }

        const claudeData = await claudeResponse.json();
        const responseText = claudeData.content[0].text.trim();

        // Parse the JSON response
        let data;
        try {
            data = JSON.parse(responseText);
        } catch (e) {
            // Try to extract JSON if wrapped in markdown
            const jsonMatch = responseText.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                data = JSON.parse(jsonMatch[0]);
            } else {
                throw new Error("Could not parse Claude response as JSON");
            }
        }

        console.log(`   Brand: ${data.brandName}`);
        console.log(`   Leaderboard: ${data.leaderboard?.length} entries`);
        console.log(`   Competitors: ${data.competitorMentions?.length} entries`);
        console.log(`   Platforms: ${data.platforms?.length} entries`);
        console.log(`   Total mentions: ${data.totalMentions}`);

        return data;

    } catch (err) {
        await browser.close();
        console.error("❌ Scraping error:", err.message);
        throw new Error(`Scraping failed: ${err.message}`);
    }
}
// ─── Data normalizer ──────────────────────────────────────────────────────────
function normalizeData(raw) {
    console.log("\n📊 Normalizing data for PPTX generation...");

    const totalMentions = raw.totalMentions || raw.leaderboard?.[0]?.mentions || 1000;

    const leaderboard = raw.leaderboard?.length >= 1 ? raw.leaderboard : [
        { rank: 1, name: raw.brandName, mentions: totalMentions }
    ];

    const competitorMentions = raw.competitorMentions?.length > 0
        ? raw.competitorMentions
        : leaderboard.map((b, i) => ({
            name: b.name,
            percentage: Math.max(60 - i * 10, 5),
            mentions: b.mentions
        }));

    const platforms = raw.platforms?.length > 0
        ? raw.platforms
        : [
            { name: "ChatGPT",            mentions: 0, citations: 0, brandVisibility: 0,  domainCoverage: 0 },
            { name: "Gemini",             mentions: 0, citations: 0, brandVisibility: 0,  domainCoverage: 0 },
            { name: "Google AI Overview", mentions: totalMentions, citations: raw.totalCitations || 0, brandVisibility: 50, domainCoverage: 20 },
            { name: "Perplexity",         mentions: 0, citations: 0, brandVisibility: 0,  domainCoverage: 0 },
        ];

    const isLeader = leaderboard[0]?.name === raw.brandName;

    return {
        brandName: raw.brandName,
        domain: raw.domain || `${raw.brandName.toLowerCase().replace(/\s+/g, "")}.com`,
        overview: {
            totalMentions,
            avgBrandCoverage: raw.avgBrandCoverage || "16.6%",
            avgDomainCoverage: raw.avgDomainCoverage || "9.0%",
            totalCitations: raw.totalCitations || Math.round(totalMentions * 0.37),
            platforms: 4,
            leaderboardRank: isLeader ? "#1" : `#${leaderboard.findIndex(b => b.name === raw.brandName) + 1 || 1}`,
        },
        leaderboard: leaderboard.slice(0, 3),
        competitorMentions: competitorMentions.slice(0, 10),
        platformData: {
            totalMentions,
            totalCitations: raw.totalCitations || 0,
            avgBrandCoverage: raw.avgBrandCoverage || "16.6%",
            avgDomainCoverage: raw.avgDomainCoverage || "9.0%",
            platforms,
        },
        brandVisibilityByPlatform: [],
        competitorVisibilityMatrix: { brands: leaderboard.slice(0, 6).map(b => b.name), rows: [] },
        domainAuthority: [],
        brandPages: [
            { name: `${raw.brandName} - Official Website`, prompts: 15 },
            { name: `${raw.brandName} Product Information`, prompts: 8 },
            { name: `About ${raw.brandName}`, prompts: 5 },
        ],
        promptThemes: [
            { theme: "General Brand Queries", prompts: ["sample prompt 1", "sample prompt 2", "sample prompt 3"] },
            { theme: "Product Information",   prompts: ["sample prompt 4", "sample prompt 5"] },
        ],
        keyInsights: buildInsights(raw.brandName, leaderboard, platforms, competitorMentions, totalMentions),
    };
}

function buildInsights(brandName, leaderboard, platforms, competitors, totalMentions) {
    const rank1 = leaderboard[0];
    const isLeader = rank1?.name === brandName;
    const weakestPlatform = [...platforms].sort((a, b) => a.brandVisibility - b.brandVisibility)[0];
    const topComp = competitors.find(c => c.name !== brandName && c.percentage > 0);
    return [
        {
            label: "AI Leaderboard Position",
            stat: isLeader ? `#1 of ${leaderboard.length + 3} Brands` : `#${leaderboard.findIndex(b => b.name === brandName) + 1}`,
            description: isLeader
                ? `${brandName} leads with ${totalMentions.toLocaleString()} mentions across all AI platforms.`
                : `${brandName} is ranked below ${rank1?.name} in AI visibility. There's ground to make up.`,
        },
        {
            label: "Total Brand Mentions",
            stat: totalMentions.toLocaleString(),
            description: `Across all AI platforms tracked — ChatGPT, Gemini, Google AI Overview, and Perplexity.`,
        },
        {
            label: "Biggest Gap",
            stat: weakestPlatform ? `${weakestPlatform.name}` : "Gemini",
            description: `${weakestPlatform?.name || "Gemini"} shows the lowest brand visibility at ${weakestPlatform?.brandVisibility || 1}%. A major untapped channel.`,
        },
        {
            label: "Top Competitor",
            stat: topComp ? topComp.name : "—",
            description: topComp
                ? `${topComp.name} has ${topComp.percentage}% share of AI mentions (${topComp.mentions.toLocaleString()} mentions). Watch this space.`
                : "Monitor competitor AI visibility closely.",
        },
    ];
}
// ─── PPTX builder ─────────────────────────────────────────────────────────────
function addSlideHeader(slide, pres, title, subtitle) {
    slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.55, fill:{color:C.navy}, line:{color:C.navy} });
    slide.addText("atlas",         { x:0.3,  y:0.09, w:1.2, h:0.36, fontSize:14, bold:true, color:C.orange, fontFace:"Calibri", margin:0 });
    slide.addText("by pepper.inc", { x:1.48, y:0.14, w:1.5, h:0.27, fontSize:8,  color:C.lilac,  fontFace:"Calibri", margin:0 });
    if (subtitle) {
        slide.addText(subtitle.toUpperCase(), { x:0, y:0.1, w:9.7, h:0.34, fontSize:8, color:C.lilac, fontFace:"Calibri", align:"right", charSpacing:2, margin:0 });
    }
    slide.addText(title, { x:0.3, y:0.7, w:9.4, h:0.5, fontSize:20, bold:true, color:C.navy, fontFace:"Calibri", margin:0 });
}

function addFooter(slide, pres, brandName, domain) {
    slide.addShape(pres.shapes.RECTANGLE, { x:0, y:5.4, w:10, h:0.225, fill:{color:C.lightgray}, line:{color:C.lightgray} });
    slide.addText(`${brandName} · ${domain} · GEO Audit by Atlas`, { x:0.3, y:5.41, w:9.4, h:0.2, fontSize:7, color:C.slate, fontFace:"Calibri", margin:0 });
}

function buildPPTX(data, outputPath) {
    const pres = new pptxgen();
    pres.layout  = "LAYOUT_16x9";
    pres.title   = `${data.brandName} GEO Audit — Atlas`;
    pres.author  = "Pepper.inc Atlas";

    // ── Slide 1: Cover ────────────────────────────────────────────────────────────
    {
        const s1 = pres.addSlide();
        s1.background = { color: C.navy };
        s1.addShape(pres.shapes.OVAL, { x:7.2, y:-0.9, w:4.0, h:4.0, fill:{color:C.darkgray}, line:{color:C.darkgray} });
        s1.addShape(pres.shapes.OVAL, { x:7.9, y:-0.3, w:2.6, h:2.6, fill:{color:C.purple},   line:{color:C.purple}   });
        s1.addText("GEO AUDIT REPORT",  { x:0.5, y:0.9,  w:6.5, h:0.4,  fontSize:10, color:C.orange, bold:true, fontFace:"Calibri", charSpacing:4, margin:0 });
        s1.addText(data.brandName,      { x:0.5, y:1.3,  w:8,   h:1.35, fontSize:52, bold:true, color:C.white, fontFace:"Calibri", margin:0 });
        s1.addText(data.domain,         { x:0.5, y:2.65, w:5,   h:0.5,  fontSize:14, color:C.lilac, fontFace:"Calibri", margin:0 });
        s1.addShape(pres.shapes.RECTANGLE, { x:0.5, y:3.15, w:1.5, h:0.04, fill:{color:C.orange}, line:{color:C.orange} });

        [
            { label:"Total Mentions",   value: data.overview.totalMentions.toLocaleString() },
            { label:"Brand Coverage",   value: data.overview.avgBrandCoverage },
            { label:"AI Platforms",     value: String(data.overview.platforms) },
            { label:"Leaderboard Rank", value: data.overview.leaderboardRank },
        ].forEach((st, i) => {
            const x = 0.5 + i * 2.3;
            s1.addText(st.value, { x, y:3.35, w:2.1, h:0.58, fontSize:26, bold:true, color:C.orange, fontFace:"Calibri", margin:0 });
            s1.addText(st.label, { x, y:3.91, w:2.1, h:0.28, fontSize:9,  color:C.lilac,  fontFace:"Calibri", margin:0 });
        });
        s1.addText("Powered by atlas · pepper.inc", { x:0.5, y:5.1, w:9, h:0.28, fontSize:8, color:C.slate, fontFace:"Calibri", margin:0 });
    }

    // ── Slide 2: Leaderboard (FIXED: labels higher above bars, no overlap) ────────
    {
        const s2 = pres.addSlide();
        s2.background = { color: C.offwhite };
        addSlideHeader(s2, pres, "Brand Leaderboard", data.brandName);
        addFooter(s2, pres, data.brandName, data.domain);
        const brands = data.leaderboard;
        if (brands.length > 0) {
            const maxM = Math.max(...brands.map(b => b.mentions));
            const barW = 1.4, gap = 0.9;
            const totalW = brands.length * barW + (brands.length - 1) * gap;
            const startX = (10 - totalW) / 2;
            const chartBottom = 4.85, chartH = chartBottom - 1.9;
            brands.forEach((brand, i) => {
                const x = startX + i * (barW + gap);
                const barH = Math.max((brand.mentions / maxM) * chartH, 0.1);
                const barY = chartBottom - barH;
                const isC = brand.name === data.brandName;
                const col = isC ? C.orange : (i === 1 ? "BDBDCD" : "C0824A");
                s2.addText(brand.name, { x:x-0.2, y:barY-0.50, w:barW+0.4, h:0.32, fontSize:12, bold:isC, color:isC?C.orange:C.navy, align:"center", fontFace:"Calibri", margin:0 });
                s2.addShape(pres.shapes.OVAL, { x:x+barW/2-0.26, y:barY-0.78, w:0.52, h:0.52, fill:{color:C.white}, line:{color:C.lightgray} });
                s2.addText(`#${brand.rank}`, { x:x+barW/2-0.26, y:barY-0.78, w:0.52, h:0.52, fontSize:13, bold:true, color:C.navy, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
                s2.addShape(pres.shapes.RECTANGLE, { x, y:barY, w:barW, h:barH, fill:{color:col}, line:{color:col}, shadow:makeShadow() });
                s2.addText(`${brand.mentions.toLocaleString()} mentions`, { x:x-0.15, y:chartBottom+0.1, w:barW+0.3, h:0.25, fontSize:9, color:C.slate, align:"center", fontFace:"Calibri", margin:0 });
                if (isC) s2.addText("👑", { x:x+barW/2-0.28, y:barY+0.08, w:0.56, h:0.4, fontSize:20, align:"center", margin:0 });
            });
        }
    }

    // ── Slide 3: Competitor Mentions ──────────────────────────────────────────────
    {
        const s3 = pres.addSlide();
        s3.background = { color: C.offwhite };
        addSlideHeader(s3, pres, `Competitor Mentions vs. ${data.brandName}`, data.brandName);
        addFooter(s3, pres, data.brandName, data.domain);
        const comps = data.competitorMentions;
        const maxPct = Math.max(...comps.map(c => c.percentage), 1);
        const rowH = 0.38, startY = 1.35, barMaxW = 6.5;
        comps.forEach((comp, i) => {
            const y = startY + i * rowH;
            const isC = comp.name === data.brandName;
            s3.addShape(pres.shapes.OVAL, { x:0.25, y:y+0.04, w:0.28, h:0.28, fill:{color:isC?C.purple:C.lightgray}, line:{color:isC?C.purple:C.lightgray} });
            s3.addText(comp.name[0].toUpperCase(), { x:0.25, y:y+0.04, w:0.28, h:0.28, fontSize:9, bold:true, color:isC?C.white:C.navy, align:"center", valign:"middle", fontFace:"Calibri", margin:0 });
            s3.addText(comp.name, { x:0.6, y:y+0.06, w:1.4, h:0.25, fontSize:10, bold:isC, color:isC?C.purple:C.navy, fontFace:"Calibri", margin:0 });
            const barW = (comp.percentage / maxPct) * barMaxW;
            s3.addShape(pres.shapes.RECTANGLE, { x:2.1, y:y+0.06, w:Math.max(barW,0.05), h:0.25, fill:{color:isC?C.navy:C.lightgray}, line:{color:isC?C.navy:C.lightgray} });
            s3.addText(`${comp.percentage}% ${comp.mentions.toLocaleString()} mentions`, { x:2.15+barW, y:y+0.06, w:2.5, h:0.25, fontSize:9, bold:isC, color:isC?C.navy:C.slate, fontFace:"Calibri", margin:0 });
        });
    }

    // ── Slide 4: AI Platform Breakdown ────────────────────────────────────────────
    {
        const s4 = pres.addSlide();
        s4.background = { color: C.offwhite };
        addSlideHeader(s4, pres, `${data.brandName} Mentions by AI Platform`, data.brandName);
        addFooter(s4, pres, data.brandName, data.domain);
        const sumStats = [
            { label:"Total Brand Mentions",  value: data.platformData.totalMentions.toLocaleString() },
            { label:"Total Domain Citations", value: data.platformData.totalCitations.toLocaleString() },
            { label:"Avg Brand Coverage",    value: data.platformData.avgBrandCoverage },
            { label:"Avg Domain Coverage",   value: data.platformData.avgDomainCoverage },
        ];
        sumStats.forEach((st, i) => {
            const x = 0.25 + i * 2.45;
            s4.addShape(pres.shapes.RECTANGLE, { x, y:1.35, w:2.3, h:0.75, fill:{color:C.white}, line:{color:C.lightgray}, shadow:makeShadow() });
            s4.addText(st.value, { x, y:1.4,  w:2.3, h:0.38, fontSize:18, bold:true, color:C.navy, align:"center", fontFace:"Calibri", margin:0 });
            s4.addText(st.label, { x, y:1.75, w:2.3, h:0.28, fontSize:8,  color:C.slate, align:"center", fontFace:"Calibri", margin:0 });
        });
        const platforms = data.platformData.platforms;
        const colX = [0.25, 2.1, 3.3, 4.5, 7.2];
        const colW = [1.8,  1.1, 1.1, 2.6, 2.6];
        const headers = ["Platform","Mentions","Citations","Brand Visibility","Domain Coverage"];
        s4.addShape(pres.shapes.RECTANGLE, { x:0.25, y:2.35, w:9.5, h:0.3, fill:{color:C.lightgray}, line:{color:C.lightgray} });
        headers.forEach((h, i) => s4.addText(h, { x:colX[i], y:2.37, w:colW[i], h:0.26, fontSize:8, bold:true, color:C.slate, fontFace:"Calibri", margin:0 }));
        platforms.forEach((p, i) => {
            const y = 2.7 + i * 0.52;
            if (i % 2 === 0) s4.addShape(pres.shapes.RECTANGLE, { x:0.25, y:y-0.04, w:9.5, h:0.5, fill:{color:"F2F1FB"}, line:{color:"F2F1FB"} });
            s4.addText(p.name, { x:colX[0], y, w:colW[0], h:0.3, fontSize:10, bold:true, color:C.navy, fontFace:"Calibri", margin:0 });
            s4.addText(p.mentions.toLocaleString(), { x:colX[1], y, w:colW[1], h:0.3, fontSize:10, color:C.navy, fontFace:"Calibri", margin:0 });
            s4.addText(p.citations.toLocaleString(), { x:colX[2], y, w:colW[2], h:0.3, fontSize:10, color:C.navy, fontFace:"Calibri", margin:0 });
            const bvW = (p.brandVisibility / 100) * 2.3;
            s4.addShape(pres.shapes.RECTANGLE, { x:colX[3],     y:y+0.06, w:Math.max(bvW,0.05), h:0.18, fill:{color:C.navy},      line:{color:C.navy}      });
            s4.addShape(pres.shapes.RECTANGLE, { x:colX[3]+bvW, y:y+0.06, w:2.3-bvW,            h:0.18, fill:{color:C.lightgray}, line:{color:C.lightgray} });
            s4.addText(`${p.brandVisibility}%`, { x:colX[3], y, w:0.5, h:0.28, fontSize:8, bold:true, color:C.navy,   fontFace:"Calibri", margin:0 });
            const dcW = (p.domainCoverage / 100) * 2.3;
            s4.addShape(pres.shapes.RECTANGLE, { x:colX[4],     y:y+0.06, w:Math.max(dcW,0.05), h:0.18, fill:{color:C.violet},    line:{color:C.violet}    });
            s4.addShape(pres.shapes.RECTANGLE, { x:colX[4]+dcW, y:y+0.06, w:2.3-dcW,            h:0.18, fill:{color:C.lightgray}, line:{color:C.lightgray} });
            s4.addText(`${p.domainCoverage}%`, { x:colX[4], y, w:0.5, h:0.28, fontSize:8, bold:true, color:C.violet, fontFace:"Calibri", margin:0 });
        });
    }

    // ── Slide 5: Key Insights ─────────────────────────────────────────────────────
    {
        const s5 = pres.addSlide();
        s5.background = { color: C.navy };
        addFooter(s5, pres, data.brandName, data.domain);
        s5.addText("KEY INSIGHTS", { x:0.5, y:0.3, w:9, h:0.35, fontSize:10, color:C.orange, bold:true, charSpacing:4, fontFace:"Calibri", margin:0 });
        s5.addText(`${data.brandName} — GEO Audit Summary`, { x:0.5, y:0.65, w:9, h:0.5, fontSize:22, color:C.white, bold:true, fontFace:"Calibri", margin:0 });
        const positions = [
            { x:0.25, y:1.3 }, { x:5.05, y:1.3 },
            { x:0.25, y:2.9 }, { x:5.05, y:2.9 },
        ];
        data.keyInsights.slice(0,4).forEach((insight, i) => {
            const pos = positions[i], cardW = 4.55, cardH = 1.45;
            s5.addShape(pres.shapes.RECTANGLE, { x:pos.x, y:pos.y, w:cardW, h:cardH, fill:{color:C.darkgray}, line:{color:C.darkgray}, shadow:makeShadow() });
            s5.addShape(pres.shapes.RECTANGLE, { x:pos.x, y:pos.y, w:0.06, h:cardH, fill:{color:C.orange}, line:{color:C.orange} });
            s5.addText(insight.label.toUpperCase(), { x:pos.x+0.15, y:pos.y+0.12, w:cardW-0.2, h:0.22, fontSize:7,  color:C.orange, bold:true, charSpacing:2, fontFace:"Calibri", margin:0 });
            s5.addText(insight.stat,               { x:pos.x+0.15, y:pos.y+0.32, w:cardW-0.2, h:0.5,  fontSize:24, bold:true,  color:C.white,  fontFace:"Calibri", margin:0 });
            s5.addText(insight.description,        { x:pos.x+0.15, y:pos.y+0.82, w:cardW-0.2, h:0.55, fontSize:9,  color:C.lilac,  fontFace:"Calibri", margin:0, wrap:true });
        });
    }

    pres.writeFile({ fileName: outputPath });
    console.log(`✅ PPTX written: ${outputPath}`);
}

// ─── API: POST /generate ──────────────────────────────────────────────────────────
app.post("/generate", async (req, res) => {
    const { url } = req.body;
    if (!url || !url.includes("atlas.pepper.inc")) {
        return res.status(400).json({ error: "Please provide a valid Atlas report URL." });
    }
    const id      = uuidv4();
    const pptxOut = path.join(TMP, `${id}.pptx`);
    const pdfOut  = path.join(TMP, `${id}.pdf`);
    try {
        const raw  = await scrapeAtlasReport(url);
        const data = normalizeData(raw);
        buildPPTX(data, pptxOut);
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
        res.status(500).json({ error: err.message || "Generation failed. Check the URL and try again." });
    }
});

// ─── Health check ──────────────────────────────────────────────────────────────────
app.get("/health", (_, res) => res.json({ status: "ok" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Atlas PDF Generator running on port ${PORT}`));
