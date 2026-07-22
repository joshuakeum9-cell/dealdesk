/* ============ DealDesk document generators ============
   Turns engine output into real .docx / .pptx / .xlsx files,
   entirely in the browser.

   Design system: synthesized "top tier consulting" style.
   Georgia (headings) + Arial (body) is the documented substitute
   standard at both McKinsey and PwC; deep ink + one electric-blue
   accent; summary first structure; action titles on slides. */

/* Theme layer: structure and logic never change; visual identity does.
   Ghost mode (default) = neutral firm style with DRAFT stamps.
   Branded mode = client colors, fonts, and logo. */

const DEFAULT_THEME = {
  ink: "0A1F2E",
  body: "1A1A1A",
  accent: "1F4CFF",
  accent2: "00A0B0",
  gray: "63666A",
  grayLight: "A6A6A6",
  rule: "D9D9D9",
  fill: "F2F5F7",
  fillAccent: "E8EDFF",
  good: "2E7D32",
  warn: "FFB600",
  bad: "C62828",
  onInk: "FFFFFF",
  onAccent: "FFFFFF",
  serif: "Georgia",
  sans: "Arial",
  logo: null, // {b64, mime, w, h} natural pixel dims
  branded: false,
};

let THEME = { ...DEFAULT_THEME };

function cleanHex(h) {
  const m = String(h || "").replace("#", "").trim();
  return /^[0-9a-f]{6}$/i.test(m) ? m.toUpperCase() : null;
}

function tint(hex, f) {
  const n = parseInt(hex, 16);
  const mix = (c) => Math.round(c + (255 - c) * f);
  const r = mix((n >> 16) & 255), g = mix((n >> 8) & 255), b = mix(n & 255);
  return ((r << 16) | (g << 8) | b).toString(16).padStart(6, "0").toUpperCase();
}

function luminance(hex) {
  const n = parseInt(hex, 16);
  return 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255);
}

function setTheme(brand) {
  if (!brand) {
    THEME = { ...DEFAULT_THEME };
    return;
  }
  const ink = cleanHex(brand.ink) || DEFAULT_THEME.ink;
  const accent = cleanHex(brand.accent) || DEFAULT_THEME.accent;
  THEME = {
    ...DEFAULT_THEME,
    ink,
    accent,
    serif: brand.serif || DEFAULT_THEME.serif,
    sans: brand.sans || DEFAULT_THEME.sans,
    logo: brand.logo || null,
    fill: tint(ink, 0.95),
    fillAccent: tint(accent, 0.88),
    onInk: luminance(ink) < 140 ? "FFFFFF" : "1A1A1A",
    onAccent: luminance(accent) < 140 ? "FFFFFF" : "1A1A1A",
    branded: true,
  };
}

function slug(name) {
  return (name || "Target").replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "");
}

function dealCode(deal) {
  const w = (deal.company || "Target").split(/\s+/)[0].replace(/[^a-z]/gi, "");
  return "PROJECT " + (w || "TARGET").toUpperCase();
}

function todayLabel() {
  return new Date().toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

function fyLabel(year, kind) {
  return "FY" + String(year).slice(2) + (kind || "A");
}

function fmtNum(v) {
  if (v === null || v === undefined) return "n/a";
  return v < 0 ? "(" + Math.abs(v).toFixed(1) + ")" : v.toFixed(1);
}

function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

/* ================================================================
   DOCX building blocks
   ================================================================ */

function run(text, opts = {}) {
  return new docx.TextRun({
    text,
    font: opts.serif ? THEME.serif : THEME.sans,
    size: opts.size || 21, // half-points; 21 = 10.5pt
    bold: opts.bold || false,
    italics: opts.italics || false,
    color: opts.color || THEME.body,
  });
}

function para(children, opts = {}) {
  return new docx.Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { line: 276, after: opts.after !== undefined ? opts.after : 120, before: opts.before || 0 },
    bullet: opts.bullet ? { level: 0 } : undefined,
    alignment: opts.align,
    border: opts.borderBottom
      ? { bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: THEME.rule, space: 4 } }
      : undefined,
    tabStops: opts.tabStops,
  });
}

function bodyPara(text, opts = {}) {
  return para(run(text, opts.runOpts || {}), opts);
}

function h1(text) {
  return para(run(text, { serif: true, size: 32, bold: true, color: THEME.ink }), {
    before: 360,
    after: 160,
    borderBottom: true,
  });
}

function h2(text) {
  return para(run(text, { serif: true, size: 26, bold: true, color: THEME.ink }), {
    before: 240,
    after: 100,
  });
}

function sourceLine(text) {
  return para(run(text, { size: 16, color: THEME.gray }), { before: 60, after: 240 });
}

function findingPara(i, lead, text) {
  return para(
    [run(`${i}. ${lead}: `, { bold: true, color: THEME.ink }), run(text)],
    { after: 140 }
  );
}

// Cover-style title block: eyebrow code name, serif title, subtitle, rule
function titleBlock(deal, docTitle, subtitle) {
  return [
    para(run(dealCode(deal), { size: 22, bold: true, color: THEME.accent }), { after: 80 }),
    para(run(docTitle, { serif: true, size: 56, color: THEME.ink }), { after: 100 }),
    para(run(subtitle, { size: 21, color: THEME.gray }), { after: 60, borderBottom: true }),
  ];
}

// Executive summary callout box: single-cell table, accent left bar, tint fill
function summaryBox(children) {
  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.SINGLE, size: 2, color: THEME.rule },
      bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: THEME.rule },
      right: { style: docx.BorderStyle.SINGLE, size: 2, color: THEME.rule },
      left: { style: docx.BorderStyle.SINGLE, size: 24, color: THEME.accent },
      insideHorizontal: { style: docx.BorderStyle.NONE },
      insideVertical: { style: docx.BorderStyle.NONE },
    },
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            shading: { fill: THEME.fill },
            margins: { top: 160, bottom: 160, left: 240, right: 240 },
            children,
          }),
        ],
      }),
    ],
  });
}

function docHeader(deal) {
  const right =
    THEME.branded && THEME.logo
      ? new docx.ImageRun({
          data: THEME.logo.b64,
          transformation: {
            height: 22,
            width: Math.min(Math.round(22 * (THEME.logo.w / THEME.logo.h)), 140),
          },
        })
      : run("Private and confidential", { size: 16, color: THEME.gray });
  return new docx.Header({
    children: [
      para(
        [
          run(dealCode(deal), { size: 16, color: THEME.gray }),
          new docx.TextRun({ text: "\t" }),
          right,
        ],
        {
          after: 0,
          borderBottom: true,
          tabStops: [{ type: docx.TabStopType.RIGHT, position: 9360 }],
        }
      ),
    ],
  });
}

function docFooter() {
  return new docx.Footer({
    children: [
      new docx.Paragraph({
        tabStops: [{ type: docx.TabStopType.RIGHT, position: 9360 }],
        spacing: { before: 60, after: 0 },
        children: [
          run("Draft for discussion purposes only", { size: 16, color: THEME.gray }),
          new docx.TextRun({ text: "\t" }),
          new docx.TextRun({
            font: THEME.sans,
            size: 16,
            color: THEME.gray,
            children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
          }),
        ],
      }),
    ],
  });
}

async function packDocx(deal, children) {
  const doc = new docx.Document({
    styles: { default: { document: { run: { font: THEME.sans, size: 21, color: THEME.body } } } },
    sections: [
      {
        properties: {
          page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
        },
        headers: { default: docHeader(deal) },
        footers: { default: docFooter() },
        children,
      },
    ],
  });
  return docx.Packer.toBlob(doc);
}

/* ================================================================
   1. Business Summary (.docx)
   ================================================================ */

async function generateSummaryDocx(deal, m) {
  const P = practiceOf(deal);
  const n = buildNarrative(deal, m);
  const findings = buildKeyFindings(deal, m);
  const sample = m.source === "sample";

  const srcText = sample
    ? "Source: Illustrative sample financials; DealDesk analysis"
    : "Source: Management information; DealDesk analysis";

  // Section order follows the PwC business summary template:
  // company overview, financial summary, recent news, key people,
  // key products, other. High level by design: orient the team fast.
  const children = [
    ...titleBlock(
      deal,
      `${deal.company}: Business summary`,
      `${deal.dealType} | ${deal.industry || "Industry not specified"} | ${todayLabel()}` +
        (sample ? " | Illustrative sample financials" : "")
    ),

    h1("1. Company overview"),
    bodyPara(n.dealFrame),
    bodyPara(
      `${deal.company} operates in the ${(deal.industry || "target").toLowerCase()} sector with reported revenue of ${fmtM(m.revenueLatest)} in ${fyLabel(m.lastYear)}. ${P.overview(deal, m)}`
    ),
    ...situationSnapshotParas(deal),
    summaryBox([
      para(run("What matters most", { serif: true, size: 24, bold: true, color: THEME.ink }), { after: 120 }),
      ...findings.map((f, i) => findingPara(i + 1, f.lead, f.text)),
    ]),

    h1("2. Financial summary"),
    bodyPara(n.growthSentence),
    bodyPara(n.marginSentence),
    financialsTable(m),
    sourceLine(srcText),

    h1("3. Recent news"),
    bodyPara("Recent company news relevant to the engagement objective, with links to sources:", {}),
    bodyPara("[Headline 1: from the company newsroom or investor relations page]", { bullet: true }),
    bodyPara("[Headline 2: from sector press or analyst coverage]", { bullet: true }),

    h1("4. Key people"),
    peopleTable(),
    sourceLine("Source: Company website; investor relations"),

    h1("5. Key products"),
    productsTable(),
    sourceLine("Source: Company financial statements; segment reporting where available"),

    ...analystNotesSection(deal),

    h1(analystNotesSection(deal).length ? "7. Other" : "6. Other"),
    bodyPara("The most important remaining topics to orient the team, prioritized for this engagement:", {}),
    (() => {
      const risk = situationRisk(deal, m);
      return bodyPara(`${risk.lead}: ${risk.text}`, { bullet: true });
    })(),
    bodyPara("Customer and supplier concentration: confirm the share of revenue tied to the top five relationships.", { bullet: true }),
    bodyPara("Next step: management interviews using the accompanying interview guide.", { bullet: true }),
    para(run("Prepared with DealDesk. High level by design; figures may not sum due to rounding.", { size: 16, color: THEME.gray }), { before: 300 }),
  ];
  return packDocx(deal, children);
}

function placeholderTable(header, rows) {
  const cell = (text, isHeader) =>
    new docx.TableCell({
      shading: isHeader ? { fill: THEME.ink } : undefined,
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [
        new docx.Paragraph({
          spacing: { after: 0 },
          children: [run(text, { size: 19, bold: isHeader, color: isHeader ? THEME.onInk : THEME.body })],
        }),
      ],
    });
  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.SINGLE, size: 8, color: THEME.ink },
      bottom: { style: docx.BorderStyle.SINGLE, size: 8, color: THEME.ink },
      left: { style: docx.BorderStyle.NONE },
      right: { style: docx.BorderStyle.NONE },
      insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 4, color: THEME.rule },
      insideVertical: { style: docx.BorderStyle.NONE },
    },
    rows: [
      new docx.TableRow({ children: header.map((h) => cell(h, true)) }),
      ...rows.map((r) => new docx.TableRow({ children: r.map((c) => cell(c, false)) })),
    ],
  });
}

// One line summary of the structured situation answers, when any exist
function situationSnapshotParas(deal) {
  const s = situationOf(deal);
  const bits = [];
  if (s.goal) bits.push(`stated goal is ${s.goal.toLowerCase()}`);
  if (s.marginTrend) bits.push(`margin ${s.marginTrend.toLowerCase()}`);
  if (s.revenueDriver === "Price") bits.push("growth driven by price and mix");
  else if (s.revenueDriver === "Volume") bits.push("growth driven by volume");
  else if (s.revenueDriver === "Declining") bits.push("demand softening");
  if (s.concentration === "Yes") bits.push("significant customer concentration");
  if (s.capacity === "Excess") bits.push("excess capacity in the network");
  else if (s.capacity === "Tight") bits.push("capacity running tight");
  if (s.costProgram === "Yes") bits.push("cost program underway");
  if (s.competition === "Intensifying") bits.push("competition intensifying");
  if (!bits.length) return [];
  return [bodyPara("Situation assessment: " + bits.join("; ") + ".")];
}

// Long pasted context becomes its own clean section instead of being
// squeezed into a framing sentence. House style: no dashes.
function analystNotesSection(deal) {
  const ctx = (deal.context || "").trim();
  if (ctx.length <= 160) return [];
  const softened = ctx
    .replace(/(\d)\s*[-–—]\s*(\d)/g, "$1 to $2")
    .replace(/[—–]/g, ", ")
    .replace(/-/g, " ");
  const paras = softened.split(/\n+/).map((p) => p.trim()).filter(Boolean);
  return [
    h1("6. Analyst notes"),
    bodyPara("Notes provided with the engagement request, reproduced for the team:", {}),
    ...paras.map((p) => bodyPara(p)),
  ];
}

function peopleTable() {
  return placeholderTable(
    ["Role", "Name"],
    [
      ["Chief Executive Officer", "[Name]"],
      ["Chief Financial Officer", "[Name]"],
      ["Chief Operating Officer", "[Name]"],
      ["[Other key role]", "[Name]"],
    ]
  );
}

function productsTable() {
  return placeholderTable(
    ["Name", "Revenue", "Description"],
    [
      ["[Product or segment]", "[$M]", "[Brief description]"],
      ["[Product or segment]", "[$M]", "[Brief description]"],
      ["[Product or segment]", "[$M]", "[Brief description]"],
    ]
  );
}

function financialsTable(m) {
  // PwC template: historical years plus a projection column, with an
  // EBITDA margin % row. The projection follows the ACTUAL trend, even
  // when it is negative; growing a declining business would contradict
  // the narrative. Large companies switch to $B units.
  const g = Math.max(-0.5, Math.min(0.5, m.revenueCAGR));
  const last = m.series.revenue.length - 1;
  const div = m.revenueLatest >= 10000 ? 1000 : 1;
  const unit = div === 1000 ? "$B" : "$M";
  const proj = (series) => series[last] * (1 + g);
  const fN = (v) => fmtNum(v / div);

  const header = [unit, ...m.years.map((y) => fyLabel(y)), fyLabel(m.lastYear + 1, "F")];
  const rows = [["Revenue", ...m.series.revenue.map(fN), fN(proj(m.series.revenue))]];
  if (m.series.cogs) rows.push(["COGS", ...m.series.cogs.map((v) => fN(-v)), fN(-proj(m.series.cogs))]);
  if (m.series.opex) rows.push(["SG&A / Opex", ...m.series.opex.map((v) => fN(-v)), fN(-proj(m.series.opex))]);
  const totalRow = m.ebitdaSeries
    ? ["EBITDA", ...m.ebitdaSeries.map(fN), fN(proj(m.ebitdaSeries))]
    : null;
  const marginRow = m.ebitdaSeries
    ? [
        "EBITDA margin (%)",
        ...m.ebitdaSeries.map((e, i) => fmtPct(e / m.series.revenue[i])),
        fmtPct(proj(m.ebitdaSeries) / proj(m.series.revenue)),
      ]
    : null;

  const cell = (text, { headerRow, total, first } = {}) =>
    new docx.TableCell({
      shading: headerRow ? { fill: THEME.ink } : total ? { fill: THEME.fill } : undefined,
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [
        new docx.Paragraph({
          alignment: first ? docx.AlignmentType.LEFT : docx.AlignmentType.RIGHT,
          spacing: { after: 0 },
          children: [
            run(text, {
              size: 19,
              bold: headerRow || total,
              color: headerRow ? THEME.onInk : THEME.body,
            }),
          ],
        }),
      ],
    });

  const mkRow = (cells, opts = {}) =>
    new docx.TableRow({
      children: cells.map((c, i) => cell(c, { ...opts, first: i === 0 })),
    });

  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.SINGLE, size: 8, color: THEME.ink },
      bottom: { style: docx.BorderStyle.SINGLE, size: 8, color: THEME.ink },
      left: { style: docx.BorderStyle.NONE },
      right: { style: docx.BorderStyle.NONE },
      insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 4, color: THEME.rule },
      insideVertical: { style: docx.BorderStyle.NONE },
    },
    rows: [
      mkRow(header, { headerRow: true }),
      ...rows.map((r) => mkRow(r)),
      ...(totalRow ? [mkRow(totalRow, { total: true })] : []),
      ...(marginRow ? [mkRow(marginRow)] : []),
    ],
  });
}

/* ================================================================
   2. Interview Guide (.docx)
   ================================================================ */

async function generateGuideDocx(deal, m) {
  const sections = buildInterviewQuestions(deal);
  const total = sections.reduce((s, sec) => s + sec.questions.length, 0);

  // PwC interview guide template: interviewee header block, topics
  // lettered, questions numbered continuously (Q1..Qn), space for answers.
  const children = [
    ...titleBlock(
      deal,
      `${deal.company}: Interview guide`,
      `${deal.dealType} | ${total} questions | Private and confidential`
    ),
    para([run("Interviewee: ", { bold: true }), run("[Name]")], { after: 60 }),
    para([run("Position, company: ", { bold: true }), run(`[Title], ${deal.company}`)], { after: 60 }),
    para([run("Date: ", { bold: true }), run("[Date]")], { after: 200, borderBottom: true }),
    summaryBox([
      para(
        [
          run("Purpose: ", { bold: true, color: THEME.ink }),
          run(
            "structure the conversation and stay focused on the key questions we need answered. Questions are ordered for maximum impact; capture answers in the space provided and flag follow ups for the tracker."
          ),
        ],
        { after: 0 }
      ),
    ]),
    para(run(""), { after: 120 }),
  ];

  let qNum = 0;
  sections.forEach((s, i) => {
    children.push(h1(`Topic ${String.fromCharCode(65 + i)}: ${s.section}`));
    s.questions.forEach((q) => {
      qNum += 1;
      children.push(para([run(`Q${qNum}. `, { bold: true, color: THEME.accent }), run(q)], { after: 60 }));
      children.push(para(run("Answer:", { size: 18, italics: true, color: THEME.gray }), { after: 260 }));
    });
  });
  return packDocx(deal, children);
}

/* ================================================================
   3. Synergy Presentation (.pptx)
   ================================================================ */

const SLIDE_W = 13.33;
const MARGIN = 0.45;

function addFurniture(slide, { kicker, title, source, takeaway }) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: MARGIN, y: 0.18, w: 9, h: 0.28,
      fontFace: THEME.sans, fontSize: 10, color: THEME.accent, bold: true, charSpacing: 2,
    });
  }
  if (THEME.branded && THEME.logo) {
    // Client branded final: logo replaces the draft sticker (the "ghost deck" flips)
    const h = 0.32;
    const w = Math.min(h * (THEME.logo.w / THEME.logo.h), 1.8);
    slide.addImage({
      data: `${THEME.logo.mime};base64,${THEME.logo.b64}`,
      x: SLIDE_W - MARGIN - w, y: 0.16, w, h,
    });
  } else {
    slide.addText("DRAFT | FOR DISCUSSION", {
      x: 10.4, y: 0.18, w: 2.5, h: 0.28, align: "right",
      fontFace: THEME.sans, fontSize: 9, color: THEME.grayLight,
    });
  }
  // Action title: full-sentence takeaway, fixed position, largest text on slide
  slide.addText(title, {
    x: MARGIN, y: 0.5, w: SLIDE_W - 2 * MARGIN, h: 0.85,
    fontFace: THEME.serif, fontSize: 19, color: THEME.ink, bold: false, valign: "top",
  });
  if (takeaway) {
    slide.addShape("rect", { x: MARGIN, y: 6.42, w: SLIDE_W - 2 * MARGIN, h: 0.52, fill: { color: THEME.fill }, line: { type: "none" } });
    slide.addShape("rect", { x: MARGIN, y: 6.42, w: 0.06, h: 0.52, fill: { color: THEME.accent }, line: { type: "none" } });
    slide.addText(takeaway, {
      x: MARGIN + 0.18, y: 6.42, w: SLIDE_W - 2 * MARGIN - 0.3, h: 0.52,
      fontFace: THEME.sans, fontSize: 12.5, bold: true, color: THEME.ink, valign: "middle",
    });
  }
  if (source) {
    slide.addText(source, {
      x: MARGIN, y: 7.02, w: 9, h: 0.25,
      fontFace: THEME.sans, fontSize: 8, color: THEME.gray,
    });
  }
}

async function generateSynergyPptx(deal, m) {
  const P = practiceOf(deal);
  const groups = buildOpportunities(deal, m).groups;
  const all = groups.flatMap((g, gi) => g.items.map((s) => ({ ...s, kind: g.label, gi })));
  const score = { High: 3, Medium: 2, Low: 1 };
  const srcLine = m.source === "sample"
    ? "Source: Illustrative sample financials; DealDesk analysis"
    : "Source: Management financials; DealDesk analysis";

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.defineSlideMaster({
    title: "CONTENT",
    background: { color: "FFFFFF" },
    objects: [
      { text: { text: `${dealCode(deal)} | Private and confidential`, options: { x: MARGIN, y: 7.24, w: 6, h: 0.22, fontFace: THEME.sans, fontSize: 8, color: THEME.grayLight } } },
    ],
    slideNumber: { x: 12.55, y: 7.22, fontFace: THEME.sans, fontSize: 9, color: THEME.gray },
  });

  /* --- Title slide: full-bleed ink, white serif title --- */
  let s = pptx.addSlide();
  s.background = { color: THEME.ink };
  s.addText(dealCode(deal), { x: 0.6, y: 2.0, w: 11, h: 0.4, fontFace: THEME.sans, fontSize: 13, bold: true, color: THEME.accent, charSpacing: 3 });
  s.addText(P.deckTitle.charAt(0).toUpperCase() + P.deckTitle.slice(1), { x: 0.6, y: 2.45, w: 12, h: 1.6, fontFace: THEME.serif, fontSize: 38, color: THEME.onInk });
  s.addText(`${deal.company} | ${deal.dealType} | ${todayLabel()}`, { x: 0.6, y: 4.1, w: 11, h: 0.4, fontFace: THEME.sans, fontSize: 13, color: "B8C2CC" });
  s.addText("Private and confidential. Draft for discussion purposes only." + (m.source === "sample" ? " Prepared on illustrative sample financials." : ""), {
    x: 0.6, y: 6.9, w: 11, h: 0.3, fontFace: THEME.sans, fontSize: 9, color: "8A97A3",
  });

  /* --- Slide 1: summary that signposts slides 2 and 3 (PwC 3-slide loop) --- */
  s = pptx.addSlide({ masterName: "CONTENT" });
  const topSyn = [...all].sort((a, b) => score[b.impact] + score[b.ease] - (score[a.impact] + score[a.ease]))[0];
  const n = buildNarrative(deal, m);
  addFurniture(s, {
    kicker: "Summary",
    title: `${deal.company} presents ${groups[0].items.length} ${groups[0].label.toLowerCase()} and ${groups[1].items.length} ${groups[1].label.toLowerCase()}; ${topSyn.name.toLowerCase()} offers the highest confidence value`,
    source: srcLine,
    takeaway: "The next two slides expand each category in the order shown here",
  });
  s.addText(n.dealFrame, {
    x: MARGIN, y: 1.45, w: SLIDE_W - 2 * MARGIN, h: 0.65,
    fontFace: THEME.sans, fontSize: 12.5, color: "404850", valign: "top",
  });
  groups.forEach((g, gi) => {
    const x = MARGIN + gi * 6.3;
    s.addText(`${gi + 1}. ${g.label.toUpperCase()}`, {
      x, y: 2.25, w: 6.1, h: 0.35,
      fontFace: THEME.sans, fontSize: 12, bold: true, color: THEME.accent, charSpacing: 1,
    });
    s.addText(
      g.items.map((it) => ({
        text: `${all.findIndex((a) => a.name === it.name) + 1}. ${it.name}`,
        options: { breakLine: true, color: THEME.ink, bold: false },
      })),
      { x, y: 2.65, w: 6.1, h: 2.4, fontFace: THEME.sans, fontSize: 13, lineSpacing: 26, valign: "top" }
    );
    s.addText(`Expanded on slide ${gi + 2}`, {
      x, y: 5.35, w: 6.1, h: 0.3, fontFace: THEME.sans, fontSize: 10, italic: true, color: THEME.gray,
    });
  });

  /* --- Slides 2 and 3: expand each category, same order as slide 1 --- */
  groups.forEach((g, gi) => {
    const sl = pptx.addSlide({ masterName: "CONTENT" });
    addFurniture(sl, {
      kicker: `${gi + 1}. ${g.label}`,
      title: g.slideTitle,
      source: srcLine,
      takeaway: g.slideTakeaway,
    });
    g.items.forEach((it, i) => {
      const y = 1.5 + i * 1.55;
      const num = all.findIndex((a) => a.name === it.name) + 1;
      sl.addText(String(num), {
        shape: "ellipse", x: MARGIN, y: y + 0.04, w: 0.42, h: 0.42,
        fill: { color: gi === 0 ? THEME.accent2 : THEME.accent },
        align: "center", valign: "middle", fontFace: THEME.sans, fontSize: 13, bold: true, color: THEME.onAccent,
      });
      sl.addText(
        [
          { text: it.name, options: { bold: true, color: THEME.ink, fontSize: 15 } },
          { text: `   Impact: ${it.impact}  |  Ease: ${it.ease}`, options: { color: THEME.gray, fontSize: 10.5 } },
        ],
        { x: 1.05, y, w: 11.7, h: 0.42, fontFace: THEME.sans, valign: "middle" }
      );
      sl.addText(it.rationale, {
        x: 1.05, y: y + 0.44, w: 11.7, h: 0.95,
        fontFace: THEME.sans, fontSize: 12.5, color: "404850", valign: "top", lineSpacing: 18,
      });
    });
  });

  /* --- Achievability 2x2 matrix --- */
  s = pptx.addSlide({ masterName: "CONTENT" });
  const ranked = [...all].sort((a, b) => score[b.impact] + score[b.ease] - (score[a.impact] + score[a.ease]));
  addFurniture(s, {
    kicker: "Appendix: prioritization",
    title: `Pursue the top ${Math.min(3, ranked.length)} opportunities first; hold the remainder for the roadmap`,
    source: srcLine,
  });
  const px = 1.1, py = 1.55, pw = 7.0, ph = 4.5;
  // Quadrant fills: highlight top-right (do first)
  s.addShape("rect", { x: px, y: py, w: pw, h: ph, fill: { color: "FAFBFC" }, line: { color: THEME.rule, pt: 0.75 } });
  s.addShape("rect", { x: px + pw / 2, y: py, w: pw / 2, h: ph / 2, fill: { color: THEME.fillAccent }, line: { type: "none" } });
  s.addShape("line", { x: px + pw / 2, y: py, w: 0, h: ph, line: { color: THEME.rule, pt: 0.75, dashType: "dash" } });
  s.addShape("line", { x: px, y: py + ph / 2, w: pw, h: 0, line: { color: THEME.rule, pt: 0.75, dashType: "dash" } });
  s.addText("Do first", { x: px + pw / 2, y: py + 0.05, w: pw / 2 - 0.1, h: 0.3, align: "right", fontFace: THEME.sans, fontSize: 10, bold: true, color: THEME.accent });
  s.addText("Value impact", { x: px - 0.95, y: py + ph / 2 - 0.2, w: 1.7, h: 0.4, rotate: 270, fontFace: THEME.sans, fontSize: 11, color: THEME.gray, align: "center" });
  s.addText("Ease of capture", { x: px, y: py + ph + 0.1, w: pw, h: 0.3, align: "center", fontFace: THEME.sans, fontSize: 11, color: THEME.gray });
  const frac = { High: 0.8, Medium: 0.5, Low: 0.2 };
  const seen = {};
  all.forEach((it, i) => {
    const key = it.impact + it.ease;
    const bump = (seen[key] = (seen[key] || 0) + 1) - 1;
    const cx = px + frac[it.ease] * pw + bump * 0.55 - 0.21;
    const cy = py + (1 - frac[it.impact]) * ph - 0.21;
    s.addText(String(i + 1), {
      shape: "ellipse", x: cx, y: cy, w: 0.42, h: 0.42,
      fill: { color: it.gi === 0 ? THEME.accent2 : THEME.accent },
      align: "center", valign: "middle", fontFace: THEME.sans, fontSize: 12, bold: true, color: THEME.onAccent,
    });
  });
  // Legend list, right column
  s.addText("Opportunities", { x: 8.55, y: 1.55, w: 4.3, h: 0.3, fontFace: THEME.sans, fontSize: 11, bold: true, color: THEME.ink });
  s.addText(
    all.map((it, i) => ({
      text: `${i + 1}. ${it.name}`,
      options: { breakLine: true, color: "404850" },
    })),
    { x: 8.55, y: 1.9, w: 4.3, h: 4.2, fontFace: THEME.sans, fontSize: 11, lineSpacing: 20, valign: "top" }
  );
  s.addText([
    { text: "● ", options: { color: THEME.accent2 } }, { text: groups[0].label + "   ", options: { color: THEME.gray } },
    { text: "● ", options: { color: THEME.accent } }, { text: groups[1].label, options: { color: THEME.gray } },
  ], { x: 8.55, y: 5.7, w: 4, h: 0.3, fontFace: THEME.sans, fontSize: 10 });

  return pptx.write({ outputType: "blob" });
}

/* ================================================================
   4. Excel model (.xlsx) + email (.docx)
   ================================================================ */

function generateModelXlsx(deal, m) {
  // PwC Task 4 structure: inputs on their own sheet in labeled cells, the
  // model on a second sheet as pure formulas (no hardcoding), three
  // scenarios: conservative (low end), midpoint, aggressive (high end).
  const wb = XLSX.utils.book_new();
  const term = practiceTerm(deal);
  const Term = term.charAt(0).toUpperCase() + term.slice(1);
  const sc = computeScenarios(deal, m);
  const inp = sc.inputs;
  const NUMFMT = "#,##0.0;(#,##0.0)";
  const PCT = "0.0%";

  /* --- Inputs sheet: every changeable number lives here, labeled --- */
  const aoaI = [];
  const refs = {};
  const push = (row) => aoaI.push(row);
  push([`${deal.company}: ${term} model inputs ($M)`, null]);
  push([null, null]);
  push([`Client data (target), ${fyLabel(m.lastYear)}`, null]);
  push(["Revenue", inp.revenue]);
  refs.revenue = "B" + aoaI.length;
  inp.lines.forEach((l) => {
    push([l.label, l.base]);
    refs[l.key + "Base"] = "B" + aoaI.length;
  });
  push([null, null]);
  push([`Revenue ${term} assumptions (INPUTS: change these, the model recalculates)`, null]);
  push(["Revenue uplift, low end", inp.revLow]);
  refs.revLow = "B" + aoaI.length;
  push(["Revenue uplift, high end", inp.revHigh]);
  refs.revHigh = "B" + aoaI.length;
  push(["Flow through margin on new revenue" + (inp.flowDefaulted ? " (defaulted: company margin unavailable or negative)" : ""), inp.flowMargin]);
  refs.margin = "B" + aoaI.length;
  push([null, null]);
  push([`Cost ${term} assumptions, each applied to its own cost line only (INPUTS)`, null]);
  inp.lines.forEach((l) => {
    push([`${l.label}: ${term} %, low end`, l.low]);
    refs[l.key + "Low"] = "B" + aoaI.length;
    push([`${l.label}: ${term} %, high end`, l.high]);
    refs[l.key + "High"] = "B" + aoaI.length;
  });
  push([null, null]);
  push(["Conservative = low end of every range. Aggressive = high end. Midpoint = average.", null]);

  const wsI = XLSX.utils.aoa_to_sheet(aoaI);
  [refs.revLow, refs.revHigh, refs.margin, ...inp.lines.flatMap((l) => [refs[l.key + "Low"], refs[l.key + "High"]])]
    .forEach((c) => { if (wsI[c]) wsI[c].z = PCT; });
  [refs.revenue, ...inp.lines.map((l) => refs[l.key + "Base"])].forEach((c) => { if (wsI[c]) wsI[c].z = NUMFMT; });
  wsI["!cols"] = [{ wch: 58 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsI, "Inputs");

  /* --- Model sheet: formulas only, three scenario columns --- */
  const I = (ref) => `Inputs!$${ref.slice(0, 1)}$${ref.slice(1)}`;
  const scenCols = ["B", "C", "D"];
  const scen = [sc.conservative, sc.midpoint, sc.aggressive];
  const revF = [
    `${I(refs.revenue)}*${I(refs.revLow)}*${I(refs.margin)}`,
    `${I(refs.revenue)}*AVERAGE(${I(refs.revLow)},${I(refs.revHigh)})*${I(refs.margin)}`,
    `${I(refs.revenue)}*${I(refs.revHigh)}*${I(refs.margin)}`,
  ];
  const aoaM = [
    [`${deal.company}: directional ${term} model ($M, annual EBIT impact)`, null, null, null],
    [null, "Conservative", "Midpoint", "Aggressive"],
    [`Revenue ${term} (at flow through margin)`, null, null, null],
    ...inp.lines.map((l) => [`${l.label} ${term}`, null, null, null]),
    [`Total ${term} value`, null, null, null],
    [null, null, null, null],
    ["All cells on this sheet are formulas driven by the Inputs sheet.", null, null, null],
  ];
  const wsM = XLSX.utils.aoa_to_sheet(aoaM);
  const lineRow = (i) => 4 + i; // first cost line sits on sheet row 4
  const totalRow = 4 + inp.lines.length;
  scenCols.forEach((c, si) => {
    wsM[c + "3"] = { t: "n", v: scen[si].revenueImpact, f: revF[si], z: NUMFMT };
    inp.lines.forEach((l, li) => {
      const lf = [
        `${I(refs[l.key + "Base"])}*${I(refs[l.key + "Low"])}`,
        `${I(refs[l.key + "Base"])}*AVERAGE(${I(refs[l.key + "Low"])},${I(refs[l.key + "High"])})`,
        `${I(refs[l.key + "Base"])}*${I(refs[l.key + "High"])}`,
      ][si];
      const lv = l.base * [l.low, (l.low + l.high) / 2, l.high][si];
      wsM[c + lineRow(li)] = { t: "n", v: lv, f: lf, z: NUMFMT };
    });
    wsM[c + totalRow] = {
      t: "n",
      v: scen[si].total,
      f: `SUM(${c}3:${c}${totalRow - 1})`,
      z: NUMFMT,
    };
  });
  wsM["!cols"] = [{ wch: 42 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
  wsM["!ref"] = `A1:D${aoaM.length}`;
  XLSX.utils.book_append_sheet(wb, wsM, `${Term} Model`);

  // Historicals sheet
  const hist = [
    ["Historical financials ($M)", ...m.years.map((y) => fyLabel(y))],
    ["Revenue", ...m.series.revenue],
  ];
  if (m.series.cogs) hist.push(["COGS", ...m.series.cogs]);
  if (m.series.opex) hist.push(["Opex", ...m.series.opex]);
  if (m.ebitdaSeries) hist.push(["EBITDA", ...m.ebitdaSeries]);
  const wsH = XLSX.utils.aoa_to_sheet(hist);
  Object.keys(wsH).forEach((k) => {
    if (!k.startsWith("!") && wsH[k].t === "n") wsH[k].z = NUMFMT;
  });
  wsH["!cols"] = [{ wch: 24 }, ...m.years.map(() => ({ wch: 11 }))];
  XLSX.utils.book_append_sheet(wb, wsH, "Historicals");

  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

async function generateEmailDocx(deal, m) {
  const email = buildEmail(deal, m);
  const rule = { borderBottom: true, after: 0 };
  const children = [
    para(run("DealDesk | Deal team memorandum", { size: 18, bold: true, color: THEME.gray }), { after: 160 }),
    para([run("To: ", { bold: true }), run("Engagement leadership")], rule),
    para([run("From: ", { bold: true }), run("Deal team")], rule),
    para([run("Date: ", { bold: true }), run(todayLabel())], rule),
    para([run("Re: ", { bold: true }), run(email.subject)], { borderBottom: true, after: 240 }),

    para(run(email.answer, { bold: true, color: THEME.ink }), { after: 140 }),
    bodyPara(email.soWhat),
    bodyPara(email.ask, { after: 240 }),

    ...email.points.map((p, i) => findingPara(i + 1, p.lead, p.text)),

    para(run("Next steps", { serif: true, size: 24, bold: true, color: THEME.ink }), { before: 200, after: 100 }),
    ...email.nextSteps.map((t) => bodyPara(t, { bullet: true })),
    bodyPara(email.close, { after: 0 }),
    para(run("Privileged and confidential | Draft for discussion purposes only", { size: 16, color: THEME.gray }), { before: 300 }),
  ];
  return packDocx(deal, children);
}

/* ================================================================
   Full package (.zip)
   ================================================================ */

async function generatePackageZip(deal, m) {
  const zip = new JSZip();
  const base = slug(deal.company);
  const [summary, guide, synergy, email] = await Promise.all([
    generateSummaryDocx(deal, m),
    generateGuideDocx(deal, m),
    generateSynergyPptx(deal, m),
    generateEmailDocx(deal, m),
  ]);
  zip.file(`${base}_Business_Summary.docx`, summary);
  zip.file(`${base}_Interview_Guide.docx`, guide);
  zip.file(`${base}_Synergy_Presentation.pptx`, synergy);
  zip.file(`${base}_Deal_Model.xlsx`, generateModelXlsx(deal, m));
  zip.file(`${base}_Email_Summary.docx`, email);
  return zip.generateAsync({ type: "blob" });
}
