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
    accent2: tint(accent, 0.4), // second series stays in the brand family
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
    : m.source === "lookup"
    ? "Source: Reported figures via Wikidata; DealDesk analysis"
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
    ...lookupProfileParas(deal),
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
    ...newsSection(deal),

    h1("4. Key people"),
    peopleTable(deal),
    sourceLine(
      deal.lookup && (deal.lookup.ceo || deal.lookup.chair)
        ? "Source: Wikidata; confirm against company investor relations"
        : "Source: Company website; investor relations"
    ),

    h1("5. Key products"),
    productsTable(deal),
    sourceLine("Source: Company financial statements; segment reporting where available"),

    ...analystNotesSection(deal),

    h1(analystNotesSection(deal).length ? "7. Other" : "6. Other"),
    bodyPara("The most important remaining topics to orient the team, prioritized for this engagement:", {}),
    // The priority risk already leads the findings box; this section adds
    // topics not covered there instead of repeating it.
    bodyPara(
      situationRisk(deal, m).lead === "Customer concentration"
        ? "Supplier dependence: map the top suppliers by spend and any single points of failure."
        : "Customer and supplier concentration: confirm the share of revenue tied to the top five relationships.",
      { bullet: true }
    ),
    bodyPara("Quality of reported figures: confirm accounting basis, one time items, and any restatements.", { bullet: true }),
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

// Real company facts from the lookup, rendered as one overview line
function lookupProfileParas(deal) {
  const L = deal.lookup;
  if (!L) return [];
  const out = [];
  if (L.description) out.push(bodyPara(L.description));
  const facts = [];
  if (L.founded) facts.push(`founded ${L.founded}`);
  if (L.hq) facts.push(`headquartered in ${L.hq}`);
  if (L.employees) facts.push(`about ${L.employees.toLocaleString("en-US")} employees`);
  if (L.ceo) facts.push(`led by CEO ${L.ceo}`);
  if (facts.length) {
    const line = facts.join("; ");
    out.push(bodyPara(line.charAt(0).toUpperCase() + line.slice(1) + ". Source: Wikidata."));
  }
  return out;
}

// House rule applies to user typed profile text too
function softenText(s) {
  return String(s)
    .replace(/(\d)\s*[-–—]\s*(\d)/g, "$1 to $2")
    .replace(/[—–]/g, ", ")
    .replace(/-/g, " ")
    .trim();
}

function newsSection(deal) {
  const manual = deal.profile && deal.profile.news && deal.profile.news.length ? deal.profile.news : null;
  const looked = deal.lookup && deal.lookup.news && deal.lookup.news.length ? deal.lookup.news : null;
  if (manual) {
    return [
      bodyPara("Recent company news relevant to the engagement objective, as provided:", {}),
      ...manual.slice(0, 4).map((n) => bodyPara(softenText(n), { bullet: true })),
      sourceLine("Source: As provided with the engagement request"),
    ];
  }
  if (looked) {
    return [
      bodyPara("Recent coverage mentioning the company, from the GDELT global news index:", {}),
      ...looked.map((n) => bodyPara(`${n.title}${n.source ? ` (${n.source}${n.date ? ", " + n.date : ""})` : ""}`, { bullet: true })),
      sourceLine("Source: GDELT news index; verify against the company newsroom"),
    ];
  }
  return [
    bodyPara("Recent company news relevant to the engagement objective, with links to sources:", {}),
    bodyPara("[Headline 1: from the company newsroom or investor relations page]", { bullet: true }),
    bodyPara("[Headline 2: from sector press or analyst coverage]", { bullet: true }),
  ];
}

function peopleTable(deal) {
  const manual = deal.profile && deal.profile.people && deal.profile.people.length ? deal.profile.people : null;
  if (manual) {
    const rows = manual.slice(0, 5).map((p) => [softenText(p.role), softenText(p.name)]);
    while (rows.length < 3) rows.push(["[Other key role]", "[Name]"]);
    return placeholderTable(["Role", "Name"], rows);
  }
  const L = deal.lookup || {};
  return placeholderTable(
    ["Role", "Name"],
    [
      ["Chief Executive Officer", L.ceo || "[Name]"],
      ["Chair of the Board", L.chair || "[Name]"],
      ["Chief Financial Officer", "[Name]"],
      ["[Other key role]", "[Name]"],
    ]
  );
}

function productsTable(deal) {
  const manual = deal.profile && deal.profile.products && deal.profile.products.length ? deal.profile.products : null;
  if (manual) {
    return placeholderTable(
      ["Name", "Revenue", "Description"],
      manual.slice(0, 5).map((p) => [softenText(p.name), "[$M]", softenText(p.desc) || "[Brief description]"])
    );
  }
  const looked = deal.lookup && deal.lookup.products && deal.lookup.products.length ? deal.lookup.products : null;
  if (looked) {
    return placeholderTable(
      ["Name", "Revenue", "Description"],
      looked.slice(0, 4).map((p) => [p.charAt(0).toUpperCase() + p.slice(1), "[$M]", "[Confirm revenue split in interviews]"])
    );
  }
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
  rows.push([
    "Revenue growth (%)",
    "n/a",
    ...m.series.revenue.slice(1).map((v, i) => fmtPct(v / m.series.revenue[i] - 1)),
    fmtPct(g),
  ]);
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

/* Deterministic text fitting: PowerPoint only applies its shrink
   autofit when a box is edited, so on first open long text would
   still spill. Instead the fitting font size is computed here, with
   conservative width math (0.6em average character, 1.4 line height,
   0.1in internal margins), and written into the file directly. */
function fitSize(text, wIn, hIn, baseSize) {
  return fitSizeLines([String(text)], wIn, hIn, baseSize);
}

function fitSizeLines(lines, wIn, hIn, baseSize) {
  const usableW = Math.max(0.3, wIn - 0.1);
  let size = baseSize;
  while (size > 8) {
    const fontIn = size / 72;
    const perLine = Math.max(1, Math.floor(usableW / (fontIn * 0.6)));
    let needed = 0;
    for (const l of lines) needed += Math.max(1, Math.ceil(String(l).length / perLine));
    if (needed * fontIn * 1.4 <= hIn) break;
    size -= 0.5;
  }
  return size;
}

// Revenue columns plus a separate EBITDA margin line on the summary
// slide. Two charts, one measure each: never a dual axis. Single
// series per chart, so no legends; values labeled directly, no
// gridlines, hidden value axes.
function addSummaryCharts(slide, m) {
  const CX = 6.85, CW = SLIDE_W - MARGIN - 6.85;
  const big = m.revenueLatest >= 10000;
  const div = big ? 1000 : 1;
  const labels = m.years.map((y) => fyLabel(y));
  const axisOpts = {
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catAxisLineShow: false,
    catAxisLabelFontFace: THEME.sans,
    catAxisLabelFontSize: 10,
    catAxisLabelColor: THEME.gray,
    showLegend: false,
    showTitle: false,
    dataLabelFontFace: THEME.sans,
    dataLabelFontSize: 9,
    dataLabelColor: "404850",
  };

  slide.addText(`REVENUE, ${big ? "$B" : "$M"}`, {
    x: CX, y: 1.42, w: CW, h: 0.26,
    fontFace: THEME.sans, fontSize: 10, bold: true, color: THEME.gray, charSpacing: 1,
  });
  slide.addChart("bar", [{ name: "Revenue", labels, values: m.series.revenue.map((v) => +(v / div).toFixed(1)) }], {
    x: CX, y: 1.7, w: CW, h: m.ebitdaSeries ? 2.15 : 4.4,
    barDir: "col", barGapWidthPct: 40,
    chartColors: [THEME.accent],
    showValue: true, dataLabelPosition: "outEnd",
    ...axisOpts,
  });

  if (m.ebitdaSeries) {
    slide.addText("EBITDA MARGIN, %", {
      x: CX, y: 4.02, w: CW, h: 0.26,
      fontFace: THEME.sans, fontSize: 10, bold: true, color: THEME.gray, charSpacing: 1,
    });
    slide.addChart("line", [{ name: "EBITDA margin", labels, values: m.ebitdaSeries.map((e, i) => +((e / m.series.revenue[i]) * 100).toFixed(1)) }], {
      x: CX, y: 4.3, w: CW, h: 1.95,
      chartColors: [THEME.accent2],
      lineSize: 2, lineSmooth: false, lineDataSymbol: "circle", lineDataSymbolSize: 6,
      showValue: true, dataLabelPosition: "t",
      ...axisOpts,
    });
  }
}

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
  // Action title: full-sentence takeaway, fixed position, largest text
  // on the slide. Font size is pre computed to fit the box.
  slide.addText(title, {
    x: MARGIN, y: 0.5, w: SLIDE_W - 2 * MARGIN, h: 0.9,
    fontFace: THEME.serif, fontSize: fitSize(title, SLIDE_W - 2 * MARGIN, 0.9, 19),
    color: THEME.ink, bold: false, valign: "top",
    fit: "shrink",
  });
  if (takeaway) {
    slide.addShape("rect", { x: MARGIN, y: 6.42, w: SLIDE_W - 2 * MARGIN, h: 0.52, fill: { color: THEME.fill }, line: { type: "none" } });
    slide.addShape("rect", { x: MARGIN, y: 6.42, w: 0.06, h: 0.52, fill: { color: THEME.accent }, line: { type: "none" } });
    slide.addText(takeaway, {
      x: MARGIN + 0.18, y: 6.42, w: SLIDE_W - 2 * MARGIN - 0.3, h: 0.52,
      fontFace: THEME.sans, fontSize: fitSize(takeaway, SLIDE_W - 2 * MARGIN - 0.3, 0.52, 12.5),
      bold: true, color: THEME.ink, valign: "middle",
      fit: "shrink",
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
    : m.source === "lookup"
    ? "Source: Reported figures via Wikidata; DealDesk analysis"
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
  const tTitle = P.deckTitle.charAt(0).toUpperCase() + P.deckTitle.slice(1);
  s.addText(tTitle, { x: 0.6, y: 2.45, w: 12, h: 1.6, fontFace: THEME.serif, fontSize: fitSize(tTitle, 12, 1.6, 38), color: THEME.onInk });
  const tSub = `${deal.company} | ${deal.dealType} | ${todayLabel()}`;
  s.addText(tSub, { x: 0.6, y: 4.1, w: 11, h: 0.4, fontFace: THEME.sans, fontSize: fitSize(tSub, 11, 0.4, 13), color: "B8C2CC" });
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
  // Left column: framing plus the two category signposts
  s.addText(n.dealFrame, {
    x: MARGIN, y: 1.42, w: 6.15, h: 0.95,
    fontFace: THEME.sans, fontSize: fitSize(n.dealFrame, 6.15, 0.95, 12),
    color: "404850", valign: "top",
    fit: "shrink",
  });
  groups.forEach((g, gi) => {
    const y = 2.5 + gi * 1.9;
    s.addText(`${gi + 1}. ${g.label.toUpperCase()} (SLIDE ${gi + 2})`, {
      x: MARGIN, y, w: 6.15, h: 0.3,
      fontFace: THEME.sans, fontSize: 11, bold: true, color: THEME.accent, charSpacing: 1,
    });
    const sign = g.items.map((it) => `${all.findIndex((a) => a.name === it.name) + 1}. ${it.name}`);
    const signSize = fitSizeLines(sign, 6.15, 1.45, 12);
    s.addText(
      sign.map((t) => ({ text: t, options: { breakLine: true, color: THEME.ink, bold: false } })),
      { x: MARGIN, y: y + 0.32, w: 6.15, h: 1.45, fontFace: THEME.sans, fontSize: signSize, lineSpacing: signSize * 1.7, valign: "top", fit: "shrink" }
    );
  });
  // Right column: the financial context as charts (one measure per chart)
  addSummaryCharts(s, m);

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
      const nameSize = fitSize(it.name + "   Impact: Medium  |  Ease: Medium", 11.7, 0.42, 15);
      sl.addText(
        [
          { text: it.name, options: { bold: true, color: THEME.ink, fontSize: nameSize } },
          { text: `   Impact: ${it.impact}  |  Ease: ${it.ease}`, options: { color: THEME.gray, fontSize: Math.min(10.5, nameSize - 3) } },
        ],
        { x: 1.05, y, w: 11.7, h: 0.42, fontFace: THEME.sans, valign: "middle", fit: "shrink" }
      );
      sl.addText(it.rationale, {
        x: 1.05, y: y + 0.44, w: 11.7, h: 0.95,
        fontFace: THEME.sans, fontSize: fitSize(it.rationale, 11.7, 0.95, 12.5),
        color: "404850", valign: "top", lineSpacing: 18,
        fit: "shrink",
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
  const legend = all.map((it, i) => `${i + 1}. ${it.name}`);
  const legendSize = fitSizeLines(legend, 4.3, 4.2, 11);
  s.addText(
    legend.map((t) => ({ text: t, options: { breakLine: true, color: "404850" } })),
    { x: 8.55, y: 1.9, w: 4.3, h: 4.2, fontFace: THEME.sans, fontSize: legendSize, lineSpacing: legendSize * 1.8, valign: "top", fit: "shrink" }
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
  // Four linked sheets, banker style: yellow input cells drive every
  // other number. Inputs (assumptions and multiples), Model
  // (historicals plus a five year projection), Synergies (three
  // scenarios), Valuation (multiple driven, with value creation).
  const wb = XLSX.utils.book_new();
  const term = practiceTerm(deal);
  const sc = computeScenarios(deal, m);
  const inp = sc.inputs;
  const NUMFMT = "#,##0.0;(#,##0.0)";
  const PCT = "0.0%";
  const MULT = '0.0"x"';

  const thin = { style: "thin", color: { rgb: "D9D9DD" } };
  const box = { top: thin, bottom: thin, left: thin, right: thin };
  const ST = {
    title: { font: { name: "Arial", sz: 13, bold: true, color: { rgb: THEME.ink } } },
    note: { font: { name: "Arial", sz: 9, italic: true, color: { rgb: "63666A" } } },
    section: { font: { name: "Arial", sz: 10, bold: true, color: { rgb: THEME.onInk } }, fill: { patternType: "solid", fgColor: { rgb: THEME.ink } } },
    label: { font: { name: "Arial", sz: 10, color: { rgb: "1A1A1A" } } },
    input: { font: { name: "Arial", sz: 10, bold: true, color: { rgb: "1F4E79" } }, fill: { patternType: "solid", fgColor: { rgb: "FFF2CC" } }, border: box, alignment: { horizontal: "right" } },
    calc: { font: { name: "Arial", sz: 10, color: { rgb: "1A1A1A" } }, alignment: { horizontal: "right" } },
    calcBold: { font: { name: "Arial", sz: 10, bold: true, color: { rgb: "1A1A1A" } }, fill: { patternType: "solid", fgColor: { rgb: "F2F5F7" } }, border: { top: { style: "thin", color: { rgb: THEME.ink } } }, alignment: { horizontal: "right" } },
    labelBold: { font: { name: "Arial", sz: 10, bold: true, color: { rgb: "1A1A1A" } }, fill: { patternType: "solid", fgColor: { rgb: "F2F5F7" } }, border: { top: { style: "thin", color: { rgb: THEME.ink } } } },
    pctRow: { font: { name: "Arial", sz: 9, italic: true, color: { rgb: "63666A" } }, alignment: { horizontal: "right" } },
    yearA: { font: { name: "Arial", sz: 10, bold: true, color: { rgb: THEME.onInk } }, fill: { patternType: "solid", fgColor: { rgb: THEME.ink } }, alignment: { horizontal: "center" } },
    yearF: { font: { name: "Arial", sz: 10, bold: true, color: { rgb: THEME.ink } }, fill: { patternType: "solid", fgColor: { rgb: THEME.fillAccent } }, alignment: { horizontal: "center" } },
  };

  function newSheet(widths) {
    return { ws: { "!cols": widths.map((wch) => ({ wch })) }, maxR: 0, maxC: 0 };
  }
  function put(sh, r, c, v, opts = {}) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = {};
    if (opts.f !== undefined) { cell.f = opts.f; cell.t = "n"; cell.v = opts.v !== undefined ? opts.v : 0; }
    else if (typeof v === "number") { cell.t = "n"; cell.v = v; }
    else { cell.t = "s"; cell.v = v === null || v === undefined ? "" : String(v); }
    if (opts.z) cell.z = opts.z;
    if (opts.s) cell.s = opts.s;
    sh.ws[addr] = cell;
    sh.maxR = Math.max(sh.maxR, r);
    sh.maxC = Math.max(sh.maxC, c);
    return addr;
  }
  function section(sh, r, text, span) {
    for (let c = 0; c < span; c++) put(sh, r, c, c === 0 ? text : "", { s: ST.section });
  }
  function finish(sh, name) {
    sh.ws["!ref"] = "A1:" + XLSX.utils.encode_cell({ r: sh.maxR, c: sh.maxC });
    XLSX.utils.book_append_sheet(wb, sh.ws, name);
  }

  /* ---------------- Sheet 1: Inputs ---------------- */
  const I = newSheet([48, 14]);
  const refs = {};
  let r = 0;
  put(I, r++, 0, `${deal.company}: model inputs ($ millions)`, { s: ST.title });
  put(I, r++, 0, "Yellow cells are inputs. Every other number in this workbook is a live formula linked to them.", { s: ST.note });
  r++;
  section(I, r++, "CLIENT DATA, LATEST ACTUAL YEAR " + fyLabel(m.lastYear), 2);
  put(I, r, 0, "Revenue", { s: ST.label });
  refs.revenue = put(I, r++, 1, inp.revenue, { z: NUMFMT, s: ST.input });
  for (const l of inp.lines) {
    put(I, r, 0, l.label, { s: ST.label });
    refs[l.key + "Base"] = put(I, r++, 1, l.base, { z: NUMFMT, s: ST.input });
  }
  r++;
  section(I, r++, "PROJECTION ASSUMPTIONS", 2);
  const growthDefault = Math.max(-0.5, Math.min(0.5, m.revenueCAGR));
  put(I, r, 0, "Revenue growth, annual", { s: ST.label });
  refs.growth = put(I, r++, 1, growthDefault, { z: PCT, s: ST.input });
  const projMarginDefault = m.ebitdaMargin !== null && m.ebitdaMargin > 0 ? Math.round(m.ebitdaMargin * 1000) / 1000 : 0.15;
  put(I, r, 0, "EBITDA margin, projection years", { s: ST.label });
  refs.projMargin = put(I, r++, 1, projMarginDefault, { z: PCT, s: ST.input });
  r++;
  section(I, r++, "SYNERGY ASSUMPTIONS, EACH APPLIED TO ITS OWN LINE", 2);
  put(I, r, 0, "Revenue uplift, low end", { s: ST.label });
  refs.revLow = put(I, r++, 1, inp.revLow, { z: PCT, s: ST.input });
  put(I, r, 0, "Revenue uplift, high end", { s: ST.label });
  refs.revHigh = put(I, r++, 1, inp.revHigh, { z: PCT, s: ST.input });
  put(I, r, 0, "Flow through margin on new revenue" + (inp.flowDefaulted ? " (defaulted)" : ""), { s: ST.label });
  refs.margin = put(I, r++, 1, inp.flowMargin, { z: PCT, s: ST.input });
  for (const l of inp.lines) {
    put(I, r, 0, `${l.label}: ${term} %, low end`, { s: ST.label });
    refs[l.key + "Low"] = put(I, r++, 1, l.low, { z: PCT, s: ST.input });
    put(I, r, 0, `${l.label}: ${term} %, high end`, { s: ST.label });
    refs[l.key + "High"] = put(I, r++, 1, l.high, { z: PCT, s: ST.input });
  }
  r++;
  section(I, r++, "VALUATION ASSUMPTIONS", 2);
  const evMultDefault = m.ebitdaMargin !== null && m.ebitdaMargin > 0.2 ? 10 : m.ebitdaMargin !== null && m.ebitdaMargin > 0.1 ? 8 : 6;
  put(I, r, 0, "EV to EBITDA multiple", { s: ST.label });
  refs.evM = put(I, r++, 1, evMultDefault, { z: MULT, s: ST.input });
  put(I, r, 0, "EV to revenue multiple, cross check", { s: ST.label });
  refs.revM = put(I, r++, 1, 1.5, { z: MULT, s: ST.input });
  finish(I, "Inputs");
  const IR = (a) => "Inputs!" + a.replace(/([A-Z]+)(\d+)/, "$$$1$$$2");

  /* ---------------- Sheet 2: Model ---------------- */
  const M = newSheet([26].concat(Array(m.years.length + 5).fill(11)));
  r = 0;
  put(M, r++, 0, `${deal.company}: historicals and five year projection ($ millions)`, { s: ST.title });
  put(M, r++, 0, "A = actual, F = forecast. Forecast cells are formulas driven by the Inputs sheet.", { s: ST.note });
  const nA = m.years.length;
  put(M, r, 0, "", { s: ST.section });
  for (let i = 0; i < nA; i++) put(M, r, 1 + i, fyLabel(m.years[i]), { s: ST.yearA });
  for (let i = 0; i < 5; i++) put(M, r, 1 + nA + i, fyLabel(m.lastYear + 1 + i, "F"), { s: ST.yearF });
  r++;
  const rev = m.series.revenue;
  const growthV = growthDefault;
  const revAddrs = [];
  put(M, r, 0, "Revenue", { s: ST.label });
  for (let i = 0; i < nA; i++) revAddrs.push(put(M, r, 1 + i, rev[i], { z: NUMFMT, s: ST.calc }));
  let prevRevV = rev[nA - 1];
  for (let i = 0; i < 5; i++) {
    const prevAddr = revAddrs[revAddrs.length - 1];
    prevRevV = prevRevV * (1 + growthV);
    revAddrs.push(put(M, r, 1 + nA + i, null, { f: `${prevAddr}*(1+${IR(refs.growth)})`, v: prevRevV, z: NUMFMT, s: ST.calc }));
  }
  r++;
  const lineAddrs = {};
  for (const key of ["cogs", "opex"]) {
    if (!m.series[key]) continue;
    const label = key === "cogs" ? "COGS" : "Opex / SG&A";
    const arr = m.series[key];
    lineAddrs[key] = [];
    put(M, r, 0, label, { s: ST.label });
    for (let i = 0; i < nA; i++) lineAddrs[key].push(put(M, r, 1 + i, arr[i], { z: NUMFMT, s: ST.calc }));
    const ratioV = arr[nA - 1] / rev[nA - 1];
    for (let i = 0; i < 5; i++) {
      lineAddrs[key].push(put(M, r, 1 + nA + i, null, {
        f: `${revAddrs[nA + i]}*(${lineAddrs[key][nA - 1]}/${revAddrs[nA - 1]})`,
        v: rev[nA - 1] * Math.pow(1 + growthV, i + 1) * ratioV, z: NUMFMT, s: ST.calc,
      }));
    }
    r++;
  }
  const hasLines = !!(m.series.cogs && m.series.opex);
  const ebitdaAddrs = [];
  put(M, r, 0, "EBITDA", { s: ST.labelBold });
  for (let i = 0; i < nA + 5; i++) {
    if (hasLines) {
      const v = i < nA
        ? rev[i] - m.series.cogs[i] - m.series.opex[i]
        : (rev[nA - 1] - m.series.cogs[nA - 1] - m.series.opex[nA - 1]) * Math.pow(1 + growthV, i - nA + 1);
      ebitdaAddrs.push(put(M, r, 1 + i, null, { f: `${revAddrs[i]}-${lineAddrs.cogs[i]}-${lineAddrs.opex[i]}`, v, z: NUMFMT, s: ST.calcBold }));
    } else if (i >= nA) {
      const v = rev[nA - 1] * Math.pow(1 + growthV, i - nA + 1) * projMarginDefault;
      ebitdaAddrs.push(put(M, r, 1 + i, null, { f: `${revAddrs[i]}*${IR(refs.projMargin)}`, v, z: NUMFMT, s: ST.calcBold }));
    } else {
      ebitdaAddrs.push(null);
      put(M, r, 1 + i, "n/a", { s: ST.pctRow });
    }
  }
  r++;
  put(M, r, 0, "Revenue growth %", { s: ST.label });
  for (let i = 1; i < nA + 5; i++) {
    const v = i < nA ? rev[i] / rev[i - 1] - 1 : growthV;
    put(M, r, 1 + i, null, { f: `(${revAddrs[i]}-${revAddrs[i - 1]})/${revAddrs[i - 1]}`, v, z: PCT, s: ST.pctRow });
  }
  r++;
  if (hasLines) {
    put(M, r, 0, "EBITDA margin %", { s: ST.label });
    for (let i = 0; i < nA + 5; i++) {
      if (!ebitdaAddrs[i]) continue;
      const revV = i < nA ? rev[i] : rev[nA - 1] * Math.pow(1 + growthV, i - nA + 1);
      const ebV = i < nA
        ? rev[i] - m.series.cogs[i] - m.series.opex[i]
        : (rev[nA - 1] - m.series.cogs[nA - 1] - m.series.opex[nA - 1]) * Math.pow(1 + growthV, i - nA + 1);
      put(M, r, 1 + i, null, { f: `${ebitdaAddrs[i]}/${revAddrs[i]}`, v: ebV / revV, z: PCT, s: ST.pctRow });
    }
    r++;
  }
  const baseEbitdaAddr = ebitdaAddrs[hasLines ? nA - 1 : nA];
  const baseEbitdaV = hasLines
    ? rev[nA - 1] - m.series.cogs[nA - 1] - m.series.opex[nA - 1]
    : rev[nA - 1] * (1 + growthV) * projMarginDefault;
  finish(M, "Model");

  /* ---------------- Sheet 3: Synergies ---------------- */
  const S = newSheet([42, 14, 14, 14]);
  r = 0;
  put(S, r++, 0, `${deal.company}: annual ${term} value scenarios ($ millions)`, { s: ST.title });
  put(S, r++, 0, "Conservative = low end of every range. Aggressive = high end. Midpoint = average.", { s: ST.note });
  put(S, r, 0, "", { s: ST.section });
  ["Conservative", "Midpoint", "Aggressive"].forEach((h, i) => put(S, r, 1 + i, h, { s: ST.yearA }));
  r++;
  const scen = [sc.conservative, sc.midpoint, sc.aggressive];
  const revFormulas = [
    `${IR(refs.revenue)}*${IR(refs.revLow)}*${IR(refs.margin)}`,
    `${IR(refs.revenue)}*AVERAGE(${IR(refs.revLow)},${IR(refs.revHigh)})*${IR(refs.margin)}`,
    `${IR(refs.revenue)}*${IR(refs.revHigh)}*${IR(refs.margin)}`,
  ];
  const synFirstExcelRow = r + 1;
  put(S, r, 0, `Revenue ${term}, at flow through margin`, { s: ST.label });
  for (let i = 0; i < 3; i++) put(S, r, 1 + i, null, { f: revFormulas[i], v: scen[i].revenueImpact, z: NUMFMT, s: ST.calc });
  r++;
  for (const l of inp.lines) {
    put(S, r, 0, `${l.label} ${term}`, { s: ST.label });
    const fns = [
      `${IR(refs[l.key + "Base"])}*${IR(refs[l.key + "Low"])}`,
      `${IR(refs[l.key + "Base"])}*AVERAGE(${IR(refs[l.key + "Low"])},${IR(refs[l.key + "High"])})`,
      `${IR(refs[l.key + "Base"])}*${IR(refs[l.key + "High"])}`,
    ];
    const vals = [l.base * l.low, l.base * (l.low + l.high) / 2, l.base * l.high];
    for (let i = 0; i < 3; i++) put(S, r, 1 + i, null, { f: fns[i], v: vals[i], z: NUMFMT, s: ST.calc });
    r++;
  }
  put(S, r, 0, `Total ${term} value`, { s: ST.labelBold });
  const synTotalAddrs = [];
  for (let i = 0; i < 3; i++) {
    const col = XLSX.utils.encode_col(1 + i);
    synTotalAddrs.push(put(S, r, 1 + i, null, { f: `SUM(${col}${synFirstExcelRow}:${col}${r})`, v: scen[i].total, z: NUMFMT, s: ST.calcBold }));
  }
  r++;
  put(S, r, 0, "Total as % of revenue", { s: ST.label });
  for (let i = 0; i < 3; i++) put(S, r, 1 + i, null, { f: `${synTotalAddrs[i]}/${IR(refs.revenue)}`, v: scen[i].total / inp.revenue, z: PCT, s: ST.pctRow });
  finish(S, "Synergies");
  const synMidRef = "Synergies!" + synTotalAddrs[1];

  /* ---------------- Sheet 4: Valuation ---------------- */
  const V = newSheet([46, 16]);
  r = 0;
  put(V, r++, 0, `${deal.company}: valuation view ($ millions)`, { s: ST.title });
  put(V, r++, 0, "Change the multiples on the Inputs sheet; every value here recalculates.", { s: ST.note });
  r++;
  section(V, r++, "ENTERPRISE VALUE AT THE EV TO EBITDA MULTIPLE", 2);
  put(V, r, 0, hasLines ? `EBITDA, ${fyLabel(m.lastYear)} actual` : `EBITDA, ${fyLabel(m.lastYear + 1, "F")} projected`, { s: ST.label });
  const vBase = put(V, r++, 1, null, { f: "Model!" + baseEbitdaAddr, v: baseEbitdaV, z: NUMFMT, s: ST.calc });
  put(V, r, 0, "Enterprise value at the multiple", { s: ST.label });
  const vEV = put(V, r++, 1, null, { f: `${vBase}*${IR(refs.evM)}`, v: baseEbitdaV * evMultDefault, z: NUMFMT, s: ST.calc });
  put(V, r, 0, `Midpoint annual ${term} value`, { s: ST.label });
  const vSyn = put(V, r++, 1, null, { f: synMidRef, v: sc.midpoint.total, z: NUMFMT, s: ST.calc });
  put(V, r, 0, `EBITDA including ${term} capture`, { s: ST.label });
  const vBaseSyn = put(V, r++, 1, null, { f: `${vBase}+${vSyn}`, v: baseEbitdaV + sc.midpoint.total, z: NUMFMT, s: ST.calc });
  put(V, r, 0, `Enterprise value including ${term} capture`, { s: ST.label });
  const vEVSyn = put(V, r++, 1, null, { f: `${vBaseSyn}*${IR(refs.evM)}`, v: (baseEbitdaV + sc.midpoint.total) * evMultDefault, z: NUMFMT, s: ST.calc });
  put(V, r, 0, `Value creation from ${term} capture`, { s: ST.labelBold });
  put(V, r++, 1, null, { f: `${vEVSyn}-${vEV}`, v: sc.midpoint.total * evMultDefault, z: NUMFMT, s: ST.calcBold });
  r++;
  section(V, r++, "REVENUE CROSS CHECK", 2);
  put(V, r, 0, `Revenue, ${fyLabel(m.lastYear)} actual`, { s: ST.label });
  const vRev = put(V, r++, 1, null, { f: IR(refs.revenue), v: inp.revenue, z: NUMFMT, s: ST.calc });
  put(V, r, 0, "Enterprise value at the revenue multiple", { s: ST.label });
  put(V, r++, 1, null, { f: `${vRev}*${IR(refs.revM)}`, v: inp.revenue * 1.5, z: NUMFMT, s: ST.calc });
  r++;

  /* Two way sensitivity: EV across growth (rows) and multiple (cols).
     Every cell is a formula, so editing the Inputs shifts the whole
     grid around the new base case. */
  section(V, r++, "SENSITIVITY: EV ON NEXT YEAR EBITDA, ACROSS GROWTH AND MULTIPLE", 6);
  const marginExpr = hasLines ? `(${vBase}/${vRev})` : IR(refs.projMargin);
  const marginVal = hasLines ? baseEbitdaV / inp.revenue : projMarginDefault;
  const dgs = [-0.02, -0.01, 0, 0.01, 0.02];
  const dms = [-2, -1, 0, 1, 2];
  put(V, r, 0, "Growth below, multiple right", { s: ST.note });
  dms.forEach((dm, ci) => {
    put(V, r, 1 + ci, null, { f: `${IR(refs.evM)}+${dm}`, v: evMultDefault + dm, z: MULT, s: ST.yearA });
  });
  r++;
  dgs.forEach((dg) => {
    put(V, r, 0, null, { f: `${IR(refs.growth)}+${dg}`, v: growthDefault + dg, z: PCT, s: ST.yearF });
    dms.forEach((dm, ci) => {
      const isBase = dg === 0 && dm === 0;
      put(V, r, 1 + ci, null, {
        f: `${vRev}*(1+${IR(refs.growth)}+${dg})*${marginExpr}*(${IR(refs.evM)}+${dm})`,
        v: inp.revenue * (1 + growthDefault + dg) * marginVal * (evMultDefault + dm),
        z: NUMFMT,
        s: isBase ? ST.calcBold : ST.calc,
      });
    });
    r++;
  });
  put(V, r, 0, "EBITDA in each cell is next year revenue at that growth times the current margin.", { s: ST.note });
  finish(V, "Valuation");

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
