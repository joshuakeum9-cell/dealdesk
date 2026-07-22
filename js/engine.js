/* ============ DealDesk analysis engine ============
   Parses uploaded financials and computes the metrics + narrative
   language that drive all four deliverables. No AI, no server. */

const SAMPLE_FINANCIALS = {
  years: [2021, 2022, 2023, 2024],
  series: {
    revenue: [142.0, 154.0, 168.0, 183.0],
    cogs: [92.3, 99.1, 107.0, 114.9],
    opex: [28.4, 30.1, 32.0, 33.8],
    capex: [5.8, 6.1, 6.4, 6.9],
  },
  source: "sample",
};

/* ---------- Parsing ---------- */

const ROW_KEYWORDS = {
  revenue: /revenue|sales|turnover/i,
  cogs: /cogs|cost of goods|cost of sales|cost of revenue/i,
  opex: /opex|operating expense|sg&a|sga|overhead/i,
  ebitda: /ebitda/i,
  capex: /capex|capital expenditure/i,
};

// Parse the first sheet of an uploaded workbook into {years, series}.
// Expects rows of "label, value, value, ..." with a header row of years.
function parseFinancialsWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  let years = null;
  const series = {};

  for (const row of aoa) {
    if (!row || row.length < 2) continue;
    const label = String(row[0] || "");
    const values = row.slice(1).filter((v) => v !== null && v !== "");

    if (!years) {
      const asYears = values.map(yearFromCell).filter(Boolean);
      if (asYears.length >= 2) {
        years = asYears;
        continue;
      }
    }
    for (const [key, re] of Object.entries(ROW_KEYWORDS)) {
      if (re.test(label) && !series[key]) {
        const nums = values.map(toNumber).filter((v) => v !== null);
        if (nums.length >= 2) series[key] = nums;
      }
    }
  }

  if (!series.revenue) return null;
  const n = series.revenue.length;
  if (!years || years.length < n) {
    const end = new Date().getFullYear() - 1;
    years = Array.from({ length: n }, (_, i) => end - n + 1 + i);
  }
  years = years.slice(0, n);
  return { years, series, source: "upload" };
}

function yearFromCell(v) {
  const s = String(v).match(/(19|20)\d{2}/);
  return s ? Number(s[0]) : null;
}

function toNumber(v) {
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/[$,()\s]/g, ""));
  return Number.isFinite(n) ? n : null;
}

/* ---------- Metrics ---------- */

function computeMetrics(fin) {
  const { years, series } = fin;
  const rev = series.revenue;
  const n = rev.length;
  const last = n - 1;

  const cagr = n > 1 ? Math.pow(rev[last] / rev[0], 1 / (n - 1)) - 1 : 0;

  const grossProfit = series.cogs ? rev.map((r, i) => r - series.cogs[i]) : null;
  const ebitda = series.ebitda
    ? series.ebitda
    : grossProfit && series.opex
    ? grossProfit.map((g, i) => g - series.opex[i])
    : null;

  return {
    years,
    series,
    source: fin.source,
    firstYear: years[0],
    lastYear: years[last],
    revenueLatest: rev[last],
    revenueCAGR: cagr,
    grossMargin: grossProfit ? grossProfit[last] / rev[last] : null,
    ebitdaLatest: ebitda ? ebitda[last] : null,
    ebitdaMargin: ebitda ? ebitda[last] / rev[last] : null,
    ebitdaSeries: ebitda,
    growthWord: growthWord(cagr),
  };
}

function growthWord(cagr) {
  if (cagr > 0.15) return "rapid";
  if (cagr > 0.08) return "healthy";
  if (cagr > 0.03) return "moderate";
  if (cagr > 0) return "modest";
  return "declining";
}

/* ---------- Formatting helpers ---------- */

function fmtM(v) {
  if (v === null || v === undefined) return "n/a";
  return "$" + v.toFixed(1) + "M";
}

function fmtPct(v) {
  if (v === null || v === undefined) return "n/a";
  return (v * 100).toFixed(1) + "%";
}

/* ---------- Practice areas: every output tailors to one ---------- */

const PRACTICES = {
  strategy: {
    label: "Strategy",
    blurb: "Long term goals, market entry, competitive positioning",
    guideTitle: "Strategy Interview Guide",
    deckTitle: "Strategic options assessment",
    deckShort: "Strategic Options Presentation",
    frame: (deal) =>
      `The client has commissioned a strategic review of ${deal.company} covering competitive position, growth options, and portfolio choices${deal.context ? `, with a focus on ${deal.context.replace(/\.$/, "")}` : ""}.`,
    overview: (deal, m) =>
      `Market attractiveness, share position, and the durability of ${deal.company}'s competitive advantages frame every recommendation in this review.`,
    finding: {
      lead: "Strategic position",
      text: "Competitive positioning and market attractiveness are unvalidated; the growth thesis depends on both (see Section 4).",
    },
    anchor: {
      lead: "Opportunity anchor",
      text: "Focusing the core while pricing to value is the highest confidence move; market entry options carry more risk and should be staged (see the options presentation).",
    },
    riskBullet: "Strategic drift: confirm which businesses are core, and whether management's growth bets align with where the market is moving.",
    emailRisk: {
      lead: "Primary risk",
      text: "The competitive response to any strategic move is unmodeled; war gaming the top two options should precede commitment.",
    },
    interviewSections: [
      { section: "Objectives and context", questions: [
        "What prompted this strategic review now, and what decision does it need to inform?",
        "What would a successful outcome look like in three years, in terms you would measure?",
        "Which options are already on or off the table, and why?",
        "Who are the key stakeholders for this decision, and where do they disagree?",
      ]},
      { section: "Business overview and strategy", questions: [
        "How does the company win today: what do customers buy from you that they cannot get elsewhere?",
        "Which businesses or products do you consider core, and which are legacy?",
        "Where has the strategy changed in the last three years, and why?",
        "Which parts of the business earn returns above their cost of capital, and which do not?",
      ]},
      { section: "Market and competition", questions: [
        "How large is the addressable market, and how fast is it growing?",
        "Who are the top three competitors, and where are they winning share?",
        "Where do you have pricing power, and where are you a price taker?",
        "What could a new entrant or substitute do to your economics in the next five years?",
      ]},
      { section: "Growth options", questions: [
        "Which adjacencies (products, segments, geographies) have you evaluated, and what stopped you?",
        "For each growth option: would you build, buy, or partner, and why?",
        "What would you do with twice the investment capacity?",
        "Which past growth bets worked, which did not, and what separated them?",
      ]},
      { section: "Organization and execution capacity", questions: [
        "Which capabilities would need to be built or acquired to execute the growth agenda?",
        "Where has execution fallen short of strategy in the past, and what caused it?",
        "How quickly can the organization reallocate people and capital when priorities change?",
      ]},
    ],
    groups: [
      { label: "Growth moves", slideTitle: "Growth moves center on pricing to value and staged entry into proven adjacencies", slideTakeaway: "Stage growth investments behind evidence; pricing is the fastest lever", items: [
        { name: "Value based pricing reset", rationale: "Pricing is the fastest profit lever; segments with demonstrated willingness to pay are underpriced today.", impact: "High", ease: "High" },
        { name: "Adjacent segment entry", rationale: "The core capability set transfers to a neighboring customer segment with limited incremental investment.", impact: "High", ease: "Medium" },
        { name: "Geographic expansion", rationale: "Demand exists in adjacent regions the current footprint does not serve.", impact: "Medium", ease: "Low" },
      ]},
      { label: "Portfolio moves", slideTitle: "Portfolio moves sharpen the core by reallocating capital from legacy positions", slideTakeaway: "Fund the growth agenda by pruning what no longer earns its capital", items: [
        { name: "Core focus and simplification", rationale: "Complexity in the tail of the portfolio absorbs management attention disproportionate to its profit.", impact: "Medium", ease: "Medium" },
        { name: "Partnership or alliance", rationale: "A partner supplies the missing capability faster than building it internally.", impact: "Medium", ease: "Medium" },
        { name: "Divest noncore assets", rationale: "Legacy positions may be worth more to another owner; proceeds fund the growth agenda.", impact: "High", ease: "Low" },
      ]},
    ],
  },

  operations: {
    label: "Operations",
    blurb: "Efficiency, supply chain, cost reduction, workflows",
    guideTitle: "Operations Interview Guide",
    deckTitle: "Operational improvement assessment",
    deckShort: "Operational Improvement Presentation",
    frame: (deal) =>
      `The client has engaged an operational review of ${deal.company} to identify efficiency, cost, and throughput improvements${deal.context ? `, with a focus on ${deal.context.replace(/\.$/, "")}` : ""}.`,
    overview: (deal, m) =>
      `The operating model, cost structure, and supply chain of ${deal.company} are the units of analysis; financial results are treated as symptoms of operational drivers.`,
    finding: {
      lead: "Operational efficiency",
      text: "The cost base has not been benchmarked against peers; margin trends suggest addressable inefficiency in procurement and overhead (see Section 4).",
    },
    anchor: {
      lead: "Opportunity anchor",
      text: "Procurement and process improvement are the highest confidence levers; footprint changes carry more disruption risk (see the improvement presentation).",
    },
    riskBullet: "Operational fragility: identify single points of failure in the supply chain and capacity constraints that cap growth.",
    emailRisk: {
      lead: "Primary risk",
      text: "Savings estimates are top down until process level data is gathered; the bottom up validation could move totals materially in either direction.",
    },
    interviewSections: [
      { section: "Objectives and context", questions: [
        "What prompted this operational review now, and what outcome would make it a success?",
        "Which operational problems does leadership already agree on, and where is there debate?",
        "What improvement programs are already underway, and how are they tracking?",
        "What constraints (customer commitments, labor agreements, capital) limit the solution space?",
      ]},
      { section: "Operating model overview", questions: [
        "Walk us through the end to end flow from order to delivery: where does it slow down?",
        "Which sites or lines run at capacity, and which are underutilized?",
        "What are the three metrics you review daily, and what are they telling you?",
        "Where do handoffs between functions create delay or error?",
      ]},
      { section: "Process and quality", questions: [
        "Where are the largest sources of rework, scrap, or error?",
        "Which processes are documented and standardized vs. dependent on individuals?",
        "What continuous improvement programs exist, and what have they delivered?",
        "How is quality measured at each stage, and who acts on the data?",
      ]},
      { section: "Supply chain and procurement", questions: [
        "Who are the top five suppliers by spend, and when were their contracts last renegotiated?",
        "How many days of inventory do you hold, and what drives it?",
        "Where have logistics costs moved over the last two years?",
        "Which supply relationships are single sourced, and what is the backup plan?",
      ]},
      { section: "Cost structure", questions: [
        "Which costs are fixed vs. variable, and where is there discretion?",
        "If you had to take out 10% of cost in a year, where would you go first, and what stops you today?",
        "Which cost lines have grown faster than revenue, and why?",
      ]},
    ],
    groups: [
      { label: "Cost reduction", slideTitle: "Procurement and process waste are the largest near term cost reduction levers", slideTakeaway: "Cost levers are within management's direct control; sequence procurement first", items: [
        { name: "Procurement renegotiation", rationale: "Concentrated supplier spend with dated contracts supports renegotiated rates and consolidation.", impact: "High", ease: "High" },
        { name: "Process and lean improvement", rationale: "Rework and unstandardized processes absorb labor hours that standard work releases.", impact: "Medium", ease: "Medium" },
        { name: "Footprint optimization", rationale: "Underutilized sites carry fixed cost the network no longer needs.", impact: "High", ease: "Low" },
      ]},
      { label: "Performance improvement", slideTitle: "Planning and automation improvements raise throughput without new capacity", slideTakeaway: "Throughput gains defer capital spend; start where data already exists", items: [
        { name: "Demand planning and forecasting", rationale: "Forecast error drives both stockouts and excess inventory; better planning releases working capital.", impact: "Medium", ease: "Medium" },
        { name: "Targeted automation", rationale: "High volume repetitive steps are automatable with proven technology and rapid payback.", impact: "Medium", ease: "Medium" },
        { name: "Quality system uplift", rationale: "Reducing defect rates lowers cost of poor quality and warranty exposure simultaneously.", impact: "Medium", ease: "High" },
      ]},
    ],
  },

  financial: {
    label: "Financial / M&A",
    blurb: "M&A advisory, restructuring, profitability, risk",
    guideTitle: "Divestiture Interview Guide",
    deckTitle: "Qualitative synergy assessment",
    deckShort: "Qualitative Synergy Presentation",
    frame: null, // uses deal type specific framing below
    overview: (deal, m) =>
      `${deal.company} operates in the ${(deal.industry || "target").toLowerCase()} sector; the analysis focuses on quality of earnings, deal perimeter, and value creation levers.`,
    finding: null, // deal type specific, built in buildKeyFindings
    anchor: {
      lead: "Synergy anchor",
      text: "Procurement scale is the highest confidence synergy opportunity; revenue synergies should be treated as upside, not base case (see the synergy presentation).",
    },
    riskBullet: null, // deal type specific
    emailRisk: null, // deal type specific
    interviewSections: null, // deal type specific, legacy builder
    groups: null, // built dynamically in buildOpportunities
  },

};

function practiceOf(deal) {
  return PRACTICES[deal.practice] || PRACTICES.financial;
}

// What the value levers are called, per practice
const PRACTICE_TERMS = { strategy: "opportunity", operations: "improvement", financial: "synergy" };

function practiceTerm(deal) {
  return PRACTICE_TERMS[deal.practice] || "synergy";
}

/* ---------- Value scenarios (PwC Task 4 structure) ----------
   Per line percentage ranges applied to their own cost lines, revenue
   uplift flowing through at an assumed margin, three scenarios:
   conservative = low end, aggressive = high end, midpoint between. */

function buildScenarioInputs(deal, m) {
  const lines = [];
  if (m.series.cogs) {
    lines.push({ key: "cogs", label: "Cost of goods sold", base: m.series.cogs[m.series.cogs.length - 1], low: 0.05, high: 0.08 });
  }
  if (m.series.opex) {
    lines.push({ key: "opex", label: "Opex / G&A (back office)", base: m.series.opex[m.series.opex.length - 1], low: 0.1, high: 0.25 });
  }
  return {
    revenue: m.revenueLatest,
    revLow: 0.05,
    revHigh: 0.1,
    flowMargin: m.ebitdaMargin !== null ? Math.round(m.ebitdaMargin * 1000) / 1000 : 0.15,
    lines,
  };
}

function computeScenarios(deal, m) {
  const inp = buildScenarioInputs(deal, m);
  const calc = (pick) => {
    const rev = inp.revenue * pick(inp.revLow, inp.revHigh) * inp.flowMargin;
    const cost = inp.lines.reduce((s, l) => s + l.base * pick(l.low, l.high), 0);
    return { revenueImpact: rev, costImpact: cost, total: rev + cost };
  };
  return {
    inputs: inp,
    conservative: calc((lo) => lo),
    midpoint: calc((lo, hi) => (lo + hi) / 2),
    aggressive: calc((lo, hi) => hi),
  };
}

/* ---------- Narrative rules ---------- */

function buildNarrative(deal, m) {
  const growthSentence =
    m.revenueCAGR > 0
      ? `Revenue grew from ${fmtM(m.series.revenue[0])} in ${m.firstYear} to ${fmtM(m.revenueLatest)} in ${m.lastYear}, a ${m.growthWord} ${fmtPct(m.revenueCAGR)} compound annual growth rate.`
      : `Revenue declined from ${fmtM(m.series.revenue[0])} in ${m.firstYear} to ${fmtM(m.revenueLatest)} in ${m.lastYear}, an area for focused diligence.`;

  const marginSentence = m.ebitdaMargin
    ? m.ebitdaMargin > 0.2
      ? `EBITDA margin of ${fmtPct(m.ebitdaMargin)} is strong for the sector and suggests durable pricing power or cost discipline.`
      : m.ebitdaMargin > 0.1
      ? `EBITDA margin of ${fmtPct(m.ebitdaMargin)} is within a typical range, with room for operational improvement after close.`
      : `EBITDA margin of ${fmtPct(m.ebitdaMargin)} is thin, making cost structure a priority diligence area.`
    : "Profitability could not be derived from the uploaded statements; request a full income statement.";

  const P = practiceOf(deal);
  const dealFrame = P.frame
    ? P.frame(deal)
    : {
        "Acquisition": `The client is evaluating the acquisition of ${deal.company} to ${deal.context || "expand its market position"}.`,
        "Divestiture": `The client is preparing ${deal.company} for divestiture and needs a clear view of separation complexity and standalone economics.`,
        "Carve out": `${deal.company} is being carved out of its parent, so entanglements and transition services drive both risk and value.`,
        "Merger": `The contemplated merger with ${deal.company} turns on integration feasibility and the credibility of the combined entity synergy case.`,
      }[deal.dealType];

  return { growthSentence, marginSentence, dealFrame };
}

/* ---------- Numbered key findings (exec summary + deck) ---------- */

function buildKeyFindings(deal, m) {
  const findings = [
    {
      lead: "Top line growth",
      text:
        m.revenueCAGR > 0
          ? `Revenue grew from ${fmtM(m.series.revenue[0])} (${m.firstYear}) to ${fmtM(m.revenueLatest)} (${m.lastYear}), a ${fmtPct(m.revenueCAGR)} CAGR; confirm how much is volume vs. price (see Section 3).`
          : `Revenue declined from ${fmtM(m.series.revenue[0])} to ${fmtM(m.revenueLatest)} over the period; the value case rests on margin and synergies (see Section 3).`,
    },
    {
      lead: "Profitability",
      text:
        m.ebitdaMargin !== null
          ? `EBITDA of ${fmtM(m.ebitdaLatest)} implies a ${fmtPct(m.ebitdaMargin)} margin in ${m.lastYear}; ${m.ebitdaMargin > 0.2 ? "strong for the sector, but validate sustainability" : m.ebitdaMargin > 0.1 ? "within a typical range, with improvement potential after close" : "thin, making the cost base a priority diligence area"} (see Section 3).`
          : "Profitability could not be derived from the provided statements; request a full income statement before proceeding.",
    },
    practiceFinding(deal),
    practiceOf(deal).anchor,
  ];
  return findings;
}

function practiceFinding(deal) {
  const P = practiceOf(deal);
  if (P.finding) return P.finding;
  // Financial/M&A: the third finding varies by deal type
  const separation = deal.dealType === "Divestiture" || deal.dealType === "Carve out";
  return {
    lead: separation ? "Separation complexity" : "Integration readiness",
    text: separation
      ? "Entangled systems, shared services, and TSA scope are the largest unquantified risks; management interviews should map them first (see Section 4)."
      : "Operational dependencies and management retention determine speed to value; both are unconfirmed at this stage (see Section 4).",
  };
}

/* ---------- Interview guide rules (vary by deal type) ---------- */

function buildInterviewQuestions(deal) {
  const P = practiceOf(deal);
  if (P.interviewSections) return P.interviewSections;

  // Financial/M&A: deal type driven guide, pitched at CFO level.
  // Covers motivation, financials, entanglement, and standalone operations.
  const separation = deal.dealType === "Divestiture" || deal.dealType === "Carve out";

  const motivation = {
    section: "Motivation and objectives",
    questions: separation
      ? [
          `What is driving the decision to ${deal.dealType === "Divestiture" ? "divest" : "carve out"} this business, and why now?`,
          "What does a successful outcome look like: price, speed, certainty, or ongoing relationship?",
          "Have you explored alternatives (sale, spin, joint venture, wind down), and what ruled them out?",
          "What are the board's expectations on timeline, and what could move them?",
        ]
      : [
          `What is the strategic rationale for this ${deal.dealType.toLowerCase()}, and why now?`,
          "What does a successful outcome look like twelve months after close?",
          "Which alternatives were considered, and what made this path the preferred one?",
          "Where does leadership disagree about this transaction?",
        ],
  };

  const overview = {
    section: "Business overview and perimeter",
    questions: [
      `Walk us through ${deal.company}'s operating model: key products, customer segments, and how revenue is generated.`,
      "Which products, sites, contracts, and people sit inside the transaction perimeter?",
      "Which shared functions or systems does the business rely on today?",
      "How was this business integrated when it was acquired or built, and how integrated is it now?",
    ],
  };

  const financials = {
    section: "Financial performance",
    questions: [
      "What drove the revenue trend over the historical period, and how much is volume vs. price?",
      "How is this business's performance reported internally if it is not broken out in the financial statements?",
      "Which costs are fixed vs. variable, and where is there discretion in the cost base?",
      "Are there any one time items, related party transactions, or accounting changes we should normalize for?",
    ],
  };

  const byType = {
    "Divestiture": {
      section: "Entanglement and shared services",
      questions: [
        "What services does the parent provide today (IT, HR, finance, procurement, real estate)?",
        "What would it cost to replicate each of these on a standalone basis?",
        "Which services will require a transition services agreement, and for how long?",
        "Are there systems, contracts, or licenses that cannot be separated by close?",
      ],
    },
    "Carve out": {
      section: "Separation and transition services",
      questions: [
        "Which entangled systems and contracts must be separated, and what is the realistic timeline?",
        "What stranded costs will remain with the parent, and who bears them?",
        "Which customer or supplier contracts require consent to transfer?",
        "Which shared teams serve both this business and the parent, and how would they split?",
      ],
    },
    "Acquisition": {
      section: "Integration and value creation",
      questions: [
        "Where do you see the clearest overlap with the acquirer's footprint, customers, or suppliers?",
        "What would break if the two organizations were combined within 12 months?",
        "Which investments has the business deferred that a new owner would need to fund?",
        "Which systems and processes would need to integrate first for the deal logic to work?",
      ],
    },
    "Merger": {
      section: "Combination feasibility",
      questions: [
        "Where are the cultural and operating model differences that could slow integration?",
        "Which leadership roles are duplicated, and how would the combined organization be structured?",
        "What customer overlap exists, and is there attrition risk from the combination?",
        "Which systems would the combined entity run on, and who decides?",
      ],
    },
  };

  const standalone = {
    section: "Standalone operations and people",
    questions: [
      "Could this business operate independently on day one: what would it lack?",
      "Which key employees are critical to the business, and would they transfer with it?",
      "How dependent is the business on current ownership or a small number of leaders?",
    ],
  };

  return [motivation, overview, financials, byType[deal.dealType], standalone];
}

/* ---------- Opportunity hypotheses (vary by practice, deal type, metrics) ---------- */

// Returns { groups: [{label, slideTitle, slideTakeaway, items}, x2] }
function buildOpportunities(deal, m) {
  const P = practiceOf(deal);
  if (P.groups) return { groups: P.groups };
  const syn = buildSynergies(deal, m);
  return {
    groups: [
      {
        label: "Revenue synergies",
        slideTitle: "Revenue synergies center on cross selling into the combined customer base and widening the addressable footprint",
        slideTakeaway: "Revenue synergies carry more execution risk than cost synergies; treat as upside, not base case",
        items: syn.revenue,
      },
      {
        label: "Cost synergies",
        slideTitle: "Procurement scale and overlapping G&A anchor the cost case, with capture feasible inside 24 months",
        slideTakeaway: "Cost synergies should anchor the value case; they are within management's direct control",
        items: syn.cost,
      },
    ],
  };
}

function buildSynergies(deal, m) {
  const revenue = [
    {
      name: "Cross sell into the combined customer base",
      rationale: `${deal.company}'s customers can be offered the acquirer's adjacent products, and vice versa.`,
      impact: "High",
      ease: "Medium",
    },
    {
      name: "Geographic and channel expansion",
      rationale: "The combined footprint opens regions and channels neither business serves alone.",
      impact: "Medium",
      ease: "Medium",
    },
  ];
  const cost = [
    {
      name: "Procurement scale and vendor consolidation",
      rationale: "Combined purchasing volume supports renegotiated rates with overlapping vendors.",
      impact: "High",
      ease: "High",
    },
    {
      name: "Overlapping G&A rationalization",
      rationale: "Duplicated back office functions (finance, HR, IT) can be consolidated over 12 to 24 months.",
      impact: "Medium",
      ease: "Medium",
    },
    {
      name: "Facility and footprint optimization",
      rationale: "Locating operations together reduces lease and logistics costs where geographies overlap.",
      impact: "Medium",
      ease: "Low",
    },
  ];
  if (m.ebitdaMargin !== null && m.ebitdaMargin < 0.12) {
    cost.unshift({
      name: "Margin improvement to peer levels",
      rationale: `EBITDA margin of ${fmtPct(m.ebitdaMargin)} trails typical sector levels, suggesting addressable cost opportunity independent of the deal.`,
      impact: "High",
      ease: "Medium",
    });
  }
  return { revenue, cost };
}

/* ---------- Email summary (BLUF: answer first, then support) ---------- */

function buildEmail(deal, m) {
  const projRev = m.revenueLatest * Math.pow(1 + Math.max(m.revenueCAGR, 0.02), 5);
  const marginText =
    m.ebitdaMargin !== null ? fmtPct(m.ebitdaMargin) + " EBITDA margin" : "profitability to be confirmed";

  return {
    subject: `${deal.company}: recommend advancing to management interviews; ${fmtM(m.revenueLatest)} revenue, ${marginText}`,
    answer: `Recommendation: advance ${deal.company} to the management interview phase. The financial profile supports the ${practiceOf(deal).frame ? "engagement" : deal.dealType.toLowerCase()} thesis, subject to the diligence items below.`,
    soWhat: `This matters now because the historical financials show ${m.growthWord} growth (${fmtPct(m.revenueCAGR)} CAGR) with ${marginText}, and the open questions are answerable within a two week interview window.`,
    ask: `Decision needed: approve interview scheduling and data room access by end of week to hold the timeline.`,
    points: [
      {
        lead: "Directional value",
        text: (() => {
          const sc = computeScenarios(deal, m);
          const term = practiceTerm(deal);
          return `The model shows annual ${term} value of ${fmtM(sc.conservative.total)} conservative to ${fmtM(sc.aggressive.total)} aggressive (${fmtM(sc.midpoint.total)} midpoint), on ${fmtM(m.revenueLatest)} of ${m.lastYear} revenue. Every percentage sits in a labeled input cell, so the assumptions can be flexed directly.`;
        })(),
      },
      {
        lead: "Key sensitivity",
        text: `A one point change in the growth assumption moves year five revenue by roughly ${fmtM(m.revenueLatest * 0.05)}; growth quality is the single most valuable interview topic.`,
      },
      practiceOf(deal).emailRisk || {
        lead: "Primary risk",
        text:
          deal.dealType === "Divestiture" || deal.dealType === "Carve out"
            ? "Separation complexity is unquantified; TSA scope and standalone costs could materially change the economics."
            : "Synergy capture depends on procurement and G&A overlap that has not yet been validated bottom up.",
      },
    ],
    nextSteps: [
      "Conduct management interviews using the attached guide (weeks 1 to 2).",
      "Size the top three synergy opportunities bottom up (weeks 2 to 3).",
      "Fold validated findings into the deal model and reconvene (week 4).",
    ],
    close: "Full analysis in the attached model, summary, and interview guide. Happy to walk through the assumptions.",
  };
}
