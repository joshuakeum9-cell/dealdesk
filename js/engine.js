/* ============ DealDesk analysis engine ============
   Parses uploaded financials and computes the metrics and language
   that drive all four deliverables. No AI, no server: an expert
   system. A structured situation assessment plus computed metrics
   select findings, interview questions, and opportunity hypotheses
   from a scored library, so two different situations produce
   genuinely different documents. */

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
// Handles the standard layout (metrics down, years across) and the
// transposed layout (years down, metrics across). Quarterly data is
// rejected rather than silently treated as annual.
function parseFinancialsWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  return parseAoa(aoa) || parseAoa(transposeAoa(aoa));
}

function transposeAoa(aoa) {
  const width = Math.max(0, ...aoa.map((r) => (r ? r.length : 0)));
  const out = [];
  for (let c = 0; c < width; c++) out.push(aoa.map((r) => (r ? r[c] : null)));
  return out;
}

function parseAoa(aoa) {
  let years = null;
  const series = {};

  for (const row of aoa) {
    if (!row || row.length < 2) continue;
    const label = String(row[0] || "");
    const values = row.slice(1).filter((v) => v !== null && v !== "");

    if (!years) {
      const asYears = values.map(yearFromCell).filter(Boolean);
      if (asYears.length >= 2) {
        // Duplicate years mean sub annual columns (Q1 2024, Q2 2024...).
        // Refuse rather than compute a fake annual growth rate.
        if (new Set(asYears).size !== asYears.length) return null;
        years = asYears;
        continue;
      }
    }
    for (const [key, re] of Object.entries(ROW_KEYWORDS)) {
      if (re.test(label) && !series[key]) {
        let nums = values.map(toNumber).filter((v) => v !== null);
        // Cost and revenue lines are stored as magnitudes; sheets that
        // show expenses as (800) or -800 mean the same cost.
        if (key !== "revenue") nums = nums.map(Math.abs);
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
  const s = String(v).trim();
  const negative = /^\(.*\)$/.test(s);
  const n = Number(s.replace(/[$,()\s]/g, ""));
  if (!Number.isFinite(n)) return null;
  return negative ? -n : n;
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

  // Is the computed margin trend itself rising or falling?
  let marginDirection = null;
  if (ebitda && n > 1) {
    const mFirst = ebitda[0] / rev[0];
    const mLast = ebitda[last] / rev[last];
    marginDirection = mLast - mFirst > 0.01 ? "rising" : mFirst - mLast > 0.01 ? "falling" : "flat";
  }

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
    marginDirection,
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
  if (Math.abs(v) >= 10000) return "$" + (v / 1000).toFixed(1) + "B";
  return "$" + v.toFixed(1) + "M";
}

function fmtPct(v) {
  if (v === null || v === undefined) return "n/a";
  return (v * 100).toFixed(1) + "%";
}

/* ---------- Context guard ----------
   The engagement context box accepts anything, including long pasted
   memos. Framing sentences only ever take the first sentence, capped;
   the full text goes into its own document section instead. */

function focusPhrase(context) {
  if (!context) return "";
  let s = String(context).replace(/\s+/g, " ").trim();
  const stop = s.search(/[.!?]\s/);
  if (stop > 10) s = s.slice(0, stop);
  if (s.length > 140) s = s.slice(0, 140).replace(/\s\S*$/, "");
  return s;
}

// If the pasted context already reads as a complete framing sentence
// ("The client is evaluating...", or it names the company), use it AS
// the frame instead of embedding it into a template, which would
// double the subject: "...acquisition of X to The client is...".
function frameFromContext(deal) {
  const fp = focusPhrase(deal.context);
  if (!fp) return null;
  const firstWord = (deal.company || "").split(/\s+/)[0].toLowerCase();
  const readsAsFrame =
    /^(the client|our client|the parent|the company|management|we )/i.test(fp) ||
    (firstWord.length > 2 && fp.toLowerCase().includes(firstWord));
  if (!readsAsFrame) return null;
  return /[.!?]$/.test(fp) ? fp : fp + ".";
}

/* ---------- Situation assessment ----------
   Structured answers from the engagement form. Every value defaults
   to "" (not answered), and the engine treats unanswered as neutral. */

function situationOf(deal) {
  const s = deal.situation || {};
  return {
    goal: s.goal || "",
    marginTrend: s.marginTrend || "",       // Improving | Stable | Declining
    revenueDriver: s.revenueDriver || "",   // Volume | Price | Both | Declining
    concentration: s.concentration || "",   // Yes | No
    capacity: s.capacity || "",             // Tight | Balanced | Excess
    costProgram: s.costProgram || "",       // Yes | No
    competition: s.competition || "",       // Intensifying | Stable
    urgency: s.urgency || "",               // Weeks | Months | Exploratory
  };
}

// Goal options shown in the form, per practice.
const PRACTICE_GOALS = {
  strategy: ["Growth strategy", "Portfolio review", "Competitive response", "Market entry"],
  operations: ["Cost reduction", "Margin turnaround", "Capacity and footprint", "Supply chain resilience"],
  financial: ["Validate acquisition thesis", "Identify synergies", "Prepare a separation", "Maximize sale value"],
};

/* ---------- Practice areas ---------- */

const PRACTICES = {
  strategy: {
    label: "Strategy",
    blurb: "Long term goals, market entry, competitive positioning",
    guideTitle: "Strategy Interview Guide",
    deckTitle: "Strategic options assessment",
    deckShort: "Strategic Options Presentation",
    frame: (deal) => {
      const fc = frameFromContext(deal);
      if (fc) return fc + " This strategic review covers competitive position, growth options, and portfolio choices.";
      const fp = focusPhrase(deal.context);
      return `The client has commissioned a strategic review of ${deal.company} covering competitive position, growth options, and portfolio choices${fp ? `, with a focus on ${fp.replace(/\.$/, "")}` : ""}.`;
    },
    overview: (deal, m) =>
      `Market attractiveness, share position, and the durability of ${deal.company}'s competitive advantages frame every recommendation in this review.`,
    defaultRisk: {
      lead: "Strategic position",
      text: "Competitive positioning and market attractiveness are unvalidated; the growth thesis depends on both.",
    },
    emailRisk: {
      lead: "Primary risk",
      text: "The competitive response to any strategic move is unmodeled; war gaming the top two options should precede commitment.",
    },
    baseSections: [
      { section: "Objectives and context", questions: [
        "What prompted this strategic review now, and what decision does it need to inform?",
        "What would a successful outcome look like in three years, in terms you would measure?",
        "Which options are already on or off the table, and why?",
      ]},
      { section: "Business overview and strategy", questions: [
        "How does the company win today: what do customers buy from you that they cannot get elsewhere?",
        "Which businesses or products do you consider core, and which are legacy?",
        "Which parts of the business earn returns above their cost of capital, and which do not?",
      ]},
      { section: "Market and competition", questions: [
        "How large is the addressable market, and how fast is it growing?",
        "Who are the top three competitors, and where are they winning share?",
        "Where do you have pricing power, and where are you a price taker?",
      ]},
      { section: "Growth options", questions: [
        "Which adjacencies (products, segments, geographies) have you evaluated, and what stopped you?",
        "For each growth option: would you build, buy, or partner, and why?",
        "Which past growth bets worked, which did not, and what separated them?",
      ]},
      { section: "Organization and execution capacity", questions: [
        "Which capabilities would need to be built or acquired to execute the growth agenda?",
        "Where has execution fallen short of strategy in the past, and what caused it?",
        "How quickly can the organization reallocate people and capital when priorities change?",
      ]},
    ],
    conditionalQuestions: [
      { topic: 2, cond: (s) => s.competition === "Intensifying", q: "Which competitor move of the last year worried you most, and how did you respond?" },
      { topic: 1, cond: (s) => s.marginTrend === "Declining", q: "Margins have been under pressure: how much is price, how much is cost, and how much is mix?" },
      { topic: 3, cond: (s) => s.revenueDriver === "Declining", q: "With demand softening, which segments would you defend at all costs, and which would you let go?" },
      { topic: 1, cond: (s) => s.concentration === "Yes", q: "How dependent is the business on its largest customers, and how would losing one change the strategy?" },
      { topic: 4, cond: (s) => s.costProgram === "Yes", q: "How does the current cost program interact with the growth agenda: where do they conflict?" },
    ],
    groups: [
      { label: "Growth moves", slideTitle: "Growth moves center on the highest scoring levers for this situation", slideTakeaway: "Stage growth investments behind evidence; sequence the highest confidence lever first", items: [
        { name: "Value based pricing reset", rationale: "Pricing is the fastest profit lever; segments with demonstrated willingness to pay are underpriced today.", impact: "High", ease: "High", boost: (s, m) => (s.revenueDriver === "Volume" ? 2 : 0) + (s.marginTrend === "Declining" ? 1 : 0) },
        { name: "Adjacent segment entry", rationale: "The core capability set transfers to a neighboring customer segment with limited incremental investment.", impact: "High", ease: "Medium", boost: (s) => (s.goal === "Growth strategy" || s.goal === "Market entry" ? 2 : 0) },
        { name: "Geographic expansion", rationale: "Demand exists in adjacent regions the current footprint does not serve.", impact: "Medium", ease: "Low", boost: (s) => (s.goal === "Market entry" ? 2 : 0) },
        { name: "Mix shift toward higher margin volume", rationale: "Replacing low margin volume with higher value business raises profitability without raising cost.", impact: "High", ease: "Medium", boost: (s) => (s.marginTrend === "Declining" ? 2 : 0) + (s.capacity === "Excess" ? 2 : 0) + (s.revenueDriver === "Price" ? 1 : 0) },
        { name: "Defend the core against attack", rationale: "Intensifying competition makes retention of profitable core customers the first priority.", impact: "High", ease: "Medium", boost: (s) => (s.competition === "Intensifying" ? 3 : 0) },
      ]},
      { label: "Portfolio moves", slideTitle: "Portfolio moves reallocate capital from legacy positions to where it earns", slideTakeaway: "Fund the agenda by pruning what no longer earns its capital", items: [
        { name: "Core focus and simplification", rationale: "Complexity in the tail of the portfolio absorbs management attention disproportionate to its profit.", impact: "Medium", ease: "Medium", boost: (s) => (s.goal === "Portfolio review" ? 2 : 0) },
        { name: "Partnership or alliance", rationale: "A partner supplies the missing capability faster than building it internally.", impact: "Medium", ease: "Medium", boost: (s) => (s.urgency === "Weeks" ? 1 : 0) },
        { name: "Divest noncore assets", rationale: "Legacy positions may be worth more to another owner; proceeds fund the growth agenda.", impact: "High", ease: "Low", boost: (s) => (s.goal === "Portfolio review" ? 2 : 0) + (s.revenueDriver === "Declining" ? 1 : 0) },
        { name: "Right size capacity to the strategy", rationale: "Excess capacity ties up capital and management attention that the growth agenda needs.", impact: "Medium", ease: "Medium", boost: (s) => (s.capacity === "Excess" ? 3 : -6) },
        { name: "Exit structurally declining segments", rationale: "Segments in structural decline consume investment that compounds better elsewhere.", impact: "High", ease: "Low", boost: (s) => (s.revenueDriver === "Declining" ? 3 : -6) },
      ]},
    ],
  },

  operations: {
    label: "Operations",
    blurb: "Efficiency, supply chain, cost reduction, workflows",
    guideTitle: "Operations Interview Guide",
    deckTitle: "Operational improvement assessment",
    deckShort: "Operational Improvement Presentation",
    frame: (deal) => {
      const fc = frameFromContext(deal);
      if (fc) return fc + " This operational review targets efficiency, cost, and throughput improvements.";
      const fp = focusPhrase(deal.context);
      return `The client has engaged an operational review of ${deal.company} to identify efficiency, cost, and throughput improvements${fp ? `, with a focus on ${fp.replace(/\.$/, "")}` : ""}.`;
    },
    overview: (deal, m) =>
      `The operating model, cost structure, and supply chain of ${deal.company} are the units of analysis; financial results are treated as symptoms of operational drivers.`,
    defaultRisk: {
      lead: "Operational efficiency",
      text: "The cost base has not been benchmarked against peers; margin trends suggest addressable inefficiency.",
    },
    emailRisk: {
      lead: "Primary risk",
      text: "Savings estimates are top down until process level data is gathered; bottom up validation could move totals materially.",
    },
    baseSections: [
      { section: "Objectives and context", questions: [
        "What prompted this operational review now, and what outcome would make it a success?",
        "Which operational problems does leadership already agree on, and where is there debate?",
        "What constraints (customer commitments, labor agreements, capital) limit the solution space?",
      ]},
      { section: "Operating model overview", questions: [
        "Walk us through the end to end flow from order to delivery: where does it slow down?",
        "Which sites or lines run at capacity, and which are underutilized?",
        "What are the three metrics you review daily, and what are they telling you?",
      ]},
      { section: "Process and quality", questions: [
        "Where are the largest sources of rework, scrap, or error?",
        "Which processes are documented and standardized vs. dependent on individuals?",
        "What continuous improvement programs exist, and what have they delivered?",
      ]},
      { section: "Supply chain and procurement", questions: [
        "Who are the top five suppliers by spend, and when were their contracts last renegotiated?",
        "How many days of inventory do you hold, and what drives it?",
        "Which supply relationships are single sourced, and what is the backup plan?",
      ]},
      { section: "Cost structure", questions: [
        "Which costs are fixed vs. variable, and where is there discretion?",
        "If you had to take out 10% of cost in a year, where would you go first, and what stops you today?",
        "Which cost lines have grown faster than revenue, and why?",
      ]},
    ],
    conditionalQuestions: [
      { topic: 1, cond: (s) => s.capacity === "Excess", q: "Which parts of the network are sized for volume you no longer carry, and what does that idle capacity cost per quarter?" },
      { topic: 4, cond: (s) => s.costProgram === "Yes", q: "How is the current cost program tracking against its target, and where is capture behind plan?" },
      { topic: 1, cond: (s) => s.marginTrend === "Declining", q: "Cost has outrun revenue: which lines moved most, and which were deliberate choices vs. drift?" },
      { topic: 3, cond: (s) => s.concentration === "Yes", q: "How much of the network and inventory is dedicated to your largest customers, and what happens to that cost if their volume changes?" },
      { topic: 4, cond: (s) => s.urgency === "Weeks", q: "Which savings could be actioned within 90 days without structural change?" },
    ],
    groups: [
      { label: "Cost reduction", slideTitle: "Cost reduction levers are ranked for this specific situation", slideTakeaway: "Cost levers are within management's direct control; sequence the highest scoring first", items: [
        { name: "Procurement renegotiation", rationale: "Concentrated supplier spend with dated contracts supports renegotiated rates and consolidation.", impact: "High", ease: "High", boost: (s) => (s.costProgram === "No" ? 1 : 0) },
        { name: "Process and lean improvement", rationale: "Rework and unstandardized processes absorb labor hours that standard work releases.", impact: "Medium", ease: "Medium", boost: () => 0 },
        { name: "Footprint and network optimization", rationale: "Underutilized sites carry fixed cost the network no longer needs; closures must track the volume glide down.", impact: "High", ease: "Low", boost: (s) => (s.capacity === "Excess" ? 3 : 0) },
        { name: "Accelerate the existing cost program", rationale: "A program is already underway; the fastest value is closing the gap between announced savings and captured savings.", impact: "High", ease: "Medium", boost: (s) => (s.costProgram === "Yes" ? 3 : -6) },
        { name: "Overhead and span of control reset", rationale: "Declining margin with stable revenue points to indirect cost growth that a spans and layers review addresses.", impact: "Medium", ease: "Medium", boost: (s) => (s.marginTrend === "Declining" ? 2 : 0) },
      ]},
      { label: "Performance improvement", slideTitle: "Performance levers raise throughput and margin quality without new capacity", slideTakeaway: "Throughput and mix gains defer capital spend; start where data already exists", items: [
        { name: "Demand planning and forecasting", rationale: "Forecast error drives both stockouts and excess inventory; better planning releases working capital.", impact: "Medium", ease: "Medium", boost: () => 0 },
        { name: "Targeted automation", rationale: "High volume repetitive steps are automatable with proven technology and rapid payback.", impact: "Medium", ease: "Medium", boost: (s) => (s.capacity === "Tight" ? 1 : 0) },
        { name: "Quality system uplift", rationale: "Reducing defect rates lowers cost of poor quality and warranty exposure simultaneously.", impact: "Medium", ease: "High", boost: () => 0 },
        { name: "Capacity rebalancing across the network", rationale: "Constrained sites cap revenue while others sit idle; rebalancing lifts throughput without capex.", impact: "High", ease: "Medium", boost: (s) => (s.capacity === "Tight" ? 3 : 0) },
        { name: "Revenue per unit and mix discipline", rationale: "Price and mix are carrying growth; protecting that gain operationally matters as much as cost.", impact: "High", ease: "Medium", boost: (s) => (s.revenueDriver === "Price" ? 3 : 0) + (s.marginTrend === "Declining" ? 1 : 0) },
      ]},
    ],
  },

  financial: {
    label: "Financial / M&A",
    blurb: "M&A advisory, divestitures, synergies, deal value",
    guideTitle: "Divestiture Interview Guide",
    deckTitle: "Qualitative synergy assessment",
    deckShort: "Qualitative Synergy Presentation",
    frame: null, // deal type specific, built in buildNarrative
    overview: (deal, m) =>
      "The analysis focuses on quality of earnings, deal perimeter, and value creation levers.",
    defaultRisk: null, // deal type specific
    emailRisk: null, // deal type specific
    baseSections: null, // deal type specific, built in buildInterviewQuestions
    conditionalQuestions: [
      { topic: 2, cond: (s) => s.concentration === "Yes", q: "What share of revenue sits with the top five customers, and how would the transaction affect those relationships?" },
      { topic: 3, cond: (s) => s.capacity === "Excess", q: "How much capacity would the combined or separated business carry relative to its volume, and who bears the stranded cost?" },
      { topic: 2, cond: (s) => s.costProgram === "Yes", q: "How much of the announced cost program is already reflected in the numbers, and how much is still to be captured?" },
      { topic: 2, cond: (s) => s.marginTrend === "Declining", q: "Walk us through the margin bridge over the last two years: what portion is structural vs. one time?" },
    ],
    groups: [
      { label: "Revenue synergies", slideTitle: "Revenue synergies are ranked by fit with this transaction", slideTakeaway: "Revenue synergies carry more execution risk than cost synergies; treat as upside, not base case", items: [
        { name: "Cross sell into the combined customer base", rationale: "Customers of each business can be offered the other's adjacent products.", impact: "High", ease: "Medium", boost: () => 0 },
        { name: "Geographic and channel expansion", rationale: "The combined footprint opens regions and channels neither business serves alone.", impact: "Medium", ease: "Medium", boost: () => 0 },
        { name: "Pricing harmonization", rationale: "Aligning price levels and discount discipline across the combined book lifts realized price.", impact: "Medium", ease: "High", boost: (s) => (s.revenueDriver === "Price" ? 2 : 0) },
        { name: "Key account retention program", rationale: "Concentrated revenue makes protecting the largest relationships through close the first revenue priority.", impact: "High", ease: "Medium", boost: (s) => (s.concentration === "Yes" ? 3 : 0) },
      ]},
      { label: "Cost synergies", slideTitle: "Cost synergies anchor the value case and are within management's control", slideTakeaway: "Cost synergies should anchor the value case; sequence procurement first", items: [
        { name: "Procurement scale and vendor consolidation", rationale: "Combined purchasing volume supports renegotiated rates with overlapping vendors.", impact: "High", ease: "High", boost: () => 0 },
        { name: "Overlapping G&A rationalization", rationale: "Duplicated back office functions can be consolidated over 12 to 24 months.", impact: "Medium", ease: "Medium", boost: () => 0 },
        { name: "Facility and footprint optimization", rationale: "Overlapping or underused facilities carry fixed cost the combined network does not need.", impact: "Medium", ease: "Low", boost: (s) => (s.capacity === "Excess" ? 3 : 0) },
        { name: "Margin improvement to peer levels", rationale: "Profitability trails typical sector levels, an addressable opportunity independent of the deal.", impact: "High", ease: "Medium", boost: (s, m) => (m.ebitdaMargin !== null && m.ebitdaMargin < 0.12 ? 3 : 0) + (s.marginTrend === "Declining" ? 1 : 0) },
        { name: "Integrate the existing cost program", rationale: "An announced cost program must be folded into the synergy case to avoid double counting and to protect capture.", impact: "Medium", ease: "Medium", boost: (s) => (s.costProgram === "Yes" ? 3 : -6) },
      ]},
    ],
  },
};

function practiceOf(deal) {
  return PRACTICES[deal.practice] || PRACTICES.financial;
}

const PRACTICE_TERMS = { strategy: "opportunity", operations: "improvement", financial: "synergy" };

// The M&A guide is named for the deal type actually selected
function guideTitleOf(deal) {
  const P = practiceOf(deal);
  return P === PRACTICES.financial ? `${deal.dealType} Interview Guide` : P.guideTitle;
}

function practiceTerm(deal) {
  return PRACTICE_TERMS[deal.practice] || "synergy";
}

/* ---------- Situation risk (shared by summary, deck, email) ---------- */

function situationRisk(deal, m) {
  const s = situationOf(deal);
  if (s.concentration === "Yes")
    return { lead: "Customer concentration", text: "A significant share of revenue depends on a small number of customers; quantify the exposure and the retention plan before anything else." };
  if (s.capacity === "Excess")
    return { lead: "Stranded capacity", text: "The network is sized for volume it no longer carries; cost must come out at least as fast as volume leaves or margin will keep compressing." };
  if (s.marginTrend === "Declining" || (m.ebitdaMargin !== null && m.ebitdaMargin < 0.08))
    return { lead: "Margin compression", text: "Profitability is moving the wrong way; build the margin bridge to separate structural pressure from one time items." };
  if (s.competition === "Intensifying")
    return { lead: "Competitive pressure", text: "Competitors are moving; every recommendation must survive a realistic view of their response." };
  if (s.costProgram === "Yes")
    return { lead: "Cost program capture", text: "Announced savings are not captured savings; track the program at the initiative level to protect the value." };
  const P = practiceOf(deal);
  if (P.defaultRisk) return P.defaultRisk;
  const separation = deal.dealType === "Divestiture" || deal.dealType === "Carve out";
  return separation
    ? { lead: "Separation complexity", text: "Entangled systems, shared services, and TSA scope are the largest unquantified risks; management interviews should map them first." }
    : { lead: "Integration readiness", text: "Operational dependencies and management retention determine speed to value; both are unconfirmed at this stage." };
}

/* ---------- Value scenarios (per line ranges, three scenarios) ---------- */

function buildScenarioInputs(deal, m) {
  const lines = [];
  if (m.series.cogs) {
    // Range depends on the cost structure, not a fixed guess: thin
    // margin businesses (retail) cannot cut COGS 8%; service and
    // software businesses have little COGS to cut at all.
    const cogsShare = m.series.cogs[m.series.cogs.length - 1] / m.revenueLatest;
    const [low, high] = cogsShare > 0.75 ? [0.01, 0.03] : cogsShare < 0.35 ? [0.03, 0.05] : [0.05, 0.08];
    lines.push({ key: "cogs", label: "Cost of goods sold", base: m.series.cogs[m.series.cogs.length - 1], low, high });
  }
  if (m.series.opex) {
    lines.push({ key: "opex", label: "Opex / G&A (back office)", base: m.series.opex[m.series.opex.length - 1], low: 0.1, high: 0.25 });
  }
  // Revenue uplift flows through at the company margin, but a negative
  // or unknown margin gets a flagged default instead of nonsense.
  const flowDefaulted = m.ebitdaMargin === null || m.ebitdaMargin < 0.05;
  return {
    revenue: m.revenueLatest,
    revLow: 0.05,
    revHigh: 0.1,
    flowMargin: flowDefaulted ? 0.15 : Math.round(m.ebitdaMargin * 1000) / 1000,
    flowDefaulted,
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
  const s = situationOf(deal);

  let growthSentence =
    m.revenueCAGR > 0
      ? `Revenue grew from ${fmtM(m.series.revenue[0])} in ${m.firstYear} to ${fmtM(m.revenueLatest)} in ${m.lastYear}, a ${m.growthWord} ${fmtPct(m.revenueCAGR)} compound annual growth rate.`
      : `Revenue declined from ${fmtM(m.series.revenue[0])} in ${m.firstYear} to ${fmtM(m.revenueLatest)} in ${m.lastYear}, an area for focused diligence.`;
  if (s.revenueDriver === "Price") {
    growthSentence += " Management attributes the growth to price and mix rather than volume; validate its durability.";
  } else if (s.revenueDriver === "Declining") {
    growthSentence += " Management reports softening demand behind the headline figures.";
  }

  let marginSentence = m.ebitdaMargin !== null
    ? m.ebitdaMargin < 0
      ? `The business is loss making at the EBITDA line (${fmtPct(m.ebitdaMargin)} margin in ${m.lastYear}); stabilizing the cost base comes before any value creation agenda.`
      : m.ebitdaMargin > 0.2
      ? `EBITDA margin of ${fmtPct(m.ebitdaMargin)} is strong for the sector and suggests durable pricing power or cost discipline.`
      : m.ebitdaMargin > 0.1
      ? `EBITDA margin of ${fmtPct(m.ebitdaMargin)} is within a typical range, with room for operational improvement after close.`
      : `EBITDA margin of ${fmtPct(m.ebitdaMargin)} is thin, making cost structure a priority diligence area.`
    : "Profitability could not be derived from the provided statements; request a full income statement.";
  if (s.marginTrend === "Declining" && m.marginDirection === "rising") {
    marginSentence += " Note: management describes margin as declining while the provided statements show it rising; reconcile the two views early.";
  } else if (s.marginTrend === "Declining") {
    marginSentence += " Management confirms margin is under pressure; the margin bridge is the first analysis to build.";
  }

  const P = practiceOf(deal);
  const fp = focusPhrase(deal.context);
  let dealFrame = P.frame
    ? P.frame(deal)
    : frameFromContext(deal) ||
      {
        "Acquisition": `The client is evaluating the acquisition of ${deal.company}${fp ? `, with a focus on ${fp.replace(/\.$/, "")}` : " to expand its market position"}.`,
        "Divestiture": `The client is preparing ${deal.company} for divestiture and needs a clear view of separation complexity and standalone economics.`,
        "Carve out": `${deal.company} is being carved out of its parent, so entanglements and transition services drive both risk and value.`,
        "Merger": `The contemplated merger with ${deal.company} turns on integration feasibility and the credibility of the combined entity synergy case.`,
      }[deal.dealType];
  if (s.goal) dealFrame += ` Stated priority: ${s.goal.toLowerCase()}.`;

  return { growthSentence, marginSentence, dealFrame };
}

/* ---------- Key findings ---------- */

function buildKeyFindings(deal, m) {
  const n = buildNarrative(deal, m);
  const risk = situationRisk(deal, m);
  const opps = buildOpportunities(deal, m);
  const top = opps.ranked[0];
  const term = practiceTerm(deal);

  return [
    { lead: "Top line", text: n.growthSentence + " See Section 2." },
    { lead: "Profitability", text: n.marginSentence + " See Section 2." },
    { lead: risk.lead, text: risk.text },
    {
      lead: `${term.charAt(0).toUpperCase() + term.slice(1)} anchor`,
      text: `${top.name} scores highest for this situation on impact and achievability; see the accompanying presentation for the ranked view.`,
    },
  ];
}

/* ---------- Interview questions ---------- */

function buildInterviewQuestions(deal) {
  const P = practiceOf(deal);
  const s = situationOf(deal);

  let sections;
  if (P.baseSections) {
    sections = P.baseSections.map((sec) => ({ section: sec.section, questions: [...sec.questions] }));
  } else {
    sections = financialBaseSections(deal);
  }

  // Situation triggered questions slot into their topics, capped at 20 total.
  let total = sections.reduce((t, sec) => t + sec.questions.length, 0);
  for (const cq of P.conditionalQuestions || []) {
    if (total >= 20) break;
    if (cq.cond(s) && sections[cq.topic]) {
      sections[cq.topic].questions.push(cq.q);
      total += 1;
    }
  }
  return sections;
}

function financialBaseSections(deal) {
  const separation = deal.dealType === "Divestiture" || deal.dealType === "Carve out";

  const motivation = {
    section: "Motivation and objectives",
    questions: separation
      ? [
          `What is driving the decision to ${deal.dealType === "Divestiture" ? "divest" : "carve out"} this business, and why now?`,
          "What does a successful outcome look like: price, speed, certainty, or ongoing relationship?",
          "Have you explored alternatives (sale, spin, joint venture, wind down), and what ruled them out?",
        ]
      : [
          `What is the strategic rationale for this ${deal.dealType.toLowerCase()}, and why now?`,
          "What does a successful outcome look like twelve months after close?",
          "Which alternatives were considered, and what made this path the preferred one?",
        ],
  };

  const overview = {
    section: "Business overview and perimeter",
    questions: [
      `Walk us through ${deal.company}'s operating model: key products, customer segments, and how revenue is generated.`,
      "Which products, sites, contracts, and people sit inside the transaction perimeter?",
      "Which shared functions or systems does the business rely on today?",
    ],
  };

  const financials = {
    section: "Financial performance",
    questions: [
      "What drove the revenue trend over the historical period, and how much is volume vs. price?",
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

/* ---------- Opportunity selection (scored library) ---------- */

// Scores every library item against the situation and metrics, keeps
// the top three per group. Returns { groups, ranked } where ranked is
// all selected items ordered by total score.
function buildOpportunities(deal, m) {
  const P = practiceOf(deal);
  const s = situationOf(deal);
  const impactScore = { High: 3, Medium: 2, Low: 1 };

  const groups = P.groups.map((g) => {
    const scored = g.items.map((item, i) => ({
      ...item,
      score:
        impactScore[item.impact] + impactScore[item.ease] +
        (item.boost ? item.boost(s, m) : 0) -
        i * 0.01, // stable tie break by library order
    }));
    scored.sort((a, b) => b.score - a.score);
    return {
      label: g.label,
      slideTitle: g.slideTitle,
      slideTakeaway: g.slideTakeaway,
      items: scored.slice(0, 3),
    };
  });

  const ranked = groups
    .flatMap((g) => g.items)
    .slice()
    .sort((a, b) => b.score - a.score);

  return { groups, ranked };
}

/* ---------- Verdict: the recommendation follows the numbers ---------- */

// proceed | caution | pause. High COGS businesses (retail, distribution)
// are structurally thin margin, so their caution floor sits lower.
function verdictOf(m) {
  if (m.revenueCAGR < -0.05 || (m.ebitdaMargin !== null && m.ebitdaMargin < 0)) return "pause";
  const cogsShare = m.series.cogs ? m.series.cogs[m.series.cogs.length - 1] / m.revenueLatest : null;
  const marginFloor = cogsShare !== null && cogsShare > 0.75 ? 0.015 : 0.05;
  if (m.revenueCAGR < 0 || (m.ebitdaMargin !== null && m.ebitdaMargin < marginFloor)) return "caution";
  return "proceed";
}

/* ---------- Email summary (BLUF: answer first, then support) ---------- */

function buildEmail(deal, m) {
  const s = situationOf(deal);
  const marginText =
    m.ebitdaMargin !== null ? fmtPct(m.ebitdaMargin) + " EBITDA margin" : "profitability to be confirmed";
  const P = practiceOf(deal);
  const term = practiceTerm(deal);
  const risk = situationRisk(deal, m);
  const opps = buildOpportunities(deal, m);
  const top = opps.ranked[0];
  const verdict = verdictOf(m);
  const thesis = P.frame ? "engagement" : deal.dealType.toLowerCase();

  const deadline =
    s.urgency === "Weeks" ? "by end of week" : s.urgency === "Exploratory" ? "when convenient this month" : "within two weeks";

  const subjects = {
    proceed: `${deal.company}: recommend advancing to management interviews; ${fmtM(m.revenueLatest)} revenue, ${marginText}`,
    caution: `${deal.company}: proceed with caution; ${fmtM(m.revenueLatest)} revenue, ${marginText}`,
    pause: `${deal.company}: recommend a focused diagnostic before advancing; ${fmtM(m.revenueLatest)} revenue, ${marginText}`,
  };
  const answers = {
    proceed: `Recommendation: advance ${deal.company} to the management interview phase, leading with ${top.name.toLowerCase()}. The financial profile supports the ${thesis} thesis, subject to the diligence items below.`,
    caution: `Recommendation: proceed to management interviews with caution. Growth or profitability is below par, so the agenda below is built to test whether the ${thesis} thesis still holds before further commitment.`,
    pause: `Recommendation: pause before advancing. The provided statements show ${m.revenueCAGR < -0.05 ? "declining revenue" : "losses at the EBITDA line"}, which the ${thesis} thesis cannot yet support; commission a focused margin and cash diagnostic first, then revisit.`,
  };
  const soWhats = {
    proceed: `This matters now because the historical financials show ${m.growthWord} growth (${fmtPct(m.revenueCAGR)} CAGR) with ${marginText}, and the open questions are answerable within a two week interview window.`,
    caution: `The historical financials show ${m.growthWord} growth (${fmtPct(m.revenueCAGR)} CAGR) with ${marginText}; the value case depends on interview answers, not on momentum.`,
    pause: `The historical financials show ${m.growthWord} growth (${fmtPct(m.revenueCAGR)} CAGR) with ${marginText}; committing before a diagnostic risks anchoring the work on a broken baseline.`,
  };

  // Sensitivity anchored to what the model actually computes
  const sc = computeScenarios(deal, m);
  const bigLine = sc.inputs.lines.slice().sort((a, b) => b.base - a.base)[0];
  const sensitivity = bigLine
    ? `One percentage point of ${term} on ${bigLine.label.toLowerCase()} is worth about ${fmtM(bigLine.base * 0.01)} a year; validating the per line ranges is the most valuable interview outcome.`
    : `One percentage point of revenue uplift at the assumed flow through margin is worth about ${fmtM(m.revenueLatest * 0.01 * sc.inputs.flowMargin)} a year; validating that margin is the most valuable interview outcome.`;

  return {
    subject: subjects[verdict],
    answer: answers[verdict],
    soWhat: soWhats[verdict],
    ask: `Decision needed: approve ${verdict === "pause" ? "the diagnostic scope" : "interview scheduling and data room access"} ${deadline} to hold the timeline.`,
    points: [
      {
        lead: "Directional value",
        text: `The model shows annual ${term} value of ${fmtM(sc.conservative.total)} conservative to ${fmtM(sc.aggressive.total)} aggressive (${fmtM(sc.midpoint.total)} midpoint), on ${fmtM(m.revenueLatest)} of ${m.lastYear} revenue.${sc.inputs.flowDefaulted ? " Note: revenue uplift uses a defaulted 15% flow through margin because the company margin is unavailable or negative; treat that line as illustrative." : ""} Every percentage sits in a labeled input cell, so the assumptions can be flexed directly.`,
      },
      { lead: "Key sensitivity", text: sensitivity },
      { lead: risk.lead, text: risk.text },
    ],
    nextSteps: [
      "Conduct management interviews using the attached guide (weeks 1 to 2).",
      `Size the top three ${term} opportunities bottom up (weeks 2 to 3).`,
      "Fold validated findings into the model and reconvene (week 4).",
    ],
    close: "Full analysis in the attached model, summary, and interview guide. Happy to walk through the assumptions.",
  };
}
