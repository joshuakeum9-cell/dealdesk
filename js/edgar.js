/* ============ DealDesk live filing data (SEC EDGAR) ============
   Pulls real quarterly figures out of a company's own SEC filings,
   in the browser, with no account and no cost.

   SEC serves the submissions endpoint with open CORS but not the XBRL
   endpoints, so requests try direct first and then fall through a list
   of free public relays. Everything successful is cached in
   localStorage, so repeat lookups are instant and a relay outage is
   survivable. If every route fails the caller falls back to manual
   entry: the site never blocks on someone else's service. */

const EDGAR_CACHE_PREFIX = "dd_edgar_";
const EDGAR_TTL_MS = 12 * 60 * 60 * 1000; // filings change daily at most

const EDGAR_ROUTES = [
  (u) => u, // direct: works for the submissions endpoint
  (u) => "https://corsproxy.io/?url=" + encodeURIComponent(u),
  (u) => "https://api.allorigins.win/raw?url=" + encodeURIComponent(u),
  (u) => "https://api.codetabs.com/v1/proxy?quest=" + encodeURIComponent(u),
];

function edgarCacheGet(key) {
  try {
    const raw = localStorage.getItem(EDGAR_CACHE_PREFIX + key);
    if (!raw) return null;
    const { t, v } = JSON.parse(raw);
    if (Date.now() - t > EDGAR_TTL_MS) return null;
    return v;
  } catch (e) { return null; }
}

function edgarCacheSet(key, value) {
  try {
    localStorage.setItem(EDGAR_CACHE_PREFIX + key, JSON.stringify({ t: Date.now(), v: value }));
  } catch (e) { /* quota: caching is an optimization, not a requirement */ }
}

// The route that last worked is tried first, so a session pays the
// discovery cost once rather than on every request.
let EDGAR_BEST_ROUTE = 0;

async function edgarFetchJson(url, { timeoutMs = 6000 } = {}) {
  const cached = edgarCacheGet(url);
  if (cached !== null) {
    if (cached === "__MISS__") throw new Error("known miss");
    return cached;
  }
  const order = [EDGAR_BEST_ROUTE, ...EDGAR_ROUTES.map((_, i) => i).filter((i) => i !== EDGAR_BEST_ROUTE)];
  let lastErr;
  for (const idx of order) {
    try {
      const res = await fetch(EDGAR_ROUTES[idx](url), { signal: AbortSignal.timeout(timeoutMs) });
      if (res.status === 404) {
        // A tag this filer does not report: remember it so the next
        // lookup does not pay for the same discovery again.
        edgarCacheSet(url, "__MISS__");
        throw new Error("not reported");
      }
      if (!res.ok) { lastErr = new Error("status " + res.status); continue; }
      const data = await res.json();
      EDGAR_BEST_ROUTE = idx;
      edgarCacheSet(url, data);
      return data;
    } catch (e) {
      if (String(e.message) === "not reported") throw e;
      lastErr = e;
    }
  }
  throw lastErr || new Error("unreachable");
}

/* ---------- Company identity ---------- */

// The ticker file is about a megabyte, so it is fetched once and kept.
async function edgarTickerMap() {
  const data = await edgarFetchJson("https://www.sec.gov/files/company_tickers.json");
  return Object.values(data);
}

function padCik(cik) {
  return "CIK" + String(cik).padStart(10, "0");
}

// Accepts a ticker or a company name and resolves to a filer.
async function edgarResolveCompany(query) {
  const q = String(query || "").trim();
  if (!q) return null;
  const rows = await edgarTickerMap();
  const upper = q.toUpperCase();
  const norm = (s) => String(s).toUpperCase().replace(/[^A-Z0-9 ]/g, "").replace(/\s+/g, " ").trim();
  const nq = norm(q);

  let hit = rows.find((r) => String(r.ticker).toUpperCase() === upper);
  if (!hit) hit = rows.find((r) => norm(r.title) === nq);
  if (!hit) {
    // Prefer the shortest title that contains the query, which favors
    // the parent company over subsidiaries and trusts.
    const contains = rows.filter((r) => norm(r.title).includes(nq));
    contains.sort((a, b) => a.title.length - b.title.length);
    hit = contains[0];
  }
  if (!hit) {
    const words = nq.split(" ")[0];
    if (words.length >= 4) {
      const starts = rows.filter((r) => norm(r.title).startsWith(words));
      starts.sort((a, b) => a.title.length - b.title.length);
      hit = starts[0];
    }
  }
  return hit ? { cik: hit.cik_str, ticker: hit.ticker, name: hit.title } : null;
}

async function edgarSubmissions(cik) {
  return edgarFetchJson("https://data.sec.gov/submissions/" + padCik(cik) + ".json");
}

/* ---------- Financial concepts ---------- */

const REVENUE_TAGS = [
  "RevenueFromContractWithCustomerExcludingAssessedTax",
  "RevenueFromContractWithCustomerIncludingAssessedTax",
  "Revenues",
  "SalesRevenueNet",
  "SalesRevenueGoodsNet",
];
const OPINCOME_TAGS = [
  "OperatingIncomeLoss",
  "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
  "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
];
const NETINCOME_TAGS = ["NetIncomeLoss"];
const COST_TAGS = ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfServices"];
const OPEX_TAGS = ["OperatingExpenses", "SellingGeneralAndAdministrativeExpense"];

async function edgarConcept(cik, tag) {
  const url = `https://data.sec.gov/api/xbrl/companyconcept/${padCik(cik)}/us-gaap/${tag}.json`;
  try {
    const data = await edgarFetchJson(url);
    return data && data.units && data.units.USD ? data.units.USD : null;
  } catch (e) { return null; }
}

// First tag that actually returns data for this filer
async function edgarFirstConcept(cik, tags) {
  for (const tag of tags) {
    const rows = await edgarConcept(cik, tag);
    if (rows && rows.length) return { tag, rows };
  }
  return null;
}

// Quarter label from the period end date. The fy and fp fields describe
// the filing that reported the fact, not the period itself, so using
// them mislabels restated figures.
function quarterLabel(endIso) {
  const d = new Date(endIso + "T00:00:00Z");
  const y = d.getUTCFullYear();
  const q = Math.floor(d.getUTCMonth() / 3) + 1;
  return { key: `${y}Q${q}`, year: y, q, label: `Q${q} ${y}` };
}

// Reduce raw XBRL facts to one clean row per period.
function edgarPeriodSeries(rows, { minDays, maxDays }) {
  const byEnd = new Map();
  for (const r of rows || []) {
    if (r.val === null || r.val === undefined) continue;
    const days = (new Date(r.end) - new Date(r.start)) / 86400000;
    if (days < minDays || days > maxDays) continue;
    const prev = byEnd.get(r.end);
    // Later filings supersede earlier ones for the same period
    if (!prev || String(r.filed) > String(prev.filed)) byEnd.set(r.end, r);
  }
  return [...byEnd.values()].sort((a, b) => a.end.localeCompare(b.end));
}

const edgarQuarterly = (rows) => edgarPeriodSeries(rows, { minDays: 80, maxDays: 100 });

// Companies report Q4 inside the annual 10-K rather than as its own
// quarter, so it is derived the way an analyst does it: the fiscal
// year total minus the three reported quarters.
function edgarWithDerivedQ4(rows) {
  const q = edgarPeriodSeries(rows, { minDays: 80, maxDays: 100 });
  const a = edgarPeriodSeries(rows, { minDays: 340, maxDays: 380 });
  const out = q.slice();
  for (const yr of a) {
    const yEnd = new Date(yr.end);
    const yStart = new Date(yr.start);
    const inYear = q.filter((r) => new Date(r.start) >= yStart && new Date(r.end) <= yEnd);
    if (inYear.length !== 3) continue; // only safe when exactly Q1 to Q3 are present
    if (out.some((r) => r.end === yr.end)) continue;
    const sum = inYear.reduce((s, r) => s + r.val, 0);
    const lastQEnd = inYear.map((r) => r.end).sort().pop();
    out.push({
      end: yr.end,
      start: lastQEnd,
      val: yr.val - sum,
      form: yr.form,
      filed: yr.filed,
      accn: yr.accn,
      derived: true,
    });
  }
  return out.sort((x, y) => x.end.localeCompare(y.end));
}
const edgarAnnual = (rows) => edgarPeriodSeries(rows, { minDays: 340, maxDays: 380 });

/* ---------- The public call ---------- */

// Returns a normalized profile: identity, quarterly series, annual
// series, and the filings the numbers came from.
async function fetchFilingProfile(query, onProgress) {
  const say = (s) => { if (onProgress) onProgress(s); };

  say("Finding the company in the SEC register");
  const co = await edgarResolveCompany(query);
  if (!co) return null;

  say("Reading " + co.name + " filing history");
  const subsPromise = edgarSubmissions(co.cik).catch(() => null);

  say("Pulling revenue and operating income from filings");
  const [revC, oiC, niC, costC, opexC] = await Promise.all([
    edgarFirstConcept(co.cik, REVENUE_TAGS),
    edgarFirstConcept(co.cik, OPINCOME_TAGS),
    edgarFirstConcept(co.cik, NETINCOME_TAGS),
    edgarFirstConcept(co.cik, COST_TAGS),
    edgarFirstConcept(co.cik, OPEX_TAGS),
  ]);
  if (!revC) return { company: co, error: "no revenue data in filings" };

  const revQ = edgarWithDerivedQ4(revC.rows);
  const oiQ = oiC ? edgarWithDerivedQ4(oiC.rows) : [];
  const niQ = niC ? edgarWithDerivedQ4(niC.rows) : [];
  const oiMap = new Map(oiQ.map((r) => [r.end, r.val]));
  const niMap = new Map(niQ.map((r) => [r.end, r.val]));

  const quarters = revQ.slice(-12).map((r) => {
    const q = quarterLabel(r.end);
    const oi = oiMap.has(r.end) ? oiMap.get(r.end) : null;
    return {
      ...q,
      end: r.end,
      revenue: r.val / 1e6,
      opIncome: oi === null ? null : oi / 1e6,
      netIncome: niMap.has(r.end) ? niMap.get(r.end) / 1e6 : null,
      margin: oi === null || r.val === 0 ? null : oi / r.val,
      form: r.form,
      filed: r.filed,
      accession: r.accn,
      derived: !!r.derived,
    };
  });

  const revA = edgarAnnual(revC.rows).slice(-5);
  const costA = costC ? new Map(edgarAnnual(costC.rows).map((r) => [r.end, r.val])) : null;
  const opexA = opexC ? new Map(edgarAnnual(opexC.rows).map((r) => [r.end, r.val])) : null;
  const oiA = oiC ? new Map(edgarAnnual(oiC.rows).map((r) => [r.end, r.val])) : null;
  const annual = revA.map((r) => ({
    year: new Date(r.end + "T00:00:00Z").getUTCFullYear(),
    end: r.end,
    revenue: r.val / 1e6,
    cost: costA && costA.has(r.end) ? costA.get(r.end) / 1e6 : null,
    opex: opexA && opexA.has(r.end) ? opexA.get(r.end) / 1e6 : null,
    opIncome: oiA && oiA.has(r.end) ? oiA.get(r.end) / 1e6 : null,
  }));

  const subs = await subsPromise;
  const latest = quarters[quarters.length - 1] || null;
  const recentFilings = subs && subs.filings && subs.filings.recent
    ? (() => {
        const f = subs.filings.recent;
        const out = [];
        for (let i = 0; i < Math.min(f.form.length, 60) && out.length < 6; i++) {
          if (["10-K", "10-Q", "8-K"].includes(f.form[i])) {
            out.push({ form: f.form[i], filed: f.filingDate[i], period: f.reportDate[i] });
          }
        }
        return out;
      })()
    : [];

  return {
    company: {
      ...co,
      sic: subs ? subs.sicDescription : null,
      fiscalYearEnd: subs ? subs.fiscalYearEnd : null,
      exchange: subs && subs.exchanges ? subs.exchanges[0] : null,
      stateOfIncorporation: subs ? subs.stateOfIncorporation : null,
    },
    revenueTag: revC.tag,
    quarters,
    annual,
    latest,
    recentFilings,
    asOf: latest ? { period: latest.label, form: latest.form, filed: latest.filed } : null,
  };
}
