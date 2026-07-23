/* ============ DealDesk company lookup ============
   Free, keyless, browser safe research from Wikipedia and Wikidata
   (both send open CORS headers). Returns a structured profile:
   description, key people, headquarters, founding year, employees,
   industry, and where Wikidata has it, revenue by year in USD.
   No AI and no server; public structured data only. */

const WD_API = "https://www.wikidata.org/w/api.php";
const WP_SUMMARY = "https://en.wikipedia.org/api/rest_v1/page/summary/";

// House rule: no dashes in anything that lands in a document.
function softenLookupText(s) {
  return String(s)
    .replace(/(\d)\s*[-–—]\s*(\d)/g, "$1 to $2")
    .replace(/[—–]/g, ", ")
    .replace(/-/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Wikimedia asks API clients to identify themselves. Browsers cannot
// set User-Agent, so the supported Api-User-Agent header is used; in
// Node both are set.
const LOOKUP_UA = "DealDesk/1.0 (https://joshuakeum9-cell.github.io/dealdesk/; joshuakeum9@gmail.com)";
const LOOKUP_HEADERS = { Accept: "application/json", "Api-User-Agent": LOOKUP_UA };
if (typeof window === "undefined") LOOKUP_HEADERS["User-Agent"] = LOOKUP_UA;

async function wdGet(params) {
  const qs = new URLSearchParams({ format: "json", origin: "*", ...params });
  const res = await fetch(WD_API + "?" + qs, { headers: LOOKUP_HEADERS });
  if (!res.ok) throw new Error("wikidata " + res.status);
  return res.json();
}

function claimValue(claims, prop) {
  const c = claims[prop];
  if (!c || !c.length) return null;
  return c[0].mainsnak && c[0].mainsnak.datavalue ? c[0].mainsnak.datavalue.value : null;
}

// Wikidata lists every historical officeholder. The current one is
// the statement with preferred rank, or failing that, one without an
// end time qualifier (P582).
function currentPersonId(claims, prop) {
  const sts = (claims[prop] || []).filter((st) => st.mainsnak && st.mainsnak.datavalue && st.mainsnak.datavalue.value.id);
  if (!sts.length) return null;
  const preferred = sts.find((st) => st.rank === "preferred");
  if (preferred) return preferred.mainsnak.datavalue.value.id;
  const open = sts.find((st) => !(st.qualifiers && st.qualifiers.P582));
  return (open || sts[sts.length - 1]).mainsnak.datavalue.value.id;
}

// Quantities like employee count carry point in time qualifiers; take
// the preferred statement or the most recent one.
function latestQuantity(claims, prop) {
  const sts = (claims[prop] || []).filter((st) => st.mainsnak && st.mainsnak.datavalue);
  if (!sts.length) return null;
  const preferred = sts.find((st) => st.rank === "preferred");
  const timeOf = (st) => {
    const q = st.qualifiers && st.qualifiers.P585 && st.qualifiers.P585[0];
    return q && q.datavalue ? q.datavalue.value.time : "";
  };
  const pick = preferred || sts.slice().sort((a, b) => timeOf(a).localeCompare(timeOf(b))).pop();
  const v = pick.mainsnak.datavalue.value;
  return v && v.amount ? Math.round(Number(v.amount)) : null;
}

// P2139 revenue: statements carry an amount, a currency unit, and a
// "point in time" qualifier. Keep USD only, one value per year.
function revenueSeriesFromClaims(claims) {
  const rows = [];
  for (const st of claims.P2139 || []) {
    const dv = st.mainsnak && st.mainsnak.datavalue && st.mainsnak.datavalue.value;
    if (!dv || !/Q4917$/.test(dv.unit || "")) continue; // USD only
    const t = st.qualifiers && st.qualifiers.P585 && st.qualifiers.P585[0];
    const time = t && t.datavalue && t.datavalue.value.time;
    const year = time && Number(time.slice(1, 5));
    const amount = Number(dv.amount);
    if (!year || !Number.isFinite(amount) || amount <= 0) continue;
    rows.push({ year, m: amount / 1e6 });
  }
  const byYear = new Map();
  for (const r of rows) byYear.set(r.year, r.m);
  const years = [...byYear.keys()].sort((a, b) => a - b).slice(-4);
  if (years.length < 2) return null;
  return { years, revenue: years.map((y) => +byYear.get(y).toFixed(1)) };
}

async function lookupCompany(name) {
  // 1. Find the entity
  const search = await wdGet({ action: "wbsearchentities", search: name, language: "en", type: "item", limit: 5 });
  const hit = (search.search || []).find((h) => h.description && /compan|corporation|business|enterprise|manufacturer|retailer|conglomerate|bank|airline|chain/i.test(h.description)) || (search.search || [])[0];
  if (!hit) return null;

  // 2. Entity claims plus sitelink for the Wikipedia article
  const ent = await wdGet({ action: "wbgetentities", ids: hit.id, props: "claims|sitelinks|descriptions" });
  const entity = ent.entities[hit.id];
  const claims = entity.claims || {};

  // 3. Resolve labels for referenced entities (CEO, HQ, industry)
  const refIds = [];
  const refOf = (prop, n) =>
    (claims[prop] || []).slice(0, n).map((st) => {
      const v = st.mainsnak && st.mainsnak.datavalue && st.mainsnak.datavalue.value;
      return v && v.id ? v.id : null;
    }).filter(Boolean);
  const ceoIds = [currentPersonId(claims, "P169")].filter(Boolean);
  const chairIds = [currentPersonId(claims, "P488")].filter(Boolean);
  const hqIds = refOf("P159", 1);
  const industryIds = refOf("P452", 2);
  refIds.push(...ceoIds, ...chairIds, ...hqIds, ...industryIds);
  let labels = {};
  if (refIds.length) {
    const lab = await wdGet({ action: "wbgetentities", ids: [...new Set(refIds)].join("|"), props: "labels", languages: "en" });
    for (const [id, e] of Object.entries(lab.entities || {})) {
      labels[id] = e.labels && e.labels.en ? e.labels.en.value : null;
    }
  }

  // 4. Wikipedia lead paragraph
  let description = null, wikiUrl = null;
  const title = entity.sitelinks && entity.sitelinks.enwiki && entity.sitelinks.enwiki.title;
  if (title) {
    try {
      const res = await fetch(WP_SUMMARY + encodeURIComponent(title), { headers: LOOKUP_HEADERS });
      if (res.ok) {
        const sum = await res.json();
        if (sum.extract) {
          const sentences = sum.extract.split(/(?<=[.!?])\s+/).slice(0, 2).join(" ");
          description = softenLookupText(sentences);
          wikiUrl = sum.content_urls && sum.content_urls.desktop ? sum.content_urls.desktop.page : null;
        }
      }
    } catch (e) { /* description stays null */ }
  }

  const founded = (() => {
    const v = claimValue(claims, "P571");
    return v && v.time ? Number(v.time.slice(1, 5)) : null;
  })();
  const employees = latestQuantity(claims, "P1128");

  return {
    id: hit.id,
    label: hit.label || name,
    shortDescription: hit.description ? softenLookupText(hit.description) : null,
    description,
    wikiUrl,
    ceo: ceoIds.length ? labels[ceoIds[0]] : null,
    chair: chairIds.length ? labels[chairIds[0]] : null,
    hq: hqIds.length ? labels[hqIds[0]] : null,
    industries: industryIds.map((i) => labels[i]).filter(Boolean).map(softenLookupText),
    founded,
    employees,
    revenueSeries: revenueSeriesFromClaims(claims),
  };
}
