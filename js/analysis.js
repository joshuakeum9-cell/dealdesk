/* ============ DealDesk diagnosis engine ============
   Reads a real quarterly series from SEC filings and derives the
   business question, the answer, and a falsifiable test.

   Nothing here is invented: every sentence is built from a computed
   figure, and every figure traces to a filing. Comparisons are year
   over year against the same quarter, which is the only honest way to
   read a seasonal business. */

function pct(x) {
  return (x * 100).toFixed(1) + "%";
}
// Margin differences are in points, never percent
function pp(x) {
  const v = Math.abs(x) * 100;
  return v.toFixed(1) + (v === 1 ? " point" : " points");
}
function titleCase(name) {
  const small = new Set(["and", "of", "the", "for"]);
  // Registration names carry suffixes that no one says out loud:
  // "COSTCO WHOLESALE CORP /NEW", "AMAZON COM INC", "NIKE, Inc."
  let s = String(name)
    .replace(/\s*\/\s*\w+\s*$/, "")
    .replace(/[,.\s]+$/, "");
  for (let i = 0; i < 3; i++) {
    // A trailing ampersand means the suffix that followed it was part
    // of the name, as in JPMorgan Chase & Co.
    s = s.replace(/[\s,]+(INC|CORP|CORPORATION|CO|COMPANY|PLC|LTD|LIMITED|HOLDINGS|GROUP|CLASS\s+[A-Z])\.?$/i, "").replace(/[,.\s]+$/, "").replace(/\s*&$/, "");
  }
  s = s.replace(/\s+COM$/i, ""); // AMAZON COM
  const keepUpper = new Set(["UPS", "IBM", "AT&T", "3M", "PNC", "USA", "US", "NIKE"]);
  return s
    .split(/\s+/)
    .map((w, i) => {
      const upper = w.toUpperCase();
      if (keepUpper.has(upper)) return upper === "NIKE" ? "Nike" : upper;
      const lower = w.toLowerCase();
      if (i > 0 && small.has(lower)) return lower;
      if (/^[A-Z]&[A-Z]$/.test(w)) return w;
      return lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join(" ");
}

function money(m) {
  const a = Math.abs(m);
  if (a >= 1000) return (m < 0 ? "-$" : "$") + (a / 1000).toFixed(1) + "B";
  return (m < 0 ? "-$" : "$") + a.toFixed(0) + "M";
}

// Year over year for a quarter against the same quarter last year
function yoyPairs(quarters) {
  const byKey = new Map(quarters.map((q) => [q.key, q]));
  return quarters.map((q) => {
    const prior = byKey.get(`${q.year - 1}Q${q.q}`);
    return {
      ...q,
      revYoY: prior && prior.revenue ? q.revenue / prior.revenue - 1 : null,
      marginDelta: prior && prior.margin !== null && q.margin !== null ? q.margin - prior.margin : null,
      priorLabel: prior ? prior.label : null,
    };
  });
}

function analyzeFilings(profile) {
  const qs = yoyPairs(profile.quarters.filter((q) => q.revenue));
  const withMargin = qs.filter((q) => q.margin !== null);
  const latest = qs[qs.length - 1];
  if (!latest) return null;

  const lm = withMargin[withMargin.length - 1] || null;
  const trailing4 = withMargin.slice(-4);
  const prior4 = withMargin.slice(-8, -4);
  const avg = (a) => (a.length ? a.reduce((s, x) => s + x.margin, 0) / a.length : null);
  const t4 = avg(trailing4);
  const p4 = avg(prior4);
  const marginShift = t4 !== null && p4 !== null ? t4 - p4 : null;

  // Revenue direction on the same basis
  const revT4 = trailing4.length ? trailing4.reduce((s, x) => s + x.revenue, 0) : null;
  const revP4 = prior4.length ? prior4.reduce((s, x) => s + x.revenue, 0) : null;
  const revShift = revT4 && revP4 ? revT4 / revP4 - 1 : (latest.revYoY !== null ? latest.revYoY : 0);

  // The trough and peak actually present in the window
  let trough = null, peak = null;
  for (const q of withMargin) {
    if (!trough || q.margin < trough.margin) trough = q;
    if (!peak || q.margin > peak.margin) peak = q;
  }
  // Has margin turned up since the trough?
  const sinceTrough = trough ? withMargin.slice(withMargin.indexOf(trough) + 1) : [];
  const recovering = sinceTrough.length >= 2 && sinceTrough.every((q, i, a) => i === 0 || q.margin >= a[i - 1].margin - 0.002);

  // Prefer a human name when the lookup found one; EDGAR names are
  // registration names like "COSTCO WHOLESALE CORP /NEW".
  const co = (profile.displayName && profile.displayName.trim()) || titleCase(profile.company.name);
  const growing = revShift > 0.02;
  const strongGrowth = revShift > 0.08;
  const shrinking = revShift < -0.02;
  const marginUp = marginShift !== null && marginShift > 0.005;
  const marginDown = marginShift !== null && marginShift < -0.005;

  let pattern, question, answer;
  if (shrinking && marginUp) {
    pattern = "volume_for_margin";
    question = `Is the volume for margin trade working at ${co}?`;
    answer = `Yes, so far. Revenue is down ${pct(Math.abs(revShift))} over the last four quarters against the prior four, but operating margin is up ${pp(marginShift)} over the same span, so the business is converting lost revenue into better economics.`;
  } else if (shrinking && marginDown) {
    pattern = "shrink_not_paying";
    question = `Is ${co}'s shrinking top line buying anything?`;
    answer = `Not yet. Revenue fell ${pct(Math.abs(revShift))} across the last four quarters and operating margin fell ${pp(marginShift)} with it, so the business is getting smaller without getting better. Cost is not leaving fast enough to hold margin.`;
  } else if (growing && marginDown) {
    pattern = "growth_diluting";
    question = `Is ${co} buying growth at the cost of margin?`;
    answer = `It looks that way. Revenue grew ${pct(revShift)} over the last four quarters while operating margin gave up ${pp(marginShift)}, so the incremental revenue is coming in below the margin of the existing book.`;
  } else if (strongGrowth) {
    pattern = "scaling";
    question = `Can ${co} hold its margin while it grows this fast?`;
    answer = `So far yes. Revenue rose ${pct(revShift)} over the last four quarters against the prior four while operating margin held near ${pct(t4)}${
      marginShift !== null ? `, moving ${marginShift >= 0 ? "up" : "down"} ${pp(marginShift)}` : ""
    }. Growth at this rate usually costs margin, so the question is whether the cost to serve stays flat as the base gets larger.`;
  } else if (growing && marginUp) {
    pattern = "compounding";
    question = `Is ${co}'s operating leverage real?`;
    answer = `Yes. Revenue rose ${pct(revShift)} over the last four quarters and operating margin expanded ${pp(marginShift)} at the same time, which is revenue growing faster than the cost to serve it.`;
  } else if (recovering && trough) {
    pattern = "recovering";
    question = `Is the margin recovery at ${co} holding?`;
    answer = `It has held for ${sinceTrough.length} quarters. Margin troughed at ${pct(trough.margin)} in ${trough.label} and has not gone backwards since.`;
  } else {
    pattern = "steady";
    question = `Where does ${co} create value from here?`;
    answer = `The business is steady: revenue moved ${pct(revShift)} over the last four quarters against the prior four, with margin within half a point. Neither growth nor margin is the live issue, so the question is where the next unit of value comes from.`;
  }

  const findings = [];
  if (lm) {
    findings.push({
      lead: "Latest quarter",
      text: `${lm.label} revenue of ${money(lm.revenue)} at ${pct(lm.margin)} operating margin${
        lm.marginDelta === null
          ? ""
          : Math.abs(lm.marginDelta) < 0.0005
          ? `, level with ${lm.priorLabel}`
          : `, ${lm.marginDelta >= 0 ? "up" : "down"} ${pp(lm.marginDelta)} against ${lm.priorLabel}`
      }. Source: ${lm.form} filed ${lm.filed}.`,
    });
  }
  if (latest.revYoY !== null) {
    findings.push({
      lead: "Top line",
      text: `Revenue ${latest.revYoY >= 0 ? "grew" : "declined"} ${pct(Math.abs(latest.revYoY))} year over year in ${latest.label}, against ${money(latest.revenue)} for the quarter.`,
    });
  }
  if (trough && peak && trough.key !== peak.key) {
    findings.push({
      lead: "Margin range",
      text: `Across ${withMargin.length} quarters margin ran from ${pct(trough.margin)} in ${trough.label} to ${pct(peak.margin)} in ${peak.label}, a spread of ${pp(peak.margin - trough.margin)}.`,
    });
  }
  if (marginShift !== null) {
    findings.push({
      lead: "Trend",
      text: `The last four quarters averaged ${pct(t4)} operating margin against ${pct(p4)} for the four before, a ${marginShift >= 0 ? "gain" : "loss"} of ${pp(marginShift)}.`,
    });
  }

  // A falsifiable test on the next print. The benchmark is the same
  // quarter a year earlier, because that is the only comparison that
  // holds seasonality constant, and a trailing average can be dragged
  // by one unusual quarter.
  let watch = null;
  if (lm) {
    const nextQ = lm.q === 4 ? 1 : lm.q + 1;
    const nextYear = lm.q === 4 ? lm.year + 1 : lm.year;
    const benchKey = `${nextYear - 1}Q${nextQ}`;
    const bench = withMargin.find((q) => q.key === benchKey);
    const target = bench ? bench.margin : t4;
    const benchName = bench ? bench.label : "the trailing four quarter average";
    watch = {
      metric: "operating margin",
      quarter: `Q${nextQ} ${nextYear}`,
      target,
      benchmark: benchName,
      text:
        target === null
          ? null
          : `Judge this on Q${nextQ} ${nextYear}. Operating margin at or above ${pct(target)}, which is what ${benchName} delivered, says the trend is intact. A print below ${pct(target - 0.01)} says it is not.`,
    };
  }

  return {
    pattern,
    question,
    answer,
    findings,
    watch,
    quarters: qs,
    withMargin,
    trough,
    peak,
    trailingAvg: t4,
    priorAvg: p4,
    revShift,
    marginShift,
    company: co,
    source: profile.asOf
      ? `Source: ${profile.company.name} ${profile.asOf.form} filed ${profile.asOf.filed} and prior filings, via SEC EDGAR. Latest reported quarter: ${profile.asOf.period}.`
      : "Source: SEC EDGAR company filings.",
  };
}
