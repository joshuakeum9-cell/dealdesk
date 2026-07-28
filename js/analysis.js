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


/* ---------- Practice lenses ----------
   The data decides the pattern; the practice decides what to ask about
   it. The same shrinking business is a cost problem to an operations
   team, a positioning problem to a strategist, and an underwriting
   problem to a deal team. */

const PRACTICE_LENS = {
  operations: {
    kicker: "OPERATIONAL DIAGNOSIS",
    docTitle: "Operational diagnosis",
    frame: "cost base, capacity, and throughput",
    questions: {
      volume_for_margin: (co) => `Is the volume for margin trade working at ${co}?`,
      shrink_not_paying: (co) => `Is ${co}'s shrinking top line buying anything?`,
      growth_diluting: (co) => `Why is growth costing ${co} margin?`,
      scaling: (co) => `Can ${co}'s operations hold margin at this growth rate?`,
      compounding: (co) => `Where is ${co}'s operating leverage coming from?`,
      recovering: (co) => `Is the margin recovery at ${co} structural or temporary?`,
      steady: (co) => `Where is the cost opportunity at ${co}?`,
    },
    implication:
      "For an operations engagement the first move is to locate where the cost actually sits, and on what schedule it can leave.",
    complicationTail:
      "Cost leaves a business on a slower clock than revenue does, which is why the timing, not the target, is usually what fails.",
  },
  strategy: {
    kicker: "STRATEGIC REVIEW",
    docTitle: "Strategic review",
    frame: "position, mix, and where the next unit of value comes from",
    questions: {
      volume_for_margin: (co) => `Is ${co} becoming a better business or just a smaller one?`,
      shrink_not_paying: (co) => `Is ${co} shrinking by choice or by loss?`,
      growth_diluting: (co) => `Is ${co} buying share it cannot monetize?`,
      scaling: (co) => `Can ${co} defend the position it is growing into?`,
      compounding: (co) => `What is ${co}'s advantage, and how long does it hold?`,
      recovering: (co) => `Has ${co}'s position actually improved, or only its numbers?`,
      steady: (co) => `Where does ${co}'s next unit of value come from?`,
    },
    implication:
      "For a strategy engagement the question is whether the shape the business is taking is the one worth defending.",
    complicationTail:
      "Financial results describe the position a business already holds; they say nothing about whether it is the position worth holding next.",
  },
  financial: {
    kicker: "TRANSACTION DIAGNOSTIC",
    docTitle: "Transaction diagnostic",
    frame: "earnings quality, valuation, and what would break the case",
    questions: {
      volume_for_margin: (co) => `Would you underwrite ${co}'s margin trade at today's multiple?`,
      shrink_not_paying: (co) => `What would you pay for ${co} while revenue and margin are both falling?`,
      growth_diluting: (co) => `Does ${co}'s growth still justify its multiple if margin keeps sliding?`,
      scaling: (co) => `What multiple does ${co}'s growth support, and what breaks it?`,
      compounding: (co) => `Is ${co}'s earnings quality good enough to underwrite?`,
      recovering: (co) => `Is ${co}'s recovery underwritable yet?`,
      steady: (co) => `What is the value case for ${co} beyond the status quo?`,
    },
    implication:
      "For a transaction the question is what this earnings profile is worth, and which assumption breaks the case first.",
    complicationTail:
      "A multiple applied to an earnings base that is still moving prices the past. The diligence question is which direction the base settles in.",
  },
};

// Recommendations differ by practice and by whether the business is
// shrinking, growing, or flat, so a deal team is not handed an
// operations plan.
const PRACTICE_ACTIONS = {
  operations: {
    shrinking: [
      ["Rebuild the cost base to the revenue that remains", "Size network, footprint, and overhead to what the business earns now rather than what it planned for, and put dates against each reduction."],
      ["Separate deliberate exits from lost business", "A planned exit and an unplanned loss look identical in the revenue line and call for opposite responses. Split them before deciding anything."],
      ["Protect the profitable core", "Before chasing recovery volume, confirm the remaining book is not also eroding on price."],
    ],
    growing: [
      ["Find where the next unit of cost enters", "Identify the step change, whether capacity, headcount, or infrastructure, that the next growth leg triggers, and plan it before it arrives."],
      ["Hold the cost to serve flat as the base grows", "Measure cost per unit of revenue by quarter. If it is rising, growth is being bought rather than earned."],
      ["Standardize before scaling", "Growth locks in whatever process exists when it happens, so fix the process first."],
    ],
    steady: [
      ["Benchmark the cost base while nothing is forcing it", "Flat performance rarely triggers scrutiny, which makes it the cheapest time to find the cost."],
      ["Take out the complexity that steady state hides", "Product, site, and process tails accumulate quietly and cost more attention than they earn."],
      ["Set a productivity target with a date", "Without a forcing function, a steady business stays exactly as expensive as it is today."],
    ],
  },
  strategy: {
    shrinking: [
      ["Decide which businesses are worth defending", "Rank what remains by whether it earns its cost of capital, and be explicit about what is being allowed to decline."],
      ["Establish whether the decline is market or share", "A shrinking market and lost share demand opposite responses, and the revenue line cannot tell them apart."],
      ["Redeploy the capital the shrink releases", "A smaller business is only a better one if the freed capital goes somewhere with a higher return."],
    ],
    growing: [
      ["Name the advantage the growth rests on", "Establish whether it is price, distribution, product, or simply a growing market, because only some of those are defensible."],
      ["Decide where not to compete", "Growth invites expansion in every direction; the strategy is which directions to refuse."],
      ["Invest ahead of the curve while economics allow", "Expansion is cheapest to fund while results are strong, not after they turn."],
    ],
    steady: [
      ["Find the adjacency the business could already take", "Identify the segment, channel, or price move available without changing what the company is."],
      ["Test whether the position is actually stable", "Flat results can mask share loss in a growing market. Check the position, not just the numbers."],
      ["Decide what to stop", "In a steady business the fastest value creation is usually ending something rather than starting something."],
    ],
  },
  financial: {
    shrinking: [
      ["Rebase earnings before applying any multiple", "Establish the run rate the business actually earns after the decline, and underwrite that rather than the trailing figure."],
      ["Quantify the downside case explicitly", "Model the scenario where the decline continues for two more years, and decide whether the price still works."],
      ["Diligence the cost reduction as a commitment, not a plan", "Savings that require actions not yet taken belong in the risk column, not the base case."],
    ],
    growing: [
      ["Test whether the growth is durable or pulled forward", "Separate structural demand from one time effects before capitalizing the growth rate."],
      ["Underwrite the margin, not just the top line", "Growth that dilutes margin changes the multiple the business deserves."],
      ["Set the price at which the case fails", "Define the entry multiple above which the return no longer clears the hurdle, before negotiating."],
    ],
    steady: [
      ["Underwrite the base case honestly", "A stable business is a bond with equity risk. Price it on cash generation rather than on a growth story."],
      ["Find the value the current owner is not taking", "Steady performance often hides an unexercised pricing, cost, or capital structure lever."],
      ["Confirm there is no deferred spending", "Flat results are sometimes maintained by underinvesting, which the buyer inherits."],
    ],
  },
};

function situationClass(a) {
  if (a.revShift < -0.02) return "shrinking";
  if (a.revShift > 0.02) return "growing";
  return "steady";
}

function analyzeFilings(profile, practice) {
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

  let pattern, answer;
  if (shrinking && marginUp) {
    pattern = "volume_for_margin";
    answer = `Yes, so far. Revenue is down ${pct(Math.abs(revShift))} over the last four quarters against the prior four, but operating margin is up ${pp(marginShift)} over the same span, so the business is converting lost revenue into better economics.`;
  } else if (shrinking && marginDown) {
    pattern = "shrink_not_paying";
    answer = `Not yet. Revenue fell ${pct(Math.abs(revShift))} across the last four quarters and operating margin fell ${pp(marginShift)} with it, so the business is getting smaller without getting better.`;
  } else if (growing && marginDown) {
    pattern = "growth_diluting";
    answer = `Revenue grew ${pct(revShift)} over the last four quarters while operating margin gave up ${pp(marginShift)}, so the incremental business is coming in below the economics of the existing book.`;
  } else if (strongGrowth) {
    pattern = "scaling";
    answer = `Revenue rose ${pct(revShift)} over the last four quarters against the prior four while operating margin held near ${pct(t4)}${
      marginShift !== null ? `, moving ${marginShift >= 0 ? "up" : "down"} ${pp(marginShift)}` : ""
    }. Growth at this rate usually costs margin, so holding it flat is the achievement and the fragile part.`;
  } else if (growing && marginUp) {
    pattern = "compounding";
    answer = `Revenue rose ${pct(revShift)} over the last four quarters and operating margin expanded ${pp(marginShift)} at the same time, which is revenue growing faster than the cost to serve it.`;
  } else if (recovering && trough) {
    pattern = "recovering";
    answer = `Margin troughed at ${pct(trough.margin)} in ${trough.label} and has held or improved for ${sinceTrough.length} quarters since.`;
  } else {
    pattern = "steady";
    answer = `Revenue moved ${pct(revShift)} over the last four quarters against the prior four, with operating margin near ${t4 === null ? "its recent level" : pct(t4)}. Neither growth nor margin is the live issue.`;
  }

  // The practice decides what to ask about the pattern the data shows
  const lens = PRACTICE_LENS[practice] || PRACTICE_LENS.operations;
  const question = lens.questions[pattern](co);
  answer = answer + " " + lens.implication;

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
    practice: practice || "operations",
    lens,
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

/* ---------- Memo narrative ----------
   Turns the diagnosis into the sections a diagnosis memo needs:
   situation, complication, recommendation, and what would change the
   call. Each block is written per pattern, so the memo argues the case
   the data supports rather than a generic one. */

const PATTERN_NARRATIVE = {
  volume_for_margin: {
    complication:
      "Volume leaves on a schedule, but the gains it is meant to buy arrive on their own. For several quarters the business carries cost built for revenue it no longer earns. The risk is stranded capacity, so the question is not whether margin is rising, it is whether the cost base is shrinking fast enough to keep it rising.",
    recommendations: [
      ["Keep cost leaving at least as fast as revenue", "Capacity freed by the exit has to come out on the same schedule, or the margin gain stalls once the easy reductions are done."],
      ["Refill with margin, not volume", "Replace what left with business that clears the current margin, rather than chasing the revenue line back to where it was."],
      ["Hold the discipline that created the gain", "The improvement came from mix and price. Protect it in the next contract cycle before pursuing new volume."],
    ],
  },
  shrink_not_paying: {
    complication:
      "The top line is contracting and margin is going with it, which means the shrink is not yet buying better economics. Either cost is leaving more slowly than revenue, or the revenue that left was more profitable than assumed. Both readings point at the cost base, and both get worse the longer they run.",
    recommendations: [
      ["Rebuild the cost base to the smaller revenue", "Size the network, footprint, and overhead to the revenue the business earns now, not the revenue it planned for."],
      ["Separate deliberate exits from lost business", "A planned exit and an unplanned loss look identical in the revenue line and call for opposite responses. Split them before deciding anything."],
      ["Protect the profitable core first", "Before pursuing recovery volume, confirm the remaining book is not also eroding on price."],
    ],
  },
  growth_diluting: {
    complication:
      "Revenue is growing while margin is falling, so the incremental business is coming in below the economics of the existing book. That is sustainable only if the new revenue is an investment with a known payback. Without that, growth is quietly buying down the margin.",
    recommendations: [
      ["Price the incremental business honestly", "Measure the margin on new revenue separately from the base. If it is dilutive, decide deliberately whether the share is worth it."],
      ["Find the fixed cost that grew with it", "Growth that costs margin usually added structure. Identify what was added and whether it scales."],
      ["Set a floor", "Define the margin the business will not go below to win volume, and hold it."],
    ],
  },
  scaling: {
    complication:
      "Growth at this rate normally costs margin, because the cost to serve rises with the base. Holding margin flat while growing is the achievement, and it is also the fragile part: the next leg of growth is usually where operating leverage breaks.",
    recommendations: [
      ["Find where the next unit of cost enters", "Identify the step change, whether capacity, headcount, or infrastructure, that the next growth leg triggers, and plan it before it arrives."],
      ["Protect the mix that holds the margin", "The margin is holding because of what is being sold. Confirm growth is not gradually shifting toward the lower margin end."],
      ["Convert scale into price or cost advantage", "Scale that does not lower unit cost or raise price is only volume."],
    ],
  },
  compounding: {
    complication:
      "Revenue and margin are rising together, which is the healthiest pattern and the easiest to take for granted. The risk is not this quarter, it is assuming the leverage continues without knowing what produces it.",
    recommendations: [
      ["Name the source of the leverage", "Establish whether the gain is price, mix, volume absorbing fixed cost, or one time relief. Only the first three repeat."],
      ["Invest while the economics are strong", "Expansion is cheapest to fund while margin is expanding, not after it has peaked."],
      ["Watch for the ceiling", "Track whether the margin gain per quarter is shrinking, the first sign the leverage is exhausted."],
    ],
  },
  recovering: {
    complication:
      "A recovery that has held for a few quarters is not yet a trend. The question is whether margin is recovering because the cost problem was fixed, or because a temporary drag rolled off. The two look identical for two or three quarters and then diverge.",
    recommendations: [
      ["Separate structural repair from a drag rolling off", "Confirm which part of the recovery is a cost change that persists and which was a one time item ending."],
      ["Lock in the structural part", "Whatever produced the durable improvement should be made permanent before attention moves elsewhere."],
      ["Set the bar above the trough", "Judge the next quarters against the pre trough level, not against the worst quarter."],
    ],
  },
  steady: {
    complication:
      "Neither the top line nor margin is moving much, which means the business is not in trouble and also not creating value from its current shape. Steady is comfortable, and it is the condition in which competitors accumulate advantages quietly.",
    recommendations: [
      ["Find the growth that is already available", "Identify the adjacent segment, channel, or price move the business could take without changing what it is."],
      ["Take the cost that steady state hides", "Flat performance rarely triggers cost scrutiny. Benchmark the cost base while no crisis is forcing it."],
      ["Decide what to stop", "In a steady business the fastest margin gain is usually ending something rather than starting something."],
    ],
  },
};

function buildMemoNarrative(a, profile) {
  const spec = PATTERN_NARRATIVE[a.pattern] || PATTERN_NARRATIVE.steady;
  const actions = (PRACTICE_ACTIONS[a.practice] || PRACTICE_ACTIONS.operations)[situationClass(a)];
  const co = a.company;
  const first = a.withMargin[0] || a.quarters[0];
  const last = a.withMargin[a.withMargin.length - 1] || a.quarters[a.quarters.length - 1];
  const sector = profile.company.sic ? profile.company.sic.toLowerCase() : null;

  return {
    situation: [
      `${co} is a ${profile.company.exchange || "US"} listed ${sector ? sector + " " : ""}filer reporting ${money(last.revenue)} of revenue in ${last.label}, its most recently reported quarter.`,
      `Across the ${a.quarters.length} quarters on file, revenue moved from ${money(first.revenue)} in ${first.label} to ${money(last.revenue)} in ${last.label}${
        a.trailingAvg !== null ? `, with operating margin averaging ${pct(a.trailingAvg)} over the last four quarters` : ""
      }.`,
    ],
    complication: spec.complication + " " + a.lens.complicationTail,
    diagnosisLead: `The question: across the ${a.withMargin.length} quarters with reported operating income, what do ${a.lens.frame} say about the answer above?`,
    recommendations: actions,
    changeTheCall: [
      a.watch && a.watch.target !== null
        ? `If ${a.watch.quarter} operating margin lands at or above ${pct(a.watch.target)}, the reading here holds and the question moves from diagnosing margin to sustaining it.`
        : null,
      a.watch && a.watch.target !== null
        ? `If it prints below ${pct(a.watch.target - 0.01)}, the pattern described here is not holding, and the cost base becomes the first priority rather than one of several.`
        : null,
      "These figures are company wide as filed. A business with divergent segments can hold a stable consolidated margin while one segment deteriorates, so segment detail would sharpen this read.",
    ].filter(Boolean),
  };
}
