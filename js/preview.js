/* ============ DealDesk on screen previews ============
   Renders what each downloaded file will look like: a Word page, the
   slides, and the spreadsheet grid. Every value comes from the same
   engine functions the writers use, and the model preview reads the
   cell level spec the workbook itself is built from, so a preview
   cannot show something different from the file. */

const PV_DOC_SCALE = 84; // px per inch for document pages
const PV_SLIDE_SCALE = 52; // px per inch for slides

function pvEsc(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function pvPt(pt, scale) {
  return ((pt / 72) * (scale || PV_DOC_SCALE)).toFixed(1) + "px";
}

function pvC(hex) {
  return "#" + hex;
}

/* ================= Word document pages ================= */

function pvPage(inner, opts = {}) {
  const pad = PV_DOC_SCALE; // one inch margins
  return `<div class="pv-page" style="font-family:'${THEME.sans}',Arial,sans-serif;color:${pvC(THEME.body)};padding:${pad * 0.5}px ${pad}px;font-size:${pvPt(10.5)}">
      <div class="pv-runhead" style="color:${pvC(THEME.gray)};font-size:${pvPt(8)};border-bottom:1px solid ${pvC(THEME.rule)}">
        <span>${pvEsc(opts.code || "")}</span><span>${pvEsc(opts.right || "Private and confidential")}</span>
      </div>
      ${inner}
      <div class="pv-runfoot" style="color:${pvC(THEME.gray)};font-size:${pvPt(8)}">
        <span>${THEME.branded ? "Private and confidential" : "Draft for discussion purposes only"}</span><span>Page 1 of 1</span>
      </div>
    </div>`;
}

function pvTitleBlock(deal, title, subtitle) {
  return `<div style="margin-bottom:${pvPt(10)}">
      <div style="color:${pvC(THEME.accent)};font-weight:700;font-size:${pvPt(11)};letter-spacing:.06em">${pvEsc(dealCode(deal))}</div>
      <div style="font-family:'${THEME.serif}',Georgia,serif;font-size:${pvPt(28)};color:${pvC(THEME.ink)};line-height:1.15;margin:${pvPt(6)} 0">${pvEsc(title)}</div>
      <div style="color:${pvC(THEME.gray)};font-size:${pvPt(10.5)};padding-bottom:${pvPt(5)};border-bottom:1px solid ${pvC(THEME.rule)}">${pvEsc(subtitle)}</div>
    </div>`;
}

function pvH1(text) {
  return `<div style="font-family:'${THEME.serif}',Georgia,serif;font-weight:700;font-size:${pvPt(16)};color:${pvC(THEME.ink)};margin:${pvPt(16)} 0 ${pvPt(7)};padding-bottom:${pvPt(4)};border-bottom:1px solid ${pvC(THEME.rule)}">${pvEsc(text)}</div>`;
}

function pvP(text, opts = {}) {
  const bullet = opts.bullet ? `<span style="display:inline-block;width:${pvPt(12)}">&bull;</span>` : "";
  return `<div style="margin:0 0 ${pvPt(6)};line-height:1.45;${opts.bullet ? `padding-left:${pvPt(14)};text-indent:-${pvPt(14)}` : ""}">${bullet}${opts.html || pvEsc(text)}</div>`;
}

function pvBox(inner) {
  return `<div style="background:${pvC(THEME.fill)};border:1px solid ${pvC(THEME.rule)};border-left:4px solid ${pvC(THEME.accent)};padding:${pvPt(9)} ${pvPt(12)};margin:${pvPt(8)} 0">${inner}</div>`;
}

function pvTable(header, rows, opts = {}) {
  const th = header
    .map((h, i) => `<th style="background:${pvC(THEME.ink)};color:${pvC(THEME.onInk)};font-size:${pvPt(9.5)};text-align:${i === 0 ? "left" : "right"};padding:${pvPt(4)} ${pvPt(6)}">${pvEsc(h)}</th>`)
    .join("");
  const body = rows
    .map((r, ri) => {
      const strong = opts.boldRows && opts.boldRows.includes(ri);
      return `<tr>${r
        .map((c, i) => `<td style="font-size:${pvPt(9.5)};text-align:${i === 0 ? "left" : "right"};padding:${pvPt(4)} ${pvPt(6)};border-bottom:1px solid ${pvC(THEME.rule)};${strong ? `font-weight:700;background:${pvC(THEME.fill)}` : ""}">${pvEsc(c)}</td>`)
        .join("")}</tr>`;
    })
    .join("");
  return `<table style="width:100%;border-collapse:collapse;margin:${pvPt(4)} 0">${`<thead><tr>${th}</tr></thead>`}<tbody>${body}</tbody></table>`;
}

function pvSource(text) {
  return `<div style="color:${pvC(THEME.gray)};font-size:${pvPt(8)};margin:${pvPt(3)} 0 ${pvPt(10)}">${pvEsc(text)}</div>`;
}

function previewSummaryHTML(deal, m) {
  const P = practiceOf(deal);
  const n = buildNarrative(deal, m);
  const findings = buildKeyFindings(deal, m);
  const ft = buildFinancialsTable(m);
  const sample = m.source === "sample";
  const L = deal.lookup;
  const profile = deal.profile || {};

  let out = pvTitleBlock(
    deal,
    `${deal.company}: Business summary`,
    `${deal.dealType} | ${deal.industry || "Industry not specified"} | ${todayLabel()}${sample ? " | Illustrative sample financials" : ""}`
  );

  out += pvH1("1. Company overview");
  out += pvP(n.dealFrame);
  out += pvP(`${deal.company} operates in the ${(deal.industry || "target").toLowerCase()} sector with reported revenue of ${fmtM(m.revenueLatest)} in ${fyLabel(m.lastYear)}. ${P.overview(deal, m)}`);
  if (L && L.description) out += pvP(L.description);
  if (L) {
    const facts = [];
    if (L.founded) facts.push(`founded ${L.founded}`);
    if (L.hq) facts.push(`headquartered in ${L.hq}`);
    if (L.employees) facts.push(`about ${L.employees.toLocaleString("en-US")} employees`);
    if (L.ceo) facts.push(`led by CEO ${L.ceo}`);
    if (facts.length) {
      const line = facts.join("; ");
      out += pvP(line.charAt(0).toUpperCase() + line.slice(1) + ". Source: Wikidata.");
    }
  }
  out += pvBox(
    `<div style="font-family:'${THEME.serif}',Georgia,serif;font-weight:700;font-size:${pvPt(12)};color:${pvC(THEME.ink)};margin-bottom:${pvPt(5)}">What matters most</div>` +
      findings
        .map((f, i) => pvP(null, { html: `<strong style="color:${pvC(THEME.ink)}">${i + 1}. ${pvEsc(f.lead)}: </strong>${pvEsc(f.text)}` }))
        .join("")
  );

  out += pvH1("2. Financial summary");
  out += pvP(n.growthSentence);
  out += pvP(n.marginSentence);
  const tRows = [...ft.rows];
  const bold = [];
  if (ft.totalRow) { tRows.push(ft.totalRow); bold.push(tRows.length - 1); }
  if (ft.marginRow) tRows.push(ft.marginRow);
  out += pvTable(ft.header, tRows, { boldRows: bold });
  out += pvSource(sourceLineFor(m).replace("Management financials", "Management information"));

  out += pvH1("3. Recent news");
  const news = (profile.news && profile.news.length && profile.news) || (L && L.news);
  if (news && news.length) {
    out += pvP("Recent company news relevant to the engagement objective:");
    out += news
      .slice(0, 4)
      .map((x) => pvP(typeof x === "string" ? x : `${x.title}${x.source ? ` (${x.source}${x.date ? ", " + x.date : ""})` : ""}`, { bullet: true }))
      .join("");
  } else {
    out += pvP("Recent company news relevant to the engagement objective, with links to sources:");
    out += pvP("[Headline 1: from the company newsroom or investor relations page]", { bullet: true });
    out += pvP("[Headline 2: from sector press or analyst coverage]", { bullet: true });
  }

  out += pvH1("4. Key people");
  const people = profile.people && profile.people.length
    ? profile.people.map((p) => [p.role, p.name])
    : [
        ["Chief Executive Officer", (L && L.ceo) || "[Name]"],
        ["Chair of the Board", (L && L.chair) || "[Name]"],
        ["Chief Financial Officer", "[Name]"],
        ["[Other key role]", "[Name]"],
      ];
  out += pvTable(["Role", "Name"], people);

  out += pvH1("5. Key products");
  const products = profile.products && profile.products.length
    ? profile.products.map((p) => [p.name, "[$M]", p.desc || "[Brief description]"])
    : L && L.products && L.products.length
    ? L.products.map((p) => [p.charAt(0).toUpperCase() + p.slice(1), "[$M]", "[Confirm revenue split in interviews]"])
    : [["[Product or segment]", "[$M]", "[Brief description]"], ["[Product or segment]", "[$M]", "[Brief description]"]];
  out += pvTable(["Name", "Revenue", "Description"], products);

  const risk = situationRisk(deal, m);
  out += pvH1("6. Other");
  out += pvP("The most important remaining topics to orient the team, prioritized for this engagement:");
  out += pvP(
    risk.lead === "Customer concentration"
      ? "Supplier dependence: map the top suppliers by spend and any single points of failure."
      : "Customer and supplier concentration: confirm the share of revenue tied to the top five relationships.",
    { bullet: true }
  );
  out += pvP("Quality of reported figures: confirm accounting basis, one time items, and any restatements.", { bullet: true });
  out += pvP("Next step: management interviews using the accompanying interview guide.", { bullet: true });

  return pvPage(out, { code: dealCode(deal) });
}

function previewGuideHTML(deal, m) {
  const sections = buildInterviewQuestions(deal);
  const total = sections.reduce((s, sec) => s + sec.questions.length, 0);
  let out = pvTitleBlock(deal, `${deal.company}: Interview guide`, `${deal.dealType} | ${total} questions | Private and confidential`);
  out += pvP(null, { html: `<strong>Interviewee:</strong> [Name]` });
  out += pvP(null, { html: `<strong>Position, company:</strong> [Title], ${pvEsc(deal.company)}` });
  out += pvP(null, { html: `<strong>Date:</strong> [Date]` });
  out += pvBox(
    pvP(null, {
      html: `<strong style="color:${pvC(THEME.ink)}">Purpose: </strong>structure the conversation and stay focused on the key questions we need answered. Questions are ordered for maximum impact; capture answers in the space provided and flag follow ups for the tracker.`,
    })
  );
  let q = 0;
  sections.forEach((s, i) => {
    out += pvH1(`Topic ${String.fromCharCode(65 + i)}: ${s.section}`);
    s.questions.forEach((text) => {
      q += 1;
      out += pvP(null, { html: `<strong style="color:${pvC(THEME.accent)}">Q${q}. </strong>${pvEsc(text)}` });
      out += `<div style="color:${pvC(THEME.gray)};font-style:italic;font-size:${pvPt(9)};margin-bottom:${pvPt(14)}">Answer:</div>`;
    });
  });
  return pvPage(out, { code: dealCode(deal) });
}

function previewEmailHTML(deal, m) {
  const e = buildEmail(deal, m);
  let out = `<div style="color:${pvC(THEME.gray)};font-weight:700;font-size:${pvPt(9)};margin-bottom:${pvPt(8)}">DealDesk | Deal team memorandum</div>`;
  [["To", "Engagement leadership"], ["From", "Deal team"], ["Date", todayLabel()], ["Re", e.subject]].forEach(([k, v]) => {
    out += `<div style="border-bottom:1px solid ${pvC(THEME.rule)};padding:${pvPt(3)} 0"><strong>${k}:</strong> ${pvEsc(v)}</div>`;
  });
  out += `<div style="height:${pvPt(10)}"></div>`;
  out += pvP(null, { html: `<strong style="color:${pvC(THEME.ink)}">${pvEsc(e.answer)}</strong>` });
  out += pvP(e.soWhat);
  out += pvP(e.ask);
  out += `<div style="height:${pvPt(6)}"></div>`;
  e.points.forEach((p, i) => {
    out += pvP(null, { html: `<strong style="color:${pvC(THEME.ink)}">${i + 1}. ${pvEsc(p.lead)}: </strong>${pvEsc(p.text)}` });
  });
  out += `<div style="font-family:'${THEME.serif}',Georgia,serif;font-weight:700;font-size:${pvPt(12)};color:${pvC(THEME.ink)};margin:${pvPt(10)} 0 ${pvPt(5)}">Next steps</div>`;
  e.nextSteps.forEach((s) => (out += pvP(s, { bullet: true })));
  out += pvP(e.close);
  return pvPage(out, { code: dealCode(deal) });
}

/* ================= Slides ================= */

function pvSlideBox(x, y, w, h, inner, style) {
  const S = PV_SLIDE_SCALE;
  return `<div style="position:absolute;left:${x * S}px;top:${y * S}px;width:${w * S}px;height:${h * S}px;${style || ""}">${inner}</div>`;
}

function pvSlide(inner, dark) {
  const S = PV_SLIDE_SCALE;
  return `<div class="pv-slide" style="width:${13.33 * S}px;height:${7.5 * S}px;background:${dark ? pvC(THEME.ink) : "#fff"}">${inner}</div>`;
}

function pvFurniture(deal, { kicker, title, takeaway, source }) {
  const S = PV_SLIDE_SCALE;
  let out = "";
  if (kicker)
    out += pvSlideBox(0.45, 0.18, 9, 0.28, pvEsc(kicker.toUpperCase()), `color:${pvC(THEME.accent)};font-weight:700;font-size:${pvPt(10, S)};letter-spacing:.08em`);
  out += pvSlideBox(10.4, 0.18, 2.5, 0.28, THEME.branded && THEME.logo ? "" : "DRAFT | FOR DISCUSSION", `color:${pvC(THEME.grayLight)};font-size:${pvPt(9, S)};text-align:right`);
  out += pvSlideBox(0.45, 0.5, 12.43, 0.9, pvEsc(title), `font-family:'${THEME.serif}',Georgia,serif;font-size:${pvPt(fitSize(title, 12.43, 0.9, 19), S)};color:${pvC(THEME.ink)};line-height:1.2`);
  if (takeaway) {
    out += pvSlideBox(0.45, 6.42, 12.43, 0.52, "", `background:${pvC(THEME.fill)}`);
    out += pvSlideBox(0.45, 6.42, 0.06, 0.52, "", `background:${pvC(THEME.accent)}`);
    out += pvSlideBox(0.63, 6.42, 12.2, 0.52, pvEsc(takeaway), `font-weight:700;font-size:${pvPt(12.5, S)};color:${pvC(THEME.ink)};display:flex;align-items:center`);
  }
  if (source) out += pvSlideBox(0.45, 7.02, 9, 0.25, pvEsc(source), `color:${pvC(THEME.gray)};font-size:${pvPt(8, S)}`);
  out += pvSlideBox(0.45, 7.24, 6, 0.22, pvEsc(dealCode(deal) + " | Private and confidential"), `color:${pvC(THEME.grayLight)};font-size:${pvPt(8, S)}`);
  return out;
}

function pvCharts(m) {
  const S = PV_SLIDE_SCALE;
  const CX = 6.85, CW = 13.33 - 0.45 - 6.85;
  const big = m.revenueLatest >= 10000;
  const div = big ? 1000 : 1;
  const vals = m.series.revenue.map((v) => v / div);
  const max = Math.max(...vals) * 1.15 || 1;
  const w = CW * S, hBars = (m.ebitdaSeries ? 2.15 : 4.4) * S;
  const n = vals.length;
  const bw = (w / n) * 0.6;
  let bars = vals
    .map((v, i) => {
      const bh = (v / max) * (hBars - 22);
      const x = (w / n) * i + (w / n - bw) / 2;
      return `<rect x="${x}" y="${hBars - bh - 12}" width="${bw}" height="${bh}" fill="${pvC(THEME.accent)}" rx="2"/>
        <text x="${x + bw / 2}" y="${hBars - bh - 15}" text-anchor="middle" font-size="9" fill="#404850">${v.toFixed(1)}</text>
        <text x="${x + bw / 2}" y="${hBars - 1}" text-anchor="middle" font-size="9" fill="${pvC(THEME.gray)}">${fyLabel(m.years[i])}</text>`;
    })
    .join("");
  let out = pvSlideBox(CX, 1.42, CW, 0.26, `REVENUE, ${big ? "$B" : "$M"}`, `color:${pvC(THEME.gray)};font-weight:700;font-size:${pvPt(10, S)};letter-spacing:.08em`);
  out += pvSlideBox(CX, 1.7, CW, m.ebitdaSeries ? 2.15 : 4.4, `<svg width="${w}" height="${hBars}">${bars}</svg>`);

  if (m.ebitdaSeries) {
    const mg = m.ebitdaSeries.map((e, i) => (e / m.series.revenue[i]) * 100);
    const hL = 1.95 * S;
    const lo = Math.min(...mg), hi = Math.max(...mg);
    const span = hi - lo || 1;
    const pts = mg.map((v, i) => {
      const x = (w / mg.length) * i + w / mg.length / 2;
      const y = 18 + (1 - (v - lo) / span) * (hL - 40);
      return { x, y, v };
    });
    const path = pts.map((p, i) => `${i ? "L" : "M"}${p.x},${p.y}`).join(" ");
    out += pvSlideBox(CX, 4.02, CW, 0.26, "EBITDA MARGIN, %", `color:${pvC(THEME.gray)};font-weight:700;font-size:${pvPt(10, S)};letter-spacing:.08em`);
    out += pvSlideBox(
      CX, 4.3, CW, 1.95,
      `<svg width="${w}" height="${hL}"><path d="${path}" fill="none" stroke="${pvC(THEME.accent2)}" stroke-width="2"/>` +
        pts.map((p) => `<circle cx="${p.x}" cy="${p.y}" r="3" fill="${pvC(THEME.accent2)}"/><text x="${p.x}" y="${p.y - 7}" text-anchor="middle" font-size="9" fill="#404850">${p.v.toFixed(1)}</text>`).join("") +
        pts.map((p, i) => `<text x="${p.x}" y="${hL - 2}" text-anchor="middle" font-size="9" fill="${pvC(THEME.gray)}">${fyLabel(m.years[i])}</text>`).join("") +
        `</svg>`
    );
  }
  return out;
}

function previewDeckHTML(deal, m) {
  const c = buildDeckContent(deal, m);
  const S = PV_SLIDE_SCALE;
  const slides = [];

  // Cover
  slides.push(
    pvSlide(
      pvSlideBox(0.6, 2.0, 11, 0.4, pvEsc(c.cover.code), `color:${pvC(THEME.accent)};font-weight:700;font-size:${pvPt(13, S)};letter-spacing:.12em`) +
        pvSlideBox(0.6, 2.45, 12, 1.6, pvEsc(c.cover.title), `font-family:'${THEME.serif}',Georgia,serif;font-size:${pvPt(fitSize(c.cover.title, 12, 1.6, 38), S)};color:${pvC(THEME.onInk)}`) +
        pvSlideBox(0.6, 4.1, 11, 0.4, pvEsc(c.cover.sub), `color:#B8C2CC;font-size:${pvPt(13, S)}`) +
        pvSlideBox(0.6, 6.9, 11, 0.3, pvEsc(c.cover.note), `color:#8A97A3;font-size:${pvPt(9, S)}`),
      true
    )
  );

  // Summary slide
  let sum = pvFurniture(deal, { kicker: c.summary.kicker, title: c.summary.title, takeaway: c.summary.takeaway, source: c.srcLine });
  sum += pvSlideBox(0.45, 1.42, 6.15, 0.95, pvEsc(c.summary.frame), `color:#404850;font-size:${pvPt(fitSize(c.summary.frame, 6.15, 0.95, 12), S)};line-height:1.35`);
  c.groups.forEach((g, gi) => {
    const y = 2.5 + gi * 1.9;
    sum += pvSlideBox(0.45, y, 6.15, 0.3, pvEsc(`${gi + 1}. ${g.label.toUpperCase()} (SLIDE ${gi + 2})`), `color:${pvC(THEME.accent)};font-weight:700;font-size:${pvPt(11, S)};letter-spacing:.06em`);
    const lines = g.items.map((it) => `${c.all.findIndex((a) => a.name === it.name) + 1}. ${it.name}`);
    sum += pvSlideBox(0.45, y + 0.32, 6.15, 1.45, lines.map((l) => `<div style="margin-bottom:3px">${pvEsc(l)}</div>`).join(""), `color:${pvC(THEME.ink)};font-size:${pvPt(fitSizeLines(lines, 6.15, 1.45, 12), S)}`);
  });
  sum += pvCharts(m);
  slides.push(pvSlide(sum));

  // Category slides
  c.groups.forEach((g, gi) => {
    let sl = pvFurniture(deal, { kicker: `${gi + 1}. ${g.label}`, title: g.slideTitle, takeaway: g.slideTakeaway, source: c.srcLine });
    g.items.forEach((it, i) => {
      const y = 1.5 + i * 1.55;
      const num = c.all.findIndex((a) => a.name === it.name) + 1;
      sl += pvSlideBox(0.45, y + 0.04, 0.42, 0.42, `<div style="width:100%;height:100%;border-radius:50%;background:${pvC(gi === 0 ? THEME.accent2 : THEME.accent)};color:${pvC(THEME.onAccent)};display:flex;align-items:center;justify-content:center;font-weight:700;font-size:${pvPt(13, S)}">${num}</div>`);
      sl += pvSlideBox(1.05, y, 11.7, 0.42, `<strong style="color:${pvC(THEME.ink)};font-size:${pvPt(15, S)}">${pvEsc(it.name)}</strong><span style="color:${pvC(THEME.gray)};font-size:${pvPt(10.5, S)}">&nbsp;&nbsp;&nbsp;Impact: ${it.impact} | Ease: ${it.ease}</span>`, "display:flex;align-items:center");
      sl += pvSlideBox(1.05, y + 0.44, 11.7, 0.95, pvEsc(it.rationale), `color:#404850;font-size:${pvPt(12.5, S)};line-height:1.35`);
    });
    slides.push(pvSlide(sl));
  });

  // Matrix slide
  let mx = pvFurniture(deal, { kicker: c.matrix.kicker, title: c.matrix.title, source: c.srcLine });
  const px = 1.1, py = 1.55, pw = 7.0, ph = 4.5;
  mx += pvSlideBox(px, py, pw, ph, "", `background:#FAFBFC;border:1px solid ${pvC(THEME.rule)}`);
  mx += pvSlideBox(px + pw / 2, py, pw / 2, ph / 2, "", `background:${pvC(THEME.fillAccent)}`);
  mx += pvSlideBox(px + pw / 2, py + 0.05, pw / 2 - 0.1, 0.3, "Do first", `color:${pvC(THEME.accent)};font-weight:700;font-size:${pvPt(10, S)};text-align:right`);
  mx += pvSlideBox(px, py + ph + 0.1, pw, 0.3, "Ease of capture", `color:${pvC(THEME.gray)};font-size:${pvPt(11, S)};text-align:center`);
  mx += pvSlideBox(px - 1.0, py + ph / 2 - 0.2, 1.7, 0.4, "Value impact", `color:${pvC(THEME.gray)};font-size:${pvPt(11, S)};text-align:center`);
  const frac = { High: 0.8, Medium: 0.5, Low: 0.2 };
  const seen = {};
  c.all.forEach((it, i) => {
    const key = it.impact + it.ease;
    const bump = (seen[key] = (seen[key] || 0) + 1) - 1;
    const cx = px + frac[it.ease] * pw + bump * 0.55 - 0.21;
    const cy = py + (1 - frac[it.impact]) * ph - 0.21;
    mx += pvSlideBox(cx, cy, 0.42, 0.42, `<div style="width:100%;height:100%;border-radius:50%;background:${pvC(it.gi === 0 ? THEME.accent2 : THEME.accent)};color:${pvC(THEME.onAccent)};display:flex;align-items:center;justify-content:center;font-weight:700;font-size:${pvPt(12, S)}">${i + 1}</div>`);
  });
  mx += pvSlideBox(8.55, 1.55, 4.3, 0.3, "Opportunities", `color:${pvC(THEME.ink)};font-weight:700;font-size:${pvPt(11, S)}`);
  mx += pvSlideBox(8.55, 1.9, 4.3, 4.2, c.all.map((it, i) => `<div style="margin-bottom:4px">${i + 1}. ${pvEsc(it.name)}</div>`).join(""), `color:#404850;font-size:${pvPt(fitSizeLines(c.all.map((it, i) => `${i + 1}. ${it.name}`), 4.3, 4.2, 11), S)}`);
  slides.push(pvSlide(mx));

  return `<div class="pv-deck">${slides.map((s, i) => `<div class="pv-slide-wrap"><div class="pv-slide-no">Slide ${i + 1}</div>${s}</div>`).join("")}</div>`;
}

/* ================= Spreadsheet ================= */

function pvCellText(cell) {
  if (cell.v === "" || cell.v === null || cell.v === undefined) return "";
  if (typeof cell.v !== "number") return pvEsc(cell.v);
  if (cell.z === "0.0%") return (cell.v * 100).toFixed(1) + "%";
  if (cell.z && cell.z.indexOf('"x"') !== -1) return cell.v.toFixed(1) + "x";
  if (cell.z) return cell.v < 0 ? "(" + Math.abs(cell.v).toFixed(1) + ")" : cell.v.toFixed(1);
  return pvEsc(String(cell.v));
}

function previewModelHTML(deal, m) {
  const specs = modelSpecs(deal, m);
  const tabs = specs
    .map((s, i) => `<button class="pv-tab${i === 0 ? " active" : ""}" data-sheet="${i}" onclick="pvShowSheet(${i})">${pvEsc(s.name)}</button>`)
    .join("");

  const sheets = specs
    .map((s, si) => {
      const grid = [];
      for (let r = 0; r <= s.maxR; r++) grid.push(new Array(s.maxC + 1).fill(null));
      s.cells.forEach((c) => (grid[c.r][c.c] = c));
      const colLetters = Array.from({ length: s.maxC + 1 }, (_, i) => String.fromCharCode(65 + i));
      const head = `<tr><th class="pv-corner"></th>${colLetters
        .map((l, i) => `<th class="pv-colhead" style="min-width:${Math.max(60, (s.cols[i] || 12) * 7)}px">${l}</th>`)
        .join("")}</tr>`;
      const body = grid
        .map((row, r) => {
          const cells = row
            .map((c) => {
              if (!c) return `<td class="pv-cell"></td>`;
              const tip = c.f ? ` title="=${pvEsc(c.f)}"` : "";
              return `<td class="pv-cell pv-${c.role}${c.f ? " pv-hasf" : ""}"${tip}>${pvCellText(c)}</td>`;
            })
            .join("");
          return `<tr><th class="pv-rowhead">${r + 1}</th>${cells}</tr>`;
        })
        .join("");
      return `<div class="pv-sheet" data-sheet="${si}" style="display:${si === 0 ? "block" : "none"}"><table class="pv-grid">${head}${body}</table></div>`;
    })
    .join("");

  return `<div class="pv-xl"><div class="pv-tabs">${tabs}</div>${sheets}
    <p class="pv-hint">Yellow cells are inputs. Every other number is a formula: hover a cell to see it.</p></div>`;
}

function pvShowSheet(i) {
  document.querySelectorAll(".pv-sheet").forEach((el) => (el.style.display = el.dataset.sheet === String(i) ? "block" : "none"));
  document.querySelectorAll(".pv-tab").forEach((el) => el.classList.toggle("active", el.dataset.sheet === String(i)));
}
