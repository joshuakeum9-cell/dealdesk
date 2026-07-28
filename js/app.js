/* ============ DealDesk app: state, workflow, generation ============ */

const state = {
  uploads: { financials: [] },
  financials: null, // parsed {years, series, source}
  metrics: null,
  deal: null,
  brandLogo: null, // {b64, mime, w, h}
  lookup: null, // description, leadership, logo colors
  filings: null, // quarterly series straight from SEC filings
  analysis: null, // the diagnosis derived from that series
};

/* ---------- Live filings: the analytical spine ----------
   Filings give the numbers and the citation; the encyclopedia lookup
   gives the description and leadership. Filings come first because the
   documents argue from them. */

function currentPractice() {
  const el = document.querySelector('input[name="practice"]:checked');
  return el ? el.value : "financial";
}

// The practice decides the question, so switching it re-runs the
// diagnosis against the same filing data.
function refreshAnalysisForPractice() {
  if (state.filings) {
    state.analysis = analyzeFilings(state.filings, currentPractice());
    const st = document.getElementById("lookup-status");
    if (st && st.classList.contains("ok") && state.analysis) {
      st.textContent = `${state.filings.quarters.length} quarters from SEC filings through ${state.filings.asOf.period}. Diagnosis: ${state.analysis.question}`;
    }
  }
}

async function loadFilings(name, say) {
  try {
    const profile = await fetchFilingProfile(name, say);
    if (!profile || profile.error || !profile.quarters || !profile.quarters.length) return null;
    state.filings = profile;
    state.analysis = analyzeFilings(profile, currentPractice());
    // The annual series keeps the existing model and deck working
    if (profile.annual && profile.annual.length >= 2) {
      const series = { revenue: profile.annual.map((a) => a.revenue) };
      if (profile.annual.every((a) => a.cost !== null)) series.cogs = profile.annual.map((a) => a.cost);
      if (profile.annual.every((a) => a.opex !== null)) series.opex = profile.annual.map((a) => a.opex);
      state.financials = { years: profile.annual.map((a) => a.year), series, source: "filings" };
    }
    return profile;
  } catch (e) {
    return null;
  }
}

/* ---------- Company lookup (Wikipedia + Wikidata, free, keyless) ---------- */

async function runLookup() {
  const name = document.getElementById("f-company").value.trim();
  const status = document.getElementById("lookup-status");
  if (!name) {
    status.className = "lookup-status err";
    status.textContent = "Enter a company name first.";
    return;
  }
  status.className = "lookup-status";
  status.textContent = "Looking up " + name + "...";
  // Track the in flight promise so generation can wait for it instead
  // of silently building documents without the arriving profile.
  state.lookupPromise = (async () => {
    try {
      const filings = await loadFilings(name, (s) => { status.textContent = s + "..."; });
      if (filings && state.analysis) {
        const a = state.analysis;
        addFileChip("financials", `SEC filings: ${filings.quarters.length} quarters through ${filings.asOf.period}`, "applied");
        status.className = "lookup-status ok";
        status.textContent = `${a.company}: ${filings.quarters.length} quarters from SEC filings through ${filings.asOf.period}. Diagnosis: ${a.question}`;
        const industryEl0 = document.getElementById("f-industry");
        if (!industryEl0.value.trim() && filings.company.sic) industryEl0.value = filings.company.sic;
        const nameEl = document.getElementById("f-company");
        if (nameEl.value.trim().length <= 5) nameEl.value = a.company;
      }
      const r = await lookupCompany(name);
      if (!r) {
        state.lookup = null;
        status.className = "lookup-status err";
        status.textContent = "No public profile found. Private companies: continue manually.";
        return;
      }
      state.lookup = r;
      const industryEl = document.getElementById("f-industry");
      if (!industryEl.value.trim() && (r.industries[0] || r.shortDescription)) {
        industryEl.value = r.industries[0] || r.shortDescription;
      }
      const found = [];
      if (r.description) found.push("description");
      if (r.ceo) found.push("CEO " + r.ceo);
      if (r.employees) found.push(r.employees.toLocaleString("en-US") + " employees");
      if (r.revenueSeries) {
        found.push("revenue " + r.revenueSeries.years[0] + " to " + r.revenueSeries.years[r.revenueSeries.years.length - 1]);
        if (!state.financials || state.financials.source === "sample") {
          state.financials = { years: r.revenueSeries.years, series: { revenue: r.revenueSeries.revenue }, source: "lookup" };
          addFileChip("financials", "Reported revenue (" + r.revenueSeries.years.join(", ") + ")", "applied");
        }
      }
      const themed = await applyCompanyBranding(r);
      if (!state.analysis) {
        status.className = "lookup-status ok";
        status.textContent = "Found " + r.label + ": " + found.join("; ") + ". Documents will use this profile.";
      } else if (themed) {
        status.textContent += " Documents themed in " + r.label + " colors.";
      }
    } catch (e) {
      if (!state.analysis) {
        state.lookup = null;
        status.className = "lookup-status err";
        status.textContent = "Lookup unavailable right now; continue manually.";
      }
    } finally {
      state.lookupPromise = null;
    }
  })();
  await state.lookupPromise;
}

/* ---------- Typography ----------
   Ten pairings that exist on stock Windows and Mac installs, so a
   document opens the way it was designed on the recipient's machine. */
const FONT_PAIRS = {
  georgia: { serif: "Georgia", sans: "Arial", label: "Georgia + Arial" },
  garamond: { serif: "Garamond", sans: "Calibri", label: "Garamond + Calibri" },
  cambria: { serif: "Cambria", sans: "Calibri", label: "Cambria + Calibri" },
  times: { serif: "Times New Roman", sans: "Arial", label: "Times New Roman + Arial" },
  bookman: { serif: "Book Antiqua", sans: "Tahoma", label: "Book Antiqua + Tahoma" },
  constantia: { serif: "Constantia", sans: "Corbel", label: "Constantia + Corbel" },
  palatino: { serif: "Palatino Linotype", sans: "Verdana", label: "Palatino + Verdana" },
  century: { serif: "Century Schoolbook", sans: "Franklin Gothic Book", label: "Century Schoolbook + Franklin Gothic" },
  inverted: { serif: "Segoe UI Semibold", sans: "Georgia", label: "Segoe UI headings + Georgia body" },
  modern: { serif: "Arial", sans: "Arial", label: "Arial only" },
};

const HEADING_FONTS = [
  "Georgia", "Garamond", "Cambria", "Times New Roman", "Book Antiqua",
  "Constantia", "Palatino Linotype", "Century Schoolbook", "Segoe UI Semibold",
  "Trebuchet MS", "Arial", "Verdana",
];
const BODY_FONTS = [
  "Arial", "Calibri", "Segoe UI", "Tahoma", "Corbel", "Verdana",
  "Franklin Gothic Book", "Trebuchet MS", "Georgia", "Times New Roman",
];

function chosenFonts() {
  const key = document.getElementById("f-fonts").value;
  if (key === "custom") {
    return {
      serif: document.getElementById("f-font-head").value || "Georgia",
      sans: document.getElementById("f-font-body").value || "Arial",
    };
  }
  return FONT_PAIRS[key] || FONT_PAIRS.georgia;
}

// Build the theme for the selected output mode. Ghost drafts keep the
// neutral palette but still honor the chosen typography.
function currentTheme() {
  const mode = document.querySelector('input[name="output-mode"]:checked');
  const fonts = chosenFonts();
  if (!mode || mode.value === "ghost") {
    return { serif: fonts.serif, sans: fonts.sans, branded: false };
  }
  return {
    ink: document.getElementById("f-brand-primary").value,
    accent: document.getElementById("f-brand-accent").value,
    serif: fonts.serif,
    sans: fonts.sans,
    logo: state.brandLogo,
    branded: true,
  };
}

/* ---------- Color readouts: hex, rgb, and hsl together ---------- */

function hexToRgb(hex) {
  const n = parseInt(String(hex).replace("#", ""), 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function rgbToHsl({ r, g, b }) {
  const R = r / 255, G = g / 255, B = b / 255;
  const mx = Math.max(R, G, B), mn = Math.min(R, G, B), d = mx - mn;
  const l = (mx + mn) / 2;
  let h = 0, s = 0;
  if (d) {
    s = d / (1 - Math.abs(2 * l - 1));
    h = mx === R ? ((G - B) / d) % 6 : mx === G ? (B - R) / d + 2 : (R - G) / d + 4;
    h = Math.round(h * 60);
    if (h < 0) h += 360;
  }
  return { h, s: Math.round(s * 100), l: Math.round(l * 100) };
}

function updateColorReadout(inputId, boxId) {
  const hex = document.getElementById(inputId).value.toUpperCase();
  const rgb = hexToRgb(hex);
  const hsl = rgbToHsl(rgb);
  document.getElementById(boxId).innerHTML =
    `<span class="cr-chip" style="background:${hex}"></span>` +
    `<code>HEX ${hex}</code>` +
    `<code>RGB ${rgb.r}, ${rgb.g}, ${rgb.b}</code>` +
    `<code>HSL ${hsl.h}, ${hsl.s}%, ${hsl.l}%</code>`;
}

function refreshBrandingUI() {
  updateColorReadout("f-brand-primary", "ro-primary");
  updateColorReadout("f-brand-accent", "ro-accent");
  // One control or the other, never both. A named pairing already fixes
  // both fonts, so the mix and match selects are hidden and disabled
  // rather than sitting there implying they still have a say.
  const custom = document.getElementById("f-fonts").value === "custom";
  document.getElementById("custom-fonts").hidden = !custom;
  ["f-font-head", "f-font-body"].forEach((id) => {
    document.getElementById(id).disabled = !custom;
  });
  const f = chosenFonts();
  const sample = document.getElementById("font-sample");
  if (sample) {
    sample.innerHTML =
      `<span style="font-family:'${f.serif}',Georgia,serif;font-size:19px">${f.serif}: heading sample</span>` +
      `<span style="font-family:'${f.sans}',Arial,sans-serif;font-size:14px">${f.sans}: body text sample, the quick brown fox jumps over the lazy dog.</span>`;
  }
}

// Theme the documents in the looked up company's own colors. Reads the
// palette from its logo, keeps a readable contrast, and leaves every
// control overridable so a user can still impose their own branding.
async function applyCompanyBranding(lookup) {
  if (!lookup || !lookup.logoUrl) return false;
  try {
    const img = await new Promise((resolve, reject) => {
      const i = new Image();
      i.crossOrigin = "anonymous";
      i.onload = () => resolve(i);
      i.onerror = reject;
      i.src = lookup.logoUrl;
    });
    const palette = extractLogoPalette(img);
    if (!palette) return false;
    document.getElementById("f-brand-primary").value = palette.primary;
    document.getElementById("f-brand-accent").value = palette.accent;
    const branded = document.querySelector('input[name="output-mode"][value="branded"]');
    if (branded) branded.checked = true;
    // Keep the logo itself for the document header
    try {
      const cv = document.createElement("canvas");
      cv.width = img.naturalWidth; cv.height = img.naturalHeight;
      cv.getContext("2d").drawImage(img, 0, 0);
      const url = cv.toDataURL("image/png");
      const m = url.match(/^data:(image\/\w+);base64,(.+)$/);
      if (m) state.brandLogo = { b64: m[2], mime: m[1], w: img.naturalWidth, h: img.naturalHeight };
    } catch (e) { /* logo optional */ }
    refreshBrandingUI();
    state.autoBranded = { primary: palette.primary, accent: palette.accent };
    return true;
  } catch (e) {
    return false;
  }
}

function initCompanyField() {
  const el = document.getElementById("f-company");
  if (!el) return;
  el.addEventListener("input", () => {
    el.classList.remove("field-error");
    document.getElementById("company-error").hidden = true;
  });
  document.querySelectorAll('.check-row input[type="checkbox"]').forEach((cb) =>
    cb.addEventListener("change", () => (document.getElementById("select-error").hidden = true))
  );
}

// Landing page lookup: the fastest possible path to a real document
async function heroLookup() {
  const name = document.getElementById("hero-company").value.trim();
  const out = document.getElementById("hero-result");
  if (!name) { out.className = "hero-result err"; out.textContent = "Type a public company name first."; return; }
  out.className = "hero-result";
  out.textContent = "Looking up " + name + "...";
  try {
    const filings = await loadFilings(name, (s) => { out.textContent = s + "..."; });
    if (filings && state.analysis) {
      const a = state.analysis;
      out.className = "hero-result ok";
      out.innerHTML =
        `<strong>${a.question}</strong>` +
        `<span class="hero-verdict">${a.answer}</span>` +
        `<span class="hero-src">${filings.quarters.length} quarters from ${a.company} SEC filings, latest ${filings.asOf.period} (${filings.asOf.form} filed ${filings.asOf.filed}).</span>` +
        `<button class="btn btn-sm btn-primary hero-go" onclick="startFromHero()">Build the documents</button>`;
      // Registered as the in flight lookup so generation waits for the
      // description and the brand colors instead of racing them.
      state.lookupPromise = lookupCompany(name)
        .then(async (r) => { if (r) { state.lookup = r; await applyCompanyBranding(r); } })
        .catch(() => {})
        .finally(() => { state.lookupPromise = null; });
      return;
    }
    const r = await lookupCompany(name);
    if (!r) {
      out.className = "hero-result err";
      out.textContent = "No filings or public profile found for " + name + ". Private companies work too: start a deal and type the numbers in.";
      return;
    }
    state.lookup = r;
    const bits = [];
    if (r.ceo) bits.push("CEO " + r.ceo);
    if (r.employees) bits.push(r.employees.toLocaleString("en-US") + " employees");
    let verdictLine = "";
    if (r.revenueSeries) {
      state.financials = { years: r.revenueSeries.years, series: { revenue: r.revenueSeries.revenue }, source: "lookup" };
      const m = computeMetrics(state.financials);
      const deal = { company: r.label, industry: r.industries[0] || "", dealType: "Acquisition", practice: "financial", context: "", situation: {}, lookup: r };
      const verdict = buildEmail(deal, m).subject.split(";")[0].split(":").slice(1).join(":").trim();
      bits.push(fmtM(m.revenueLatest) + " revenue, " + fmtPct(m.revenueCAGR) + " growth");
      verdictLine = `<span class="hero-verdict">Recommendation the engine reaches: ${verdict}</span>`;
    }
    out.className = "hero-result ok";
    out.innerHTML = `<strong>${r.label}</strong> found: ${bits.join("; ")}. ${verdictLine}
      <button class="btn btn-sm btn-primary hero-go" onclick="startFromHero()">Build the documents</button>`;
  } catch (e) {
    out.className = "hero-result err";
    out.textContent = "Lookup is unavailable right now. Start a deal and type the numbers in instead.";
  }
}

function startFromHero() {
  const name = document.getElementById("hero-company").value.trim();
  startWorkflow();
  const display = state.analysis ? state.analysis.company : state.lookup ? state.lookup.label : name;
  if (name) document.getElementById("f-company").value = display;
  if (state.filings) {
    const st = document.getElementById("lookup-status");
    st.className = "lookup-status ok";
    st.textContent = `${state.filings.quarters.length} quarters from SEC filings through ${state.filings.asOf.period}. Diagnosis: ${state.analysis.question}`;
    const ind = document.getElementById("f-industry");
    if (!ind.value.trim() && state.filings.company.sic) ind.value = state.filings.company.sic;
    addFileChip("financials", `SEC filings: ${state.filings.quarters.length} quarters through ${state.filings.asOf.period}`, "applied");
  }
  if (state.lookup) {
    const ind = document.getElementById("f-industry");
    if (!ind.value.trim() && state.lookup.industries[0]) ind.value = state.lookup.industries[0];
    const st = document.getElementById("lookup-status");
    st.className = "lookup-status ok";
    st.textContent = "Using the profile found for " + state.lookup.label + ".";
    if (state.financials && state.financials.source === "lookup") {
      addFileChip("financials", "Wikidata reported revenue (" + state.financials.years.join(", ") + ")", "applied");
    }
  }
}

function initBrandingUI() {
  const head = document.getElementById("f-font-head");
  const body = document.getElementById("f-font-body");
  if (head && !head.options.length) {
    head.innerHTML = HEADING_FONTS.map((f) => `<option>${f}</option>`).join("");
    body.innerHTML = BODY_FONTS.map((f) => `<option>${f}</option>`).join("");
  }
  const pairSel = document.getElementById("f-fonts");
  if (pairSel && !pairSel.dataset.built) {
    pairSel.innerHTML =
      Object.entries(FONT_PAIRS).map(([k, v]) => `<option value="${k}">${v.label}</option>`).join("") +
      `<option value="custom">Mix and match your own</option>`;
    pairSel.dataset.built = "1";
  }
  ["f-brand-primary", "f-brand-accent", "f-fonts", "f-font-head", "f-font-body"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("input", refreshBrandingUI);
  });
  refreshBrandingUI();
}

/* ---------- View switching ---------- */

function showView(id) {
  document.querySelectorAll(".view").forEach((v) => v.classList.remove("active"));
  document.getElementById(id).classList.add("active");
  window.scrollTo({ top: 0 });
}

function goHome(anchor) {
  showView("view-landing");
  if (anchor) {
    requestAnimationFrame(() => {
      const el = document.getElementById(anchor);
      if (el) el.scrollIntoView({ behavior: "smooth" });
    });
  }
}

function startWorkflow() {
  showView("view-workflow");
  nextStep(1);
}

/* ---------- Stepper ---------- */

function nextStep(n) {
  document.querySelectorAll(".wf-step").forEach((s) => s.classList.remove("active"));
  document.getElementById("step-" + n).classList.add("active");

  document.querySelectorAll(".stepper-item").forEach((item) => {
    const step = Number(item.dataset.step);
    item.classList.toggle("active", step === n);
    item.classList.toggle("done", step < n);
  });

  if (n === 3) fillReview();
  window.scrollTo({ top: 0 });
}

/* ---------- Real uploads ---------- */

function openPicker(zone) {
  document.getElementById("picker-" + zone).click();
}

function addFileChip(zone, name, note) {
  state.uploads[zone].push(name);
  const zoneEl = document.getElementById("zone-" + zone);
  zoneEl.classList.add("has-files");
  const chip = document.createElement("div");
  chip.className = "file-chip";
  chip.textContent = "✓ " + name + (note ? " (" + note + ")" : "");
  document.getElementById("files-" + zone).appendChild(chip);
}

function initPickers() {
  // Upload zones: reachable and operable by keyboard
  document.querySelectorAll(".upload-zone").forEach((z) => {
    z.setAttribute("tabindex", "0");
    z.setAttribute("role", "button");
    z.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        z.click();
      }
    });
  });
  ["financials"].forEach((zone) => {
    const input = document.getElementById("picker-" + zone);
    input.addEventListener("change", async () => {
      for (const file of input.files) {
        if (zone === "financials") {
          try {
            const buf = await file.arrayBuffer();
            const parsed = parseFinancialsWorkbook(buf);
            if (parsed) {
              state.financials = parsed;
              addFileChip(zone, file.name, "parsed");
            } else {
              addFileChip(zone, file.name, "could not read: sample data will be used");
            }
          } catch (e) {
            addFileChip(zone, file.name, "could not read: sample data will be used");
          }
        } else {
          // Context files are listed but not parsed in this version
          addFileChip(zone, file.name);
        }
      }
      input.value = "";
    });
  });

  // Client logo: read as base64, capture natural dimensions, and pull
  // the brand colors straight out of the pixels. No AI: a canvas
  // histogram picks a dark dominant color for primary and the most
  // saturated distinct color for accent.
  const logoInput = document.getElementById("picker-logo");
  logoInput.addEventListener("change", () => {
    const file = logoInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const [, mime, b64] = reader.result.match(/^data:(image\/\w+);base64,(.+)$/) || [];
        if (!b64) return;
        state.brandLogo = { b64, mime, w: img.naturalWidth, h: img.naturalHeight };
        const palette = extractLogoPalette(img);
        let note = "";
        if (palette) {
          document.getElementById("f-brand-primary").value = palette.primary;
          document.getElementById("f-brand-accent").value = palette.accent;
          const branded = document.querySelector('input[name="output-mode"][value="branded"]');
          if (branded) branded.checked = true;
          note = " (brand colors applied from logo)";
        }
        const list = document.getElementById("files-logo");
        list.innerHTML = "";
        const chip = document.createElement("div");
        chip.className = "file-chip";
        chip.textContent = "✓ " + file.name + note;
        list.appendChild(chip);
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
    logoInput.value = "";
  });
}

/* ---------- Logo palette extraction (pure pixel math) ---------- */

function extractLogoPalette(img) {
  try {
    const size = 48;
    const cv = document.createElement("canvas");
    cv.width = size;
    cv.height = size;
    const ctx = cv.getContext("2d");
    ctx.drawImage(img, 0, 0, size, size);
    const data = ctx.getImageData(0, 0, size, size).data;

    const buckets = new Map();
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] < 200) continue; // transparent
      const r = data[i], g = data[i + 1], b = data[i + 2];
      if (r > 242 && g > 242 && b > 242) continue; // white background
      const key = ((r >> 4) << 8) | ((g >> 4) << 4) | (b >> 4);
      const e = buckets.get(key) || { n: 0, r: 0, g: 0, b: 0 };
      e.n++; e.r += r; e.g += g; e.b += b;
      buckets.set(key, e);
    }
    const cols = [...buckets.values()]
      .map((e) => {
        const r = e.r / e.n, g = e.g / e.n, b = e.b / e.n;
        const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
        return { r, g, b, n: e.n, sat: mx === 0 ? 0 : (mx - mn) / mx, lum: 0.2126 * r + 0.7152 * g + 0.0722 * b };
      })
      .filter((c) => c.n >= 4);
    if (!cols.length) return null;

    // Primary: common and dark wins (it plays the ink role in documents)
    let primary = cols.slice().sort((a, b) => b.n * (1 - (b.lum / 255) * 0.7) - a.n * (1 - (a.lum / 255) * 0.7))[0];
    // Documents need a primary dark enough to carry text and fills
    let { r, g, b } = primary;
    while (0.2126 * r + 0.7152 * g + 0.0722 * b > 150) { r *= 0.85; g *= 0.85; b *= 0.85; }
    const hex = (R, G, B) => "#" + [R, G, B].map((v) => Math.round(v).toString(16).padStart(2, "0")).join("");
    const primaryHex = hex(r, g, b);

    // Accent: most saturated color clearly different from the primary
    const dist = (a, c) => Math.hypot(a.r - c.r, a.g - c.g, a.b - c.b);
    const accent = cols
      .filter((c) => dist(c, primary) > 60 && c.sat > 0.25 && c.lum < 235)
      .sort((a, c) => c.sat * c.n - a.sat * a.n)[0];
    return { primary: primaryHex, accent: accent ? hex(accent.r, accent.g, accent.b) : primaryHex };
  } catch (e) {
    return null;
  }
}

/* ---------- Manual financials entry ---------- */

function toggleManual() {
  const panel = document.getElementById("manual-panel");
  panel.hidden = !panel.hidden;
}

function applyManualFinancials() {
  const num = (id) => {
    const v = document.getElementById(id).value.trim();
    return v === "" ? null : Number(v);
  };
  const cols = [0, 1, 2, 3];
  const years = [], revenue = [], cogs = [], opex = [];
  for (const i of cols) {
    const r = num("mr-" + i);
    if (r === null || !Number.isFinite(r)) continue;
    years.push(num("my-" + i) || 2021 + i);
    revenue.push(r);
    cogs.push(num("mc-" + i));
    opex.push(num("mo-" + i));
  }
  const status = document.getElementById("manual-status");
  if (revenue.length < 2) {
    status.textContent = "Enter revenue for at least two years.";
    return;
  }
  // Sort chronologically in case years were typed newest first
  const order = years.map((_, i) => i).sort((a, b) => years[a] - years[b]);
  const sortBy = (arr) => order.map((i) => arr[i]);
  const series = { revenue: sortBy(revenue) };
  if (cogs.every((v) => v !== null)) series.cogs = sortBy(cogs);
  if (opex.every((v) => v !== null)) series.opex = sortBy(opex);
  state.financials = { years: sortBy(years), series, source: "manual" };
  status.textContent = "Applied: " + revenue.length + " years of financials.";
  addFileChip("financials", "Manual entry (" + years[0] + " to " + years[years.length - 1] + ")", "applied");
}

function loadSampleCompany() {
  state.financials = SAMPLE_FINANCIALS;
  const f = document.getElementById("f-company");
  if (!f.value.trim()) f.value = "Meridian Components Inc.";
  const ind = document.getElementById("f-industry");
  if (!ind.value.trim()) ind.value = "Industrial manufacturing";
  if (!state.uploads.financials.includes("Sample_Financials_FY21_24.xlsx")) {
    addFileChip("financials", "Sample_Financials_FY21_24.xlsx", "sample");
  }
}

/* ---------- Review ---------- */

function getDeal() {
  const val = (id, fallback) => document.getElementById(id).value.trim() || fallback;
  const sval = (id) => document.getElementById(id).value;
  const practice = document.querySelector('input[name="practice"]:checked');
  return {
    practice: practice ? practice.value : "financial",
    company: val("f-company", "Meridian Components Inc."),
    industry: val("f-industry", "Industrial manufacturing"),
    dealType: document.getElementById("f-dealtype").value,
    context: document.getElementById("f-context").value.trim(),
    lookup: state.lookup || null,
    filings: state.filings || null,
    analysis: state.analysis || null,
    profile: readProfile(),
    situation: {
      goal: sval("s-goal"),
      marginTrend: sval("s-margin"),
      revenueDriver: sval("s-driver"),
      concentration: sval("s-concentration"),
      capacity: sval("s-capacity"),
      costProgram: sval("s-costprog"),
      competition: sval("s-competition"),
      urgency: sval("s-urgency"),
    },
  };
}

// Optional manual company profile for private companies
function readProfile() {
  const lines = (id) =>
    document.getElementById(id).value.split(/\n+/).map((l) => l.trim()).filter(Boolean);
  const split = (l) => {
    const i = l.indexOf(":");
    return i > 0 ? { head: l.slice(0, i).trim(), tail: l.slice(i + 1).trim() } : { head: "", tail: l };
  };
  return {
    people: lines("p-people").map((l) => { const s = split(l); return { role: s.head || "[Role]", name: s.tail }; }),
    products: lines("p-products").map((l) => { const s = split(l); return { name: s.head || s.tail, desc: s.head ? s.tail : "" }; }),
    news: lines("p-news"),
  };
}

// The goal dropdown offers different options per practice
function refreshGoalOptions() {
  const practice = document.querySelector('input[name="practice"]:checked');
  const goals = PRACTICE_GOALS[practice ? practice.value : "financial"] || [];
  const sel = document.getElementById("s-goal");
  const current = sel.value;
  sel.innerHTML = '<option value="">Not specified</option>' +
    goals.map((g) => `<option>${g}</option>`).join("");
  if (goals.includes(current)) sel.value = current;
}

document.querySelectorAll('input[name="practice"]').forEach((r) =>
  r.addEventListener("change", () => { refreshGoalOptions(); refreshAnalysisForPractice(); })
);

function fillReview() {
  const deal = getDeal();
  document.getElementById("r-practice").textContent = practiceOf(deal).label;
  document.getElementById("r-company").textContent = deal.company;
  document.getElementById("r-industry").textContent = deal.industry;
  document.getElementById("r-dealtype").textContent = deal.dealType;
  const total = Object.values(state.uploads).reduce((n, arr) => n + arr.length, 0);
  document.getElementById("r-files").textContent =
    total > 0 ? total + " file" + (total > 1 ? "s" : "") : "None (sample data will be used)";

  // Warn loudly when a named company would get fictional demo numbers
  const usingSample = !state.financials || state.financials.source === "sample";
  const warn = document.getElementById("sample-warning");
  warn.hidden = !usingSample;
  if (usingSample) {
    const named = deal.company && deal.company !== "Meridian Components Inc.";
    warn.innerHTML = named
      ? `No financial statements were provided for <strong>${deal.company}</strong>. The documents will use fictional demo numbers, labeled as illustrative on every file. Go back to add real numbers, or continue.`
      : "No financial statements were provided, so the documents will use the sample company's illustrative numbers. Upload a spreadsheet or type the numbers on the previous step for real output.";
  }
}

/* ---------- Generation ---------- */

async function runGeneration() {
  // A lookup still in flight would arrive after the documents build;
  // wait for it briefly rather than dropping it.
  if (state.lookupPromise) {
    await Promise.race([state.lookupPromise, new Promise((r) => setTimeout(r, 8000))]);
  }
  const deal = getDeal();

  // The company name drives every document; never substitute silently.
  if (!document.getElementById("f-company").value.trim()) {
    nextStep(1);
    const el = document.getElementById("f-company");
    el.focus();
    el.classList.add("field-error");
    document.getElementById("company-error").hidden = false;
    return;
  }

  const selected = selectedOutputs();
  if (!selected.length) {
    document.getElementById("select-error").hidden = false;
    return;
  }

  state.deal = deal;
  // A real company name with fictional numbers is flagged in plain sight
  // on the review step and stamped on every document, so no native
  // popup is needed to interrupt the flow.
  if (!state.financials) state.financials = SAMPLE_FINANCIALS;
  state.metrics = computeMetrics(state.financials);
  state.blobs = {};
  setTheme(currentTheme());

  const P = practiceOf(state.deal);
  // Live filings produce a diagnosis rather than a brief, so the cards
  // have to name what was actually built. Promising a synergy deck and
  // handing over an operating review is the kind of small dishonesty
  // that makes someone distrust the rest of the package.
  const diag = !!state.analysis;
  const cardText = {
    summary: diag
      ? ["Diagnosis Memo", "Answer up front, then situation, complication, diagnosis, and the recommendation"]
      : ["Business Summary", "Company overview, financial profile, market position, key risks"],
    guide: [guideTitleOf(state.deal), "Structured management interview questions by workstream"],
    synergy: diag
      ? ["Diagnosis Presentation", "The argument in eight slides, with the reported quarters charted"]
      : [P.deckShort, "Synergy hypotheses ranked by impact and achievability"],
    model: diag
      ? ["Supporting Model + Email Summary", `The reported quarters and a bear, base, and bull ${state.analysis.practice === "strategy" ? "growth" : "margin"} scenario, plus an answer first email draft`]
      : ["Excel Model + Email Summary", "Scenario model with labeled inputs, plus an answer first email draft"],
  };
  for (const [key, [title, desc]] of Object.entries(cardText)) {
    const card = document.querySelector(`.output-card[data-output="${key}"]`);
    if (!card) continue;
    card.querySelector("h3").textContent = title;
    card.querySelector(".output-desc").textContent = desc;
  }

  nextStep(4);
  document.getElementById("gen-title").textContent = "Building your deal package";
  document.getElementById("gen-sub").textContent =
    "Each file is being written in your browser right now. Nothing is uploaded to a server.";
  document.getElementById("final-actions").style.visibility = "hidden";

  const cards = Array.from(document.querySelectorAll(".output-card"));
  cards.forEach((card) => {
    const on = selected.includes(card.dataset.output);
    card.hidden = !on;
    card.classList.remove("ready");
    setStatus(card, "pending", on ? "Queued" : "Skipped");
    card.querySelector(".output-bar i").style.width = "0";
    card.querySelectorAll(".output-actions .btn").forEach((b) => (b.disabled = true));
  });

  // Build the files one at a time, updating each card when its own file
  // actually exists. The bar creeps while the work runs and completes
  // when the blob is in hand, so the progress shown is real progress.
  for (const key of selected) {
    const card = cards.find((c) => c.dataset.output === key);
    setStatus(card, "working", "Writing");
    const bar = card.querySelector(".output-bar i");
    let pct = 0;
    const creep = setInterval(() => {
      pct = Math.min(88, pct + 9);
      bar.style.width = pct + "%";
    }, 90);
    const started = Date.now();
    try {
      state.blobs[key] = await buildOutput(key, state.deal, state.metrics);
    } catch (e) {
      clearInterval(creep);
      console.error("Generation failed for " + key, e);
      setStatus(card, "working", "Failed");
      bar.style.width = "100%";
      continue;
    }
    // Files build in well under a second; hold each card briefly so the
    // sequence stays readable instead of flashing past.
    const elapsed = Date.now() - started;
    if (elapsed < 320) await new Promise((r) => setTimeout(r, 320 - elapsed));
    clearInterval(creep);
    bar.style.width = "100%";
    setStatus(card, "ready", "Ready");
    card.classList.add("ready");
    card.querySelectorAll(".output-actions .btn").forEach((b) => (b.disabled = false));
  }

  const failed = selected.filter((k) => !state.blobs[k]);
  document.getElementById("gen-title").textContent = failed.length
    ? "Some files could not be built"
    : "Your deal package is ready";
  document.getElementById("gen-sub").textContent = failed.length
    ? "Try again, or download the files that did build."
    : "Preview any file below, or download it. Every file is already built, so downloads are instant.";
  document.getElementById("final-actions").style.visibility = "visible";
}

// Which deliverables the user actually asked for
function selectedOutputs() {
  const map = { "c-summary": "summary", "c-guide": "guide", "c-synergy": "synergy", "c-model": "model" };
  return Object.entries(map)
    .filter(([id]) => document.getElementById(id).checked)
    .map(([, key]) => key);
}

// One place that knows how to build each deliverable
async function buildOutput(key, deal, m) {
  // With live filing data the summary becomes a diagnosis memo, which
  // answers a real question instead of describing the company.
  if (key === "summary") {
    return deal.analysis ? generateDiagnosisMemoDocx(deal, m) : generateSummaryDocx(deal, m);
  }
  if (key === "guide") return generateGuideDocx(deal, m);
  if (key === "synergy") {
    // With live filings the deck follows the diagnosis narrative
    return deal.analysis ? generateDiagnosisDeckPptx(deal, m) : generateSynergyPptx(deal, m);
  }
  if (key === "model") {
    return { model: generateModelXlsx(deal, m), email: await generateEmailDocx(deal, m) };
  }
  throw new Error("unknown output " + key);
}

function setStatus(card, kind, label) {
  const el = card.querySelector(".output-status");
  el.className = "output-status status-" + kind;
  el.textContent = label;
}

/* ---------- Downloads ---------- */

async function downloadOutput(kind, btn) {
  const deal = state.deal || getDeal();
  const m = state.metrics || computeMetrics(state.financials || SAMPLE_FINANCIALS);
  const base = slug(deal.company);
  const original = btn ? btn.textContent : null;
  if (btn) { btn.disabled = true; btn.textContent = "Saving..."; }

  try {
    // Files were built during generation; only build here if a user
    // reached a download without going through that step.
    const blobs = state.blobs || (state.blobs = {});
    const get = async (k) => (blobs[k] = blobs[k] || (await buildOutput(k, deal, m)));

    if (kind === "summary") downloadBlob(await get("summary"), `${base}_Business_Summary.docx`);
    else if (kind === "guide") downloadBlob(await get("guide"), `${base}_Interview_Guide.docx`);
    else if (kind === "synergy") downloadBlob(await get("synergy"), `${base}_Synergy_Presentation.pptx`);
    else if (kind === "model") {
      const pair = await get("model");
      downloadBlob(pair.model, `${base}_Deal_Model.xlsx`);
      downloadBlob(pair.email, `${base}_Email_Summary.docx`);
    } else if (kind === "package") {
      const zip = new JSZip();
      const keys = selectedOutputs();
      for (const k of keys) {
        const b = await get(k);
        if (k === "summary") zip.file(`${base}_Business_Summary.docx`, b);
        if (k === "guide") zip.file(`${base}_Interview_Guide.docx`, b);
        if (k === "synergy") zip.file(`${base}_Synergy_Presentation.pptx`, b);
        if (k === "model") {
          zip.file(`${base}_Deal_Model.xlsx`, b.model);
          zip.file(`${base}_Email_Summary.docx`, b.email);
        }
      }
      downloadBlob(await zip.generateAsync({ type: "blob" }), `${base}_Deal_Package.zip`);
    }
  } catch (e) {
    console.error("Download failed:", e);
    alert("Sorry, that file could not be saved. Check the console for details.");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = original; }
  }
}

/* ---------- Preview modal ---------- */

function previewOutput(key) {
  setTheme(currentTheme());
  const deal = state.deal || getDeal();
  const m = state.metrics || computeMetrics(state.financials || SAMPLE_FINANCIALS);
  const P = practiceOf(deal);
  const email = buildEmail(deal, m);
  const srcNote =
    m.source === "sample"
      ? "Built from the sample company financials. Upload your own numbers for tailored output."
      : m.source === "lookup"
      ? "Built from reported figures found by the company lookup."
      : "Built from your financials, entirely in your browser.";

  const VISUAL = {
    summary: () => previewSummaryHTML(deal, m),
    guide: () => previewGuideHTML(deal, m),
    synergy: () => previewDeckHTML(deal, m),
    model: () =>
      previewModelHTML(deal, m) +
      '<div class="pv-divider">Executive email summary, sent with the model</div>' +
      previewEmailHTML(deal, m),
  };

  const OVERVIEW = {
    summary: () => {
      const n = buildNarrative(deal, m);
      return `<h4>Situation overview</h4><ul><li>${n.dealFrame}</li></ul>
        <h4>Financial profile</h4><ul><li>${n.growthSentence}</li><li>${n.marginSentence}</li></ul>
        <h4>Also included</h4><ul>
        <li>Historical financials table with growth and margin rows (${m.firstYear} to ${m.lastYear})</li>
        <li>Key people, key products, and recent news sections</li>
        <li>Risks and open questions prioritized for this situation</li></ul>`;
    },
    guide: () =>
      buildInterviewQuestions(deal)
        .map(
          (s) =>
            `<h4>${s.section}</h4><ul>` +
            s.questions.slice(0, 2).map((q) => `<li>${q}</li>`).join("") +
            `</ul>`
        )
        .join(""),
    synergy: () => {
      const groups = buildOpportunities(deal, m).groups;
      return (
        groups
          .map(
            (g) =>
              `<h4>${g.label}</h4><ul>` +
              g.items.map((s) => `<li><strong>${s.name}</strong>: ${s.rationale}</li>`).join("") +
              `</ul>`
          )
          .join("") + `<h4>Plus</h4><ul><li>Cover, summary signpost slide, and an achievability matrix appendix</li></ul>`
      );
    },
    model: () => {
      if (deal.analysis) {
        const a = deal.analysis;
        const specs = modelSpecs(deal, m);
        const scenarioTab = specs[2] ? specs[2].name : "Scenario";
        const line = a.practice === "strategy" ? "the growth rate" : a.practice === "financial" ? "what the earnings base is worth" : "operating margin";
        return `<h4>Workbook structure</h4>
          <ul><li><strong>${specs[0].name}</strong>: the question, how to read the model, the headline figures as live links, and the sources</li>
          <li><strong>${specs[1].name}</strong>: ${a.withMargin.length} quarters of revenue and operating income exactly as filed, with margin and year over year computed</li>
          <li><strong>${scenarioTab}</strong>: bear, base, and bull paths for the next four quarters, testing ${line}, each anchored to the same quarter a year earlier</li></ul>
          <h4>Email summary</h4><ul><li><strong>Subject:</strong> ${email.subject}</li><li>${email.answer}</li></ul>`;
      }
      const sc = computeScenarios(deal, m);
      const term = practiceTerm(deal);
      return `<h4>Workbook structure</h4>
        <ul><li><strong>Inputs</strong>: client data, projection assumptions, ${term} percentages, and valuation multiples, all editable</li>
        <li><strong>Model</strong>: historicals plus a five year projection</li>
        <li><strong>${term.charAt(0).toUpperCase() + term.slice(1)}</strong>: conservative ${fmtM(sc.conservative.total)}, midpoint ${fmtM(sc.midpoint.total)}, aggressive ${fmtM(sc.aggressive.total)} annual impact</li>
        <li><strong>Valuation</strong>: enterprise value, value creation, and a sensitivity grid</li></ul>
        <h4>Email summary</h4><ul><li><strong>Subject:</strong> ${email.subject}</li><li>${email.answer}</li></ul>`;
    },
  };

  // The preview modal reads the card it was opened from, so the two
  // can never drift apart.
  const cardTitle = (key, fallback) => {
    const el = document.querySelector(`.output-card[data-output="${key}"] h3`);
    return el && el.textContent.trim() ? el.textContent.trim() : fallback;
  };
  const titles = {
    summary: cardTitle("summary", "Business Summary"),
    guide: cardTitle("guide", guideTitleOf(deal)),
    synergy: cardTitle("synergy", P.deckShort),
    model: cardTitle("model", "Excel Model + Email Summary"),
  };

  document.getElementById("modal-title").textContent = titles[key];
  document.getElementById("modal-body").innerHTML =
    `<div class="pv-switch">
       <button class="pv-swbtn active" data-pane="visual" onclick="pvPane('visual')">Preview</button>
       <button class="pv-swbtn" data-pane="overview" onclick="pvPane('overview')">Overview</button>
       <span class="pv-note">${srcNote}</span>
     </div>
     <div id="pv-pane-visual">${VISUAL[key]()}</div>
     <div id="pv-pane-overview" hidden><div class="pv-overview">${OVERVIEW[key]()}</div></div>`;
  document.getElementById("modal").classList.add("open");
}

function pvPane(which) {
  document.getElementById("pv-pane-visual").hidden = which !== "visual";
  document.getElementById("pv-pane-overview").hidden = which !== "overview";
  document.querySelectorAll(".pv-swbtn").forEach((b) => b.classList.toggle("active", b.dataset.pane === which));
}

function closeModal() {
  document.getElementById("modal").classList.remove("open");
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeModal();
});

initPickers();
initBrandingUI();
initCompanyField();
refreshGoalOptions();
