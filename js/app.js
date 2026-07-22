/* ============ DealDesk app: state, workflow, generation ============ */

const state = {
  uploads: { financials: [], org: [], notes: [] },
  financials: null, // parsed {years, series, source}
  metrics: null,
  deal: null,
  brandLogo: null, // {b64, mime, w, h}
};

const FONT_PAIRS = {
  georgia: { serif: "Georgia", sans: "Arial" },
  garamond: { serif: "Garamond", sans: "Calibri" },
  times: { serif: "Times New Roman", sans: "Segoe UI" },
  modern: { serif: "Arial", sans: "Arial" },
};

// Build the theme for the selected output mode: null = ghost draft
function currentTheme() {
  const mode = document.querySelector('input[name="output-mode"]:checked');
  if (!mode || mode.value === "ghost") return null;
  const fonts = FONT_PAIRS[document.getElementById("f-fonts").value] || FONT_PAIRS.georgia;
  return {
    ink: document.getElementById("f-brand-primary").value,
    accent: document.getElementById("f-brand-accent").value,
    serif: fonts.serif,
    sans: fonts.sans,
    logo: state.brandLogo,
  };
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
  ["financials", "org", "notes"].forEach((zone) => {
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

  // Client logo: read as base64, capture natural dimensions for aspect ratio
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
        const list = document.getElementById("files-logo");
        list.innerHTML = "";
        const chip = document.createElement("div");
        chip.className = "file-chip";
        chip.textContent = "✓ " + file.name;
        list.appendChild(chip);
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
    logoInput.value = "";
  });
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
  const practice = document.querySelector('input[name="practice"]:checked');
  return {
    practice: practice ? practice.value : "financial",
    company: val("f-company", "Meridian Components Inc."),
    industry: val("f-industry", "Industrial manufacturing"),
    dealType: document.getElementById("f-dealtype").value,
    revenueBand: document.getElementById("f-revenue").value,
    context: document.getElementById("f-context").value.trim(),
  };
}

function fillReview() {
  const deal = getDeal();
  document.getElementById("r-practice").textContent = practiceOf(deal).label;
  document.getElementById("r-company").textContent = deal.company;
  document.getElementById("r-industry").textContent = deal.industry;
  document.getElementById("r-dealtype").textContent = deal.dealType;
  document.getElementById("r-revenue").textContent = deal.revenueBand;
  const total = Object.values(state.uploads).reduce((n, arr) => n + arr.length, 0);
  document.getElementById("r-files").textContent =
    total > 0 ? total + " file" + (total > 1 ? "s" : "") : "None (sample data will be used)";
}

/* ---------- Generation ---------- */

function runGeneration() {
  state.deal = getDeal();
  if (!state.financials) state.financials = SAMPLE_FINANCIALS;
  state.metrics = computeMetrics(state.financials);

  // Retitle the output cards for the selected practice
  const P = practiceOf(state.deal);
  document.querySelector('.output-card[data-output="guide"] h3').textContent = P.guideTitle;
  document.querySelector('.output-card[data-output="synergy"] h3').textContent = P.deckShort;

  nextStep(4);
  document.getElementById("gen-title").textContent = "Generating your deal package";
  document.getElementById("gen-sub").textContent =
    "Each deliverable is drafted from your inputs, entirely in your browser. Nothing is uploaded to a server.";
  document.getElementById("final-actions").style.visibility = "hidden";

  const cards = Array.from(document.querySelectorAll(".output-card"));
  cards.forEach((card) => {
    card.classList.remove("ready");
    setStatus(card, "pending", "Queued");
    card.querySelector(".output-bar i").style.width = "0";
    card.querySelectorAll(".output-actions .btn").forEach((b) => (b.disabled = true));
  });

  cards.forEach((card, i) => {
    setTimeout(() => animateCard(card), 400 + i * 700);
  });

  const totalTime = 400 + cards.length * 700 + 2200;
  setTimeout(() => {
    document.getElementById("gen-title").textContent = "Your deal package is ready";
    document.getElementById("gen-sub").textContent =
      "Preview each deliverable below, or download real Word, PowerPoint, and Excel files.";
    document.getElementById("final-actions").style.visibility = "visible";
  }, totalTime);
}

function animateCard(card) {
  setStatus(card, "working", "Drafting");
  const bar = card.querySelector(".output-bar i");
  let pct = 0;
  const tick = setInterval(() => {
    pct += 10 + Math.random() * 16;
    if (pct >= 100) {
      pct = 100;
      clearInterval(tick);
      setStatus(card, "ready", "Ready");
      card.classList.add("ready");
      card.querySelectorAll(".output-actions .btn").forEach((b) => (b.disabled = false));
    }
    bar.style.width = pct + "%";
  }, 220);
}

function setStatus(card, kind, label) {
  const el = card.querySelector(".output-status");
  el.className = "output-status status-" + kind;
  el.textContent = label;
}

/* ---------- Downloads ---------- */

async function downloadOutput(kind, btn) {
  setTheme(currentTheme());
  const deal = state.deal || getDeal();
  const m = state.metrics || computeMetrics(state.financials || SAMPLE_FINANCIALS);
  const base = slug(deal.company);
  const original = btn ? btn.textContent : null;
  if (btn) { btn.disabled = true; btn.textContent = "Building..."; }

  try {
    if (kind === "summary") {
      downloadBlob(await generateSummaryDocx(deal, m), `${base}_Business_Summary.docx`);
    } else if (kind === "guide") {
      downloadBlob(await generateGuideDocx(deal, m), `${base}_Interview_Guide.docx`);
    } else if (kind === "synergy") {
      downloadBlob(await generateSynergyPptx(deal, m), `${base}_Synergy_Presentation.pptx`);
    } else if (kind === "model") {
      downloadBlob(generateModelXlsx(deal, m), `${base}_Deal_Model.xlsx`);
      downloadBlob(await generateEmailDocx(deal, m), `${base}_Email_Summary.docx`);
    } else if (kind === "package") {
      downloadBlob(await generatePackageZip(deal, m), `${base}_Deal_Package.zip`);
    }
  } catch (e) {
    console.error("Generation failed:", e);
    alert("Sorry, that document failed to generate. Check the console for details.");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = original; }
  }
}

/* ---------- Preview modal ---------- */

function previewOutput(key) {
  const deal = state.deal || getDeal();
  const m = state.metrics || computeMetrics(state.financials || SAMPLE_FINANCIALS);
  const P = practiceOf(deal);
  const n = buildNarrative(deal, m);
  const email = buildEmail(deal, m);
  const srcNote =
    m.source === "sample"
      ? "Built from the sample company financials. Upload your own Excel file for tailored output."
      : "Built from your uploaded financials, entirely in your browser.";

  const PREVIEWS = {
    summary: {
      title: "Business Summary (preview)",
      body: `
        <div class="preview-note">${srcNote}</div>
        <h4>Situation overview</h4><ul><li>${n.dealFrame}</li></ul>
        <h4>Financial profile</h4>
        <ul><li>${n.growthSentence}</li><li>${n.marginSentence}</li></ul>
        <h4>Also included</h4>
        <ul><li>Historical financials table (${m.firstYear} to ${m.lastYear})</li>
        <li>Key risks and open questions tailored to a ${deal.dealType.toLowerCase()}</li>
        <li>Recommended next steps</li></ul>`,
    },
    guide: {
      title: `${P.guideTitle} (preview)`,
      body:
        `<div class="preview-note">${srcNote}</div>` +
        buildInterviewQuestions(deal)
          .map(
            (s) =>
              `<h4>${s.section}</h4><ul>` +
              s.questions.slice(0, 2).map((q) => `<li>${q}</li>`).join("") +
              `</ul>`
          )
          .join(""),
    },
    synergy: {
      title: `${P.deckShort} (preview)`,
      body: (() => {
        const groups = buildOpportunities(deal, m).groups;
        return (
          `<div class="preview-note">${srcNote}</div>` +
          groups
            .map(
              (g) =>
                `<h4>${g.label}</h4><ul>` +
                g.items.map((s) => `<li><strong>${s.name}</strong>: ${s.rationale}</li>`).join("") +
                `</ul>`
            )
            .join("") +
          `<h4>Plus</h4><ul><li>Executive summary and achievability matrix slides</li></ul>`
        );
      })(),
    },
    model: {
      title: "Excel Model + Email Summary (preview)",
      body: (() => {
        const sc = computeScenarios(deal, m);
        const term = practiceTerm(deal);
        return `
        <div class="preview-note">${srcNote}</div>
        <h4>Workbook structure</h4>
        <ul><li><strong>Inputs</strong>: client data plus every ${term} percentage in a labeled, editable cell</li>
        <li><strong>${term.charAt(0).toUpperCase() + term.slice(1)} model</strong>: conservative ${fmtM(sc.conservative.total)}, midpoint ${fmtM(sc.midpoint.total)}, aggressive ${fmtM(sc.aggressive.total)} annual EBIT impact, all formulas, no hardcoding</li>
        <li><strong>Historicals</strong>: the parsed financials (${m.firstYear} to ${m.lastYear})</li></ul>
        <h4>Email summary (excerpt)</h4>
        <ul><li><strong>Subject:</strong> ${email.subject}</li>
        <li>${email.answer}</li></ul>`;
      })(),
    },
  };

  const p = PREVIEWS[key];
  document.getElementById("modal-title").textContent = p.title;
  document.getElementById("modal-body").innerHTML = p.body;
  document.getElementById("modal").classList.add("open");
}

function closeModal() {
  document.getElementById("modal").classList.remove("open");
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeModal();
});

initPickers();
