(() => {
  /* ====== CONSTANTS ====== */
  const PX_PER_CM = 37.7952755906; // 96dpi
  const PREVIEW_GAP_CM_NUM = 0; // geen ruimte tussen etiketten in preview
  const PDF_MARGIN_CM = 0.5; // witmarge rondom in PDF
  const LABEL_PADDING_CM = 0.5; // binnenmarge in label (cm)

  // Typografie / fit
  const WRAP_THRESHOLD_PX = 10; // onder 10px pas zachte afbreking aanzetten
  const MIN_FS_PX = 6; // noodrem bij heel kleine labels
  const CODE_MULT = 1.6; // productcode ≈ 1.6 × body

  // PDF rand (alleen voor capture-visual; fysieke rand komt uit DOM-stijl)
  const BORDER_PX = 2;

  /* ====== DOM HOOKS ====== */
  const labelsGrid = document.getElementById("labelsGrid");
  const controlInfo = document.getElementById("controlInfo");
  const canvasEl = document.getElementById("canvas");
  const btnGen = document.getElementById("btnGenerate");
  const btnPDF = document.getElementById("btnPDF");

  // Batch DOM (optioneel aanwezig)
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const btnPickFile = document.getElementById("btnPickFile");
  const btnTemplateCsv = document.getElementById("btnTemplateCsv");
  const btnTemplateXlsx = document.getElementById("btnTemplateXlsx");
  const mappingWrap = document.getElementById("mappingWrap");
  const mappingGrid = document.getElementById("mappingGrid");
  const previewWrap = document.getElementById("previewWrap");
  const tablePreview = document.getElementById("tablePreview");
  const normWrap = document.getElementById("normWrap");
  const chkComma = document.getElementById("optCommaDecimal");
  const chkTrim = document.getElementById("optTrimSpaces");
  const batchControls = document.getElementById("batchControls");
  const btnRunBatch = document.getElementById("btnRunBatch");
  const btnAbortBatch = document.getElementById("btnAbortBatch");
  const progressWrap = document.getElementById("progressWrap");
  const progressBar = document.getElementById("progressBar");
  const progressLabel = document.getElementById("progressLabel");
  const progressPhase = document.getElementById("progressPhase");
  const logWrap = document.getElementById("logWrap");
  const logList = document.getElementById("logList");
  const labelFactory = document.getElementById("labelFactory");

  /* ====== STATE ====== */
  let currentPreviewScale = 1;
  let parsedRows = [];
  let headers = [];
  let mapping = {};
  let abortFlag = false;

  /* ====== HELPERS ====== */
  const el = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === "style" && typeof v === "object") Object.assign(node.style, v);
      else if (k.startsWith("on") && typeof v === "function")
        node.addEventListener(k.substring(2), v);
      else node.setAttribute(k, v);
    });
    children
      .flat()
      .forEach((c) =>
        node.appendChild(typeof c === "string" ? document.createTextNode(c) : c)
      );
    return node;
  };
  const line = (lab, val) => [
    el("div", { class: "lab" }, lab),
    el("div", { class: "val" }, val),
  ];
  const pad2 = (n) => String(n).padStart(2, "0");
  const ts = (d = new Date()) =>
    `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(
      d.getHours()
    )}.${pad2(d.getMinutes())}.${pad2(d.getSeconds())}`;

  // Compat-wrapper: project kan readValues óf readValuesSingle hebben
  function getFormValues() {
    if (typeof readValues === "function") return readValues();
    if (typeof readValuesSingle === "function") return readValuesSingle();
    throw new Error("readValues / readValuesSingle niet gevonden");
  }

  // 1 frame wachten zodat layout/afmetingen kloppen
  function nextFrame() {
    return new Promise((r) => requestAnimationFrame(r));
  }

  // Fit alle label-inhouden in container
  function fitAllIn(container) {
    container.querySelectorAll(".label-inner").forEach((inner) => {
      inner.classList.add("nowrap-mode");
      inner.classList.remove("softwrap-mode");

      // font-fit Top-Box + Detail-Box
      fitContentToBoxConditional(inner);

      // C/N-lijn breedte laten volgen op "IN CHINA"
      updateCnLine(inner);
    });
  }

  // Extra robuust: twee fit-rondes met tussentijds frame
  async function mountThenFit(container) {
    await nextFrame();
    fitAllIn(container);
    await nextFrame();
    fitAllIn(container);
  }

  /* ====== ENKELVOUDIG INLEZEN ====== */
  function readValuesSingle() {
    const get = (id) => document.getElementById(id).value.trim();

    const L = parseFloat(get("boxLength"));
    const W = parseFloat(get("boxWidth"));
    const H = parseFloat(get("boxHeight"));

    const vals = {
      L,
      W,
      H,
      code: get("prodCode"),
      desc: get("prodDesc"),
      ean: get("ean"),
      qty: String(Math.max(0, Math.floor(Number(get("qty")) || 0))),
      gw: ((v) => (isFinite(+v) ? (+v).toFixed(2) : v))(get("gw")),
      cbm: get("cbm"),
      batch: get("batch"), // VERPLICHT
    };

    // Verplichte velden check
    if (
      !vals.code ||
      !vals.desc ||
      !vals.ean ||
      !vals.qty ||
      !vals.gw ||
      !vals.cbm ||
      !vals.batch
    ) {
      throw new Error("Vul alle verplichte velden in (inclusief Batch).");
    }

    // Maatvalidatie: 5–100 cm
    ["L", "W", "H"].forEach((k) => {
      const v = vals[k];
      if (!isFinite(v) || v < 5 || v > 100) {
        throw new Error(
          "Lengte (L), Breedte (W) en Hoogte (H) moeten tussen 5 en 100 cm liggen (5–100)."
        );
      }
    });

    return vals;
  }

  /* ====== LABELMATEN & PREVIEW SCALE ====== */
  function computeLabelSizes({ L, W, H }) {
    // 10% kleiner aan elke zijde
    const lw = Math.max(5, Math.min(100, L));
    const ww = Math.max(5, Math.min(100, W));
    const hh = Math.max(5, Math.min(100, H));

    const fb = { w: lw * 0.9, h: hh * 0.9 }; // front/back = L × H
    const sd = { w: ww * 0.9, h: hh * 0.9 }; // side       = W × H

    return [
      { idx: 1, kind: "front/back", ...fb }, // C/N straks hier
      { idx: 2, kind: "front/back", ...fb }, // C/N straks hier
      { idx: 3, kind: "side", ...sd }, // Made in China hier
      { idx: 4, kind: "side", ...sd }, // Made in China hier
    ];
  }

  function updateControlInfo(sizes) {
    const n2 = (x) => (Math.round(x * 100) / 100).toFixed(2);
    const [s1, s2, s3, s4] = sizes;
    controlInfo.innerHTML = `
      <h3>Berekende labelafmetingen (werkelijke cm)</h3>
      <div class="control-grid-2x2">
        <div class="control-item">Etiket 1 (${s1.kind}): ${n2(s1.w)} × ${n2(
      s1.h
    )} cm</div>
        <div class="control-item">Etiket 3 (${s3.kind}): ${n2(s3.w)} × ${n2(
      s3.h
    )} cm</div>
        <div class="control-item">Etiket 2 (${s2.kind}): ${n2(s2.w)} × ${n2(
      s2.h
    )} cm</div>
        <div class="control-item">Etiket 4 (${s4.kind}): ${n2(s4.w)} × ${n2(
      s4.h
    )} cm</div>
      </div>`;
  }

  function computePreviewScale(sizes) {
    const gapPx = PREVIEW_GAP_CM_NUM * PX_PER_CM;
    const w1 = sizes[0].w * PX_PER_CM,
      w3 = sizes[2].w * PX_PER_CM;
    const requiredW = Math.max(w1 + gapPx + w1, w3 + gapPx + w3);
    const cs = getComputedStyle(canvasEl);
    const innerW =
      canvasEl.clientWidth -
      parseFloat(cs.paddingLeft) -
      parseFloat(cs.paddingRight);
    return Math.min(innerW / requiredW, 1);
  }

  /* ====== FONT-FIT ===========================================
   Doelen:
   - Grotere startwaarde bij grote etiketten → zichtbaar grotere tekst.
   - Eerst agressief omhoog groeien, dán pas binair finetunen.
   - Wrap pas inzetten als body < WRAP_THRESHOLD_PX (no-wrap voorkeur).
   - Houd een kleine “guard” (binnenmarge) aan tegen clipping.
   - Code-box schaalt mee (geen cap), body via --fs.
========================================================================= */

  /** past alles binnen 'innerEl' met een veiligheidsmarge? */
  function fitsWithGuard(innerEl, guardX, guardY) {
    return (
      innerEl.scrollWidth <= innerEl.clientWidth - guardX &&
      innerEl.scrollHeight <= innerEl.clientHeight - guardY
    );
  }

  function fitsTopAndDetail(innerEl, guardX, guardY) {
    const topBox = innerEl.querySelector(".top-box");
    const detailBox = innerEl.querySelector(".detail-box");
    const detailInner = innerEl.querySelector(".detail-box-inner");

    // Als we de nieuwe structuur niet vinden, val terug op de oude check:
    if (!topBox || !detailBox || !detailInner) {
      return fitsWithGuard(innerEl, guardX, guardY);
    }

    // TOP-BOX: ERP + omschrijving
    // Minimaal ~8% verticale marge zodat het nooit tegen de rand geplakt zit
    const topGuardY = Math.max(guardY, topBox.clientHeight * 0.08);

    const topOk =
      topBox.scrollWidth <= topBox.clientWidth - guardX &&
      topBox.scrollHeight <= topBox.clientHeight - topGuardY;

    if (!topOk) return false;

    // DETAIL-BOX: EAN/QTY/.../Made in China
    // Minimaal ~7% marge zodat de onderste regel niet de onderrand raakt
    const detailGuardY = Math.max(guardY, detailBox.clientHeight * 0.07);

    const detailOkBox =
      detailInner.scrollWidth <= detailBox.clientWidth - guardX &&
      detailInner.scrollHeight <= detailBox.clientHeight - detailGuardY;

    if (!detailOkBox) return false;

    // EXTRA: geen enkele detail-value (inclusief EAN) mag horizontaal overlopen
    const values = detailInner.querySelectorAll(".detail-value");
    for (const v of values) {
      if (v.scrollWidth > v.clientWidth - guardX) {
        return false; // font is nog te groot → verder verkleinen
      }
    }

    return true;
  }

  function updateCnLine(innerEl) {
    const detailInner = innerEl.querySelector(".detail-box-inner");
    const cnLine = innerEl.querySelector(".cn-line");

    if (!detailInner || !cnLine) return;

    // Meet-probe met tekst "IN CHINA" in dezelfde context als de details
    const probe = document.createElement("span");
    probe.textContent = "IN CHINA";
    probe.style.visibility = "hidden";
    probe.style.position = "absolute";
    probe.style.whiteSpace = "nowrap";

    detailInner.appendChild(probe);
    const width = probe.getBoundingClientRect().width;
    detailInner.removeChild(probe);

    if (width > 0) {
      cnLine.style.width = width + "px";
    }
  }

  function searchBaseFontSize(innerEl, minFs, startHi, guardX, guardY) {
    // 1) agressief omhoog groeien vanaf startHi
    applyFontSizes(innerEl, startHi);
    if (fitsTopAndDetail(innerEl, guardX, guardY)) {
      let grow = startHi;
      for (let i = 0; i < 48; i++) {
        const next = grow * 1.08;
        applyFontSizes(innerEl, next);
        if (!fitsTopAndDetail(innerEl, guardX, guardY)) {
          applyFontSizes(innerEl, grow); // stap terug naar laatste passende
          return grow;
        }
        grow = next;
      }
      return grow; // plafond bereikt zonder clip
    }

    // 2) startHi paste al niet → binair omlaag tussen [minFs, startHi]
    let lo = minFs,
      hi = startHi,
      best = lo;
    while (hi - lo > 0.5) {
      const mid = (lo + hi) / 2;
      applyFontSizes(innerEl, mid);
      if (fitsTopAndDetail(innerEl, guardX, guardY)) {
        best = mid;
        lo = mid;
      } else {
        hi = mid;
      }
    }
    applyFontSizes(innerEl, best);
    return best;
  }

  //const FONT_STEP_PX = 4;  // 1 stap = 4px, ERP = basis + 2 stappen

  function applyFontSizes(innerEl, fsPx) {
    // Bodytekst (details, productomschrijving) = basis
    const base = fsPx;

    // ERP/productnaam: ~1,4 × basis (tussen 1,3 en 1,6)
    const erp = base * 1.4;

    // Detailtekst: gelijk aan basis
    const detail = base;

    // Kleinere details (bijv. batch/datum): ~0,75 × basis
    const small = base * 0.75;

    innerEl.style.setProperty("--fs-base", base + "px");
    innerEl.style.setProperty("--fs-erp", erp + "px");
    innerEl.style.setProperty("--fs-detail", detail + "px");
    innerEl.style.setProperty("--fs-small", small + "px");

    // Backwards compatibiliteit
    innerEl.style.setProperty("--fs", base + "px");
  }

  /** zoek een passende fontgrootte (eerst groeien, daarna finetunen naar beneden) */
  function searchFontSize(innerEl, minFs, startHi, guardX, guardY) {
    // 1) agressief omhoog groeien vanaf startHi (groeifactor 1.08)
    applyFontSizes(innerEl, startHi);
    if (fitsWithGuard(innerEl, guardX, guardY)) {
      let grow = startHi;
      for (let i = 0; i < 48; i++) {
        const next = grow * 1.08; // iets sneller groeien
        applyFontSizes(innerEl, next);
        if (!fitsWithGuard(innerEl, guardX, guardY)) {
          applyFontSizes(innerEl, grow); // stap terug naar laatste passende
          return grow;
        }
        grow = next;
      }
      return grow; // plafond bereikt zonder clip
    }

    // 2) paste startHi al niet? dan binair omlaag tussen [minFs, startHi]
    let lo = minFs,
      hi = startHi,
      best = lo;
    while (hi - lo > 0.5) {
      const mid = (lo + hi) / 2;
      applyFontSizes(innerEl, mid);
      if (fitsWithGuard(innerEl, guardX, guardY)) {
        best = mid;
        lo = mid;
      } else {
        hi = mid;
      }
    }
    applyFontSizes(innerEl, best);
    return best;
  }

  /** hoofd-fit: eerst no-wrap ≥ WRAP_THRESHOLD_PX, anders soft-wrap ≥ MIN_FS_PX */
  function fitContentToBoxConditional(innerEl) {
    // Effectieve box (na padding)
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;

    // Veiligheidsmarges (px): 2% van kant + absolute ondergrens
    const guardX = Math.max(8, w * 0.02);
    const guardY = Math.max(8, h * 0.02);

    // BASIS FONTGROOTTE OP BASIS VAN HOOGTE:
    // Bodytekst ≈ 10% van de label-hoogte (in px)
    const baseFromBox = h * 0.1; // jouw "labelHeightPx * 0.10"
    const startHi = Math.max(16, baseFromBox);

    // Fase 1: no-wrap (voorkeur)
    innerEl.classList.add("nowrap-mode");
    innerEl.classList.remove("softwrap-mode");

    let fs = searchBaseFontSize(
      innerEl,
      WRAP_THRESHOLD_PX, // ondergrens voor no-wrap
      startHi,
      guardX,
      guardY
    );

    // Als we boven de wrap-drempel blijven is dit prima
    if (fs >= WRAP_THRESHOLD_PX) {
      return;
    }

    // Fase 2: soft-wrap (als we kleiner dan WRAP_THRESHOLD_PX moesten)
    innerEl.classList.remove("nowrap-mode");
    innerEl.classList.add("softwrap-mode");

    fs = searchBaseFontSize(innerEl, MIN_FS_PX, fs, guardX, guardY);

    // Noodrem
    if (fs < MIN_FS_PX) {
      fs = MIN_FS_PX;
      applyFontSizes(innerEl, fs);
    }
  }

  /* ====== UI OPBOUW ====== */
  function buildLeftBlock(values, size) {
    const rowsData = [
      { label: "EAN:", value: values.ean || "" },
      { label: "QTY:", value: values.qty || "" },
      { label: "G.W:", value: values.gw || "" },
      { label: "CBM:", value: values.cbm || "" },
      { label: "Batch:", value: values.batch || "" },
      { label: "C/N:", value: null, kind: "cn" },
      { label: "Made in China", value: "", kind: "made" },
    ];

    const detailBoxInner = el("div", { class: "detail-box-inner" });

    rowsData.forEach((r) => {
      const rowClasses = ["detail-row"];
      if (r.kind === "cn") rowClasses.push("cn-row");
      if (r.kind === "made") rowClasses.push("made-row");

      const row = el("div", { class: rowClasses.join(" ") });

      const labelEl = el("div", { class: "detail-label" }, r.label);

      let valueEl;
      if (r.kind === "cn") {
        // C/N: lijn om op te schrijven
        const lineEl = el("span", { class: "cn-line" });
        valueEl = el("div", { class: "detail-value" }, lineEl);
      } else {
        // normale cases, inclusief Made in China (lege value)
        valueEl = el("div", { class: "detail-value" }, r.value || "");
      }

      row.append(labelEl, valueEl);
      detailBoxInner.append(row);
    });

    return detailBoxInner;
  }

  function createLabelEl(size, values, previewScale) {
    // Afmeting in pixels voor de PREVIEW (geschaald naar het canvas)
    const widthPx = Math.round(size.w * PX_PER_CM * previewScale);
    const heightPx = Math.round(size.h * PX_PER_CM * previewScale);

    const wrap = el("div", { class: "label-wrap" });
    const label = el("div", {
      class: "label",
      style: { width: widthPx + "px", height: heightPx + "px" },
    });
    label.dataset.idx = String(size.idx);

    const inner = el("div", { class: "label-inner nowrap-mode" });

    // Padding op de label-rand, mee schalen met de preview
    const padPx = LABEL_PADDING_CM * PX_PER_CM * previewScale;
    label.style.padding = padPx + "px";

    // --- TOP-BOX: ERP-box boven, daaronder productomschrijving ---
    const topBox = el(
      "div",
      { class: "top-box" },
      el(
        "div",
        { class: "erp-box" },
        el("div", { class: "code-box line" }, values.code)
      ),
      el("div", { class: "product-desc line" }, values.desc)
    );

    // --- BOTTOM-BOX: Detail-Box met de bestaande leftblock-inhoud ---
    const detailContent = buildLeftBlock(values, size);
    const detailBox = el("div", { class: "detail-box" }, detailContent);
    const bottomBox = el("div", { class: "bottom-box" }, detailBox);

    // Alles in elkaar klikken
    inner.append(topBox, bottomBox);
    label.append(inner);
    wrap.append(label, el("div", { class: "label-num" }, `Etiket ${size.idx}`));

    return wrap;
  }

  // Bouw een 1:1 label (cm → px, géén previewScale)
  // Dit is de bron voor font-fit en PDF.
  function createMasterLabelEl(size, values) {
    const widthPx = Math.round(size.w * PX_PER_CM); // 1:1
    const heightPx = Math.round(size.h * PX_PER_CM);

    const wrap = el("div", { class: "label-wrap" });
    const label = el("div", {
      class: "label",
      style: { width: widthPx + "px", height: heightPx + "px" },
    });
    label.dataset.idx = String(size.idx);

    const inner = el("div", { class: "label-inner nowrap-mode" });

    // Padding op de label-rand (niet meegeschaald)
    const padPx = LABEL_PADDING_CM * PX_PER_CM;
    label.style.padding = padPx + "px";

    // Referentiematen voor font-fit (kunnen blijven zoals je ze nu hebt)
    const REF_W = 100;
    const REF_H = 60;
    inner.style.setProperty("--ref-w", REF_W + "px");
    inner.style.setProperty("--ref-h", REF_H + "px");

    // Geen --k meer nodig voor master; font-fit werkt op echte px
    inner.style.removeProperty("--k");

    // TOP-BOX: ERP boven, daaronder productomschrijving
    const topBox = el(
      "div",
      { class: "top-box" },
      el(
        "div",
        { class: "erp-box" },
        el("div", { class: "code-box line" }, values.code)
      ),
      el("div", { class: "product-desc line" }, values.desc)
    );

    // BOTTOM-BOX: Detail-Box met bestaande left/right inhoud
    const detailContent = buildLeftBlock(values, size);
    const detailBox = el("div", { class: "detail-box" }, detailContent);
    const bottomBox = el("div", { class: "bottom-box" }, detailBox);

    inner.append(topBox, bottomBox);
    label.append(inner);

    // In de factory hebben we label-num niet nodig, maar kan geen kwaad
    wrap.append(label, el("div", { class: "label-num" }, `Etiket ${size.idx}`));
    return wrap;
  }

  // Maak een geschaalde preview op basis van een masterlabel
  function createPreviewFromMaster(masterWrap, size, previewScale) {
    // masterWrap is de <div class="label-wrap"> uit de factory
    const masterLabel = masterWrap.querySelector(".label");
    const clonedLabel = masterLabel.cloneNode(true);

    const widthPx = size.w * PX_PER_CM;
    const heightPx = size.h * PX_PER_CM;

    const scaledW = widthPx * previewScale;
    const scaledH = heightPx * previewScale;

    // Wrapper voor preview
    const wrap = el("div", { class: "label-wrap" });
    const num = el("div", { class: "label-num" }, `Etiket ${size.idx}`);

    // Schaal het volledige label uniform omlaag
    clonedLabel.style.transformOrigin = "top left";
    clonedLabel.style.transform = `scale(${previewScale})`;

    // Zorg dat de wrapper de geschaalde afmetingen heeft
    wrap.style.width = scaledW + "px";
    wrap.style.height = scaledH + "px";

    wrap.append(clonedLabel, num);
    return wrap;
  }

  /* ====== PREVIEW PIPELINE (single & batch) ====== */
  async function renderPreviewFor(vals) {
    const sizes = computeLabelSizes(vals);

    // 1. Maak 1:1 masterlabels in de factory
    if (labelFactory) {
      labelFactory.innerHTML = "";
      // volgorde is hier nog gewoon 1,2,3,4
      sizes.forEach((size) => {
        const masterWrap = createMasterLabelEl(size, vals);
        labelFactory.appendChild(masterWrap);
      });

      // Font-fit + C/N op de masterlabels
      await mountThenFit(labelFactory);
    }

    // 2. Preview-schaalfactor op basis van echte cm-afmetingen
    const scale = computePreviewScale(sizes);
    currentPreviewScale = scale;

    updateControlInfo(sizes);
    labelsGrid.style.gap = "0";
    labelsGrid.innerHTML = "";

    // 3. Maak geschaalde previews op basis van de masterlabels
    // Volgorde: 1 & 3 boven, 2 & 4 onder (zoals je had)
    const order = [0, 2, 1, 3];
    order.forEach((i) => {
      const size = sizes[i];
      const masterWrap = labelFactory
        ? labelFactory
            .querySelector(`.label-wrap .label[data-idx="${size.idx}"]`)
            ?.closest(".label-wrap")
        : null;

      // safety: als er geen factory is, val terug op oude createLabelEl
      let previewWrap;
      if (masterWrap) {
        previewWrap = createPreviewFromMaster(masterWrap, size, scale);
      } else {
        // fallback naar oude gedrag (voor het geval)
        previewWrap = createLabelEl(size, vals, scale);
      }
      labelsGrid.appendChild(previewWrap);
    });

    return { sizes, scale };
  }

  async function renderSingle() {
    const vals = getFormValues();
    await renderPreviewFor(vals);
  }

  /* ====== jsPDF / html2canvas ====== */
  function loadJsPDF() {
    return new Promise((res, rej) => {
      if (window.jspdf?.jsPDF) return res(window.jspdf.jsPDF);
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js";
      s.onload = () => res(window.jspdf.jsPDF);
      s.onerror = () => rej(new Error("Kon jsPDF niet laden."));
      document.head.appendChild(s);
    });
  }
  function loadHtml2Canvas() {
    return new Promise((res, rej) => {
      if (window.html2canvas) return res(window.html2canvas);
      const s = document.createElement("script");
      s.src =
        "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js";
      s.onload = () => res(window.html2canvas);
      s.onerror = () => rej(new Error("Kon html2canvas niet laden."));
      document.head.appendChild(s);
    });
  }

  async function capturePreviewLabelToImage(h2c, idx) {
    // Eerst proberen in de factory (1:1 masterlabel)
    let src = labelFactory
      ? labelFactory.querySelector(`.label[data-idx="${idx}"]`)
      : null;

    // Fallback: gebruik de zichtbare preview zoals voorheen
    if (!src) {
      src = document.querySelector(`.label[data-idx="${idx}"]`);
    }
    if (!src) throw new Error("Label niet gevonden voor capture.");

    const clone = src.cloneNode(true);

    // Fysieke randen voor PDF
    clone.style.borderTop = `${BORDER_PX}px solid #000`;
    clone.style.borderRight = `${BORDER_PX}px solid #000`;
    clone.style.borderBottom = `${BORDER_PX}px solid #000`;
    clone.style.borderLeft = `${BORDER_PX}px solid #000`;

    const wrap = document.createElement("div");
    wrap.style.position = "fixed";
    wrap.style.left = "-10000px";
    wrap.style.top = "0";
    wrap.style.background = "#fff";

    document.body.appendChild(wrap);
    wrap.appendChild(clone);

    // Cap-scale hoeft geen rekening meer te houden met previewScale
    const capScale = Math.max(2, window.devicePixelRatio || 1);
    const canvas = await h2c(clone, {
      backgroundColor: "#fff",
      scale: capScale,
    });

    // 90° met de klok mee
    const rot = document.createElement("canvas");
    rot.width = canvas.height;
    rot.height = canvas.width;
    const ctx = rot.getContext("2d");
    ctx.translate(rot.width, 0);
    ctx.rotate(Math.PI / 2);
    ctx.drawImage(canvas, 0, 0);

    document.body.removeChild(wrap);
    return rot.toDataURL("image/png");
  }

  /* ====== SINGLE PDF (via preview) ====== */
  async function generatePDFSingle() {
    const vals = getFormValues();
    const { sizes } = await renderPreviewFor(vals);

    const jsPDF = await loadJsPDF();
    const h2c = await loadHtml2Canvas();

    // Na rotatie: breedte = s.h, hoogte = s.w (cm)
    const contentW = Math.max(...sizes.map((s) => s.h));
    const contentH = sizes.reduce((sum, s) => sum + s.w, 0);
    const pageW = contentW + PDF_MARGIN_CM * 2;
    const pageH = contentH + PDF_MARGIN_CM * 2;

    const A4W = 21.0,
      A4H = 29.7;
    const doc = new jsPDF({
      unit: "cm",
      orientation: "portrait",
      format: pageW <= A4W && pageH <= A4H ? "a4" : [pageW, pageH],
    });
    doc.setFont("helvetica", "normal");

    const orderIdx = [1, 3, 2, 4];
    for (let i = 0, y = PDF_MARGIN_CM; i < orderIdx.length; i++) {
      const img = await capturePreviewLabelToImage(h2c, orderIdx[i], i === 0);
      const s = sizes[orderIdx[i] - 1];
      const wRot = s.h,
        hRot = s.w;
      doc.addImage(img, "PNG", PDF_MARGIN_CM, y, wRot, hRot, undefined, "FAST");
      y += hRot;
    }
    doc.save(`${vals.code} - ${ts()}.pdf`);
  }

  /* ====== BATCH PIPELINE (via dezelfde preview) ====== */
  function setHidden(elm, hidden) {
    if (elm) elm.classList.toggle("hidden", hidden);
  }
  function log(msg, type = "info") {
    if (!logList) return;
    const div = el(
      "div",
      { class: type === "error" ? "err" : type === "ok" ? "ok" : "" },
      msg
    );
    logList.appendChild(div);
  }
  function resetLog() {
    if (logList) logList.innerHTML = "";
  }

  async function parseFile(file) {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  }

  const REQUIRED_FIELDS = [
    ["productcode", "ERP"],
    ["omschrijving", "Omschrijving"],
    ["ean", "EAN"],
    ["qty", "QTY"],
    ["gw", "G.W"],
    ["cbm", "CBM"],
    ["lengte", "Length (L)"],
    ["breedte", "Width (W)"],
    ["hoogte", "Height (H)"],
    ["batch", "Batch"],
  ];
  const ALL_FIELDS = [...REQUIRED_FIELDS];

  const SYNONYMS = {
    productcode: [
      "erp",
      "erp#",
      "erp #",
      "productcode",
      "code",
      "sku",
      "artikelcode",
      "itemcode",
      "prodcode",
    ],
    omschrijving: [
      "omschrijving",
      "description",
      "product",
      "naam",
      "title",
      "titel",
    ],
    ean: ["ean", "barcode", "gtin", "ean13", "ean_13"],
    qty: ["qty", "aantal", "quantity", "pcs", "stuks"],
    gw: [
      "gw",
      "g.w",
      "gewicht",
      "weight",
      "grossweight",
      "gweight",
      "brutogewicht",
    ],
    cbm: ["cbm", "m3", "volume", "kub", "kubiekemeter", "kubiekemeters"],
    lengte: ["lengte", "l", "depth", "diepte", "length"],
    breedte: ["breedte", "b", "width", "w"],
    hoogte: ["hoogte", "h", "height"],
    batch: ["batch", "lot", "lotno", "batchno", "batchnr"],
  };

  function slugify(s) {
    return String(s || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "")
      .trim();
  }

  function guessMapping(hdrs) {
    const m = {};
    ALL_FIELDS.forEach(([key]) => (m[key] = ""));
    const slugs = hdrs.map((h) => slugify(String(h).replace(/\([^)]*\)/g, "")));
    for (let i = 0; i < hdrs.length; i++) {
      const h = hdrs[i],
        s = slugs[i];
      for (const [key, syns] of Object.entries(SYNONYMS)) {
        if (syns.includes(s)) {
          m[key] = h;
          break;
        }
      }
    }
    for (const [key] of ALL_FIELDS) {
      if (!m[key]) {
        const sset = (SYNONYMS[key] || [key]).filter((tok) => tok.length >= 2);
        const idx = slugs.findIndex((s) => sset.some((tok) => s.includes(tok)));
        if (idx >= 0) m[key] = hdrs[idx];
      }
    }
    return m;
  }

  function buildMappingUI(hdrs, mapping) {
    if (!mappingGrid) return;
    mappingGrid.innerHTML = "";
    const makeRow = (key, labelText) => {
      const row = el("div", { class: "map-row" });
      const lab = el("label", {}, labelText + " *");
      const sel = el("select", { "data-key": key });
      sel.appendChild(el("option", { value: "" }, "-- kies kolom --"));
      hdrs.forEach((h) => {
        const opt = el("option", { value: h }, h);
        if (mapping[key] === h) opt.selected = true;
        sel.appendChild(opt);
      });
      row.append(lab, sel);
      mappingGrid.appendChild(row);
    };
    REQUIRED_FIELDS.forEach(([k, l]) => makeRow(k, l));
    mappingGrid.querySelectorAll("select").forEach((sel) => {
      sel.addEventListener("change", () => {
        mapping[sel.getAttribute("data-key")] = sel.value;
      });
    });
  }

  function showTablePreview(rows) {
    if (!tablePreview) return;
    if (!rows.length) {
      tablePreview.innerHTML = "<em>Geen data gevonden.</em>";
      return;
    }
    const cols = Object.keys(rows[0]);
    const n = Math.min(5, rows.length);
    const table = el("table");
    const thead = el("thead");
    const trh = el("tr");
    cols.forEach((c) => trh.appendChild(el("th", {}, c)));
    thead.appendChild(trh);
    const tbody = el("tbody");
    for (let i = 0; i < n; i++) {
      const tr = el("tr");
      cols.forEach((c) =>
        tr.appendChild(el("td", {}, String(rows[i][c] ?? "")))
      );
      tbody.appendChild(tr);
    }
    table.append(thead, tbody);
    tablePreview.innerHTML = "";
    tablePreview.appendChild(table);
  }

  function normalizeNumber(val) {
    if (typeof val !== "string") val = String(val ?? "");
    if (chkTrim?.checked) val = val.trim();
    if (chkComma?.checked) val = val.replace(",", ".");
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(val)) {
      val = val.replace(/\./g, "");
    }
    const num = parseFloat(val);
    return isFinite(num) ? num : NaN;
  }

  function readRowWithMapping(row, mapping) {
    const get = (key) => {
      const hdr = mapping[key] || "";
      return hdr ? row[hdr] ?? "" : "";
    };
    const vals = {
      code: String(get("productcode") ?? "").trim(),
      desc: String(get("omschrijving") ?? "").trim(),
      ean: String(get("ean") ?? "").trim(),
      qty: String(Math.max(0, Math.floor(normalizeNumber(String(get("qty")))))),
      gw: (() => {
        const n = normalizeNumber(String(get("gw")));
        return isFinite(n) ? n.toFixed(2) : "";
      })(),
      cbm: String(get("cbm") ?? "").trim(),
      L: normalizeNumber(String(get("lengte"))),
      W: normalizeNumber(String(get("breedte"))),
      H: normalizeNumber(String(get("hoogte"))),
      batch: String(get("batch") ?? "").trim(),
    };

    const missing = [];
    if (!vals.code) missing.push("Productcode");
    if (!vals.desc) missing.push("Omschrijving");
    if (!vals.ean) missing.push("EAN");
    if (!vals.qty || isNaN(+vals.qty)) missing.push("QTY");
    if (!vals.gw) missing.push("G.W");
    if (!vals.cbm) missing.push("CBM");
    if (!isFinite(vals.L) || vals.L <= 0) missing.push("Lengte");
    if (!isFinite(vals.W) || vals.W <= 0) missing.push("Breedte");
    if (!isFinite(vals.H) || vals.H <= 0) missing.push("Hoogte");
    if (!vals.batch) missing.push("Batch");

    if (missing.length) {
      return {
        ok: false,
        error: `Ontbrekende/ongeldige velden: ${missing.join(", ")}`,
      };
    }
    vals.L = +vals.L;
    vals.W = +vals.W;
    vals.H = +vals.H;
    return { ok: true, vals };
  }

  // Render één PDF via de zichtbare PREVIEW (parity met single)
  async function renderOnePdfBlobViaPreview(vals) {
    const oldOpacity = canvasEl.style.opacity;
    canvasEl.style.opacity = "0.15"; // demp UI tijdens batch-render (layout blijft zichtbaar)

    const { sizes } = await renderPreviewFor(vals);

    const jsPDF = await loadJsPDF();
    const h2c = await loadHtml2Canvas();

    // Na rotatie: breedte = s.h, hoogte = s.w (cm)
    const contentW = Math.max(...sizes.map((s) => s.h));
    const contentH = sizes.reduce((sum, s) => sum + s.w, 0);
    const pageW = contentW + PDF_MARGIN_CM * 2;
    const pageH = contentH + PDF_MARGIN_CM * 2;

    const A4W = 21.0,
      A4H = 29.7;
    const doc = new jsPDF({
      unit: "cm",
      orientation: "portrait",
      format: pageW <= A4W && pageH <= A4H ? "a4" : [pageW, pageH],
    });
    doc.setFont("helvetica", "normal");

    const orderIdx = [1, 3, 2, 4];
    for (let i = 0, y = PDF_MARGIN_CM; i < orderIdx.length; i++) {
      const img = await capturePreviewLabelToImage(h2c, orderIdx[i], i === 0);
      const s = sizes[orderIdx[i] - 1];
      const wRot = s.h,
        hRot = s.w;
      doc.addImage(img, "PNG", PDF_MARGIN_CM, y, wRot, hRot, undefined, "FAST");
      y += hRot;
    }

    canvasEl.style.opacity = oldOpacity || "";
    return doc.output("blob");
  }

  /* ====== BATCH UI EVENTS ====== */
  if (btnPickFile && fileInput)
    btnPickFile.addEventListener("click", () => fileInput.click());

  if (dropzone) {
    ["dragover", "dragenter"].forEach((ev) => {
      dropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      dropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
      });
    });
    dropzone.addEventListener("drop", async (e) => {
      const f = e.dataTransfer.files?.[0];
      if (f) await handleFile(f);
    });
  }

  if (fileInput) {
    fileInput.addEventListener("change", async () => {
      const f = fileInput.files?.[0];
      if (f) await handleFile(f);
    });
  }

  async function handleFile(file) {
    resetLog();
    if (logWrap) logWrap.classList.remove("hidden");
    try {
      log(`Bestand: ${file.name}`);
      const rows = await parseFile(file);
      if (!rows.length) {
        log("Geen rijen gevonden.", "error");
        return;
      }

      headers = Object.keys(rows[0]);
      mapping = guessMapping(headers);

      buildMappingUI(headers, mapping);
      setHidden(mappingWrap, false);

      showTablePreview(rows);
      setHidden(previewWrap, false);

      setHidden(normWrap, false);
      setHidden(batchControls, false);
      setHidden(progressWrap, true);

      parsedRows = rows;
      log(`Gelezen rijen: ${rows.length}`, "ok");
    } catch (err) {
      log(`Fout bij lezen bestand: ${err.message || err}`, "error");
    }
  }

  if (btnTemplateCsv) {
    btnTemplateCsv.addEventListener("click", () => {
      // Alleen headers — géén voorbeeldregels
      const hdrs = [
        "ERP",
        "Omschrijving",
        "EAN",
        "QTY",
        "G.W",
        "CBM",
        "Length (L)",
        "Width (W)",
        "Height (H)",
        "Batch",
      ];
      const blob = new Blob([hdrs.join(",") + "\n"], {
        type: "text/csv;charset=utf-8;",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "etiketten-template.csv";
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(url);
        a.remove();
      }, 0);
    });
  }

  if (btnTemplateXlsx) {
    btnTemplateXlsx.addEventListener("click", () => {
      // Alleen headers — géén voorbeeldregels
      const hdrs = [
        "ERP",
        "Omschrijving",
        "EAN",
        "QTY",
        "G.W",
        "CBM",
        "Length (L)",
        "Width (W)",
        "Height (H)",
        "Batch",
      ];
      const ws = XLSX.utils.aoa_to_sheet([hdrs]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Etiketten");
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "etiketten-template.xlsx";
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(url);
        a.remove();
      }, 0);
    });
  }

  if (btnRunBatch) {
    btnRunBatch.addEventListener("click", async () => {
      if (!parsedRows.length) {
        log("Geen dataset geladen.", "error");
        return;
      }

      const missingMap = REQUIRED_FIELDS.filter(([k]) => !mapping[k]).map(
        ([, label]) => label
      );
      if (missingMap.length) {
        log(`Koppel alle verplichte velden: ${missingMap.join(", ")}`, "error");
        return;
      }

      abortFlag = false;
      if (btnAbortBatch) btnAbortBatch.disabled = false;
      setHidden(progressWrap, false);
      progressBar.style.width = "0%";
      progressLabel.textContent = `${0} / ${parsedRows.length}`;
      progressPhase.textContent = "Voorbereiden…";

      const zip = new JSZip();
      const batchTime = ts();
      let okCount = 0,
        errCount = 0;

      for (let i = 0; i < parsedRows.length; i++) {
        if (abortFlag) {
          log(`Batch afgebroken op rij ${i + 1}.`, "error");
          break;
        }

        const r = readRowWithMapping(parsedRows[i], mapping);
        if (!r.ok) {
          errCount++;
          log(`Rij ${i + 1}: ${r.error}`, "error");
        } else {
          try {
            progressPhase.textContent = `Rij ${i + 1}: PDF renderen…`;
            const blob = await renderOnePdfBlobViaPreview(r.vals); // <<< parity pad
            const safeCode = r.vals.code.replace(/[^\w.-]+/g, "_");
            const name = `${safeCode} - ${batchTime} - R${String(
              i + 1
            ).padStart(3, "0")}.pdf`;
            zip.file(name, blob);
            okCount++;
          } catch (err) {
            errCount++;
            log(`Rij ${i + 1}: renderfout: ${err.message || err}`, "error");
          }
        }
        progressBar.style.width = `${Math.round(
          ((i + 1) / parsedRows.length) * 100
        )}%`;
        progressLabel.textContent = `${i + 1} / ${parsedRows.length}`;
        await new Promise((r) => setTimeout(r, 0)); // UI ademruimte
      }

      if (btnAbortBatch) btnAbortBatch.disabled = true;
      progressPhase.textContent = "Bundelen als ZIP…";

      if (okCount > 0) {
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `etiketten-batch - ${batchTime}.zip`;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          URL.revokeObjectURL(url);
          a.remove();
        }, 0);
        log(`Gereed: ${okCount} PDF’s succesvol, ${errCount} fouten.`, "ok");
      } else {
        log(`Geen PDF’s gegenereerd. (${errCount} fouten)`, "error");
      }
      progressPhase.textContent = "Klaar.";
    });
  }

  if (btnAbortBatch) {
    btnAbortBatch.addEventListener("click", () => {
      abortFlag = true;
      btnAbortBatch.disabled = true;
      progressPhase.textContent = "Afbreken…";
    });
  }

  /* ====== EVENTS & INIT ====== */
  const safeRender = () => renderSingle().catch((e) => alert(e.message || e));
  if (btnGen) btnGen.addEventListener("click", safeRender);
  if (btnPDF)
    btnPDF.addEventListener("click", async () => {
      try {
        await generatePDFSingle();
      } catch (e) {
        alert(e.message || e);
      }
    });
  window.addEventListener("resize", () => {
    renderSingle().catch(() => {});
  });

  // (optioneel) demo-waarden voor snelle start
  try {
    if (document.getElementById("prodCode")) {
      document.getElementById("prodCode").value = "LG1000843";
      document.getElementById("prodDesc").value =
        "Combination Lock - Orange - 1 Pack (YF20610B)";
      document.getElementById("ean").value = "8719632951889";
      document.getElementById("qty").value = "12";
      document.getElementById("gw").value = "18.00";
      document.getElementById("cbm").value = "0.02";
      document.getElementById("boxLength").value = "39";
      document.getElementById("boxWidth").value = "19.5";
      document.getElementById("boxHeight").value = "22";
      document.getElementById("batch").value = "IOR2500307";
      safeRender();
    }
  } catch (_) {}
})();
