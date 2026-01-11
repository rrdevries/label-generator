(() => {
  /*
    Refactor (bucket typography):
    - Bucket config loaded from ./labelBuckets.json
    - Preview / PDF / Batch flows unchanged
    - Typography is driven by bucket anchors; final safety net keeps "always render" behavior
  */

  /* ====== CONSTANTS ====== */
  const PX_PER_CM = 37.7952755906;
  const PX_PER_PT = 96 / 72; // 1pt = 1/72 inch; CSS px = 1/96 inch

  // Label padding (rondom, binnen het label) in cm
  const LABEL_PADDING_CM = 0.6;

  // Threshold: als we onder 10px content-tekst komen, dan pas zachte afbreking aanzetten
  const WRAP_THRESHOLD_PX = 10;
  const MIN_SCALE_K = 0.02; // absolute bodem voor fallback-scale

  // ===== PDF =====
  const PDF_MARGIN_CM = 0.5;
  const BORDER_PX = 1;
  let currentPreviewScale = 1;

  /* ====== Bucket config ====== */
  let BUCKET_CONFIG = null;
  let BUCKET_BY_KEY = new Map();

  async function loadBucketConfig(url = "./labelBuckets.json") {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) {
      throw new Error(
        `Bucket-config kon niet worden geladen (${res.status} ${res.statusText}). ` +
          `Tip: open dit via een (lokale) webserver i.p.v. file://.`
      );
    }
    const cfg = await res.json();
    if (!cfg || !Array.isArray(cfg.anchors)) {
      throw new Error("Bucket-config is ongeldig: anchors ontbreken.");
    }
    return cfg;
  }

  function indexBucketConfig(cfg) {
    BUCKET_BY_KEY = new Map();
    cfg.anchors.forEach((a) => {
      if (!a || !a.key) return;
      BUCKET_BY_KEY.set(String(a.key).toUpperCase(), a);
    });
  }

  function clamp01(x) {
    return Math.max(0, Math.min(1, x));
  }

  function ptToPx(pt) {
    return (Number(pt) || 0) * PX_PER_PT;
  }

  function selectBucketKeyFor(W_cm, H_cm) {
    // Inputs are label face dimensions (cm)
    const w = Number(W_cm);
    const h = Number(H_cm);
    if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
      return null;
    }

    const rules = BUCKET_CONFIG?.bucketRules;
    const eps = Number(rules?.familySelection?.squareTolerance?.eps) || 0.1;
    const r = w / h;
    const q = h / w;
    const D = Math.max(w, h);

    // family
    let family = "SQUARE";
    if (!(1 - eps <= r && r <= 1 + eps)) {
      family = r < 1 - eps ? "PORTRAIT" : "LANDSCAPE";
    }

    // variant
    let variant = null;
    if (family === "PORTRAIT") {
      if (r < 0.4) variant = "NARROW";
      else if (r < 0.65) variant = "STANDARD";
      else variant = "WIDE";
    } else if (family === "LANDSCAPE") {
      if (q < 0.4) variant = "SHORT";
      else if (q < 0.65) variant = "STANDARD";
      else variant = "HIGH";
    }

    // size class
    let sizeClass = null;
    if (D >= 5 && D < 10) sizeClass = "MICRO";
    else if (D >= 10 && D < 25) sizeClass = "SMALL";
    else if (D >= 25 && D < 40) sizeClass = "MEDIUM";
    else if (D >= 40 && D < 70) sizeClass = "LARGE";
    else if (D >= 70 && D <= 100) sizeClass = "EXTRA_LARGE";

    if (!sizeClass) return null;

    if (family === "SQUARE") return `${family}_${sizeClass}`;
    return `${family}_${sizeClass}_${variant}`;
  }

  function getBucketAnchorFor(W_cm, H_cm) {
    const key = selectBucketKeyFor(W_cm, H_cm);
    if (!key) return null;
    const lookup = (k) => BUCKET_BY_KEY.get(String(k).toUpperCase()) || null;

    // 1) exact
    let a = lookup(key);

    // 2) variant fallback (als bepaalde varianten niet bestaan / niet gewenst zijn)
    //    - PORTRAIT_*_NARROW -> PORTRAIT_*_STANDARD
    //    - LANDSCAPE_*_SHORT -> LANDSCAPE_*_STANDARD
    //    - PORTRAIT/LANDSCAPE -> als STANDARD ontbreekt: probeer WIDE/HIGH
    if (!a) {
      if (/_NARROW$/i.test(key))
        a = lookup(key.replace(/_NARROW$/i, "_STANDARD"));
      else if (/_SHORT$/i.test(key))
        a = lookup(key.replace(/_SHORT$/i, "_STANDARD"));
    }
    if (!a) {
      if (/^PORTRAIT_/i.test(key))
        a = lookup(key.replace(/_(NARROW|STANDARD|WIDE)$/i, "_WIDE"));
      else if (/^LANDSCAPE_/i.test(key))
        a = lookup(key.replace(/_(SHORT|STANDARD|HIGH)$/i, "_HIGH"));
    }

    // 3) ignore disabled
    if (a && a.enabled === false) return null;
    return a || null;
  }

  function applyBucketTypography(innerEl) {
    const W_cm = Number(innerEl.dataset.wcm);
    const H_cm = Number(innerEl.dataset.hcm);
    const anchor = getBucketAnchorFor(W_cm, H_cm);

    // Fallback: geen bucket -> zet niets (oude auto-fit kan dan nog werken)
    if (!anchor) return null;

    const D = Math.max(W_cm, H_cm);
    const Dref = Number(anchor.D_ref_cm) || D;
    const k = Dref > 0 ? D / Dref : 1;

    const pick = (name) => Number(anchor.anchors?.[name]?.pt) || 0;

    const erpPx = ptToPx(pick("erp_text") * k);
    const descPx = ptToPx(pick("product_description") * k);
    const labelPx = ptToPx(pick("content_label") * k);
    const textPx = ptToPx(pick("content_text") * k);
    const footerPx = ptToPx(pick("footer") * k);

    innerEl.style.setProperty("--fs-erp", `${erpPx}px`);
    innerEl.style.setProperty("--fs-desc", `${descPx}px`);
    innerEl.style.setProperty("--fs-label", `${labelPx}px`);
    innerEl.style.setProperty("--fs-text", `${textPx}px`);
    innerEl.style.setProperty("--fs-footer", `${footerPx}px`);

    innerEl.dataset.bucketKey = String(anchor.key || "");
    innerEl.dataset.bucketK = String(k);
    innerEl.dataset.family = String(anchor.requirements?.family || "");
    innerEl.dataset.variant = String(anchor.requirements?.variant || "");
    innerEl.dataset.sizeClass = String(anchor.requirements?.sizeClass || "");

    return {
      key: anchor.key,
      k,
      sizesPx: { erpPx, descPx, labelPx, textPx, footerPx },
    };
  }

  /* ====== Helpers ====== */
  const $ = (sel) => document.querySelector(sel);

  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === "class") node.className = v;
      else if (k === "style") Object.assign(node.style, v);
      else node.setAttribute(k, v);
    }
    children.flat().forEach((c) => {
      if (c == null) return;
      node.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    });
    return node;
  }

  function parseNumber(val) {
    if (val == null) return "";
    const s = String(val).trim();
    if (!s) return "";
    // accepteer komma of punt
    const n = Number(s.replace(",", "."));
    return Number.isFinite(n) ? n : "";
  }

  function format2(val) {
    const n = parseNumber(val);
    if (n === "") return "";
    return n.toFixed(2);
  }

  function getFormValues() {
    return {
      code: ($("#erp")?.value || "").trim(),
      desc: ($("#desc")?.value || "").trim(),
      ean: ($("#ean")?.value || "").trim(),
      qty: ($("#qty")?.value || "").trim(),
      gw: ($("#gw")?.value || "").trim(),
      cbm: ($("#cbm")?.value || "").trim(),
      len: parseNumber($("#len")?.value),
      wid: parseNumber($("#wid")?.value),
      hei: parseNumber($("#hei")?.value),
      batch: ($("#batch")?.value || "").trim(),
    };
  }

  function syncDescWidthToSpecs(innerEl) {
    const grid = innerEl.querySelector(".specs-grid");
    const desc = innerEl.querySelector(".label-desc");
    if (!grid || !desc) return;

    // offsetWidth is layout-breedte (niet beïnvloed door transform scale)
    const w = grid.offsetWidth || grid.getBoundingClientRect().width;
    desc.style.setProperty("--desc-w", w + "px");
  }

  function descFitsInTwoLines(descEl) {
    const cs = getComputedStyle(descEl);
    const lh = parseFloat(cs.lineHeight);
    if (!Number.isFinite(lh) || lh <= 0) return true; // fallback

    const maxH = lh * 2 + 0.5; // toleranties
    return descEl.scrollHeight <= maxH;
  }

  function shrinkDescToTwoLines(innerEl) {
    const desc = innerEl.querySelector(".label-desc");
    if (!desc) return;

    // Eerst breedte syncen, anders klopt wrap niet
    syncDescWidthToSpecs(innerEl);

    // Reset naar CSS var
    desc.style.fontSize = "";

    if (descFitsInTwoLines(desc)) return;

    const basePx = parseFloat(getComputedStyle(desc).fontSize) || 12;

    // Binary search: verklein alleen de omschrijving tot hij binnen 2 regels past
    let lo = 2;
    let hi = Math.max(2, basePx);
    let best = lo;

    for (let i = 0; i < 18; i++) {
      const mid = (lo + hi) / 2;
      desc.style.fontSize = mid + "px";
      if (descFitsInTwoLines(desc)) {
        best = mid;
        hi = mid;
      } else {
        lo = mid;
      }
    }

    desc.style.fontSize = best + "px";
  }

  /* ====== Label sizes (Optie A: 0.9) ====== */
  function calcLabelSizes(values) {
    const L = values.len || 0;
    const W = values.wid || 0;
    const H = values.hei || 0;

    const FACTOR = 0.9; // Optie A: 10% kleiner aan elke zijde

    // Etiket 1 & 2: lengte x hoogte
    const fbW = L * FACTOR;
    const fbH = H * FACTOR;

    // Etiket 3 & 4: breedte x hoogte
    const sideW = W * FACTOR;
    const sideH = H * FACTOR;

    return [
      { idx: 1, name: "Etiket 1 (front/back)", w: fbW, h: fbH, type: "fb" },
      { idx: 2, name: "Etiket 2 (front/back)", w: fbW, h: fbH, type: "fb" },
      { idx: 3, name: "Etiket 3 (side)", w: sideW, h: sideH, type: "side" },
      { idx: 4, name: "Etiket 4 (side)", w: sideW, h: sideH, type: "side" },
    ];
  }

  function renderDims(sizes) {
    const dims = $("#dims");
    if (!dims) return;
    dims.innerHTML = "";
    sizes.forEach((s) => {
      dims.append(
        el(
          "div",
          { class: "pill" },
          `${s.name}: ${format2(s.w)} × ${format2(s.h)} cm`
        )
      );
    });
  }

  /* ====== Fit / guard / fallback ====== */
  function fitsWithGuard(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;

    // Detecteer overflow in grid-cellen (EAN/waarden)
    const valOverflow = Array.from(
      content.querySelectorAll(".specs-grid .val")
    ).some((v) => v.scrollWidth > v.clientWidth + 0.5);
    if (valOverflow) return false;

    return (
      content.scrollWidth <= innerEl.clientWidth - guardX &&
      content.scrollHeight <= innerEl.clientHeight - guardY
    );
  }

  // Laatste redmiddel: schaal de volledige content zodat alles altijd zichtbaar blijft.
  function applyScaleFallback(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;

    const availW = Math.max(1, innerEl.clientWidth - guardX);
    const availH = Math.max(1, innerEl.clientHeight - guardY);

    const sw = Math.max(1, content.scrollWidth);
    const sh = Math.max(1, content.scrollHeight);

    let scaleW = availW / sw;
    let scaleH = availH / sh;

    // Extra: corrigeer voor overflow die alleen in grid-cellen zichtbaar is
    let scaleVal = 1;
    content.querySelectorAll(".specs-grid .val").forEach((v) => {
      const vSw = v.scrollWidth;
      const vCw = v.clientWidth;
      if (vSw > vCw + 0.5) {
        scaleVal = Math.min(scaleVal, vCw / vSw);
      }
    });

    const k = Math.max(MIN_SCALE_K, Math.min(1, scaleW, scaleH, scaleVal));
    content.style.setProperty("--k", String(k));
    return k;
  }

  function applyBucketThenFit(innerEl) {
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;

    const guardX = Math.max(2, w * 0.015);
    const guardY = Math.max(2, h * 0.015);

    // 1) apply bucket typography
    const info = applyBucketTypography(innerEl);

    // 2) wrap mode: start no-wrap; enable soft-wrap when text becomes small
    innerEl.classList.add("nowrap-mode");
    innerEl.classList.remove("softwrap-mode");

    const textPx = info?.sizesPx?.textPx;
    if (Number.isFinite(textPx) && textPx < WRAP_THRESHOLD_PX) {
      innerEl.classList.remove("nowrap-mode");
      innerEl.classList.add("softwrap-mode");
    }

    // Reset fallback-scale bij nieuwe metingen
    const content = innerEl.querySelector(".label-content");
    if (content) content.style.setProperty("--k", "1");

    // 3) ensure desc is max 2 lines
    syncDescWidthToSpecs(innerEl);
    shrinkDescToTwoLines(innerEl);

    // 4) final safety net: scale down whole content if needed
    if (!fitsWithGuard(innerEl, guardX, guardY)) {
      applyScaleFallback(innerEl, guardX, guardY);
    } else {
      if (content) content.style.setProperty("--k", "1");
    }
  }

  async function mountThenFit(container) {
    await new Promise((r) => requestAnimationFrame(r));
    await new Promise((r) => requestAnimationFrame(r));
    fitAllIn(container);
  }

  function fitAllIn(container) {
    container.querySelectorAll(".label-inner").forEach((inner) => {
      applyBucketThenFit(inner);
    });
  }

  /* ====== Build label DOM ====== */
  function buildLeftBlock(values, size, largestTwo) {
    const block = el("div", { class: "specs-grid" });

    block.append(
      el("div", { class: "key" }, "EAN:"),
      el("div", { class: "val" }, values.ean || ""),
      el("div", { class: "key" }, "QTY:"),
      el("div", { class: "val" }, `${values.qty || ""} PCS`),
      el("div", { class: "key" }, "G.W:"),
      el("div", { class: "val" }, `${values.gw || ""} KGS`),
      el("div", { class: "key" }, "CBM:"),
      el("div", { class: "val" }, values.cbm || "")
    );

    // Als largestTwo null is => fallback naar default (oude gedrag: fb => C/N, side => Made in China)
    const useMadeInChina = largestTwo
      ? largestTwo.has(size.idx)
      : size.type !== "fb";

    if (useMadeInChina) {
      block.append(
        el("div", { class: "key footer-key" }, ""),
        el("div", { class: "val footer-val" }, "Made in China")
      );
    } else {
      block.append(
        el("div", { class: "key footer-key" }, "C/N:"),
        el("div", { class: "val footer-val" }, "___________________")
      );
    }

    return block;
  }

  function createLabelEl(size, values, previewScale, largestTwo) {
    const widthPx = Math.round(size.w * PX_PER_CM * previewScale);
    const heightPx = Math.round(size.h * PX_PER_CM * previewScale);

    const wrap = el("div", { class: "label-wrap" });
    const label = el("div", {
      class: "label",
      style: { width: widthPx + "px", height: heightPx + "px" },
    });
    label.dataset.idx = String(size.idx);

    const inner = el("div", { class: "label-inner nowrap-mode" });
    inner.dataset.wcm = String(size.w);
    inner.dataset.hcm = String(size.h);

    const padPx = LABEL_PADDING_CM * PX_PER_CM * previewScale;
    label.style.padding = padPx + "px";

    const head = el(
      "div",
      { class: "label-head" },
      el("div", { class: "code-box line" }, values.code),
      el("div", { class: "line label-desc" }, values.desc)
    );

    const content = el("div", { class: "label-content" });
    content.append(
      head,
      el("div", { class: "block-spacer" }),
      buildLeftBlock(values, size, largestTwo)
    );
    inner.append(content);
    label.append(inner);
    wrap.append(label, el("div", { class: "label-num" }, `Etiket ${size.idx}`));

    return wrap;
  }

  /* ====== Preview render ====== */
  function computePreviewScale(sizes) {
    const labelsGrid = $("#labelsGrid");
    if (!labelsGrid) return 1;

    const containerWidth = labelsGrid.clientWidth || 900;
    const maxWcm = Math.max(...sizes.map((s) => s.w));
    const maxHcm = Math.max(...sizes.map((s) => s.h));

    const maxWpx = maxWcm * PX_PER_CM;
    const maxHpx = maxHcm * PX_PER_CM;

    const cellW = containerWidth / 2;
    const cellH = Math.max(280, window.innerHeight * 0.35);

    const sW = cellW / maxWpx;
    const sH = cellH / maxHpx;

    const s = Math.min(sW, sH) * 0.98;
    return Math.max(0.08, Math.min(1, s));
  }

  async function renderPreviewFor(values) {
    const labelsGrid = $("#labelsGrid");
    if (!labelsGrid) return;

    const sizes = calcLabelSizes(values);

    // Bepaal welke 2 etiketten “Made in China” krijgen op basis van grootste oppervlak.
    // Als er een tie op de grens is, gebruiken we default gedrag (size.type).
    function pickTwoLargestIdxOrNull(sizes) {
      const eps = 1e-9;
      const ranked = [...sizes]
        .map((s) => ({ idx: s.idx, area: (s.w || 0) * (s.h || 0) }))
        .sort((a, b) => b.area - a.area);

      if (Math.abs(ranked[1].area - ranked[2].area) <= eps) return null;
      return new Set([ranked[0].idx, ranked[1].idx]);
    }

    const largestTwo = pickTwoLargestIdxOrNull(sizes);

    renderDims(sizes);

    const scale = computePreviewScale(sizes);
    currentPreviewScale = scale;

    labelsGrid.innerHTML = "";
    const fragments = document.createDocumentFragment();
    sizes.forEach((size) => {
      fragments.append(createLabelEl(size, values, scale, largestTwo));
    });
    labelsGrid.append(fragments);

    await mountThenFit(labelsGrid);
    return { sizes, scale };
  }

  async function renderSingle() {
    const vals = getFormValues();
    await renderPreviewFor(vals);
  }

  /* ====== jsPDF / html2canvas ====== */
  function loadJsPDF() {
    return window.jspdf?.jsPDF;
  }

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function buildTimestamp(d = new Date()) {
    return (
      `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ` +
      `${pad2(d.getHours())}.${pad2(d.getMinutes())}.${pad2(d.getSeconds())}`
    );
  }

  function buildPdfFileName(code) {
    const safeCode = (code || "export").trim() || "export";
    return `${safeCode} - ${buildTimestamp()}.pdf`;
  }

  function rotateCanvas90CW(srcCanvas) {
    const dst = document.createElement("canvas");
    dst.width = srcCanvas.height;
    dst.height = srcCanvas.width;

    const ctx = dst.getContext("2d");
    ctx.translate(dst.width, 0);
    ctx.rotate(Math.PI / 2);
    ctx.drawImage(srcCanvas, 0, 0);

    return dst;
  }

  async function captureLabelToRotatedPng(labelIdx) {
    const src = document.querySelector(`.label[data-idx="${labelIdx}"]`);
    if (!src)
      throw new Error(`Label ${labelIdx} niet gevonden voor PDF-capture.`);

    const clone = src.cloneNode(true);

    clone.style.borderTop = `${BORDER_PX}px solid #000`;
    clone.style.borderRight = `${BORDER_PX}px solid #000`;
    clone.style.borderBottom = `${BORDER_PX}px solid #000`;
    clone.style.borderLeft = `${BORDER_PX}px solid #000`;

    const host = document.createElement("div");
    host.style.position = "fixed";
    host.style.left = "-10000px";
    host.style.top = "0";
    host.style.background = "#fff";
    host.style.padding = "0";
    host.style.margin = "0";

    document.body.appendChild(host);
    host.appendChild(clone);

    const capScale = Math.max(
      2,
      window.devicePixelRatio || 1,
      1 / (currentPreviewScale || 1)
    );

    const canvas = await html2canvas(clone, {
      scale: capScale,
      backgroundColor: "#ffffff",
      useCORS: true,
    });

    const rot = rotateCanvas90CW(canvas);

    document.body.removeChild(host);

    return rot.toDataURL("image/png");
  }

  async function generatePDFSingle() {
    const JsPDF = loadJsPDF();
    if (!JsPDF) throw new Error("jsPDF niet geladen");
    if (!window.html2canvas) throw new Error("html2canvas niet geladen");

    const vals = getFormValues();
    const result = await renderPreviewFor(vals);
    if (!result) throw new Error("Kon preview niet renderen voor PDF.");
    const { sizes } = result;

    const order = [1, 3, 2, 4];

    const pageW = Math.max(...sizes.map((s) => s.h)) + PDF_MARGIN_CM * 2;
    const pageH = sizes.reduce((sum, s) => sum + s.w, 0) + PDF_MARGIN_CM * 2;

    const pdf = new JsPDF({
      unit: "cm",
      orientation: "portrait",
      format: [pageW, pageH],
    });

    let y = PDF_MARGIN_CM;

    for (const idx of order) {
      const s = sizes[idx - 1];
      const imgData = await captureLabelToRotatedPng(idx);

      const wRot = s.h;
      const hRot = s.w;

      pdf.addImage(
        imgData,
        "PNG",
        PDF_MARGIN_CM,
        y,
        wRot,
        hRot,
        undefined,
        "FAST"
      );
      y += hRot;
    }

    pdf.save(buildPdfFileName(vals.code));
  }

  /* ====== BATCH (Excel / CSV) ====== */
  // Batch state
  let parsedRows = [];
  let headers = [];
  let mapping = {};
  let abortFlag = false;

  // Batch DOM refs
  let dropzone,
    fileInput,
    btnPickFile,
    btnTemplateCsv,
    btnTemplateXlsx,
    mappingWrap,
    mappingGrid,
    previewWrap,
    tablePreview,
    normWrap,
    chkComma,
    chkTrim,
    batchControls,
    btnRunBatch,
    btnAbortBatch,
    progressWrap,
    progressBar,
    progressLabel,
    progressPhase,
    logWrap,
    logList;

  function setHidden(elm, hidden) {
    if (elm) elm.classList.toggle("hidden", !!hidden);
  }

  function resetLog() {
    if (logList) logList.innerHTML = "";
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

  async function parseFile(file) {
    if (!window.XLSX) throw new Error("XLSX library niet geladen.");
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
    REQUIRED_FIELDS.forEach(([key]) => (m[key] = ""));
    const slugs = hdrs.map((h) => slugify(String(h).replace(/\([^)]*\)/g, "")));

    // exact synonym match
    for (let i = 0; i < hdrs.length; i++) {
      const h = hdrs[i];
      const s = slugs[i];
      for (const [key, syns] of Object.entries(SYNONYMS)) {
        if (syns.includes(s)) {
          m[key] = h;
          break;
        }
      }
    }

    // substring fallback
    for (const [key] of REQUIRED_FIELDS) {
      if (!m[key]) {
        const sset = (SYNONYMS[key] || [key]).filter((tok) => tok.length >= 2);
        const idx = slugs.findIndex((s) => sset.some((tok) => s.includes(tok)));
        if (idx >= 0) m[key] = hdrs[idx];
      }
    }

    return m;
  }

  function buildMappingUI(hdrs, mappingObj) {
    if (!mappingGrid) return;
    mappingGrid.innerHTML = "";

    const makeRow = (key, labelText) => {
      const row = el("div", { class: "map-row" });
      const lab = el("label", {}, labelText + " *");
      const sel = el("select", { "data-key": key });
      sel.appendChild(el("option", { value: "" }, "-- kies kolom --"));
      hdrs.forEach((h) => {
        const opt = el("option", { value: h }, h);
        if (mappingObj[key] === h) opt.selected = true;
        sel.appendChild(opt);
      });
      row.append(lab, sel);
      mappingGrid.appendChild(row);
    };

    REQUIRED_FIELDS.forEach(([k, l]) => makeRow(k, l));

    mappingGrid.querySelectorAll("select").forEach((sel) => {
      sel.addEventListener("change", () => {
        mappingObj[sel.getAttribute("data-key")] = sel.value;
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
    // 1.234.567,89 -> 1234567.89
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(val)) val = val.replace(/\./g, "");
    const num = parseFloat(val);
    return isFinite(num) ? num : NaN;
  }

  function readRowWithMapping(row, mappingObj) {
    const get = (key) => {
      const hdr = mappingObj[key] || "";
      return hdr ? row[hdr] ?? "" : "";
    };

    return {
      code: String(get("productcode") ?? "").trim(),
      desc: String(get("omschrijving") ?? "").trim(),
      ean: String(get("ean") ?? "").trim(),
      qty: String(get("qty") ?? "").trim(),
      gw: String(get("gw") ?? "").trim(),
      cbm: String(get("cbm") ?? "").trim(),
      len: normalizeNumber(get("lengte")),
      wid: normalizeNumber(get("breedte")),
      hei: normalizeNumber(get("hoogte")),
      batch: String(get("batch") ?? "").trim(),
    };
  }

  function validateMapping(mappingObj) {
    const missing = REQUIRED_FIELDS.filter(([k]) => !mappingObj[k]);
    return missing.map(([, label]) => label);
  }

  function buildTemplateRows() {
    return [
      {
        ERP: "LG1000843",
        Omschrijving: "Combination Lock - Orange - 1 Pack (YF20610B)",
        EAN: "8719632951889",
        QTY: "12",
        "G.W": "18,00",
        CBM: "0.02",
        "Length (L)": "39",
        "Width (W)": "19,5",
        "Height (H)": "22",
        Batch: "IOR2500307",
      },
    ];
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 0);
  }

  function downloadTemplateCsv() {
    const rows = buildTemplateRows();
    const cols = Object.keys(rows[0]);
    const lines = [cols.join(";")].concat(
      rows.map((r) =>
        cols.map((c) => String(r[c] ?? "").replace(/;/g, ",")).join(";")
      )
    );
    const blob = new Blob([lines.join("\n")], {
      type: "text/csv;charset=utf-8",
    });
    downloadBlob(blob, "etiketten-template.csv");
  }

  function downloadTemplateXlsx() {
    if (!window.XLSX) {
      alert("XLSX library niet geladen.");
      return;
    }
    const rows = buildTemplateRows();
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    downloadBlob(blob, "etiketten-template.xlsx");
  }

  function initBatchUI() {
    dropzone = $("#dropzone");
    fileInput = $("#fileInput");
    btnPickFile = $("#btnPickFile");
    btnTemplateCsv = $("#btnTemplateCsv");
    btnTemplateXlsx = $("#btnTemplateXlsx");
    mappingWrap = $("#mappingWrap");
    mappingGrid = $("#mappingGrid");
    previewWrap = $("#previewWrap");
    tablePreview = $("#tablePreview");
    normWrap = $("#normWrap");
    chkComma = $("#optCommaDecimal");
    chkTrim = $("#optTrimSpaces");
    batchControls = $("#batchControls");
    btnRunBatch = $("#btnRunBatch");
    btnAbortBatch = $("#btnAbortBatch");
    progressWrap = $("#progressWrap");
    progressBar = $("#progressBar");
    progressLabel = $("#progressLabel");
    progressPhase = $("#progressPhase");
    logWrap = $("#logWrap");
    logList = $("#logList");

    const pickFile = () => fileInput?.click();

    btnPickFile?.addEventListener("click", pickFile);
    dropzone?.addEventListener("click", pickFile);

    btnTemplateCsv?.addEventListener("click", downloadTemplateCsv);
    btnTemplateXlsx?.addEventListener("click", downloadTemplateXlsx);

    dropzone?.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
    dropzone?.addEventListener("dragleave", () => {
      dropzone.classList.remove("dragover");
    });
    dropzone?.addEventListener("drop", (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
      const f = e.dataTransfer?.files?.[0];
      if (f) handleFileSelected(f).catch((err) => alert(err.message || err));
    });

    fileInput?.addEventListener("change", () => {
      const f = fileInput.files?.[0];
      if (f) handleFileSelected(f).catch((err) => alert(err.message || err));
    });

    async function handleFileSelected(file) {
      resetLog();
      log(`Bestand laden: ${file.name}`);

      parsedRows = await parseFile(file);
      if (!parsedRows.length) {
        setHidden(mappingWrap, true);
        setHidden(previewWrap, true);
        setHidden(normWrap, true);
        setHidden(batchControls, true);
        setHidden(progressWrap, true);
        setHidden(logWrap, false);
        log("Geen rijen gevonden.", "error");
        return;
      }

      headers = Object.keys(parsedRows[0]);
      mapping = guessMapping(headers);
      buildMappingUI(headers, mapping);
      showTablePreview(parsedRows);

      setHidden(mappingWrap, false);
      setHidden(previewWrap, false);
      setHidden(normWrap, false);
      setHidden(batchControls, false);
      setHidden(progressWrap, true);
      setHidden(logWrap, false);
      log(`Rijen geladen: ${parsedRows.length}`, "ok");
    }

    btnRunBatch?.addEventListener("click", async () => {
      try {
        if (!window.JSZip) throw new Error("JSZip library niet geladen.");
        if (!window.html2canvas) throw new Error("html2canvas niet geladen.");
        if (!loadJsPDF()) throw new Error("jsPDF niet geladen.");

        const missing = validateMapping(mapping);
        if (missing.length) {
          alert("Ontbrekende mapping: " + missing.join(", "));
          return;
        }

        abortFlag = false;
        btnAbortBatch.disabled = false;
        setHidden(progressWrap, false);
        setHidden(logWrap, false);

        if (progressBar) progressBar.style.width = "0%";
        if (progressLabel)
          progressLabel.textContent = `0 / ${parsedRows.length}`;
        if (progressPhase) progressPhase.textContent = "Renderen…";

        const zip = new JSZip();
        const batchTime = buildTimestamp();

        let okCount = 0;
        let errCount = 0;

        for (let i = 0; i < parsedRows.length; i++) {
          if (abortFlag) break;
          const row = parsedRows[i];
          try {
            const vals = readRowWithMapping(row, mapping);
            // Render en capture met dezelfde pipeline als single
            const result = await renderPreviewFor(vals);
            if (!result) throw new Error("Kon preview niet renderen.");

            // Maak PDF als blob
            const { sizes } = result;
            const order = [1, 3, 2, 4];

            const pageW =
              Math.max(...sizes.map((s) => s.h)) + PDF_MARGIN_CM * 2;
            const pageH =
              sizes.reduce((sum, s) => sum + s.w, 0) + PDF_MARGIN_CM * 2;

            const JsPDF = loadJsPDF();
            const pdf = new JsPDF({
              unit: "cm",
              orientation: "portrait",
              format: [pageW, pageH],
            });

            let y = PDF_MARGIN_CM;
            for (const idx of order) {
              const s = sizes[idx - 1];
              const imgData = await captureLabelToRotatedPng(idx);
              pdf.addImage(
                imgData,
                "PNG",
                PDF_MARGIN_CM,
                y,
                s.h,
                s.w,
                undefined,
                "FAST"
              );
              y += s.w;
            }

            const blob = pdf.output("blob");
            const safeCode = (vals.code || "export").trim() || "export";
            const name = `${safeCode} - ${batchTime} - R${String(
              i + 1
            ).padStart(3, "0")}.pdf`;
            zip.file(name, blob);
            okCount++;
          } catch (err) {
            errCount++;
            log(`Rij ${i + 1}: renderfout: ${err.message || err}`, "error");
          }

          if (progressBar)
            progressBar.style.width = `${Math.round(
              ((i + 1) / parsedRows.length) * 100
            )}%`;
          if (progressLabel)
            progressLabel.textContent = `${i + 1} / ${parsedRows.length}`;
          await new Promise((r) => setTimeout(r, 0));
        }

        if (btnAbortBatch) btnAbortBatch.disabled = true;
        if (progressPhase) progressPhase.textContent = "Bundelen als ZIP…";

        if (abortFlag) {
          log("Batch afgebroken.", "error");
        }

        if (okCount > 0) {
          const zipBlob = await zip.generateAsync({ type: "blob" });
          downloadBlob(zipBlob, `etiketten-batch - ${batchTime}.zip`);
          log(`Gereed: ${okCount} PDF’s succesvol, ${errCount} fouten.`, "ok");
        } else {
          log(`Geen PDF’s gegenereerd. (${errCount} fouten)`, "error");
        }

        if (progressPhase) progressPhase.textContent = "Klaar.";
      } catch (e) {
        alert(e.message || e);
      }
    });

    btnAbortBatch?.addEventListener("click", () => {
      abortFlag = true;
      btnAbortBatch.disabled = true;
      if (progressPhase) progressPhase.textContent = "Afbreken…";
    });
  }

  /* ====== init ====== */
  async function init() {
    try {
      BUCKET_CONFIG = await loadBucketConfig("./labelBuckets.json");
      indexBucketConfig(BUCKET_CONFIG);
    } catch (e) {
      console.error(e);
      alert(e.message || e);
      // Zonder config kan de rest nog draaien, maar bucket-typografie zal ontbreken.
    }

    initBatchUI();

    const btnGen = $("#btnGen");
    const btnPDF = $("#btnPDF");

    const safeRender = () =>
      renderSingle().catch((err) => alert(err.message || err));
    btnGen?.addEventListener("click", safeRender);

    btnPDF?.addEventListener("click", async () => {
      try {
        await generatePDFSingle();
      } catch (err) {
        alert(err.message || err);
      }
    });

    window.addEventListener("resize", () => {
      renderSingle().catch(() => {});
    });

    renderSingle().catch(() => {});
  }

  document.addEventListener("DOMContentLoaded", () => {
    init().catch((e) => alert(e.message || e));
  });
})();
