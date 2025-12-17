(() => {
  /* ====== CONSTANTS ====== */
  const PX_PER_CM = 37.7952755906;

  // Label padding (rondom, binnen het label) in cm
  const LABEL_PADDING_CM = 0.6;

  // Threshold: als we onder 10px body-font moeten, dan pas zachte afbreking aanzetten
  const WRAP_THRESHOLD_PX = 10;
  const MIN_FS_PX = 2; // absolute bodem (moet altijd alles kunnen tonen)

  // Code-box is 1.6× de body-font
  const CODE_MULT = 1.6;

  // ===== PDF (terug naar v0.31 gedrag) =====
  const PDF_MARGIN_CM = 0.5;
  const BORDER_PX = 1;
  let currentPreviewScale = 1;

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

  /* ====== Auto-fit logic (font sizing) ====== */

  function fitsWithGuard(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;

    // Detecteer overflow in grid-cellen (EAN/waarden). Dit kan gebeuren zonder dat
    // content.scrollWidth groter wordt (grid-cel overflow).
    const valOverflow = Array.from(
      content.querySelectorAll(".specs-grid .val")
    ).some((v) => v.scrollWidth > v.clientWidth + 0.5);

    if (valOverflow) return false;

    return (
      content.scrollWidth <= innerEl.clientWidth - guardX &&
      content.scrollHeight <= innerEl.clientHeight - guardY
    );
  }

  function applyFontSizes(innerEl, fsPx) {
    innerEl.style.setProperty("--fs", fsPx + "px");

    // Reset eventuele fallback-scale bij nieuwe metingen
    const content = innerEl.querySelector(".label-content");
    if (content) content.style.setProperty("--k", "1");

    const codeEl = innerEl.querySelector(".code-box");
    if (codeEl) {
      codeEl.style.fontSize = fsPx * CODE_MULT + "px";
    }
  }

  function searchFontSize(innerEl, minFs, startHi, guardX, guardY) {
    // 1) agressief omhoog groeien vanaf startHi (groeifactor 1.08)
    let hi = Math.max(minFs, startHi);
    applyFontSizes(innerEl, hi);

    while (fitsWithGuard(innerEl, guardX, guardY) && hi < 500) {
      const next = hi * 1.08;
      applyFontSizes(innerEl, next);
      if (!fitsWithGuard(innerEl, guardX, guardY)) {
        applyFontSizes(innerEl, hi);
        break;
      }
      hi = next;
    }

    // 2) binaire search naar max die past
    let lo = minFs;
    let best = lo;

    // Zorg dat hi in elk geval "te groot" is (of gelijk)
    applyFontSizes(innerEl, hi);
    if (fitsWithGuard(innerEl, guardX, guardY)) {
      best = hi;
      return best;
    }

    // lo laten passen (minFs moet altijd passen, zo niet dan toch de bodem)
    applyFontSizes(innerEl, lo);
    if (!fitsWithGuard(innerEl, guardX, guardY)) {
      return minFs;
    }
    best = lo;

    for (let i = 0; i < 22; i++) {
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

  // Laatste redmiddel: als zelfs MIN_FS_PX + (soft)wrap niet past, schaal de volledige content
  // met CSS transform zodat alles altijd zichtbaar blijft.
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

    const k = Math.max(0.02, Math.min(1, scaleW, scaleH, scaleVal));
    content.style.setProperty("--k", String(k));
    return k;
  }

  function fitContentToBoxConditional(innerEl) {
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;

    // Guard: kleiner minimum zodat mini-labels nog bruikbare ruimte hebben,
    // maar wel voldoende marge tegen rand-clipping.
    const guardX = Math.max(2, w * 0.015);
    const guardY = Math.max(2, h * 0.015);

    const baseFromBox = Math.min(w, h) * 0.11;
    const startHi = Math.max(16, baseFromBox);

    // Fase 1: no-wrap (voorkeur)
    innerEl.classList.add("nowrap-mode");
    innerEl.classList.remove("softwrap-mode");

    let best = searchFontSize(innerEl, MIN_FS_PX, startHi, guardX, guardY);

    // Fase 2: als erg klein, probeer soft-wrap (kan horizontale overflow oplossen)
    if (best < WRAP_THRESHOLD_PX) {
      innerEl.classList.remove("nowrap-mode");
      innerEl.classList.add("softwrap-mode");

      best = searchFontSize(innerEl, MIN_FS_PX, startHi, guardX, guardY);
    }

    // Fase 3 (altijd alles tonen): als het nog steeds niet past op MIN_FS_PX,
    // schaal dan de volledige content met transform.
    if (!fitsWithGuard(innerEl, guardX, guardY)) {
      applyScaleFallback(innerEl, guardX, guardY);
    } else {
      const content = innerEl.querySelector(".label-content");
      if (content) content.style.setProperty("--k", "1");
    }

    return best;
  }

  async function mountThenFit(container) {
    await new Promise((r) => requestAnimationFrame(r));
    await new Promise((r) => requestAnimationFrame(r));
    fitAllIn(container);
  }

  function fitAllIn(container) {
    container.querySelectorAll(".label-inner").forEach((inner) => {
      inner.classList.add("nowrap-mode");
      inner.classList.remove("softwrap-mode");
      fitContentToBoxConditional(inner);
    });
  }

  /* ====== Build label DOM ====== */
  function buildLeftBlock(values, size) {
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

    if (size.type === "fb") {
      block.append(
        el("div", { class: "key" }, "C/N:"),
        el("div", { class: "val" }, "___________________")
      );
    } else {
      block.append(
        el("div", { class: "key" }, ""),
        el("div", { class: "val" }, "Made in China")
      );
    }
    return block;
  }

  function createLabelEl(size, values, previewScale) {
    const widthPx = Math.round(size.w * PX_PER_CM * previewScale);
    const heightPx = Math.round(size.h * PX_PER_CM * previewScale);

    const wrap = el("div", { class: "label-wrap" });
    const label = el("div", {
      class: "label",
      style: { width: widthPx + "px", height: heightPx + "px" },
    });
    label.dataset.idx = String(size.idx);

    const inner = el("div", { class: "label-inner nowrap-mode" });

    const padPx = LABEL_PADDING_CM * PX_PER_CM * previewScale;
    label.style.padding = padPx + "px";

    const head = el(
      "div",
      { class: "label-head" },
      el("div", { class: "code-box line" }, values.code),
      el("div", { class: "line" }, values.desc)
    );

    const content = el("div", { class: "label-content" });
    content.append(head, el("div", { class: "block-spacer" }), buildLeftBlock(values, size));

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
    renderDims(sizes);

    const scale = computePreviewScale(sizes);
    currentPreviewScale = scale;

    labelsGrid.innerHTML = "";
    const fragments = document.createDocumentFragment();

    sizes.forEach((size) => {
      fragments.append(createLabelEl(size, values, scale));
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
    if (!src) throw new Error(`Label ${labelIdx} niet gevonden voor PDF-capture.`);

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

    const capScale = Math.max(2, window.devicePixelRatio || 1, 1 / (currentPreviewScale || 1));

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

      pdf.addImage(imgData, "PNG", PDF_MARGIN_CM, y, wRot, hRot, undefined, "FAST");
      y += hRot;
    }

    pdf.save(buildPdfFileName(vals.code));
  }


  /* ====== BATCH (Excel / CSV) ======
     Hersteld vanuit v0.70, passend gemaakt op v0.78:
     - Leest XLSX/XLS/CSV (XLSX lib)
     - Kolommen mappen (auto + handmatig)
     - Genereert per rij dezelfde PDF als "PDF genereren" (v0.78 PDF pipeline)
     - Bundelt alles in één ZIP (JSZip)
  */

  // Batch state
  let parsedRows = [];
  let headers = [];
  let mapping = {};
  let abortFlag = false;

  // Batch DOM refs (worden in initBatchUI gezet)
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
      cols.forEach((c) => tr.appendChild(el("td", {}, String(rows[i][c] ?? ""))));
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
      len: normalizeNumber(String(get("lengte"))),
      wid: normalizeNumber(String(get("breedte"))),
      hei: normalizeNumber(String(get("hoogte"))),
      batch: String(get("batch") ?? "").trim(),
    };

    const missing = [];
    if (!vals.code) missing.push("ERP");
    if (!vals.desc) missing.push("Omschrijving");
    if (!vals.ean) missing.push("EAN");
    if (!vals.qty || isNaN(+vals.qty)) missing.push("QTY");
    if (!vals.gw) missing.push("G.W");
    if (!vals.cbm) missing.push("CBM");
    if (!isFinite(vals.len) || vals.len <= 0) missing.push("Length (L)");
    if (!isFinite(vals.wid) || vals.wid <= 0) missing.push("Width (W)");
    if (!isFinite(vals.hei) || vals.hei <= 0) missing.push("Height (H)");
    if (!vals.batch) missing.push("Batch");

    if (missing.length) {
      return { ok: false, error: `Ontbrekende/ongeldige velden: ${missing.join(", ")}` };
    }

    vals.len = +vals.len;
    vals.wid = +vals.wid;
    vals.hei = +vals.hei;
    return { ok: true, vals };
  }

  // Render één PDF als Blob via dezelfde pipeline als "PDF genereren"
  async function renderOnePdfBlobViaPreview(vals) {
    const JsPDF = loadJsPDF();
    if (!JsPDF) throw new Error("jsPDF niet geladen");
    if (!window.html2canvas) throw new Error("html2canvas niet geladen");

    const result = await renderPreviewFor(vals);
    if (!result) throw new Error("Kon preview niet renderen voor batch.");
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
      pdf.addImage(imgData, "PNG", PDF_MARGIN_CM, y, wRot, hRot, undefined, "FAST");
      y += hRot;
    }

    return pdf.output("blob");
  }

  function initBatchUI() {
    // DOM refs
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

    // Als batch-sectie niet aanwezig is, niets doen.
    if (!dropzone || !fileInput || !btnRunBatch) return;

    const handleFile = async (file) => {
      resetLog();
      setHidden(logWrap, false);

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
    };

    // pick file button
    btnPickFile?.addEventListener("click", () => fileInput.click());

    // drag & drop
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

    // input change
    fileInput.addEventListener("change", async () => {
      const f = fileInput.files?.[0];
      if (f) await handleFile(f);
    });

    // templates
    btnTemplateCsv?.addEventListener("click", () => {
      const hdrs = ["ERP","Omschrijving","EAN","QTY","G.W","CBM","Length (L)","Width (W)","Height (H)","Batch"];
      const blob = new Blob([hdrs.join(",") + "\n"], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "etiketten-template.csv";
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 0);
    });

    btnTemplateXlsx?.addEventListener("click", () => {
      if (!window.XLSX) {
        alert("XLSX library niet geladen.");
        return;
      }
      const hdrs = ["ERP","Omschrijving","EAN","QTY","G.W","CBM","Length (L)","Width (W)","Height (H)","Batch"];
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
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 0);
    });

    // run batch
    btnRunBatch.addEventListener("click", async () => {
      if (!parsedRows.length) {
        log("Geen dataset geladen.", "error");
        return;
      }
      if (!window.JSZip) {
        log("JSZip library niet geladen.", "error");
        return;
      }

      const missingMap = REQUIRED_FIELDS.filter(([k]) => !mapping[k]).map(([, label]) => label);
      if (missingMap.length) {
        log(`Koppel alle verplichte velden: ${missingMap.join(", ")}`, "error");
        return;
      }

      abortFlag = false;
      if (btnAbortBatch) btnAbortBatch.disabled = false;

      setHidden(progressWrap, false);
      if (progressBar) progressBar.style.width = "0%";
      if (progressLabel) progressLabel.textContent = `${0} / ${parsedRows.length}`;
      if (progressPhase) progressPhase.textContent = "Voorbereiden…";

      const zip = new JSZip();
      const batchTime = buildTimestamp();
      let okCount = 0;
      let errCount = 0;

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
            if (progressPhase) progressPhase.textContent = `Rij ${i + 1}: PDF renderen…`;
            const blob = await renderOnePdfBlobViaPreview(r.vals);

            const safeCode = String(r.vals.code || "export").replace(/[^\w.-]+/g, "_");
            const name = `${safeCode} - ${batchTime} - R${String(i + 1).padStart(3, "0")}.pdf`;
            zip.file(name, blob);
            okCount++;
          } catch (err) {
            errCount++;
            log(`Rij ${i + 1}: renderfout: ${err.message || err}`, "error");
          }
        }

        if (progressBar) progressBar.style.width = `${Math.round(((i + 1) / parsedRows.length) * 100)}%`;
        if (progressLabel) progressLabel.textContent = `${i + 1} / ${parsedRows.length}`;
        await new Promise((r) => setTimeout(r, 0));
      }

      if (btnAbortBatch) btnAbortBatch.disabled = true;
      if (progressPhase) progressPhase.textContent = "Bundelen als ZIP…";

      if (okCount > 0) {
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `etiketten-batch - ${batchTime}.zip`;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 0);
        log(`Gereed: ${okCount} PDF’s succesvol, ${errCount} fouten.`, "ok");
      } else {
        log(`Geen PDF’s gegenereerd. (${errCount} fouten)`, "error");
      }

      if (progressPhase) progressPhase.textContent = "Klaar.";
    });

    // abort
    btnAbortBatch?.addEventListener("click", () => {
      abortFlag = true;
      btnAbortBatch.disabled = true;
      if (progressPhase) progressPhase.textContent = "Afbreken…";
    });
  }

  /* ====== init ====== */
  function init() {
    initBatchUI();

    const btnGen = $("#btnGen");
    const btnPDF = $("#btnPDF");

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

    renderSingle().catch(() => {});
  }

  document.addEventListener("DOMContentLoaded", init);
})();