(() => {
  /* ====== CONSTANTS ====== */
  const PX_PER_CM = 37.7952755906;

  // Label padding (rondom, binnen het label) in cm
  const LABEL_PADDING_CM = 0.6;

  // Threshold: als we onder 10px body-font moeten, dan pas zachte afbreking aanzetten
  const WRAP_THRESHOLD_PX = 10;
  const MIN_FS_PX = 6; // absolute bodem

  // Code-box is 1.6× de body-font
  const CODE_MULT = 1.6;

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

  /* ====== Label sizes ======
     Uitgangspunt v0.31: labels berekend op basis van doosafmetingen:
       - front/back: (L + W) x H
       - side: (W + H) x H  (of varianten; hier: W+H x H)
     In je screenshot: 99x99x99 → 89.10 x 89.10 cm (dat komt door (L+W)*0.45 etc).
     We laten je bestaande logica intact en gebruiken wat er nu al in je v0.70 staat.
  */

  function calcLabelSizes(values) {
    // Huidige v0.70 logica (zoals in jouw code)
    const L = values.len || 0;
    const W = values.wid || 0;
    const H = values.hei || 0;

    // Factoren uit je huidige file
    const FRONT_BACK_FACTOR = 0.45;
    const SIDE_FACTOR = 0.45;

    const fbW = (L + W) * FRONT_BACK_FACTOR;
    const fbH = H * FRONT_BACK_FACTOR;

    const sideW = (W + H) * SIDE_FACTOR;
    const sideH = H * SIDE_FACTOR;

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

  /* ====== Auto-fit logic (font sizing) ======
=========================================================================
  Doel:
   - Tekst zo groot mogelijk maken binnen de beschikbare box.
   - Eerst no-wrap proberen; als nodig onder 10px -> soft-wrap.
   - Guard tegen clipping.
   - Code-box schaalt mee (1.6× body).
========================================================================= */

  /** past alles binnen 'innerEl' met een veiligheidsmarge? */
  function fitsWithGuard(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;
    return (
      content.scrollWidth <= innerEl.clientWidth - guardX &&
      content.scrollHeight <= innerEl.clientHeight - guardY
    );
  }

  /** zet actuele body-font (via --fs) + code-box (1.6× body) */
  function applyFontSizes(innerEl, fsPx) {
    innerEl.style.setProperty("--fs", fsPx + "px");
    const codeEl = innerEl.querySelector(".code-box");
    if (codeEl) {
      codeEl.style.fontSize = fsPx * CODE_MULT + "px";
    }
  }

  /** zoek een passende fontgrootte (eerst groeien, daarna finetunen naar beneden) */
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
      // hi past nog, dan is hi best (maar we laten de loop hierboven al stoppen)
      best = hi;
      return best;
    }

    // lo laten passen (minFs moet altijd passen, zo niet dan toch de bodem)
    applyFontSizes(innerEl, lo);
    if (!fitsWithGuard(innerEl, guardX, guardY)) {
      return minFs;
    }
    best = lo;

    // Binary refine
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

  function fitContentToBoxConditional(innerEl) {
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;

    const guardX = Math.max(8, w * 0.02);
    const guardY = Math.max(8, h * 0.02);

    const baseFromBox = Math.min(w, h) * 0.11;
    const startHi = Math.max(16, baseFromBox);

    // Fase 1: no-wrap (voorkeur)
    innerEl.classList.add("nowrap-mode");
    innerEl.classList.remove("softwrap-mode");

    let best = searchFontSize(innerEl, MIN_FS_PX, startHi, guardX, guardY);

    // Als we onder threshold zouden eindigen, probeer soft-wrap
    if (best < WRAP_THRESHOLD_PX) {
      innerEl.classList.remove("nowrap-mode");
      innerEl.classList.add("softwrap-mode");

      best = searchFontSize(innerEl, MIN_FS_PX, startHi, guardX, guardY);
    }

    return best;
  }

  async function mountThenFit(container) {
    await new Promise((r) => requestAnimationFrame(r));
    await new Promise((r) => requestAnimationFrame(r));
    fitAllIn(container);
  }

  // Fit alle label-inhouden in container
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

    // Padding op de label-rand, niet op de content die we schalen:
    const padPx = LABEL_PADDING_CM * PX_PER_CM * previewScale;
    label.style.padding = padPx + "px";

    const head = el(
      "div",
      { class: "label-head" },
      el("div", { class: "code-box line" }, values.code),
      el("div", { class: "line" }, values.desc)
    );

    // ====== wijziging #1: content wrapper zodat we content kunnen meten ======
    const content = el("div", { class: "label-content" });

    content.append(
      head,
      el("div", { class: "block-spacer" }),
      buildLeftBlock(values, size)
    );

    inner.append(content);
    label.append(inner);
    wrap.append(label, el("div", { class: "label-num" }, `Etiket ${size.idx}`));

    return wrap;
  }

  /* ====== Preview render ====== */
  function computePreviewScale(sizes) {
    // Houd je bestaande schaal-logica aan (zoals in v0.70)
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

  function buildPdfFileName(code) {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const ts =
      `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ` +
      `${pad(d.getHours())}.${pad(d.getMinutes())}.${pad(d.getSeconds())}`;

    const safeCode = (code || "export").trim() || "export";
    return `${safeCode} - ${ts}.pdf`;
  }

  async function generatePDFSingle() {
    const labelsGrid = $("#labelsGrid");
    if (!labelsGrid) throw new Error("labelsGrid niet gevonden");

    const JsPDF = loadJsPDF();
    if (!JsPDF) throw new Error("jsPDF niet geladen");

    // Render zeker up-to-date
    const vals = getFormValues();
    await renderPreviewFor(vals);

    const canvas = await html2canvas(labelsGrid, {
      scale: 2,
      backgroundColor: "#ffffff",
      useCORS: true,
    });

    const imgData = canvas.toDataURL("image/png");

    // A4 portrait (zoals je huidige), maar afbeelding 90° roteren zoals je “goede” PDF
    const pdf = new JsPDF({
      orientation: "portrait",
      unit: "pt",
      format: "a4",
    });

    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    const margin = 24;
    const maxW = pageW - margin * 2;
    const maxH = pageH - margin * 2;

    const imgW = canvas.width;
    const imgH = canvas.height;

    // Na rotatie wisselen breedte/hoogte om
    const rot = 90; // als hij de verkeerde kant op draait: maak dit -90
    const rotW = imgH;
    const rotH = imgW;

    const ratio = Math.min(maxW / rotW, maxH / rotH);
    const drawW = rotW * ratio;
    const drawH = rotH * ratio;

    const x = (pageW - drawW) / 2;
    const y = (pageH - drawH) / 2;

    // jsPDF: addImage(..., alias, compression, rotation)
    pdf.addImage(imgData, "PNG", x, y, drawW, drawH, undefined, "FAST", rot);

    pdf.save(buildPdfFileName(vals.code));
  }

  /* ====== init ====== */
  function init() {
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
