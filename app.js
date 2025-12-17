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

  // ===== PDF (terug naar v0.31 gedrag) =====
  const PDF_MARGIN_CM = 0.5; // klopt met je oude PDF (20.8cm breedte bij 19.8cm labelhoogte)
  const BORDER_PX = 1; // safety: borders altijd meenemen in capture
  let currentPreviewScale = 1; // nodig voor scherpe PDF-capture bij geschaalde preview

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

  // /** past alles binnen 'innerEl' met een veiligheidsmarge? */
  // function fitsWithGuard(innerEl, guardX, guardY) {
  //   const content = innerEl.querySelector(".label-content") || innerEl;
  //   return (
  //     content.scrollWidth <= innerEl.clientWidth - guardX &&
  //     content.scrollHeight <= innerEl.clientHeight - guardY
  //   );
  // }

  function fitsWithGuard(innerEl, guardX, guardY) {
  const content = innerEl.querySelector(".label-content") || innerEl;

  // Belangrijk: detecteer overflow in grid-cellen (EAN/waarden)
  const valOverflow = Array.from(
    content.querySelectorAll(".specs-grid .val")
  ).some((v) => v.scrollWidth > v.clientWidth + 0.5);

  if (valOverflow) return false;

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

    let lo = minFs;
    let best = lo;

    applyFontSizes(innerEl, hi);
    if (fitsWithGuard(innerEl, guardX, guardY)) {
      best = hi;
      return best;
    }

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

  function fitContentToBoxConditional(innerEl) {
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;

    const guardX = Math.max(8, w * 0.02);
    const guardY = Math.max(8, h * 0.02);

    const baseFromBox = Math.min(w, h) * 0.11;
    const startHi = Math.max(16, baseFromBox);

    innerEl.classList.add("nowrap-mode");
    innerEl.classList.remove("softwrap-mode");

    let best = searchFontSize(innerEl, MIN_FS_PX, startHi, guardX, guardY);

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
    currentPreviewScale = scale; // belangrijk voor PDF capture kwaliteit

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
    // 90° met de klok mee
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

    // Forceer borders (html2canvas wil bij scaling soms 1 zijde "kwijtraken")
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

    // Scherpte: compenseer preview-schaal zodat PDF niet “zacht” wordt bij grote labels
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

    // Zorg dat preview up-to-date is (en currentPreviewScale gezet is)
    const vals = getFormValues();
    const result = await renderPreviewFor(vals);
    if (!result) throw new Error("Kon preview niet renderen voor PDF.");
    const { sizes } = result;

    // v0.31 gedrag:
    // - Elk label apart capturen
    // - Canvas 90° CW roteren
    // - In PDF stapelen (1,3,2,4) op een lange pagina
    const order = [1, 3, 2, 4];

    // Na rotatie: width = originele hoogte (h), height = originele breedte (w)
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
