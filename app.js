(() => {
  /*
  ============================================================================
  LABEL GENERATOR — app.js (client-side controller)
  ----------------------------------------------------------------------------
  What this script does
  - Handles tab navigation (Single / Bulk upload / Help).
  - Single flow: read inputs -> compute 4 label faces -> render preview -> fit typography.
  - PDF flow: render labels in a stable offscreen layout -> capture via html2canvas -> build a single PDF (jsPDF).
  - Bulk flow: parse XLSX/CSV -> map columns -> validate rows -> generate many PDFs -> zip as download (JSZip).

  Non-obvious invariants (business rules)
  - Box dimensions are validated: 5–120 cm inclusive.
  - Label face size is always 0.9 * corresponding box face (10% smaller each side).
  - Layout (Standard / Stacked / Columns) is derived from the bucket key:
      * PORTRAIT: NARROW/STANDARD => Stacked, WIDE => Standard
      * LANDSCAPE: SHORT => Columns, STANDARD/HIGH => Standard
      * SQUARE: Standard
  - Typography is bucket-driven: labelBuckets.json provides anchor font sizes (pt) per bucket.
    We scale anchors to the actual label size and then apply a final fit/scale fallback if needed.

  Maintenance tips
  - HTML element IDs used below must remain stable (see index.html).
  - Keep color/styling out of JS: use CSS classes and CSS variables instead.
  ============================================================================
  */

  /* ====== CONSTANTS ======
   Unit conversions, thresholds, and shared state.
   Keep these centralized: they affect preview sizing, PDF sizing, and fit logic.
*/
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

  // Enforce allowed box dimension range (cm)
  const BOX_CM_MIN = 5;
  const BOX_CM_MAX = 120;
  /* ====== BUCKET CONFIG ======
   Bucket anchors are loaded from labelBuckets.json.
   BUCKET_BY_KEY provides O(1) lookup for typography anchors.
*/
  let BUCKET_CONFIG = null;
  let BUCKET_BY_KEY = new Map();

  /**
   * Fetches the bucket typography config (labelBuckets.json).
   * This MUST be served via http(s) due to fetch() restrictions; file:// will fail in most browsers.
   * @param {string} url - Relative/absolute URL to the JSON config.
   * @returns {Promise<object>} Parsed config object with an `anchors` array.
   */

  async function loadBucketConfig(url = "./labelBuckets.json") {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) {
      throw new Error(
        `Bucket-config kon niet worden geladen (${res.status} ${res.statusText}). ` +
          `Tip: open dit via een (lokale) webserver i.p.v. file://.`,
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

  /**
   * Computes a bucket key from label face dimensions.
   * Bucket key format examples:
   *   - SQUARE_SMALL
   *   - PORTRAIT_MEDIUM_STANDARD
   *   - LANDSCAPE_EXTRA_LARGE_SHORT
   * Note: EXTRA_LARGE is two tokens and must be handled carefully when parsing later.
   * @param {number} W_cm - label width in cm
   * @param {number} H_cm - label height in cm
   * @returns {string|null} Bucket key or null if inputs invalid.
   */

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
    else if (D >= 70 && D <= 120) sizeClass = "EXTRA_LARGE";

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

  /* ====== DOM HELPERS ======
   Small helpers for querying and building DOM safely.
   NOTE: Avoid adding styling inline; prefer classes handled by styles.css.
*/
  const $ = (sel) => document.querySelector(sel);

  function initTabs() {
    const tabButtons = Array.from(
      document.querySelectorAll(".tab-btn[data-tab-target]"),
    );
    const panels = Array.from(document.querySelectorAll(".tab-panel"));

    if (!tabButtons.length || !panels.length) return;

    const idToHash = (id) => {
      if (id === "tab-bulk") return "bulk";
      if (id === "tab-doc") return "help";
      return "single";
    };

    const hashToId = (hash) => {
      const h = (hash || "").replace("#", "").toLowerCase();
      if (h === "bulk") return "tab-bulk";
      if (
        h === "help" ||
        h === "hulp" ||
        h === "doc" ||
        h === "docs" ||
        h === "documentatie"
      )
        return "tab-doc";
      return "tab-single";
    };

    const setActive = (targetId, { updateHash = true } = {}) => {
      // Panels
      panels.forEach((p) => {
        const isActive = p.id === targetId;
        p.classList.toggle("hidden", !isActive);
      });

      // Buttons + ARIA
      tabButtons.forEach((b) => {
        const isActive = b.dataset.tabTarget === targetId;
        b.classList.toggle("active", isActive);
        b.setAttribute("aria-selected", isActive ? "true" : "false");
        b.tabIndex = isActive ? 0 : -1;
      });

      if (updateHash) {
        const newHash = idToHash(targetId);
        if (location.hash !== "#" + newHash) {
          history.replaceState(null, "", "#" + newHash);
        }
      }

      // UX: keep all tabs starting at the same viewport position.
      // This prevents perceived "jumping" between tabs caused by retained scroll.
      requestAnimationFrame(() => {
        window.scrollTo({ top: 0, left: 0, behavior: "auto" });
      });
    };

    tabButtons.forEach((b) => {
      b.addEventListener("click", () => {
        setActive(b.dataset.tabTarget, { updateHash: true });
      });

      // Keyboard nav (Left/Right)
      b.addEventListener("keydown", (e) => {
        if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
        e.preventDefault();
        const idx = tabButtons.indexOf(b);
        const delta = e.key === "ArrowRight" ? 1 : -1;
        const next =
          tabButtons[(idx + delta + tabButtons.length) % tabButtons.length];
        next.focus();
        setActive(next.dataset.tabTarget, { updateHash: true });
      });
    });

    // Initial state from hash
    setActive(hashToId(location.hash), { updateHash: false });

    // React on hash changes (back/forward)
    window.addEventListener("hashchange", () => {
      setActive(hashToId(location.hash), { updateHash: false });
    });
  }

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

  /* ====== SINGLE INPUT VALIDATION (5–120 cm) ======
   Only len/wid/hei are range-validated.
   Errors are shown inline (adds .is-invalid and a .field-error element).
*/
  function ensureInlineError(inputEl) {
    if (!inputEl) return null;
    const host = inputEl.closest(".field") || inputEl.parentElement;
    if (!host) return null;

    let el = host.querySelector(`.field-error[data-for="${inputEl.id}"]`);
    if (!el) {
      el = document.createElement("div");
      el.className = "field-error";
      el.dataset.for = inputEl.id;
      el.setAttribute("aria-live", "polite");
      host.appendChild(el);
    }
    return el;
  }

  function setFieldError(inputEl, msg) {
    const errEl = ensureInlineError(inputEl);
    if (!inputEl) return;

    const isInvalid = Boolean(msg);
    inputEl.classList.toggle("is-invalid", isInvalid);
    inputEl.setAttribute("aria-invalid", isInvalid ? "true" : "false");

    if (errEl) errEl.textContent = msg || "";
  }

  function validateBoxFieldCm(inputEl) {
    if (!inputEl) return true;

    const raw = inputEl.value;
    const n = parseNumber(raw);

    if (n === "") {
      setFieldError(inputEl, "Vul een getal in (cm).");
      return false;
    }
    if (n < BOX_CM_MIN || n > BOX_CM_MAX) {
      setFieldError(
        inputEl,
        `Moet tussen ${BOX_CM_MIN} en ${BOX_CM_MAX} cm liggen.`,
      );
      return false;
    }

    setFieldError(inputEl, "");
    return true;
  }

  function validateSingleBoxFields({ focusFirst = true } = {}) {
    const ids = ["len", "wid", "hei"];
    let firstInvalid = null;
    let ok = true;

    ids.forEach((id) => {
      const el = document.getElementById(id);
      const valid = validateBoxFieldCm(el);
      if (!valid && !firstInvalid) firstInvalid = el;
      ok = ok && valid;
    });

    if (!ok && focusFirst && firstInvalid) firstInvalid.focus();
    return ok;
  }

  function initSingleBoxValidation() {
    ["len", "wid", "hei"].forEach((id) => {
      const el = document.getElementById(id);
      if (!el) return;

      // Show error on blur; clear as soon as the value becomes valid again.
      el.addEventListener("blur", () => validateBoxFieldCm(el));
      el.addEventListener("input", () => {
        if (el.classList.contains("is-invalid")) validateBoxFieldCm(el);
      });
    });
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
    const desc = innerEl.querySelector(".label-desc");
    if (!desc) return;

    // In columns-layout the description should use the natural left-column width.
    if (innerEl.classList.contains("layout-columns")) {
      desc.style.setProperty("--desc-w", "auto");
      return;
    }

    // In Standard/Stacked we want ERP + productomschrijving to be as wide as the label allows
    // (inside the label padding). This reduces unnecessary wrapping and prevents extra fallback
    // scaling that makes the top block look "too narrow".
    const innerW =
      innerEl.clientWidth || innerEl.getBoundingClientRect().width || 0;
    if (innerW > 0) {
      desc.style.setProperty("--desc-w", Math.floor(innerW) + "px");
      return;
    }

    // Fallback: if layout has not stabilized yet, sync to specs grid width (legacy behavior).
    const grid = innerEl.querySelector(".specs-grid");
    if (!grid) return;

    const w = grid.offsetWidth || grid.getBoundingClientRect().width;
    desc.style.setProperty("--desc-w", w + "px");
  }

  function descFitsInMaxLines(descEl, maxLines = 3) {
    const cs = getComputedStyle(descEl);
    const lh = parseFloat(cs.lineHeight);
    const lineH =
      Number.isFinite(lh) && lh > 0
        ? lh
        : (parseFloat(cs.fontSize) || 12) * 1.15;
    const maxH = lineH * maxLines + 0.5; // tolerantie
    return descEl.scrollHeight <= maxH;
  }

  function shrinkDescToMaxLines(innerEl, maxLines = 3) {
    const desc = innerEl.querySelector(".label-desc");
    if (!desc) return;

    // Eerst breedte syncen, anders klopt wrap niet (zeker bij Standard/Stacked)
    syncDescWidthToSpecs(innerEl);

    // Reset naar CSS var (anchor-typografie)
    desc.style.fontSize = "";

    if (descFitsInMaxLines(desc, maxLines)) return;

    const basePx = parseFloat(getComputedStyle(desc).fontSize) || 12;

    // Binary search: verklein alleen de omschrijving tot hij binnen maxLines past
    // Begrens: nooit kleiner dan 70% van het anchor, en nooit onder 8px.
    const minPx = Math.max(8, basePx * 0.7);

    let lo = minPx;
    let hi = Math.max(minPx, basePx);
    let best = lo;

    for (let i = 0; i < 18; i++) {
      const mid = (lo + hi) / 2;
      desc.style.fontSize = mid + "px";
      if (descFitsInMaxLines(desc, maxLines)) {
        best = mid;
        hi = mid;
      } else {
        lo = mid;
      }
    }

    desc.style.fontSize = best + "px";
  }

  /* ====== LABEL GEOMETRY ======
   Calculates the four label faces based on box dimensions.
   Business rule: 0.9 scaling factor in both directions.
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

  function layoutToUiName(layout) {
    const l = String(layout || "").toUpperCase();
    if (l === "STACKED") return "Stacked";
    if (l === "COLUMNS") return "Columns";
    return "Standard";
  }

  function determineBucket(W_cm, H_cm) {
    // UI helper: return the effective bucket key (incl. fallbacks) and layout.
    const anchor = BUCKET_CONFIG ? getBucketAnchorFor(W_cm, H_cm) : null;
    const bucketKey = String(
      anchor?.key || selectBucketKeyFor(W_cm, H_cm) || "",
    );
    const layout = layoutForBucketKey(bucketKey);
    return { bucketKey, layout };
  }

  function renderDims(sizes, opts = {}) {
    const { includeBucket = false } = opts;
    const dims = $("#dims");
    if (!dims) return;

    dims.innerHTML = "";

    sizes.forEach((s) => {
      let bucketUi = "";
      if (includeBucket) {
        const picked = determineBucket(s.w, s.h);
        const name = bucketKeyToUiName(picked.bucketKey);
        bucketUi =
          name === "—" ? "—" : `${name} (${layoutToUiName(picked.layout)})`;
      }

      dims.append(
        el("div", { class: "dim" }, s.name),
        el("div", { class: "dim" }, format2(s.w)),
        el("div", { class: "dim" }, format2(s.h)),
        el("div", { class: "dim" }, bucketUi),
      );
    });
  }

  function bucketKeyToUiName(bucketKey) {
    if (!bucketKey) return "—";
    const parts = String(bucketKey).split("_").filter(Boolean);
    if (parts.length === 0) return "—";

    const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
    const family = cap(parts[0]);

    // Keys like: LANDSCAPE_SMALL_HIGH, SQUARE_EXTRA_LARGE, PORTRAIT_SMALL_NARROW
    let size = "";
    let idx = 1;
    if (parts[1] === "EXTRA" && parts[2] === "LARGE") {
      size = "Extra-Large";
      idx = 3;
    } else if (parts[1]) {
      size = cap(parts[1]);
      idx = 2;
    }

    const rest = parts.slice(idx).map(cap);
    return [family, size, ...rest].filter(Boolean).join("-");
  }

  /**
   * Maps a bucket key to a layout class.
   * Layout drives CSS structure:
   *   - STANDARD: ERP + description + specs stacked vertically
   *   - STACKED: portrait narrow/standard uses a tighter vertical stacking
   *   - COLUMNS: landscape short uses a two-column layout for specs
   * IMPORTANT: Keys containing EXTRA_LARGE split into two tokens (EXTRA, LARGE).
   * @param {string} bucketKey
   * @returns {"STANDARD"|"STACKED"|"COLUMNS"}
   */

  function layoutForBucketKey(bucketKey) {
    const parts = String(bucketKey || "").split("_");
    const family = (parts[0] || "").toUpperCase();

    // bucketKey examples:
    // - LANDSCAPE_SMALL_SHORT
    // - LANDSCAPE_EXTRA_LARGE_SHORT  (sizeClass is two tokens)
    // - PORTRAIT_MEDIUM_STANDARD
    // - SQUARE_LARGE
    let variant = "";
    if (
      parts[1] &&
      parts[2] &&
      parts[1].toUpperCase() === "EXTRA" &&
      parts[2].toUpperCase() === "LARGE"
    ) {
      variant = parts.slice(3).join("_").toUpperCase();
    } else {
      variant = parts.slice(2).join("_").toUpperCase();
    }

    // Layout mapping (per bucket family + variant)
    // Square: always Standard
    // Portrait: Narrow + Standard => Stacked, Wide => Standard
    // Landscape: Short => Columns, Standard + High => Standard
    if (family === "SQUARE") return "STANDARD";
    if (family === "PORTRAIT")
      return variant === "WIDE" ? "STANDARD" : "STACKED";
    if (family === "LANDSCAPE")
      return variant === "SHORT" ? "COLUMNS" : "STANDARD";
    return "STANDARD";
  }

  function applyBucketLayout(innerEl, bucketKey) {
    const layout = layoutForBucketKey(bucketKey);

    innerEl.classList.remove("layout-stacked", "layout-columns");
    if (layout === "STACKED") innerEl.classList.add("layout-stacked");
    if (layout === "COLUMNS") innerEl.classList.add("layout-columns");

    innerEl.dataset.layout = layout;
    return layout;
  }

  function readCurrentVariants(labelsGrid) {
    const out = [];
    for (let i = 1; i <= 4; i++) {
      const inner = labelsGrid?.querySelector(`#label${i} .label-inner`);
      const key = inner?.dataset?.bucketKey || "";
      out.push({ label: i, bucketKey: key, name: bucketKeyToUiName(key) });
    }
    return out;
  }

  function renderVariants(variantItems) {
    const variant = $("#variant");
    if (!variant) return;
    variant.innerHTML = "";
    (variantItems || []).forEach((v) => {
      variant.append(
        el("div", { class: "pill" }, `Label ${v.label}: ${v.name}`),
      );
    });
  }

  /* ====== FITTING / OVERFLOW CONTROL ======
   Bucket typography sets target font sizes.
   Then we verify the content fits; if not, we scale down the entire content as a last resort.
   This ensures: 'always render' (no clipped text), even for worst-case inputs.
*/

  /* ====== FITTING / OVERFLOW CONTROL ======
   Bucket typography sets target font sizes.
   Then we verify the content fits; if not, we scale down the entire content as a last resort.
   This ensures: 'always render' (no clipped text), even for worst-case inputs.

   Regression note (2026-02):
   Some edge-case labels still showed overflow after the PDF aspect-ratio fix. Root cause:
   - A single child (e.g. description/ERP/footer) can overflow without increasing the parent's scrollWidth/scrollHeight
     (depends on layout mode and browser rounding), so the old fit-check could miss it.
   - The PDF capture path previously cloned DOM nodes, which can slightly change line wrapping in tight cases.
   We therefore:
   - detect overflow on key child elements, and
   - verify the *visual* bounding box after scaling.
*/
  function elementOverflows(el, tol = 0.5) {
    if (!el) return false;
    return (
      el.scrollWidth > el.clientWidth + tol ||
      el.scrollHeight > el.clientHeight + tol
    );
  }

  function intrinsicFits(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;

    // Detect overflow in individual cells/blocks (these can overflow without affecting parent scroll metrics).
    const watch = [
      ...content.querySelectorAll(".specs-grid .val"),
      ...content.querySelectorAll(".code-box"),
      ...content.querySelectorAll(".label-desc"),
      ...content.querySelectorAll(".footer-text"),
    ];
    if (watch.some((n) => elementOverflows(n))) return false;

    const grid = content.querySelector(".specs-grid");
    if (grid && elementOverflows(grid)) return false;

    return (
      content.scrollWidth <= innerEl.clientWidth - guardX &&
      content.scrollHeight <= innerEl.clientHeight - guardY
    );
  }

  function visualFits(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;

    const ir = innerEl.getBoundingClientRect();
    const cr = content.getBoundingClientRect();

    // Tolerantie: we allow subpixel rounding drift.
    const tol = 0.75;

    return (
      cr.left >= ir.left + guardX - tol &&
      cr.right <= ir.right - guardX + tol &&
      cr.top >= ir.top + guardY - tol &&
      cr.bottom <= ir.bottom - guardY + tol
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

    // Corrigeer extra voor overflow in specifieke blokken (vals/desc/ERP/footer)
    let scaleChild = 1;
    const candidates = [
      ...content.querySelectorAll(".specs-grid .val"),
      ...content.querySelectorAll(".code-box"),
      ...content.querySelectorAll(".label-desc"),
      ...content.querySelectorAll(".footer-text"),
    ];
    candidates.forEach((el) => {
      const elSw = el.scrollWidth;
      const elCw = el.clientWidth;
      if (elSw > elCw + 0.5) {
        scaleChild = Math.min(scaleChild, elCw / elSw);
      }
      const elSh = el.scrollHeight;
      const elCh = el.clientHeight;
      if (elSh > elCh + 0.5) {
        scaleChild = Math.min(scaleChild, elCh / elSh);
      }
    });

    const k = Math.max(MIN_SCALE_K, Math.min(1, scaleW, scaleH, scaleChild));
    content.style.setProperty("--k", String(k));
    return k;
  }

  function ensureContentFits(innerEl, guardX, guardY) {
    const content = innerEl.querySelector(".label-content") || innerEl;

    // Reset fallback-scale bij nieuwe metingen
    if (content) content.style.setProperty("--k", "1");

    // 1) Intrinsic fit (no transform scaling needed)
    if (intrinsicFits(innerEl, guardX, guardY)) {
      if (content) content.style.setProperty("--k", "1");
      return 1;
    }

    // 2) Apply fallback scaling (based on intrinsic scroll metrics)
    let k = applyScaleFallback(innerEl, guardX, guardY);

    // 3) Verify visual bounding box (after transforms). If still clipped, apply an extra safety factor.
    if (!visualFits(innerEl, guardX, guardY)) {
      const ir = innerEl.getBoundingClientRect();
      const cr = content.getBoundingClientRect();

      const availW = Math.max(1, ir.width - guardX);
      const availH = Math.max(1, ir.height - guardY);

      const extraW = availW / Math.max(1, cr.width);
      const extraH = availH / Math.max(1, cr.height);
      const extra = Math.min(1, extraW, extraH);

      k = Math.max(MIN_SCALE_K, Math.min(1, k * extra * 0.995));
      content.style.setProperty("--k", String(k));
    }

    return k;
  }

  function applyBucketThenFit(innerEl) {
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;

    const guardX = Math.max(2, w * 0.015);
    const guardY = Math.max(2, h * 0.015);

    // 1) apply bucket typography
    const info = applyBucketTypography(innerEl);

    // 1b) apply bucket layout (standard / stacked / columns)
    if (innerEl.dataset.bucketKey) {
      applyBucketLayout(innerEl, innerEl.dataset.bucketKey);
    }

    // 2) wrap mode: start no-wrap; enable soft-wrap when text becomes small
    innerEl.classList.add("nowrap-mode");
    innerEl.classList.remove("softwrap-mode");

    const textPx = info?.sizesPx?.textPx;
    if (Number.isFinite(textPx) && textPx < WRAP_THRESHOLD_PX) {
      innerEl.classList.remove("nowrap-mode");
      innerEl.classList.add("softwrap-mode");
    }

    // 3) ensure description stays within maxLines (wrap width matters per layout)
    syncDescWidthToSpecs(innerEl);
    shrinkDescToMaxLines(innerEl, 3);

    // 4) final safety net: scale down whole content if needed (intrinsic + visual check)
    ensureContentFits(innerEl, guardX, guardY);
  }

  async function mountThenFit(container) {
    // Wait for fonts + layout. This improves stability of wrapping/line metrics in tight cases.
    if (document.fonts && document.fonts.ready) {
      try {
        await document.fonts.ready;
      } catch (_) {
        /* ignore */
      }
    }
    await new Promise((r) => requestAnimationFrame(r));
    await new Promise((r) => requestAnimationFrame(r));
    fitAllIn(container);
  }

  function fitAllIn(container) {
    container.querySelectorAll(".label-inner").forEach((inner) => {
      applyBucketThenFit(inner);
    });
  }

  /* ====== LABEL DOM CONSTRUCTION ======
   createLabelEl() builds the DOM structure that CSS expects:
     .label (size) -> .label-inner (layout + vars) -> .label-content (actual content)
*/

  function buildSpecsGrid(values) {
    const grid = el("div", { class: "specs-grid" });

    grid.append(
      el("div", { class: "key" }, "EAN:"),
      el("div", { class: "val" }, values.ean || ""),
      el("div", { class: "key" }, "QTY:"),
      el("div", { class: "val" }, `${values.qty || ""} PCS`),
      el("div", { class: "key" }, "G.W:"),
      el("div", { class: "val" }, `${values.gw || ""} KGS`),
      el("div", { class: "key" }, "CBM:"),
      el("div", { class: "val" }, values.cbm || ""),
      el("div", { class: "key" }, "BATCH:"),
      el("div", { class: "val" }, values.batch || ""),
    );

    return grid;
  }

  function footerTextForLabel(size, largestTwo) {
    // If largestTwo is null => fallback to old behavior (fb => C/N, side => Made in China)
    const useMadeInChina = largestTwo
      ? largestTwo.has(size.idx)
      : size.type !== "fb";
    return useMadeInChina ? "MADE IN CHINA" : "C/N: ___________________";
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

    const erpBox = el(
      "div",
      { class: "erp-box" },
      el("div", { class: "code-box line" }, values.code),
    );

    const descEl = el(
      "div",
      { class: "line label-desc product-desc" },
      values.desc,
    );

    const specs = buildSpecsGrid(values);

    const footerEl = el(
      "div",
      { class: "footer-text" },
      footerTextForLabel(size, largestTwo),
    );

    const content = el(
      "div",
      { class: "label-content" },
      erpBox,
      descEl,
      specs,
      footerEl,
    );

    // Apply an initial bucket/layout guess early (helps preview layout before fitting).
    const guessedKey = selectBucketKeyFor(size.w, size.h);
    if (guessedKey) {
      inner.dataset.bucketKey = guessedKey;
      applyBucketLayout(inner, guessedKey);
    }

    inner.append(content);

    label.append(inner);
    wrap.append(label, el("div", { class: "label-num" }, `Etiket ${size.idx}`));

    return wrap;
  }

  /* ====== PREVIEW RENDERING ======
   Renders 4 labels in a 2x2 grid.
   Uses a computed previewScale so large labels still fit in the viewport.
*/
  function computePreviewScale(sizes, containerEl) {
    const labelsGrid = containerEl || $("#labelsGrid");
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

  /**
   * Renders the 4 labels into a target container (Single preview grid or offscreen batch host).
   * Also updates the calculated dimensions table unless opts.renderDims === false.
   * @param {object} values - Form values (dimensions + label fields)
   * @param {object} [opts]
   * @returns {Promise<{sizes:Array, scale:number, container:Element}|void>}
   */

  async function renderPreviewFor(values, opts = {}) {
    const visibleGrid = $("#labelsGrid");
    const target = opts.targetEl || visibleGrid;
    if (!target) return;

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

    // UI-updates alleen als expliciet gewenst
    if (opts.renderDims !== false) {
      renderDims(sizes, { includeBucket: !!opts.showVariant });
    }

    const scale =
      typeof opts.previewScale === "number"
        ? opts.previewScale
        : computePreviewScale(sizes, target);

    currentPreviewScale = scale;

    target.innerHTML = "";
    const fragments = document.createDocumentFragment();
    sizes.forEach((size) => {
      fragments.append(createLabelEl(size, values, scale, largestTwo));
    });
    target.append(fragments);

    await mountThenFit(target);

    return { sizes, scale, container: target };
  }
  async function renderSingle() {
    const vals = getFormValues();

    // Empty state: no box dimensions entered yet
    if (vals.len === "" && vals.wid === "" && vals.hei === "") {
      const grid = $("#labelsGrid");
      const dims = $("#dims");

      if (grid) {
        grid.innerHTML = "";
        grid.append(
          el("div", { class: "preview-placeholder" }, "Hier komt de preview"),
        );
      }
      if (dims) dims.innerHTML = "";
      return;
    }

    await renderPreviewFor(vals, { showVariant: true });
  }

  /* ====== PDF GENERATION ======
   1) Render labels (often offscreen) at a stable scale.
   2) Capture each label DOM node to a canvas via html2canvas.
   3) Rotate (labels are placed rotated in the PDF).
   4) Add images to jsPDF with cm units for true-size output.
*/
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

  async function captureLabelToRotatedPng(labelIdx, scopeEl) {
    const scope = scopeEl || document;
    const src = scope.querySelector(`.label[data-idx="${labelIdx}"]`);
    if (!src)
      throw new Error(`Label ${labelIdx} niet gevonden voor PDF-capture.`);

    // Ensure fonts are ready before capture (prevents late reflow that can change wrapping).
    // document.fonts is widely supported in modern browsers; we guard it for older ones.
    if (document.fonts && document.fonts.ready) {
      try {
        await document.fonts.ready;
      } catch (_) {
        /* ignore */
      }
    }

    // Capture the *already fitted* DOM node directly. Cloning can subtly change layout
    // (e.g. line-wrapping / subpixel rounding), which can reintroduce overflow in edge cases.
    const capScale = Math.max(
      2,
      window.devicePixelRatio || 1,
      1 / (currentPreviewScale || 1),
    );

    const canvas = await html2canvas(src, {
      scale: capScale,
      backgroundColor: "#ffffff",
      useCORS: true,
    });

    const rot = rotateCanvas90CW(canvas);
    return rot.toDataURL("image/png");
  }

  /**
   * Generates a single combined PDF for the current Single form values.
   * Uses cm units in jsPDF so the output can be printed at 100% scale (true size).
   * The Help tab contains vendor instructions: print at 100%, do not auto-scale.
   */

  async function generatePDFSingle() {
    const JsPDF = loadJsPDF();
    if (!JsPDF) throw new Error("jsPDF niet geladen");
    if (!window.html2canvas) throw new Error("html2canvas niet geladen");

    const vals = getFormValues();
    const pdfHost = ensurePdfRenderHost();
    const result = await renderPreviewFor(vals, {
      targetEl: pdfHost,
      previewScale: 1,
      renderDims: false,
    });
    if (!result) throw new Error("Kon preview niet renderen voor PDF.");
    const { sizes, container } = result;

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
      const imgData = await captureLabelToRotatedPng(idx, container);

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
        "FAST",
      );
      y += hRot;
    }

    pdf.save(buildPdfFileName(vals.code));
  }

  /* ====== BULK UPLOAD / BATCH ======
   Parses an uploaded sheet and maps columns to required fields.
   Validates all rows before generating any PDFs to avoid partial output.
   Generates one PDF per row and bundles them in a ZIP.
*/
  // Batch state
  let parsedRows = [];
  let isBatchRunning = false;
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

  // Offscreen render host for batch: avoids UI reflow/jumping by rendering labels outside viewport.
  let batchRenderHost = null;

  function ensureBatchRenderHost() {
    if (batchRenderHost) return batchRenderHost;
    const host = document.createElement("div");
    host.id = "batchRenderHost";
    host.className = "batch-render-host pdf-mode";
    host.setAttribute("aria-hidden", "true");
    document.body.appendChild(host);
    batchRenderHost = host;
    return host;
  }

  // Dedicated off-screen render host for single-PDF generation.
  // We render labels here with PDF-specific box model rules (see .pdf-mode in CSS),
  // so html2canvas captures the correct physical aspect ratio (no padding/border expansion).
  let pdfRenderHost = null;

  function ensurePdfRenderHost() {
    if (pdfRenderHost) return pdfRenderHost;
    const host = document.createElement("div");
    host.id = "pdfRenderHost";
    host.className = "pdf-render-host pdf-mode";
    host.setAttribute("aria-hidden", "true");
    document.body.appendChild(host);
    pdfRenderHost = host;
    return host;
  }

  /* ====== PDF GEOMETRY SELF-TEST (regression guard) ======
   Why this exists
   - The generator depends on a DOM->canvas->PDF pipeline (html2canvas + jsPDF).
   - A subtle box-model mismatch (content-box vs border-box) can introduce anisotropic scaling
     in the embedded PNGs, which becomes visible only for extreme aspect ratios.
   This self-test generates a curated edge-case set and validates:
   (a) PDF MediaBox page size (true-size output), and
   (b) embedded image pixel ratio (Width/Height) ≈ expected label ratio (H/W) within 0.2%.
  */

  const CM_TO_PT = 72 / 2.54; // PDF points per centimeter
  const PDF_GEOMETRY_RATIO_TOL_REL = 0.002; // 0.2% relative tolerance
  const PDF_GEOMETRY_PAGE_TOL_PT = 0.5; // ~0.18mm tolerance on MediaBox

  // Curated edge-case set:
  // - Small H (5–15cm) with large L/W
  // - Small L/W with large H
  // - Boundaries (5 and 100)
  // - A few decimal dimensions to exercise rounding
  const PDF_GEOMETRY_EDGE_CASES = [
    // Small H, large L/W
    { name: "H05_L120_W120", len: 120, wid: 120, hei: 5 },
    { name: "H05_L120_W60", len: 120, wid: 60, hei: 5 },
    { name: "H06_L110_W30", len: 110, wid: 30, hei: 6 },
    { name: "H08_L100_W20", len: 100, wid: 20, hei: 8 },
    { name: "H10_L120_W50", len: 120, wid: 50, hei: 10 },
    { name: "H12_L90_W25", len: 90, wid: 25, hei: 12 },
    { name: "H15_L120_W40", len: 120, wid: 40, hei: 15 },

    // Small L/W, large H
    { name: "L05_W05_H120", len: 5, wid: 5, hei: 120 },
    { name: "L06_W10_H120", len: 6, wid: 10, hei: 120 },
    { name: "L10_W15_H100", len: 10, wid: 15, hei: 100 },
    { name: "L12_W12_H90", len: 12, wid: 12, hei: 90 },

    // Very narrow / very wide faces (within allowed 5–120cm)
    { name: "L120_W05_H15", len: 120, wid: 5, hei: 15 },
    { name: "L05_W120_H15", len: 5, wid: 120, hei: 15 },

    // Square-ish / stable baselines
    { name: "L60_W60_H60", len: 60, wid: 60, hei: 60 },
    { name: "L25_W25_H10", len: 25, wid: 25, hei: 10 },

    // Realistic max-width case (matches the 120×120 scenario)
    { name: "L120_W120_H60", len: 120, wid: 120, hei: 60 },

    // Decimals / rounding surfaces
    { name: "L63_5_W47_2_H12_3", len: 63.5, wid: 47.2, hei: 12.3 },
    { name: "L119_9_W5_1_H14_9", len: 119.9, wid: 5.1, hei: 14.9 },
  ];
  // Stress strings for layout overflow regression (long tokens + hyphens + unicode dash).
  const PDF_LAYOUT_STRESS_DESC =
    "8719327417447 - BarDeluxe - 5-piece - Boston - Black / Slow Juicer – Cherry Red";
  const PDF_LAYOUT_STRESS_EAN = "8719327417447";

  function _pdfLayoutCheck(container) {
    const issues = [];
    for (let i = 1; i <= 4; i++) {
      const inner = container?.querySelector(`#label${i} .label-inner`);
      if (!inner) continue;

      const w = inner.clientWidth;
      const h = inner.clientHeight;
      const guardX = Math.max(2, w * 0.015);
      const guardY = Math.max(2, h * 0.015);

      // After renderPreviewFor, fitAllIn() already ran. We validate that nothing is clipped.
      let ok =
        visualFits(inner, guardX, guardY) &&
        ![
          ...inner.querySelectorAll(
            ".specs-grid .val, .code-box, .label-desc, .footer-text",
          ),
        ].some((n) => elementOverflows(n));

      // Regression guard: in Standard/Stacked the description should not be artificially narrowed
      // to the specs-grid width (which increases wrapping and can trigger fallback scaling).
      // Use offsetWidth (unaffected by transform scale) for a stable check.
      const desc = inner.querySelector(".label-desc");
      if (ok && desc && !inner.classList.contains("layout-columns")) {
        const innerW = inner.clientWidth || 1;
        const descW = desc.offsetWidth || 0;
        const ratio = descW / innerW;
        if (ratio < 0.85) ok = false;
      }

      if (!ok) {
        issues.push({
          label: i,
          bucketKey: inner.dataset.bucketKey || "",
          layout: inner.dataset.layout || "",
          w,
          h,
        });
      }
    }
    return issues;
  }

  function _pdfGeomRelDiff(a, b) {
    const denom = Math.abs(b) > 0 ? Math.abs(b) : 1;
    return Math.abs(a - b) / denom;
  }

  function _pdfGeomFmtPct(x) {
    return `${(x * 100).toFixed(3)}%`;
  }

  function _pdfGeomDecodePdfText(arrayBuffer) {
    // We only need ASCII tokens (MediaBox, Width/Height). UTF-8 decoding is sufficient.
    try {
      return new TextDecoder("iso-8859-1").decode(new Uint8Array(arrayBuffer));
    } catch {
      return new TextDecoder("utf-8").decode(new Uint8Array(arrayBuffer));
    }
  }

  function _pdfGeomExtractMediaBoxPt(pdfText) {
    // Match first MediaBox: [x0 y0 x1 y1]
    const re =
      /\/MediaBox\s*\[\s*([0-9.+-]+)\s+([0-9.+-]+)\s+([0-9.+-]+)\s+([0-9.+-]+)\s*\]/;
    const m = re.exec(pdfText);
    if (!m) return null;
    const x0 = parseFloat(m[1]);
    const y0 = parseFloat(m[2]);
    const x1 = parseFloat(m[3]);
    const y1 = parseFloat(m[4]);
    if (![x0, y0, x1, y1].every(Number.isFinite)) return null;
    return { w: x1 - x0, h: y1 - y0, raw: { x0, y0, x1, y1 } };
  }

  function _pdfGeomExtractTopImageRatios(pdfText, topN = 4) {
    const hits = [];
    const re = /\/Subtype\s*\/Image\b/g;
    let m;
    while ((m = re.exec(pdfText))) {
      const chunk = pdfText.slice(m.index, m.index + 2500);
      const wm = /\/Width\s+(\d+)/.exec(chunk);
      const hm = /\/Height\s+(\d+)/.exec(chunk);
      if (!wm || !hm) continue;
      const w = parseInt(wm[1], 10);
      const h = parseInt(hm[1], 10);
      if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0)
        continue;
      hits.push({ w, h, area: w * h, ratio: w / h });
    }
    hits.sort((a, b) => b.area - a.area);
    return hits.slice(0, topN);
  }

  async function _buildPdfArrayBufferForValues(vals, opts = {}) {
    const JsPDF = loadJsPDF();
    if (!JsPDF) throw new Error("jsPDF niet geladen");
    if (!window.html2canvas) throw new Error("html2canvas niet geladen");

    const pdfHost = ensurePdfRenderHost();

    // Self-test uses fixed previewScale=1 for stable capture geometry.
    const result = await renderPreviewFor(vals, {
      targetEl: pdfHost,
      previewScale: 1,
      renderDims: false,
      showVariant: false,
    });
    if (!result)
      throw new Error("Kon preview niet renderen voor PDF self-test.");

    const { sizes, container } = result;
    const layoutIssues = _pdfLayoutCheck(container);
    const order = [1, 3, 2, 4];

    const pageW = Math.max(...sizes.map((s) => s.h)) + PDF_MARGIN_CM * 2;
    const pageH = sizes.reduce((sum, s) => sum + s.w, 0) + PDF_MARGIN_CM * 2;

    // Disable compression for deterministic, parse-friendly output.
    const pdf = new JsPDF({
      unit: "cm",
      orientation: "portrait",
      format: [pageW, pageH],
      compress: false,
    });

    let y = PDF_MARGIN_CM;

    for (const idx of order) {
      const s = sizes[idx - 1];
      const imgData = await captureLabelToRotatedPng(idx, container);

      // Rotated placement: width = H, height = W
      pdf.addImage(
        imgData,
        "PNG",
        PDF_MARGIN_CM,
        y,
        s.h,
        s.w,
        undefined,
        "FAST",
      );
      y += s.w;
    }

    const expectedRatios = order.map((idx) => {
      const s = sizes[idx - 1];
      // Rotated PNG ratio in PDF should be H/W
      return s.h / s.w;
    });

    const arrayBuffer = pdf.output("arraybuffer");
    return {
      arrayBuffer,
      layoutIssues,
      expected: { pageWcm: pageW, pageHcm: pageH, ratios: expectedRatios },
      pdf, // optional (debug)
    };
  }

  async function runPdfGeometrySelfTestUI() {
    const btnRun = $("#btnRunPdfGeometryTest");
    const btnDl = $("#btnDownloadPdfGeometryFailures");
    const summaryEl = $("#pdfGeometryTestSummary");
    const logEl = $("#pdfGeometryTestLog");

    if (!btnRun || !summaryEl || !logEl) return;

    const setBusy = (busy) => {
      btnRun.disabled = !!busy;
      btnRun.textContent = busy ? "Test draait…" : "Run PDF geometry test";
      if (busy && btnDl) {
        btnDl.disabled = true;
        btnDl.dataset.ready = "";
        btnDl.onclick = null;
      }
    };

    const logLine = (msg, cls) => {
      const div = document.createElement("div");
      if (cls) div.className = cls;
      div.textContent = msg;
      logEl.appendChild(div);
      logEl.scrollTop = logEl.scrollHeight;
    };

    // Reset UI
    logEl.textContent = "";
    summaryEl.textContent = "";
    setBusy(true);

    const failures = [];
    let pass = 0;
    let fail = 0;

    try {
      logLine(
        `Start PDF geometry self-test: ${PDF_GEOMETRY_EDGE_CASES.length} cases…`,
      );

      for (let i = 0; i < PDF_GEOMETRY_EDGE_CASES.length; i++) {
        const tc = PDF_GEOMETRY_EDGE_CASES[i];
        const code = `SELFTEST_${tc.name}`;

        const vals = {
          code,
          desc: PDF_LAYOUT_STRESS_DESC,
          ean: PDF_LAYOUT_STRESS_EAN,
          qty: "1",
          gw: "0",
          cbm: "0",
          len: tc.len,
          wid: tc.wid,
          hei: tc.hei,
          batch: "",
        };

        // Hard guard: keep cases within allowed ranges.
        // (If someone edits the list incorrectly, we want a clear error.)
        const dims = [vals.len, vals.wid, vals.hei];
        if (
          dims.some(
            (d) => !Number.isFinite(d) || d < BOX_CM_MIN || d > BOX_CM_MAX,
          )
        ) {
          throw new Error(
            `Edge-case buiten bereik (${BOX_CM_MIN}–${BOX_CM_MAX}cm): ${tc.name}`,
          );
        }

        const { arrayBuffer, expected, layoutIssues } =
          await _buildPdfArrayBufferForValues(vals);

        if (layoutIssues && layoutIssues.length) {
          fail++;
          failures.push({
            case: tc.name,
            kind: "LAYOUT_OVERFLOW",
            details: layoutIssues,
          });
          logLine(
            `FAIL ${tc.name}: layout overflow in labels ${layoutIssues
              .map((x) => x.label)
              .join(", ")}`,
            "fail",
          );
          continue;
        }

        const txt = _pdfGeomDecodePdfText(arrayBuffer);
        const mb = _pdfGeomExtractMediaBoxPt(txt);
        const imgs = _pdfGeomExtractTopImageRatios(txt, 4);

        // Page size check (MediaBox)
        const expWpt = expected.pageWcm * CM_TO_PT;
        const expHpt = expected.pageHcm * CM_TO_PT;

        let ok = true;
        let reason = [];

        if (!mb) {
          ok = false;
          reason.push("MediaBox niet gevonden");
        } else {
          const dw = Math.abs(mb.w - expWpt);
          const dh = Math.abs(mb.h - expHpt);
          if (dw > PDF_GEOMETRY_PAGE_TOL_PT || dh > PDF_GEOMETRY_PAGE_TOL_PT) {
            ok = false;
            reason.push(
              `MediaBox mismatch: ΔW=${dw.toFixed(2)}pt, ΔH=${dh.toFixed(2)}pt`,
            );
          }
        }

        // Image ratio check (PNG Width/Height ≈ expected H/W within 0.2%)
        if (!imgs.length) {
          ok = false;
          reason.push("Geen images gevonden");
        } else {
          // We validate that every embedded top image ratio matches at least one expected ratio,
          // and that every expected ratio is represented by at least one embedded ratio.
          const actualRatios = imgs.map((x) => x.ratio);
          const expectedRatios = expected.ratios;

          const matchesExpected = (ar) =>
            expectedRatios.some(
              (er) => _pdfGeomRelDiff(ar, er) <= PDF_GEOMETRY_RATIO_TOL_REL,
            );
          const coversActual = actualRatios.every(matchesExpected);

          const coversExpected = expectedRatios.every((er) =>
            actualRatios.some(
              (ar) => _pdfGeomRelDiff(ar, er) <= PDF_GEOMETRY_RATIO_TOL_REL,
            ),
          );

          if (!coversActual || !coversExpected) {
            ok = false;

            // Provide a useful delta readout: best match per expected ratio
            const deltas = expectedRatios.map((er) => {
              const best = Math.min(
                ...actualRatios.map((ar) => _pdfGeomRelDiff(ar, er)),
              );
              return best;
            });

            reason.push(
              `Image ratio mismatch (tol=${_pdfGeomFmtPct(PDF_GEOMETRY_RATIO_TOL_REL)}). ` +
                `bestΔ per expected: [${deltas.map(_pdfGeomFmtPct).join(", ")}]`,
            );
          }
        }

        if (ok) {
          pass++;
          logLine(
            `OK  ${tc.name}  (L=${tc.len}, W=${tc.wid}, H=${tc.hei})`,
            "ok",
          );
        } else {
          fail++;
          logLine(
            `ERR ${tc.name}  (L=${tc.len}, W=${tc.wid}, H=${tc.hei})  -> ${reason.join(" | ")}`,
            "err",
          );
          failures.push({
            name: tc.name,
            vals,
            arrayBuffer,
            reason: reason.join(" | "),
          });
        }

        // Yield to UI
        await new Promise((r) => setTimeout(r, 0));
      }

      summaryEl.textContent = `Resultaat: ${pass} OK, ${fail} fouten.`;
      if (failures.length && window.JSZip && btnDl) {
        btnDl.disabled = false;
        btnDl.dataset.ready = "1";
        btnDl.onclick = async () => {
          const zip = new JSZip();
          const stamp = buildTimestamp().replace(/[:]/g, ".");
          const folder = zip.folder(`pdf-geometry-failures - ${stamp}`);

          failures.forEach((f, idx) => {
            const blob = new Blob([f.arrayBuffer], { type: "application/pdf" });
            folder.file(
              `${String(idx + 1).padStart(2, "0")} - ${f.name}.pdf`,
              blob,
            );
            folder.file(
              `${String(idx + 1).padStart(2, "0")} - ${f.name}.txt`,
              f.reason,
            );
          });

          const zipBlob = await zip.generateAsync({ type: "blob" });
          downloadBlob(zipBlob, `pdf-geometry-failures - ${stamp}.zip`);
        };
      }
    } catch (err) {
      summaryEl.textContent = `Self-test afgebroken: ${err.message || err}`;
      logLine(String(err?.stack || err?.message || err), "err");
    } finally {
      setBusy(false);
    }
  }

  function initPdfGeometrySelfTestUI() {
    const btnRun = $("#btnRunPdfGeometryTest");
    if (!btnRun) return; // UI not present (future: separate builds)
    btnRun.addEventListener("click", () => {
      runPdfGeometrySelfTestUI().catch((e) => alert(e.message || e));
    });
  }

  function resetLog() {
    if (logList) logList.innerHTML = "";
  }

  function log(msg, type = "info") {
    if (!logList) return;
    const div = el(
      "div",
      { class: type === "error" ? "err" : type === "ok" ? "ok" : "" },
      msg,
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
        tr.appendChild(el("td", {}, String(rows[i][c] ?? ""))),
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
      return hdr ? (row[hdr] ?? "") : "";
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

  function validateRequiredRowValues(rows, mappingObj) {
    const errors = [];

    // Build "Field label (Excel column name)" messages so users can fix their sheet quickly.
    const fmt = (key, label) => {
      const col = mappingObj[key];
      return col ? `${label} (kolom: ${col})` : `${label} (kolom: —)`;
    };

    for (let i = 0; i < rows.length; i++) {
      const vals = readRowWithMapping(rows[i], mappingObj);
      const missing = [];

      if (!vals.code) missing.push(fmt("productcode", "ERP"));
      if (!vals.desc) missing.push(fmt("omschrijving", "Omschrijving"));
      if (!vals.ean) missing.push(fmt("ean", "EAN"));
      if (!vals.qty) missing.push(fmt("qty", "QTY"));
      if (!vals.gw) missing.push(fmt("gw", "G.W"));
      if (!vals.cbm) missing.push(fmt("cbm", "CBM"));
      if (!isFinite(vals.len)) missing.push(fmt("lengte", "Length (L)"));
      if (!isFinite(vals.wid)) missing.push(fmt("breedte", "Width (W)"));
      if (!isFinite(vals.hei)) missing.push(fmt("hoogte", "Height (H)"));
      if (!vals.batch) missing.push(fmt("batch", "Batch"));

      // +1 for header row in the uploaded sheet (Excel is 1-indexed and row 1 is headers)
      if (missing.length) {
        errors.push({ row: i + 2, missing });
      }
    }
    return errors;
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
        cols.map((c) => String(r[c] ?? "").replace(/;/g, ",")).join(";"),
      ),
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

  /**
   * Wires the Bulk upload UI:
   * - Drag/drop + file picker
   * - Column mapping dropdowns
   * - Preview of first rows
   * - Run/Abort generation with progress + log
   */

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

        const rowErrors = validateRequiredRowValues(parsedRows, mapping);
        if (rowErrors.length) {
          resetLog();
          setHidden(logWrap, false);
          log(
            `Fout: ${rowErrors.length} rij(en) missen verplichte velden. Er worden geen PDF’s gegenereerd.`,
            "error",
          );
          rowErrors.slice(0, 20).forEach((e) => {
            log(`Rij ${e.row}: ontbreekt ${e.missing.join(", ")}`, "error");
          });
          if (rowErrors.length > 20) {
            log(`... en nog ${rowErrors.length - 20} rijen.`, "error");
          }

          alert(
            `Bulk-upload bevat ${rowErrors.length} rij(en) met lege verplichte velden.\n` +
              `Er worden geen PDF’s gegenereerd.\n\n` +
              rowErrors
                .slice(0, 10)
                .map((e) => `Rij ${e.row}: ${e.missing.join(", ")}`)
                .join("\n"),
          );
          return;
        }

        // Bulk: geen debug-variant tonen.
        renderVariants([]);

        abortFlag = false;
        isBatchRunning = true;
        btnAbortBatch.disabled = false;
        setHidden(progressWrap, false);
        setHidden(logWrap, false);

        if (progressBar) progressBar.style.width = "0%";
        if (progressLabel)
          progressLabel.textContent = `0 / ${parsedRows.length}`;
        if (progressPhase) progressPhase.textContent = "Renderen…";

        const zip = new JSZip();
        const batchTime = buildTimestamp();
        const batchHost = ensureBatchRenderHost();

        let okCount = 0;
        let errCount = 0;

        for (let i = 0; i < parsedRows.length; i++) {
          if (abortFlag) break;
          const row = parsedRows[i];
          try {
            const vals = readRowWithMapping(row, mapping);
            // Render en capture met dezelfde pipeline als single
            // Zorg dat Bulk exact dezelfde previewScale gebruikt als Single,
            // zodat wrapping/fitting/bucket-typografie identiek uitpakken.
            const visibleGrid = document.querySelector("#labelsGrid");
            const stableScale = computePreviewScale(
              calcLabelSizes(vals),
              visibleGrid || batchHost,
            );

            const result = await renderPreviewFor(vals, {
              targetEl: batchHost,
              renderDims: false,
              previewScale: stableScale,
            });
            if (!result) throw new Error("Kon preview niet renderen.");

            // Maak PDF als blob
            const { sizes, container } = result;
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
              const imgData = await captureLabelToRotatedPng(idx, container);
              pdf.addImage(
                imgData,
                "PNG",
                PDF_MARGIN_CM,
                y,
                s.h,
                s.w,
                undefined,
                "FAST",
              );
              y += s.w;
            }

            const blob = pdf.output("blob");
            const safeCode = (vals.code || "export").trim() || "export";
            const name = `${safeCode} - ${batchTime} - R${String(
              i + 1,
            ).padStart(3, "0")}.pdf`;
            zip.file(name, blob);
            okCount++;
          } catch (err) {
            errCount++;
            log(`Rij ${i + 1}: renderfout: ${err.message || err}`, "error");
          }

          if (progressBar)
            progressBar.style.width = `${Math.round(
              ((i + 1) / parsedRows.length) * 100,
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

        isBatchRunning = false;

        if (progressPhase) progressPhase.textContent = "Klaar.";
      } catch (e) {
        isBatchRunning = false;
        alert(e.message || e);
      }
    });

    btnAbortBatch?.addEventListener("click", () => {
      abortFlag = true;
      btnAbortBatch.disabled = true;
      if (progressPhase) progressPhase.textContent = "Afbreken…";
    });
  }

  /* ====== INIT ======
   Entry point on DOMContentLoaded.
   - Loads bucket config (optional; app still works with fallback fitting)
   - Initializes tabs and batch UI
   - Wires Single buttons (Preview, Clear, PDF)
*/
  async function init() {
    try {
      BUCKET_CONFIG = await loadBucketConfig("./labelBuckets.json");
      indexBucketConfig(BUCKET_CONFIG);
    } catch (e) {
      console.error(e);
      alert(e.message || e);
      // Zonder config kan de rest nog draaien, maar bucket-typografie zal ontbreken.
    }

    initTabs();
    initBatchUI();
    initPdfGeometrySelfTestUI();

    initSingleBoxValidation();
    const btnGen = $("#btnGen");
    const btnPDF = $("#btnPDF");

    const safeRender = () => {
      if (!validateSingleBoxFields()) return;
      renderSingle().catch((err) => alert(err.message || err));
    };
    btnGen?.addEventListener("click", safeRender);

    btnPDF?.addEventListener("click", async () => {
      try {
        if (!validateSingleBoxFields()) return;
        await generatePDFSingle();
      } catch (err) {
        alert(err.message || err);
      }
    });

    window.addEventListener("resize", () => {
      if (isBatchRunning) return;
      renderSingle().catch(() => {});
    });

    renderSingle().catch(() => {});
  }

  document.addEventListener("DOMContentLoaded", () => {
    init().catch((e) => alert(e.message || e));
  });
})();
