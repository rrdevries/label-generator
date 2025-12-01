// app.js v0.33 – layout + scaling + (single & batch) PDF-generatie

(() => {
  'use strict';

  // ====== CONSTANTEN & UTILITIES ======

  const DOC = document;
  const ROOT = DOC.documentElement;

  const PX_PER_CM = (() => {
    const v = getComputedStyle(ROOT).getPropertyValue('--px-per-cm');
    const n = parseFloat(v.replace(',', '.'));
    return Number.isFinite(n) && n > 0 ? n : 37.7952755906;
  })();

  const SHRINK_FACTOR = 0.9;          // 0,9 × doosmaat → labelmaat per zijde
  const CODE_MULT = 1.35;             // ERP-font = basis × CODE_MULT
  const MIN_FS = 6;                   // minimale basis-fontgrootte (px)
  const MAX_FS_ABS = 40;              // absolute max (wordt nog begrensd door labelhoogte)

  const TARGET_FIELDS = [
    { key: 'boxLength', label: 'Lengte (L) [cm]', required: true },
    { key: 'boxWidth',  label: 'Breedte (W) [cm]', required: true },
    { key: 'boxHeight', label: 'Hoogte (H) [cm]', required: true },
    { key: 'prodCode',  label: 'Productcode (ERP)', required: true },
    { key: 'prodDesc',  label: 'Productomschrijving', required: true },
    { key: 'ean',       label: 'EAN', required: true },
    { key: 'qty',       label: 'QTY (PCS)', required: true },
    { key: 'gw',        label: 'G.W (KGS)', required: false },
    { key: 'cbm',       label: 'CBM', required: false },
    { key: 'batch',     label: 'Batch', required: false }
  ];

  const TEMPLATE_HEADERS = [
    'L', 'W', 'H',
    'Productcode', 'Omschrijving',
    'EAN', 'QTY', 'GW', 'CBM', 'Batch'
  ];
  const TEMPLATE_EXAMPLE_ROW = [
    '38', '55.5', '13',
    'ABC-123-XYZ', 'Voorbeeld product',
    '8712345678901', '120', '8.5', '0.095', 'B2025-01'
  ];

  function cmToPx(cm) {
    return cm * PX_PER_CM;
  }

  function fmt1(value) {
    if (!Number.isFinite(value)) return '';
    return value.toLocaleString('nl-NL', {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1
    });
  }

  function delay(ms) {
    return new Promise(res => setTimeout(res, ms));
  }

  function slugify(str) {
    return (str || '')
      .toString()
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '') || 'etiket';
  }

  function buildFileName(data, index) {
    const base = slugify(data.prodCode || data.prodDesc || 'etiket');
    if (typeof index === 'number') {
      return `${String(index + 1).padStart(3, '0')}-${base}.pdf`;
    }
    return `${base}.pdf`;
  }

  // ====== EXTERNE LIBS DYNAMISCH LADEN (jsPDF + html2canvas) ======

  let jsPdfPromise = null;
  let html2canvasPromise = null;

  function ensureJsPdfLoaded() {
    if (window.jspdf && window.jspdf.jsPDF) {
      return Promise.resolve(window.jspdf.jsPDF);
    }
    if (!jsPdfPromise) {
      jsPdfPromise = new Promise((resolve, reject) => {
        const script = DOC.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        script.onload = () => {
          if (window.jspdf && window.jspdf.jsPDF) {
            resolve(window.jspdf.jsPDF);
          } else {
            reject(new Error('jsPDF is niet beschikbaar na laden.'));
          }
        };
        script.onerror = () => reject(new Error('Kon jsPDF niet laden.'));
        DOC.head.appendChild(script);
      });
    }
    return jsPdfPromise;
  }

  function ensureHtml2canvasLoaded() {
    if (window.html2canvas) {
      return Promise.resolve(window.html2canvas);
    }
    if (!html2canvasPromise) {
      html2canvasPromise = new Promise((resolve, reject) => {
        const script = DOC.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        script.onload = () => {
          if (window.html2canvas) {
            resolve(window.html2canvas);
          } else {
            reject(new Error('html2canvas is niet beschikbaar na laden.'));
          }
        };
        script.onerror = () => reject(new Error('Kon html2canvas niet laden.'));
        DOC.head.appendChild(script);
      });
    }
    return html2canvasPromise;
  }

  // ====== DOM REFERENTIES ======

  const form = DOC.getElementById('labelForm');

  const boxLengthEl = DOC.getElementById('boxLength');
  const boxWidthEl  = DOC.getElementById('boxWidth');
  const boxHeightEl = DOC.getElementById('boxHeight');

  const prodCodeEl = DOC.getElementById('prodCode');
  const prodDescEl = DOC.getElementById('prodDesc');
  const eanEl      = DOC.getElementById('ean');
  const qtyEl      = DOC.getElementById('qty');
  const gwEl       = DOC.getElementById('gw');
  const cbmEl      = DOC.getElementById('cbm');
  const batchEl    = DOC.getElementById('batch');

  const btnGen = DOC.getElementById('btnGen');
  const btnPDF = DOC.getElementById('btnPDF');

  const controlInfo = DOC.getElementById('controlInfo');
  const labelsGrid  = DOC.getElementById('labelsGrid');

  // Batch DOM
  const btnPickFile   = DOC.getElementById('btnPickFile');
  const fileInput     = DOC.getElementById('fileInput');
  const dropzone      = DOC.getElementById('dropzone');
  const btnTemplateCsv  = DOC.getElementById('btnTemplateCsv');
  const btnTemplateXlsx = DOC.getElementById('btnTemplateXlsx');

  const mappingWrap = DOC.getElementById('mappingWrap');
  const mappingGrid = DOC.getElementById('mappingGrid');
  const previewWrap = DOC.getElementById('previewWrap');
  const tablePreview = DOC.getElementById('tablePreview');

  const normWrap = DOC.getElementById('normWrap');
  const optCommaDecimal = DOC.getElementById('optCommaDecimal');
  const optTrimSpaces   = DOC.getElementById('optTrimSpaces');

  const batchControls = DOC.getElementById('batchControls');
  const btnRunBatch   = DOC.getElementById('btnRunBatch');
  const btnAbortBatch = DOC.getElementById('btnAbortBatch');

  const progressWrap  = DOC.getElementById('progressWrap');
  const progressPhase = DOC.getElementById('progressPhase');
  const progressBar   = DOC.getElementById('progressBar');
  const progressLabel = DOC.getElementById('progressLabel');

  const logWrap = DOC.getElementById('logWrap');
  const logList = DOC.getElementById('logList');

  // ====== STATE ======

  const state = {
    batchHeaders: [],
    batchRows: [],
    batchMapping: {},
    batchOptions: { commaDecimal: false, trimSpaces: true },
    batchAbort: false
  };

  // ====== FORM & DATA ======

  function parseNumberInput(el) {
    const raw = (el.value || '').toString().trim().replace(',', '.');
    const n = parseFloat(raw);
    return Number.isFinite(n) ? n : NaN;
  }

  function collectFormData() {
    if (!form.reportValidity()) return null;

    const L = parseNumberInput(boxLengthEl);
    const W = parseNumberInput(boxWidthEl);
    const H = parseNumberInput(boxHeightEl);

    if (!(L > 0 && W > 0 && H > 0)) {
      alert('Vul geldige doosafmetingen in (L, W, H > 0).');
      return null;
    }

    return {
      L, W, H,
      prodCode: (prodCodeEl.value || '').trim(),
      prodDesc: (prodDescEl.value || '').trim(),
      ean:      (eanEl.value || '').trim(),
      qty:      (qtyEl.value || '').trim(),
      gw:       (gwEl.value || '').trim(),
      cbm:      (cbmEl.value || '').trim(),
      batch:    (batchEl.value || '').trim()
    };
  }

  // ====== LABEL-DIMS & CONTROL INFO ======

  function computeLabelDims(data) {
    const { L, W, H } = data;
    const labelHeightCm = H * SHRINK_FACTOR;
    const frontBackWidthCm = W * SHRINK_FACTOR;
    const sideWidthCm = L * SHRINK_FACTOR;

    return {
      labelHeightCm,
      frontBackWidthCm,
      sideWidthCm,
      topBoxHeightCm: labelHeightCm * 0.3,
      bottomBoxHeightCm: labelHeightCm * 0.7
    };
  }

  function updateControlInfoDims(data) {
    const d = computeLabelDims(data);

    controlInfo.innerHTML = `
      <h3>Berekende labelafmetingen</h3>
      <div class="control-grid-2x2">
        <div class="control-item">
          Front / Back: ${fmt1(d.frontBackWidthCm)} × ${fmt1(d.labelHeightCm)} cm
        </div>
        <div class="control-item">
          Zijkant links / rechts: ${fmt1(d.sideWidthCm)} × ${fmt1(d.labelHeightCm)} cm
        </div>
        <div class="control-item">
          Top-Box (30%): ca. ${fmt1(d.topBoxHeightCm)} cm hoog
        </div>
        <div class="control-item">
          Bottom-Box (70%): ca. ${fmt1(d.bottomBoxHeightCm)} cm hoog
        </div>
      </div>
    `;
  }

  function updateControlInfoFont(fsPx, softwrap) {
    if (!controlInfo) return;

    const old = controlInfo.querySelector('[data-font-footnote="1"]');
    if (old) old.remove();

    const foot = DOC.createElement('div');
    foot.dataset.fontFootnote = '1';
    foot.style.marginTop = '6px';
    foot.style.fontSize = '.8rem';
    foot.style.opacity = '0.85';

    const erpPx = fsPx * CODE_MULT;
    foot.textContent = `Basis-font: ${fsPx.toFixed(1)} px, ERP ≈ ${erpPx.toFixed(1)} px, detailmodus: ${softwrap ? 'soft wrap (waarden mogen afbreken)' : 'no wrap (waarden blijven op 1 regel)'}.`;

    controlInfo.appendChild(foot);
  }

  // ====== LABEL DOM-STRUCTUUR ======

  function createLabelElement(widthCm, heightCm, labelTitle, data) {
    const wrap = DOC.createElement('div');
    wrap.className = 'label-wrap';

    const frame = DOC.createElement('div');
    frame.className = 'label';
    // fysieke maat via cm (browser rekent px)
    frame.style.width = `${widthCm}cm`;
    frame.style.height = `${heightCm}cm`;

    const inner = DOC.createElement('div');
    inner.className = 'label-inner nowrap-mode'; // start in nowrap-mode

    // Top-box: ERP + productomschrijving
    const topBox = DOC.createElement('div');
    topBox.className = 'top-box';

    const head = DOC.createElement('div');
    head.className = 'label-head';

    const codeBox = DOC.createElement('div');
    codeBox.className = 'code-box';
    codeBox.textContent = data.prodCode || '';

    const prodDesc = DOC.createElement('div');
    prodDesc.className = 'product-desc';
    prodDesc.textContent = data.prodDesc || '';

    head.appendChild(codeBox);
    head.appendChild(prodDesc);
    topBox.appendChild(head);

    // Bottom-box: Detail-box met specs-grid
    const bottomBox = DOC.createElement('div');
    bottomBox.className = 'bottom-box';

    const detailBox = DOC.createElement('div');
    detailBox.className = 'detail-box';

    const specsGrid = DOC.createElement('div');
    specsGrid.className = 'specs-grid';

    function addRow(label, valueNodeOrStr) {
      const lab = DOC.createElement('div');
      lab.className = 'lab';
      lab.textContent = label;

      const val = DOC.createElement('div');
      val.className = 'val';

      if (valueNodeOrStr instanceof Node) {
        val.appendChild(valueNodeOrStr);
      } else {
        val.textContent = valueNodeOrStr || '';
      }

      specsGrid.appendChild(lab);
      specsGrid.appendChild(val);
    }

    addRow('EAN:',   data.ean || '');
    addRow('QTY:',   data.qty || '');
    addRow('G.W:',   data.gw  || '');
    addRow('CBM:',   data.cbm || '');
    addRow('Batch:', data.batch || '');

    // C/N speciale lijn (breedte ~ "IN CHINA")
    const cnLine = DOC.createElement('span');
    cnLine.className = 'cn-line';
    addRow('C/N:', cnLine);

    // "MADE IN: CHINA"
    addRow('MADE IN:', 'CHINA');

    detailBox.appendChild(specsGrid);
    bottomBox.appendChild(detailBox);

    inner.appendChild(topBox);
    inner.appendChild(bottomBox);

    frame.appendChild(inner);

    const num = DOC.createElement('div');
    num.className = 'label-num';
    num.textContent = labelTitle;

    wrap.appendChild(frame);
    wrap.appendChild(num);

    return { wrap, inner };
  }

  function buildLabelsInContainer(data, gridEl) {
    const dims = computeLabelDims(data);
    const cfgs = [
      { title: '1. Front', widthCm: dims.frontBackWidthCm, heightCm: dims.labelHeightCm },
      { title: '2. Back',  widthCm: dims.frontBackWidthCm, heightCm: dims.labelHeightCm },
      { title: '3. Side L', widthCm: dims.sideWidthCm, heightCm: dims.labelHeightCm },
      { title: '4. Side R', widthCm: dims.sideWidthCm, heightCm: dims.labelHeightCm }
    ];

    const innerEls = [];
    cfgs.forEach(cfg => {
      const { wrap, inner } = createLabelElement(cfg.widthCm, cfg.heightCm, cfg.title, data);
      gridEl.appendChild(wrap);
      innerEls.push(inner);
    });

    return innerEls;
  }

  // ====== FONT-SCALING / FIT ======

  function fitsLabel(innerEl, guardFrac = 0.02) {
    const w = innerEl.clientWidth;
    const h = innerEl.clientHeight;
    if (!w || !h) return true;

    const guardX = w * guardFrac;
    const guardY = h * guardFrac;

    const bodyFits =
      innerEl.scrollWidth <= (w - guardX + 0.5) &&
      innerEl.scrollHeight <= (h - guardY + 0.5);

    const topBox = innerEl.querySelector('.top-box');
    const detailBox = innerEl.querySelector('.detail-box');

    let topFits = true;
    let detailFits = true;

    if (topBox) {
      topFits = topBox.scrollHeight <= (topBox.clientHeight + 0.5);
    }
    if (detailBox) {
      detailFits = detailBox.scrollHeight <= (detailBox.clientHeight + 0.5);
    }

    return bodyFits && topFits && detailFits;
  }

  function labelsFit(innerEls) {
    return innerEls.every(el => fitsLabel(el));
  }

  function applyFontSizesAll(innerEls, fsPx) {
    innerEls.forEach(innerEl => {
      innerEl.style.setProperty('--fs', `${fsPx}px`);
      const codeEl = innerEl.querySelector('.code-box');
      if (codeEl) {
        codeEl.style.fontSize = `${fsPx * CODE_MULT}px`;
      }
    });
  }

  function fitFontsForLabels(innerEls) {
    if (!innerEls || !innerEls.length) return;

    innerEls.forEach(el => {
      el.classList.add('nowrap-mode');
      el.classList.remove('softwrap-mode');
    });

    let minH = Infinity;
    innerEls.forEach(el => {
      if (el.clientHeight > 0) {
        minH = Math.min(minH, el.clientHeight);
      }
    });
    if (!Number.isFinite(minH) || minH <= 0) return;

    // hi-limit schaalbaar aan labelhoogte
    const maxFs = Math.min(
      MAX_FS_ABS,
      Math.max(12, Math.floor(minH / 3))
    );
    const minFs = MIN_FS;

    let chosen = null;

    // 1) probeer zonder afbreken in waarden (nowrap-mode)
    for (let fs = maxFs; fs >= minFs; fs--) {
      applyFontSizesAll(innerEls, fs);
      if (labelsFit(innerEls)) {
        chosen = { fs, softwrap: false };
        break;
      }
    }

    // 2) indien nodig: sta afbreken van waarden toe (softwrap-mode)
    if (!chosen) {
      innerEls.forEach(el => {
        el.classList.remove('nowrap-mode');
        el.classList.add('softwrap-mode');
      });

      for (let fs = maxFs; fs >= minFs; fs--) {
        applyFontSizesAll(innerEls, fs);
        if (labelsFit(innerEls)) {
          chosen = { fs, softwrap: true };
          break;
        }
      }
    }

    // 3) nood-fallback
    if (!chosen) {
      applyFontSizesAll(innerEls, minFs);
      innerEls.forEach(el => {
        el.classList.remove('nowrap-mode');
        el.classList.add('softwrap-mode');
      });
      chosen = { fs: minFs, softwrap: true };
    }

    updateControlInfoFont(chosen.fs, chosen.softwrap);
  }

  // ====== PREVIEW RENDERING ======

  function renderPreview() {
    const data = collectFormData();
    if (!data) return;

    labelsGrid.innerHTML = '';

    updateControlInfoDims(data);

    const innerEls = buildLabelsInContainer(data, labelsGrid);

    // na DOM-insertie font-fitting uitvoeren
    fitFontsForLabels(innerEls);
  }

  // ====== SINGLE PDF ======

  async function generateSinglePdf() {
    const data = collectFormData();
    if (!data) return;

    // Preview eerst updaten zodat PDF hetzelfde gebruikt
    renderPreview();

    if (!labelsGrid || !labelsGrid.firstElementChild) {
      alert('Er is geen voorbeeld om naar PDF te schrijven.');
      return;
    }

    try {
      const [jsPDF, html2canvas] = await Promise.all([
        ensureJsPdfLoaded(),
        ensureHtml2canvasLoaded()
      ]);

      await delay(30); // kleine pauze voor layout

      const canvas = await html2canvas(labelsGrid, {
        scale: 2,
        backgroundColor: '#ffffff'
      });

      const imgData = canvas.toDataURL('image/png');

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
      });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();

      const imgRatio = canvas.width / canvas.height;
      let renderW = pageW - 20; // 10mm marge links/rechts
      let renderH = renderW / imgRatio;

      if (renderH > pageH - 20) {
        renderH = pageH - 20;
        renderW = renderH * imgRatio;
      }

      const x = (pageW - renderW) / 2;
      const y = (pageH - renderH) / 2;

      doc.addImage(imgData, 'PNG', x, y, renderW, renderH);

      const fileName = buildFileName(data);
      doc.save(fileName);
    } catch (err) {
      console.error(err);
      alert('Er ging iets mis bij het genereren van de PDF: ' + (err.message || err));
    }
  }

  // ====== BATCH: TEMPLATES ======

  function downloadTemplateCsv() {
    const lines = [
      TEMPLATE_HEADERS.join(';'),
      TEMPLATE_EXAMPLE_ROW.join(';')
    ];
    const blob = new Blob([lines.join('\r\n')], {
      type: 'text/csv;charset=utf-8;'
    });
    const a = DOC.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'etiketten-template.csv';
    DOC.body.appendChild(a);
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    a.remove();
  }

  function downloadTemplateXlsx() {
    const wb = XLSX.utils.book_new();
    const data = [TEMPLATE_HEADERS, TEMPLATE_EXAMPLE_ROW];
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Etiketten');

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    const a = DOC.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'etiketten-template.xlsx';
    DOC.body.appendChild(a);
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    a.remove();
  }

  // ====== BATCH: FILE PARSING ======

  function handleFileList(fileList) {
    if (!fileList || !fileList.length) return;
    const file = fileList[0];
    if (!file) return;

    const name = (file.name || '').toLowerCase();
    const isCsv = name.endsWith('.csv');

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        if (isCsv) {
          parseCsvString(String(e.target.result || ''));
        } else {
          parseXlsxArrayBuffer(e.target.result);
        }
        afterBatchDataParsed();
      } catch (err) {
        console.error(err);
        alert('Kon het bestand niet verwerken: ' + (err.message || err));
      }
    };

    if (isCsv) {
      reader.readAsText(file, 'utf-8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  }

  function parseCsvString(text) {
    const lines = text.split(/\r\n|\n|\r/).filter(l => l.trim() !== '');
    if (!lines.length) throw new Error('Leeg CSV-bestand.');

    // simpele delimiter-detectie
    const first = lines[0];
    const delim = first.includes(';') ? ';' : ',';

    const rows = lines.map(line =>
      line.split(delim).map(cell => cell.replace(/^"|"$/g, ''))
    );

    state.batchHeaders = rows[0].map(h => String(h || '').trim());
    state.batchRows = rows.slice(1).filter(row =>
      row.some(cell => String(cell || '').trim() !== '')
    );
  }

  function parseXlsxArrayBuffer(buffer) {
    const data = new Uint8Array(buffer);
    const wb = XLSX.read(data, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

    if (!rows.length) throw new Error('Leeg werkblad.');

    state.batchHeaders = rows[0].map(h => String(h || '').trim());
    state.batchRows = rows.slice(1).filter(row =>
      row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')
    );
  }

  // ====== BATCH: UI-OPBOUW ======

  function clearBatchUi() {
    mappingGrid.innerHTML = '';
    tablePreview.innerHTML = '';
    state.batchMapping = {};
    state.batchAbort = false;

    mappingWrap.classList.add('hidden');
    previewWrap.classList.add('hidden');
    normWrap.classList.add('hidden');
    batchControls.classList.add('hidden');
    progressWrap.classList.add('hidden');
    logWrap.classList.add('hidden');
    logList.innerHTML = '';
  }

  function buildMappingUi() {
    mappingGrid.innerHTML = '';

    const headers = state.batchHeaders;
    if (!headers.length) return;

    function guessIndex(fieldKey, fieldLabel) {
      const labelLower = fieldLabel.toLowerCase();
      const keyLower = fieldKey.toLowerCase();

      let bestIdx = -1;

      headers.forEach((h, idx) => {
        const hl = String(h || '').toLowerCase();
        if (!hl) return;

        if (keyLower.includes('length') || keyLower.includes('boxlength')) {
          if (/(lengte|length|l\b)/.test(hl)) bestIdx = idx;
        } else if (keyLower.includes('width') || keyLower.includes('boxwidth')) {
          if (/(breedte|width|w\b)/.test(hl)) bestIdx = idx;
        } else if (keyLower.includes('height') || keyLower.includes('boxheight')) {
          if (/(hoogte|height|h\b)/.test(hl)) bestIdx = idx;
        } else if (keyLower === 'prodcode') {
          if (/(productcode|erp|code)/.test(hl)) bestIdx = idx;
        } else if (keyLower === 'proddesc') {
          if (/(omschrijving|description|desc)/.test(hl)) bestIdx = idx;
        } else {
          const labelTokens = labelLower.split(/\s+/);
          const match = labelTokens.some(t => t && hl.includes(t));
          if (match && bestIdx === -1) bestIdx = idx;
        }
      });

      return bestIdx;
    }

    TARGET_FIELDS.forEach(field => {
      const row = DOC.createElement('div');
      row.className = 'map-row';

      const lab = DOC.createElement('label');
      lab.textContent = field.label + (field.required ? ' *' : '');

      const select = DOC.createElement('select');
      select.id = `map_${field.key}`;

      const optEmpty = DOC.createElement('option');
      optEmpty.value = '';
      optEmpty.textContent = '— Kies kolom —';
      select.appendChild(optEmpty);

      headers.forEach((h, idx) => {
        const opt = DOC.createElement('option');
        opt.value = String(idx);
        opt.textContent = h || `(kolom ${idx + 1})`;
        select.appendChild(opt);
      });

      const guessed = guessIndex(field.key, field.label);
      if (guessed >= 0) {
        select.value = String(guessed);
        state.batchMapping[field.key] = guessed;
      }

      select.addEventListener('change', () => {
        const val = select.value;
        state.batchMapping[field.key] =
          val === '' ? null : parseInt(val, 10);
      });

      row.appendChild(lab);
      row.appendChild(select);

      mappingGrid.appendChild(row);
    });

    mappingWrap.classList.remove('hidden');
  }

  function buildPreviewTable() {
    const headers = state.batchHeaders;
    const rows = state.batchRows;
    if (!headers.length) return;

    const maxRows = Math.min(8, rows.length);

    const table = DOC.createElement('table');
    const thead = DOC.createElement('thead');
    const trHead = DOC.createElement('tr');
    headers.forEach(h => {
      const th = DOC.createElement('th');
      th.textContent = h;
      trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    const tbody = DOC.createElement('tbody');
    for (let i = 0; i < maxRows; i++) {
      const tr = DOC.createElement('tr');
      const r = rows[i] || [];
      headers.forEach((_, colIdx) => {
        const td = DOC.createElement('td');
        td.textContent = String(
          colIdx < r.length ? (r[colIdx] ?? '') : ''
        );
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }

    table.appendChild(thead);
    table.appendChild(tbody);

    tablePreview.innerHTML = '';
    tablePreview.appendChild(table);
    previewWrap.classList.remove('hidden');
  }

  function afterBatchDataParsed() {
    clearBatchUi();

    if (!state.batchHeaders.length || !state.batchRows.length) {
      alert('Geen bruikbare rijen gevonden in het bestand.');
      return;
    }

    buildMappingUi();
    buildPreviewTable();

    normWrap.classList.remove('hidden');
    batchControls.classList.remove('hidden');
  }

  // ====== BATCH: DATA MAPPING & NORMALISATIE ======

  function getCellValue(row, idx) {
    if (idx == null || idx < 0 || idx >= row.length) return '';
    const raw = row[idx];
    let s = raw == null ? '' : String(raw);
    if (state.batchOptions.trimSpaces) {
      s = s.trim();
    }
    return s;
  }

  function parseNumberValue(str) {
    if (!str) return NaN;
    let s = String(str);
    if (state.batchOptions.commaDecimal) {
      s = s.replace(',', '.');
    }
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : NaN;
  }

  function mapRowToData(row) {
    const m = state.batchMapping;

    function v(key) {
      const idx = m[key];
      return getCellValue(row, idx);
    }
    function n(key) {
      return parseNumberValue(v(key));
    }

    const L = n('boxLength');
    const W = n('boxWidth');
    const H = n('boxHeight');

    return {
      L, W, H,
      prodCode: v('prodCode'),
      prodDesc: v('prodDesc'),
      ean:      v('ean'),
      qty:      v('qty'),
      gw:       v('gw'),
      cbm:      v('cbm'),
      batch:    v('batch')
    };
  }

  // ====== BATCH: LOG & PROGRESS ======

  function appendLog(msg, type) {
    const div = DOC.createElement('div');
    if (type) div.classList.add(type);
    div.textContent = msg;
    logList.appendChild(div);
    logList.scrollTop = logList.scrollHeight;
    logWrap.classList.remove('hidden');
  }

  function updateProgress(done, total, phaseLabel) {
    progressWrap.classList.remove('hidden');
    progressPhase.textContent = phaseLabel || 'Bezig...';
    const pct = total ? Math.round((done / total) * 100) : 0;
    progressBar.style.width = pct + '%';
    progressLabel.textContent = `${done} / ${total}`;
  }

  // ====== BATCH: RUN ======

  async function runBatch() {
    const rows = state.batchRows;
    if (!rows.length) {
      alert('Geen data geladen voor batch.');
      return;
    }

    const missing = TARGET_FIELDS.filter(f =>
      f.required &&
      (state.batchMapping[f.key] == null || state.batchMapping[f.key] === '')
    );
    if (missing.length) {
      alert(
        'Koppel eerst de kolommen voor:\n- ' +
        missing.map(m => m.label).join('\n- ')
      );
      return;
    }

    state.batchAbort = false;
    btnAbortBatch.disabled = false;
    btnRunBatch.disabled = true;

    progressWrap.classList.remove('hidden');
    progressBar.style.width = '0%';
    progressLabel.textContent = `0 / ${rows.length}`;
    progressPhase.textContent = 'Batch starten…';

    logWrap.classList.remove('hidden');
    logList.innerHTML = '';

    try {
      const [jsPDF, html2canvas] = await Promise.all([
        ensureJsPdfLoaded(),
        ensureHtml2canvasLoaded()
      ]);

      const zip = new JSZip();

      const tmpContainer = DOC.createElement('div');
      tmpContainer.style.position = 'fixed';
      tmpContainer.style.left = '-9999px';
      tmpContainer.style.top = '-9999px';
      tmpContainer.style.width = '800px';
      tmpContainer.style.background = '#ffffff';
      tmpContainer.style.padding = '0';
      DOC.body.appendChild(tmpContainer);

      for (let i = 0; i < rows.length; i++) {
        if (state.batchAbort) break;

        const rowIdxDisplay = i + 2; // rekening houden met header
        const row = rows[i];

        try {
          const data = mapRowToData(row);
          if (!(data.L > 0 && data.W > 0 && data.H > 0)) {
            appendLog(
              `Rij ${rowIdxDisplay}: overgeslagen (ongeldige doosafmetingen).`,
              'err'
            );
            continue;
          }

          tmpContainer.innerHTML = '';
          const localGrid = DOC.createElement('div');
          localGrid.className = 'labels-grid';
          tmpContainer.appendChild(localGrid);

          const innerEls = buildLabelsInContainer(data, localGrid);
          fitFontsForLabels(innerEls);

          await delay(20);

          const canvas = await html2canvas(localGrid, {
            scale: 2,
            backgroundColor: '#ffffff'
          });

          const imgData = canvas.toDataURL('image/png');

          const doc = new jsPDF({
            orientation: 'landscape',
            unit: 'mm',
            format: 'a4'
          });

          const pageW = doc.internal.pageSize.getWidth();
          const pageH = doc.internal.pageSize.getHeight();
          const imgRatio = canvas.width / canvas.height;

          let renderW = pageW - 20;
          let renderH = renderW / imgRatio;
          if (renderH > pageH - 20) {
            renderH = pageH - 20;
            renderW = renderH * imgRatio;
          }
          const x = (pageW - renderW) / 2;
          const y = (pageH - renderH) / 2;

          doc.addImage(imgData, 'PNG', x, y, renderW, renderH);
          const pdfBlob = doc.output('blob');

          const fileName = buildFileName(data, i);
          zip.file(fileName, pdfBlob);
          appendLog(`Rij ${rowIdxDisplay}: ok (${fileName})`, 'ok');
        } catch (rowErr) {
          console.error(rowErr);
          appendLog(
            `Rij ${rowIdxDisplay}: fout (${rowErr.message || rowErr})`,
            'err'
          );
        }

        updateProgress(i + 1, rows.length, 'PDF’s genereren…');
        await delay(20);
      }

      DOC.body.removeChild(tmpContainer);

      if (!state.batchAbort) {
        progressPhase.textContent = 'ZIP-bestand maken…';
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const a = DOC.createElement('a');
        a.href = URL.createObjectURL(zipBlob);
        a.download = 'etiketten-batch.zip';
        DOC.body.appendChild(a);
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 4000);
        a.remove();
        appendLog('Batch voltooid – ZIP gedownload.', 'ok');
      } else {
        appendLog('Batch afgebroken door gebruiker.', 'err');
      }
    } catch (err) {
      console.error(err);
      alert('Fout tijdens batchverwerking: ' + (err.message || err));
    } finally {
      btnAbortBatch.disabled = true;
      btnRunBatch.disabled = false;
    }
  }

  // ====== BATCH: EVENTS ======

  function setupBatchEvents() {
    if (btnPickFile) {
      btnPickFile.addEventListener('click', () => {
        fileInput && fileInput.click();
      });
    }

    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        handleFileList(e.target.files);
      });
    }

    if (dropzone) {
      ['dragenter', 'dragover'].forEach(ev => {
        dropzone.addEventListener(ev, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropzone.classList.add('dragover');
        });
      });

      ['dragleave', 'drop'].forEach(ev => {
        dropzone.addEventListener(ev, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropzone.classList.remove('dragover');
        });
      });

      dropzone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        if (dt && dt.files && dt.files.length) {
          handleFileList(dt.files);
        }
      });
    }

    if (btnTemplateCsv) {
      btnTemplateCsv.addEventListener('click', downloadTemplateCsv);
    }
    if (btnTemplateXlsx) {
      btnTemplateXlsx.addEventListener('click', downloadTemplateXlsx);
    }

    if (optCommaDecimal) {
      optCommaDecimal.addEventListener('change', () => {
        state.batchOptions.commaDecimal = !!optCommaDecimal.checked;
      });
    }
    if (optTrimSpaces) {
      optTrimSpaces.checked = true;
      state.batchOptions.trimSpaces = true;
      optTrimSpaces.addEventListener('change', () => {
        state.batchOptions.trimSpaces = !!optTrimSpaces.checked;
      });
    }

    if (btnRunBatch) {
      btnRunBatch.addEventListener('click', () => {
        runBatch();
      });
    }

    if (btnAbortBatch) {
      btnAbortBatch.addEventListener('click', () => {
        state.batchAbort = true;
        btnAbortBatch.disabled = true;
      });
    }
  }

  // ====== INIT ======

  function init() {
    if (btnGen) {
      btnGen.addEventListener('click', renderPreview);
    }

    if (btnPDF) {
      btnPDF.addEventListener('click', generateSinglePdf);
    }

    // Enter in form → preview
    if (form) {
      form.addEventListener('submit', (e) => {
        e.preventDefault();
        renderPreview();
      });
    }

    setupBatchEvents();
  }

  if (DOC.readyState === 'loading') {
    DOC.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
