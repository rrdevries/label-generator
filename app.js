(() => {
  /* ====== CONSTANTS ====== */
  const PX_PER_CM = 37.7952755906; // 96dpi
  const PREVIEW_GAP_CM_NUM = 0;    // geen ruimte tussen etiketten in preview
  const PDF_MARGIN_CM = 0.5;       // witmarge rondom in PDF
  const LABEL_PADDING_CM = 0.5;    // binnenmarge in label (cm)

  // Typografie / fit
  const WRAP_THRESHOLD_PX = 10;    // onder 10px pas zachte afbreking aanzetten
  const MIN_FS_PX = 6;             // noodrem bij heel kleine labels
  const CODE_MULT = 1.6;           // productcode ≈ 1.6 × body

  // PDF rand (alleen voor capture-visual; fysieke rand komt uit DOM-stijl)
  const BORDER_PX = 2;

  /* ====== DOM HOOKS ====== */
  const labelsGrid  = document.getElementById('labelsGrid');
  const controlInfo = document.getElementById('controlInfo');
  const canvasEl    = document.getElementById('canvas');
  const btnGen      = document.getElementById('btnGenerate');
  const btnPDF      = document.getElementById('btnPDF');

  // Batch DOM (optioneel aanwezig)
  const dropzone    = document.getElementById('dropzone');
  const fileInput   = document.getElementById('fileInput');
  const btnPickFile = document.getElementById('btnPickFile');
  const btnTemplateCsv  = document.getElementById('btnTemplateCsv');
  const btnTemplateXlsx = document.getElementById('btnTemplateXlsx');
  const mappingWrap = document.getElementById('mappingWrap');
  const mappingGrid = document.getElementById('mappingGrid');
  const previewWrap = document.getElementById('previewWrap');
  const tablePreview= document.getElementById('tablePreview');
  const normWrap    = document.getElementById('normWrap');
  const chkComma    = document.getElementById('optCommaDecimal');
  const chkTrim     = document.getElementById('optTrimSpaces');
  const batchControls = document.getElementById('batchControls');
  const btnRunBatch   = document.getElementById('btnRunBatch');
  const btnAbortBatch = document.getElementById('btnAbortBatch');
  const progressWrap  = document.getElementById('progressWrap');
  const progressBar   = document.getElementById('progressBar');
  const progressLabel = document.getElementById('progressLabel');
  const progressPhase = document.getElementById('progressPhase');
  const logWrap   = document.getElementById('logWrap');
  const logList   = document.getElementById('logList');

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
      if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.substring(2), v);
      else node.setAttribute(k, v);
    });
    children.flat().forEach(c => node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c));
    return node;
  };
  const line = (lab, val) => [el('div',{class:'lab'}, lab), el('div',{class:'val'}, val)];
  const pad2 = n => String(n).padStart(2,'0');
  const ts = (d=new Date()) => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())} ${pad2(d.getHours())}.${pad2(d.getMinutes())}.${pad2(d.getSeconds())}`;

  // Compat-wrapper: project kan readValues óf readValuesSingle hebben
  function getFormValues(){
    if (typeof readValues === 'function')       return readValues();
    if (typeof readValuesSingle === 'function') return readValuesSingle();
    throw new Error('readValues / readValuesSingle niet gevonden');
  }

  // 1 frame wachten zodat layout/afmetingen kloppen
  function nextFrame(){ return new Promise(r => requestAnimationFrame(r)); }

  // Fit alle label-inhouden in container
  function fitAllIn(container){
    container.querySelectorAll('.label-inner').forEach(inner => {
      inner.classList.add('nowrap-mode');
      inner.classList.remove('softwrap-mode');
      fitContentToBoxConditional(inner);
    });
  }

  // Extra robuust: twee fit-rondes met tussentijds frame
  async function mountThenFit(container){
    await nextFrame();
    fitAllIn(container);
    await nextFrame();
    fitAllIn(container);
  }

  /* ====== ENKELVOUDIG INLEZEN ====== */
  function readValuesSingle(){
    const g = id => document.getElementById(id).value.trim();
    const L = parseFloat(g('boxLength'));
    const W = parseFloat(g('boxWidth'));
    const H = parseFloat(g('boxHeight'));

    const vals = {
      L,W,H,
      code:g('prodCode'),
      desc:g('prodDesc'),
      ean:g('ean'),
      qty:String(Math.max(0, Math.floor(Number(g('qty'))||0))),
      gw:(v=>isFinite(+v)?(+v).toFixed(2):v)(g('gw')),
      cbm:g('cbm'),
      batch:g('batch') // verplicht
    };

    if ([L,W,H].some(v=>!isFinite(v)||v<=0) ||
        !vals.code || !vals.desc || !vals.ean || !vals.qty || !vals.gw || !vals.cbm || !vals.batch){
      throw new Error('Controleer de verplichte velden (incl. Batch).');
    }
    return vals;
  }

  /* ====== LABELMATEN & PREVIEW SCALE ====== */
  function computeLabelSizes({ L, W, H }){
    // Etiket = 80% van respectievelijk L×H (1&2) en W×H (3&4) → 10% marge rondom
    const scale=0.8;
    const fb={ w:L*scale, h:H*scale };
    const sd={ w:W*scale, h:H*scale };
    const cnText = 'C/N: ___________________';
    return [
      { idx:1, kind:'front/back', ...fb, under:'Made in China' },
      { idx:2, kind:'front/back', ...fb, under: cnText },
      { idx:3, kind:'side',       ...sd, under:'Made in China' },
      { idx:4, kind:'side',       ...sd, under: cnText }
    ];
  }

  function updateControlInfo(sizes){
    const n2=x=>(Math.round(x*100)/100).toFixed(2);
    const [s1,s2,s3,s4]=sizes;
    controlInfo.innerHTML = `
      <h3>Berekende labelafmetingen (werkelijke cm)</h3>
      <div class="control-grid-2x2">
        <div class="control-item">Etiket 1 (${s1.kind}): ${n2(s1.w)} × ${n2(s1.h)} cm</div>
        <div class="control-item">Etiket 3 (${s3.kind}): ${n2(s3.w)} × ${n2(s3.h)} cm</div>
        <div class="control-item">Etiket 2 (${s2.kind}): ${n2(s2.w)} × ${n2(s2.h)} cm</div>
        <div class="control-item">Etiket 4 (${s4.kind}): ${n2(s4.w)} × ${n2(s4.h)} cm</div>
      </div>`;
  }

  function computePreviewScale(sizes){
    const gapPx = PREVIEW_GAP_CM_NUM * PX_PER_CM;
    const w1 = sizes[0].w * PX_PER_CM, w3 = sizes[2].w * PX_PER_CM;
    const requiredW = Math.max(w1 + gapPx + w1, w3 + gapPx + w3);
    const cs = getComputedStyle(canvasEl);
    const innerW = canvasEl.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    return Math.min(innerW/requiredW, 1);
  }

  /* ====== FONT-FIT ====== */
  function fitsWithGuard(innerEl, guardX, guardY){
    return (
      innerEl.scrollWidth  <= (innerEl.clientWidth  - guardX) &&
      innerEl.scrollHeight <= (innerEl.clientHeight - guardY)
    );
  }

  function applyFontSizes(innerEl, fsPx){
    // basis (body)
    innerEl.style.setProperty('--fs', fsPx + 'px');
    // productcode schaalt mee (geen cap)
    const codeEl = innerEl.querySelector('.code-box');
    if (codeEl){
      codeEl.style.fontSize = (fsPx * CODE_MULT) + 'px';
    }
  }

  function searchFontSize(innerEl, minFs, hi, guardX, guardY){
    // probeer eerst omhoog te groeien
    applyFontSizes(innerEl, hi);
    if (fitsWithGuard(innerEl, guardX, guardY)){
      let grow = hi;
      for (let i=0;i<36;i++){
        const next = grow * 1.06;
        applyFontSizes(innerEl, next);
        if (!fitsWithGuard(innerEl, guardX, guardY)){ applyFontSizes(innerEl, grow); return grow; }
        grow = next;
      }
      return grow;
    }
    // anders binair omlaag zoeken
    let lo = minFs, best = lo;
    while (hi - lo > 0.5){
      const mid = (lo + hi)/2;
      applyFontSizes(innerEl, mid);
      if (fitsWithGuard(innerEl, guardX, guardY)){ best = mid; lo = mid; } else { hi = mid; }
    }
    applyFontSizes(innerEl, best);
    return best;
  }

  // eerst no-wrap (>=10px), dan indien nodig soft-wrap (>=6px)
  function fitContentToBoxConditional(innerEl){
    const box = innerEl.getBoundingClientRect();
    const guardX = Math.max(6, box.width  * 0.02);
    const guardY = Math.max(6, box.height * 0.02);

    const hi = Math.max(16, Math.min(box.height * 0.22, box.width * 0.18)); // geen hard cap

    innerEl.classList.add('nowrap-mode');
    innerEl.classList.remove('softwrap-mode');
    let fs = searchFontSize(innerEl, WRAP_THRESHOLD_PX, hi, guardX, guardY);
    if (fs >= WRAP_THRESHOLD_PX) return;

    innerEl.classList.remove('nowrap-mode');
    innerEl.classList.add('softwrap-mode');
    searchFontSize(innerEl, MIN_FS_PX, hi, guardX, guardY);
  }

  /* ====== UI OPBOUW ====== */
  function buildLeftBlock(values, size){
    const block = el('div', { class:'label-leftblock' });
    const grid  = el('div', { class:'specs-grid' });
    [
      ...line('EAN:', values.ean),
      ...line('QTY:', `${values.qty} PCS`),
      ...line('G.W:', `${values.gw} KGS`),
      ...line('CBM:', values.cbm)
    ].forEach(n => grid.append(n));
    block.append(grid);
    block.append(el('div',{class:'line'}, `Batch: ${values.batch}`));
    block.append(el('div',{class:'line'}, size.under));
    return block;
  }

  function createLabelEl(size, values, previewScale){
    const widthPx  = Math.round(size.w * PX_PER_CM * previewScale);
    const heightPx = Math.round(size.h * PX_PER_CM * previewScale);

    const wrap  = el('div', { class:'label-wrap' });
    const label = el('div', { class:'label', style:{ width:widthPx+'px', height:heightPx+'px' }});
    label.dataset.idx = String(size.idx);

    const inner = el('div', { class:'label-inner nowrap-mode' });
    inner.style.padding = (LABEL_PADDING_CM * PX_PER_CM * previewScale) + 'px';

    const head = el('div', { class:'label-head' },
      el('div', { class:'code-box line' }, values.code),
      el('div', { class:'line' }, values.desc)
    );

    inner.append(head, el('div',{class:'block-spacer'}), buildLeftBlock(values, size));
    label.append(inner);
    wrap.append(label, el('div',{class:'label-num'}, `Etiket ${size.idx}`));

    // Alleen startwaarde — echte fit na mount
    applyFontSizes(inner, Math.max(10, Math.floor(heightPx * 0.06)));

    return wrap;
  }

  /* ====== PREVIEW PIPELINE (single & batch) ====== */
  async function renderPreviewFor(vals){
    const sizes = computeLabelSizes(vals);
    const scale = computePreviewScale(sizes);
    currentPreviewScale = scale;

    updateControlInfo(sizes);
    labelsGrid.style.gap = '0'; // géén witruimte tussen etiketten
    labelsGrid.innerHTML = '';

    // Volgorde: 1 & 3 boven, 2 & 4 onder
    [0,2,1,3].forEach(i => labelsGrid.appendChild(createLabelEl(sizes[i], vals, scale)));

    // Fit ná mount (twee frames voor robuustheid)
    await mountThenFit(labelsGrid);

    return { sizes, scale };
  }

  async function renderSingle(){
    const vals = getFormValues();
    await renderPreviewFor(vals);
  }

  /* ====== jsPDF / html2canvas ====== */
  function loadJsPDF(){ return new Promise((res,rej)=>{ if (window.jspdf?.jsPDF) return res(window.jspdf.jsPDF);
    const s=document.createElement('script'); s.src='https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
    s.onload=()=>res(window.jspdf.jsPDF); s.onerror=()=>rej(new Error('Kon jsPDF niet laden.')); document.head.appendChild(s); }); }
  function loadHtml2Canvas(){ return new Promise((res,rej)=>{ if (window.html2canvas) return res(window.html2canvas);
    const s=document.createElement('script'); s.src='https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js';
    s.onload=()=>res(window.html2canvas); s.onerror=()=>rej(new Error('Kon html2canvas niet laden.')); document.head.appendChild(s); }); }

  async function capturePreviewLabelToImage(h2c, idx, isFirst){
    const src = document.querySelector(`.label[data-idx="${idx}"]`);
    if (!src) throw new Error('Label niet gevonden voor capture.');

    const clone = src.cloneNode(true);
    // randen zó zetten dat het visueel klopt; liever te veel dan te weinig (geen missing edge)
    clone.style.borderTop    = `${BORDER_PX}px solid #000`;
    clone.style.borderRight  = `${BORDER_PX}px solid #000`;
    clone.style.borderBottom = `${BORDER_PX}px solid #000`;
    clone.style.borderLeft   = `${BORDER_PX}px solid #000`;

    const wrap = document.createElement('div');
    wrap.style.position='fixed'; wrap.style.left='-10000px'; wrap.style.top='0'; wrap.style.background='#fff';
    document.body.appendChild(wrap);
    wrap.appendChild(clone);

    // scherpteschaal — tast fysieke cm in PDF niet aan
    const capScale = Math.max(2, window.devicePixelRatio || 1, 1/currentPreviewScale);
    const canvas = await h2c(clone, { backgroundColor:'#fff', scale: capScale });

    // 90° met de klok mee
    const rot = document.createElement('canvas');
    rot.width = canvas.height; rot.height = canvas.width;
    const ctx = rot.getContext('2d');
    ctx.translate(rot.width, 0); ctx.rotate(Math.PI/2); ctx.drawImage(canvas, 0, 0);

    document.body.removeChild(wrap);
    return rot.toDataURL('image/png');
  }

  /* ====== SINGLE PDF (via preview) ====== */
  async function generatePDFSingle(){
    const vals = getFormValues();
    const { sizes } = await renderPreviewFor(vals);

    const jsPDF = await loadJsPDF();
    const h2c   = await loadHtml2Canvas();

    // Na rotatie: breedte = s.h, hoogte = s.w (cm)
    const contentW = Math.max(...sizes.map(s => s.h));
    const contentH = sizes.reduce((sum, s) => sum + s.w, 0);
    const pageW = contentW + PDF_MARGIN_CM*2;
    const pageH = contentH + PDF_MARGIN_CM*2;

    const A4W=21.0, A4H=29.7;
    const doc = new jsPDF({ unit:'cm', orientation:'portrait', format: (pageW<=A4W && pageH<=A4H) ? 'a4' : [pageW, pageH] });
    doc.setFont('helvetica','normal');

    const orderIdx = [1,3,2,4];
    for (let i=0, y=PDF_MARGIN_CM; i<orderIdx.length; i++){
      const img = await capturePreviewLabelToImage(h2c, orderIdx[i], i===0);
      const s = sizes[orderIdx[i]-1];
      const wRot = s.h, hRot = s.w;
      doc.addImage(img, 'PNG', PDF_MARGIN_CM, y, wRot, hRot, undefined, 'FAST');
      y += hRot;
    }
    doc.save(`${vals.code} - ${ts()}.pdf`);
  }

  /* ====== BATCH PIPELINE (via dezelfde preview) ====== */
  function setHidden(elm, hidden){ if (elm) elm.classList.toggle('hidden', hidden); }
  function log(msg, type='info'){
    if (!logList) return;
    const div = el('div', { class: type==='error'?'err': (type==='ok'?'ok':'') }, msg);
    logList.appendChild(div);
  }
  function resetLog(){ if (logList) logList.innerHTML=''; }

  async function parseFile(file){
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type:'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval:'', raw:false });
  }

  const REQUIRED_FIELDS = [
    ['productcode','ERP'],
    ['omschrijving','Omschrijving'],
    ['ean','EAN'],
    ['qty','QTY'],
    ['gw','G.W'],
    ['cbm','CBM'],
    ['lengte','Length (L)'],
    ['breedte','Width (W)'],
    ['hoogte','Height (H)'],
    ['batch','Batch']
  ];
  const ALL_FIELDS = [...REQUIRED_FIELDS];

  const SYNONYMS = {
    productcode: ['erp','erp#','erp #','productcode','code','sku','artikelcode','itemcode','prodcode'],
    omschrijving:['omschrijving','description','product','naam','title','titel'],
    ean:         ['ean','barcode','gtin','ean13','ean_13'],
    qty:         ['qty','aantal','quantity','pcs','stuks'],
    gw:          ['gw','g.w','gewicht','weight','grossweight','gweight','brutogewicht'],
    cbm:         ['cbm','m3','volume','kub','kubiekemeter','kubiekemeters'],
    lengte:      ['lengte','l','depth','diepte','length'],
    breedte:     ['breedte','b','width','w'],
    hoogte:      ['hoogte','h','height'],
    batch:       ['batch','lot','lotno','batchno','batchnr']
  };

  function slugify(s){
    return String(s||'')
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .replace(/[^a-z0-9]+/g,'')
      .trim();
  }

  function guessMapping(hdrs){
    const m = {}; ALL_FIELDS.forEach(([key]) => m[key] = '');
    const slugs = hdrs.map(slugify);
    for (let i=0;i<hdrs.length;i++){
      const h = hdrs[i], s = slugs[i];
      for (const [key, syns] of Object.entries(SYNONYMS)){
        if (syns.includes(s)) { m[key] = h; break; }
      }
    }
    for (const [key] of ALL_FIELDS){
      if (!m[key]) {
        const sset = SYNONYMS[key] || [key];
        const idx = slugs.findIndex(s => sset.some(tok => s.includes(tok)));
        if (idx >= 0) m[key] = hdrs[idx];
      }
    }
    return m;
  }

  function buildMappingUI(hdrs, mapping){
    if (!mappingGrid) return;
    mappingGrid.innerHTML='';
    const makeRow = (key, labelText) => {
      const row = el('div', { class:'map-row' });
      const lab = el('label', {}, labelText + ' *');
      const sel = el('select', { 'data-key': key });
      sel.appendChild(el('option', { value:'' }, '-- kies kolom --'));
      hdrs.forEach(h => {
        const opt = el('option', { value:h }, h);
        if (mapping[key] === h) opt.selected = true;
        sel.appendChild(opt);
      });
      row.append(lab, sel);
      mappingGrid.appendChild(row);
    };
    REQUIRED_FIELDS.forEach(([k,l])=> makeRow(k,l));
    mappingGrid.querySelectorAll('select').forEach(sel=>{
      sel.addEventListener('change', ()=>{ mapping[sel.getAttribute('data-key')] = sel.value; });
    });
  }

  function showTablePreview(rows){
    if (!tablePreview) return;
    if (!rows.length){ tablePreview.innerHTML='<em>Geen data gevonden.</em>'; return; }
    const cols = Object.keys(rows[0]);
    const n = Math.min(5, rows.length);
    const table = el('table');
    const thead = el('thead'); const trh = el('tr');
    cols.forEach(c=> trh.appendChild(el('th',{}, c)));
    thead.appendChild(trh);
    const tbody = el('tbody');
    for (let i=0;i<n;i++){
      const tr = el('tr');
      cols.forEach(c=> tr.appendChild(el('td',{}, String(rows[i][c]??''))));
      tbody.appendChild(tr);
    }
    table.append(thead, tbody);
    tablePreview.innerHTML=''; tablePreview.appendChild(table);
  }

  function normalizeNumber(val){
    if (typeof val !== 'string') val = String(val??'');
    if (chkTrim?.checked) val = val.trim();
    if (chkComma?.checked) val = val.replace(',', '.');
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(val)) { val = val.replace(/\./g, ''); }
    const num = parseFloat(val);
    return isFinite(num) ? num : NaN;
  }

  function readRowWithMapping(row, mapping){
    const get = key => {
      const hdr = mapping[key] || '';
      return hdr ? (row[hdr] ?? '') : '';
    };
    const vals = {
      code: String(get('productcode') ?? '').trim(),
      desc: String(get('omschrijving') ?? '').trim(),
      ean:  String(get('ean') ?? '').trim(),
      qty:  String(Math.max(0, Math.floor(normalizeNumber(String(get('qty')))))),
      gw:   (()=>{ const n = normalizeNumber(String(get('gw'))); return isFinite(n) ? n.toFixed(2) : ''; })(),
      cbm:  String(get('cbm') ?? '').trim(),
      L:    normalizeNumber(String(get('lengte'))),
      W:    normalizeNumber(String(get('breedte'))),
      H:    normalizeNumber(String(get('hoogte'))),
      batch:String(get('batch') ?? '').trim()
    };

    const missing = [];
    if (!vals.code) missing.push('Productcode');
    if (!vals.desc) missing.push('Omschrijving');
    if (!vals.ean)  missing.push('EAN');
    if (!vals.qty || isNaN(+vals.qty)) missing.push('QTY');
    if (!vals.gw)  missing.push('G.W');
    if (!vals.cbm) missing.push('CBM');
    if (!isFinite(vals.L) || vals.L<=0) missing.push('Lengte');
    if (!isFinite(vals.W) || vals.W<=0) missing.push('Breedte');
    if (!isFinite(vals.H) || vals.H<=0) missing.push('Hoogte');
    if (!vals.batch) missing.push('Batch');

    if (missing.length){
      return { ok:false, error:`Ontbrekende/ongeldige velden: ${missing.join(', ')}` };
    }
    vals.L = +vals.L; vals.W = +vals.W; vals.H = +vals.H;
    return { ok:true, vals };
  }

  // Render één PDF via de zichtbare PREVIEW (parity met single)
  async function renderOnePdfBlobViaPreview(vals){
    const oldOpacity = canvasEl.style.opacity;
    canvasEl.style.opacity = '0.15'; // demp UI tijdens batch-render (layout blijft zichtbaar)

    const { sizes } = await renderPreviewFor(vals);

    const jsPDF = await loadJsPDF();
    const h2c   = await loadHtml2Canvas();

    // Na rotatie: breedte = s.h, hoogte = s.w (cm)
    const contentW = Math.max(...sizes.map(s => s.h));
    const contentH = sizes.reduce((sum, s) => sum + s.w, 0);
    const pageW = contentW + PDF_MARGIN_CM*2;
    const pageH = contentH + PDF_MARGIN_CM*2;

    const A4W=21.0, A4H=29.7;
    const doc = new jsPDF({ unit:'cm', orientation:'portrait', format: (pageW<=A4W && pageH<=A4H) ? 'a4' : [pageW, pageH] });
    doc.setFont('helvetica','normal');

    const orderIdx = [1,3,2,4];
    for (let i=0, y=PDF_MARGIN_CM; i<orderIdx.length; i++){
      const img = await capturePreviewLabelToImage(h2c, orderIdx[i], i===0);
      const s = sizes[orderIdx[i]-1];
      const wRot = s.h, hRot = s.w;
      doc.addImage(img, 'PNG', PDF_MARGIN_CM, y, wRot, hRot, undefined, 'FAST');
      y += hRot;
    }

    canvasEl.style.opacity = oldOpacity || '';
    return doc.output('blob');
  }

  /* ====== BATCH UI EVENTS ====== */
  if (btnPickFile && fileInput) btnPickFile.addEventListener('click', ()=> fileInput.click());

  if (dropzone){
    ['dragover','dragenter'].forEach(ev=>{
      dropzone.addEventListener(ev, e=>{ e.preventDefault(); dropzone.classList.add('dragover'); });
    });
    ['dragleave','drop'].forEach(ev=>{
      dropzone.addEventListener(ev, e=>{ e.preventDefault(); dropzone.classList.remove('dragover'); });
    });
    dropzone.addEventListener('drop', async (e)=>{ const f = e.dataTransfer.files?.[0]; if (f) await handleFile(f); });
  }

  if (fileInput){
    fileInput.addEventListener('change', async ()=>{ const f = fileInput.files?.[0]; if (f) await handleFile(f); });
  }

  async function handleFile(file){
    resetLog(); if (logWrap) logWrap.classList.remove('hidden');
    try{
      log(`Bestand: ${file.name}`);
      const rows = await parseFile(file);
      if (!rows.length){ log('Geen rijen gevonden.', 'error'); return; }

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
      log(`Gelezen rijen: ${rows.length}`, 'ok');
    }catch(err){
      log(`Fout bij lezen bestand: ${err.message||err}`, 'error');
    }
  }

  if (btnTemplateCsv){
    btnTemplateCsv.addEventListener('click', ()=>{
      // Alleen headers — géén voorbeeldregels
      const hdrs = ['Productcode','Omschrijving','EAN','QTY','G.W','CBM','Lengte (L)','Breedte (W)','Hoogte (H)','Batch'];
      const blob = new Blob([hdrs.join(',') + '\n'], {type:'text/csv;charset=utf-8;'});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = 'etiketten-template.csv';
      document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
    });
  }

  if (btnTemplateXlsx){
    btnTemplateXlsx.addEventListener('click', ()=>{
      // Alleen headers — géén voorbeeldregels
      const hdrs = ['Productcode','Omschrijving','EAN','QTY','G.W','CBM','Lengte (L)','Breedte (W)','Hoogte (H)','Batch'];
      const ws = XLSX.utils.aoa_to_sheet([hdrs]);
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Etiketten');
      const wbout = XLSX.write(wb, { bookType:'xlsx', type:'array' });
      const blob = new Blob([wbout], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = 'etiketten-template.xlsx';
      document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
    });
  }

  if (btnRunBatch){
    btnRunBatch.addEventListener('click', async ()=>{
      if (!parsedRows.length){ log('Geen dataset geladen.', 'error'); return; }

      const missingMap = REQUIRED_FIELDS.filter(([k]) => !mapping[k]).map(([,label]) => label);
      if (missingMap.length){ log(`Koppel alle verplichte velden: ${missingMap.join(', ')}`, 'error'); return; }

      abortFlag = false;
      if (btnAbortBatch) btnAbortBatch.disabled = false;
      setHidden(progressWrap, false);
      progressBar.style.width = '0%';
      progressLabel.textContent = `${0} / ${parsedRows.length}`;
      progressPhase.textContent = 'Voorbereiden…';

      const zip = new JSZip();
      const batchTime = ts();
      let okCount = 0, errCount = 0;

      for (let i=0; i<parsedRows.length; i++){
        if (abortFlag){ log(`Batch afgebroken op rij ${i+1}.`, 'error'); break; }

        const r = readRowWithMapping(parsedRows[i], mapping);
        if (!r.ok){
          errCount++; log(`Rij ${i+1}: ${r.error}`, 'error');
        } else {
          try{
            progressPhase.textContent = `Rij ${i+1}: PDF renderen…`;
            const blob = await renderOnePdfBlobViaPreview(r.vals); // <<< parity pad
            const safeCode = r.vals.code.replace(/[^\w.-]+/g,'_');
            const name = `${safeCode} - ${batchTime} - R${String(i+1).padStart(3,'0')}.pdf`;
            zip.file(name, blob);
            okCount++;
          }catch(err){
            errCount++; log(`Rij ${i+1}: renderfout: ${err.message||err}`, 'error');
          }
        }
        progressBar.style.width = `${Math.round(((i+1)/parsedRows.length)*100)}%`;
        progressLabel.textContent = `${i+1} / ${parsedRows.length}`;
        await new Promise(r=>setTimeout(r,0)); // UI ademruimte
      }

      if (btnAbortBatch) btnAbortBatch.disabled = true;
      progressPhase.textContent = 'Bundelen als ZIP…';

      if (okCount>0){
        const zipBlob = await zip.generateAsync({type:'blob'});
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a'); a.href = url; a.download = `etiketten-batch - ${batchTime}.zip`;
        document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
        log(`Gereed: ${okCount} PDF’s succesvol, ${errCount} fouten.`, 'ok');
      }else{
        log(`Geen PDF’s gegenereerd. (${errCount} fouten)`, 'error');
      }
      progressPhase.textContent = 'Klaar.';
    });
  }

  if (btnAbortBatch){
    btnAbortBatch.addEventListener('click', ()=>{
      abortFlag = true;
      btnAbortBatch.disabled = true;
      progressPhase.textContent = 'Afbreken…';
    });
  }

  /* ====== EVENTS & INIT ====== */
  const safeRender = () => renderSingle().catch(e => alert(e.message||e));
  if (btnGen) btnGen.addEventListener('click', safeRender);
  if (btnPDF) btnPDF.addEventListener('click', async ()=>{ try{ await generatePDFSingle(); } catch(e){ alert(e.message||e); } });
  window.addEventListener('resize', ()=>{ renderSingle().catch(()=>{}); });

  // (optioneel) demo-waarden voor snelle start
  try{
    if (document.getElementById('prodCode')){
      document.getElementById('prodCode').value   = 'LG1000843';
      document.getElementById('prodDesc').value   = 'Combination Lock - Orange - 1 Pack (YF20610B)';
      document.getElementById('ean').value        = '8719632951889';
      document.getElementById('qty').value        = '12';
      document.getElementById('gw').value         = '18.00';
      document.getElementById('cbm').value        = '0.02';
      document.getElementById('boxLength').value  = '39';
      document.getElementById('boxWidth').value   = '19.5';
      document.getElementById('boxHeight').value  = '22';
      document.getElementById('batch').value      = 'IOR2500307';
      safeRender();
    }
  }catch(_){}
})();
