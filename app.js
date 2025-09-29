/*
  Etiketten Generator — JS
  Update:
  - Wrap-drempel: 10 px (daarboven altijd proberen met no-wrap).
  - Absolute max body-font: 28 px.
  - Code-box = min(1.6× body, 28 px).
*/

(() => {
  const PX_PER_CM = 37.7952755906;   // 96dpi
  const PREVIEW_GAP_CM_NUM = 0.5;
  const PDF_MARGIN_CM = 0.5;
  const LABEL_PADDING_CM = 0.5;      // cm binnenmarge

  // ===== typografie-parameters =====
  const WRAP_THRESHOLD_PX = 10;      // pas onder 10px naar soft-wrap
  const MIN_FS_PX = 6;               // absolute minima (soft-wrap noodrem)
  const MAX_BODY_PX = 28;            // bovengrens body in px
  const CODE_MULT    = 1.6;          // code ≈ 1.6× body
  const CODE_CAP_PX  = 28;           // code nooit groter dan 28 px

  const labelsGrid  = document.getElementById('labelsGrid');
  const controlInfo = document.getElementById('controlInfo');
  const canvasEl    = document.getElementById('canvas');
  const btnGen      = document.getElementById('btnGenerate');
  const btnPDF      = document.getElementById('btnPDF');

  let currentPreviewScale = 1;

  const el = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.substring(2), v);
      else node.setAttribute(k, v);
    });
    children.flat().forEach(c => { if (c!=null) node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return node;
  };
  const line = (lab, val) => [el('div',{class:'lab'}, lab), el('div',{class:'val'}, val)];

  function readValues(){
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
      batch:g('batch'),
    };
    if ([L,W,H].some(v=>!isFinite(v)||v<=0) ||
        !vals.code || !vals.desc || !vals.ean || !vals.qty || !vals.gw || !vals.cbm || !vals.batch){
      throw new Error('Controleer de verplichte velden.');
    }
    return vals;
  }

  // Labelmaten (cm): 80% van respectievelijk L×H (1&2) en W×H (3&4)
  function computeLabelSizes({ L, W, H, cn }){
    const scale=0.8;
    const fb={ w:L*scale, h:H*scale };
    const sd={ w:W*scale, h:H*scale };
    return [
      { idx:1, kind:'front/back', ...fb, underType:'china' },
      { idx:2, kind:'front/back', ...fb, underType:'cn' },
      { idx:3, kind:'side',       ...sd, underType:'china' },
      { idx:4, kind:'side',       ...sd, underType:'cn' }
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
    const codeEl = innerEl.querySelector('.code-box');
    innerEl.style.setProperty('--fs', fsPx + 'px');
    if (codeEl){
      const codePx = Math.min(fsPx * CODE_MULT, CODE_CAP_PX);
      codeEl.style.fontSize = codePx + 'px';
    }
  }

  function searchFontSize(innerEl, minFs, hi, guardX, guardY){
    applyFontSizes(innerEl, hi);

    if (fitsWithGuard(innerEl, guardX, guardY)){
      let grow = hi;
      for (let i=0;i<24;i++){
        const next = grow * 1.06;
        applyFontSizes(innerEl, next);
        if (!fitsWithGuard(innerEl, guardX, guardY)){ applyFontSizes(innerEl, grow); return grow; }
        grow = next;
      }
      return grow;
    }

    let lo = minFs, best = lo;
    while (hi - lo > 0.5){
      const mid = (lo + hi)/2;
      applyFontSizes(innerEl, mid);
      if (fitsWithGuard(innerEl, guardX, guardY)){ best = mid; lo = mid; } else { hi = mid; }
    }
    applyFontSizes(innerEl, best);
    return best;
  }

  // Fase 1 (no-wrap ≥ 10 px) → Fase 2 (soft-wrap ≥ 6 px)
  function fitContentToBoxConditional(innerEl){
    const box = innerEl.getBoundingClientRect();
    const guardX = Math.max(6, box.width  * 0.02);
    const guardY = Math.max(6, box.height * 0.02);

    // container-gestuurde bovengrens én globale body-cap van 28 px
    const fracHi = Math.max(16, Math.min(box.height * 0.22, box.width * 0.18));
    const hi = Math.min(fracHi, MAX_BODY_PX);

    // Fase 1: no-wrap
    innerEl.classList.add('nowrap-mode');
    innerEl.classList.remove('softwrap-mode');
    let fs = searchFontSize(innerEl, WRAP_THRESHOLD_PX, hi, guardX, guardY);
    if (fs >= WRAP_THRESHOLD_PX) return;

    // Fase 2: soft-wrap
    innerEl.classList.remove('nowrap-mode');
    innerEl.classList.add('softwrap-mode');
    searchFontSize(innerEl, MIN_FS_PX, hi, guardX, guardY);
  }

  /* ====== UI ====== */

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
    
    if (size.underType === 'china'){
      block.append(el('div',{class:'line'}, 'Made in China'));
    } else {
      // C/N met elastische lijn die nooit wrapt
      const row = el('div',{class:'cn-row'},
        el('div',{class:'lab'}, 'C/N:'),
        el('div',{class:'cn-line'}, '')
      );
      block.append(row);
    }
    return block;
  }

  function createLabelEl(size, values, previewScale){
    const widthPx  = Math.round(size.w * PX_PER_CM * previewScale);
    const heightPx = Math.round(size.h * PX_PER_CM * previewScale);

    const wrap  = el('div', { class:'label-wrap' });
    const label = el('div', { class:'label', style:{ width:widthPx+'px', height:heightPx+'px' }});
    label.dataset.idx = String(size.idx);
    const inner = el('div', { class:'label-inner nowrap-mode' });

    // fysieke padding → px in preview
    inner.style.padding = (LABEL_PADDING_CM * PX_PER_CM * previewScale) + 'px';

    const head = el('div', { class:'label-head' },
      el('div', { class:'code-box line' }, values.code),
      el('div', { class:'line' }, values.desc)
    );

    inner.append(head, el('div',{class:'block-spacer'}), buildLeftBlock(values, size));
    label.append(inner);
    wrap.append(label, el('div',{class:'label-num'}, `Etiket ${size.idx}`));

    // startwaarde (ruwe schatting) en fit
    applyFontSizes(inner, Math.max(10, Math.floor(heightPx * 0.06)));
    fitContentToBoxConditional(inner);

    return wrap;
  }

  function render(){
    const vals = readValues();
    const sizes = computeLabelSizes(vals);
    const scale = computePreviewScale(sizes);
    currentPreviewScale = scale;

    updateControlInfo(sizes);
    labelsGrid.style.gap = '0.5cm';
    labelsGrid.innerHTML = '';

    // Preview volgorde: 1 & 3 boven, 2 & 4 onder
    [0,2,1,3].forEach(i => labelsGrid.appendChild(createLabelEl(sizes[i], vals, scale)));
  }

  /* ====== PDF (preview capture) ====== */
  function loadJsPDF(){ return new Promise((res,rej)=>{ if (window.jspdf?.jsPDF) return res(window.jspdf.jsPDF);
    const s=document.createElement('script'); s.src='https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
    s.onload=()=>res(window.jspdf.jsPDF); s.onerror=()=>rej(new Error('Kon jsPDF niet laden.')); document.head.appendChild(s); }); }
  function loadHtml2Canvas(){ return new Promise((res,rej)=>{ if (window.html2canvas) return res(window.html2canvas);
    const s=document.createElement('script'); s.src='https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js';
    s.onload=()=>res(window.html2canvas); s.onerror=()=>rej(new Error('Kon html2canvas niet laden.')); document.head.appendChild(s); }); }
  const ts=(d=new Date())=>{const p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}.${p(d.getMinutes())}.${p(d.getSeconds())}`;};

  async function capturePreviewLabelToImage(h2c, idx, isFirst) {
    const src = document.querySelector(`.label[data-idx="${idx}"]`);
    if (!src) throw new Error('Genereer eerst de preview.');

    const clone = src.cloneNode(true);
    clone.style.borderTop = '1px solid #000';
    clone.style.borderRight = '1px solid #000';
    clone.style.borderBottom = '1px solid #000';
    clone.style.borderLeft = isFirst ? '1px solid #000' : '0';

    const wrap = document.createElement('div');
    wrap.style.position='fixed'; wrap.style.left='-10000px'; wrap.style.top='0'; wrap.style.background='#fff';
    document.body.appendChild(wrap);
    wrap.appendChild(clone);

    const capScale = Math.max(2, (1/currentPreviewScale));
    const canvas = await h2c(clone, { backgroundColor:'#fff', scale: capScale });

    const rot = document.createElement('canvas');
    rot.width = canvas.height; rot.height = canvas.width;
    const ctx = rot.getContext('2d');
    ctx.translate(rot.width, 0); ctx.rotate(Math.PI/2); ctx.drawImage(canvas, 0, 0);

    document.body.removeChild(wrap);
    return rot.toDataURL('image/png');
  }

  async function generatePDF(){
    if (!document.querySelector('.label')) { try { render(); } catch (e) { alert(e.message||e); return; } }

    const vals  = readValues();
    const sizes = computeLabelSizes(vals);
    const jsPDF = await loadJsPDF();
    const h2c   = await loadHtml2Canvas();

    const contentW = Math.max(...sizes.map(s => s.h));
    const contentH = sizes.reduce((sum, s) => sum + s.w, 0);
    const pageW = contentW + PDF_MARGIN_CM*2;
    const pageH = contentH + PDF_MARGIN_CM*2;

    const A4W=21.0, A4H=29.7;
    const doc = new jsPDF({ unit:'cm', orientation:'portrait', format: (pageW<=A4W && pageH<=A4H) ? 'a4' : [pageW, pageH] });
    doc.setFont('helvetica','normal');

    // PDF volgorde: 1 → 3 → 2 → 4
    const orderIdx = [1,3,2,4];
    const imgs = [];
    for (let i=0;i<orderIdx.length;i++){
      imgs.push(await capturePreviewLabelToImage(h2c, orderIdx[i], i===0));
    }

    let y = PDF_MARGIN_CM, x = PDF_MARGIN_CM;
    for (let i=0;i<orderIdx.length;i++){
      const s = sizes[orderIdx[i]-1];
      const wRot = s.h, hRot = s.w; // cm
      doc.addImage(imgs[i], 'PNG', x, y, wRot, hRot, undefined, 'FAST');
      y += hRot;
    }
    doc.save(`${vals.code} - ${ts()}.pdf`);
  }

  /* ====== events + demo ====== */
  const safeRender = () => { try{ render(); } catch(e){ alert(e.message||e); } };
  btnGen.addEventListener('click', safeRender);
  btnPDF.addEventListener('click', async ()=>{ try{ await generatePDF(); } catch(e){ alert(e.message||e); } });
  window.addEventListener('resize', ()=>{ try{ render(); } catch(_){} });

  // Demo
  document.getElementById('prodCode').value   = 'LG1000843';
  document.getElementById('prodDesc').value   = 'Combination Lock - Orange - 1 Pack (YF20610B)';
  document.getElementById('ean').value        = '8719632951889';
  document.getElementById('qty').value        = '12';
  document.getElementById('gw').value         = '18.00';
  document.getElementById('cbm').value        = '0.02';
  document.getElementById('boxLength').value  = '39';
  document.getElementById('boxWidth').value   = '19.5';
  document.getElementById('boxHeight').value  = '22';
  document.getElementById('cnValue').value    = '';
  document.getElementById('batch').value      = '';

  safeRender();
})();
