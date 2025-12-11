// ===== Gewijzigd: renderSingle() hersteld naar v0.28 gedrag =====
async function renderSingle() {
  const vals = getFormValues();
  const sizes = computeLabelSizes(vals);
  const labels = sizes.map((s) => createLabelEl(s, vals, 1)); // 1:1 schaal

  const labelsGrid = document.getElementById("labelsGrid");
  const canvasEl = document.getElementById("previewCanvas");

  labelsGrid.innerHTML = "";
  labels.forEach((el) => labelsGrid.appendChild(el));

  await mountThenFit(labelsGrid);

  // Meet daadwerkelijke breedte voor schaalbepaling (etiket 1 en 3)
  const label1 = labelsGrid.querySelector('.label[data-idx="1"]');
  const label3 = labelsGrid.querySelector('.label[data-idx="3"]');
  const w1 = label1?.offsetWidth || 0;
  const w3 = label3?.offsetWidth || 0;

  const gapPx = 12;
  const canvasWidth = canvasEl.clientWidth;
  const totalWidth = w1 + gapPx + w3;
  const scale = Math.min(canvasWidth / totalWidth, 1);

  const gridRect = labelsGrid.getBoundingClientRect();
  const scaledW = gridRect.width * scale;
  const scaledH = gridRect.height * scale;
  const shiftX = (canvasWidth - scaledW) / 2;

  labelsGrid.style.transform = `translateX(${shiftX}px) scale(${scale})`;
  labelsGrid.style.transformOrigin = "top left";

  canvasEl.style.height = `${scaledH + 24}px`;
}

// ===== Gewijzigd: createLabelEl aangepast voor preview (geen --k) =====
function createLabelEl(size, values, previewScale = 1) {
  const widthPx = Math.round(size.w * PX_PER_CM);
  const heightPx = Math.round(size.h * PX_PER_CM);

  const wrap = el("div", { class: "label-wrap" });
  const label = el("div", {
    class: "label",
    "data-idx": String(size.idx),
    style: {
      width: widthPx + "px",
      height: heightPx + "px",
      padding: LABEL_PADDING_CM * PX_PER_CM + "px",
    },
  });

  const inner = el("div", { class: "label-inner nowrap-mode" });
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

  const detailContent = buildLeftBlock(values, size);
  const detailBox = el("div", { class: "detail-box" }, detailContent);
  const bottomBox = el("div", { class: "bottom-box" }, detailBox);

  inner.append(topBox, bottomBox);
  label.append(inner);
  wrap.append(label, el("div", { class: "label-num" }, `Etiket ${size.idx}`));

  return wrap;
}
