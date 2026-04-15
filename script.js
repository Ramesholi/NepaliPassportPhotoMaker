/* Pro Passport Photo Maker Nepal — Cyber Studio Edition */
/* Frontend only, no backend, no login, no payment */

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const PRESETS = {
  nepal_small: { label: "नेपाली पासपोर्ट स्मल साइज", wIn: 1, hIn: 1.2, wMm: 25.4, hMm: 30.48, unit: "inch", dpi: 300, note: "Default Nepal passport small size" },
  passport: { label: "पासपोर्ट साइज", wIn: 1.5, hIn: 1.5, wMm: 38.1, hMm: 38.1, unit: "inch", dpi: 300, note: "Square passport size" },
  visa: { label: "भिसा फोटो", wIn: 2, hIn: 2, wMm: 50.8, hMm: 50.8, unit: "inch", dpi: 300, note: "Common visa print" },
  citizenship: { label: "नागरिकता फोटो", wIn: 35 / 25.4, hIn: 45 / 25.4, wMm: 35, hMm: 45, unit: "mm", dpi: 300, note: "35mm × 45mm" },
  id: { label: "ID कार्ड फोटो", wIn: 1, hIn: 1, wMm: 25.4, hMm: 25.4, unit: "inch", dpi: 300, note: "One inch ID size" },
};

const state = {
  theme: "dark",
  mode: "editor",
  activePhotoId: null,
  photos: [],
  history: [],
  historyIndex: -1,
  cropper: null,
  bodyPixModel: null,
  busy: false,
  settings: {
    preset: "nepal_small",
    customWidth: 1,
    customHeight: 1.2,
    sizeUnit: "inch",
    dpi: 300,
    customDpi: 300,
    copies: 4,
    customCopies: 12,
    marginMm: 6,
    spacingMm: 3,
    brightness: 0,
    contrast: 0,
    saturation: 0,
    sharpness: 0,
    temperature: 0,
    bgColorPreset: "#ffffff",
    customBgColor: "#ffffff",
    bgRemoval: false,
    snapGrid: true,
    showSafe: true,
  },
  layout: {
    objects: [],
    pageDpi: 300,
    pageWIn: 8.27,
    pageHIn: 11.69,
    zoom: 1,
  },
};

const els = {
  html: document.documentElement,
  themeToggle: $("#themeToggle"),
  fileInput: $("#fileInput"),
  dropZone: $("#dropZone"),
  photoList: $("#photoList"),
  activePhotoPill: $("#activePhotoPill"),
  sizeInfoPill: $("#sizeInfoPill"),
  modeLabel: $("#modeLabel"),
  copiesLabel: $("#copiesLabel"),
  canvasLabel: $("#canvasLabel"),
  toast: $("#toast"),
  editorImage: $("#editorImage"),
  editorEmptyState: $("#editorEmptyState"),
  editorStage: $(".editor-stage"),
  singlePreviewTab: null,
  layoutImagePicker: $("#layoutImagePicker"),
  sheetCanvas: $("#sheetCanvas"),
  layoutCanvas: $("#layoutCanvas"),
  generateBtn: $("#generateBtn"),
  printBtn: $("#printBtn"),
  loadProjectBtn: $("#loadProjectBtn"),
  loadProjectInput: $("#loadProjectInput"),
  saveProjectBtn: $("#saveProjectBtn"),
  clearAllBtn: $("#clearAllBtn"),
  refreshPreviewBtn: $("#refreshPreviewBtn"),
  downloadPngBtn: $("#downloadPngBtn"),
  downloadPdfBtn: $("#downloadPdfBtn"),
  printPreviewBtn: $("#printPreviewBtn"),
  addToLayoutBtn: $("#addToLayoutBtn"),
  clearLayoutBtn: $("#clearLayoutBtn"),
  undoBtn: $("#undoBtn"),
  redoBtn: $("#redoBtn"),
  zoomInBtn: $("#zoomInBtn"),
  zoomOutBtn: $("#zoomOutBtn"),
  fitBtn: $("#fitBtn"),
  rotateLeftBtn: $("#rotateLeftBtn"),
  rotateRightBtn: $("#rotateRightBtn"),
  bgRemoveBtn: $("#bgRemoveBtn"),
  autoEnhanceBtn: $("#autoEnhanceBtn"),
  resetEditsBtn: $("#resetEditsBtn"),
  brightness: $("#brightness"),
  contrast: $("#contrast"),
  saturation: $("#saturation"),
  sharpness: $("#sharpness"),
  temperature: $("#temperature"),
  brightnessVal: $("#brightnessVal"),
  contrastVal: $("#contrastVal"),
  saturationVal: $("#saturationVal"),
  sharpnessVal: $("#sharpnessVal"),
  temperatureVal: $("#temperatureVal"),
  bgColorPreset: $("#bgColorPreset"),
  customBgColor: $("#customBgColor"),
  bgRemovalToggle: $("#bgRemovalToggle"),
  snapGridToggle: $("#snapGridToggle"),
  showSafeToggle: $("#showSafeToggle"),
  sizePreset: $("#sizePreset"),
  customWidth: $("#customWidth"),
  customHeight: $("#customHeight"),
  sizeUnit: $("#sizeUnit"),
  dpiSelect: $("#dpiSelect"),
  customDpiWrap: $("#customDpiWrap"),
  customDpi: $("#customDpi"),
  copiesSelect: $("#copiesSelect"),
  customCopiesWrap: $("#customCopiesWrap"),
  customCopies: $("#customCopies"),
  marginMm: $("#marginMm"),
  spacingMm: $("#spacingMm"),
  presetCards: $("#presetCards"),
  uploadPanel: $("#uploadPanel"),
  sizesPanel: $("#sizesPanel"),
  editorPanel: $("#editorPanel"),
  layoutPanel: $("#layoutPanel"),
  exportPanel: $("#exportPanel"),
};

let bodyPixReady = false;
let layoutFabric = null;
let resizeRAF = null;
let currentPreviewCanvas = null;

function toast(msg) {
  els.toast.textContent = msg;
  els.toast.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => els.toast.classList.remove("show"), 2200);
}

function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }
function uuid() { return `${Date.now()}_${Math.random().toString(16).slice(2)}`; }

function mmToIn(mm) { return mm / 25.4; }
function inToPx(inches, dpi) { return Math.round(inches * dpi); }
function mmToPx(mm, dpi) { return Math.round(mmToIn(mm) * dpi); }

function getDpi() {
  return els.dpiSelect.value === "custom" ? Number(els.customDpi.value || 300) : Number(els.dpiSelect.value);
}

function getPresetDims() {
  const preset = els.sizePreset.value;
  if (preset !== "custom" && PRESETS[preset]) {
    const p = PRESETS[preset];
    return { ...p, dpi: getDpi(), wIn: p.wIn, hIn: p.hIn, unit: "inch" };
  }

  const unit = els.sizeUnit.value;
  let w = Number(els.customWidth.value || 1);
  let h = Number(els.customHeight.value || 1);
  if (unit === "mm") {
    w = mmToIn(w);
    h = mmToIn(h);
  }
  return { label: "Custom", wIn: w, hIn: h, unit, dpi: getDpi() };
}

function currentBackground() {
  return els.bgColorPreset.value === "custom" ? els.customBgColor.value : els.bgColorPreset.value;
}

function updateSizeInfo() {
  const d = getPresetDims();
  const pxW = inToPx(d.wIn, d.dpi);
  const pxH = inToPx(d.hIn, d.dpi);
  els.sizeInfoPill.textContent = `${d.label} · ${d.wIn.toFixed(2)} × ${d.hIn.toFixed(2)} in · ${d.dpi} DPI · ${pxW}×${pxH}px`;
  els.copiesLabel.textContent = String(getCopies());
  els.canvasLabel.textContent = `A4 @ ${getDpi()} DPI`;
}

function setTheme(theme) {
  state.theme = theme;
  els.html.dataset.theme = theme;
}

function showPanel(panelId) {
  $$(".panel-section").forEach(p => p.classList.remove("active"));
  $("#" + panelId)?.classList.add("active");
  $$(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.panel === panelId));
}

function setPreviewTab(id) {
  $$(".preview-pane").forEach(p => p.classList.remove("active"));
  $("#" + id)?.classList.add("active");
  $$(".tab-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.preview === id));
  if (id === "layoutPreview") {
    resizeLayoutCanvas();
  }
}

function getActivePhoto() {
  return state.photos.find(p => p.id === state.activePhotoId) || null;
}

function setActivePhoto(id) {
  state.activePhotoId = id;
  const photo = getActivePhoto();
  if (!photo) {
    els.activePhotoPill.textContent = "No photo selected";
    els.editorStage.classList.remove("ready");
    els.editorImage.removeAttribute("src");
    destroyCropper();
    renderPhotoList();
    return;
  }
  els.activePhotoPill.textContent = photo.name;
  renderPhotoList();
  loadPhotoIntoEditor(photo);
}

function renderPhotoList() {
  els.photoList.innerHTML = "";
  state.photos.forEach((photo, index) => {
    const card = document.createElement("div");
    card.className = `photo-card ${photo.id === state.activePhotoId ? "active" : ""}`;
    card.innerHTML = `
      <img class="photo-thumb" src="${photo.dataUrl}" alt="">
      <div class="photo-meta">
        <div class="name">${photo.name}</div>
        <div class="details">${photo.width || "?"} × ${photo.height || "?"} px</div>
      </div>
      <div class="photo-actions">
        <button title="Select">◎</button>
        <button title="Delete">✕</button>
      </div>
    `;
    card.addEventListener("click", (e) => {
      const btn = e.target.closest("button");
      if (!btn) return setActivePhoto(photo.id);
      if (btn.title === "Delete") deletePhoto(photo.id);
      if (btn.title === "Select") setActivePhoto(photo.id);
    });
    els.photoList.appendChild(card);
  });
}

function deletePhoto(id) {
  const idx = state.photos.findIndex(p => p.id === id);
  if (idx === -1) return;
  state.photos.splice(idx, 1);
  if (state.activePhotoId === id) state.activePhotoId = state.photos[0]?.id || null;
  renderPhotoList();
  if (state.activePhotoId) setActivePhoto(state.activePhotoId);
  else setActivePhoto(null);
  saveLocalProject();
  toast("Photo removed");
}

function addPhotosFromFiles(fileList) {
  [...fileList].forEach(file => {
    if (!file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const photo = {
          id: uuid(),
          name: file.name.replace(/\.[^.]+$/, ""),
          fileName: file.name,
          dataUrl: reader.result,
          width: img.width,
          height: img.height,
          crop: { x: 0, y: 0, zoom: 1, rotate: 0 },
          filters: { brightness: 0, contrast: 0, saturation: 0, sharpness: 0, temperature: 0 },
          bgRemovedDataUrl: null,
          customBg: currentBackground(),
        };
        state.photos.unshift(photo);
        state.activePhotoId = photo.id;
        renderPhotoList();
        setActivePhoto(photo.id);
        saveHistory();
        saveLocalProject();
        toast(`Loaded ${file.name}`);
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function destroyCropper() {
  if (state.cropper) {
    state.cropper.destroy();
    state.cropper = null;
  }
}

function loadPhotoIntoEditor(photo) {
  destroyCropper();
  els.editorStage.classList.remove("ready");
  els.editorImage.src = photo.bgRemovedDataUrl || photo.dataUrl;
  els.editorImage.onload = () => {
    els.editorStage.classList.add("ready");
    state.cropper = new Cropper(els.editorImage, {
      viewMode: 1,
      dragMode: "move",
      autoCropArea: 0.92,
      background: false,
      responsive: true,
      restore: false,
      guides: true,
      center: true,
      highlight: false,
      cropBoxMovable: true,
      cropBoxResizable: true,
      toggleDragModeOnDblclick: false,
      ready() {
        applyCropperOverlay();
        syncEditorStateFromPhoto(photo);
        updatePreviewDebounced();
      },
      crop() {
        syncCropToPhoto(photo);
        updatePreviewDebounced();
      }
    });
  };
}

function applyCropperOverlay() {
  const cropper = state.cropper;
  if (!cropper) return;
  const container = cropper.cropper;
  if (!container) return;
  const cropBox = container.querySelector(".cropper-crop-box");
  if (cropBox) {
    cropBox.style.boxShadow = "0 0 0 9999px rgba(0,0,0,.15)";
  }
}

function syncEditorStateFromPhoto(photo) {
  if (!photo) return;
  els.brightness.value = photo.filters.brightness ?? 0;
  els.contrast.value = photo.filters.contrast ?? 0;
  els.saturation.value = photo.filters.saturation ?? 0;
  els.sharpness.value = photo.filters.sharpness ?? 0;
  els.temperature.value = photo.filters.temperature ?? 0;
  updateSliderLabels();
}

function syncCropToPhoto(photo) {
  if (!state.cropper || !photo) return;
  const d = state.cropper.getData(true);
  photo.crop = {
    x: d.x,
    y: d.y,
    width: d.width,
    height: d.height,
    rotate: d.rotate || 0,
    scaleX: d.scaleX || 1,
    scaleY: d.scaleY || 1,
  };
}

function updateSliderLabels() {
  els.brightnessVal.textContent = els.brightness.value;
  els.contrastVal.textContent = els.contrast.value;
  els.saturationVal.textContent = els.saturation.value;
  els.sharpnessVal.textContent = els.sharpness.value;
  els.temperatureVal.textContent = els.temperature.value;
}

function saveHistory() {
  const snapshot = {
    settings: { ...state.settings },
    photos: JSON.parse(JSON.stringify(state.photos)),
    activePhotoId: state.activePhotoId,
    layout: layoutFabric ? layoutFabric.toJSON(["lockMovementX", "lockMovementY"]) : { objects: state.layout.objects },
  };
  state.history = state.history.slice(0, state.historyIndex + 1);
  state.history.push(snapshot);
  if (state.history.length > 30) state.history.shift();
  state.historyIndex = state.history.length - 1;
}

function restoreSnapshot(snapshot) {
  if (!snapshot) return;
  Object.assign(state.settings, snapshot.settings);
  applySettingsToUI();
  state.photos = snapshot.photos || [];
  state.activePhotoId = snapshot.activePhotoId || null;
  renderPhotoList();
  if (layoutFabric) {
    layoutFabric.clear();
    layoutFabric.loadFromJSON(snapshot.layout, () => {
      layoutFabric.renderAll();
      enableLayoutGrid();
    });
  }
  if (state.activePhotoId) setActivePhoto(state.activePhotoId);
  else setActivePhoto(null);
  updateSizeInfo();
  updatePreviewDebounced();
  saveLocalProject(false);
}

function undo() {
  if (state.historyIndex <= 0) return toast("Nothing to undo");
  state.historyIndex -= 1;
  restoreSnapshot(state.history[state.historyIndex]);
  toast("Undo");
}

function redo() {
  if (state.historyIndex >= state.history.length - 1) return toast("Nothing to redo");
  state.historyIndex += 1;
  restoreSnapshot(state.history[state.historyIndex]);
  toast("Redo");
}

function mmFromInput() {
  return { marginPx: mmToPx(Number(els.marginMm.value || 0), getDpi()), spacingPx: mmToPx(Number(els.spacingMm.value || 0), getDpi()) };
}

function getCopies() {
  return els.copiesSelect.value === "custom" ? Number(els.customCopies.value || 1) : Number(els.copiesSelect.value);
}

function createWorkingCanvas(width, height) {
  const c = document.createElement("canvas");
  c.width = width;
  c.height = height;
  return c;
}

function sharpenImageData(imageData, amount) {
  if (!amount) return imageData;
  const src = imageData.data;
  const w = imageData.width, h = imageData.height;
  const out = new ImageData(w, h);
  const dst = out.data;
  const kernel = [
    0, -1, 0,
    -1, 5 + amount / 50, -1,
    0, -1, 0,
  ];
  const get = (x, y, c) => {
    x = clamp(x, 0, w - 1); y = clamp(y, 0, h - 1);
    return src[(y * w + x) * 4 + c];
  };
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      for (let c = 0; c < 4; c++) {
        if (c === 3) {
          dst[(y * w + x) * 4 + c] = get(x, y, c);
          continue;
        }
        let sum = 0, k = 0;
        for (let ky = -1; ky <= 1; ky++) {
          for (let kx = -1; kx <= 1; kx++) {
            sum += get(x + kx, y + ky, c) * kernel[k++];
          }
        }
        dst[(y * w + x) * 4 + c] = clamp(sum, 0, 255);
      }
    }
  }
  return out;
}

function applyFiltersToCanvas(canvas, filters, bgColor) {
  const ctx = canvas.getContext("2d");
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = img.data;
  const brightness = filters.brightness || 0;
  const contrast = filters.contrast || 0;
  const saturation = filters.saturation || 0;
  const temperature = filters.temperature || 0;
  const sharpness = filters.sharpness || 0;

  const b = brightness * 2.55;
  const c = (contrast + 100) / 100;
  const s = (saturation + 100) / 100;
  const t = temperature / 100;

  for (let i = 0; i < data.length; i += 4) {
    let r = data[i], g = data[i + 1], bl = data[i + 2];
    r = ((r - 128) * c) + 128 + b;
    g = ((g - 128) * c) + 128 + b;
    bl = ((bl - 128) * c) + 128 + b;

    const gray = (r + g + bl) / 3;
    r = gray + (r - gray) * s;
    g = gray + (g - gray) * s;
    bl = gray + (bl - gray) * s;

    r = r + t * 32;
    bl = bl - t * 24;

    data[i] = clamp(r, 0, 255);
    data[i + 1] = clamp(g, 0, 255);
    data[i + 2] = clamp(bl, 0, 255);
  }

  let out = img;
  if (sharpness > 0) {
    out = sharpenImageData(out, sharpness);
  }

  const outCanvas = createWorkingCanvas(canvas.width, canvas.height);
  const outCtx = outCanvas.getContext("2d");
  if (bgColor) {
    outCtx.fillStyle = bgColor;
    outCtx.fillRect(0, 0, outCanvas.width, outCanvas.height);
  }
  outCtx.putImageData(out, 0, 0);
  return outCanvas;
}

async function removeBackgroundCanvas(sourceCanvas, bgColor = "#ffffff") {
  if (!state.bodyPixModel) {
    try {
      if (!window.bodyPix) throw new Error("BodyPix unavailable");
      state.bodyPixModel = await bodyPix.load({ architecture: "MobileNetV1", outputStride: 16, multiplier: 0.75, quantBytes: 2 });
      bodyPixReady = true;
    } catch (e) {
      console.warn(e);
      toast("Background removal library unavailable");
      return sourceCanvas;
    }
  }
  const segmentation = await state.bodyPixModel.segmentPerson(sourceCanvas, {
    internalResolution: "medium",
    segmentationThreshold: 0.7,
    scoreThreshold: 0.3,
  });

  const out = createWorkingCanvas(sourceCanvas.width, sourceCanvas.height);
  const ctx = out.getContext("2d");
  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, out.width, out.height);

  const imageData = sourceCanvas.getContext("2d").getImageData(0, 0, sourceCanvas.width, sourceCanvas.height);
  const outData = ctx.getImageData(0, 0, out.width, out.height);
  for (let i = 0; i < segmentation.data.length; i++) {
    const mask = segmentation.data[i];
    const idx = i * 4;
    if (mask === 1) {
      outData.data[idx] = imageData.data[idx];
      outData.data[idx + 1] = imageData.data[idx + 1];
      outData.data[idx + 2] = imageData.data[idx + 2];
      outData.data[idx + 3] = 255;
    }
  }
  ctx.putImageData(outData, 0, 0);
  return out;
}

function getCurrentCropCanvas(width, height) {
  const photo = getActivePhoto();
  if (!photo || !state.cropper) return null;
  const cropCanvas = state.cropper.getCroppedCanvas({
    width,
    height,
    imageSmoothingEnabled: true,
    imageSmoothingQuality: "high",
  });
  return cropCanvas;
}

async function buildSinglePreview() {
  const photo = getActivePhoto();
  if (!photo || !state.cropper) {
    clearCanvas(els.sheetCanvas);
    return null;
  }
  const { wIn, hIn, dpi } = getPresetDims();
  const w = inToPx(wIn, dpi);
  const h = inToPx(hIn, dpi);
  let canvas = getCurrentCropCanvas(w, h);
  if (!canvas) return null;

  const filters = {
    brightness: Number(els.brightness.value),
    contrast: Number(els.contrast.value),
    saturation: Number(els.saturation.value),
    sharpness: Number(els.sharpness.value),
    temperature: Number(els.temperature.value),
  };

  if (photo.bgRemovedDataUrl) {
    canvas = await loadImageToCanvas(photo.bgRemovedDataUrl, w, h);
  }

  canvas = applyFiltersToCanvas(canvas, filters, currentBackground());
  currentPreviewCanvas = canvas;
  return canvas;
}

function clearCanvas(canvas) {
  const ctx = canvas.getContext("2d");
  canvas.width = 1200;
  canvas.height = 1200;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}

function drawSafeMargin(ctx, w, h, margin) {
  ctx.save();
  ctx.strokeStyle = "rgba(0,0,0,.25)";
  ctx.setLineDash([12, 10]);
  ctx.lineWidth = 2;
  ctx.strokeRect(margin, margin, w - margin * 2, h - margin * 2);
  ctx.restore();
}

function computeGrid(count) {
  const cols = Math.ceil(Math.sqrt(count));
  const rows = Math.ceil(count / cols);
  return { cols, rows };
}

async function buildSheetPreview() {
  const single = await buildSinglePreview();
  if (!single) return null;

  const dpi = getDpi();
  const a4W = inToPx(8.27, dpi);
  const a4H = inToPx(11.69, dpi);
  const canvas = createWorkingCanvas(a4W, a4H);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, a4W, a4H);

  const { marginPx, spacingPx } = mmFromInput();
  const copies = getCopies();
  const { cols, rows } = computeGrid(copies);
  const usableW = a4W - marginPx * 2 - spacingPx * (cols - 1);
  const usableH = a4H - marginPx * 2 - spacingPx * (rows - 1);

  const tileW = Math.floor(usableW / cols);
  const tileH = Math.floor(usableH / rows);

  const target = await buildSinglePreview();
  if (!target) return null;

  let index = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (index >= copies) break;
      const x = marginPx + c * (tileW + spacingPx);
      const y = marginPx + r * (tileH + spacingPx);
      ctx.drawImage(target, x, y, tileW, tileH);
      index++;
    }
  }

  if (els.showSafeToggle.checked) drawSafeMargin(ctx, a4W, a4H, marginPx);
  return canvas;
}

function renderCanvasOnElement(canvas, element) {
  if (!canvas || !element) return;
  const url = canvas.toDataURL("image/png");
  if (element.tagName === "CANVAS") {
    const ctx = element.getContext("2d");
    const ratio = Math.min(element.clientWidth / canvas.width, element.clientHeight / canvas.height);
    const drawW = Math.max(1, Math.floor(canvas.width * ratio));
    const drawH = Math.max(1, Math.floor(canvas.height * ratio));
    element.width = canvas.width;
    element.height = canvas.height;
    ctx.clearRect(0, 0, element.width, element.height);
    ctx.drawImage(canvas, 0, 0);
    element.style.maxHeight = "72vh";
    element.dataset.previewUrl = url;
  }
  return url;
}

async function updatePreviewDebounced() {
  clearTimeout(updatePreviewDebounced.t);
  updatePreviewDebounced.t = setTimeout(async () => {
    const photo = getActivePhoto();
    if (!photo) {
      clearCanvas(els.sheetCanvas);
      clearCanvas(els.layoutCanvas);
      return;
    }
    const canvas = await buildSinglePreview();
    if (canvas) {
      renderCanvasOnElement(canvas, els.sheetCanvas);
    }
    await updateLayoutCanvasPreview();
  }, 100);
}

async function updateLayoutCanvasPreview() {
  if (!layoutFabric) return;
  layoutFabric.renderAll();
}

function applyPreset(preset) {
  const p = PRESETS[preset];
  if (!p) return;
  if (p.unit === "mm") {
    els.customWidth.value = p.wMm.toFixed(0);
    els.customHeight.value = p.hMm.toFixed(0);
  } else {
    els.customWidth.value = p.wIn.toFixed(2);
    els.customHeight.value = p.hIn.toFixed(2);
  }
  els.sizeUnit.value = p.unit;
  if (els.dpiSelect.value !== "custom") els.dpiSelect.value = String(p.dpi);
  updateSizeInfo();
  saveHistory();
  saveLocalProject();
  updatePreviewDebounced();
}

function applySettingsToUI() {
  els.sizePreset.value = state.settings.preset;
  els.customWidth.value = state.settings.customWidth;
  els.customHeight.value = state.settings.customHeight;
  els.sizeUnit.value = state.settings.sizeUnit;
  els.dpiSelect.value = String(state.settings.dpi);
  els.customDpi.value = state.settings.customDpi;
  els.copiesSelect.value = String(state.settings.copies);
  els.customCopies.value = state.settings.customCopies;
  els.marginMm.value = state.settings.marginMm;
  els.spacingMm.value = state.settings.spacingMm;
  els.brightness.value = state.settings.brightness;
  els.contrast.value = state.settings.contrast;
  els.saturation.value = state.settings.saturation;
  els.sharpness.value = state.settings.sharpness;
  els.temperature.value = state.settings.temperature;
  els.bgColorPreset.value = state.settings.bgColorPreset;
  els.customBgColor.value = state.settings.customBgColor;
  els.bgRemovalToggle.checked = state.settings.bgRemoval;
  els.snapGridToggle.checked = state.settings.snapGrid;
  els.showSafeToggle.checked = state.settings.showSafe;
  updateSliderLabels();
  updateToggles();
  updateSizeInfo();
}

function updateToggles() {
  els.customDpiWrap.style.display = els.dpiSelect.value === "custom" ? "grid" : "none";
  els.customCopiesWrap.style.display = els.copiesSelect.value === "custom" ? "grid" : "none";
}

function saveStateFromUI() {
  state.settings = {
    preset: els.sizePreset.value,
    customWidth: Number(els.customWidth.value),
    customHeight: Number(els.customHeight.value),
    sizeUnit: els.sizeUnit.value,
    dpi: getDpi(),
    customDpi: Number(els.customDpi.value),
    copies: getCopies(),
    customCopies: Number(els.customCopies.value),
    marginMm: Number(els.marginMm.value),
    spacingMm: Number(els.spacingMm.value),
    brightness: Number(els.brightness.value),
    contrast: Number(els.contrast.value),
    saturation: Number(els.saturation.value),
    sharpness: Number(els.sharpness.value),
    temperature: Number(els.temperature.value),
    bgColorPreset: els.bgColorPreset.value,
    customBgColor: els.customBgColor.value,
    bgRemoval: els.bgRemovalToggle.checked,
    snapGrid: els.snapGridToggle.checked,
    showSafe: els.showSafeToggle.checked,
  };
  updateSizeInfo();
}

async function loadImageToCanvas(url, width, height) {
  const img = new Image();
  img.crossOrigin = "anonymous";
  img.src = url;
  await img.decode();
  const canvas = createWorkingCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  const ratio = Math.max(width / img.width, height / img.height);
  const drawW = img.width * ratio;
  const drawH = img.height * ratio;
  const dx = (width - drawW) / 2;
  const dy = (height - drawH) / 2;
  ctx.drawImage(img, dx, dy, drawW, drawH);
  return canvas;
}

async function removeBackgroundForActive() {
  const photo = getActivePhoto();
  if (!photo) return toast("Upload and select a photo first");
  if (!state.cropper) return;
  toast("Removing background...");
  const { wIn, hIn, dpi } = getPresetDims();
  const w = inToPx(wIn, dpi);
  const h = inToPx(hIn, dpi);
  let canvas = getCurrentCropCanvas(Math.max(w, 1200), Math.max(h, 1200));
  if (!canvas) return;
  const color = currentBackground();
  canvas = await removeBackgroundCanvas(canvas, color);
  photo.bgRemovedDataUrl = canvas.toDataURL("image/png");
  renderPhotoList();
  await updatePreviewDebounced();
  saveHistory();
  saveLocalProject();
  toast("Background removed");
}

function autoEnhanceActive() {
  const photo = getActivePhoto();
  if (!photo) return;
  const img = new Image();
  img.onload = async () => {
    const canvas = createWorkingCanvas(img.width, img.height);
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const pixels = data.data;
    let sum = 0;
    for (let i = 0; i < pixels.length; i += 4) sum += (pixels[i] + pixels[i + 1] + pixels[i + 2]) / 3;
    const avg = sum / (pixels.length / 4);
    const target = 150;
    const brightness = clamp(Math.round((target - avg) / 2), -30, 30);
    const contrast = avg < 110 ? 18 : 10;
    photo.filters.brightness = brightness;
    photo.filters.contrast = contrast;
    photo.filters.saturation = 8;
    photo.filters.sharpness = 12;
    photo.filters.temperature = 0;
    els.brightness.value = brightness;
    els.contrast.value = contrast;
    els.saturation.value = 8;
    els.sharpness.value = 12;
    els.temperature.value = 0;
    updateSliderLabels();
    saveHistory();
    saveLocalProject();
    await updatePreviewDebounced();
    toast("Auto enhanced");
  };
  img.src = photo.bgRemovedDataUrl || photo.dataUrl;
}

function resetEdits() {
  const photo = getActivePhoto();
  if (!photo) return;
  photo.filters = { brightness: 0, contrast: 0, saturation: 0, sharpness: 0, temperature: 0 };
  photo.bgRemovedDataUrl = null;
  photo.crop = { x: 0, y: 0, zoom: 1, rotate: 0 };
  els.brightness.value = 0;
  els.contrast.value = 0;
  els.saturation.value = 0;
  els.sharpness.value = 0;
  els.temperature.value = 0;
  updateSliderLabels();
  loadPhotoIntoEditor(photo);
  saveHistory();
  saveLocalProject();
  toast("Edits reset");
}

function rotateActive(deg) {
  const photo = getActivePhoto();
  if (!photo || !state.cropper) return;
  state.cropper.rotate(deg);
  saveHistory();
  saveLocalProject();
}

function getSingleExportCanvas() {
  return buildSinglePreview();
}

function exportSheetPNG() {
  buildSheetPreview().then(canvas => {
    if (!canvas) return toast("Nothing to export");
    const url = canvas.toDataURL("image/png", 1.0);
    downloadDataUrl(url, "passport_sheet.png");
    toast("PNG downloaded");
  });
}

function exportSheetPDF() {
  buildSheetPreview().then(canvas => {
    if (!canvas) return toast("Nothing to export");
    const url = canvas.toDataURL("image/png", 1.0);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const imgProps = pdf.getImageProperties(url);
    const pdfW = 210;
    const pdfH = (imgProps.height * pdfW) / imgProps.width;
    pdf.addImage(url, "PNG", 0, 0, pdfW, pdfH, undefined, "FAST");
    pdf.save("passport_sheet.pdf");
    toast("PDF downloaded");
  });
}

function directPrint() {
  buildSheetPreview().then(canvas => {
    if (!canvas) return toast("Nothing to print");
    const url = canvas.toDataURL("image/png");
    const win = window.open("", "_blank");
    win.document.write(`
      <html><head><title>Print Passport Sheet</title>
      <style>
        @page{size:A4;margin:0}
        html,body{height:100%;margin:0}
        body{display:grid;place-items:center;background:#fff}
        img{width:210mm;height:auto;display:block}
      </style></head>
      <body>
        <img src="${url}" onload="setTimeout(()=>{window.print();}, 250)">
      </body></html>
    `);
    win.document.close();
  });
}

function downloadDataUrl(url, fileName) {
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
}

async function generatePreview() {
  saveStateFromUI();
  await updatePreviewDebounced();
  if (state.mode === "layout") layoutFabric?.renderAll();
  toast("Preview updated");
}

function localProjectKey() {
  return "pppm_nepal_cyber_studio_project_v1";
}

function saveLocalProject(showToast = true) {
  const project = serializeProject();
  localStorage.setItem(localProjectKey(), JSON.stringify(project));
  if (showToast) toast("Project saved locally");
}

function serializeProject() {
  return {
    version: 1,
    settings: { ...state.settings },
    theme: state.theme,
    photos: state.photos,
    activePhotoId: state.activePhotoId,
    layout: layoutFabric ? layoutFabric.toJSON(["objectCaching"]) : null,
  };
}

async function loadProjectFromObject(project) {
  if (!project || !project.photos) return;
  state.settings = { ...state.settings, ...(project.settings || {}) };
  applySettingsToUI();
  setTheme(project.theme || "dark");
  state.photos = project.photos || [];
  state.activePhotoId = project.activePhotoId || state.photos[0]?.id || null;
  renderPhotoList();
  if (state.activePhotoId) setActivePhoto(state.activePhotoId);
  if (layoutFabric && project.layout) {
    layoutFabric.clear();
    layoutFabric.loadFromJSON(project.layout, () => {
      enableLayoutGrid();
      layoutFabric.renderAll();
    });
  }
  updatePreviewDebounced();
  saveHistory();
  toast("Project loaded");
}

function downloadJSONProject() {
  const project = serializeProject();
  const blob = new Blob([JSON.stringify(project, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  downloadDataUrl(url, "passport_photo_project.json");
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function loadProjectFile(file) {
  const text = await file.text();
  const data = JSON.parse(text);
  await loadProjectFromObject(data);
}

function setCanvasSize(canvas, width, height) {
  canvas.width = width;
  canvas.height = height;
  canvas.style.aspectRatio = `${width} / ${height}`;
}

function initLayoutFabric() {
  const dpi = getDpi();
  const w = inToPx(8.27, dpi);
  const h = inToPx(11.69, dpi);
  layoutFabric = new fabric.Canvas("layoutCanvas", {
    preserveObjectStacking: true,
    selection: true,
    stopContextMenu: true,
    width: w,
    height: h,
    backgroundColor: "#ffffff",
  });
  layoutFabric.setDimensions({ width: w, height: h });
  layoutFabric.on("object:modified", () => {
    if (els.snapGridToggle.checked) snapActiveObject();
    saveLayoutState();
    saveHistory();
    saveLocalProject(false);
  });
  layoutFabric.on("object:moving", () => {
    if (els.snapGridToggle.checked) snapActiveObject();
  });
  layoutFabric.on("object:scaling", () => {
    if (els.snapGridToggle.checked) snapActiveObject();
  });
  layoutFabric.on("selection:created", e => {
    const obj = e.selected?.[0];
    if (obj) updateModeLabel("Layout");
  });
  enableLayoutGrid();
}

function enableLayoutGrid() {
  if (!layoutFabric) return;
  const dpi = getDpi();
  const step = mmToPx(10, dpi);
  const bg = new fabric.Rect({
    left: 0, top: 0, width: layoutFabric.width, height: layoutFabric.height,
    fill: "#ffffff", selectable: false, evented: false
  });
  layoutFabric.setBackgroundColor("#ffffff", layoutFabric.renderAll.bind(layoutFabric));
  layoutFabric.gridStep = step;
  layoutFabric.renderAll();
}

function snapActiveObject() {
  if (!layoutFabric) return;
  const obj = layoutFabric.getActiveObject();
  if (!obj) return;
  const step = layoutFabric.gridStep || 20;
  obj.set({
    left: Math.round(obj.left / step) * step,
    top: Math.round(obj.top / step) * step,
    scaleX: Math.max(0.1, Math.round(obj.scaleX * 20) / 20),
    scaleY: Math.max(0.1, Math.round(obj.scaleY * 20) / 20),
  });
  obj.setCoords();
}

function resizeLayoutCanvas() {
  if (!layoutFabric) return;
  const dpi = getDpi();
  const w = inToPx(8.27, dpi);
  const h = inToPx(11.69, dpi);
  layoutFabric.setWidth(w);
  layoutFabric.setHeight(h);
  layoutFabric.requestRenderAll();
  enableLayoutGrid();
  updateSizeInfo();
}

async function addCurrentToLayout() {
  const photo = getActivePhoto();
  if (!photo || !state.cropper) return toast("Select a photo first");
  const canvas = await buildSinglePreview();
  if (!canvas) return;
  const dataUrl = canvas.toDataURL("image/png");
  fabric.Image.fromURL(dataUrl, img => {
    img.set({
      left: 120,
      top: 120,
      scaleX: 0.5,
      scaleY: 0.5,
      cornerStyle: "circle",
      transparentCorners: false,
      borderColor: "#06b6d4",
      cornerColor: "#7c3aed",
      objectCaching: false,
    });
    layoutFabric.add(img);
    layoutFabric.setActiveObject(img);
    layoutFabric.requestRenderAll();
    saveLayoutState();
    saveHistory();
    saveLocalProject();
    toast("Added to layout");
  }, { crossOrigin: "anonymous" });
}

function saveLayoutState() {
  if (!layoutFabric) return;
  state.layout.objects = layoutFabric.toJSON();
}

function clearLayoutPage() {
  if (!layoutFabric) return;
  layoutFabric.clear();
  enableLayoutGrid();
  layoutFabric.requestRenderAll();
  saveHistory();
  saveLocalProject();
  toast("Layout cleared");
}

function refreshPreview() {
  saveStateFromUI();
  updatePreviewDebounced();
  toast("Preview refreshed");
}

function bindInputs() {
  const uiInputs = [
    els.sizePreset, els.customWidth, els.customHeight, els.sizeUnit, els.dpiSelect, els.customDpi,
    els.copiesSelect, els.customCopies, els.marginMm, els.spacingMm,
    els.brightness, els.contrast, els.saturation, els.sharpness, els.temperature,
    els.bgColorPreset, els.customBgColor, els.bgRemovalToggle, els.snapGridToggle, els.showSafeToggle
  ];
  uiInputs.forEach(el => el.addEventListener("input", () => {
    updateToggles();
    saveStateFromUI();
    if (["sizePreset", "customWidth", "customHeight", "sizeUnit", "dpiSelect", "customDpi", "copiesSelect", "customCopies", "marginMm", "spacingMm"].includes(el.id)) {
      if (el === els.sizePreset) applyPreset(els.sizePreset.value);
      if (layoutFabric && (el === els.dpiSelect || el === els.customDpi)) resizeLayoutCanvas();
    }
    if (["brightness", "contrast", "saturation", "sharpness", "temperature"].includes(el.id)) {
      updateSliderLabels();
    }
    updateSizeInfo();
    updatePreviewDebounced();
    saveLocalProject(false);
  }));

  els.bgColorPreset.addEventListener("change", () => {
    if (els.bgColorPreset.value !== "custom") els.customBgColor.value = els.bgColorPreset.value;
  });

  els.fileInput.addEventListener("change", e => addPhotosFromFiles(e.target.files));
  els.layoutImagePicker.addEventListener("change", e => addPhotosFromFiles(e.target.files));

  els.dropZone.addEventListener("dragover", e => {
    e.preventDefault();
    els.dropZone.classList.add("dragover");
  });
  els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragover"));
  els.dropZone.addEventListener("drop", e => {
    e.preventDefault();
    els.dropZone.classList.remove("dragover");
    addPhotosFromFiles(e.dataTransfer.files);
  });
  els.dropZone.addEventListener("click", e => {
    if (e.target.tagName !== "INPUT") els.fileInput.click();
  });

  $$(".nav-btn").forEach(btn => btn.addEventListener("click", () => showPanel(btn.dataset.panel)));
  $$(".tab-btn").forEach(btn => btn.addEventListener("click", () => setPreviewTab(btn.dataset.preview)));

  els.themeToggle.addEventListener("click", () => setTheme(state.theme === "dark" ? "light" : "dark"));
  els.generateBtn.addEventListener("click", generatePreview);
  els.printBtn.addEventListener("click", directPrint);
  els.loadProjectBtn.addEventListener("click", () => els.loadProjectInput.click());
  els.loadProjectInput.addEventListener("change", async e => {
    const file = e.target.files[0];
    if (file) await loadProjectFile(file);
  });
  els.saveProjectBtn.addEventListener("click", downloadJSONProject);
  els.downloadPngBtn.addEventListener("click", exportSheetPNG);
  els.downloadPdfBtn.addEventListener("click", exportSheetPDF);
  els.printPreviewBtn.addEventListener("click", directPrint);
  els.clearAllBtn.addEventListener("click", () => {
    if (!confirm("Remove all imported photos?")) return;
    state.photos = [];
    state.activePhotoId = null;
    destroyCropper();
    renderPhotoList();
    setActivePhoto(null);
    saveLocalProject();
    toast("All photos cleared");
  });
  els.refreshPreviewBtn.addEventListener("click", refreshPreview);
  els.addToLayoutBtn.addEventListener("click", addCurrentToLayout);
  els.clearLayoutBtn.addEventListener("click", clearLayoutPage);

  els.undoBtn.addEventListener("click", undo);
  els.redoBtn.addEventListener("click", redo);
  els.zoomInBtn.addEventListener("click", () => state.cropper?.zoom(0.1));
  els.zoomOutBtn.addEventListener("click", () => state.cropper?.zoom(-0.1));
  els.fitBtn.addEventListener("click", () => {
    if (!state.cropper) return;
    state.cropper.reset();
    toast("Fitted");
  });

  els.rotateLeftBtn.addEventListener("click", () => rotateActive(-90));
  els.rotateRightBtn.addEventListener("click", () => rotateActive(90));
  els.bgRemoveBtn.addEventListener("click", removeBackgroundForActive);
  els.autoEnhanceBtn.addEventListener("click", autoEnhanceActive);
  els.resetEditsBtn.addEventListener("click", resetEdits);

  els.sizePreset.addEventListener("change", () => {
    const preset = els.sizePreset.value;
    if (preset !== "custom") applyPreset(preset);
    updateToggles();
    updateSizeInfo();
  });

  document.addEventListener("keydown", e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") { e.preventDefault(); undo(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") { e.preventDefault(); redo(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") { e.preventDefault(); els.saveProjectBtn.click(); }
    if (e.key === "Delete" && layoutFabric) {
      const active = layoutFabric.getActiveObject();
      if (active) {
        layoutFabric.remove(active);
        layoutFabric.discardActiveObject();
        layoutFabric.requestRenderAll();
        saveLayoutState();
        saveHistory();
      }
    }
  });
}

function buildPresetCards() {
  els.presetCards.innerHTML = "";
  Object.entries(PRESETS).forEach(([key, p]) => {
    const card = document.createElement("div");
    card.className = "size-card";
    card.dataset.preset = key;
    card.innerHTML = `
      <div class="title">${p.label}</div>
      <div class="meta">${p.wIn.toFixed(2)} × ${p.hIn.toFixed(2)} in<br>${p.note}</div>
    `;
    card.addEventListener("click", () => {
      els.sizePreset.value = key;
      applyPreset(key);
      $$(".size-card").forEach(c => c.classList.toggle("active", c.dataset.preset === key));
    });
    els.presetCards.appendChild(card);
  });
  $(`.size-card[data-preset="${state.settings.preset}"]`)?.classList.add("active");
}

function syncBodyPixLazy() {
  // library will load only when requested
  bodyPixReady = !!window.bodyPix;
}

function fitPreviewCanvases() {
  const canvases = [els.sheetCanvas, els.layoutCanvas];
  canvases.forEach(c => {
    if (!c.width || !c.height) return;
    const parent = c.parentElement;
    if (!parent) return;
    const maxW = parent.clientWidth - 2;
    const maxH = Math.min(window.innerHeight * 0.72, 980);
    const ratio = Math.min(maxW / c.width, maxH / c.height, 1);
    c.style.width = `${Math.floor(c.width * ratio)}px`;
    c.style.height = `${Math.floor(c.height * ratio)}px`;
  });
}

function resizeObserverInit() {
  const ro = new ResizeObserver(() => {
    clearTimeout(resizeRAF);
    resizeRAF = setTimeout(fitPreviewCanvases, 80);
  });
  ro.observe(document.body);
}

async function init() {
  setTheme("dark");
  applySettingsToUI();
  bindInputs();
  buildPresetCards();
  initLayoutFabric();
  resizeObserverInit();
  syncBodyPixLazy();
  updateSizeInfo();

  const saved = localStorage.getItem(localProjectKey());
  if (saved) {
    try {
      await loadProjectFromObject(JSON.parse(saved));
      toast("Last project restored");
    } catch (e) {
      console.warn(e);
    }
  }

  renderPhotoList();
  setActivePhoto(state.activePhotoId);

  window.addEventListener("resize", fitPreviewCanvases);
  setPreviewTab("editorPreview");
  showPanel("uploadPanel");
  fitPreviewCanvases();
  toast("Ready");
}

init();
