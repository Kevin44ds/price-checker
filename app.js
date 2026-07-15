/* ===================== State ===================== */
const STORAGE_KEY = "tci_price_data_v1";

let state = {
  items: [],          // [{barcode, style, price, desc}]
  fileName: "",
  updatedAt: null
};
let barcodeMap = new Map();
let styleMap = new Map();
let rawRows = null;     // rows parsed from the currently-uploading file (header:1 format)
let scanHistory = [];

/* ===================== Elements ===================== */
const screens = {
  empty: document.getElementById("screen-empty"),
  mapping: document.getElementById("screen-mapping"),
  main: document.getElementById("screen-main"),
};
const fileInput = document.getElementById("fileInput");
const settingsBtn = document.getElementById("settingsBtn");
const settingsModal = document.getElementById("settingsModal");
const closeSettingsBtn = document.getElementById("closeSettingsBtn");
const replaceFileBtn = document.getElementById("replaceFileBtn");
const clearDataBtn = document.getElementById("clearDataBtn");
const settingsInfo = document.getElementById("settingsInfo");

const mapBarcode = document.getElementById("mapBarcode");
const mapStyle = document.getElementById("mapStyle");
const mapPrice = document.getElementById("mapPrice");
const mapDesc = document.getElementById("mapDesc");
const previewTable = document.getElementById("previewTable");
const confirmMapping = document.getElementById("confirmMapping");
const cancelMapping = document.getElementById("cancelMapping");
const mappingError = document.getElementById("mappingError");

const itemCount = document.getElementById("itemCount");
const updatedAtEl = document.getElementById("updatedAt");
const scanBtn = document.getElementById("scanBtn");
const searchInput = document.getElementById("searchInput");
const searchBtn = document.getElementById("searchBtn");
const resultCard = document.getElementById("result");
const resultStatus = document.getElementById("resultStatus");
const resultPrice = document.getElementById("resultPrice");
const resultStyle = document.getElementById("resultStyle");
const resultBarcode = document.getElementById("resultBarcode");
const resultDesc = document.getElementById("resultDesc");
const resultDescRow = document.getElementById("resultDescRow");
const historyEl = document.getElementById("history");

const cameraModal = document.getElementById("cameraModal");
const video = document.getElementById("video");
const closeCameraBtn = document.getElementById("closeCameraBtn");

const toastEl = document.getElementById("toast");

/* ===================== Utilities ===================== */
function showScreen(name) {
  Object.values(screens).forEach(s => s.classList.remove("active"));
  screens[name].classList.add("active");
}

function toast(msg, ms = 2200) {
  toastEl.textContent = msg;
  toastEl.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (toastEl.hidden = true), ms);
}

function normalizeBarcode(v) {
  return String(v ?? "").trim().replace(/\s+/g, "");
}
function normalizeStyle(v) {
  return String(v ?? "").trim().toUpperCase();
}
function formatPrice(v) {
  const n = Number(String(v).replace(/[^0-9.\-]/g, ""));
  if (!isNaN(n) && String(v).match(/[0-9]/)) {
    return "$" + n.toFixed(2);
  }
  return String(v ?? "—");
}

function beep() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.type = "sine";
    o.frequency.value = 1000;
    g.gain.setValueAtTime(0.15, ctx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.15);
    o.connect(g); g.connect(ctx.destination);
    o.start(); o.stop(ctx.currentTime + 0.15);
  } catch (e) { /* audio not available, ignore */ }
}

/* ===================== Persistence ===================== */
function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (!parsed.items || !parsed.items.length) return false;
    state = parsed;
    buildMaps();
    return true;
  } catch (e) {
    console.error("Failed to load stored data", e);
    return false;
  }
}

function saveData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function buildMaps() {
  barcodeMap = new Map();
  styleMap = new Map();
  for (const item of state.items) {
    if (item.barcode) barcodeMap.set(normalizeBarcode(item.barcode), item);
    if (item.style) styleMap.set(normalizeStyle(item.style), item);
  }
}

/* ===================== Init ===================== */
function init() {
  if (loadData()) {
    renderMainStats();
    showScreen("main");
  } else {
    showScreen("empty");
  }
}
init();

/* ===================== File upload -> mapping ===================== */
fileInput.addEventListener("change", handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
      if (!rows.length) {
        toast("That file looks empty.");
        return;
      }
      rawRows = rows;
      pendingFileName = file.name;
      openMappingScreen(rows);
    } catch (err) {
      console.error(err);
      toast("Couldn't read that file. Try exporting as .xlsx or .csv.");
    }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = ""; // allow re-selecting the same file later
}

let pendingFileName = "";

function openMappingScreen(rows) {
  const headers = rows[0].map(h => String(h ?? "").trim() || "(blank)");
  const dataRows = rows.slice(1, 6); // preview first 5

  [mapBarcode, mapStyle, mapPrice, mapDesc].forEach(sel => (sel.innerHTML = ""));

  const optionsFor = (sel, includeNone) => {
    if (includeNone) {
      const opt = document.createElement("option");
      opt.value = "-1";
      opt.textContent = "None";
      sel.appendChild(opt);
    }
    headers.forEach((h, i) => {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = h;
      sel.appendChild(opt);
    });
  };
  optionsFor(mapBarcode, false);
  optionsFor(mapStyle, false);
  optionsFor(mapPrice, false);
  optionsFor(mapDesc, true);

  // best-effort auto-detect by header name
  const guess = (patterns) => headers.findIndex(h => patterns.some(p => h.toLowerCase().includes(p)));
  const bIdx = guess(["barcode", "upc", "ean", "sku"]);
  const sIdx = guess(["style", "item no", "item number", "model"]);
  const pIdx = guess(["price", "cost", "wholesale", "retail"]);
  const dIdx = guess(["desc", "name", "title"]);
  if (bIdx >= 0) mapBarcode.value = String(bIdx);
  if (sIdx >= 0) mapStyle.value = String(sIdx);
  else mapStyle.value = "0";
  if (pIdx >= 0) mapPrice.value = String(pIdx);
  mapDesc.value = dIdx >= 0 ? String(dIdx) : "-1";

  // preview table
  const thead = previewTable.querySelector("thead");
  const tbody = previewTable.querySelector("tbody");
  thead.innerHTML = "<tr>" + headers.map(h => `<th>${escapeHtml(h)}</th>`).join("") + "</tr>";
  tbody.innerHTML = dataRows.map(r =>
    "<tr>" + headers.map((_, i) => `<td>${escapeHtml(String(r[i] ?? ""))}</td>`).join("") + "</tr>"
  ).join("");

  mappingError.hidden = true;
  showScreen("mapping");
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

cancelMapping.addEventListener("click", () => {
  rawRows = null;
  if (state.items.length) showScreen("main"); else showScreen("empty");
});

confirmMapping.addEventListener("click", () => {
  const bIdx = Number(mapBarcode.value);
  const sIdx = Number(mapStyle.value);
  const pIdx = Number(mapPrice.value);
  const dIdx = Number(mapDesc.value);

  if (bIdx === sIdx || bIdx === pIdx || sIdx === pIdx) {
    mappingError.textContent = "Barcode, style, and price should each be different columns.";
    mappingError.hidden = false;
    return;
  }

  const dataRows = rawRows.slice(1);
  const items = [];
  for (const row of dataRows) {
    const barcode = normalizeBarcode(row[bIdx]);
    if (!barcode) continue; // skip rows without a barcode
    items.push({
      barcode,
      style: String(row[sIdx] ?? "").trim(),
      price: row[pIdx],
      desc: dIdx >= 0 ? String(row[dIdx] ?? "").trim() : ""
    });
  }

  if (!items.length) {
    mappingError.textContent = "No rows had a value in the barcode column you picked.";
    mappingError.hidden = false;
    return;
  }

  state = {
    items,
    fileName: pendingFileName,
    updatedAt: new Date().toISOString()
  };
  saveData();
  buildMaps();
  renderMainStats();
  rawRows = null;
  hideSettings();
  showScreen("main");
  toast(`Loaded ${items.length} items`);
});

/* ===================== Settings ===================== */
settingsBtn.addEventListener("click", () => {
  settingsInfo.textContent = state.items.length
    ? `${state.items.length} items loaded from "${state.fileName}", last updated ${new Date(state.updatedAt).toLocaleString()}.`
    : "No pricing file loaded yet.";
  settingsModal.hidden = false;
});
closeSettingsBtn.addEventListener("click", hideSettings);
function hideSettings() { settingsModal.hidden = true; }

replaceFileBtn.addEventListener("click", () => {
  hideSettings();
  fileInput.click();
});

clearDataBtn.addEventListener("click", () => {
  if (!confirm("Clear the loaded pricing file? You'll need to upload it again to keep scanning.")) return;
  localStorage.removeItem(STORAGE_KEY);
  state = { items: [], fileName: "", updatedAt: null };
  barcodeMap = new Map();
  styleMap = new Map();
  scanHistory = [];
  historyEl.innerHTML = "";
  resultCard.hidden = true;
  hideSettings();
  showScreen("empty");
});

/* ===================== Main screen ===================== */
function renderMainStats() {
  itemCount.textContent = `${state.items.length} item${state.items.length === 1 ? "" : "s"} loaded`;
  updatedAtEl.textContent = state.updatedAt ? `updated ${new Date(state.updatedAt).toLocaleDateString()}` : "";
}

searchBtn.addEventListener("click", runSearch);
searchInput.addEventListener("keydown", (e) => { if (e.key === "Enter") runSearch(); });

function runSearch() {
  const q = searchInput.value.trim();
  if (!q) return;
  lookup(q);
  searchInput.value = "";
  searchInput.blur();
}

function lookup(query) {
  const byBarcode = barcodeMap.get(normalizeBarcode(query));
  const byStyle = styleMap.get(normalizeStyle(query));
  const item = byBarcode || byStyle;
  showResult(item, query);
  if (item) {
    beep();
    if (navigator.vibrate) navigator.vibrate(60);
    scanHistory.unshift(item);
    scanHistory = scanHistory.slice(0, 6);
    renderHistory();
  }
}

function showResult(item, query) {
  resultCard.hidden = false;
  if (item) {
    resultStatus.textContent = "match found";
    resultStatus.classList.remove("not-found");
    resultPrice.textContent = formatPrice(item.price);
    resultStyle.textContent = item.style || "—";
    resultBarcode.textContent = item.barcode || "—";
    if (item.desc) {
      resultDesc.textContent = item.desc;
      resultDescRow.style.display = "flex";
    } else {
      resultDescRow.style.display = "none";
    }
  } else {
    resultStatus.textContent = "not found";
    resultStatus.classList.add("not-found");
    resultPrice.textContent = "—";
    resultStyle.textContent = "—";
    resultBarcode.textContent = query;
    resultDescRow.style.display = "none";
  }
}

function renderHistory() {
  historyEl.innerHTML = scanHistory.map(it => `
    <div class="history-item">
      <span class="h-style">${escapeHtml(it.style || it.barcode)}</span>
      <span class="h-price">${formatPrice(it.price)}</span>
    </div>
  `).join("");
}

/* ===================== Camera scanning ===================== */
let zxingReader = null;
let usingNativeDetector = false;
let nativeDetectLoop = null;
let mediaStream = null;

scanBtn.addEventListener("click", openCamera);
closeCameraBtn.addEventListener("click", closeCamera);

async function openCamera() {
  if (!state.items.length) {
    toast("Load a pricing file first.");
    return;
  }
  cameraModal.hidden = false;
  try {
    if ("BarcodeDetector" in window) {
      usingNativeDetector = true;
      await startNativeDetector();
    } else {
      usingNativeDetector = false;
      await startZXing();
    }
  } catch (err) {
    console.error(err);
    toast("Couldn't access the camera. Check camera permission in your browser settings.");
    closeCamera();
  }
}

async function startNativeDetector() {
  mediaStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
  video.srcObject = mediaStream;
  await video.play();
  const detector = new BarcodeDetector({
    formats: ["ean_13", "ean_8", "upc_a", "upc_e", "code_128", "code_39", "qr_code", "itf"]
  });
  const tick = async () => {
    if (cameraModal.hidden) return;
    try {
      const codes = await detector.detect(video);
      if (codes.length) {
        onScanSuccess(codes[0].rawValue);
        return;
      }
    } catch (e) { /* keep trying */ }
    nativeDetectLoop = requestAnimationFrame(tick);
  };
  nativeDetectLoop = requestAnimationFrame(tick);
}

async function startZXing() {
  zxingReader = new ZXing.BrowserMultiFormatReader();
  const constraints = { video: { facingMode: "environment" } };
  await zxingReader.decodeFromConstraints(constraints, video, (result, err) => {
    if (result) onScanSuccess(result.getText());
  });
}

function onScanSuccess(text) {
  closeCamera();
  lookup(text);
}

function closeCamera() {
  cameraModal.hidden = true;
  if (nativeDetectLoop) { cancelAnimationFrame(nativeDetectLoop); nativeDetectLoop = null; }
  if (zxingReader) { try { zxingReader.reset(); } catch (e) {} zxingReader = null; }
  if (mediaStream) { mediaStream.getTracks().forEach(t => t.stop()); mediaStream = null; }
  if (video.srcObject) { video.srcObject.getTracks?.().forEach(t => t.stop()); video.srcObject = null; }
}

/* ===================== Service worker ===================== */
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  });
}
