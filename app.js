/* ===================== State ===================== */
const STORAGE_KEY = "tci_price_data_v1";
const SCAN_LOG_KEY = "tci_scan_log_v1";

let state = {
  items: [],
  fileName: "",
  updatedAt: null
};
let barcodeMap = new Map();
let styleMap = new Map();
let rawRows = null;
let scanHistory = [];
let scanLog = [];

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
const exportScansBtn = document.getElementById("exportScansBtn");
const settingsInfo = document.getElementById("settingsInfo");

const mapBarcode = document.getElementById("mapBarcode");
const mapStyle = document.getElementById("mapStyle");
const mapPrice = document.getElementById("mapPrice");
const mapDesc = document.getElementById("mapDesc");
const mapSilver = document.getElementById("mapSilver");
const mapWeight = document.getElementById("mapWeight");
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
const resultSilver = document.getElementById("resultSilver");
const resultSilverRow = document.getElementById("resultSilverRow");
const resultWeight = document.getElementById("resultWeight");
const resultWeightRow = document.getElementById("resultWeightRow");
const historyEl = document.getElementById("history");

const cameraModal = document.getElementById("cameraModal");
const video = document.getElementById("video");
const closeCameraBtn = document.getElementById("closeCameraBtn");

const toastEl = document.getElementById("toast");

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
  } catch (e) {}
}

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

function loadScanLog() {
  try {
    const raw = localStorage.getItem(SCAN_LOG_KEY);
    scanLog = raw ? JSON.parse(raw) : [];
  } catch (e) {
    scanLog = [];
  }
}

function saveScanLog() {
  localStorage.setItem(SCAN_LOG_KEY, JSON.stringify(scanLog));
}

function logScan(item) {
  scanLog.push({
    time: new Date().toISOString(),
    barcode: item.barcode || "",
    style: item.style || "",
    price: item.price ?? "",
    desc: item.desc || ""
  });
  saveScanLog();
}

function buildMaps() {
  barcodeMap = new Map();
  styleMap = new Map();
  for (const item of state.items) {
    if (item.barcode) barcodeMap.set(normalizeBarcode(item.barcode), item);
    if (item.style) styleMap.set(normalizeStyle(item.style), item);
  }
}

function init() {
  loadScanLog();
  if (loadData()) {
    renderMainStats();
    showScreen("main");
  } else {
    showScreen("empty");
  }
}
init();

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
  e.target.value = "";
}

let pendingFileName = "";

function openMappingScreen(rows) {
  const headers = rows[0].map(h => String(h ?? "").trim() || "(blank)");
  const dataRows = rows.slice(1, 6);

  [mapBarcode, mapStyle, mapPrice, mapDesc, mapSilver, mapWeight].forEach(sel => (sel.innerHTML = ""));

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
  optionsFor(mapSilver, true);
  optionsFor(mapWeight, true);

  const guess = (patterns, exclude = []) => headers.findIndex(h => {
    const lh = h.toLowerCase();
    if (exclude.some(x => lh.includes(x))) return false;
    return patterns.some(p => lh.includes(p));
  });
  const bIdx = guess(["barcode", "upc", "ean", "sku"]);
  const sIdx = guess(["style", "item no", "item number", "model"]);
  const silverIdxGuess = guess(["silver"]);
  const weightIdxGuess = guess(["weight"]);
  const pIdx = guess(["price", "cost", "wholesale", "retail"], ["silver"]);
  const dIdx = guess(["desc", "name", "title", "product_type", "product type", "type"]);
  if (bIdx >= 0) mapBarcode.value = String(bIdx);
  if (sIdx >= 0) mapStyle.value = String(sIdx);
  else mapStyle.value = "0";
  if (pIdx >= 0) mapPrice.value = String(pIdx);
  mapDesc.value = dIdx >= 0 ? String(dIdx) : "-1";
  mapSilver.value = silverIdxGuess >= 0 ? String(silverIdxGuess) : "-1";
  mapWeight.value = weightIdxGuess >= 0 ? String(weightIdxGuess) : "-1";

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
  const silverIdx = Number(mapSilver.value);
  const weightIdx = Number(mapWeight.value);

  if (bIdx === sIdx || bIdx === pIdx || sIdx === pIdx) {
    mappingError.textContent = "Barcode, style, and price should each be different columns.";
    mappingError.hidden = false;
    return;
  }

  const dataRows = rawRows.slice(1);
  const items = [];
  for (const row of dataRows) {
    const barcode = normalizeBarcode(row[bIdx]);
    if (!barcode) continue;
    items.push({
      barcode,
      style: String(row[sIdx] ?? "").trim(),
      price: row[pIdx],
      desc: dIdx >= 0 ? String(row[dIdx] ?? "").trim() : "",
      silver: silverIdx >= 0 ? row[silverIdx] : "",
      weight: weightIdx >= 0 ? row[weightIdx] : ""
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

settingsBtn.addEventListener("click", () => {
  const fileInfo = state.items.length
    ? `${state.items.length} items loaded from "${state.fileName}", last updated ${new Date(state.updatedAt).toLocaleString()}.`
    : "No pricing file loaded yet.";
  const scanInfo = scanLog.length
    ? ` ${scanLog.length} scanned item${scanLog.length === 1 ? "" : "s"} ready to export.`
    : " No items scanned yet.";
  settingsInfo.textContent = fileInfo + scanInfo;
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

exportScansBtn.addEventListener("click", async () => {
  if (!scanLog.length) {
    toast("No scanned items yet.");
    return;
  }
  try {
    const rows = scanLog.map(s => ({
      Barcode: s.barcode,
      Style: s.style,
      Price: s.price,
      Description: s.desc,
      "Scanned At": new Date(s.time).toLocaleString()
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Scanned Items");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    const filename = `scanned-items-${stamp}.xlsx`;
    const file = new File([blob], filename, { type: blob.type });

    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({
        files: [file],
        title: "Scanned items",
        text: `${scanLog.length} scanned item${scanLog.length === 1 ? "" : "s"} — attach this file to your email.`
      });
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      toast("Downloaded — attach it to an email yourself.");
    }
  } catch (e) {
    console.error(e);
    toast("Couldn't export right now.");
  }
});

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
    logScan(item);
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
    if (item.silver) {
      resultSilver.textContent = formatPrice(item.silver);
      resultSilverRow.style.display = "flex";
    } else {
      resultSilverRow.style.display = "none";
    }
    if (item.weight) {
      resultWeight.textContent = String(item.weight);
      resultWeightRow.style.display = "flex";
    } else {
      resultWeightRow.style.display = "none";
    }
  } else {
    resultStatus.textContent = "not found";
    resultStatus.classList.add("not-found");
    resultPrice.textContent = "—";
    resultStyle.textContent = "—";
    resultBarcode.textContent = query;
    resultDescRow.style.display = "none";
    resultSilverRow.style.display = "none";
    resultWeightRow.style.display = "none";
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
    } catch (e) {}
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

const modeTabPrice = document.getElementById("modeTabPrice");
const modeTabCards = document.getElementById("modeTabCards");
const priceModeEl = document.getElementById("priceMode");
const cardModeEl = document.getElementById("cardMode");

modeTabPrice.addEventListener("click", () => switchMode("price"));
modeTabCards.addEventListener("click", () => switchMode("cards"));

function switchMode(mode) {
  const toCards = mode === "cards";
  priceModeEl.hidden = toCards;
  cardModeEl.hidden = !toCards;
  modeTabPrice.classList.toggle("active", !toCards);
  modeTabCards.classList.toggle("active", toCards);
  if (toCards) {
    showCardScreen("screen-card-list");
    renderCardList();
  }
}

const CARDS_KEY = "tci_cards_v1";
let cards = [];
let editingCardId = null;

function loadCards() {
  try {
    const raw = localStorage.getItem(CARDS_KEY);
    cards = raw ? JSON.parse(raw) : [];
  } catch (e) { cards = []; }
}
function saveCards() {
  localStorage.setItem(CARDS_KEY, JSON.stringify(cards));
}
loadCards();

const cardScreens = {
  "screen-card-list": document.getElementById("screen-card-list"),
  "screen-card-review": document.getElementById("screen-card-review"),
  "screen-card-detail": document.getElementById("screen-card-detail"),
};
function showCardScreen(name) {
  Object.values(cardScreens).forEach(s => s.classList.remove("active"));
  cardScreens[name].classList.add("active");
}

function renderCardList() {
  const listEl = document.getElementById("cardList");
  const emptyEl = document.getElementById("cardListEmpty");
  if (!cards.length) {
    emptyEl.hidden = false;
    listEl.innerHTML = "";
    return;
  }
  emptyEl.hidden = true;
  listEl.innerHTML = cards.slice().reverse().map(c => `
    <div class="card-list-item" data-id="${c.id}">
      <span class="c-name">${escapeHtml(c.name || "(no name)")}</span>
      <span class="c-sub">${escapeHtml(c.company || c.phone || c.email || "")}</span>
    </div>
  `).join("");
  listEl.querySelectorAll(".card-list-item").forEach(el => {
    el.addEventListener("click", () => openCardDetail(el.dataset.id));
  });
}

const scanCardBtn = document.getElementById("scanCardBtn");
const cardCaptureModal = document.getElementById("cardCaptureModal");
const cardVideo = document.getElementById("cardVideo");
const cardCanvas = document.getElementById("cardCanvas");
const cardCaptureBtn = document.getElementById("cardCaptureBtn");
const closeCardCameraBtn = document.getElementById("closeCardCameraBtn");
const cardProcessing = document.getElementById("cardProcessing");
const cardProcessingText = document.getElementById("cardProcessingText");
let cardStream = null;
let lastCardPhotoDataUrl = null;

scanCardBtn.addEventListener("click", openCardCamera);
closeCardCameraBtn.addEventListener("click", closeCardCamera);
cardCaptureBtn.addEventListener("click", captureCardPhoto);

async function openCardCamera() {
  cardCaptureModal.hidden = false;
  try {
    cardStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment", width: { ideal: 1920 }, height: { ideal: 1080 } } });
    cardVideo.srcObject = cardStream;
    await cardVideo.play();
  } catch (err) {
    console.error(err);
    toast("Couldn't access the camera. Check camera permission in your browser settings.");
    closeCardCamera();
  }
}

function closeCardCamera() {
  cardCaptureModal.hidden = true;
  if (cardStream) { cardStream.getTracks().forEach(t => t.stop()); cardStream = null; }
  if (cardVideo.srcObject) { cardVideo.srcObject = null; }
}

async function captureCardPhoto() {
  const w = cardVideo.videoWidth, h = cardVideo.videoHeight;
  if (!w || !h) return;
  cardCanvas.width = w;
  cardCanvas.height = h;
  const ctx = cardCanvas.getContext("2d");
  ctx.drawImage(cardVideo, 0, 0, w, h);
  lastCardPhotoDataUrl = cardCanvas.toDataURL("image/jpeg", 0.92);
  closeCardCamera();
  await runCardOcr(lastCardPhotoDataUrl);
}

async function runCardOcr(photoDataUrl) {
  cardProcessing.hidden = false;
  cardProcessingText.textContent = "Reading the card…";
  try {
    const result = await Tesseract.recognize(photoDataUrl, "eng", {
      logger: (m) => {
        if (m.status === "recognizing text") {
          cardProcessingText.textContent = `Reading the card… ${Math.round((m.progress || 0) * 100)}%`;
        }
      }
    });
    const text = result?.data?.text || "";
    const parsed = parseCardText(text);
    openCardReview(parsed, photoDataUrl);
  } catch (err) {
    console.error(err);
    toast("Couldn't read the card. You can still enter details manually.");
    openCardReview({ name: "", company: "", title: "", phone: "", email: "", notes: "" }, photoDataUrl);
  } finally {
    cardProcessing.hidden = true;
  }
}

function parseCardText(text) {
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
  const emailMatch = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
  const phoneMatch = text.match(/(\+?\d[\d\s().-]{7,}\d)/);
  const email = emailMatch ? emailMatch[0] : "";
  const phone = phoneMatch ? phoneMatch[0].trim() : "";

  const remaining = lines.filter(l => l !== email && l !== phone && !l.includes(email) && (!phone || !l.includes(phone)));

  const name = remaining[0] || "";
  const title = remaining[1] || "";
  const company = remaining[2] || remaining[1] || "";

  return { name, company, title, phone, email, notes: "" };
}

const cardPhotoPreview = document.getElementById("cardPhotoPreview");
const cardNameInput = document.getElementById("cardName");
const cardCompanyInput = document.getElementById("cardCompany");
const cardTitleInput = document.getElementById("cardTitle");
const cardPhoneInput = document.getElementById("cardPhone");
const cardEmailInput = document.getElementById("cardEmail");
const cardNotesInput = document.getElementById("cardNotes");
const cardDiscardBtn = document.getElementById("cardDiscardBtn");
const cardSaveBtn = document.getElementById("cardSaveBtn");

function openCardReview(parsed, photoDataUrl) {
  cardPhotoPreview.innerHTML = photoDataUrl ? `<img src="${photoDataUrl}" alt="Captured card">` : "";
  cardNameInput.value = parsed.name || "";
  cardCompanyInput.value = parsed.company || "";
  cardTitleInput.value = parsed.title || "";
  cardPhoneInput.value = parsed.phone || "";
  cardEmailInput.value = parsed.email || "";
  cardNotesInput.value = parsed.notes || "";
  editingCardId = null;
  showCardScreen("screen-card-review");
}

cardDiscardBtn.addEventListener("click", () => {
  showCardScreen("screen-card-list");
  renderCardList();
});

cardSaveBtn.addEventListener("click", () => {
  const card = {
    id: editingCardId || ("card_" + Date.now()),
    name: cardNameInput.value.trim(),
    company: cardCompanyInput.value.trim(),
    title: cardTitleInput.value.trim(),
    phone: cardPhoneInput.value.trim(),
    email: cardEmailInput.value.trim(),
    notes: cardNotesInput.value.trim(),
    photo: lastCardPhotoDataUrl,
    savedAt: new Date().toISOString()
  };
  if (!card.name && !card.company && !card.phone && !card.email) {
    toast("Add at least a name, company, phone, or email.");
    return;
  }
  const idx = cards.findIndex(c => c.id === card.id);
  if (idx >= 0) cards[idx] = card; else cards.push(card);
  saveCards();
  showCardScreen("screen-card-list");
  renderCardList();
  toast("Card saved");
});

const cardDetailBody = document.getElementById("cardDetailBody");
const cardDetailBackBtn = document.getElementById("cardDetailBackBtn");
const cardDetailShareBtn = document.getElementById("cardDetailShareBtn");
const cardDetailDeleteBtn = document.getElementById("cardDetailDeleteBtn");
let detailCardId = null;

function openCardDetail(id) {
  const card = cards.find(c => c.id === id);
  if (!card) return;
  detailCardId = id;
  cardDetailBody.innerHTML = `
    ${card.photo ? `<div class="card-photo-preview"><img src="${card.photo}" alt="Card photo"></div>` : ""}
    <div class="card-detail-name">${escapeHtml(card.name || "(no name)")}</div>
    ${card.company ? `<div class="card-detail-company">${escapeHtml(card.company)}${card.title ? " — " + escapeHtml(card.title) : ""}</div>` : ""}
    ${card.phone ? `<div class="card-detail-row"><span class="d-label">Phone</span><span class="d-value">${escapeHtml(card.phone)}</span></div>` : ""}
    ${card.email ? `<div class="card-detail-row"><span class="d-label">Email</span><span class="d-value">${escapeHtml(card.email)}</span></div>` : ""}
    ${card.notes ? `<div class="card-detail-row"><span class="d-label">Notes</span><span class="d-value">${escapeHtml(card.notes)}</span></div>` : ""}
  `;
  showCardScreen("screen-card-detail");
}

cardDetailBackBtn.addEventListener("click", () => {
  showCardScreen("screen-card-list");
  renderCardList();
});

cardDetailDeleteBtn.addEventListener("click", () => {
  if (!confirm("Delete this business card?")) return;
  cards = cards.filter(c => c.id !== detailCardId);
  saveCards();
  showCardScreen("screen-card-list");
  renderCardList();
});

cardDetailShareBtn.addEventListener("click", async () => {
  const card = cards.find(c => c.id === detailCardId);
  if (!card) return;
  const vcard = buildVCard(card);
  const file = new File([vcard], (card.name || "contact") + ".vcf", { type: "text/vcard" });
  if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: card.name || "Contact" });
      return;
    } catch (e) {}
  }
  const url = URL.createObjectURL(file);
  const a = document.createElement("a");
  a.href = url;
  a.download = file.name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

function buildVCard(card) {
  return [
    "BEGIN:VCARD",
    "VERSION:3.0",
    `FN:${card.name || ""}`,
    card.company ? `ORG:${card.company}` : "",
    card.title ? `TITLE:${card.title}` : "",
    card.phone ? `TEL;TYPE=CELL:${card.phone}` : "",
    card.email ? `EMAIL:${card.email}` : "",
    card.notes ? `NOTE:${card.notes}` : "",
    "END:VCARD"
  ].filter(Boolean).join("\n");
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  });
}
