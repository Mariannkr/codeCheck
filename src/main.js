import './style.css';
import { Html5Qrcode, Html5QrcodeSupportedFormats as F } from 'html5-qrcode';
import * as XLSX from 'xlsx';

/* =========================================================
   Constantes y Estado
   ========================================================= */
const COLS = ['CODIGO', 'FAMILIA', 'ARTICULO', 'STOCK_DIS', 'STOCK_IND', 'STOCK_REM',
  'STOCK_RES', 'STOCK_TRA', 'STOCK_EXH', 'TOTAL_TRA', 'FISICO', 'DIFEREN'];
const LS_BASE = 'stockscan.base.v1';
const LS_SCANS = 'stockscan.scans.v1';

const FORMATS = [F.CODE_128, F.CODE_39, F.CODE_93, F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E,
  F.ITF, F.CODABAR, F.QR_CODE, F.DATA_MATRIX];

const state = {
  base: null,          // { fileName, loadedAt, rows: [] }
  index: new Map(),    // CODIGO exacto norm -> row
  looseIndex: new Map(), // CODIGO sin guiones/símbolos -> row
  scans: [],           // [{ code, row|null, found, time, count }]
  listViewMode: 'scanned', // 'scanned' | 'base'
};

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = (s) => String(s ?? '').trim().toUpperCase();
const loose = (s) => norm(s).replace(/[^A-Z0-9]/g, '');

/* =========================================================
   Persistencia local (localStorage)
   ========================================================= */
function saveScans() {
  try {
    localStorage.setItem(LS_SCANS, JSON.stringify(state.scans));
  } catch { toast('No se pudo guardar el conteo en el celular', 'err'); }
}

function saveBase() {
  try {
    if (state.base) localStorage.setItem(LS_BASE, JSON.stringify(state.base));
    else localStorage.removeItem(LS_BASE);
  } catch { toast('La base es grande; funcionará durante la sesión actual', 'warn'); }
}

function loadSavedData() {
  try {
    const b = JSON.parse(localStorage.getItem(LS_BASE) || 'null');
    if (b?.rows) setBase(b, false);
    state.scans = JSON.parse(localStorage.getItem(LS_SCANS) || '[]');
  } catch { state.scans = []; }
}

/* =========================================================
   Manejo de Base de Stock (Excel importado)
   ========================================================= */
function setBase(base, persist = true) {
  state.base = base;
  state.index.clear();
  state.looseIndex.clear();

  for (const r of base?.rows || []) {
    state.index.set(norm(r.CODIGO), r);
    state.looseIndex.set(loose(r.CODIGO), r);
  }

  if (persist) saveBase();

  // Re-vincular lecturas existentes con la nueva base
  for (const s of state.scans) {
    const r = lookup(s.code);
    s.row = r;
    s.found = !!r;
  }

  saveScans();
  renderBaseState();
  renderList();
}

function lookup(code) {
  return state.index.get(norm(code)) || state.looseIndex.get(loose(code)) || null;
}

async function parseBaseFile(fileOrBuffer, fileName = 'INVENTARIO.xlsx') {
  const data = fileOrBuffer instanceof File ? await fileOrBuffer.arrayBuffer() : fileOrBuffer;
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

  const headerKey = (h) => norm(h).replace(/\s+/g, '_');
  const hIdx = aoa.findIndex((row) => row.some((c) => headerKey(c) === 'CODIGO'));
  if (hIdx < 0) throw new Error('No se encontró la columna CODIGO en la primera hoja.');

  const headers = aoa[hIdx].map(headerKey);
  const colIndex = Object.fromEntries(COLS.map((c) => [c, headers.indexOf(c)]));
  const rows = [];

  for (const line of aoa.slice(hIdx + 1)) {
    const code = String(line[colIndex.CODIGO] ?? '').trim();
    if (!code) continue;

    const r = {};
    for (const c of COLS) {
      r[c] = colIndex[c] >= 0 ? line[colIndex[c]] : '';
    }
    r.CODIGO = code;
    rows.push(r);
  }

  if (!rows.length) throw new Error('El archivo Excel no contiene filas de datos.');
  return { fileName: fileOrBuffer.name || fileName, loadedAt: Date.now(), rows };
}

async function loadDefaultSampleBase() {
  try {
    const res = await fetch('./INVENTARIO BABLES.xlsx');
    if (!res.ok) return;
    const buf = await res.arrayBuffer();
    const base = await parseBaseFile(buf, 'INVENTARIO BABLES.xlsx');
    setBase(base, true);
    toast('Se cargó la base predeterminada INVENTARIO BABLES.xlsx', 'ok');
  } catch (err) {
    console.log('No se pudo cargar la base por defecto:', err);
  }
}

/* =========================================================
   Cálculos del Conteo Físico (FISICO & DIFEREN)
   ========================================================= */
function getScanCountMap() {
  const map = new Map();
  for (const s of state.scans) {
    const key = s.row ? norm(s.row.CODIGO) : norm(s.code);
    map.set(key, (map.get(key) || 0) + (s.count || 1));
  }
  return map;
}

function calculateItemStats(code, row, scanEntry = null) {
  const scanMap = getScanCountMap();
  const key = row ? norm(row.CODIGO) : norm(code);
  const fisico = scanEntry ? scanEntry.count : (scanMap.get(key) || 0);

  const totalTraRaw = row ? row.TOTAL_TRA : 0;
  const totalTra = Number(totalTraRaw || 0);
  const diferenVal = fisico - totalTra;

  let diferenFmt = '0';
  if (diferenVal > 0) diferenFmt = `+${diferenVal}`;
  else if (diferenVal < 0) diferenFmt = `${diferenVal}`;

  return { fisico, totalTra, diferenVal, diferenFmt };
}

/* =========================================================
   Registro de Lectura
   ========================================================= */
function registerCode(rawCode, source) {
  const code = String(rawCode || '').trim();
  if (!code) return null;

  const row = lookup(code);
  const key = row ? norm(row.CODIGO) : norm(code);
  const existing = state.scans.find((s) => norm(s.row?.CODIGO ?? s.code) === key);

  let result;
  if (existing) {
    existing.count = (existing.count || 1) + 1;
    existing.time = Date.now();
    // Mover a la cima de la lista
    state.scans = [existing, ...state.scans.filter((s) => s !== existing)];
    result = { ...existing, dup: true };
    feedback('dup');

    const stats = calculateItemStats(code, row, existing);
    toast(`${row ? row.CODIGO : code} → FISICO: ${stats.fisico} (DIFEREN: ${stats.diferenFmt})`, 'warn');
  } else {
    const entry = { code: row ? row.CODIGO : code, row, found: !!row, time: Date.now(), count: 1, source };
    state.scans.unshift(entry);
    result = entry;
    feedback(row ? 'ok' : 'miss');

    const stats = calculateItemStats(code, row, entry);
    toast(row
      ? `Leído: ${row.CODIGO} (FISICO: ${stats.fisico} / DIFEREN: ${stats.diferenFmt})`
      : `Código ${code} no estaba en base (FISICO: 1 / DIFEREN: +1)`, row ? 'ok' : 'warn');
  }

  saveScans();
  renderList();
  bumpBadge();
  return result;
}

/* =========================================================
   Feedback Sonoro y Visual
   ========================================================= */
let audioCtx;
function beep(freq = 1800, dur = 0.09, when = 0) {
  try {
    audioCtx ||= new (window.AudioContext || window.webkitAudioContext)();
    const o = audioCtx.createOscillator();
    const g = audioCtx.createGain();
    o.type = 'square';
    o.frequency.value = freq;
    g.gain.value = 0.08;
    o.connect(g).connect(audioCtx.destination);
    const t = audioCtx.currentTime + when;
    o.start(t); o.stop(t + dur);
  } catch { /* Sin soporte de audio */ }
}

function feedback(kind) {
  const flash = $('#scan-flash');
  if (flash) {
    flash.className = 'scan-flash';
    void flash.offsetWidth;
    if (kind === 'ok') { beep(1850); navigator.vibrate?.(70); flash.classList.add('ok'); }
    else if (kind === 'dup') { beep(1200, 0.07); beep(1200, 0.07, 0.12); navigator.vibrate?.([40, 60, 40]); flash.classList.add('warn'); }
    else { beep(520, 0.18); navigator.vibrate?.([120, 60, 120]); flash.classList.add('warn'); }
  }
}

/* =========================================================
   Componentes Visuales (HTML generators)
   ========================================================= */
function itemCardHTML(s, { dup = false, compact = false, deletable = false } = {}) {
  const r = s.row;
  const { fisico, totalTra, diferenVal, diferenFmt } = calculateItemStats(s.code, r, s);

  const cls = ['item-card', !s.found && 'missing', dup && 'dup'].filter(Boolean).join(' ');
  const tag = dup ? `<span class="tag dup">Lectura ×${s.count}</span>`
    : s.found ? '<span class="tag ok">En base</span>' : '<span class="tag warn">No en base</span>';

  const diffCls = diferenVal === 0 ? 'ok' : diferenVal < 0 ? 'neg' : '';

  const summaryBar = `
    <div class="fisico-summary-bar">
      <div class="fisico-box expected">
        <span>TOTAL_TRA</span>
        <b>${totalTra}</b>
      </div>
      <div class="fisico-box physical">
        <span>FISICO</span>
        <b>${fisico}</b>
      </div>
      <div class="fisico-box diff ${diffCls}">
        <span>DIFEREN</span>
        <b>${diferenFmt}</b>
      </div>
    </div>`;

  const time = s.time ? new Date(s.time).toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' }) : '';

  return `<article class="${cls}">
    <div class="item-head">
      <div>
        <div class="item-code">${esc(s.code)}</div>
        ${r ? `<p class="item-name">${esc(r.ARTICULO)}</p><div class="item-fam">${esc(r.FAMILIA)}</div>`
            : '<p class="item-name muted">Código sin registrar en base de stock</p>'}
      </div>
      <div class="item-actions">
        ${compact ? '' : tag}
        ${deletable ? `<button class="btn-del" data-del="${esc(s.code)}" title="Quitar este artículo" type="button"><svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg></button>` : ''}
      </div>
    </div>
    ${summaryBar}
    ${compact && time ? `<div class="item-meta">Última lectura: ${time}${s.found ? '' : ' · <span style="color:var(--warn)">agregado fuera de base</span>'}</div>` : ''}
  </article>`;
}

function baseRowCardHTML(r, scanCount) {
  const totalTra = Number(r.TOTAL_TRA || 0);
  const diferenVal = scanCount - totalTra;
  let diferenFmt = '0';
  if (diferenVal > 0) diferenFmt = `+${diferenVal}`;
  else if (diferenVal < 0) diferenFmt = `${diferenVal}`;

  const diffCls = diferenVal === 0 ? 'ok' : diferenVal < 0 ? 'neg' : '';
  const isRead = scanCount > 0;

  return `<article class="item-card ${isRead ? '' : 'missing'}">
    <div class="item-head">
      <div>
        <div class="item-code">${esc(r.CODIGO)}</div>
        <p class="item-name">${esc(r.ARTICULO)}</p>
        <div class="item-fam">${esc(r.FAMILIA)}</div>
      </div>
      <div class="item-actions">
        ${isRead ? `<span class="tag ok">FISICO: ${scanCount}</span>` : '<span class="tag warn">Pendiente (0)</span>'}
      </div>
    </div>
    <div class="fisico-summary-bar">
      <div class="fisico-box expected">
        <span>TOTAL_TRA</span>
        <b>${totalTra}</b>
      </div>
      <div class="fisico-box physical">
        <span>FISICO</span>
        <b>${scanCount}</b>
      </div>
      <div class="fisico-box diff ${diffCls}">
        <span>DIFEREN</span>
        <b>${diferenFmt}</b>
      </div>
    </div>
  </article>`;
}

function showResult(target, result) {
  if (!result) return;
  const el = $(target);
  el.classList.remove('empty');
  el.innerHTML = itemCardHTML(result, { dup: result.dup });
}

/* =========================================================
   Renderizado de Lista y Estadísticas
   ========================================================= */
function renderList() {
  const q = norm($('#input-search').value);
  const scanMap = getScanCountMap();

  let totalPhysicalUnits = 0;
  for (const count of scanMap.values()) totalPhysicalUnits += count;

  let countWithDiff = 0;
  if (state.base) {
    for (const r of state.base.rows) {
      const c = scanMap.get(norm(r.CODIGO)) || 0;
      const t = Number(r.TOTAL_TRA || 0);
      if (c !== t) countWithDiff++;
    }
  }

  $('#stat-total').textContent = state.scans.length;
  $('#stat-units').textContent = totalPhysicalUnits;
  $('#stat-diffs').textContent = countWithDiff;

  $('#cnt-scanned').textContent = state.scans.length;
  $('#cnt-base').textContent = state.base?.rows.length || 0;

  const listContainer = $('#scan-list');

  if (state.listViewMode === 'scanned') {
    const list = state.scans.filter((s) => !q || norm(s.code).includes(q) || norm(s.row?.ARTICULO).includes(q) || norm(s.row?.FAMILIA).includes(q));

    listContainer.innerHTML = list.length
      ? list.map((s) => `<li>${itemCardHTML(s, { compact: true, deletable: true })}</li>`).join('')
      : `<li class="empty-state"><svg viewBox="0 0 24 24"><path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2M7 8v8M10 8v8M14 8v8M17 8v8"/></svg><p>${q ? 'Sin resultados para la búsqueda.' : 'Todavía no escaneaste ningún código.'}</p></li>`;
  } else {
    // Vista 'base completa'
    if (!state.base?.rows.length) {
      listContainer.innerHTML = `<li class="empty-state"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/></svg><p>No hay base de stock cargada.<br>Importá un archivo Excel para ver el inventario completo.</p></li>`;
      return;
    }

    const filteredRows = state.base.rows.filter((r) => !q || norm(r.CODIGO).includes(q) || norm(r.ARTICULO).includes(q) || norm(r.FAMILIA).includes(q));

    listContainer.innerHTML = filteredRows.length
      ? filteredRows.map((r) => {
          const count = scanMap.get(norm(r.CODIGO)) || 0;
          return `<li>${baseRowCardHTML(r, count)}</li>`;
        }).join('')
      : `<li class="empty-state"><p>Sin resultados en la base de stock.</p></li>`;
  }

  const badge = $('#badge-count');
  badge.hidden = !state.scans.length;
  badge.textContent = state.scans.length > 999 ? '999+' : state.scans.length;
  $('#btn-export').disabled = !state.scans.length && !state.base?.rows.length;
}

function bumpBadge() {
  const b = $('#badge-count');
  b.classList.remove('bump'); void b.offsetWidth; b.classList.add('bump');
}

function renderBaseState() {
  const chip = $('#btn-base');
  const n = state.base?.rows.length || 0;
  chip.classList.toggle('loaded', !!n);
  $('#base-chip-text').textContent = n ? `Base: ${n} art.` : 'Sin base';

  const banner = $('#no-base-banner');
  if (banner) banner.style.display = n ? 'none' : 'flex';

  $('#base-info').innerHTML = n
    ? `Archivo actual: <b>${esc(state.base.fileName)}</b><br>Total artículos registrados: <b>${n}</b><br><span class="muted">Cargado ${new Date(state.base.loadedAt).toLocaleString('es-AR')}</span>`
    : '<span class="muted">No se ha cargado ninguna base. Si escaneás códigos, se guardarán y exportarán con FISICO=cantidad y DIFEREN=+cantidad.</span>';

  $('#btn-base-clear').hidden = !n;
}

/* =========================================================
   Toasts (Notificaciones flotantes)
   ========================================================= */
const ICONS = {
  ok: '<svg viewBox="0 0 24 24"><path d="M20 6 9 17l-5-5"/></svg>',
  warn: '<svg viewBox="0 0 24 24"><path d="M12 9v4M12 17h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>',
  err: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M15 9l-6 6M9 9l6 6"/></svg>',
};

function toast(msg, type = 'ok', ms = 2400) {
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.innerHTML = `${ICONS[type] || ''}<span>${esc(msg)}</span>`;
  const box = $('#toasts');
  box.prepend(t);
  while (box.children.length > 4) box.lastChild.remove();
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 300); }, ms);
}

/* =========================================================
   Navegación
   ========================================================= */
function showView(id) {
  document.querySelectorAll('.view').forEach((v) => v.classList.toggle('active', v.id === id));
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === id));
  if (id !== 'view-live') stopCamera();
  if (id === 'view-manual') setTimeout(() => $('#input-code').focus(), 50);
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* =========================================================
   3) Escaneo EN VIVO (Cámara)
   ========================================================= */
let scanner = null;
let running = false;
let lastCode = '';
let lastAt = 0;
let torchOn = false;

function scannerConfig() {
  return { formatsToSupport: FORMATS, verbose: false, useBarCodeDetectorIfSupported: true,
    experimentalFeatures: { useBarCodeDetectorIfSupported: true } };
}

async function startCamera() {
  if (running) return;
  if (!window.isSecureContext) {
    toast('La cámara requiere HTTPS en celular. Abrí la app con la URL https://', 'err', 4000);
    return;
  }
  scanner ||= new Html5Qrcode('reader', scannerConfig());
  const btn = $('#btn-cam-toggle');
  btn.disabled = true;

  try {
    await scanner.start(
      { facingMode: 'environment' },
      {
        fps: 15,
        disableFlip: true,
        videoConstraints: {
          facingMode: 'environment',
          width: { ideal: 1920 }, height: { ideal: 1080 },
          advanced: [{ focusMode: 'continuous' }],
        },
      },
      onLiveDecode,
      () => {}
    );
    running = true;
    $('.scanner-card').classList.add('running');
    setCamButton();
    setupTorch();
  } catch (e) {
    console.error(e);
    toast('No se pudo iniciar la cámara. Verificá los permisos.', 'err', 4000);
  } finally {
    btn.disabled = false;
  }
}

async function stopCamera() {
  if (!running || !scanner) return;
  running = false;
  try { await scanner.stop(); } catch { /* noop */ }
  $('.scanner-card').classList.remove('running');
  torchOn = false;
  $('#btn-torch').classList.remove('on');
  $('#btn-torch').disabled = true;
  setCamButton();
}

function onLiveDecode(text) {
  const now = Date.now();
  // Antirrebote: mismo código < 2.5s o cualquier código < 0.9s
  if ((text === lastCode && now - lastAt < 2500) || now - lastAt < 900) return;
  lastCode = text; lastAt = now;
  showResult('#last-scan', registerCode(text, 'vivo'));
}

function setCamButton() {
  const btn = $('#btn-cam-toggle');
  btn.className = `btn ${running ? 'btn-stop' : 'btn-primary'}`;
  btn.innerHTML = running
    ? '<svg viewBox="0 0 24 24"><rect x="6" y="6" width="12" height="12" rx="2"/></svg><span>Detener cámara</span>'
    : '<svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg><span>Iniciar cámara</span>';
}

function setupTorch() {
  const btn = $('#btn-torch');
  try {
    const torch = scanner.getRunningTrackCameraCapabilities().torchFeature();
    btn.disabled = !torch.isSupported();
  } catch { btn.disabled = true; }
}

async function toggleTorch() {
  try {
    const torch = scanner.getRunningTrackCameraCapabilities().torchFeature();
    torchOn = !torchOn;
    await torch.apply(torchOn);
    $('#btn-torch').classList.toggle('on', torchOn);
  } catch { toast('Linterna no disponible en este dispositivo', 'warn'); }
}

/* =========================================================
   1) Lectura por FOTO
   ========================================================= */
async function decodeWithNative(file) {
  if (!('BarcodeDetector' in window)) return null;
  try {
    const supported = await window.BarcodeDetector.getSupportedFormats();
    const detector = new window.BarcodeDetector({ formats: supported });
    const bmp = await createImageBitmap(file);
    const found = await detector.detect(bmp);
    return found[0]?.rawValue || null;
  } catch { return null; }
}

async function downscale(file, max = 1600) {
  const bmp = await createImageBitmap(file);
  const k = Math.min(1, max / Math.max(bmp.width, bmp.height));
  if (k === 1) return file;
  const c = document.createElement('canvas');
  c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.92));
  return new File([blob], 'img.jpg', { type: 'image/jpeg' });
}

async function handlePhoto(file) {
  if (!file) return;
  const preview = $('#photo-preview');
  const status = $('#photo-status');
  $('#photo-img').src = URL.createObjectURL(file);
  preview.hidden = false;
  $('#photo-result').innerHTML = '';
  status.textContent = '🔍 Leyendo código de barras...';

  let text = await decodeWithNative(file);
  if (!text) {
    const fileScanner = new Html5Qrcode('file-reader', scannerConfig());
    for (const candidate of [file, await downscale(file, 1600), await downscale(file, 900)]) {
      try { text = await fileScanner.scanFile(candidate, false); break; } catch { /* siguiente */ }
    }
    try { fileScanner.clear(); } catch { /* noop */ }
  }

  if (text) {
    status.textContent = `✅ Código: ${text}`;
    showResult('#photo-result', registerCode(text, 'foto'));
  } else {
    status.textContent = '❌ No se detectó código de barras';
    feedback('miss');
    toast('No se pudo leer el código. Enfocá con buena iluminación.', 'err', 3500);
  }
}

/* =========================================================
   2) Carga MANUAL
   ========================================================= */
function renderSuggestions() {
  const q = norm($('#input-code').value);
  const ul = $('#suggestions');
  if (q.length < 2 || !state.base) { ul.innerHTML = ''; return; }
  const ql = loose(q);
  const matches = state.base.rows.filter((r) =>
    loose(r.CODIGO).includes(ql) || norm(r.ARTICULO).includes(q)).slice(0, 8);

  ul.innerHTML = matches.map((r) =>
    `<li data-code="${esc(r.CODIGO)}"><b>${esc(r.CODIGO)}</b><span>${esc(r.ARTICULO)}</span></li>`).join('');
}

function submitManual(code) {
  const value = (code ?? $('#input-code').value).trim();
  if (!value) { toast('Ingresá un código o nombre', 'warn'); return; }
  showResult('#manual-result', registerCode(value, 'manual'));
  $('#input-code').value = '';
  $('#suggestions').innerHTML = '';
  $('#input-code').focus();
}

/* =========================================================
   EXPORTACIÓN A EXCEL (.XLSX)
   Controla TOTAL_TRA, FISICO y DIFEREN
   ========================================================= */
function exportExcel() {
  const scanMap = getScanCountMap();
  const exportedRows = [];
  const processedBaseKeys = new Set();

  // 1. Exportar TODAS las filas de la base original cargada
  if (state.base?.rows) {
    for (const r of state.base.rows) {
      const key = norm(r.CODIGO);
      processedBaseKeys.add(key);
      const fisico = scanMap.get(key) || 0;
      const totalTra = Number(r.TOTAL_TRA || 0);
      const diferenVal = fisico - totalTra;

      let diferenStr = diferenVal === 0 ? 0 : (diferenVal > 0 ? `+${diferenVal}` : diferenVal);

      exportedRows.push({
        CODIGO: r.CODIGO,
        FAMILIA: r.FAMILIA,
        ARTICULO: r.ARTICULO,
        STOCK_DIS: r.STOCK_DIS,
        STOCK_IND: r.STOCK_IND,
        STOCK_REM: r.STOCK_REM,
        STOCK_RES: r.STOCK_RES,
        STOCK_TRA: r.STOCK_TRA,
        STOCK_EXH: r.STOCK_EXH,
        TOTAL_TRA: totalTra,
        FISICO: fisico,
        DIFEREN: diferenStr,
      });
    }
  }

  // 2. Exportar los códigos leídos que NO estaban en la base
  for (const s of state.scans) {
    const key = s.row ? norm(s.row.CODIGO) : norm(s.code);
    if (!processedBaseKeys.has(key)) {
      processedBaseKeys.add(key);
      const fisico = s.count || 1;
      exportedRows.push({
        CODIGO: s.code,
        FAMILIA: 'NUEVO / NO EN BASE',
        ARTICULO: 'Artículo leído fuera de la base de datos',
        STOCK_DIS: 0,
        STOCK_IND: 0,
        STOCK_REM: 0,
        STOCK_RES: 0,
        STOCK_TRA: 0,
        STOCK_EXH: 0,
        TOTAL_TRA: 0,
        FISICO: fisico,
        DIFEREN: `+${fisico}`,
      });
    }
  }

  if (!exportedRows.length) {
    toast('No hay datos para exportar.', 'warn');
    return;
  }

  // Construir matriz para el libro de Excel
  const aoa = [
    COLS,
    ...exportedRows.map((row) => COLS.map((col) => row[col] ?? '')),
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = COLS.map((c, i) => ({ wch: i === 0 ? 14 : i < 3 ? 32 : 11 }));
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: COLS.length - 1 } }) };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'INVENTARIO_FISICO');

  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const filename = `INVENTARIO_FISICO_${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}.xlsx`;

  XLSX.writeFile(wb, filename, { compression: true });
  toast(`Excel exportado con éxito (${exportedRows.length} filas)`, 'ok');
}

/* =========================================================
   Asignación de Eventos
   ========================================================= */
function bindEvents() {
  // Pestañas inferiores
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => showView(t.dataset.view)));

  // Importar Excel (Botón superior + Chip + Modal)
  const dlg = $('#dlg-base');
  const openImport = () => dlg.showModal();
  $('#btn-open-import')?.addEventListener('click', openImport);
  $('#btn-base')?.addEventListener('click', openImport);
  dlg.addEventListener('click', (e) => { if (e.target === dlg) dlg.close(); });

  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try {
      const base = await parseBaseFile(file);
      setBase(base, true);
      toast(`Base importada: ${base.rows.length} artículos`, 'ok');
      dlg.close();
    } catch (err) {
      toast(err.message || 'Error al procesar el archivo Excel', 'err', 4000);
    }
  };

  $('#input-base')?.addEventListener('change', handleFileChange);
  $('#input-base-direct')?.addEventListener('change', handleFileChange);

  $('#btn-load-sample')?.addEventListener('click', async () => {
    await loadDefaultSampleBase();
    dlg.close();
  });

  $('#btn-base-clear')?.addEventListener('click', () => {
    if (!confirm('¿Seguro que deseas quitar la base de stock?')) return;
    setBase(null, true);
    toast('Base eliminada', 'warn');
  });

  // En vivo
  $('#cam-placeholder')?.addEventListener('click', startCamera);
  $('#btn-cam-toggle')?.addEventListener('click', () => (running ? stopCamera() : startCamera()));
  $('#btn-torch')?.addEventListener('click', toggleTorch);
  document.addEventListener('visibilitychange', () => { if (document.hidden) stopCamera(); });

  // Fotos
  for (const id of ['#input-photo', '#input-gallery']) {
    $(id)?.addEventListener('change', (e) => { handlePhoto(e.target.files[0]); e.target.value = ''; });
  }

  // Manual
  $('#form-manual')?.addEventListener('submit', (e) => { e.preventDefault(); submitManual(); });
  $('#input-code')?.addEventListener('input', renderSuggestions);
  $('#suggestions')?.addEventListener('click', (e) => {
    const li = e.target.closest('li[data-code]');
    if (li) submitManual(li.dataset.code);
  });

  // Sub-tabs de la lista
  $('#tab-scanned-only')?.addEventListener('click', () => {
    state.listViewMode = 'scanned';
    $('#tab-scanned-only').classList.add('active');
    $('#tab-all-base').classList.remove('active');
    renderList();
  });

  $('#tab-all-base')?.addEventListener('click', () => {
    state.listViewMode = 'base';
    $('#tab-all-base').classList.add('active');
    $('#tab-scanned-only').classList.remove('active');
    renderList();
  });

  // Lista & Exportación
  $('#input-search')?.addEventListener('input', renderList);
  $('#scan-list')?.addEventListener('click', (e) => {
    const b = e.target.closest('[data-del]');
    if (!b) return;
    state.scans = state.scans.filter((s) => s.code !== b.dataset.del);
    saveScans();
    renderList();
    toast(`Eliminado ${b.dataset.del}`, 'warn');
  });

  $('#btn-export')?.addEventListener('click', exportExcel);
  $('#btn-clear')?.addEventListener('click', () => {
    if (!state.scans.length) return;
    if (!confirm(`¿Vaciar la lista de ${state.scans.length} lecturas efectuadas?`)) return;
    state.scans = [];
    saveScans();
    renderList();
    $('#last-scan').className = 'last-scan empty';
    $('#last-scan').innerHTML = '<p class="muted">Los artículos leídos aparecerán acá.</p>';
    toast('Lista de lecturas vaciada.', 'warn');
  });
}

/* =========================================================
   Inicialización
   ========================================================= */
async function init() {
  loadSavedData();
  bindEvents();
  renderBaseState();
  renderList();

  // Si no había base guardada en localStorage, cargar la base por defecto de la carpeta
  if (!state.base) {
    await loadDefaultSampleBase();
  }
}

init();
