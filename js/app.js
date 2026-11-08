/* Inventory Console — client-side data admin for CSV / Excel / Google Sheets.
   No build step, no server: everything runs in the browser. */
(() => {
'use strict';

// ════════════════════════════════════════════════════════════════
//  Helpers
// ════════════════════════════════════════════════════════════════
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
const fmt = (n, d = 0) => (n == null || isNaN(n)) ? '—' : n.toLocaleString(undefined, { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad = n => String(n).padStart(2, '0');
const fmtDate = (ts, time = true) => {
  if (ts == null || isNaN(ts)) return '—';
  const d = new Date(ts);
  const s = `${pad(d.getDate())} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
  return time && (d.getHours() || d.getMinutes()) ? `${s} ${pad(d.getHours())}:${pad(d.getMinutes())}` : s;
};
const dayKey = ts => { const d = new Date(ts); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const isoDay = ts => dayKey(ts);
const startOfDay = ts => { const d = new Date(ts); d.setHours(0, 0, 0, 0); return d.getTime(); };
const shortLabel = (s, n = 30) => { s = String(s).replace(/\s*\(\d{5,}\)\s*$/, ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; };
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const toneOf = v => {
  const s = String(v).toLowerCase();
  if (/cancel|reject|fail|error|lost|^no$/.test(s)) return 'bad';
  if (/pend|hold|wait|draft|estimated/.test(s)) return 'warn';
  if (/confirm|active|complete|done|deliver|^yes$|actual|sailed/.test(s)) return 'ok';
  return '';
};
const pill = v => v === '' || v == null ? '<span class="faint">—</span>' : `<span class="pill ${toneOf(v)}">${esc(v)}</span>`;
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } },
  del(k) { try { localStorage.removeItem(k); } catch {} },
};
/** IndexedDB key/value store — holds the (merged) dataset; far larger than localStorage's ~5 MB. */
const idb = (() => {
  let dbp;
  const open = () => dbp ||= new Promise((res, rej) => {
    const rq = indexedDB.open('inventory-console', 1);
    rq.onupgradeneeded = () => rq.result.createObjectStore('kv');
    rq.onsuccess = () => res(rq.result);
    rq.onerror = () => rej(rq.error);
  });
  const run = (mode, fn) => open().then(db => new Promise((res, rej) => {
    const tx = db.transaction('kv', mode), rq = fn(tx.objectStore('kv'));
    tx.oncomplete = () => res(rq.result); tx.onerror = tx.onabort = () => rej(tx.error);
  }));
  return {
    get: k => run('readonly', s => s.get(k)).catch(() => undefined),
    set: (k, v) => run('readwrite', s => s.put(v, k)).then(() => true, e => { console.error(e); return false; }),
    del: k => run('readwrite', s => s.delete(k)).catch(() => {}),
  };
})();
let toastT;
function toast(msg, error = false) {
  const t = $('#toast'); t.textContent = msg; t.className = 'toast show' + (error ? ' error' : '');
  clearTimeout(toastT); toastT = setTimeout(() => t.className = 'toast' + (error ? ' error' : ''), error ? 6000 : 2600);
}
const ICON = {
  x: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  upload: '<svg viewBox="0 0 24 24"><path d="M12 16V4M7 9l5-5 5 5M4 20h16"/></svg>',
  sheet: '<svg viewBox="0 0 24 24"><path d="M5 3h10l4 4v14H5zM5 10h14M5 15h14M10 10v11"/></svg>',
  cols: '<svg viewBox="0 0 24 24"><path d="M4 4h5v16H4zM10 4h5v16h-5zM16 4h4v16h-4z"/></svg>',
};

// ════════════════════════════════════════════════════════════════
//  State
// ════════════════════════════════════════════════════════════════
const prefs = Object.assign({ theme: null, hidden: {}, sheets: [], refreshMin: 0, uploadMode: 'merge' }, store.get('ic.prefs', {}));
const savePrefs = () => store.set('ic.prefs', prefs);

const S = {
  ds: null,                 // current dataset
  view: 'overview',
  f: { cats: new Map(), ranges: new Map(), search: '' },
  filtered: [],
  facetCounts: new Map(),
  tables: {},
  charts: {},
  ui: { openFacets: new Set(), facetQ: {} },
  pivot: { row: null, col: null, measure: 'count' },
  workbook: null,
  refreshTimer: null,
};

// ════════════════════════════════════════════════════════════════
//  Parsing
// ════════════════════════════════════════════════════════════════
const BLANKS = new Set(['', 'not available', 'n/a', 'na', '-', '--', 'null', 'none', 'nil', '#n/a']);
const isBlank = v => BLANKS.has(String(v).trim().toLowerCase());
const NUM_RE = /^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/;
const DMY_RE = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?)?/i;
const ISO_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/;
const ID_LABEL_RE = /(number|\bno\b|\bno\.|\bid\b|code|#|phone|mobile|zip|pin|postal|serial|sku|barcode|imei|gst|pan)/i;

function cleanCell(v) {
  let s = v == null ? '' : String(v);
  const m = s.match(/^="([\s\S]*)"$/);       // Excel text guard  ="00123"
  if (m) s = m[1];
  return s.replace(/\s+/g, ' ').trim();
}

function parseDateStr(s, dayFirst) {
  let m = ISO_RE.exec(s), y, mo, d, h = 0, mi = 0, se = 0;
  if (m) { [y, mo, d] = [+m[1], +m[2], +m[3]]; h = +(m[4] || 0); mi = +(m[5] || 0); se = +(m[6] || 0); }
  else if ((m = DMY_RE.exec(s))) {
    const a = +m[1], b = +m[2];
    [d, mo] = dayFirst ? [a, b] : [b, a];
    y = +m[3]; if (y < 100) y += 2000;
    h = +(m[4] || 0); mi = +(m[5] || 0); se = +(m[6] || 0);
    if (m[7]) { const pm = m[7].toLowerCase() === 'pm'; if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
  } else return NaN;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return NaN;
  return new Date(y, mo - 1, d, h, mi, se).getTime();
}

/** Turn a raw text file into an array-of-arrays. Handles Excel's `sep=;` line and BOM. */
function textToAoa(text) {
  text = text.replace(/^﻿/, '');
  let delimiter = '';
  const m = text.match(/^sep=(.)\s*\r?\n/i);
  if (m) { delimiter = m[1]; text = text.slice(m[0].length); }
  const res = Papa.parse(text, { delimiter, skipEmptyLines: 'greedy' });
  return res.data;
}

/** Drop empty rows and locate the header: the first row with ≥2 filled cells within the first 10 rows. */
function splitAoa(aoa) {
  aoa = aoa.filter(r => Array.isArray(r) && r.some(c => cleanCell(c) !== ''));
  if (!aoa.length) throw new Error('The file has no rows.');
  let hi = aoa.findIndex((r, i) => i < 10 && r.filter(c => cleanCell(c) !== '').length >= 2);
  if (hi < 0) hi = 0;
  return { header: aoa[hi], body: aoa.slice(hi + 1) };
}

// ════════════════════════════════════════════════════════════════
//  Merging uploads (unique by Booking Number, latest file wins)
// ════════════════════════════════════════════════════════════════
const headerKey = h => norm(cleanCell(h).replace(/\*+$/, ''));
const KEY_NAMES = ['Booking Number', 'Booking No', 'Booking #', 'Booking'].map(norm);
function keyIndex(header) {
  const L = header.map(headerKey);
  for (const k of KEY_NAMES) { const i = L.indexOf(k); if (i >= 0) return i; }
  return -1;
}

/**
 * Merge a new upload into the stored data.
 * Every booking in the new file replaces ALL old rows of that booking (so container lines that
 * changed, or went from "Unassigned" to a real number, never duplicate). Bookings that are not in
 * the new file are kept untouched. Columns are matched by name, so column order may differ.
 * Returns null when either side has no Booking Number column.
 */
function mergeAoa(oldAoa, newAoa) {
  const A = splitAoa(oldAoa), B = splitAoa(newAoa);
  const ka = keyIndex(A.header), kb = keyIndex(B.header);
  if (ka < 0 || kb < 0) return null;

  // union of columns, old order first
  const header = [...A.header], pos = new Map(header.map((h, i) => [headerKey(h), i]));
  const mapB = B.header.map(h => {
    const k = headerKey(h);
    if (!k) return -1;
    if (!pos.has(k)) { pos.set(k, header.length); header.push(h); }
    return pos.get(k);
  });
  const W = header.length;
  const widen = r => { const o = r.slice(0, W); while (o.length < W) o.push(''); return o; };
  const remap = r => { const o = new Array(W).fill(''); mapB.forEach((j, i) => { if (j >= 0 && r[i] != null) o[j] = r[i]; }); return o; };
  const keyOf = (r, k) => cleanCell(r[k]);
  const rowSig = r => r.map(cleanCell).join('\u0001');

  // group the new rows by booking
  const incoming = new Map(), looseNew = new Set();
  const newRows = B.body.map(remap);
  for (const r of newRows) {
    const k = keyOf(r, ka);
    if (k) (incoming.get(k) || incoming.set(k, []).get(k)).push(r); else looseNew.add(rowSig(r));
  }
  // group the old rows by booking
  const oldByKey = new Map(), kept = [];
  for (const r0 of A.body) {
    const r = widen(r0), k = keyOf(r, ka);
    if (k && incoming.has(k)) { (oldByKey.get(k) || oldByKey.set(k, []).get(k)).push(r); continue; }   // replaced
    if (!k && looseNew.has(rowSig(r))) continue;                                                      // same key-less row re-uploaded
    kept.push(r);
  }
  const oldKeys = new Set(A.body.map(r => keyOf(r, ka)).filter(Boolean));
  let added = 0, updated = 0, unchanged = 0;
  for (const [k, rows] of incoming) {
    if (!oldKeys.has(k)) { added++; continue; }
    const before = (oldByKey.get(k) || []).map(rowSig).sort().join('\n');
    const after = rows.map(rowSig).sort().join('\n');
    before === after ? unchanged++ : updated++;
  }
  const total = new Set([...oldKeys, ...incoming.keys()]).size;
  return { aoa: [header, ...kept, ...newRows], stats: { added, updated, unchanged, inFile: incoming.size, total, rows: kept.length + newRows.length } };
}

/** Build a typed, column-major dataset from an array-of-arrays. */
function buildDataset(aoa, meta) {
  const { header: rawHeader, body } = splitAoa(aoa);
  const header = rawHeader.map(cleanCell);
  const width = Math.max(header.length, ...body.map(r => r.length));
  const seen = {};
  const cols = [];
  for (let c = 0; c < width; c++) {
    let label = (header[c] || '').replace(/\*+$/, '').trim();
    const values = body.map(r => cleanCell(r[c]));
    if (!label && values.every(v => v === '')) continue;      // trailing empty column
    if (!label) label = `Column ${c + 1}`;
    if (seen[label]) label = `${label} (${++seen[label]})`; else seen[label] = 1;
    cols.push({ label, values });
  }
  const n = body.length;
  for (const col of cols) inferColumn(col, n);

  const ds = { ...meta, cols, n, loadedAt: meta.loadedAt || Date.now(), aoa: [rawHeader, ...body] };
  ds.profile = detectBookingProfile(ds);
  if (ds.profile) addBookingDerived(ds);
  finalizeDataset(ds);
  return ds;
}

function inferColumn(col, n) {
  const vals = col.values;
  // strip trailing "*" when every filled value has it (e.g. "Confirmed*")
  const filled = vals.filter(v => v !== '');
  if (filled.length && filled.every(v => v.endsWith('*'))) col.values = vals.map(v => v.replace(/\*+$/, ''));
  const V = col.values;
  let nonBlank = 0, numHits = 0, dateHits = 0, aGt12 = false, bGt12 = false;
  for (const v of V) {
    if (isBlank(v)) continue;
    nonBlank++;
    if (NUM_RE.test(v)) { numHits++; continue; }
    const m = DMY_RE.exec(v);
    if (m) { dateHits++; if (+m[1] > 12) aGt12 = true; if (+m[2] > 12) bGt12 = true; }
    else if (ISO_RE.test(v)) dateHits++;
  }
  col.type = 'text';
  if (nonBlank && dateHits / nonBlank >= 0.8) {
    col.type = 'date';
    const dayFirst = !(bGt12 && !aGt12);
    col.ts = Float64Array.from(V, v => isBlank(v) ? NaN : parseDateStr(v, dayFirst));
  } else if (nonBlank && numHits / nonBlank >= 0.9 && !ID_LABEL_RE.test(col.label)) {
    col.type = 'number';
    col.num = Float64Array.from(V, v => isBlank(v) || !NUM_RE.test(v) ? NaN : parseFloat(v.replace(/,/g, '')));
  }
}

function finalizeDataset(ds) {
  for (const col of ds.cols) {
    const u = new Map();
    for (const v of col.values) u.set(v, (u.get(v) || 0) + 1);
    col.uniques = u;
    col.isCat = col.type === 'text' && u.size >= 2 && u.size <= 80 && u.size <= Math.max(12, ds.n * 0.6);
    if (col.type === 'number') {
      let mn = Infinity, mx = -Infinity;
      for (const x of col.num) if (!isNaN(x)) { if (x < mn) mn = x; if (x > mx) mx = x; }
      col.min = mn; col.max = mx;
      col.isRange = isFinite(mn) && mx > mn;
    }
    if (col.type === 'date') {
      let mn = Infinity, mx = -Infinity;
      for (const x of col.ts) if (!isNaN(x)) { if (x < mn) mn = x; if (x > mx) mx = x; }
      col.min = mn; col.max = mx;
      col.isRange = isFinite(mn);
    }
  }
  ds.cols.forEach((c, i) => c.i = i);
  ds.sig = norm(ds.cols.filter(c => !c.derived).map(c => c.label).join('|')).slice(0, 200);
  ds.text = null;   // lazy search index

  // facet ordering
  const P = ds.profile;
  const preferred = P ? [P.status, P.etdType, P.pol, P.pod, P.voyage, P.service, P.size, P.event, P.assigned, P.shipper, P.forwarder, P.origin, P.dest]
    .filter(c => c != null) : [];
  const noise = P ? /time zone|unit$|^total container quantity of/i : /^$/;
  const cats = ds.cols.filter(c => c.isCat && !noise.test(c.label)).map(c => c.i);
  ds.facetCols = [...preferred.filter(i => cats.includes(i)), ...cats.filter(i => !preferred.includes(i))];
  ds.rangeCols = ds.cols.filter(c => c.isRange).map(c => c.i);
  if (P) { // put ETD first among date ranges
    ds.rangeCols.sort((a, b) => (b === P.etd) - (a === P.etd) || (ds.cols[b].type === 'date') - (ds.cols[a].type === 'date'));
  }
}

// ════════════════════════════════════════════════════════════════
//  Booking / shipment profile (auto-detected by column names)
// ════════════════════════════════════════════════════════════════
const STAGES = [
  ['no event', 0, 'No event'],
  ['empty container picked', 1, 'Empty picked up'],
  ['full container received', 2, 'Gated in (full)'],
  ['intermodal', 3, 'Left inland hub'],
  ['arrival at first port', 4, 'Arrived at POL'],
  ['loaded on board', 5, 'Loaded on board'],
  ['vessel departure', 6, 'Sailed'],
];
const stageRank = ev => { const s = String(ev).toLowerCase(); for (const [k, r] of STAGES) if (s.includes(k)) return r; return s ? -1 : 0; };

function detectBookingProfile(ds) {
  const L = ds.cols.map(c => norm(c.label));
  const find = (...names) => {
    for (const n of names) { const i = L.indexOf(norm(n)); if (i >= 0) return i; }
    for (const n of names) { const i = L.findIndex(l => l.includes(norm(n))); if (i >= 0) return i; }
    return null;
  };
  const p = {
    booking: find('Booking Number', 'Booking No', 'Booking'),
    container: find('Container Number', 'Container No'),
  };
  if (p.booking == null || p.container == null) return null;
  Object.assign(p, {
    status: find('Booking Status', 'Status'),
    service: find('First Vessel Service', 'Service'),
    voyage: find('First Vessel Voyage', 'Vessel Voyage', 'Vessel'),
    etd: find('Departure at First Port of Load Date/Time', 'ETD'),
    origin: find('Origin City'),
    pol: find('Port of Load', 'POL'),
    pod: find('Port of Discharge', 'POD'),
    dest: find('Destination City'),
    size: find('Container Size/Type', 'Container Type', 'Size/Type'),
    weight: find('Container Weight Quantity', 'Weight'),
    volume: find('Container Volume Quantity', 'Volume'),
    emptyAt: find('Empty Container Picked Up Date/Time'),
    emptyLoc: find('Empty Container Picked Up Location'),
    gateAt: find('Full Container Received by Carrier at Origin Date/Time'),
    gateLoc: find('Full Container Received by Carrier at Origin Location'),
    polAt: find('Arrival at First Port of Load Date/Time'),
    polLoc: find('Arrival at First Port of Load Location'),
    event: find('Latest Container Event'),
    eventAt: find('Latest Container Event Date/Time'),
    eventLoc: find('Latest Container Event Location'),
    forwarder: find('Forwarder'),
    shipper: find('Shipper'),
  });
  if (p.event != null && p.event === p.eventAt) p.eventAt = null;
  return p;
}

function addBookingDerived(ds) {
  const P = ds.profile, n = ds.n, C = ds.cols;
  const add = (label, values) => { C.push({ label, values, derived: true }); inferColumn(C[C.length - 1], n); return C.length - 1; };
  if (P.etd != null) P.etdType = add('ETD Type', C[P.etd].values.map(v => (v.match(/\((actual|estimated)\)/i) || [, ''])[1].replace(/^./, c => c.toUpperCase())));
  P.assigned = add('Container Assigned', C[P.container].values.map(v => (!v || /unassigned/i.test(v)) ? 'No' : 'Yes'));
  if (P.size != null) P.teu = add('TEU', C[P.size].values.map(v => /^\s*(40|45)/.test(v) ? '2' : /^\s*20/.test(v) ? '1' : ''));
  if (P.event != null) P.stage = add('Milestone', C[P.event].values.map(v => { const r = stageRank(v); return r < 0 ? 'Other' : STAGES.find(s => s[1] === r)[2]; }));
}

// convenient accessors
const val = (c, r) => S.ds.cols[c].values[r];
const numv = (c, r) => { const col = S.ds.cols[c]; return col.num ? col.num[r] : NaN; };
const tsv = (c, r) => { const col = S.ds.cols[c]; return col.ts ? col.ts[r] : NaN; };

// ════════════════════════════════════════════════════════════════
//  Filtering & search
// ════════════════════════════════════════════════════════════════
const ALIASES = { pod: 'pod', pol: 'pol', status: 'status', vessel: 'voyage', voyage: 'voyage', service: 'service', shipper: 'shipper', forwarder: 'forwarder', container: 'container', cntr: 'container', booking: 'booking', bkg: 'booking', event: 'event', type: 'size', size: 'size', origin: 'origin', dest: 'dest' };

function resolveField(name) {
  const D = S.ds, k = norm(name);
  if (D.profile && ALIASES[k] && D.profile[ALIASES[k]] != null) return D.profile[ALIASES[k]];
  let i = D.cols.findIndex(c => norm(c.label) === k);
  if (i < 0) i = D.cols.findIndex(c => norm(c.label).startsWith(k));
  if (i < 0) i = D.cols.findIndex(c => norm(c.label).includes(k));
  return i < 0 ? null : i;
}

function parseSearch(q) {
  const out = [];
  for (const m of q.matchAll(/(-?)(?:([\w/]+):)?(?:"([^"]+)"|(\S+))/g)) {
    const neg = !!m[1], field = m[2], term = (m[3] ?? m[4] ?? '').toLowerCase();
    if (!term) continue;
    const col = field ? resolveField(field) : null;
    if (field && col == null) { out.push({ neg, col: null, term: `${field}:${term}`.toLowerCase() }); continue; }
    out.push({ neg, col, term });
  }
  return out;
}

function buildTextIndex() {
  const D = S.ds;
  D.text = new Array(D.n);
  for (let r = 0; r < D.n; r++) D.text[r] = D.cols.map(c => c.values[r]).join(' \u0001 ').toLowerCase();
}

function applyFilters() {
  const D = S.ds;
  if (!D) { S.filtered = []; return; }
  const tokens = parseSearch(S.f.search);
  if (tokens.length && !D.text) buildTextIndex();
  const cats = [...S.f.cats].filter(([, s]) => s.size);
  const ranges = [...S.f.ranges];
  const counts = new Map(D.facetCols.map(c => [c, new Map()]));
  const inc = (m, v) => m.set(v, (m.get(v) || 0) + 1);
  const out = [];
  rows: for (let r = 0; r < D.n; r++) {
    for (const t of tokens) {
      const hay = t.col == null ? D.text[r] : D.cols[t.col].values[r].toLowerCase();
      if (hay.includes(t.term) === t.neg) continue rows;
    }
    for (const [c, rg] of ranges) {
      const col = D.cols[c], x = col.type === 'date' ? col.ts[r] : col.num[r];
      if (rg.min != null && !(x >= rg.min)) continue rows;
      if (rg.max != null && !(x <= rg.max)) continue rows;
    }
    let fails = 0, failCol = -1;
    for (const [c, set] of cats) {
      if (!set.has(D.cols[c].values[r])) { fails++; failCol = c; if (fails > 1) break; }
    }
    if (fails === 0) { out.push(r); for (const c of D.facetCols) inc(counts.get(c), D.cols[c].values[r]); }
    else if (fails === 1 && counts.has(failCol)) inc(counts.get(failCol), D.cols[failCol].values[r]);
  }
  S.filtered = out;
  S.facetCounts = counts;
}

function activeFilterCount() {
  let n = 0;
  for (const [, s] of S.f.cats) if (s.size) n++;
  n += S.f.ranges.size;
  return n;
}

function setCat(c, v, on) {
  let s = S.f.cats.get(c); if (!s) S.f.cats.set(c, s = new Set());
  if (on === undefined) on = !s.has(v);
  on ? s.add(v) : s.delete(v);
  if (!s.size) S.f.cats.delete(c);
}
function onlyCat(c, v) { S.f.cats.set(c, new Set([v])); }

const refresh = () => {
  applyFilters();
  for (const t of Object.values(S.tables)) t.page = 1;
  renderChrome(); renderView();
};

// ════════════════════════════════════════════════════════════════
//  Loading data
// ════════════════════════════════════════════════════════════════
function setDataset(ds, { keepFilters = false, persist = true } = {}) {
  const sameShape = S.ds && S.ds.sig === ds.sig;
  S.ds = ds;
  if (!(keepFilters && sameShape)) {
    S.f = { cats: new Map(), ranges: new Map(), search: '' };
    $('#search').value = '';
    S.tables = {};
    S.pivot = { row: null, col: null, measure: 'count' };
  }
  if (persist) persistDataset(ds);
  setupAutoRefresh();
  applyFilters();
  renderChrome();
  if (persist && S.view === 'sources' && !keepFilters) go('overview'); else renderView();
}

async function persistDataset(ds) {
  const ok = await idb.set('ic.data', { name: ds.name, source: ds.source, loadedAt: ds.loadedAt, aoa: ds.aoa, imports: ds.imports || [] });
  if (!ok) toast('Loaded, but the browser refused to save it — it will be gone after reload.', true);
}

/** Add an uploaded table: merged into the stored data (unique by Booking Number) or replacing it. */
function ingestUpload(aoa, { name, sheet }) {
  const prev = S.ds, at = Date.now();
  const canMerge = prefs.uploadMode === 'merge' && prev && prev.source?.type === 'file';
  const merged = canMerge ? mergeAoa(prev.aoa, aoa) : null;
  if (merged) {
    const { stats } = merged;
    const imports = [...(prev.imports || []), { name, sheet, at, ...stats }];
    const ds = buildDataset(merged.aoa, { name: `Merged bookings · ${imports.length} uploads`, sheet, source: { type: 'file' }, imports, loadedAt: at });
    setDataset(ds, { keepFilters: true });
    toast(`${name}: ${fmt(stats.added)} new, ${fmt(stats.updated)} updated, ${fmt(stats.unchanged)} unchanged → ${fmt(stats.total)} unique bookings (${fmt(ds.n)} rows)`);
    return;
  }
  if (canMerge) toast(`${name} has no Booking Number column, so it replaced the saved data instead of merging.`, true);
  const ds = buildDataset(aoa, { name, sheet, source: { type: 'file' }, loadedAt: at });
  const k = keyIndex(ds.aoa[0]);
  const bookings = k < 0 ? 0 : new Set(ds.aoa.slice(1).map(r => cleanCell(r[k])).filter(Boolean)).size;
  ds.imports = [{ name, sheet, at, added: bookings, updated: 0, unchanged: 0, inFile: bookings, total: bookings, rows: ds.n }];
  setDataset(ds);
  if (!canMerge) toast(`Loaded ${fmt(ds.n)} rows from ${name}${bookings ? ` · ${fmt(bookings)} bookings` : ''}`);
}

/** Ask before anything other than a merge throws away saved uploads. */
function okToReplace(what) {
  const D = S.ds;
  if (!D || D.source?.type !== 'file') return true;
  return confirm(`Loading ${what} will replace your saved data (${fmt(D.n)} rows${D.imports?.length ? ` from ${D.imports.length} upload${D.imports.length > 1 ? 's' : ''}` : ''}). Continue?`);
}

async function loadFile(file) {
  if (prefs.uploadMode !== 'merge' && !okToReplace(`“${file.name}”`)) return;
  try {
    const name = file.name;
    if (/\.(xlsx|xls|xlsm|ods)$/i.test(name)) {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
      S.workbook = { wb, name };
      ingestUpload(sheetAoa(wb, wb.SheetNames[0]), { name, sheet: wb.SheetNames[0] });
    } else {
      S.workbook = null;
      ingestUpload(textToAoa(await file.text()), { name });
    }
  } catch (e) { console.error(e); toast('Could not read file: ' + e.message, true); }
}
const sheetAoa = (wb, sheet) => XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, raw: false, defval: '' });

function sheetCandidates(input) {
  const u = input.trim();
  const gid = (u.match(/[#&?]gid=(\d+)/) || [])[1];
  if (/\/spreadsheets\/d\/e\//.test(u)) {               // "Publish to web" link
    const base = u.replace(/\/pub(html)?([?#].*)?$/, '/pub').replace(/\/pubhtml.*$/, '/pub');
    return [`${base}?output=csv${gid ? `&gid=${gid}&single=true` : ''}`];
  }
  const id = (u.match(/\/spreadsheets\/d\/([\w-]+)/) || [])[1] || (/^[\w-]{25,}$/.test(u) ? u : null);
  if (id) return [
    `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid || 0}`,
    `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&gid=${gid || 0}`,
  ];
  if (/^https?:\/\//.test(u)) return [u];               // any other CSV URL
  return [];
}

async function loadSheet(url, { silent = false, name } = {}) {
  const cands = sheetCandidates(url);
  if (!cands.length) { toast('That does not look like a Google Sheets link.', true); return false; }
  if (!silent && !okToReplace('this Google Sheet')) return false;
  const btn = $('#refreshBtn'); btn.disabled = true;
  let lastErr;
  for (const c of cands) {
    try {
      const res = await fetch(c, { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      if (/^\s*<(!doctype|html)/i.test(text)) throw new Error('Sheet is not public');
      const ds = buildDataset(textToAoa(text), { name: name || S.ds?.source?.url === url && S.ds.name || 'Google Sheet', source: { type: 'gsheet', url } });
      setDataset(ds, { keepFilters: S.ds?.source?.url === url });
      if (!silent) toast(`Loaded ${fmt(ds.n)} rows from Google Sheets`);
      btn.disabled = false;
      return true;
    } catch (e) { lastErr = e; }
  }
  btn.disabled = false;
  console.warn(lastErr);
  if (!silent) toast('Could not read the sheet. Share it as "Anyone with the link → Viewer" (or File → Share → Publish to web) and try again.', true);
  return false;
}

function setupAutoRefresh() {
  clearInterval(S.refreshTimer);
  const isSheet = S.ds?.source?.type === 'gsheet';
  $('#refreshBtn').hidden = !isSheet;
  if (isSheet && prefs.refreshMin > 0) S.refreshTimer = setInterval(() => loadSheet(S.ds.source.url, { silent: true }), prefs.refreshMin * 60000);
}

// ════════════════════════════════════════════════════════════════
//  Chrome: nav, chips, filter panel, dataset info
// ════════════════════════════════════════════════════════════════
function renderChrome() {
  const D = S.ds;
  $$('#nav a[data-profile]').forEach(a => a.hidden = !(D && D.profile));
  const n = activeFilterCount();
  const fc = $('#filterCount'); fc.hidden = !n; fc.textContent = n;
  $('#dsInfo').innerHTML = D
    ? `<b>${esc(D.name)}</b>${fmt(S.filtered.length)} of ${fmt(D.n)} rows${D.source.type === 'gsheet' ? ' · Google Sheet' : ''}<br><span class="faint">Loaded ${fmtDate(D.loadedAt)}</span>`
    : '<b>No data yet</b>Upload a file or connect a Google Sheet.';
  renderChips();
  renderFilters();
}

function renderChips() {
  const D = S.ds, el = $('#chips');
  if (!D) { el.innerHTML = ''; return; }
  const chips = [];
  for (const [c, set] of S.f.cats) for (const v of set)
    chips.push(`<span class="chip" title="${esc(D.cols[c].label)}: ${esc(v || '(blank)')}"><span><b>${esc(shortLabel(D.cols[c].label, 22))}:</b> ${esc(shortLabel(v || '(blank)', 40))}</span><button data-chip-cat="${c}" data-v="${esc(v)}">${ICON.x}</button></span>`);
  for (const [c, rg] of S.f.ranges) {
    const col = D.cols[c], f = col.type === 'date' ? x => fmtDate(x, false) : x => fmt(x, 2).replace(/\.00$/, '');
    const txt = rg.min != null && rg.max != null ? `${f(rg.min)} → ${f(rg.max)}` : rg.min != null ? `≥ ${f(rg.min)}` : `≤ ${f(rg.max)}`;
    chips.push(`<span class="chip"><span><b>${esc(shortLabel(col.label, 22))}:</b> ${esc(txt)}</span><button data-chip-range="${c}">${ICON.x}</button></span>`);
  }
  if (S.f.search) chips.push(`<span class="chip"><span><b>Search:</b> ${esc(S.f.search)}</span><button data-chip-search>${ICON.x}</button></span>`);
  if (chips.length > 1) chips.push(`<button class="link" data-chip-clear>Clear all</button>`);
  el.innerHTML = chips.join('');
}

function renderFilters() {
  const D = S.ds, body = $('#filtersBody');
  if (!D) { body.innerHTML = '<p class="muted" style="padding:16px">Load data to see filters.</p>'; return; }
  const scroll = body.scrollTop;
  const html = [];
  const firstTime = !S.ui.facetsInit;
  if (firstTime) { S.ui.facetsInit = true; D.facetCols.slice(0, 3).forEach(c => S.ui.openFacets.add('c' + c)); }

  for (const c of D.facetCols) {
    const col = D.cols[c], sel = S.f.cats.get(c) || new Set(), counts = S.facetCounts.get(c) || new Map();
    const q = (S.ui.facetQ[c] || '').toLowerCase();
    let items = [...col.uniques.keys()].map(v => ({ v, n: counts.get(v) || 0 }));
    items.sort((a, b) => sel.has(b.v) - sel.has(a.v) || b.n - a.n || String(a.v).localeCompare(b.v, undefined, { numeric: true }));
    if (q) items = items.filter(it => String(it.v || '(blank)').toLowerCase().includes(q));
    const id = 'c' + c;
    html.push(`<details class="facet" data-facet="${id}" ${S.ui.openFacets.has(id) || sel.size ? 'open' : ''}>
      <summary><span>${esc(col.label)}</span>${sel.size ? `<span class="badge count">${sel.size}</span>` : `<span class="faint count">${col.uniques.size}</span>`}</summary>
      <div class="facet-body">
        ${col.uniques.size > 8 ? `<input class="input facet-search" placeholder="Search ${esc(col.label.toLowerCase())}…" data-facet-q="${c}" value="${esc(S.ui.facetQ[c] || '')}">` : ''}
        <div class="facet-list">
          ${items.map(it => `<label class="facet-item ${it.n ? '' : 'zero'}" title="${esc(it.v || '(blank)')}">
            <input type="checkbox" data-cat="${c}" data-v="${esc(it.v)}" ${sel.has(it.v) ? 'checked' : ''}>
            <span class="lbl">${esc(it.v || '(blank)')}</span><span class="c">${fmt(it.n)}</span></label>`).join('') || '<div class="faint" style="font-size:13px">No matches</div>'}
        </div>
        ${sel.size ? `<button class="link" data-clear-cat="${c}" style="margin-top:6px">Clear</button>` : ''}
      </div></details>`);
  }

  for (const c of D.rangeCols) {
    const col = D.cols[c], rg = S.f.ranges.get(c) || {}, id = 'r' + c, isDate = col.type === 'date';
    const v = x => x == null ? '' : isDate ? isoDay(x) : x;
    html.push(`<details class="facet" data-facet="${id}" ${S.ui.openFacets.has(id) || S.f.ranges.has(c) ? 'open' : ''}>
      <summary><span>${esc(col.label)}</span>${S.f.ranges.has(c) ? '<span class="badge count">1</span>' : `<span class="faint count">${isDate ? 'date' : '#'}</span>`}</summary>
      <div class="facet-body">
        <div class="range">
          <input class="input" type="${isDate ? 'date' : 'number'}" data-range="${c}" data-end="min" value="${v(rg.min)}" placeholder="${isDate ? '' : fmt(col.min)}" ${isDate ? `min="${isoDay(col.min)}" max="${isoDay(col.max)}"` : 'step="any"'}>
          <span class="faint">→</span>
          <input class="input" type="${isDate ? 'date' : 'number'}" data-range="${c}" data-end="max" value="${v(rg.max)}" placeholder="${isDate ? '' : fmt(col.max)}" ${isDate ? `min="${isoDay(col.min)}" max="${isoDay(col.max)}"` : 'step="any"'}>
        </div>
        ${isDate ? `<div class="quick">
          <button data-quick="${c}" data-q="today">Today</button><button data-quick="${c}" data-q="next7">Next 7 days</button>
          <button data-quick="${c}" data-q="last7">Last 7 days</button><button data-quick="${c}" data-q="month">This month</button>
        </div><div class="faint" style="font-size:12px;margin-top:6px">Data: ${fmtDate(col.min, false)} – ${fmtDate(col.max, false)}</div>` : ''}
        ${S.f.ranges.has(c) ? `<button class="link" data-clear-range="${c}" style="margin-top:6px">Clear</button>` : ''}
      </div></details>`);
  }
  body.innerHTML = html.join('') || '<p class="muted" style="padding:16px">No filterable columns detected.</p>';
  body.scrollTop = scroll;
  if (S.ui.focusFacet != null) {
    const inp = $(`[data-facet-q="${S.ui.focusFacet}"]`, body);
    if (inp) { inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length); }
    S.ui.focusFacet = null;
  }
}

// ════════════════════════════════════════════════════════════════
//  Generic sortable / paged grid
// ════════════════════════════════════════════════════════════════
function cmp(a, b) {
  const an = a == null || a === '' || (typeof a === 'number' && isNaN(a));
  const bn = b == null || b === '' || (typeof b === 'number' && isNaN(b));
  if (an || bn) return an - bn;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}

/**
 * columns: [{ id, label, value(row) -> sortable, render?(row) -> html, align:'n' }]
 */
function renderGrid(el, id, { columns, rows, onRow, toolbar = '', empty = 'No rows match your filters.', after }) {
  const st = S.tables[id] ||= { sort: null, dir: 1, page: 1, size: 50 };
  let data = rows;
  const sc = columns.find(c => c.id === st.sort);
  if (sc) { const f = sc.sort || sc.value; data = rows.map(r => [f(r), r]).sort((a, b) => cmp(a[0], b[0]) * st.dir).map(x => x[1]); }
  const pages = Math.max(1, Math.ceil(data.length / st.size));
  st.page = Math.min(st.page, pages);
  const start = (st.page - 1) * st.size, slice = data.slice(start, start + st.size);
  el.innerHTML = `<div class="card">${toolbar}
    <div class="table-wrap"><table class="data"><thead><tr>
      ${columns.map(c => `<th class="${c.align || ''}" data-sort="${esc(c.id)}" title="${esc(c.label)}">${esc(c.label)}${st.sort === c.id ? `<span class="arrow">${st.dir > 0 ? '▲' : '▼'}</span>` : ''}</th>`).join('')}
    </tr></thead><tbody>
      ${slice.length ? slice.map((r, i) => `<tr data-i="${i}">${columns.map(c => `<td class="${c.align || ''}">${c.render ? c.render(r) : esc(c.value(r))}</td>`).join('')}</tr>`).join('')
      : `<tr><td class="empty-row" colspan="${columns.length}">${empty}</td></tr>`}
    </tbody></table></div>
    <div class="pager">
      <span>Rows per page</span>
      <select class="select" data-size>${[25, 50, 100, 250, 1000].map(s => `<option ${s === st.size ? 'selected' : ''}>${s}</option>`).join('')}</select>
      <span class="spacer"></span>
      <span class="num">${data.length ? `${fmt(start + 1)}–${fmt(start + slice.length)} of ${fmt(data.length)}` : '0 rows'}</span>
      <button class="btn sm" data-page="-1" ${st.page <= 1 ? 'disabled' : ''}>‹ Prev</button>
      <button class="btn sm" data-page="1" ${st.page >= pages ? 'disabled' : ''}>Next ›</button>
    </div></div>`;
  $$('th[data-sort]', el).forEach(th => th.onclick = () => {
    const c = th.dataset.sort;
    if (st.sort === c) { if (st.dir > 0) st.dir = -1; else { st.sort = null; st.dir = 1; } } else { st.sort = c; st.dir = 1; }
    renderGrid(el, id, { columns, rows, onRow, toolbar, empty, after });
  });
  $$('[data-page]', el).forEach(b => b.onclick = () => { st.page += +b.dataset.page; renderGrid(el, id, { columns, rows, onRow, toolbar, empty, after }); el.querySelector('.table-wrap').scrollTop = 0; });
  $('[data-size]', el).onchange = e => { st.size = +e.target.value; st.page = 1; renderGrid(el, id, { columns, rows, onRow, toolbar, empty, after }); };
  if (onRow) $$('tbody tr[data-i]', el).forEach(tr => tr.onclick = () => onRow(slice[+tr.dataset.i]));
  after && after(el);
}

function highlighter() {
  const terms = parseSearch(S.f.search).filter(t => !t.neg).map(t => t.term).filter(t => t.length > 1);
  if (!terms.length) return esc;
  const re = new RegExp('(' + terms.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')', 'gi');
  return s => esc(s).replace(re, '<mark>$1</mark>');
}

// ════════════════════════════════════════════════════════════════
//  Charts
// ════════════════════════════════════════════════════════════════
const PALETTE = ['#2563eb', '#12b76a', '#f79009', '#7a5af8', '#06aed4', '#ee46bc', '#f04438', '#84cc16', '#64748b', '#eab308', '#0e9384', '#c11574'];
const toneColor = v => ({ ok: '#12b76a', bad: '#f04438', warn: '#f79009' }[toneOf(v)]);

function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
function chartDefaults() {
  Chart.defaults.color = cssVar('--text-2');
  Chart.defaults.borderColor = cssVar('--border');
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.font.size = 12;
  Chart.defaults.plugins.legend.labels.boxWidth = 10;
  Chart.defaults.plugins.legend.labels.boxHeight = 10;
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.maintainAspectRatio = false;
}
function destroyCharts() { for (const k in S.charts) { S.charts[k].destroy(); delete S.charts[k]; } }

function countBy(rows, c, weight) {
  const m = new Map();
  for (const r of rows) { const v = val(c, r); m.set(v, (m.get(v) || 0) + (weight ? weight(r) : 1)); }
  return [...m].sort((a, b) => b[1] - a[1]);
}

function doughnut(id, entries, onClick, { colorFn } = {}) {
  const cv = document.getElementById(id); if (!cv) return;
  const labels = entries.map(e => e[0] || '(blank)');
  S.charts[id] = new Chart(cv, {
    type: 'doughnut',
    data: { labels: labels.map(l => shortLabel(l, 26)), datasets: [{ data: entries.map(e => e[1]), backgroundColor: entries.map((e, i) => (colorFn && colorFn(e[0])) || PALETTE[i % PALETTE.length]), borderColor: cssVar('--surface'), borderWidth: 2 }] },
    options: {
      cutout: '64%',
      plugins: { legend: { position: 'right' }, tooltip: { callbacks: { title: it => labels[it[0].dataIndex] } } },
      onClick: (e, els) => els.length && onClick && onClick(entries[els[0].index][0]),
      onHover: (e, els) => { if (e.native) e.native.target.style.cursor = els.length && onClick ? 'pointer' : 'default'; },
    },
  });
}

function bars(id, entries, onClick, { horizontal = true, color = PALETTE[0], label = 'Count', stacks } = {}) {
  const cv = document.getElementById(id); if (!cv) return;
  const labels = entries.map(e => e[0] || '(blank)');
  const datasets = stacks
    ? stacks.map((s, i) => ({ label: s.label, data: s.data, backgroundColor: s.color || PALETTE[i % PALETTE.length], borderRadius: 3, maxBarThickness: 28 }))
    : [{ label, data: entries.map(e => e[1]), backgroundColor: color, borderRadius: 4, maxBarThickness: 22 }];
  S.charts[id] = new Chart(cv, {
    type: 'bar',
    data: { labels: labels.map(l => shortLabel(l, horizontal ? 28 : 14)), datasets },
    options: {
      indexAxis: horizontal ? 'y' : 'x',
      scales: {
        x: { stacked: !!stacks, grid: { display: !horizontal }, ticks: { precision: 0 } },
        y: { stacked: !!stacks, grid: { display: horizontal ? false : true }, ticks: { precision: 0, autoSkip: !horizontal } },
      },
      plugins: { legend: { display: !!stacks, position: 'bottom' }, tooltip: { callbacks: { title: it => labels[it[0].dataIndex] } } },
      onClick: (e, els) => els.length && onClick && onClick(entries[els[0].index][0], els[0].datasetIndex),
      onHover: (e, els) => { if (e.native) e.native.target.style.cursor = els.length && onClick ? 'pointer' : 'default'; },
    },
  });
}

/** Timeline of counts by day (or month) for a date column, optionally stacked by a category column. */
function timeline(id, dateCol, rows, stackCol) {
  const col = S.ds.cols[dateCol];
  const dayRows = rows.filter(r => !isNaN(col.ts[r]));
  if (!dayRows.length) return false;
  let tmin = Infinity, tmax = -Infinity;
  for (const r of dayRows) { const t = col.ts[r]; if (t < tmin) tmin = t; if (t > tmax) tmax = t; }
  const byMonth = (tmax - tmin) / 864e5 > 120;
  const key = ts => { const d = new Date(ts); return byMonth ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}` : dayKey(ts); };
  const keys = [...new Set(dayRows.map(r => key(col.ts[r])))].sort();
  const lbl = k => { const [y, m, d] = k.split('-'); return byMonth ? `${MONTHS[+m - 1]} ${y.slice(2)}` : `${d} ${MONTHS[+m - 1]}`; };
  const stackVals = stackCol != null ? countBy(dayRows, stackCol).map(e => e[0]).slice(0, 8) : [null];
  const kIdx = new Map(keys.map((k, i) => [k, i]));
  const sIdx = new Map(stackVals.map((v, i) => [v, i]));
  const stacks = stackVals.map((sv, i) => ({ label: sv ?? 'Count', color: PALETTE[i % PALETTE.length], data: new Array(keys.length).fill(0) }));
  for (const r of dayRows) {
    const si = stackCol == null ? 0 : sIdx.get(val(stackCol, r));
    if (si != null) stacks[si].data[kIdx.get(key(col.ts[r]))]++;
  }
  bars(id, keys.map(k => [lbl(k), 0]), null, { horizontal: false, stacks: stackCol != null ? stacks : null, label: 'Rows' });
  const ch = S.charts[id];
  if (stackCol == null) ch.data.datasets[0].data = stacks[0].data, ch.update();
  ch.options.onClick = (e, els) => {
    if (!els.length) return;
    const k = keys[els[0].index];
    const [y, m, d] = k.split('-').map(Number);
    const from = new Date(y, m - 1, d || 1).getTime();
    const to = byMonth ? new Date(y, m, 1).getTime() - 1 : from + 864e5 - 1;
    S.f.ranges.set(dateCol, { min: from, max: to });
    if (stackCol != null && els[0].datasetIndex != null && stackVals[els[0].datasetIndex] != null) onlyCat(stackCol, stackVals[els[0].datasetIndex]);
    refresh();
  };
  return true;
}

const filterTo = c => v => { setCat(c, v, true); refresh(); };

// ════════════════════════════════════════════════════════════════
//  Views
// ════════════════════════════════════════════════════════════════
function renderView() {
  destroyCharts();
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.view === S.view));
  $$('.view').forEach(v => v.classList.toggle('active', v.id === 'view-' + S.view));
  const el = $('#view-' + S.view);
  if (!S.ds && S.view !== 'sources') { renderWelcome(el); return; }
  chartDefaults();
  ({ overview: renderOverview, table: renderTable, bookings: renderBookings, voyages: renderVoyages, pivot: renderPivot, sources: renderSources })[S.view](el);
}

function go(view) {
  if (!S.ds?.profile && (view === 'bookings' || view === 'voyages')) view = 'overview';
  S.view = view;
  if (location.hash.slice(1) !== view) history.replaceState(null, '', '#' + view);
  $('#app').classList.remove('nav-open');
  renderView();
}

const kpi = (label, value, sub = '') => `<div class="card kpi"><div class="label">${label}</div><div class="value">${value}</div><div class="sub">${sub}</div></div>`;
const card = (title, body, span = 6, sub = '') => `<div class="card span-${span}"><div class="card-head"><h3>${title}</h3><small>${sub}</small></div><div class="card-body">${body}</div></div>`;
const chartCard = (id, title, span = 6, sub = '', tall = false) => card(title, `<div class="chart-box ${tall ? 'tall' : ''}"><canvas id="${id}"></canvas></div>`, span, sub);

// ── Overview ─────────────────────────────────────────────────────
function renderOverview(el) {
  const D = S.ds, rows = S.filtered;
  const head = `<div class="page-head"><div><h1>Overview</h1><p>${esc(D.name)} · ${fmt(rows.length)} of ${fmt(D.n)} rows${activeFilterCount() || S.f.search ? ' (filtered)' : ''} · click any chart to filter</p></div></div>`;
  if (D.profile) return renderBookingOverview(el, head);

  // Generic overview
  const nums = D.cols.filter(c => c.type === 'number').slice(0, 4);
  const cats = D.facetCols.map(c => D.cols[c]).filter(c => c.uniques.size <= 40).slice(0, 6);
  const dateCol = D.cols.find(c => c.type === 'date' && c.isRange);
  el.innerHTML = head + `<div class="kpis">
      ${kpi('Rows', fmt(rows.length), `of ${fmt(D.n)} total`)}
      ${kpi('Columns', fmt(D.cols.length), `${D.facetCols.length} filterable`)}
      ${nums.map(c => { let s = 0, k = 0; for (const r of rows) { const x = c.num[r]; if (!isNaN(x)) { s += x; k++; } } return kpi('Σ ' + esc(shortLabel(c.label, 24)), fmt(s, s % 1 ? 2 : 0), `avg ${fmt(k ? s / k : NaN, 2)}`); }).join('')}
    </div>
    <div class="grid">
      ${dateCol ? chartCard('g-time', esc(dateCol.label), 12, 'rows per period') : ''}
      ${cats.map((c, i) => chartCard('g-c' + i, esc(c.label), 6, `${c.uniques.size} values`)).join('')}
    </div>`;
  if (dateCol) timeline('g-time', dateCol.i, rows, cats.length && cats[0].uniques.size <= 8 ? cats[0].i : null);
  cats.forEach((c, i) => {
    const e = countBy(rows, c.i);
    e.length <= 6 ? doughnut('g-c' + i, e, filterTo(c.i), { colorFn: toneColor }) : bars('g-c' + i, e.slice(0, 12), filterTo(c.i), { color: PALETTE[i % PALETTE.length] });
  });
}

function bookingStats(rows) {
  const D = S.ds, P = D.profile;
  const bk = new Map();
  let assigned = 0, teu = 0, weight = 0, atPol = 0, sailed = 0, live = 0;
  for (const r of rows) {
    const b = val(P.booking, r);
    if (!bk.has(b)) bk.set(b, P.status != null ? val(P.status, r) : '');
    if (val(P.assigned, r) === 'Yes') assigned++;
    if (P.teu != null) teu += +val(P.teu, r) || 0;
    if (P.weight != null) { const w = numv(P.weight, r); if (!isNaN(w)) weight += w; }
    const cancelled = P.status != null && /cancel/i.test(val(P.status, r));
    if (!cancelled) {
      live++;
      const rk = P.event != null ? stageRank(val(P.event, r)) : 0;
      if (rk >= 4) atPol++;
      if (rk >= 6) sailed++;
    }
  }
  const status = new Map();
  for (const s of bk.values()) status.set(s, (status.get(s) || 0) + 1);
  return { bookings: bk.size, status, assigned, teu, weight, atPol, sailed, live };
}

function renderBookingOverview(el, head) {
  const D = S.ds, P = D.profile, rows = S.filtered, st = bookingStats(rows);
  const statusTxt = [...st.status].sort((a, b) => b[1] - a[1]).map(([s, n]) => `${fmt(n)} ${esc(s.toLowerCase() || 'blank')}`).join(' · ');
  const today = startOfDay(Date.now());
  const upcoming = voyageGroups(rows).filter(v => v.etd >= today || isNaN(v.etd)).sort((a, b) => cmp(a.etd, b.etd)).slice(0, 8);

  el.innerHTML = head + `<div class="kpis">
      ${kpi('Bookings', fmt(st.bookings), statusTxt)}
      ${kpi('Containers', fmt(rows.length), `${fmt(st.teu)} TEU`)}
      ${kpi('Containers assigned', fmt(st.assigned), `${pct(st.assigned, rows.length)}% of lines · ${fmt(rows.length - st.assigned)} unassigned`)}
      ${kpi('At port of load', fmt(st.atPol), `${pct(st.atPol, st.live)}% of active containers`)}
      ${kpi('Sailed', fmt(st.sailed), `${pct(st.sailed, st.live)}% of active containers`)}
      ${kpi('Cargo weight', fmt(st.weight / 1000, 1) + ' t', `avg ${fmt(rows.length ? st.weight / rows.length / 1000 : NaN, 1)} t / container`)}
    </div>
    <div class="grid">
      ${P.etd != null ? chartCard('b-time', 'Departures (ETD)', 8, 'containers per day, by size/type') : ''}
      ${P.status != null ? chartCard('b-status', 'Booking status', 4, 'unique bookings') : ''}
      ${P.event != null ? chartCard('b-stage', 'Container milestone', 6, 'latest event per container (active)') : ''}
      ${P.size != null ? chartCard('b-size', 'Container size / type', 6, 'containers') : ''}
      ${P.pod != null ? chartCard('b-pod', 'Top ports of discharge', 6, 'containers', true) : ''}
      ${P.voyage != null ? chartCard('b-voy', 'Vessel / voyage', 6, 'containers', true) : ''}
      ${P.shipper != null ? chartCard('b-ship', 'Top shippers', 6, 'containers', true) : ''}
      ${P.forwarder != null ? chartCard('b-fwd', 'Forwarders', 6, 'containers', true) : ''}
      <div class="card span-12" id="b-upcoming"></div>
    </div>`;

  if (P.etd != null) timeline('b-time', P.etd, rows, P.size);
  if (P.status != null) {
    const seen = new Set(), e = new Map();
    for (const r of rows) { const b = val(P.booking, r); if (seen.has(b)) continue; seen.add(b); const s = val(P.status, r); e.set(s, (e.get(s) || 0) + 1); }
    doughnut('b-status', [...e].sort((a, b) => b[1] - a[1]), filterTo(P.status), { colorFn: toneColor });
  }
  if (P.event != null) {
    const active = rows.filter(r => !(P.status != null && /cancel/i.test(val(P.status, r))));
    const e = countBy(active, P.stage);
    const order = ['No event', ...STAGES.slice(1).map(s => s[2]), 'Other'];
    e.sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0]));
    bars('b-stage', e, filterTo(P.stage), { color: PALETTE[1] });
  }
  if (P.size != null) doughnut('b-size', countBy(rows, P.size), filterTo(P.size));
  if (P.pod != null) bars('b-pod', countBy(rows, P.pod).slice(0, 12), filterTo(P.pod), { color: PALETTE[0] });
  if (P.voyage != null) bars('b-voy', countBy(rows, P.voyage).slice(0, 12), filterTo(P.voyage), { color: PALETTE[3] });
  if (P.shipper != null) bars('b-ship', countBy(rows, P.shipper).slice(0, 12), filterTo(P.shipper), { color: PALETTE[4] });
  if (P.forwarder != null) bars('b-fwd', countBy(rows, P.forwarder).slice(0, 12), filterTo(P.forwarder), { color: PALETTE[2] });

  const up = $('#b-upcoming');
  up.innerHTML = `<div class="card-head"><h3>Upcoming departures</h3><button class="link" data-go="voyages">All voyages →</button></div>
    <div class="card-body" style="overflow:auto">${upcoming.length ? `<table class="mini-table"><thead><tr><th>Vessel / voyage</th><th>Service</th><th>POL</th><th>ETD</th><th class="n">Bookings</th><th class="n">Containers</th><th>Ready at POL</th></tr></thead><tbody>
    ${upcoming.map(v => `<tr data-voy="${esc(v.voyage)}"><td><b>${esc(v.voyage)}</b></td><td>${esc(v.service)}</td><td>${esc(v.pol)}</td><td>${fmtDate(v.etd)} ${pill(v.etdType)}</td><td class="n">${fmt(v.bookings.size)}</td><td class="n">${fmt(v.rows.length)}</td>
      <td><div class="row" style="flex-wrap:nowrap"><div class="progress" style="flex:1"><div style="width:${pct(v.atPol, v.live)}%"></div></div><span class="faint num">${v.atPol}/${v.live}</span></div></td></tr>`).join('')}
    </tbody></table>` : '<p class="muted">No upcoming departures in the current selection.</p>'}</div>`;
  $$('[data-voy]', up).forEach(tr => tr.onclick = () => { onlyCat(P.voyage, tr.dataset.voy); refresh(); go('bookings'); });
}

// ── Data table ───────────────────────────────────────────────────
function defaultHidden(D) {
  if (!D.profile) return D.cols.length > 14 ? D.cols.filter(c => c.uniques.size === 1).map(c => c.label) : [];
  return D.cols.filter(c => /time zone|unit$|^total container quantity of|^teu$|^container assigned$|^milestone$/i.test(c.label) || (c.uniques.size === 1 && !c.derived)).map(c => c.label);
}
const hiddenSet = () => new Set(prefs.hidden[S.ds.sig] ?? defaultHidden(S.ds));

function renderTable(el) {
  const D = S.ds, hidden = hiddenSet(), hl = highlighter();
  const cols = D.cols.filter(c => !hidden.has(c.label));
  const columns = cols.map(c => ({
    id: 'c' + c.i, label: c.label, align: c.type === 'number' ? 'n' : '',
    value: r => c.values[r],
    sort: c.type === 'number' ? r => c.num[r] : c.type === 'date' ? r => c.ts[r] : r => c.values[r],
    render: (D.profile && c.i === D.profile.status) || /status/i.test(c.label) && c.isCat ? r => pill(c.values[r]) : r => hl(c.values[r]),
  }));
  const toolbar = `<div class="toolbar"><b>${fmt(S.filtered.length)} rows</b><span class="faint">· ${cols.length} of ${D.cols.length} columns</span><span class="spacer"></span>
    <div class="dropdown"><button class="btn sm" data-colmenu>${ICON.cols}Columns</button>
      <div class="dropdown-menu col-menu">
        <div class="row" style="padding:4px 6px 8px"><button class="btn sm" data-cols="all">Show all</button><button class="btn sm" data-cols="reset">Reset</button></div>
        ${D.cols.map(c => `<label><input type="checkbox" data-col="${esc(c.label)}" ${hidden.has(c.label) ? '' : 'checked'}> ${esc(c.label)}${c.derived ? ' <span class="faint">(derived)</span>' : ''}</label>`).join('')}
      </div></div></div>`;
  renderGrid(el, 'data', {
    columns, rows: S.filtered, toolbar, onRow: openRow,
    after: g => {
      const menu = $('.col-menu', g);
      $('[data-colmenu]', g).onclick = e => { e.stopPropagation(); menu.classList.toggle('open'); };
      menu.onclick = e => e.stopPropagation();
      $$('[data-col]', g).forEach(cb => cb.onchange = () => {
        const h = hiddenSet(); cb.checked ? h.delete(cb.dataset.col) : h.add(cb.dataset.col);
        prefs.hidden[D.sig] = [...h]; savePrefs(); renderTable(el); $('.col-menu', el).classList.add('open');
      });
      $$('[data-cols]', g).forEach(b => b.onclick = () => {
        if (b.dataset.cols === 'all') prefs.hidden[D.sig] = []; else delete prefs.hidden[D.sig];
        savePrefs(); renderTable(el); $('.col-menu', el).classList.add('open');
      });
    },
  });
}

// ── Bookings (grouped) ───────────────────────────────────────────
function bookingGroups(rows) {
  const P = S.ds.profile, m = new Map();
  for (const r of rows) {
    const b = val(P.booking, r);
    let g = m.get(b);
    if (!g) m.set(b, g = { booking: b, rows: [], first: r });
    g.rows.push(r);
  }
  for (const g of m.values()) {
    const r = g.first, get = c => c != null ? val(c, r) : '';
    g.status = get(P.status); g.shipper = get(P.shipper); g.forwarder = get(P.forwarder);
    g.voyage = get(P.voyage); g.pol = get(P.pol); g.pod = get(P.pod);
    g.etd = P.etd != null ? tsv(P.etd, r) : NaN; g.etdType = get(P.etdType);
    g.assigned = g.rows.filter(x => val(P.assigned, x) === 'Yes').length;
    g.weight = P.weight != null ? g.rows.reduce((s, x) => s + (numv(P.weight, x) || 0), 0) : NaN;
    const sizes = new Map(); g.rows.forEach(x => { const s = P.size != null ? val(P.size, x) : ''; sizes.set(s, (sizes.get(s) || 0) + 1); });
    g.sizes = [...sizes].map(([s, n]) => `${n}×${s}`).join(', ');
    g.ranks = g.rows.map(x => P.event != null ? stageRank(val(P.event, x)) : 0);
    g.atPol = g.ranks.filter(k => k >= 4).length;
    g.best = Math.max(...g.ranks);
    let latest = NaN, latestEv = '';
    if (P.eventAt != null) g.rows.forEach(x => { const t = tsv(P.eventAt, x); if (!isNaN(t) && !(t <= latest)) { latest = t; latestEv = val(P.event, x); } });
    g.latestAt = latest; g.latestEv = latestEv || (P.event != null ? val(P.event, r) : '');
  }
  return [...m.values()];
}

function renderBookings(el) {
  const groups = bookingGroups(S.filtered), hl = highlighter();
  el.innerHTML = `<div class="page-head"><div><h1>Bookings</h1><p>${fmt(groups.length)} bookings · ${fmt(S.filtered.length)} container lines · click a booking for details</p></div></div><div id="bk-grid"></div>`;
  renderGrid($('#bk-grid'), 'bookings', {
    rows: groups, onRow: g => openBooking(g.booking),
    columns: [
      { id: 'b', label: 'Booking #', value: g => g.booking, render: g => `<b>${hl(g.booking)}</b>` },
      { id: 's', label: 'Status', value: g => g.status, render: g => pill(g.status) },
      { id: 'sh', label: 'Shipper', value: g => g.shipper, render: g => hl(shortLabel(g.shipper, 34)) },
      { id: 'v', label: 'Vessel / voyage', value: g => g.voyage, render: g => hl(g.voyage) },
      { id: 'r', label: 'Route', value: g => g.pol + ' → ' + g.pod, render: g => `${hl(g.pol)} <span class="faint">→</span> ${hl(g.pod)}` },
      { id: 'e', label: 'ETD', value: g => g.etd, render: g => `${fmtDate(g.etd)} ${g.etdType ? `<span class="pill ${toneOf(g.etdType)}">${esc(g.etdType[0])}</span>` : ''}` },
      { id: 'c', label: 'Containers', align: 'n', value: g => g.rows.length, render: g => `<span title="${esc(g.sizes)}">${g.rows.length} <span class="faint">${esc(shortLabel(g.sizes, 22))}</span></span>` },
      { id: 'a', label: 'Assigned', align: 'n', value: g => g.assigned / g.rows.length, render: g => `${g.assigned}/${g.rows.length}` },
      { id: 'p', label: 'At POL', value: g => g.atPol / g.rows.length, render: g => `<div class="row" style="flex-wrap:nowrap"><div class="progress" style="width:70px"><div style="width:${pct(g.atPol, g.rows.length)}%"></div></div><span class="faint num">${g.atPol}/${g.rows.length}</span></div>` },
      { id: 'w', label: 'Weight (t)', align: 'n', value: g => g.weight, render: g => fmt(g.weight / 1000, 1) },
      { id: 'l', label: 'Latest event', value: g => g.latestAt, render: g => `${esc(g.latestEv)}<br><span class="faint">${isNaN(g.latestAt) ? '' : fmtDate(g.latestAt)}</span>` },
      { id: 'f', label: 'Forwarder', value: g => g.forwarder, render: g => hl(shortLabel(g.forwarder, 30)) },
    ],
  });
}

// ── Voyages ──────────────────────────────────────────────────────
function voyageGroups(rows) {
  const P = S.ds.profile, m = new Map();
  if (P.voyage == null) return [];
  for (const r of rows) {
    const v = val(P.voyage, r);
    let g = m.get(v);
    if (!g) m.set(v, g = { voyage: v, rows: [], bookings: new Set(), etd: NaN, etdType: '', service: P.service != null ? val(P.service, r) : '', pol: P.pol != null ? val(P.pol, r) : '', pods: new Set(), teu: 0, weight: 0, live: 0, picked: 0, gated: 0, atPol: 0, sailed: 0 });
    g.rows.push(r); g.bookings.add(val(P.booking, r));
    if (P.pod != null) g.pods.add(val(P.pod, r));
    if (P.etd != null) { const t = tsv(P.etd, r); if (!isNaN(t) && (isNaN(g.etd) || t < g.etd)) { g.etd = t; g.etdType = P.etdType != null ? val(P.etdType, r) : ''; } }
    if (P.teu != null) g.teu += +val(P.teu, r) || 0;
    if (P.weight != null) g.weight += numv(P.weight, r) || 0;
    if (!(P.status != null && /cancel/i.test(val(P.status, r)))) {
      g.live++;
      const k = P.event != null ? stageRank(val(P.event, r)) : 0;
      if (k >= 1) g.picked++; if (k >= 2) g.gated++; if (k >= 4) g.atPol++; if (k >= 6) g.sailed++;
    }
  }
  return [...m.values()];
}

function renderVoyages(el) {
  const P = S.ds.profile, groups = voyageGroups(S.filtered), hl = highlighter();
  el.innerHTML = `<div class="page-head"><div><h1>Voyages</h1><p>${fmt(groups.length)} vessel voyages · readiness counts exclude cancelled bookings · click to see its bookings</p></div></div><div id="vy-grid"></div>`;
  const prog = (a, b) => `<div class="row" style="flex-wrap:nowrap"><div class="progress" style="width:70px"><div style="width:${pct(a, b)}%"></div></div><span class="faint num">${a}/${b}</span></div>`;
  S.tables.voyages ||= { sort: 'e', dir: 1, page: 1, size: 50 };
  renderGrid($('#vy-grid'), 'voyages', {
    rows: groups, onRow: g => { onlyCat(P.voyage, g.voyage); refresh(); go('bookings'); },
    columns: [
      { id: 'v', label: 'Vessel / voyage', value: g => g.voyage, render: g => `<b>${hl(g.voyage)}</b>` },
      { id: 's', label: 'Service', value: g => g.service, render: g => hl(g.service) },
      { id: 'p', label: 'POL', value: g => g.pol, render: g => hl(g.pol) },
      { id: 'e', label: 'ETD', value: g => g.etd, render: g => `${fmtDate(g.etd)} ${pill(g.etdType)}` },
      { id: 'd', label: 'Discharge ports', value: g => g.pods.size, render: g => `<span title="${esc([...g.pods].join(', '))}">${esc(shortLabel([...g.pods].join(', '), 34))}</span>` },
      { id: 'b', label: 'Bookings', align: 'n', value: g => g.bookings.size },
      { id: 'c', label: 'Containers', align: 'n', value: g => g.rows.length },
      { id: 't', label: 'TEU', align: 'n', value: g => g.teu },
      { id: 'w', label: 'Weight (t)', align: 'n', value: g => g.weight, render: g => fmt(g.weight / 1000, 1) },
      { id: 'pk', label: 'Empty picked', value: g => g.picked / (g.live || 1), render: g => prog(g.picked, g.live) },
      { id: 'gt', label: 'Gated in', value: g => g.gated / (g.live || 1), render: g => prog(g.gated, g.live) },
      { id: 'ap', label: 'At POL', value: g => g.atPol / (g.live || 1), render: g => prog(g.atPol, g.live) },
      { id: 'sl', label: 'Sailed', value: g => g.sailed / (g.live || 1), render: g => prog(g.sailed, g.live) },
    ],
  });
}

// ── Pivot ────────────────────────────────────────────────────────
function renderPivot(el) {
  const D = S.ds, pv = S.pivot;
  const groupable = D.cols.filter(c => c.type !== 'number' && c.uniques.size <= 2000 && c.uniques.size > 1);
  const colable = D.cols.filter(c => c.isCat && c.uniques.size <= 15);
  const numeric = D.cols.filter(c => c.type === 'number');
  if (pv.row == null || !groupable.some(c => c.i === pv.row)) pv.row = (D.profile?.pod ?? D.facetCols[0] ?? groupable[0]?.i) ?? null;
  if (pv.col != null && !colable.some(c => c.i === pv.col)) pv.col = null;
  if (pv.row == null) { el.innerHTML = '<p class="muted">No columns suitable for grouping.</p>'; return; }

  const measures = [['count', 'Count rows'], ...D.cols.filter(c => c.type === 'text' && c.uniques.size > 1).map(c => ['u' + c.i, 'Count unique ' + c.label]), ...numeric.map(c => ['s' + c.i, 'Sum of ' + c.label]), ...numeric.map(c => ['a' + c.i, 'Average of ' + c.label])];
  if (!measures.some(m => m[0] === pv.measure)) pv.measure = 'count';
  const opt = (list, cur, none) => (none ? `<option value="">${none}</option>` : '') + list.map(c => `<option value="${c.i}" ${c.i === cur ? 'selected' : ''}>${esc(c.label)}</option>`).join('');

  // compute
  const rows = S.filtered, mk = pv.measure, mi = mk === 'count' ? null : +mk.slice(1), kind = mk[0];
  const colVals = pv.col != null ? countBy(rows, pv.col).map(e => e[0]) : [];
  const cells = new Map();   // rowVal -> Map(colVal|'__t' -> accumulator)
  const acc = () => kind === 'u' ? new Set() : { s: 0, n: 0 };
  const add = (a, r) => { if (kind === 'u') a.add(val(mi, r)); else if (kind === 'c') a.n++; else { const x = numv(mi, r); if (!isNaN(x)) { a.s += x; a.n++; } } };
  const out = a => !a ? NaN : kind === 'u' ? a.size : kind === 'c' ? a.n : kind === 's' ? a.s : (a.n ? a.s / a.n : NaN);
  const total = acc(), colTotals = new Map(colVals.map(v => [v, acc()]));
  for (const r of rows) {
    const rv = val(pv.row, r);
    let m = cells.get(rv); if (!m) cells.set(rv, m = new Map([['__t', acc()]]));
    add(m.get('__t'), r); add(total, r);
    if (pv.col != null) { const cv = val(pv.col, r); if (!m.has(cv)) m.set(cv, acc()); add(m.get(cv), r); add(colTotals.get(cv), r); }
  }
  const dec = kind === 'a' ? 2 : (kind === 's' && numeric.find(c => c.i === mi)?.values.some(v => /\.\d*[1-9]/.test(v))) ? 2 : 0;
  const list = [...cells].map(([rv, m]) => ({ rv, m, t: out(m.get('__t')) })).sort((a, b) => cmp(b.t, a.t));

  el.innerHTML = `<div class="page-head"><div><h1>Pivot</h1><p>Summarise the filtered rows · click a cell to drill down</p></div></div>
    <div class="card" style="margin-bottom:16px"><div class="toolbar" style="border:0">
      <label class="muted">Rows</label><select class="select" id="pvRow">${opt(groupable, pv.row)}</select>
      <label class="muted">Columns</label><select class="select" id="pvCol">${opt(colable, pv.col, '— none —')}</select>
      <label class="muted">Value</label><select class="select" id="pvMeasure">${measures.map(([k, l]) => `<option value="${k}" ${k === mk ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>
    </div></div>
    <div class="grid">
      <div class="card span-12"><div class="table-wrap" style="max-height:520px"><table class="data" id="pvTable"><thead><tr>
        <th>${esc(D.cols[pv.row].label)}</th>${colVals.map(v => `<th class="n">${esc(shortLabel(v || '(blank)', 22))}</th>`).join('')}<th class="n">Total</th></tr></thead>
        <tbody>${list.map(({ rv, m, t }) => `<tr><td data-r="${esc(rv)}"><b>${esc(rv || '(blank)')}</b></td>${colVals.map(cv => `<td class="n" data-r="${esc(rv)}" data-c="${esc(cv)}">${m.has(cv) ? fmt(out(m.get(cv)), dec) : '<span class="faint">·</span>'}</td>`).join('')}<td class="n" data-r="${esc(rv)}"><b>${fmt(t, dec)}</b></td></tr>`).join('')}
        <tr><td><b>Total</b></td>${colVals.map(cv => `<td class="n"><b>${fmt(out(colTotals.get(cv)), dec)}</b></td>`).join('')}<td class="n"><b>${fmt(out(total), dec)}</b></td></tr></tbody>
      </table></div></div>
      ${card('Top ' + esc(D.cols[pv.row].label), `<div class="chart-box" style="height:${Math.max(260, Math.min(list.length, 20) * 26 + 60)}px"><canvas id="pv-chart"></canvas></div>`, 12, measures.find(m => m[0] === mk)[1])}
    </div>`;
  $('#pvRow').onchange = e => { pv.row = +e.target.value; renderPivot(el); };
  $('#pvCol').onchange = e => { pv.col = e.target.value === '' ? null : +e.target.value; renderPivot(el); };
  $('#pvMeasure').onchange = e => { pv.measure = e.target.value; renderPivot(el); };
  $$('#pvTable td[data-r]').forEach(td => td.onclick = () => { onlyCat(pv.row, td.dataset.r); if (td.dataset.c != null) onlyCat(pv.col, td.dataset.c); refresh(); go('table'); });
  const top = list.slice(0, 20);
  if (pv.col != null && kind !== 'a') {
    bars('pv-chart', top.map(x => [x.rv, x.t]), filterTo(pv.row), { stacks: colVals.slice(0, 10).map((cv, i) => ({ label: shortLabel(cv || '(blank)', 20), data: top.map(x => x.m.has(cv) ? out(x.m.get(cv)) : 0), color: toneColor(cv) || PALETTE[i % PALETTE.length] })) });
  } else bars('pv-chart', top.map(x => [x.rv, x.t]), filterTo(pv.row), { label: measures.find(m => m[0] === mk)[1] });
}

// ── Sources & welcome ────────────────────────────────────────────
function sourcesHtml() {
  return `<div class="grid">
    <div class="card span-6"><div class="card-head"><h3>Upload a file</h3><small>CSV · TSV · Excel</small></div><div class="card-body">
      <div class="dropzone" id="dropzone">${ICON.upload}<h3>Drop a file here or click to browse</h3><p>.csv, .tsv, .xlsx, .xls — processed locally, nothing is uploaded to a server</p></div>
      <div class="row" style="margin-top:12px"><label class="muted">New uploads</label>
        <select class="select" id="uploadMode">
          <option value="merge" ${prefs.uploadMode === 'merge' ? 'selected' : ''}>Merge into saved data (unique by Booking Number)</option>
          <option value="replace" ${prefs.uploadMode === 'replace' ? 'selected' : ''}>Replace saved data</option>
        </select></div>
      <div class="help">${prefs.uploadMode === 'merge'
        ? 'Each upload is added to what is saved. A booking that is already saved is replaced by its rows from the newest file, so nothing is counted twice. Bookings missing from the new file are kept.'
        : 'Each upload replaces everything that is saved.'}</div>
      ${S.workbook && S.workbook.wb.SheetNames.length > 1 ? `<div class="row" style="margin-top:12px"><label class="muted">Sheet</label><select class="select" id="sheetPick">${S.workbook.wb.SheetNames.map(n => `<option ${n === S.ds?.sheet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></div>` : ''}
    </div></div>
    <div class="card span-6"><div class="card-head"><h3>Google Sheets</h3><small>live link</small></div><div class="card-body">
      <div class="row"><input class="input" id="gsUrl" placeholder="https://docs.google.com/spreadsheets/d/…" value="${esc(S.ds?.source?.type === 'gsheet' ? S.ds.source.url : '')}"><button class="btn primary" id="gsLoad">${ICON.sheet}Connect</button></div>
      <div class="row" style="margin-top:10px"><label class="muted">Auto-refresh</label>
        <select class="select" id="gsRefresh">${[[0, 'Off'], [1, 'Every minute'], [5, 'Every 5 min'], [15, 'Every 15 min'], [60, 'Every hour']].map(([v, l]) => `<option value="${v}" ${+prefs.refreshMin === v ? 'selected' : ''}>${l}</option>`).join('')}</select></div>
      <div class="help">The sheet must be readable without signing in:
        <ol><li>In Google Sheets: <b>Share → General access → Anyone with the link (Viewer)</b>, <i>or</i> <b>File → Share → Publish to web → CSV</b>.</li>
        <li>Paste the link. Open the right tab first so the link contains its <code>gid</code>.</li></ol></div>
      ${prefs.sheets.length ? `<div class="section-title" style="margin-top:16px">Saved sheets</div><div class="saved-list">${prefs.sheets.map((s, i) => `<div class="saved">${ICON.sheet}<span class="name" title="${esc(s.url)}">${esc(s.name)}</span><button class="btn sm" data-gs-open="${i}">Load</button><button class="icon-btn" data-gs-del="${i}" title="Remove">${ICON.x}</button></div>`).join('')}</div>` : ''}
    </div></div>
  </div>`;
}

function bindSources(el) {
  const dz = $('#dropzone', el);
  dz.onclick = () => $('#fileInput').click();
  dz.ondragover = e => { e.preventDefault(); dz.classList.add('over'); };
  dz.ondragleave = () => dz.classList.remove('over');
  dz.ondrop = e => { e.preventDefault(); dz.classList.remove('over'); const f = e.dataTransfer.files[0]; f && loadFile(f); };
  const sp = $('#sheetPick', el);
  if (sp) sp.onchange = () => { const { wb, name } = S.workbook; ingestUpload(sheetAoa(wb, sp.value), { name, sheet: sp.value }); };
  $('#uploadMode', el).onchange = e => { prefs.uploadMode = e.target.value; savePrefs(); renderView(); toast(prefs.uploadMode === 'merge' ? 'New uploads will be merged by Booking Number' : 'New uploads will replace the saved data'); };
  const connect = async () => {
    const url = $('#gsUrl', el).value.trim(); if (!url) return;
    const ok = await loadSheet(url);
    if (ok && !prefs.sheets.some(s => s.url === url)) {
      const name = prompt('Name this sheet so you can reload it later:', 'Google Sheet ' + (prefs.sheets.length + 1));
      if (name) { prefs.sheets.push({ name, url }); savePrefs(); S.ds.name = name; persistDataset(S.ds); renderChrome(); }
    }
  };
  $('#gsLoad', el).onclick = connect;
  $('#gsUrl', el).onkeydown = e => e.key === 'Enter' && connect();
  $('#gsRefresh', el).onchange = e => { prefs.refreshMin = +e.target.value; savePrefs(); setupAutoRefresh(); toast(prefs.refreshMin ? `Auto-refresh every ${prefs.refreshMin} min` : 'Auto-refresh off'); };
  $$('[data-gs-open]', el).forEach(b => b.onclick = () => { const s = prefs.sheets[+b.dataset.gsOpen]; loadSheet(s.url, { name: s.name }); });
  $$('[data-gs-del]', el).forEach(b => b.onclick = () => { prefs.sheets.splice(+b.dataset.gsDel, 1); savePrefs(); renderView(); });
}

function renderWelcome(el) {
  el.innerHTML = `<div class="welcome"><h1>Bring in your data</h1><p>Upload a CSV / Excel file or connect a Google Sheet to get a searchable, filterable admin view with dashboards. Booking reports with containers, vessels and ports are recognised automatically.</p>${sourcesHtml()}</div>`;
  bindSources(el);
}

function renderSources(el) {
  const D = S.ds;
  el.innerHTML = `<div class="page-head"><div><h1>Data sources</h1><p>Replace or refresh the dataset. Your last dataset is remembered in this browser.</p></div></div>${sourcesHtml()}
    ${D ? `<div class="card" style="margin-top:16px"><div class="card-head"><h3>Current dataset</h3><button class="btn sm" id="clearData">Remove data</button></div><div class="card-body">
      <div class="kv" style="margin-bottom:14px">
        <div><small>Name</small><span>${esc(D.name)}${D.sheet ? ` · sheet “${esc(D.sheet)}”` : ''}</span></div>
        <div><small>Source</small><span>${D.source.type === 'gsheet' ? `Google Sheet · <a href="${esc(D.source.url)}" target="_blank" rel="noopener">open</a>` : 'Uploaded file'}</span></div>
        <div><small>Rows × columns</small><span>${fmt(D.n)} × ${D.cols.length}</span></div>
        <div><small>Loaded</small><span>${fmtDate(D.loadedAt)}</span></div>
        <div class="full"><small>Detected layout</small><span>${D.profile ? 'Shipping bookings report — Bookings &amp; Voyages views enabled' : 'Generic table'}</span></div>
      </div>
      ${D.imports?.length ? `<div class="section-title">Upload history</div>
      <div style="overflow:auto;margin-bottom:14px"><table class="mini-table"><thead><tr><th>File</th><th>Uploaded</th><th class="n">Bookings in file</th><th class="n">New</th><th class="n">Updated</th><th class="n">Unchanged</th><th class="n">Total after</th></tr></thead><tbody>
        ${D.imports.map(h => `<tr><td>${esc(h.name)}${h.sheet ? ` <span class="faint">· ${esc(h.sheet)}</span>` : ''}</td><td>${fmtDate(h.at)}</td><td class="n">${fmt(h.inFile)}</td><td class="n">${fmt(h.added)}</td><td class="n">${fmt(h.updated)}</td><td class="n">${fmt(h.unchanged)}</td><td class="n"><b>${fmt(h.total)}</b></td></tr>`).join('')}
      </tbody></table></div>` : ''}
      <div class="section-title">Columns</div>
      <div class="col-list">${D.cols.map(c => `<div><span title="${esc(c.label)}">${esc(c.label)}</span><span class="pill ${c.derived ? 'info' : ''}">${c.derived ? 'derived' : c.isCat ? 'category' : c.type}</span></div>`).join('')}</div>
    </div></div>` : ''}`;
  bindSources(el);
  const cd = $('#clearData', el);
  if (cd) cd.onclick = () => { if (!confirm('Remove all saved data (every merged upload) from this browser? This cannot be undone.')) return; idb.del('ic.data'); store.del('ic.data'); S.ds = null; S.workbook = null; setupAutoRefresh(); applyFilters(); renderChrome(); renderView(); };
}

// ════════════════════════════════════════════════════════════════
//  Drawer (details)
// ════════════════════════════════════════════════════════════════
function openDrawer(title, body) {
  $('#drawerTitle').innerHTML = title; $('#drawerBody').innerHTML = body; $('#drawerBody').scrollTop = 0;
  document.body.classList.add('drawer-open');
}
const closeDrawer = () => document.body.classList.remove('drawer-open');
const kvHtml = pairs => `<div class="kv">${pairs.filter(p => p[1] !== '' && p[1] != null).map(([k, v, full]) => `<div class="${full ? 'full' : ''}"><small>${esc(k)}</small><span>${v}</span></div>`).join('')}</div>`;

function allFields(r) {
  return kvHtml(S.ds.cols.map(c => [c.label, c.values[r] === '' ? '<span class="faint">—</span>' : esc(c.values[r]), c.values[r].length > 28]));
}

function openRow(r) {
  if (S.ds.profile) return openBooking(val(S.ds.profile.booking, r), r);
  const D = S.ds, first = D.cols[0];
  openDrawer(`<h2>${esc(first.values[r] || 'Row ' + (r + 1))}</h2><p>Row ${fmt(r + 1)} of ${fmt(D.n)}</p>`, allFields(r));
}

function openBooking(bno, focusRow) {
  const D = S.ds, P = D.profile;
  const all = []; for (let r = 0; r < D.n; r++) if (val(P.booking, r) === bno) all.push(r);
  if (!all.length) return;
  const r0 = focusRow ?? all[0], g = (c, r = r0) => c != null ? val(c, r) : '';
  const etd = P.etd != null ? tsv(P.etd, r0) : NaN;
  const statusPill = pill(g(P.status));

  let focus = '';
  if (focusRow != null) {
    const ms = [
      ['Empty container picked up', P.emptyAt, P.emptyLoc],
      ['Full container received by carrier', P.gateAt, P.gateLoc],
      ['Arrived at port of load', P.polAt, P.polLoc],
    ].filter(m => m[1] != null);
    const rank = P.event != null ? stageRank(val(P.event, focusRow)) : 0;
    const sailed = rank >= 6;
    focus = `<div class="section-title">Container ${esc(g(P.container, focusRow))}</div>
      ${kvHtml([['Size / type', esc(g(P.size, focusRow))], ['Weight', esc(g(P.weight, focusRow)) + ' kg'], ['Volume', esc(g(P.volume, focusRow)) + (g(P.volume, focusRow) && !isBlank(g(P.volume, focusRow)) ? ' CBM' : '')], ['Latest event', esc(g(P.event, focusRow)) + (P.eventAt != null && g(P.eventAt, focusRow) ? ` <span class="faint">· ${esc(g(P.eventAt, focusRow))}</span>` : '')]])}
      <div class="section-title">Milestones</div>
      <div class="timeline">
        ${ms.map(([label, at, loc]) => { const v = val(at, focusRow); return `<div class="tl-item ${v ? 'done' : ''}"><b>${label}</b><small>${v ? `${esc(v)}${loc != null && val(loc, focusRow) ? ' · ' + esc(val(loc, focusRow)) : ''}` : 'Pending'}</small></div>`; }).join('')}
        <div class="tl-item ${sailed ? 'done' : ''}"><b>Vessel departure</b><small>${esc(g(P.etd, focusRow))}${g(P.pol, focusRow) ? ' · ' + esc(g(P.pol, focusRow)) : ''}</small></div>
      </div>`;
  }

  const rowsHtml = all.map(r => `<tr data-r="${r}" ${r === focusRow ? 'style="background:var(--primary-soft)"' : ''}><td><b>${esc(val(P.container, r))}</b></td><td>${esc(g(P.size, r))}</td><td class="num">${esc(g(P.weight, r))}</td><td>${esc(g(P.event, r))}<br><span class="faint">${esc(P.eventAt != null ? val(P.eventAt, r) : '')}</span></td></tr>`).join('');

  openDrawer(
    `<h2>Booking ${esc(bno)}</h2><p>${statusPill} &nbsp;${all.length} container line${all.length > 1 ? 's' : ''} · ${esc(shortLabel(g(P.shipper), 40))}</p>`,
    `${focus}
    <div class="section-title">Shipment</div>
    ${kvHtml([
      ['Vessel / voyage', esc(g(P.voyage))], ['Service', esc(g(P.service))],
      ['Origin city', esc(g(P.origin))], ['Port of load', esc(g(P.pol))],
      ['Port of discharge', esc(g(P.pod))], ['Destination city', esc(g(P.dest))],
      ['ETD', `${fmtDate(etd)} ${pill(g(P.etdType))}`, true],
      ['Shipper', esc(g(P.shipper)), true], ['Forwarder', esc(g(P.forwarder)), true],
    ])}
    <div class="section-title">Containers</div>
    <table class="mini-table"><thead><tr><th>Container</th><th>Type</th><th>Weight (kg)</th><th>Latest event</th></tr></thead><tbody>${rowsHtml}</tbody></table>
    <details style="margin-top:18px"><summary class="section-title" style="cursor:pointer;margin:0">All fields${focusRow != null ? ' (this container)' : ''}</summary><div style="margin-top:10px">${allFields(r0)}</div></details>`
  );
  $$('#drawerBody tr[data-r]').forEach(tr => tr.onclick = () => openBooking(bno, +tr.dataset.r));
}

// ════════════════════════════════════════════════════════════════
//  Export
// ════════════════════════════════════════════════════════════════
function exportRows(kind) {
  const D = S.ds; if (!D) return toast('Nothing to export yet.', true);
  const hidden = hiddenSet(), cols = D.cols.filter(c => !hidden.has(c.label));
  const aoa = [cols.map(c => c.label), ...S.filtered.map(r => cols.map(c => c.values[r]))];
  const base = (D.name || 'export').replace(/\.[^.]+$/, '').replace(/[^\w-]+/g, '_').slice(0, 60) + '_filtered';
  if (kind === 'xlsx') {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Data'); XLSX.writeFile(wb, base + '.xlsx');
  } else {
    const blob = new Blob(['﻿' + Papa.unparse(aoa)], { type: 'text/csv;charset=utf-8' });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: base + '.csv' });
    a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  toast(`Exported ${fmt(S.filtered.length)} rows`);
}

// ════════════════════════════════════════════════════════════════
//  Events
// ════════════════════════════════════════════════════════════════
function applyTheme() {
  const t = prefs.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  document.documentElement.dataset.theme = t;
}

function bind() {
  $('#nav').onclick = e => { const a = e.target.closest('a[data-view]'); if (a) go(a.dataset.view); };
  $('#menuBtn').onclick = () => $('#app').classList.toggle('nav-open');
  $('#scrim').onclick = () => $('#app').classList.remove('nav-open');
  $('#filterBtn').onclick = () => { $('#app').classList.toggle('filters-open'); setTimeout(() => Object.values(S.charts).forEach(c => c.resize()), 220); };
  $('#closeFilters').onclick = () => { $('#app').classList.remove('filters-open'); setTimeout(() => Object.values(S.charts).forEach(c => c.resize()), 220); };
  $('#clearFilters').onclick = () => { S.f.cats.clear(); S.f.ranges.clear(); refresh(); };
  $('#themeBtn').onclick = () => { prefs.theme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'; savePrefs(); applyTheme(); renderView(); };
  $('#refreshBtn').onclick = () => S.ds?.source?.url && loadSheet(S.ds.source.url);
  $('#fileInput').onchange = e => { const f = e.target.files[0]; f && loadFile(f); e.target.value = ''; };
  $('#closeDrawer').onclick = closeDrawer;
  $('#drawerScrim').onclick = closeDrawer;

  const exportMenu = $('#exportMenu');
  $('#exportBtn').onclick = e => { e.stopPropagation(); exportMenu.classList.toggle('open'); };
  exportMenu.onclick = e => { const b = e.target.closest('[data-export]'); if (b) { exportMenu.classList.remove('open'); exportRows(b.dataset.export); } };
  document.addEventListener('click', () => $$('.dropdown-menu.open').forEach(m => m.classList.remove('open')));

  const search = $('#search');
  const doSearch = debounce(() => { S.f.search = search.value.trim(); refresh(); }, 180);
  search.oninput = doSearch;

  // chips
  $('#chips').onclick = e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.chipCat != null) setCat(+b.dataset.chipCat, b.dataset.v, false);
    else if (b.dataset.chipRange != null) S.f.ranges.delete(+b.dataset.chipRange);
    else if (b.hasAttribute('data-chip-search')) { search.value = ''; S.f.search = ''; }
    else if (b.hasAttribute('data-chip-clear')) { S.f.cats.clear(); S.f.ranges.clear(); search.value = ''; S.f.search = ''; }
    refresh();
  };

  // filter panel (delegated)
  const fb = $('#filtersBody');
  fb.addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.cat != null) { setCat(+t.dataset.cat, t.dataset.v, t.checked); refresh(); }
    else if (t.dataset.range != null) {
      const c = +t.dataset.range, col = S.ds.cols[c], rg = { ...(S.f.ranges.get(c) || {}) };
      let x = t.value === '' ? null : col.type === 'date' ? new Date(t.value + 'T00:00').getTime() : +t.value;
      if (x != null && col.type === 'date' && t.dataset.end === 'max') x += 864e5 - 1;
      rg[t.dataset.end] = x;
      if (rg.min == null && rg.max == null) S.f.ranges.delete(c); else S.f.ranges.set(c, rg);
      refresh();
    }
  });
  fb.addEventListener('input', debounce(e => {
    const t = e.target; if (t.dataset.facetQ == null) return;
    S.ui.facetQ[t.dataset.facetQ] = t.value; S.ui.focusFacet = t.dataset.facetQ; renderFilters();
  }, 120));
  fb.addEventListener('toggle', e => {
    const d = e.target; if (!d.dataset?.facet) return;
    d.open ? S.ui.openFacets.add(d.dataset.facet) : S.ui.openFacets.delete(d.dataset.facet);
  }, true);
  fb.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    if (b.dataset.clearCat != null) { S.f.cats.delete(+b.dataset.clearCat); refresh(); }
    if (b.dataset.clearRange != null) { S.f.ranges.delete(+b.dataset.clearRange); refresh(); }
    if (b.dataset.quick != null) {
      const c = +b.dataset.quick, t0 = startOfDay(Date.now()), d = 864e5;
      const q = b.dataset.q, now = new Date();
      const rg = q === 'today' ? { min: t0, max: t0 + d - 1 } : q === 'next7' ? { min: t0, max: t0 + 7 * d - 1 } : q === 'last7' ? { min: t0 - 6 * d, max: t0 + d - 1 }
        : { min: new Date(now.getFullYear(), now.getMonth(), 1).getTime(), max: new Date(now.getFullYear(), now.getMonth() + 1, 1).getTime() - 1 };
      S.f.ranges.set(c, rg); refresh();
    }
  });

  document.addEventListener('click', e => { const g = e.target.closest('[data-go]'); if (g) go(g.dataset.go); });

  // global drag & drop
  document.addEventListener('dragover', e => e.preventDefault());
  document.addEventListener('drop', e => { if (e.target.closest('#dropzone')) return; e.preventDefault(); const f = e.dataTransfer.files[0]; f && loadFile(f); });

  document.addEventListener('keydown', e => {
    if (e.key === '/' && !/input|select|textarea/i.test(document.activeElement.tagName)) { e.preventDefault(); search.focus(); search.select(); }
    if (e.key === 'Escape') { closeDrawer(); if (document.activeElement === search && search.value) { search.value = ''; doSearch(); } }
  });
  window.addEventListener('hashchange', () => { const v = location.hash.slice(1); if (v && v !== S.view) go(v); });
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (!prefs.theme) { applyTheme(); renderView(); } });
}

// ════════════════════════════════════════════════════════════════
//  Boot
// ════════════════════════════════════════════════════════════════
async function boot() {
  applyTheme();
  bind();
  const v = location.hash.slice(1);
  if (['overview', 'table', 'bookings', 'voyages', 'pivot', 'sources'].includes(v)) S.view = v;
  let saved = await idb.get('ic.data');
  if (!saved) {                                   // one-time move from the old localStorage copy
    saved = store.get('ic.data');
    if (saved?.aoa && await idb.set('ic.data', saved)) store.del('ic.data');
  }
  if (saved?.aoa) {
    try {
      setDataset(buildDataset(saved.aoa, { name: saved.name, source: saved.source || { type: 'file' }, loadedAt: saved.loadedAt, imports: saved.imports || [] }), { persist: false });
      if (saved.source?.type === 'gsheet') loadSheet(saved.source.url, { silent: true, name: saved.name });   // pull fresh data
    } catch (e) { console.error(e); }            // keep the saved copy; never delete data on a read error
  }
  applyFilters(); renderChrome(); go(S.view);
}

// Expose a tiny API for debugging in the console
window.InventoryConsole = { state: S, loadSheet };
boot();
})();
