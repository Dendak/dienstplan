/* Dienstplan Assistenz – liest die Excel-Dienstpläne direkt im Browser. */
'use strict';

// ---------------------------------------------------------------------------
// Konfiguration
// ---------------------------------------------------------------------------

// Kürzel → Name (Assistenzärzt:innen, laut ooeg.at)
const ASSISTENZ = {
  Al:   'Dr.in Kholod Alshebli',
  Ei:   'Dr.in Anna-Maria Eicher',
  Khri: 'Dr.in Ghizlane Khribach',
  Ma:   'Dr.in Mariana Machynska',
  Mas:  'Dr.in Karolina Masarova',
  Mo:   'Dr. Majed Mohammad',
  San:  'Dr.in Sofia Santesteban',
  Ah:   'Aharanwa',
  El:   'El Shiri',
};

// Letzter Arbeitstag – danach werden Einträge ignoriert (z. B. Reste in der VB-Formel)
const AUSTRITT = {
  Ah: '2026-07-31',
  El: '2026-06-30',
};

// Tippvarianten aus den Excel-Dateien → Kürzel
const ALIAS = { kh: 'Khri', sa: 'San', alshebly: 'Al', wa: 'Was', bra: 'Br', put: 'Pu' };

// Primar/Oberärzt:innen – werden als "OA im Dienst" angezeigt (Spalte ND1)
const OA_NAMEN = {
  Br:  'Prim. Dr. Helge Brandmeier',
  Was: 'OA Dr. Wolfram Wasserfaller',
  Co:  'OA Dr. Ludovit Cobirka',
  Gi:  'OÄ Dr.in Elke Gierlinger-Plöderl',
  Ki:  'OA Dr. Dovydas Kindurys',
  Pu:  'OA Dr. Christian Puttinger',
  Ra:  'OÄ Dr.in Jovana Radojevic',
  Sd:  'OÄ Dr.in Anita Schild',
  Pf:  'Pfleger',
  Bl:  'Bleier',
  Ba:  'Baschata',
  Sch: 'Schweiger',
  Kr:  'Krupitz',
  Le:  'Lechner',
};
// weitere bekannte Kürzel (kein Tippfehler): Ko, Grü, Kl/LK = Turnusärzt:innen
const OA_KUERZEL = [...Object.keys(OA_NAMEN), 'Ko', 'Grü', 'Kl', 'LK'];

// Spaltenüberschrift (Zeile 1 im Excel) → Status
const SPALTEN = {
  'nd1': 'oadienst', 'nd2': 'dienst', 'nd3': 'dienst3',
  'vb': 'vb', 'bis16.00': 'bis16', 'bz': 'brz',
  'fb': 'fb',
  'urlaub': 'urlaub',
  'krank': 'abw',
  'wrt': 'wrt',
  'e,frei': 'efrei',
  'ausdienst': 'frei',
  'vb-ass': 'vb',
  'gm': 'gm',
  'brz': 'brz',
  '16:00': 'bis16',
  'besonderheiten': 'note',
};

// Reihenfolge = Priorität für die Anzeige
const STATUS = {
  dienst: { label: 'Dienst',            short: 'D',  cls: 's-dienst', ics: 'Dienst' },
  abw:    { label: 'Abwesend',          short: '–',  cls: 's-abw' },
  urlaub: { label: 'Urlaub',            short: 'U',  cls: 's-urlaub', ics: 'Urlaub' },
  fb:     { label: 'Fortbildung',       cal: 'FB', short: 'FB', cls: 's-fb',     ics: 'Fortbildung' },
  wrt:    { label: 'Wochenruhetag',     cal: 'WRT', short: 'W',  cls: 's-wrt',    ics: 'Wochenruhetag' },
  efrei:  { label: 'Ersatzfrei',        short: 'E',  cls: 's-efrei',  ics: 'Ersatzfrei' },
  frei:   { label: 'Frei nach Dienst',  cal: 'frei', short: 'f',  cls: 's-frei',   ics: 'Frei nach Dienst' },
  gm:     { label: 'Gmunden',           short: 'GM', cls: 's-gm',     ics: 'Gmunden' },
  brz:    { label: 'BRZ',               short: 'BZ', cls: 's-gm' },
  bis16:  { label: 'bis 16:00',         short: '16', cls: 's-fb' },
  vb:     { label: 'Im Haus',           cal: 'Haus', short: '·',  cls: 's-vb' },
};
const PRIO = Object.keys(STATUS);

const MONATE = ['Jänner', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
const WT = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

// ---------------------------------------------------------------------------
// Hilfsfunktionen
// ---------------------------------------------------------------------------

const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = n => String(n).padStart(2, '0');
const iso = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
const parseIso = s => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const addDays = (s, n) => { const d = parseIso(s); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const dow = s => parseIso(s).getUTCDay();
const todayIso = () => { const d = new Date(); return iso(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
const fmtDay = s => { const d = parseIso(s); return `${WT[d.getUTCDay()]}, ${d.getUTCDate()}.${d.getUTCMonth() + 1}.`; };
const fmtLong = s => { const d = parseIso(s); return `${['Sonntag','Montag','Dienstag','Mittwoch','Donnerstag','Freitag','Samstag'][d.getUTCDay()]}, ${d.getUTCDate()}. ${MONATE[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
const monthLabel = key => { const [y, m] = key.split('-').map(Number); return `${MONATE[m - 1]} ${y}`; };
const personName = a => ASSISTENZ[a] || OA_NAMEN[a] || a;
const initials = a => a;

function serialToIso(n) {
  const d = new Date(Math.round((n - 25569) * 86400000));
  return d.toISOString().slice(0, 10);
}
function headerText(v) {
  if (typeof v === 'number' && v > 0 && v < 1) { // Uhrzeit, z. B. 16:00
    const min = Math.round(v * 1440);
    return `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
  }
  return v == null ? '' : String(v);
}

// Österreichische Feiertage
const holidayCache = {};
function holidays(y) {
  if (holidayCache[y]) return holidayCache[y];
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const em = Math.floor((h + l - 7 * m + 114) / 31), ed = ((h + l - 7 * m + 114) % 31) + 1;
  const easter = iso(y, em, ed);
  const map = {
    [iso(y, 1, 1)]: 'Neujahr', [iso(y, 1, 6)]: 'Heilige Drei Könige',
    [addDays(easter, 1)]: 'Ostermontag', [iso(y, 5, 1)]: 'Staatsfeiertag',
    [addDays(easter, 39)]: 'Christi Himmelfahrt', [addDays(easter, 50)]: 'Pfingstmontag',
    [addDays(easter, 60)]: 'Fronleichnam', [iso(y, 8, 15)]: 'Mariä Himmelfahrt',
    [iso(y, 10, 26)]: 'Nationalfeiertag', [iso(y, 11, 1)]: 'Allerheiligen',
    [iso(y, 12, 8)]: 'Mariä Empfängnis', [iso(y, 12, 25)]: 'Christtag', [iso(y, 12, 26)]: 'Stefanitag',
  };
  return (holidayCache[y] = map);
}
const holidayOf = s => holidays(+s.slice(0, 4))[s];
const isWeekendOrHoliday = s => { const w = dow(s); return w === 0 || w === 6 || !!holidayOf(s); };

// ---------------------------------------------------------------------------
// Excel einlesen
// ---------------------------------------------------------------------------

/** "Pu,Wa", "Pu Mo", "PuWa", "Br(K)", "San?" … → einzelne Kürzel */
function splitNames(v) {
  if (v == null) return [];
  const raw = String(v);
  return raw.replace(/\([^)]*\)/g, ' ')
    .replace(/([a-zäöüß])([A-ZÄÖÜ])/g, '$1,$2')
    .split(/[,;\/+\n\s]+/)
    .map(t => t.replace(/[.?!:]+$/, ''))
    .filter(t => /^[A-Za-zÄÖÜäöüß]{1,5}$/.test(t) || ALIAS[t.toLowerCase()])
    .map(t => ({ raw, name: ALIAS[t.toLowerCase()] || t }));
}

/** Liest ein Tabellenblatt; liefert null, wenn es kein Monatsblatt ist. */
function parseSheet(ws, sheetName, source) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  // Kopfzeile = erste Zeile mit „ND2“ (neue Pläne: Zeile 1, alte: Zeile 3)
  const norm = v => headerText(v).toLowerCase().replace(/\s+/g, '');
  const h = rows.slice(0, 6).findIndex(r => r && r.some(v => norm(v) === 'nd2'));
  if (h < 0) return null;
  const baseRow = rows.slice(0, h + 1).find(r => typeof r[0] === 'number' && r[0] > 30000);
  if (!baseRow) return null;
  const base = serialToIso(baseRow[0]);

  const cols = rows[h].map(v => SPALTEN[norm(v)] || null);
  const days = [];
  const tokens = []; // für die Prüfung
  let r = h + 1, prevDay = 0;
  for (; r < rows.length; r++) {
    const row = rows[r];
    if (!row || typeof row[0] !== 'number') break;
    let date;
    if (row[0] > 30000) date = serialToIso(row[0]);
    else if (row[0] >= 1 && row[0] <= 31 && row[0] > prevDay) date = base.slice(0, 8) + pad(row[0]); // altes Format: nur Tageszahl
    else break;
    prevDay = +date.slice(8);
    if (date.slice(0, 7) !== base.slice(0, 7)) continue;
    const day = { date, entries: {}, note: '' };
    cols.forEach((st, c) => {
      if (!st || row[c] == null || row[c] === '') return;
      if (st === 'note') { day.note = String(row[c]).trim(); return; }
      for (const t of splitNames(row[c])) {
        tokens.push({ date, status: st, raw: t.raw, name: t.name });
        (day.entries[t.name] ||= new Set()).add(st);
      }
    });
    days.push(day);
  }
  if (!days.length) return null;

  // Assistenz-Liste der Planerin (Spalte E unter der Tabelle, mit Zählung in F)
  const listed = [];
  for (; r < rows.length; r++) {
    const v = rows[r] && rows[r][4];
    if (typeof v === 'string' && v.trim() && v.trim().length <= 6) listed.push(v.trim());
  }

  const first = days[0].date;
  return { key: first.slice(0, 7), days, tokens, listed, source: `${source} · Blatt „${sheetName}“` };
}

function parseWorkbook(buf, source) {
  const wb = XLSX.read(buf, { type: 'array' });
  return wb.SheetNames.map(n => parseSheet(wb.Sheets[n], n, source)).filter(Boolean);
}

/** Baut aus Monatsblättern den Datensatz. */
function buildData(sheets) {
  const months = new Map();
  for (const s of sheets) months.set(s.key, s); // spätere überschreiben frühere

  // Wer gehört zur Assistenz?
  const known = new Set(Object.keys(ASSISTENZ));
  const oa = new Set(OA_KUERZEL.map(x => x.toLowerCase()));
  const found = new Set();
  const nd2 = new Set();
  for (const m of months.values()) {
    m.listed.forEach(a => found.add(a));
    m.tokens.forEach(t => { if (t.status === 'vb') found.add(t.name); if (t.status === 'dienst') nd2.add(t.name); });
  }
  const canon = new Map();
  const nd2l = new Set([...nd2].map(x => x.toLowerCase()));
  [...known, ...found, ...nd2].forEach(a => {
    const k = a.toLowerCase();
    if (!oa.has(k) || nd2l.has(k)) canon.set(k, canon.get(k) || [...known].find(x => x.toLowerCase() === k) || a);
  });

  // Einträge normalisieren (Groß/Klein, Leerzeichen) und auf Assistenz beschränken
  for (const m of months.values()) {
    for (const d of m.days) {
      const norm = {};
      d.oa = [];
      for (const [n, set] of Object.entries(d.entries)) {
        if (set.has('oadienst')) d.oa.push(n);
        const a = canon.get(n.toLowerCase());
        if (!a || (AUSTRITT[a] && d.date > AUSTRITT[a])) continue;
        norm[a] ||= new Set();
        set.forEach(s => norm[a].add(s === 'dienst3' ? 'dienst' : s));
      }
      d.entries = norm;
    }
    // Monate ohne einen einzigen Dienst gelten als leer (z. B. leere Jahresvorlage)
    m.hasDienst = m.days.some(d => Object.values(d.entries).some(s => s.has('dienst')));
    m.hasData = m.hasDienst || m.days.some(d => Object.values(d.entries).some(s => [...s].some(x => x !== 'vb')));
  }
  for (const [k, m] of months) if (!m.hasData) months.delete(k);

  // Personen sortieren: bekannte zuerst in Konfig-Reihenfolge, dann weitere
  const active = new Set();
  // (bekannte Assistenz, sobald sie vorkommt; andere Kürzel nur, wenn sie Dienste haben)
  for (const m of months.values()) for (const d of m.days) for (const [a, set] of Object.entries(d.entries)) {
    if (known.has(a) || (set.has('dienst') && nd2.has(a))) active.add(a);
  }
  const people = [...canon.values()].filter(a => active.has(a))
    .sort((a, b) => (known.has(b) - known.has(a)) || a.localeCompare(b, 'de'));

  return { months, people, canon, oa };
}

function statusesOf(day, a) {
  const set = day && day.entries[a];
  if (!set) return [];
  return PRIO.filter(p => set.has(p));
}
const primary = (day, a) => statusesOf(day, a)[0] || null;

// ---------------------------------------------------------------------------
// Laden
// ---------------------------------------------------------------------------

const state = {
  data: null, published: null, preview: null,
  view: safeGet('dp.view') || 'me',
  person: safeGet('dp.person') || '',
  month: null, day: todayIso(), sourceInfo: '',
  enc: null, password: null, files: [], pending: null, locked: false, lockError: '',
  planMode: safeGet('dp.planMode') || 'list', onlyMine: safeGet('dp.onlyMine') === '1', choosing: false, scrollToday: false, year: null, yearMetric: safeGet('dp.yearMetric') || 'urlaub',
};
if (!['me', 'plan', 'day', 'year'].includes(state.view)) state.view = 'me';
function safeSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* egal */ } }
function safeGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
function safeDel(k) { try { localStorage.removeItem(k); } catch (e) { /* egal */ } }

// ---- Verschlüsselung (AES-256-GCM, Schlüssel per PBKDF2-SHA256 aus dem Passwort) ----
const ENC_FILE = 'plaene/plan.enc';
const GITHUB_UPLOAD = 'https://github.com/Dendak/dienstplan/upload/main/plaene';
const b64 = buf => { let s = ''; const b = new Uint8Array(buf); for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));

async function deriveKey(pw, salt, iter) {
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
async function decryptBundle(enc, pw) {
  const key = await deriveKey(pw, unb64(enc.salt), enc.iter);
  const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(enc.iv) }, key, unb64(enc.ct));
  return JSON.parse(new TextDecoder().decode(plain)).files; // [{name, data(b64)}]
}
async function encryptBundle(files, pw) {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12)), iter = 310000;
  const key = await deriveKey(pw, salt, iter);
  const plain = new TextEncoder().encode(JSON.stringify({ created: new Date().toISOString(), files }));
  const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, plain);
  return { v: 1, kdf: 'PBKDF2-SHA256', iter, salt: b64(salt), iv: b64(iv), ct: b64(ct) };
}

function dataFromFiles(files) {
  const sheets = files.flatMap(f => parseWorkbook(unb64(f.data), f.name));
  state.sourceInfo = files.length ? `Quelle: ${files.map(f => f.name).join(', ')}` : '';
  return buildData(sheets);
}

async function loadPublished() {
  const res = await fetch(ENC_FILE, { cache: 'no-cache' });
  if (!res.ok) return buildData([]);
  state.enc = await res.json();
  const saved = safeGet('dp.pw');
  if (saved && await unlock(saved, true)) return state.published;
  state.locked = true;
  return null;
}

async function unlock(pw, silent) {
  try {
    state.files = await decryptBundle(state.enc, pw);
  } catch (e) {
    if (!silent) state.lockError = 'Passwort falsch.';
    return false;
  }
  state.password = pw;
  state.locked = false;
  state.lockError = '';
  state.published = state.data = dataFromFiles(state.files);
  state.month = pickMonth(state.data);
  return true;
}

function viewLock() {
  return `<section class="card lock">
    <div class="lock-icon" aria-hidden="true"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg></div>
    <h2>Geschützter Dienstplan</h2>
    <p class="muted">Bitte das Team-Passwort eingeben.</p>
    <form id="lockForm">
      <input type="password" id="pwInput" autocomplete="current-password" placeholder="Passwort" required>
      <label class="small muted"><input type="checkbox" id="pwRemember" checked> Auf diesem Gerät merken</label>
      <button class="btn" type="submit">Öffnen</button>
      ${state.lockError ? `<p class="lock-err">${esc(state.lockError)}</p>` : ''}
    </form>
  </section>`;
}

function pickMonth(data) {
  const keys = [...data.months.keys()].sort();
  if (!keys.length) return null;
  const cur = todayIso().slice(0, 7);
  if (state.month && data.months.has(state.month)) return state.month;
  return keys.find(k => k >= cur) || keys[keys.length - 1];
}

// ---------------------------------------------------------------------------
// Ansichten
// ---------------------------------------------------------------------------

function render() {
  const data = state.data;
  const app = $('#app');
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.view === state.view));
  $('#sourceInfo').textContent = state.preview ? '' : state.sourceInfo;
  $('.foot').hidden = state.locked;
  $('.tabs').hidden = $('#monthNav').hidden = state.locked;
  $('#publishBtn').hidden = !state.pending;
  $('#logoutBtn').hidden = !state.password;

  if (state.locked) {
    app.innerHTML = viewLock();
    $('#pwInput').focus();
    return;
  }
  if (!data || !data.months.size) {
    app.innerHTML = `<div class="empty">Noch kein Dienstplan veröffentlicht.</div>`;
    $('#monthLabel').textContent = '–';
    renderChecks();
    return;
  }
  const keys = [...data.months.keys()].sort();
  const idx = keys.indexOf(state.month);
  $('#monthLabel').textContent = monthLabel(state.month);
  $('#prevMonth').disabled = idx <= 0;
  $('#nextMonth').disabled = idx >= keys.length - 1;

  const month = data.months.get(state.month);
  $('#monthNav').hidden = state.view === 'day' || state.view === 'year';
  if (state.view === 'plan') app.innerHTML = viewPlan(data, month);
  else if (state.view === 'year') app.innerHTML = viewYear(data);
  else if (state.view === 'day') app.innerHTML = viewDay(data);
  else app.innerHTML = viewMe(data, month);
  renderChecks();
  if (state.scrollToday) {
    state.scrollToday = false;
    const t = $('#app .is-today');
    if (t) t.scrollIntoView({ block: 'center' });
  }
}

// ---- Namen ----
const plainName = a => personName(a).replace(/^(Prim\.|OA|OÄ) /, '').replace(/^Dr\.(in)? /, '');
const surname = a => /^(Dr|OA|OÄ|Prim)/.test(personName(a)) ? plainName(a).split(' ').slice(-1)[0] : plainName(a);
const oaName = k => OA_NAMEN[k] || k;
const oaShort = k => OA_NAMEN[k] ? OA_NAMEN[k].replace(/^(Prim\.|OA|OÄ) (Dr\.(in)? )?/, '$1 ').replace(/ \S+ (\S+)$/, ' $1') : k;
const avatar = a => `<span class="av" aria-hidden="true">${esc(a)}</span>`;

function dienstOf(day, data) { return day ? data.people.filter(a => statusesOf(day, a).includes('dienst')) : []; }
function oaOf(day) { return (day && day.oa) || []; }

function allDays(data) {
  return [...data.months.values()].sort((x, y) => x.key.localeCompare(y.key)).flatMap(m => m.days);
}
function relDays(date) {
  const diff = Math.round((parseIso(date) - parseIso(todayIso())) / 86400000);
  if (diff === 0) return 'heute';
  if (diff === 1) return 'morgen';
  if (diff === -1) return 'gestern';
  if (diff > 1) return `in ${diff} Tagen`;
  return `vor ${-diff} Tagen`;
}
const dayTags = date => {
  const hol = holidayOf(date);
  if (hol) return `<span class="chip chip-hol">${esc(hol)}</span>`;
  const w = dow(date);
  return (w === 0 || w === 6) ? `<span class="chip chip-we">Wochenende</span>` : '';
};
const oaLine = (day, prefix = 'mit ') => {
  const oa = oaOf(day);
  return oa.length ? `${prefix}${oa.map(k => `<strong title="${esc(oaName(k))}">${esc(oaShort(k))}</strong>`).join(' / ')}` : '<span class="muted">OA: noch offen</span>';
};

function peoplePicker(data, big) {
  return `<div class="people${big ? ' people-big' : ''}">${data.people.map(a => `
    <button class="person${a === state.person ? ' active' : ''}" data-person="${esc(a)}">
      ${avatar(a)}<span>${esc(plainName(a))}</span>
    </button>`).join('')}</div>`;
}

function legend(keys) {
  return `<div class="legend">${keys.map(k => `<span><i class="${STATUS[k].cls}"${k === 'vb' ? ' style="border:1px solid var(--line)"' : ''}></i>${STATUS[k].label}</span>`).join('')}</div>`;
}

// ---------------------------------------------------------------------------
// Meine Dienste
// ---------------------------------------------------------------------------

function viewMe(data, month) {
  const a = state.person;
  if (!a || !data.people.includes(a) || state.choosing) {
    return `<section class="card onboarding">
      <h2>Wer bist du?</h2>
      <p class="muted">Tippe auf deinen Namen – danach siehst du sofort deine Dienste. Die Auswahl wird nur auf diesem Gerät gespeichert.</p>
      ${peoplePicker(data, true)}
    </section>`;
  }
  const today = todayIso();
  const days = allDays(data);
  const next = days.find(d => d.date >= today && statusesOf(d, a).includes('dienst'));
  const myMonth = month.days.filter(d => statusesOf(d, a).includes('dienst'));
  const we = myMonth.filter(d => isWeekendOrHoliday(d.date)).length;
  const count = st => month.days.filter(d => statusesOf(d, a).includes(st)).length;

  // Hero: nächster Dienst
  let hero;
  if (next) {
    const rel = relDays(next.date);
    hero = `<section class="hero">
      <div class="hero-top"><span class="hero-label">${rel === 'heute' ? 'Heute hast du Dienst' : 'Dein nächster Dienst'}</span>
        <button class="hero-who" data-action="choose" title="Person wechseln">${avatar(a)}${esc(plainName(a).split(' ')[0])} <span aria-hidden="true">▾</span></button></div>
      <div class="hero-date">${fmtLong(next.date).replace(/ \d{4}$/, '')}</div>
      <div class="hero-meta"><span class="hero-rel">${rel}</span>${dayTags(next.date)}</div>
      <div class="hero-oa">${oaLine(next, 'mit ')}</div>
    </section>`;
  } else {
    hero = `<section class="hero">
      <div class="hero-top"><span class="hero-label">Keine weiteren Dienste</span>
        <button class="hero-who" data-action="choose">${avatar(a)}${esc(plainName(a).split(' ')[0])} <span aria-hidden="true">▾</span></button></div>
      <div class="hero-date small-date">Im veröffentlichten Plan ist kein weiterer Dienst eingetragen.</div>
    </section>`;
  }

  // Liste der Dienste im Monat
  const list = myMonth.length ? `<ul class="dlist">${myMonth.map(d => {
    const past = d.date < today;
    return `<li class="${past ? 'past' : ''}${d.date === today ? ' is-today' : ''}" data-goto="${d.date}">
      <div class="dl-date${isWeekendOrHoliday(d.date) ? ' we' : ''}"><b>${+d.date.slice(8)}</b><span>${WT[dow(d.date)]}</span></div>
      <div class="dl-main">
        <div class="dl-title">Dienst ${d.date >= today ? `<span class="muted">· ${relDays(d.date)}</span>` : ''}</div>
        <div class="dl-sub">${oaLine(d)}</div>
      </div>
      <div class="dl-side">${dayTags(d.date)}</div>
    </li>`;
  }).join('')}</ul>` : '<p class="muted">In diesem Monat hast du keinen Dienst.</p>';

  // Weitere Einträge (Urlaub, FB, WRT, Ersatzfrei, …)
  const other = [];
  for (const d of month.days) for (const s of statusesOf(d, a)) if (!['vb', 'dienst', 'frei'].includes(s)) {
    const last = other[other.length - 1];
    if (last && last.s === s && addDays(last.to, 1) === d.date) last.to = d.date; else other.push({ s, from: d.date, to: d.date });
  }

  // Kalender (Montag zuerst)
  const lead = (dow(month.days[0].date) + 6) % 7;
  let cal = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'].map(w => `<div class="dow">${w}</div>`).join('');
  cal += '<div class="cell blank"></div>'.repeat(lead);
  for (const d of month.days) {
    const sts = statusesOf(d, a);
    const p = sts[0];
    const hol = holidayOf(d.date);
    const cls = ['cell', isWeekendOrHoliday(d.date) ? 'we' : '', hol ? 'hol' : '', d.date === today ? 'today' : '', p && p !== 'vb' ? 'has ' + STATUS[p].cls : ''].join(' ');
    const sub = p === 'dienst' ? (oaOf(d).map(esc).join('/') || '') : p && p !== 'vb' ? (STATUS[p].cal || STATUS[p].label) : '';
    cal += `<button class="${cls}" data-goto="${d.date}" title="${esc([fmtDay(d.date), hol, ...sts.filter(s => s !== 'vb').map(s => STATUS[s].label)].filter(Boolean).join(' · '))}">
      <span class="d">${+d.date.slice(8)}</span>${p === 'dienst' ? '<span class="c-main">Dienst</span>' : ''}<span class="c-sub">${sub}</span></button>`;
  }

  return `
    ${hero}
    <section class="card">
      <div class="card-head">
        <h2>Meine Dienste <span class="muted">· ${monthLabel(state.month)}</span></h2>
        <span class="summary">${myMonth.length} Dienste${we ? ` · ${we} WE/Feiertag` : ''}</span>
      </div>
      ${list}
    </section>
    ${other.length ? `<section class="card">
      <h2>Weitere Einträge</h2>
      <ul class="olist">${other.map(o => `<li><span class="tag ${STATUS[o.s].cls}">${STATUS[o.s].label}</span>
        <span>${o.from === o.to ? fmtDay(o.from) : `${fmtDay(o.from)} – ${fmtDay(o.to)}`}</span></li>`).join('')}</ul>
    </section>` : ''}
    <section class="card">
      <h2>Kalender</h2>
      <div class="cal">${cal}</div>
      ${legend(['dienst', 'frei', 'urlaub', 'fb', 'wrt', 'efrei', 'abw'])}
      <p class="muted small">Tippe auf einen Tag, um zu sehen, wer an dem Tag Dienst hat.</p>
    </section>
    <section class="card">
      <h2>In den eigenen Kalender</h2>
      <p class="muted small">Lädt alle veröffentlichten Dienste als .ics-Datei – öffnen mit Outlook, Google- oder iPhone-Kalender.</p>
      <div class="row-actions">
        <button class="btn" id="icsBtn">Dienste exportieren (.ics)</button>
        <button class="btn secondary" onclick="window.print()">Drucken</button>
      </div>
    </section>`;
}

// ---------------------------------------------------------------------------
// Monatsplan
// ---------------------------------------------------------------------------

function viewPlan(data, month) {
  const me = data.people.includes(state.person) ? state.person : null;
  const mode = state.planMode;
  const onlyMine = state.onlyMine && me;
  const today = todayIso();

  const controls = `<div class="plan-controls">
    <div class="seg" role="group" aria-label="Darstellung">
      <button class="${mode === 'list' ? 'on' : ''}" data-mode="list">Liste</button>
      <button class="${mode === 'table' ? 'on' : ''}" data-mode="table">Tabelle</button>
    </div>
    <div class="pc-right">
      ${me && mode === 'list' ? `<label class="switch"><input type="checkbox" id="onlyMine" ${onlyMine ? 'checked' : ''}> nur meine</label>` : ''}
      <button class="btn secondary btn-sm${state.exportOpen ? ' on' : ''}" data-action="export-toggle" aria-expanded="${!!state.exportOpen}">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>Export</button>
    </div>
  </div>${exportPanel(data)}`;

  if (mode === 'table') return `<section class="card">${controls}${planTable(data, month, me)}</section>`;

  const rows = month.days.map(d => {
    const dienst = dienstOf(d, data);
    const mine = me && dienst.includes(me);
    if (onlyMine && !mine) return '';
    const hol = holidayOf(d.date);
    const extra = [];
    for (const [st, lbl] of [['frei', 'frei n. D.'], ['urlaub', 'Urlaub'], ['efrei', 'Ersatzfrei'], ['wrt', 'WRT'], ['fb', 'FB'], ['gm', 'Gmunden'], ['abw', 'abwesend']]) {
      const who = data.people.filter(a => statusesOf(d, a).includes(st) && !dienst.includes(a));
      if (who.length) extra.push(`<span class="x"><span class="x-l">${lbl}:</span> ${who.map(a => a === me ? '<b>du</b>' : esc(surname(a))).join(', ')}</span>`);
    }
    return `<li class="${mine ? 'mine' : ''}${d.date === today ? ' is-today' : ''}${d.date < today ? ' past' : ''}" data-goto="${d.date}">
      <div class="dl-date${isWeekendOrHoliday(d.date) ? ' we' : ''}${hol ? ' hol' : ''}"><b>${+d.date.slice(8)}</b><span>${WT[dow(d.date)]}</span></div>
      <div class="dl-main">
        <div class="dl-title">${dienst.length ? dienst.map(a => a === me ? `<span class="me-badge">Du</span>` : esc(plainName(a))).join(', ') : '<span class="muted">kein Dienst eingetragen</span>'}</div>
        <div class="dl-sub">${oaOf(d).length ? `mit ${oaOf(d).map(k => `<strong title="${esc(oaName(k))}">${esc(oaShort(k))}</strong>`).join(' / ')}` : ''}${hol ? ` <span class="hol-name">${esc(hol)}</span>` : ''}</div>
        ${extra.length ? `<div class="dl-extra">${extra.join('')}</div>` : ''}
      </div>
    </li>`;
  }).join('');

  return `<section class="card">
    <div class="card-head"><h2>Monatsplan <span class="muted">· ${monthLabel(state.month)}</span></h2></div>
    ${controls}
    <ul class="dlist plan">${rows || '<li class="muted">Keine Dienste.</li>'}</ul>
  </section>`;
}

// ---------------------------------------------------------------------------
// Jahresübersicht (Urlaub / Dienste pro Person und Monat)
// ---------------------------------------------------------------------------

const isWorkday = s => { const w = dow(s); return w !== 0 && w !== 6 && !holidayOf(s); };
const YEAR_METRICS = {
  urlaub:   { label: 'Urlaub', unit: 'Arbeitstage', hint: 'Urlaubstage Mo–Fr ohne Feiertage', test: (d, a) => statusesOf(d, a).includes('urlaub') && isWorkday(d.date) },
  urlaubKT: { label: 'Urlaub (Kalendertage)', unit: 'Kalendertage', hint: 'alle eingetragenen Urlaubstage inkl. Wochenende', test: (d, a) => statusesOf(d, a).includes('urlaub') },
  dienst:   { label: 'Dienste', unit: 'Dienste', hint: 'Nachtdienste (ND2 + ND3)', test: (d, a) => statusesOf(d, a).includes('dienst') },
  weDienst: { label: 'WE/Feiertag-Dienste', unit: 'Dienste', hint: 'Dienste an Sa, So und Feiertagen', test: (d, a) => statusesOf(d, a).includes('dienst') && isWeekendOrHoliday(d.date) },
  efrei:    { label: 'Ersatzfrei + WRT', unit: 'Tage', hint: 'Ersatzfrei und Wochenruhetage', test: (d, a) => statusesOf(d, a).some(s => s === 'efrei' || s === 'wrt') },
};

function yearsOf(data) { return [...new Set([...data.months.keys()].map(k => k.slice(0, 4)))].sort(); }

/** Werte je Person × Monat für ein Jahr; null = Monat nicht im Plan. */
function yearTable(data, year, metric) {
  const test = YEAR_METRICS[metric].test;
  const keys = Array.from({ length: 12 }, (_, i) => `${year}-${pad(i + 1)}`);
  const rows = data.people.map(a => {
    const vals = keys.map(k => { const m = data.months.get(k); return m ? m.days.filter(d => test(d, a)).length : null; });
    return { a, vals, sum: vals.reduce((x, v) => x + (v || 0), 0) };
  }).filter(r => r.vals.some(v => v));
  return { keys, rows, missing: keys.filter(k => !data.months.has(k)) };
}

/** Datumsbereiche für Tooltip, z. B. „3.–7.8., 24.8.“ */
function ranges(days) {
  const out = [];
  for (const d of days) {
    const last = out[out.length - 1];
    if (last && addDays(last[1], 1) === d) last[1] = d; else out.push([d, d]);
  }
  const f = s => `${+s.slice(8)}.${+s.slice(5, 7)}.`;
  return out.map(([x, y]) => x === y ? f(x) : x.slice(0, 7) === y.slice(0, 7) ? `${+x.slice(8)}.–${f(y)}` : `${f(x)}–${f(y)}`).join(', ');
}

function viewYear(data) {
  const years = yearsOf(data);
  if (!years.includes(state.year)) state.year = years.includes(state.month.slice(0, 4)) ? state.month.slice(0, 4) : years[years.length - 1];
  const metric = YEAR_METRICS[state.yearMetric] ? state.yearMetric : 'urlaub';
  const { keys, rows, missing } = yearTable(data, state.year, metric);
  const max = Math.max(1, ...rows.flatMap(r => r.vals.map(v => v || 0)));
  const me = state.person;
  const test = YEAR_METRICS[metric].test;

  const head = keys.map(k => `<th class="${data.months.has(k) ? '' : 'miss'}" title="${data.months.has(k) ? '' : 'Kein Plan geladen'}">${MONATE[+k.slice(5) - 1].slice(0, 3)}</th>`).join('');
  const body = rows.map(r => `<tr class="${r.a === me ? 'me' : ''}">
      <th class="name" title="${esc(personName(r.a))}">${esc(plainName(r.a))}</th>
      ${r.vals.map((v, i) => {
        if (v === null) return '<td class="miss">–</td>';
        if (!v) return '<td class="zero">·</td>';
        const days = data.months.get(keys[i]).days.filter(d => test(d, r.a)).map(d => d.date);
        return `<td class="val" data-month="${keys[i]}" style="--i:${(0.18 + 0.82 * v / max).toFixed(2)};color:${v / max > 0.45 ? '#fff' : 'var(--text)'}" title="${esc(ranges(days))}">${v}</td>`;
      }).join('')}
      <td class="sum">${r.sum}</td></tr>`).join('');
  const colSum = keys.map((k, i) => data.months.has(k) ? rows.reduce((x, r) => x + (r.vals[i] || 0), 0) : '');

  return `<section class="card">
    <div class="card-head">
      <h2>Jahresübersicht ${years.length > 1 ? `<select id="yearSel" aria-label="Jahr">${years.map(y => `<option ${y === state.year ? 'selected' : ''}>${y}</option>`).join('')}</select>` : state.year}</h2>
      <button class="btn secondary btn-sm" data-action="year-xlsx">
        <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>Excel</button>
    </div>
    <div class="metric-chips" role="group" aria-label="Kennzahl">${Object.entries(YEAR_METRICS).map(([k, m]) =>
      `<button class="mchip${k === metric ? ' on' : ''}" data-metric="${k}">${m.label}</button>`).join('')}</div>
    <p class="muted small">${YEAR_METRICS[metric].hint}. Auf eine Zahl tippen öffnet den Monat; Maus darüber zeigt die Tage.</p>
    ${rows.length ? `<div class="matrix-wrap"><table class="ytable">
      <thead><tr><th class="name"></th>${head}<th>Σ</th></tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr><th class="name">Summe</th>${colSum.map(v => `<td>${v}</td>`).join('')}<td class="sum">${colSum.reduce((x, v) => x + (v || 0), 0)}</td></tr></tfoot>
    </table></div>` : '<p class="muted">Keine Einträge.</p>'}
    ${missing.length ? `<p class="muted small">Ohne Plan: ${missing.map(k => MONATE[+k.slice(5) - 1]).join(', ')}.</p>` : ''}
  </section>`;
}

function exportYearXlsx() {
  const data = state.data, year = state.year;
  const wb = XLSX.utils.book_new();
  for (const [k, m] of Object.entries(YEAR_METRICS)) {
    const { keys, rows, missing } = yearTable(data, year, k);
    const aoa = [[`${m.label} ${year} – ${m.hint}`], [],
      ['Name', 'Kürzel', ...keys.map(x => MONATE[+x.slice(5) - 1]), 'Summe'],
      ...rows.map(r => [plainName(r.a), r.a, ...r.vals.map(v => v === null ? '' : v), r.sum]),
      [], ...(missing.length ? [[`Kein Plan für: ${missing.map(x => MONATE[+x.slice(5) - 1]).join(', ')}`]] : [])];
    if (k === 'urlaub' || k === 'urlaubKT') {
      aoa.push([], ['Urlaubstage im Detail']);
      const t = YEAR_METRICS[k].test;
      for (const r of rows) {
        const days = keys.filter(x => data.months.has(x)).flatMap(x => data.months.get(x).days.filter(d => t(d, r.a)).map(d => d.date));
        aoa.push([plainName(r.a), r.a, ranges(days)]);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 24 }, { wch: 7 }, ...keys.map(() => ({ wch: 9 })), { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, ws, m.label.replace(/[\/\\?*\[\]:]/g, '-').slice(0, 31));
  }
  XLSX.writeFile(wb, `Jahresuebersicht_Assistenz_${year}.xlsx`);
}

// ---------------------------------------------------------------------------
// Export ganzer Monate (Excel / PDF)
// ---------------------------------------------------------------------------

function exportPanel(data) {
  if (!state.exportOpen) return '';
  const keys = [...data.months.keys()].sort();
  if (!state.exportSel) state.exportSel = new Set([state.month]);
  const dis = state.exportSel.size ? '' : 'disabled';
  return `<div class="export-panel">
    <h3>Monate exportieren</h3>
    <div class="exp-months">${keys.map(k => `<label class="exp-month${state.exportSel.has(k) ? ' on' : ''}">
      <input type="checkbox" class="exp-check" value="${k}" ${state.exportSel.has(k) ? 'checked' : ''}>${monthLabel(k)}</label>`).join('')}</div>
    <div class="row-actions">
      <button class="btn" data-action="export-pdf" ${dis}>PDF / Drucken</button>
      <button class="btn secondary" data-action="export-xlsx" ${dis}>Excel (.xlsx)</button>
    </div>
    <p class="muted small">PDF: im Druckdialog „Als PDF speichern“ wählen. Jeder Monat kommt auf eine eigene Seite (A4 quer).</p>
  </div>`;
}

const EXPORT_COLS = [
  ['dienst', 'Dienst Assistenz'], ['oa', 'OA'], ['frei', 'Frei nach Dienst'], ['urlaub', 'Urlaub'],
  ['wrt', 'Wochenruhetag'], ['efrei', 'Ersatzfrei'], ['fb', 'Fortbildung'], ['gm', 'Gmunden'], ['abw', 'Abwesend'],
];

/** Eine Zeile pro Tag, je Spalte die Namen. */
function exportRows(data, m) {
  return m.days.map(d => {
    const row = { date: d.date, hol: holidayOf(d.date) || '', note: d.note || '' };
    for (const [st] of EXPORT_COLS) {
      row[st] = st === 'oa' ? oaOf(d).map(oaShort).join(', ')
        : data.people.filter(a => statusesOf(d, a).includes(st) && (st === 'dienst' || !statusesOf(d, a).includes('dienst')))
            .map(st === 'dienst' ? plainName : surname).join(', ');
    }
    return row;
  });
}
function dienstCounts(data, m) {
  return data.people.map(a => [a, m.days.filter(d => statusesOf(d, a).includes('dienst')).length]).filter(([, n]) => n);
}
const selectedMonths = data => [...state.exportSel].filter(k => data.months.has(k)).sort().map(k => data.months.get(k));

function exportXlsx() {
  const data = state.data;
  const wb = XLSX.utils.book_new();
  const months = selectedMonths(data);
  for (const m of months) {
    const head = ['Datum', 'Tag', ...EXPORT_COLS.map(c => c[1]), 'Feiertag', 'Besonderheiten'];
    const rows = exportRows(data, m).map(r => {
      const [y, mo, d] = r.date.split('-');
      return [`${d}.${mo}.${y}`, WT[dow(r.date)], ...EXPORT_COLS.map(c => r[c[0]]), r.hol, r.note];
    });
    const counts = dienstCounts(data, m);
    const aoa = [[`Dienstplan Assistenz – ${monthLabel(m.key)}`], [], head, ...rows, [], ['Dienste'], ...counts.map(([a, n]) => [plainName(a), n])];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 11 }, { wch: 4 }, { wch: 22 }, { wch: 14 }, ...EXPORT_COLS.slice(2).map(() => ({ wch: 16 })), { wch: 18 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, ws, monthLabel(m.key).slice(0, 31));
  }
  const keys = months.map(m => m.key);
  const name = keys.length === 1 ? keys[0] : `${keys[0]}_bis_${keys[keys.length - 1]}`;
  XLSX.writeFile(wb, `Dienstplan_Assistenz_${name}.xlsx`);
}

function exportPdf() {
  const data = state.data;
  const pages = selectedMonths(data).map(m => {
    const rows = exportRows(data, m);
    const used = EXPORT_COLS.filter(([st]) => st === 'dienst' || st === 'oa' || rows.some(r => r[st]));
    const body = rows.map(r => `<tr class="${isWeekendOrHoliday(r.date) ? 'we' : ''}">
      <td class="pd">${+r.date.slice(8)}.</td><td class="pw">${WT[dow(r.date)]}</td>
      ${used.map(([st]) => `<td class="${st === 'dienst' ? 'pdienst' : ''}">${esc(r[st])}</td>`).join('')}
      <td class="pn">${esc([r.hol, r.note].filter(Boolean).join(' · '))}</td></tr>`).join('');
    const counts = dienstCounts(data, m).map(([a, n]) => `${esc(surname(a))} ${n}`).join(' · ');
    return `<section class="ppage">
      <header><h1>Dienstplan Assistenz – ${monthLabel(m.key)}</h1><span>Frauenheilkunde &amp; Geburtshilfe · Vöcklabruck</span></header>
      <table><thead><tr><th colspan="2">Datum</th>${used.map(([, l]) => `<th>${l}</th>`).join('')}<th>Feiertag / Bemerkung</th></tr></thead><tbody>${body}</tbody></table>
      <footer>Dienste: ${counts} · Stand ${new Date().toLocaleDateString('de-AT')}</footer>
    </section>`;
  }).join('');
  $('#printArea').innerHTML = pages;
  document.body.classList.add('printing');
  const done = () => { document.body.classList.remove('printing'); $('#printArea').innerHTML = ''; window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  setTimeout(() => window.print(), 50);
}

function planTable(data, month, me) {
  const today = todayIso();
  const head = month.days.map(d => {
    const hol = holidayOf(d.date);
    const c = [isWeekendOrHoliday(d.date) ? 'we' : '', hol ? 'hol' : '', d.date === today ? 'today' : ''].join(' ');
    return `<th class="${c}" title="${esc(hol || '')}">${WT[dow(d.date)]}<br>${+d.date.slice(8)}</th>`;
  }).join('');
  const oaRow = `<tr class="oa-row"><th class="name">OA</th>${month.days.map(d => `<td title="${esc(oaOf(d).map(oaName).join(', '))}">${esc(oaOf(d).join('/'))}</td>`).join('')}<td class="cnt"></td></tr>`;
  const body = data.people.map(a => {
    let n = 0;
    const cells = month.days.map(d => {
      const p = primary(d, a);
      if (p === 'dienst') n++;
      const we = isWeekendOrHoliday(d.date) ? ' we' : '';
      if (!p) return `<td class="${we}"></td>`;
      return `<td class="${STATUS[p].cls}${p === 'vb' ? we : ''}" title="${esc(fmtDay(d.date) + ': ' + statusesOf(d, a).map(s => STATUS[s].label).join(', '))}">${STATUS[p].short}</td>`;
    }).join('');
    return `<tr class="${a === me ? 'me' : ''}"><th class="name" title="${esc(personName(a))}">${esc(surname(a))}</th>${cells}<td class="cnt" title="Dienste">${n}</td></tr>`;
  }).join('');
  return `<div class="matrix-wrap"><table class="matrix">
      <thead><tr><th class="name"></th>${head}<th title="Dienste">Σ</th></tr></thead>
      <tbody>${oaRow}${body}</tbody>
    </table></div>
    ${legend(['dienst', 'frei', 'urlaub', 'fb', 'wrt', 'efrei', 'gm', 'abw', 'vb'])}
    <p class="muted small">D = Dienst · f = frei nach Dienst · U = Urlaub · W = Wochenruhetag · E = Ersatzfrei · · = im Haus</p>`;
}

function findDay(data, date) {
  const m = data.months.get(date.slice(0, 7));
  return m && m.days.find(d => d.date === date);
}

// ---------------------------------------------------------------------------
// Tag
// ---------------------------------------------------------------------------

function viewDay(data) {
  const days = allDays(data);
  const min = days[0].date, max = days[days.length - 1].date;
  if (!findDay(data, state.day)) state.day = todayIso() >= min && todayIso() <= max ? todayIso() : state.month + '-01';
  const d = findDay(data, state.day);
  const me = state.person;
  const who = st => data.people.filter(a => st === 'vb' ? primary(d, a) === 'vb' : statusesOf(d, a).includes(st));
  const names = arr => arr.length ? arr.map(a => `<span class="nm${a === me ? ' me' : ''}">${esc(plainName(a))}</span>`).join('') : '<span class="muted">–</span>';
  const dienst = who('dienst');
  const next = findDay(data, addDays(state.day, 1));
  const hol = holidayOf(state.day);

  const slots = [
    ['frei', 'Frei nach Dienst'], ['vb', 'Im Haus'], ['gm', 'Gmunden'], ['urlaub', 'Urlaub'],
    ['fb', 'Fortbildung'], ['wrt', 'Wochenruhetag'], ['efrei', 'Ersatzfrei'], ['abw', 'Abwesend'],
  ].filter(([st]) => who(st).length)
    .map(([st, t]) => `<div class="slot"><h3>${t}</h3><div class="names">${names(who(st))}</div></div>`).join('');

  return `
    <section class="card">
      <div class="day-picker">
        <button class="btn secondary" data-day="-1" aria-label="Vorheriger Tag">‹</button>
        <input type="date" id="dayInput" value="${state.day}" min="${min}" max="${max}" aria-label="Datum">
        <button class="btn secondary" data-day="1" aria-label="Nächster Tag">›</button>
        <button class="btn secondary" data-day="today">Heute</button>
      </div>
      <h2 class="day-title">${fmtLong(state.day)} <span class="muted">· ${relDays(state.day)}</span>${hol ? ` <span class="chip chip-hol">${esc(hol)}</span>` : ''}</h2>
      <div class="duty">
        <div class="duty-col"><h3>Dienst Assistenz</h3><div class="duty-name">${dienst.length ? dienst.map(a => esc(plainName(a)) + (a === me ? ' <span class="me-badge">Du</span>' : '')).join('<br>') : '–'}</div></div>
        <div class="duty-col"><h3>OA im Dienst</h3><div class="duty-name">${oaOf(d).length ? oaOf(d).map(k => esc(oaName(k))).join('<br>') : '<span class="dim">nicht eingetragen</span>'}</div></div>
      </div>
      ${slots ? `<div class="day-grid">${slots}</div>` : ''}
      ${next && dienstOf(next, data).length ? `<p class="muted small" style="margin-top:12px">Dienst am ${fmtDay(next.date)}: <strong>${dienstOf(next, data).map(a => esc(plainName(a))).join(', ')}</strong>${oaOf(next).length ? ` mit ${oaOf(next).map(k => esc(oaShort(k))).join(' / ')}` : ''}</p>` : ''}
      ${d.note ? `<div class="note">${esc(d.note)}</div>` : ''}
    </section>`;
}

// ---------------------------------------------------------------------------
// Kalender-Export (.ics)
// ---------------------------------------------------------------------------

function downloadIcs() {
  const a = state.person, data = state.data;
  const ev = [];
  for (const m of [...data.months.values()].sort((x, y) => x.key.localeCompare(y.key))) {
    for (const d of m.days) for (const s of statusesOf(d, a)) {
      if (!STATUS[s].ics) continue;
      // aufeinanderfolgende Tage (außer Dienste) zu einem Termin zusammenfassen
      const prev = [...ev].reverse().find(e => e.s === s);
      if (s !== 'dienst' && prev && prev.end === d.date) { prev.end = addDays(d.date, 1); continue; }
      ev.push({ s, start: d.date, end: addDays(d.date, 1) });
    }
  }
  const stamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15) + 'Z';
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Dienstplan Assistenz Gyn VB//DE', 'CALSCALE:GREGORIAN',
    `X-WR-CALNAME:Dienstplan ${a}`];
  for (const e of ev) {
    lines.push('BEGIN:VEVENT',
      `UID:${a}-${e.s}-${e.start}@dienstplan-gyn-vb`,
      `DTSTAMP:${stamp}`,
      `DTSTART;VALUE=DATE:${e.start.replace(/-/g, '')}`,
      `DTEND;VALUE=DATE:${e.end.replace(/-/g, '')}`,
      `SUMMARY:${STATUS[e.s].ics}`,
      'TRANSP:' + (e.s === 'dienst' ? 'OPAQUE' : 'TRANSPARENT'),
      'END:VEVENT');
  }
  lines.push('END:VCALENDAR');
  const blob = new Blob([lines.join('\r\n')], { type: 'text/calendar' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `dienstplan-${a}.ics`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

// ---------------------------------------------------------------------------
// Prüfung für die Planerin
// ---------------------------------------------------------------------------

function runChecks(data) {
  const out = [];
  const months = [...data.months.values()].sort((x, y) => x.key.localeCompare(y.key));
  const absent = ['abw', 'urlaub', 'fb', 'wrt', 'efrei', 'frei'];
  const cur = todayIso().slice(0, 7);
  for (const m of months) {
    if (m.key < cur) continue; // vergangene Monate nicht mehr prüfen
    const ml = monthLabel(m.key);
    const spaces = new Set();
    const unknown = new Set();
    for (const t of m.tokens) {
      if (t.status === 'dienst' && /\s$/.test(t.raw)) spaces.add(t.name);
      if (['dienst', 'dienst3', 'urlaub', 'frei', 'efrei', 'wrt', 'fb', 'abw'].includes(t.status) && !data.canon.has(t.name.toLowerCase()) && !data.oa.has(t.name.toLowerCase())) unknown.add(`„${t.name}“ (${fmtDay(t.date)})`);
    }
    if (unknown.size) out.push(['err', `${ml}: Unbekannte Kürzel ${[...unknown].join(', ')} – Tippfehler?`]);
    if (spaces.size) out.push(['warn', `${ml}: Kürzel mit Leerzeichen am Ende im Dienst (${[...spaces].map(s => `„${s} “`).join(', ')}). Die Web-Ansicht zählt sie richtig, aber ZÄHLENWENN im Excel übersieht diese Dienste.`]);

    for (const d of m.days) {
      const dienst = data.people.filter(a => statusesOf(d, a).includes('dienst'));
      if (!dienst.length && m.hasDienst) out.push(['warn', `${ml}: ${fmtDay(d.date)} ist kein Assistenz-Dienst eingetragen.`]);
      for (const a of dienst) {
        const sts = statusesOf(d, a).filter(s => absent.includes(s));
        if (sts.length) out.push(['err', `${fmtDay(d.date)}: ${a} hat Dienst, ist aber auch als ${sts.map(s => STATUS[s].label).join(', ')} eingetragen.`]);
        const prev = findDay(data, addDays(d.date, -1));
        if (prev && statusesOf(prev, a).includes('dienst')) out.push(['err', `${fmtDay(d.date)}: ${a} hat zwei Dienste hintereinander.`]);
        const next = findDay(data, addDays(d.date, 1));
        if (next && !statusesOf(next, a).includes('frei') && !statusesOf(next, a).includes('dienst'))
          out.push(['warn', `${fmtDay(next.date)}: ${a} ist nach dem Dienst nicht als „aus Dienst“ eingetragen.`]);
      }
    }
  }
  return out;
}

function renderChecks() {
  const box = $('#checks');
  if (!state.data || !state.data.months.size) { box.innerHTML = ''; return; }
  const checks = runChecks(state.data);
  const counts = state.data.people.map(a => {
    let n = 0, we = 0;
    for (const m of state.data.months.values()) for (const d of m.days) if (statusesOf(d, a).includes('dienst')) { n++; if (isWeekendOrHoliday(d.date)) we++; }
    return `${a}: ${n} (${we} WE/FT)`;
  }).join(' · ');
  box.innerHTML = `<div class="checks">
    ${checks.length ? checks.map(([t, msg]) => `<div class="check ${t}">${esc(msg)}</div>`).join('') : '<div class="check ok">Keine Auffälligkeiten gefunden.</div>'}
    <p class="muted small">Dienste gesamt (alle geladenen Monate): ${esc(counts)}</p></div>`;
}

// ---------------------------------------------------------------------------
// Vorschau & Veröffentlichen (für die Planerin)
// ---------------------------------------------------------------------------

async function previewFiles(fileList) {
  const files = [...fileList].filter(f => /\.xls[xm]?$/i.test(f.name));
  if (!files.length) return;
  try {
    const pending = await Promise.all(files.map(async f => ({ name: f.name, data: b64(await f.arrayBuffer()) })));
    const data = buildData(pending.flatMap(f => parseWorkbook(unb64(f.data), f.name)));
    if (!data.months.size) { alert('In dieser Datei wurde kein Monat mit Diensten gefunden.'); return; }
    state.pending = pending;
    state.preview = pending.map(f => f.name).join(', ');
    state.data = data;
    state.month = null;
    state.month = pickMonth(data);
    $('#previewName').textContent = state.preview;
    $('#previewBanner').hidden = false;
    $('.foot details').open = true;
    render();
  } catch (e) {
    alert('Die Datei konnte nicht gelesen werden: ' + e.message);
  }
}
function endPreview() {
  state.preview = null;
  state.pending = null;
  state.data = state.published;
  state.month = pickMonth(state.data);
  $('#previewBanner').hidden = true;
  $('#publishHelp').hidden = true;
  render();
}

/** Bisherige Pläne + neue Dateien (gleicher Dateiname wird ersetzt, neuere Dateien gewinnen). */
async function publishPending() {
  let pw = state.password;
  if (!pw) {
    pw = prompt('Noch kein Plan veröffentlicht. Neues Team-Passwort festlegen:');
    if (!pw) return;
  }
  const names = new Set(state.pending.map(f => f.name));
  const files = [...state.files.filter(f => !names.has(f.name)), ...state.pending];
  const enc = await encryptBundle(files, pw);
  const link = document.createElement('a');
  link.href = URL.createObjectURL(new Blob([JSON.stringify(enc)], { type: 'application/octet-stream' }));
  link.download = 'plan.enc';
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  $('#publishHelp').hidden = false;
}

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

document.addEventListener('submit', async e => {
  if (e.target.id !== 'lockForm') return;
  e.preventDefault();
  const pw = $('#pwInput').value;
  const btn = e.target.querySelector('button'); btn.disabled = true; btn.textContent = 'Öffne …';
  if (await unlock(pw) && $('#pwRemember').checked) safeSet('dp.pw', pw);
  render();
});

document.addEventListener('click', e => {
  const t = e.target.closest('button, [data-person], [data-goto], [data-month]');
  if (!t) return;
  if (t.classList.contains('tab')) {
    state.view = t.dataset.view; safeSet('dp.view', state.view);
    if (state.view === 'day') state.day = todayIso();
    state.scrollToday = state.view === 'plan';
    render(); window.scrollTo(0, 0);
  }
  else if (t.dataset.person) { state.person = t.dataset.person; state.choosing = false; safeSet('dp.person', state.person); render(); window.scrollTo(0, 0); }
  else if (t.dataset.action === 'choose') { state.choosing = true; render(); }
  else if (t.dataset.metric) { state.yearMetric = t.dataset.metric; safeSet('dp.yearMetric', state.yearMetric); render(); }
  else if (t.dataset.action === 'year-xlsx') exportYearXlsx();
  else if (t.dataset.month) { state.month = t.dataset.month; state.view = 'plan'; state.planMode = 'table'; render(); window.scrollTo(0, 0); }
  else if (t.dataset.action === 'export-toggle') { state.exportOpen = !state.exportOpen; state.exportSel = null; render(); }
  else if (t.dataset.action === 'export-xlsx') exportXlsx();
  else if (t.dataset.action === 'export-pdf') exportPdf();
  else if (t.dataset.mode) { state.planMode = t.dataset.mode; safeSet('dp.planMode', state.planMode); render(); }
  else if (t.dataset.goto) { state.day = t.dataset.goto; state.month = state.day.slice(0, 7); state.view = 'day'; render(); window.scrollTo(0, 0); }
  else if (t.id === 'prevMonth' || t.id === 'nextMonth') {
    if (!state.data) return;
    const keys = [...state.data.months.keys()].sort();
    const i = keys.indexOf(state.month) + (t.id === 'nextMonth' ? 1 : -1);
    if (keys[i]) { state.month = keys[i]; state.day = state.month + '-01'; render(); }
  }
  else if (t.id === 'icsBtn') downloadIcs();
  else if (t.id === 'endPreview') endPreview();
  else if (t.id === 'publishBtn') publishPending();
  else if (t.id === 'logoutBtn') { safeDel('dp.pw'); location.reload(); }
  else if (t.dataset.day) {
    state.day = t.dataset.day === 'today' ? todayIso() : addDays(state.day, +t.dataset.day);
    if (!findDay(state.data, state.day)) { alert('Für diesen Tag gibt es noch keinen Plan.'); state.day = addDays(state.day, t.dataset.day === 'today' ? 0 : -t.dataset.day); if (!findDay(state.data, state.day)) state.day = state.month + '-01'; }
    state.month = state.day.slice(0, 7);
    render();
  }
});
document.addEventListener('change', e => {
  if (e.target.id === 'dayInput' && e.target.value) {
    state.day = e.target.value; state.month = state.day.slice(0, 7);
    if (!state.data.months.has(state.month)) state.month = pickMonth(state.data);
    render();
  }
  if (e.target.id === 'yearSel') { state.year = e.target.value; render(); }
  if (e.target.classList.contains('exp-check')) {
    if (e.target.checked) state.exportSel.add(e.target.value); else state.exportSel.delete(e.target.value);
    render();
  }
  if (e.target.id === 'onlyMine') { state.onlyMine = e.target.checked; safeSet('dp.onlyMine', state.onlyMine ? '1' : '0'); render(); }
  if (e.target.id === 'fileInput' && e.target.files.length) { previewFiles(e.target.files); e.target.value = ''; }
});

let dragDepth = 0;
const canUpload = () => !state.locked;
window.addEventListener('dragenter', e => { if (canUpload() && [...e.dataTransfer.types].includes('Files')) { dragDepth++; $('#dropOverlay').hidden = false; } });
window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('#dropOverlay').hidden = true; } });
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => {
  e.preventDefault(); dragDepth = 0; $('#dropOverlay').hidden = true;
  if (canUpload() && e.dataTransfer.files.length) previewFiles(e.dataTransfer.files);
});

// Start
(async () => {
  try {
    const data = await loadPublished();
    if (data) {
      state.published = state.data = data;
      state.month = pickMonth(state.data);
    }
    if (state.data && state.view === 'day' && state.data.months.size && !findDay(state.data, state.day)) state.day = state.month + '-01';
  } catch (e) {
    $('#app').innerHTML = `<div class="empty">Fehler beim Laden: ${esc(e.message)}</div>`;
    return;
  }
  render();
})();
