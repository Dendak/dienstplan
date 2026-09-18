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
};

// Kürzel von Primar/Oberärzt:innen – werden ignoriert (nicht als "unbekannt" gemeldet)
const OA_KUERZEL = ['Br', 'Was', 'Co', 'Ki', 'Ra', 'Pf', 'Pu', 'Bl', 'Ba', 'Gi', 'Ko', 'Sch'];

// Spaltenüberschrift (Zeile 1 im Excel) → Status
const SPALTEN = {
  'nd1': 'dienst', 'nd2': 'dienst', 'nd3': 'dienst',
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
const personName = a => ASSISTENZ[a] || a;
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

function splitNames(v) {
  if (v == null) return [];
  return String(v).split(/[,;\/+\n]/).map(t => ({ raw: t, name: t.trim() })).filter(t => t.name);
}

/** Liest ein Tabellenblatt; liefert null, wenn es kein Monatsblatt ist. */
function parseSheet(ws, sheetName, source) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!rows.length || typeof rows[0][0] !== 'number') return null;

  const cols = rows[0].map(v => SPALTEN[headerText(v).toLowerCase().replace(/\s+/g, '')] || null);
  const days = [];
  const tokens = []; // für die Prüfung
  let r = 1;
  for (; r < rows.length; r++) {
    const row = rows[r];
    if (typeof row[0] !== 'number' || row[0] < 30000) break;
    const date = serialToIso(row[0]);
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
  for (const m of months.values()) {
    m.listed.forEach(a => found.add(a));
    m.tokens.forEach(t => { if (t.status === 'vb') found.add(t.name); });
  }
  const canon = new Map();
  [...known, ...found].forEach(a => { if (!oa.has(a.toLowerCase())) canon.set(a.toLowerCase(), canon.get(a.toLowerCase()) || a); });

  // Einträge normalisieren (Groß/Klein, Leerzeichen) und auf Assistenz beschränken
  for (const m of months.values()) {
    for (const d of m.days) {
      const norm = {};
      for (const [n, set] of Object.entries(d.entries)) {
        const a = canon.get(n.toLowerCase());
        if (!a) continue;
        norm[a] ||= new Set();
        set.forEach(s => norm[a].add(s));
      }
      d.entries = norm;
    }
    // Monate ohne einen einzigen Dienst gelten als leer (z. B. leere Jahresvorlage)
    m.hasDienst = m.days.some(d => Object.values(d.entries).some(s => s.has('dienst')));
  }
  for (const [k, m] of months) if (!m.hasDienst) months.delete(k);

  // Personen sortieren: bekannte zuerst in Konfig-Reihenfolge, dann weitere
  const active = new Set();
  for (const m of months.values()) for (const d of m.days) Object.keys(d.entries).forEach(a => active.add(a));
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
};
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
  if (state.view === 'team') app.innerHTML = viewTeam(data, month);
  else if (state.view === 'day') app.innerHTML = viewDay(data);
  else app.innerHTML = viewMe(data, month);
  renderChecks();
}

function peoplePicker(data) {
  return `<div class="people">${data.people.map(a => `
    <button class="person${a === state.person ? ' active' : ''}" data-person="${esc(a)}" title="${esc(personName(a))}">
      <span class="av">${esc(initials(a))}</span>${esc(ASSISTENZ[a] ? ASSISTENZ[a].replace(/^Dr\.(in)? /, '') : a)}
    </button>`).join('')}</div>`;
}

function legend(keys) {
  return `<div class="legend">${keys.map(k => `<span><i class="${STATUS[k].cls}"${k === 'vb' ? ' style="border:1px solid var(--line)"' : ''}></i>${STATUS[k].short} = ${STATUS[k].label}</span>`).join('')}</div>`;
}

function viewMe(data, month) {
  const a = state.person;
  if (!a || !data.people.includes(a)) {
    return `<section class="card"><h2>Wer bist du?</h2>${peoplePicker(data)}
      <p class="muted small">Die Auswahl wird nur auf diesem Gerät gespeichert.</p></section>`;
  }
  const days = month.days;
  const dienste = days.filter(d => statusesOf(d, a).includes('dienst'));
  const weDienste = dienste.filter(d => isWeekendOrHoliday(d.date));
  const count = st => days.filter(d => statusesOf(d, a).includes(st)).length;
  const today = todayIso();

  // Kalender (Montag zuerst)
  const lead = (dow(days[0].date) + 6) % 7;
  let cal = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'].map(w => `<div class="dow">${w}</div>`).join('');
  cal += '<div class="cell blank"></div>'.repeat(lead);
  for (const d of days) {
    const sts = statusesOf(d, a).filter(s => s !== 'vb' || statusesOf(d, a).length === 1);
    const hol = holidayOf(d.date);
    const cls = ['cell', isWeekendOrHoliday(d.date) && !hol ? 'we' : '', hol ? 'hol' : '', d.date === today ? 'today' : ''].join(' ');
    cal += `<div class="${cls}" title="${esc(hol || '')}"><span class="d">${+d.date.slice(8)}</span>
      ${sts.slice(0, 2).map(s => `<span class="pill ${STATUS[s].cls}" title="${STATUS[s].label}">${STATUS[s].cal || STATUS[s].label}</span>`).join('')}</div>`;
  }

  // Termine
  const items = [];
  for (const d of days) for (const s of statusesOf(d, a)) if (s !== 'vb') items.push({ d, s });
  const upcoming = items.filter(i => i.d.date >= today);
  const list = (upcoming.length ? upcoming : items);

  return `
    <section class="card"><h2>Person</h2>${peoplePicker(data)}</section>
    <section class="card">
      <h2>${esc(personName(a))} <span class="muted">· ${monthLabel(state.month)}</span></h2>
      <div class="stats">
        <div class="stat"><b>${dienste.length}</b><span>Dienste</span></div>
        <div class="stat"><b>${weDienste.length}</b><span>davon Wochenende/Feiertag</span></div>
        <div class="stat"><b>${count('urlaub')}</b><span>Urlaubstage</span></div>
        <div class="stat"><b>${count('fb') + count('wrt') + count('efrei')}</b><span>FB / WRT / Ersatzfrei</span></div>
      </div>
    </section>
    <section class="card">
      <div class="cal">${cal}</div>
      ${legend(['dienst', 'frei', 'urlaub', 'fb', 'wrt', 'efrei', 'gm', 'abw'])}
    </section>
    <section class="card">
      <h2>${upcoming.length ? 'Nächste Einträge' : 'Einträge im Monat'}</h2>
      ${list.length ? `<ul class="list">${list.map(i => `
        <li><span class="when">${fmtDay(i.d.date)}</span><span class="tag ${STATUS[i.s].cls}">${STATUS[i.s].label}</span>
        ${holidayOf(i.d.date) ? `<span class="muted small">${esc(holidayOf(i.d.date))}</span>` : ''}</li>`).join('')}</ul>`
        : '<p class="muted">Keine Einträge.</p>'}
      <div class="row-actions">
        <button class="btn" id="icsBtn">In meinen Kalender (.ics)</button>
        <button class="btn secondary" onclick="window.print()">Drucken</button>
      </div>
      <p class="muted small">Die .ics-Datei enthält alle veröffentlichten Monate und lässt sich in Outlook, Google- oder iPhone-Kalender importieren.</p>
    </section>`;
}

function viewTeam(data, month) {
  const today = todayIso();
  const head = month.days.map(d => {
    const hol = holidayOf(d.date);
    const c = [isWeekendOrHoliday(d.date) && !hol ? 'we' : '', hol ? 'hol' : '', d.date === today ? 'today' : ''].join(' ');
    return `<th class="${c}" title="${esc(hol || '')}">${WT[dow(d.date)].slice(0, 2)}<br>${+d.date.slice(8)}</th>`;
  }).join('');
  const body = data.people.map(a => {
    let n = 0;
    const cells = month.days.map(d => {
      const p = primary(d, a);
      if (p === 'dienst') n++;
      const we = isWeekendOrHoliday(d.date) ? ' we' : '';
      if (!p) return `<td class="${we}"></td>`;
      return `<td class="${STATUS[p].cls}${p === 'vb' ? we : ''}" title="${esc(fmtDay(d.date) + ': ' + statusesOf(d, a).map(s => STATUS[s].label).join(', '))}">${STATUS[p].short}</td>`;
    }).join('');
    return `<tr class="${a === state.person ? 'me' : ''}"><th class="name" title="${esc(personName(a))}">${esc(a)}</th>${cells}<td class="cnt" title="Dienste">${n}</td></tr>`;
  }).join('');

  // Dienst-Zeile oben: wer hat an welchem Tag Dienst
  return `
    <section class="card">
      <h2>Team · ${monthLabel(state.month)}</h2>
      <div class="matrix-wrap"><table class="matrix">
        <thead><tr><th class="name"></th>${head}<th title="Dienste">Σ</th></tr></thead>
        <tbody>${body}</tbody>
      </table></div>
      ${legend(['dienst', 'frei', 'urlaub', 'fb', 'wrt', 'efrei', 'gm', 'abw', 'vb'])}
    </section>
    <section class="card">
      <h2>Kürzel</h2>
      <ul class="list">${data.people.map(a => `<li><span class="when">${esc(a)}</span>${esc(personName(a))}</li>`).join('')}</ul>
    </section>`;
}

function findDay(data, date) {
  const m = data.months.get(date.slice(0, 7));
  return m && m.days.find(d => d.date === date);
}

function viewDay(data) {
  const keys = [...data.months.keys()].sort();
  const min = data.months.get(keys[0]).days[0].date;
  const lastM = data.months.get(keys[keys.length - 1]);
  const max = lastM.days[lastM.days.length - 1].date;
  if (state.day < min || state.day > max || !findDay(data, state.day)) {
    state.day = state.month + '-01';
  }
  const d = findDay(data, state.day);
  // "Im Haus" nur, wer an dem Tag nichts anderes eingetragen hat
  const who = st => data.people.filter(a => st === 'vb' ? primary(d, a) === 'vb' : statusesOf(d, a).includes(st));
  const names = arr => arr.length ? arr.map(a => `<span title="${esc(personName(a))}">${esc(personName(a).replace(/^Dr\.(in)? /, ''))}</span>`).join('<br>') : '<span class="muted">–</span>';
  const next = findDay(data, addDays(state.day, 1));
  const nextDienst = next ? data.people.filter(a => statusesOf(next, a).includes('dienst')) : [];
  const hol = holidayOf(state.day);

  const slots = [
    ['frei', 'Frei nach Dienst'], ['vb', 'Im Haus (VB)'], ['gm', 'Gmunden'], ['urlaub', 'Urlaub'],
    ['fb', 'Fortbildung'], ['wrt', 'Wochenruhetag'], ['efrei', 'Ersatzfrei'], ['abw', 'Abwesend'],
  ].filter(([st]) => st === 'vb' || st === 'frei' || who(st).length)
    .map(([st, t]) => `<div class="slot"><h3>${t}</h3><div class="names">${names(who(st))}</div></div>`).join('');

  return `
    <section class="card">
      <div class="day-picker">
        <button class="btn secondary" data-day="-1" aria-label="Vorheriger Tag">‹</button>
        <input type="date" id="dayInput" value="${state.day}" min="${min}" max="${max}">
        <button class="btn secondary" data-day="1" aria-label="Nächster Tag">›</button>
        <button class="btn secondary" data-day="today">Heute</button>
      </div>
      <h2 style="margin-top:14px">${fmtLong(state.day)}${hol ? ` <span class="muted">· ${esc(hol)}</span>` : ''}</h2>
      <div class="day-grid">
        <div class="slot big"><h3>Dienst</h3><div class="names">${names(who('dienst'))}</div></div>
        ${slots}
      </div>
      ${nextDienst.length ? `<p class="muted small" style="margin-top:12px">Dienst am ${fmtDay(next.date)}: <strong>${nextDienst.map(a => esc(personName(a))).join(', ')}</strong></p>` : ''}
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
  for (const m of months) {
    const ml = monthLabel(m.key);
    const spaces = new Set();
    const unknown = new Set();
    for (const t of m.tokens) {
      if (t.status === 'dienst' && /\s$/.test(t.raw)) spaces.add(t.name);
      if (!data.canon.has(t.name.toLowerCase()) && !data.oa.has(t.name.toLowerCase())) unknown.add(`„${t.name}“ (${fmtDay(t.date)})`);
    }
    if (unknown.size) out.push(['err', `${ml}: Unbekannte Kürzel ${[...unknown].join(', ')} – Tippfehler?`]);
    if (spaces.size) out.push(['warn', `${ml}: Kürzel mit Leerzeichen am Ende im Dienst (${[...spaces].map(s => `„${s} “`).join(', ')}). Die Web-Ansicht zählt sie richtig, aber ZÄHLENWENN im Excel übersieht diese Dienste.`]);

    for (const d of m.days) {
      const dienst = data.people.filter(a => statusesOf(d, a).includes('dienst'));
      if (!dienst.length) out.push(['warn', `${ml}: ${fmtDay(d.date)} ist kein Assistenz-Dienst eingetragen.`]);
      if (dienst.length > 1) out.push(['warn', `${fmtDay(d.date)}: mehrere Assistenz-Dienste (${dienst.join(', ')}).`]);
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
  const t = e.target.closest('button, [data-person]');
  if (!t) return;
  if (t.classList.contains('tab')) { state.view = t.dataset.view; safeSet('dp.view', state.view); render(); }
  else if (t.dataset.person) { state.person = t.dataset.person; safeSet('dp.person', state.person); render(); }
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
