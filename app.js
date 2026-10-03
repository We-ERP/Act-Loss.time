'use strict';

/* ════════════════════════════════════════════════════════════════════════════
   CC Performance Dashboard
   ────────────────────────────────────────────────────────────────────────────
   المعادلات (لكل موظف ولكل يوم):
     Assigning Tkts = عدد تذاكر IR اللي assigned_to  = TTS User  (عمود X)
     TKT            = عدد تذاكر IR اللي added_by     = TTS User  (عمود Y)
     التاريخ        = added_on (عمود Z)
     System         = TKT × 0.00104166666666667 يوم  (= TKT × 90 ثانية)
     Talk Time      = Hold + Other + AUX + ACW من UTL (بالـ Login ID)
     Tele-SCH       = مجموع Duration للأكواد المختارة من Schedule
     Comp           = من Compensation
     Loss Time      = Tele-SCH × 90%  −  (System + Talk Time + Comp)   (لا يقل عن صفر)
   ════════════════════════════════════════════════════════════════════════════ */

const DEFAULT_CONFIG = {
  teleSchFactor: 0.9,
  secondsPerTicket: 90,
  scheduleCodes: []
};

// أعمدة IR الاحتياطية لو الهيدر مش موجود (A=0 ... X=23, Y=24, Z=25)
const IR_COL = { assigned: 23, added: 24, date: 25 };
// أعمدة Schedule الخام (نفس الماكرو): B=ID/Login, C=التاريخ أو الكود, J=Duration
const SCHEDULE_COL = { id: 1, dateOrCode: 2, duration: 9 };
const PARTIAL_MATCH_MIN_LENGTH = 5;
const CODES_STORAGE_KEY = 'ccScheduleCodes';

let CONFIG = { ...DEFAULT_CONFIG };
let configPromise = null;

let processedMatrixData = [];
let visibleMatrixData = [];
let dateGroups = [];
let collapsedDays = {};
let activeColumns = [];
let inferredYear = new Date().getFullYear();

const sourceRows = { structure: [], utl: [], ir: [], comp: [] };
const filesState = { struct: null, schedule: null, utl: null, ir: null, comp: null };
const sourceLabels = { structure: 'STR Loss.xlsx من المستودع', schedule: 'لم يتم رفع Schedule' };

let scheduleCache = { file: null, parsed: null };
let scheduleCodeTotals = new Map(); // lower → { label, secs }
let selectedCodes = new Set();       // lower-case codes

const FIELD_ALIASES = {
  structureId: ['Teleopti ID', 'Teleopti', 'ST_ID'],
  loginId: ['Login ID', 'Login', 'UL_lo', 'Username'],
  perm: ['Perm'],
  ttsUser: ['TTS User', 'TTS'],
  bssUser: ['BSS User', 'BSS'],
  group: ['Group'],
  agentName: ['Agent Name', 'Agent'],
  status: ['Status'],
  tlId: ['TL ID', 'TL Id'],
  tlName: ['TL Name', 'Team Leader'],
  structureDate: ['ST_D'],
  structureDuration: ['ST_Du', 'ST Duration'],
  irAssigned: ['assigned_to'],
  irAdded: ['added_by', 'IR_L_E'],
  irDate: ['added_on', 'Date'],
  utlUser: ['UL_lo', 'Login ID', 'Login'],
  utlDate: ['UL_Date', 'Date'],
  compId: ['Comp_ID', 'Comp ID', 'Teleopti ID', 'ST_ID'],
  compDate: ['Comp_Da', 'Date'],
  compDuration: ['Comp_Du', 'Comp Duration', 'Duration'],
  scheduleAgent: ['Agent', 'Agent Name', 'Employee', 'Employee Name'],
  scheduleDate: ['Date', 'Scheduled Date'],
  scheduleDuration: ['Scheduled time', 'Scheduled Time', 'Scheduled-Time', 'Scheduled_Time']
};

const SCHEDULE_SHEET_HINTS = ['schedule', 'scheduled time', 'scheduled time per agent', 'scheduled', 'rd'];
const TOTALS_WORDS = ['totals', 'total', 'الاجمالي', 'الإجمالي', 'اجمالي', 'إجمالي', 'المجموع'];

const FIXED_COLUMNS = [
  { key: 'teleoptiId', label: 'Teleopti ID', always: true },
  { key: 'loginId', label: 'Login ID', always: true },
  { key: 'perm', label: 'Perm' },
  { key: 'ttsUser', label: 'TTS User', always: true },
  { key: 'bssUser', label: 'BSS User' },
  { key: 'group', label: 'Group' },
  { key: 'agentName', label: 'Agent Name', always: true },
  { key: 'status', label: 'Status', always: true },
  { key: 'tlId', label: 'TL ID' },
  { key: 'tlName', label: 'TL Name', always: true }
];

const DAY_METRIC_LABELS = ['Assigning Tkts', 'TKT', 'System', 'Talk Time', 'Tele-SCH', 'Comp', 'Loss Time'];

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

/* ─── ربط حقول رفع الملفات ─────────────────────────────────────────────── */
['struct', 'schedule', 'utl', 'ir', 'comp'].forEach(key => {
  const input = document.getElementById(`file-${key}`);
  if (!input) return;

  input.addEventListener('change', event => {
    const file = event.target.files[0];
    if (!file) return;

    filesState[key] = file;
    const name = document.getElementById(`name-${key}`);
    const card = document.getElementById(`card-${key}`);
    if (name) name.textContent = file.name;
    if (card) card.dataset.ready = 'true';

    if (key === 'schedule') prepareSchedule(file);
  });
});

/* ─── أدوات عامة ───────────────────────────────────────────────────────── */
function normalise(value) {
  let v = value;
  if (typeof v === 'number' && Number.isInteger(v)) v = String(v);
  return String(v ?? '').trim().toLowerCase().replace(/[^a-z0-9\u0600-\u06ff]+/gi, '');
}

// بحث عن اسم العمود: تطابق كامل أولاً (بترتيب الأولوية)، وبعدين تطابق جزئي للأسماء الطويلة فقط
function findKey(row, aliases) {
  if (!row) return undefined;
  const keys = Object.keys(row);
  const wanted = aliases.map(normalise).filter(Boolean);

  for (const w of wanted) {
    const hit = keys.find(k => normalise(k) === w);
    if (hit !== undefined) return hit;
  }
  for (const w of wanted) {
    if (w.length < PARTIAL_MATCH_MIN_LENGTH) continue;
    const hit = keys.find(k => normalise(k).includes(w));
    if (hit !== undefined) return hit;
  }
  return undefined;
}

function findValue(row, aliases, fallback = '') {
  const key = findKey(row, aliases);
  return key === undefined ? fallback : row[key];
}

function formatDateParts(year, month, day) {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function formatLocalDate(date) {
  return formatDateParts(date.getFullYear(), date.getMonth() + 1, date.getDate());
}

// التواريخ عندك Month-First (9/1/2026 = 1 سبتمبر). لو الرقم الأول > 12 يتعامل معاه كيوم.
// أي نص مش تاريخ صريح (زي "Break 1") بيرجع فاضي — مفيش new Date(text) عشان ميفهمش الأكواد كتواريخ.
function dateKey(value) {
  if (value === null || value === undefined || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatLocalDate(value);

  if (typeof value === 'number') {
    if (value > 20000 && value < 80000 && typeof XLSX !== 'undefined' && XLSX.SSF) {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return formatDateParts(parsed.y, parsed.m, parsed.d);
    }
    return '';
  }

  const text = String(value).trim();
  if (!text) return '';

  let m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  if (m) return formatDateParts(+m[1], +m[2], +m[3]);

  m = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:\s.*)?$/);
  if (m) {
    let a = +m[1];
    let b = +m[2];
    let year = +m[3];
    if (year < 100) year += 2000;
    let month = a;
    let day = b;
    if (a > 12 && b <= 12) { month = b; day = a; }
    if (month < 1 || month > 12 || day < 1 || day > 31) return '';
    return formatDateParts(year, month, day);
  }

  m = text.match(/^(\d{1,2})[\s\-\/]([A-Za-z]{3})[A-Za-z]*(?:[\s\-\/,]+(\d{2,4}))?$/);
  if (m) {
    const month = MONTHS[m[2].toLowerCase()];
    if (!month) return '';
    let year = m[3] ? +m[3] : inferredYear;
    if (year < 100) year += 2000;
    return formatDateParts(year, month, +m[1]);
  }

  return '';
}

function displayDate(value) {
  if (!value) return '';
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }).replace(' ', '-');
}

function parseSeconds(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') {
    return value > 0 && value < 1 ? Math.round(value * 86400) : value;
  }

  const text = String(value).trim();
  const lower = text.toLowerCase();
  if (!text || ['unpaid', 'maternity', 'planned sick'].includes(lower)) return 0;

  const parts = text.split(':').map(Number);
  if (parts.some(Number.isNaN)) return 0;
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  if (parts.length === 2) return parts[0] * 3600 + parts[1] * 60; // hh:mm
  return Number(text) || 0;
}

function formatTime(seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) return '0:00:00';
  const total = Math.round(seconds);
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

/* ─── قراءة الملفات ────────────────────────────────────────────────────── */
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        resolve(XLSX.read(new Uint8Array(event.target.result), { type: 'array' }));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function readBundledStructure() {
  let response;
  try {
    response = await fetch(encodeURI('STR Loss.xlsx'), { cache: 'no-store' });
  } catch (error) {
    throw new Error('تعذر تحميل STR Loss.xlsx تلقائياً. شغّل الصفحة عبر GitHub Pages أو localhost، أو ارفع Structure يدوياً.');
  }
  if (!response.ok) throw new Error('لم يتم العثور على STR Loss.xlsx داخل جذر المستودع');

  return XLSX.read(new Uint8Array(await response.arrayBuffer()), { type: 'array' });
}

function getOrderedSheetNames(workbook, words) {
  const names = workbook?.SheetNames || [];
  const hits = names.filter(n => words.some(w => normalise(n).includes(normalise(w))));
  return [...hits, ...names.filter(n => !hits.includes(n))];
}

// يرجّع صفوف أول شيت صالح (يفضّل الشيتات اللي اسمها قريب من hints)
function chooseSheet(workbook, hints, options = {}) {
  const { validate } = options;
  for (const name of getOrderedSheetNames(workbook, hints)) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '', raw: false });
    if (!rows.length) continue;
    if (!validate || validate(rows)) return rows;
  }
  return [];
}

function hasAgentNameColumn(rows) {
  return findKey(rows[0], FIELD_ALIASES.agentName) !== undefined;
}

async function ensureConfig() {
  if (configPromise) return configPromise;

  configPromise = (async () => {
    if (typeof fetch === 'function') {
      try {
        const response = await fetch('config.json', { cache: 'no-store' });
        if (response.ok) CONFIG = { ...DEFAULT_CONFIG, ...(await response.json()) };
      } catch (error) { /* نكمل بالقيم الافتراضية */ }
    }

    let saved = null;
    try {
      if (typeof localStorage !== 'undefined') saved = JSON.parse(localStorage.getItem(CODES_STORAGE_KEY) || 'null');
    } catch (error) { saved = null; }

    const initial = Array.isArray(saved) ? saved : CONFIG.scheduleCodes || [];
    selectedCodes = new Set(initial.map(c => String(c).trim().toLowerCase()).filter(Boolean));
    return CONFIG;
  })();

  return configPromise;
}

/* ─── الفهارس ──────────────────────────────────────────────────────────── */
function makeLookupKey(value, day) {
  const n = normalise(value);
  return n && day ? `${n}|${day}` : '';
}

function addToIndex(map, value, day, amount) {
  const key = makeLookupKey(value, day);
  if (key) map.set(key, (map.get(key) || 0) + amount);
}

function buildSumIndex(rows, userAliases, dateAliases, valueAliases) {
  const map = new Map();
  rows.forEach(row => {
    addToIndex(
      map,
      findValue(row, userAliases),
      dateKey(findValue(row, dateAliases)),
      parseSeconds(findValue(row, valueAliases, 0))
    );
  });
  return map;
}

function buildTalkTimeIndex(rows) {
  const map = new Map();
  rows.forEach(row => {
    const total =
      parseSeconds(findValue(row, ['Hold Time', 'HoldTime'], 0)) +
      parseSeconds(findValue(row, ['Other Time', 'OtherTime'], 0)) +
      parseSeconds(findValue(row, ['AUXOUTOFFTIME'], 0)) +
      parseSeconds(findValue(row, ['ACWOUTOFFTIME'], 0));
    addToIndex(map, findValue(row, FIELD_ALIASES.utlUser), dateKey(findValue(row, FIELD_ALIASES.utlDate)), total);
  });
  return map;
}

// IR: assigned_to (X) → Assigning | added_by (Y) → TKT | added_on (Z) → التاريخ
// الأولوية لأسماء الأعمدة، ولو مش موجودة نرجع لمكان العمود X/Y/Z.
function buildIRIndexes(rows) {
  const assigning = new Map();
  const tkt = new Map();
  if (!rows.length) return { assigning, tkt };

  const keys = Object.keys(rows[0]);
  const colAssigned = findKey(rows[0], FIELD_ALIASES.irAssigned) ?? keys[IR_COL.assigned];
  const colAdded = findKey(rows[0], FIELD_ALIASES.irAdded) ?? keys[IR_COL.added];
  const colDate = findKey(rows[0], FIELD_ALIASES.irDate) ?? keys[IR_COL.date];

  rows.forEach(row => {
    const day = dateKey(row[colDate]);
    if (!day) return;
    addToIndex(assigning, row[colAssigned], day, 1);
    addToIndex(tkt, row[colAdded], day, 1);
  });

  return { assigning, tkt };
}

function lookupFirst(map, candidates, day) {
  for (const candidate of candidates) {
    const key = makeLookupKey(candidate, day);
    if (key && map.has(key)) return map.get(key);
  }
  return 0;
}

/* ─── Schedule ─────────────────────────────────────────────────────────── */
// يقرأ التقرير الخام (RD) بنفس منطق الماكرو:
//   B = "ID + اسم" يبدأ بلوك موظف | C = تاريخ (صف يوم) أو كود النشاط | J = Duration
function parseScheduleMatrix(matrix) {
  const index = new Map();
  const days = new Set();
  const codeTotals = new Map();
  let login = '';
  let day = '';

  const entryFor = (id, d) => {
    const key = makeLookupKey(id, d);
    if (!index.has(key)) index.set(key, { total: 0, hasTotal: false, codes: new Map() });
    return index.get(key);
  };

  for (const row of matrix) {
    const cells = [0, 1, 2].map(i => String(row[i] ?? '').trim());

    if (cells.some(c => c && TOTALS_WORDS.some(w => normalise(c) === normalise(w)))) {
      login = '';
      day = '';
      continue;
    }

    const idMatch = cells[1].match(/\b(\d{5,6})\b/);
    if (idMatch) {
      login = idMatch[1];
      day = '';
    }

    if (!login || cells[2] === '') continue;

    const rawTime = row[SCHEDULE_COL.duration];
    const hasTime = String(rawTime ?? '').trim() !== '';
    const d = dateKey(row[SCHEDULE_COL.dateOrCode]);

    if (d) {
      day = d;
      days.add(d);
      const entry = entryFor(login, d);
      if (hasTime) {
        entry.total += parseSeconds(rawTime);
        entry.hasTotal = true;
      }
      continue;
    }

    if (!day || !hasTime) continue;

    const secs = parseSeconds(rawTime);
    if (!secs) continue;

    const lower = cells[2].toLowerCase();
    const entry = entryFor(login, day);
    entry.codes.set(lower, (entry.codes.get(lower) || 0) + secs);

    const total = codeTotals.get(lower) || { label: cells[2], secs: 0 };
    total.secs += secs;
    codeTotals.set(lower, total);
  }

  return { index, days, codeTotals };
}

// جدول مسطح: Agent / Date / Scheduled time
function parseFlatSchedule(rows) {
  const index = new Map();
  const days = new Set();

  rows.forEach(row => {
    const agent = findValue(row, FIELD_ALIASES.scheduleAgent);
    const d = dateKey(findValue(row, FIELD_ALIASES.scheduleDate));
    const raw = findValue(row, FIELD_ALIASES.scheduleDuration);
    const key = makeLookupKey(agent, d);
    if (!key || String(raw ?? '').trim() === '') return;

    if (!index.has(key)) index.set(key, { total: 0, hasTotal: true, codes: new Map() });
    index.get(key).total += parseSeconds(raw);
    days.add(d);
  });

  return { index, days, codeTotals: new Map() };
}

function parseScheduleSheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
  const positional = parseScheduleMatrix(matrix);
  if (positional.index.size) return positional;

  return parseFlatSchedule(XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false }));
}

async function loadSchedule(file) {
  const workbook = await readWorkbook(file);
  for (const name of getOrderedSheetNames(workbook, SCHEDULE_SHEET_HINTS)) {
    const parsed = parseScheduleSheet(workbook.Sheets[name]);
    if (parsed.index.size) return parsed;
  }
  return { index: new Map(), days: new Set(), codeTotals: new Map() };
}

async function getSchedule(file) {
  if (scheduleCache.file === file && scheduleCache.parsed) return scheduleCache.parsed;
  const parsed = await loadSchedule(file);
  scheduleCache = { file, parsed };
  return parsed;
}

async function prepareSchedule(file) {
  const status = document.getElementById('statusText');
  try {
    await ensureConfig();
    status.textContent = 'جاري قراءة Schedule واستخراج الأكواد...';
    const parsed = await getSchedule(file);
    scheduleCodeTotals = parsed.codeTotals;
    renderCodeChips();
    status.textContent = parsed.index.size
      ? `تم قراءة Schedule • ${scheduleCodeTotals.size} كود • اختار الأكواد اللي تتحسب في Tele-SCH`
      : 'تعذر قراءة Schedule - راجع شكل الملف';
  } catch (error) {
    console.error(error);
    status.textContent = 'تعذر قراءة Schedule';
  }
}

/* ─── اختيار الأكواد ───────────────────────────────────────────────────── */
function persistSelectedCodes() {
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(CODES_STORAGE_KEY, JSON.stringify([...selectedCodes]));
    }
  } catch (error) { /* مش مشكلة */ }
}

function renderCodeChips() {
  const panel = document.getElementById('codesPanel');
  const box = document.getElementById('codeChips');
  if (!box) return;
  if (panel) panel.hidden = false;

  box.innerHTML = '';
  const entries = [...scheduleCodeTotals.entries()].sort((a, b) => b[1].secs - a[1].secs);

  if (!entries.length) {
    box.innerHTML = '<span class="codes-empty">لا توجد أكواد داخل ملف Schedule</span>';
    return;
  }

  entries.forEach(([lower, info]) => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = `code-chip${selectedCodes.has(lower) ? ' selected' : ''}`;
    chip.innerHTML = `${info.label}<small>${formatTime(info.secs)}</small>`;
    chip.onclick = () => {
      if (selectedCodes.has(lower)) selectedCodes.delete(lower);
      else selectedCodes.add(lower);
      persistSelectedCodes();
      renderCodeChips();
    };
    box.appendChild(chip);
  });
}

function selectAllCodes() {
  selectedCodes = new Set(scheduleCodeTotals.keys());
  persistSelectedCodes();
  renderCodeChips();
}

function clearCodes() {
  selectedCodes = new Set();
  persistSelectedCodes();
  renderCodeChips();
}

// Tele-SCH بالثواني: لو في أكواد مختارة نجمعها، غير كده نستخدم إجمالي اليوم
function scheduleSecondsFor(entry, codes) {
  if (!entry) return null;
  if (codes.size) {
    let sum = 0;
    codes.forEach(code => { sum += entry.codes.get(code) || 0; });
    return sum;
  }
  if (entry.hasTotal) return entry.total;
  let sum = 0;
  entry.codes.forEach(v => { sum += v; });
  return sum;
}

/* ─── حساب المصفوفة (دالة نقية قابلة للاختبار) ──────────────────────────── */
function getDatesFromRows(rows, aliases) {
  const dates = new Set();
  rows.forEach(row => {
    const key = dateKey(findValue(row, aliases));
    if (key) dates.add(key);
  });
  return dates;
}

function getStructureDates(rows) {
  const dates = getDatesFromRows(rows, FIELD_ALIASES.structureDate);
  const year = [...dates][0]?.slice(0, 4) || String(inferredYear);

  rows.forEach(row => {
    Object.keys(row).forEach(key => {
      const m = key.match(/^(\d{1,2})[-\/](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$/i);
      if (m) dates.add(formatDateParts(year, MONTHS[m[2].toLowerCase()], +m[1]));
    });
  });
  return dates;
}

function getPreferredDateGroups(schedule) {
  const utlDates = [...getDatesFromRows(sourceRows.utl, FIELD_ALIASES.utlDate)].sort();
  if (utlDates.length) return utlDates;

  const all = new Set([
    ...getDatesFromRows(sourceRows.ir, FIELD_ALIASES.irDate),
    ...getDatesFromRows(sourceRows.comp, FIELD_ALIASES.compDate),
    ...getStructureDates(sourceRows.structure),
    ...(schedule?.days || [])
  ]);
  return [...all].sort();
}

function buildMatrix({ structureRows, utlRows, irRows, compRows, schedule, days, codes, config }) {
  const cfg = { ...DEFAULT_CONFIG, ...(config || {}) };
  const ir = buildIRIndexes(irRows);
  const talkIndex = buildTalkTimeIndex(utlRows);
  const structureDuration = buildSumIndex(
    structureRows, FIELD_ALIASES.structureId, FIELD_ALIASES.structureDate, FIELD_ALIASES.structureDuration
  );
  const compIndex = buildSumIndex(
    compRows, FIELD_ALIASES.compId, FIELD_ALIASES.compDate, FIELD_ALIASES.compDuration
  );
  const scheduleIndex = schedule?.index || new Map();

  return structureRows
    .map(row => ({
      teleoptiId: findValue(row, FIELD_ALIASES.structureId),
      loginId: findValue(row, FIELD_ALIASES.loginId),
      perm: findValue(row, FIELD_ALIASES.perm),
      ttsUser: findValue(row, FIELD_ALIASES.ttsUser),
      bssUser: findValue(row, FIELD_ALIASES.bssUser),
      group: findValue(row, FIELD_ALIASES.group),
      agentName: findValue(row, FIELD_ALIASES.agentName),
      status: findValue(row, FIELD_ALIASES.status, 'Active'),
      tlId: findValue(row, FIELD_ALIASES.tlId),
      tlName: findValue(row, FIELD_ALIASES.tlName)
    }))
    .filter(agent => String(agent.agentName).trim() || String(agent.loginId).trim())
    .map(agent => {
      const isActive = String(agent.status).trim().toLowerCase() === 'active';
      const scheduleCandidates = [agent.teleoptiId, agent.loginId, agent.agentName];
      const perDay = {};

      days.forEach(day => {
        const assigning = lookupFirst(ir.assigning, [agent.ttsUser], day);
        const tkt = lookupFirst(ir.tkt, [agent.ttsUser], day);
        const systemSeconds = Math.round(tkt * cfg.secondsPerTicket);
        const talkSeconds = lookupFirst(talkIndex, [agent.loginId], day);
        const compSeconds = lookupFirst(compIndex, [agent.teleoptiId, agent.loginId], day);

        let scheduleSeconds = null;
        for (const candidate of scheduleCandidates) {
          const entry = scheduleIndex.get(makeLookupKey(candidate, day));
          if (entry) { scheduleSeconds = scheduleSecondsFor(entry, codes); break; }
        }
        const teleSchSeconds = scheduleSeconds ?? lookupFirst(structureDuration, [agent.teleoptiId], day);

        const lossSeconds = Math.max(
          0,
          teleSchSeconds * cfg.teleSchFactor - (systemSeconds + talkSeconds + compSeconds)
        );

        perDay[day] = {
          assigning,
          tkt,
          system: formatTime(systemSeconds),
          talkTime: formatTime(talkSeconds),
          teleSch: formatTime(teleSchSeconds),
          comp: formatTime(compSeconds),
          lossTime: isActive ? formatTime(lossSeconds) : agent.status
        };
      });

      return { ...agent, days: perDay };
    });
}

function collapseAllDateGroups() {
  collapsedDays = {};
  dateGroups.forEach(day => { collapsedDays[day] = true; });
}

async function processData() {
  const progress = document.getElementById('progressBar');
  const status = document.getElementById('statusText');

  try {
    await ensureConfig();
    progress.style.width = '8%';
    status.textContent = 'جاري تحميل Structure...';

    const structureWorkbook = filesState.struct
      ? await readWorkbook(filesState.struct)
      : await readBundledStructure();
    sourceRows.structure = chooseSheet(
      structureWorkbook,
      ['structure', 'str', 'loss', 'master', 'sep', 'updated'],
      { validate: hasAgentNameColumn }
    );
    sourceLabels.structure = filesState.struct ? filesState.struct.name : 'STR Loss.xlsx من المستودع';

    if (!sourceRows.structure.length) {
      throw new Error('ملف Structure فاضي أو مفيهوش عمود Agent Name');
    }
    progress.style.width = '20%';

    let schedule = null;
    if (filesState.schedule) {
      status.textContent = 'جاري قراءة Schedule...';
      schedule = await getSchedule(filesState.schedule);
      scheduleCodeTotals = schedule.codeTotals;
      sourceLabels.schedule = filesState.schedule.name;
    } else {
      sourceLabels.schedule = 'لم يتم رفع Schedule';
    }
    progress.style.width = '35%';

    const loadRows = async (file, hints) => (file ? chooseSheet(await readWorkbook(file), hints) : []);
    sourceRows.utl = await loadRows(filesState.utl, ['utl', 'log']);
    progress.style.width = '50%';
    sourceRows.ir = await loadRows(filesState.ir, ['ir', 'ticket']);
    progress.style.width = '62%';
    sourceRows.comp = await loadRows(filesState.comp, ['comp', 'compensation']);

    dateGroups = getPreferredDateGroups(schedule);
    if (!dateGroups.length) throw new Error('لم يتم العثور على أي تاريخ داخل الشيتات');

    status.textContent = `جاري حساب ${dateGroups.length} يوم...`;
    progress.style.width = '80%';

    processedMatrixData = buildMatrix({
      structureRows: sourceRows.structure,
      utlRows: sourceRows.utl,
      irRows: sourceRows.ir,
      compRows: sourceRows.comp,
      schedule,
      days: dateGroups,
      codes: selectedCodes,
      config: CONFIG
    });

    collapseAllDateGroups();
    progress.style.width = '100%';
    status.textContent =
      `تم التحديث • ${processedMatrixData.length} موظف • ${dateGroups.length} يوم • ` +
      `UTL ${sourceRows.utl.length} • IR ${sourceRows.ir.length} • Comp ${sourceRows.comp.length} • ` +
      `Schedule: ${sourceLabels.schedule}` +
      (selectedCodes.size ? ` (${selectedCodes.size} كود)` : '');

    buildGroupToggles();
    renderMatrixTable(processedMatrixData);
  } catch (error) {
    console.error(error);
    progress.style.width = '0%';
    status.textContent = 'حدث خطأ أثناء المعالجة';
    alert(error.message);
  }
}

/* ─── العرض ────────────────────────────────────────────────────────────── */
function getActiveColumns(rows) {
  return FIXED_COLUMNS.filter(col => col.always || rows.some(r => String(r[col.key] ?? '').trim() !== ''));
}

function getVisibleDayColumnCount() {
  return dateGroups.reduce((t, day) => t + (collapsedDays[day] ? 1 : DAY_METRIC_LABELS.length), 0);
}

function getEmptyStateColspan() {
  return activeColumns.length + getVisibleDayColumnCount();
}

function buildGroupToggles() {
  const container = document.getElementById('groupToggles');
  container.innerHTML = '<span class="control-label">عرض/طي الأيام:</span>';

  dateGroups.forEach(day => {
    const button = document.createElement('button');
    button.className = `day-btn${collapsedDays[day] ? ' collapsed' : ''}`;
    button.textContent = displayDate(day);
    button.onclick = () => {
      collapsedDays[day] = !collapsedDays[day];
      buildGroupToggles();
      renderMatrixTable(processedMatrixData);
    };
    container.appendChild(button);
  });
}

function renderMatrixTable(rows) {
  const head = document.getElementById('tableHead');
  const body = document.getElementById('tableBody');
  visibleMatrixData = [...rows];

  const basis = processedMatrixData.length ? processedMatrixData : rows;
  if (basis.length) activeColumns = getActiveColumns(basis);

  document.getElementById('rowCount').textContent = `عدد الموظفين: ${rows.length}`;
  head.innerHTML = '';
  body.innerHTML = '';

  if (!rows.length) {
    body.innerHTML = `
      <tr>
        <td colspan="${getEmptyStateColspan()}" class="empty-state">لا توجد بيانات للعرض</td>
      </tr>`;
    return;
  }

  const firstHeader = document.createElement('tr');

  activeColumns.forEach(col => {
    const th = document.createElement('th');
    th.rowSpan = 2;
    th.className = 'th-base';
    th.textContent = col.label;
    firstHeader.appendChild(th);
  });

  dateGroups.forEach(day => {
    const isCollapsed = Boolean(collapsedDays[day]);
    const th = document.createElement('th');
    th.colSpan = isCollapsed ? 1 : DAY_METRIC_LABELS.length;
    th.rowSpan = isCollapsed ? 2 : 1;
    th.className = 'th-date-group';

    const content = document.createElement('span');
    content.className = 'th-date-content';

    const dateText = document.createElement('span');
    dateText.textContent = displayDate(day);

    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'th-day-toggle';
    toggle.setAttribute('aria-label', `${isCollapsed ? 'توسيع' : 'طي'} يوم ${displayDate(day)}`);
    toggle.innerHTML = `<i class="fa-solid ${isCollapsed ? 'fa-chevron-left' : 'fa-chevron-down'}"></i>`;
    toggle.onclick = event => {
      event.stopPropagation();
      collapsedDays[day] = !collapsedDays[day];
      buildGroupToggles();
      renderMatrixTable(processedMatrixData);
    };

    content.appendChild(dateText);
    content.appendChild(toggle);
    th.appendChild(content);
    firstHeader.appendChild(th);
  });
  head.appendChild(firstHeader);

  const secondHeader = document.createElement('tr');
  dateGroups.forEach(day => {
    if (collapsedDays[day]) return;
    DAY_METRIC_LABELS.forEach((label, index) => {
      const th = document.createElement('th');
      th.className = index === 0 ? 'th-sub-orange' : 'th-sub-purple';
      th.textContent = label;
      secondHeader.appendChild(th);
    });
  });
  head.appendChild(secondHeader);

  rows.forEach(row => {
    const tr = document.createElement('tr');

    activeColumns.forEach(col => {
      const td = document.createElement('td');
      td.textContent = row[col.key] ?? '';
      tr.appendChild(td);
    });

    dateGroups.forEach(day => {
      if (collapsedDays[day]) {
        const td = document.createElement('td');
        td.className = 'day-collapsed-cell';
        td.textContent = '—';
        tr.appendChild(td);
        return;
      }

      const v = row.days[day] || {};
      const lossText = String(v.lossTime ?? '0:00:00');
      const isTime = /^\d+:\d{2}:\d{2}$/.test(lossText);
      const lossClass = !isTime ? 'cell-unpaid' : lossText === '0:00:00' ? 'cell-zero-loss' : 'cell-loss';

      [v.assigning ?? 0, v.tkt ?? 0, v.system ?? '0:00:00', v.talkTime ?? '0:00:00',
        v.teleSch ?? '0:00:00', v.comp ?? '0:00:00', lossText
      ].forEach((value, index) => {
        const td = document.createElement('td');
        td.textContent = value;
        td.classList.add('day-metric-cell');
        if (index === DAY_METRIC_LABELS.length - 1) td.classList.add(lossClass);
        tr.appendChild(td);
      });
    });

    body.appendChild(tr);
  });
}

function filterData() {
  const query = document.getElementById('searchInput').value.trim().toLowerCase();

  if (!query) {
    renderMatrixTable(processedMatrixData);
    return;
  }

  renderMatrixTable(processedMatrixData.filter(row =>
    [row.agentName, row.loginId, row.teleoptiId, row.ttsUser, row.tlName]
      .some(value => String(value ?? '').toLowerCase().includes(query))
  ));
}

/* ─── التصدير (كل الأيام، بغض النظر عن الطي) ───────────────────────────── */
function buildExportRows(rows) {
  return rows.map(row => {
    const result = {};
    activeColumns.forEach(col => { result[col.label] = row[col.key]; });

    dateGroups.forEach(day => {
      const v = row.days[day] || {};
      const label = displayDate(day);
      result[`${label} - Assigning Tkts`] = v.assigning;
      result[`${label} - TKT`] = v.tkt;
      result[`${label} - System`] = v.system;
      result[`${label} - Talk Time`] = v.talkTime;
      result[`${label} - Tele-SCH`] = v.teleSch;
      result[`${label} - Comp`] = v.comp;
      result[`${label} - Loss Time`] = v.lossTime;
    });
    return result;
  });
}

function currentExportRows() {
  return buildExportRows(visibleMatrixData.length ? visibleMatrixData : processedMatrixData);
}

function exportToExcel() {
  const exportRows = currentExportRows();
  if (!exportRows.length) { alert('لا توجد بيانات للتصدير'); return; }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(exportRows), 'Matrix_Report');
  XLSX.writeFile(workbook, 'CallCenter_Daily_Performance_Report.xlsx');
}

function exportToCSV() {
  const exportRows = currentExportRows();
  if (!exportRows.length) { alert('لا توجد بيانات للتصدير'); return; }

  const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(exportRows));
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'CallCenter_Daily_Performance_Report.csv';
  link.click();
  URL.revokeObjectURL(url);
}
