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
     Tele-SCH       = مجموع Duration للأكواد المختارة من Schedule × 90%  (بيظهر بعد الخصم)
     Comp           = من Compensation
     Loss Time      = Tele-SCH  −  (System + Talk Time + Comp)   (لا يقل عن صفر)
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
let scheduleDebugText = '';
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
  compUser: ['User Name', 'Username', 'Comp_User', 'TTS User'],
  compId: ['ID', 'Comp_ID', 'Comp ID', 'Teleopti ID', 'ST_ID'],
  compDate: ['Date', 'Comp_Da'],
  compDuration: ['Code Time', 'Comp_Du', 'Comp Duration'],
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
function serialToDateKey(serial) {
  if (!(serial > 20000 && serial < 80000)) return '';
  const d = new Date(Date.UTC(1899, 11, 30 + Math.floor(serial)));
  return formatDateParts(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}

function dateKey(value) {
  if (value === null || value === undefined || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatLocalDate(value);

  if (typeof value === 'number') return serialToDateKey(value);

  const text = String(value).trim();
  if (!text) return '';

  // رقم تاريخ إكسيل (مثلاً 46277.0736) لما الخلية تتقري كنص
  if (/^\d{5}(?:\.\d+)?$/.test(text)) return serialToDateKey(Number(text));

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

  m = text.match(/^(\d{1,2})[\s\-\/]([A-Za-z]{3})[A-Za-z]*(?:[\s\-\/,]+(\d{2,4}))?(?:\s.*)?$/);
  if (m) {
    const month = MONTHS[m[2].toLowerCase()];
    if (!month) return '';
    let year = m[3] ? +m[3] : inferredYear;
    if (year < 100) year += 2000;
    return formatDateParts(year, month, +m[1]);
  }

  // تواريخ جوه نص (مثلاً "Tue 9/1/2026" أو "Tuesday, September 01, 2026")
  m = text.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})\b/);
  if (m) return dateKey(`${m[1]}/${m[2]}/${m[3]}`);

  const monthPattern = '(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|Sept?(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)';
  m = text.match(new RegExp(`\\b${monthPattern}\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})\\b`, 'i'));
  if (m) return formatDateParts(+m[3], MONTHS[m[1].slice(0, 3).toLowerCase()], +m[2]);
  m = text.match(new RegExp(`\\b(\\d{1,2})(?:st|nd|rd|th)?\\s+${monthPattern}\\.?,?\\s+(\\d{4})\\b`, 'i'));
  if (m) return formatDateParts(+m[3], MONTHS[m[2].slice(0, 3).toLowerCase()], +m[1]);

  return '';
}

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function displayDate(value) {
  if (!value) return '';
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return `${MONTH_NAMES[date.getMonth()]}-${String(date.getDate()).padStart(2, '0')}`;
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

// يدوّر على صف العناوين (أول 40 صف) اللي فيه كل مجموعات الأعمدة المطلوبة
function rowsByHeader(sheet, requiredGroups) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
  const limit = Math.min(matrix.length, 40);

  for (let i = 0; i < limit; i += 1) {
    const cells = matrix[i].map(normalise);
    const ok = requiredGroups.every(group => group.some(alias => cells.includes(normalise(alias))));
    if (!ok) continue;

    const seen = new Map();
    const headers = matrix[i].map((h, idx) => {
      const base = String(h ?? '').trim() || `Column ${idx + 1}`;
      const n = seen.get(base) || 0;
      seen.set(base, n + 1);
      return n ? `${base} ${n + 1}` : base;
    });

    return matrix.slice(i + 1)
      .filter(row => row.some(v => String(v ?? '').trim() !== ''))
      .map(row => Object.fromEntries(headers.map((h, idx) => [h, row[idx] ?? ''])));
  }
  return null;
}

// يرجّع صفوف أول شيت صالح (يفضّل الشيتات اللي اسمها قريب من hints)
function chooseSheet(workbook, hints, options = {}) {
  const { validate, requiredGroups } = options;
  const names = getOrderedSheetNames(workbook, hints);

  if (requiredGroups) {
    for (const name of names) {
      const rows = rowsByHeader(workbook.Sheets[name], requiredGroups);
      if (rows && rows.length) return rows;
    }
  }

  for (const name of names) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '', raw: false });
    if (!rows.length) continue;
    if (!validate || validate(rows)) return rows;
  }
  return [];
}

const HEADER_GROUPS = {
  ir: [[...FIELD_ALIASES.irAssigned, ...FIELD_ALIASES.irAdded], ['added_on']],
  utl: [FIELD_ALIASES.utlUser, FIELD_ALIASES.utlDate],
  comp: [[...FIELD_ALIASES.compUser, ...FIELD_ALIASES.compId], FIELD_ALIASES.compDuration]
};

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

    const initial = Array.isArray(saved) && saved.length ? saved : CONFIG.scheduleCodes || [];
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
function userId(value) {
  return String(value ?? '').split('@')[0];
}

// تشخيص: هل مستخدمين IR بيطابقوا TTS User في الـ Structure؟ وهل التواريخ اتقرت؟
function diagnoseIR(irRows, structureRows) {
  const result = { rows: irRows.length, addedMatched: 0, assignedMatched: 0, datesRead: 0, unmatched: [], first: '', last: '' };
  if (!irRows.length) return result;

  const tts = new Set(structureRows.map(r => normalise(userId(findValue(r, FIELD_ALIASES.ttsUser)))).filter(Boolean));
  const keys = Object.keys(irRows[0]);
  const colAssigned = findKey(irRows[0], FIELD_ALIASES.irAssigned) ?? keys[IR_COL.assigned];
  const colAdded = findKey(irRows[0], FIELD_ALIASES.irAdded) ?? keys[IR_COL.added];
  const colDate = findKey(irRows[0], FIELD_ALIASES.irDate) ?? keys[IR_COL.date];
  const unmatched = new Set();
  const dates = [];

  irRows.forEach(row => {
    const a = normalise(userId(row[colAssigned]));
    const b = normalise(userId(row[colAdded]));
    if (a && tts.has(a)) result.assignedMatched += 1;
    if (b && tts.has(b)) result.addedMatched += 1;
    else if (b && unmatched.size < 5) unmatched.add(String(row[colAdded]));
    const d = dateKey(row[colDate]);
    if (d) { result.datesRead += 1; dates.push(d); }
  });

  dates.sort();
  result.first = dates[0] || '';
  result.last = dates[dates.length - 1] || '';
  result.unmatched = [...unmatched];
  result.columns = { assigned: colAssigned, added: colAdded, date: colDate };
  result.sampleDate = `${String(irRows[0][colDate])} (${typeof irRows[0][colDate]}) • عمود: ${colDate}`;
  return result;
}

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
    addToIndex(assigning, userId(row[colAssigned]), day, 1);
    addToIndex(tkt, userId(row[colAdded]), day, 1);
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
// يحدد عمود Scheduled time وأعمدة العناوين (Label) من صف عناوين التقرير الخام
function detectScheduleLayout(matrix) {
  for (const row of matrix.slice(0, 80)) {
    const cells = row.map(normalise);
    const duration = cells.findIndex(c => /^scheduledtime(hhmm)?$/.test(c)); // مش عنوان التقرير 'Scheduled Time per Agent'
    if (duration < 0) continue;
    const firstValue = cells.findIndex(c => /^(contracttime|worktime|paidtime)/.test(c));
    return { duration, labelEnd: firstValue > 1 ? firstValue : Math.min(duration, 4) };
  }
  return { duration: SCHEDULE_COL.duration, labelEnd: 4 };
}

// يقرأ التقرير الخام (RD) بنفس منطق الماكرو:
//   B = "ID + اسم" يبدأ بلوك موظف (أي نص تاني في B = فريق/إجمالي → بنفصل البلوك)
//   C = تاريخ (صف يوم) أو كود نشاط | Scheduled time = Duration
function parseScheduleMatrix(matrix) {
  const index = new Map();
  const days = new Set();
  const codeTotals = new Map();
  const agents = new Set();
  const layout = detectScheduleLayout(matrix);
  let login = '';
  let day = '';
  let agentDepth = -1;

  const entryFor = (id, d) => {
    const key = makeLookupKey(id, d);
    if (!index.has(key)) index.set(key, { total: 0, hasTotal: false, codes: new Map() });
    return index.get(key);
  };

  for (const row of matrix) {
    // أول خلية فيها نص في أعمدة العناوين = مستوى الصف (فريق / موظف / يوم / كود)
    let depth = -1;
    let label = '';
    let labelValue = '';
    for (let i = 0; i < layout.labelEnd; i += 1) {
      const text = String(row[i] ?? '').trim();
      if (text) { depth = i; label = text; labelValue = row[i]; break; }
    }
    if (depth < 0) continue;

    if (TOTALS_WORDS.some(w => normalise(label) === normalise(w))) {
      login = '';
      day = '';
      agentDepth = -1;
      continue;
    }

    const labelDay = dateKey(labelValue);
    const idMatch = !labelDay && label.match(/\b(\d{5,6})\b/);

    if (idMatch) {                       // صف موظف: "156958 Hesham ... 86466"
      login = idMatch[1];
      agentDepth = depth;
      day = '';
      agents.add(login);
      continue;
    }

    if (!login || depth <= agentDepth) { // صف فريق أو عنوان → بنقفل بلوك الموظف
      login = '';
      day = '';
      agentDepth = -1;
      continue;
    }

    const rawTime = row[layout.duration];
    const hasTime = String(rawTime ?? '').trim() !== '';

    if (labelDay) {
      day = labelDay;
      days.add(labelDay);
      const entry = entryFor(login, labelDay);
      if (hasTime) {
        entry.total += parseSeconds(rawTime);
        entry.hasTotal = true;
      }
      continue;
    }

    if (!day || !hasTime) continue;

    const secs = parseSeconds(rawTime);
    if (!secs) continue;

    const lower = label.toLowerCase();
    const entry = entryFor(login, day);
    entry.codes.set(lower, (entry.codes.get(lower) || 0) + secs);

    const total = codeTotals.get(lower) || { label, secs: 0 };
    total.secs += secs;
    codeTotals.set(lower, total);
  }

  return { index, days, codeTotals, agents };
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

// شيت Final بعد الماكرو: ID | Date | Duration | Agent Name | TL | Code
function parseFinalScheduleMatrix(matrix) {
  const limit = Math.min(matrix.length, 40);

  for (let i = 0; i < limit; i += 1) {
    const cells = matrix[i].map(normalise);
    const col = name => cells.indexOf(normalise(name));
    const idCol = col('ID');
    const dateCol = col('Date');
    const durCol = col('Duration');
    const codeCol = col('Code');
    if (idCol < 0 || dateCol < 0 || durCol < 0) continue;

    const index = new Map();
    const days = new Set();
    const codeTotals = new Map();

    matrix.slice(i + 1).forEach(row => {
      const id = String(row[idCol] ?? '').trim();
      const day = dateKey(row[dateCol]);
      if (!id || !day) return;

      const secs = parseSeconds(row[durCol]);
      const key = makeLookupKey(id, day);
      if (!index.has(key)) index.set(key, { total: 0, hasTotal: false, codes: new Map() });
      const entry = index.get(key);
      days.add(day);

      const code = codeCol >= 0 ? String(row[codeCol] ?? '').trim() : '';
      if (!code) {
        entry.total += secs;
        entry.hasTotal = true;
        return;
      }
      const lower = code.toLowerCase();
      entry.codes.set(lower, (entry.codes.get(lower) || 0) + secs);
      const total = codeTotals.get(lower) || { label: code, secs: 0 };
      total.secs += secs;
      codeTotals.set(lower, total);
    });

    if (index.size) return { index, days, codeTotals };
  }
  return null;
}

function parseScheduleSheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
  const finalFormat = parseFinalScheduleMatrix(matrix);
  if (finalFormat) return finalFormat;

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
  return { index: new Map(), days: new Set(), codeTotals: new Map(), agents: new Set(), debug: describeSchedule(workbook) };
}

// لو الملف مش مفهوم: نعرض أول صفوف فيها نص عشان نعرف شكله
function describeSchedule(workbook) {
  try {
    const name = workbook.SheetNames[0];
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: true });
    const lines = [];
    matrix.forEach((row, r) => {
      if (lines.length >= 14) return;
      const cells = row.map((v, c) => [c, String(v ?? '').trim()]).filter(([, t]) => t).slice(0, 6);
      if (cells.length) lines.push(`ص${r + 1}: ` + cells.map(([c, t]) => `[${c}] ${t.slice(0, 28)}`).join(' | '));
    });
    return `شيت "${name}" (${matrix.length} صف) — ${lines.join(' ▸ ')}`;
  } catch (error) {
    return '';
  }
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
    scheduleDebugText = parsed.debug || '';
    renderCodeChips();
    status.textContent = parsed.index.size
      ? `تم قراءة Schedule • ${parsed.agents ? parsed.agents.size + ' موظف • ' : ''}${parsed.days.size} يوم • ${scheduleCodeTotals.size} كود • اختار الأكواد اللي تتحسب في Tele-SCH`
      : 'تعذر قراءة Schedule - مفيش صفوف موظفين (ID من 5-6 أرقام في عمود B) ولا أعمدة ID/Date/Duration';
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
    box.innerHTML = '';
    const note = document.createElement('span');
    note.className = 'codes-empty';
    note.textContent = 'لا توجد أكواد داخل ملف Schedule' + (scheduleDebugText ? ` — شكل الملف: ${scheduleDebugText}` : '');
    box.appendChild(note);
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
  // Compensation بيتربط بـ User Name (= TTS User) أولاً، وبعدين بالـ ID (= Teleopti ID)
  const compByUser = buildSumIndex(
    compRows, FIELD_ALIASES.compUser, FIELD_ALIASES.compDate, FIELD_ALIASES.compDuration
  );
  const compById = buildSumIndex(
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
        const ttsKey = userId(agent.ttsUser);
        const assigning = lookupFirst(ir.assigning, [ttsKey], day);
        const tkt = lookupFirst(ir.tkt, [ttsKey], day);
        const systemSeconds = Math.round(tkt * cfg.secondsPerTicket);
        const talkSeconds = lookupFirst(talkIndex, [agent.loginId], day);
        const compKey = makeLookupKey(ttsKey, day);
        const compSeconds = compKey && compByUser.has(compKey)
          ? compByUser.get(compKey)
          : lookupFirst(compById, [agent.teleoptiId], day);

        let scheduleSeconds = null;
        for (const candidate of scheduleCandidates) {
          const entry = scheduleIndex.get(makeLookupKey(candidate, day));
          if (entry) { scheduleSeconds = scheduleSecondsFor(entry, codes); break; }
        }
        const teleSchSeconds = scheduleSeconds ?? lookupFirst(structureDuration, [agent.teleoptiId], day);

        // نفس معادلتك: Assigning = IF(Tele-SCH=0, 0, COUNTIFS(...)) — لما فيه Schedule مرفوع
        const assigningShown = scheduleIndex.size > 0 && teleSchSeconds === 0 ? 0 : assigning;
        // Tele-SCH بيظهر بعد خصم الـ 90% (7:12:00 بدل 8:00:00)
        const teleSchAfterFactor = Math.round(teleSchSeconds * cfg.teleSchFactor);
        const lossSeconds = Math.max(0, teleSchAfterFactor - (systemSeconds + talkSeconds + compSeconds));

        perDay[day] = {
          assigning: assigningShown,
          tkt,
          system: formatTime(systemSeconds),
          talkTime: formatTime(talkSeconds),
          teleSch: formatTime(teleSchAfterFactor),
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

    const loadRows = async (file, hints, groups) =>
      (file ? chooseSheet(await readWorkbook(file), hints, { requiredGroups: groups }) : []);
    sourceRows.utl = await loadRows(filesState.utl, ['utl', 'log'], HEADER_GROUPS.utl);
    progress.style.width = '50%';
    sourceRows.ir = await loadRows(filesState.ir, ['ir', 'ticket'], HEADER_GROUPS.ir);
    progress.style.width = '62%';
    sourceRows.comp = await loadRows(filesState.comp, ['comp', 'compensation'], HEADER_GROUPS.comp);

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

    const diag = diagnoseIR(sourceRows.ir, sourceRows.structure);
    if (sourceRows.ir.length) console.log('IR diagnostics', diag);

    let irText = filesState.ir ? '' : ' • IR: لم يتم رفع الملف';
    if (filesState.ir) {
      irText = ` • IR ${diag.rows} صف (TKT مطابق ${diag.addedMatched} | Assigning مطابق ${diag.assignedMatched} | تواريخ ${diag.datesRead}/${diag.rows}`;
      if (diag.first) irText += ` | ${diag.first} → ${diag.last}`;
      irText += ')';
      if (!diag.addedMatched && diag.unmatched.length) irText += ` ⚠ أمثلة added_by مش مطابقة: ${diag.unmatched.join(', ')}`;
      if (diag.rows && !diag.datesRead) irText += ` ⚠ مفيش تاريخ اتقرا من added_on — أول قيمة: ${diag.sampleDate}`;
    }

    if (diag.last && collapsedDays[diag.last] !== undefined) collapsedDays[diag.last] = false; // نفتح آخر يوم فيه IR
    const scheduleText = schedule ? `${schedule.index.size} يوم-موظف` : sourceLabels.schedule;

    status.textContent =
      `تم التحديث • ${processedMatrixData.length} موظف • ${dateGroups.length} يوم • ` +
      `UTL ${sourceRows.utl.length} • Comp ${sourceRows.comp.length}${irText} • ` +
      `Schedule: ${scheduleText}` +
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

/* ─── التصدير ──────────────────────────────────────────────────────────── */
// نفس شكل شيت الإكسيل: الأعمدة الثابتة ثم 7 أعمدة لكل يوم، واليوم المطوي بيفضل ظاهر منه Loss Time بس
const EXPORT_FIXED = [
  { key: 'teleoptiId', label: 'Teleopti ID', width: 12 },
  { key: 'loginId', label: 'Login ID', width: 11 },
  { key: 'agentName', label: 'Agent Name', width: 38 },
  { key: 'ttsUser', label: 'TTS User', width: 22 },
  { key: 'tlName', label: 'TL Name', width: 18 },
  { key: 'status', label: 'Status', width: 12 }
];
const EXPORT_COLORS = {
  purple: 'FF7030A0',
  orange: 'FFE46C0A',
  white: 'FFFFFFFF',
  greenFill: 'FFC6EFCE',
  greenText: 'FF006100',
  redFill: 'FFFFC7CE',
  redText: 'FF9C0006'
};

function shortDate(day) {
  const d = new Date(`${day}T00:00:00`);
  if (Number.isNaN(d.getTime())) return day;
  return `${d.getDate()}-${MONTH_NAMES[d.getMonth()]}`;
}

function buildExportRows(rows) {
  return rows.map(row => {
    const result = {};
    EXPORT_FIXED.forEach(col => { result[col.label] = row[col.key]; });

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

function currentExportSource() {
  return visibleMatrixData.length ? visibleMatrixData : processedMatrixData;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

async function exportToExcel() {
  const source = currentExportSource();
  if (!source.length) { alert('لا توجد بيانات للتصدير'); return; }
  if (typeof ExcelJS === 'undefined') { alert('مكتبة ExcelJS لم يتم تحميلها - تأكد من الاتصال بالإنترنت'); return; }

  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const center = { horizontal: 'center', vertical: 'middle' };
  const fixedCount = EXPORT_FIXED.length;
  const metricCount = DAY_METRIC_LABELS.length;
  const timeFormat = '[h]:mm:ss';

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Matrix_Report', {
    views: [{ state: 'frozen', xSplit: fixedCount, ySplit: 4 }],
    properties: { outlineLevelCol: 1, outlineProperties: { summaryRight: true, summaryBelow: false } }
  });

  const styleHeader = (cell, fill) => {
    cell.font = { bold: true, color: { argb: EXPORT_COLORS.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    cell.alignment = center;
    cell.border = border;
  };

  EXPORT_FIXED.forEach((col, i) => {
    ws.getColumn(i + 1).width = col.width;
    const cell = ws.getCell(4, i + 1);
    cell.value = col.label;
    styleHeader(cell, EXPORT_COLORS.purple);
  });

  dateGroups.forEach((day, d) => {
    const start = fixedCount + d * metricCount + 1;
    ws.mergeCells(3, start, 3, start + metricCount - 1);
    const dateCell = ws.getCell(3, start);
    dateCell.value = shortDate(day);
    dateCell.font = { bold: true };
    dateCell.alignment = center;
    for (let c = start; c < start + metricCount; c += 1) ws.getCell(3, c).border = border;

    DAY_METRIC_LABELS.forEach((label, m) => {
      const col = ws.getColumn(start + m);
      col.width = m === 0 ? 14 : 11;
      const cell = ws.getCell(4, start + m);
      cell.value = label;
      styleHeader(cell, m === 0 ? EXPORT_COLORS.orange : EXPORT_COLORS.purple);

      // أول 6 أعمدة في اليوم = مجموعة (Outline) قابلة للطي، وLoss Time هو عمود الملخص
      if (m < metricCount - 1) {
        col.outlineLevel = 1;
        col.hidden = Boolean(collapsedDays[day]);
      }
    });
  });

  const toDays = text => parseSeconds(text) / 86400;

  source.forEach((row, r) => {
    const rowIndex = 5 + r;

    EXPORT_FIXED.forEach((col, i) => {
      const cell = ws.getCell(rowIndex, i + 1);
      const raw = row[col.key] ?? '';
      cell.value = /^\d+$/.test(String(raw)) ? Number(raw) : raw;
      cell.alignment = center;
      cell.border = border;
    });

    dateGroups.forEach((day, d) => {
      const v = row.days[day] || {};
      const start = fixedCount + d * metricCount + 1;
      const lossIsTime = /^\d+:\d{2}:\d{2}$/.test(String(v.lossTime ?? ''));
      const values = [
        v.assigning ?? 0,
        v.tkt ?? 0,
        toDays(v.system),
        toDays(v.talkTime),
        toDays(v.teleSch),
        toDays(v.comp),
        lossIsTime ? toDays(v.lossTime) : (v.lossTime ?? '')
      ];

      values.forEach((value, m) => {
        const cell = ws.getCell(rowIndex, start + m);
        cell.value = value;
        cell.alignment = center;
        cell.border = border;
        if (m >= 2 && (m < metricCount - 1 || lossIsTime)) cell.numFmt = timeFormat;

        if (m === metricCount - 1) {
          const bad = lossIsTime && value > 0;
          cell.font = { bold: true, color: { argb: bad ? EXPORT_COLORS.redText : EXPORT_COLORS.greenText } };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: bad ? EXPORT_COLORS.redFill : EXPORT_COLORS.greenFill }
          };
        }
      });
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    'CallCenter_Daily_Performance_Report.xlsx'
  );
}

function exportToCSV() {
  const exportRows = buildExportRows(currentExportSource());
  if (!exportRows.length) { alert('لا توجد بيانات للتصدير'); return; }

  const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(exportRows));
  downloadBlob(new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' }), 'CallCenter_Daily_Performance_Report.csv');
}
