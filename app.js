let rawStructureData = [];
let processedMatrixData = [];
let visibleMatrixData = [];
let dateGroups = [];
let collapsedDays = {};

const sourceRows = {
  structure: [],
  schedule: [],
  utl: [],
  ir: [],
  comp: []
};

const filesState = {
  struct: null,
  schedule: null,
  utl: null,
  ir: null,
  comp: null
};

const sourceLabels = {
  structure: 'STR Loss.xlsx من المستودع',
  schedule: 'لم يتم رفع Schedule - سيتم الاعتماد على Structure'
};

const FIELD_ALIASES = {
  structureId: ['Teleopti ID', 'Teleopti', 'ST_ID'],
  loginId: ['Login ID', 'Login', 'UL_lo', 'User', 'Username'],
  perm: ['Perm'],
  ttsUser: ['TTS User', 'TTS'],
  bssUser: ['BSS User', 'BSS'],
  group: ['Group'],
  agentName: ['Agent Name', 'Agent', 'Name'],
  status: ['Status'],
  tlId: ['TL ID', 'TL Id'],
  tlName: ['TL Name', 'Team Leader', 'TL'],
  irUser: ['added_by', 'IR_L_E', 'User', 'TTS User', 'TTS', 'Agent Name', 'Login ID', 'Created By'],
  irAssigned: ['assigned_to', 'assigned to', 'Assignee', 'TTS User', 'TTS', 'User', 'Login ID', 'Agent Name'],
  irDate: ['added_on', 'Date', 'Date/Time', 'Created', 'Created On', 'تاريخ'],
  utlUser: ['UL_lo', 'Login ID', 'Login', 'User'],
  utlDate: ['UL_Date', 'Date'],
  compId: ['Comp_ID', 'Comp ID', 'Teleopti ID', 'ST_ID', 'Login ID'],
  compDate: ['Comp_Da', 'Date', 'Date/Time'],
  compDuration: ['Comp_Du', 'Comp Duration', 'Duration', 'Time'],
  structureDate: ['ST_D', 'Date'],
  structureDuration: ['ST_Du', 'ST Duration', 'Duration', 'Tele-SCH', 'Tele_SCH', 'Scheduled time'],
  scheduleAgent: ['Agent', 'Agent Name', 'Employee', 'Employee Name'],
  scheduleDate: ['Date', 'Scheduled Date'],
  scheduleDuration: ['Scheduled time', 'Scheduled Time', 'Scheduled time (hh:mm:ss)', 'Scheduled-Time']
};

const BASE_COLUMN_LABELS = [
  'Teleopti ID',
  'Login ID',
  'TTS User',
  'Agent Name',
  'Status',
  'TL Name'
];
const DAY_METRIC_LABELS = [
  'Assigning Tkts',
  'TKT',
  'System',
  'Talk Time',
  'Tele-SCH',
  'Comp',
  'Loss Time'
];

['struct', 'schedule', 'utl', 'ir', 'comp'].forEach(key => {
  const input = document.getElementById(`file-${key}`);
  if (input) {
    input.addEventListener('change', event => {
      const file = event.target.files[0];
      if (!file) return;

      filesState[key] = file;

      const name = document.getElementById(`name-${key}`);
      const card = document.getElementById(`card-${key}`);

      if (name) name.textContent = file.name;
      if (card) card.dataset.ready = 'true';
    });
  }
});

function normalise(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u0600-\u06ff]+/gi, '');
}

function matchesNormalisedValue(left, right) {
  if (!left || !right) return false;
  return left === right || left.includes(right) || right.includes(left);
}

function matchesAnyAlias(value, aliases) {
  const current = normalise(value);
  return aliases.some(alias => matchesNormalisedValue(current, normalise(alias)));
}

function findValue(row, aliases, fallback = '') {
  if (!row) return fallback;

  const key = Object.keys(row).find(name => matchesAnyAlias(name, aliases));
  return key === undefined ? fallback : row[key];
}

function formatDateParts(year, month, day) {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function formatLocalDate(date) {
  return formatDateParts(
    date.getFullYear(),
    date.getMonth() + 1,
    date.getDate()
  );
}

function flexibleDateKey(value) {
  if (value === null || value === undefined || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatLocalDate(value);
  }

  if (typeof value === 'number' && window.XLSX?.SSF) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }

  let text = String(value).trim().replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ');
  if (!text) return '';

  let datePart = text.split(' ')[0];

  const d = new Date(datePart);
  if (!Number.isNaN(d.getTime())) {
    return formatLocalDate(d);
  }

  const fullD = new Date(text);
  if (!Number.isNaN(fullD.getTime())) {
    return formatLocalDate(fullD);
  }

  const match = datePart.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    return formatDateParts(year, Number(match[2]), Number(match[1]));
  }

  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    return formatDateParts(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));
  }

  return '';
}

function parseGroupedScheduleSeconds(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return parseSeconds(value);

  const text = String(value).trim();
  if (!text) return 0;

  const parts = text.split(':').map(Number);
  if (parts.length === 2) {
    return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60;
  }
  return parseSeconds(value);
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
    return value > 0 && value < 1 ? value * 86400 : value;
  }

  const text = String(value).trim();
  const lower = text.toLowerCase();

  if (!text || lower === 'unpaid' || lower === 'maternity' || lower === 'planned sick') {
    return 0;
  }

  const parts = text.split(':').map(Number);
  if (parts.length === 3) return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0);
  if (parts.length === 2) return (parts[0] || 0) * 60 + (parts[1] || 0);

  return Number(text) || 0;
}

function formatTime(seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) return '0:00:00';
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remaining = Math.floor(seconds % 60);
  return `${hours}:${String(minutes).padStart(2, '0')}:${String(remaining).padStart(2, '0')}`;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        resolve(XLSX.read(new Uint8Array(event.target.result), { type: 'array', cellDates: true, raw: false }));
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
    throw new Error('تعذر تحميل STR Loss.xlsx');
  }
  if (!response.ok) throw new Error('لم يتم العثور على STR Loss.xlsx');
  const buffer = await response.arrayBuffer();
  return XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true, raw: false });
}

function getOrderedSheetNames(workbook, words) {
  const sheetNames = workbook?.SheetNames || [];
  const preferred = [];
  const fallback = [];
  sheetNames.forEach(name => {
    if (words.some(word => matchesAnyAlias(name, [word]))) preferred.push(name);
    else fallback.push(name);
  });
  return [...preferred, ...fallback];
}

function matrixFromSheet(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
}

function detectHeaderRow(matrix, aliasGroups) {
  let bestIndex = -1;
  let bestScore = 0;

  matrix.forEach((row, index) => {
    const cells = row.map(cell => normalise(cell)).filter(Boolean);
    if (!cells.length) return;
    const score = aliasGroups.filter(aliases =>
      aliases.some(alias => cells.some(cell => matchesNormalisedValue(cell, normalise(alias))))
    ).length;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore === aliasGroups.length ? bestIndex : -1;
}

function rowsFromMatrix(matrix, headerRowIndex) {
  const rawHeaders = matrix[headerRowIndex] || [];
  const headerCounts = new Map();
  const headers = rawHeaders.map((value, index) => {
    const base = String(value ?? '').trim() || `Column ${index + 1}`;
    const seen = headerCounts.get(base) || 0;
    headerCounts.set(base, seen + 1);
    return seen ? `${base} ${seen + 1}` : base;
  });

  return matrix
    .slice(headerRowIndex + 1)
    .filter(row => row.some(value => String(value ?? '').trim() !== ''))
    .map(row => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row[index] ?? '';
      });
      return record;
    });
}

function parseSheetRows(sheet, options = {}) {
  const { requiredHeaderAliases, label = 'الشيت' } = options;
  if (!sheet) return [];

  const matrix = matrixFromSheet(sheet);
  let headerRowIndex = -1;

  if (requiredHeaderAliases) {
    headerRowIndex = detectHeaderRow(matrix, requiredHeaderAliases);
  }

  // Generic Header Detection Fallback for IR/Comp/UTL
  if (headerRowIndex === -1) {
    for (let i = 0; i < Math.min(25, matrix.length); i++) {
      const row = matrix[i];
      const strCount = row.filter(c => typeof c === 'string' && String(c).trim().length > 1).length;
      if (strCount >= 3) {
        const hasCommon = row.some(c => /date|user|login|agent|time|duration|id|ticket|assign|comp|name/i.test(String(c)));
        if (hasCommon) {
          headerRowIndex = i;
          break;
        }
      }
    }
  }

  if (headerRowIndex === -1) headerRowIndex = 0; // Absolute fallback

  return rowsFromMatrix(matrix, headerRowIndex);
}

function chooseSheet(workbook, words, options = {}) {
  const orderedNames = getOrderedSheetNames(workbook, words);
  let lastError = null;
  for (const name of orderedNames) {
    try {
      const rows = parseSheetRows(workbook.Sheets[name], { ...options, sheetName: name });
      if (rows.length || !options.requiredHeaderAliases) return rows;
    } catch (error) {
      lastError = error;
    }
  }
  if (lastError) throw lastError;
  return [];
}

function getDatesFromRows(rows, aliases) {
  const dates = new Set();
  rows.forEach(row => {
    const key = flexibleDateKey(findValue(row, aliases, ''));
    if (key) dates.add(key);
  });
  return dates;
}

function getStructureDates(rows) {
  const dates = getDatesFromRows(rows, FIELD_ALIASES.structureDate);
  const fallbackYear = [...dates][0]?.slice(0, 4) || String(new Date().getFullYear());

  rows.forEach(row => {
    Object.keys(row).forEach(key => {
      const match = key.match(/^(\d{1,2})[-\/](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$/i);
      if (match) {
        const parsed = new Date(`${match[1]} ${match[2]} ${fallbackYear}`);
        if (!Number.isNaN(parsed.getTime())) dates.add(formatLocalDate(parsed));
      }
    });
  });
  return dates;
}

function makeLookupKey(value, day) {
  const normalised = normalise(value);
  return normalised && day ? `${normalised}|${day}` : '';
}

function addToIndex(map, value, day, amount) {
  const key = makeLookupKey(value, day);
  if (!key) return;
  map.set(key, (map.get(key) || 0) + amount);
}

function buildCountIndex(rows, userAliases, dateAliases) {
  const map = new Map();
  rows.forEach(row => {
    let user = '';
    for (const alias of userAliases) {
      const val = findValue(row, [alias], '');
      if (String(val ?? '').trim() !== '') {
        user = val;
        break;
      }
    }
    if (!user) {
      user = findValue(row, ['TTS User', 'TTS', 'added_by', 'assigned_to', 'User', 'Login ID', 'Login', 'Agent Name', 'Agent', 'Created By'], '');
    }
    const day = flexibleDateKey(findValue(row, dateAliases, ''));
    addToIndex(map, user, day, 1);
  });
  return map;
}

function buildSumIndex(rows, userAliases, dateAliases, valueAliases) {
  const map = new Map();
  rows.forEach(row => {
    const user = findValue(row, userAliases, '');
    const day = flexibleDateKey(findValue(row, dateAliases, ''));
    const amount = parseSeconds(findValue(row, valueAliases, 0));
    addToIndex(map, user, day, amount);
  });
  return map;
}

function buildTalkTimeIndex(rows) {
  const map = new Map();
  rows.forEach(row => {
    const user = findValue(row, FIELD_ALIASES.utlUser, '');
    const day = flexibleDateKey(findValue(row, FIELD_ALIASES.utlDate, ''));
    const total =
      parseSeconds(findValue(row, ['Hold Time', 'HoldTime'], 0)) +
      parseSeconds(findValue(row, ['Other Time', 'OtherTime'], 0)) +
      parseSeconds(findValue(row, ['AUXOUTOFFTIME'], 0)) +
      parseSeconds(findValue(row, ['ACWOUTOFFTIME'], 0));
    addToIndex(map, user, day, total);
  });
  return map;
}

function getIndexedValue(map, value, day) {
  return map.get(makeLookupKey(value, day)) || 0;
}

function getPresentIndexedValueByCandidates(map, presenceSet, candidates, day) {
  for (const candidate of candidates) {
    const key = makeLookupKey(candidate, day);
    if (presenceSet.has(key)) return map.get(key) || 0;
  }
  return 0;
}

function hasPresentCandidate(presenceSet, candidates, day) {
  return candidates.some(candidate => presenceSet.has(makeLookupKey(candidate, day)));
}

// دالة متطورة لقراءة شيت Schedule أياً كان تنسيقه (مدمج في عمود A أو مفصول)
function buildScheduleIndexByPosition(sheet) {
  if (!sheet) return { sumMap: new Map(), presenceSet: new Set() };
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const sumMap = new Map();
  const presenceSet = new Set();

  let timeColIndex = -1;
  // البحث الديناميكي عن عمود Scheduled time
  for (let i = 0; i < Math.min(30, matrix.length); i++) {
    for (let j = 0; j < matrix[i].length; j++) {
      const cell = String(matrix[i][j] || '').trim().toLowerCase();
      if (cell.includes('scheduled time') || cell === 'scheduled_time') {
        if (!cell.includes('overtime')) {
          timeColIndex = j;
          break;
        }
      }
    }
    if (timeColIndex !== -1) break;
  }
  
  if (timeColIndex === -1) timeColIndex = 11; // العمود البديل L

  let currentLogin = '';
  let currentDate = '';
  let currentDaySecs = 0;

  const saveDay = () => {
    if (currentLogin && currentDate && currentDaySecs > 0) {
      addToIndex(sumMap, currentLogin, currentDate, currentDaySecs);
      presenceSet.add(makeLookupKey(currentLogin, currentDate));
    }
  };

  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i];
    const colA = String(row[0] ?? '').trim();
    const colB = String(row[1] ?? '').trim();
    const colC = String(row[2] ?? '').trim();
    
    // دمج محتوى الأعمدة الأولى لالتقاط البيانات سواء كانت مقسمة أو Grouped بـ Indentation
    const identityOrDateCol = colA || colB || colC;
    const timeVal = String(row[timeColIndex] ?? row[11] ?? row[9] ?? '').trim();

    if (!identityOrDateCol) continue;
    const lowerId = identityOrDateCol.toLowerCase();

    if (!lowerId.includes('totals:') && !lowerId.includes('welcome call') && !lowerId.includes('site:')) {
      const day = flexibleDateKey(identityOrDateCol);
      
      if (day && day.includes('-')) {
        saveDay();
        currentDate = day;
        currentDaySecs = parseGroupedScheduleSeconds(timeVal);
      } else {
        const loginMatches = identityOrDateCol.match(/\b\d{3,6}\b/g);
        if (loginMatches && loginMatches.length >= 1 && /[a-zA-Z]/.test(identityOrDateCol)) {
          saveDay();
          currentLogin = loginMatches[loginMatches.length - 1];
          currentDate = '';
          currentDaySecs = 0;
        } else if (currentDate) {
          if (currentDaySecs === 0 && timeVal) {
            if (!['unpaid', 'maternity', 'planned sick', 'absent', 'accident', 'annual'].includes(lowerId)) {
              currentDaySecs += parseGroupedScheduleSeconds(timeVal);
            }
          }
        }
      }
    }
  }
  
  saveDay();
  return { sumMap, presenceSet };
}

function getSourceSummary() {
  return [
    `Structure: ${sourceLabels.structure}`,
    `Schedule: ${sourceLabels.schedule}`
  ].join(' | ');
}

function getVisibleDayColumnCount() {
  return dateGroups.reduce((total, day) => total + (collapsedDays[day] ? 1 : DAY_METRIC_LABELS.length), 0);
}

function getEmptyStateColspan() {
  return BASE_COLUMN_LABELS.length + getVisibleDayColumnCount();
}

function getPreferredDateGroups() {
  const utlDates = [...getDatesFromRows(sourceRows.utl, FIELD_ALIASES.utlDate)].sort();
  if (utlDates.length) return utlDates;
  return [...getStructureDates(sourceRows.structure)].sort();
}

function collapseAllDateGroups() {
  collapsedDays = {};
  dateGroups.forEach(day => { collapsedDays[day] = true; });
}

async function processData() {
  const progress = document.getElementById('progressBar');
  const status = document.getElementById('statusText');

  try {
    status.textContent = 'جاري تحميل Structure...';
    const structureSheetOptions = {
      requiredHeaderAliases: [FIELD_ALIASES.agentName],
      label: 'ملف Structure Master'
    };

    if (!filesState.struct) {
      sourceRows.structure = chooseSheet(await readBundledStructure(), ['structure', 'str', 'loss', 'master', 'sep', 'updated'], structureSheetOptions);
      sourceLabels.structure = 'STR Loss.xlsx من المستودع';
    } else {
      sourceRows.structure = chooseSheet(await readWorkbook(filesState.struct), ['structure', 'str', 'loss', 'master', 'sep', 'updated'], structureSheetOptions);
      sourceLabels.structure = filesState.struct.name;
    }

    progress.style.width = '18%';

    let _scheduleWorkbook = null;
    if (filesState.schedule) {
      status.textContent = 'جاري قراءة Schedule...';
      _scheduleWorkbook = await readWorkbook(filesState.schedule);
      sourceLabels.schedule = filesState.schedule.name;
    } else {
      sourceRows.schedule = [];
      sourceLabels.schedule = 'لم يتم رفع Schedule - سيتم الاعتماد على Structure';
    }

    progress.style.width = '34%';

    if (filesState.utl) {
      sourceRows.utl = chooseSheet(await readWorkbook(filesState.utl), ['utl', 'log']);
    } else { sourceRows.utl = []; }

    progress.style.width = '48%';

    if (filesState.ir) {
      sourceRows.ir = chooseSheet(await readWorkbook(filesState.ir), ['ir', 'ticket']);
    } else { sourceRows.ir = []; }

    progress.style.width = '62%';

    if (filesState.comp) {
      sourceRows.comp = chooseSheet(await readWorkbook(filesState.comp), ['comp', 'compensation']);
    } else { sourceRows.comp = []; }

    rawStructureData = sourceRows.structure;
    dateGroups = getPreferredDateGroups();

    if (!dateGroups.length) throw new Error('لم يتم العثور على أي تاريخ داخل الشيتات');

    status.textContent = `جاري بناء الفهارس وتجهيز ${dateGroups.length} تاريخ...`;
    progress.style.width = '75%';

    const irAssigningIndex = buildCountIndex(sourceRows.ir, FIELD_ALIASES.irAssigned, FIELD_ALIASES.irDate);
    const irTktIndex = buildCountIndex(sourceRows.ir, FIELD_ALIASES.irUser, FIELD_ALIASES.irDate);

    const talkTimeIndex = buildTalkTimeIndex(sourceRows.utl);
    const structureDurationIndex = buildSumIndex(sourceRows.structure, FIELD_ALIASES.structureId, FIELD_ALIASES.structureDate, FIELD_ALIASES.structureDuration);
    const compIndex = buildSumIndex(sourceRows.comp, FIELD_ALIASES.compId, FIELD_ALIASES.compDate, FIELD_ALIASES.compDuration);

    const _schedSheetName = _scheduleWorkbook ? getOrderedSheetNames(_scheduleWorkbook, ['schedule', 'scheduled time', 'scheduled time per agent', 'scheduled'])[0] : null;
    const _schedSheet = _schedSheetName ? _scheduleWorkbook.Sheets[_schedSheetName] : null;
    
    const { sumMap: scheduleIndex, presenceSet: scheduleDurationPresenceIndex } = buildScheduleIndexByPosition(_schedSheet);
    const hasSchedule = scheduleDurationPresenceIndex.size > 0;

    processedMatrixData = sourceRows.structure.map(row => {
      const teleoptiId = findValue(row, FIELD_ALIASES.structureId);
      const loginId = findValue(row, FIELD_ALIASES.loginId);
      const perm = findValue(row, FIELD_ALIASES.perm);
      const ttsUser = findValue(row, FIELD_ALIASES.ttsUser);
      const bssUser = findValue(row, FIELD_ALIASES.bssUser);
      const group = findValue(row, FIELD_ALIASES.group);
      const agentName = findValue(row, FIELD_ALIASES.agentName);
      const statusValue = findValue(row, FIELD_ALIASES.status, 'Active');
      const tlId = findValue(row, FIELD_ALIASES.tlId);
      const tlName = findValue(row, FIELD_ALIASES.tlName);
      const scheduleCandidates = [agentName, loginId, ttsUser, teleoptiId];
      const days = {};

      dateGroups.forEach(day => {
        const lookupCandidates = [ttsUser, loginId, agentName, teleoptiId].filter(Boolean);
        let rawAssigning = 0;
        let tkt = 0;

        for (const candidate of lookupCandidates) {
          const key = makeLookupKey(candidate, day);
          if (irAssigningIndex.has(key)) {
            rawAssigning = irAssigningIndex.get(key);
            break;
          }
        }

        for (const candidate of lookupCandidates) {
          const key = makeLookupKey(candidate, day);
          if (irTktIndex.has(key)) {
            tkt = irTktIndex.get(key);
            break;
          }
        }
        
        const systemDecimal = tkt * 0.00104166666666667;
        const systemSeconds = tkt * 90; 

        const talkTime = getIndexedValue(talkTimeIndex, loginId, day);

        const scheduleSeconds = hasSchedule
          ? getPresentIndexedValueByCandidates(scheduleIndex, scheduleDurationPresenceIndex, scheduleCandidates, day)
          : 0;
          
        const hasScheduleDuration = hasSchedule && hasPresentCandidate(scheduleDurationPresenceIndex, scheduleCandidates, day);
        
        const teleScheduleBase = hasSchedule
          ? (hasScheduleDuration ? scheduleSeconds : getIndexedValue(structureDurationIndex, teleoptiId, day))
          : getIndexedValue(structureDurationIndex, teleoptiId, day);
        
        const teleSchedule90 = teleScheduleBase * 0.9;
        const comp = getIndexedValue(compIndex, teleoptiId, day);
        const assigning = teleScheduleBase === 0 ? 0 : rawAssigning;

        const loss = String(statusValue).trim().toLowerCase() !== 'active'
          ? statusValue
          : formatTime(Math.max(0, teleSchedule90 - (systemSeconds + talkTime + comp)));

        days[day] = {
          assigning,
          tkt,
          system: systemDecimal,
          talkTime: formatTime(talkTime),
          teleSch: formatTime(teleScheduleBase), 
          comp: formatTime(comp),
          lossTime: loss
        };
      });

      return {
        teleoptiId, loginId, perm, ttsUser, bssUser, group, agentName, status: statusValue, tlId, tlName, days
      };
    });

    collapseAllDateGroups();

    progress.style.width = '100%';
    status.textContent = `تم تحديث التقرير بنجاح • ${getSourceSummary()}`;

    buildGroupToggles();
    renderMatrixTable(processedMatrixData);
  } catch (error) {
    console.error(error);
    progress.style.width = '0%';
    status.textContent = 'حدث خطأ أثناء المعالجة';
    alert(error.message);
  }
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

  document.getElementById('rowCount').textContent = `عدد الموظفين: ${rows.length}`;
  head.innerHTML = '';
  body.innerHTML = '';

  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="${getEmptyStateColspan()}" class="empty-state">لا توجد بيانات للعرض</td></tr>`;
    return;
  }

  const firstHeader = document.createElement('tr');
  BASE_COLUMN_LABELS.forEach(label => {
    const th = document.createElement('th');
    th.rowSpan = 2;
    th.className = 'th-base';
    th.textContent = label;
    firstHeader.appendChild(th);
  });

  dateGroups.forEach(day => {
    const isCollapsed = Boolean(collapsedDays[day]);
    const th = document.createElement('th');
    th.colSpan = isCollapsed ? 1 : 7;
    th.rowSpan = isCollapsed ? 2 : 1;
    th.className = 'th-date-group';

    const content = document.createElement('span');
    content.className = 'th-date-content';

    const dateText = document.createElement('span');
    dateText.textContent = displayDate(day);

    const toggleButton = document.createElement('button');
    toggleButton.type = 'button';
    toggleButton.className = 'th-day-toggle';
    toggleButton.setAttribute('aria-label', isCollapsed ? `توسيع يوم ${displayDate(day)}` : `طي يوم ${displayDate(day)}`);
    toggleButton.innerHTML = `<i class="fa-solid ${isCollapsed ? 'fa-chevron-left' : 'fa-chevron-down'}"></i>`;
    toggleButton.onclick = event => {
      event.stopPropagation();
      collapsedDays[day] = !collapsedDays[day];
      buildGroupToggles();
      renderMatrixTable(processedMatrixData);
    };

    content.appendChild(dateText);
    content.appendChild(toggleButton);
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

    [row.teleoptiId, row.loginId, row.ttsUser, row.agentName, row.status, row.tlName].forEach(value => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
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

      const values = row.days[day] || {};
      const lossText = String(values.lossTime ?? '');
      const isStatus = ['Unpaid', 'Maternity', 'Planned sick'].includes(lossText);
      const lossClass = isStatus ? 'cell-unpaid' : lossText === '0:00:00' ? 'cell-zero-loss' : 'cell-loss';

      [
        values.assigning ?? 0,
        values.tkt ?? 0,
        values.system ?? 0,
        values.talkTime ?? '0:00:00',
        values.teleSch ?? '0:00:00',
        values.comp ?? '0:00:00',
        values.lossTime ?? '0:00:00'
      ].forEach((value, index) => {
        const td = document.createElement('td');
        if (index === 2 && typeof value === 'number') {
          td.textContent = Number.isInteger(value) ? value : Number(value.toFixed(5));
        } else {
          td.textContent = value;
        }
        td.classList.add('day-metric-cell');
        if (index === 6) td.classList.add(lossClass);
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
  const filtered = processedMatrixData.filter(row =>
    [row.agentName, row.loginId, row.teleoptiId, row.ttsUser, row.tlName].some(value =>
      String(value ?? '').toLowerCase().includes(query)
    )
  );
  renderMatrixTable(filtered);
}

function exportToExcel() {
  const exportRows = buildExportRows(visibleMatrixData.length ? visibleMatrixData : processedMatrixData);
  if (!exportRows.length) {
    alert('لا توجد بيانات للتصدير');
    return;
  }
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Matrix_Report');
  XLSX.writeFile(workbook, 'CallCenter_Daily_Performance_Report.xlsx');
}

function exportToCSV() {
  const exportRows = buildExportRows(visibleMatrixData.length ? visibleMatrixData : processedMatrixData);
  if (!exportRows.length) {
    alert('لا توجد بيانات للتصدير');
    return;
  }
  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'CallCenter_Daily_Performance_Report.csv';
  link.click();
  URL.revokeObjectURL(url);
}

function buildExportRows(rows) {
  return rows.map(row => {
    const result = {
      'Teleopti ID': row.teleoptiId,
      'Login ID': row.loginId,
      'TTS User': row.ttsUser,
      'Agent Name': row.agentName,
      'Status': row.status,
      'TL Name': row.tlName
    };
    dateGroups.forEach(day => {
      if (collapsedDays[day]) return;
      const values = row.days[day] || {};
      const label = displayDate(day);
      result[`${label} - Assigning Tkts`] = values.assigning;
      result[`${label} - TKT`] = values.tkt;
      result[`${label} - System`] = values.system; 
      result[`${label} - Talk Time`] = values.talkTime;
      result[`${label} - Tele-SCH`] = values.teleSch;
      result[`${label} - Comp`] = values.comp;
      result[`${label} - Loss Time`] = values.lossTime;
    });
    return result;
  });
}