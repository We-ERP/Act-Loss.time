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
  agentName: ['Agent Name', 'Agent', 'Name'],
  ttsUser: ['TTS User', 'TTS'],
  tlName: ['TL Name', 'Team Leader', 'TL'],
  status: ['Status'],
  irUser: ['added_by', 'IR_L_E', 'User', 'Login ID'],
  irAssigned: ['assigned_to'],
  irDate: ['added_on', 'Date'],
  utlUser: ['UL_lo', 'Login ID', 'Login', 'User'],
  utlDate: ['UL_Date', 'Date'],
  compId: ['Comp_ID', 'Comp ID', 'Teleopti ID', 'ST_ID'],
  compDate: ['Comp_Da', 'Date'],
  compDuration: ['Comp_Du', 'Comp Duration', 'Duration'],
  structureDate: ['ST_D', 'Date'],
  structureDuration: ['ST_Du', 'ST Duration', 'Duration'],
  scheduleAgent: ['Agent', 'Agent Name', 'Employee', 'Employee Name'],
  scheduleDate: ['Date', 'Scheduled Date'],
  scheduleDuration: ['Scheduled time', 'Scheduled Time', 'Scheduled time (hh:mm:ss)', 'Scheduled Time (hh:mm:ss)', 'Scheduled-Time', 'Scheduled_Time']
};

const REQUIRED_HEADER_ALIASES = {
  schedule: [
    FIELD_ALIASES.scheduleAgent,
    FIELD_ALIASES.scheduleDate,
    FIELD_ALIASES.scheduleDuration
  ]
};

const GROUPED_SCHEDULE_HEADER_ALIASES = [
  ['Contract time', 'Contract time (hh:mm)', 'Contract time (hh:mm:ss)'],
  ['Work time', 'Work time (hh:mm)', 'Work time (hh:mm:ss)'],
  ['Paid time', 'Paid time (hh:mm)', 'Paid time (hh:mm:ss)'],
  ['Scheduled overtime', 'Scheduled overtime (hh:mm)', 'Scheduled overtime (hh:mm:ss)'],
  ['Contract absence time', 'Contract absence time (hh:mm)', 'Contract absence time (hh:mm:ss)'],
  FIELD_ALIASES.scheduleDuration,
  ['Planned overtime', 'Planned overtime (hh:mm)', 'Planned overtime (hh:mm:ss)']
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

function dateKey(value) {
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

  const text = String(value).trim();
  if (!text) return '';

  const match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += 2000;

    return formatDateParts(
      year,
      Number(match[2]),
      Number(match[1])
    );
  }

  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    return formatDateParts(
      Number(isoMatch[1]),
      Number(isoMatch[2]),
      Number(isoMatch[3])
    );
  }

  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) {
    return formatLocalDate(direct);
  }

  return '';
}

function monthFirstDateKey(value) {
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

  const text = String(value).trim();
  const match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (!match) {
    const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    return isoMatch
      ? formatDateParts(
        Number(isoMatch[1]),
        Number(isoMatch[2]),
        Number(isoMatch[3])
      )
      : '';
  }

  let year = Number(match[3]);
  if (year < 100) year += 2000;

  return formatDateParts(
    year,
    Number(match[1]),
    Number(match[2])
  );
}

function parseGroupedScheduleSeconds(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') {
    return parseSeconds(value);
  }

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

  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short'
  }).replace(' ', '-');
}

function parseSeconds(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') {
    return value > 0 && value < 1 ? value * 86400 : value;
  }

  const text = String(value).trim();
  const lower = text.toLowerCase();

  if (
    !text ||
    lower === 'unpaid' ||
    lower === 'maternity' ||
    lower === 'planned sick'
  ) {
    return 0;
  }

  const parts = text.split(':').map(Number);

  if (parts.length === 3) {
    return (parts[0] || 0) * 3600 +
      (parts[1] || 0) * 60 +
      (parts[2] || 0);
  }

  if (parts.length === 2) {
    return (parts[0] || 0) * 60 + (parts[1] || 0);
  }

  return Number(text) || 0;
}

function formatTime(seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) {
    return '0:00:00';
  }

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
        resolve(
          XLSX.read(
            new Uint8Array(event.target.result),
            {
              type: 'array',
              cellDates: true,
              raw: false
            }
          )
        );
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
    response = await fetch(
      encodeURI('STR Loss.xlsx'),
      { cache: 'no-store' }
    );
  } catch (error) {
    throw new Error(
      'تعذر تحميل STR Loss.xlsx تلقائياً. إذا كنت تفتح الصفحة مباشرة من الملفات المحلية فشغّلها عبر localhost أو ارفع Structure يدوياً.'
    );
  }

  if (!response.ok) {
    throw new Error('لم يتم العثور على STR Loss.xlsx داخل جذر المستودع');
  }

  const buffer = await response.arrayBuffer();
  return XLSX.read(
    new Uint8Array(buffer),
    {
      type: 'array',
      cellDates: true,
      raw: false
    }
  );
}

function getOrderedSheetNames(workbook, words) {
  const sheetNames = workbook?.SheetNames || [];
  const preferred = [];
  const fallback = [];

  sheetNames.forEach(name => {
    if (words.some(word => matchesAnyAlias(name, [word]))) {
      preferred.push(name);
    } else {
      fallback.push(name);
    }
  });

  return [...preferred, ...fallback];
}

function matrixFromSheet(sheet) {
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: false
  });
}

function detectHeaderRow(matrix, aliasGroups) {
  let bestIndex = -1;
  let bestScore = 0;

  matrix.forEach((row, index) => {
    const cells = row
      .map(cell => normalise(cell))
      .filter(Boolean);

    if (!cells.length) return;

    const score = aliasGroups.filter(aliases =>
      aliases.some(alias => {
        const expected = normalise(alias);
        return cells.some(cell => matchesNormalisedValue(cell, expected));
      })
    ).length;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore === aliasGroups.length
    ? bestIndex
    : -1;
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

function detectGroupedScheduleHeader(matrix) {
  let bestMatch = null;

  matrix.forEach((row, index) => {
    const scheduledTimeColumnIndex = row.findIndex(cell =>
      matchesAnyAlias(cell, FIELD_ALIASES.scheduleDuration)
    );

    if (scheduledTimeColumnIndex === -1) return;

    const score = GROUPED_SCHEDULE_HEADER_ALIASES.filter(aliases =>
      row.some(cell => matchesAnyAlias(cell, aliases))
    ).length;

    if (!bestMatch || score > bestMatch.score) {
      bestMatch = { headerRowIndex: index, scheduledTimeColumnIndex, score };
    }
  });

  return bestMatch && bestMatch.score >= 3
    ? bestMatch
    : null;
}

function getGroupedScheduleLabelCell(row, scheduledTimeColumnIndex) {
  const limit = scheduledTimeColumnIndex > 0 ? scheduledTimeColumnIndex : row.length;

  for (let index = 0; index < limit; index += 1) {
    const value = row[index];
    const text = String(value ?? '').trim();

    if (text) {
      return { value, text, index };
    }
  }

  return null;
}

function extractGroupedScheduleAgentName(label) {
  const match = String(label ?? '')
    .trim()
    .match(/^\d+\s+(.+?)(?:\s+\d+)?$/);

  if (!match) return '';

  const agentName = match[1].trim();
  return /[a-z\u0600-\u06ff]/i.test(agentName)
    ? agentName
    : '';
}

function rowsFromGroupedScheduleMatrix(matrix) {
  const headerInfo = detectGroupedScheduleHeader(matrix);
  if (!headerInfo) return null;

  const { headerRowIndex, scheduledTimeColumnIndex } = headerInfo;
  const rows = [];
  let currentAgent = '';
  let pendingDay = null;

  function flushPendingDay() {
    if (!pendingDay?.agent || !pendingDay.day) {
      pendingDay = null;
      return;
    }

    const durationSeconds = pendingDay.hasDirectValue
      ? parseGroupedScheduleSeconds(pendingDay.rawDuration)
      : pendingDay.activityDurationSeconds;

    rows.push({
      [FIELD_ALIASES.scheduleAgent[0]]: pendingDay.agent,
      [FIELD_ALIASES.scheduleDate[0]]: pendingDay.day,
      [FIELD_ALIASES.scheduleDuration[0]]: formatTime(durationSeconds)
    });

    pendingDay = null;
  }

  matrix.slice(headerRowIndex + 1).forEach(row => {
    const labelCell = getGroupedScheduleLabelCell(row, scheduledTimeColumnIndex);
    if (!labelCell) return;

    const labelText = labelCell.text;
    const day = monthFirstDateKey(labelCell.value);
    const agentName = extractGroupedScheduleAgentName(labelText);

    if (
      pendingDay &&
      !day &&
      !agentName &&
      labelCell.index <= pendingDay.labelIndex
    ) {
      flushPendingDay();
    }

    if (matchesAnyAlias(labelText, ['Totals'])) {
      flushPendingDay();
      currentAgent = '';
      return;
    }

    if (agentName) {
      flushPendingDay();
      currentAgent = agentName;
      return;
    }

    if (day) {
      flushPendingDay();

      if (!currentAgent) return;

      const rawDuration = row[scheduledTimeColumnIndex] ?? '';
      pendingDay = {
        agent: currentAgent,
        day,
        labelIndex: labelCell.index,
        rawDuration,
        hasDirectValue: String(rawDuration ?? '').trim() !== '',
        activityDurationSeconds: 0
      };
      return;
    }

    if (!pendingDay || labelCell.index <= pendingDay.labelIndex) {
      return;
    }

    const activityDuration = row[scheduledTimeColumnIndex] ?? '';
    if (String(activityDuration ?? '').trim() === '') return;

    pendingDay.activityDurationSeconds += parseGroupedScheduleSeconds(activityDuration);
  });

  flushPendingDay();
  return rows;
}

function parseSheetRows(sheet, options = {}) {
  const {
    requiredHeaderAliases,
    label = 'الشيت',
    sheetName = '',
    allowGroupedScheduleFallback = false
  } = options;

  if (!sheet) return [];

  if (!requiredHeaderAliases) {
    return XLSX.utils.sheet_to_json(sheet, {
      defval: '',
      raw: false
    });
  }

  const matrix = matrixFromSheet(sheet);
  const headerRowIndex = detectHeaderRow(matrix, requiredHeaderAliases);

  if (headerRowIndex === -1) {
    if (allowGroupedScheduleFallback) {
      // Some WFM exports are grouped reports where Agent/Date are row labels, not flat columns.
      // In that case we reshape the hierarchy back into the flat rows expected by the rest of the app.
      const groupedRows = rowsFromGroupedScheduleMatrix(matrix);
      if (groupedRows) return groupedRows;
    }

    throw new Error(
      `تعذر اكتشاف صف العناوين في ${label}${sheetName ? ` (${sheetName})` : ''}. تأكد من وجود الأعمدة Agent و Date و Scheduled time.`
    );
  }

  const rows = rowsFromMatrix(matrix, headerRowIndex);
  const headerNames = (matrix[headerRowIndex] || [])
    .map(value => String(value ?? '').trim())
    .filter(Boolean);
  const missingColumns = requiredHeaderAliases
    .filter(aliases => !headerNames.some(header => matchesAnyAlias(header, aliases)))
    .map(aliases => aliases[0]);

  if (missingColumns.length) {
    throw new Error(
      `تعذر العثور على الأعمدة المطلوبة في ${label}${sheetName ? ` (${sheetName})` : ''}: ${missingColumns.join(', ')}`
    );
  }

  return rows;
}

function chooseSheet(workbook, words, options = {}) {
  const orderedNames = getOrderedSheetNames(workbook, words);
  let lastError = null;

  for (const name of orderedNames) {
    try {
      const rows = parseSheetRows(workbook.Sheets[name], {
        ...options,
        sheetName: name
      });

      if (rows.length || !options.requiredHeaderAliases) {
        return rows;
      }
    } catch (error) {
      lastError = error;
    }
  }

  if (lastError) {
    throw lastError;
  }

  return [];
}

function getDatesFromRows(rows, aliases) {
  const dates = new Set();

  rows.forEach(row => {
    const value = findValue(row, aliases, '');
    const key = dateKey(value);
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
        if (!Number.isNaN(parsed.getTime())) {
          dates.add(formatLocalDate(parsed));
        }
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
    const user = findValue(row, userAliases, '');
    const day = dateKey(findValue(row, dateAliases, ''));
    addToIndex(map, user, day, 1);
  });

  return map;
}

function buildSumIndex(rows, userAliases, dateAliases, valueAliases) {
  const map = new Map();

  rows.forEach(row => {
    const user = findValue(row, userAliases, '');
    const day = dateKey(findValue(row, dateAliases, ''));
    const amount = parseSeconds(findValue(row, valueAliases, 0));
    addToIndex(map, user, day, amount);
  });

  return map;
}

function buildDurationPresenceIndex(rows, userAliases, dateAliases, valueAliases) {
  const set = new Set();

  rows.forEach(row => {
    const rawValue = findValue(row, valueAliases, '');
    const key = makeLookupKey(
      findValue(row, userAliases, ''),
      dateKey(findValue(row, dateAliases, ''))
    );

    if (key && String(rawValue ?? '').trim() !== '') {
      set.add(key);
    }
  });

  return set;
}

function buildTalkTimeIndex(rows) {
  const map = new Map();

  rows.forEach(row => {
    const user = findValue(row, FIELD_ALIASES.utlUser, '');
    const day = dateKey(findValue(row, FIELD_ALIASES.utlDate, ''));
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
    if (presenceSet.has(key)) {
      return map.get(key) || 0;
    }
  }

  return 0;
}

function hasPresentCandidate(presenceSet, candidates, day) {
  return candidates.some(candidate => presenceSet.has(makeLookupKey(candidate, day)));
}

function getSourceSummary() {
  return [
    `Structure: ${sourceLabels.structure}`,
    `Schedule: ${sourceLabels.schedule}`
  ].join(' | ');
}

async function processData() {
  const progress = document.getElementById('progressBar');
  const status = document.getElementById('statusText');

  try {
    status.textContent = 'جاري تحميل Structure...';

    if (!filesState.struct) {
      sourceRows.structure = chooseSheet(
        await readBundledStructure(),
        ['structure', 'str', 'loss', 'master']
      );
      sourceLabels.structure = 'STR Loss.xlsx من المستودع';
    } else {
      sourceRows.structure = chooseSheet(
        await readWorkbook(filesState.struct),
        ['structure', 'str', 'loss', 'master']
      );
      sourceLabels.structure = filesState.struct.name;
    }

    progress.style.width = '18%';

    if (filesState.schedule) {
      status.textContent = 'جاري قراءة Schedule / Scheduled Time per Agent...';
      sourceRows.schedule = chooseSheet(
        await readWorkbook(filesState.schedule),
        ['schedule', 'scheduled time', 'scheduled time per agent', 'scheduled'],
        {
          allowGroupedScheduleFallback: true,
          label: 'ملف Schedule / Scheduled Time per Agent',
          requiredHeaderAliases: REQUIRED_HEADER_ALIASES.schedule
        }
      );
      sourceLabels.schedule = filesState.schedule.name;
    } else {
      sourceRows.schedule = [];
      sourceLabels.schedule = 'لم يتم رفع Schedule - سيتم الاعتماد على Structure';
    }

    progress.style.width = '34%';

    if (filesState.utl) {
      sourceRows.utl = chooseSheet(
        await readWorkbook(filesState.utl),
        ['utl', 'log']
      );
    } else {
      sourceRows.utl = [];
    }

    progress.style.width = '48%';

    if (filesState.ir) {
      sourceRows.ir = chooseSheet(
        await readWorkbook(filesState.ir),
        ['ir', 'ticket']
      );
    } else {
      sourceRows.ir = [];
    }

    progress.style.width = '62%';

    if (filesState.comp) {
      sourceRows.comp = chooseSheet(
        await readWorkbook(filesState.comp),
        ['comp', 'compensation']
      );
    } else {
      sourceRows.comp = [];
    }

    rawStructureData = sourceRows.structure;

    const dates = new Set();

    getStructureDates(sourceRows.structure).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.schedule, FIELD_ALIASES.scheduleDate).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.utl, FIELD_ALIASES.utlDate).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.ir, FIELD_ALIASES.irDate).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.comp, FIELD_ALIASES.compDate).forEach(day => dates.add(day));

    dateGroups = [...dates].sort();

    if (!dateGroups.length) {
      throw new Error('لم يتم العثور على أي تاريخ داخل الشيتات');
    }

    status.textContent = `جاري بناء الفهارس وتجهيز ${dateGroups.length} تاريخ...`;
    progress.style.width = '75%';

    const irAssigningIndex = buildCountIndex(
      sourceRows.ir,
      FIELD_ALIASES.irAssigned,
      FIELD_ALIASES.irDate
    );
    const irTktIndex = buildCountIndex(
      sourceRows.ir,
      ['IR_L_E'],
      FIELD_ALIASES.irDate
    );
    const talkTimeIndex = buildTalkTimeIndex(sourceRows.utl);
    const structureDurationIndex = buildSumIndex(
      sourceRows.structure,
      FIELD_ALIASES.structureId,
      FIELD_ALIASES.structureDate,
      FIELD_ALIASES.structureDuration
    );
    const compIndex = buildSumIndex(
      sourceRows.comp,
      FIELD_ALIASES.compId,
      FIELD_ALIASES.compDate,
      FIELD_ALIASES.compDuration
    );
    const scheduleIndex = buildSumIndex(
      sourceRows.schedule,
      FIELD_ALIASES.scheduleAgent,
      FIELD_ALIASES.scheduleDate,
      FIELD_ALIASES.scheduleDuration
    );
    const scheduleDurationPresenceIndex = buildDurationPresenceIndex(
      sourceRows.schedule,
      FIELD_ALIASES.scheduleAgent,
      FIELD_ALIASES.scheduleDate,
      FIELD_ALIASES.scheduleDuration
    );
    const hasSchedule = scheduleDurationPresenceIndex.size > 0;

    processedMatrixData = sourceRows.structure.map(row => {
      const teleoptiId = findValue(row, FIELD_ALIASES.structureId);
      const loginId = findValue(row, FIELD_ALIASES.loginId);
      const agentName = findValue(row, FIELD_ALIASES.agentName);
      const ttsUser = findValue(row, FIELD_ALIASES.ttsUser);
      const tlName = findValue(row, FIELD_ALIASES.tlName);
      const statusValue = findValue(row, FIELD_ALIASES.status, 'Active');
      const scheduleCandidates = [agentName, loginId, ttsUser, teleoptiId];
      const days = {};

      dateGroups.forEach(day => {
        const assigning = getIndexedValue(irAssigningIndex, loginId, day);
        const tkt = getIndexedValue(irTktIndex, loginId, day);
        const system = tkt * 0.00104166666666667;
        const talkTime = getIndexedValue(talkTimeIndex, loginId, day);

        const scheduleSeconds = hasSchedule
          ? getPresentIndexedValueByCandidates(
            scheduleIndex,
            scheduleDurationPresenceIndex,
            scheduleCandidates,
            day
          )
          : 0;
        const hasScheduleDuration = hasSchedule && hasPresentCandidate(
          scheduleDurationPresenceIndex,
          scheduleCandidates,
          day
        );
        const teleScheduleBase = hasSchedule
          ? (hasScheduleDuration
            ? scheduleSeconds
            : getIndexedValue(structureDurationIndex, teleoptiId, day))
          : getIndexedValue(structureDurationIndex, teleoptiId, day);
        const teleSchedule = teleScheduleBase * 0.9;
        const comp = getIndexedValue(compIndex, teleoptiId, day);

        const loss = String(statusValue).trim().toLowerCase() !== 'active'
          ? statusValue
          : formatTime(Math.max(0, teleSchedule - talkTime - comp));

        days[day] = {
          assigning,
          tkt,
          system,
          talkTime: formatTime(talkTime),
          teleSch: formatTime(teleSchedule),
          comp: formatTime(comp),
          lossTime: loss
        };
      });

      return {
        teleoptiId,
        loginId,
        agentName,
        ttsUser,
        tlName,
        status: statusValue,
        days
      };
    });

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

  document.getElementById('rowCount').textContent =
    `عدد الموظفين: ${rows.length}`;

  head.innerHTML = '';
  body.innerHTML = '';

  if (!rows.length) {
    body.innerHTML = `
      <tr>
        <td colspan="15" class="empty-state">
          لا توجد بيانات للعرض
        </td>
      </tr>
    `;
    return;
  }

  const firstHeader = document.createElement('tr');

  [
    'Teleopti ID',
    'Login ID',
    'Agent Name',
    'TTS User',
    'TL Name',
    'Status'
  ].forEach(label => {
    const th = document.createElement('th');
    th.rowSpan = 2;
    th.className = 'th-base';
    th.textContent = label;
    firstHeader.appendChild(th);
  });

  dateGroups.forEach(day => {
    if (collapsedDays[day]) return;

    const th = document.createElement('th');
    th.colSpan = 7;
    th.className = 'th-date-group';
    th.textContent = displayDate(day);
    firstHeader.appendChild(th);
  });

  head.appendChild(firstHeader);

  const secondHeader = document.createElement('tr');

  dateGroups.forEach(day => {
    if (collapsedDays[day]) return;

    [
      'Assigning Tkts',
      'TKT',
      'System',
      'Talk Time',
      'Tele-SCH',
      'Comp',
      'Loss Time'
    ].forEach((label, index) => {
      const th = document.createElement('th');
      th.className = index === 0 ? 'th-sub-orange' : 'th-sub-purple';
      th.textContent = label;
      secondHeader.appendChild(th);
    });
  });

  head.appendChild(secondHeader);

  rows.forEach(row => {
    const tr = document.createElement('tr');

    [
      row.teleoptiId,
      row.loginId,
      row.agentName,
      row.ttsUser,
      row.tlName,
      row.status
    ].forEach(value => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });

    dateGroups.forEach(day => {
      if (collapsedDays[day]) return;

      const values = row.days[day] || {};
      const lossText = String(values.lossTime ?? '');
      const isStatus = ['Unpaid', 'Maternity', 'Planned sick']
        .includes(lossText);

      const lossClass = isStatus
        ? 'cell-unpaid'
        : lossText === '0:00:00'
          ? 'cell-zero-loss'
          : 'cell-loss';

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
        td.textContent = value;

        if (index === 6) {
          td.className = lossClass;
        }

        tr.appendChild(td);
      });
    });

    body.appendChild(tr);
  });
}

function filterData() {
  const query = document
    .getElementById('searchInput')
    .value
    .trim()
    .toLowerCase();

  if (!query) {
    renderMatrixTable(processedMatrixData);
    return;
  }

  const filtered = processedMatrixData.filter(row =>
    [
      row.agentName,
      row.loginId,
      row.teleoptiId,
      row.ttsUser,
      row.tlName
    ].some(value =>
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

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    'Matrix_Report'
  );

  XLSX.writeFile(
    workbook,
    'CallCenter_Daily_Performance_Report.xlsx'
  );
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
      'Agent Name': row.agentName,
      'TTS User': row.ttsUser,
      'TL Name': row.tlName,
      'Status': row.status
    };

    dateGroups.forEach(day => {
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
