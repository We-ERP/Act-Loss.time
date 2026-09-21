let rawStructureData = [];
let processedMatrixData = [];
let dateGroups = [];
let collapsedDays = {};

const sourceRows = {
  structure: [],
  utl: [],
  ir: [],
  comp: []
};

const filesState = {
  struct: null,
  utl: null,
  ir: null,
  comp: null
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
  structureDuration: ['ST_Du', 'ST Duration', 'Duration']
};

['struct', 'utl', 'ir', 'comp'].forEach(key => {
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
    .replace(/[\s_\-\/]+/g, '');
}

function findValue(row, aliases, fallback = '') {
  if (!row) return fallback;

  const wanted = aliases.map(normalise);
  const key = Object.keys(row).find(name =>
    wanted.includes(normalise(name))
  );

  return key === undefined ? fallback : row[key];
}

function dateKey(value) {
  if (value === null || value === undefined || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number' && window.XLSX?.SSF) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }

  const text = String(value).trim();
  if (!text) return '';

  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) {
    return direct.toISOString().slice(0, 10);
  }

  const match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += 2000;

    return `${year}-${String(Number(match[2])).padStart(2, '0')}-${String(Number(match[1])).padStart(2, '0')}`;
  }

  return '';
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
    // Excel time values are fractions of a day.
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
        const workbook = XLSX.read(
          new Uint8Array(event.target.result),
          {
            type: 'array',
            cellDates: true,
            raw: false
          }
        );

        const result = {};

        workbook.SheetNames.forEach(sheetName => {
          result[sheetName] = XLSX.utils.sheet_to_json(
            workbook.Sheets[sheetName],
            {
              defval: '',
              raw: false
            }
          );
        });

        resolve(result);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function readBundledStructure() {
  const response = await fetch(
    encodeURI('STR Loss.xlsx'),
    { cache: 'no-store' }
  );

  if (!response.ok) {
    throw new Error('لم يتم العثور على STR Loss.xlsx');
  }

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(
    new Uint8Array(buffer),
    {
      type: 'array',
      cellDates: true,
      raw: false
    }
  );

  const result = {};

  workbook.SheetNames.forEach(sheetName => {
    result[sheetName] = XLSX.utils.sheet_to_json(
      workbook.Sheets[sheetName],
      {
        defval: '',
        raw: false
      }
    );
  });

  return result;
}

function firstSheet(workbook) {
  return Object.values(workbook || {})[0] || [];
}

function chooseSheet(workbook, words) {
  const found = Object.entries(workbook || {}).find(([name]) =>
    words.some(word => normalise(name).includes(normalise(word)))
  );

  return found ? found[1] : firstSheet(workbook);
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

  // Supports wide STR Loss sheets that have date columns in their headers.
  rows.forEach(row => {
    Object.keys(row).forEach(key => {
      const match = key.match(/^(\d{1,2})[-\/](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$/i);

      if (match) {
        const parsed = new Date(`${match[1]} ${match[2]} 2024`);
        if (!Number.isNaN(parsed.getTime())) {
          dates.add(parsed.toISOString().slice(0, 10));
        }
      }
    });
  });

  return dates;
}

function countRows(rows, userAliases, dateAliases, user, day) {
  return rows.filter(row => {
    const rowUser = findValue(row, userAliases, '');
    const rowDate = findValue(row, dateAliases, '');

    return String(rowUser).trim() === String(user).trim() &&
      dateKey(rowDate) === day;
  }).length;
}

function sumRows(rows, userAliases, dateAliases, user, day, valueAliases) {
  return rows.reduce((total, row) => {
    const rowUser = findValue(row, userAliases, '');
    const rowDate = findValue(row, dateAliases, '');

    if (
      String(rowUser).trim() === String(user).trim() &&
      dateKey(rowDate) === day
    ) {
      return total + parseSeconds(findValue(row, valueAliases, 0));
    }

    return total;
  }, 0);
}

function sumById(rows, idAliases, dateAliases, id, day, valueAliases) {
  return rows.reduce((total, row) => {
    const rowId = findValue(row, idAliases, '');
    const rowDate = findValue(row, dateAliases, '');

    if (
      String(rowId).trim() === String(id).trim() &&
      dateKey(rowDate) === day
    ) {
      return total + parseSeconds(findValue(row, valueAliases, 0));
    }

    return total;
  }, 0);
}

async function processData() {
  const progress = document.getElementById('progressBar');
  const status = document.getElementById('statusText');

  try {
    if (!filesState.struct) {
      status.textContent = 'جاري تحميل STR Loss.xlsx من المستودع...';
      const bundled = await readBundledStructure();
      sourceRows.structure = chooseSheet(bundled, ['structure', 'str', 'loss', 'master']);
    } else {
      const workbook = await readWorkbook(filesState.struct);
      sourceRows.structure = chooseSheet(workbook, ['structure', 'str', 'loss', 'master']);
    }

    progress.style.width = '20%';

    if (filesState.utl) {
      sourceRows.utl = chooseSheet(
        await readWorkbook(filesState.utl),
        ['utl', 'log']
      );
    }

    progress.style.width = '35%';

    if (filesState.ir) {
      sourceRows.ir = chooseSheet(
        await readWorkbook(filesState.ir),
        ['ir', 'ticket']
      );
    }

    progress.style.width = '50%';

    if (filesState.comp) {
      sourceRows.comp = chooseSheet(
        await readWorkbook(filesState.comp),
        ['comp', 'compensation']
      );
    }

    rawStructureData = sourceRows.structure;

    const dates = new Set();

    getStructureDates(sourceRows.structure).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.utl, FIELD_ALIASES.utlDate).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.ir, FIELD_ALIASES.irDate).forEach(day => dates.add(day));
    getDatesFromRows(sourceRows.comp, FIELD_ALIASES.compDate).forEach(day => dates.add(day));

    dateGroups = [...dates].sort();

    if (!dateGroups.length) {
      throw new Error('لم يتم العثور على أي تاريخ داخل الشيتات');
    }

    status.textContent = `جاري حساب ${dateGroups.length} تاريخ...`;
    progress.style.width = '70%';

    processedMatrixData = sourceRows.structure.map(row => {
      const teleoptiId = findValue(row, FIELD_ALIASES.structureId);
      const loginId = findValue(row, FIELD_ALIASES.loginId);
      const statusValue = findValue(row, FIELD_ALIASES.status, 'Active');

      const days = {};

      dateGroups.forEach(day => {
        // Assigning Tkts = COUNTIFS(assigned_to, Login ID, Date, Day)
        const assigning = countRows(
          sourceRows.ir,
          FIELD_ALIASES.irAssigned,
          FIELD_ALIASES.irDate,
          loginId,
          day
        );

        // TKT = COUNTIFS(IR_L_E, Login ID, Date, Day)
        const tkt = countRows(
          sourceRows.ir,
          ['IR_L_E'],
          FIELD_ALIASES.irDate,
          loginId,
          day
        );

        // System = TKT * 0.00104166666666667
        const system = tkt * 0.00104166666666667;

        // Talk Time = Hold + Other + AUXOUTOFFTIME + ACWOUTOFFTIME
        const talkTime =
          sumRows(sourceRows.utl, FIELD_ALIASES.utlUser, FIELD_ALIASES.utlDate, loginId, day, ['Hold Time', 'HoldTime']) +
          sumRows(sourceRows.utl, FIELD_ALIASES.utlUser, FIELD_ALIASES.utlDate, loginId, day, ['Other Time', 'OtherTime']) +
          sumRows(sourceRows.utl, FIELD_ALIASES.utlUser, FIELD_ALIASES.utlDate, loginId, day, ['AUXOUTOFFTIME']) +
          sumRows(sourceRows.utl, FIELD_ALIASES.utlUser, FIELD_ALIASES.utlDate, loginId, day, ['ACWOUTOFFTIME']);

        // Tele-SCH = SUMIFS(ST_Du, ST_ID, Teleopti ID, ST_D, Day) * 0.9
        const teleSchedule =
          sumById(
            sourceRows.structure,
            FIELD_ALIASES.structureId,
            FIELD_ALIASES.structureDate,
            teleoptiId,
            day,
            FIELD_ALIASES.structureDuration
          ) * 0.9;

        // Comp = SUMIFS(Comp_Du, Comp_ID, Teleopti ID, Comp_Da, Day)
        const comp =
          sumById(
            sourceRows.comp,
            FIELD_ALIASES.compId,
            FIELD_ALIASES.compDate,
            teleoptiId,
            day,
            FIELD_ALIASES.compDuration
          );

        // Loss Time =
        // IF(Status <> Active, Status, MAX(0, Tele-SCH - Talk Time - Comp))
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
        agentName: findValue(row, FIELD_ALIASES.agentName),
        ttsUser: findValue(row, FIELD_ALIASES.ttsUser),
        tlName: findValue(row, FIELD_ALIASES.tlName),
        status: statusValue,
        days
      };
    });

    progress.style.width = '100%';
    status.textContent = 'تم تحديث التقرير حتى آخر تاريخ موجود';

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
  if (!processedMatrixData.length) {
    alert('لا توجد بيانات للتصدير');
    return;
  }

  const rows = processedMatrixData.map(row => {
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

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows);

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
  exportToExcel();
}