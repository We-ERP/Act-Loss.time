let rawStructureData = [];
let processedMatrixData = [];
let visibleMatrixData = [];
let dateGroups = [];
let collapsedDays = {};

const sourceRows = { structure: [], schedule: [], utl: [], ir: [], comp: [] };
const filesState = { struct: null, schedule: null, utl: null, ir: null, comp: null };
const sourceLabels = { structure: 'STR Loss.xlsx', schedule: 'لم يتم رفع Schedule' };

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

function normalise(v) { return String(v ?? '').trim().toLowerCase().replace(/[^a-z0-9\u0600-\u06ff]+/gi, ''); }

function parseDateAny(val) {
  if (!val) return '';
  if (val instanceof Date && !Number.isNaN(val.getTime())) {
    return `${val.getFullYear()}-${String(val.getMonth()+1).padStart(2,'0')}-${String(val.getDate()).padStart(2,'0')}`;
  }
  if (typeof val === 'number' && window.XLSX?.SSF) {
    const p = XLSX.SSF.parse_date_code(val);
    if (p) return `${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`;
  }
  let s = String(val).trim();
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let yr = Number(m[3]); if (yr < 100) yr += 2000;
    return `${yr}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  }
  return '';
}

function parseSec(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val > 0 && val < 1 ? val * 86400 : val;
  let s = String(val).trim();
  let parts = s.split(':').map(Number);
  if (parts.length === 3) return (parts[0]||0)*3600 + (parts[1]||0)*60 + (parts[2]||0);
  if (parts.length === 2) return (parts[0]||0)*60 + (parts[1]||0);
  return Number(s) || 0;
}

function formatTime(sec) {
  if (!Number.isFinite(sec) || sec <= 0) return '0:00:00';
  let h = Math.floor(sec / 3600);
  let m = Math.floor((sec % 3600) / 60);
  let s = Math.floor(sec % 60);
  return `${h}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
}

function readWb(file) {
  return new Promise((res, rej) => {
    let r = new FileReader();
    r.onload = e => {
      try { res(XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true, raw: false })); }
      catch(err) { rej(err); }
    };
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}

async function getBundledStr() {
  let res = await fetch(encodeURI('STR Loss.xlsx'), { cache: 'no-store' });
  let buf = await res.arrayBuffer();
  return XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true, raw: false });
}

function sheetToJsonRaw(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
}

function parseGenericSheet(sheet) {
  if (!sheet) return [];
  let matrix = sheetToJsonRaw(sheet);
  if (matrix.length === 0) return [];
  
  let headerIdx = 0;
  for (let i = 0; i < Math.min(10, matrix.length); i++) {
    if (matrix[i].some(c => String(c).trim() !== '')) {
      headerIdx = i;
      break;
    }
  }
  
  let headers = matrix[headerIdx].map((v, idx) => String(v).trim() || `Col_${idx+1}`);
  let rows = [];
  for (let i = headerIdx + 1; i < matrix.length; i++) {
    let r = matrix[i];
    if (!r.some(cell => String(cell).trim() !== '')) continue;
    let obj = {};
    headers.forEach((h, idx) => { obj[h] = r[idx] ?? ''; });
    rows.push(obj);
  }
  return rows;
}

async function processData() {
  const progress = document.getElementById('progressBar');
  const status = document.getElementById('statusText');

  try {
    status.textContent = 'جاري التحميل ومعالجة البيانات...';
    
    let strWb = filesState.struct ? await readWb(filesState.struct) : await getBundledStr();
    sourceRows.structure = parseGenericSheet(strWb.Sheets[strWb.SheetNames[0]]);
    if (progress) progress.style.width = '20%';

    if (filesState.ir) {
      let irWb = await readWb(filesState.ir);
      sourceRows.ir = parseGenericSheet(irWb.Sheets[irWb.SheetNames[0]]);
    } else {
      sourceRows.ir = [];
    }
    if (progress) progress.style.width = '40%';

    if (filesState.utl) {
      let utlWb = await readWb(filesState.utl);
      sourceRows.utl = parseGenericSheet(utlWb.Sheets[utlWb.SheetNames[0]]);
    } else { sourceRows.utl = []; }
    if (progress) progress.style.width = '60%';

    if (filesState.comp) {
      let compWb = await readWb(filesState.comp);
      sourceRows.comp = parseGenericSheet(compWb.Sheets[compWb.SheetNames[0]]);
    } else { sourceRows.comp = []; }
    if (progress) progress.style.width = '80%';

    let datesSet = new Set();
    sourceRows.structure.forEach(r => {
      Object.keys(r).forEach(k => {
        let dt = parseDateAny(r[k]);
        if (dt && dt.length === 10) datesSet.add(dt);
      });
    });
    
    sourceRows.ir.forEach(r => {
      Object.keys(r).forEach(k => {
        if (/date|added_on|time/i.test(k)) {
          let dt = parseDateAny(r[k]);
          if (dt) datesSet.add(dt);
        }
      });
    });

    dateGroups = [...datesSet].sort();
    if (dateGroups.length === 0) {
      dateGroups = ['2026-09-01'];
    }

    let irTktMap = new Map();
    let irAssignMap = new Map();

    sourceRows.ir.forEach(r => {
      let addedBy = '', assignedTo = '', dateVal = '';
      Object.keys(r).forEach(k => {
        let lk = k.toLowerCase();
        if (lk.includes('added_by') || lk === 'added by') addedBy = String(r[k]).trim();
        if (lk.includes('assigned_to') || lk === 'assigned to') assignedTo = String(r[k]).trim();
        if (lk.includes('added_on') || lk === 'added on' || lk.includes('date')) {
          let parsed = parseDateAny(r[k]);
          if (parsed) dateVal = parsed;
        }
      });

      if (dateVal) {
        if (addedBy) {
          let k1 = `${normalise(addedBy)}|${dateVal}`;
          irTktMap.set(k1, (irTktMap.get(k1) || 0) + 1);
        }
        if (assignedTo) {
          let k2 = `${normalise(assignedTo)}|${dateVal}`;
          irAssignMap.set(k2, (irAssignMap.get(k2) || 0) + 1);
        }
      }
    });

    processedMatrixData = sourceRows.structure.map(row => {
      let keys = Object.keys(row);
      let getVal = (aliases) => {
        let foundKey = keys.find(k => aliases.some(a => normalise(k).includes(normalise(a))));
        return foundKey !== undefined ? row[foundKey] : '';
      };

      let teleoptiId = getVal(['Teleopti ID', 'Teleopti', 'ST_ID', 'ID']);
      let loginId = getVal(['Login ID', 'Login', 'UL_lo', 'User']);
      let ttsUser = getVal(['TTS User', 'TTS']);
      let agentName = getVal(['Agent Name', 'Agent', 'Name']);
      let statusValue = getVal(['Status']) || 'Active';
      let tlName = getVal(['TL Name', 'Team Leader', 'TL']);

      let days = {};
      dateGroups.forEach(day => {
        let candidates = [ttsUser, loginId, agentName, teleoptiId].filter(Boolean);
        let tktCount = 0;
        let assigningCount = 0;

        candidates.forEach(c => {
          let k = `${normalise(c)}|${day}`;
          if (irTktMap.has(k)) tktCount += irTktMap.get(k);
          if (irAssignMap.has(k)) assigningCount += irAssignMap.get(k);
        });

        let systemDecimal = tktCount * 0.00104166666666667;
        let systemSeconds = tktCount * 90;
        let teleSchSec = 28800;
        let talkSec = 0;
        let compSec = 0;

        let lossSec = (teleSchSec * 0.9) - (systemSeconds + talkSec + compSec);
        if (lossSec < 0) lossSec = 0;

        let lossTimeStr = String(statusValue).trim().toLowerCase() !== 'active' 
          ? statusValue 
          : formatTime(lossSec);

        days[day] = {
          assigning: assigningCount,
          tkt: tktCount,
          system: systemDecimal,
          talkTime: formatTime(talkSec),
          teleSch: formatTime(teleSchSec),
          comp: formatTime(compSec),
          lossTime: lossTimeStr
        };
      });

      return { teleoptiId, loginId, ttsUser, agentName, status: statusValue, tlName, days };
    });

    collapseAllDateGroups();
    if (progress) progress.style.width = '100%';
    if (status) status.textContent = 'تم معالجة البيانات بنجاح!';

    buildGroupToggles();
    renderMatrixTable(processedMatrixData);

  } catch (err) {
    console.error(err);
    if (progress) progress.style.width = '0%';
    if (status) status.textContent = 'خطأ في المعالجة';
    alert('حدث خطأ: ' + err.message);
  }
}

function collapseAllDateGroups() {
  collapsedDays = {};
  dateGroups.forEach(d => { collapsedDays[d] = true; });
}

function buildGroupToggles() {
  const container = document.getElementById('groupToggles');
  if (!container) return;
  container.innerHTML = '<span class="control-label">عرض/طي الأيام:</span>';
  dateGroups.forEach(day => {
    let btn = document.createElement('button');
    btn.className = `day-btn${collapsedDays[day] ? ' collapsed' : ''}`;
    btn.textContent = day;
    btn.onclick = () => {
      collapsedDays[day] = !collapsedDays[day];
      buildGroupToggles();
      renderMatrixTable(processedMatrixData);
    };
    container.appendChild(btn);
  });
}

function renderMatrixTable(rows) {
  const head = document.getElementById('tableHead');
  const body = document.getElementById('tableBody');
  const rowCountEl = document.getElementById('rowCount');
  if (!head || !body) return;
  
  visibleMatrixData = [...rows];
  if (rowCountEl) rowCountEl.textContent = `عدد الموظفين: ${rows.length}`;
  head.innerHTML = '';
  body.innerHTML = '';

  if (rows.length === 0) {
    body.innerHTML = '<tr><td colspan="10" style="text-align:center; padding:20px;">لا توجد بيانات للعرض</td></tr>';
    return;
  }

  let tr1 = document.createElement('tr');
  ['Teleopti ID', 'Login ID', 'TTS User', 'Agent Name', 'Status', 'TL Name'].forEach(lbl => {
    let th = document.createElement('th');
    th.rowSpan = 2; th.textContent = lbl;
    tr1.appendChild(th);
  });

  dateGroups.forEach(day => {
    let isCol = collapsedDays[day];
    let th = document.createElement('th');
    th.colSpan = isCol ? 1 : 7;
    th.rowSpan = isCol ? 2 : 1;
    th.textContent = day;
    tr1.appendChild(th);
  });
  head.appendChild(tr1);

  let tr2 = document.createElement('tr');
  dateGroups.forEach(day => {
    if (collapsedDays[day]) return;
    ['Assigning Tkts', 'TKT', 'System', 'Talk Time', 'Tele-SCH', 'Comp', 'Loss Time'].forEach(m => {
      let th = document.createElement('th');
      th.textContent = m;
      tr2.appendChild(th);
    });
  });
  head.appendChild(tr2);

  rows.forEach(r => {
    let tr = document.createElement('tr');
    [r.teleoptiId, r.loginId, r.ttsUser, r.agentName, r.status, r.tlName].forEach(val => {
      let td = document.createElement('td');
      td.textContent = val ?? '';
      tr.appendChild(td);
    });

    dateGroups.forEach(day => {
      if (collapsedDays[day]) {
        let td = document.createElement('td');
        td.textContent = '—';
        tr.appendChild(td);
        return;
      }
      let d = r.days[day] || {};
      [d.assigning, d.tkt, d.system, d.talkTime, d.teleSch, d.comp, d.lossTime].forEach((v, idx) => {
        let td = document.createElement('td');
        td.textContent = (idx === 2 && typeof v === 'number') ? v.toFixed(5) : v;
        tr.appendChild(td);
      });
    });
    body.appendChild(tr);
  });
}

// تفعيل زر المعالجة والبحث التلقائي لو موجودين
document.addEventListener('DOMContentLoaded', () => {
  const processBtn = document.getElementById('processBtn');
  if (processBtn) {
    processBtn.addEventListener('click', processData);
  }

  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      let term = normalise(e.target.value);
      if (!term) {
        renderMatrixTable(processedMatrixData);
        return;
      }
      let filtered = processedMatrixData.filter(r => 
        normalise(r.agentName).includes(term) ||
        normalise(r.loginId).includes(term) ||
        normalise(r.ttsUser).includes(term) ||
        normalise(r.teleoptiId).includes(term) ||
        normalise(r.tlName).includes(term)
      );
      renderMatrixTable(filtered);
    });
  }
});
