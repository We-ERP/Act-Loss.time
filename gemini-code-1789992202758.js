let filesData = { utl: null, ir: null, comp: null, sched: null, struct: null };
let processedRows = [];

// تحويل الوقت من hh:mm:ss إلى ثواني
function parseTimeToSeconds(val) {
  if (!val || val === 'Unpaid' || val === 'Maternity' || val === 'Planned sick') return 0;
  let strVal = String(val).trim();
  let parts = strVal.split(':');
  if (parts.length === 3) {
    return (parseInt(parts[0], 10) || 0) * 3600 + (parseInt(parts[1], 10) || 0) * 60 + (parseInt(parts[2], 10) || 0);
  } else if (parts.length === 2) {
    return (parseInt(parts[0], 10) || 0) * 60 + (parseInt(parts[1], 10) || 0);
  }
  return 0;
}

// تحويل الثواني إلى hh:mm:ss
function secondsToHHMMSS(sec) {
  if (isNaN(sec) || sec <= 0) return "0:00:00";
  let h = Math.floor(sec / 3600);
  let m = Math.floor((sec % 3600) / 60);
  let s = sec % 60;
  return `${h}:${m < 10 ? '0' : ''}${m}:${s < 10 ? '0' : ''}${s}`;
}

// ربط أزرار اختيار الملفات
['utl', 'ir', 'comp', 'sched', 'struct'].forEach(key => {
  const input = document.getElementById(`file-${key}`);
  input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      filesData[key] = file;
      document.getElementById(`name-${key}`).innerText = file.name;
      document.getElementById(`card-${key}`).setAttribute('data-ready', 'true');
    }
  });
});

async function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const json = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });
      resolve(json);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function processAllData() {
  if (!filesData.struct) {
    alert("يرجى اختيار شيت Structure الخامس لكي يتم تحديثه!");
    return;
  }

  const progressBar = document.getElementById('progressBar');
  const runHint = document.getElementById('runHint');
  progressBar.style.width = '30%';
  runHint.innerText = 'جاري قراءة الشيتات...';

  try {
    const structJson = await readExcel(filesData.struct);
    const utlJson = filesData.utl ? await readExcel(filesData.utl) : [];
    const irJson = filesData.ir ? await readExcel(filesData.ir) : [];
    const compJson = filesData.comp ? await readExcel(filesData.comp) : [];

    progressBar.style.width = '70%';
    runHint.innerText = 'جاري معالجة وتحديث الـ Structure...';

    // 1. خريطة الـ IR
    let irAssignMap = {};
    let irClosedMap = {};
    irJson.forEach(row => {
      let user = row['assigned_to'] || row['added_by'];
      if (user) {
        irAssignMap[user] = (irAssignMap[user] || 0) + 1;
        if (row['ticket_status'] === 'Reached') {
          irClosedMap[user] = (irClosedMap[user] || 0) + 1;
        }
      }
    });

    // 2. خريطة Compensation
    let compMap = {};
    compJson.forEach(row => {
      let id = String(row['ID'] || row['Login ID'] || '').trim();
      if (id) compMap[id] = row['Code Time'] || "0:00:00";
    });

    // 3. خريطة UTL
    let utlMap = {};
    utlJson.forEach(row => {
      let id = String(row['Login ID'] || '').trim();
      if (id) utlMap[id] = row;
    });

    let activeCount = 0;
    let totalAssigned = 0;
    let totalTalkSec = 0;
    let totalLossSec = 0;

    processedRows = structJson.map(row => {
      let loginId = String(row['Login ID'] || '').trim();
      let ttsUser = String(row['TTS User'] || '').trim();
      let status = row['Status'] || 'Active';

      if (status === 'Active') activeCount++;

      let utlData = utlMap[loginId] || {};
      let talkTime = utlData['Talk Time'] || row['Talk Time'] || "0:00:00";
      let assignedTkts = irAssignMap[ttsUser] !== undefined ? irAssignMap[ttsUser] : (row['Assigning Tkts'] || 0);
      let tktCount = irClosedMap[ttsUser] !== undefined ? irClosedMap[ttsUser] : (row['TKT'] || 0);
      let compVal = compMap[loginId] || row['Comp'] || "0:00:00";

      totalAssigned += Number(assignedTkts);
      let talkSec = parseTimeToSeconds(talkTime);
      totalTalkSec += talkSec;

      // حساب Loss Time = Tele-SCH - (Talk Time + Comp)
      let teleSchSec = parseTimeToSeconds(row['Tele-SCH'] || "7:12:00");
      let compSec = parseTimeToSeconds(compVal);
      let lossSec = Math.max(0, teleSchSec - (talkSec + compSec));
      
      let lossTimeStr = status === 'Active' ? secondsToHHMMSS(lossSec) : status;
      if (status === 'Active') totalLossSec += lossSec;

      return {
        ...row,
        "Talk Time": talkTime,
        "Assigning Tkts": assignedTkts,
        "TKT": tktCount,
        "Comp": compVal,
        "Loss Time": lossTimeStr
      };
    });

    progressBar.style.width = '100%';
    runHint.innerText = 'تم التحديث بنجاح!';

    document.getElementById('statTotalAgents').innerText = processedRows.length;
    document.getElementById('statActiveAgents').innerText = activeCount;
    document.getElementById('statTickets').innerText = totalAssigned;
    document.getElementById('statTalkTime').innerText = secondsToHHMMSS(totalTalkSec);
    document.getElementById('statLossTime').innerText = secondsToHHMMSS(totalLossSec);

    renderTable(processedRows);
    document.getElementById('results').style.display = 'block';

  } catch (err) {
    alert("حدث خطأ أثناء معالجة البيانات: " + err.message);
    console.error(err);
  }
}

function renderTable(data) {
  if (!data || !data.length) return;
  const headerTr = document.getElementById('tableHeader');
  const body = document.getElementById('tableBody');

  headerTr.innerHTML = '';
  body.innerHTML = '';

  const headers = Object.keys(data[0]);
  headers.forEach(h => {
    let th = document.createElement('th');
    th.innerText = h;
    headerTr.appendChild(th);
  });

  data.forEach(row => {
    let tr = document.createElement('tr');
    headers.forEach(h => {
      let td = document.createElement('td');
      td.innerText = row[h] !== undefined ? row[h] : '';
      if (h.includes('Loss Time')) td.classList.add('loss-cell');
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function filterTable() {
  let q = document.getElementById('searchInput').value.toLowerCase();
  let filtered = processedRows.filter(r => 
    String(r['Agent Name'] || '').toLowerCase().includes(q) ||
    String(r['Login ID'] || '').toLowerCase().includes(q)
  );
  renderTable(filtered);
}

function exportToExcel() {
  const ws = XLSX.utils.json_to_sheet(processedRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Final_Report");
  XLSX.writeFile(wb, "Final_Updated_CallCenter_Report.xlsx");
}

function exportToCSV() {
  const ws = XLSX.utils.json_to_sheet(processedRows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Final_Updated_CallCenter_Report.csv");
  document.body.appendChild(link);
  link.click();
}