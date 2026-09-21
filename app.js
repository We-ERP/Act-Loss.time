// حالة البيانات العالمية
let rawStructureData = [];
let processedMatrixData = [];
let dateGroups = ["1-Sep", "2-Sep"]; // المجموعات الافتراضية، وسيتم تحسينها ديناميكياً
let collapsedDays = {};

// الربط الأوتوماتيكي عند اختيار الملفات
const filesState = { struct: null, utl: null, ir: null, comp: null };

document.getElementById('file-struct').addEventListener('change', (e) => handleFileSelect(e, 'struct'));
document.getElementById('file-utl').addEventListener('change', (e) => handleFileSelect(e, 'utl'));
document.getElementById('file-ir').addEventListener('change', (e) => handleFileSelect(e, 'ir'));
document.getElementById('file-comp').addEventListener('change', (e) => handleFileSelect(e, 'comp'));

function handleFileSelect(event, key) {
  const file = event.target.files[0];
  if (file) {
    filesState[key] = file;
    document.getElementById(`name-${key}`).innerText = file.name;
    document.getElementById(`card-${key}`).setAttribute('data-ready', 'true');
  }
}

// قراءة شيت إكسيل وتحويله إلى JSON
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// تحويل صيغ الوقت إلى ثواني
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

// معالجة البيانات الرئيسية لبناء الجدول المطابق للصورة
async function processData() {
  if (!filesState.struct) {
    alert("يرجى اختيار ملف Structure Master أولاً!");
    return;
  }

  const progressBar = document.getElementById('progressBar');
  const statusText = document.getElementById('statusText');

  progressBar.style.width = '30%';
  statusText.innerText = 'جاري قراءة الشيتات...';

  try {
    const structJson = await readExcelFile(filesState.struct);
    rawStructureData = structJson;

    progressBar.style.width = '60%';
    statusText.innerText = 'جاري معالجة الهيكل ومجموعات الأيام...';

    // اكتشاف أياّم التقرير المتاحة داخل الشيت تلقائياً
    detectDateGroups(structJson);

    // تجميع وترتيب الصفوف
    processedMatrixData = structJson.map(row => {
      let baseInfo = {
        teleoptiId: row['Teleopti ID'] || row['Teleopti'] || '',
        loginId: row['Login ID'] || '',
        agentName: row['Agent Name'] || '',
        ttsUser: row['TTS User'] || '',
        tlName: row['TL Name'] || '',
        status: row['Status'] || 'Active'
      };

      let daysData = {};

      dateGroups.forEach(day => {
        // قراءة البيانات إما بالاسم المباشر أو الملحق برقم المجموعات كما بالصورة
        let suffix = day === '1-Sep' ? '' : (day === '2-Sep' ? '2' : '');
        let assignKey = `Assigning Tkts${suffix}`;
        let tktKey = `TKT${day === '1-Sep' ? '' : (day === '2-Sep' ? '3' : '')}`;
        let systemKey = `System${day === '1-Sep' ? '' : (day === '2-Sep' ? '4' : '')}`;
        let talkKey = `Talk Time${day === '1-Sep' ? '' : (day === '2-Sep' ? '5' : '')}`;
        let teleKey = `Tele-SCH${day === '1-Sep' ? '' : (day === '2-Sep' ? '6' : '')}`;
        let compKey = `Comp${day === '1-Sep' ? '' : (day === '2-Sep' ? '7' : '')}`;
        let lossKey = `Loss Time${day === '1-Sep' ? '' : (day === '2-Sep' ? '8' : '')}`;

        let statusVal = baseInfo.status;
        let talkTime = row[talkKey] || row['Talk Time'] || "0:00:00";
        let teleSch = row[teleKey] || row['Tele-SCH'] || "7:12:00";
        let comp = row[compKey] || row['Comp'] || "0:00:00";

        // حساب Loss Time إذا كانت الحالة Active
        let lossTimeCalculated = row[lossKey];
        if (!lossTimeCalculated || lossTimeCalculated === '') {
          if (statusVal === 'Active') {
            let teleSec = parseTimeToSeconds(teleSch);
            let talkSec = parseTimeToSeconds(talkTime);
            let compSec = parseTimeToSeconds(comp);
            let lossSec = Math.max(0, teleSec - (talkSec + compSec));
            lossTimeCalculated = secondsToHHMMSS(lossSec);
          } else {
            lossTimeCalculated = statusVal;
          }
        }

        daysData[day] = {
          assigning: row[assignKey] || row['Assigning Tkts'] || 0,
          tkt: row[tktKey] || row['TKT'] || 0,
          system: row[systemKey] || row['System'] || "0:00:00",
          talkTime: talkTime,
          teleSch: teleSch,
          comp: comp,
          lossTime: lossTimeCalculated
        };
      });

      return { ...baseInfo, days: daysData };
    });

    progressBar.style.width = '100%';
    statusText.innerText = 'تم التحديث وبناء الجدول بنجاح!';

    buildGroupToggles();
    renderMatrixTable(processedMatrixData);

  } catch (err) {
    console.error(err);
    alert("حدث خطأ أثناء معالجة البيانات: " + err.message);
  }
}

// تحديد الأيام المتاحة
function detectDateGroups(json) {
  if (!json || !json.length) return;
  const keys = Object.keys(json[0]);
  let detected = [];
  
  keys.forEach(k => {
    if (k.includes('-Sep') || k.includes('-Aug') || k.includes('-Oct') || k.toLowerCase().includes('day')) {
      if (!detected.includes(k)) detected.push(k);
    }
  });

  if (detected.length > 0) {
    dateGroups = detected;
  } else {
    dateGroups = ["1-Sep", "2-Sep"];
  }
}

// بناء أزرار طي وفتح الأيام
function buildGroupToggles() {
  const container = document.getElementById('groupToggles');
  container.innerHTML = '<span class="control-label">عرض/طي أيام التقرير:</span>';

  dateGroups.forEach(day => {
    let btn = document.createElement('button');
    btn.className = 'day-btn';
    btn.innerText = `مجموعة ${day}`;
    btn.onclick = () => toggleDayGroup(day, btn);
    container.appendChild(btn);
  });
}

function toggleDayGroup(day, btn) {
  collapsedDays[day] = !collapsedDays[day];
  btn.classList.toggle('collapsed', collapsedDays[day]);
  renderMatrixTable(processedMatrixData);
}

// بناء رندر الجدول المطابق لصورة المستخدم بالكامل
function renderMatrixTable(data) {
  const thead = document.getElementById('tableHead');
  const tbody = document.getElementById('tableBody');

  document.getElementById('rowCount').innerText = `عدد الموظفين: ${data.length}`;

  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="15" class="empty-state"><p>لا توجد بيانات للعرض</p></td></tr>`;
    return;
  }

  // Row 1 Header: Fixed Headers + Date Group Headers
  let tr1 = document.createElement('tr');

  const baseHeaders = ["Teleopti ID", "Login ID", "Agent Name", "TTS User", "TL Name", "Status"];
  baseHeaders.forEach(h => {
    let th = document.createElement('th');
    th.rowSpan = 2;
    th.className = 'th-base';
    th.innerText = h;
    tr1.appendChild(th);
  });

  dateGroups.forEach(day => {
    if (!collapsedDays[day]) {
      let th = document.createElement('th');
      th.colSpan = 7;
      th.className = 'th-date-group';
      th.innerText = day;
      tr1.appendChild(th);
    }
  });

  thead.appendChild(tr1);

  // Row 2 Header: Sub-columns under each day
  let tr2 = document.createElement('tr');

  dateGroups.forEach(day => {
    if (!collapsedDays[day]) {
      const subCols = [
        { name: "Assigning Tkts", cls: "th-sub-orange" },
        { name: "TKT", cls: "th-sub-purple" },
        { name: "System", cls: "th-sub-purple" },
        { name: "Talk Time", cls: "th-sub-purple" },
        { name: "Tele-SCH", cls: "th-sub-purple" },
        { name: "Comp", cls: "th-sub-purple" },
        { name: "Loss Time", cls: "th-sub-purple" }
      ];

      subCols.forEach(sc => {
        let th = document.createElement('th');
        th.className = sc.cls;
        th.innerText = sc.name;
        tr2.appendChild(th);
      });
    }
  });

  thead.appendChild(tr2);

  // Render Table Rows Body
  data.forEach(row => {
    let tr = document.createElement('tr');

    tr.innerHTML = `
      <td>${row.teleoptiId}</td>
      <td>${row.loginId}</td>
      <td style="text-align: right; font-weight: 600;">${row.agentName}</td>
      <td>${row.ttsUser}</td>
      <td>${row.tlName}</td>
      <td><span class="status-tag">${row.status}</span></td>
    `;

    dateGroups.forEach(day => {
      if (!collapsedDays[day]) {
        let d = row.days[day] || {};
        
        // Dynamic Loss Time Styling
        let lossClass = "";
        if (d.lossTime === 'Unpaid' || d.lossTime === 'Planned sick' || d.lossTime === 'Maternity') {
          lossClass = "cell-unpaid";
        } else if (d.lossTime === '0:00:00' || d.lossTime === '00:00:00') {
          lossClass = "cell-zero-loss";
        } else {
          lossClass = "cell-loss";
        }

        tr.innerHTML += `
          <td>${d.assigning || 0}</td>
          <td>${d.tkt || 0}</td>
          <td>${d.system || "0:00:00"}</td>
          <td>${d.talkTime || "0:00:00"}</td>
          <td>${d.teleSch || "0:00:00"}</td>
          <td>${d.comp || "0:00:00"}</td>
          <td class="${lossClass}">${d.lossTime || "0:00:00"}</td>
        `;
      }
    });

    tbody.appendChild(tr);
  });
}

// التصفية والبحث
function filterData() {
  let q = document.getElementById('searchInput').value.toLowerCase().trim();
  if (!q) {
    renderMatrixTable(processedMatrixData);
    return;
  }

  let filtered = processedMatrixData.filter(r => 
    String(r.agentName).toLowerCase().includes(q) ||
    String(r.loginId).toLowerCase().includes(q) ||
    String(r.teleoptiId).toLowerCase().includes(q) ||
    String(r.tlName).toLowerCase().includes(q)
  );

  renderMatrixTable(filtered);
}

// تصدير الشيت إلى إكسيل بالكامل بنفس المجموعات
function exportToExcel() {
  if (!processedMatrixData.length) return alert("لا توجد بيانات للتصدير!");

  let exportRows = processedMatrixData.map(r => {
    let flatObj = {
      "Teleopti ID": r.teleoptiId,
      "Login ID": r.loginId,
      "Agent Name": r.agentName,
      "TTS User": r.ttsUser,
      "TL Name": r.tlName,
      "Status": r.status
    };

    dateGroups.forEach(day => {
      let d = r.days[day] || {};
      flatObj[`${day}_Assigning Tkts`] = d.assigning;
      flatObj[`${day}_TKT`] = d.tkt;
      flatObj[`${day}_System`] = d.system;
      flatObj[`${day}_Talk Time`] = d.talkTime;
      flatObj[`${day}_Tele-SCH`] = d.teleSch;
      flatObj[`${day}_Comp`] = d.comp;
      flatObj[`${day}_Loss Time`] = d.lossTime;
    });

    return flatObj;
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Matrix_Report");
  XLSX.writeFile(wb, "CallCenter_Matrix_Daily_Report.xlsx");
}

// تصدير CSV
function exportToCSV() {
  if (!processedMatrixData.length) return alert("لا توجد بيانات للتصدير!");
  exportToExcel(); // التصدير عبر Excel يضمن التوافق التام مع اللغة العربية وتنسيق الجدول
}