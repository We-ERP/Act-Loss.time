const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');

class FakeElement {
  constructor(tagName = 'div', id = '') {
    this.tagName = tagName.toUpperCase();
    this.id = id;
    this.children = [];
    this.attributes = {};
    this.dataset = {};
    this.style = {};
    this.value = '';
    this.className = '';
    this.onclick = null;
    this.rowSpan = 1;
    this.colSpan = 1;
    this.type = '';
    this._textContent = '';
    this._innerHTML = '';
    this.classList = {
      values: [],
      add: (...tokens) => {
        tokens.forEach(token => {
          if (!this.classList.values.includes(token)) {
            this.classList.values.push(token);
          }
        });
      }
    };
  }

  appendChild(child) {
    this.children.push(child);
    child.parentNode = this;
    return child;
  }

  setAttribute(name, value) {
    this.attributes[name] = value;
  }

  addEventListener() {}

  click() {}

  get textContent() {
    return this._textContent;
  }

  set textContent(value) {
    this._textContent = String(value);
  }

  get innerHTML() {
    return this._innerHTML;
  }

  set innerHTML(value) {
    this._innerHTML = String(value);
    this.children = [];
  }
}

function createSandbox(fileName) {
  const code = fs.readFileSync(path.join(repoRoot, fileName), 'utf8');
  const elements = new Map();

  function ensureElement(id, tagName = 'div') {
    if (!elements.has(id)) {
      elements.set(id, new FakeElement(tagName, id));
    }
    return elements.get(id);
  }

  [
    'progressBar',
    'statusText',
    'rowCount',
    'tableHead',
    'tableBody',
    'searchInput',
    'groupToggles',
    'file-struct',
    'file-schedule',
    'file-utl',
    'file-ir',
    'file-comp',
    'name-struct',
    'name-schedule',
    'name-utl',
    'name-ir',
    'name-comp',
    'card-struct',
    'card-schedule',
    'card-utl',
    'card-ir',
    'card-comp'
  ].forEach(id => ensureElement(id));

  const document = {
    getElementById(id) {
      return ensureElement(id);
    },
    createElement(tagName) {
      return new FakeElement(tagName);
    }
  };

  const sandbox = {
    console,
    setTimeout,
    clearTimeout,
    Blob: function Blob() {},
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL: () => {} },
    alert: () => {},
    window: {},
    XLSX: { SSF: { parse_date_code: () => null } },
    document
  };

  sandbox.window.XLSX = sandbox.XLSX;
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox, { filename: fileName });

  return { sandbox, elements };
}

async function runProcessDataScenario(fileName, { structureRows, utlRows, includeUtlFile }) {
  const { sandbox } = createSandbox(fileName);
  sandbox.__fixtures = {
    structureRows,
    scheduleRows: [],
    utlRows,
    irRows: [],
    compRows: [],
    includeUtlFile
  };

  vm.runInContext(`
    readBundledStructure = async () => ({ fixture: 'structure' });
    readWorkbook = async file => ({ fixture: file.kind });
    chooseSheet = (workbook, hints) => {
      const joined = Array.isArray(hints) ? hints.join('|') : '';
      if (joined.includes('structure')) return __fixtures.structureRows;
      if (joined.includes('schedule')) return __fixtures.scheduleRows;
      if (joined.includes('utl')) return __fixtures.utlRows;
      if (joined.includes('ir')) return __fixtures.irRows;
      if (joined.includes('comp')) return __fixtures.compRows;
      return [];
    };
    filesState.struct = null;
    filesState.schedule = null;
    filesState.utl = __fixtures.includeUtlFile ? { name: 'utl.xlsx', kind: 'utl' } : null;
    filesState.ir = null;
    filesState.comp = null;
  `, sandbox);

  await vm.runInContext('processData()', sandbox);

  return JSON.parse(JSON.stringify({
    processedMatrixData: vm.runInContext('processedMatrixData', sandbox),
    dateGroups: vm.runInContext('dateGroups', sandbox),
    collapsedDays: vm.runInContext('collapsedDays', sandbox)
  }));
}

async function assertProcessDataRequirements(fileName) {
  const structureRows = [
    {
      'Teleopti ID': '1001',
      'Login ID': 'agent01',
      Perm: 'OPS',
      'TTS User': 'tts-agent01',
      'BSS User': 'bss-agent01',
      Group: 'Cairo',
      'Agent Name': 'Agent One',
      Status: 'Active',
      'TL ID': 'TL-77',
      'TL Name': 'Leader One',
      ST_D: '2026-09-03',
      ST_Du: '08:00:00'
    }
  ];
  const utlRows = [
    {
      UL_lo: 'agent01',
      UL_Date: '2026-09-01',
      'Hold Time': '01:00:00',
      'Other Time': '00:30:00',
      AUXOUTOFFTIME: '00:15:00',
      ACWOUTOFFTIME: '00:00:00'
    }
  ];

  const result = await runProcessDataScenario(fileName, {
    structureRows,
    utlRows,
    includeUtlFile: true
  });

  assert.deepStrictEqual(
    result.dateGroups,
    ['2026-09-01'],
    `${fileName}: displayed dates should come from UTL only when valid UTL dates exist`
  );
  assert.deepStrictEqual(
    result.collapsedDays,
    { '2026-09-01': true },
    `${fileName}: all displayed dates should start collapsed after processing`
  );
  assert.strictEqual(result.processedMatrixData.length, 1, `${fileName}: one structure row should be processed`);
  assert.strictEqual(result.processedMatrixData[0].perm, 'OPS', `${fileName}: Perm should be extracted from structure rows`);
  assert.strictEqual(result.processedMatrixData[0].bssUser, 'bss-agent01', `${fileName}: BSS User should be extracted from structure rows`);
  assert.strictEqual(result.processedMatrixData[0].group, 'Cairo', `${fileName}: Group should be extracted from structure rows`);
  assert.strictEqual(result.processedMatrixData[0].tlId, 'TL-77', `${fileName}: TL ID should be extracted from structure rows`);

  const fallback = await runProcessDataScenario(fileName, {
    structureRows,
    utlRows: [{ UL_lo: 'agent01', UL_Date: '', 'Hold Time': '00:00:00' }],
    includeUtlFile: true
  });

  assert.deepStrictEqual(
    fallback.dateGroups,
    ['2026-09-03'],
    `${fileName}: structure dates should be used only as fallback when UTL has no valid dates`
  );
  assert.deepStrictEqual(
    fallback.collapsedDays,
    { '2026-09-03': true },
    `${fileName}: fallback dates should also start collapsed`
  );
}

function assertRenderAndExportRequirements(fileName) {
  const { sandbox, elements } = createSandbox(fileName);
  sandbox.__rows = [
    {
      teleoptiId: '1001',
      loginId: 'agent01',
      perm: 'OPS',
      ttsUser: 'tts-agent01',
      bssUser: 'bss-agent01',
      group: 'Cairo',
      agentName: 'Agent One',
      status: 'Active',
      tlId: 'TL-77',
      tlName: 'Leader One',
      days: {
        '2026-09-01': {
          assigning: 1,
          tkt: 2,
          system: 0.002,
          talkTime: '1:45:00',
          teleSch: '7:12:00',
          comp: '0:15:00',
          lossTime: '5:12:00'
        },
        '2026-09-02': {
          assigning: 3,
          tkt: 4,
          system: 0.004,
          talkTime: '2:00:00',
          teleSch: '7:00:00',
          comp: '0:30:00',
          lossTime: '4:30:00'
        }
      }
    }
  ];

  vm.runInContext(`
    dateGroups = ['2026-09-01', '2026-09-02'];
    collapsedDays = { '2026-09-01': true, '2026-09-02': false };
  `, sandbox);

  vm.runInContext('renderMatrixTable(__rows)', sandbox);

  const headerLabels = elements.get('tableHead').children[0].children
    .slice(0, 10)
    .map(cell => cell.textContent);
  assert.deepStrictEqual(
    headerLabels,
    ['Teleopti ID', 'Login ID', 'Perm', 'TTS User', 'BSS User', 'Group', 'Agent Name', 'Status', 'TL ID', 'TL Name'],
    `${fileName}: table header should show the full Structure Master fixed-column order`
  );
  assert.strictEqual(
    elements.get('tableHead').children[1].children.length,
    7,
    `${fileName}: only expanded dates should render day metric subcolumns`
  );
  assert.strictEqual(
    elements.get('tableBody').children[0].children.length,
    18,
    `${fileName}: row cells should match fixed columns plus one collapsed day and one expanded day`
  );

  const exportRows = JSON.parse(JSON.stringify(vm.runInContext('buildExportRows(__rows)', sandbox)));
  const exportKeys = Object.keys(exportRows[0]);
  const collapsedLabel = vm.runInContext(`displayDate('2026-09-01')`, sandbox);
  const expandedLabel = vm.runInContext(`displayDate('2026-09-02')`, sandbox);
  assert.deepStrictEqual(
    exportKeys.slice(0, 10),
    ['Teleopti ID', 'Login ID', 'Perm', 'TTS User', 'BSS User', 'Group', 'Agent Name', 'Status', 'TL ID', 'TL Name'],
    `${fileName}: exports should begin with the same fixed columns shown in the table`
  );
  assert(exportKeys.some(key => key.startsWith(`${expandedLabel} - `)), `${fileName}: expanded days should be exported`);
  assert(!exportKeys.some(key => key.startsWith(`${collapsedLabel} - `)), `${fileName}: collapsed days should be omitted from export output`);

  vm.runInContext('renderMatrixTable([])', sandbox);
  assert(
    elements.get('tableBody').innerHTML.includes('colspan="18"'),
    `${fileName}: empty-state colspan should match fixed columns plus currently visible day columns`
  );
}

(async () => {
  for (const fileName of ['app.js', 'app_Version2.js']) {
    await assertProcessDataRequirements(fileName);
    assertRenderAndExportRequirements(fileName);
  }

  console.log('Structure Master/date-group/export regression tests passed for app.js and app_Version2.js');
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
