const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');

class FakeElement {
  constructor(tagName = 'div', id = '') {
    Object.assign(this, {
      tagName: tagName.toUpperCase(), id, children: [], attributes: {}, dataset: {}, style: {},
      value: '', className: '', onclick: null, rowSpan: 1, colSpan: 1, type: '', hidden: false,
      _textContent: '', _innerHTML: ''
    });
    this.classList = {
      values: [],
      add: (...t) => t.forEach(x => { if (!this.classList.values.includes(x)) this.classList.values.push(x); })
    };
  }
  appendChild(c) { this.children.push(c); c.parentNode = this; return c; }
  setAttribute(n, v) { this.attributes[n] = v; }
  addEventListener() {}
  click() {}
  get textContent() { return this._textContent; }
  set textContent(v) { this._textContent = String(v); }
  get innerHTML() { return this._innerHTML; }
  set innerHTML(v) { this._innerHTML = String(v); this.children = []; }
}

function createSandbox() {
  const code = fs.readFileSync(path.join(repoRoot, 'app.js'), 'utf8');
  const elements = new Map();
  const ensure = id => { if (!elements.has(id)) elements.set(id, new FakeElement('div', id)); return elements.get(id); };
  const document = { getElementById: ensure, createElement: t => new FakeElement(t) };
  const sandbox = {
    console, setTimeout, clearTimeout, Blob: function Blob() {},
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL: () => {} },
    alert: () => {}, window: {}, XLSX: { SSF: { parse_date_code: () => null } }, document
  };
  sandbox.window.XLSX = sandbox.XLSX;
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox, { filename: 'app.js' });
  return { sandbox, elements };
}

const run = (sandbox, code) => vm.runInContext(code, sandbox);
const plain = v => JSON.parse(JSON.stringify(v));

/* ── 1) التواريخ ─────────────────────────────────────────────────────────── */
function testDates() {
  const { sandbox } = createSandbox();
  const k = v => run(sandbox, `dateKey(${JSON.stringify(v)})`);
  assert.strictEqual(k('9/1/2026 8:41:42 AM'), '2026-09-01', 'M/D/YYYY with time → 1 Sep');
  assert.strictEqual(k('9/1/2026'), '2026-09-01');
  assert.strictEqual(k('9/1/26'), '2026-09-01');
  assert.strictEqual(k('13/9/2026'), '2026-09-13', 'first > 12 is treated as day');
  assert.strictEqual(k('2026-09-03'), '2026-09-03');
  assert.strictEqual(k('Break 1'), '', 'activity codes must never parse as dates');
  assert.strictEqual(k('Phone'), '');
}

/* ── 2) مطابقة أسماء الأعمدة ─────────────────────────────────────────────── */
function testColumnMatching() {
  const { sandbox } = createSandbox();
  sandbox.__row = { 'TL ID': 'TL-77', 'TL Name': 'Leader One', 'TTS User': 't', 'BSS User': 'b', 'Login ID': 'l' };
  assert.strictEqual(run(sandbox, `findValue(__row, FIELD_ALIASES.tlName)`), 'Leader One');
  assert.strictEqual(run(sandbox, `findValue(__row, FIELD_ALIASES.tlId)`), 'TL-77');
  assert.strictEqual(run(sandbox, `findValue(__row, FIELD_ALIASES.loginId)`), 'l');
}

/* ── 3) Schedule: الأكواد ────────────────────────────────────────────────── */
function testScheduleCodes() {
  const { sandbox } = createSandbox();
  sandbox.__matrix = [
    ['', 'Agent', 'Date', '', '', '', '', '', '', 'Scheduled time'],
    ['', '73935 Mostafa Hanafy', '', '', '', '', '', '', '', '8:00:00'],
    ['', '', '9/1/2026', '', '', '', '', '', '', '8:00:00'],
    ['', '', 'Phone', '', '', '', '', '', '', '6:00:00'],
    ['', '', 'Break 1', '', '', '', '', '', '', '1:00:00'],
    ['', '', 'Lunch', '', '', '', '', '', '', '1:00:00'],
    ['', 'Totals', '', '', '', '', '', '', '', '8:00:00']
  ];
  const parsed = run(sandbox, `parseScheduleMatrix(__matrix)`);
  const entry = parsed.index.get('73935|2026-09-01');
  assert(entry, 'agent/day entry exists keyed by ID from column B');
  assert.strictEqual(entry.total, 8 * 3600);
  assert.strictEqual(parsed.codeTotals.size, 3, 'Phone, Break 1, Lunch');

  sandbox.__entry = entry;
  assert.strictEqual(run(sandbox, `scheduleSecondsFor(__entry, new Set())`), 8 * 3600, 'no codes → day total');
  assert.strictEqual(run(sandbox, `scheduleSecondsFor(__entry, new Set(['phone']))`), 6 * 3600);
  assert.strictEqual(run(sandbox, `scheduleSecondsFor(__entry, new Set(['phone','break 1']))`), 7 * 3600);
}

/* ── 3b) شيت Final بعد الماكرو + الأكواد ───────────────────────────────────── */
function testFinalSchedule() {
  const { sandbox } = createSandbox();
  sandbox.__m = [
    ['ID', 'Date', 'Duration', 'Agent Name', 'TL', 'Code'],
    [73935, '9/1/2026', 0.25, 'A', 'T', 'Phone'],                       // 6:00:00
    [73935, '9/1/2026', 1 / 24, 'A', 'T', 'Covering PC Pro'],           // 1:00:00
    [73935, '9/1/2026', 1 / 48, 'A', 'T', 'Break']                      // 0:30:00 (غير مختار)
  ];
  const parsed = run(sandbox, 'parseFinalScheduleMatrix(__m)');
  assert(parsed, 'Final-format schedule detected');
  sandbox.__e = parsed.index.get('73935|2026-09-01');
  assert(sandbox.__e, 'entry keyed by ID + date');
  assert.strictEqual(
    run(sandbox, `scheduleSecondsFor(__e, new Set(['phone', 'covering pc pro']))`),
    7 * 3600,
    'only the selected codes are summed'
  );
}

/* ── 3c) تشخيص IR ───────────────────────────────────────────────────────── */
function testDiagnoseIR() {
  const { sandbox } = createSandbox();
  sandbox.__ir = [
    { assigned_to: 'Mostafa.M69994', added_by: 'Mostafa.M69994@co.com', added_on: '9/1/2026 8:41:42 AM' },
    { assigned_to: 'nobody', added_by: 'nobody', added_on: 'bad' }
  ];
  sandbox.__st = [{ 'TTS User': 'Mostafa.M69994', 'Agent Name': 'A' }];
  const d = plain(run(sandbox, 'diagnoseIR(__ir, __st)'));
  assert.strictEqual(d.addedMatched, 1, 'email suffix is ignored');
  assert.strictEqual(d.assignedMatched, 1);
  assert.strictEqual(d.datesRead, 1);
  assert.deepStrictEqual(d.unmatched, ['nobody']);
}


/* ── 3d) Compensation بالـ User Name + Code Time ────────────────────────────── */
function testCompensationByUser() {
  const { sandbox } = createSandbox();
  sandbox.__ctx = {
    structureRows: [{ 'Teleopti ID': 68261, 'Login ID': 83957, 'Agent Name': 'M', 'TTS User': 'Mahmoud.a.aly', Status: 'Active', 'TL Name': 'T' }],
    utlRows: [], irRows: [],
    compRows: [
      { ID: 68261, 'User Name ': 'Mahmoud.a.aly', TL: 'Atef Shoaib', Reason: 'INQ Task', Date: '9/1/2026', Shift: '9:00 AM - 6:00 PM', 'Code Time': '8:00:00' },
      { ID: 68261, 'User Name ': 'Mahmoud.a.aly', TL: 'Atef Shoaib', Reason: 'x', Date: '9/1/2026', Shift: '', 'Code Time': '0:45:00' }
    ],
    schedule: { index: new Map([['68261|2026-09-01', { total: 32400, hasTotal: true, codes: new Map() }]]) },
    days: ['2026-09-01'], codes: new Set(), config: {}
  };
  const [a] = plain(run(sandbox, 'buildMatrix(__ctx)'));
  const d = a.days['2026-09-01'];
  assert.strictEqual(d.comp, '8:45:00', 'two comp rows summed by user + date');
  assert.strictEqual(d.teleSch, '8:06:00');
  assert.strictEqual(d.lossTime, '0:00:00', 'comp bigger than Tele-SCH → clamped to zero');
}

/* ── 3e) Schedule الخام بنفس شكل Teleopti ────────────────────────────────────── */
function testRawTeleoptiSchedule() {
  const { sandbox } = createSandbox();
  const pad = (cells) => { const r = Array(11).fill(''); Object.entries(cells).forEach(([i, v]) => { r[i] = v; }); return r; };
  sandbox.__m = [
    pad({ 1: 'Scheduled Time per Agent', 10: '9/20/2026 12:42:44 PM' }),
    pad({ 1: 'Date:', 3: '9/1/2026 – 9/19/2026' }),
    pad({ 4: 'Contract time (hh:mm)', 5: 'Work time (hh:mm)', 9: 'Scheduled time (hh:mm)' }),
    pad({ 1: 'Totals:', 4: '16793:00', 9: '16793:00' }),
    pad({ 1: 'TEData Welcome Call-New Profile 1', 4: '120:00', 9: '120:00' }),
    pad({ 1: '156958 Hesham nabil mohamed ali 86466', 4: '120:00', 9: '120:00' }),
    pad({ 2: 'Tuesday, September 01, 2026', 9: '8:00' }),
    pad({ 3: 'Phone', 9: '6:00' }),
    pad({ 3: 'Covering PC Pro', 9: '1:00' }),
    pad({ 3: 'Lunch', 9: '1:00' }),
    pad({ 1: 'TEData Welcome Call-New Profile 2', 4: '60:00', 9: '60:00' }),
    pad({ 1: '200410 Another Agent 1234', 4: '8:00', 9: '8:00' }),
    pad({ 2: '9/2/2026', 9: '8:00' }),
    pad({ 2: 'Phone', 9: '8:00' })
  ];
  const parsed = run(sandbox, 'parseScheduleMatrix(__m)');
  const e1 = parsed.index.get('156958|2026-09-01');
  assert(e1, 'agent 156958 / 1 Sep found (date written as text with weekday)');
  assert.strictEqual(e1.codes.get('phone'), 6 * 3600);
  assert.strictEqual(e1.codes.get('covering pc pro'), 3600);
  assert(!parsed.codeTotals.has('teData welcome call-new profile 2'.toLowerCase()), 'team rows are never codes');
  assert.strictEqual(parsed.codeTotals.size, 3, 'Phone, Covering PC Pro, Lunch only');
  assert(parsed.index.get('200410|2026-09-02'), 'second agent parsed after team row');
}

/* ── 4) المعادلة الكاملة ─────────────────────────────────────────────────── */
function testLossFormula() {
  const { sandbox } = createSandbox();
  sandbox.__ctx = {
    structureRows: [{ 'Teleopti ID': 73935, 'Login ID': 601492, 'Agent Name': 'A', 'TTS User': 'Mostafa.M69994', Status: 'Active', 'TL Name': 'T' },
                    { 'Teleopti ID': 2, 'Login ID': 3, 'Agent Name': 'B', 'TTS User': 'b.b', Status: 'Unpaid', 'TL Name': 'T' }],
    utlRows: [{ UL_lo: '601492', UL_Date: '9/1/2026', 'Hold Time': '1:00:00', 'Other Time': '0:30:00', AUXOUTOFFTIME: '0:15:00', ACWOUTOFFTIME: '0:00:00' }],
    irRows: [
      { assigned_to: 'Mostafa.M69994', added_by: 'someone', added_on: '9/1/2026 8:41:42 AM' },
      { assigned_to: 'Mostafa.M69994', added_by: 'Mostafa.M69994', added_on: '9/1/2026 9:02:17 AM' },
      { assigned_to: 'x', added_by: 'Mostafa.M69994', added_on: '9/1/2026 9:30:00 AM' }
    ],
    compRows: [{ Comp_ID: '73935', Comp_Da: '9/1/2026', Comp_Du: '0:15:00' }],
    schedule: { index: new Map([['73935|2026-09-01', { total: 28800, hasTotal: true, codes: new Map() }]]) },
    days: ['2026-09-01'],
    codes: new Set(),
    config: { teleSchFactor: 0.9, secondsPerTicket: 90 }
  };
  const [a, b] = plain(run(sandbox, `buildMatrix(__ctx)`));
  const d = a.days['2026-09-01'];
  assert.strictEqual(d.assigning, 2);
  assert.strictEqual(d.tkt, 2);
  assert.strictEqual(d.system, '0:03:00', '2 tickets × 90s');
  assert.strictEqual(d.talkTime, '1:45:00');
  assert.strictEqual(d.teleSch, '7:12:00', 'Tele-SCH is shown after the 90% factor');
  assert.strictEqual(d.comp, '0:15:00');
  // Tele-SCH 7:12:00 − (0:03:00 + 1:45:00 + 0:15:00) = 5:09:00
  assert.strictEqual(d.lossTime, '5:09:00');
  assert.strictEqual(b.days['2026-09-01'].lossTime, 'Unpaid', 'non-active shows status');
  // 0.00104166666666667 day = 90 seconds
  assert.strictEqual(Math.round(0.00104166666666667 * 86400), 90);
}

/* ── 5) التواريخ المعروضة ────────────────────────────────────────────────── */
async function runProcessScenario({ structureRows, utlRows }) {
  const { sandbox } = createSandbox();
  sandbox.__f = { structureRows, utlRows };
  run(sandbox, `
    readBundledStructure = async () => ({ fixture: 'structure' });
    readWorkbook = async file => ({ fixture: file.kind });
    chooseSheet = (wb, hints) => {
      const j = hints.join('|');
      if (j.includes('structure')) return __f.structureRows;
      if (j.includes('utl')) return __f.utlRows;
      return [];
    };
    filesState.utl = { name: 'utl.xlsx', kind: 'utl' };
  `);
  await run(sandbox, 'processData()');
  return plain({
    rows: run(sandbox, 'processedMatrixData'),
    dateGroups: run(sandbox, 'dateGroups'),
    collapsedDays: run(sandbox, 'collapsedDays')
  });
}

async function testProcessData() {
  const structureRows = [{
    'Teleopti ID': '1001', 'Login ID': 'agent01', Perm: 'OPS', 'TTS User': 'tts-agent01', 'BSS User': 'bss-agent01',
    Group: 'Cairo', 'Agent Name': 'Agent One', Status: 'Active', 'TL ID': 'TL-77', 'TL Name': 'Leader One',
    ST_D: '2026-09-03', ST_Du: '08:00:00'
  }];

  const r1 = await runProcessScenario({ structureRows, utlRows: [{ UL_lo: 'agent01', UL_Date: '2026-09-01', 'Hold Time': '01:00:00' }] });
  assert.deepStrictEqual(r1.dateGroups, ['2026-09-01'], 'dates come from UTL when valid');
  assert.deepStrictEqual(r1.collapsedDays, { '2026-09-01': true });
  const row = r1.rows[0];
  assert.strictEqual(row.perm, 'OPS');
  assert.strictEqual(row.bssUser, 'bss-agent01');
  assert.strictEqual(row.group, 'Cairo');
  assert.strictEqual(row.tlId, 'TL-77');
  assert.strictEqual(row.tlName, 'Leader One', 'TL Name must not pick TL ID');

  const r2 = await runProcessScenario({ structureRows, utlRows: [{ UL_lo: 'agent01', UL_Date: '', 'Hold Time': '0:00:00' }] });
  assert.deepStrictEqual(r2.dateGroups, ['2026-09-03'], 'fallback to other sources when UTL has no dates');
}

/* ── 6) العرض والتصدير ───────────────────────────────────────────────────── */
function testRenderAndExport() {
  const { sandbox, elements } = createSandbox();
  sandbox.__rows = [{
    teleoptiId: '1001', loginId: 'agent01', perm: 'OPS', ttsUser: 'tts', bssUser: 'bss', group: 'Cairo',
    agentName: 'Agent One', status: 'Active', tlId: 'TL-77', tlName: 'Leader One',
    days: {
      '2026-09-01': { assigning: 1, tkt: 2, system: '0:03:00', talkTime: '1:45:00', teleSch: '7:12:00', comp: '0:15:00', lossTime: '5:12:00' },
      '2026-09-02': { assigning: 3, tkt: 4, system: '0:06:00', talkTime: '2:00:00', teleSch: '7:00:00', comp: '0:30:00', lossTime: '4:30:00' }
    }
  }];
  run(sandbox, `dateGroups = ['2026-09-01','2026-09-02']; collapsedDays = {'2026-09-01': true, '2026-09-02': false};`);
  run(sandbox, 'renderMatrixTable(__rows)');

  const labels = elements.get('tableHead').children[0].children.slice(0, 10).map(c => c.textContent);
  assert.deepStrictEqual(labels, ['Teleopti ID', 'Login ID', 'Perm', 'TTS User', 'BSS User', 'Group', 'Agent Name', 'Status', 'TL ID', 'TL Name']);
  assert.strictEqual(elements.get('tableHead').children[1].children.length, 7, 'only expanded day has sub-columns');
  assert.strictEqual(elements.get('tableBody').children[0].children.length, 18);

  const exportRows = plain(run(sandbox, 'buildExportRows(__rows)'));
  const keys = Object.keys(exportRows[0]);
  assert.deepStrictEqual(keys.slice(0, 6), ['Teleopti ID', 'Login ID', 'Agent Name', 'TTS User', 'TL Name', 'Status']);
  const l1 = run(sandbox, `displayDate('2026-09-01')`);
  const l2 = run(sandbox, `displayDate('2026-09-02')`);
  assert(keys.some(k => k.startsWith(`${l1} - `)) && keys.some(k => k.startsWith(`${l2} - `)),
    'export includes ALL days, even collapsed ones');

  run(sandbox, 'renderMatrixTable([])');
  assert(elements.get('tableBody').innerHTML.includes('colspan="18"'));
}

(async () => {
  testDates();
  testColumnMatching();
  testScheduleCodes();
  testFinalSchedule();
  testDiagnoseIR();
  testCompensationByUser();
  testRawTeleoptiSchedule();
  testLossFormula();
  await testProcessData();
  testRenderAndExport();
  console.log('✅ All tests passed');
})().catch(e => { console.error(e); process.exitCode = 1; });
