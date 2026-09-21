const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');

function loadScheduleHelpers(fileName) {
  const code = fs.readFileSync(path.join(repoRoot, fileName), 'utf8');
  const sandbox = {
    console,
    setTimeout,
    clearTimeout,
    Blob: function Blob() {},
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL: () => {} },
    alert: () => {},
    window: {},
    XLSX: {},
    document: {
      getElementById: () => null,
      createElement: () => ({
        click() {},
        appendChild() {},
        setAttribute() {},
        style: {},
        dataset: {}
      })
    }
  };

  vm.createContext(sandbox);
  vm.runInContext(code, sandbox, { filename: fileName });

  return {
    detectHeaderRow: sandbox.detectHeaderRow,
    detectGroupedScheduleHeader: sandbox.detectGroupedScheduleHeader,
    rowsFromGroupedScheduleMatrix: sandbox.rowsFromGroupedScheduleMatrix,
    rowsFromMatrix: sandbox.rowsFromMatrix,
    extractGroupedScheduleNumberPrefixedAgentName: sandbox.extractGroupedScheduleNumberPrefixedAgentName
  };
}

function runSharedAssertions(fileName) {
  const helpers = loadScheduleHelpers(fileName);

  const minimalGroupedMatrix = [
    ['Team Schedule'],
    ['Agent', 'Day', 'Activity', 'Scheduled time (hh:mm)'],
    ['156958 Hesham nabil mohamed ali 86466', '', '', ''],
    ['', '9/1/2026', '', '9:00'],
    ['', '', 'Break 2', '1:00'],
    ['', '', 'Phone', '8:00'],
    ['', '9/2/2026', '', ''],
    ['', '', 'Break 2', '1:00'],
    ['', '', 'Phone', '8:00'],
    ['Totals', '', '', '18:00']
  ];

  const headerInfo = helpers.detectGroupedScheduleHeader(minimalGroupedMatrix);
  assert(headerInfo, `${fileName}: grouped header should be detected when Scheduled time is the only matching duration header`);
  assert.strictEqual(headerInfo.headerRowIndex, 1, `${fileName}: grouped header row index should point at the Scheduled time header row`);
  assert.strictEqual(headerInfo.scheduledTimeColumnIndex, 3, `${fileName}: Scheduled time column index should be detected correctly`);

  assert.strictEqual(
    helpers.extractGroupedScheduleNumberPrefixedAgentName('156958 Hesham nabil mohamed ali 86466'),
    'Hesham nabil mohamed ali',
    `${fileName}: number-prefixed agent names should preserve the agent name and ignore the trailing numeric token`
  );

  const groupedRows = helpers.rowsFromGroupedScheduleMatrix(minimalGroupedMatrix);
  assert.strictEqual(groupedRows.length, 2, `${fileName}: two grouped day rows should be produced`);
  assert.strictEqual(groupedRows[0].Agent, 'Hesham nabil mohamed ali', `${fileName}: grouped parser should resolve the agent name from ID + name + trailing number rows`);
  assert.strictEqual(groupedRows[0].Date, '2026-09-01', `${fileName}: first grouped row date should be normalised`);
  assert.strictEqual(groupedRows[0]['Scheduled time'], '9:00:00', `${fileName}: direct day values should win over child activity totals`);
  assert.strictEqual(groupedRows[1].Agent, 'Hesham nabil mohamed ali', `${fileName}: second grouped row should keep the current agent`);
  assert.strictEqual(groupedRows[1].Date, '2026-09-02', `${fileName}: second grouped row date should be normalised`);
  assert.strictEqual(groupedRows[1]['Scheduled time'], '9:00:00', `${fileName}: missing direct day values should fall back to summed child activity durations`);

  const richerHeaderMatrix = [
    ['Scheduled time'],
    ['Label', 'Contract time', 'Scheduled time', 'Work time'],
    ['156958 Hesham nabil mohamed ali 86466', '', '', ''],
    ['9/3/2026', '', '9:00', '']
  ];
  const richerHeaderInfo = helpers.detectGroupedScheduleHeader(richerHeaderMatrix);
  assert(richerHeaderInfo, `${fileName}: richer grouped header should still be detected`);
  assert.strictEqual(richerHeaderInfo.headerRowIndex, 1, `${fileName}: richer grouped header rows should still outrank weaker candidates by score`);
  assert.strictEqual(richerHeaderInfo.scheduledTimeColumnIndex, 2, `${fileName}: richer grouped header Scheduled time column should be returned`);

  const flatMatrix = [
    ['ignore me'],
    ['Agent', 'Date', 'Scheduled time'],
    ['Jane Doe', '2026-09-01', '08:30:00']
  ];
  const flatHeaderIndex = helpers.detectHeaderRow(flatMatrix, [
    ['Agent', 'Agent Name', 'Employee', 'Employee Name'],
    ['Date', 'Scheduled Date'],
    ['Scheduled time', 'Scheduled Time', 'Scheduled time (hh:mm:ss)', 'Scheduled Time (hh:mm:ss)', 'Scheduled-Time', 'Scheduled_Time']
  ]);
  assert.strictEqual(flatHeaderIndex, 1, `${fileName}: flat schedule header detection should remain unchanged`);
  const flatRows = helpers.rowsFromMatrix(flatMatrix, flatHeaderIndex);
  assert.strictEqual(flatRows.length, 1, `${fileName}: flat schedule parsing should still produce one row`);
  assert.strictEqual(flatRows[0].Agent, 'Jane Doe', `${fileName}: flat schedule parsing should preserve Agent values`);
  assert.strictEqual(flatRows[0].Date, '2026-09-01', `${fileName}: flat schedule parsing should preserve Date values`);
  assert.strictEqual(flatRows[0]['Scheduled time'], '08:30:00', `${fileName}: flat schedule parsing should preserve Scheduled time values`);
}

['app.js', 'app_Version2.js'].forEach(runSharedAssertions);
console.log('Grouped schedule regression tests passed for app.js and app_Version2.js');
