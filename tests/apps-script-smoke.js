'use strict';

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const root = path.resolve(__dirname, '..');
const gasDir = path.join(root, 'google-apps-script');
const serverFiles = [
  'code.gs',
  'setup.gs',
  'scheduler.gs',
  'teacher-schedule-adapter.gs',
  'class-schedule.gs',
  'web-ui-data.gs',
  'field-trip-ui.gs',
  'handout.gs'
];

function read(name) {
  return fs.readFileSync(path.join(gasDir, name), 'utf8');
}

const sources = Object.fromEntries(serverFiles.map(name => [name, read(name)]));
const combined = serverFiles.map(name => '// FILE: ' + name + '\n' + sources[name]).join('\n');

// Apps Script server files share one global namespace, so parse them together.
new Function(combined);

const index = read('index.html');
const scriptBlocks = [...index.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(match => match[1]);
assert(scriptBlocks.length > 0, 'index.html must contain a browser script');
scriptBlocks.forEach((script, i) => {
  assert.doesNotThrow(() => new Function(script), 'index.html script block ' + (i + 1) + ' must parse');
});

function numericConst(source, pattern, label) {
  const match = source.match(pattern);
  assert(match, 'Missing ' + label);
  return Number(match[1]);
}

const serverApi = numericConst(
  sources['code.gs'],
  /const\s+COVERAGE_WEB_API_VERSION\s*=\s*(\d+)\s*;/,
  'COVERAGE_WEB_API_VERSION'
);
const clientApi = numericConst(
  index,
  /const\s+REQUIRED_WEB_API_VERSION\s*=\s*(\d+)\s*;/,
  'REQUIRED_WEB_API_VERSION'
);
assert.strictEqual(clientApi, serverApi, 'browser/server API versions must match');

const handoutModule = numericConst(
  sources['handout.gs'],
  /const\s+FIELD_TRIP_HANDOUT_MODULE_VERSION\s*=\s*(\d+)\s*;/,
  'FIELD_TRIP_HANDOUT_MODULE_VERSION'
);
assert(
  sources['code.gs'].includes('FIELD_TRIP_HANDOUT_MODULE_VERSION !== ' + handoutModule),
  'bootstrap must reject a stale handout module'
);

function functionBody(source, name) {
  const start = source.indexOf('function ' + name + '(');
  assert(start >= 0, 'Missing function ' + name);
  const next = source.indexOf('\nfunction ', start + 10);
  return source.slice(start, next < 0 ? source.length : next);
}

const bootstrap = functionBody(sources['code.gs'], 'getCoverageBootstrap_');
assert(
  !bootstrap.includes('getFieldTripCoverageStaffForDate_'),
  'date bootstrap must not materialize field-trip coverage candidates'
);
assert(
  !bootstrap.includes("readSheetObjects_('Teacher Schedule')"),
  'date bootstrap must not read the full Teacher Schedule'
);

const handoutWeb = functionBody(sources['code.gs'], 'webCreateHandoutFromRows');
assert(
  handoutWeb.includes('validateCoveragePlanForSave_'),
  'browser handout creation must revalidate the reviewed plan'
);
const handoutPreview = functionBody(sources['handout.gs'], 'createCoverageHandoutDocWideFromLatestPreview_');
assert(
  handoutPreview.includes('validateCoveragePlanForSave_'),
  'latest-preview handout creation must revalidate before printing'
);

const generatePreview = functionBody(sources['scheduler.gs'], 'generateCoveragePreview');
assert(
  !generatePreview.includes("'Teacher Schedule'"),
  'Generate must not preload or read the full Teacher Schedule'
);
assert(
  generatePreview.includes('teacherScheduleRowsForDate_'),
  'Generate must use the prepared weekday schedule cache'
);

const liveContext = functionBody(sources['scheduler.gs'], 'buildCoverageLiveContext_');
assert(
  liveContext.includes('teacherScheduleRowsForDate_'),
  'save/manual validation must use the weekday schedule cache'
);

const dayCacheReader = functionBody(sources['scheduler.gs'], 'readTeacherScheduleDayCached_');
assert(
  dayCacheReader.includes('rebuildTeacherScheduleDayCache_'),
  'weekday schedule cache must self-heal when stale'
);

const dayCacheBuilder = functionBody(sources['scheduler.gs'], 'rebuildTeacherScheduleDayCache_');
assert(
  dayCacheBuilder.includes("Math.min(8, sheet.getLastColumn())") ||
  sources['scheduler.gs'].includes("Math.min(8, sheet.getLastColumn())"),
  'schedule cache source read must exclude the unused Lead/Co formula column'
);
assert(
  dayCacheBuilder.includes('normalizeTeacherScheduleRow_'),
  'weekday caches must store prepared/derived schedule data'
);

const poolRefresh = functionBody(sources['scheduler.gs'], 'ensureFieldTripCoveragePoolFresh_');
assert(
  poolRefresh.includes('withCoverageLock_'),
  'full field-trip pool rebuild must be serialized'
);
assert(
  poolRefresh.includes('fieldTripCoveragePoolIsDirty_'),
  'full field-trip pool rebuild must re-check freshness'
);

const datePoolRefresh = functionBody(sources['scheduler.gs'], 'fieldTripCoveragePoolRowsForDate_');
assert(
  datePoolRefresh.includes('withCoverageLock_'),
  'interactive date pool refresh must be serialized'
);
assert(
  datePoolRefresh.includes('rebuildFieldTripCoveragePoolForDate_'),
  'interactive Generate must rebuild only the requested date when stale'
);
assert(
  !datePoolRefresh.includes('rebuildAllFieldTripCoveragePool_'),
  'interactive Generate must not trigger a full-workbook pool rebuild'
);

const targetedPoolRebuild = functionBody(sources['scheduler.gs'], 'rebuildFieldTripCoveragePoolForDate_');
assert(
  targetedPoolRebuild.includes('getFieldTripsForDate_'),
  'targeted pool rebuild must limit itself to field trips active on the requested date'
);
assert(
  targetedPoolRebuild.includes('markFieldTripCoveragePoolDateFresh_'),
  'targeted pool rebuild must mark only the requested date revision fresh'
);

assert(
  sources['handout.gs'].includes('Field trip form could not remove unused assignment rows.'),
  'handout generation must verify assignment-row compaction'
);
assert(
  sources['handout.gs'].includes('Field trip form compaction failed:'),
  'handout generation must verify unused form removal'
);

// Every browser gas("method") call must have a server-side function.
const clientSources = [index, sources['code.gs'], sources['field-trip-ui.gs']];
const rpcCalls = new Set();
for (const source of clientSources) {
  for (const match of source.matchAll(/\bgas\(\s*['"]([A-Za-z_$][\w$]*)['"]/g)) {
    rpcCalls.add(match[1]);
  }
}
const serverFunctions = new Set();
for (const source of Object.values(sources)) {
  for (const match of source.matchAll(/\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g)) {
    serverFunctions.add(match[1]);
  }
}
const missingRpc = [...rpcCalls].filter(name => !serverFunctions.has(name));
assert.deepStrictEqual(missingRpc, [], 'browser RPCs missing on server: ' + missingRpc.join(', '));

// Evaluate core scheduler helpers with Apps Script services stubbed out.
const schedulerFactory = new Function(
  'SpreadsheetApp',
  'Utilities',
  'Session',
  'PropertiesService',
  'CacheService',
  'LockService',
  sources['scheduler.gs'] + `
    return {
      displayTimeToMinutes_,
      inferGradeFromClass_,
      buildCoverageNeedsByTeacher_,
      candidateCanCoverBlock_,
      makeEmptyState_,
      recordAssignment_,
      makeManualScheduleCandidate_
    };
  `
);
const noop = () => {};
const scheduler = schedulerFactory(
  { getActiveSpreadsheet: () => { throw new Error('Unexpected SpreadsheetApp call in pure smoke test'); } },
  { formatDate: (date, tz, fmt) => fmt === 'yyyy-MM-dd' ? date.toISOString().slice(0, 10) : '' },
  { getScriptTimeZone: () => 'America/New_York' },
  { getScriptProperties: () => ({ getProperty: () => null, setProperty: noop, deleteProperty: noop }) },
  { getScriptCache: () => ({ get: () => null, put: noop, remove: noop }) },
  { getScriptLock: () => ({ tryLock: () => true, releaseLock: noop }) }
);

assert.strictEqual(scheduler.displayTimeToMinutes_('12:00 AM'), 0, 'midnight parsing');
assert.strictEqual(scheduler.displayTimeToMinutes_('12:00 PM'), 720, 'noon parsing');
assert.strictEqual(scheduler.inferGradeFromClass_('3A'), '3', 'elementary class grade inference');
assert.strictEqual(scheduler.inferGradeFromClass_('8th Grade'), '8', 'ordinal grade inference');
assert.strictEqual(scheduler.inferGradeFromClass_('Algebra 1'), '', 'course number must not become a grade');
assert.strictEqual(scheduler.inferGradeFromClass_('(Art in Rm. 24)'), '', 'room number must not become a grade');

const schedule = [{
  Staff_Name: 'Teacher A',
  Term: 'All Year',
  Day: 'R',
  Start: '9:00 AM',
  End: '10:00 AM',
  Class: '3A',
  Subject: 'Science',
  Room: '14'
}];
const needs = scheduler.buildCoverageNeedsByTeacher_([{
  staffName: 'Teacher A',
  absenceType: 'Partial Day',
  startOverride: '9:15 AM',
  endOverride: '9:45 AM',
  emergency: false,
  preferredCoverage: ''
}], schedule, 'R', []);
assert.deepStrictEqual(
  [needs['Teacher A'][0].startMinutes, needs['Teacher A'][0].endMinutes],
  [555, 585],
  'partial absence must clip the scheduled block'
);

function candidate(name) {
  const c = scheduler.makeManualScheduleCandidate_(name, 'Teacher');
  c.maxBlocksPerDay = 8;
  c.maxTeachersPerDay = 2;
  c.activeToday = true;
  return c;
}
const block = {
  staffName: 'Absent A',
  startMinutes: 540,
  endMinutes: 585,
  className: '3A — Science',
  grade: '3',
  subject: 'Science',
  assignmentType: 'Class',
  room: '14',
  emergencyOverride: false
};
const state = scheduler.makeEmptyState_();
state.absencesByCandidate = {};
assert.strictEqual(
  scheduler.candidateCanCoverBlock_(
    candidate('Candidate B'),
    'Absent A',
    block,
    [{ Staff_Name: 'Candidate B', Day: 'R', Start: '9:00 AM', End: '10:00 AM', Class: '4A', Subject: 'Math', Room: '10' }],
    'R',
    state,
    { Use_Lunch_For_Coverage: 'FALSE' }
  ),
  false,
  'a candidate teaching another class must be rejected'
);
assert.strictEqual(
  scheduler.candidateCanCoverBlock_(
    candidate('Candidate B'),
    'Absent A',
    block,
    [{ Staff_Name: 'Candidate B', Day: 'R', Start: '9:00 AM', End: '10:00 AM', Subject: 'Break' }],
    'R',
    state,
    { Use_Lunch_For_Coverage: 'FALSE' }
  ),
  true,
  'a candidate on break may cover when otherwise eligible'
);

const absentState = scheduler.makeEmptyState_();
absentState.absencesByCandidate = { 'Candidate B': [{ startMinutes: 530, endMinutes: 600 }] };
assert.strictEqual(
  scheduler.candidateCanCoverBlock_(
    candidate('Candidate B'),
    'Absent A',
    block,
    [],
    'R',
    absentState,
    { Use_Lunch_For_Coverage: 'FALSE' }
  ),
  false,
  'an absent candidate must be rejected'
);

const limited = candidate('Candidate B');
limited.maxBlocksPerDay = 1;
const limitedState = scheduler.makeEmptyState_();
limitedState.absencesByCandidate = {};
scheduler.recordAssignment_(
  limitedState,
  limited,
  { startMinutes: 480, endMinutes: 525, className: '2A', fieldTripBreakMove: null },
  'Other Teacher'
);
assert.strictEqual(
  scheduler.candidateCanCoverBlock_(
    limited,
    'Absent A',
    block,
    [],
    'R',
    limitedState,
    { Use_Lunch_For_Coverage: 'FALSE' }
  ),
  false,
  'daily block limit must be enforced'
);

console.log(
  'CoverageScheduler production smoke tests passed. API ' + serverApi +
  ', handout module ' + handoutModule + ', ' + rpcCalls.size + ' browser RPCs checked.'
);
