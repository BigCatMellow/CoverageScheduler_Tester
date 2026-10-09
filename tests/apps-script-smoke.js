'use strict';

const fs = require('fs');
const path = require('path');
const assert = require('assert');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const gasDir = path.join(root, 'google-apps-script');
const serverFiles = fs.readdirSync(gasDir)
  .filter(name => name.endsWith('.gs'))
  .sort();

function read(name) {
  return fs.readFileSync(path.join(gasDir, name), 'utf8');
}

const sources = Object.fromEntries(serverFiles.map(name => [name, read(name)]));
const combined = serverFiles.map(name => '// FILE: ' + name + '\n' + sources[name]).join('\n');

// Parse each Apps Script file as a true script first. This catches syntax that
// Function() incorrectly permits, including a top-level return from a bad
// extraction boundary.
serverFiles.forEach(name => {
  assert.doesNotThrow(
    () => new vm.Script(sources[name], { filename: name }),
    name + ' must parse as an Apps Script file'
  );
});

// Apps Script server files share one global namespace, so also parse them together.
new Function(combined);

const uiPartialFiles = [
  'styles.html',
  'ui-core.html',
  'ui-plan.html',
  'ui-people.html',
  'ui-manual.html',
  'ui-actions.html'
];

function expandIndexTemplateForTest(template) {
  return template.replace(
    /<\?!=\s*includeCoveragePartial_\('([^']+)'\);\s*\?>/g,
    (_, name) => read(name + '.html')
  );
}

const indexTemplate = read('index.html');
const index = expandIndexTemplateForTest(indexTemplate);
assert(
  !index.includes('<?!= includeCoveragePartial_'),
  'all index partials must resolve in smoke tests'
);
const scriptBlocks = [...index.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(match => match[1]);
assert(scriptBlocks.length > 0, 'expanded index.html must contain a browser script');
scriptBlocks.forEach((script, i) => {
  assert.doesNotThrow(() => new Function(script), 'expanded index.html script block ' + (i + 1) + ' must parse');
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

const schedulerModule = numericConst(
  sources['scheduler.gs'],
  /const\s+COVERAGE_SCHEDULER_MODULE_VERSION\s*=\s*(\d+)\s*;/,
  'COVERAGE_SCHEDULER_MODULE_VERSION'
);
assert(
  sources['code.gs'].includes('COVERAGE_SCHEDULER_MODULE_VERSION !== ' + schedulerModule),
  'bootstrap must reject a stale scheduler module'
);

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

function sourceForFunction(name) {
  const entry = Object.entries(sources).find(([, source]) =>
    source.includes('function ' + name + '(')
  );
  assert(entry, 'Missing server function ' + name);
  return entry[1];
}

function functionBodyAny(name) {
  return functionBody(sourceForFunction(name), name);
}

const includePartial = functionBody(sources['code.gs'], 'includeCoveragePartial_');
assert(
  includePartial.includes('HtmlService.createHtmlOutputFromFile'),
  'web UI partials must be composed through the Apps Script template include helper'
);
assert(
  indexTemplate.includes("includeCoveragePartial_('styles')") &&
  indexTemplate.includes("includeCoveragePartial_('ui-core')") &&
  indexTemplate.includes("includeCoveragePartial_('ui-plan')") &&
  indexTemplate.includes("includeCoveragePartial_('ui-people')") &&
  indexTemplate.includes("includeCoveragePartial_('ui-manual')") &&
  indexTemplate.includes("includeCoveragePartial_('ui-actions')"),
  'index.html must remain a thin shell that composes all UI partials'
);

const bootstrap = functionBody(sources['code.gs'], 'getCoverageBootstrap_');
assert(
  !bootstrap.includes('getFieldTripCoverageStaffForDate_'),
  'date bootstrap must not materialize field-trip coverage candidates'
);
assert(
  !bootstrap.includes("readSheetObjects_('Teacher Schedule')"),
  'date bootstrap must not read the full Teacher Schedule'
);

assert(
  !bootstrap.includes("'_Preview'"),
  'web bootstrap must not read the shared _Preview scratch sheet'
);
assert(
  bootstrap.includes("'Coverage Output'"),
  'web bootstrap must include saved Coverage Output in its request snapshot'
);
assert(
  bootstrap.includes('getSavedCoverageForDate_(today, dayCode)'),
  'web bootstrap must restore the saved plan for the selected date'
);
assert(
  bootstrap.includes('dayStaffSchedules'),
  'web bootstrap must batch selected-day staff schedules for local absence editing'
);

const webGenerate = functionBody(sources['code.gs'], 'webGenerateCoverage');
assert(
  webGenerate.includes('persistPreview: false'),
  'web Generate must explicitly disable shared preview persistence'
);

const combinedHandoutWeb = functionBody(sources['code.gs'], 'webSaveCoverageAndCreateHandout');
assert(
  combinedHandoutWeb.indexOf('saveCoveragePlan({ rows: planRows })') <
    combinedHandoutWeb.indexOf('createCoverageHandoutPackage_'),
  'Create Handout must persist the reviewed plan before generating documents'
);
assert(
  index.includes("gas('webSaveCoverageAndCreateHandout',S.plan)"),
  'browser Create Handout must use the save-first handout endpoint'
);
assert(
  !index.includes("async function createHandout(){try{const r=await gas('webCreateHandoutFromRows',S.plan)"),
  'browser Create Handout must not use the print-only endpoint'
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

const savedCoverageReader = functionBodyAny('getSavedCoverageForDate_');
assert(
  savedCoverageReader.includes("readSheetObjects_('Coverage Output')"),
  'saved-plan reload must read durable Coverage Output rather than _Preview'
);
assert(
  savedCoverageReader.includes('normalizeDateKey_(row.Date) === dateKey'),
  'saved-plan reload must be date-scoped'
);
assert(
  savedCoverageReader.includes("String(row.Day || '').trim() === dayCode"),
  'saved-plan reload must preserve the selected day code'
);
assert(
  savedCoverageReader.includes("String(row.Status || '').trim() === 'Assigned'"),
  'saved-plan reload must rebuild its assignment summary from saved rows'
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
assert(
  generatePreview.includes('payload.persistPreview === true'),
  'Generate may write _Preview only when explicitly requested by the spreadsheet-menu workflow'
);
assert(
  generatePreview.includes('buildManualChoiceWorkspace_'),
  'Generate must prepare the bounded browser reassignment workspace'
);

const liveContext = functionBodyAny('buildCoverageLiveContext_');
assert(
  liveContext.includes('teacherScheduleRowsForDate_'),
  'save/manual validation must use the weekday schedule cache'
);

const dayCacheReader = functionBodyAny('readTeacherScheduleDayCached_');
assert(
  dayCacheReader.includes('rebuildTeacherScheduleDayCache_'),
  'weekday schedule cache must self-heal when stale'
);

const dayCacheBuilder = functionBodyAny('rebuildTeacherScheduleDayCache_');
assert(
  dayCacheBuilder.includes("Math.min(8, sheet.getLastColumn())") ||
  combined.includes("Math.min(8, sheet.getLastColumn())"),
  'schedule cache source read must exclude the unused Lead/Co formula column'
);
assert(
  dayCacheBuilder.includes('normalizeTeacherScheduleRow_'),
  'weekday caches must store prepared/derived schedule data'
);

const classGradeLookup = functionBody(sources['class-schedule.gs'], 'classScheduleGradeFor_');
assert(
  classGradeLookup.includes('classScheduleIndexForDay_'),
  'runtime grade inference must use the weekday Class Schedule index'
);
const classDayReader = functionBody(sources['class-schedule.gs'], 'readClassScheduleDayCache_');
assert(
  classDayReader.includes('rebuildClassScheduleDayCache_'),
  'weekday Class Schedule cache must self-heal when stale'
);
const classDayBuilder = functionBody(sources['class-schedule.gs'], 'rebuildClassScheduleDayCacheUnlocked_');
assert(
  classDayBuilder.includes('Resolved_Staff') || sources['class-schedule.gs'].includes("'Resolved_Staff'"),
  'Class Schedule cache must pre-resolve teacher names'
);

const poolRefresh = functionBodyAny('ensureFieldTripCoveragePoolFresh_');
assert(
  poolRefresh.includes('withCoverageLock_'),
  'full field-trip pool rebuild must be serialized'
);
assert(
  poolRefresh.includes('fieldTripCoveragePoolIsDirty_'),
  'full field-trip pool rebuild must re-check freshness'
);

const datePoolRefresh = functionBodyAny('fieldTripCoveragePoolRowsForDate_');
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

const fastPoolRead = functionBodyAny('readFieldTripCoveragePoolRowsForDateFast_');
assert(
  datePoolRefresh.includes('readFieldTripCoveragePoolRowsForDateFast_'),
  'fresh field-trip pool reads must be date-scoped'
);
assert(
  fastPoolRead.includes("getRange(2, dateColumn + 1"),
  'date-scoped field-trip pool reader must scan only the Date column before reading matching rows'
);

const targetedPoolRebuild = functionBodyAny('rebuildFieldTripCoveragePoolForDate_');
assert(
  targetedPoolRebuild.includes('getFieldTripsForDate_'),
  'targeted pool rebuild must limit itself to field trips active on the requested date'
);
assert(
  targetedPoolRebuild.includes('markFieldTripCoveragePoolDateFresh_'),
  'targeted pool rebuild must mark only the requested date revision fresh'
);

const fieldTripPackage = functionBody(sources['handout.gs'], 'createCoverageHandoutPackage_');
assert(
  !fieldTripPackage.includes('getFieldTripFormTemplateFile_'),
  'field-trip handouts must not open the old Drive template at runtime'
);
assert(
  !sources['handout.gs'].includes('makeCopy(outputName, folder)'),
  'field-trip handouts must be generated rather than copied from a template'
);
assert(
  !sources['handout.gs'].includes('removeTrailingUnusedFieldTripForm_'),
  'generated field-trip handouts must not depend on destructive template trimming'
);

const fieldTripBuilder = functionBody(sources['handout.gs'], 'createFieldTripCoverageFormDocs_');
assert(
  fieldTripBuilder.includes('DocumentApp.create(outputName)'),
  'field-trip handouts must start from a fresh Google Doc'
);
assert(
  fieldTripBuilder.includes('appendGeneratedFieldTripForm_'),
  'field-trip handouts must build each required form explicitly'
);
assert(
  fieldTripBuilder.includes('body.getTables().length !== formUnits.length'),
  'field-trip handout generation must verify that every form produced exactly one table'
);
assert(
  fieldTripBuilder.includes('removeGeneratedFieldTripStarterParagraph_(body)'),
  'field-trip handouts must remove the empty starter paragraph from a newly created Doc'
);

const starterCleanup = functionBody(sources['handout.gs'], 'removeGeneratedFieldTripStarterParagraph_');
assert(
  starterCleanup.includes('body.getChild(0)') &&
  starterCleanup.includes('body.removeChild(first)'),
  'starter cleanup must target only the initial empty paragraph'
);

const fieldTripForm = functionBody(sources['handout.gs'], 'appendGeneratedFieldTripForm_');
assert(
  fieldTripForm.includes("'FIELD TRIP COVERAGE FORM"),
  'generated field-trip forms must include the approved form heading'
);
assert(
  fieldTripForm.includes("'SPECIAL INSTRUCTIONS'") &&
  fieldTripForm.includes("'WITH'"),
  'generated field-trip forms must retain the approved five-column table'
);

const destinationHelper = functionBody(sources['handout.gs'], 'fieldTripHandoutDestination_');
assert(
  destinationHelper.includes('trip && trip.destination'),
  'generated field-trip forms must retain the destination helper'
);
const resolveFieldTripDestination = new Function(
  destinationHelper + '; return fieldTripHandoutDestination_;'
)();
assert.strictEqual(
  resolveFieldTripDestination({ destination: 'National Gallery', name: '6th Grade - Other' }),
  'National Gallery',
  'field-trip destination helper must prefer the explicit Destination field'
);
assert.strictEqual(
  resolveFieldTripDestination({ destination: '', name: '6th Grade - National Gallery' }),
  'National Gallery',
  'field-trip destination helper must preserve the legacy event-name fallback'
);

const fieldTripPage = functionBody(sources['handout.gs'], 'formatGeneratedFieldTripHandoutPage_');
assert(
  fieldTripPage.includes('body.setPageWidth(612)') &&
  fieldTripPage.includes('body.setPageHeight(792)') &&
  fieldTripPage.includes('body.setMarginTop(36)'),
  'generated field-trip forms must retain the approved US Letter portrait geometry'
);

assert(
  index.includes('S.manualWorkspace=result.manualWorkspace||null'),
  'browser must retain the generated manual-choice workspace'
);
assert(
  index.includes('.sub-row:hover,.sub-row:focus-visible') &&
  index.includes('.sub-time{'),
  'By Sub rows must visibly advertise that coverage blocks are interactive'
);
assert(
  index.includes('title="Click to change coverage"') &&
  index.includes('tabindex="0" role="button"'),
  'clickable coverage blocks must expose an interaction cue and keyboard focus'
);
assert(
  index.includes("$('planBody').addEventListener('keydown'") &&
  index.includes("e.key!=='Enter'&&e.key!==' '"),
  'coverage blocks must support Enter/Space activation'
);
assert(
  index.includes('.tl-blk:hover,.tl-blk:focus-visible') &&
  index.includes('.plan-tbl tbody tr:hover td'),
  'timeline and table coverage blocks must have hover feedback'
);

assert(
  index.includes('(data.dayStaffSchedules||[]).forEach'),
  'browser must hydrate its local staff schedule cache from bootstrap'
);
const saveBlockStart = index.indexOf('function saveBlock()');
const saveBlockEnd = index.indexOf('\nasync function generate()', saveBlockStart);
const saveBlockBody = index.slice(saveBlockStart, saveBlockEnd);
assert(
  !saveBlockBody.includes('webValidateManualCoverageAssignment'),
  'manual assignment modal Save must not make a redundant validation RPC'
);
assert(
  saveBlockBody.includes('S.manualWorkspace=null'),
  'manual changes must invalidate precomputed choices for subsequent blocks'
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
  combined + `
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
