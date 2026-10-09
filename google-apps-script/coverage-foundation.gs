// Shared schemas, request caches, sheet I/O, config, and low-level utilities.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

const HEADER_ALIASES = {
  'Teacher Schedule': {
    Staff_Name: ['Staff_Name', 'Teacher', 'Teacher_Name', 'Name'],
    Role: ['Role'],
    Term: ['Term', 'Semester', 'Schedule_Term', 'Schedule Term'],
    Day: ['Day'],
    Start: ['Start', 'Start_Time', 'Start Time'],
    End: ['End', 'End_Time', 'End Time'],
    Class: ['Class', 'Class_Name', 'Class Name', 'Current_Class', 'Current Class', 'Section', 'Group', 'Course', 'Course_Name', 'Course Name', 'Activity'],
    Grade: ['Grade'],
    Subject: ['Subject', 'Course', 'Course_Name', 'Course Name', 'Activity', 'Department', 'Content_Area', 'Content Area'],
    Assignment_Type: ['Assignment_Type', 'Assignment Type', 'Type'],
    Room: ['Room'],
    Needs_Coverage_If_Absent: ['Needs_Coverage_If_Absent', 'Needs Coverage If Absent'],
    Cover_Eligible_This_Block: ['Cover_Eligible_This_Block', 'Available_To_Cover', 'Coverage_Eligible', 'Cover Eligible This Block']
  },
  'Coverage Staff': {
    Name: ['Name', 'Substitute_Name', 'Staff_Name', 'Teacher'],
    Role: ['Role'],
    Coverage_Tier: ['Coverage_Tier', 'Tier', 'Priority'],
    Can_Cover_All_Day: ['Can_Cover_All_Day', 'All_Day', 'Can Cover All Day'],
    Active_Today: ['Active_Today', 'Active'],
    Available_Days: ['Available_Days', 'Available Days', 'Days', 'Days_Available'],
    Default_Start: ['Default_Start', 'Default Start', 'Available_From', 'Available From'],
    Default_End: ['Default_End', 'Default End', 'Available_Until', 'Available Until'],
    Allowed_Grades: ['Allowed_Grades', 'Grades'],
    Allowed_Subjects: ['Allowed_Subjects', 'Subjects'],
    Allowed_Assignment_Types: ['Allowed_Assignment_Types', 'Assignment_Types'],
    Max_Blocks_Per_Day: ['Max_Blocks_Per_Day'],
    Max_Teachers_Per_Day: ['Max_Teachers_Per_Day'],
    Can_Be_Split_Across_Teachers: ['Can_Be_Split_Across_Teachers', 'Can Split'],
    Notes: ['Notes']
  },
  'Substitute Availability': {
    Date: ['Date'],
    Day: ['Day'],
    Name: ['Name', 'Substitute_Name', 'Staff_Name', 'Teacher'],
    Available: ['Available', 'Active', 'Active_Today', 'Available_Today'],
    Start: ['Start', 'Start_Time', 'Available_From', 'Available From'],
    End: ['End', 'End_Time', 'Available_Until', 'Available Until'],
    Notes: ['Notes']
  },
  'Daily Absences': {
    Date: ['Date'],
    Day: ['Day'],
    Staff_Name: ['Staff_Name', 'Teacher', 'Name'],
    Absence_Type: ['Absence_Type', 'Type'],
    Start_Override: ['Start_Override', 'Start', 'Start Time'],
    End_Override: ['End_Override', 'End', 'End Time'],
    Notes: ['Notes'],
    Preferred_Coverage: ['Preferred_Coverage', 'Preferred Coverage', 'Assigned_To']
  },
  'Field Trips': {
    Event_ID: ['Event_ID', 'Event ID', 'ID'],
    Name: ['Name', 'Event_Name', 'Event Name'],
    Destination: ['Destination', 'Trip_Destination', 'Trip Destination'],
    Date: ['Date', 'Start_Date', 'Start Date'],
    End_Date: ['End_Date', 'End Date'],
    Start: ['Start', 'Start_Time', 'Start Time'],
    End: ['End', 'End_Time', 'End Time'],
    Grades: ['Grades', 'Students_Away', 'Students Away'],
    Staff: ['Staff', 'Teachers', 'Staff_Away', 'Staff Away'],
    Notes: ['Notes']
  },
  'Field Trip Coverage Pool': {
    Event_ID: ['Event_ID', 'Event ID'],
    Event_Name: ['Event_Name', 'Event Name'],
    Date: ['Date'],
    Day: ['Day'],
    Absent_Staff: ['Absent_Staff', 'Absent Staff'],
    Start: ['Start'],
    End: ['End'],
    Class: ['Class'],
    Grade: ['Grade'],
    Subject: ['Subject'],
    Assignment_Type: ['Assignment_Type', 'Assignment Type'],
    Room: ['Room'],
    Candidate: ['Candidate', 'Coverage Person', 'Staff'],
    Candidate_Role: ['Candidate_Role', 'Candidate Role'],
    Candidate_Tier: ['Candidate_Tier', 'Candidate Tier', 'Tier'],
    Source: ['Source'],
    Baseline_Score: ['Baseline_Score', 'Baseline Score'],
    Reason: ['Reason'],
    Enabled: ['Enabled', 'Use', 'Active'],
    Priority_Adjustment: ['Priority_Adjustment', 'Priority Adjustment'],
    Built_At: ['Built_At', 'Built At']
  },
  'Coverage Output': {
    Date: ['Date'],
    Day: ['Day'],
    Event_ID: ['Event_ID', 'Event ID', 'Field_Trip_ID', 'Field Trip ID'],
    Start: ['Start'],
    End: ['End'],
    Absent_Staff: ['Absent_Staff'],
    Class: ['Class'],
    Grade: ['Grade'],
    Subject: ['Subject'],
    Assignment_Type: ['Assignment_Type'],
    Room: ['Room'],
    Assigned_Coverage: ['Assigned_Coverage'],
    Coverage_Mode: ['Coverage_Mode'],
    Coverage_Tier_Used: ['Coverage_Tier_Used'],
    Status: ['Status'],
    Notes: ['Notes']
  },
  '_Preview': {
    Date: ['Date'],
    Day: ['Day'],
    Event_ID: ['Event_ID', 'Event ID', 'Field_Trip_ID', 'Field Trip ID'],
    Start: ['Start'],
    End: ['End'],
    Absent_Staff: ['Absent_Staff'],
    Class: ['Class'],
    Grade: ['Grade'],
    Subject: ['Subject'],
    Assignment_Type: ['Assignment_Type'],
    Room: ['Room'],
    Assigned_Coverage: ['Assigned_Coverage'],
    Coverage_Mode: ['Coverage_Mode'],
    Coverage_Tier_Used: ['Coverage_Tier_Used'],
    Status: ['Status'],
    Notes: ['Notes']
  },
  'Config': {
    Setting: ['Setting'],
    Value: ['Value'],
    Description: ['Description']
  }
};

// ── Request-local read cache + lightweight performance tracing ─────────────
// Spreadsheet service calls dominate Apps Script latency. Keep one canonical
// copy of each sheet read for the lifetime of a server request, then invalidate
// it immediately after any write. Config also uses a short Script Cache entry;
// Teacher Schedule deliberately stays request-local so schedule dropdowns and
// eligibility checks always receive the exact live Sheet values/data types.
const COVERAGE_PERSISTENT_CACHE_TTL_SECONDS_ = 60;
const COVERAGE_PERSISTENT_CACHE_MAX_CHARS_ = 85000;
const COVERAGE_PERSISTENT_CACHE_KEYS_ = {
  'Config': 'coverage:v1:config',
  'Coverage Staff': 'coverage:v1:coverage-staff',
  'Substitutes': 'coverage:v1:substitutes',
  'Substitute Availability': 'coverage:v1:sub-availability',
  'Daily Absences': 'coverage:v1:daily-absences',
  'Field Trips': 'coverage:v1:field-trips'
};

const FIELD_TRIP_COVERAGE_POOL_SHEET_ = 'Field Trip Coverage Pool';
const FIELD_TRIP_COVERAGE_POOL_DIRTY_PROPERTY_ = 'FIELD_TRIP_COVERAGE_POOL_DIRTY';
const FIELD_TRIP_COVERAGE_POOL_REVISION_PROPERTY_ = 'FIELD_TRIP_COVERAGE_POOL_REVISION';
const FIELD_TRIP_COVERAGE_POOL_DATE_REVISION_PREFIX_ = 'FIELD_TRIP_COVERAGE_POOL_DATE_REVISION:';

const TEACHER_SCHEDULE_CACHE_DAYS_ = ['M', 'T', 'W', 'R', 'F'];
const TEACHER_SCHEDULE_CACHE_SHEET_PREFIX_ = '_Schedule_';
const TEACHER_SCHEDULE_CACHE_REVISION_PROPERTY_ = 'TEACHER_SCHEDULE_CACHE_REVISION';
const TEACHER_SCHEDULE_CACHE_DAY_REVISION_PREFIX_ = 'TEACHER_SCHEDULE_CACHE_DAY_REVISION:';
const TEACHER_SCHEDULE_CACHE_HEADERS_ = [
  'Teacher',
  'Term',
  'Day',
  'Start',
  'End',
  'Class',
  'Subject',
  'Room',
  'Grade',
  'Assignment_Type',
  'Needs_Coverage_If_Absent',
  'Cover_Eligible_This_Block'
];

let COVERAGE_REQUEST_SHEET_CACHE_ = {};
let COVERAGE_REQUEST_METRICS_ = null;
let COVERAGE_PERF_TRACE_ = null;

function beginCoverageRequest_(name) {
  COVERAGE_REQUEST_SHEET_CACHE_ = {};
  COVERAGE_CONFIG_CACHE_ = null;
  COVERAGE_REQUEST_METRICS_ = {
    request: String(name || 'coverage-request'),
    sheetReads: 0,
    requestCacheHits: 0,
    persistentCacheHits: 0,
    sheetWrites: 0
  };
  COVERAGE_PERF_TRACE_ = {
    name: String(name || 'coverage-request'),
    started: Date.now(),
    marks: []
  };
}

function coveragePerfMark_(label) {
  if (!COVERAGE_PERF_TRACE_) return;
  COVERAGE_PERF_TRACE_.marks.push({
    label: String(label || ''),
    elapsedMs: Date.now() - COVERAGE_PERF_TRACE_.started
  });
}

function endCoverageRequest_() {
  if (!COVERAGE_PERF_TRACE_) return null;
  const result = {
    request: COVERAGE_PERF_TRACE_.name,
    elapsedMs: Date.now() - COVERAGE_PERF_TRACE_.started,
    marks: COVERAGE_PERF_TRACE_.marks.slice(),
    io: COVERAGE_REQUEST_METRICS_ || {}
  };
  console.log('[Coverage Performance] ' + JSON.stringify(result));
  COVERAGE_PERF_TRACE_ = null;
  return result;
}

function incrementCoverageMetric_(name) {
  if (!COVERAGE_REQUEST_METRICS_) return;
  COVERAGE_REQUEST_METRICS_[name] = Number(COVERAGE_REQUEST_METRICS_[name] || 0) + 1;
}

function persistentCoverageCacheKey_(sheetName) {
  const base = COVERAGE_PERSISTENT_CACHE_KEYS_[String(sheetName || '')] || '';
  if (!base) return '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workbookId = ss ? ss.getId() : 'unbound';
  return base + ':' + workbookId;
}

function persistentCacheSafeRows_(rows) {
  return (rows || []).map(row => {
    const out = {};
    Object.keys(row || {}).forEach(key => {
      const value = row[key];
      if (Object.prototype.toString.call(value) === '[object Date]') {
        const year = value.getFullYear();
        out[key] = year <= 1900 ? timeToDisplay_(value) : normalizeDateKey_(value);
      } else {
        out[key] = value;
      }
    });
    return out;
  });
}

function readPersistentSheetObjectsCache_(sheetName) {
  const key = persistentCoverageCacheKey_(sheetName);
  if (!key) return null;
  try {
    const cached = CacheService.getScriptCache().get(key);
    if (!cached) return null;
    const parsed = JSON.parse(cached);
    if (!Array.isArray(parsed)) return null;
    incrementCoverageMetric_('persistentCacheHits');
    return parsed;
  } catch (error) {
    return null;
  }
}

function writePersistentSheetObjectsCache_(sheetName, rows) {
  const key = persistentCoverageCacheKey_(sheetName);
  if (!key) return;
  try {
    const json = JSON.stringify(persistentCacheSafeRows_(rows));
    if (json.length > COVERAGE_PERSISTENT_CACHE_MAX_CHARS_) return;
    CacheService.getScriptCache().put(key, json, COVERAGE_PERSISTENT_CACHE_TTL_SECONDS_);
  } catch (error) {
    // Cache is an optimization only. Never let it block operational work.
  }
}

function invalidateCoverageSheetCache_(sheetName) {
  const name = String(sheetName || '');
  delete COVERAGE_REQUEST_SHEET_CACHE_[name];
  if (name === 'Config') COVERAGE_CONFIG_CACHE_ = null;
  const key = persistentCoverageCacheKey_(name);
  if (key) {
    try { CacheService.getScriptCache().remove(key); } catch (error) {}
  }
}

function primeCoverageRequestSnapshot_(sheetNames) {
  (sheetNames || []).forEach(name => readSheetObjects_(name));
}

// ── Write safety ──────────────────────────────────────────────────────────
// Several front-office users can use the web app at once, and most saves are
// read-modify-write over a whole sheet. Without a lock, two overlapping saves
// can silently drop one person's change (or, for field trips, both pass the
// duplicate check and create two events). Nested calls within one execution
// reuse the lock already held.
let COVERAGE_LOCK_DEPTH_ = 0;

function withCoverageLock_(fn) {
  if (COVERAGE_LOCK_DEPTH_ > 0) return fn();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another Coverage Scheduler change is still being saved. Wait a few seconds and try again.');
  }
  COVERAGE_LOCK_DEPTH_++;
  try {
    return fn();
  } finally {
    COVERAGE_LOCK_DEPTH_--;
    lock.releaseLock();
  }
}

// Replaces a sheet's data rows without an empty window: the new rows are
// written first and only leftover old rows below them are cleared. The old
// clear-then-write pattern would leave the sheet empty (losing every date's
// absences or saved coverage) if the write failed partway.
function rewriteSheetRows_(sheetName, headers, rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet: ' + sheetName);
  const previousLastRow = sheet.getLastRow();
  if (rows.length) writeObjectsToSheet_(sheetName, headers, rows, false);
  const firstStaleRow = rows.length + 2;
  const lastCol = sheet.getLastColumn();
  if (previousLastRow >= firstStaleRow && lastCol > 0) {
    sheet.getRange(firstStaleRow, 1, previousLastRow - firstStaleRow + 1, lastCol).clearContent();
    incrementCoverageMetric_('sheetWrites');
  }
  invalidateCoverageSheetCache_(sheetName);
}

// A generated plan is only valid for the absences and field trips it was
// built from. When those change, drop the stored preview for affected dates so
// reopening the app shows "no plan" instead of quietly restoring a stale one.
function invalidatePreviewForDateRange_(startDate, endDate) {
  const start = normalizeDateKey_(startDate);
  const end = normalizeDateKey_(endDate || startDate);
  if (!start || !end) return;
  const affected = readSheetObjects_('_Preview').some(row => {
    const key = normalizeDateKey_(row.Date);
    return key && key >= start && key <= end;
  });
  if (affected) clearSheetDataKeepingHeader_('_Preview');
}

function readSheetObjects_(sheetName) {
  const name = String(sheetName || '');
  if (Object.prototype.hasOwnProperty.call(COVERAGE_REQUEST_SHEET_CACHE_, name)) {
    incrementCoverageMetric_('requestCacheHits');
    return COVERAGE_REQUEST_SHEET_CACHE_[name];
  }

  const persistent = readPersistentSheetObjectsCache_(name);
  if (persistent !== null) {
    COVERAGE_REQUEST_SHEET_CACHE_[name] = persistent;
    return persistent;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) {
    COVERAGE_REQUEST_SHEET_CACHE_[name] = [];
    return COVERAGE_REQUEST_SHEET_CACHE_[name];
  }

  incrementCoverageMetric_('sheetReads');
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const aliasMap = HEADER_ALIASES[name] || (name === 'Substitutes' ? HEADER_ALIASES['Coverage Staff'] : {});
  const columnByCanonical = {};

  Object.keys(aliasMap).forEach(canonical => {
    const aliases = aliasMap[canonical];
    let foundIndex = -1;
    aliases.some(alias => {
      foundIndex = headers.indexOf(alias);
      return foundIndex !== -1;
    });
    columnByCanonical[canonical] = foundIndex;
  });

  const rows = values.slice(1).map(row => {
    const out = {};
    Object.keys(columnByCanonical).forEach(canonical => {
      const foundIndex = columnByCanonical[canonical];
      out[canonical] = foundIndex === -1 ? '' : row[foundIndex];
    });
    return out;
  }).filter(obj => Object.keys(obj).some(key => String(obj[key] || '').trim() !== ''));

  COVERAGE_REQUEST_SHEET_CACHE_[name] = rows;
  writePersistentSheetObjectsCache_(name, rows);
  return rows;
}

function writeObjectsToSheet_(sheetName, headers, rows, append) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet: ' + sheetName);

  // Existing managed sheets can gain new columns over time. setup.gs appends
  // a missing header rather than reordering old columns, so always write using
  // the sheet's actual header order. This keeps migrated workbooks aligned.
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const actualHeaders = sheet.getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(value => String(value || '').trim());

  while (actualHeaders.length && !actualHeaders[actualHeaders.length - 1]) {
    actualHeaders.pop();
  }

  const writeHeaders = actualHeaders.length ? actualHeaders : headers;
  const startRow = append ? sheet.getLastRow() + 1 : 2;
  const values = rows.map(row =>
    writeHeaders.map(header => row[header] != null ? row[header] : '')
  );

  if (!values.length) return;
  sheet.getRange(startRow, 1, values.length, writeHeaders.length).setValues(values);
  incrementCoverageMetric_('sheetWrites');
  invalidateCoverageSheetCache_(sheetName);
}

function getConfigMap_() {
  const rows = readSheetObjects_('Config');
  const map = {};
  rows.forEach(row => {
    const key = String(row.Setting || row['Setting'] || '').trim();
    const value = row.Value != null ? row.Value : row['Value'];
    if (key) map[key] = value;
  });
  return map;
}

// ── Config ────────────────────────────────────────────────────────────────
// Read once per server call. setup.gs clears this after seeding Config.
let COVERAGE_CONFIG_CACHE_ = null;

function coverageConfig_() {
  if (!COVERAGE_CONFIG_CACHE_) {
    try {
      COVERAGE_CONFIG_CACHE_ = getConfigMap_();
    } catch (err) {
      // No active workbook yet (e.g. before the web app attaches to it).
      return {};
    }
  }
  return COVERAGE_CONFIG_CACHE_;
}

// Script_Time_Zone, when it is a real zone name; otherwise the project zone.
// An unrecognized name is ignored rather than trusted, because Apps Script
// silently treats unknown zones as GMT, which would shift every time.
function coverageTimeZone_() {
  const configured = String(coverageConfig_().Script_Time_Zone || '').trim();
  if (/^(?:[A-Za-z_]+(?:\/[A-Za-z0-9_+\-]+)+|UTC|GMT|Etc\/[A-Za-z0-9+\-]+)$/.test(configured)) return configured;
  return Session.getScriptTimeZone();
}

function lunchCoverageAllowed_() {
  return normalizeYesNo_(coverageConfig_().Use_Lunch_For_Coverage, false);
}

// Availability_Override_Mode: DATE (default) applies Substitute Availability
// rows for the selected date; OFF ignores them and uses Coverage Staff
// defaults only.
function availabilityOverridesEnabled_() {
  const mode = String(coverageConfig_().Availability_Override_Mode || 'DATE').trim().toUpperCase();
  return ['OFF', 'NONE', 'IGNORE', 'DEFAULTS'].indexOf(mode) === -1;
}

function parseDailyLimit_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return null;
  const num = Number(raw);
  return Number.isInteger(num) && num >= 1 ? num : null;
}

// A person's own Max_* value wins; blank (or unreadable) falls back to the
// Config default; a blank default means no limit.
function resolveDailyLimit_(rawValue, config, defaultKey) {
  const own = parseDailyLimit_(rawValue);
  if (own != null) return own;
  const fallback = parseDailyLimit_((config || coverageConfig_())[defaultKey]);
  return fallback != null ? fallback : Infinity;
}

// True when two absences describe the same window for the same person.
function sameAbsenceWindow_(a, b) {
  if (String(a.staffName || '').trim() !== String(b.staffName || '').trim()) return false;
  if (!!a.emergency !== !!b.emergency) return false;
  const shape = x => ({
    absenceType: x.allDay === true || String(x.absenceType || '').trim() === 'Full Day' || x.emergency ? 'Full Day' : 'Partial Day',
    startOverride: x.startOverride,
    endOverride: x.endOverride
  });
  const wa = absenceWindowMinutes_(shape(a));
  const wb = absenceWindowMinutes_(shape(b));
  return wa.startMinutes === wb.startMinutes && wa.endMinutes === wb.endMinutes;
}

function normalizeYesNo_(value, defaultValue) {
  if (value === '' || value == null) return !!defaultValue;
  const str = String(value).trim().toLowerCase();
  if (['yes', 'y', 'true', '1'].indexOf(str) !== -1) return true;
  if (['no', 'n', 'false', '0'].indexOf(str) !== -1) return false;
  return !!defaultValue;
}

function normalizeDateKey_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, coverageTimeZone_(), 'yyyy-MM-dd');
  }
  const str = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const parsed = new Date(str);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, coverageTimeZone_(), 'yyyy-MM-dd');
  }
  return str;
}

function guessDayCodeFromDate_(dateStr) {
  const d = new Date(dateStr + 'T12:00:00');
  const day = d.getDay();
  return ['U', 'M', 'T', 'W', 'R', 'F', 'S'][day] || '';
}

function timeToMinutes_(value) {
  if (value == null || value === '') return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    // Read the clock time in the same zone used to display it, so parsing and
    // display can never disagree when Script_Time_Zone is set.
    const zone = coverageTimeZone_();
    if (zone === Session.getScriptTimeZone()) return value.getHours() * 60 + value.getMinutes();
    const parts = Utilities.formatDate(value, zone, 'H:mm').split(':');
    return Number(parts[0]) * 60 + Number(parts[1]);
  }
  const str = String(value).trim();
  const match = str.match(/^(\d{1,2}):(\d{2})(?:\s*([AP]M))?$/i);
  if (match) {
    let hour = Number(match[1]);
    const minute = Number(match[2]);
    const meridian = (match[3] || '').toUpperCase();
    if (meridian === 'PM' && hour < 12) hour += 12;
    if (meridian === 'AM' && hour === 12) hour = 0;
    return hour * 60 + minute;
  }
  return null;
}

function displayTimeToMinutes_(value) {
  return timeToMinutes_(value);
}

function timeToDisplay_(value) {
  if (value == null || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, coverageTimeZone_(), 'h:mm a');
  }
  const mins = timeToMinutes_(value);
  return mins == null ? String(value) : minutesToDisplay_(mins);
}

function minutesToDisplay_(mins) {
  if (mins == null) return '';
  let hours = Math.floor(mins / 60);
  const minutes = mins % 60;
  const meridian = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return hours + ':' + Utilities.formatString('%02d', minutes) + ' ' + meridian;
}

function timesOverlap_(startA, endA, startB, endB) {
  return startA < endB && startB < endA;
}
