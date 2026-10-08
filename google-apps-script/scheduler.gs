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
  'Config': 'coverage:v1:config'
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

function ensureFieldTripsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Field Trips');
  if (!sheet) {
    sheet = ss.insertSheet('Field Trips');
    sheet.getRange(1, 1, 1, SHEET_SCHEMAS['Field Trips'].headers.length)
      .setValues([SHEET_SCHEMAS['Field Trips'].headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_SCHEMAS['Field Trips'].headers.length)
      .setFontWeight('bold')
      .setBackground('#d9eaf7');
  } else {
    ensureHeaderRow_(sheet, SHEET_SCHEMAS['Field Trips'].headers);
  }
  return sheet;
}

function normalizeGradeKey_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const compact = raw.toLowerCase().replace(/[.\s_-]+/g, '');

  if (compact.indexOf('beginner') === 0 || compact.indexOf('beg') === 0) return 'Beg';
  if (compact.indexOf('prekindergarten') === 0 || compact.indexOf('prekind') === 0 || compact.indexOf('prek') === 0) return 'PreK';
  if (/^k[a-z]?$/.test(compact) || compact.indexOf('kindergarten') === 0) return 'K';

  const words = {
    first: '1',
    second: '2',
    third: '3',
    fourth: '4',
    fifth: '5',
    sixth: '6',
    seventh: '7',
    eighth: '8'
  };
  const lower = raw.toLowerCase();
  for (const word in words) {
    if (lower.indexOf(word) !== -1) return words[word];
  }

  const numeric = raw.match(/\d+/);
  if (numeric) return String(Number(numeric[0]));

  return raw;
}

function splitFieldTripStaffList_(value) {
  if (Array.isArray(value)) return value.map(v => String(v || '').trim()).filter(Boolean);

  // Staff names are commonly stored as "Last, First". Commas are part of
  // the person's name, not a list separator. Persisted field-trip staff are
  // separated with pipes, and semicolons are accepted as a manual fallback.
  return String(value || '')
    .split(/\s*[|;]\s*/)
    .map(v => v.trim())
    .filter(Boolean);
}

function normalizeFieldTripGrades_(value) {
  const seen = {};
  const values = Array.isArray(value)
    ? value
    : String(value || '').split(/\s*[|,;]\s*/);

  return values
    .map(normalizeGradeKey_)
    .filter(grade => {
      if (!grade || seen[grade]) return false;
      seen[grade] = true;
      return true;
    });
}

function normalizeFieldTripRow_(row) {
  const startDate = normalizeDateKey_(row.Date);
  const endDate = normalizeDateKey_(row.End_Date) || startDate;
  return {
    eventId: String(row.Event_ID || '').trim(),
    name: String(row.Name || '').trim(),
    destination: String(row.Destination || '').trim(),
    date: startDate,
    startDate: startDate,
    endDate: endDate,
    start: timeToDisplay_(row.Start),
    end: timeToDisplay_(row.End),
    grades: normalizeFieldTripGrades_(row.Grades),
    staffNames: splitFieldTripStaffList_(row.Staff),
    notes: String(row.Notes || '').trim()
  };
}

function getFieldTripsInRange_(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Field Trips')) return [];

  const start = normalizeDateKey_(startDate);
  const end = normalizeDateKey_(endDate || startDate);
  return readSheetObjects_('Field Trips')
    .map(normalizeFieldTripRow_)
    .filter(row =>
      row.eventId &&
      row.startDate &&
      row.endDate &&
      (!start || row.endDate >= start) &&
      (!end || row.startDate <= end)
    )
    .sort((a, b) =>
      a.startDate.localeCompare(b.startDate) ||
      a.start.localeCompare(b.start) ||
      a.name.localeCompare(b.name)
    );
}

function fieldTripForDate_(trip, date) {
  const key = normalizeDateKey_(date);
  if (!trip || !key || key < trip.startDate || key > trip.endDate) return null;

  const sameDay = trip.startDate === trip.endDate;
  let effectiveStartMinutes = 0;
  let effectiveEndMinutes = 1440;
  let allDayForDate = !sameDay && key !== trip.startDate && key !== trip.endDate;

  if (sameDay) {
    effectiveStartMinutes = displayTimeToMinutes_(trip.start);
    effectiveEndMinutes = displayTimeToMinutes_(trip.end);
    allDayForDate = false;
  } else if (key === trip.startDate) {
    effectiveStartMinutes = displayTimeToMinutes_(trip.start);
    effectiveEndMinutes = 1440;
    allDayForDate = false;
  } else if (key === trip.endDate) {
    effectiveStartMinutes = 0;
    effectiveEndMinutes = displayTimeToMinutes_(trip.end);
    allDayForDate = false;
  }

  return Object.assign({}, trip, {
    activeDate: key,
    effectiveStartMinutes: effectiveStartMinutes,
    effectiveEndMinutes: effectiveEndMinutes,
    allDayForDate: allDayForDate
  });
}

function getFieldTripsForDate_(date) {
  const key = normalizeDateKey_(date);
  return getFieldTripsInRange_(key, key)
    .map(trip => fieldTripForDate_(trip, key))
    .filter(Boolean);
}

function fieldTripIdentityKey_(trip) {
  trip = trip || {};
  const startDate = normalizeDateKey_(trip.startDate || trip.date || trip.Date);
  const endDate = normalizeDateKey_(trip.endDate || trip.End_Date || trip.startDate || trip.date || trip.Date);
  const name = String(trip.name || trip.Name || '').trim().toLowerCase();
  const destination = String(trip.destination || trip.Destination || '').trim().toLowerCase();
  const start = timeToDisplay_(trip.start || trip.Start);
  const end = timeToDisplay_(trip.end || trip.End);
  const grades = normalizeFieldTripGrades_(trip.grades || trip.Grades).slice().sort().join('|');
  const staff = splitFieldTripStaffList_(trip.staffNames || trip.staff || trip.Staff)
    .map(name => String(name || '').trim().toLowerCase())
    .filter(Boolean)
    .sort()
    .join('|');

  return [
    startDate || '',
    endDate || '',
    start || '',
    end || '',
    name,
    destination,
    grades,
    staff
  ].join('::');
}

function saveFieldTrip_(payload) {
  return withCoverageLock_(() => saveFieldTripUnlocked_(payload));
}

function saveFieldTripUnlocked_(payload) {
  payload = payload || {};
  const startDate = normalizeDateKey_(payload.startDate || payload.date);
  const endDate = normalizeDateKey_(payload.endDate || payload.startDate || payload.date);
  const name = String(payload.name || '').trim();
  const destination = String(payload.destination || '').trim();
  const start = timeToDisplay_(payload.start);
  const end = timeToDisplay_(payload.end);
  const grades = normalizeFieldTripGrades_(payload.grades);
  const staffNames = splitFieldTripStaffList_(payload.staffNames || payload.staff);
  const notes = String(payload.notes || '').trim();

  if (!startDate) throw new Error('Choose a valid field trip start date.');
  if (!endDate) throw new Error('Choose a valid field trip end date.');
  if (endDate < startDate) throw new Error('Field trip end date must be the same as or after the start date.');
  if (!name) throw new Error('Enter a field trip name.');
  if (!start || !end) throw new Error('Enter both the departure and return time.');

  const startMinutes = displayTimeToMinutes_(start);
  const endMinutes = displayTimeToMinutes_(end);
  if (startMinutes == null || endMinutes == null) {
    throw new Error('Enter valid field trip departure and return times.');
  }
  if (startDate === endDate && endMinutes <= startMinutes) {
    throw new Error('For a one-day field trip, return time must be after departure time.');
  }

  if (!grades.length) throw new Error('Choose at least one student grade for the field trip.');
  if (!staffNames.length) throw new Error('Choose at least one staff member going on the field trip.');

  const sheet = ensureFieldTripsSheet_();
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(h => String(h || '').trim());
  const eventIdCol = findColumnByAliases_(headers, HEADER_ALIASES['Field Trips'].Event_ID);
  if (eventIdCol === -1) throw new Error('Field Trips sheet is missing Event_ID.');

  let eventId = String(payload.eventId || '').trim();

  const incomingIdentity = fieldTripIdentityKey_({
    name: name,
    destination: destination,
    startDate: startDate,
    endDate: endDate,
    start: start,
    end: end,
    grades: grades,
    staffNames: staffNames
  });
  const duplicate = getFieldTripsInRange_(startDate, endDate).find(trip =>
    trip.eventId !== eventId &&
    fieldTripIdentityKey_(trip) === incomingIdentity
  );

  if (duplicate) {
    if (eventId) {
      throw new Error(
        'Another field trip already matches this event exactly (' +
        (duplicate.name || 'Field Trip') + ', ' + duplicate.startDate + ').'
      );
    }
    return Object.assign({}, duplicate, { deduplicated: true });
  }

  if (!eventId) {
    eventId = 'FT-' + startDate.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  }

  const values = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][eventIdCol] || '').trim() === eventId) {
      targetRow = r + 1;
      break;
    }
  }
  if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

  // An edit can move a trip to different dates; any preview built for either
  // the old or the new dates no longer reflects this trip.
  const previousTrip = getFieldTripsInRange_('', '').find(trip => trip.eventId === eventId);
  if (previousTrip) invalidatePreviewForDateRange_(previousTrip.startDate, previousTrip.endDate);
  invalidatePreviewForDateRange_(startDate, endDate);

  setSheetRowObject_(sheet, targetRow, headers, HEADER_ALIASES['Field Trips'], {
    Event_ID: eventId,
    Name: name,
    Destination: destination,
    Date: startDate,
    End_Date: endDate,
    Start: start,
    End: end,
    Grades: grades.join(' | '),
    Staff: staffNames.join(' | '),
    Notes: notes
  });

  const savedTrip = normalizeFieldTripRow_({
    Event_ID: eventId,
    Name: name,
    Destination: destination,
    Date: startDate,
    End_Date: endDate,
    Start: start,
    End: end,
    Grades: grades.join(' | '),
    Staff: staffNames.join(' | '),
    Notes: notes
  });
  rebuildFieldTripCoveragePoolForTrip_(savedTrip);
  return savedTrip;
}

function deleteFieldTrip_(eventId) {
  return withCoverageLock_(() => deleteFieldTripUnlocked_(eventId));
}

function deleteFieldTripUnlocked_(eventId) {
  const id = String(eventId || '').trim();
  if (!id) throw new Error('No field trip ID was provided.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Field Trips');
  if (!sheet || sheet.getLastRow() < 2) return { deleted: false, eventId: id };

  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(h => String(h || '').trim());
  const eventIdCol = findColumnByAliases_(headers, HEADER_ALIASES['Field Trips'].Event_ID);
  if (eventIdCol === -1) throw new Error('Field Trips sheet is missing Event_ID.');

  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][eventIdCol] || '').trim() === id) {
      const trip = getFieldTripsInRange_('', '').find(item => item.eventId === id);
      sheet.deleteRow(r + 1);
      incrementCoverageMetric_('sheetWrites');
      invalidateCoverageSheetCache_('Field Trips');
      removeFieldTripCoveragePoolForEvent_(id);
      if (trip) invalidatePreviewForDateRange_(trip.startDate, trip.endDate);
      return { deleted: true, eventId: id };
    }
  }
  return { deleted: false, eventId: id };
}

function getCoverageStaffSheetName_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Coverage Staff')) return 'Coverage Staff';
  if (ss.getSheetByName('Substitutes')) return 'Substitutes';
  throw new Error('Missing sheet: Coverage Staff (or Substitutes)');
}

// ── Persistent per-day Teacher Schedule cache ─────────────────────────────
// Teacher Schedule is the authoritative source, but normal front-office actions
// should never rescan all ~5,500 rows. Each weekday gets a hidden prepared sheet
// with the derived fields normalizeTeacherScheduleRow_ would otherwise rebuild
// on every request. Editing Teacher Schedule or Class Schedule changes the
// revision; only the next requested weekday is rebuilt.
function teacherScheduleCacheSheetName_(day) {
  const code = String(day || '').trim().toUpperCase();
  return TEACHER_SCHEDULE_CACHE_DAYS_.indexOf(code) !== -1
    ? TEACHER_SCHEDULE_CACHE_SHEET_PREFIX_ + code
    : '';
}

function teacherScheduleCacheDayRevisionKey_(day) {
  return TEACHER_SCHEDULE_CACHE_DAY_REVISION_PREFIX_ + String(day || '').trim().toUpperCase();
}

function teacherScheduleCacheRevision_() {
  const properties = PropertiesService.getScriptProperties();
  let revision = properties.getProperty(TEACHER_SCHEDULE_CACHE_REVISION_PROPERTY_);
  if (revision) return revision;

  revision = Utilities.getUuid();
  properties.setProperty(TEACHER_SCHEDULE_CACHE_REVISION_PROPERTY_, revision);

  // A copied/test workbook may already contain cache tabs built from its current
  // Teacher Schedule. Trust them only when they contain the full prepared header.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  TEACHER_SCHEDULE_CACHE_DAYS_.forEach(day => {
    const sheet = ss && ss.getSheetByName(teacherScheduleCacheSheetName_(day));
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < TEACHER_SCHEDULE_CACHE_HEADERS_.length) return;
    const headers = sheet.getRange(1, 1, 1, TEACHER_SCHEDULE_CACHE_HEADERS_.length)
      .getValues()[0]
      .map(value => String(value || '').trim());
    if (TEACHER_SCHEDULE_CACHE_HEADERS_.every((header, index) => headers[index] === header)) {
      properties.setProperty(teacherScheduleCacheDayRevisionKey_(day), revision);
    }
  });
  return revision;
}

function markTeacherScheduleCacheDirty_() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(TEACHER_SCHEDULE_CACHE_REVISION_PROPERTY_, Utilities.getUuid());
  invalidateCoverageSheetCache_('Teacher Schedule');
  TEACHER_SCHEDULE_CACHE_DAYS_.forEach(day => {
    invalidateCoverageSheetCache_(teacherScheduleCacheSheetName_(day));
  });
}

function teacherScheduleCacheDayIsFresh_(day) {
  const code = String(day || '').trim().toUpperCase();
  const sheetName = teacherScheduleCacheSheetName_(code);
  if (!sheetName) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < TEACHER_SCHEDULE_CACHE_HEADERS_.length) return false;
  return PropertiesService.getScriptProperties()
    .getProperty(teacherScheduleCacheDayRevisionKey_(code)) === teacherScheduleCacheRevision_();
}

function markTeacherScheduleCacheDayFresh_(day) {
  const code = String(day || '').trim().toUpperCase();
  if (!teacherScheduleCacheSheetName_(code)) return;
  PropertiesService.getScriptProperties().setProperty(
    teacherScheduleCacheDayRevisionKey_(code),
    teacherScheduleCacheRevision_()
  );
}

function ensureTeacherScheduleCacheSheet_(day) {
  const code = String(day || '').trim().toUpperCase();
  const name = teacherScheduleCacheSheetName_(code);
  if (!name) throw new Error('Invalid Teacher Schedule cache day: ' + day);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getMaxColumns() < TEACHER_SCHEDULE_CACHE_HEADERS_.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TEACHER_SCHEDULE_CACHE_HEADERS_.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, TEACHER_SCHEDULE_CACHE_HEADERS_.length)
    .setValues([TEACHER_SCHEDULE_CACHE_HEADERS_])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}

function rawTeacherScheduleRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher Schedule');
  if (!sheet) throw new Error('Missing Teacher Schedule sheet.');
  if (sheet.getLastRow() < 2) return { headers: [], rows: [] };

  // Deliberately stop at H. Column I is the Lead/Co spill formula and is not
  // consumed by the scheduler; reading it forces needless formula evaluation.
  const width = Math.min(8, sheet.getLastColumn());
  incrementCoverageMetric_('sheetReads');
  const values = sheet.getRange(1, 1, sheet.getLastRow(), width).getValues();
  return {
    headers: values[0].map(value => String(value || '').trim()),
    rows: values.slice(1)
  };
}

function rebuildTeacherScheduleDayCache_(day) {
  const code = String(day || '').trim().toUpperCase();
  if (!teacherScheduleCacheSheetName_(code)) return [];

  return withCoverageLock_(() => {
    const source = rawTeacherScheduleRows_();
    const headers = source.headers;
    const aliasMap = HEADER_ALIASES['Teacher Schedule'];
    const col = {};
    Object.keys(aliasMap).forEach(key => {
      col[key] = findColumnByAliases_(headers, aliasMap[key]);
    });
    if (col.Staff_Name === -1 || col.Day === -1 || col.Start === -1 || col.End === -1) {
      throw new Error('Teacher Schedule needs Teacher, Day, Start, and End columns.');
    }

    const value = (row, key) => col[key] === -1 ? '' : row[col[key]];
    const prepared = [];
    source.rows.forEach(row => {
      if (String(value(row, 'Day') || '').trim().toUpperCase() !== code) return;
      const raw = {
        Staff_Name: value(row, 'Staff_Name'),
        Role: value(row, 'Role'),
        Term: value(row, 'Term'),
        Day: value(row, 'Day'),
        Start: value(row, 'Start'),
        End: value(row, 'End'),
        Class: value(row, 'Class'),
        Grade: value(row, 'Grade'),
        Subject: value(row, 'Subject'),
        Assignment_Type: value(row, 'Assignment_Type'),
        Room: value(row, 'Room'),
        Needs_Coverage_If_Absent: value(row, 'Needs_Coverage_If_Absent'),
        Cover_Eligible_This_Block: value(row, 'Cover_Eligible_This_Block')
      };
      const normalized = normalizeTeacherScheduleRow_(raw);
      prepared.push([
        normalized.staffName,
        raw.Term || '',
        code,
        raw.Start || '',
        raw.End || '',
        raw.Class || '',
        normalized.subject,
        normalized.room,
        normalized.grade,
        normalized.assignmentType,
        normalized.needsCoverageIfAbsent ? 'Yes' : 'No',
        normalized.coverEligibleThisBlock ? 'Yes' : 'No'
      ]);
    });

    const sheet = ensureTeacherScheduleCacheSheet_(code);
    const previousLastRow = sheet.getLastRow();
    if (prepared.length) {
      sheet.getRange(2, 1, prepared.length, TEACHER_SCHEDULE_CACHE_HEADERS_.length).setValues(prepared);
      sheet.getRange(2, 4, prepared.length, 2).setNumberFormat('h:mm AM/PM');
      incrementCoverageMetric_('sheetWrites');
    }
    if (previousLastRow > prepared.length + 1) {
      sheet.getRange(
        prepared.length + 2,
        1,
        previousLastRow - prepared.length - 1,
        TEACHER_SCHEDULE_CACHE_HEADERS_.length
      ).clearContent();
      incrementCoverageMetric_('sheetWrites');
    }
    if (!sheet.isSheetHidden()) sheet.hideSheet();

    invalidateCoverageSheetCache_(teacherScheduleCacheSheetName_(code));
    markTeacherScheduleCacheDayFresh_(code);
    coveragePerfMark_('teacher-schedule-' + code + '-rebuilt');
    return prepared;
  });
}

function readTeacherScheduleDayCached_(day) {
  const code = String(day || '').trim().toUpperCase();
  const sheetName = teacherScheduleCacheSheetName_(code);
  if (!sheetName) return readSheetObjects_('Teacher Schedule');

  if (!teacherScheduleCacheDayIsFresh_(code)) {
    rebuildTeacherScheduleDayCache_(code);
  }

  // Cache sheets use the same canonical aliases as Teacher Schedule.
  const name = sheetName;
  if (Object.prototype.hasOwnProperty.call(COVERAGE_REQUEST_SHEET_CACHE_, name)) {
    incrementCoverageMetric_('requestCacheHits');
    return COVERAGE_REQUEST_SHEET_CACHE_[name];
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  incrementCoverageMetric_('sheetReads');
  const values = sheet.getRange(
    1, 1, sheet.getLastRow(), TEACHER_SCHEDULE_CACHE_HEADERS_.length
  ).getValues();
  const headers = values[0].map(value => String(value || '').trim());
  const aliasMap = HEADER_ALIASES['Teacher Schedule'];
  const columnByCanonical = {};
  Object.keys(aliasMap).forEach(canonical => {
    columnByCanonical[canonical] = findColumnByAliases_(headers, aliasMap[canonical]);
  });

  const rows = values.slice(1).map(row => {
    const out = {};
    Object.keys(columnByCanonical).forEach(canonical => {
      const index = columnByCanonical[canonical];
      out[canonical] = index === -1 ? '' : row[index];
    });
    return out;
  }).filter(row => String(row.Staff_Name || '').trim());

  COVERAGE_REQUEST_SHEET_CACHE_[name] = rows;
  coveragePerfMark_('teacher-schedule-' + code + '-cache-read');
  return rows;
}

function teacherScheduleRowsForDate_(date, day, config) {
  const code = String(day || guessDayCodeFromDate_(normalizeDateKey_(date)) || '').trim().toUpperCase();
  const rows = readTeacherScheduleDayCached_(code);
  return filterTeacherScheduleForDate_(rows, date, config || getConfigMap_());
}

// Compatibility helper for older code paths that intentionally need the full
// master schedule (validation/reporting). Hot paths should use the day helper.
function readTeacherScheduleCached_() {
  return readSheetObjects_('Teacher Schedule');
}

function rebuildAllTeacherScheduleDayCaches_() {
  const counts = {};
  TEACHER_SCHEDULE_CACHE_DAYS_.forEach(day => {
    counts[day] = rebuildTeacherScheduleDayCache_(day).length;
  });
  return counts;
}

function menuRebuildTeacherScheduleDayCaches() {
  const counts = rebuildAllTeacherScheduleDayCaches_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Schedule caches rebuilt: ' +
      TEACHER_SCHEDULE_CACHE_DAYS_.map(day => day + ' ' + counts[day]).join(' · '),
    APP_TITLE,
    8
  );
  return counts;
}

function getAllSchedulableStaff_(dayCode, date) {
  const config = getConfigMap_();
  const rows = dayCode
    ? (date
        ? teacherScheduleRowsForDate_(date, dayCode, config)
        : readTeacherScheduleDayCached_(dayCode))
    : readSheetObjects_('Teacher Schedule');
  const map = {};

  rows.forEach(row => {
    const normalized = normalizeTeacherScheduleRow_(row);
    if (dayCode && normalized.day !== String(dayCode).trim()) return;
    const name = normalized.staffName;
    if (!name) return;
    if (!map[name]) {
      map[name] = {
        name: name,
        role: normalized.role,
        days: {},
        blocks: []
      };
    }
    map[name].days[normalized.day] = true;
    if (normalized.startMinutes != null && normalized.endMinutes != null) {
      map[name].blocks.push({
        start: minutesToDisplay_(normalized.startMinutes),
        end: minutesToDisplay_(normalized.endMinutes),
        className: normalized.className,
        grade: normalized.grade,
        subject: normalized.subject,
        assignmentType: normalized.assignmentType,
        room: normalized.room,
        needsCoverageIfAbsent: normalized.needsCoverageIfAbsent
      });
    }
  });

  return Object.keys(map).sort().map(name => {
    map[name].blocks.sort((a, b) => displayTimeToMinutes_(a.start) - displayTimeToMinutes_(b.start));
    return map[name];
  });
}

function getActiveCoverageStaff_(date, day) {
  return getCoverageStaffForDate_(date, day, getConfigMap_())
    .filter(row => row.name && row.activeToday)
    .map(row => ({
      name: row.name,
      role: row.role,
      tier: row.tier,
      allDay: row.canCoverAllDay,
      selectedStart: row.selectedStart,
      selectedEnd: row.selectedEnd,
      notes: row.notes
    }))
    .sort((a, b) => a.tier - b.tier || a.name.localeCompare(b.name));
}

function getAllCoverageStaff_(date, day) {
  const config = getConfigMap_();
  return getCoverageStaffForDate_(date, day, config)
    .map(row => ({
      name: row.name,
      role: row.role,
      tier: row.tier,
      allDay: row.canCoverAllDay,
      activeToday: row.activeToday,
      availableDays: row.availableDays === '*' ? '' : row.availableDays,
      defaultStart: row.defaultStart,
      defaultEnd: row.defaultEnd,
      selectedStart: row.selectedStart,
      selectedEnd: row.selectedEnd,
      hasDateOverride: row.hasDateOverride,
      availabilityNotes: row.availabilityNotes,
      allowedGrades: row.allowedGrades === '*' ? '' : row.allowedGrades,
      allowedSubjects: row.allowedSubjects === '*' ? '' : row.allowedSubjects,
      allowedAssignmentTypes: row.allowedAssignmentTypes === '*' ? '' : row.allowedAssignmentTypes,
      // Show the person's own setting; blank means the Config default applies.
      // Showing the effective value would write the default into their row the
      // next time someone saved the edit form.
      maxBlocksPerDay: row.rawMaxBlocksPerDay,
      maxTeachersPerDay: row.rawMaxTeachersPerDay,
      canBeSplitAcrossTeachers: row.canBeSplitAcrossTeachers,
      notes: row.notes
    }))
    .filter(row => row.name)
    .sort((a, b) => a.tier - b.tier || a.name.localeCompare(b.name));
}

function toggleCoverageStaffActive(payload) {
  return withCoverageLock_(() => toggleCoverageStaffActiveUnlocked_(payload));
}

function toggleCoverageStaffActiveUnlocked_(payload) {
  payload = payload || {};
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('No staff name provided.');

  const date = normalizeDateKey_(payload.date);
  const day = String(payload.day || guessDayCodeFromDate_(date) || '').trim();

  if (date && day) {
    const current = getAllCoverageStaff_(date, day).find(row => row.name === name);
    const hasExplicitAvailable = payload.available !== undefined && payload.available !== null;
    const nextAvailable = hasExplicitAvailable
      ? normalizeYesNo_(payload.available, true)
      : !(current && current.activeToday);
    upsertSubstituteAvailability({
      date: date,
      day: day,
      name: name,
      available: nextAvailable,
      start: current ? current.selectedStart : '',
      end: current ? current.selectedEnd : '',
      notes: current ? current.availabilityNotes : ''
    });
    return getAllCoverageStaff_(date, day);
  }

  const sheetName = getCoverageStaffSheetName_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet: ' + sheetName);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const aliasMap = HEADER_ALIASES['Coverage Staff'];

  const nameCol = findColumnByAliases_(headers, aliasMap.Name);
  if (nameCol === -1) throw new Error('Cannot find Name column.');

  const activeCol = findColumnByAliases_(headers, aliasMap.Active_Today);
  if (activeCol === -1) throw new Error('Cannot find Active_Today column.');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][nameCol] || '').trim() === name) {
      const current = normalizeYesNo_(values[r][activeCol], true);
      sheet.getRange(r + 1, activeCol + 1).setValue(current ? 'No' : 'Yes');
      incrementCoverageMetric_('sheetWrites');
      invalidateCoverageSheetCache_(getCoverageStaffSheetName_());
      break;
    }
  }

  return getAllCoverageStaff_();
}

function findColumnByAliases_(headers, aliases) {
  if (!aliases) return -1;
  for (let i = 0; i < aliases.length; i++) {
    const idx = headers.indexOf(aliases[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function getDailyAbsencesForDate_(dateStr, dayCode) {
  const rows = readSheetObjects_('Daily Absences');
  return rows
    .filter(row => normalizeDateKey_(row.Date) === dateStr && String(row.Day || '').trim() === String(dayCode || '').trim())
    .map(row => ({
      date: normalizeDateKey_(row.Date),
      day: String(row.Day || '').trim(),
      staffName: String(row.Staff_Name || '').trim(),
      absenceType: String(row.Absence_Type || 'Full Day').trim(),
      startOverride: timeToDisplay_(row.Start_Override),
      endOverride: timeToDisplay_(row.End_Override),
      notes: stripEmergencyNoteMarker_(row.Notes),
      emergency: isEmergencyAbsence_(row),
      preferredCoverage: String(row.Preferred_Coverage || '').trim()
    }))
    .filter(row => row.staffName)
    .sort((a, b) =>
      a.staffName.localeCompare(b.staffName) ||
      (displayTimeToMinutes_(a.startOverride) || 0) - (displayTimeToMinutes_(b.startOverride) || 0)
    );
}

function replaceDailyAbsences(payload) {
  return withCoverageLock_(() => replaceDailyAbsencesUnlocked_(payload));
}

function replaceDailyAbsencesUnlocked_(payload) {
  payload = payload || {};

  const date = payload.date || Utilities.formatDate(new Date(), coverageTimeZone_(), 'yyyy-MM-dd');
  const day = String(payload.day || guessDayCodeFromDate_(date) || '').trim();
  const incomingAbsences = (payload.absences || []).filter(row => row && row.staffName);
  const staffNames = (payload.staffNames || []).filter(Boolean);

  if (!date) throw new Error('No date available for Daily Absences.');
  if (!day) throw new Error('No day code available for Daily Absences.');

  const rows = readSheetObjects_('Daily Absences');
  const kept = rows.filter(
    row => !(normalizeDateKey_(row.Date) === date && String(row.Day || '').trim() === day)
  );

  const added = incomingAbsences.length
    ? incomingAbsences.map(row => {
      const allDay = normalizeYesNo_(row.allDay, false);
      const emergency = normalizeYesNo_(row.emergency, false);
      return {
        Date: date,
        Day: day,
        Staff_Name: String(row.staffName || '').trim(),
        Absence_Type: allDay ? 'Full Day' : 'Partial Day',
        Start_Override: allDay ? '' : String(row.startOverride || '').trim(),
        End_Override: allDay ? '' : String(row.endOverride || '').trim(),
        Notes: composeAbsenceNotes_(emergency, row.notes),
        Preferred_Coverage: String(row.preferredCoverage || '').trim()
      };
    })
    : staffNames.map(name => ({
      Date: date,
      Day: day,
      Staff_Name: name,
      Absence_Type: 'Full Day',
      Start_Override: '',
      End_Override: '',
      Notes: '',
      Preferred_Coverage: ''
    }));

  rewriteSheetRows_(
    'Daily Absences',
    SHEET_SCHEMAS['Daily Absences'].headers,
    kept.concat(added)
  );
  invalidatePreviewForDateRange_(date, date);

  return getDailyAbsencesForDate_(date, day);
}

function fieldTripTimes_(trip) {
  return {
    startMinutes: trip && trip.effectiveStartMinutes != null
      ? Number(trip.effectiveStartMinutes)
      : displayTimeToMinutes_(trip && trip.start),
    endMinutes: trip && trip.effectiveEndMinutes != null
      ? Number(trip.effectiveEndMinutes)
      : displayTimeToMinutes_(trip && trip.end)
  };
}

function fieldTripGradeMatches_(trip, gradeValue) {
  const grade = normalizeGradeKey_(gradeValue);
  return !!grade && (trip.grades || []).some(value => normalizeGradeKey_(value) === grade);
}

function isInstructionalGradeBlock_(row) {
  const type = String(row.assignmentType || '').trim().toLowerCase();
  if (type === 'planning' || type === 'break' || type === 'lunch' || type === 'meeting' || type === 'duty') return false;
  return !!String(row.grade || '').trim();
}

function blockOverlapsFieldTrip_(row, trip) {
  const times = fieldTripTimes_(trip);
  if (times.startMinutes == null || times.endMinutes == null) return false;
  return row.startMinutes < times.endMinutes && row.endMinutes > times.startMinutes;
}

function blockIsCancelledByFieldTrip_(row, fieldTrips) {
  return (fieldTrips || []).some(trip =>
    fieldTripGradeMatches_(trip, row.grade) &&
    isInstructionalGradeBlock_(row) &&
    blockOverlapsFieldTrip_(row, trip)
  );
}

function coveragePersonNameKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function canonicalTeacherScheduleName_(name, teacherSchedule, day) {
  const key = coveragePersonNameKey_(name);
  if (!key) return String(name || '').trim();
  const match = (teacherSchedule || [])
    .map(row => normalizeTeacherScheduleRow_(row))
    .find(row =>
      row.staffName &&
      (!day || row.day === day) &&
      coveragePersonNameKey_(row.staffName) === key
    );
  return match ? match.staffName : String(name || '').trim();
}

function buildFieldTripParticipantAbsences_(fieldTrips, teacherSchedule, day) {
  const rows = [];
  (fieldTrips || []).forEach(trip => {
    const times = fieldTripTimes_(trip);
    (trip.staffNames || []).forEach(name => {
      rows.push({
        staffName: canonicalTeacherScheduleName_(name, teacherSchedule, day),
        absenceType: trip.allDayForDate ? 'Full Day' : 'Partial Day',
        startOverride: trip.allDayForDate ? '' : minutesToDisplay_(Math.min(times.startMinutes, 1439)),
        endOverride: trip.allDayForDate ? '' : minutesToDisplay_(Math.min(times.endMinutes, 1439)),
        notes: trip.name || 'Field Trip',
        emergency: false,
        preferredCoverage: '',
        fieldTripEventId: trip.eventId,
        fieldTripName: trip.name,
        fieldTripGrades: (trip.grades || []).slice()
      });
    });
  });
  return rows;
}

function markFieldTripCoveragePoolDirty_() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(FIELD_TRIP_COVERAGE_POOL_DIRTY_PROPERTY_, '1');
  properties.setProperty(
    FIELD_TRIP_COVERAGE_POOL_REVISION_PROPERTY_,
    Utilities.getUuid()
  );
  invalidateCoverageSheetCache_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
}

function clearFieldTripCoveragePoolDirty_() {
  PropertiesService.getScriptProperties()
    .deleteProperty(FIELD_TRIP_COVERAGE_POOL_DIRTY_PROPERTY_);
}

function fieldTripCoveragePoolIsDirty_() {
  return PropertiesService.getScriptProperties()
    .getProperty(FIELD_TRIP_COVERAGE_POOL_DIRTY_PROPERTY_) === '1';
}

function fieldTripCoveragePoolRevision_() {
  const properties = PropertiesService.getScriptProperties();
  let revision = properties.getProperty(FIELD_TRIP_COVERAGE_POOL_REVISION_PROPERTY_);
  if (!revision) {
    revision = Utilities.getUuid();
    properties.setProperty(FIELD_TRIP_COVERAGE_POOL_REVISION_PROPERTY_, revision);
  }
  return revision;
}

function fieldTripCoveragePoolDateRevisionKey_(date) {
  return FIELD_TRIP_COVERAGE_POOL_DATE_REVISION_PREFIX_ + normalizeDateKey_(date);
}

function fieldTripCoveragePoolDateIsFresh_(date) {
  const key = normalizeDateKey_(date);
  if (!key || !fieldTripCoveragePoolIsDirty_()) return true;
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty(fieldTripCoveragePoolDateRevisionKey_(key)) ===
    fieldTripCoveragePoolRevision_();
}

function markFieldTripCoveragePoolDateFresh_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty(
    fieldTripCoveragePoolDateRevisionKey_(key),
    fieldTripCoveragePoolRevision_()
  );
}

function ensureFieldTripCoveragePoolSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(FIELD_TRIP_COVERAGE_POOL_SHEET_);
    sheet.getRange(1, 1, 1, SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers.length)
      .setValues([SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers]);
  } else {
    ensureHeaderRow_(sheet, SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers);
  }
  if (typeof formatFieldTripCoveragePoolSheet_ === 'function') formatFieldTripCoveragePoolSheet_(sheet);
  if (sheet.isSheetHidden()) sheet.showSheet();
  return sheet;
}

function fieldTripDateKeys_(trip) {
  if (!trip || !trip.startDate || !trip.endDate) return [];
  const start = new Date(trip.startDate + 'T12:00:00');
  const end = new Date(trip.endDate + 'T12:00:00');
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return [];
  const keys = [];
  const cursor = new Date(start.getTime());
  while (cursor <= end) {
    const key = Utilities.formatDate(cursor, coverageTimeZone_(), 'yyyy-MM-dd');
    const day = guessDayCodeFromDate_(key);
    if (day === 'M' || day === 'T' || day === 'W' || day === 'R' || day === 'F') keys.push(key);
    cursor.setDate(cursor.getDate() + 1);
  }
  return keys;
}

function fieldTripAffectedNamesForDate_(trip, date, teacherScheduleRows, config) {
  const activeTrip = fieldTripForDate_(trip, date);
  if (!activeTrip) return [];

  const day = guessDayCodeFromDate_(date);
  const filtered = filterTeacherScheduleForDate_(teacherScheduleRows, date, config);
  const participantSet = {};
  (activeTrip.staffNames || []).forEach(name => { participantSet[coveragePersonNameKey_(name)] = true; });

  const affected = {};
  filtered.forEach(raw => {
    const row = normalizeTeacherScheduleRow_(raw);
    if (!row.staffName || row.day !== day) return;
    if (participantSet[coveragePersonNameKey_(row.staffName)]) return;
    if (!fieldTripGradeMatches_(activeTrip, row.grade)) return;
    if (!isInstructionalGradeBlock_(row)) return;
    if (!blockOverlapsFieldTrip_(row, activeTrip)) return;
    affected[row.staffName] = true;
  });

  return Object.keys(affected).sort();
}

function makeFieldTripReleaseCandidate_(name, trip, date) {
  const activeTrip = fieldTripForDate_(trip, date) || trip;
  const times = fieldTripTimes_(activeTrip);
  return {
    name: name,
    role: 'Field Trip Release',
    tier: 3,
    canCoverAllDay: false,
    baseActive: true,
    activeToday: true,
    availableDays: '*',
    defaultStart: '',
    defaultEnd: '',
    selectedStart: '',
    selectedEnd: '',
    hasDateOverride: false,
    availabilityNotes: '',
    allowedGrades: '*',
    allowedSubjects: '*',
    allowedAssignmentTypes: '*',
    maxBlocksPerDay: Infinity,
    maxTeachersPerDay: Infinity,
    canBeSplitAcrossTeachers: true,
    notes: 'Temporary coverage availability created by a field trip.',
    fieldTripEvents: [{
      eventId: trip.eventId,
      name: trip.name,
      grades: (trip.grades || []).slice(),
      startMinutes: times.startMinutes,
      endMinutes: times.endMinutes
    }],
    fieldTripOnly: true
  };
}

function buildFieldTripCoverageCandidatesLive_(trip, date, teacherScheduleRows, config, activeCoverageStaff, configuredCoverageStaff) {
  const configuredByName = {};
  (configuredCoverageStaff || []).forEach(candidate => {
    if (candidate && candidate.name) configuredByName[candidate.name] = candidate;
  });

  const byName = {};
  (activeCoverageStaff || []).forEach(candidate => {
    byName[candidate.name] = Object.assign({}, candidate, {
      fieldTripEvents: (candidate.fieldTripEvents || []).slice(),
      fieldTripOnly: false
    });
  });

  fieldTripAffectedNamesForDate_(trip, date, teacherScheduleRows, config).forEach(name => {
    const configured = configuredByName[name];
    if (configured && !configured.activeToday) return;

    let candidate = byName[name];
    if (!candidate) {
      candidate = makeFieldTripReleaseCandidate_(name, trip, date);
      byName[name] = candidate;
      return;
    }

    if (!candidate.fieldTripEvents.some(event => event.eventId === trip.eventId)) {
      candidate.fieldTripEvents = candidate.fieldTripEvents.concat(
        makeFieldTripReleaseCandidate_(name, trip, date).fieldTripEvents
      );
    }
  });

  return Object.keys(byName).map(name => byName[name]);
}

function fieldTripCoveragePoolRowKey_(row) {
  return [
    String(row.Event_ID || '').trim(),
    normalizeDateKey_(row.Date),
    coveragePersonNameKey_(row.Absent_Staff),
    timeToDisplay_(row.Start),
    timeToDisplay_(row.End),
    coveragePersonNameKey_(row.Candidate)
  ].join('\u0000');
}

function normalizePoolEnabled_(value, fallback) {
  const raw = String(value == null ? '' : value).trim().toLowerCase();
  if (!raw) return fallback !== false;
  return ['true', 'yes', 'y', '1', 'enabled', 'on'].indexOf(raw) !== -1;
}

function poolPriorityAdjustment_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function fieldTripCoveragePoolRowsForTrip_(trip, existingRows) {
  const config = getConfigMap_();
  const overrides = {};
  (existingRows || []).forEach(row => {
    if (String(row.Event_ID || '').trim() !== String(trip.eventId || '').trim()) return;
    overrides[fieldTripCoveragePoolRowKey_(row)] = row;
  });

  const builtAt = Utilities.formatDate(new Date(), coverageTimeZone_(), "yyyy-MM-dd'T'HH:mm:ss");
  const rows = [];

  fieldTripDateKeys_(trip).forEach(date => {
    const activeTrip = fieldTripForDate_(trip, date);
    if (!activeTrip) return;
    const day = guessDayCodeFromDate_(date);
    const teacherSchedule = teacherScheduleRowsForDate_(date, day, config);
    const configuredCoverageStaff = getCoverageStaffForDate_(date, day, config);
    const activeCoverageStaff = configuredCoverageStaff.filter(row => row.name && row.activeToday);
    const candidates = buildFieldTripCoverageCandidatesLive_(
      activeTrip,
      date,
      teacherSchedule,
      config,
      activeCoverageStaff,
      configuredCoverageStaff
    );

    const tripAbsences = buildFieldTripParticipantAbsences_([activeTrip], teacherSchedule, day);
    const needsByTeacher = buildCoverageNeedsByTeacher_(tripAbsences, teacherSchedule, day, [activeTrip]);
    const state = makeEmptyState_();
    state.absencesByCandidate = buildAbsenceWindowsByStaff_(tripAbsences);

    Object.keys(needsByTeacher).sort().forEach(absentName => {
      (needsByTeacher[absentName] || []).forEach(block => {
        if (String(block.fieldTripEventId || '') !== String(trip.eventId || '')) return;

        candidates.forEach(candidate => {
          if (!candidateCanCoverBlock_(candidate, absentName, block, teacherSchedule, day, state, config)) return;

          const availability = candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state);
          const scoreInfo = scoreCandidateForBlock_(candidate, absentName, block, state);
          const runwayMinutes = Math.max(0, Number(availability.fieldTripRunwayMinutes || 0));
          const baselineScore = scoreInfo.score +
            Number(availability.fieldTripPriority || 0) * 250 +
            Math.min(runwayMinutes, 60);
          const source = candidateHasFieldTripEvent_(candidate, trip.eventId)
            ? 'Field Trip Pool'
            : 'Coverage Staff';
          const reason = availability.fieldTripReason ||
            scoreInfo.reason ||
            (candidate.canCoverAllDay
              ? 'Available Coverage Staff for this block.'
              : 'Available during a cover-eligible schedule block.');

          const row = {
            Event_ID: trip.eventId,
            Event_Name: trip.name || '',
            Date: date,
            Day: day,
            Absent_Staff: absentName,
            Start: minutesToDisplay_(block.startMinutes),
            End: minutesToDisplay_(block.endMinutes),
            Class: block.className || '',
            Grade: block.grade || '',
            Subject: block.subject || '',
            Assignment_Type: block.assignmentType || '',
            Room: block.room || '',
            Candidate: candidate.name,
            Candidate_Role: candidate.role || '',
            Candidate_Tier: candidate.tier,
            Source: source,
            Baseline_Score: baselineScore,
            Reason: reason,
            Enabled: 'Yes',
            Priority_Adjustment: 0,
            Built_At: builtAt
          };

          const previous = overrides[fieldTripCoveragePoolRowKey_(row)];
          if (previous) {
            row.Enabled = previous.Enabled === '' || previous.Enabled == null ? 'Yes' : previous.Enabled;
            row.Priority_Adjustment = previous.Priority_Adjustment === '' || previous.Priority_Adjustment == null
              ? 0
              : previous.Priority_Adjustment;
          }
          rows.push(row);
        });
      });
    });
  });

  return rows;
}

function sortFieldTripCoveragePoolRows_(rows) {
  return (rows || []).slice().sort((a, b) =>
    String(a.Date || '').localeCompare(String(b.Date || '')) ||
    displayTimeToMinutes_(a.Start) - displayTimeToMinutes_(b.Start) ||
    String(a.Absent_Staff || '').localeCompare(String(b.Absent_Staff || '')) ||
    Number(b.Baseline_Score || 0) - Number(a.Baseline_Score || 0) ||
    String(a.Candidate || '').localeCompare(String(b.Candidate || ''))
  );
}

function rebuildFieldTripCoveragePoolForTrip_(trip) {
  if (!trip || !trip.eventId) return [];
  ensureFieldTripCoveragePoolSheet_();

  const existing = readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  const kept = existing.filter(row =>
    String(row.Event_ID || '').trim() !== String(trip.eventId || '').trim()
  );
  const rebuilt = fieldTripCoveragePoolRowsForTrip_(trip, existing);
  const allRows = sortFieldTripCoveragePoolRows_(kept.concat(rebuilt));

  rewriteSheetRows_(
    FIELD_TRIP_COVERAGE_POOL_SHEET_,
    SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers,
    allRows
  );
  coveragePerfMark_('field-trip-pool-rebuilt');
  return rebuilt;
}

function removeFieldTripCoveragePoolForEvent_(eventId) {
  const id = String(eventId || '').trim();
  if (!id) return;
  ensureFieldTripCoveragePoolSheet_();
  const existing = readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  const kept = existing.filter(row => String(row.Event_ID || '').trim() !== id);
  if (kept.length === existing.length) return;
  rewriteSheetRows_(
    FIELD_TRIP_COVERAGE_POOL_SHEET_,
    SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers,
    kept
  );
}

function rebuildFieldTripCoveragePoolForDate_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return [];

  ensureFieldTripCoveragePoolSheet_();
  const existing = readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  const trips = getFieldTripsForDate_(key);
  const eventIds = new Set(trips.map(trip => String(trip.eventId || '').trim()).filter(Boolean));

  // Replace only the events active on the requested date. A multi-day trip is
  // rebuilt as a whole so its other dates stay internally consistent.
  let kept = existing.filter(row => !eventIds.has(String(row.Event_ID || '').trim()));
  let rebuilt = [];
  trips.forEach(trip => {
    rebuilt = rebuilt.concat(fieldTripCoveragePoolRowsForTrip_(trip, existing));
  });

  rewriteSheetRows_(
    FIELD_TRIP_COVERAGE_POOL_SHEET_,
    SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers,
    sortFieldTripCoveragePoolRows_(kept.concat(rebuilt))
  );
  markFieldTripCoveragePoolDateFresh_(key);
  coveragePerfMark_('field-trip-pool-date-rebuilt');
  return rebuilt.filter(row => normalizeDateKey_(row.Date) === key);
}

function rebuildAllFieldTripCoveragePool_() {
  ensureFieldTripCoveragePoolSheet_();
  const trips = getFieldTripsInRange_('', '');
  const existing = readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  let rows = [];
  trips.forEach(trip => {
    rows = rows.concat(fieldTripCoveragePoolRowsForTrip_(trip, existing));
  });

  rewriteSheetRows_(
    FIELD_TRIP_COVERAGE_POOL_SHEET_,
    SHEET_SCHEMAS[FIELD_TRIP_COVERAGE_POOL_SHEET_].headers,
    sortFieldTripCoveragePoolRows_(rows)
  );
  clearFieldTripCoveragePoolDirty_();
  coveragePerfMark_('field-trip-pool-full-rebuild');
  return rows;
}

function ensureFieldTripCoveragePoolFresh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_) && !fieldTripCoveragePoolIsDirty_()) {
    return readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  }

  // Full rebuild remains available from the spreadsheet menu. Interactive
  // Generate uses fieldTripCoveragePoolRowsForDate_() so it never rebuilds the
  // entire workbook just because one source sheet changed.
  return withCoverageLock_(() => {
    const current = SpreadsheetApp.getActiveSpreadsheet();
    if (current.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_) && !fieldTripCoveragePoolIsDirty_()) {
      return readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_);
    }
    return rebuildAllFieldTripCoveragePool_();
  });
}

function fieldTripCoveragePoolRowsForDate_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasSheet = !!ss.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  if (hasSheet && fieldTripCoveragePoolDateIsFresh_(key)) {
    return readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_)
      .filter(row => normalizeDateKey_(row.Date) === key);
  }

  // Source/config changes invalidate the pool globally, but Generate only needs
  // one date. Rebuild that date under the existing re-entrant lock, then remember
  // its revision so repeated Generate clicks remain fast while other dates stay
  // correctly marked stale until they are used.
  return withCoverageLock_(() => {
    const current = SpreadsheetApp.getActiveSpreadsheet();
    if (current.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_) &&
        fieldTripCoveragePoolDateIsFresh_(key)) {
      return readSheetObjects_(FIELD_TRIP_COVERAGE_POOL_SHEET_)
        .filter(row => normalizeDateKey_(row.Date) === key);
    }
    return rebuildFieldTripCoveragePoolForDate_(key);
  });
}

function fieldTripCoveragePoolEntriesForBlock_(poolRows, eventId, absentName, block) {
  const id = String(eventId || '').trim();
  return (poolRows || []).filter(row =>
    String(row.Event_ID || '').trim() === id &&
    coveragePersonNameKey_(row.Absent_Staff) === coveragePersonNameKey_(absentName) &&
    displayTimeToMinutes_(row.Start) === Number(block.startMinutes) &&
    displayTimeToMinutes_(row.End) === Number(block.endMinutes) &&
    normalizePoolEnabled_(row.Enabled, true)
  );
}

function buildFieldTripCoverageCandidatesFromPool_(fieldTrips, date, activeCoverageStaff, configuredCoverageStaff, poolRows) {
  const configuredByName = {};
  (configuredCoverageStaff || []).forEach(candidate => {
    if (candidate && candidate.name) configuredByName[candidate.name] = candidate;
  });

  const byName = {};
  (activeCoverageStaff || []).forEach(candidate => {
    byName[candidate.name] = Object.assign({}, candidate, {
      fieldTripEvents: (candidate.fieldTripEvents || []).slice(),
      fieldTripOnly: false
    });
  });

  if (!(fieldTrips || []).length) {
    return Object.keys(byName).map(name => byName[name]);
  }

  const tripsById = {};
  (fieldTrips || []).forEach(trip => { tripsById[String(trip.eventId || '').trim()] = trip; });

  (poolRows || []).forEach(row => {
    if (!normalizePoolEnabled_(row.Enabled, true)) return;
    const eventId = String(row.Event_ID || '').trim();
    const name = String(row.Candidate || '').trim();
    const trip = tripsById[eventId];
    if (!eventId || !name || !trip) return;

    const configured = configuredByName[name];
    if (configured && !configured.activeToday) return;

    let candidate = byName[name];
    if (!candidate) {
      candidate = makeFieldTripReleaseCandidate_(name, trip, date);
      candidate.fieldTripEvents = [];
      byName[name] = candidate;
    }

    if (String(row.Source || '') === 'Field Trip Pool' &&
        !candidate.fieldTripEvents.some(event => event.eventId === eventId)) {
      candidate.fieldTripEvents = candidate.fieldTripEvents.concat(
        makeFieldTripReleaseCandidate_(name, trip, date).fieldTripEvents
      );
    }
  });

  return Object.keys(byName).map(name => byName[name]);
}

function fieldTripPoolCandidateSubset_(poolRows, eventId, absentName, block, coverageStaff, releasedOnly) {
  const entries = fieldTripCoveragePoolEntriesForBlock_(poolRows, eventId, absentName, block)
    .filter(row => !releasedOnly || String(row.Source || '') === 'Field Trip Pool');
  const entryByName = {};
  entries.forEach(row => {
    entryByName[coveragePersonNameKey_(row.Candidate)] = row;
  });

  return (coverageStaff || [])
    .filter(candidate => !!entryByName[coveragePersonNameKey_(candidate.name)])
    .map(candidate => {
      const row = entryByName[coveragePersonNameKey_(candidate.name)];
      const copy = Object.assign({}, candidate);
      copy.poolPriorityAdjustment = poolPriorityAdjustment_(row.Priority_Adjustment);
      copy.poolBaselineScore = Number(row.Baseline_Score || 0);
      copy.poolSource = String(row.Source || '');
      return copy;
    });
}

function menuRebuildFieldTripCoveragePool() {
  return withCoverageLock_(() => {
    const rows = rebuildAllFieldTripCoveragePool_();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Field Trip Coverage Pool rebuilt: ' + rows.length + ' candidate rows.',
      APP_TITLE,
      6
    );
    return rows.length;
  });
}

function candidateHasFieldTripEvent_(candidate, eventId) {
  const id = String(eventId || '').trim();
  if (!id) return false;
  return (candidate.fieldTripEvents || []).some(event => String(event.eventId || '').trim() === id);
}

function assignmentTypeIsBreak_(row) {
  return String(row && row.assignmentType || '').trim().toLowerCase() === 'break';
}

function assignmentTypeIsPlanning_(row) {
  return String(row && row.assignmentType || '').trim().toLowerCase() === 'planning';
}

function candidateBreakReservations_(candidateName, state) {
  if (!state || !state.breakReservationsByCandidate) return [];
  return state.breakReservationsByCandidate[candidateName] || [];
}

// A moved break must land in some open slot inside cancelled trip-grade time,
// not in one particular slot. A reservation remembers every slot it may use
// (eligibleRows); when a block needs its current spot, it can move again.
// Without this, whether a block "conflicted" with a moved break depended on
// the order assignments were made: the manual picker and save validation
// (which replay the plan without the row being checked) could pick a
// different spot than generation did and reject a valid plan.
function relocateBreakReservation_(candidateName, reservation, extraBusy, state) {
  const duration = reservation.endMinutes - reservation.startMinutes;
  if (duration <= 0 || !(reservation.eligibleRows || []).length) return null;
  const shadow = Object.assign({}, state, {
    assignmentsByCandidate: Object.assign({}, state.assignmentsByCandidate, {
      [candidateName]: ((state.assignmentsByCandidate || {})[candidateName] || []).concat(extraBusy || [])
    }),
    breakReservationsByCandidate: Object.assign({}, state.breakReservationsByCandidate, {
      [candidateName]: candidateBreakReservations_(candidateName, state).filter(item => item !== reservation)
    })
  });
  for (const row of reservation.eligibleRows) {
    const slot = findOpenBreakSlotInReleasedRow_(candidateName, row, duration, shadow);
    if (slot) return slot;
  }
  return null;
}

// Finds a slot for a newly moved break. If every slot is held by another
// moved break that could itself shift to a different cancelled slot, plans
// that shift (applied by recordAssignment_). Without this, whether a break
// could be moved depended on which assignment happened to be placed first.
function findBreakSlotAllowingShuffle_(candidateName, rows, duration, state, extraBusy) {
  for (const row of rows) {
    const slot = findOpenBreakSlotInReleasedRow_(candidateName, row, duration, state);
    if (slot) return { slot: slot, row: row, relocations: [] };
  }
  const reservations = candidateBreakReservations_(candidateName, state);
  for (const reservation of reservations) {
    if (!(reservation.eligibleRows || []).length) continue;
    const without = Object.assign({}, state, {
      breakReservationsByCandidate: Object.assign({}, state.breakReservationsByCandidate, {
        [candidateName]: reservations.filter(item => item !== reservation)
      })
    });
    for (const row of rows) {
      const slot = findOpenBreakSlotInReleasedRow_(candidateName, row, duration, without);
      if (!slot) continue;
      const moved = relocateBreakReservation_(candidateName, reservation, (extraBusy || []).concat([slot]), state);
      if (moved) {
        return {
          slot: slot,
          row: row,
          relocations: [{ reservation: reservation, startMinutes: moved.startMinutes, endMinutes: moved.endMinutes }]
        };
      }
    }
  }
  return null;
}

function candidateHasReservedBreakConflict_(candidateName, block, state) {
  const reservations = candidateBreakReservations_(candidateName, state);
  if (!reservations.some(item => timesOverlap_(item.startMinutes, item.endMinutes, block.startMinutes, block.endMinutes))) {
    return false;
  }
  // Try moving each blocking reservation elsewhere, on a copy.
  const simulated = reservations.map(item => Object.assign({}, item));
  const simState = Object.assign({}, state, {
    breakReservationsByCandidate: Object.assign({}, state.breakReservationsByCandidate, { [candidateName]: simulated })
  });
  for (const item of simulated) {
    if (!timesOverlap_(item.startMinutes, item.endMinutes, block.startMinutes, block.endMinutes)) continue;
    const slot = relocateBreakReservation_(candidateName, item, [block], simState);
    if (!slot) return true;
    item.startMinutes = slot.startMinutes;
    item.endMinutes = slot.endMinutes;
  }
  return false;
}

function intervalConflictsWithCandidateState_(candidateName, startMinutes, endMinutes, state) {
  const assignments = state && state.assignmentsByCandidate && state.assignmentsByCandidate[candidateName] || [];
  if (assignments.some(existing =>
    timesOverlap_(existing.startMinutes, existing.endMinutes, startMinutes, endMinutes)
  )) return true;

  return candidateBreakReservations_(candidateName, state).some(reservation =>
    timesOverlap_(reservation.startMinutes, reservation.endMinutes, startMinutes, endMinutes)
  );
}

function findOpenBreakSlotInReleasedRow_(candidateName, releasedRow, durationMinutes, state) {
  if (!releasedRow || durationMinutes <= 0) return null;
  if ((releasedRow.endMinutes - releasedRow.startMinutes) < durationMinutes) return null;

  const busy = [];
  const assignments = state && state.assignmentsByCandidate && state.assignmentsByCandidate[candidateName] || [];
  assignments.forEach(item => busy.push({
    startMinutes: item.startMinutes,
    endMinutes: item.endMinutes
  }));
  candidateBreakReservations_(candidateName, state).forEach(item => busy.push({
    startMinutes: item.startMinutes,
    endMinutes: item.endMinutes
  }));

  const absenceWindows = absenceWindowsForCandidate_(candidateName, state);
  for (const absence of absenceWindows) {
    if (absence.allDay) return null;
    if (absence.startMinutes == null || absence.endMinutes == null) return null;
    busy.push({
      startMinutes: absence.startMinutes,
      endMinutes: absence.endMinutes
    });
  }

  busy.sort((a, b) => a.startMinutes - b.startMinutes || a.endMinutes - b.endMinutes);

  let cursor = releasedRow.startMinutes;
  for (const interval of busy) {
    if (interval.endMinutes <= cursor) continue;
    if (interval.startMinutes >= releasedRow.endMinutes) break;
    if (interval.startMinutes - cursor >= durationMinutes) {
      return {
        startMinutes: cursor,
        endMinutes: cursor + durationMinutes
      };
    }
    cursor = Math.max(cursor, interval.endMinutes);
    if (cursor + durationMinutes > releasedRow.endMinutes) return null;
  }

  if (cursor + durationMinutes <= releasedRow.endMinutes) {
    return {
      startMinutes: cursor,
      endMinutes: cursor + durationMinutes
    };
  }
  return null;
}

function mergeTimeSegments_(segments) {
  const sorted = (segments || [])
    .filter(segment => segment && segment.endMinutes > segment.startMinutes)
    .sort((a, b) => a.startMinutes - b.startMinutes || a.endMinutes - b.endMinutes);

  const merged = [];
  sorted.forEach(segment => {
    const last = merged[merged.length - 1];
    if (!last || segment.startMinutes > last.endMinutes) {
      merged.push({
        startMinutes: segment.startMinutes,
        endMinutes: segment.endMinutes
      });
      return;
    }
    last.endMinutes = Math.max(last.endMinutes, segment.endMinutes);
  });
  return merged;
}

function uncoveredIntervals_(startMinutes, endMinutes, coveredSegments) {
  const merged = mergeTimeSegments_(coveredSegments);
  const gaps = [];
  let cursor = startMinutes;

  merged.forEach(segment => {
    if (segment.endMinutes <= cursor || segment.startMinutes >= endMinutes) return;
    const start = Math.max(segment.startMinutes, startMinutes);
    const end = Math.min(segment.endMinutes, endMinutes);
    if (start > cursor) gaps.push({ startMinutes: cursor, endMinutes: start });
    cursor = Math.max(cursor, end);
  });

  if (cursor < endMinutes) gaps.push({ startMinutes: cursor, endMinutes: endMinutes });
  return gaps.filter(gap => gap.endMinutes > gap.startMinutes);
}

function segmentsCoverIntervals_(segments, intervals) {
  const merged = mergeTimeSegments_(segments);
  return (intervals || []).every(interval => {
    let cursor = interval.startMinutes;
    for (const segment of merged) {
      if (segment.endMinutes <= cursor) continue;
      if (segment.startMinutes > cursor) return false;
      cursor = Math.max(cursor, segment.endMinutes);
      if (cursor >= interval.endMinutes) return true;
    }
    return cursor >= interval.endMinutes;
  });
}

function fieldTripReleasedRows_(availabilityRows, event) {
  return (availabilityRows || []).filter(row =>
    isInstructionalGradeBlock_(row) &&
    (event.grades || []).some(grade => normalizeGradeKey_(grade) === normalizeGradeKey_(row.grade)) &&
    row.startMinutes < event.endMinutes &&
    row.endMinutes > event.startMinutes
  );
}

function fieldTripCoverageComposition_(candidate, block, event, availabilityRows, state) {
  const releasedRows = fieldTripReleasedRows_(availabilityRows, event);
  const releasedSegments = releasedRows
    .map(row => ({
      startMinutes: Math.max(row.startMinutes, block.startMinutes),
      endMinutes: Math.min(row.endMinutes, block.endMinutes),
      row: row,
      kind: 'released'
    }))
    .filter(segment => segment.endMinutes > segment.startMinutes);

  const planningSegments = (availabilityRows || [])
    .filter(row => row.coverEligibleThisBlock && assignmentTypeIsPlanning_(row))
    .map(row => ({
      startMinutes: Math.max(row.startMinutes, block.startMinutes),
      endMinutes: Math.min(row.endMinutes, block.endMinutes),
      row: row,
      kind: 'planning'
    }))
    .filter(segment => segment.endMinutes > segment.startMinutes);

  const nonBreakSegments = releasedSegments.concat(planningSegments);
  const uncovered = uncoveredIntervals_(block.startMinutes, block.endMinutes, nonBreakSegments);

  if (!uncovered.length) {
    const releasedNames = releasedSegments
      .filter(segment => segment.startMinutes < block.endMinutes && segment.endMinutes > block.startMinutes)
      .map(segment => segment.row.className || (normalizeGradeKey_(segment.row.grade) + ' class'));
    const planningUsed = planningSegments.length > 0;

    const directReleaseWindows = mergeTimeSegments_(
      releasedRows.map(row => ({
        startMinutes: Math.max(row.startMinutes, event.startMinutes),
        endMinutes: Math.min(row.endMinutes, event.endMinutes)
      }))
    );
    const containingRelease = directReleaseWindows.find(window =>
      window.startMinutes <= block.startMinutes &&
      window.endMinutes >= block.endMinutes
    );
    const runwayMinutes = containingRelease
      ? Math.max(0, containingRelease.endMinutes - block.endMinutes)
      : 0;

    return {
      available: true,
      priority: releasedSegments.length ? 3 : 1,
      reason: releasedSegments.length
        ? 'Available because ' + Array.from(new Set(releasedNames)).join(' + ') +
          ' ' + (releasedNames.length > 1 ? 'are' : 'is') +
          ' cancelled by ' + (event.name || 'the field trip')
        : 'Available during planning; included in the field-trip pool because this teacher teaches ' +
          ((event.grades || []).join('/') || 'the trip grade'),
      breakMove: null,
      runwayMinutes: runwayMinutes
    };
  }

  const breakSegments = (availabilityRows || [])
    .filter(row => assignmentTypeIsBreak_(row))
    .map(row => ({
      startMinutes: Math.max(row.startMinutes, block.startMinutes),
      endMinutes: Math.min(row.endMinutes, block.endMinutes),
      row: row,
      kind: 'break'
    }))
    .filter(segment => segment.endMinutes > segment.startMinutes);

  if (!segmentsCoverIntervals_(breakSegments, uncovered)) {
    return { available: false, priority: 0, reason: '', breakMove: null };
  }

  const displacedBreakMinutes = uncovered.reduce(
    (sum, gap) => sum + (gap.endMinutes - gap.startMinutes),
    0
  );
  if (displacedBreakMinutes <= 0) {
    return { available: false, priority: 0, reason: '', breakMove: null };
  }

  const replacementAnchor = uncovered.length
    ? uncovered[uncovered.length - 1].endMinutes
    : block.endMinutes;

  const replacementRows = releasedRows
    .filter(row => !timesOverlap_(row.startMinutes, row.endMinutes, block.startMinutes, block.endMinutes))
    .map(row => Object.assign({}, row, {
      startMinutes: Math.max(row.startMinutes, event.startMinutes),
      endMinutes: Math.min(row.endMinutes, event.endMinutes)
    }))
    .filter(row => row.endMinutes > row.startMinutes)
    .sort((a, b) => {
      const aAfter = a.startMinutes >= replacementAnchor;
      const bAfter = b.startMinutes >= replacementAnchor;
      if (aAfter !== bAfter) return aAfter ? -1 : 1;
      if (aAfter) return a.startMinutes - b.startMinutes;
      return b.endMinutes - a.endMinutes;
    });

  const found = findBreakSlotAllowingShuffle_(candidate.name, replacementRows, displacedBreakMinutes, state, [block]);
  const replacement = found ? found.slot : null;
  const replacementRow = found ? found.row : null;

  if (!replacement || !replacementRow) {
    return { available: false, priority: 0, reason: '', breakMove: null };
  }

  const displacedText = uncovered
    .map(gap => minutesToDisplay_(gap.startMinutes) + '–' + minutesToDisplay_(gap.endMinutes))
    .join(' + ');

  const releasedUsed = releasedSegments
    .map(segment => segment.row.className || (normalizeGradeKey_(segment.row.grade) + ' class'));
  const releasedText = Array.from(new Set(releasedUsed)).join(' + ');

  const reason =
    (releasedText
      ? 'Available because ' + releasedText + ' is cancelled by ' + (event.name || 'the field trip') + '; '
      : '') +
    'break time used for ' + displacedText +
    ' and moved to ' +
    minutesToDisplay_(replacement.startMinutes) + '–' + minutesToDisplay_(replacement.endMinutes) +
    ' inside cancelled ' +
    (replacementRow.className || (normalizeGradeKey_(replacementRow.grade) + ' class'));

  return {
    available: true,
    priority: 2,
    reason: reason,
    runwayMinutes: 0,
    breakMove: {
      eventId: event.eventId,
      originalBreakStartMinutes: uncovered[0].startMinutes,
      originalBreakEndMinutes: uncovered[uncovered.length - 1].endMinutes,
      replacementStartMinutes: replacement.startMinutes,
      replacementEndMinutes: replacement.endMinutes,
      replacementGrade: normalizeGradeKey_(replacementRow.grade),
      replacementClass: replacementRow.className || replacementRow.subject || 'trip-grade class',
      // Every cancelled trip-grade slot (inside the trip window) this break
      // could live in, so it can move again if a later block needs this spot.
      eligibleRows: replacementRows.map(row => ({ startMinutes: row.startMinutes, endMinutes: row.endMinutes })),
      relocations: found.relocations,
      reason: reason
    }
  };
}

function findFieldTripBreakMove_(candidate, block, event, availabilityRows, state) {
  const currentBreak = availabilityRows.find(row =>
    assignmentTypeIsBreak_(row) &&
    row.startMinutes <= block.startMinutes &&
    row.endMinutes >= block.endMinutes
  );
  if (!currentBreak) return null;

  const breakDuration = currentBreak.endMinutes - currentBreak.startMinutes;
  if (breakDuration <= 0) return null;

  const releasedRows = availabilityRows
    .filter(row =>
      isInstructionalGradeBlock_(row) &&
      (event.grades || []).some(grade => normalizeGradeKey_(grade) === normalizeGradeKey_(row.grade)) &&
      row.startMinutes < event.endMinutes &&
      row.endMinutes > event.startMinutes &&
      !timesOverlap_(row.startMinutes, row.endMinutes, block.startMinutes, block.endMinutes)
    )
    .sort((a, b) => a.startMinutes - b.startMinutes);

  const found = findBreakSlotAllowingShuffle_(candidate.name, releasedRows, breakDuration, state, [block]);
  if (found) {
    const released = found.row;
    const slot = found.slot;
    return {
      eventId: event.eventId,
      originalBreakStartMinutes: currentBreak.startMinutes,
      originalBreakEndMinutes: currentBreak.endMinutes,
      replacementStartMinutes: slot.startMinutes,
      replacementEndMinutes: slot.endMinutes,
      replacementGrade: normalizeGradeKey_(released.grade),
      replacementClass: released.className || released.subject || 'trip-grade class',
      // Every cancelled trip-grade slot this break could live in.
      eligibleRows: releasedRows.map(row => ({ startMinutes: row.startMinutes, endMinutes: row.endMinutes })),
      relocations: found.relocations,
      reason:
        'Break moved from ' +
        minutesToDisplay_(currentBreak.startMinutes) + '–' + minutesToDisplay_(currentBreak.endMinutes) +
        ' to ' +
        minutesToDisplay_(slot.startMinutes) + '–' + minutesToDisplay_(slot.endMinutes) +
        ' because ' + (released.className || (normalizeGradeKey_(released.grade) + ' class')) +
        ' is cancelled by ' + (event.name || 'the field trip')
    };
  }

  return null;
}

function candidateFieldTripAvailability_(candidate, block, availabilityRows, state) {
  if (!block.fieldTripEventId || !candidateHasFieldTripEvent_(candidate, block.fieldTripEventId)) {
    return { available: false, fieldTripPriority: 0, fieldTripReason: '', fieldTripBreakMove: null, fieldTripRunwayMinutes: 0 };
  }

  const event = (candidate.fieldTripEvents || []).find(item => item.eventId === block.fieldTripEventId);
  if (!event) {
    return { available: false, fieldTripPriority: 0, fieldTripReason: '', fieldTripBreakMove: null, fieldTripRunwayMinutes: 0 };
  }

  if (candidateHasReservedBreakConflict_(candidate.name, block, state)) {
    return {
      available: false,
      fieldTripPriority: 0,
      fieldTripReason: 'Reserved as replacement break for an earlier field-trip coverage assignment',
      fieldTripBreakMove: null,
      fieldTripRunwayMinutes: 0
    };
  }

  const composition = fieldTripCoverageComposition_(
    candidate,
    block,
    event,
    availabilityRows,
    state
  );

  return {
    available: !!composition.available,
    fieldTripPriority: Number(composition.priority || 0),
    fieldTripReason: composition.reason || '',
    fieldTripBreakMove: composition.breakMove || null,
    fieldTripRunwayMinutes: Number(composition.runwayMinutes || 0)
  };
}

function planRowToCoverageBlock_(row) {
  row = row || {};
  return {
    staffName: String(row.Absent_Staff || '').trim(),
    startMinutes: displayTimeToMinutes_(row.Start),
    endMinutes: displayTimeToMinutes_(row.End),
    className: String(row.Class || '').trim(),
    grade: String(row.Grade || '').trim() || inferGradeFromClass_(row.Class || ''),
    subject: String(row.Subject || '').trim(),
    assignmentType: String(row.Assignment_Type || '').trim() || 'Class',
    room: String(row.Room || '').trim(),
    fieldTripEventId: String(row.Event_ID || '').trim(),
    emergencyOverride: false
  };
}

function makeManualScheduleCandidate_(name, role) {
  return {
    name: String(name || '').trim(),
    role: String(role || 'Staff').trim(),
    tier: 3,
    canCoverAllDay: false,
    baseActive: true,
    activeToday: true,
    availableDays: '*',
    defaultStart: '',
    defaultEnd: '',
    selectedStart: '',
    selectedEnd: '',
    hasDateOverride: false,
    availabilityNotes: '',
    allowedGrades: '*',
    allowedSubjects: '*',
    allowedAssignmentTypes: '*',
    maxBlocksPerDay: Infinity,
    maxTeachersPerDay: Infinity,
    canBeSplitAcrossTeachers: true,
    notes: 'Manual placement candidate derived from Teacher Schedule.',
    fieldTripEvents: [],
    fieldTripOnly: false,
    manualSource: 'Available Staff'
  };
}

function buildManualCoverageCandidates_(date, day, config, fieldTrips, teacherSchedule, fieldTripPoolRows) {
  const configuredCoverageStaff = getCoverageStaffForDate_(date, day, config);
  const activeCoverageStaff = configuredCoverageStaff.filter(row => row.name && row.activeToday);

  const poolRows = fieldTrips.length ? (fieldTripPoolRows || fieldTripCoveragePoolRowsForDate_(date)) : [];
  const automaticPool = buildFieldTripCoverageCandidatesFromPool_(
    fieldTrips,
    date,
    activeCoverageStaff,
    configuredCoverageStaff,
    poolRows
  );

  const configuredNames = {};
  configuredCoverageStaff.forEach(candidate => {
    if (candidate && candidate.name) configuredNames[candidate.name] = candidate;
  });

  const byName = {};
  automaticPool.forEach(candidate => {
    const copy = Object.assign({}, candidate);
    const role = String(copy.role || '').trim().toLowerCase();
    copy.manualSource = (copy.fieldTripEvents || []).length
      ? 'Field Trip Pool'
      : (role === 'substitute' ? 'Subs' : 'Coverage Staff');
    byName[copy.name] = copy;
  });

  teacherSchedule
    .map(row => normalizeTeacherScheduleRow_(row))
    .filter(row => row.day === day && row.staffName)
    .forEach(row => {
      if (byName[row.staffName]) return;

      // If this person is explicitly managed in Coverage Staff and has been
      // marked unavailable for the date, do not reintroduce them through their
      // Teacher Schedule row.
      if (configuredNames[row.staffName] && !configuredNames[row.staffName].activeToday) return;

      const candidate = makeManualScheduleCandidate_(row.staffName, row.role || 'Staff');
      if (configuredNames[row.staffName]) {
        candidate.manualSource = String(configuredNames[row.staffName].role || '').trim().toLowerCase() === 'substitute'
          ? 'Subs'
          : 'Coverage Staff';
      }
      byName[row.staffName] = candidate;
    });

  return Object.keys(byName)
    .map(name => byName[name])
    .filter(candidate => candidate && candidate.name);
}

function manualCoverageStateFromPlan_(planRows, excludedIndex, candidates, teacherSchedule, day, effectiveAbsences) {
  const state = makeEmptyState_();
  state.absencesByCandidate = buildAbsenceWindowsByStaff_(effectiveAbsences);

  const candidateByName = {};
  (candidates || []).forEach(candidate => {
    if (candidate && candidate.name) candidateByName[candidate.name] = candidate;
  });

  (planRows || [])
    .map((row, index) => ({ row: row || {}, index: index }))
    .filter(item =>
      item.index !== excludedIndex &&
      String(item.row.Status || '').trim() === 'Assigned' &&
      String(item.row.Assigned_Coverage || '').trim()
    )
    .sort((a, b) => {
      const ab = planRowToCoverageBlock_(a.row);
      const bb = planRowToCoverageBlock_(b.row);
      return (ab.startMinutes || 0) - (bb.startMinutes || 0) ||
        (ab.endMinutes || 0) - (bb.endMinutes || 0) ||
        a.index - b.index;
    })
    .forEach(item => {
      const name = String(item.row.Assigned_Coverage || '').trim();
      const block = planRowToCoverageBlock_(item.row);
      if (!name || block.startMinutes == null || block.endMinutes == null) return;

      let candidate = candidateByName[name] || makeManualScheduleCandidate_(name, 'Staff');

      // Reconstruct any field-trip break reservation used by an existing row
      // so later manual choices cannot consume that replacement break.
      const availability = candidateAvailabilityForBlock_(
        candidate,
        block,
        teacherSchedule,
        day,
        state
      );
      if (availability && availability.fieldTripBreakMove) {
        candidate = Object.assign({}, candidate, {
          fieldTripBreakMove: availability.fieldTripBreakMove
        });
      }

      recordAssignment_(
        state,
        candidate,
        block,
        String(item.row.Absent_Staff || '').trim()
      );
    });

  return state;
}

function coverageNeedKey_(staffName, startMinutes, endMinutes, className) {
  return [
    String(staffName || '').trim(),
    startMinutes,
    endMinutes,
    String(className == null ? '' : className).trim()
  ].join('\u0000');
}

// Everything the manual picker and save-time validation need from the live
// workbook, read once. needByKey maps each current coverage need to its
// block so a plan row can recover facts the row itself doesn't store
// (notably whether the absence is an emergency).
function buildCoverageLiveContext_(date, day) {
  const config = getConfigMap_();
  const teacherSchedule = teacherScheduleRowsForDate_(date, day, config);
  const absences = getDailyAbsencesForDate_(date, day);
  const fieldTrips = getFieldTripsForDate_(date);
  const fieldTripPoolRows = fieldTrips.length ? fieldTripCoveragePoolRowsForDate_(date) : [];
  const effectiveAbsences = absences.concat(buildFieldTripParticipantAbsences_(fieldTrips, teacherSchedule, day));
  const candidates = buildManualCoverageCandidates_(date, day, config, fieldTrips, teacherSchedule, fieldTripPoolRows);
  const needsByTeacher = buildCoverageNeedsByTeacher_(effectiveAbsences, teacherSchedule, day, fieldTrips);

  const needByKey = {};
  const needCounts = {};
  Object.keys(needsByTeacher).forEach(name => {
    (needsByTeacher[name] || []).forEach(need => {
      const key = coverageNeedKey_(name, need.startMinutes, need.endMinutes, need.className);
      needByKey[key] = need;
      needCounts[key] = (needCounts[key] || 0) + 1;
    });
  });

  return {
    date: date,
    day: day,
    config: config,
    teacherSchedule: teacherSchedule,
    absences: absences,
    fieldTrips: fieldTrips,
    fieldTripPoolRows: fieldTripPoolRows,
    effectiveAbsences: effectiveAbsences,
    candidates: candidates,
    needsByTeacher: needsByTeacher,
    needByKey: needByKey,
    needCounts: needCounts
  };
}

// Plan rows don't record emergency status, so the manual picker and save
// validation look it up from the live need. Without this, an emergency block
// was judged by non-emergency rules: the picker offered nobody the scheduler
// itself had been allowed to use, and flagged the scheduler's own choice as
// "currently unavailable".
function planRowToLiveBlock_(row, liveContext) {
  const block = planRowToCoverageBlock_(row);
  const need = liveContext.needByKey[
    coverageNeedKey_(block.staffName, block.startMinutes, block.endMinutes, block.className)
  ];
  if (need) block.emergencyOverride = !!need.emergencyOverride;
  return block;
}

function manualCoverageContext_(payload) {
  payload = payload || {};
  const date = normalizeDateKey_(payload.date);
  const day = String(payload.day || guessDayCodeFromDate_(date) || '').trim();
  const planRows = Array.isArray(payload.rows) ? payload.rows : [];
  const blockIndex = Number(payload.blockIndex);
  if (!date) throw new Error('Choose a valid date before reassigning coverage.');
  if (!day) throw new Error('The selected date does not have a valid school-day code.');
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= planRows.length) {
    throw new Error('The coverage block could not be identified.');
  }

  const live = buildCoverageLiveContext_(date, day);
  const row = planRows[blockIndex] || {};
  const block = planRowToLiveBlock_(row, live);
  if (block.startMinutes == null || block.endMinutes == null) {
    throw new Error('The selected coverage block has an invalid start or end time.');
  }

  const state = manualCoverageStateFromPlan_(
    planRows,
    blockIndex,
    live.candidates,
    live.teacherSchedule,
    day,
    live.effectiveAbsences
  );

  return Object.assign({}, live, {
    row: row,
    block: block,
    state: state
  });
}

// Server-side check run on every save. The plan in the browser can outlive
// the data it was built from (absences or trips edited afterwards, a
// coverage person switched off, a stale preview restored on reload), and the
// save path used to trust whatever rows it was sent.
function validateCoveragePlanForSave_(date, day, rows) {
  const live = buildCoverageLiveContext_(date, day);
  const planCounts = {};
  const planLabels = {};
  rows.forEach(row => {
    const block = planRowToCoverageBlock_(row);
    const key = coverageNeedKey_(block.staffName, block.startMinutes, block.endMinutes, block.className);
    planCounts[key] = (planCounts[key] || 0) + 1;
    planLabels[key] = block.staffName + ' ' + String(row.Start || '') + '–' + String(row.End || '');
  });

  const describeNeed = key => {
    const need = live.needByKey[key];
    return need
      ? need.staffName + ' ' + minutesToDisplay_(need.startMinutes) + '–' + minutesToDisplay_(need.endMinutes)
      : planLabels[key];
  };
  const missing = Object.keys(live.needCounts).filter(key => (planCounts[key] || 0) < live.needCounts[key]);
  const extra = Object.keys(planCounts).filter(key => planCounts[key] > (live.needCounts[key] || 0));
  if (missing.length || extra.length) {
    const parts = [];
    if (missing.length) parts.push('not in this plan: ' + missing.slice(0, 4).map(describeNeed).join('; '));
    if (extra.length) parts.push('no longer needed: ' + extra.slice(0, 4).map(describeNeed).join('; '));
    throw new Error(
      'This plan is out of date — absences, field trips, or the schedule changed after it was generated (' +
      parts.join(' | ') + '). Click Re-generate, review, then save again.'
    );
  }

  const candidateByName = {};
  live.candidates.forEach(candidate => { candidateByName[candidate.name] = candidate; });
  const configuredCoverageNames = new Set(
    getCoverageStaffForDate_(date, day, live.config).map(candidate => candidate.name)
  );

  const problems = [];
  rows.forEach((row, index) => {
    const name = String(row.Assigned_Coverage || '').trim();
    if (String(row.Status || '').trim() !== 'Assigned' || !name) return;

    const block = planRowToLiveBlock_(row, live);
    const absentName = String(row.Absent_Staff || '').trim();
    const label = name + ' at ' + String(row.Start || '') + ' (' + absentName + ')';
    const candidate = candidateByName[name];
    if (!candidate) {
      problems.push(label + (configuredCoverageNames.has(name)
        ? ': marked unavailable for this date'
        : ': not in Coverage Staff or the Teacher Schedule'));
      return;
    }
    if (name === absentName) {
      problems.push(label + ': cannot cover their own absence');
      return;
    }
    if (block.fieldTripEventId) {
      const allowed = fieldTripCoveragePoolEntriesForBlock_(
        live.fieldTripPoolRows,
        block.fieldTripEventId,
        absentName,
        block
      ).some(entry => coveragePersonNameKey_(entry.Candidate) === coveragePersonNameKey_(name));
      if (!allowed) {
        problems.push(label + ': disabled or no longer listed in the Field Trip Coverage Pool');
        return;
      }
    }

    const state = manualCoverageStateFromPlan_(rows, index, live.candidates, live.teacherSchedule, day, live.effectiveAbsences);
    if (candidateIsAbsentForBlock_(name, block, state)) {
      problems.push(label + ': absent or on a field trip at that time');
    } else if (!candidateCanCoverBlock_(candidate, absentName, block, live.teacherSchedule, day, state, live.config)) {
      problems.push(label + ': no longer available (availability, schedule, limits, or another assignment)');
    }
  });

  if (problems.length) {
    throw new Error(
      'This plan can\'t be saved because ' + problems.length + ' assignment' + (problems.length === 1 ? ' is' : 's are') +
      ' no longer valid: ' + problems.slice(0, 5).join('; ') + (problems.length > 5 ? '; and ' + (problems.length - 5) + ' more' : '') +
      '. Reassign those blocks or click Re-generate, then save again.'
    );
  }
}

// A person's day, simplified for the inline schedule preview in the manual
// reassignment modal: one entry per Teacher Schedule row, in display time.
// "cancelled" marks a row a field trip has cancelled today (students away) —
// shown differently from an ordinary break, since it explains why someone who
// normally teaches then is free today specifically.
//
// breakReservations (state.breakReservationsByCandidate[staffName], optional)
// lets this also reflect a MOVED break: without it, a break that's actually
// being used to cover something else still shows as plain "Break" (looking
// free when it isn't), and the cancelled slot it moved to shows as generic
// "freed by a trip" with no indication it's now this person's real break.
function describeStaffDayForDisplay_(teacherSchedule, day, staffName, fieldTrips, breakReservations) {
  const entries = normalizedScheduleRowsFor_(teacherSchedule, staffName, day)
    .filter(row => row.startMinutes != null && row.endMinutes != null)
    .slice()
    .sort((a, b) => a.startMinutes - b.startMinutes)
    .map(row => ({
      start: minutesToDisplay_(row.startMinutes),
      end: minutesToDisplay_(row.endMinutes),
      startMinutes: row.startMinutes,
      endMinutes: row.endMinutes,
      label: row.className || row.subject || row.assignmentType || 'Class',
      type: row.assignmentType || 'Class',
      cancelled: blockIsCancelledByFieldTrip_(row, fieldTrips)
    }));

  (breakReservations || []).forEach(reservation => {
    entries.forEach(entry => {
      if (timesOverlap_(entry.startMinutes, entry.endMinutes, reservation.originalBreakStartMinutes, reservation.originalBreakEndMinutes)) {
        entry.label = entry.label + ' — moved to ' + minutesToDisplay_(reservation.startMinutes);
        entry.breakMovedAway = true;
      }
      if (timesOverlap_(entry.startMinutes, entry.endMinutes, reservation.startMinutes, reservation.endMinutes)) {
        entry.label = 'Break (moved from ' + minutesToDisplay_(reservation.originalBreakStartMinutes) + ')';
        entry.breakMovedHere = true;
      }
    });
  });

  return entries.map(entry => {
    delete entry.startMinutes;
    delete entry.endMinutes;
    return entry;
  });
}

function getManualCoverageChoices_(payload) {
  const context = manualCoverageContext_(payload);
  const row = context.row;
  const block = context.block;
  const absentName = String(row.Absent_Staff || '').trim();
  const currentName = String(row.Assigned_Coverage || '').trim();
  const choices = [];
  const absentDuringBlock = [];
  const poolEntries = block.fieldTripEventId
    ? fieldTripCoveragePoolEntriesForBlock_(
        context.fieldTripPoolRows,
        block.fieldTripEventId,
        absentName,
        block
      )
    : [];
  const poolEntryByName = {};
  poolEntries.forEach(entry => {
    poolEntryByName[coveragePersonNameKey_(entry.Candidate)] = entry;
  });

  context.candidates.forEach(candidate => {
    if (!candidate || !candidate.name || candidate.name === absentName) return;
    const poolEntry = block.fieldTripEventId
      ? poolEntryByName[coveragePersonNameKey_(candidate.name)]
      : null;
    if (block.fieldTripEventId && !poolEntry) return;

    if (candidateIsAbsentForBlock_(candidate.name, block, context.state)) {
      absentDuringBlock.push(candidate.name);
      return;
    }

    if (!candidateCanCoverBlock_(
      candidate,
      absentName,
      block,
      context.teacherSchedule,
      context.day,
      context.state,
      context.config
    )) return;

    const availability = candidateAvailabilityForBlock_(
      candidate,
      block,
      context.teacherSchedule,
      context.day,
      context.state
    );

    const fieldTripPool = !!(
      block.fieldTripEventId &&
      candidateHasFieldTripEvent_(candidate, block.fieldTripEventId)
    );
    const source = poolEntry
      ? String(poolEntry.Source || (fieldTripPool ? 'Field Trip Pool' : 'Coverage Staff'))
      : (fieldTripPool ? 'Field Trip Pool' : (candidate.manualSource || 'Available Staff'));

    const scoreInfo = scoreCandidateForBlock_(candidate, absentName, block, context.state);
    const fieldTripBoost = Number(availability.fieldTripPriority || 0) * 250;
    const runwayBoost = Math.min(Math.max(0, Number(availability.fieldTripRunwayMinutes || 0)), 60);

    choices.push({
      name: candidate.name,
      role: candidate.role || '',
      tier: candidate.tier,
      source: source,
      recommended: fieldTripPool,
      reason: availability.fieldTripReason ||
        (candidate.canCoverAllDay
          ? 'Available as Coverage Staff for the full block.'
          : 'Available during a cover-eligible schedule block.'),
      score: scoreInfo.score + fieldTripBoost + runwayBoost + poolPriorityAdjustment_(poolEntry && poolEntry.Priority_Adjustment)
    });
  });

  choices.sort((a, b) => {
    const sourceRank = source =>
      source === 'Field Trip Pool' ? 0 :
      source === 'Coverage Staff' ? 1 :
      source === 'Available Staff' ? 2 :
      source === 'Subs' ? 3 : 4;
    return sourceRank(a.source) - sourceRank(b.source) ||
      b.score - a.score ||
      a.name.localeCompare(b.name);
  });

  // One schedule per name that could appear in the dropdown, including the
  // current assignee even if they're shown disabled as "currently
  // unavailable" — seeing their day often explains why.
  const namesToDescribe = new Set(choices.map(choice => choice.name));
  if (currentName) namesToDescribe.add(currentName);
  const schedules = {};
  namesToDescribe.forEach(name => {
    let reservations = candidateBreakReservations_(name, context.state);
    // The state replay above deliberately excludes this block's own current
    // assignment (so every other candidate can be evaluated as if it were
    // open) — which means if this person is the current assignee, their own
    // break move caused by filling this exact block is missing from
    // context.state. Compute it fresh, just for display.
    if (name === currentName) {
      const candidateForDisplay = context.candidates.find(c => c.name === name) || makeManualScheduleCandidate_(name, 'Staff');
      const freshAvailability = candidateAvailabilityForBlock_(
        candidateForDisplay,
        block,
        context.teacherSchedule,
        context.day,
        context.state
      );
      if (freshAvailability && freshAvailability.fieldTripBreakMove) {
        const move = freshAvailability.fieldTripBreakMove;
        reservations = reservations.concat([{
          eventId: move.eventId,
          startMinutes: move.replacementStartMinutes,
          endMinutes: move.replacementEndMinutes,
          originalBreakStartMinutes: move.originalBreakStartMinutes,
          originalBreakEndMinutes: move.originalBreakEndMinutes
        }]);
      }
    }
    schedules[name] = describeStaffDayForDisplay_(
      context.teacherSchedule,
      context.day,
      name,
      context.fieldTrips,
      reservations
    );
  });

  return {
    date: context.date,
    day: context.day,
    blockIndex: Number(payload.blockIndex),
    currentName: currentName,
    currentEligible: !currentName || choices.some(choice => choice.name === currentName),
    choices: choices,
    schedules: schedules,
    excludedAbsentNames: Array.from(new Set(absentDuringBlock)).sort()
  };
}

function validateManualCoverageAssignment_(payload) {
  payload = payload || {};
  const name = String(payload.name || '').trim();

  if (!name) {
    return {
      valid: true,
      name: '',
      tier: '',
      source: '',
      reason: 'Manually left unfilled.'
    };
  }

  const result = getManualCoverageChoices_(payload);
  const choice = result.choices.find(item => item.name === name);
  if (!choice) {
    if (result.excludedAbsentNames.indexOf(name) !== -1) {
      throw new Error(name + ' is absent during this coverage block and cannot be assigned.');
    }
    throw new Error(name + ' is not available for this coverage block because of schedule, absence, trip, or another coverage conflict.');
  }

  return {
    valid: true,
    name: choice.name,
    tier: choice.tier,
    source: choice.source,
    reason: 'Manually assigned. ' + choice.reason
  };
}

function scheduleFieldTripNeedsChronologically_(
  needsByTeacher,
  coverageStaff,
  teacherSchedule,
  day,
  state,
  config,
  date,
  planRows,
  summary,
  fieldTripPoolRows
) {
  const tripNeeds = [];

  Object.keys(needsByTeacher || {}).forEach(absentName => {
    (needsByTeacher[absentName] || []).forEach(block => {
      if (!block.fieldTripEventId) return;
      tripNeeds.push({
        absentName: absentName,
        block: block
      });
    });
  });

  tripNeeds.sort((a, b) =>
    a.block.startMinutes - b.block.startMinutes ||
    a.block.endMinutes - b.block.endMinutes ||
    a.absentName.localeCompare(b.absentName)
  );

  const deferred = [];
  const assignedAbsentNames = {};

  // First pass: use only people released by this field trip. This lets normal
  // Coverage Staff remain available for ordinary absences unless needed later.
  tripNeeds.forEach(item => {
    const eventPool = fieldTripPoolCandidateSubset_(
      fieldTripPoolRows,
      item.block.fieldTripEventId,
      item.absentName,
      item.block,
      coverageStaff,
      true
    );
    const best = pickBestBlockCandidate_(
      item.absentName,
      item.block,
      eventPool,
      teacherSchedule,
      day,
      state,
      config
    );

    if (!best) {
      deferred.push(item);
      return;
    }

    planRows.push(makePlanRow_(
      date,
      day,
      item.block,
      best,
      'Field Trip Coverage',
      'Assigned',
      best.reason
    ));
    recordAssignment_(state, best, item.block, item.absentName);
    summary.assignedBlocks += 1;
    assignedAbsentNames[item.absentName] = true;
  });

  return {
    deferred: deferred,
    assignedAbsentNames: assignedAbsentNames,
    totalBlocks: tripNeeds.length
  };
}

function fillDeferredFieldTripNeeds_(
  deferred,
  coverageStaff,
  teacherSchedule,
  day,
  state,
  config,
  date,
  planRows,
  summary,
  fieldTripPoolRows
) {
  const assignedAbsentNames = {};

  (deferred || []).forEach(item => {
    const pooledCandidates = fieldTripPoolCandidateSubset_(
      fieldTripPoolRows,
      item.block.fieldTripEventId,
      item.absentName,
      item.block,
      coverageStaff,
      false
    );
    const best = pickBestBlockCandidate_(
      item.absentName,
      item.block,
      pooledCandidates,
      teacherSchedule,
      day,
      state,
      config
    );

    if (best) {
      planRows.push(makePlanRow_(
        date,
        day,
        item.block,
        best,
        'Field Trip Coverage',
        'Assigned',
        best.reason
      ));
      recordAssignment_(state, best, item.block, item.absentName);
      summary.assignedBlocks += 1;
      assignedAbsentNames[item.absentName] = true;
    } else {
      planRows.push(makePlanRow_(
        date,
        day,
        item.block,
        null,
        'Field Trip Coverage',
        'Unfilled',
        'No eligible field-trip coverage person or fallback coverage person found.'
      ));
      summary.unfilledBlocks += 1;
    }
  });

  return assignedAbsentNames;
}

function generateCoveragePreview(payload) {
  const generateStartedAt = Date.now();
  payload = payload || {};
  const date = payload.date || Utilities.formatDate(new Date(), coverageTimeZone_(), 'yyyy-MM-dd');
  const day = payload.day || guessDayCodeFromDate_(date);

  primeCoverageRequestSnapshot_([
    getCoverageStaffSheetName_(),
    'Substitute Availability',
    'Daily Absences',
    'Field Trips',
    'Config'
  ]);
  coveragePerfMark_('snapshot-loaded');

  const config = getConfigMap_();
  const teacherSchedule = teacherScheduleRowsForDate_(date, day, config);
  const configuredCoverageStaff = getCoverageStaffForDate_(date, day, config);
  const activeCoverageStaff = configuredCoverageStaff.filter(row => row.name && row.activeToday);
  const absences = getDailyAbsencesForDate_(date, day);
  const fieldTrips = getFieldTripsForDate_(date);
  const fieldTripAbsences = buildFieldTripParticipantAbsences_(fieldTrips, teacherSchedule, day);
  const effectiveAbsences = absences.concat(fieldTripAbsences);
  const fieldTripPoolRows = fieldTrips.length ? fieldTripCoveragePoolRowsForDate_(date) : [];
  const coverageStaff = buildFieldTripCoverageCandidatesFromPool_(
    fieldTrips,
    date,
    activeCoverageStaff,
    configuredCoverageStaff,
    fieldTripPoolRows
  );

  const needsByTeacher = buildCoverageNeedsByTeacher_(
    effectiveAbsences,
    teacherSchedule,
    day,
    fieldTrips
  );

  const state = makeEmptyState_();
  state.absencesByCandidate = buildAbsenceWindowsByStaff_(effectiveAbsences);
  const planRows = [];
  const summary = {
    totalAbsentStaff: Object.keys(needsByTeacher).filter(name => (needsByTeacher[name] || []).length).length,
    totalBlocks: 0,
    assignedBlocks: 0,
    unfilledBlocks: 0,
    wholeDayAssignments: 0,
    splitAssignments: 0,
    fieldTrips: fieldTrips.length
  };

  // ── Manual/preferred ordinary assignments first ──
  // Each need carries the preferred person from the absence window it came
  // from, so two windows for the same person can prefer different people and
  // a preference never spills onto another window's blocks.
  const preferredAssigned = new Set();
  const preferredGroups = {};
  Object.keys(needsByTeacher).forEach(absentName => {
    (needsByTeacher[absentName] || []).forEach(block => {
      if (block.fieldTripEventId || !block.preferredCoverage) return;
      const key = absentName + '\u0000' + block.preferredCoverage;
      if (!preferredGroups[key]) {
        preferredGroups[key] = { absentName: absentName, preferred: block.preferredCoverage, blocks: [] };
      }
      preferredGroups[key].blocks.push(block);
    });
  });

  Object.keys(preferredGroups).sort().forEach(key => {
    const group = preferredGroups[key];
    const absentName = group.absentName;
    const preferred = group.preferred;
    const blocks = group.blocks.slice().sort((a, b) => a.startMinutes - b.startMinutes);
    const candidate = coverageStaff.find(c => c.name === preferred);

    // A preferred/manual person is a preference, not permission to bypass
    // absences, schedule conflicts, availability windows, or daily limits.
    if (!candidate || !candidateCanCoverAllBlocks_(
      candidate,
      absentName,
      blocks,
      teacherSchedule,
      day,
      state,
      config
    )) {
      return;
    }

    const candidateInfo = {
      name: candidate.name,
      tier: candidate.tier,
      role: candidate.role
    };

    blocks.forEach(block => {
      planRows.push(makePlanRow_(date, day, block, candidateInfo, 'Manual', 'Assigned', 'Manually preferred: ' + preferred + '.'));
      recordAssignment_(state, candidate, block, absentName);
      preferredAssigned.add(block);
    });

    summary.totalBlocks += blocks.length;
    summary.assignedBlocks += blocks.length;
  });

  // ── Field trip first pass: event-created staff pool, global chronological order ──
  const tripPass = scheduleFieldTripNeedsChronologically_(
    needsByTeacher,
    coverageStaff,
    teacherSchedule,
    day,
    state,
    config,
    date,
    planRows,
    summary,
    fieldTripPoolRows
  );
  summary.totalBlocks += tripPass.totalBlocks;
  if (Object.keys(tripPass.assignedAbsentNames).length) {
    summary.splitAssignments += Object.keys(tripPass.assignedAbsentNames).length;
  }

  // ── Ordinary automatic coverage ──
  const orderedNeeds = Object.keys(needsByTeacher)
    .map(absentName => ({
      absentName: absentName,
      blocks: (needsByTeacher[absentName] || []).filter(block =>
        !block.fieldTripEventId && !preferredAssigned.has(block)
      )
    }))
    .filter(item => item.blocks.length)
    .map(item => ({
      absentName: item.absentName,
      blocks: item.blocks,
      difficulty: estimateDifficulty_(item.absentName, item.blocks, coverageStaff, teacherSchedule, day, config)
    }))
    .sort((a, b) =>
      a.difficulty.fullDayCandidateCount - b.difficulty.fullDayCandidateCount ||
      b.blocks.length - a.blocks.length ||
      a.absentName.localeCompare(b.absentName)
    );

  orderedNeeds.forEach(item => {
    const absentName = item.absentName;
    const blocks = item.blocks.slice().sort((a, b) => a.startMinutes - b.startMinutes);
    summary.totalBlocks += blocks.length;
    if (!blocks.length) return;

    let assignedWholeDay = false;

    if (normalizeYesNo_(config.Whole_Day_First, true)) {
      const fullDayCandidate = pickBestWholeDayCandidate_(
        absentName,
        blocks,
        coverageStaff,
        teacherSchedule,
        day,
        state,
        config
      );
      if (fullDayCandidate) {
        blocks.forEach(block => {
          planRows.push(makePlanRow_(
            date,
            day,
            block,
            fullDayCandidate,
            'Whole Day',
            'Assigned',
            fullDayCandidate.reason || 'Whole-day assignment.'
          ));
          recordAssignment_(state, fullDayCandidate, block, absentName);
        });
        assignedWholeDay = true;
        summary.assignedBlocks += blocks.length;
        summary.wholeDayAssignments += 1;
      }
    }

    if (assignedWholeDay) return;

    if (!normalizeYesNo_(config.Allow_Split_Coverage, true)) {
      blocks.forEach(block => {
        planRows.push(makePlanRow_(
          date,
          day,
          block,
          null,
          'Unfilled',
          'Unfilled',
          'Split coverage disabled and no full-day match found.'
        ));
        summary.unfilledBlocks += 1;
      });
      return;
    }

    let usedSomeone = false;
    const remainingBlocks = [];
    const primary = pickPrimarySplitCandidate_(
      absentName,
      blocks,
      coverageStaff,
      teacherSchedule,
      day,
      state,
      config
    );

    if (primary) {
      blocks.forEach(block => {
        if (candidateCanCoverBlock_(primary, absentName, block, teacherSchedule, day, state, config)) {
          planRows.push(makePlanRow_(
            date,
            day,
            block,
            primary,
            'Split Coverage',
            'Assigned',
            primary.reason
          ));
          recordAssignment_(state, primary, block, absentName);
          summary.assignedBlocks += 1;
          usedSomeone = true;
        } else {
          remainingBlocks.push(block);
        }
      });
    } else {
      remainingBlocks.push(...blocks);
    }

    remainingBlocks.forEach(block => {
      const best = pickBestBlockCandidate_(
        absentName,
        block,
        coverageStaff,
        teacherSchedule,
        day,
        state,
        config
      );
      if (best) {
        planRows.push(makePlanRow_(
          date,
          day,
          block,
          best,
          'Split Coverage',
          'Assigned',
          best.reason
        ));
        recordAssignment_(state, best, block, absentName);
        summary.assignedBlocks += 1;
        usedSomeone = true;
      } else {
        planRows.push(makePlanRow_(
          date,
          day,
          block,
          null,
          'Split Coverage',
          'Unfilled',
          'No eligible coverage person found.'
        ));
        summary.unfilledBlocks += 1;
      }
    });

    if (usedSomeone) summary.splitAssignments += 1;
  });

  // ── Field trip fallback: regular Coverage Staff only after ordinary needs ──
  const fallbackAssigned = fillDeferredFieldTripNeeds_(
    tripPass.deferred,
    coverageStaff,
    teacherSchedule,
    day,
    state,
    config,
    date,
    planRows,
    summary,
    fieldTripPoolRows
  );
  if (Object.keys(fallbackAssigned).length) {
    summary.splitAssignments += Object.keys(fallbackAssigned).length;
  }

  coveragePerfMark_('schedule-built');
  writePreview_(planRows);
  coveragePerfMark_('preview-written');
  summary.elapsedMs = Date.now() - generateStartedAt;

  return {
    date: date,
    day: day,
    summary: summary,
    rows: planRows,
    absences: absences,
    fieldTrips: fieldTrips,
    activeCoverageStaff: coverageStaff.map(row => ({
      name: row.name,
      tier: row.tier,
      role: row.role,
      allDay: row.canCoverAllDay,
      activeToday: !!row.activeToday,
      fieldTripOnly: !!row.fieldTripOnly,
      fieldTripEvents: (row.fieldTripEvents || []).map(event => ({
        eventId: event.eventId,
        name: event.name
      }))
    }))
  };
}

function saveCoveragePlan(payload) {
  return withCoverageLock_(() => saveCoveragePlanUnlocked_(payload));
}

function saveCoveragePlanUnlocked_(payload) {
  payload = payload || {};
  const rows = payload.rows || getLatestPreview_().rows || [];
  if (!rows.length) {
    throw new Error('No preview rows available to save.');
  }

  // Rows restored from _Preview can carry a spreadsheet Date object here;
  // comparing that to normalized keys never matched, so the old rows for this
  // date were kept and the save appended a duplicate set.
  const date = normalizeDateKey_(rows[0].Date || payload.date);
  const day = String(rows[0].Day || payload.day || guessDayCodeFromDate_(date) || '').trim();
  if (!date || !day) throw new Error('The plan does not have a valid date.');
  if (rows.some(row => normalizeDateKey_(row.Date) !== date)) {
    throw new Error('A coverage plan can only be saved for one date at a time.');
  }

  validateCoveragePlanForSave_(date, day, rows);

  const existing = readSheetObjects_('Coverage Output').filter(row => !(normalizeDateKey_(row.Date) === date && String(row.Day || '').trim() === day));
  rewriteSheetRows_('Coverage Output', SHEET_SCHEMAS['Coverage Output'].headers, existing.concat(rows));

  return {
    date: date,
    day: day,
    savedRows: rows.length
  };
}

function getLatestPreview_(date, day) {
  let rows = readSheetObjects_('_Preview');
  if (date) rows = rows.filter(r => normalizeDateKey_(r.Date) === date);
  if (day) rows = rows.filter(r => String(r.Day || '').trim() === String(day).trim());
  return {
    rows: rows,
    summary: {
      totalBlocks: rows.length,
      assignedBlocks: rows.filter(r => String(r.Status || '').trim() === 'Assigned').length,
      unfilledBlocks: rows.filter(r => String(r.Status || '').trim() !== 'Assigned').length
    }
  };
}

function writePreview_(rows) {
  withCoverageLock_(() => {
    rewriteSheetRows_('_Preview', SHEET_SCHEMAS['_Preview'].headers, rows);
  });
}

// The part of the day an absence removes someone, in minutes. Full-day and
// malformed partial absences cover the whole day (fail closed: a bad time
// entry means cover everything rather than silently cover nothing).
function absenceWindowMinutes_(absence) {
  const type = String(absence.absenceType || 'Full Day').trim();
  if (type === 'Full Day') return { startMinutes: 0, endMinutes: 1440 };
  const startMinutes = displayTimeToMinutes_(absence.startOverride);
  const endMinutes = displayTimeToMinutes_(absence.endOverride);
  if (startMinutes == null || endMinutes == null || endMinutes <= startMinutes) {
    return { startMinutes: 0, endMinutes: 1440 };
  }
  return { startMinutes: startMinutes, endMinutes: endMinutes };
}

// Builds each absent person's coverage needs from the union of all of their
// absence windows (ordinary absences and field-trip participation), so
// overlapping or adjacent windows never produce overlapping needs for the same
// class. Each scheduled block is cut into non-overlapping segments, and
// adjacent segments are merged whenever they share the same field trip,
// emergency status, and preferred coverage person.
function buildCoverageNeedsByTeacher_(absences, teacherSchedule, day, fieldTrips) {
  const needs = {};
  const tripsById = {};
  (fieldTrips || []).forEach(trip => {
    if (trip && trip.eventId) tripsById[String(trip.eventId)] = trip;
  });

  const windowsByStaff = {};
  (absences || []).forEach(absence => {
    const name = String(absence.staffName || '').trim();
    if (!name) return;
    const trip = absence.fieldTripEventId
      ? tripsById[String(absence.fieldTripEventId)] || null
      : null;
    (windowsByStaff[name] || (windowsByStaff[name] = [])).push({
      window: absenceWindowMinutes_(absence),
      trip: trip,
      emergency: !!absence.emergency,
      // Preferred coverage is a per-absence preference for ordinary blocks.
      preferred: trip ? '' : String(absence.preferredCoverage || '').trim()
    });
  });

  Object.keys(windowsByStaff).forEach(name => {
    const windows = windowsByStaff[name];
    const rows = normalizedScheduleRowsFor_(teacherSchedule, name, day)
      .filter(row =>
        row.needsCoverageIfAbsent &&
        row.startMinutes != null &&
        row.endMinutes != null &&
        row.endMinutes > row.startMinutes &&
        // Any class whose students are away on a field trip is cancelled and
        // never becomes a coverage need, whether the teacher is on the trip or
        // absent for another reason.
        !blockIsCancelledByFieldTrip_(row, fieldTrips)
      )
      .slice()
      .sort((a, b) => a.startMinutes - b.startMinutes || a.endMinutes - b.endMinutes);

    rows.forEach(row => {
      const pieces = windows
        .map(item => ({
          startMinutes: Math.max(row.startMinutes, item.window.startMinutes),
          endMinutes: Math.min(row.endMinutes, item.window.endMinutes),
          trip: item.trip,
          emergency: item.emergency,
          preferred: item.preferred
        }))
        .filter(piece => piece.endMinutes > piece.startMinutes);
      if (!pieces.length) return;

      const points = Array.from(new Set(
        pieces.reduce((all, piece) => all.concat([piece.startMinutes, piece.endMinutes]), [])
      )).sort((a, b) => a - b);

      const segments = [];
      for (let i = 0; i < points.length - 1; i++) {
        const segStart = points[i];
        const segEnd = points[i + 1];
        const covering = pieces.filter(piece => piece.startMinutes <= segStart && piece.endMinutes >= segEnd);
        if (!covering.length) continue;

        const tripPiece = covering.find(piece => piece.trip);
        const preferredNames = Array.from(new Set(
          covering.map(piece => piece.preferred).filter(Boolean)
        ));
        const label = {
          trip: tripPiece ? tripPiece.trip : null,
          emergency: covering.some(piece => piece.emergency),
          // Two windows asking for different people over the same minutes is
          // ambiguous, so neither preference applies to that stretch.
          preferred: tripPiece || preferredNames.length !== 1 ? '' : preferredNames[0]
        };

        const last = segments[segments.length - 1];
        if (
          last &&
          last.endMinutes === segStart &&
          last.trip === label.trip &&
          last.emergency === label.emergency &&
          last.preferred === label.preferred
        ) {
          last.endMinutes = segEnd;
        } else {
          segments.push(Object.assign({ startMinutes: segStart, endMinutes: segEnd }, label));
        }
      }

      segments.forEach(segment => {
        const need = Object.assign({}, row, {
          startMinutes: segment.startMinutes,
          endMinutes: segment.endMinutes,
          emergencyOverride: segment.emergency,
          preferredCoverage: segment.preferred
        });
        if (segment.startMinutes !== row.startMinutes || segment.endMinutes !== row.endMinutes) {
          need.originalStartMinutes = row.startMinutes;
          need.originalEndMinutes = row.endMinutes;
        }
        if (segment.trip) {
          need.fieldTripEventId = segment.trip.eventId;
          need.fieldTripName = segment.trip.name;
          need.fieldTripGrades = (segment.trip.grades || []).slice();
        }
        (needs[name] || (needs[name] = [])).push(need);
      });
    });
  });

  return needs;
}

// Daily Absences has no dedicated emergency column, so emergency status is
// stored as a marker at the start of Notes. It used to be detected by the word
// "emergency" appearing anywhere, which meant an ordinary absence noted as
// "family emergency" silently switched on emergency rules (restriction
// bypass and 7th/8th-grade pulls), and saving any absence replaced an
// emergency absence's notes with the marker text. The marker is now
// deliberate and the user's notes are kept after it. Rows written by earlier
// versions contain exactly "Emergency coverage" and are still recognized.
const EMERGENCY_NOTE_MARKER_ = '[Emergency]';
const LEGACY_EMERGENCY_NOTE_ = 'emergency coverage';

function isEmergencyAbsence_(absence) {
  const notes = String(absence.Notes || absence.notes || '').trim().toLowerCase();
  return notes.indexOf(EMERGENCY_NOTE_MARKER_.toLowerCase()) === 0 || notes === LEGACY_EMERGENCY_NOTE_;
}

function stripEmergencyNoteMarker_(notes) {
  const text = String(notes || '').trim();
  if (text.toLowerCase() === LEGACY_EMERGENCY_NOTE_) return '';
  if (text.toLowerCase().indexOf(EMERGENCY_NOTE_MARKER_.toLowerCase()) !== 0) return text;
  return text.slice(EMERGENCY_NOTE_MARKER_.length).trim();
}

function composeAbsenceNotes_(emergency, notes) {
  const clean = stripEmergencyNoteMarker_(notes);
  if (!emergency) return clean;
  return clean ? EMERGENCY_NOTE_MARKER_ + ' ' + clean : EMERGENCY_NOTE_MARKER_;
}

function estimateDifficulty_(absentName, blocks, coverageStaff, teacherSchedule, day, config) {
  const fullDayCandidateCount = coverageStaff.filter(candidate => candidateCanCoverAllBlocks_(candidate, absentName, blocks, teacherSchedule, day, makeEmptyState_(), config)).length;
  return {
    fullDayCandidateCount: fullDayCandidateCount,
    blocks: blocks.length
  };
}

function pickBestWholeDayCandidate_(absentName, blocks, coverageStaff, teacherSchedule, day, state, config) {
  const candidates = coverageStaff
    .filter(candidate => candidateCanCoverAllBlocks_(candidate, absentName, blocks, teacherSchedule, day, state, config))
    .map(candidate => {
      const usesEmergencyPull = blocks.some(block =>
        candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state).emergencyPull
      );
      return {
        name: candidate.name,
        tier: candidate.tier,
        role: candidate.role,
        canCoverAllDay: candidate.canCoverAllDay,
        emergencyPull: usesEmergencyPull,
        score: scoreCandidateForWholeDay_(candidate, absentName, blocks, state) - (usesEmergencyPull ? 55 : 0),
        reason: usesEmergencyPull ? 'Emergency pull from 7th/8th.' : 'Best whole-day fit.'
      };
    })
    .sort((a, b) => b.score - a.score || a.tier - b.tier || a.name.localeCompare(b.name));

  return candidates[0] || null;
}


function pickPrimarySplitCandidate_(absentName, blocks, coverageStaff, teacherSchedule, day, state, config) {
  const eventIds = Array.from(new Set(blocks.map(block => block.fieldTripEventId).filter(Boolean)));
  let sourceCandidates = coverageStaff;

  if (eventIds.length === 1) {
    const eventPool = coverageStaff.filter(candidate => candidateHasFieldTripEvent_(candidate, eventIds[0]));
    const eventPoolCanCover = eventPool.some(candidate =>
      blocks.some(block => candidateCanCoverBlock_(candidate, absentName, block, teacherSchedule, day, state, config))
    );
    if (eventPoolCanCover) sourceCandidates = eventPool;
  }

  let candidates = sourceCandidates
    .map(candidate => {
      const coverableBlocks = blocks.filter(block =>
        candidateCanCoverBlock_(candidate, absentName, block, teacherSchedule, day, state, config)
      );

      if (!coverableBlocks.length) return null;

      const availabilities = coverableBlocks.map(block =>
        candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state)
      );
      const usesEmergencyPull = availabilities.some(info => info.emergencyPull);
      const bestFieldTrip = availabilities
        .filter(info => info.fieldTripPriority)
        .sort((a, b) => b.fieldTripPriority - a.fieldTripPriority)[0];

      let score = coverableBlocks.length * 100;
      score += tierBaseScore_(candidate.tier);
      if (candidate.canCoverAllDay) score += 30;
      if (usesEmergencyPull) score -= 55;
      if (bestFieldTrip) score += bestFieldTrip.fieldTripPriority * 250;
      score -= (state.blocksByCandidate[candidate.name] || 0) * 3;
      score -= Object.keys(state.teachersByCandidate[candidate.name] || {}).length * 10;

      const fieldTripReason = bestFieldTrip ? bestFieldTrip.fieldTripReason : '';
      return {
        name: candidate.name,
        tier: candidate.tier,
        role: candidate.role,
        canCoverAllDay: candidate.canCoverAllDay,
        activeToday: candidate.activeToday,
        allowedGrades: candidate.allowedGrades,
        allowedSubjects: candidate.allowedSubjects,
        allowedAssignmentTypes: candidate.allowedAssignmentTypes,
        maxBlocksPerDay: candidate.maxBlocksPerDay,
        maxTeachersPerDay: candidate.maxTeachersPerDay,
        canBeSplitAcrossTeachers: candidate.canBeSplitAcrossTeachers,
        selectedStart: candidate.selectedStart,
        selectedEnd: candidate.selectedEnd,
        notes: candidate.notes,
        fieldTripEvents: candidate.fieldTripEvents || [],
        fieldTripOnly: !!candidate.fieldTripOnly,
        fieldTripReason: fieldTripReason,
        // Mirrors pickBestBlockCandidate_: without this, a break-time move
        // implied by bestFieldTrip would be scored and reported in the
        // reason text but never reserved in
        // state.breakReservationsByCandidate, letting a later assignment
        // double-book the same replacement break slot. Note this batch can
        // cover several blocks at once; if more than one of them implies a
        // distinct break move, only the highest-priority one (bestFieldTrip)
        // is reserved here.
        fieldTripBreakMove: bestFieldTrip ? bestFieldTrip.fieldTripBreakMove || null : null,
        emergencyPull: usesEmergencyPull,
        coverableCount: coverableBlocks.length,
        score: score,
        reason: (fieldTripReason ? fieldTripReason + '; ' : '') +
          (usesEmergencyPull ? 'Emergency pull from 7th/8th; ' : '') +
          'Primary split candidate for ' + coverableBlocks.length + '/' + blocks.length + ' blocks'
      };
    })
    .filter(Boolean);

  if (candidates.some(candidate => !candidate.emergencyPull)) {
    candidates = candidates.filter(candidate => !candidate.emergencyPull);
  }

  candidates = candidates.sort((a, b) =>
    b.coverableCount - a.coverableCount ||
    b.score - a.score ||
    a.tier - b.tier ||
    a.name.localeCompare(b.name)
  );

  return candidates[0] || null;
}

function pickBestBlockCandidate_(absentName, block, coverageStaff, teacherSchedule, day, state, config) {
  let eligible = coverageStaff
    .filter(candidate => candidateCanCoverBlock_(candidate, absentName, block, teacherSchedule, day, state, config));

  if (block.fieldTripEventId) {
    const fieldTripPool = eligible.filter(candidate => candidateHasFieldTripEvent_(candidate, block.fieldTripEventId));
    if (fieldTripPool.length) eligible = fieldTripPool;
  }

  const candidates = eligible
    .map(candidate => {
      const availability = candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state);
      const scoredCandidate = Object.assign({}, candidate, {
        emergencyPull: availability.emergencyPull
      });
      const scoreInfo = scoreCandidateForBlock_(scoredCandidate, absentName, block, state);
      const fieldTripBoost = Number(availability.fieldTripPriority || 0) * 250;
      const runwayMinutes = Math.max(0, Number(availability.fieldTripRunwayMinutes || 0));
      const runwayBoost = Math.min(runwayMinutes, 60);
      const fieldTripReason = availability.fieldTripReason || '';
      const runwayReason = runwayMinutes > 0
        ? runwayMinutes + ' min of released trip-grade time continues after this block'
        : '';
      return {
        name: candidate.name,
        tier: candidate.tier,
        role: candidate.role,
        canCoverAllDay: candidate.canCoverAllDay,
        activeToday: candidate.activeToday,
        allowedGrades: candidate.allowedGrades,
        allowedSubjects: candidate.allowedSubjects,
        allowedAssignmentTypes: candidate.allowedAssignmentTypes,
        maxBlocksPerDay: candidate.maxBlocksPerDay,
        maxTeachersPerDay: candidate.maxTeachersPerDay,
        canBeSplitAcrossTeachers: candidate.canBeSplitAcrossTeachers,
        selectedStart: candidate.selectedStart,
        selectedEnd: candidate.selectedEnd,
        fieldTripEvents: candidate.fieldTripEvents || [],
        fieldTripOnly: !!candidate.fieldTripOnly,
        emergencyPull: availability.emergencyPull,
        fieldTripReason: fieldTripReason,
        fieldTripBreakMove: availability.fieldTripBreakMove || null,
        fieldTripRunwayMinutes: runwayMinutes,
        score: scoreInfo.score + fieldTripBoost + runwayBoost + poolPriorityAdjustment_(candidate.poolPriorityAdjustment),
        reason:
          (fieldTripReason ? fieldTripReason + '; ' : '') +
          (runwayReason ? runwayReason + '; ' : '') +
          scoreInfo.reason
      };
    })
    .sort((a, b) => b.score - a.score || a.tier - b.tier || a.name.localeCompare(b.name));

  return candidates[0] || null;
}

function blocksOverlapEachOther_(blocks) {
  const sorted = (blocks || []).slice().sort((a, b) => a.startMinutes - b.startMinutes);
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i].startMinutes < sorted[i - 1].endMinutes) return true;
  }
  return false;
}

function candidateCanCoverAllBlocks_(candidate, absentName, blocks, teacherSchedule, day, state, config) {
  if (!candidate || !candidate.name) return false;
  if (!candidate.activeToday) return false;
  // Each block is otherwise checked only against assignments already made,
  // so one person could be handed two of these blocks at the same time.
  if (blocksOverlapEachOther_(blocks)) return false;
  if (!candidateCanTakeTeacher_(candidate, absentName, blocks.length, state, config)) return false;
  return blocks.every(block => candidateCanCoverBlock_(candidate, absentName, block, teacherSchedule, day, state, config));
}

function candidateCanCoverBlock_(candidate, absentName, block, teacherSchedule, day, state, config) {
  if (!candidate || !candidate.name) return false;
  if (!candidate.activeToday) return false;
  if (candidate.fieldTripOnly && !block.fieldTripEventId) return false;
  if (candidateIsAbsentForBlock_(candidate.name, block, state)) return false;
  if (candidateHasReservedBreakConflict_(candidate.name, block, state)) return false;
  if (!candidateCanTakeTeacher_(candidate, absentName, 1, state, config)) return false;
  if (!block.emergencyOverride) {
    if (!matchesAllowedValue_(candidate.allowedAssignmentTypes, block.assignmentType)) return false;
    if (!matchesAllowedValue_(candidate.allowedSubjects, block.subject)) return false;
    if (!matchesAllowedGrade_(candidate.allowedGrades, block.grade)) return false;
  }
  if (!candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state).available) return false;
  if (candidateHasTimeConflict_(candidate.name, block, state)) return false;
  if ((state.blocksByCandidate[candidate.name] || 0) + 1 > candidate.maxBlocksPerDay) return false;
  return true;
}

function candidateCanTakeTeacher_(candidate, absentName, additionalBlocks, state, config) {
  const currentTeachers = state.teachersByCandidate[candidate.name] || {};
  const alreadyHasTeacher = !!currentTeachers[absentName];
  const distinctTeachers = Object.keys(currentTeachers).length + (alreadyHasTeacher ? 0 : 1);
  if (distinctTeachers > candidate.maxTeachersPerDay) return false;
  if (!candidate.canBeSplitAcrossTeachers && !alreadyHasTeacher && Object.keys(currentTeachers).length > 0) return false;
  if ((state.blocksByCandidate[candidate.name] || 0) + additionalBlocks > candidate.maxBlocksPerDay) return false;
  return true;
}

function candidateHasAvailabilityForBlock_(candidate, block, teacherSchedule, day, state) {
  return candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state).available;
}

// Normalized Teacher Schedule rows, indexed by staff name + day code and
// cached per schedule array. Every server call builds a fresh schedule array,
// so the cache can never serve rows from an earlier request. Consumers treat
// these rows as read-only (they copy before changing anything).
//
// Before this index, candidateAvailabilityForBlock_ re-normalized the entire
// schedule on every call, and it is called for every candidate x block pair
// (several times over). On a 60-teacher schedule that was ~4.4 million row
// normalizations per Generate Plan.
const SCHEDULE_INDEX_CACHE_ = new WeakMap();

function normalizedScheduleRowsFor_(teacherSchedule, staffName, day) {
  const source = teacherSchedule || [];
  let index = SCHEDULE_INDEX_CACHE_.get(source);
  if (!index) {
    index = {};
    source.forEach(raw => {
      const row = normalizeTeacherScheduleRow_(raw);
      if (!row.staffName) return;
      const key = row.staffName + '\u0000' + row.day;
      (index[key] || (index[key] = [])).push(row);
    });
    SCHEDULE_INDEX_CACHE_.set(source, index);
  }
  return index[String(staffName || '') + '\u0000' + String(day || '')] || [];
}

// A scheduled obligation (a class, homeroom, or duty the person must still
// attend) that overlaps the block. Time inside a field trip that cancels this
// person's own trip-grade class is not an obligation. Used so that "can cover
// any block" never overrides where the Teacher Schedule says someone is.
function candidateOwnScheduleConflict_(candidate, block, availabilityRows) {
  return (availabilityRows || []).find(row => {
    if (!row.needsCoverageIfAbsent || row.coverEligibleThisBlock) return false;
    if (row.startMinutes == null || row.endMinutes == null) return false;
    const overlapStart = Math.max(row.startMinutes, block.startMinutes);
    const overlapEnd = Math.min(row.endMinutes, block.endMinutes);
    if (overlapEnd <= overlapStart) return false;

    const releasedByTrip = isInstructionalGradeBlock_(row) &&
      (candidate.fieldTripEvents || []).some(event =>
        (event.grades || []).some(grade => normalizeGradeKey_(grade) === normalizeGradeKey_(row.grade)) &&
        event.startMinutes <= overlapStart &&
        event.endMinutes >= overlapEnd
      );
    return !releasedByTrip;
  }) || null;
}

function candidateAvailabilityForBlock_(candidate, block, teacherSchedule, day, state) {
  if (!candidateAvailabilityWindowAllowsBlock_(candidate, block)) {
    return { available: false, emergencyPull: false, fieldTripPriority: 0, fieldTripReason: '', fieldTripBreakMove: null, fieldTripRunwayMinutes: 0 };
  }

  const availabilityRows = normalizedScheduleRowsFor_(teacherSchedule, candidate.name, day);

  const fieldTripAvailability = candidateFieldTripAvailability_(candidate, block, availabilityRows, state);
  if (fieldTripAvailability.available) {
    return {
      available: true,
      emergencyPull: false,
      fieldTripPriority: fieldTripAvailability.fieldTripPriority,
      fieldTripReason: fieldTripAvailability.fieldTripReason,
      fieldTripBreakMove: fieldTripAvailability.fieldTripBreakMove || null,
      fieldTripRunwayMinutes: Number(fieldTripAvailability.fieldTripRunwayMinutes || 0)
    };
  }

  if (candidate.fieldTripOnly) {
    return {
      available: false,
      emergencyPull: false,
      fieldTripPriority: 0,
      fieldTripReason: fieldTripAvailability.fieldTripReason || '',
      fieldTripBreakMove: null,
      fieldTripRunwayMinutes: 0
    };
  }

  // Teacher Schedule is authoritative for where people are. "Can cover any
  // block" describes a person's hours, not permission to leave their own class
  // or duty. Emergency pulls (below) deliberately bypass this.
  const ownConflict = candidateOwnScheduleConflict_(candidate, block, availabilityRows);

  if (!ownConflict && candidate.canCoverAllDay) {
    return { available: true, emergencyPull: false, fieldTripPriority: 0, fieldTripReason: '', fieldTripBreakMove: null, fieldTripRunwayMinutes: 0 };
  }

  if (!ownConflict && availabilityRows.some(row =>
    row.coverEligibleThisBlock &&
    row.startMinutes <= block.startMinutes &&
    row.endMinutes >= block.endMinutes &&
    (
      !assignmentTypeIsBreak_(row) ||
      !block.fieldTripEventId ||
      !candidateHasFieldTripEvent_(candidate, block.fieldTripEventId)
    )
  )) {
    return { available: true, emergencyPull: false, fieldTripPriority: 0, fieldTripReason: '', fieldTripBreakMove: null, fieldTripRunwayMinutes: 0 };
  }

  if (!block.emergencyOverride || Number(candidate.tier) !== 3) {
    return { available: false, emergencyPull: false, fieldTripPriority: 0, fieldTripReason: '', fieldTripBreakMove: null, fieldTripRunwayMinutes: 0 };
  }

  const emergencyRow = availabilityRows.find(row =>
    isSeventhOrEighthGradeBlock_(row) &&
    row.startMinutes < block.endMinutes &&
    row.endMinutes > block.startMinutes
  );

  return {
    available: !!emergencyRow,
    emergencyPull: !!emergencyRow,
    fieldTripPriority: 0,
    fieldTripReason: '',
    fieldTripBreakMove: null,
    fieldTripRunwayMinutes: 0
  };
}

function candidateHasTimeConflict_(candidateName, block, state) {
  const assignments = state.assignmentsByCandidate[candidateName] || [];
  return assignments.some(existing => timesOverlap_(existing.startMinutes, existing.endMinutes, block.startMinutes, block.endMinutes));
}

function scoreCandidateForWholeDay_(candidate, absentName, blocks, state) {
  let score = tierBaseScore_(candidate.tier) + 100;
  if (candidate.canCoverAllDay) score += 30;
  score -= (state.blocksByCandidate[candidate.name] || 0) * 3;
  score -= Object.keys(state.teachersByCandidate[candidate.name] || {}).length * 8;
  return score;
}

function scoreCandidateForBlock_(candidate, absentName, block, state) {
  let score = tierBaseScore_(candidate.tier);
  const reasons = ['Tier ' + candidate.tier];
  if (candidate.emergencyPull) {
    score -= 55;
    reasons.push('emergency pull from 7th/8th');
  }
  if (candidate.canCoverAllDay) {
    score += 20;
    reasons.push('all-day available');
  }

  const teacherAssignments = state.teachersByCandidate[candidate.name] || {};
  const blocksByCandidate = state.assignmentsByCandidate[candidate.name] || [];
  const sameEventAssignments = block.fieldTripEventId
    ? blocksByCandidate.filter(existing => existing.fieldTripEventId === block.fieldTripEventId)
    : [];

  if (block.fieldTripEventId && sameEventAssignments.length) {
    score += 55;
    reasons.push('already helping this field trip');
  }

  if (teacherAssignments[absentName]) {
    score += block.fieldTripEventId ? 35 : 80;
    reasons.push('already covering same teacher');
  }

  const adjacentSameEvent = block.fieldTripEventId && sameEventAssignments.some(existing =>
    existing.endMinutes === block.startMinutes ||
    existing.startMinutes === block.endMinutes
  );
  const adjacentSameTeacher = blocksByCandidate.some(existing =>
    existing.absentName === absentName &&
    (existing.endMinutes === block.startMinutes || existing.startMinutes === block.endMinutes)
  );

  if (adjacentSameEvent) {
    score += 120;
    reasons.push('keeps field-trip coverage continuous');
  } else if (adjacentSameTeacher) {
    score += 30;
    reasons.push('keeps continuity');
  }

  if (!teacherAssignments[absentName] && Object.keys(teacherAssignments).length > 0) {
    if (!(block.fieldTripEventId && sameEventAssignments.length)) {
      score -= 40;
      reasons.push('new teacher handoff');
    }
  }

  if (!matchesAllowedValue_(candidate.allowedSubjects, block.subject)) {
    score -= 30;
  } else if (candidate.allowedSubjects !== '*') {
    score += 8;
    reasons.push('subject fit');
  }

  if (!matchesAllowedGrade_(candidate.allowedGrades, block.grade)) {
    score -= 30;
  } else if (candidate.allowedGrades !== '*') {
    score += 6;
    reasons.push('grade fit');
  }

  score -= (state.blocksByCandidate[candidate.name] || 0) * 4;
  score -= Object.keys(teacherAssignments).length * 8;

  return {
    score: score,
    reason: reasons.join(', ')
  };
}

function tierBaseScore_(tier) {
  if (Number(tier) === 1) return 100;
  if (Number(tier) === 2) return 60;
  return 20;
}

function isSeventhOrEighthGradeBlock_(row) {
  const grade = extractGradeNumber_(row.grade || row.className);
  if (grade !== 7 && grade !== 8) return false;
  const type = String(row.assignmentType || '').trim().toLowerCase();
  return type === 'class' || type === 'homeroom' || type === '';
}

function recordAssignment_(state, candidate, block, absentName) {
  // Shifts of existing moved breaks planned by the slot search for this block.
  const move = candidate.fieldTripBreakMove;
  (move && move.relocations || []).forEach(item => {
    if (candidateBreakReservations_(candidate.name, state).indexOf(item.reservation) === -1) return;
    item.reservation.startMinutes = item.startMinutes;
    item.reservation.endMinutes = item.endMinutes;
  });

  if (!state.assignmentsByCandidate[candidate.name]) state.assignmentsByCandidate[candidate.name] = [];
  state.assignmentsByCandidate[candidate.name].push({
    absentName: absentName,
    startMinutes: block.startMinutes,
    endMinutes: block.endMinutes,
    fieldTripEventId: block.fieldTripEventId || ''
  });

  if (candidate.fieldTripBreakMove) {
    if (!state.breakReservationsByCandidate[candidate.name]) {
      state.breakReservationsByCandidate[candidate.name] = [];
    }
    const move = candidate.fieldTripBreakMove;
    const duplicate = state.breakReservationsByCandidate[candidate.name].some(existing =>
      existing.eventId === move.eventId &&
      existing.startMinutes === move.replacementStartMinutes &&
      existing.endMinutes === move.replacementEndMinutes
    );
    if (!duplicate) {
      state.breakReservationsByCandidate[candidate.name].push({
        eventId: move.eventId,
        startMinutes: move.replacementStartMinutes,
        endMinutes: move.replacementEndMinutes,
        originalBreakStartMinutes: move.originalBreakStartMinutes,
        originalBreakEndMinutes: move.originalBreakEndMinutes,
        eligibleRows: move.eligibleRows || [],
        reason: move.reason
      });
    }
  }

  // If this block sits on a moved break, shift that break to another open
  // cancelled slot (the conflict check already confirmed one exists).
  candidateBreakReservations_(candidate.name, state).forEach(reservation => {
    if (!timesOverlap_(reservation.startMinutes, reservation.endMinutes, block.startMinutes, block.endMinutes)) return;
    const slot = relocateBreakReservation_(candidate.name, reservation, [block], state);
    if (slot) {
      reservation.startMinutes = slot.startMinutes;
      reservation.endMinutes = slot.endMinutes;
    }
  });

  state.blocksByCandidate[candidate.name] = (state.blocksByCandidate[candidate.name] || 0) + 1;
  if (!state.teachersByCandidate[candidate.name]) state.teachersByCandidate[candidate.name] = {};
  state.teachersByCandidate[candidate.name][absentName] = true;
}

function makeEmptyState_() {
  return {
    assignmentsByCandidate: {},
    blocksByCandidate: {},
    teachersByCandidate: {},
    absencesByCandidate: {},
    breakReservationsByCandidate: {}
  };
}

function buildAbsenceWindowsByStaff_(absences) {
  const map = {};

  (absences || []).forEach(absence => {
    const name = String(absence.staffName || '').trim();
    if (!name) return;
    if (!map[name]) map[name] = [];

    const type = String(absence.absenceType || 'Full Day').trim();
    if (type === 'Full Day') {
      map[name].push({
        allDay: true,
        startMinutes: null,
        endMinutes: null,
        source: String(absence.fieldTripEventId || absence.notes || 'Absence')
      });
      return;
    }

    map[name].push({
      allDay: false,
      startMinutes: displayTimeToMinutes_(absence.startOverride),
      endMinutes: displayTimeToMinutes_(absence.endOverride),
      source: String(absence.fieldTripEventId || absence.notes || 'Absence')
    });
  });

  return map;
}

function absenceWindowsForCandidate_(candidateName, state) {
  const raw = state && state.absencesByCandidate
    ? state.absencesByCandidate[String(candidateName || '').trim()]
    : null;
  if (!raw) return [];
  return Array.isArray(raw) ? raw : [raw];
}

function candidateIsAbsentForBlock_(candidateName, block, state) {
  const windows = absenceWindowsForCandidate_(candidateName, state);
  if (!windows.length) return false;

  return windows.some(absence => {
    if (absence.allDay) return true;

    // If a partial absence is malformed, fail closed rather than scheduling
    // an absent person as coverage.
    if (absence.startMinutes == null || absence.endMinutes == null) return true;

    return timesOverlap_(
      absence.startMinutes,
      absence.endMinutes,
      block.startMinutes,
      block.endMinutes
    );
  });
}

function makePlanRow_(date, day, block, candidate, coverageMode, status, notes) {
  return {
    Date: date,
    Day: day,
    Event_ID: block.fieldTripEventId || '',
    Start: minutesToDisplay_(block.startMinutes),
    End: minutesToDisplay_(block.endMinutes),
    Absent_Staff: block.staffName,
    Class: block.className,
    Grade: block.grade,
    Subject: block.subject,
    Assignment_Type: block.assignmentType,
    Room: block.room,
    Assigned_Coverage: candidate ? candidate.name : '',
    Coverage_Mode: coverageMode,
    Coverage_Tier_Used: candidate ? String(candidate.tier || '') : '',
    Status: status,
    Notes: notes || ''
  };
}

function normalizeScheduleTerm_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const compact = raw.toLowerCase().replace(/[.\s_-]+/g, '');

  // Full-year rows are active in both semesters. Normalize them to blank so
  // date-based semester filtering preserves them alongside the active term.
  if (
    compact === 'allyear' ||
    compact === 'fullyear' ||
    compact === 'yearlong' ||
    compact === 'annual'
  ) return '';

  if (compact === 's1' || compact === 'semester1' || compact === 'term1' || compact === 'firstsemester') return 'S1';
  if (compact === 's2' || compact === 'semester2' || compact === 'term2' || compact === 'secondsemester') return 'S2';
  return raw;
}

function activeScheduleTermForDate_(date, config) {
  const override = normalizeScheduleTerm_(
    config && (config.Schedule_Term_Override || config.Active_Term || config.Schedule_Term)
  );
  if (override) return override;

  const key = normalizeDateKey_(date);
  const parsed = key ? new Date(key + 'T12:00:00') : null;
  if (!parsed || isNaN(parsed)) return '';

  // Coverage Scheduler follows the school's academic-year pattern:
  // S1 is the fall semester; S2 is the spring semester.
  return parsed.getMonth() >= 6 ? 'S1' : 'S2';
}

function filterTeacherScheduleForDate_(rows, date, config) {
  const source = rows || [];
  const termsPresent = new Set(
    source.map(row => normalizeScheduleTerm_(row.Term)).filter(Boolean)
  );

  // Preserve backward compatibility for schedules with no semester column or
  // schedules using unrelated custom term labels.
  if (!termsPresent.has('S1') && !termsPresent.has('S2')) return source.slice();

  const activeTerm = activeScheduleTermForDate_(date, config);
  if (!activeTerm) return source.slice();

  return source.filter(row => {
    const term = normalizeScheduleTerm_(row.Term);
    return !term || term === activeTerm;
  });
}

function normalizeTeacherScheduleRow_(row) {
  const subject = String(row.Subject || '').trim();
  const rawClassName = String(row.Class || '').trim();
  const assignmentType = String(row.Assignment_Type || '').trim() || inferAssignmentType_(row);
  const className = buildClassDisplayName_(rawClassName, subject, assignmentType);

  const explicitNeeds = String(row.Needs_Coverage_If_Absent || '').trim();
  const explicitCover = String(row.Cover_Eligible_This_Block || '').trim();

  const normalized = {
    staffName: String(row.Staff_Name || '').trim(),
    role: String(row.Role || '').trim(),
    term: normalizeScheduleTerm_(row.Term),
    day: String(row.Day || '').trim(),
    startMinutes: timeToMinutes_(row.Start),
    endMinutes: timeToMinutes_(row.End),
    className: className,
    // Only the Class column (or an explicit Grade) identifies students; the
    // Subject text never does.
    grade: String(row.Grade || '').trim() || inferGradeFromClass_(rawClassName),
    subject: subject,
    assignmentType: assignmentType,
    room: String(row.Room || '').trim(),
    needsCoverageIfAbsent: explicitNeeds !== ''
      ? normalizeYesNo_(explicitNeeds, false)
      : inferNeedsCoverageIfAbsent_(assignmentType, subject),
    coverEligibleThisBlock: explicitCover !== ''
      ? normalizeYesNo_(explicitCover, false)
      : inferCoverEligibleThisBlock_(assignmentType, subject)
  };

  // A row with no Class (e.g. "(Art Thursdays)") still has students: Class
  // Schedule says which section this teacher has then. With that grade, a
  // field trip for the grade cancels the row and releases the teacher, just
  // like a row whose Class names the section.
  if (!normalized.grade && normalized.needsCoverageIfAbsent) {
    const match = classScheduleGradeFor_(normalized);
    if (match && match.grade) {
      normalized.grade = match.grade;
      normalized.gradeSource = 'Class Schedule';
      normalized.sections = match.sections;
      if (!rawClassName) {
        normalized.className = buildClassDisplayName_(match.sections.join('/'), subject, assignmentType);
      }
    }
  }

  return normalized;
}

function buildClassDisplayName_(className, subject, assignmentType) {
  const rawClass = String(className || '').trim();
  const rawSubject = String(subject || '').trim();
  const type = String(assignmentType || '').trim();

  // Keep non-teaching blocks readable without pretending they are classes.
  if (!rawClass && !rawSubject) return type;
  if (!rawClass) return rawSubject;
  if (!rawSubject) return rawClass;

  const classLower = rawClass.toLowerCase();
  const subjectLower = rawSubject.toLowerCase();
  if (classLower === subjectLower || classLower.indexOf(subjectLower) !== -1 || subjectLower.indexOf(classLower) !== -1) {
    return rawClass;
  }
  return rawClass + ' — ' + rawSubject;
}

// Reads a grade only from a Class label that names a student group: a
// section or grade label at the start ("7D", "6B — Homeroom", "7th Grade",
// "Grade 3", "KA", "PreK A", "Beg B"). Anything else has no grade.
//
// This used to pull the first number out of any text, including the Subject
// when Class was blank. On the real schedule that turned "8:30-9:15" into
// grade 8, "Office 4" into grade 4 and "(Art in Rm. 24)" into grade 24, and
// every one of those rows would be cancelled as a class whenever that grade
// went on a field trip. Course names like "Algebra 1" or "Spanish 2" would
// have been read as grades 1 and 2 the same way. An unknown grade is the safe
// failure: the row still gets coverage instead of being silently cancelled.
function inferGradeFromClass_(className) {
  const raw = String(className || '').trim();
  if (!raw) return '';
  // Plan rows use "Class — Subject" display names; only the Class part counts.
  const label = raw.split(/\s+[—–-]\s+/)[0].trim();

  if (/^beg(inners?)?(\s+[a-z])?$/i.test(label)) return 'Beg';
  if (/^pre[\s-]?k(indergarten)?(\s*[a-z])?$/i.test(label)) return 'PreK';
  if (/^k[a-z]?$/i.test(label) || /^kindergarten(\s+[a-z])?$/i.test(label) || /^k\s+[a-z]$/i.test(label)) return 'K';

  const leading = label.match(/^(\d{1,2})(?:st|nd|rd|th)?(?:\s*grade)?\s*[a-z]?(?=\s|$)/i);
  if (leading) {
    const n = Number(leading[1]);
    return n >= 1 && n <= 12 ? String(n) : '';
  }
  const named = label.match(/^grade\s*(\d{1,2})(?:\s*[a-z])?(?=\s|$)/i);
  if (named) {
    const n = Number(named[1]);
    return n >= 1 && n <= 12 ? String(n) : '';
  }
  return '';
}

function inferNeedsCoverageIfAbsent_(assignmentType, subject) {
  const type = String(assignmentType || '').trim().toLowerCase();
  const s = String(subject || '').trim().toLowerCase();

  if (type === 'break' || type === 'lunch' || type === 'planning') return false;
  if (type === 'class' || type === 'homeroom' || type === 'duty') return true;
  if (s.indexOf('break') !== -1 || s.indexOf('lunch') !== -1 || s.indexOf('plan') !== -1) return false;
  return true;
}

function inferCoverEligibleThisBlock_(assignmentType, subject) {
  const type = String(assignmentType || '').trim().toLowerCase();
  const s = String(subject || '').trim().toLowerCase();

  if (type === 'class' || type === 'homeroom' || type === 'duty') return false;
  if (type === 'lunch') return lunchCoverageAllowed_();
  if (type === 'planning' || type === 'break') return true;
  if (s.indexOf('plan') !== -1 || s.indexOf('break') !== -1) return true;
  if (s.indexOf('lunch') !== -1) return lunchCoverageAllowed_();
  return false;
}

function normalizeCoverageStaffRow_(row, config) {
  const rawMaxBlocks = String(row.Max_Blocks_Per_Day || '').trim();
  const rawMaxTeachers = String(row.Max_Teachers_Per_Day || '').trim();

  return {
    name: String(row.Name || '').trim(),
    role: String(row.Role || '').trim(),
    tier: Number(row.Coverage_Tier || 3),
    canCoverAllDay: normalizeYesNo_(row.Can_Cover_All_Day, false),
    baseActive: normalizeYesNo_(row.Active_Today, true),
    activeToday: normalizeYesNo_(row.Active_Today, true),
    availableDays: normalizeAvailableDays_(row.Available_Days),
    defaultStart: timeToDisplay_(row.Default_Start),
    defaultEnd: timeToDisplay_(row.Default_End),
    selectedStart: timeToDisplay_(row.Default_Start),
    selectedEnd: timeToDisplay_(row.Default_End),
    hasDateOverride: false,
    availabilityNotes: '',
    allowedGrades: normalizeAllowedField_(row.Allowed_Grades),
    allowedSubjects: normalizeAllowedField_(row.Allowed_Subjects),
    allowedAssignmentTypes: normalizeAllowedField_(row.Allowed_Assignment_Types),
    maxBlocksPerDay: resolveDailyLimit_(rawMaxBlocks, config, 'Default_Max_Blocks_Per_Day'),
    maxTeachersPerDay: resolveDailyLimit_(rawMaxTeachers, config, 'Default_Max_Teachers_Per_Day'),
    // What the sheet actually says (blank = use the Config default), for editing.
    rawMaxBlocksPerDay: rawMaxBlocks,
    rawMaxTeachersPerDay: rawMaxTeachers,
    canBeSplitAcrossTeachers: normalizeYesNo_(row.Can_Be_Split_Across_Teachers, true),
    notes: String(row.Notes || '').trim()
  };
}


function getFieldTripCoverageStaffForDate_(date, day) {
  const fieldTrips = getFieldTripsForDate_(date);
  if (!fieldTrips.length) return [];

  const config = getConfigMap_();
  const configuredCoverageStaff = getCoverageStaffForDate_(date, day, config);
  const activeCoverageStaff = configuredCoverageStaff.filter(row => row.name && row.activeToday);
  const poolRows = fieldTripCoveragePoolRowsForDate_(date);
  return buildFieldTripCoverageCandidatesFromPool_(
    fieldTrips,
    date,
    activeCoverageStaff,
    configuredCoverageStaff,
    poolRows
  )
    .filter(candidate => (candidate.fieldTripEvents || []).length)
    .map(candidate => ({
      name: candidate.name,
      role: candidate.role,
      tier: candidate.tier,
      activeToday: candidate.activeToday,
      fieldTripOnly: !!candidate.fieldTripOnly,
      fieldTripEvents: (candidate.fieldTripEvents || []).map(event => ({
        eventId: event.eventId,
        name: event.name
      }))
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function getCoverageStaffForDate_(date, day, config) {
  const rows = readSheetObjects_(getCoverageStaffSheetName_())
    .map(row => normalizeCoverageStaffRow_(row, config));
  return applySubstituteAvailabilityOverrides_(rows, date, day);
}

function applySubstituteAvailabilityOverrides_(coverageStaff, date, day) {
  const dateKey = normalizeDateKey_(date);
  const dayCode = String(day || guessDayCodeFromDate_(dateKey) || '').trim();
  const overrides = availabilityOverridesEnabled_() ? getSubstituteAvailabilityMap_(dateKey, dayCode) : {};

  return coverageStaff.map(candidate => {
    const dayAllowed = isAvailableOnDay_(candidate.availableDays, dayCode);
    let active = !!candidate.baseActive && dayAllowed;
    let selectedStart = candidate.defaultStart || '';
    let selectedEnd = candidate.defaultEnd || '';
    let availabilityNotes = '';
    let hasDateOverride = false;

    const override = overrides[candidate.name];
    if (override) {
      hasDateOverride = true;
      active = normalizeYesNo_(override.Available, active);
      selectedStart = timeToDisplay_(override.Start) || selectedStart;
      selectedEnd = timeToDisplay_(override.End) || selectedEnd;
      availabilityNotes = String(override.Notes || '').trim();
    }

    return Object.assign({}, candidate, {
      activeToday: active,
      selectedStart: selectedStart,
      selectedEnd: selectedEnd,
      hasDateOverride: hasDateOverride,
      availabilityNotes: availabilityNotes
    });
  });
}

function getSubstituteAvailabilityMap_(date, day) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Substitute Availability')) return {};
  const dateKey = normalizeDateKey_(date);
  const dayCode = String(day || '').trim();
  const map = {};
  readSheetObjects_('Substitute Availability')
    .filter(row => normalizeDateKey_(row.Date) === dateKey && (!dayCode || String(row.Day || '').trim() === dayCode))
    .forEach(row => {
      const name = String(row.Name || '').trim();
      if (name) map[name] = row;
    });
  return map;
}

function upsertSubstituteAvailability(payload) {
  return withCoverageLock_(() => upsertSubstituteAvailabilityUnlocked_(payload));
}

function upsertSubstituteAvailabilityUnlocked_(payload) {
  payload = payload || {};
  if (!availabilityOverridesEnabled_()) {
    throw new Error('Per-day availability changes are turned off (Config: Availability_Override_Mode is OFF), so this change would be ignored. Set it to DATE to use the availability switches.');
  }
  const date = normalizeDateKey_(payload.date);
  const day = String(payload.day || guessDayCodeFromDate_(date) || '').trim();
  const name = String(payload.name || '').trim();
  if (!date) throw new Error('No date provided for substitute availability.');
  if (!day) throw new Error('No day code provided for substitute availability.');
  if (!name) throw new Error('No substitute name provided for availability.');
  const sheet = ensureSubstituteAvailabilitySheet_();
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(h => String(h || '').trim());
  const aliasMap = HEADER_ALIASES['Substitute Availability'];
  const dateCol = findColumnByAliases_(headers, aliasMap.Date);
  const nameCol = findColumnByAliases_(headers, aliasMap.Name);
  if (dateCol === -1 || nameCol === -1) throw new Error('Substitute Availability must include Date and Name columns.');
  const values = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let r = 1; r < values.length; r++) {
    const rowDate = normalizeDateKey_(values[r][dateCol]);
    const rowName = String(values[r][nameCol] || '').trim();
    if (rowDate === date && rowName === name) {
      targetRow = r + 1;
      break;
    }
  }
  const rowObject = {
    Date: date,
    Day: day,
    Name: name,
    Available: normalizeYesNo_(payload.available, true) ? 'Yes' : 'No',
    Start: timeToDisplay_(payload.start),
    End: timeToDisplay_(payload.end),
    Notes: String(payload.notes || '').trim()
  };
  if (targetRow === -1) targetRow = sheet.getLastRow() + 1;
  setSheetRowObject_(sheet, targetRow, headers, aliasMap, rowObject);
  markFieldTripCoveragePoolDirty_();
  return getAllCoverageStaff_(date, day);
}

function ensureSubstituteAvailabilitySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Substitute Availability');
  if (!sheet) {
    sheet = ss.insertSheet('Substitute Availability');
    sheet.getRange(1, 1, 1, SHEET_SCHEMAS['Substitute Availability'].headers.length)
      .setValues([SHEET_SCHEMAS['Substitute Availability'].headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setSheetRowObject_(sheet, rowNumber, headers, aliasMap, rowObject) {
  const actualHeaders = (headers || []).map(h => String(h || '').trim());
  const width = Math.max(actualHeaders.length, 1);
  let existing;
  if (rowNumber <= sheet.getLastRow()) {
    const existingRange = sheet.getRange(rowNumber, 1, 1, width);
    existing = existingRange.getValues()[0];
    const formulas = existingRange.getFormulas()[0];
    // Preserve formulas in custom/unmanaged columns when the row is rewritten.
    formulas.forEach((formula, index) => {
      if (formula) existing[index] = formula;
    });
  } else {
    existing = new Array(width).fill('');
  }

  Object.keys(rowObject || {}).forEach(canonical => {
    const col = findColumnByAliases_(actualHeaders, aliasMap[canonical]);
    if (col !== -1) existing[col] = rowObject[canonical];
  });

  sheet.getRange(rowNumber, 1, 1, width).setValues([existing]);
  incrementCoverageMetric_('sheetWrites');
  invalidateCoverageSheetCache_(sheet.getName());
}

function normalizeAvailableDays_(value) {
  const str = String(value || '').trim();
  if (!str || /^all$/i.test(str)) return '*';
  return str;
}

function isAvailableOnDay_(availableDays, day) {
  if (!availableDays || availableDays === '*') return true;
  const dayCode = normalizeDayToken_(day);
  if (!dayCode) return true;
  return String(availableDays || '')
    .split(/[;,/|\s]+/)
    .map(normalizeDayToken_)
    .filter(Boolean)
    .indexOf(dayCode) !== -1;
}

function normalizeDayToken_(value) {
  const str = String(value || '').trim().toUpperCase();
  if (!str) return '';
  if (str === 'MON' || str === 'MONDAY') return 'M';
  if (str === 'TUE' || str === 'TUES' || str === 'TUESDAY') return 'T';
  if (str === 'WED' || str === 'WEDNESDAY') return 'W';
  if (str === 'THU' || str === 'THUR' || str === 'THURS' || str === 'THURSDAY') return 'R';
  if (str === 'FRI' || str === 'FRIDAY') return 'F';
  return str.charAt(0);
}

function candidateAvailabilityWindowAllowsBlock_(candidate, block) {
  const start = timeToMinutes_(candidate.selectedStart);
  const end = timeToMinutes_(candidate.selectedEnd);
  if (start != null && block.startMinutes < start) return false;
  if (end != null && block.endMinutes > end) return false;
  return true;
}

function inferAssignmentType_(row) {
  const subject = String(row.Subject || '').toLowerCase();
  // Duty first: "Lunch Duty" is supervision that needs coverage, not the
  // teacher's own lunch.
  if (subject.indexOf('duty') !== -1) return 'Duty';
  if (subject.indexOf('lunch') !== -1) return 'Lunch';
  if (subject.indexOf('break') !== -1) return 'Break';
  if (subject.indexOf('plan') !== -1) return 'Planning';
  if (subject.indexOf('homeroom') !== -1) return 'Homeroom';
  return 'Class';
}

function normalizeAllowedField_(value) {
  const str = String(value || '').trim();
  if (!str || /^all$/i.test(str)) return '*';
  return str;
}

function matchesAllowedValue_(allowedField, value) {
  if (allowedField === '*') return true;
  const target = normalizeMatchText_(value);
  if (!target) return true;
  return String(allowedField || '')
    .split(',')
    .map(v => normalizeMatchText_(v))
    .filter(Boolean)
    .some(token => matchTextContainsToken_(target, token) || matchTextContainsToken_(token, target));
}

function normalizeMatchText_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function matchTextContainsToken_(target, token) {
  if (!target || !token) return false;
  return target === token || (' ' + target + ' ').indexOf(' ' + token + ' ') !== -1;
}

function matchesAllowedGrade_(allowedField, gradeValue) {
  if (allowedField === '*') return true;
  const targetRaw = String(gradeValue || '').trim();
  if (!targetRaw) return true;
  const target = targetRaw.toLowerCase();
  const targetNumeric = extractGradeNumber_(targetRaw);

  return String(allowedField || '')
    .split(',')
    .map(v => v.trim())
    .filter(Boolean)
    .some(token => {
      const lower = token.toLowerCase();
      if (lower === target) return true;
      if (lower.indexOf('-') !== -1) {
        const parts = lower.split('-').map(s => s.trim());
        const start = gradeToComparable_(parts[0]);
        const end = gradeToComparable_(parts[1]);
        const current = gradeToComparable_(targetRaw);
        return start != null && end != null && current != null && current >= start && current <= end;
      }
      if (targetNumeric != null) {
        const tokenNumeric = extractGradeNumber_(token);
        return tokenNumeric != null && tokenNumeric === targetNumeric;
      }
      return false;
    });
}

function extractGradeNumber_(value) {
  const match = String(value || '').match(/\d+/);
  return match ? Number(match[0]) : null;
}

function gradeToComparable_(value) {
  const str = String(value || '').trim().toUpperCase();
  if (str === 'K') return 0;
  const num = extractGradeNumber_(str);
  return num != null ? num : null;
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


function backfillTeacherScheduleDerivedFields() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher Schedule');
  if (!sheet) throw new Error('Missing Teacher Schedule sheet.');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h || '').trim());
  const required = ['Grade', 'Assignment_Type', 'Needs_Coverage_If_Absent', 'Cover_Eligible_This_Block'];

  required.forEach(h => {
    if (headers.indexOf(h) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      headers.push(h);
    }
  });

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const out = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] || '');
    const normalized = normalizeTeacherScheduleRow_(obj);

    const newRow = new Array(headers.length).fill('');
    headers.forEach((h, i) => newRow[i] = row[i] || '');

    newRow[idx['Grade']] = normalized.grade;
    newRow[idx['Assignment_Type']] = normalized.assignmentType;
    newRow[idx['Needs_Coverage_If_Absent']] = normalized.needsCoverageIfAbsent ? 'Yes' : 'No';
    newRow[idx['Cover_Eligible_This_Block']] = normalized.coverEligibleThisBlock ? 'Yes' : 'No';

    return newRow;
  });

  sheet.getRange(2, 1, out.length, headers.length).setValues(out);
  incrementCoverageMetric_('sheetWrites');
  invalidateCoverageSheetCache_('Teacher Schedule');
  markFieldTripCoveragePoolDirty_();
}

