// Prepared weekday teacher-schedule cache and schedule normalization.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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
