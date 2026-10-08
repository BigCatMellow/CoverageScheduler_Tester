const TEACHER_SCHEDULE_SOURCE_HEADERS = ['Teacher', 'Term', 'Day', 'Start', 'End', 'Class', 'Subject', 'Room'];

/**
 * Sets up the scheduler around an existing Teacher Schedule source sheet.
 *
 * The source schedule is intentionally left in its native eight-column shape:
 * Teacher | Term | Day | Start | End | Class | Subject | Room
 *
 * Grade, assignment type, coverage need, and cover eligibility are inferred at
 * runtime by scheduler.gs. This avoids maintaining a second transformed copy of
 * the schedule and avoids adding columns to the operational source sheet.
 */
function setupCoverageWorkbookFromTeacherSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Open the Google Sheet that will hold Coverage Scheduler before running setup.');
  }

  // Keep the standalone web app connected even when this lower-level setup
  // function is run directly from the Apps Script editor.
  if (typeof rememberCoverageSpreadsheet_ === 'function') {
    rememberCoverageSpreadsheet_();
  }

  let teacherSheet = ss.getSheetByName('Teacher Schedule');

  if (!teacherSheet) {
    teacherSheet = ss.insertSheet('Teacher Schedule');
    teacherSheet.getRange(1, 1, 1, TEACHER_SCHEDULE_SOURCE_HEADERS.length)
      .setValues([TEACHER_SCHEDULE_SOURCE_HEADERS]);
    teacherSheet.setFrozenRows(1);
    teacherSheet.getRange(1, 1, 1, TEACHER_SCHEDULE_SOURCE_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#d9eaf7')
      .setWrap(true);
  }

  const inspection = validateTeacherScheduleSource_();

  const legacyPool = ss.getSheetByName('Field Trip Coverage Cache');
  if (legacyPool && !ss.getSheetByName('Field Trip Coverage Pool')) {
    legacyPool.setName('Field Trip Coverage Pool');
    legacyPool.clearContents();
  }

  const managedSheets = [
    'Coverage Staff',
    'Substitute Availability',
    'Daily Absences',
    'Field Trip Coverage Pool',
    'Coverage Output',
    'Lists',
    'Config',
    '_Preview'
  ];

  managedSheets.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    ensureHeaderRow_(sheet, SHEET_SCHEMAS[name].headers);
    formatSheet_(sheet);
    if (name === 'Field Trip Coverage Pool') formatFieldTripCoveragePoolSheet_(sheet);
  });

  seedLists_();
  seedConfig_();
  applyDataValidation_();
  hideHelperSheets_();
  if (typeof markTeacherScheduleCacheDirty_ === 'function') {
    markTeacherScheduleCacheDirty_();
  }
  if (typeof markFieldTripCoveragePoolDirty_ === 'function') {
    markFieldTripCoveragePoolDirty_();
  }

  if (typeof ensureStaffListForWeb_ === 'function') {
    ensureStaffListForWeb_();
  }

  const warningSuffix = inspection.warnings.length
    ? ' ' + inspection.warnings.length + ' schedule warning(s) found; use Coverage Scheduler → Validate teacher schedule.'
    : '';
  ss.toast(
    'Coverage Scheduler is ready. Using ' + inspection.teacherCount + ' teacher schedule(s).' + warningSuffix,
    APP_TITLE,
    8
  );

  return inspection;
}

/**
 * Validates the source structure without changing Teacher Schedule.
 * The runtime scheduler already accepts the older Staff_Name-style schema, so
 * the validator keeps those aliases for backward compatibility.
 */
function validateTeacherScheduleSource_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher Schedule');
  if (!sheet) throw new Error('Missing Teacher Schedule sheet.');
  if (sheet.getLastColumn() < 1) throw new Error('Teacher Schedule has no columns.');

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(value => String(value || '').trim());

  const headerGroups = {
    Teacher: ['Teacher', 'Staff_Name', 'Teacher_Name', 'Name'],
    Term: ['Term'],
    Day: ['Day'],
    Start: ['Start', 'Start_Time', 'Start Time'],
    End: ['End', 'End_Time', 'End Time'],
    Class: ['Class', 'Class_Name', 'Class Name', 'Section', 'Group'],
    Subject: ['Subject', 'Course', 'Course_Name', 'Course Name', 'Activity'],
    Room: ['Room']
  };

  const positions = {};
  Object.keys(headerGroups).forEach(key => {
    positions[key] = findTeacherScheduleHeader_(headers, headerGroups[key]);
  });

  const required = ['Teacher', 'Day', 'Start', 'End', 'Subject'];
  const missing = required.filter(key => positions[key] === -1);
  if (missing.length) {
    throw new Error(
      'Teacher Schedule is missing required column(s): ' + missing.join(', ') +
      '. Expected the source format Teacher | Term | Day | Start | End | Class | Subject | Room.'
    );
  }

  const warnings = [];
  ['Term', 'Class', 'Room'].forEach(key => {
    if (positions[key] === -1) warnings.push('Recommended column missing: ' + key + '.');
  });

  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (!rowCount) {
    return {
      rowCount: 0,
      teacherCount: 0,
      terms: [],
      warnings: warnings.concat(['Teacher Schedule has no schedule rows yet.'])
    };
  }

  const values = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
  const teachers = {};
  const terms = {};
  let invalidDayCount = 0;
  let invalidTimeCount = 0;
  let blankTeacherCount = 0;

  values.forEach(row => {
    const teacher = String(row[positions.Teacher] || '').trim();
    if (!teacher) {
      blankTeacherCount++;
      return;
    }
    teachers[teacher] = true;

    if (positions.Term !== -1) {
      const term = String(row[positions.Term] || '').trim();
      if (term) terms[term] = true;
    }

    const day = String(row[positions.Day] || '').trim().toUpperCase();
    if (['M', 'T', 'W', 'R', 'F'].indexOf(day) === -1) invalidDayCount++;

    const start = timeToMinutes_(row[positions.Start]);
    const end = timeToMinutes_(row[positions.End]);
    if (start == null || end == null || end <= start) invalidTimeCount++;
  });

  if (blankTeacherCount) warnings.push(blankTeacherCount + ' row(s) have no teacher name and will be ignored.');
  if (invalidDayCount) warnings.push(invalidDayCount + ' row(s) use a day code other than M/T/W/R/F.');
  if (invalidTimeCount) warnings.push(invalidTimeCount + ' row(s) have an invalid or reversed Start/End time.');

  const overlapInfo = teacherScheduleOverlaps_(readSheetObjects_('Teacher Schedule'));
  if (overlapInfo.count) {
    warnings.push(
      overlapInfo.count + ' pair(s) of rows put the same person in two places at once (' +
      overlapInfo.people + ' staff). Examples: ' + overlapInfo.examples.join('; ') +
      '. The scheduler treats these as two separate rooms, so one coverage person can never take both; ' +
      'if they are data-entry artifacts, fix the rows or several blocks may show Unfilled.'
    );
  }

  const termList = Object.keys(terms).sort();
  const recognizedTerms = { '': true, S1: true, S2: true };
  const unrecognizedTerms = termList.filter(term => !recognizedTerms[normalizeScheduleTerm_(term)]);
  const usesRecognizedSemesters = termList.some(term => {
    const normalized = normalizeScheduleTerm_(term);
    return normalized === 'S1' || normalized === 'S2';
  });

  if (unrecognizedTerms.length) {
    warnings.push(
      'Unrecognized Term value(s): ' + unrecognizedTerms.join(', ') +
      '. Recognized values are blank/"All Year" and "S1"/"S2" (also accepts "Semester 1/2", "Term 1/2") — those are already date-filtered automatically. ' +
      (usesRecognizedSemesters
        ? 'Because this sheet also has recognized S1/S2 rows, rows using an unrecognized term will never be treated as active for coverage — fix the term value or these rows will be silently skipped every day.'
        : 'Right now these rows are treated as active on every date, since no recognized S1/S2 rows are present elsewhere to trigger semester filtering.')
    );
  }

  const normalizedRows = readSheetObjects_('Teacher Schedule').map(row => normalizeTeacherScheduleRow_(row));
  const classCheck = classScheduleConsistency_(normalizedRows);
  if (classCheck.error) warnings.push(classCheck.error);
  if (classCheck.sheetName) {
    const enriched = normalizedRows.filter(row => row.gradeSource === 'Class Schedule').length;
    const counts = {};
    classCheck.findings.forEach(f => { counts[f.type] = (counts[f.type] || 0) + 1; });
    const fixes = classCheck.findings.filter(f => f.severity === 'Fix').length;
    warnings.push(
      'Class Schedule ("' + classCheck.sheetName + '"): ' + enriched +
      ' Teacher Schedule row(s) without a Class got their grade from it. ' +
      (classCheck.findings.length
        ? fixes + ' issue(s) to fix, ' + (classCheck.findings.length - fixes) + ' note(s): ' +
          Object.keys(counts).sort().map(type => counts[type] + ' ' + type.toLowerCase()).join(', ') +
          '. Full list is on the ' + SCHEDULE_CHECK_SHEET_ + ' sheet.'
        : 'It agrees with Teacher Schedule.')
    );
  }

  return {
    rowCount: rowCount,
    teacherCount: Object.keys(teachers).length,
    terms: termList,
    warnings: warnings,
    classScheduleFindings: classCheck.findings
  };
}

function findTeacherScheduleHeader_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = headers.indexOf(aliases[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}


// Pairs of rows needing coverage that overlap in time for the same person,
// day, and term. Rows in different semesters never overlap in practice;
// All Year rows overlap with both semesters.
function teacherScheduleOverlaps_(rawRows) {
  const groups = {};
  (rawRows || []).forEach(raw => {
    const row = normalizeTeacherScheduleRow_(raw);
    if (!row.staffName || !row.needsCoverageIfAbsent || row.startMinutes == null || row.endMinutes == null) return;
    const key = row.staffName + '\u0000' + row.day;
    (groups[key] || (groups[key] = [])).push(row);
  });

  let count = 0;
  const people = new Set();
  const examples = [];
  Object.keys(groups).forEach(key => {
    const rows = groups[key].sort((a, b) => a.startMinutes - b.startMinutes);
    for (let i = 0; i < rows.length; i++) {
      for (let j = i + 1; j < rows.length && rows[j].startMinutes < rows[i].endMinutes; j++) {
        const a = rows[i], b = rows[j];
        if (a.term && b.term && a.term !== b.term) continue;
        count++;
        people.add(a.staffName);
        if (examples.length < 3) {
          examples.push(a.staffName + ' (' + a.day + ') ' + minutesToDisplay_(a.startMinutes) + ' ' + a.className +
            ' / ' + minutesToDisplay_(b.startMinutes) + ' ' + b.className);
        }
      }
    }
  });
  return { count: count, people: people.size, examples: examples };
}
