// Coverage-staff roster, date overrides, eligibility normalization, and matching.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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
