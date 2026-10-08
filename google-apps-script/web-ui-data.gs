function ensureStaffListSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Staff List');
  if (!sheet) {
    sheet = ss.insertSheet('Staff List');
    sheet.getRange('A1').setValue('Teacher').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  if (sheet.getLastRow() < 2) {
    const names = {};
    readSheetObjects_('Teacher Schedule')
      .map(row => normalizeTeacherScheduleRow_(row))
      .forEach(row => {
        const name = String(row.staffName || '').trim();
        if (name) names[name] = true;
      });
    const sorted = Object.keys(names).sort((a, b) => a.localeCompare(b));
    if (sorted.length) {
      sheet.getRange(2, 1, sorted.length, 1).setValues(sorted.map(name => [name]));
    }
  }
  return sheet;
}

function getWebStaffRosterSummary_() {
  const staffSheet = ensureStaffListSheet_();
  const values = staffSheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value || '').trim());
  const teacherCol = findRosterColumn_(headers, ['Teacher', 'Staff', 'Staff Name', 'Name']);
  if (teacherCol === -1) {
    throw new Error('Staff List needs a column named Teacher.');
  }

  const seen = {};
  return values.slice(1)
    .map(row => String(row[teacherCol] || '').trim())
    .filter(name => {
      const key = rosterNameKey_(name);
      if (!name || seen[key]) return false;
      seen[key] = true;
      return true;
    })
    .sort((a, b) => a.localeCompare(b))
    .map(name => ({
      name: name,
      displayName: name,
      scheduleName: name,
      role: 'Teacher',
      subject: '',
      blocks: [],
      scheduleMatched: null,
      scheduleLoaded: false
    }));
}

function getWebStaffSchedule_(staffName, date, dayCode) {
  const requestedName = String(staffName || '').trim();
  if (!requestedName) throw new Error('Choose a staff member.');

  const targetKey = rosterNameKey_(requestedName);
  const config = getConfigMap_();
  const dateKey = normalizeDateKey_(date);
  const day = String(dayCode || guessDayCodeFromDate_(dateKey) || '').trim();

  const rows = teacherScheduleRowsForDate_(dateKey, day, config)
    .map(row => normalizeTeacherScheduleRow_(row))
    .filter(row =>
      (!day || row.day === day) &&
      rosterNameKey_(row.staffName) === targetKey
    );

  const blocks = rows
    .filter(row => row.startMinutes != null && row.endMinutes != null)
    .map(row => ({
      start: minutesToDisplay_(row.startMinutes),
      end: minutesToDisplay_(row.endMinutes),
      className: row.className,
      grade: row.grade,
      subject: row.subject,
      assignmentType: row.assignmentType,
      room: row.room,
      needsCoverageIfAbsent: row.needsCoverageIfAbsent,
      coverEligibleThisBlock: row.coverEligibleThisBlock
    }))
    .sort((a, b) => displayTimeToMinutes_(a.start) - displayTimeToMinutes_(b.start));

  const first = rows[0] || null;
  return {
    name: requestedName,
    displayName: requestedName,
    scheduleName: first ? first.staffName : requestedName,
    role: first ? (first.role || 'Teacher') : 'Teacher',
    subject: summarizeRosterSubjects_(blocks),
    blocks: blocks,
    scheduleMatched: !!first,
    scheduleLoaded: true
  };
}

function getWebStaffRoster_(dayCode) {
  const staffSheet = ensureStaffListSheet_();
  const values = staffSheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value || '').trim());
  const teacherCol = findRosterColumn_(headers, ['Teacher', 'Staff', 'Staff Name', 'Name']);
  if (teacherCol === -1) {
    throw new Error('Staff List needs a column named Teacher.');
  }

  const rosterNames = [];
  const seenRoster = {};
  values.slice(1).forEach(row => {
    const name = String(row[teacherCol] || '').trim();
    const key = rosterNameKey_(name);
    if (!name || seenRoster[key]) return;
    seenRoster[key] = true;
    rosterNames.push(name);
  });

  const scheduleSource = dayCode
    ? readTeacherScheduleDayCached_(dayCode)
    : readSheetObjects_('Teacher Schedule');
  const scheduleRows = scheduleSource
    .map(row => normalizeTeacherScheduleRow_(row))
    .filter(row => !dayCode || row.day === String(dayCode).trim());

  const scheduleByKey = {};
  scheduleRows.forEach(row => {
    const key = rosterNameKey_(row.staffName);
    if (!key) return;
    if (!scheduleByKey[key]) {
      scheduleByKey[key] = {
        scheduleName: row.staffName,
        role: row.role || 'Teacher',
        blocks: []
      };
    }
    if (row.startMinutes != null && row.endMinutes != null) {
      scheduleByKey[key].blocks.push({
        start: minutesToDisplay_(row.startMinutes),
        end: minutesToDisplay_(row.endMinutes),
        className: row.className,
        grade: row.grade,
        subject: row.subject,
        assignmentType: row.assignmentType,
        room: row.room,
        needsCoverageIfAbsent: row.needsCoverageIfAbsent
      });
    }
  });

  return rosterNames
    .sort((a, b) => a.localeCompare(b))
    .map(displayName => {
      const match = scheduleByKey[rosterNameKey_(displayName)] || null;
      const blocks = match ? match.blocks.slice() : [];
      blocks.sort((a, b) => displayTimeToMinutes_(a.start) - displayTimeToMinutes_(b.start));
      return {
        name: displayName,
        displayName: displayName,
        scheduleName: match ? match.scheduleName : displayName,
        role: match ? match.role : 'Teacher',
        subject: summarizeRosterSubjects_(blocks),
        blocks: blocks,
        scheduleMatched: !!match
      };
    });
}

function findRosterColumn_(headers, aliases) {
  const normalized = headers.map(header => String(header || '').trim().toLowerCase());
  for (let i = 0; i < aliases.length; i++) {
    const idx = normalized.indexOf(String(aliases[i]).toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

function rosterNameKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function summarizeRosterSubjects_(blocks) {
  const seen = {};
  const subjects = [];
  (blocks || []).forEach(block => {
    const type = String(block.assignmentType || '').toLowerCase();
    const subject = String(block.subject || '').trim();
    if (!subject || ['break', 'lunch', 'planning', 'duty'].indexOf(type) !== -1) return;
    const key = subject.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    subjects.push(subject);
  });
  if (!subjects.length) return '';
  if (subjects.length <= 2) return subjects.join(' / ');
  return subjects.slice(0, 2).join(' / ') + ' +' + (subjects.length - 2);
}

function parsePositiveWholeNumberField_(value, label) {
  const raw = String(value == null ? '' : value).trim();
  if (raw === '') return ''; // blank means no limit — preserved as-is downstream

  const num = Number(raw);
  if (!Number.isFinite(num) || !Number.isInteger(num) || num < 1) {
    throw new Error(label + ' must be a whole number of 1 or more, or left blank for no limit.');
  }
  return String(num);
}

function saveCoverageStaffFromWeb_(payload) {
  return withCoverageLock_(() => saveCoverageStaffFromWebUnlocked_(payload));
}

function saveCoverageStaffFromWebUnlocked_(payload) {
  payload = payload || {};
  const name = String(payload.name || '').trim();
  const originalName = String(payload.originalName || '').trim();
  if (!name) throw new Error('Enter a name for the coverage staff member.');

  const sheetName = getCoverageStaffSheetName_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet: ' + sheetName);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const aliasMap = HEADER_ALIASES['Coverage Staff'];
  const nameCol = findColumnByAliases_(headers, aliasMap.Name);
  if (nameCol === -1) throw new Error('Coverage Staff needs a Name column.');

  let targetRow = -1;
  if (originalName) {
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][nameCol] || '').trim() === originalName) {
        targetRow = r + 1;
        break;
      }
    }
  }

  let duplicateRow = -1;
  for (let r = 1; r < values.length; r++) {
    if (r + 1 === targetRow) continue;
    if (String(values[r][nameCol] || '').trim() === name) {
      duplicateRow = r + 1;
      break;
    }
  }

  // This must reject a collision whether we're adding a brand-new person
  // (originalName blank) or renaming an existing one into another name
  // that's already taken — either way, writing to targetRow would silently
  // overwrite that other person's tier, availability, and restrictions.
  if (duplicateRow !== -1) {
    throw new Error('A coverage staff member named ' + name + ' already exists.');
  }

  if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

  const maxBlocksPerDay = parsePositiveWholeNumberField_(payload.maxBlocksPerDay, 'Max Blocks Per Day');
  const maxTeachersPerDay = parsePositiveWholeNumberField_(payload.maxTeachersPerDay, 'Max Teachers Per Day');

  const rowObject = {
    Name: name,
    Role: String(payload.role || 'Substitute').trim(),
    Coverage_Tier: String(Number(payload.tier) || 1),
    Can_Cover_All_Day: normalizeYesNo_(payload.canCoverAllDay, false) ? 'Yes' : 'No',
    Active_Today: normalizeYesNo_(payload.activeByDefault, true) ? 'Yes' : 'No',
    Available_Days: String(payload.availableDays || '').trim(),
    Default_Start: timeToDisplay_(payload.defaultStart),
    Default_End: timeToDisplay_(payload.defaultEnd),
    Allowed_Grades: String(payload.allowedGrades || '').trim(),
    Allowed_Subjects: String(payload.allowedSubjects || '').trim(),
    Allowed_Assignment_Types: String(payload.allowedAssignmentTypes || '').trim(),
    Max_Blocks_Per_Day: maxBlocksPerDay,
    Max_Teachers_Per_Day: maxTeachersPerDay,
    Can_Be_Split_Across_Teachers: normalizeYesNo_(payload.canBeSplitAcrossTeachers, true) ? 'Yes' : 'No',
    Notes: String(payload.notes || '').trim()
  };

  setSheetRowObject_(sheet, targetRow, headers, aliasMap, rowObject);
  markFieldTripCoveragePoolDirty_();

  if (originalName && originalName !== name) {
    renameCoverageAvailabilityRows_(originalName, name);
  }

  const selectedDate = normalizeDateKey_(payload.date);
  const selectedDay = String(payload.day || guessDayCodeFromDate_(selectedDate) || '').trim();
  return getAllCoverageStaff_(selectedDate, selectedDay);
}

function deleteCoverageStaffFromWeb_(payload) {
  return withCoverageLock_(() => deleteCoverageStaffFromWebUnlocked_(payload));
}

function deleteCoverageStaffFromWebUnlocked_(payload) {
  payload = payload || {};
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('No coverage staff name provided.');

  const sheetName = getCoverageStaffSheetName_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const nameCol = findColumnByAliases_(headers, HEADER_ALIASES['Coverage Staff'].Name);
  if (nameCol === -1) throw new Error('Coverage Staff needs a Name column.');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][nameCol] || '').trim() === name) {
      sheet.deleteRow(r + 1);
      incrementCoverageMetric_('sheetWrites');
      invalidateCoverageSheetCache_(sheetName);
      markFieldTripCoveragePoolDirty_();
      break;
    }
  }

  const selectedDate = normalizeDateKey_(payload.date);
  const selectedDay = String(payload.day || guessDayCodeFromDate_(selectedDate) || '').trim();
  return getAllCoverageStaff_(selectedDate, selectedDay);
}

function renameCoverageAvailabilityRows_(oldName, newName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Substitute Availability');
  if (!sheet || sheet.getLastRow() < 2) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const nameCol = findColumnByAliases_(headers, HEADER_ALIASES['Substitute Availability'].Name);
  if (nameCol === -1) return;
  let changed = false;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][nameCol] || '').trim() === oldName) {
      sheet.getRange(r + 1, nameCol + 1).setValue(newName);
      changed = true;
    }
  }
  if (changed) {
    incrementCoverageMetric_('sheetWrites');
    invalidateCoverageSheetCache_('Substitute Availability');
  }
}
