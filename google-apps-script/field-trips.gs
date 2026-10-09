// Field-trip persistence, normalization, date behavior, and participant absences.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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
