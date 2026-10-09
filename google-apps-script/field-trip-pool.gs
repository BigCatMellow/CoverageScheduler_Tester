// Field-trip coverage-pool materialization, freshness, and released-staff availability.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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

function readFieldTripCoveragePoolRowsForDateFast_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return [];

  // If this request already needed the full pool, reuse it.
  if (Object.prototype.hasOwnProperty.call(COVERAGE_REQUEST_SHEET_CACHE_, FIELD_TRIP_COVERAGE_POOL_SHEET_)) {
    incrementCoverageMetric_('requestCacheHits');
    return COVERAGE_REQUEST_SHEET_CACHE_[FIELD_TRIP_COVERAGE_POOL_SHEET_]
      .filter(row => normalizeDateKey_(row.Date) === key);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(value => String(value || '').trim());
  const aliases = HEADER_ALIASES[FIELD_TRIP_COVERAGE_POOL_SHEET_];
  const dateColumn = findColumnByAliases_(headers, aliases.Date);
  if (dateColumn === -1) return [];

  incrementCoverageMetric_('sheetReads');
  const dates = sheet.getRange(2, dateColumn + 1, sheet.getLastRow() - 1, 1).getValues();
  const matches = [];
  dates.forEach((row, index) => {
    if (normalizeDateKey_(row[0]) === key) matches.push(index + 2);
  });
  if (!matches.length) return [];

  // Pool rows are sorted, so a date is normally contiguous. Reading the span
  // and filtering is still safe if an operator has manually disturbed the sort.
  const first = matches[0];
  const last = matches[matches.length - 1];
  incrementCoverageMetric_('sheetReads');
  const values = sheet.getRange(first, 1, last - first + 1, headers.length).getValues();

  const columnByCanonical = {};
  Object.keys(aliases).forEach(canonical => {
    columnByCanonical[canonical] = findColumnByAliases_(headers, aliases[canonical]);
  });
  return values.map(row => {
    const out = {};
    Object.keys(columnByCanonical).forEach(canonical => {
      const index = columnByCanonical[canonical];
      out[canonical] = index === -1 ? '' : row[index];
    });
    return out;
  }).filter(row => normalizeDateKey_(row.Date) === key);
}

function fieldTripCoveragePoolRowsForDate_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasSheet = !!ss.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_);
  if (hasSheet && fieldTripCoveragePoolDateIsFresh_(key)) {
    return readFieldTripCoveragePoolRowsForDateFast_(key);
  }

  // Source/config changes invalidate the pool globally, but Generate only needs
  // one date. Rebuild that date under the existing re-entrant lock, then remember
  // its revision so repeated Generate clicks remain fast while other dates stay
  // correctly marked stale until they are used.
  return withCoverageLock_(() => {
    const current = SpreadsheetApp.getActiveSpreadsheet();
    if (current.getSheetByName(FIELD_TRIP_COVERAGE_POOL_SHEET_) &&
        fieldTripCoveragePoolDateIsFresh_(key)) {
      return readFieldTripCoveragePoolRowsForDateFast_(key);
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
