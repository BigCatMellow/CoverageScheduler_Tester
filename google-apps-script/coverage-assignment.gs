// Coverage-need construction, automatic assignment, scoring, and plan-row creation.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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
