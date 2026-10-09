// Manual coverage choice workspace and save-time assignment validation.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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

function manualCoverageContext_(payload, liveOverride) {
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

  const live = liveOverride || buildCoverageLiveContext_(date, day);
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

function getManualCoverageChoices_(payload, liveOverride, options) {
  options = options || {};
  const context = manualCoverageContext_(payload, liveOverride);
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
  // unavailable" — seeing their day often explains why. Generate can skip
  // this per-block copy and build one shared browser workspace instead.
  const namesToDescribe = new Set(choices.map(choice => choice.name));
  if (currentName) namesToDescribe.add(currentName);
  const schedules = {};
  if (options.includeSchedules !== false) namesToDescribe.forEach(name => {
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

function buildManualChoiceWorkspace_(planRows, liveContext) {
  const rows = Array.isArray(planRows) ? planRows : [];
  if (!rows.length || rows.length > 40) return null;

  const blocks = {};
  const names = new Set();

  rows.forEach((row, index) => {
    try {
      const result = getManualCoverageChoices_({
        date: liveContext.date,
        day: liveContext.day,
        rows: rows,
        blockIndex: index
      }, liveContext, { includeSchedules: false });

      blocks[index] = {
        blockIndex: index,
        currentName: result.currentName || '',
        currentEligible: result.currentEligible !== false,
        choices: result.choices || [],
        excludedAbsentNames: result.excludedAbsentNames || []
      };
      (result.choices || []).forEach(choice => {
        if (choice && choice.name) names.add(choice.name);
      });
      if (result.currentName) names.add(result.currentName);
    } catch (error) {
      // A single unusual block should not slow/fail Generate. The browser will
      // transparently fall back to the live endpoint for that block.
    }
  });

  const fullState = manualCoverageStateFromPlan_(
    rows,
    -1,
    liveContext.candidates,
    liveContext.teacherSchedule,
    liveContext.day,
    liveContext.effectiveAbsences
  );
  const schedules = {};
  names.forEach(name => {
    schedules[name] = describeStaffDayForDisplay_(
      liveContext.teacherSchedule,
      liveContext.day,
      name,
      liveContext.fieldTrips,
      candidateBreakReservations_(name, fullState)
    );
  });

  return {
    date: liveContext.date,
    day: liveContext.day,
    blocks: blocks,
    schedules: schedules
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
