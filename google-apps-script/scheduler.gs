const COVERAGE_SCHEDULER_MODULE_VERSION = 1;

// Top-level coverage generation orchestration.
// Supporting storage, field-trip, cache, candidate, assignment, manual-placement,
// and save/reload helpers live in focused sibling .gs modules. Apps Script loads
// all server .gs files into the same global namespace.

function generateCoveragePreview(payload) {
  const generateStartedAt = Date.now();
  let timingCursor = generateStartedAt;
  const timingMs = {};
  const timingMark = name => {
    const now = Date.now();
    timingMs[name] = now - timingCursor;
    timingCursor = now;
  };
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
  timingMark('snapshot');

  const config = getConfigMap_();
  const teacherSchedule = teacherScheduleRowsForDate_(date, day, config);
  timingMark('teacherSchedule');
  const configuredCoverageStaff = getCoverageStaffForDate_(date, day, config);
  const activeCoverageStaff = configuredCoverageStaff.filter(row => row.name && row.activeToday);
  const absences = getDailyAbsencesForDate_(date, day);
  const fieldTrips = getFieldTripsForDate_(date);
  const fieldTripAbsences = buildFieldTripParticipantAbsences_(fieldTrips, teacherSchedule, day);
  const effectiveAbsences = absences.concat(fieldTripAbsences);
  const fieldTripPoolRows = fieldTrips.length ? fieldTripCoveragePoolRowsForDate_(date) : [];
  timingMark('dayInputsAndFieldTripPool');
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
  timingMark('coverageNeeds');

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
  timingMark('scheduler');

  let manualWorkspace = null;
  if (planRows.length && planRows.length <= 40) {
    const liveForWorkspace = buildCoverageLiveContext_(date, day);
    manualWorkspace = buildManualChoiceWorkspace_(planRows, liveForWorkspace);
  }
  timingMark('manualWorkspace');

  if (payload.persistPreview === true) {
    writePreview_(planRows);
    coveragePerfMark_('preview-written');
  }
  timingMark('previewWrite');

  summary.elapsedMs = Date.now() - generateStartedAt;
  summary.timingMs = timingMs;

  return {
    date: date,
    day: day,
    summary: summary,
    rows: planRows,
    manualWorkspace: manualWorkspace,
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
