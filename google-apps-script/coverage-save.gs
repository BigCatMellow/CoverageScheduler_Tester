// Durable coverage-plan save/reload and legacy preview helpers.
// Structural extraction from scheduler.gs; behavior intentionally unchanged.

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

function getSavedCoverageForDate_(date, day) {
  const dateKey = normalizeDateKey_(date);
  const dayCode = String(day || guessDayCodeFromDate_(dateKey) || '').trim();

  let rows = readSheetObjects_('Coverage Output');
  if (dateKey) {
    rows = rows.filter(row => normalizeDateKey_(row.Date) === dateKey);
  }
  if (dayCode) {
    rows = rows.filter(row => String(row.Day || '').trim() === dayCode);
  }

  return {
    rows: rows,
    summary: {
      totalBlocks: rows.length,
      assignedBlocks: rows.filter(row => String(row.Status || '').trim() === 'Assigned').length,
      unfilledBlocks: rows.filter(row => String(row.Status || '').trim() !== 'Assigned').length
    }
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
