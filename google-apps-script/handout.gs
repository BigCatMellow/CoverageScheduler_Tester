const FIELD_TRIP_HANDOUT_MODULE_VERSION = 4;
const FIELD_TRIP_FORM_MAX_ASSIGNMENTS_ = 6;
const FIELD_TRIP_FORMS_PER_PAGE_ = 2;
const FIELD_TRIP_FORM_COLUMN_WIDTHS_ = [105, 87, 98, 175, 75];

function createCoverageHandoutDocWideFromLatestPreview_() {
  const preview = getLatestPreview_();
  const planRows = preview.rows || [];
  if (!planRows.length) {
    throw new Error('No preview rows found to build handouts.');
  }

  const date = normalizeDateKey_(planRows[0].Date) || Utilities.formatDate(new Date(), coverageTimeZone_(), 'yyyy-MM-dd');
  const day = String(planRows[0].Day || guessDayCodeFromDate_(date) || '').trim();
  if (!date || !day) throw new Error('The preview does not have a valid date.');
  if (planRows.some(row => normalizeDateKey_(row.Date) !== date)) {
    throw new Error('A handout can only be created for one date at a time.');
  }

  // Creating a handout is independent from Save Plan, but it must meet the
  // same live safety rules. Refuse stale/invalid assignments instead of
  // printing a plan that Save would reject.
  validateCoveragePlanForSave_(date, day, planRows);

  const assignedRows = planRows.filter(row =>
    String(row.Status || '').trim() === 'Assigned' &&
    String(row.Assigned_Coverage || '').trim()
  );
  if (!assignedRows.length) {
    throw new Error('No assigned preview rows found to build handouts.');
  }
  return createCoverageHandoutPackage_(assignedRows, date, day);
}

function createCoverageHandoutPackage_(rows, date, day) {
  const assignedRows = (rows || []).filter(row =>
    String(row.Status || '').trim() === 'Assigned' &&
    String(row.Assigned_Coverage || '').trim()
  );

  if (!assignedRows.length) {
    throw new Error('No assigned coverage rows are available for a handout.');
  }

  const ordinaryRows = [];
  const fieldTripRowsByEvent = {};

  assignedRows.forEach(row => {
    const eventId = String(row.Event_ID || '').trim();
    if (!eventId) {
      ordinaryRows.push(row);
      return;
    }
    if (!fieldTripRowsByEvent[eventId]) fieldTripRowsByEvent[eventId] = [];
    fieldTripRowsByEvent[eventId].push(row);
  });

  const fieldTripEventIds = Object.keys(fieldTripRowsByEvent);
  const folder = getCoverageHandoutFolder_();
  const documents = [];

  let fieldTripContext = null;

  if (fieldTripEventIds.length) {
    fieldTripContext = buildFieldTripFormContext_(date, day);
  }

  if (ordinaryRows.length) {
    const ordinary = createCoverageHandoutDocWide_(ordinaryRows, date, day);
    documents.push(Object.assign({
      type: 'coverage',
      label: 'Coverage Handout'
    }, ordinary));
  }

  fieldTripEventIds.sort().forEach(eventId => {
    const trip = fieldTripContext.tripsById[eventId];
    if (!trip) {
      throw new Error(
        'Field trip ' + eventId + ' is no longer available for ' + date +
        '. Re-generate the plan before creating the handout.'
      );
    }

    const tripDocs = createFieldTripCoverageFormDocs_(
      fieldTripRowsByEvent[eventId],
      trip,
      date,
      day,
      fieldTripContext,
      folder
    );
    tripDocs.forEach(doc => documents.push(doc));
  });

  if (!documents.length) {
    throw new Error('No handout documents were created.');
  }

  const first = documents[0];
  return {
    id: first.id,
    url: first.url,
    name: documents.length === 1 ? first.name : ('Coverage Handouts - ' + date),
    documents: documents,
    folderId: folder.getId(),
    folderUrl: folder.getUrl(),
    folderName: folder.getName()
  };
}

function buildFieldTripFormContext_(date, day) {
  const config = getConfigMap_();
  const teacherSchedule = teacherScheduleRowsForDate_(date, day, config)
    .map(row => normalizeTeacherScheduleRow_(row))
    .filter(row => row.day === day && row.staffName);

  const trips = getFieldTripsForDate_(date);
  const tripsById = {};
  trips.forEach(trip => {
    if (trip.eventId) tripsById[trip.eventId] = trip;
  });

  const effectiveAbsences = getDailyAbsencesForDate_(date, day)
    .concat(buildFieldTripParticipantAbsences_(trips, teacherSchedule, day));
  const absenceState = makeEmptyState_();
  absenceState.absencesByCandidate = buildAbsenceWindowsByStaff_(effectiveAbsences);

  return {
    teacherSchedule: teacherSchedule,
    tripsById: tripsById,
    absenceState: absenceState
  };
}

function createFieldTripCoverageFormDocs_(rows, trip, date, day, context, folder) {
  const byTeacher = {};
  (rows || []).forEach(row => {
    const teacher = String(row.Absent_Staff || '').trim();
    if (!teacher) return;
    if (!byTeacher[teacher]) byTeacher[teacher] = [];
    byTeacher[teacher].push(row);
  });

  const formUnits = [];
  Object.keys(byTeacher).sort((a, b) => a.localeCompare(b)).forEach(teacher => {
    const teacherRows = byTeacher[teacher]
      .slice()
      .sort((a, b) =>
        timeToMinutes_(a.Start) - timeToMinutes_(b.Start) ||
        String(a.Class || '').localeCompare(String(b.Class || ''))
      );

    for (let offset = 0; offset < teacherRows.length; offset += FIELD_TRIP_FORM_MAX_ASSIGNMENTS_) {
      formUnits.push({
        coverageFor: teacher,
        rows: teacherRows.slice(offset, offset + FIELD_TRIP_FORM_MAX_ASSIGNMENTS_)
      });
    }
  });

  if (!formUnits.length) return [];

  const outputName = 'Field Trip Coverage - ' + (trip.name || 'Field Trip') + ' - ' + date;
  const doc = DocumentApp.create(outputName);
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  const body = doc.getBody();

  formatGeneratedFieldTripHandoutPage_(body);

  formUnits.forEach((unit, formIndex) => {
    appendGeneratedFieldTripForm_(body, unit, trip, date, context);

    if (formIndex >= formUnits.length - 1) return;

    if ((formIndex + 1) % FIELD_TRIP_FORMS_PER_PAGE_ === 0) {
      body.appendPageBreak();
    } else {
      appendGeneratedFieldTripFormSpacer_(body);
    }
  });

  if (body.getTables().length !== formUnits.length) {
    throw new Error(
      'Field trip handout build failed: expected ' + formUnits.length +
      ' table(s), found ' + body.getTables().length + '.'
    );
  }

  doc.saveAndClose();

  return [{
    type: 'field-trip',
    label: 'Field Trip Form - ' + (trip.name || 'Field Trip'),
    eventId: trip.eventId || '',
    id: doc.getId(),
    url: doc.getUrl(),
    name: doc.getName(),
    folderId: folder.getId(),
    folderUrl: folder.getUrl(),
    folderName: folder.getName()
  }];
}

function formatGeneratedFieldTripHandoutPage_(body) {
  // Mirrors the approved reference form: US Letter portrait with 0.5" margins.
  body.setPageWidth(612);
  body.setPageHeight(792);
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);
}

function appendGeneratedFieldTripForm_(body, formUnit, trip, date, context) {
  const heading = appendGeneratedFieldTripMetaLine_(
    body,
    'FIELD TRIP COVERAGE FORM\t\t\t\tPREPARED BY:  C. SCHREMBS'
  );
  const headingText = heading.editAsText();
  headingText.setUnderline(0, 'FIELD TRIP COVERAGE FORM'.length - 1, true);

  appendGeneratedFieldTripMetaLine_(
    body,
    'COVERAGE FOR:  ' + String(formUnit.coverageFor || '').trim() +
      '\t\t\t\tDATE:  ' + formatFieldTripFormDate_(date)
  );
  appendGeneratedFieldTripMetaLine_(
    body,
    'CLASS(ES) TAKING TRIP: ' + formatFieldTripFormGrades_(trip.grades)
  );
  appendGeneratedFieldTripMetaLine_(
    body,
    'TRIP DESTINATION:  ' + fieldTripHandoutDestination_(trip)
  );
  appendGeneratedFieldTripMetaLine_(
    body,
    'DEPARTURE TIME: ' + String(trip.start || '').trim() +
      '\t\t\tAPPROXIMATE RETURN TIME: ' + String(trip.end || '').trim()
  );

  const tableData = [[
    'NAME',
    'TIME NEEDED',
    'RM#/SUBJECT',
    'SPECIAL INSTRUCTIONS',
    'WITH'
  ]];

  (formUnit.rows || []).forEach(row => {
    tableData.push([
      String(row.Assigned_Coverage || '').trim(),
      formatWideHandoutTimeRange_(row.Start, row.End),
      formatFieldTripRoomSubject_(row),
      String(trip.notes || '').trim(),
      resolveFieldTripWithTeacher_(row, trip, context)
    ]);
  });

  const table = body.appendTable(tableData);
  styleGeneratedFieldTripTable_(table);
}

function appendGeneratedFieldTripMetaLine_(body, text) {
  const paragraph = body.appendParagraph(String(text == null ? '' : text));
  paragraph
    .setSpacingBefore(0)
    .setSpacingAfter(0)
    .setLineSpacing(1.5);

  paragraph.editAsText()
    .setFontFamily('Cambria')
    .setFontSize(12)
    .setBold(true);

  return paragraph;
}

function appendGeneratedFieldTripFormSpacer_(body) {
  // The original two-up form used six blank 11pt lines between half-forms.
  for (let i = 0; i < 6; i++) {
    const paragraph = body.appendParagraph('');
    paragraph
      .setSpacingBefore(0)
      .setSpacingAfter(0)
      .setLineSpacing(1);
    paragraph.editAsText()
      .setFontFamily('Cambria')
      .setFontSize(11);
  }
}

function styleGeneratedFieldTripTable_(table) {
  FIELD_TRIP_FORM_COLUMN_WIDTHS_.forEach((width, columnIndex) => {
    table.setColumnWidth(columnIndex, width);
  });

  table.setBorderColor('#000000');
  table.setBorderWidth(1);

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    row.setMinimumHeight(r === 0 ? 14 : 15);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
      cell.setPaddingTop(1);
      cell.setPaddingBottom(1);
      cell.setPaddingLeft(2);
      cell.setPaddingRight(2);

      for (let p = 0; p < cell.getNumChildren(); p++) {
        const child = cell.getChild(p);
        if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

        child.asParagraph()
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setSpacingBefore(0)
          .setSpacingAfter(0)
          .setLineSpacing(1);
      }

      cell.editAsText()
        .setFontFamily('Cambria')
        .setFontSize(11)
        .setBold(r === 0 || (r > 0 && c === 0));
    }
  }
}

function formatFieldTripFormDate_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return String(date || '');

  const parsed = new Date(key + 'T12:00:00');
  if (isNaN(parsed)) return key;
  return Utilities.formatDate(parsed, coverageTimeZone_(), 'M/d/yyyy');
}

function formatFieldTripFormGrades_(grades) {
  return (grades || []).map(value => String(value || '').trim()).filter(Boolean).join(', ');
}

function formatFieldTripRoomSubject_(row) {
  const room = String(row.Room || '').trim();
  const subject = String(row.Subject || row.Class || row.Assignment_Type || '').trim();

  if (room && subject) return room + ' / ' + subject;
  return room || subject;
}

function normalizeFieldTripClassMatch_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function fieldTripClassSectionKey_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const section = raw.split(/\s+[—–-]\s+/)[0].trim();
  return normalizeFieldTripClassMatch_(section);
}

function resolveFieldTripWithTeacher_(row, trip, context) {
  const absentTeacher = String(row.Absent_Staff || '').trim();
  const assignedCoverage = String(row.Assigned_Coverage || '').trim();
  const block = planRowToCoverageBlock_(row);
  if (block.startMinutes == null || block.endMinutes == null) return '';

  const excluded = {};
  excluded[absentTeacher] = true;
  excluded[assignedCoverage] = true;
  (trip.staffNames || []).forEach(name => { excluded[String(name || '').trim()] = true; });

  const classKey = normalizeFieldTripClassMatch_(row.Class);
  const sectionKey = fieldTripClassSectionKey_(row.Class);
  const roomKey = normalizeFieldTripClassMatch_(row.Room);
  const subjectKey = normalizeFieldTripClassMatch_(row.Subject);

  const availableRows = (context.teacherSchedule || []).filter(scheduleRow => {
    if (!scheduleRow.staffName || excluded[scheduleRow.staffName]) return false;
    if (scheduleRow.startMinutes == null || scheduleRow.endMinutes == null) return false;
    if (scheduleRow.startMinutes > block.startMinutes || scheduleRow.endMinutes < block.endMinutes) return false;
    if (candidateIsAbsentForBlock_(scheduleRow.staffName, block, context.absenceState)) return false;

    const type = String(scheduleRow.assignmentType || '').trim().toLowerCase();
    if (['planning', 'break', 'lunch', 'meeting', 'duty'].indexOf(type) !== -1) return false;
    return true;
  });

  let matches = availableRows.filter(scheduleRow =>
    classKey &&
    normalizeFieldTripClassMatch_(scheduleRow.className) === classKey
  );

  // Co-teacher rows sometimes encode one another in Subject, e.g.
  // "3D — Logic w/Vlattas" / "3D — Logic w/Harrington". In that case the
  // full display names differ even though the section, room, and time are the
  // same. Use section + room as the next conservative match.
  if (!matches.length && sectionKey && roomKey) {
    matches = availableRows.filter(scheduleRow =>
      fieldTripClassSectionKey_(scheduleRow.className) === sectionKey &&
      normalizeFieldTripClassMatch_(scheduleRow.room) === roomKey
    );
  }

  if (!matches.length && roomKey && subjectKey) {
    matches = availableRows.filter(scheduleRow =>
      normalizeFieldTripClassMatch_(scheduleRow.room) === roomKey &&
      normalizeFieldTripClassMatch_(scheduleRow.subject) === subjectKey
    );
  }

  const names = Array.from(new Set(matches.map(scheduleRow => scheduleRow.staffName).filter(Boolean)));
  return names.length === 1 ? names[0] : '';
}

function createCoverageHandoutDocWide_(rows, date, day) {
  const grouped = {};

  rows.forEach(row => {
    const person = String(row.Assigned_Coverage || '').trim();
    if (!person) return;
    if (!grouped[person]) grouped[person] = [];
    grouped[person].push(row);
  });

  const names = Object.keys(grouped).sort((a, b) => a.localeCompare(b));
  if (!names.length) {
    throw new Error('No assigned coverage staff found to build handouts.');
  }

  const folder = getCoverageHandoutFolder_();
  const doc = DocumentApp.create('Coverage Handouts - ' + date + (day ? ' - ' + day : ''));
  DriveApp.getFileById(doc.getId()).moveTo(folder);

  const body = doc.getBody();
  formatWideHandoutPage_(body);

  names.forEach((name, index) => {
    const personRows = grouped[name]
      .slice()
      .sort((a, b) => timeToMinutes_(a.Start) - timeToMinutes_(b.Start));

    appendWideHandoutHeader_(body, name, date, personRows.length);

    const tableData = [['Time', 'Absent Teacher', 'Subject', 'Room']];
    personRows.forEach(row => {
      tableData.push([
        formatWideHandoutTimeRange_(row.Start, row.End),
        String(row.Absent_Staff || ''),
        String(row.Subject || row.Assignment_Type || ''),
        String(row.Room || '')
      ]);
    });

    const table = body.appendTable(tableData);
    styleWideHandoutTable_(table);

    if (index < names.length - 1) body.appendPageBreak();
  });

  doc.saveAndClose();

  return {
    id: doc.getId(),
    url: doc.getUrl(),
    name: doc.getName(),
    folderId: folder.getId(),
    folderUrl: folder.getUrl(),
    folderName: folder.getName()
  };
}

function getCoverageHandoutFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active Coverage Scheduler workbook is available.');

  const propertyKey = 'COVERAGE_HANDOUT_FOLDER_ID_' + ss.getId();
  const properties = PropertiesService.getScriptProperties();
  const savedFolderId = String(properties.getProperty(propertyKey) || '').trim();

  if (savedFolderId) {
    try {
      const savedFolder = DriveApp.getFolderById(savedFolderId);
      if (!savedFolder.isTrashed()) return savedFolder;
    } catch (error) {
      // The folder may have been deleted or access may have changed.
    }
    properties.deleteProperty(propertyKey);
  }

  const folderName = ss.getName() + ' - Handouts';
  let parent = null;

  try {
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parents = spreadsheetFile.getParents();
    if (parents.hasNext()) parent = parents.next();
  } catch (error) {
    // Fall back to My Drive root if the spreadsheet parent cannot be resolved.
  }

  if (!parent) parent = DriveApp.getRootFolder();

  const existingFolders = parent.getFoldersByName(folderName);
  while (existingFolders.hasNext()) {
    const existing = existingFolders.next();
    if (!existing.isTrashed()) {
      properties.setProperty(propertyKey, existing.getId());
      return existing;
    }
  }

  const folder = parent.createFolder(folderName);
  properties.setProperty(propertyKey, folder.getId());
  return folder;
}

function formatWideHandoutPage_(body) {
  // Letter landscape: 11 x 8.5 inches at 72 points per inch.
  body.setPageWidth(792);
  body.setPageHeight(612);

  // Small, print-safe margins leave 756 pt for the assignment table.
  body.setMarginTop(22);
  body.setMarginBottom(22);
  body.setMarginLeft(18);
  body.setMarginRight(18);
}

function appendWideHandoutHeader_(body, name, date, assignmentCount) {
  const eyebrow = body.appendParagraph('COVERAGE ASSIGNMENTS');
  eyebrow
    .setSpacingBefore(0)
    .setSpacingAfter(2);
  eyebrow.editAsText()
    .setFontSize(8)
    .setBold(true)
    .setForegroundColor('#64748B');

  const title = body.appendParagraph(name);
  title
    .setSpacingBefore(0)
    .setSpacingAfter(2);
  title.editAsText()
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor('#0F172A');

  const prettyDate = formatWideHandoutDate_(date);
  const countText = assignmentCount + ' assignment' + (assignmentCount === 1 ? '' : 's');
  const subtitle = body.appendParagraph(prettyDate + '  •  ' + countText);
  subtitle
    .setSpacingBefore(0)
    .setSpacingAfter(10);
  subtitle.editAsText()
    .setFontSize(9)
    .setForegroundColor('#475569');
}

function formatWideHandoutDate_(date) {
  const key = normalizeDateKey_(date);
  if (!key) return String(date || '');

  const parsed = new Date(key + 'T12:00:00');
  if (isNaN(parsed)) return key;

  return Utilities.formatDate(parsed, coverageTimeZone_(), 'EEEE, MMMM d, yyyy');
}

function formatWideHandoutTimeRange_(startValue, endValue) {
  const start = timeToDisplay_(startValue);
  const end = timeToDisplay_(endValue);
  if (!start && !end) return '';
  if (!start) return end;
  if (!end) return start;
  return start + '–' + end;
}

function styleWideHandoutTable_(table) {
  // Usable page width is 756 pt. These widths intentionally consume all of it.
  // The Time column is deliberately generous so normal ranges stay on one line.
  table.setColumnWidth(0, 150);
  table.setColumnWidth(1, 195);
  table.setColumnWidth(2, 315);
  table.setColumnWidth(3, 96);
  table.setBorderColor('#CBD5E1');
  table.setBorderWidth(0.75);

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    row.setMinimumHeight(r === 0 ? 24 : 26);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.setPaddingTop(3);
      cell.setPaddingBottom(3);
      cell.setPaddingLeft(6);
      cell.setPaddingRight(6);

      if (r === 0) {
        cell.setBackgroundColor('#E2E8F0');
      } else if (r % 2 === 0) {
        cell.setBackgroundColor('#F8FAFC');
      }

      for (let p = 0; p < cell.getNumChildren(); p++) {
        const child = cell.getChild(p);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          child.asParagraph()
            .setSpacingBefore(0)
            .setSpacingAfter(0)
            .setLineSpacing(1);
        }
      }

      const text = cell.editAsText();
      text.setFontSize(r === 0 ? 9 : 10);
      text.setForegroundColor(r === 0 ? '#334155' : '#0F172A');
      text.setBold(r === 0 || (r > 0 && c === 0));
    }
  }
}
