const FIELD_TRIP_HANDOUT_MODULE_VERSION = 2;

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

  let templateFile = null;
  let fieldTripContext = null;

  if (fieldTripEventIds.length) {
    templateFile = getFieldTripFormTemplateFile_();
    validateFieldTripFormTemplate_(DocumentApp.openById(templateFile.getId()).getBody());
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
      templateFile,
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

function getFieldTripFormTemplateFile_() {
  const templateId = String(coverageConfig_().Field_Trip_Form_Template_ID || '').trim();
  if (!templateId) {
    throw new Error(
      'Field trip handout template is not configured. Set Config → Field_Trip_Form_Template_ID to the Google Doc template ID.'
    );
  }

  let file;
  try {
    file = DriveApp.getFileById(templateId);
  } catch (error) {
    throw new Error('The configured field trip handout template could not be opened.');
  }

  if (file.isTrashed()) {
    throw new Error('The configured field trip handout template is in the trash.');
  }
  if (file.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('The field trip handout template must be a native Google Doc.');
  }

  return file;
}

function validateFieldTripFormTemplate_(body) {
  const tables = body.getTables();
  if (tables.length < 2) {
    throw new Error('The field trip handout template no longer contains both coverage tables.');
  }

  for (let i = 0; i < 2; i++) {
    if (tables[i].getNumRows() < 7 || tables[i].getRow(0).getNumCells() < 5) {
      throw new Error('The field trip handout template table structure has changed.');
    }
  }

  [
    'COVERAGE FOR:',
    'DATE:',
    'CLASS(ES) TAKING TRIP:',
    'TRIP DESTINATION:',
    'DEPARTURE TIME:',
    'APPROXIMATE RETURN TIME:'
  ].forEach(label => {
    if (fieldTripTemplateParagraphs_(body, label).length < 2) {
      throw new Error('The field trip handout template is missing the expected label: ' + label);
    }
  });
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

function createFieldTripCoverageFormDocs_(rows, trip, date, day, context, templateFile, folder) {
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

    // The template has six assignment rows per teacher. Keep the existing
    // safety split only for unusually long days, but all form units now live
    // in one output document instead of separate Part 1 / Part 2 files.
    for (let offset = 0; offset < teacherRows.length; offset += 6) {
      formUnits.push({
        coverageFor: teacher,
        rows: teacherRows.slice(offset, offset + 6)
      });
    }
  });

  if (!formUnits.length) return [];

  const outputName = 'Field Trip Coverage - ' + (trip.name || 'Field Trip') + ' - ' + date;
  const copy = templateFile.makeCopy(outputName, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const templateBody = DocumentApp.openById(templateFile.getId()).getBody();

  // The source template contains two form slots. Append additional copies of
  // that same template structure to this one document when more slots are
  // needed. No forced page breaks are inserted; after unused rows are trimmed,
  // Google Docs can naturally fit as many compact forms on a page as space
  // allows.
  const requiredTemplateCopies = Math.ceil(formUnits.length / 2);
  for (let copyIndex = 1; copyIndex < requiredTemplateCopies; copyIndex++) {
    appendFieldTripTemplateBodyCopy_(body, templateBody);
  }

  validateFieldTripFormTemplate_(body);

  formUnits.forEach((unit, formIndex) => {
    fillFieldTripForm_(body, formIndex, unit, trip, date, context);
  });

  // An odd number of form units leaves the final template's second form empty.
  // Remove that entire unused form rather than printing a blank half-page.
  if (formUnits.length % 2 === 1) {
    removeTrailingUnusedFieldTripForm_(body);
  }

  const remainingTables = body.getTables().length;
  if (remainingTables !== formUnits.length) {
    throw new Error(
      'Field trip form compaction failed: expected ' + formUnits.length +
      ' filled table(s), found ' + remainingTables + '.'
    );
  }

  doc.saveAndClose();

  return [{
    type: 'field-trip',
    label: 'Field Trip Form - ' + (trip.name || 'Field Trip'),
    eventId: trip.eventId || '',
    id: copy.getId(),
    url: copy.getUrl(),
    name: copy.getName(),
    folderId: folder.getId(),
    folderUrl: folder.getUrl(),
    folderName: folder.getName()
  }];
}

function appendFieldTripTemplateBodyCopy_(targetBody, sourceBody) {
  for (let i = 0; i < sourceBody.getNumChildren(); i++) {
    const child = sourceBody.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetBody.appendParagraph(child.copy().asParagraph());
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(child.copy().asTable());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetBody.appendListItem(child.copy().asListItem());
    } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
      targetBody.appendPageBreak();
    } else if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      targetBody.appendHorizontalRule();
    }
  }
}

function removeTrailingUnusedFieldTripForm_(body) {
  let startIndex = -1;

  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const text = String(child.asParagraph().getText() || '').trim();
    if (text.indexOf('FIELD TRIP COVERAGE FORM') !== -1) {
      startIndex = i;
      break;
    }
  }

  if (startIndex === -1) return;

  for (let i = body.getNumChildren() - 1; i >= startIndex; i--) {
    body.removeChild(body.getChild(i));
  }
}


function fillFieldTripForm_(body, formIndex, formUnit, trip, date, context) {
  insertFieldTripTemplateValue_(body, 'COVERAGE FOR:', formIndex, formUnit.coverageFor);
  insertFieldTripTemplateValue_(body, 'DATE:', formIndex, formatFieldTripFormDate_(date));
  insertFieldTripTemplateValue_(body, 'CLASS(ES) TAKING TRIP:', formIndex, formatFieldTripFormGrades_(trip.grades));
  insertFieldTripTemplateValue_(body, 'TRIP DESTINATION:', formIndex, fieldTripHandoutDestination_(trip));
  insertFieldTripTemplateValue_(body, 'DEPARTURE TIME:', formIndex, trip.start || '');
  insertFieldTripTemplateValue_(body, 'APPROXIMATE RETURN TIME:', formIndex, trip.end || '');
  compactFieldTripTimeLine_(body, formIndex);

  const table = body.getTables()[formIndex];
  const formRows = formUnit.rows || [];
  formRows.forEach((row, rowIndex) => {
    const withTeacher = resolveFieldTripWithTeacher_(row, trip, context);
    const values = [
      String(row.Assigned_Coverage || '').trim(),
      formatWideHandoutTimeRange_(row.Start, row.End),
      formatFieldTripRoomSubject_(row),
      String(trip.notes || '').trim(),
      withTeacher
    ];

    values.forEach((value, columnIndex) => {
      appendFieldTripTemplateCellText_(table.getCell(rowIndex + 1, columnIndex), value);
    });
  });

  // The source template reserves six rows. Keep only the rows this teacher
  // actually needs so the handout does not print large blank areas.
  while (table.getNumRows() > formRows.length + 1) {
    table.removeRow(table.getNumRows() - 1);
  }

  if (table.getNumRows() !== formRows.length + 1) {
    throw new Error('Field trip form could not remove unused assignment rows.');
  }
}

function compactFieldTripTimeLine_(body, occurrence) {
  const paragraphs = body.getParagraphs().filter(paragraph => {
    const text = String(paragraph.getText() || '');
    return text.indexOf('DEPARTURE TIME:') !== -1 &&
      text.indexOf('APPROXIMATE RETURN TIME:') !== -1;
  });

  if (paragraphs.length <= occurrence) return;

  // The source template uses a long tab run between the two time fields.
  // Once values are inserted, that run can wrap the return time to a second
  // line. Shorten only that spacer in the generated copy.
  paragraphs[occurrence].editAsText().replaceText('\\t{4,}', '\t\t\t');
}

function fieldTripHandoutDestination_(trip) {
  const destination = String(trip && trip.destination || '').trim();
  if (destination) return destination;

  // Older trip records sometimes put the destination in the event name
  // ("6th Grade - National Gallery") and leave Destination blank.
  const name = String(trip && trip.name || '').trim();
  const parts = name.split(/\s+[—–-]\s+/).map(part => part.trim()).filter(Boolean);
  return parts.length > 1 ? parts[parts.length - 1] : '';
}

function fieldTripTemplateParagraphs_(body, label) {
  return body.getParagraphs().filter(paragraph =>
    String(paragraph.getText() || '').indexOf(label) !== -1
  );
}

function insertFieldTripTemplateValue_(body, label, occurrence, value) {
  const textValue = String(value == null ? '' : value).trim();
  if (!textValue) return;

  const paragraphs = fieldTripTemplateParagraphs_(body, label);
  if (paragraphs.length <= occurrence) {
    throw new Error('The field trip handout template is missing ' + label);
  }

  const text = paragraphs[occurrence].editAsText();
  const current = text.getText();
  const labelIndex = current.indexOf(label);
  if (labelIndex === -1) {
    throw new Error('The field trip handout template is missing ' + label);
  }

  text.insertText(labelIndex + label.length, ' ' + textValue);
}

function appendFieldTripTemplateCellText_(cell, value) {
  const textValue = String(value == null ? '' : value).trim();
  if (!textValue) return;

  let paragraph = null;
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      paragraph = child.asParagraph();
      break;
    }
  }

  if (!paragraph) {
    paragraph = cell.appendParagraph('');
  }

  paragraph.editAsText().setText(textValue);

  // Template cells sometimes contain extra empty paragraphs. They create
  // visible blank lines in exported DOCX/PDF, so remove them after setting
  // the actual value.
  for (let i = cell.getNumChildren() - 1; i >= 0; i--) {
    const child = cell.getChild(i);
    if (child === paragraph) continue;
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH &&
        !String(child.asParagraph().getText() || '').trim()) {
      cell.removeChild(child);
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
