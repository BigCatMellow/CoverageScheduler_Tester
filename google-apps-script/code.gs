const COVERAGE_WEB_API_VERSION = 12;

const APP_TITLE = 'Coverage Scheduler';
const COVERAGE_SPREADSHEET_PROPERTY = 'COVERAGE_SPREADSHEET_ID';
const COVERAGE_SCHEMA_VERSION_PROPERTY = 'COVERAGE_SCHEMA_VERSION';
const COVERAGE_SCHEMA_VERSION = '2026-10-08-api12-hardening-1';
let COVERAGE_WEB_READY_THIS_REQUEST_ = false;

function onOpen() {
  rememberCoverageSpreadsheet_();

  SpreadsheetApp.getUi()
    .createMenu(APP_TITLE)
    .addItem('Set up workbook', 'setupCoverageScheduler')
    .addItem('Validate teacher schedule', 'menuValidateTeacherScheduleSource')
    .addItem('Rebuild schedule day caches', 'menuRebuildTeacherScheduleDayCaches')
    .addSeparator()
    .addItem('Generate preview for selected day', 'generateCoveragePreviewFromPrompt')
    .addItem('Rebuild field trip coverage pool', 'menuRebuildFieldTripCoveragePool')
    .addItem('Create handout doc from latest preview', 'menuCreateHandoutDoc')
    .addItem('Clear Coverage Output', 'clearCoverageOutput')
    .addToUi();
}

// Manual edits to source/config sheets must be visible immediately. Teacher
// Schedule hot paths use hidden per-day caches; edits to either schedule source
// invalidate those caches so the next requested weekday rebuilds exactly once.
function onEdit(e) {
  try {
    const sheet = e && e.range ? e.range.getSheet() : null;
    if (sheet) {
      const sheetName = sheet.getName();
      invalidateCoverageSheetCache_(sheetName);
      if (['Teacher Schedule', 'Class Schedule', 'Copy of Class Schedule'].indexOf(sheetName) !== -1 &&
          typeof markTeacherScheduleCacheDirty_ === 'function') {
        markTeacherScheduleCacheDirty_();
      }
      if (['Teacher Schedule', 'Class Schedule', 'Copy of Class Schedule', 'Field Trips', 'Config', 'Coverage Staff', 'Substitute Availability', 'Substitutes'].indexOf(sheetName) !== -1) {
        markFieldTripCoveragePoolDirty_();
      }
    }
  } catch (error) {
    // Cache invalidation is best-effort; an edit must never be blocked by it.
  }
}

function rememberCoverageSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  PropertiesService.getScriptProperties()
    .setProperty(COVERAGE_SPREADSHEET_PROPERTY, ss.getId());

  return ss;
}

function setupCoverageScheduler() {
  const ss = rememberCoverageSpreadsheet_();
  if (!ss) {
    throw new Error('Open the Google Sheet that will hold Coverage Scheduler, then run setup again from Extensions → Apps Script.');
  }

  SpreadsheetApp.setActiveSpreadsheet(ss);
  const result = setupCoverageWorkbookFromTeacherSchedule();
  markCoverageSchemaReady_();
  return result;
}

function doGet() {
  const output = HtmlService.createTemplateFromFile('index').evaluate();

  output.append(getSchoolThemeCss_());
  output.append(getWorkingOverlayUi_());
  output.append(getAbsenceRangeUi_());
  output.append(getFieldTripUi_());
  output.append(getHandoutOpenLinkUi_());

  return output
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSchoolThemeCss_() {
  return [
    '<style>',
    ':root{',
    '  --primary:#214289;',
    '  --primary-hover:#1a356c;',
    '  --accent:#f6b756;',
    '  --accent-light:rgba(246,183,86,.12);',
    '  --bg:#f0f2f7;',
    '  --surface:#ffffff;',
    '  --border:#e2e5eb;',
    '  --text:#1f2937;',
    '  --text-muted:#6b7280;',
    '  --success:#059669;',
    '  --error:#dc2626;',
    '  --called-color:#f59e0b;',
    '  --dismissed-color:#059669;',
    '  --returned-color:#6b7280;',
    '  --radius:14px;',
    '  --radius-sm:10px;',
    '  --shadow-card:0 2px 5px rgba(0,0,0,.06),0 1px 3px rgba(0,0,0,.04);',
    '  --shadow-lifted:0 10px 30px rgba(0,0,0,.12),0 3px 8px rgba(0,0,0,.06);',
    '  --shadow-sheet:0 -10px 35px rgba(0,0,0,.18);',
    '  --bottom-nav-h:80px;',
    '  --header-h:72px;',
    '  --safe-bottom:env(safe-area-inset-bottom,0px);',
    '  --ink:var(--text);',
    '  --muted:var(--text-muted);',
    '  --line:var(--border);',
    '  --panel:var(--surface);',
    '  --soft:#f7f8fb;',
    '  --indigo:var(--primary);',
    '  --indigo-dark:var(--primary-hover);',
    '  --green:var(--success);',
    '  --red:var(--error);',
    '  --amber:var(--called-color);',
    '}',
    'body{background:var(--bg);color:var(--text)}',
    '.topbar{background:var(--primary);border-bottom:3px solid var(--accent);box-shadow:0 2px 8px rgba(33,66,137,.16)}',
    '.brand-icon{background:var(--accent);color:var(--primary);border-radius:var(--radius-sm)}',
    '.day-tag{background:var(--accent-light);color:#ffe0aa;border:1px solid rgba(246,183,86,.28)}',
    '.btn-primary{background:var(--accent);color:var(--primary)}',
    '.btn-primary:hover{background:#efaa3f;color:var(--primary-hover)}',
    '.btn-ghost{border-color:rgba(255,255,255,.22)}',
    '.panel-left,.panel-right,.pnl-hd,.plan-hd,.pnl-search{background:var(--surface)}',
    '.panel-center{background:var(--bg)}',
    '.panel-left{border-right-color:var(--border)}',
    '.panel-right{border-left-color:var(--border)}',
    '.pnl-hd,.plan-hd,.pnl-search{border-bottom-color:var(--border)}',
    '.pnl-hd h2,.plan-hd h2,.fl,.section-title{color:var(--text-muted)}',
    '.pnl-count{background:#edf0f5;color:var(--text-muted)}',
    '.mini-add{background:var(--accent-light);color:var(--primary);border:1px solid rgba(246,183,86,.4)}',
    '.mini-add:hover{background:rgba(246,183,86,.22)}',
    '.link-btn{color:var(--primary)}',
    '.link-btn:hover{color:var(--primary-hover)}',
    '.add-btn{border-color:#c8ced8;color:var(--text-muted);border-radius:var(--radius-sm)}',
    '.add-btn:hover{border-color:var(--primary);color:var(--primary);background:rgba(33,66,137,.045)}',
    '.abs-row:hover,.staff-row:hover,.plan-tbl tbody tr:hover td{background:#f7f8fb}',
    '.t1{background:rgba(33,66,137,.10);color:var(--primary)}',
    '.t2{background:var(--accent-light);color:#9a650e}',
    '.t3{background:#eef0f4;color:var(--returned-color)}',
    '.vtabs{background:#e7eaf0;border-radius:var(--radius-sm)}',
    '.vtab{color:var(--text-muted)}',
    '.vtab.on{background:var(--surface);color:var(--primary);box-shadow:var(--shadow-card)}',
    '.spinner{border-color:#d9dee7;border-top-color:var(--primary)}',
    '.plan-tbl,.sub-card{background:var(--surface);border-radius:var(--radius-sm);box-shadow:var(--shadow-card)}',
    '.plan-tbl th,.sub-card-hd{background:#f7f8fb;border-color:var(--border)}',
    '.plan-tbl th{color:#566171}',
    '.plan-tbl tbody tr:nth-child(odd) td{background:#ffffff}',
    '.plan-tbl tbody tr:nth-child(even) td{background:#f4f6fa}',
    '.plan-tbl tbody tr.uf td{background:#fff5f5}',
    '.tl-row{border-radius:7px;padding:2px 4px}',
    '.tl-group .tl-row:nth-child(odd){background:rgba(255,255,255,.55)}',
    '.tl-group .tl-row:nth-child(even){background:rgba(33,66,137,.045)}',
    '.tl-group .tl-row:nth-child(odd) .tl-track{background:#edf0f5}',
    '.tl-group .tl-row:nth-child(even) .tl-track{background:#e3e8f0}',
    '.modal{background:var(--surface);border-radius:var(--radius);box-shadow:var(--shadow-lifted)}',
    '.m-hd{background:var(--surface);border-bottom-color:var(--border)}',
    '.m-ft{background:#f7f8fb;border-top-color:var(--border)}',
    '.fi,.fs,.fta,.t-opt,.day-check span,.btn-cancel{border-color:#d7dce5;border-radius:var(--radius-sm)}',
    '.fi:focus,.fs:focus,.fta:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(33,66,137,.10)}',
    '.t-opt.sel{border-color:var(--primary);background:rgba(33,66,137,.08);color:var(--primary)}',
    '.checkline input{accent-color:var(--primary)}',
    '.day-check input:checked+span{background:rgba(33,66,137,.08);border-color:var(--primary);color:var(--primary)}',
    '.advanced{border-color:var(--border);border-radius:var(--radius-sm)}',
    '.advanced summary{background:#f7f8fb;color:#566171;border-radius:var(--radius-sm)}',
    '.advanced[open] summary{border-bottom-color:var(--border)}',
    '.btn-save{background:var(--primary);color:#fff;border-radius:var(--radius-sm)}',
    '.btn-save:hover{background:var(--primary-hover)}',
    '.btn-cancel,.btn-del{border-radius:var(--radius-sm)}',
    '.block-info{background:#f7f8fb;border-radius:var(--radius-sm)}',
    '.error-box{border-radius:var(--radius);box-shadow:var(--shadow-card)}',
    '</style>'
  ].join('\n');
}

function getWorkingOverlayUi_() {
  return [
    '<style>',
    '.working-overlay{position:fixed;inset:0;z-index:95;display:grid;place-items:center;padding:20px;background:rgba(15,23,42,.34);backdrop-filter:blur(2px)}',
    '.working-overlay.hidden{display:none!important}',
    '.working-card{display:flex;align-items:center;gap:14px;min-width:280px;max-width:92vw;padding:20px 22px;background:var(--surface,#fff);border:1px solid var(--border,#e2e5eb);border-radius:var(--radius,14px);box-shadow:var(--shadow-lifted,0 10px 30px rgba(0,0,0,.12))}',
    '.working-spinner{width:32px;height:32px;flex:0 0 32px;border:3px solid #dce2eb;border-top-color:var(--primary,#214289);border-radius:50%;animation:working-spin .72s linear infinite}',
    '.working-title{font-size:13px;font-weight:800;color:var(--text,#1f2937);line-height:1.25}',
    '.working-sub{margin-top:4px;font-size:10px;color:var(--text-muted,#6b7280);line-height:1.35}',
    '@keyframes working-spin{to{transform:rotate(360deg)}}',
    '</style>',
    '<script>',
    '(function(){',
    '  var workingCount=0;',
    '  var messages={',
    "    webGetBootstrap:['Loading date…','Loading absences, coverage staff, and the saved plan.'],",
    "    webGenerateCoverage:['Generating coverage plan…','Matching schedules and available coverage staff.'],",
    "    webCreateHandout:['Building handout…','Creating and formatting the Google Doc.'],",
    "    webCreateHandoutFromRows:['Building handout…','Creating the Google Doc from the reviewed plan.'],",
    "    webSaveCoverage:['Saving coverage plan…','Writing the plan to Coverage Output.'],",
    "    webSaveCoverageAndCreateHandout:['Saving plan & building handout…','Writing Coverage Output and creating the Google Doc from the same plan.'],",
    "    webSaveAbsences:['Saving absences…','Updating the selected day.'],",
    "    webSaveAbsenceRange:['Saving absence dates…','Adding the absence to each selected school day.'],",
    "    webSaveFieldTrip:['Saving field trip…','Updating the event and recalculating its scheduling context.'],",
    "    webDeleteFieldTrip:['Removing field trip…','Deleting the event from the calendar.'],",
    "    webSaveCoverageStaff:['Saving coverage staff…','Updating the coverage team.'],",
    "    webDeleteCoverageStaff:['Removing coverage staff…','Updating the coverage team.'],",
    "    webToggleCoverageStaff:['Updating availability…','Saving the daily availability change.'],",
    "    webValidateTeacherSchedule:['Checking teacher schedule…','Validating the schedule source.']",
    '  };',
    '  function ensureOverlay(){',
    "    var el=document.getElementById('workingOverlay');",
    '    if(el)return el;',
    "    el=document.createElement('div');",
    "    el.id='workingOverlay';",
    "    el.className='working-overlay hidden';",
    "    el.setAttribute('role','status');",
    "    el.setAttribute('aria-live','polite');",
    "    el.innerHTML='<div class=\"working-card\"><div class=\"working-spinner\" aria-hidden=\"true\"></div><div><div id=\"workingTitle\" class=\"working-title\">Working…</div><div id=\"workingSub\" class=\"working-sub\">This can take a few seconds.</div></div></div>';",
    '    document.body.appendChild(el);',
    '    return el;',
    '  }',
    '  window.showWorking=function(title,sub){',
    '    workingCount++;',
    '    var el=ensureOverlay();',
    "    document.getElementById('workingTitle').textContent=title||'Working…';",
    "    document.getElementById('workingSub').textContent=sub||'This can take a few seconds.';",
    "    el.classList.remove('hidden');",
    '  };',
    '  window.hideWorking=function(){',
    '    workingCount=Math.max(0,workingCount-1);',
    '    if(workingCount===0){',
    "      var el=document.getElementById('workingOverlay');",
    "      if(el)el.classList.add('hidden');",
    '    }',
    '  };',
    "  if(typeof gas==='function'){",
    '    var baseGas=gas;',
    '    gas=function(method,payload){',
    '      var message=messages[method];',
    '      if(!message)return baseGas(method,payload);',
    '      window.showWorking(message[0],message[1]);',
    '      return baseGas(method,payload).then(function(result){',
    '        window.hideWorking();',
    '        return result;',
    '      },function(error){',
    '        window.hideWorking();',
    '        throw error;',
    '      });',
    '    };',
    '  }',
    "  if(typeof setLoading==='function' && typeof renderPlan==='function'){",
    '    var baseSetLoading=setLoading;',
    '    setLoading=function(on){',
    '      baseSetLoading(on);',
    '      if(!on){',
    "        var body=document.getElementById('planBody');",
    "        if(body && body.querySelector('.spinner')) renderPlan();",
    '      }',
    '    };',
    '  }',
    '})();',
    '</script>'
  ].join('\n');
}

function getAbsenceRangeUi_() {
  return [
    '<style>',
    '.absence-range-box{padding:11px;background:#f7f8fb;border:1px solid var(--border,#e2e5eb);border-radius:var(--radius-sm,10px)}',
    '.absence-range-box .fg{margin-bottom:0}',
    '.absence-range-presets{display:flex;gap:6px;margin-top:8px}',
    '.absence-range-presets button{border:1px solid #d7dce5;background:#fff;color:var(--primary,#214289);border-radius:7px;padding:5px 8px;font-size:10px;font-weight:750}',
    '.absence-range-presets button:hover{border-color:var(--primary,#214289);background:rgba(33,66,137,.045)}',
    '.absence-range-error{color:var(--error,#dc2626)!important}',
    '.add-row{display:flex;gap:7px;flex-wrap:wrap}',
    '.add-row .add-btn{width:auto;flex:1 1 120px}',
    '</style>',
    '<script>',
    '(function(){',
    "  var staff=document.getElementById('absenceStaff');",
    '  if(!staff)return;',
    "  var staffGroup=staff.closest('.fg');",
    "  var range=document.createElement('div');",
    "  range.id='absenceDateRange';",
    "  range.className='fg absence-range-box';",
    "  range.innerHTML='<label class=\"fl\">Dates</label><div class=\"f-row\"><div class=\"fg\"><label class=\"fl\" style=\"text-transform:none;letter-spacing:0\">From</label><input id=\"absenceDateFrom\" type=\"date\" class=\"fi\"></div><div class=\"fg\"><label class=\"fl\" style=\"text-transform:none;letter-spacing:0\">Through</label><input id=\"absenceDateThrough\" type=\"date\" class=\"fi\"></div></div><div class=\"absence-range-presets\"><button type=\"button\" id=\"absenceOneDay\">This day only</button><button type=\"button\" id=\"absenceFiveDays\">5 school days</button></div><div id=\"absenceDateHint\" class=\"hint\"></div>';",
    '  staffGroup.parentNode.insertBefore(range,staffGroup.nextSibling);',
    '',
    '  function parseDate(value){',
    "    var p=String(value||'').split('-').map(Number);",
    '    if(p.length!==3||!p[0]||!p[1]||!p[2])return null;',
    '    return new Date(p[0],p[1]-1,p[2],12,0,0,0);',
    '  }',
    '  function dateKey(d){',
    "    return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');",
    '  }',
    '  function schoolDayCount(from,to){',
    '    var a=parseDate(from),b=parseDate(to);',
    '    if(!a||!b||b<a)return 0;',
    '    var count=0,guard=0;',
    '    while(a<=b&&guard<70){var day=a.getDay();if(day!==0&&day!==6)count++;a.setDate(a.getDate()+1);guard++;}',
    '    return count;',
    '  }',
    '  function fiveSchoolDayEnd(start){',
    '    var d=parseDate(start);if(!d)return start;',
    '    var count=0,guard=0;',
    '    while(count<5&&guard<14){var day=d.getDay();if(day!==0&&day!==6)count++;if(count<5)d.setDate(d.getDate()+1);guard++;}',
    '    return dateKey(d);',
    '  }',
    '  function updateRangeHint(){',
    "    var from=document.getElementById('absenceDateFrom').value;",
    "    var through=document.getElementById('absenceDateThrough').value;",
    "    var hint=document.getElementById('absenceDateHint');",
    "    hint.classList.remove('absence-range-error');",
    "    if(!from||!through){hint.textContent='Choose the date or range. Weekends are skipped automatically.';return;}",
    "    if(through<from){hint.textContent='Through date must be the same as or after From.';hint.classList.add('absence-range-error');return;}",
    '    var count=schoolDayCount(from,through);',
    "    hint.textContent=count+' school day'+(count===1?'':'s')+' selected. Weekends are skipped automatically.';",
    '  }',
    "  document.getElementById('absenceDateFrom').addEventListener('change',function(){",
    "    var through=document.getElementById('absenceDateThrough');",
    '    if(!through.value||through.value<this.value)through.value=this.value;',
    '    updateRangeHint();',
    "    if(typeof updateAbsenceScheduleHint==='function')updateAbsenceScheduleHint();",
    '  });',
    "  document.getElementById('absenceDateThrough').addEventListener('change',updateRangeHint);",
    "  document.getElementById('absenceOneDay').addEventListener('click',function(){var from=document.getElementById('absenceDateFrom');var through=document.getElementById('absenceDateThrough');through.value=from.value||S.date;updateRangeHint();});",
    "  document.getElementById('absenceFiveDays').addEventListener('click',function(){var from=document.getElementById('absenceDateFrom');var through=document.getElementById('absenceDateThrough');if(!from.value)from.value=S.date;through.value=fiveSchoolDayEnd(from.value);updateRangeHint();});",
    '',
    '  var baseUpdateAbsenceScheduleHint=updateAbsenceScheduleHint;',
    '  updateAbsenceScheduleHint=function(){',
    "    var from=document.getElementById('absenceDateFrom');",
    '    if(!S.editingAbsence&&from&&from.value&&from.value!==S.date){',
    "      document.getElementById('absenceScheduleHint').textContent='Coverage blocks will use the Teacher Schedule for each selected date.';",
    '      return;',
    '    }',
    '    return baseUpdateAbsenceScheduleHint();',
    '  };',
    '',
    '  var baseOpenAbsence=openAbsence;',
    '  openAbsence=function(index){',
    '    baseOpenAbsence(index);',
    '    var editing=!!S.editingAbsence;',
    "    range.classList.toggle('hidden',editing);",
    "    var sub=document.querySelector('#absenceModal .m-sub');",
    "    if(sub)sub.textContent=editing?'Edit this absence for the selected day.':'Choose a staff member, date or range, and the part of the day they will miss.';",
    '    if(!editing){',
    "      document.getElementById('absenceDateFrom').value=S.date;",
    "      document.getElementById('absenceDateThrough').value=S.date;",
    '      updateRangeHint();',
    '      updateAbsenceScheduleHint();',
    '    }',
    '  };',
    '',
    '  var baseSaveAbsence=saveAbsence;',
    '  saveAbsence=async function(){',
    '    if(S.editingAbsence)return baseSaveAbsence();',
    "    var staffName=document.getElementById('absenceStaff').value;",
    "    if(!staffName){flash('Choose a staff member.','warn');return;}",
    "    var type=document.getElementById('absenceTypes').dataset.selected||'Full Day';",
    "    var next={staffName:staffName,absenceType:type==='Emergency'?'Full Day':type,allDay:type!=='Partial Day',emergency:type==='Emergency',startOverride:type==='Partial Day'?document.getElementById('absenceStart').value.trim():'',endOverride:type==='Partial Day'?document.getElementById('absenceEnd').value.trim():'',preferredCoverage:document.getElementById('preferredCoverage').value,notes:document.getElementById('absenceNotes').value.trim()};",
    "    if(type==='Partial Day'&&(!next.startOverride||!next.endOverride)){flash('Enter both a start and end time for a partial-day absence.','warn');return;}",
    "    var startDate=document.getElementById('absenceDateFrom').value;",
    "    var endDate=document.getElementById('absenceDateThrough').value||startDate;",
    "    if(!startDate||!endDate){flash('Choose the absence date.','warn');return;}",
    "    if(endDate<startDate){flash('Through date must be the same as or after From.','warn');return;}",
    '    try{',
    "      var result=await gas('webSaveAbsenceRange',{startDate:startDate,endDate:endDate,currentDate:S.date,absence:next});",
    "      closeModal('absenceModal');",
    '      if(result&&result.currentDateUpdated){',
    '        S.absences=normalizeAbsences(result.currentAbsences||[]);',
    '        S.plan=[];',
    '        renderAll();',
    '      }',
    "      var count=(result&&result.count)||schoolDayCount(startDate,endDate);",
    "      flash('Absence saved for '+count+' school day'+(count===1?'':'s')+'.');",
    '    }catch(e){fail(e)}',
    '  };',
    '})();',
    '</script>'
  ].join('\n');
}
function getHandoutOpenLinkUi_() {
  return [
    '<style>',
    '.notif .handout-open{display:inline-flex;align-items:center;margin-left:6px;padding:4px 8px;border-radius:6px;background:#166534;color:#fff;text-decoration:none;font-size:11px;font-weight:800;white-space:nowrap}',
    '.notif .handout-open:hover{background:#14532d}',
    '.notif .handout-folder-open{background:#0f766e}',
    '.notif .handout-folder-open:hover{background:#115e59}',
    '</style>',
    '<script>',
    'createHandout = async function(){',
    '  try{',
    "    const handout=await gas('webCreateHandoutFromRows',S.plan);",
    "    const n=$('notif');",
    "    const docs=handout&&Array.isArray(handout.documents)&&handout.documents.length?handout.documents:[handout];",
    "    const folderUrl=handout&&handout.folderUrl?handout.folderUrl:'';",
    "    const links=docs.filter(function(doc){return doc&&doc.url;}).map(function(doc){var label=doc.label||doc.name||'Open Handout';return ' <a class=\"handout-open\" href=\"'+esc(doc.url)+'\" target=\"_blank\" rel=\"noopener noreferrer\">'+esc(label)+' ↗</a>';}).join('');",
    "    n.className='notif ok';",
    "    n.innerHTML='✓ Created '+docs.length+' handout document'+(docs.length===1?'':'s')+' in the Handouts folder.'+links+(folderUrl?' <a class=\"handout-open handout-folder-open\" href=\"'+esc(folderUrl)+'\" target=\"_blank\" rel=\"noopener noreferrer\">Open Handouts Folder ↗</a>':'')+'<button onclick=\"this.parentElement.classList.add(\\'hidden\\')\">×</button>';",
    "    n.classList.remove('hidden');",
    "    clearTimeout(flash.t);",
    '  }catch(e){fail(e)}',
    '};',
    '</script>'
  ].join('\n');
}

function activateCoverageSpreadsheetForWeb_() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(COVERAGE_SPREADSHEET_PROPERTY);

  if (!spreadsheetId) {
    throw new Error(
      'Coverage Scheduler has not been connected to a workbook yet. Reload the Google Sheet that contains this Apps Script project, then reopen the web app.'
    );
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  SpreadsheetApp.setActiveSpreadsheet(ss);
  return ss;
}

function coverageSchemaMarker_(ss) {
  const workbook = ss || SpreadsheetApp.getActiveSpreadsheet();
  return COVERAGE_SCHEMA_VERSION + ':' + (workbook ? workbook.getId() : '');
}

function markCoverageSchemaReady_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties()
    .setProperty(COVERAGE_SCHEMA_VERSION_PROPERTY, coverageSchemaMarker_(ss));
}

function ensureCoverageWorkbookReadyForWeb_() {
  if (COVERAGE_WEB_READY_THIS_REQUEST_) return SpreadsheetApp.getActiveSpreadsheet();

  const ss = activateCoverageSpreadsheetForWeb_();
  const properties = PropertiesService.getScriptProperties();
  const schemaVersion = properties.getProperty(COVERAGE_SCHEMA_VERSION_PROPERTY);

  // Existing installations run the full structural check once after this
  // optimization is deployed. Normal requests then use the version marker
  // instead of probing every managed sheet on every click.
  if (schemaVersion !== coverageSchemaMarker_(ss)) {
    const requiredSheets = [
      'Teacher Schedule',
      'Coverage Staff',
      'Substitute Availability',
      'Daily Absences',
      'Field Trip Coverage Pool',
      'Coverage Output',
      'Lists',
      'Config',
      '_Preview'
    ];
    const missing = requiredSheets.filter(name => !ss.getSheetByName(name));
    if (missing.length) setupCoverageWorkbookFromTeacherSchedule();
    ensureStaffListSheet_();
    markFieldTripCoveragePoolDirty_();
    markCoverageSchemaReady_();
  } else if (!ss.getSheetByName('Teacher Schedule')) {
    // Fail clearly if the one authoritative source tab was removed after setup.
    throw new Error('Coverage Scheduler cannot find the Teacher Schedule sheet. Run Coverage Scheduler → Set up workbook.');
  }

  COVERAGE_WEB_READY_THIS_REQUEST_ = true;
  return ss;
}

function runCoverageWebRequest_(name, fn) {
  beginCoverageRequest_(name);
  COVERAGE_WEB_READY_THIS_REQUEST_ = false;
  try {
    ensureCoverageWorkbookReadyForWeb_();
    coveragePerfMark_('workbook-ready');
    return fn();
  } finally {
    endCoverageRequest_();
    COVERAGE_WEB_READY_THIS_REQUEST_ = false;
  }
}

function makeWebSafe_(value) {
  if (value === null || value === undefined) return value === undefined ? null : value;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    const timeZone = coverageTimeZone_();
    const year = Number(Utilities.formatDate(value, timeZone, 'yyyy'));
    return year <= 1900
      ? Utilities.formatDate(value, timeZone, 'h:mm a')
      : Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }

  if (Array.isArray(value)) {
    return value.map(item => makeWebSafe_(item));
  }

  if (typeof value === 'number') {
    return isFinite(value) ? value : '';
  }

  if (typeof value === 'string' || typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(key => {
      out[key] = makeWebSafe_(value[key]);
    });
    return out;
  }

  return String(value);
}

function webGetBootstrap(payload) {
  return runCoverageWebRequest_('webGetBootstrap', () => {
    if (typeof FIELD_TRIP_HANDOUT_MODULE_VERSION === 'undefined' ||
        FIELD_TRIP_HANDOUT_MODULE_VERSION !== 2) {
      throw new Error(
        'Coverage Scheduler deployment is out of sync: handout.gs is not the current compact field-trip form module. Deploy every file in google-apps-script together as one new version.'
      );
    }

    const data = getCoverageBootstrap_(payload || {});
    if (!data) throw new Error('Coverage Scheduler could not build its startup data.');

    coveragePerfMark_('bootstrap-built');
    const safe = makeWebSafe_(data);
    if (!safe || !safe.today) {
      throw new Error('Coverage Scheduler startup data was invalid before it reached the browser.');
    }

    safe.webApiVersion = COVERAGE_WEB_API_VERSION;
    safe.webCapabilities = {
      manualCoveragePlacement: true,
      unifiedFieldTripPlan: true,
      fieldTripContinuityScheduling: true
    };
    return safe;
  });
}

function webGetStaffSchedule(payload) {
  return runCoverageWebRequest_('webGetStaffSchedule', () => {
    payload = payload || {};
    return makeWebSafe_(getWebStaffSchedule_(
      payload.staffName,
      payload.date,
      payload.day
    ));
  });
}

function webSaveFieldTrip(payload) {
  return runCoverageWebRequest_('webSaveFieldTrip', () => makeWebSafe_(saveFieldTrip_(payload || {})));
}

function webDeleteFieldTrip(payload) {
  return runCoverageWebRequest_('webDeleteFieldTrip', () => {
    payload = payload || {};
    return makeWebSafe_(deleteFieldTrip_(payload.eventId));
  });
}

function webGetCalendarData(payload) {
  return runCoverageWebRequest_('webGetCalendarData', () => {
    payload = payload || {};
    const startDate = normalizeDateKey_(payload.startDate);
    const endDate = normalizeDateKey_(payload.endDate || payload.startDate);

    if (!startDate || !endDate) throw new Error('Calendar requires a valid date range.');
    if (endDate < startDate) throw new Error('Calendar end date must be after its start date.');

    primeCoverageRequestSnapshot_(['Field Trips', 'Daily Absences']);
    const fieldTrips = getFieldTripsInRange_(startDate, endDate);
    const absences = readSheetObjects_('Daily Absences')
      .map(row => ({
        date: normalizeDateKey_(row.Date),
        staffName: String(row.Staff_Name || '').trim(),
        absenceType: String(row.Absence_Type || 'Full Day').trim(),
        start: timeToDisplay_(row.Start_Override),
        end: timeToDisplay_(row.End_Override),
        notes: String(row.Notes || '').trim()
      }))
      .filter(row => row.date && row.date >= startDate && row.date <= endDate && row.staffName);

    return makeWebSafe_({ startDate: startDate, endDate: endDate, fieldTrips: fieldTrips, absences: absences });
  });
}

function webSaveAbsences(payload) {
  return runCoverageWebRequest_('webSaveAbsences', () => replaceDailyAbsences(payload || {}));
}

function parseAbsenceRangeDate_(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0, 0);
  if (
    date.getFullYear() !== Number(match[1]) ||
    date.getMonth() !== Number(match[2]) - 1 ||
    date.getDate() !== Number(match[3])
  ) return null;
  return date;
}

function webSaveAbsenceRange(payload) {
  return runCoverageWebRequest_('webSaveAbsenceRange', () =>
    withCoverageLock_(() => webSaveAbsenceRangeUnlocked_(payload))
  );
}

function webSaveAbsenceRangeUnlocked_(payload) {
  payload = payload || {};

  const start = parseAbsenceRangeDate_(payload.startDate);
  const end = parseAbsenceRangeDate_(payload.endDate || payload.startDate);
  const absence = payload.absence || {};
  const staffName = String(absence.staffName || '').trim();

  if (!start || !end) throw new Error('Choose a valid absence date or date range.');
  if (end < start) throw new Error('Through date must be the same as or after From.');
  if (!staffName) throw new Error('Choose a staff member.');

  const calendarDays = Math.floor((end.getTime() - start.getTime()) / 86400000) + 1;
  if (calendarDays > 63) throw new Error('Please keep one absence entry to 63 calendar days or fewer.');

  const timeZone = coverageTimeZone_();
  const savedDates = [];
  const rows = readSheetObjects_('Daily Absences').slice();
  let changed = false;
  const cursor = new Date(start.getTime());

  while (cursor <= end) {
    const weekday = cursor.getDay();
    if (weekday !== 0 && weekday !== 6) {
      const dateKey = Utilities.formatDate(cursor, timeZone, 'yyyy-MM-dd');
      const day = guessDayCodeFromDate_(dateKey);
      if (day) {
        const duplicate = rows.some(row => {
          if (normalizeDateKey_(row.Date) !== dateKey || String(row.Day || '').trim() !== day) return false;
          return sameAbsenceWindow_({
            staffName: String(row.Staff_Name || '').trim(),
            absenceType: String(row.Absence_Type || 'Full Day').trim(),
            startOverride: timeToDisplay_(row.Start_Override),
            endOverride: timeToDisplay_(row.End_Override),
            emergency: isEmergencyAbsence_(row)
          }, absence);
        });

        if (!duplicate) {
          const allDay = normalizeYesNo_(absence.allDay, false);
          const emergency = normalizeYesNo_(absence.emergency, false);
          rows.push({
            Date: dateKey,
            Day: day,
            Staff_Name: staffName,
            Absence_Type: allDay ? 'Full Day' : 'Partial Day',
            Start_Override: allDay ? '' : String(absence.startOverride || '').trim(),
            End_Override: allDay ? '' : String(absence.endOverride || '').trim(),
            Notes: composeAbsenceNotes_(emergency, absence.notes),
            Preferred_Coverage: String(absence.preferredCoverage || '').trim()
          });
          changed = true;
        }
        savedDates.push(dateKey);
      }
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  if (!savedDates.length) throw new Error('That range does not contain a Monday–Friday school day.');

  if (changed) {
    rewriteSheetRows_('Daily Absences', SHEET_SCHEMAS['Daily Absences'].headers, rows);
    invalidatePreviewForDateRange_(savedDates[0], savedDates[savedDates.length - 1]);
  }
  coveragePerfMark_('absence-range-written');

  const currentDate = String(payload.currentDate || '').trim();
  const currentDateUpdated = savedDates.indexOf(currentDate) !== -1;
  const currentDay = currentDateUpdated ? guessDayCodeFromDate_(currentDate) : '';

  return makeWebSafe_({
    count: savedDates.length,
    savedDates: savedDates,
    currentDateUpdated: currentDateUpdated,
    currentAbsences: currentDateUpdated ? getDailyAbsencesForDate_(currentDate, currentDay) : []
  });
}

function webGenerateCoverage(payload) {
  return runCoverageWebRequest_('webGenerateCoverage', () => generateCoveragePreview(payload || {}));
}

function webGetManualCoverageChoices(payload) {
  return runCoverageWebRequest_('webGetManualCoverageChoices', () =>
    makeWebSafe_(getManualCoverageChoices_(payload || {}))
  );
}

function webValidateManualCoverageAssignment(payload) {
  return runCoverageWebRequest_('webValidateManualCoverageAssignment', () =>
    makeWebSafe_(validateManualCoverageAssignment_(payload || {}))
  );
}

function webToggleCoverageStaff(payload) {
  return runCoverageWebRequest_('webToggleCoverageStaff', () => toggleCoverageStaffActive(payload || {}));
}

function webSaveCoverageStaff(payload) {
  return runCoverageWebRequest_('webSaveCoverageStaff', () => saveCoverageStaffFromWeb_(payload || {}));
}

function webDeleteCoverageStaff(payload) {
  return runCoverageWebRequest_('webDeleteCoverageStaff', () => deleteCoverageStaffFromWeb_(payload || {}));
}

function webSaveCoverage(rows) {
  return runCoverageWebRequest_('webSaveCoverage', () => saveCoveragePlan({ rows: rows || [] }));
}

function webSaveCoverageAndCreateHandout(rows) {
  return runCoverageWebRequest_('webSaveCoverageAndCreateHandout', () => {
    const planRows = rows || [];
    const saveResult = saveCoveragePlan({ rows: planRows });
    const assignedRows = planRows.filter(row =>
      String(row.Status || '').trim() === 'Assigned' &&
      String(row.Assigned_Coverage || '').trim()
    );

    let handout = null;
    let handoutError = '';
    if (assignedRows.length) {
      try {
        handout = createCoverageHandoutPackage_(assignedRows, saveResult.date, saveResult.day);
      } catch (error) {
        handoutError = error && error.message ? error.message : String(error || 'Unknown handout error');
      }
    }

    return makeWebSafe_({
      date: saveResult.date,
      day: saveResult.day,
      savedRows: saveResult.savedRows,
      handout: handout,
      handoutError: handoutError
    });
  });
}

function webCreateHandout() {
  return runCoverageWebRequest_('webCreateHandout', () => createCoverageHandoutDocWideFromLatestPreview_());
}

function webCreateHandoutFromRows(rows) {
  return runCoverageWebRequest_('webCreateHandoutFromRows', () => {
    const planRows = rows || [];
    if (!planRows.length) throw new Error('No reviewed coverage rows are available for a handout.');

    const date = normalizeDateKey_(planRows[0].Date);
    const day = String(planRows[0].Day || guessDayCodeFromDate_(date) || '').trim();
    if (!date || !day) throw new Error('The reviewed plan does not have a valid date.');
    if (planRows.some(row => normalizeDateKey_(row.Date) !== date)) {
      throw new Error('A handout can only be created for one date at a time.');
    }

    // Handout creation remains independent from Save Plan, but it must pass
    // the same live validation so stale browser state can never be printed.
    validateCoveragePlanForSave_(date, day, planRows);

    const assignedRows = planRows.filter(row =>
      String(row.Status || '').trim() === 'Assigned' &&
      String(row.Assigned_Coverage || '').trim()
    );
    if (!assignedRows.length) throw new Error('No assigned coverage rows are available for a handout.');

    return makeWebSafe_(createCoverageHandoutPackage_(assignedRows, date, day));
  });
}

function webValidateTeacherSchedule() {
  return runCoverageWebRequest_('webValidateTeacherSchedule', () => validateTeacherScheduleSource_());
}

function generateCoveragePreviewFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const datePrompt = ui.prompt('Generate Coverage Preview', 'Enter date as YYYY-MM-DD', ui.ButtonSet.OK_CANCEL);
  if (datePrompt.getSelectedButton() !== ui.Button.OK) return;
  const dateStr = (datePrompt.getResponseText() || '').trim();
  if (!dateStr) return;

  const dayCode = guessDayCodeFromDate_(dateStr);
  if (!dayCode) {
    ui.alert('Could not derive a day code from that date. Use YYYY-MM-DD, such as 2026-04-27.');
    return;
  }

  const result = generateCoveragePreview({ date: dateStr, day: dayCode });
  ui.alert(
    'Preview complete',
    'Generated ' + result.summary.totalBlocks + ' block(s). Open the Coverage Scheduler web app to review, adjust, and save the plan.',
    ui.ButtonSet.OK
  );
}

function getCoverageBootstrap_(payload) {
  payload = payload || {};
  const today = payload.date || Utilities.formatDate(new Date(), coverageTimeZone_(), 'yyyy-MM-dd');
  const dayCode = payload.day || guessDayCodeFromDate_(today);

  primeCoverageRequestSnapshot_([
    getCoverageStaffSheetName_(),
    'Substitute Availability',
    'Daily Absences',
    'Field Trips',
    '_Preview',
    'Config'
  ]);
  coveragePerfMark_('snapshot-loaded');

  const includeRoster = payload.includeRoster !== false;
  return {
    today: today,
    day: dayCode,
    allStaff: includeRoster && typeof getWebStaffRosterSummary_ === 'function'
      ? getWebStaffRosterSummary_()
      : null,
    allCoverageStaff: getAllCoverageStaff_(today, dayCode),
    // Date navigation must stay cheap. Field-trip candidate materialization is
    // needed only by Generate/manual reassignment, not to display the day.
    currentAbsences: getDailyAbsencesForDate_(today, dayCode),
    currentFieldTrips: getFieldTripsForDate_(today),
    currentPreview: getLatestPreview_(today, dayCode),
    config: getConfigMap_()
  };
}

function menuValidateTeacherScheduleSource() {
  const result = validateTeacherScheduleSource_();
  if (classScheduleSheet_()) writeScheduleCheckSheet_(result.classScheduleFindings || []);
  const warningText = result.warnings.length
    ? '\n\nWarnings:\n- ' + result.warnings.join('\n- ')
    : '\n\nNo structural issues found.';

  SpreadsheetApp.getUi().alert(
    'Teacher Schedule check',
    'Rows: ' + result.rowCount +
      '\nTeachers: ' + result.teacherCount +
      '\nTerms: ' + (result.terms.length ? result.terms.join(', ') : 'none') +
      warningText,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function menuCreateHandoutDoc() {
  const result = createCoverageHandoutDocWideFromLatestPreview_();
  const docs = result.documents && result.documents.length ? result.documents : [result];
  const docText = docs.map(doc =>
    (doc.label || doc.name || 'Handout') + ':\n' + String(doc.url || '')
  ).join('\n\n');

  SpreadsheetApp.getUi().alert(
    'Handout docs created',
    docText + '\n\nHandouts folder:\n' + result.folderUrl,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
