// ════════════════════════════════════════════════════════════════════════════
//  DOVETAIL JOB PORTAL — Google Apps Script backend  v2
//  Deploy: Extensions → Apps Script → Deploy → Manage deployments
//         → edit existing → New version → Deploy  (URL stays the same)
// ════════════════════════════════════════════════════════════════════════════

// ─── ENTRY POINTS ────────────────────────────────────────────────────────────

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const raw  = (e && e.postData) ? e.postData.contents : '{}';
    const body = JSON.parse(raw || '{}');
    let result;

    switch (body.action) {
      case 'getJobs':         result = getJobs();                         break;
      case 'saveJob':         result = saveJob(body.job);                 break;
      case 'deleteJob':       result = deleteJob(body.id);                break;
      case 'getSchedule':     result = getSchedule();                     break;
      case 'saveSchedule':    result = saveSchedule(body.entries);        break;
      case 'getContractors':  result = getContractors();                  break;
      case 'saveContractors': result = saveContractors(body.contractors); break;
      case 'addClientNote':     result = addClientNote(body.jobId, body.note); break;
      case 'validatePasscode':  result = validatePasscode(body.passcode);    break;
      case 'shortUrl':          result = { short: body.url || '' };          break;
      default:                  result = { error: 'Unknown action: ' + body.action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('handleRequest error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  JOBS
//
//  Adding a new field to the portal:
//    1. Add the key to JOB_COLS below
//    2. Redeploy the script (Manage deployments → edit → New version)
//    3. The column will be added to the sheet automatically on the next save
// ════════════════════════════════════════════════════════════════════════════

const JOB_COLS = [
  'id',
  'ref',
  'name',
  'client',
  'phone',
  'email',
  'address',
  'value',
  'startDate',
  'startTime',
  'prestartDate',
  'prestartTime',
  'durationNum',
  'durationUnit',
  'notes',
  'stage',
  'declined',
  'created',
  'updated',
  'gcalAdded',
  'depositReceived',
  'activityLog',     // JSON string
  'siteLog',         // JSON string
  'promptHistory',   // JSON string
  'bom',             // JSON string
  'bomFilename',
  'extras',          // JSON string
  'paymentLinks',    // JSON string  ← Monzo payment links
  'keyCode',
  'contractorNotes',
  'clientNotes',    // JSON string — notes posted by client from the client portal
  'bomDropboxUrl',
];

// Find the Jobs sheet by name (tries common variants), falls back to first sheet.
function getJobsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Jobs')
      || ss.getSheetByName('jobs')
      || ss.getSheetByName('Sheet1')
      || ss.getSheets()[0];
}

// Return lowercase header array, automatically appending any columns from
// JOB_COLS that are missing.  Case-insensitive — works with existing sheets
// that have headers like 'ID', 'Ref', 'Name', etc.
function getAndEnsureHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  const rawHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim())
    : [];
  const lcHeaders = rawHeaders.map(h => h.toLowerCase());

  JOB_COLS.forEach(col => {
    if (!lcHeaders.includes(col.toLowerCase())) {
      const nextCol = rawHeaders.length + 1;
      sheet.getRange(1, nextCol).setValue(col);
      rawHeaders.push(col);
      lcHeaders.push(col.toLowerCase());
    }
  });

  return lcHeaders; // always lowercase for consistent key lookup
}

function getJobs() {
  try {
    const sheet = getJobsSheet();
    if (!sheet || sheet.getLastRow() < 2) return { jobs: [] };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return { jobs: [] };

    const allData  = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const lcHeaders = allData[0].map(v => String(v).trim().toLowerCase());
    const idIdx    = lcHeaders.indexOf('id');

    if (idIdx === -1) {
      Logger.log('getJobs: no "id" column found. Headers: ' + lcHeaders.join(', '));
      return { jobs: [] };
    }

    const jobs = [];
    for (let r = 1; r < allData.length; r++) {
      const row = allData[r];
      const idVal = String(row[idIdx] || '').trim();
      if (!idVal) continue;

      const job = {};
      lcHeaders.forEach((h, i) => {
        if (h) job[h] = (row[i] === null || row[i] === undefined) ? '' : String(row[i]);
      });
      // Fill any fields not yet in the sheet with empty string
      JOB_COLS.forEach(col => { if (job[col] === undefined) job[col] = ''; });
      jobs.push(job);
    }

    return { jobs };
  } catch (err) {
    Logger.log('getJobs error: ' + err.toString());
    return { jobs: [], error: err.toString() };
  }
}

function saveJob(job) {
  if (!job || !job.id) return { error: 'Missing job id' };
  try {
    const sheet     = getJobsSheet();
    const lcHeaders = getAndEnsureHeaders(sheet); // adds missing columns

    // Normalise the incoming job to a lowercase-keyed map so camelCase fields
    // (e.g. siteLog, activityLog) match the lowercased sheet headers.
    const jobLc = {};
    Object.keys(job).forEach(k => { jobLc[k.toLowerCase()] = job[k]; });

    const rowData = lcHeaders.map(h => {
      const v = jobLc[h];
      return (v !== undefined && v !== null) ? v : '';
    });

    // Find the existing row for this job (match by id, case-insensitive)
    const idColIdx = lcHeaders.indexOf('id') + 1; // 1-based
    let targetRow  = -1;
    const lastRow  = sheet.getLastRow();

    if (lastRow > 1 && idColIdx > 0) {
      const ids = sheet.getRange(2, idColIdx, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(job.id).trim()) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true };
  } catch (err) {
    Logger.log('saveJob error: ' + err.toString());
    return { error: err.toString() };
  }
}

function deleteJob(id) {
  if (!id) return { error: 'Missing id' };
  try {
    const sheet   = getJobsSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true };

    const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(v => String(v).trim().toLowerCase());
    const idColIdx = headers.indexOf('id') + 1;
    if (idColIdx === 0) return { error: 'No id column in sheet' };

    const ids = sheet.getRange(2, idColIdx, lastRow - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
    return { success: true };
  } catch (err) {
    Logger.log('deleteJob error: ' + err.toString());
    return { error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  SCHEDULE  — stored as a JSON blob (all entries sent/received as one array)
// ════════════════════════════════════════════════════════════════════════════

function getScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Schedule');
  if (!sheet) {
    sheet = ss.insertSheet('Schedule');
    sheet.getRange(1, 1).setValue('entries_json');
  }
  return sheet;
}

function getSchedule() {
  try {
    const sheet = getScheduleSheet();
    // Try cell B2 first (header in A1, data in A2)
    const raw = String(sheet.getRange(2, 1).getValue() || '').trim();
    if (raw) {
      const entries = JSON.parse(raw);
      return { entries: Array.isArray(entries) ? entries : [] };
    }
    return { entries: [] };
  } catch (err) {
    Logger.log('getSchedule error: ' + err.toString());
    return { entries: [] };
  }
}

function saveSchedule(entries) {
  try {
    const sheet = getScheduleSheet();
    sheet.getRange(2, 1).setValue(JSON.stringify(Array.isArray(entries) ? entries : []));
    return { success: true };
  } catch (err) {
    Logger.log('saveSchedule error: ' + err.toString());
    return { error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  CONTRACTORS — stored as a JSON blob
// ════════════════════════════════════════════════════════════════════════════

function getContractorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Contractors');
  if (!sheet) {
    sheet = ss.insertSheet('Contractors');
    sheet.getRange(1, 1).setValue('contractors_json');
  }
  return sheet;
}

function getContractors() {
  try {
    const sheet = getContractorsSheet();
    const raw   = String(sheet.getRange(2, 1).getValue() || '').trim();
    if (raw) {
      const contractors = JSON.parse(raw);
      return { contractors: Array.isArray(contractors) ? contractors : [] };
    }
    return { contractors: [] };
  } catch (err) {
    Logger.log('getContractors error: ' + err.toString());
    return { contractors: [] };
  }
}

function saveContractors(contractors) {
  try {
    const sheet = getContractorsSheet();
    sheet.getRange(2, 1).setValue(JSON.stringify(Array.isArray(contractors) ? contractors : []));
    return { success: true };
  } catch (err) {
    Logger.log('saveContractors error: ' + err.toString());
    return { error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  CLIENT NOTES — appended by the client portal without a full job overwrite
// ════════════════════════════════════════════════════════════════════════════

function addClientNote(jobId, note) {
  if (!jobId || !note || !note.text) return { error: 'Missing jobId or note text' };
  try {
    const sheet   = getJobsSheet();
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return { error: 'Empty sheet' };

    const headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                       .map(v => String(v).trim().toLowerCase());
    const idIdx    = headers.indexOf('id') + 1;
    const notesIdx = headers.indexOf('clientnotes') + 1;

    if (idIdx === 0) return { error: 'No id column' };

    // Find the job row
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: 'No data' };
    const ids = sheet.getRange(2, idIdx, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(jobId).trim()) { targetRow = i + 2; break; }
    }
    if (targetRow < 0) return { error: 'Job not found' };

    // Read existing notes
    let existing = [];
    if (notesIdx > 0) {
      const raw = String(sheet.getRange(targetRow, notesIdx).getValue() || '').trim();
      try { existing = raw ? JSON.parse(raw) : []; } catch(e) { existing = []; }
      if (!Array.isArray(existing)) existing = [];
    }

    // Append new note (unread by default)
    existing.push({ id: Utilities.getUuid(), text: note.text, ts: Date.now(), read: false });
    const json = JSON.stringify(existing);

    if (notesIdx > 0) {
      sheet.getRange(targetRow, notesIdx).setValue(json);
    } else {
      // Column doesn't exist yet — ensure headers then write
      getAndEnsureHeaders(sheet);
      const newHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                           .map(v => String(v).trim().toLowerCase());
      const newIdx = newHeaders.indexOf('clientnotes') + 1;
      if (newIdx > 0) sheet.getRange(targetRow, newIdx).setValue(json);
    }

    return { success: true };
  } catch (err) {
    Logger.log('addClientNote error: ' + err.toString());
    return { error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  AUTH — passcode stored in Script Properties, never in source code
//  To set/change: Apps Script editor → Project Settings → Script Properties
//                 Add property  PASSCODE  with your chosen value
// ════════════════════════════════════════════════════════════════════════════

function validatePasscode(entered) {
  const stored = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!stored) return { error: 'Passcode not configured — set PASSCODE in Script Properties.' };
  return { valid: (entered || '').trim() === stored.trim() };
}

// ════════════════════════════════════════════════════════════════════════════
//  DIAGNOSTICS — run this from the Apps Script editor to check sheet state
// ════════════════════════════════════════════════════════════════════════════

function runDiagnostics() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());
  Logger.log('Sheets in workbook: ' + sheets.join(', '));

  const jobSheet = getJobsSheet();
  Logger.log('Jobs sheet name: ' + (jobSheet ? jobSheet.getName() : 'NOT FOUND'));

  if (jobSheet && jobSheet.getLastRow() > 0) {
    const headers = jobSheet.getRange(1, 1, 1, jobSheet.getLastColumn()).getValues()[0];
    Logger.log('Jobs headers (' + headers.length + '): ' + headers.join(' | '));
    Logger.log('Jobs data rows: ' + (jobSheet.getLastRow() - 1));

    const missing = JOB_COLS.filter(c => !headers.map(h => String(h).toLowerCase()).includes(c.toLowerCase()));
    if (missing.length) {
      Logger.log('MISSING columns (will be added on next save): ' + missing.join(', '));
    } else {
      Logger.log('All required columns present.');
    }
  }
}
