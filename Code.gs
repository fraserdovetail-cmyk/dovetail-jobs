// ════════════════════════════════════════════════════════════════════════════
//  DOVETAIL JOB PORTAL — Google Apps Script backend
//  Deploy: Extensions → Apps Script → Deploy → New deployment → Web app
//  Execute as: Me   |   Who has access: Anyone
// ════════════════════════════════════════════════════════════════════════════

// ─── ENTRY POINTS ────────────────────────────────────────────────────────────

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const raw  = e.postData ? e.postData.contents : (e.parameter ? JSON.stringify(e.parameter) : '{}');
    const body = JSON.parse(raw || '{}');
    let result;

    switch (body.action) {
      case 'getJobs':          result = getJobs();                          break;
      case 'saveJob':          result = saveJob(body.job);                  break;
      case 'deleteJob':        result = deleteJob(body.id);                 break;
      case 'getSchedule':      result = getSchedule();                      break;
      case 'saveSchedule':     result = saveSchedule(body.entries);         break;
      case 'getContractors':   result = getContractors();                   break;
      case 'saveContractors':  result = saveContractors(body.contractors);  break;
      case 'shortUrl':         result = shortUrl(body.url);                 break;
      default:                 result = { error: 'Unknown action: ' + body.action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  JOBS
//  Every field the portal uses is listed here.  Adding a new field to the
//  portal only requires adding its key to JOB_COLS below — the script will
//  automatically append the column to the sheet on the next save.
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
  'activityLog',      // stored as JSON string
  'siteLog',          // stored as JSON string
  'promptHistory',    // stored as JSON string
  'bom',              // stored as JSON string
  'bomFilename',
  'extras',           // stored as JSON string
  'paymentLinks',     // stored as JSON string  ← Monzo payment links
  'keyCode',
  'contractorNotes',
];

function getJobsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Jobs');
  if (!sheet) {
    sheet = ss.insertSheet('Jobs');
    sheet.getRange(1, 1, 1, JOB_COLS.length).setValues([JOB_COLS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Reads the current header row and appends any columns from JOB_COLS that are
 * missing.  Returns the up-to-date headers array.
 */
function ensureJobColumns(sheet) {
  const lastCol = sheet.getLastColumn() || 0;
  const headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    : [];

  JOB_COLS.forEach(col => {
    if (!headers.includes(col)) {
      const newIdx = headers.length + 1;
      sheet.getRange(1, newIdx).setValue(col);
      headers.push(col);
    }
  });

  return headers;
}

function getJobs() {
  const sheet   = getJobsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { jobs: [] };

  const headers = ensureJobColumns(sheet);
  const lastCol = headers.length;
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const idIdx   = headers.indexOf('id');

  const jobs = data
    .filter(row => row[idIdx] !== '' && row[idIdx] !== undefined)
    .map(row => {
      const job = {};
      headers.forEach((h, i) => {
        const v = row[i];
        job[h] = (v === null || v === undefined) ? '' : String(v);
      });
      return job;
    });

  return { jobs };
}

function saveJob(job) {
  if (!job || !job.id) return { error: 'Missing job id' };

  const sheet   = getJobsSheet();
  const headers = ensureJobColumns(sheet);
  const lastCol = headers.length;
  const idIdx   = headers.indexOf('id') + 1; // 1-based

  // Find existing row by id
  let targetRow = -1;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, idIdx, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(job.id)) {
        targetRow = i + 2; // offset by header row
        break;
      }
    }
  }

  const rowData = [headers.map(h => (job[h] !== undefined && job[h] !== null) ? job[h] : '')];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, lastCol).setValues(rowData);
  } else {
    sheet.appendRow(rowData[0]);
  }

  return { success: true };
}

function deleteJob(id) {
  if (!id) return { error: 'Missing id' };

  const sheet   = getJobsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idCol   = headers.indexOf('id') + 1;
  const ids     = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();

  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }

  return { success: true };
}

// ════════════════════════════════════════════════════════════════════════════
//  SCHEDULE
//  Stored as a single JSON blob (cell B2) — the whole array is sent and
//  received in one go, matching how the portal manages the schedule.
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
  const sheet = getScheduleSheet();
  const raw   = String(sheet.getRange(2, 1).getValue() || '');
  try {
    const entries = JSON.parse(raw);
    return { entries: Array.isArray(entries) ? entries : [] };
  } catch (e) {
    return { entries: [] };
  }
}

function saveSchedule(entries) {
  const sheet = getScheduleSheet();
  sheet.getRange(2, 1).setValue(JSON.stringify(Array.isArray(entries) ? entries : []));
  return { success: true };
}

// ════════════════════════════════════════════════════════════════════════════
//  CONTRACTORS
//  Custom contractor list stored as a single JSON blob (cell B2).
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
  const sheet = getContractorsSheet();
  const raw   = String(sheet.getRange(2, 1).getValue() || '');
  try {
    const contractors = JSON.parse(raw);
    return { contractors: Array.isArray(contractors) ? contractors : [] };
  } catch (e) {
    return { contractors: [] };
  }
}

function saveContractors(contractors) {
  const sheet = getContractorsSheet();
  sheet.getRange(2, 1).setValue(JSON.stringify(Array.isArray(contractors) ? contractors : []));
  return { success: true };
}

// ════════════════════════════════════════════════════════════════════════════
//  SHORT URL
//  Returns the URL unchanged.  Replace with a URL-shortening API if needed.
// ════════════════════════════════════════════════════════════════════════════

function shortUrl(url) {
  return { short: url || '' };
}
