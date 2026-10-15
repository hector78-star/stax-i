/**
 * STAX-I debate-night votes. Runs on its own Google Sheet, separate from sign-ups.
 *
 * Setup: new Google Sheet → Extensions → Apps Script → paste this file → Save →
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) → copy the /exec URL.
 *
 * The Debate and Followups tabs are created automatically on first use.
 * Tick or untick the "accepting" boxes in the Debate tab to open and close each vote.
 * Votes are only counters, never rows, so no vote can be linked to anyone.
 */
const PHASES = ['open', 'final'];
const OPTIONS = ['for', 'against', 'undecided'];

function doGet() {
  return json_(stats_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.type === 'debate') return json_(vote_(p));
    if (p.type === 'followup') {
      // Stored apart from the vote, with no link to it.
      const contact = clean_(p.contact, 120), suggestion = clean_(p.suggestion, 400);
      if (!contact && !suggestion) throw new Error('empty');
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName('Followups') || ss.insertSheet('Followups');
      if (sh.getLastRow() === 0) sh.appendRow(['timestamp', 'contact', 'motion_suggestion']).setFrozenRows(1);
      sh.appendRow([new Date(), contact, suggestion]);
      SpreadsheetApp.flush();
      return json_({ ok: true });
    }
    throw new Error('bad type');
  } catch (err) {
    return json_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Debate');
  if (!sh) {
    sh = ss.insertSheet('Debate', 0);
    sh.getRange(1, 1, 3, 5).setValues([
      ['phase', 'for', 'against', 'undecided', 'accepting'],
      ['open', 0, 0, 0, true],
      ['final', 0, 0, 0, false]
    ]);
    sh.getRange(2, 5, 2, 1).insertCheckboxes();
    sh.setFrozenRows(1);
  }
  return sh;
}

function stats_() {
  const out = {};
  sheet_().getRange(2, 1, 2, 5).getValues().forEach(r => {
    out[r[0]] = { for: Number(r[1]) || 0, against: Number(r[2]) || 0, undecided: Number(r[3]) || 0, accepting: r[4] === true };
  });
  return out;
}

function vote_(p) {
  if (!PHASES.includes(p.phase)) throw new Error('bad phase');
  if (!OPTIONS.includes(p.vote)) throw new Error('bad vote');
  const sh = sheet_();
  const row = 2 + PHASES.indexOf(p.phase);
  if (sh.getRange(row, 5).getValue() !== true) return { error: 'closed' };
  const cell = sh.getRange(row, 2 + OPTIONS.indexOf(p.vote));
  cell.setValue(Number(cell.getValue()) + 1);
  SpreadsheetApp.flush();
  return Object.assign({ ok: true }, stats_());
}

// Stops input being read as a spreadsheet formula (=, +, -, @).
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
