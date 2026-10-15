/**
 * STAX-I sign-up backend. Paste into Extensions → Apps Script on the Google Sheet,
 * run setup() once, then Deploy → New deployment → Web app (Execute as: Me, Access: Anyone).
 *
 * Privacy design: votes are never stored as rows. Each vote only increments a counter in
 * the Tally sheet, so no vote can be matched to a sign-up.
 */
const SEED_PLACES = 32;
const SEED_VOTES = { yes: 22, no: 14 }; // starting count, added to real votes (mirror in index.html)

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tally = sheet_(ss, 'Tally', ['vote', 'count']);
  if (tally.getLastRow() === 1) tally.getRange(2, 1, 2, 2).setValues([['yes', 0], ['no', 0]]);
  sheet_(ss, 'Signups', ['timestamp', 'attending', 'first_name', 'email', 'heard', 'future_updates', 'speak_future']);
}

function sheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers).setFrozenRows(1);
  return sh;
}

function doGet() {
  return json_(stats_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.type === 'vote') {
      if (!['yes', 'no'].includes(p.vote)) throw new Error('bad vote');
      const cell = ss.getSheetByName('Tally').getRange(p.vote === 'yes' ? 2 : 3, 2);
      cell.setValue(Number(cell.getValue()) + 1);
    } else if (p.type === 'signup') {
      const email = clean_(p.email, 120);
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('bad email');
      const signups = ss.getSheetByName('Signups');
      if (!signups.getRange(1, 7).getValue()) signups.getRange(1, 7).setValue('speak_future'); // column added after launch
      signups.appendRow([
        new Date(), p.attending === 'yes' ? 'yes' : 'no', clean_(p.first_name, 60), email,
        clean_(p.heard, 40), p.updates ? 'yes' : '', ['yes', 'no'].includes(p.speak) ? p.speak : ''
      ]);
    } else {
      throw new Error('bad type');
    }
    SpreadsheetApp.flush();
    return json_(stats_());
  } catch (err) {
    return json_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function stats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = ss.getSheetByName('Tally').getRange(2, 2, 2, 1).getValues();
  const attending = ss.getSheetByName('Signups').getDataRange().getValues().slice(1)
    .filter(r => r[1] === 'yes').length;
  return {
    yes: SEED_VOTES.yes + Number(t[0][0]),
    no: SEED_VOTES.no + Number(t[1][0]),
    places: SEED_PLACES + attending
  };
}

// Stops a sign-up being read as a spreadsheet formula (=, +, -, @).
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
