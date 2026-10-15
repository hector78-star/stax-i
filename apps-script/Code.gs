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

function doGet(e) {
  if (e && e.parameter && e.parameter.debate) return json_(debateStats_());
  return json_(stats_());
}

// --- Debate night: opening and final votes (For / Against / Undecided) ---
// The Debate tab is created automatically on first use. Tick or untick the
// "accepting" boxes in that tab to open and close each vote.
const DEBATE_PHASES = ['open', 'final'];
const DEBATE_OPTIONS = ['for', 'against', 'undecided'];

function debateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Debate');
  if (!sh) {
    sh = ss.insertSheet('Debate');
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

function debateStats_() {
  const rows = debateSheet_().getRange(2, 1, 2, 5).getValues();
  const out = {};
  rows.forEach(r => {
    out[r[0]] = { for: Number(r[1]) || 0, against: Number(r[2]) || 0, undecided: Number(r[3]) || 0, accepting: r[4] === true };
  });
  return out;
}

function debateVote_(p) {
  if (!DEBATE_PHASES.includes(p.phase)) throw new Error('bad phase');
  if (!DEBATE_OPTIONS.includes(p.vote)) throw new Error('bad vote');
  const sh = debateSheet_();
  const row = 2 + DEBATE_PHASES.indexOf(p.phase);
  if (sh.getRange(row, 5).getValue() !== true) return { error: 'closed' };
  const cell = sh.getRange(row, 2 + DEBATE_OPTIONS.indexOf(p.vote));
  cell.setValue(Number(cell.getValue()) + 1);
  SpreadsheetApp.flush();
  return Object.assign({ ok: true }, debateStats_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.type === 'debate') {
      return json_(debateVote_(p));
    } else if (p.type === 'followup') {
      // Stored apart from the vote, with no link to it, so votes stay anonymous.
      const contact = clean_(p.contact, 120), suggestion = clean_(p.suggestion, 400);
      if (!contact && !suggestion) throw new Error('empty');
      sheet_(ss, 'Followups', ['timestamp', 'contact', 'motion_suggestion']).appendRow([new Date(), contact, suggestion]);
      SpreadsheetApp.flush();
      return json_({ ok: true });
    } else if (p.type === 'vote') {
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
