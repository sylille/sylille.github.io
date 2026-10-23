/**
 * Study 2 prototype, remote admin relay.
 * Paste into Extensions > Apps Script of a Google Sheet, change SECRET, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Copy the /exec URL into CONFIG.syncUrl in index.html.
 *
 * The participant page sends a state summary and log events (metadata only, no chat text).
 * The admin page sends commands and reads the state and log back.
 * Log rows are written to the "log" tab of this Sheet.
 */
const SECRET = 'change-this-key';   // admin key typed once on admin.html
const KEEP_CMDS = 50;               // commands kept per room
const HEAD = ['room', 'ts_ms', 'time', 'event', 'detail', 'participant', 'persona', 'lang', 'sim_time'];
const props = PropertiesService.getScriptProperties();

function doGet() { return out({ ok: true, service: 'study2-sync' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const room = String(req.room || 'lab').replace(/[^\w-]/g, '').slice(0, 40) || 'lab';
    return out(req.op === 'a' ? admin(req, room) : participant(req, room));
  } catch (err) {
    return out({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function participant(req, room) {
  const seq = Number(props.getProperty('seq:' + room) || 0);
  if (req.state) props.setProperty('state:' + room, JSON.stringify(req.state));
  props.setProperty('seen:' + room, String(Date.now()));
  if (req.logs && req.logs.length) {
    const sh = logSheet();
    const rows = req.logs.map(x => [room, String(x.ts), new Date(x.ts).toISOString(), x.event, x.detail, x.pid, x.persona, x.lang, x.sim]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows);
  }
  if (Number(req.since) < 0) return { seq, cmds: [] };
  const cmds = JSON.parse(props.getProperty('cmds:' + room) || '[]').filter(c => c.seq > Number(req.since));
  return { seq, cmds };
}

function admin(req, room) {
  if (req.key !== SECRET) return { error: 'key' };
  let seq = Number(props.getProperty('seq:' + room) || 0);
  let cmds = JSON.parse(props.getProperty('cmds:' + room) || '[]');
  (req.cmds || []).forEach(c => { seq += 1; cmds.push({ seq, act: String(c.act), data: c.data || {} }); });
  if (req.cmds && req.cmds.length) {
    props.setProperty('cmds:' + room, JSON.stringify(cmds.slice(-KEEP_CMDS)));
    props.setProperty('seq:' + room, String(seq));
  }
  if (req.clearLog) deleteRoomLog(room);
  const rows = roomLog(room);
  const from = Number(req.logFrom) > rows.length ? 0 : Number(req.logFrom || 0);
  return {
    seq,
    state: JSON.parse(props.getProperty('state:' + room) || 'null'),
    seen: Number(props.getProperty('seen:' + room) || 0),
    now: Date.now(),
    logFrom: from,
    logs: rows.slice(from)
  };
}

function logSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('log');
  if (!sh) {
    sh = ss.insertSheet('log');
    sh.getRange('A:I').setNumberFormat('@');   // keep every value as plain text
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function roomLog(room) {
  const sh = logSheet();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues().filter(r => r[0] === room);
}

function deleteRoomLog(room) {
  const sh = logSheet();
  if (sh.getLastRow() < 2) return;
  const keep = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues().filter(r => r[0] !== room);
  sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, HEAD.length).setValues(keep);
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
