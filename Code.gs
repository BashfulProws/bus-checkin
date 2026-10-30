/**
 * Bus check-in backend for "CWTV Field Trip 1 Nov – Bus Check-in".
 * Bound to the Google Sheet (Extensions > Apps Script).
 *
 * Sheet layout (tab "Check-in"):
 *   A Sr | B Name | C-F four checkpoint checkboxes | G Roll No | H Notes
 *   Row 1 = headers, roster starts row 2 and ends at the first empty Name cell.
 *
 * Endpoints (GET only, so no CORS preflight):
 *   ?action=list                      -> roster + current ticks
 *   ?action=set&row=5&cp=0&val=1      -> set checkpoint cp (0-3) for sheet row to ticked (1) / unticked (0)
 */
const SHEET_NAME = 'Check-in';
const FIRST_ROW = 2;
const CP_COUNT = 4;
const CP_FIRST_COL = 3; // column C

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function lastRosterRow_(sh) {
  const names = sh.getRange(FIRST_ROW, 2, sh.getMaxRows() - FIRST_ROW + 1, 1).getValues();
  let n = 0;
  while (n < names.length && String(names[n][0]).trim() !== '') n++;
  return FIRST_ROW + n - 1;
}

function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sh) return json_({ ok: false, error: 'Sheet tab "' + SHEET_NAME + '" not found' });
    const last = lastRosterRow_(sh);

    if (p.action === 'set') {
      const row = parseInt(p.row, 10);
      const cp = parseInt(p.cp, 10);
      if (!(row >= FIRST_ROW && row <= last && cp >= 0 && cp < CP_COUNT)) {
        return json_({ ok: false, error: 'Bad row or checkpoint' });
      }
      const lock = LockService.getScriptLock();
      lock.waitLock(8000);
      try {
        sh.getRange(row, CP_FIRST_COL + cp).setValue(p.val === '1');
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
      return json_({ ok: true });
    }

    // default: list
    const n = last - FIRST_ROW + 1;
    const head = sh.getRange(1, 1, 1, 8).getValues()[0];
    const data = sh.getRange(FIRST_ROW, 1, n, 8).getValues();
    const rows = data.map(function (r, i) {
      return {
        row: FIRST_ROW + i,
        sr: r[0],
        name: String(r[1]),
        roll: String(r[6] || ''),
        notes: String(r[7] || ''),
        checks: [isTrue_(r[2]), isTrue_(r[3]), isTrue_(r[4]), isTrue_(r[5])]
      };
    });
    return json_({
      ok: true,
      checkpoints: head.slice(2, 6).map(String),
      rows: rows,
      ts: new Date().toISOString()
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
