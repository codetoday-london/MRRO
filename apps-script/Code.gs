const S = 'Import status', B = 'All books', H = 'Import history';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('MRRO administration')
    .addItem('Validate selected publisher(s)', 'validate')
    .addItem('Import and freeze selected publisher(s)', 'freezeImport')
    .addItem('Refresh selected publisher import(s)', 'refresh')
    .addItem('Reset selected publisher(s)', 'reset')
    .addItem('Reopen selected publisher submission(s)', 'reopen')
    .addToUi();
}

function rows_() {
  const m = SpreadsheetApp.getActive(), range = m.getActiveRange(), s = range.getSheet();
  if (s.getName() !== S || range.getRow() < 2) throw Error('Select one or more publisher rows in Import status.');
  return Array.from({ length: range.getNumRows() }, (_, i) => {
    const r = range.getRow() + i, x = s.getRange(r, 1, 1, 8).getDisplayValues()[0];
    return x[0] ? [m, s, r, x] : null;
  }).filter(Boolean);
}

function log_(m, p, a, d) { m.getSheetByName(H).appendRow([new Date(), p, a, d]); }

function source_(x) {
  const id = (x[3].match(/[-\w]{25,}/) || [])[0];
  if (!id) throw Error('Publisher sheet link is missing.');
  return SpreadsheetApp.openById(id);
}

function submission_(x) {
  const sh = source_(x).getSheetByName('Submission');
  const status = sh.getRange('B3').getDisplayValue().trim();
  const data = sh.getRange('A9:G1008').getValues().filter(a => a.some(v => String(v).trim())).map(a => [x[0], ...a]);
  if (status !== 'READY TO SUBMIT' || !data.length) throw Error('Source is not READY TO SUBMIT: ' + status);
  return { data };
}

function confirm_(ui, mode, rows) {
  if (rows.length < 2 || !['freeze', 'refresh', 'reset'].includes(mode)) return true;
  const verb = mode === 'freeze' ? 'import and freeze' : mode === 'reset' ? 'reset' : 'refresh';
  const names = rows.map(([, , , x]) => x[0]).join('\n');
  return ui.alert('Confirm batch ' + verb, 'Apply this to ' + rows.length + ' publishers?\n\n' + names, ui.ButtonSet.YES_NO) === ui.Button.YES;
}

function batch_(mode) {
  const ui = SpreadsheetApp.getUi(), rows = rows_();
  if (!confirm_(ui, mode, rows)) return;
  const done = [], failed = [];
  rows.forEach(entry => { try { one_(mode, entry); done.push(entry[3][0]); } catch (error) { failed.push(entry[3][0] + ': ' + error.message); } });
  const title = mode === 'validate' ? 'Validation' : mode === 'freeze' ? 'Import and freeze' : mode === 'refresh' ? 'Refresh' : mode === 'reset' ? 'Reset' : 'Reopen';
  const report = ['Completed (' + done.length + '): ' + (done.join(', ') || 'none')];
  if (failed.length) report.push('Not changed (' + failed.length + '):\n' + failed.join('\n'));
  ui.alert(title, report.join('\n\n'), ui.ButtonSet.OK);
}

function validate() { batch_('validate'); }
function freezeImport() { batch_('freeze'); }
function refresh() { batch_('refresh'); }
function reset() { batch_('reset'); }
function reopen() { batch_('reopen'); }

function one_(mode, entry) {
  const [m, s, r, x] = entry, p = x[0];
  if (/archive/i.test(x[1])) throw Error('Archive rows cannot be changed.');
  if (mode === 'reset') {
    const d = m.getSheetByName(B);
    d.getRange(2, 1, d.getMaxRows() - 1, 1).getDisplayValues().forEach((a, i) => { if (a[0] === p) d.getRange(i + 2, 1, 1, 8).clearContent(); });
    s.getRange(r, 2).setValue('OPEN - awaiting administrator validation');
    s.getRange(r, 3).clearContent();
    s.getRange(r, 5).setValue('PENDING - not yet verified');
    s.getRange(r, 6).setValue('Not locked');
    s.getRange(r, 7, 1, 2).clearContent();
    log_(m, p, 'Reset', 'Cleared imported books; publisher sheet unchanged');
    return;
  }
  if (mode === 'reopen') {
    const src = source_(x), viewers = src.getViewers();
    if (viewers.length > 1) throw Error('Source has multiple viewers.');
    viewers.forEach(viewer => { const email = viewer.getEmail(); src.removeViewer(email); src.addEditor(email); });
    s.getRange(r, 2).setValue('OPEN - publisher submission reopened');
    s.getRange(r, 5).setValue('PENDING - awaiting resubmission');
    s.getRange(r, 6).setValue('OPEN - publisher is editor');
    s.getRange(r, 7, 1, 2).clearContent();
    log_(m, p, 'Reopened', 'Publisher viewer restored to editor');
    return;
  }
  const { data } = submission_(x);
  if (mode === 'validate') {
    s.getRange(r, 2).setValue('READY - validated');
    s.getRange(r, 5).setValue('READY TO SUBMIT - verified');
    s.getRange(r, 6).setValue('Not locked');
    log_(m, p, 'Validated', data.length + ' records ready');
    return;
  }
  if (mode === 'refresh' && !x[7]) throw Error('No prior import is recorded; use import and freeze.');
  const src = source_(x), editors = src.getEditors();
  if (editors.length > 1) throw Error('Source has multiple editors.');
  editors.forEach(editor => { const email = editor.getEmail(); src.removeEditor(email); src.addViewer(email); });
  const d = m.getSheetByName(B);
  if (mode === 'refresh') d.getRange(2, 1, d.getMaxRows() - 1, 1).getDisplayValues().forEach((a, i) => { if (a[0] === p) d.getRange(i + 2, 1, 1, 8).clearContent(); });
  const first = d.getRange(2, 1, d.getMaxRows() - 1, 1).getValues();
  let z = 1;
  first.forEach((a, i) => { if (a[0] !== '') z = i + 2; });
  d.getRange(z + 1, 1, data.length, 8).setValues(data);
  const now = new Date();
  s.getRange(r, 2).setValue('IMPORTED - frozen snapshot');
  s.getRange(r, 3).setValue(data.length);
  s.getRange(r, 5).setValue('READY TO SUBMIT - verified');
  s.getRange(r, 6).setValue('LOCKED - owner-only');
  s.getRange(r, 7).setValue(now);
  s.getRange(r, 8).setValue(now);
  log_(m, p, mode === 'refresh' ? 'Refreshed' : 'Imported', data.length + ' records; frozen snapshot');
}
