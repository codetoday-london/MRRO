const S = 'Import status', B = 'All books', H = 'Import history';
const ADMIN_TEMPLATE_ID = '1PfqNEmnMEdBNturxJLYdJdAl7HC2InHuSNl_yJOX0OY';
const PUBLISHER_TEMPLATE_IDS = ['1OyMQRCF9kK2yuek8tizyabBrcR7F6z0Ztyo1wr-AEV8', '1vJNfXU_i9QDAa2JJfsgVRR7MRHHNGj0DhLroZKnBcEc'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('MRRO administration')
    .addItem('Validate selected publisher(s)', 'validate')
    .addItem('Import and freeze selected publisher(s)', 'freezeImport')
    .addItem('Refresh selected publisher import(s)', 'refresh')
    .addItem('Reopen and reset selected publisher submission(s)', 'reopenAndReset')
    .addToUi();
}

function rows_() {
  const m = SpreadsheetApp.getActive(), range = m.getActiveRange(), s = range.getSheet();
  if (s.getName() !== S || range.getRow() < 2) throw Error('Select one or more publisher rows in Import status.');
  return Array.from({ length: range.getNumRows() }, (_, i) => {
    const r = range.getRow() + i, x = s.getRange(r, 1, 1, 9).getDisplayValues()[0];
    return x[0] && x[3] ? [m, s, r, x] : null;
  }).filter(Boolean);
}

function log_(m, p, a, d) { m.getSheetByName(H).appendRow([new Date(), p, a, d]); }

function source_(x) {
  const id = (x[3].match(/[-\w]{25,}/) || [])[0];
  if (!id) throw Error('Publisher sheet link is missing.');
  return SpreadsheetApp.openById(id);
}

function years_(m) {
  const sh = m.getSheetByName('Settings');
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  let start, end;
  values.forEach((row, index) => {
    const label = String(row[0]).toLowerCase();
    const year = Number(row[1]);
    if (!Number.isInteger(year)) return;
    if (/(earliest|start|first)/.test(label) && /year/.test(label)) start = { value: year, ref: 'Settings!$B$' + (index + 1) };
    if (/(latest|end|last)/.test(label) && /year/.test(label)) end = { value: year, ref: 'Settings!$B$' + (index + 1) };
  });
  if (!start || !end || start.value > end.value) throw Error('Settings must contain valid earliest/start and latest/end publication years.');
  return { start, end };
}

function publisherEmail_(x) {
  const email = String(x[8] || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw Error('Enter the publisher email address in Import status column I before using this action.');
  return email;
}

function setPublisherAccess_(src, email, access) {
  const editors = src.getEditors().map(user => user.getEmail());
  const viewers = src.getViewers().map(user => user.getEmail());
  if (access === 'viewer') {
    if (viewers.includes(email)) return;
    if (!editors.includes(email)) throw Error('The recorded publisher email is not an editor of this publisher sheet.');
    src.removeEditor(email);
    src.addViewer(email);
    return;
  }
  if (editors.includes(email)) return;
  if (!viewers.includes(email)) throw Error('The recorded publisher email is not a viewer of this publisher sheet.');
  src.removeViewer(email);
  src.addEditor(email);
}

function primePublisher_(src, endYear) {
  const sh = src.getSheetByName('Submission');
  sh.getRange('A6:B6').setValues([['Submission year', endYear]]);
  sh.getRange('A6').setFontWeight('bold');
  sh.getRange('B6').setNumberFormat('0');
  const rule = SpreadsheetApp.newDataValidation().requireNumberEqualTo(endYear).setAllowInvalid(false).setHelpText('Enter the current submission year: ' + endYear + '.').build();
  sh.getRange('A9:A1008').setDataValidation(rule);
  const formula = '=IF(COUNTA(A9:G9)=0,"",IF(NOT(AND(ISNUMBER(A9),A9=INT(A9),A9=$B$6)),"Year must be the current submission year ("&$B$6&"). ","")&IF(B9="","Book title is required. ","")&IF(OR(D9="",REGEXMATCH(D9,"(?i)(^| )and( |$)|[&%;]|^,|,$|,,")),"Use comma-separated author names only. ","")&IF(NOT(AND(ISNUMBER(E9),E9=INT(E9),E9>0)),"Pages must be a positive whole number. ","")&IF(NOT(AND(ISNUMBER(F9),F9>0)),"Price must be a positive number. ","")&IF(NOT(AND(ISNUMBER(G9),G9=INT(G9),G9>=1,G9<=3)),"Classification must be 1, 2, or 3.",""))';
  sh.getRange('H9').setFormula(formula);
  sh.getRange('H9').copyTo(sh.getRange('H9:H1008'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
}

function submission_(m, x) {
  const year = years_(m).end.value;
  const sh = source_(x).getSheetByName('Submission');
  const status = sh.getRange('B3').getDisplayValue().trim();
  const data = sh.getRange('A9:G1008').getValues().filter(a => a.some(v => String(v).trim())).map(a => [x[0], ...a]);
  if (status !== 'READY TO SUBMIT' || !data.length) throw Error('Source is not READY TO SUBMIT: ' + status);
  if (data.some(row => Number(row[1]) !== year)) throw Error('Every submitted book must have the current submission year: ' + year + '.');
  return { data, year };
}

function currentYearRows_(m, publisher, year) {
  const sh = m.getSheetByName(B), last = lastBookDataRow_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 8).getValues().map((row, i) => ({ row: i + 2, values: row })).filter(item => item.values[0] === publisher && Number(item.values[1]) === year);
}

function lastBookDataRow_(sh) {
  let last = 1;
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).getValues().forEach((row, i) => { if (String(row[0]).trim()) last = i + 2; });
  return last;
}

function removeCurrentYearRows_(m, publisher, year) {
  const sh = m.getSheetByName(B), rows = currentYearRows_(m, publisher, year);
  rows.reverse().forEach(item => sh.deleteRow(item.row));
  return rows.length;
}

function confirm_(ui, mode, rows, year) {
  const action = mode === 'freeze' ? 'import and freeze' : mode === 'refresh' ? 'refresh' : 'reopen and reset';
  const names = rows.map(([, , , x]) => x[0]).join('\n');
  return ui.alert('Confirm ' + action, 'Submission year: ' + year + '\n\nApply this to ' + rows.length + ' publisher(s)?\n\n' + names, ui.ButtonSet.YES_NO) === ui.Button.YES;
}

function batch_(mode) {
  const ui = SpreadsheetApp.getUi(), rows = rows_(), year = years_(SpreadsheetApp.getActive()).end.value;
  if (!rows.length) throw Error('Select one or more publisher rows with a publisher sheet link.');
  if (!confirm_(ui, mode, rows, year)) return;
  const done = [], failed = [];
  rows.forEach(entry => { try { one_(mode, entry); done.push(entry[3][0]); } catch (error) { failed.push(entry[3][0] + ': ' + error.message); } });
  const title = mode === 'validate' ? 'Validation' : mode === 'freeze' ? 'Import and freeze' : mode === 'refresh' ? 'Refresh' : 'Reopen and reset';
  const report = ['Completed (' + done.length + '): ' + (done.join(', ') || 'none')];
  if (failed.length) report.push('Not changed (' + failed.length + '):\n' + failed.join('\n'));
  ui.alert(title, report.join('\n\n'), ui.ButtonSet.OK);
}

function validate() { batch_('validate'); }
function freezeImport() { batch_('freeze'); }
function refresh() { batch_('refresh'); }
function reopenAndReset() { batch_('reopen'); }

function one_(mode, entry) {
  const [m, s, r, x] = entry, publisher = x[0], email = publisherEmail_(x), year = years_(m).end.value;
  if (/archive/i.test(x[1])) throw Error('Archive rows cannot be changed.');
  if (mode === 'reopen') {
    const src = source_(x);
    setPublisherAccess_(src, email, 'editor');
    primePublisher_(src, year);
    src.getSheetByName('Submission').getRange('A9:G1008').clearContent();
    s.getRange(r, 2).setValue('OPEN - current-year submission reopened');
    s.getRange(r, 3).clearContent();
    s.getRange(r, 5).setValue('PENDING - awaiting ' + year + ' submission');
    s.getRange(r, 6).setValue('OPEN - publisher is editor');
    s.getRange(r, 7, 1, 2).clearContent();
    log_(m, publisher, 'Reopened and reset', 'Prepared blank ' + year + ' submission; publisher restored to editor');
    return;
  }
  const { data } = submission_(m, x);
  const existing = currentYearRows_(m, publisher, year);
  if (mode === 'validate') {
    s.getRange(r, 2).setValue('READY - validated for ' + year);
    s.getRange(r, 5).setValue('READY TO SUBMIT - verified');
    s.getRange(r, 6).setValue('Not locked');
    log_(m, publisher, 'Validated', data.length + ' records ready for ' + year);
    return;
  }
  if (mode === 'freeze' && existing.length) throw Error('This publisher already has ' + existing.length + ' imported book(s) for ' + year + '. Use Refresh instead.');
  if (mode === 'refresh' && !existing.length) throw Error('No imported books exist for ' + year + '; use Import and freeze.');
  const src = source_(x);
  setPublisherAccess_(src, email, 'viewer');
  if (mode === 'refresh') removeCurrentYearRows_(m, publisher, year);
  const allBooks = m.getSheetByName(B), first = lastBookDataRow_(allBooks) + 1;
  allBooks.getRange(first, 1, data.length, 8).setValues(data);
  rebuildPaymentTabs_(m);
  const now = new Date();
  s.getRange(r, 2).setValue('IMPORTED - ' + year + ' frozen snapshot');
  s.getRange(r, 3).setValue(data.length);
  s.getRange(r, 5).setValue('READY TO SUBMIT - verified');
  s.getRange(r, 6).setValue('LOCKED - publisher is viewer');
  s.getRange(r, 7).setValue(now);
  s.getRange(r, 8).setValue(now);
  log_(m, publisher, mode === 'refresh' ? 'Refreshed ' + year : 'Imported ' + year, data.length + ' records; frozen snapshot');
}

function configureAnnualWorkflow() {
  const master = SpreadsheetApp.getActive();
  configureMaster_(master);
  configureMaster_(SpreadsheetApp.openById(ADMIN_TEMPLATE_ID));
  const links = master.getSheetByName(S).getRange(2, 4, master.getSheetByName(S).getLastRow() - 1, 1).getDisplayValues().flat();
  const ids = links.map(link => (link.match(/[-\w]{25,}/) || [])[0]).filter(Boolean).concat(PUBLISHER_TEMPLATE_IDS);
  [...new Set(ids)].forEach(id => primePublisher_(SpreadsheetApp.openById(id), years_(master).end.value));
}

function formatPublisherEmailInputs() {
  [SpreadsheetApp.getActive(), SpreadsheetApp.openById(ADMIN_TEMPLATE_ID)].forEach(m => {
    const status = m.getSheetByName(S);
    status.getRange('I1').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    status.setColumnWidth(9, 150);
    status.getRange('I2:I19').setBackground('#fff2cc');
  });
}

function repairCalculations_(m) {
  const sh = m.getSheetByName('Calculations'), end = sh.getMaxRows();
  const formulas = [
    '=IF(\'All books\'!F2="","",INT((\'All books\'!F2-1)/100)+1)',
    '=IF(\'All books\'!G2="","",IF(\'All books\'!I2<>"",\'All books\'!I2,LOOKUP(\'All books\'!G2,{0,0.0000001,10.0000001,20.0000001,50.0000001,80.0000001,100.0000001,150.0000001},{0,1,2,3,4,5,6,7})))',
    '=IF(\'All books\'!E2="","",LEN(\'All books\'!E2)-LEN(SUBSTITUTE(\'All books\'!E2,",",""))+1)',
    '=IF(C2="","",1/C2)',
    '=IF(\'All books\'!H2="","",\'All books\'!H2)',
    '=IF(A2="","",IF(AND(\'All books\'!B2>=Settings!$B$14,\'All books\'!B2<=Settings!$B$15),(A2+B2)*E2,0))',
    '=IF(F2="","",Settings!$B$4)',
    '=IF(F2="","",F2*G2)'
  ];
  formulas.forEach((formula, index) => {
    const seed = sh.getRange(2, index + 1);
    seed.setFormula(formula);
    seed.copyTo(sh.getRange(2, index + 1, end - 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  });
}

function points_(pages, price, classification, override) {
  const a = Math.floor((Number(pages) - 1) / 100) + 1;
  const p = override === '' || override === null ? Number(price) : Number(override);
  const b = p <= 0 ? 0 : p <= 10 ? 1 : p <= 20 ? 2 : p <= 50 ? 3 : p <= 80 ? 4 : p <= 100 ? 5 : p <= 150 ? 6 : 7;
  return (a + b) * Number(classification);
}

function rebuildPaymentTabs_(m) {
  const years = years_(m), all = m.getSheetByName(B), last = lastBookDataRow_(all);
  const rows = last < 2 ? [] : all.getRange(2, 1, last - 1, 9).getValues().filter(row => String(row[0]).trim());
  const allocation = [], publishers = [];
  const publisherSet = new Set();
  rows.forEach(row => {
    const [publisher, year, , , authors, pages, price, classification, override] = row;
    if (!publisherSet.has(publisher)) { publisherSet.add(publisher); publishers.push(publisher); }
    if (Number(year) < years.start.value || Number(year) > years.end.value || !String(authors).trim()) return;
    const share = points_(pages, price, classification, override) / String(authors).split(',').length;
    String(authors).split(',').forEach(author => allocation.push(["['" + publisher + "'] " + author.trim(), share]));
  });
  const authorAllocation = m.getSheetByName('Author allocation');
  authorAllocation.getRange(2, 1, authorAllocation.getMaxRows() - 1, 2).clearContent();
  if (allocation.length) authorAllocation.getRange(2, 1, allocation.length, 2).setValues(allocation);
  const authorPayments = m.getSheetByName('Author payments');
  authorPayments.getRange(2, 1, authorPayments.getMaxRows() - 1, 2).clearContent();
  const authors = [...new Set(allocation.map(row => row[0]))].sort();
  if (authors.length) {
    authorPayments.getRange(2, 1, authors.length, 1).setValues(authors.map(author => [author]));
    authorPayments.getRange('B2').setFormula('=SUMIF(\'Author allocation\'!A:A,A2,\'Author allocation\'!B:B)*Settings!$B$4*0.5');
    authorPayments.getRange('B2').copyTo(authorPayments.getRange(2, 2, authors.length, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  }
  const publisherPayments = m.getSheetByName('Publisher payments');
  publisherPayments.getRange(2, 1, publisherPayments.getMaxRows() - 1, 2).clearContent();
  if (publishers.length) {
    publisherPayments.getRange(2, 1, publishers.length, 1).setValues(publishers.map(publisher => [publisher]));
    publisherPayments.getRange('B2').setFormula('=SUMIF(\'All books\'!A:A,A2,Calculations!H:H)*0.5');
    publisherPayments.getRange('B2').copyTo(publisherPayments.getRange(2, 2, publishers.length, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  }
}

function configureMaster_(m) {
  const years = years_(m), status = m.getSheetByName(S), start = m.getSheetByName('Start here'), books = m.getSheetByName(B), settings = m.getSheetByName('Settings');
  settings.getRange('A1').setValue(m.getId() === ADMIN_TEMPLATE_ID ? 'MRRO Administrator Template' : 'MRRO Administrator Master');
  settings.getRange('B4').setFormula('=IFERROR(B2/B3,0)');
  settings.getRange('B5').setValue('Annual submissions are ready to be prepared.');
  settings.getRange('A8').setValue('Historic-data note');
  settings.getRange('B8').setValue('All books retains earlier imports. The eligibility years decide which books are included in the calculation.');
  repairCalculations_(m);
  rebuildPaymentTabs_(m);
  status.getRange('I1').setValue('Publisher email address');
  status.getRange('I1').setFontWeight('bold').setBackground('#fff2cc').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  status.setColumnWidth(9, 150);
  status.getRange('I2:I19').setBackground('#fff2cc');
  status.getRange(2, 9, status.getMaxRows() - 1, 1).setNumberFormat('@');
  start.getRange('B7').setValue('In Import status, select one or more publisher rows. Validate checks the submission year and entries. Import and freeze adds the current-year books to All books and changes only the recorded publisher email from editor to viewer.');
  start.getRange('A10').setValue('6. Prepare the next annual submission');
  start.getRange('B10').setValue('First update the end year in Settings. Then use Reopen and reset to restore the recorded publisher email as editor, clear that publisher workbook, and set it to accept the new submission year only. Refresh replaces only that publisher’s imported books for the current year.');
  start.getRange('B12').setValue('Yellow = administrator input or action required. In All books: white = current submission year; pale blue = an earlier eligible year; pale red = outside the eligibility range and excluded from calculations.');
  start.getRange('B18').setValue('Create the new publisher workbook from the publisher template, name it for the publisher, enter the publisher’s email address in Import status column I, and put its link in column D. Share it with that publisher as an editor when appropriate.');
  start.getRange('B20').setValue('Publisher sheets check: the publication year exactly matches the current submission year; a book title; an author name using commas between multiple names; positive whole-number pages; a positive price; and classification 1, 2 or 3. Duplicate ISBNs are checked only in All books.');
  const dataRange = books.getRange(2, 1, books.getMaxRows() - 1, 10);
  const retained = books.getConditionalFormatRules().filter(rule => {
    const condition = rule.getBooleanCondition();
    if (!condition) return true;
    return !condition.getCriteriaValues().some(value => String(value).includes('ISNUMBER($B2)'));
  });
  const startRef = 'INDIRECT("' + years.start.ref + '")', endRef = 'INDIRECT("' + years.end.ref + '")';
  const excluded = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($B2),OR($B2<' + startRef + ',$B2>' + endRef + '))').setBackground('#f4cccc').setRanges([dataRange]).build();
  const historic = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($B2),$B2>=' + startRef + ',$B2<' + endRef + ')').setBackground('#d9eaf7').setRanges([dataRange]).build();
  const current = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($B2),$B2=' + endRef + ')').setBackground('#ffffff').setRanges([dataRange]).build();
  books.setConditionalFormatRules(retained.concat([excluded, historic, current]));
}
