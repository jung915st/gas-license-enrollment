// Node harness: loads the .gs files into one shared global scope (like Apps Script)
// with in-memory fakes of SpreadsheetApp / FormApp / MailApp, then drives the real flows.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dir = process.argv[2] || path.join(__dirname, '..');
const files = ['code.gs', 'allocate.gs', 'setupForm.gs', 'confirm.gs', 'tests_enrollment.gs'];

// ---------- fake spreadsheet ----------
class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr: nr || 1, nc: nc || 1 }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sheet.get(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  getValue() { return this.sheet.get(this.r, this.c); }
  setValues(v) {
    assert.strictEqual(v.length, this.nr, `setValues rows ${v.length} != ${this.nr} on ${this.sheet.name}`);
    v.forEach((row, i) => {
      assert.strictEqual(row.length, this.nc, `setValues cols ${row.length} != ${this.nc} on ${this.sheet.name}`);
      row.forEach((val, j) => this.sheet.set(this.r + i, this.c + j, val));
    });
    return this;
  }
  setValue(v) { this.sheet.set(this.r, this.c, v); return this; }
  setNumberFormat() { return this; }
  insertCheckboxes() { return this; }
  setDataValidation() { return this; }
  getRow() { return this.r; }
  getSheet() { return this.sheet; }
}
class Sheet {
  constructor(name, formUrl, id) { this.name = name; this.cells = []; this.formUrl = formUrl || null; this.id = id; }
  getSheetId() { return this.id; }
  get(r, c) { const row = this.cells[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined ? '' : v; }
  set(r, c, v) {
    assert.ok(r >= 1 && c >= 1, 'bad cell');
    while (this.cells.length < r) this.cells.push([]);
    this.cells[r - 1][c - 1] = v;
  }
  getRange(r, c, nr, nc) {
    assert.ok(Number.isInteger(r) && Number.isInteger(c), 'getRange needs integers');
    assert.ok((nr === undefined || nr >= 1) && (nc === undefined || nc >= 1), `getRange empty range on ${this.name}`);
    return new Range(this, r, c, nr, nc);
  }
  getLastRow() {
    for (let i = this.cells.length; i >= 1; i--) {
      if ((this.cells[i - 1] || []).some((v) => v !== '' && v !== undefined && v !== false)) return i;
    }
    return 0;
  }
  getLastColumn() { return Math.max(0, ...this.cells.map((r) => (r ? r.length : 0))); }
  getMaxRows() { return 1000; }
  clearContents() { this.cells = []; return this; }
  setFrozenRows() {}
  getFormUrl() { return this.formUrl; }
  appendRow(row) { const n = this.getLastRow() + 1; row.forEach((v, j) => this.set(n, j + 1, v)); }
  getName() { return this.name; }
}
class Spreadsheet {
  constructor() { this.sheets = []; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n, formUrl) { this.nextId = (this.nextId || 100) + 1; const s = new Sheet(n, formUrl, this.nextId); this.sheets.push(s); return s; }
  getSheets() { return this.sheets; }
  getFormUrl() { const s = this.sheets.find((x) => x.formUrl); return s ? s.formUrl : null; }
  getId() { return 'ss-id'; }
}

// ---------- fake form ----------
const ItemType = { TEXT: 'TEXT', PARAGRAPH_TEXT: 'PARAGRAPH_TEXT', MULTIPLE_CHOICE: 'MULTIPLE_CHOICE', CHECKBOX: 'CHECKBOX', LIST: 'LIST', PAGE_BREAK: 'PAGE_BREAK' };
class Item {
  constructor(form, type, title) { Object.assign(this, { form, type, title: title || '', choices: [], required: false, help: '' }); }
  getTitle() { return this.title; } setTitle(t) { this.title = t; return this; }
  getType() { return this.type; }
  getIndex() { return this.form.items.indexOf(this); }
  setHelpText(t) { this.help = t; return this; }
  setRequired(b) { this.required = b; return this; }
  setChoiceValues(v) { this.choices = Array.from(v); return this; }
  createChoice(value, page) { return { value, page }; }
  setChoices(list) { this.choices = Array.from(list, (c) => c.value); this.nav = Array.from(list, (c) => c.page); return this; }
  setGoToPage(page) { assert.strictEqual(this.type, ItemType.PAGE_BREAK); this.goTo = page; return this; }
  setValidation(v) { this.validation = v; return this; }
  asTextItem() { assert.strictEqual(this.type, ItemType.TEXT); return this; }
  asParagraphTextItem() { assert.strictEqual(this.type, ItemType.PARAGRAPH_TEXT); return this; }
  asCheckboxItem() { assert.strictEqual(this.type, ItemType.CHECKBOX); return this; }
  asMultipleChoiceItem() { assert.strictEqual(this.type, ItemType.MULTIPLE_CHOICE); return this; }
  asListItem() { assert.strictEqual(this.type, ItemType.LIST); return this; }
}
class Form {
  constructor(title, id) { this.title = title; this.items = []; this.description = ''; this.id = id; this.url = 'https://forms/' + id; }
  getId() { return this.id; }
  addPageBreakItem() { return this.add(ItemType.PAGE_BREAK); }
  getItems() { return this.items.slice(); }
  add(type) { const it = new Item(this, type); this.items.push(it); return it; }
  addTextItem() { return this.add(ItemType.TEXT); }
  addCheckboxItem() { return this.add(ItemType.CHECKBOX); }
  addMultipleChoiceItem() { return this.add(ItemType.MULTIPLE_CHOICE); }
  deleteItem(it) { this.items.splice(this.items.indexOf(it), 1); }
  moveItem(from, to) {
    assert.ok(Number.isInteger(from) && Number.isInteger(to), 'moveItem expects indexes');
    const [it] = this.items.splice(from, 1); this.items.splice(to, 0, it); return it;
  }
  setDescription(d) { this.description = d; return this; }
  setDestination(type, id) { this.destination = id; if (this.onDestination) this.onDestination(this); return this; }
  getPublishedUrl() { return 'https://forms/published'; }
  getEditUrl() { return 'https://forms/edit'; }
}

function makeEnv(opts) {
  opts = opts || {};
  const ss = new Spreadsheet();
  const mails = [];
  const alerts = [];
  const forms = {};
  const createdForms = [];
  const props = {};
  const triggers = [];
  // what Google does when a form is linked to a spreadsheet: a new response tab with the question titles as headers
  const linkForm = (form) => {
    const n = ss.sheets.filter((s) => s.formUrl).length + 1;
    const sheet = ss.insertSheet('表單回應 ' + n, form.url);
    const headers = ['時間戳記'].concat(form.items.filter((i) => i.type !== ItemType.PAGE_BREAK).map((i) => i.title));
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  };
  const newForm = (title) => {
    const f = new Form(title, 'form-' + (Object.keys(forms).length + 1));
    forms[f.url] = f;
    return f;
  };
  const env = {
    console, JSON, Date, Math, Object, Array, String, Number, Error, isNaN,
    Logger: { log: () => {} },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 16).replace('T', ' '), sleep: () => {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      flush: () => {},
      getUi: () => {
        if (opts.noUi) throw new Error('no ui');
        const menu = { addItem: () => menu, addSeparator: () => menu, addSubMenu: () => menu, addToUi: () => {} };
        return {
          alert: (a, b, c) => { alerts.push(b === undefined ? a : a + '\n' + b); return 'YES'; },
          ButtonSet: { YES_NO: 'YES_NO' }, Button: { YES: 'YES' }, createMenu: () => menu,
        };
      },
      newDataValidation: () => { const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; },
    },
    FormApp: {
      ItemType, DestinationType: { SPREADSHEET: 'SPREADSHEET' },
      openByUrl: (url) => { assert.ok(forms[url], 'unknown form ' + url); return forms[url]; },
      openById: (id) => { const f = Object.values(forms).find((x) => x.id === id); if (!f) throw new Error('no form ' + id); return f; },
      create: (title) => {
        const f = newForm(title);
        createdForms.push(f);
        f.onDestination = (form) => { if (opts.lateResponseSheet) env.__pendingLink = () => linkForm(form); else linkForm(form); };
        return f;
      },
      createTextValidation: () => { const b = { setHelpText: () => b, requireTextIsEmail: () => b, build: () => 'email' }; return b; },
    },
    MailApp: {
      sendEmail: (a, b, c) => {
        const m = typeof a === 'object' ? a : { to: a, subject: b, body: c };
        if (opts.failMailTo && m.to === opts.failMailTo) throw new Error('mail failed');
        mails.push(m);
      },
      getRemainingDailyQuota: () => (opts.mailQuota === undefined ? 100 : opts.mailQuota),
    },
    GmailApp: { sendEmail: () => {} },
    HtmlService: {},
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (props[k] === undefined ? null : props[k]),
      setProperty: (k, v) => { props[k] = v; },
      deleteProperty: (k) => { delete props[k]; },
    }) },
    ScriptApp: {
      EventType: { ON_FORM_SUBMIT: 'ON_FORM_SUBMIT' },
      getProjectTriggers: () => triggers,
      newTrigger: (fn) => { const b = { forSpreadsheet: () => b, onFormSubmit: () => b, create: () => triggers.push({ getHandlerFunction: () => fn, getEventType: () => 'ON_FORM_SUBMIT' }) }; return b; },
    },
  };
  const ctx = vm.createContext(env);
  files.forEach((f) => vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f }));
  const call = (expr, vars) => { Object.assign(env, vars || {}); return vm.runInContext(expr, ctx); };
  return { ss, mails, alerts, forms, createdForms, props, triggers, env, call, newForm, linkForm };
}

// submit a form response: append to that form's response tab, then fire the spreadsheet trigger
function submitTo(E, sheet, answers, when) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map((h) => (h === '時間戳記' ? when : (answers[h] === undefined ? '' : answers[h])));
  sheet.appendRow(row);
  const r = sheet.getLastRow();
  const namedValues = {};
  headers.forEach((h, i) => { (namedValues[h] = namedValues[h] || []).push(h === '時間戳記' ? when.toISOString() : String(row[i])); });
  E.call('onFormSubmit(__evt)', { __evt: { namedValues, range: sheet.getRange(r, 1, 1, headers.length) } });
}
const Q = {
  name: '您的大名', email: '您要用來開通授權的 Google 帳號', role: '您的職務（學年／科任／行政）', grade: '任教學年',
  subject: '任教科目', office: '所屬處室', software: '您要申請的軟體授權', workspace: '申請 Google AI Pro 教育版的學校帳號',
};
const SCHOOL = '@mail.lres.tc.edu.tw';
// who: {name, email, role:'學年'|'科任'|'行政', detail, workspace}
function answersOf(who, software) {
  const a = { [Q.name]: who.name, [Q.email]: who.email, [Q.role]: who.role, [Q.software]: software, [Q.workspace]: who.workspace || '' };
  a[{ 學年: Q.grade, 科任: Q.subject, 行政: Q.office }[who.role]] = who.detail;
  return a;
}
const regRows = (E) => { const s = E.ss.getSheetByName('申請登記'); const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, 12).getValues(); };
const t = (min) => new Date(2026, 10, 17, 8, min, 0);
let passed = 0;
const ok = (name, fn) => { fn(); passed++; console.log('ok   ' + name); };

// A spreadsheet as it is today: the old form + its response tab, old serials, the Wordwall sheet.
function schoolEnv(opts) {
  const E = makeEnv(opts);
  const oldForm = E.newForm('舊表單');
  oldForm.addTextItem().setTitle('您的大名');
  oldForm.addTextItem().setTitle('您的常用email信箱');
  oldForm.addTextItem().setTitle('您的職務');
  oldForm.addMultipleChoiceItem().setTitle('您要申請的軟體授權').setChoiceValues(['Hiteach', 'Wordwall']);
  const oldSheet = E.linkForm(oldForm);
  oldSheet.appendRow([new Date(2025, 8, 1), '去年', 'old@x.tw', '教師', 'Hiteach']);
  const lic = E.ss.insertSheet('序號列表');
  lic.getRange(1, 1, 1, 3).setValues([['序號', '使用者Email', '分配時間']]);
  lic.appendRow(['HT-OLD', 'old@x.tw', new Date(2025, 8, 1)]);
  lic.appendRow(['HT-A', '', '']);
  lic.appendRow(['HT-B', '', '']);
  const ww = E.ss.insertSheet('wordwall授權確認');
  ww.getRange(1, 1, 1, 3).setValues([['姓名', '帳號', '狀態']]);
  return { E, oldForm, oldSheet, lic, ww };
}
function readyEnv(opts) {
  const S = schoolEnv(opts);
  S.E.call('setupSheets()');
  S.E.call('setupForm()');
  S.E.call('createFormSubmitTrigger()');
  S.E.call('createFormSubmitTrigger()');
  S.form = S.E.createdForms[0];
  S.sheet = S.E.ss.sheets.find((s) => s.formUrl === S.form.url);
  S.send = (i, who, software) => submitTo(S.E, S.sheet, answersOf(who, software), t(i));
  return S;
}
const teacher = (n, grade, extra) => Object.assign({ name: '老師' + n, email: `t${n}@gmail.com`, role: '學年', detail: grade }, extra || {});

// ============ 1. unit tests shipped to the user ============
ok('runEnrollmentTests (shipped unit tests)', () => {
  const E = makeEnv();
  const res = E.call('runEnrollmentTests()');
  assert.ok(res.length >= 45 && res.every((r) => r.ok));
  console.log('     ' + res.length + ' checks');
});

// ============ 2. a brand-new form, old form untouched ============
ok('setupForm creates a NEW form with role branching; the old form and tab are untouched', () => {
  const { E, oldForm, oldSheet, form, sheet, lic } = readyEnv();
  assert.strictEqual(E.createdForms.length, 1);
  assert.notStrictEqual(form, oldForm);
  assert.deepStrictEqual(oldForm.items.map((i) => i.title), ['您的大名', '您的常用email信箱', '您的職務', '您要申請的軟體授權']);
  assert.deepStrictEqual(oldForm.items[3].choices, ['Hiteach', 'Wordwall']);
  assert.strictEqual(oldSheet.getLastRow(), 2);

  const titles = form.items.map((i) => (i.type === 'PAGE_BREAK' ? '#' + i.title : i.title));
  assert.deepStrictEqual(titles, [Q.name, Q.email, Q.role, '#【學年】', Q.grade, '#【科任】', Q.subject, '#【行政】', Q.office,
    '#申請軟體授權', Q.software, Q.workspace]);
  const [, email, role, pGrade, grade, pSubject, subject, pOffice, office, pSoftware, software, workspace] = form.items;
  assert.deepStrictEqual(role.choices, ['學年', '科任', '行政']);
  assert.deepStrictEqual(role.nav, [pGrade, pSubject, pOffice], 'each role jumps to its own section');
  assert.strictEqual(pSubject.goTo, pSoftware, 'after the 學年 section -> software');
  assert.strictEqual(pOffice.goTo, pSoftware, 'after the 科任 section -> software');
  assert.strictEqual(pSoftware.goTo, undefined, '行政 section flows straight on');
  assert.deepStrictEqual(grade.choices, ['一年級', '二年級', '三年級', '四年級', '五年級', '六年級']);
  assert.strictEqual(subject.type, 'TEXT');
  assert.deepStrictEqual(office.choices, ['教務處', '學務處', '總務處', '輔導室']);
  assert.ok(email.required && email.validation === 'email');
  assert.ok(!workspace.required && workspace.validation === 'email');
  assert.strictEqual(software.type, 'CHECKBOX');
  assert.strictEqual(software.choices.length, 6);
  assert.ok(software.help.includes('・HiTeach 專業版：剩餘 15／15'));

  assert.strictEqual(form.destination, 'ss-id');
  assert.strictEqual(sheet.name, '表單回應 2', 'responses go to a new tab of the same spreadsheet');
  assert.strictEqual(E.props.ENROLL_FORM_ID, form.id);
  assert.strictEqual(E.props.ENROLL_RESPONSE_SHEET_ID, String(sheet.id));
  assert.strictEqual(lic.get(1, 4), '軟體');
  assert.strictEqual(E.triggers.length, 1, 'trigger must not be duplicated');

  // running it again syncs choices instead of creating yet another form
  E.ss.getSheetByName('軟體清單').appendRow(['Canva 教育版', 3, '帳號', '']);
  E.call('setupForm()');
  assert.strictEqual(E.createdForms.length, 1);
  assert.strictEqual(software.choices.length, 7);
  assert.ok(software.help.includes('・Canva 教育版：剩餘 3／3'));
});

// ============ 3. no duplicates from the old form ============
ok('responses from the old form are ignored (before and after the new form exists)', () => {
  const S = schoolEnv();
  S.E.call('setupSheets()');
  const oldAnswers = { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的職務': '教師', '您要申請的軟體授權': 'Hiteach' };
  submitTo(S.E, S.oldSheet, oldAnswers, t(1)); // code pushed, form not created yet
  assert.strictEqual(regRows(S.E).length, 0);
  assert.strictEqual(S.E.mails.length, 0, 'no error mail either');

  const R = readyEnv();
  submitTo(R.E, R.oldSheet, oldAnswers, t(1));
  // even an old-form response that happens to carry valid titles must not register
  submitTo(R.E, R.oldSheet, { '您的大名': '乙', '您要申請的軟體授權': 'HiTeach 專業版' }, t(2));
  assert.strictEqual(regRows(R.E).length, 0);
  assert.strictEqual(R.E.mails.length, 0);
  assert.strictEqual(R.lic.get(3, 2), '', 'no serial handed out');
  R.send(3, teacher(1, '一年級'), 'HiTeach 專業版');
  assert.strictEqual(regRows(R.E).length, 1);
});

// ============ 4. end-to-end registration ============
ok('registration: role branches, ranks, waitlist, duplicates, receipts with CC, remaining seats', () => {
  const { E, form, send } = readyEnv();
  for (let i = 1; i <= 5; i++) send(i, teacher(i, '一年級'), 'Kahoot 教師版, HiTeach 專業版');
  send(6, teacher(6, '二年級', { workspace: 't6' + SCHOOL }), 'Kahoot 教師版, Google AI Pro 教育版');
  send(7, teacher(7, '二年級'), 'Google AI Pro 教育版');
  send(8, { name: '科任8', email: 't8@gmail.com', role: '科任', detail: '自然', workspace: 't8' + SCHOOL }, 'Google AI Pro 教育版');
  send(9, { name: '科任9', email: 't9@gmail.com', role: '科任', detail: '社會', workspace: 't9' + SCHOOL }, 'Google AI Pro 教育版');
  send(10, { name: '行政10', email: 't10@gmail.com', role: '行政', detail: '教務處', workspace: 't10' + SCHOOL }, 'Google AI Pro 教育版, Wordwall 專業版');
  send(11, teacher(1, '一年級', { email: 'T1@gmail.com' }), 'Kahoot 教師版, Padlet 學校版'); // duplicate Kahoot

  const rows = regRows(E);
  const find = (email, sw) => rows.find((r) => String(r[2]).toLowerCase() === email && r[4] === sw);
  assert.strictEqual(rows.length, 10 + 2 + 1 + 1 + 1 + 2 + 1);
  assert.deepStrictEqual(find('t5@gmail.com', 'Kahoot 教師版').slice(5, 7), ['正取', 5]);
  assert.deepStrictEqual(find('t6@gmail.com', 'Kahoot 教師版').slice(5, 7), ['備取', 1]);
  assert.strictEqual(find('t1@gmail.com', 'HiTeach 專業版')[3], '學年／一年級');
  assert.strictEqual(find('t8@gmail.com', 'Google AI Pro 教育版')[3], '科任／自然');
  assert.strictEqual(find('t10@gmail.com', 'Wordwall 專業版')[3], '行政／教務處');
  // Google AI Pro: one per grade, one for all subject teachers, none for 行政
  assert.strictEqual(find('t6@gmail.com', 'Google AI Pro 教育版')[5], '正取');
  assert.strictEqual(find('t7@gmail.com', 'Google AI Pro 教育版')[5], '備取');
  assert.strictEqual(find('t8@gmail.com', 'Google AI Pro 教育版')[5], '正取');
  assert.strictEqual(find('t9@gmail.com', 'Google AI Pro 教育版')[5], '備取', 'second 科任 waits even with another subject');
  assert.strictEqual(find('t10@gmail.com', 'Google AI Pro 教育版')[5], '不符資格');
  assert.strictEqual(find('t10@gmail.com', 'Wordwall 專業版')[5], '正取');
  // school account column + notes
  assert.strictEqual(find('t6@gmail.com', 'Google AI Pro 教育版')[11], 't6' + SCHOOL);
  assert.strictEqual(find('t6@gmail.com', 'Google AI Pro 教育版')[10], '');
  assert.strictEqual(find('t7@gmail.com', 'Google AI Pro 教育版')[10], '未填學校帳號');
  assert.strictEqual(rows.filter((r) => String(r[2]).toLowerCase() === 't1@gmail.com' && r[4] === 'Kahoot 教師版').length, 1);
  assert.ok(rows.every((r) => r[7] === false && r[8] === ''), 'no approval, no serial at registration');

  // receipts go to the Google account, CC the school account when one was given
  const receipts = E.mails.filter((m) => m.subject === '【已登記】教學軟體授權申請');
  assert.strictEqual(receipts.length, 11);
  assert.ok(receipts.every((m) => !/HT-/.test(m.htmlBody)));
  assert.deepStrictEqual([receipts[0].to, receipts[0].cc], ['t1@gmail.com', undefined]);
  assert.deepStrictEqual([receipts[5].to, receipts[5].cc], ['t6@gmail.com', 't6' + SCHOOL]);
  assert.ok(receipts[5].htmlBody.includes('備取（備取第 1 位）'));
  assert.ok(receipts[6].htmlBody.includes('未填寫學校帳號'));
  assert.ok(!receipts[5].htmlBody.includes('未填寫學校帳號'));
  assert.ok(receipts[9].htmlBody.includes('不符申請資格') && receipts[9].cc === 't10' + SCHOOL);
  assert.ok(receipts[10].htmlBody.includes('您先前已登記過') && receipts[10].htmlBody.includes('<b>Padlet 學校版</b>：正取（第 1 位）'));
  const adminMails = E.mails.filter((m) => m.to === 'lres712@mail.lres.tc.edu.tw');
  assert.strictEqual(adminMails.length, 2);
  assert.ok(adminMails[0].subject.includes('Kahoot 教師版') && adminMails[1].subject.includes('Google AI Pro 教育版'));

  const help = form.items.find((i) => i.title === Q.software).help;
  assert.ok(help.includes('・Kahoot 教師版：已額滿（備取 1 人）'), help);
  assert.ok(help.includes('・HiTeach 專業版：剩餘 10／15'), help);
  assert.ok(help.includes('・Google AI Pro 教育版：剩餘 5／7｜限各學年、科任各一位，尚可申請：一年級、三年級、四年級、五年級、六年級'), help);
  const srow = E.ss.getSheetByName('申請統計').getRange(1, 1, 7, 10).getValues().find((r) => r[0] === 'Kahoot 教師版');
  assert.deepStrictEqual(srow.slice(1, 8), [5, 6, 5, 1, 0, 0, 0]);
});

// ============ 5. approval + issuing ============
ok('issuing: approved rows only, HiTeach serials, account wording, one mail per person with CC, cancel promotes', () => {
  const { E, lic, ww, send } = readyEnv();
  send(1, { name: '甲', email: 'a@gmail.com', role: '學年', detail: '一年級', workspace: 'a' + SCHOOL }, 'HiTeach 專業版, Wordwall 專業版, Google AI Pro 教育版');
  send(2, { name: '乙', email: 'b@gmail.com', role: '科任', detail: '自然' }, 'HiTeach 專業版, Kahoot 教師版');
  send(3, { name: '丙', email: 'c@gmail.com', role: '科任', detail: '社會' }, 'HiTeach 專業版');
  send(4, { name: '丁', email: 'd@gmail.com', role: '行政', detail: '總務處' }, 'HiTeach 專業版');
  const before = E.mails.length;

  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, before);
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可發放的名單'));

  E.call('menuApproveAllAccepted()');
  assert.ok(regRows(E).every((r) => r[7] === true));
  E.call('menuPreviewIssue()');
  assert.strictEqual(E.mails.length, before, 'preview must not send');

  E.call('menuIssueApproved()');
  let rows = regRows(E);
  const by = (email, sw) => rows.find((r) => r[2] === email && r[4] === sw);
  assert.deepStrictEqual([by('a@gmail.com', 'HiTeach 專業版')[5], by('a@gmail.com', 'HiTeach 專業版')[8]], ['已發放', 'HT-A']);
  assert.deepStrictEqual([by('b@gmail.com', 'HiTeach 專業版')[5], by('b@gmail.com', 'HiTeach 專業版')[8]], ['已發放', 'HT-B']);
  assert.deepStrictEqual([by('c@gmail.com', 'HiTeach 專業版')[5], by('c@gmail.com', 'HiTeach 專業版')[8]], ['正取', '']); // out of serials
  assert.deepStrictEqual([by('b@gmail.com', 'Kahoot 教師版')[5], by('b@gmail.com', 'Kahoot 教師版')[8]], ['已發放', ''], 'account mode needs no serial');
  assert.strictEqual(by('a@gmail.com', 'Wordwall 專業版')[5], '已發放');
  assert.strictEqual(by('a@gmail.com', 'Google AI Pro 教育版')[5], '已發放');
  assert.deepStrictEqual([lic.get(2, 1), lic.get(2, 2)], ['HT-OLD', 'old@x.tw']);
  assert.deepStrictEqual([lic.get(3, 2), lic.get(4, 2)], ['a@gmail.com', 'b@gmail.com']);
  assert.deepStrictEqual(ww.getRange(2, 1, 1, 2).getValues()[0], ['甲', 'a@gmail.com']);

  const issued = E.mails.slice(before).filter((m) => m.subject === '【授權核准】教學軟體授權已發放');
  assert.strictEqual(issued.length, 2, 'one mail per person');
  const mailA = issued.find((m) => m.to === 'a@gmail.com');
  assert.strictEqual(mailA.cc, 'a' + SCHOOL, 'school account is CCed');
  assert.ok(mailA.htmlBody.includes('HT-A') && !mailA.htmlBody.includes('HT-B'));
  assert.ok(mailA.htmlBody.includes('<b>Google AI Pro 教育版</b>－授權將開通於您的學校帳號（a' + SCHOOL + '）'));
  assert.ok(mailA.htmlBody.includes('<b>Wordwall 專業版</b>－授權將開通於您登記的 Google 帳號（a@gmail.com）'));
  const mailB = issued.find((m) => m.to === 'b@gmail.com');
  assert.strictEqual(mailB.cc, undefined);
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可用序號'));

  const n = E.mails.length;
  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, n, 're-running sends nothing new');

  const reg = E.ss.getSheetByName('申請登記');
  const cRow = rows.findIndex((r) => r[2] === 'c@gmail.com') + 2;
  reg.getRange(cRow, 6).setValue('取消');
  E.call('menuRerank()');
  rows = regRows(E);
  assert.strictEqual(rows[cRow - 2][5], '取消');
  assert.deepStrictEqual(rows.find((r) => r[2] === 'd@gmail.com').slice(5, 7), ['正取', 1]);
});

ok('mail quota counts CC addresses and blocks before any write', () => {
  const { E, lic, send } = readyEnv({ mailQuota: 1 });
  send(1, { name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然', workspace: 'a' + SCHOOL }, 'HiTeach 專業版');
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  assert.ok(E.alerts[E.alerts.length - 1].includes('含副本共 2 個收件人'));
  assert.strictEqual(lic.get(3, 2), '');
  assert.strictEqual(regRows(E)[0][5], '正取');
});

ok('a failed notification mail is recorded, the licence stays issued', () => {
  const { E, send } = readyEnv({ failMailTo: 'a@gmail.com' });
  send(1, { name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然' }, 'HiTeach 專業版');
  assert.strictEqual(regRows(E).length, 1);
  const warn = E.mails[E.mails.length - 1];
  assert.ok(warn.subject.includes('回條未寄出') && !warn.subject.includes('【錯誤】'), warn.subject);
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  const row = regRows(E)[0];
  assert.deepStrictEqual([row[5], row[8], row[10]], ['已發放', 'HT-A', '通知信寄送失敗，請人工通知']);
});

// ============ 6. failure paths ============
ok('renamed question: admin is mailed, nothing registered, backfill recovers from the new tab only', () => {
  const { E, sheet, oldSheet } = readyEnv();
  const a = answersOf({ name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然' }, 'Kahoot 教師版');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map((h) => (h === '時間戳記' ? t(1) : (a[h] === undefined ? '' : a[h]))));
  const broken = Object.assign({}, a); delete broken[Q.role]; broken['職務'] = '科任';
  const nv = {}; Object.keys(broken).forEach((k) => { nv[k] = [broken[k]]; });
  E.call('onFormSubmit(__evt)', { __evt: { namedValues: nv, range: sheet.getRange(2, 1, 1, headers.length) } });
  assert.strictEqual(regRows(E).length, 0);
  const err = E.mails[E.mails.length - 1];
  assert.ok(err.subject.includes('【錯誤】') && err.body.includes(Q.role));
  const n = E.mails.length;
  E.call('menuBackfill()');
  const rows = regRows(E);
  assert.strictEqual(rows.length, 1, 'the old tab\'s response is never imported');
  assert.deepStrictEqual([rows[0][1], rows[0][3], rows[0][4], rows[0][5]], ['甲', '科任／自然', 'Kahoot 教師版', '正取']);
  assert.strictEqual(E.mails.length, n, 'backfill sends no mail');
  E.call('menuBackfill()');
  assert.strictEqual(regRows(E).length, 1, 'backfill is idempotent');
  assert.strictEqual(oldSheet.getLastRow(), 2);
});

ok('response tab that appears late is recognised on the first submission', () => {
  const S = schoolEnv({ lateResponseSheet: true });
  S.E.call('setupSheets()');
  S.E.call('setupForm()');
  assert.ok(S.E.props.ENROLL_FORM_ID && S.E.props.ENROLL_RESPONSE_SHEET_ID === undefined);
  const sheet = S.E.env.__pendingLink();
  submitTo(S.E, S.oldSheet, { '您的大名': '舊', '您要申請的軟體授權': 'HiTeach 專業版' }, t(1));
  assert.strictEqual(regRows(S.E).length, 0);
  submitTo(S.E, sheet, answersOf(teacher(1, '三年級'), 'HiTeach 專業版'), t(2));
  assert.strictEqual(regRows(S.E).length, 1);
  assert.strictEqual(S.E.props.ENROLL_RESPONSE_SHEET_ID, String(sheet.id));
});

ok('manual run / no selection / bad account / no UI / reset are all safe', () => {
  const { E, send } = readyEnv();
  E.call('onFormSubmit()');
  send(1, teacher(1, '一年級'), '');
  assert.strictEqual(regRows(E).length, 0);
  send(2, teacher(2, '一年級', { email: 'not-an-email' }), 'Padlet 學校版');
  let rows = regRows(E);
  assert.strictEqual(rows.length, 1);
  assert.ok(rows[0][10].includes('Google 帳號格式有誤'));
  assert.ok(E.mails.every((m) => m.to !== 'not-an-email'));

  const N = readyEnv({ noUi: true });
  N.send(1, teacher(1, '一年級'), 'HiTeach 專業版');
  N.E.ss.getSheetByName('申請登記').getRange(2, 8).setValue(true);
  const n = N.E.mails.length;
  N.E.call('menuIssueApproved()');
  assert.strictEqual(N.E.mails.length, n, 'no UI -> no confirmation -> nothing issued');
  assert.strictEqual(regRows(N.E)[0][5], '正取');

  E.call('resetEnrollmentForm()');
  send(3, teacher(3, '一年級'), 'Padlet 學校版');
  assert.strictEqual(regRows(E).length, 1, 'after reset nothing registers until a form is bound again');
});

ok('quota edited in 軟體清單 takes effect on the next submission', () => {
  const { E, send } = readyEnv();
  E.ss.getSheetByName('軟體清單').getRange(3, 2).setValue(1); // Kahoot -> 1
  send(1, teacher(1, '一年級'), 'Kahoot 教師版');
  send(2, teacher(2, '一年級'), 'Kahoot 教師版');
  assert.deepStrictEqual(regRows(E).map((r) => r[5]), ['正取', '備取']);
});

console.log('\n' + passed + ' scenario groups passed');
