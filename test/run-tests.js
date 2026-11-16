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
  constructor(name, formUrl) { this.name = name; this.cells = []; this.formUrl = formUrl || null; }
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
  insertSheet(n, formUrl) { const s = new Sheet(n, formUrl); this.sheets.push(s); return s; }
  getSheets() { return this.sheets; }
  getFormUrl() { const s = this.sheets.find((x) => x.formUrl); return s ? s.formUrl : null; }
  getId() { return 'ss-id'; }
}

// ---------- fake form ----------
const ItemType = { TEXT: 'TEXT', PARAGRAPH_TEXT: 'PARAGRAPH_TEXT', MULTIPLE_CHOICE: 'MULTIPLE_CHOICE', CHECKBOX: 'CHECKBOX', LIST: 'LIST' };
class Item {
  constructor(form, type, title) { Object.assign(this, { form, type, title: title || '', choices: [], required: false, help: '' }); }
  getTitle() { return this.title; } setTitle(t) { this.title = t; return this; }
  getType() { return this.type; }
  getIndex() { return this.form.items.indexOf(this); }
  setHelpText(t) { this.help = t; return this; }
  setRequired(b) { this.required = b; return this; }
  setChoiceValues(v) { this.choices = v.slice(); return this; }
  setValidation(v) { this.validation = v; return this; }
  asTextItem() { assert.strictEqual(this.type, ItemType.TEXT); return this; }
  asParagraphTextItem() { assert.strictEqual(this.type, ItemType.PARAGRAPH_TEXT); return this; }
  asCheckboxItem() { assert.strictEqual(this.type, ItemType.CHECKBOX); return this; }
  asMultipleChoiceItem() { assert.strictEqual(this.type, ItemType.MULTIPLE_CHOICE); return this; }
  asListItem() { assert.strictEqual(this.type, ItemType.LIST); return this; }
}
class Form {
  constructor(title) { this.title = title; this.items = []; this.description = ''; }
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
  setDestination(type, id) { this.destination = id; return this; }
  getPublishedUrl() { return 'https://forms/published'; }
  getEditUrl() { return 'https://forms/edit'; }
}

function makeEnv(opts) {
  opts = opts || {};
  const ss = new Spreadsheet();
  const mails = [];
  const alerts = [];
  const forms = {};
  const props = {};
  const triggers = [];
  const env = {
    console, JSON, Date, Math, Object, Array, String, Number, Error, isNaN,
    Logger: { log: () => {} },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 16).replace('T', ' ') },
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
      create: (title) => { const f = new Form(title); forms['https://forms/new'] = f; env.__createdForm = f; return f; },
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
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    ScriptApp: {
      EventType: { ON_FORM_SUBMIT: 'ON_FORM_SUBMIT' },
      getProjectTriggers: () => triggers,
      newTrigger: (fn) => { const b = { forSpreadsheet: () => b, onFormSubmit: () => b, create: () => triggers.push({ getHandlerFunction: () => fn, getEventType: () => 'ON_FORM_SUBMIT' }) }; return b; },
    },
  };
  const ctx = vm.createContext(env);
  files.forEach((f) => vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f }));
  const call = (expr, vars) => { Object.assign(env, vars || {}); return vm.runInContext(expr, ctx); };
  return { ss, mails, alerts, forms, props, triggers, env, call };
}

// submit a form response: append to the response sheet, then fire the trigger
function submit(E, respSheetName, answers, when) {
  const sheet = E.ss.getSheetByName(respSheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map((h) => (h === '時間戳記' ? when : (answers[h] === undefined ? '' : answers[h])));
  sheet.appendRow(row);
  const r = sheet.getLastRow();
  const namedValues = {};
  headers.forEach((h, i) => { (namedValues[h] = namedValues[h] || []).push(h === '時間戳記' ? when.toISOString() : String(row[i])); });
  E.call('onFormSubmit(__evt)', { __evt: { namedValues, range: sheet.getRange(r, 1, 1, headers.length) } });
}
const regRows = (E) => { const s = E.ss.getSheetByName('申請登記'); const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, 11).getValues(); };
const t = (min) => new Date(2026, 10, 17, 8, min, 0);
let passed = 0;
const ok = (name, fn) => { fn(); passed++; console.log('ok   ' + name); };

// ============ 1. unit tests shipped to the user ============
ok('runEnrollmentTests (shipped unit tests)', () => {
  const E = makeEnv();
  const res = E.call('runEnrollmentTests()');
  assert.ok(res.length >= 30 && res.every((r) => r.ok));
  console.log('     ' + res.length + ' checks');
});

// ============ 2. setup on an existing form (old radio software question) ============
const H = ['時間戳記', '您的大名', '您的常用email信箱', '您的身分（學年／科任）', '您要申請的軟體授權'];
function existingFormEnv(opts) {
  const E = makeEnv(opts);
  const form = new Form('舊表單');
  E.forms['https://forms/old'] = form;
  form.addTextItem().setTitle('您的大名');
  form.addTextItem().setTitle('您的常用email信箱');
  form.addMultipleChoiceItem().setTitle('您要申請的軟體授權').setChoiceValues(['Hiteach', 'Wordwall']);
  form.addTextItem().setTitle('備註');
  const resp = E.ss.insertSheet('表單回應 1', 'https://forms/old');
  resp.getRange(1, 1, 1, H.length).setValues([H]);
  // last year's response must never be imported
  resp.appendRow([new Date(2025, 8, 1), '去年', 'old@x.tw', '', 'Hiteach']);
  const lic = E.ss.insertSheet('序號列表');
  lic.getRange(1, 1, 1, 3).setValues([['序號', '使用者Email', '分配時間']]);
  lic.appendRow(['HT-OLD', 'old@x.tw', new Date(2025, 8, 1)]);
  lic.appendRow(['HT-A', '', '']);
  lic.appendRow(['HT-B', '', '']);
  const ww = E.ss.insertSheet('wordwall授權確認');
  ww.getRange(1, 1, 1, 3).setValues([['姓名', '帳號', '狀態']]);
  E.call('setupSheets()');
  E.call('setupForm()');
  E.call('createFormSubmitTrigger()');
  E.call('createFormSubmitTrigger()');
  return { E, form, resp, lic, ww };
}

ok('setupSheets + setupForm reuse and repair the existing form', () => {
  const { E, form, lic } = existingFormEnv();
  assert.deepStrictEqual(form.items.map((i) => i.title), ['您的大名', '您的常用email信箱', '您的身分（學年／科任）', '您要申請的軟體授權', '備註']);
  const sw = form.items[3];
  assert.strictEqual(sw.type, 'CHECKBOX');
  assert.strictEqual(sw.choices.length, 6);
  assert.ok(sw.help.includes('・HiTeach 專業版：剩餘 15／15'));
  assert.strictEqual(form.items[2].type, 'MULTIPLE_CHOICE');
  assert.strictEqual(form.items[2].choices.length, 8);
  assert.strictEqual(form.items[1].validation, 'email');
  assert.strictEqual(lic.get(1, 4), '軟體');
  assert.strictEqual(E.triggers.length, 1, 'trigger must not be duplicated');
  assert.ok(E.props.ENROLL_REG_START);
  assert.strictEqual(E.ss.getSheetByName('軟體清單').getLastRow(), 7);
  assert.strictEqual(E.env.__createdForm, undefined, 'must not create a second form');
});

ok('setupForm creates and links a new form when none is linked', () => {
  const E = makeEnv();
  E.call('setupSheets()');
  E.call('setupForm()');
  const f = E.env.__createdForm;
  assert.ok(f && f.destination === 'ss-id');
  assert.deepStrictEqual(f.items.map((i) => i.title), ['您的大名', '您的常用email信箱', '您的身分（學年／科任）', '您要申請的軟體授權']);
});

// ============ 3. end-to-end registration ============
ok('registration: ranks, waitlist, duplicates, receipts, form remaining, admin notice', () => {
  const { E, form } = existingFormEnv();
  const Q = { n: '您的大名', e: '您的常用email信箱', g: '您的身分（學年／科任）', s: '您要申請的軟體授權' };
  const send = (i, name, email, group, software) => submit(E, '表單回應 1', { [Q.n]: name, [Q.e]: email, [Q.g]: group, [Q.s]: software }, t(i));

  for (let i = 1; i <= 5; i++) send(i, '老師' + i, `t${i}@x.tw`, '一年級', 'Kahoot 教師版, HiTeach 專業版');
  send(6, '老師6', 't6@x.tw', '二年級', 'Kahoot 教師版, Google AI Pro 教育版');
  send(7, '老師7', 't7@x.tw', '二年級', 'Google AI Pro 教育版');
  send(8, '行政8', 't8@x.tw', '行政／其他', 'Google AI Pro 教育版, Wordwall 專業版');
  send(9, '老師1', 'T1@x.tw', '一年級', 'Kahoot 教師版, Padlet 學校版'); // duplicate Kahoot

  const rows = regRows(E);
  const find = (email, sw) => rows.find((r) => String(r[2]).toLowerCase() === email && r[4] === sw);
  assert.strictEqual(rows.length, 5 * 2 + 2 + 1 + 2 + 1);
  assert.deepStrictEqual([find('t5@x.tw', 'Kahoot 教師版')[5], find('t5@x.tw', 'Kahoot 教師版')[6]], ['正取', 5]);
  assert.deepStrictEqual([find('t6@x.tw', 'Kahoot 教師版')[5], find('t6@x.tw', 'Kahoot 教師版')[6]], ['備取', 1]);
  assert.strictEqual(find('t6@x.tw', 'Google AI Pro 教育版')[5], '正取');
  assert.strictEqual(find('t7@x.tw', 'Google AI Pro 教育版')[5], '備取');
  assert.strictEqual(find('t8@x.tw', 'Google AI Pro 教育版')[5], '不符資格');
  assert.strictEqual(rows.filter((r) => String(r[2]).toLowerCase() === 't1@x.tw' && r[4] === 'Kahoot 教師版').length, 1);
  assert.ok(rows.every((r) => r[7] === false && r[8] === ''), 'no approval, no serial at registration');
  assert.ok(rows.every((r) => r[0] instanceof Date));

  // receipts: one per submission, none carries a serial
  const receipts = E.mails.filter((m) => m.subject === '【已登記】教學軟體授權申請');
  assert.strictEqual(receipts.length, 9);
  assert.ok(receipts.every((m) => !/HT-/.test(m.htmlBody)));
  assert.ok(receipts[5].htmlBody.includes('備取（備取第 1 位）'));
  assert.ok(receipts[7].htmlBody.includes('不符申請資格'));
  assert.ok(receipts[8].htmlBody.includes('您先前已登記過') && receipts[8].htmlBody.includes('<b>Padlet 學校版</b>：正取（第 1 位）'));
  // admin told once per software when it first overflows
  const adminMails = E.mails.filter((m) => m.to === 'lres712@mail.lres.tc.edu.tw');
  assert.strictEqual(adminMails.length, 2);
  assert.ok(adminMails[0].subject.includes('Kahoot 教師版') && adminMails[1].subject.includes('Google AI Pro 教育版'));

  // form shows remaining
  const help = form.items[3].help;
  assert.ok(help.includes('・Kahoot 教師版：已額滿（備取 1 人）'), help);
  assert.ok(help.includes('・HiTeach 專業版：剩餘 10／15'), help);
  assert.ok(help.includes('・Google AI Pro 教育版：剩餘 6／7'), help);
  assert.ok(help.includes('・Padlet 學校版：剩餘 9／10'), help);

  // stats sheet
  const stats = E.ss.getSheetByName('申請統計');
  const srow = stats.getRange(1, 1, 7, 10).getValues().find((r) => r[0] === 'Kahoot 教師版');
  assert.deepStrictEqual(srow.slice(1, 8), [5, 6, 5, 1, 0, 0, 0]);
});

// ============ 4. approval + issuing ============
ok('issuing: only approved rows, serial by software, one mail per person, cancel promotes waitlist', () => {
  const { E, lic, ww } = existingFormEnv();
  const Q = { n: '您的大名', e: '您的常用email信箱', g: '您的身分（學年／科任）', s: '您要申請的軟體授權' };
  const send = (i, name, email, group, software) => submit(E, '表單回應 1', { [Q.n]: name, [Q.e]: email, [Q.g]: group, [Q.s]: software }, t(i));
  send(1, '甲', 'a@x.tw', '一年級', 'HiTeach 專業版, Wordwall 專業版, Google AI Pro 教育版');
  send(2, '乙', 'b@x.tw', '科任', 'HiTeach 專業版');
  send(3, '丙', 'c@x.tw', '科任', 'HiTeach 專業版');
  send(4, '丁', 'd@x.tw', '三年級', 'HiTeach 專業版');
  const before = E.mails.length;

  // nothing approved yet -> nothing issued
  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, before);
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可發放的名單'));

  E.call('menuApproveAllAccepted()');
  let rows = regRows(E);
  assert.ok(rows.every((r) => r[7] === true));
  E.call('menuPreviewIssue()');
  assert.strictEqual(E.mails.length, before, 'preview must not send');
  assert.ok(E.alerts[E.alerts.length - 1].includes('【預覽，尚未發放】'));

  E.call('menuIssueApproved()');
  rows = regRows(E);
  const by = (email, sw) => rows.find((r) => r[2] === email && r[4] === sw);
  assert.deepStrictEqual([by('a@x.tw', 'HiTeach 專業版')[5], by('a@x.tw', 'HiTeach 專業版')[8]], ['已發放', 'HT-A']);
  assert.deepStrictEqual([by('b@x.tw', 'HiTeach 專業版')[5], by('b@x.tw', 'HiTeach 專業版')[8]], ['已發放', 'HT-B']);
  assert.deepStrictEqual([by('c@x.tw', 'HiTeach 專業版')[5], by('c@x.tw', 'HiTeach 專業版')[8]], ['正取', '']); // out of serials
  assert.strictEqual(by('a@x.tw', 'Wordwall 專業版')[5], '已發放');
  assert.strictEqual(by('a@x.tw', 'Google AI Pro 教育版')[5], '已發放');
  assert.ok(by('a@x.tw', 'HiTeach 專業版')[9] instanceof Date);
  // serial sheet: old row untouched, new rows stamped
  assert.deepStrictEqual([lic.get(2, 1), lic.get(2, 2)], ['HT-OLD', 'old@x.tw']);
  assert.deepStrictEqual([lic.get(3, 2), lic.get(4, 2)], ['a@x.tw', 'b@x.tw']);
  // wordwall confirm sheet synced once
  assert.deepStrictEqual(ww.getRange(2, 1, 1, 2).getValues()[0], ['甲', 'a@x.tw']);
  const issued = E.mails.slice(before).filter((m) => m.subject === '【授權核准】教學軟體授權已發放');
  assert.strictEqual(issued.length, 2, 'one mail per person');
  const mailA = issued.find((m) => m.to === 'a@x.tw').htmlBody;
  assert.ok(mailA.includes('HT-A') && mailA.includes('Wordwall 專業版') && mailA.includes('Google AI Pro 教育版'));
  assert.ok(!mailA.includes('HT-B'));
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可用序號'));

  // re-running issues nothing new and does not re-send
  const n = E.mails.length;
  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, n);
  assert.strictEqual(lic.get(3, 2), 'a@x.tw');

  // cancel + rerank
  const reg = E.ss.getSheetByName('申請登記');
  const cRow = rows.findIndex((r) => r[2] === 'c@x.tw') + 2;
  reg.getRange(cRow, 6).setValue('取消');
  E.call('menuRerank()');
  rows = regRows(E);
  assert.strictEqual(rows[cRow - 2][5], '取消');
  assert.deepStrictEqual([by2(rows, 'd@x.tw')[5], by2(rows, 'd@x.tw')[6]], ['正取', 1]);
  function by2(rs, email) { return rs.find((r) => r[2] === email); }
});

ok('issuing stops before any write when the mail quota is too small', () => {
  const { E, lic } = existingFormEnv({ mailQuota: 0 });
  submit(E, '表單回應 1', { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': 'HiTeach 專業版' }, t(1));
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  assert.strictEqual(lic.get(3, 2), '');
  assert.strictEqual(regRows(E)[0][5], '正取');
});

ok('a failed notification mail is recorded, the licence stays issued', () => {
  const { E } = existingFormEnv({ failMailTo: 'a@x.tw' });
  // receipt mail to a@x.tw also fails -> onFormSubmit must still have registered the row
  submit(E, '表單回應 1', { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': 'HiTeach 專業版' }, t(1));
  assert.strictEqual(regRows(E).length, 1);
  const warn = E.mails[E.mails.length - 1];
  assert.ok(warn.subject.includes('回條未寄出') && !warn.subject.includes('【錯誤】'), warn.subject);
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  const row = regRows(E)[0];
  assert.deepStrictEqual([row[5], row[8], row[10]], ['已發放', 'HT-A', '通知信寄送失敗，請人工通知']);
});

// ============ 5. failure paths ============
ok('missing question title: admin is mailed, nothing registered, backfill recovers it', () => {
  const { E, resp } = existingFormEnv();
  // simulate a renamed question on the form: trigger gets a different title
  resp.appendRow([t(1), '甲', 'a@x.tw', '科任', 'Kahoot 教師版']);
  E.call('onFormSubmit(__evt)', { __evt: { namedValues: { '時間戳記': ['x'], '姓名': ['甲'], '您的常用email信箱': ['a@x.tw'], '您要申請的軟體授權': ['Kahoot 教師版'] }, range: resp.getRange(3, 1, 1, 5) } });
  assert.strictEqual(regRows(E).length, 0);
  const err = E.mails[E.mails.length - 1];
  assert.ok(err.subject.includes('【錯誤】') && err.body.includes('您的大名'));
  const n = E.mails.length;
  E.call('menuBackfill()');
  const rows = regRows(E);
  assert.strictEqual(rows.length, 1, 'only this year\'s response is imported');
  assert.deepStrictEqual([rows[0][1], rows[0][4], rows[0][5]], ['甲', 'Kahoot 教師版', '正取']);
  assert.strictEqual(E.mails.length, n, 'backfill sends no mail');
  E.call('menuBackfill()');
  assert.strictEqual(regRows(E).length, 1, 'backfill is idempotent');
});

ok('manual run / no selection / bad email / no UI are all safe', () => {
  const { E } = existingFormEnv();
  E.call('onFormSubmit()');
  submit(E, '表單回應 1', { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': '' }, t(1));
  assert.strictEqual(regRows(E).length, 0);
  submit(E, '表單回應 1', { '您的大名': '乙', '您的常用email信箱': 'not-an-email', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': 'Padlet 學校版' }, t(2));
  const rows = regRows(E);
  assert.strictEqual(rows.length, 1);
  assert.ok(rows[0][10].includes('Email 格式有誤'));
  assert.ok(E.mails.every((m) => m.to !== 'not-an-email'));

  const N = existingFormEnv({ noUi: true }).E;
  submit(N, '表單回應 1', { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': 'HiTeach 專業版' }, t(1));
  const sheet = N.ss.getSheetByName('申請登記');
  sheet.getRange(2, 8).setValue(true);
  const n = N.mails.length;
  N.call('menuIssueApproved()');
  assert.strictEqual(N.mails.length, n, 'no UI -> no confirmation -> nothing issued');
  assert.strictEqual(regRows(N)[0][5], '正取');
});

ok('quota edited in 軟體清單 takes effect on the next submission', () => {
  const { E } = existingFormEnv();
  E.ss.getSheetByName('軟體清單').getRange(3, 2).setValue(1); // Kahoot -> 1
  submit(E, '表單回應 1', { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': 'Kahoot 教師版' }, t(1));
  submit(E, '表單回應 1', { '您的大名': '乙', '您的常用email信箱': 'b@x.tw', '您的身分（學年／科任）': '科任', '您要申請的軟體授權': 'Kahoot 教師版' }, t(2));
  assert.deepStrictEqual(regRows(E).map((r) => r[5]), ['正取', '備取']);
});

console.log('\n' + passed + ' scenario groups passed');
