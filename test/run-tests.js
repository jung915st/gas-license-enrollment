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
  setName(n) { this.name = n; return this; }
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
// Apps Script has two kinds of item objects: the typed item returned by form.addXxxItem()
// (CheckboxItem, PageBreakItem ...; no asXxxItem() methods) and the generic Item returned by
// form.getItems() (has asXxxItem(), which returns the typed item). The fake keeps that split
// so a call like addPageBreakItem().asPageBreakItem() fails here the way it fails in Apps Script.
const ItemType = { TEXT: 'TEXT', PARAGRAPH_TEXT: 'PARAGRAPH_TEXT', MULTIPLE_CHOICE: 'MULTIPLE_CHOICE', CHECKBOX: 'CHECKBOX', LIST: 'LIST', PAGE_BREAK: 'PAGE_BREAK' };
const AS = { asTextItem: 'TEXT', asParagraphTextItem: 'PARAGRAPH_TEXT', asCheckboxItem: 'CHECKBOX', asMultipleChoiceItem: 'MULTIPLE_CHOICE', asListItem: 'LIST', asPageBreakItem: 'PAGE_BREAK' };
class Item {
  constructor(form, type) { Object.assign(this, { form, type, title: '', choices: [], required: false, help: '' }); }
  getTitle() { return this.title; } setTitle(t) { this.title = t; return this; }
  getType() { return this.type; }
  getIndex() { return this.form.items.indexOf(this); }
  getId() { return this.form.items.indexOf(this) + 1000; }
  setHelpText(t) { this.help = t; return this; }
  setRequired(b) { this.required = b; return this; }
  setChoiceValues(v) { this.choices = Array.from(v); this.nav = undefined; return this; }
  getChoices() { return this.choices.map((value) => ({ getValue: () => value })); }
  createChoice(value, nav) { return { value, nav }; }
  setChoices(list) { this.choices = Array.from(list, (c) => c.value); this.nav = Array.from(list, (c) => c.nav); return this; }
  setGoToPage(page) {
    assert.strictEqual(this.type, ItemType.PAGE_BREAK);
    assert.ok(page === 'SUBMIT' || page === 'CONTINUE' || (page instanceof Item && page.type === ItemType.PAGE_BREAK),
      'setGoToPage needs a PageBreakItem or a PageNavigationType');
    this.goTo = page === 'CONTINUE' ? undefined : page; return this;
  }
  setValidation(v) { this.validation = v; return this; }
  createResponse(values) { assert.strictEqual(this.type, ItemType.CHECKBOX); return { item: this, values: Array.from(values) }; }
  generic() {
    const raw = this;
    const view = { getTitle: () => raw.title, getType: () => raw.type, getIndex: () => raw.getIndex(), getId: () => raw.getId(),
      setTitle: (t) => { raw.title = t; return view; }, setHelpText: (t) => { raw.help = t; return view; } };
    Object.keys(AS).forEach((m) => { view[m] = () => { assert.strictEqual(raw.type, AS[m], `${m} on a ${raw.type} item`); return raw; }; });
    return view;
  }
}
class Form {
  constructor(title, id) { this.title = title; this.items = []; this.description = ''; this.id = id; this.url = 'https://forms/' + id; this.published = false; }
  getId() { return this.id; }
  supportsAdvancedResponderPermissions() { return true; }
  isPublished() { return this.published; }
  setPublished(b) { this.published = b; return this; }
  getItems() { return this.items.map((it) => it.generic()); }
  add(type) { const it = new Item(this, type); this.items.push(it); return it; }
  addTextItem() { return this.add(ItemType.TEXT); }
  addCheckboxItem() { return this.add(ItemType.CHECKBOX); }
  addMultipleChoiceItem() { return this.add(ItemType.MULTIPLE_CHOICE); }
  addPageBreakItem() { return this.add(ItemType.PAGE_BREAK); }
  addListItem() { return this.add(ItemType.LIST); }
  deleteItem(x) {
    const index = typeof x === 'number' ? x : x.getIndex();
    assert.ok(index >= 0 && index < this.items.length, 'deleteItem: no item at ' + index);
    const [gone] = this.items.splice(index, 1);
    // like Forms, a deleted page break can no longer be a jump target
    this.items.forEach((it) => {
      assert.ok(it.goTo !== gone && !(it.nav || []).includes(gone), 'deleted page "' + gone.title + '" is still a jump target of "' + it.title + '"');
    });
  }
  moveItem(from, to) {
    assert.ok(Number.isInteger(from) && Number.isInteger(to), 'moveItem expects indexes');
    const [it] = this.items.splice(from, 1); this.items.splice(to, 0, it); return it;
  }
  setDescription(d) { this.description = d; return this; }
  setDestination(type, id) { this.destination = id; if (this.onDestination) this.onDestination(this); return this; }
  getPublishedUrl() { return this.url + '/viewform'; }
  getEditUrl() { return this.url + '/edit'; }
  createResponse() {
    const parts = [];
    const r = {
      withItemResponse: (resp) => { assert.strictEqual(resp.item.form, this); parts.push('entry.' + resp.item.getId() + '=' + encodeURIComponent(resp.values.join(','))); return r; },
      toPrefilledUrl: () => this.url + '/viewform?usp=pp_url&' + parts.join('&'),
    };
    return r;
  }
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
  const drivePermissions = [];
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
          alert: (a, b, c) => { alerts.push(b === undefined ? a : a + '\n' + b); return env.__confirm || 'YES'; },
          prompt: (title, message) => {
            alerts.push(title + '\n' + message);
            const reply = env.__prompt || { button: 'CANCEL', text: '' };
            return { getSelectedButton: () => reply.button, getResponseText: () => reply.text };
          },
          ButtonSet: { YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' }, Button: { YES: 'YES', OK: 'OK', CANCEL: 'CANCEL' }, createMenu: () => menu,
        };
      },
      newDataValidation: () => { const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; },
    },
    FormApp: {
      ItemType, DestinationType: { SPREADSHEET: 'SPREADSHEET' }, PageNavigationType: { SUBMIT: 'SUBMIT', CONTINUE: 'CONTINUE' },
      openByUrl: (url) => { assert.ok(forms[url], 'unknown form ' + url); return forms[url]; },
      openById: (id) => { const f = Object.values(forms).find((x) => x.id === id); if (!f) throw new Error('no form ' + id); return f; },
      create: (title) => {
        const f = newForm(title);
        createdForms.push(f);
        f.onDestination = (form) => { if (opts.lateResponseSheet) env.__pendingLink = () => linkForm(form); else linkForm(form); };
        return f;
      },
      createTextValidation: () => {
        const v = {};
        const b = { setHelpText: (h) => { v.help = h; return b; }, requireTextIsEmail: () => { v.kind = 'email'; return b; },
          requireTextMatchesPattern: (p) => { v.kind = 'pattern'; v.pattern = p; return b; },
          requireNumberBetween: (min, max) => { v.kind = 'number'; v.min = min; v.max = max; return b; },
          build: () => (v.kind === 'email' ? 'email' : v) };
        return b;
      },
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
    Drive: { Permissions: { create: (resource, fileId) => {
      if (opts.drivePolicyBlocks) throw new Error('The user does not have sufficient permissions for this file.');
      assert.ok(Object.values(forms).some((f) => f.id === fileId), 'permission on an unknown file ' + fileId);
      drivePermissions.push({ resource: Object.assign({}, resource), fileId });
      return { id: 'perm-' + drivePermissions.length };
    } } },
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
  return { ss, mails, alerts, forms, createdForms, props, triggers, env, call, newForm, linkForm, drivePermissions };
}

// submit a form response: append to that form's response tab, then fire the spreadsheet trigger
function submitTo(E, sheet, answers, when) {
  Object.keys(answers).forEach((key) => {
    const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (current.indexOf(key) < 0) sheet.set(1, current.length + 1, key);
  });
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
  subject: '任教科目', office: '所屬處室', teacherSay: 'TeacherSay 專業版（必選）',
  pick: '您要另外申請的軟體（限選一項）', pickOffice: '您要另外申請的軟體（行政，限選一項）',
  workspace: '申請 Google AI Pro 教育版的學校帳號',
  kind: '申請項目', tabletUnit: '您的職務與任教班級（或科目、處室）', tabletCart: '借用的平板車', tabletMode: '借用方式',
  tabletCount: '借用數量（台）',
  // v3.1 / v3.2 questions, only present on old forms
  software: '您要申請的軟體授權', aiPro: '是否申請 Google AI Pro 教育版？',
};
const SCHOOL = '@mail.lres.tc.edu.tw';
const TS = 'TeacherSay 專業版';
const AI = 'Google AI Pro 教育版';
const NONE = '不申請其他軟體（只申請 TeacherSay）';
const OTHERS = ['Wayground (Quizizz)', 'Kahoot 教師版', 'Padlet 學校版', 'Wordwall 專業版', 'HiTeach 專業版'];
// who: {name, email, role:'學年'|'科任'|'行政', detail, workspace}; software: what is requested besides TeacherSay.
// The v3.3 form has one single-choice dropdown, so at most one program; 行政 gets a dropdown without Google AI Pro.
// TeacherSay is ticked on its own required checkbox; opts.noTeacherSay simulates leaving it unticked.
function answersOf(who, software, opts) {
  opts = opts || {};
  const list = String(software || '').split(',').map((x) => x.trim()).filter((x) => x && x !== TS);
  if (list.length > 1) throw new Error('test bug: the form allows one program besides TeacherSay, got ' + list.join(' + '));
  const office = who.role === '行政';
  if (office && list[0] === AI) throw new Error('test bug: the 行政 dropdown has no Google AI Pro');
  const a = {
    [Q.name]: who.name, [Q.email]: who.email, [Q.kind]: '教學軟體授權', [Q.role]: who.role,
    [Q.teacherSay]: opts.noTeacherSay ? '' : TS,
    [office ? Q.pickOffice : Q.pick]: list[0] || NONE,
    [Q.workspace]: list[0] === AI ? (who.workspace || '') : '',
  };
  a[{ 學年: Q.grade, 科任: Q.subject, 行政: Q.office }[who.role]] = who.detail;
  return a;
}
const regRows = (E) => { const s = E.ss.getSheetByName('申請登記'); const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, 12).getValues(); };
const rowOf = (rows, email, sw) => rows.find((r) => String(r[2]).toLowerCase() === email.toLowerCase() && r[4] === sw);
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
  S.send = (i, who, software, o) => submitTo(S.E, S.sheet, answersOf(who, software, o), t(i));
  return S;
}
const teacher = (n, grade, extra) => Object.assign({ name: '老師' + n, email: `t${n}@gmail.com`, role: '學年', detail: grade }, extra || {});
const item = (form, title) => form.items.find((i) => i.title === title);

// Walk the form like a respondent: which sections are shown for these answers, and can the form be submitted?
// Mirrors Google Forms: required questions block "next"; a multiple-choice question with per-choice
// navigation decides where the section goes; otherwise the next page break's goTo, else the next section.
// An answer that is not among a choice question's options cannot be given (reported as blocked).
function walkForm(form, answers) {
  const sections = [[]];
  form.items.forEach((it) => { if (it.type === 'PAGE_BREAK') sections.push([it]); else sections[sections.length - 1].push(it); });
  const visited = [];
  let idx = 0;
  for (let guard = 0; guard < 20; guard++) {
    const sec = sections[idx];
    visited.push(idx === 0 ? '基本資料' : sec[0].title);
    const questions = sec.filter((it) => it.type !== 'PAGE_BREAK');
    for (const q of questions) {
      const v = answers[q.title];
      if (q.required && (v === undefined || v === '')) return { visited, submitted: false, blockedAt: q.title };
      if (v && (q.type === 'LIST' || q.type === 'MULTIPLE_CHOICE') && q.choices.indexOf(v) < 0) {
        return { visited, submitted: false, blockedAt: q.title, notOffered: v };
      }
      if (v && q.validation && q.validation.kind === 'pattern' && !new RegExp(q.validation.pattern).test(v)) {
        return { visited, submitted: false, blockedAt: q.title };
      }
      if (v && q.validation && q.validation.kind === 'number' && !(Number(v) >= q.validation.min && Number(v) <= q.validation.max)) {
        return { visited, submitted: false, blockedAt: q.title };
      }
    }
    let next = null;
    const branching = questions.find((q) => q.nav && q.nav.some((n) => n));
    if (branching) {
      const target = branching.nav[branching.choices.indexOf(answers[branching.title])];
      if (target === 'SUBMIT') return { visited, submitted: true };
      if (target) next = sections.findIndex((x) => x[0] === target);
    }
    if (next === null) {
      if (idx + 1 >= sections.length) return { visited, submitted: true };
      const pb = sections[idx + 1][0];
      if (pb.goTo === 'SUBMIT') return { visited, submitted: true };
      next = pb.goTo ? sections.findIndex((x) => x[0] === pb.goTo) : idx + 1;
    }
    idx = next;
  }
  throw new Error('form navigation loops');
}
const FORM_LAYOUT = [Q.name, Q.email, Q.kind, '#【TeacherSay 專業版】', Q.teacherSay, '#【職務】', Q.role,
  '#【學年】', Q.grade, '#【科任】', Q.subject, '#【行政】', Q.office,
  '#申請軟體授權', Q.pick, '#申請軟體授權（行政）', Q.pickOffice, '#【Google AI Pro 學校帳號】', Q.workspace,
  '#【平板借用】', Q.tabletUnit, Q.tabletCart, Q.tabletMode, '#【平板借用數量】', Q.tabletCount];
const pageOf = (form, title) => form.items.find((i) => i.type === 'PAGE_BREAK' && i.title === title);
// the v3.3 jumps: 學年/科任 -> software; 行政 -> software (行政); AI Pro -> school account; everything else submits
function assertJumps(form) {
  const p = (t) => pageOf(form, t);
  assert.deepStrictEqual(item(form, Q.role).nav, [p('【學年】'), p('【科任】'), p('【行政】')], 'role jumps');
  assert.strictEqual(p('【科任】').goTo, p('申請軟體授權'), 'after 學年 -> software');
  assert.strictEqual(p('【行政】').goTo, p('申請軟體授權'), 'after 科任 -> software');
  assert.strictEqual(p('申請軟體授權').goTo, p('申請軟體授權（行政）'), 'after 行政 -> software (行政)');
  assert.strictEqual(p('【Google AI Pro 學校帳號】').goTo, 'SUBMIT', 'after the 行政 dropdown -> submit');
  const pick = item(form, Q.pick);
  assert.deepStrictEqual(pick.nav, pick.choices.map((c) => (c === AI ? p('【Google AI Pro 學校帳號】') : 'SUBMIT')), 'dropdown jumps');
  assert.ok(item(form, Q.pickOffice).nav.every((n) => n === 'SUBMIT'));
  assert.deepStrictEqual(item(form, Q.kind).nav, [p('【TeacherSay 專業版】'), p('【平板借用】')], 'kind jumps');
  assert.strictEqual(p('【平板借用】').goTo, 'SUBMIT', 'after the school account -> submit (not into the tablet section)');
  assert.deepStrictEqual(item(form, Q.tabletMode).nav, ['SUBMIT', p('【平板借用數量】')], 'whole cart submits, partial asks the count');
}
const titlesOf = (form) => form.items.map((i) => (i.type === 'PAGE_BREAK' ? '#' + i.title : i.title));

// ============ 1. unit tests shipped to the user ============
ok('runEnrollmentTests (shipped unit tests)', () => {
  const E = makeEnv();
  const res = E.call('runEnrollmentTests()');
  assert.ok(res.length >= 60 && res.every((r) => r.ok));
  console.log('     ' + res.length + ' checks');
});

// ============ 2. a brand-new form, old form untouched ============
ok('setupForm creates a NEW form: TeacherSay gate, role branches, one-pick dropdowns, AI Pro only for 學年/科任, prefilled link', () => {
  const { E, oldForm, oldSheet, form, sheet, lic } = readyEnv();
  assert.strictEqual(E.createdForms.length, 1);
  assert.notStrictEqual(form, oldForm);
  assert.deepStrictEqual(oldForm.items.map((i) => i.title), ['您的大名', '您的常用email信箱', '您的職務', '您要申請的軟體授權']);
  assert.strictEqual(oldSheet.getLastRow(), 2);

  assert.deepStrictEqual(titlesOf(form), FORM_LAYOUT);
  assert.ok(form.description.includes('每次填表最多申請 2 項軟體：已預設申請 TeacherSay 專業版，只能再選填 1 項'));
  assert.ok(form.description.includes('本次軟體使用申請需協助提交使用成果，同時配合取得並使用TeacherSay。'));
  assert.ok(form.description.includes('要更改申請的軟體，請用同一個 Google 帳號重新填寫，以最後一次為準'), form.description);
  assertJumps(form);
  assert.deepStrictEqual(item(form, Q.office).choices, ['教務處', '學務處', '總務處', '輔導室', '幼兒園']);
  assert.ok(pageOf(form, '【行政】').help.includes('由資訊組另行分配'));
  assert.strictEqual(item(form, Q.subject).type, 'TEXT');

  const ts = item(form, Q.teacherSay);
  assert.deepStrictEqual([ts.type, ts.choices, ts.required], ['CHECKBOX', [TS], true]);
  assert.ok(pageOf(form, '【TeacherSay 專業版】').help.includes('TeacherSay'));
  const pick = item(form, Q.pick);
  assert.deepStrictEqual([pick.type, pick.required, pick.choices], ['LIST', true, OTHERS.concat([AI, NONE])]);
  const pickOffice = item(form, Q.pickOffice);
  assert.deepStrictEqual([pickOffice.type, pickOffice.required, pickOffice.choices], ['LIST', true, OTHERS.concat([NONE])]);
  assert.ok(pick.help.includes('除 TeacherSay 外只能再選一項'));
  assert.ok(pick.help.includes('・HiTeach 專業版：剩餘 15／15') && pick.help.includes('・Google AI Pro 教育版：剩餘 7／7'), pick.help);
  assert.ok(!pick.help.includes('TeacherSay 專業版：'), 'TeacherSay is not in the dropdown, so not in its seat list');
  assert.ok(pickOffice.help.includes('・HiTeach 專業版：剩餘 15／15') && !pickOffice.help.includes('Google AI Pro 教育版：'));
  const ws = item(form, Q.workspace);
  assert.deepStrictEqual([ws.type, ws.required, ws.validation.kind], ['TEXT', true, 'pattern']);
  assert.ok(new RegExp(ws.validation.pattern).test('a' + SCHOOL) && !new RegExp(ws.validation.pattern).test('a@gmail.com'));

  // prefilled link ticks TeacherSay
  const prefill = E.props.ENROLL_PREFILL_URL;
  assert.ok(prefill.includes('usp=pp_url') && prefill.includes('entry.' + ts.getId() + '=' + encodeURIComponent(TS)), prefill);
  assert.ok(E.alerts.some((a) => a.includes('已預先勾選') && a.includes(prefill)), 'setupForm shows the prefilled link');
  E.call('menuShowFormUrls()');
  assert.ok(E.alerts[E.alerts.length - 1].includes(prefill));

  assert.strictEqual(form.destination, 'ss-id');
  assert.strictEqual(form.published, true, 'new form is published');
  assert.deepStrictEqual(E.drivePermissions, [{ resource: { type: 'anyone', role: 'reader', view: 'published' }, fileId: form.id }],
    'anyone with the link may respond');
  assert.ok(E.alerts.some((a) => a.includes('作答者權限：知道連結的任何人')));
  assert.strictEqual(sheet.name, '表單回應 2', 'responses go to a new tab of the same spreadsheet');
  assert.strictEqual(E.props.ENROLL_FORM_ID, form.id);
  assert.strictEqual(E.props.ENROLL_RESPONSE_SHEET_ID, String(sheet.id));
  assert.strictEqual(lic.get(1, 4), '軟體');
  assert.strictEqual(E.triggers.length, 1, 'trigger must not be duplicated');

  // running it again updates the same form instead of creating another, and leaves responder access alone
  E.ss.getSheetByName('軟體清單').appendRow(['Canva 教育版', 3, '帳號', '']);
  E.call('setupForm()');
  assert.strictEqual(E.createdForms.length, 1);
  assert.strictEqual(E.drivePermissions.length, 1, 'updating does not touch responder access');
  assert.deepStrictEqual(titlesOf(form), FORM_LAYOUT, 'no duplicated questions');
  assert.deepStrictEqual(pick.choices, OTHERS.concat([AI, 'Canva 教育版', NONE]));
  assert.deepStrictEqual(pickOffice.choices, OTHERS.concat(['Canva 教育版', NONE]));
  assertJumps(form);
  assert.ok(pick.help.includes('・Canva 教育版：剩餘 3／3') && pickOffice.help.includes('・Canva 教育版：剩餘 3／3'));
});

ok('respondent paths: TeacherSay gate first; one program at most; AI Pro via the 學年/科任 dropdown; 行政 never offered it', () => {
  const { form } = readyEnv();
  const T = '【TeacherSay 專業版】', R = '【職務】', G = '【學年】', S = '【科任】', O = '【行政】',
    W = '申請軟體授權', WO = '申請軟體授權（行政）', AW = '【Google AI Pro 學校帳號】';
  const base = (role, detail, sw) => answersOf({ name: '甲', email: 'a@gmail.com', role, detail, workspace: 'a' + SCHOOL }, sw);
  assert.deepStrictEqual(walkForm(form, base('學年', '三年級', 'Kahoot 教師版')), { visited: ['基本資料', T, R, G, W], submitted: true });
  assert.deepStrictEqual(walkForm(form, base('科任', '自然', '')), { visited: ['基本資料', T, R, S, W], submitted: true });
  assert.deepStrictEqual(walkForm(form, base('行政', '幼兒園', 'Wordwall 專業版')), { visited: ['基本資料', T, R, O, WO], submitted: true });
  // TeacherSay unticked: stuck on the TeacherSay page, no software choice is ever shown
  const noTs = Object.assign(base('學年', '一年級', 'Kahoot 教師版'), { [Q.teacherSay]: '' });
  assert.deepStrictEqual(walkForm(form, noTs), { visited: ['基本資料', T], submitted: false, blockedAt: Q.teacherSay });
  // the dropdown must be answered (the "none" option is there for TeacherSay only)
  assert.strictEqual(walkForm(form, Object.assign(base('學年', '一年級', ''), { [Q.pick]: '' })).blockedAt, Q.pick);
  // Google AI Pro from the dropdown: the school-account page follows and must hold a school account
  const aiYes = base('學年', '二年級', AI);
  assert.deepStrictEqual(walkForm(form, aiYes), { visited: ['基本資料', T, R, G, W, AW], submitted: true });
  assert.deepStrictEqual(walkForm(form, base('科任', '社會', AI)).visited, ['基本資料', T, R, S, W, AW]);
  assert.strictEqual(walkForm(form, Object.assign({}, aiYes, { [Q.workspace]: '' })).blockedAt, Q.workspace);
  assert.strictEqual(walkForm(form, Object.assign({}, aiYes, { [Q.workspace]: 't1@gmail.com' })).blockedAt, Q.workspace);
  // 行政 is not offered Google AI Pro
  const officeAi = walkForm(form, Object.assign(base('行政', '教務處', ''), { [Q.pickOffice]: AI }));
  assert.deepStrictEqual([officeAi.blockedAt, officeAi.notOffered], [Q.pickOffice, AI]);
});

// ============ 3. upgrading the v3.1 form in place ============
ok('upgrade a v3.1 form in place: same form and tab, new questions slotted in, old responses still parse', () => {
  const S = schoolEnv();
  // the v3.1 software list (no TeacherSay row) and the v3.1 form layout, already bound
  const swSheet = S.E.ss.insertSheet('軟體清單');
  swSheet.getRange(1, 1, 1, 4).setValues([['軟體名稱', '可用授權數', '分配方式', '申請限制']]);
  [['Wayground (Quizizz)', 5, '帳號', ''], ['Kahoot 教師版', 5, '帳號', ''], ['Padlet 學校版', 10, '帳號', ''],
    ['Wordwall 專業版', 10, '帳號', ''], ['HiTeach 專業版', 15, '序號', ''], [AI, 7, '帳號', '各學年、科任各一位']]
    .forEach((r) => swSheet.appendRow(r));
  const f = S.E.newForm('教學軟體授權申請表');
  f.addTextItem().setTitle(Q.name); f.addTextItem().setTitle(Q.email);
  const role = f.addMultipleChoiceItem().setTitle(Q.role);
  const pG = f.addPageBreakItem().setTitle('【學年】'); f.addMultipleChoiceItem().setTitle(Q.grade);
  const pS = f.addPageBreakItem().setTitle('【科任】'); f.addTextItem().setTitle(Q.subject);
  const pO = f.addPageBreakItem().setTitle('【行政】'); f.addMultipleChoiceItem().setTitle(Q.office).setChoiceValues(['教務處', '學務處', '總務處', '輔導室']);
  const pW = f.addPageBreakItem().setTitle('申請軟體授權');
  f.addCheckboxItem().setTitle(Q.software).setChoiceValues(['Kahoot 教師版', AI]).setRequired(true);
  f.addTextItem().setTitle(Q.workspace).setRequired(false);
  role.setChoices([{ value: '學年', nav: pG }, { value: '科任', nav: pS }, { value: '行政', nav: pO }]);
  pS.setGoToPage(pW); pO.setGoToPage(pW);
  const tab = S.E.linkForm(f);
  S.E.props.ENROLL_FORM_ID = f.id;
  S.E.props.ENROLL_RESPONSE_SHEET_ID = String(tab.id);
  // a response submitted on the v3.1 form before the upgrade (AI Pro inside the checkbox)
  const v31 = { [Q.name]: '舊回應', [Q.email]: 'v31@gmail.com', [Q.role]: '科任', [Q.subject]: '自然',
    [Q.software]: 'Kahoot 教師版, ' + AI, [Q.workspace]: 'v31' + SCHOOL };
  submitTo(S.E, tab, v31, t(0)); // no setupSheets yet -> 申請登記 is created on the fly

  S.E.call('setupSheets()');
  assert.ok(S.E.alerts[S.E.alerts.length - 1].includes('已加入「TeacherSay 專業版」'));
  assert.strictEqual(swSheet.get(8, 1), TS);
  S.E.call('setupSheets()');
  assert.strictEqual(swSheet.getLastRow(), 8, 'TeacherSay row added only once');
  S.E.call('setupForm()');

  assert.strictEqual(S.E.createdForms.length, 0, 'no new form');
  assert.strictEqual(S.E.drivePermissions.length, 0, 'an existing form keeps its responder access');
  assert.strictEqual(S.E.props.ENROLL_RESPONSE_SHEET_ID, String(tab.id), 'same response tab');
  assert.deepStrictEqual(titlesOf(f), FORM_LAYOUT);
  assertJumps(f);
  assert.strictEqual(pageOf(f, '申請軟體授權'), pW, 'the v3.1 software page is reused');
  assert.deepStrictEqual(role.nav, [pG, pS, pO], 'role choices still jump to their sections');
  assert.strictEqual(item(f, Q.software), undefined, 'the old checkbox is deleted');
  assert.deepStrictEqual(item(f, Q.office).choices.slice(-1), ['幼兒園']);
  assert.deepStrictEqual(item(f, Q.pick).choices, OTHERS.concat([AI, NONE]));
  assert.strictEqual(item(f, Q.workspace).required, true);
  const msg = S.E.alerts.find((a) => a.includes('已更新申請表單'));
  assert.ok(msg.includes('新增「' + Q.teacherSay + '」') && msg.includes('刪除舊題「' + Q.software + '」'), msg);
  const walked = walkForm(f, answersOf({ name: '甲', email: 'a@gmail.com', role: '行政', detail: '幼兒園' }, 'Padlet 學校版'));
  assert.deepStrictEqual(walked, { visited: ['基本資料', '【TeacherSay 專業版】', '【職務】', '【行政】', '申請軟體授權（行政）'], submitted: true });

  // the pre-upgrade response was registered by the trigger; a new one goes through the new questions
  let rows = regRows(S.E);
  assert.deepStrictEqual(rows.map((r) => r[4]), ['Kahoot 教師版', AI]);
  assert.ok(rows[1][10].includes('此帳號另登記了：Kahoot 教師版'), 'a v3.1 response with two programs is flagged');
  submitTo(S.E, tab, answersOf({ name: '新', email: 'n@gmail.com', role: '行政', detail: '幼兒園' }, 'Wordwall 專業版'), t(1));
  // a v3.1 response from 行政 that ticked AI Pro in the old checkbox is still refused
  submitTo(S.E, tab, { [Q.name]: '舊行政', [Q.email]: 'o@gmail.com', [Q.role]: '行政', [Q.office]: '教務處',
    [Q.software]: AI, [Q.workspace]: 'o' + SCHOOL }, t(2));
  rows = regRows(S.E);
  assert.deepStrictEqual(rows.slice(2).map((r) => [r[3], r[4], r[5]]),
    [['行政／幼兒園', 'Wordwall 專業版', '正取'], ['行政／幼兒園', TS, '正取'], ['行政／教務處', AI, '不符資格']],
    'order follows 軟體清單, where the upgrade appended TeacherSay at the bottom');
});

ok('upgrade the live v3.2 form in place: checkbox and yes/no question removed, dropdowns added, v3.2 responses still parse', () => {
  const S = schoolEnv();
  S.E.call('setupSheets()'); // the v3.2 software list is the default one
  const f = S.E.newForm('教學軟體授權申請表');
  f.addTextItem().setTitle(Q.name); f.addTextItem().setTitle(Q.email);
  const pT = f.addPageBreakItem().setTitle('【TeacherSay 專業版】');
  const ts = f.addCheckboxItem().setTitle(Q.teacherSay).setChoiceValues([TS]).setRequired(true);
  const pW = f.addPageBreakItem().setTitle('申請軟體授權');
  f.addCheckboxItem().setTitle(Q.software).setChoiceValues(OTHERS);
  const pR = f.addPageBreakItem().setTitle('【職務】');
  const role = f.addMultipleChoiceItem().setTitle(Q.role);
  const pG = f.addPageBreakItem().setTitle('【學年】'); f.addMultipleChoiceItem().setTitle(Q.grade).setChoiceValues(['一年級', '二年級', '三年級', '四年級', '五年級', '六年級']);
  const pS = f.addPageBreakItem().setTitle('【科任】'); f.addTextItem().setTitle(Q.subject);
  const pO = f.addPageBreakItem().setTitle('【行政】'); f.addMultipleChoiceItem().setTitle(Q.office).setChoiceValues(['教務處', '學務處', '總務處', '輔導室', '幼兒園']);
  const pA = f.addPageBreakItem().setTitle('【Google AI Pro 教育版】');
  const aiQ = f.addMultipleChoiceItem().setTitle(Q.aiPro);
  const pWs = f.addPageBreakItem().setTitle('【Google AI Pro 學校帳號】');
  f.addTextItem().setTitle(Q.workspace).setRequired(true);
  role.setChoices([{ value: '學年', nav: pG }, { value: '科任', nav: pS }, { value: '行政', nav: pO }]);
  aiQ.setChoices([{ value: '是', nav: pWs }, { value: '否', nav: 'SUBMIT' }]);
  pS.setGoToPage(pA); pO.setGoToPage(pA); pA.setGoToPage('SUBMIT');
  const tab = S.E.linkForm(f);
  S.E.props.ENROLL_FORM_ID = f.id;
  S.E.props.ENROLL_RESPONSE_SHEET_ID = String(tab.id);

  // submitted on the v3.2 form after the v3.3 code is pushed but before setupForm() runs
  submitTo(S.E, tab, { [Q.name]: '舊', [Q.email]: 'v32@gmail.com', [Q.teacherSay]: TS, [Q.software]: 'Kahoot 教師版, HiTeach 專業版',
    [Q.role]: '學年', [Q.grade]: '二年級', [Q.aiPro]: '是', [Q.workspace]: 'v32' + SCHOOL }, t(0));
  let rows = regRows(S.E);
  assert.deepStrictEqual(rows.map((r) => [r[4], r[5]]), [[TS, '正取'], ['Kahoot 教師版', '正取'], ['HiTeach 專業版', '正取'], [AI, '正取']]);
  assert.strictEqual(rows[2][10], '此帳號另登記了：Kahoot 教師版（每人除 TeacherSay 外限一項，請確認）');
  assert.strictEqual(rows[3][10], '此帳號另登記了：Kahoot 教師版、HiTeach 專業版（每人除 TeacherSay 外限一項，請確認）');

  S.E.call('setupForm()');
  assert.strictEqual(S.E.createdForms.length, 0, 'no new form');
  assert.strictEqual(S.E.drivePermissions.length, 0, 'responder access untouched');
  assert.strictEqual(S.E.props.ENROLL_RESPONSE_SHEET_ID, String(tab.id), 'same response tab');
  assert.deepStrictEqual(titlesOf(f), FORM_LAYOUT);
  assertJumps(f);
  assert.ok([pT, ts, pW, pR, pG, pS, pO, pWs].every((it) => f.items.includes(it)), 'existing items are kept, not recreated');
  assert.ok(!f.items.includes(pA) && !f.items.includes(aiQ) && item(f, Q.software) === undefined, 'old questions removed');
  const msg = S.E.alerts.find((a) => a.includes('已更新申請表單'));
  ['刪除舊題「' + Q.software + '」', '刪除舊題「' + Q.aiPro + '」', '刪除舊題「【Google AI Pro 教育版】」', '新增「' + Q.pick + '」']
    .forEach((text) => assert.ok(msg.includes(text), text));
  assert.ok(tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0].includes(Q.aiPro), 'old answers stay in the response tab');
  assert.deepStrictEqual(walkForm(f, answersOf(teacher(1, '二年級', { workspace: 't1' + SCHOOL }), AI)).visited,
    ['基本資料', '【TeacherSay 專業版】', '【職務】', '【學年】', '申請軟體授權', '【Google AI Pro 學校帳號】']);

  // after the upgrade
  submitTo(S.E, tab, answersOf({ name: '新', email: 'n@gmail.com', role: '科任', detail: '自然' }, 'Padlet 學校版'), t(1));
  rows = regRows(S.E);
  assert.deepStrictEqual(rows.slice(4).map((r) => [r[2], r[4], r[5], r[10]]),
    [['n@gmail.com', TS, '正取', ''], ['n@gmail.com', 'Padlet 學校版', '正取', '']]);
});

// ============ 4. no duplicates from the old form ============
ok('responses from the old form are ignored (before and after the new form exists)', () => {
  const S = schoolEnv();
  S.E.call('setupSheets()');
  const oldAnswers = { '您的大名': '甲', '您的常用email信箱': 'a@x.tw', '您的職務': '教師', '您要申請的軟體授權': 'Hiteach' };
  submitTo(S.E, S.oldSheet, oldAnswers, t(1));
  assert.strictEqual(regRows(S.E).length, 0);
  assert.strictEqual(S.E.mails.length, 0, 'no error mail either');

  const R = readyEnv();
  submitTo(R.E, R.oldSheet, oldAnswers, t(1));
  submitTo(R.E, R.oldSheet, { '您的大名': '乙', '您要申請的軟體授權': 'HiTeach 專業版' }, t(2));
  assert.strictEqual(regRows(R.E).length, 0);
  assert.strictEqual(R.E.mails.length, 0);
  assert.strictEqual(R.lic.get(3, 2), '', 'no serial handed out');
  R.send(3, teacher(1, '一年級'), 'HiTeach 專業版');
  assert.deepStrictEqual(regRows(R.E).map((r) => r[4]), [TS, 'HiTeach 專業版']);
});

// ============ 5. end-to-end registration ============
ok('registration: TeacherSay + one program, ranks, waitlist, AI Pro rules, CC, remaining seats', () => {
  const { E, form, send } = readyEnv();
  for (let i = 1; i <= 5; i++) send(i, teacher(i, '一年級'), 'Kahoot 教師版');
  send(6, teacher(6, '二年級', { workspace: 't6' + SCHOOL }), AI);
  send(7, teacher(7, '二年級', { workspace: 't7' + SCHOOL }), AI);
  send(8, { name: '科任8', email: 't8@gmail.com', role: '科任', detail: '自然', workspace: 't8' + SCHOOL }, AI);
  send(9, { name: '科任9', email: 't9@gmail.com', role: '科任', detail: '社會', workspace: 't9' + SCHOOL }, AI);
  send(10, { name: '幼10', email: 't10@gmail.com', role: '行政', detail: '幼兒園' }, 'Wordwall 專業版');
  send(11, teacher(11, '一年級'), 'Padlet 學校版');
  send(12, teacher(12, '三年級'), ''); // TeacherSay only
  send(13, teacher(13, '四年級'), 'Kahoot 教師版'); // Kahoot is full by now
  send(14, teacher(14, '五年級'), 'HiTeach 專業版');

  const rows = regRows(E);
  assert.strictEqual(rows.filter((r) => r[4] === TS).length, 14, 'one TeacherSay row per person');
  assert.ok(rows.filter((r) => r[4] === TS).every((r) => r[5] === '正取'), 'TeacherSay is unlimited');
  assert.strictEqual(rows.length, 14 + 6 + 4 + 1 + 1 + 1); // TeacherSay, Kahoot, AI Pro, Wordwall, Padlet, HiTeach
  assert.deepStrictEqual(rowOf(rows, 't5@gmail.com', 'Kahoot 教師版').slice(5, 7), ['正取', 5]);
  assert.deepStrictEqual(rowOf(rows, 't13@gmail.com', 'Kahoot 教師版').slice(5, 7), ['備取', 1]);
  assert.strictEqual(rowOf(rows, 't1@gmail.com', 'Kahoot 教師版')[3], '學年／一年級');
  assert.strictEqual(rowOf(rows, 't8@gmail.com', AI)[3], '科任／自然');
  assert.strictEqual(rowOf(rows, 't10@gmail.com', 'Wordwall 專業版')[3], '行政／幼兒園');
  assert.deepStrictEqual(['t6', 't7', 't8', 't9'].map((p) => rowOf(rows, p + '@gmail.com', AI)[5]), ['正取', '備取', '正取', '備取'],
    'one per grade, one for all 科任');
  assert.strictEqual(rowOf(rows, 't10@gmail.com', AI), undefined);
  assert.strictEqual(rowOf(rows, 't10@gmail.com', 'Wordwall 專業版')[11], '', 'no school account recorded for 行政');
  assert.deepStrictEqual([rowOf(rows, 't6@gmail.com', AI)[11], rowOf(rows, 't6@gmail.com', AI)[10]], ['t6' + SCHOOL, '']);
  assert.deepStrictEqual(rows.filter((r) => r[2] === 't12@gmail.com').map((r) => r[4]), [TS]);
  assert.ok(rows.every((r) => r[7] === false && r[8] === ''), 'no approval, no serial at registration');

  const receipts = E.mails.filter((m) => m.subject === '【已登記】教學軟體授權申請');
  assert.strictEqual(receipts.length, 14);
  assert.ok(receipts.every((m) => m.htmlBody.includes('<b>TeacherSay 專業版</b>')));
  assert.ok(receipts.every((m) => m.htmlBody.includes('請用同一個 Google 帳號重新填寫，以最後一次為準')), 'every receipt explains how to change');
  assert.deepStrictEqual([receipts[0].to, receipts[0].cc], ['t1@gmail.com', undefined]);
  assert.deepStrictEqual([receipts[5].to, receipts[5].cc], ['t6@gmail.com', 't6' + SCHOOL]);
  assert.ok(receipts[6].htmlBody.includes('備取（備取第 1 位）') && receipts[6].cc === 't7' + SCHOOL);
  assert.ok(!receipts[9].htmlBody.includes(AI) && receipts[9].cc === undefined, '行政 receipt has no AI Pro and no CC');
  assert.ok(!receipts[10].htmlBody.includes('您先前已登記過') && receipts[10].htmlBody.includes('<b>Padlet 學校版</b>：正取（第 1 位）'));
  assert.ok(receipts.every((m) => !m.htmlBody.includes('已取消')), 'first submissions cancel nothing');
  assert.ok(receipts[12].htmlBody.includes('<b>Kahoot 教師版</b>：備取（備取第 1 位）'));
  const adminSubjects = E.mails.filter((m) => m.to === 'lres712@mail.lres.tc.edu.tw').map((m) => m.subject);
  assert.strictEqual(adminSubjects.length, 2);
  assert.ok(adminSubjects.some((x) => x.includes('Kahoot 教師版')) && adminSubjects.some((x) => x.includes(AI)));

  const help = item(form, Q.pick).help;
  assert.ok(help.includes('・Kahoot 教師版：已額滿（備取 1 人）'), help);
  assert.ok(help.includes('・HiTeach 專業版：剩餘 14／15'), help);
  assert.ok(help.includes('・Google AI Pro 教育版：剩餘 5／7｜限各學年、科任各一位，尚可申請：一年級、三年級、四年級、五年級、六年級'), help);
  const officeHelp = item(form, Q.pickOffice).help;
  assert.ok(officeHelp.includes('・Kahoot 教師版：已額滿（備取 1 人）') && !officeHelp.includes('Google AI Pro'), officeHelp);
  const stats = E.ss.getSheetByName('申請統計').getRange(1, 1, 8, 10).getValues();
  assert.deepStrictEqual(stats.find((r) => r[0] === 'Kahoot 教師版').slice(1, 8), [5, 6, 5, 1, 0, 0, 0]);
  assert.deepStrictEqual(stats.find((r) => r[0] === TS).slice(1, 8), ['不限', 14, 14, 0, 0, 0, '不限']);
});

ok('resubmitting: the latest form counts — dropped programs cancelled, approved ones flagged, issued ones kept, backfill replays the same', () => {
  const { E, form, sheet, send } = readyEnv();
  const regSheet = E.ss.getSheetByName('申請登記');
  for (let i = 1; i <= 6; i++) send(i, teacher(i, '一年級'), 'Kahoot 教師版'); // t6 is 備取 1
  assert.deepStrictEqual(rowOf(regRows(E), 't6@gmail.com', 'Kahoot 教師版').slice(5, 7), ['備取', 1]);
  const rowNo = (email, sw) => regRows(E).findIndex((r) => r[2] === email && r[4] === sw) + 2;
  regSheet.getRange(rowNo('t3@gmail.com', 'Kahoot 教師版'), 8).setValue(true);     // approved, not issued yet
  regSheet.getRange(rowNo('t4@gmail.com', 'Kahoot 教師版'), 6).setValue('已發放'); // already issued
  const before = E.mails.length;

  send(7, teacher(1, '一年級', { email: 'T1@gmail.com' }), 'Padlet 學校版'); // switches program (address case differs)
  send(8, teacher(2, '一年級'), '');                                         // withdraws: TeacherSay only
  send(9, teacher(3, '一年級'), 'Wordwall 專業版');                          // approved Kahoot is not touched
  send(10, teacher(4, '一年級'), '');                                        // issued Kahoot is not touched
  send(11, teacher(5, '一年級'), 'Kahoot 教師版');                           // same choice again: nothing changes

  let rows = regRows(E);
  const of = (email, sw) => rowOf(rows, email, sw);
  assert.deepStrictEqual(of('t1@gmail.com', 'Kahoot 教師版').slice(5, 7), ['取消', '']);
  assert.ok(of('t1@gmail.com', 'Kahoot 教師版')[10].includes('重新填寫，未再選此項，已自動取消'));
  assert.deepStrictEqual([of('t1@gmail.com', 'Padlet 學校版')[5], of('t1@gmail.com', 'Padlet 學校版')[10]], ['正取', ''],
    'no second-program note once the old program is cancelled');
  assert.deepStrictEqual(of('t2@gmail.com', 'Kahoot 教師版')[5], '取消');
  assert.deepStrictEqual(of('t3@gmail.com', 'Kahoot 教師版').slice(5, 8), ['正取', 1, true]);
  assert.ok(of('t3@gmail.com', 'Kahoot 教師版')[10].includes('重新填寫時未選此項（已核准未發放），請確認是否保留'));
  assert.strictEqual(of('t3@gmail.com', 'Wordwall 專業版')[10], '此帳號另登記了：Kahoot 教師版（每人除 TeacherSay 外限一項，請確認）');
  assert.deepStrictEqual([of('t4@gmail.com', 'Kahoot 教師版')[5], of('t4@gmail.com', 'Kahoot 教師版')[10]], ['已發放', '']);
  assert.deepStrictEqual([of('t5@gmail.com', 'Kahoot 教師版').slice(5, 7), of('t5@gmail.com', 'Kahoot 教師版')[10]], [['正取', 2], '']);
  assert.deepStrictEqual(of('t6@gmail.com', 'Kahoot 教師版').slice(5, 7), ['正取', 3], 'the waitlist moves up');
  assert.ok(rows.filter((r) => r[4] === TS).every((r) => r[5] === '正取'), 'TeacherSay is never cancelled');
  assert.strictEqual(rows.filter((r) => r[4] === TS).length, 6, 'no second TeacherSay row');

  const receipts = E.mails.slice(before).filter((m) => m.subject === '【已登記】教學軟體授權申請');
  assert.strictEqual(receipts.length, 5);
  assert.ok(receipts[0].htmlBody.includes('<b>已取消</b>：Kahoot 教師版') && receipts[0].htmlBody.includes('<b>Padlet 學校版</b>：正取'));
  assert.ok(receipts[1].htmlBody.includes('<b>已取消</b>：Kahoot 教師版'));
  assert.ok(receipts[2].htmlBody.includes('資訊組會再與您確認是否保留：Kahoot 教師版') && !receipts[2].htmlBody.includes('已取消'));
  assert.ok(!receipts[3].htmlBody.includes('已取消') && !receipts[3].htmlBody.includes('確認是否保留'));
  assert.ok(receipts[4].htmlBody.includes('您先前已登記過') && !receipts[4].htmlBody.includes('已取消'));
  assert.ok(item(form, Q.pick).help.includes('・Kahoot 教師版：剩餘 1／5'), item(form, Q.pick).help);

  // a v3.2 response with four programs ticked, then the same person refills with one
  submitTo(E, sheet, {
    [Q.name]: '盧', [Q.email]: 'lu@gmail.com', [Q.role]: '科任', [Q.subject]: '英語', [Q.teacherSay]: TS,
    [Q.software]: 'Wayground (Quizizz), Padlet 學校版, Wordwall 專業版, HiTeach 專業版', [Q.aiPro]: '是', [Q.workspace]: 'lu' + SCHOOL,
  }, t(20));
  const lu = () => Object.fromEntries(regRows(E).filter((r) => r[2] === 'lu@gmail.com').map((r) => [r[4], r[5]]));
  assert.strictEqual(Object.keys(lu()).length, 6);
  assert.ok(rowOf(regRows(E), 'lu@gmail.com', 'Wordwall 專業版')[10].startsWith('此帳號另登記了'));
  send(21, { name: '盧', email: 'lu@gmail.com', role: '科任', detail: '英語' }, 'Wordwall 專業版');
  assert.deepStrictEqual(lu(), {
    [TS]: '正取', 'Wayground (Quizizz)': '取消', 'Padlet 學校版': '取消', 'Wordwall 專業版': '正取', 'HiTeach 專業版': '取消', [AI]: '取消',
  });
  assert.strictEqual(rowOf(regRows(E), 'lu@gmail.com', 'Wordwall 專業版')[10], '', 'the stale second-program note is cleared');
  const luReceipt = E.mails.filter((m) => m.to === 'lu@gmail.com').pop().htmlBody;
  assert.ok(luReceipt.includes('<b>已取消</b>：') && ['Wayground (Quizizz)', 'Padlet 學校版', 'HiTeach 專業版', AI].every((x) => luReceipt.includes(x)));

  // backfill replays every response in order and lands on the same sheet, sending nothing
  const snapshot = JSON.stringify(regRows(E));
  const mails = E.mails.length;
  E.call('menuBackfill()');
  assert.strictEqual(JSON.stringify(regRows(E)), snapshot);
  assert.strictEqual(E.mails.length, mails);
  assert.ok(E.alerts[E.alerts.length - 1].includes('新增 0 筆登記') && !E.alerts[E.alerts.length - 1].includes('取消'));
});

ok('a refill that came in before this version is applied by backfill', () => {
  const { E, sheet } = readyEnv();
  const a1 = answersOf(teacher(1, '一年級'), 'Kahoot 教師版');
  const a2 = answersOf(teacher(1, '一年級'), 'HiTeach 專業版');
  // both responses are on the tab but neither was registered (e.g. the trigger failed)
  [[a1, t(1)], [a2, t(5)]].forEach(([a, when]) => {
    Object.keys(a).forEach((key) => {
      const cur = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (cur.indexOf(key) < 0) sheet.set(1, cur.length + 1, key);
    });
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map((h) => (h === '時間戳記' ? when : (a[h] === undefined ? '' : a[h]))));
  });
  E.call('menuBackfill()');
  const rows = regRows(E);
  assert.deepStrictEqual(rows.map((r) => [r[4], r[5]]), [[TS, '正取'], ['Kahoot 教師版', '取消'], ['HiTeach 專業版', '正取']]);
  assert.ok(E.alerts[E.alerts.length - 1].includes('新增 3 筆登記；依重新填寫（以最後一次為準）取消 1 筆舊登記'), E.alerts[E.alerts.length - 1]);
});

ok('resending receipts: same format from the current registrations, CC the school account, confirm first', () => {
  const { E, send } = readyEnv();
  send(1, teacher(1, '一年級', { workspace: 't1' + SCHOOL }), AI);
  send(2, { name: '科任2', email: 't2@gmail.com', role: '科任', detail: '英語' }, 'Kahoot 教師版');
  send(3, { name: '科任2', email: 't2@gmail.com', role: '科任', detail: '英語' }, ''); // withdrew Kahoot
  let n = E.mails.length;

  E.call('menuResendReceipts()', { __prompt: { button: 'CANCEL', text: 't1@gmail.com' } });
  assert.strictEqual(E.mails.length, n, 'cancelled prompt sends nothing');
  E.call('menuResendReceipts()', { __prompt: { button: 'OK', text: 't1@gmail.com' }, __confirm: 'NO' });
  assert.strictEqual(E.mails.length, n, 'declined confirmation sends nothing');
  E.call('menuResendReceipts()', { __prompt: { button: 'OK', text: 'nobody@x.tw' }, __confirm: 'YES' });
  assert.strictEqual(E.mails.length, n);
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可補寄的回條') && E.alerts[E.alerts.length - 1].includes('nobody@x.tw'));

  E.call('menuResendReceipts()', { __prompt: { button: 'OK', text: 'T1@gmail.com，t2@gmail.com\nnobody@x.tw' }, __confirm: 'YES' });
  const confirmText = E.alerts[E.alerts.length - 2];
  assert.ok(confirmText.includes('將寄出 2 封') && confirmText.includes('副本 t1' + SCHOOL) && confirmText.includes('不寄：nobody@x.tw'), confirmText);
  const sent = E.mails.slice(n);
  assert.strictEqual(sent.length, 2);
  assert.deepStrictEqual(sent.map((m) => [m.to, m.cc, m.subject]),
    [['t1@gmail.com', 't1' + SCHOOL, '【已登記】教學軟體授權申請'], ['t2@gmail.com', undefined, '【已登記】教學軟體授權申請']]);
  assert.ok(sent[0].htmlBody.includes('此為補寄的登記回條') && sent[0].htmlBody.includes('<b>' + AI + '</b>：正取（第 1 位）'));
  assert.ok(sent[0].htmlBody.includes('請用同一個 Google 帳號重新填寫，以最後一次為準'));
  assert.ok(!sent[1].htmlBody.includes('Kahoot') && sent[1].htmlBody.includes('<b>TeacherSay 專業版</b>：正取'), 'cancelled rows are left out');
  assert.ok(!sent.some((m) => m.htmlBody.includes('您先前已登記過')));
  assert.ok(E.alerts[E.alerts.length - 1].includes('已補寄 2 封回條：老師1、科任2'), E.alerts[E.alerts.length - 1]);

  const low = readyEnv({ mailQuota: 1 });
  low.send(1, teacher(1, '一年級', { workspace: 't1' + SCHOOL }), AI);
  n = low.E.mails.length;
  low.E.call('menuResendReceipts()', { __prompt: { button: 'OK', text: 't1@gmail.com' }, __confirm: 'YES' });
  assert.strictEqual(low.E.mails.length, n);
  assert.ok(low.E.alerts[low.E.alerts.length - 1].includes('剩餘寄信額度為 1'));
});

const tabletAnswers = (who, cart, mode, count) => ({
  [Q.name]: who.name, [Q.email]: who.email, [Q.kind]: '平板借用', [Q.tabletUnit]: who.unit || '三年二班導師',
  [Q.tabletCart]: cart, [Q.tabletMode]: mode, [Q.tabletCount]: count === undefined ? '' : String(count),
});
const WHOLE = '整車借用';
const PARTIAL = '部分借用（填寫數量）';
const regRows14 = (E) => { const s = E.ss.getSheetByName('申請登記'); const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, 14).getValues(); };

ok('tablet borrowing: own branch, cart + count, 已登記 with no seats, independent of software, latest wins, backfill replays', () => {
  const { E, form, sheet, send } = readyEnv();
  // ---- respondent paths ----
  const lu = { name: '盧老師', email: 'lu@gmail.com', unit: '英語科任' };
  let w = walkForm(form, tabletAnswers(lu, '平板車 A（28 台）', WHOLE));
  assert.deepStrictEqual([w.visited, w.submitted], [['基本資料', '【平板借用】'], true], 'whole cart: no TeacherSay, no count page');
  w = walkForm(form, tabletAnswers(lu, '平板車 G（30 台）', PARTIAL, 12));
  assert.deepStrictEqual([w.visited, w.submitted], [['基本資料', '【平板借用】', '【平板借用數量】'], true]);
  [0, 31, ''].forEach((n) => assert.strictEqual(walkForm(form, tabletAnswers(lu, '平板車 G（30 台）', PARTIAL, n)).blockedAt, Q.tabletCount, 'count ' + n));
  assert.deepStrictEqual(item(form, Q.tabletCart).choices.slice(0, 2).concat(item(form, Q.tabletCart).choices.slice(-2)),
    ['平板車 A（28 台）', '平板車 B（28 台）', '黑車（30 台）', '白車（28 台）']);
  w = walkForm(form, answersOf(teacher(9, '一年級'), 'Kahoot 教師版'));
  assert.ok(w.submitted && !w.visited.includes('【平板借用】'), 'software path never enters the tablet sections');
  assert.ok(form.description.includes('平板借用') && item(form, Q.email).help.includes('填報成果（含平板借用）'));

  // ---- registering ----
  const before = E.mails.length;
  submitTo(E, sheet, tabletAnswers(lu, '平板車 A（28 台）', WHOLE), t(1));
  let rows = regRows14(E);
  assert.deepStrictEqual(rows.map((r) => [r[1], r[2], r[3], r[4], r[5], r[6], r[12], r[13]]),
    [['盧老師', 'lu@gmail.com', '英語科任', '平板借用', '已登記', '', '平板車 A', 28]]);
  assert.strictEqual(E.ss.getSheetByName('申請登記').getRange(1, 13, 1, 2).getValues()[0].join(), '借用平板車,借用數量');
  let mail = E.mails[E.mails.length - 1];
  assert.deepStrictEqual([mail.to, mail.subject], ['lu@gmail.com', '【已登記】平板借用']);
  assert.ok(mail.htmlBody.includes('平板車 A，28 台（整車借用）') && mail.htmlBody.includes('填報時請填這個 Google 帳號（lu@gmail.com）'));
  assert.ok(!E.ss.getSheetByName('申請統計').getRange(1, 1, 20, 1).getValues().some((r) => r[0] === '平板借用'), 'no seats, not in the stats');

  // the same person also applies for software: both registrations stay
  send(2, { name: '盧老師', email: 'lu@gmail.com', role: '科任', detail: '英語' }, 'Padlet 學校版');
  rows = regRows14(E);
  assert.deepStrictEqual(rows.map((r) => [r[4], r[5]]), [['平板借用', '已登記'], [TS, '正取'], ['Padlet 學校版', '正取']]);
  assert.ok(!E.mails[E.mails.length - 1].htmlBody.includes('平板'), 'the software receipt is about software only');

  // same tablet answers again: nothing changes
  submitTo(E, sheet, tabletAnswers(lu, '平板車 A（28 台）', WHOLE), t(3));
  assert.strictEqual(regRows14(E).length, 3);
  assert.ok(E.mails[E.mails.length - 1].htmlBody.includes('您先前已登記過相同內容'));

  // a different cart and count: the old one is cancelled, software untouched
  submitTo(E, sheet, tabletAnswers(lu, '平板車 B（28 台）', PARTIAL, 10), t(4));
  rows = regRows14(E);
  assert.deepStrictEqual(rows.map((r) => [r[4], r[5], r[12], r[13]]),
    [['平板借用', '取消', '平板車 A', 28], [TS, '正取', '', ''], ['Padlet 學校版', '正取', '', ''], ['平板借用', '已登記', '平板車 B', 10]]);
  assert.ok(rows[0][10].includes('重新填寫平板借用'));
  mail = E.mails[E.mails.length - 1];
  assert.ok(mail.htmlBody.includes('平板車 B，10 台') && mail.htmlBody.includes('原本的登記（平板車 A，28 台）<b>已取消</b>'));

  // a later software refill does not cancel the tablet registration
  send(5, { name: '盧老師', email: 'lu@gmail.com', role: '科任', detail: '英語' }, 'Kahoot 教師版');
  rows = regRows14(E);
  assert.strictEqual(rows.find((r) => r[4] === '平板借用' && r[5] === '已登記')[12], '平板車 B');
  assert.strictEqual(rows.find((r) => r[4] === 'Padlet 學校版')[5], '取消');

  // more than the cart holds: counted as the whole cart, with a note
  submitTo(E, sheet, tabletAnswers({ name: '張老師', email: 'chang@gmail.com', unit: '教務處' }, '平板車 C（28 台）', PARTIAL, 40), t(6));
  const chang = regRows14(E).find((r) => r[2] === 'chang@gmail.com');
  assert.deepStrictEqual([chang[12], chang[13]], ['平板車 C', 28]);
  assert.ok(chang[10].includes('填寫 40 台超過該車台數，以整車 28 台計'));
  assert.ok(E.mails[E.mails.length - 1].htmlBody.includes('提醒：填寫 40 台超過該車台數'));

  // approving all 正取 never ticks a tablet row; the status survives re-ranking
  E.call('menuApproveAllAccepted()');
  rows = regRows14(E);
  assert.ok(rows.filter((r) => r[4] === '平板借用').every((r) => r[7] !== true), 'tablet rows are not approved');
  assert.deepStrictEqual(rows.filter((r) => r[4] === '平板借用').map((r) => r[5]), ['取消', '已登記', '已登記']);

  // resend to a tablet-only person uses the tablet receipt
  const n = E.mails.length;
  E.call('menuResendReceipts()', { __prompt: { button: 'OK', text: 'chang@gmail.com' }, __confirm: 'YES' });
  assert.deepStrictEqual([E.mails.length - n, E.mails[n].subject], [1, '【已登記】平板借用']);
  assert.ok(E.mails[n].htmlBody.includes('此為補寄的登記回條') && E.mails[n].htmlBody.includes('平板車 C，28 台'));

  // backfill replays every response and lands on the same sheet
  const snapshot = JSON.stringify(regRows14(E));
  const mails = E.mails.length;
  E.call('menuBackfill()');
  assert.strictEqual(JSON.stringify(regRows14(E)), snapshot);
  assert.strictEqual(E.mails.length, mails);
  assert.ok(before < mails);
});

// ============ 6. approval + issuing ============
ok('issuing: approved rows only, HiTeach serials, account wording, one mail per person with CC, revoking promotes the waitlist', () => {
  const { E, lic, ww, send } = readyEnv();
  E.ss.getSheetByName('軟體清單').getRange(7, 2).setValue(2); // HiTeach -> 2 seats (rows: header, TS, Wayground, Kahoot, Padlet, Wordwall, HiTeach)
  send(1, { name: '甲', email: 'a@gmail.com', role: '學年', detail: '一年級', workspace: 'a' + SCHOOL }, AI);
  send(2, { name: '乙', email: 'b@gmail.com', role: '科任', detail: '自然' }, 'HiTeach 專業版');
  send(3, { name: '丙', email: 'c@gmail.com', role: '科任', detail: '社會' }, 'HiTeach 專業版');
  send(4, { name: '丁', email: 'd@gmail.com', role: '行政', detail: '總務處' }, 'HiTeach 專業版');
  send(5, { name: '戊', email: 'e@gmail.com', role: '學年', detail: '二年級' }, 'Wordwall 專業版');
  const before = E.mails.length;

  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, before);
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可發放的名單'));

  E.call('menuApproveAllAccepted()');
  let rows = regRows(E);
  assert.ok(rows.filter((r) => r[5] === '正取').every((r) => r[7] === true));
  assert.deepStrictEqual(rowOf(rows, 'd@gmail.com', 'HiTeach 專業版').slice(5, 8), ['備取', 1, false], 'the waitlist is not ticked');
  E.call('menuPreviewIssue()');
  assert.strictEqual(E.mails.length, before, 'preview must not send');

  E.call('menuIssueApproved()');
  rows = regRows(E);
  assert.deepStrictEqual([rowOf(rows, 'b@gmail.com', 'HiTeach 專業版')[5], rowOf(rows, 'b@gmail.com', 'HiTeach 專業版')[8]], ['已發放', 'HT-A']);
  assert.deepStrictEqual([rowOf(rows, 'c@gmail.com', 'HiTeach 專業版')[5], rowOf(rows, 'c@gmail.com', 'HiTeach 專業版')[8]], ['已發放', 'HT-B']);
  assert.deepStrictEqual([rowOf(rows, 'd@gmail.com', 'HiTeach 專業版')[5], rowOf(rows, 'd@gmail.com', 'HiTeach 專業版')[8]], ['備取', '']);
  assert.ok(['a', 'b', 'c', 'd', 'e'].every((p) => rowOf(rows, p + '@gmail.com', TS)[5] === '已發放'));
  assert.strictEqual(rowOf(rows, 'a@gmail.com', AI)[5], '已發放');
  assert.strictEqual(rowOf(rows, 'e@gmail.com', 'Wordwall 專業版')[5], '已發放');
  assert.deepStrictEqual([lic.get(2, 1), lic.get(2, 2)], ['HT-OLD', 'old@x.tw']);
  assert.deepStrictEqual([lic.get(3, 2), lic.get(4, 2)], ['b@gmail.com', 'c@gmail.com']);
  assert.deepStrictEqual(ww.getRange(2, 1, 1, 2).getValues()[0], ['戊', 'e@gmail.com']);

  const issued = E.mails.slice(before).filter((m) => m.subject === '【授權核准】教學軟體授權已發放');
  assert.deepStrictEqual(issued.map((m) => m.to).sort(), ['a@gmail.com', 'b@gmail.com', 'c@gmail.com', 'd@gmail.com', 'e@gmail.com']);
  const mailA = issued.find((m) => m.to === 'a@gmail.com');
  assert.strictEqual(mailA.cc, 'a' + SCHOOL);
  assert.ok(mailA.htmlBody.includes('<b>Google AI Pro 教育版</b>－授權將開通於您的學校帳號（a' + SCHOOL + '）'));
  assert.ok(mailA.htmlBody.includes('<b>TeacherSay 專業版</b>－授權將開通於您登記的 Google 帳號（a@gmail.com）'));
  const mailB = issued.find((m) => m.to === 'b@gmail.com');
  assert.ok(mailB.htmlBody.includes('HT-A') && !mailB.htmlBody.includes('HT-B') && mailB.cc === undefined);
  const mailD = issued.find((m) => m.to === 'd@gmail.com');
  assert.ok(mailD.htmlBody.includes('TeacherSay') && !mailD.htmlBody.includes('HiTeach'), 'the waitlisted program is not in the notice');

  const n = E.mails.length;
  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, n, 're-running sends nothing new');

  // revoke 丙's HiTeach: the waitlisted 丁 moves up
  const reg = E.ss.getSheetByName('申請登記');
  const cRow = rows.findIndex((r) => r[2] === 'c@gmail.com' && r[4] === 'HiTeach 專業版') + 2;
  reg.getRange(cRow, 6).setValue('取消');
  E.call('menuRerank()');
  rows = regRows(E);
  assert.strictEqual(rows[cRow - 2][5], '取消');
  assert.deepStrictEqual(rowOf(rows, 'd@gmail.com', 'HiTeach 專業版').slice(5, 7), ['正取', 1]);
});

ok('mail quota counts CC addresses and blocks before any write', () => {
  const { E, send } = readyEnv({ mailQuota: 1 });
  send(1, { name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然', workspace: 'a' + SCHOOL }, AI);
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  assert.ok(E.alerts[E.alerts.length - 1].includes('含副本共 2 個收件人'));
  assert.ok(regRows(E).every((r) => r[5] === '正取' && r[9] === ''));
});

ok('a failed notification mail is recorded, the licence stays issued', () => {
  const { E, send } = readyEnv({ failMailTo: 'a@gmail.com' });
  send(1, { name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然' }, 'HiTeach 專業版');
  assert.strictEqual(regRows(E).length, 2);
  const warn = E.mails[E.mails.length - 1];
  assert.ok(warn.subject.includes('回條未寄出') && !warn.subject.includes('【錯誤】'), warn.subject);
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  const row = rowOf(regRows(E), 'a@gmail.com', 'HiTeach 專業版');
  assert.deepStrictEqual([row[5], row[8], row[10]], ['已發放', 'HT-A', '通知信寄送失敗，請人工通知']);
});

// ============ 7. failure paths ============
ok('renamed question: admin is mailed, nothing registered, backfill recovers from the new tab only', () => {
  const { E, sheet, oldSheet } = readyEnv();
  const a = answersOf({ name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然' }, 'Kahoot 教師版');
  Object.keys(a).forEach((key) => {
    const cur = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (cur.indexOf(key) < 0) sheet.set(1, cur.length + 1, key);
  });
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
  assert.deepStrictEqual(rows.map((r) => [r[1], r[3], r[4], r[5]]), [['甲', '科任／自然', TS, '正取'], ['甲', '科任／自然', 'Kahoot 教師版', '正取']]);
  assert.strictEqual(E.mails.length, n, 'backfill sends no mail');
  E.call('menuBackfill()');
  assert.strictEqual(regRows(E).length, 2, 'backfill is idempotent');
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
  assert.strictEqual(regRows(S.E).length, 2);
  assert.strictEqual(S.E.props.ENROLL_RESPONSE_SHEET_ID, String(sheet.id));
});

ok('manual run / no selection / bad account / no UI / reset are all safe', () => {
  const { E, send } = readyEnv();
  E.call('onFormSubmit()');
  send(1, teacher(1, '一年級'), '', { noTeacherSay: true });
  assert.strictEqual(regRows(E).length, 0, 'nothing ticked -> nothing registered');
  send(2, teacher(2, '一年級', { email: 'not-an-email' }), 'Padlet 學校版');
  const rows = regRows(E);
  assert.strictEqual(rows.length, 2);
  assert.ok(rows.every((r) => r[10].includes('Google 帳號格式有誤')));
  assert.ok(E.mails.every((m) => m.to !== 'not-an-email'));

  const N = readyEnv({ noUi: true });
  N.send(1, teacher(1, '一年級'), 'HiTeach 專業版');
  N.E.ss.getSheetByName('申請登記').getRange(3, 8).setValue(true);
  const n = N.E.mails.length;
  N.E.call('menuIssueApproved()');
  assert.strictEqual(N.E.mails.length, n, 'no UI -> no confirmation -> nothing issued');
  assert.ok(regRows(N.E).every((r) => r[5] === '正取'));

  E.call('resetEnrollmentForm()');
  send(3, teacher(3, '一年級'), 'Padlet 學校版');
  assert.strictEqual(regRows(E).length, 2, 'after reset nothing registers until a form is bound again');
});

ok('upgrade from v3.0: old registrations are archived to a renamed tab, the new tab starts empty', () => {
  const S = schoolEnv();
  const old = S.E.ss.insertSheet('申請登記');
  old.getRange(1, 1, 1, 11).setValues([['登記時間', '姓名', 'Email', '身分', '軟體', '狀態', '順位', '核准', '序號', '發放時間', '備註']]);
  old.appendRow([t(0), '測試', 'test@x.tw', '一年級', 'Kahoot 教師版', '正取', 1, false, '', '', '']);
  old.appendRow([t(0), '測試', 'test@x.tw', '一年級', AI, '正取', 1, false, '', '', '']);
  S.E.call('setupSheets()');
  S.E.call('setupForm()');
  assert.ok(S.E.alerts.some((a) => a.includes('已有 2 筆來自舊表單的登記')), 'asks before archiving');
  const archived = S.E.ss.sheets.find((s) => s.name.startsWith('申請登記_舊_'));
  assert.ok(archived === old && archived.getLastRow() === 3, 'old rows kept in the renamed tab');
  const fresh = S.E.ss.getSheetByName('申請登記');
  assert.notStrictEqual(fresh, old);
  assert.strictEqual(fresh.getLastRow(), 1);
  assert.strictEqual(fresh.get(1, 12), '學校帳號');
  const form = S.E.createdForms[0];
  const sheet = S.E.ss.sheets.find((s) => s.formUrl === form.url);
  assert.ok(item(form, Q.pick).help.includes('・Kahoot 教師版：剩餘 5／5'), 'archived rows no longer use seats');
  submitTo(S.E, sheet, answersOf(teacher(1, '一年級', { workspace: 't1' + SCHOOL }), AI), t(1));
  assert.deepStrictEqual(regRows(S.E).map((r) => [r[4], r[5], r[6]]), [[TS, '正取', 1], [AI, '正取', 1]]);

  const N = schoolEnv({ noUi: true });
  const keep = N.E.ss.insertSheet('申請登記');
  keep.getRange(1, 1, 1, 2).setValues([['登記時間', '姓名']]);
  keep.appendRow([t(0), '測試']);
  N.E.call('setupForm()');
  assert.strictEqual(N.E.createdForms.length, 0);
  assert.strictEqual(N.E.ss.getSheetByName('申請登記'), keep);
});

ok('admin hands a 行政 seat out by hand: a row marked 已發放 uses a seat, is never mailed or re-issued', () => {
  const { E, send } = readyEnv();
  const reg = E.ss.getSheetByName('申請登記');
  // the admin types a row for an office staff member directly into 申請登記
  reg.appendRow([t(0), '主任', 'boss@gmail.com', '行政／教務處', AI, '已發放', '', false, '', t(0), '管理者手動分配', 'boss' + SCHOOL]);
  for (let i = 1; i <= 6; i++) send(i, teacher(i, ['一年級', '二年級', '三年級', '四年級', '五年級', '六年級'][i - 1], { workspace: 't' + i + SCHOOL }), AI);
  send(7, { name: '科任7', email: 't7@gmail.com', role: '科任', detail: '自然', workspace: 't7' + SCHOOL }, AI);
  const rows = regRows(E);
  const ai = rows.filter((r) => r[4] === AI);
  assert.deepStrictEqual(ai.map((r) => r[5]), ['已發放', '正取', '正取', '正取', '正取', '正取', '正取', '備取'],
    'the manual seat counts: 7 seats = 1 manual + 6 grades, so the 科任 applicant waits');
  const before = E.mails.length;
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  const issuedTo = E.mails.slice(before).filter((m) => m.subject.includes('授權核准')).map((m) => m.to);
  assert.ok(issuedTo.indexOf('boss@gmail.com') < 0, 'no mail for the hand-entered row');
  assert.strictEqual(regRows(E).find((r) => r[2] === 'boss@gmail.com')[10], '管理者手動分配');
});

ok('when the domain blocks sharing outside, the form is still created and the admin is told what to set by hand', () => {
  const S = schoolEnv({ drivePolicyBlocks: true });
  S.E.call('setupSheets()');
  S.E.call('setupForm()');
  const form = S.E.createdForms[0];
  assert.ok(form && S.E.props.ENROLL_FORM_ID === form.id, 'form created and bound');
  const msg = S.E.alerts.find((a) => a.includes('已建立新的申請表單'));
  assert.ok(msg.includes('【注意】無法自動開放作答者權限') && msg.includes('知道連結的任何人'), msg);
  assert.ok(msg.includes('已預先勾選'), 'the rest of the setup message still shows');
  const sheet = S.E.ss.sheets.find((s) => s.formUrl === form.url);
  submitTo(S.E, sheet, answersOf(teacher(1, '一年級'), 'Kahoot 教師版'), t(1));
  assert.strictEqual(regRows(S.E).length, 2, 'registration works regardless');
});

ok('quota edited in 軟體清單 takes effect on the next submission', () => {
  const { E, send } = readyEnv();
  E.ss.getSheetByName('軟體清單').getRange(4, 2).setValue(1); // Kahoot -> 1 (row 4: header, TeacherSay, Wayground, Kahoot)
  send(1, teacher(1, '一年級'), 'Kahoot 教師版');
  send(2, teacher(2, '一年級'), 'Kahoot 教師版');
  assert.deepStrictEqual(regRows(E).filter((r) => r[4] === 'Kahoot 教師版').map((r) => r[5]), ['正取', '備取']);
});

console.log('\n' + passed + ' scenario groups passed');
