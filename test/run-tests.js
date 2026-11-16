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
  constructor(title, id) { this.title = title; this.items = []; this.description = ''; this.id = id; this.url = 'https://forms/' + id; }
  getId() { return this.id; }
  getItems() { return this.items.map((it) => it.generic()); }
  add(type) { const it = new Item(this, type); this.items.push(it); return it; }
  addTextItem() { return this.add(ItemType.TEXT); }
  addCheckboxItem() { return this.add(ItemType.CHECKBOX); }
  addMultipleChoiceItem() { return this.add(ItemType.MULTIPLE_CHOICE); }
  addPageBreakItem() { return this.add(ItemType.PAGE_BREAK); }
  deleteItem(it) { this.items.splice(this.items.indexOf(it), 1); }
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
  subject: '任教科目', office: '所屬處室', teacherSay: 'TeacherSay 專業版（必選）', software: '您要申請的軟體授權',
  aiPro: '是否申請 Google AI Pro 教育版？', workspace: '申請 Google AI Pro 教育版的學校帳號',
};
const SCHOOL = '@mail.lres.tc.edu.tw';
const TS = 'TeacherSay 專業版';
const AI = 'Google AI Pro 教育版';
// who: {name, email, role:'學年'|'科任'|'行政', detail, workspace}; software: the programs ticked, comma separated.
// Google AI Pro is answered on its own yes/no question and TeacherSay on its own required checkbox,
// exactly as the v3.2 form asks them. opts.noTeacherSay simulates a response that skipped it.
function answersOf(who, software, opts) {
  opts = opts || {};
  const list = String(software || '').split(',').map((x) => x.trim()).filter(Boolean);
  const ai = list.indexOf(AI) >= 0;
  const a = {
    [Q.name]: who.name, [Q.email]: who.email, [Q.role]: who.role,
    [Q.teacherSay]: opts.noTeacherSay ? '' : TS,
    [Q.software]: list.filter((x) => x !== AI && x !== TS).join(', '),
    [Q.aiPro]: who.role === '行政' ? '' : (ai ? '是' : '否'),
    [Q.workspace]: who.role !== '行政' && ai ? (who.workspace || '') : '',
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
      if (v && q.validation && q.validation.kind === 'pattern' && !new RegExp(q.validation.pattern).test(v)) {
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
const FORM_LAYOUT = [Q.name, Q.email, '#【TeacherSay 專業版】', Q.teacherSay, '#申請軟體授權', Q.software, '#【職務】', Q.role,
  '#【學年】', Q.grade, '#【科任】', Q.subject, '#【行政】', Q.office, '#【Google AI Pro 教育版】', Q.aiPro, '#【Google AI Pro 學校帳號】', Q.workspace];
const titlesOf = (form) => form.items.map((i) => (i.type === 'PAGE_BREAK' ? '#' + i.title : i.title));

// ============ 1. unit tests shipped to the user ============
ok('runEnrollmentTests (shipped unit tests)', () => {
  const E = makeEnv();
  const res = E.call('runEnrollmentTests()');
  assert.ok(res.length >= 60 && res.every((r) => r.ok));
  console.log('     ' + res.length + ' checks');
});

// ============ 2. a brand-new form, old form untouched ============
ok('setupForm creates a NEW form: TeacherSay gate, software, role branches, AI Pro for 學年/科任 only, prefilled link', () => {
  const { E, oldForm, oldSheet, form, sheet, lic } = readyEnv();
  assert.strictEqual(E.createdForms.length, 1);
  assert.notStrictEqual(form, oldForm);
  assert.deepStrictEqual(oldForm.items.map((i) => i.title), ['您的大名', '您的常用email信箱', '您的職務', '您要申請的軟體授權']);
  assert.strictEqual(oldSheet.getLastRow(), 2);

  assert.deepStrictEqual(titlesOf(form), FORM_LAYOUT);
  assert.ok(form.description.includes('本次軟體使用申請需協助提交使用成果，同時配合取得並使用TeacherSay。'));
  const role = item(form, Q.role);
  const page = (t) => form.items.find((i) => i.type === 'PAGE_BREAK' && i.title === t);
  assert.deepStrictEqual(role.nav, [page('【學年】'), page('【科任】'), page('【行政】')]);
  assert.deepStrictEqual(item(form, Q.office).choices, ['教務處', '學務處', '總務處', '輔導室', '幼兒園']);
  assert.ok(page('【行政】').help.includes('由資訊組另行分配'));
  assert.strictEqual(page('【科任】').goTo, page('【Google AI Pro 教育版】'), 'after 學年 -> AI Pro');
  assert.strictEqual(page('【行政】').goTo, page('【Google AI Pro 教育版】'), 'after 科任 -> AI Pro');
  assert.strictEqual(page('【Google AI Pro 教育版】').goTo, 'SUBMIT', 'after 行政 -> submit');
  assert.strictEqual(item(form, Q.subject).type, 'TEXT');

  const ts = item(form, Q.teacherSay);
  assert.deepStrictEqual([ts.type, ts.choices, ts.required], ['CHECKBOX', [TS], true]);
  assert.ok(page('【TeacherSay 專業版】').help.includes('TeacherSay'));
  const sw = item(form, Q.software);
  assert.deepStrictEqual([sw.type, sw.required], ['CHECKBOX', false]);
  assert.deepStrictEqual(sw.choices, ['Wayground (Quizizz)', 'Kahoot 教師版', 'Padlet 學校版', 'Wordwall 專業版', 'HiTeach 專業版']);
  assert.ok(sw.help.includes('・TeacherSay 專業版：不限名額（每位申請者必選）') && sw.help.includes('・HiTeach 專業版：剩餘 15／15'));
  const ai = item(form, Q.aiPro);
  assert.deepStrictEqual([ai.type, ai.required, ai.choices, ai.nav], ['MULTIPLE_CHOICE', true, ['是', '否'], [page('【Google AI Pro 學校帳號】'), 'SUBMIT']]);
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
  assert.strictEqual(sheet.name, '表單回應 2', 'responses go to a new tab of the same spreadsheet');
  assert.strictEqual(E.props.ENROLL_FORM_ID, form.id);
  assert.strictEqual(E.props.ENROLL_RESPONSE_SHEET_ID, String(sheet.id));
  assert.strictEqual(lic.get(1, 4), '軟體');
  assert.strictEqual(E.triggers.length, 1, 'trigger must not be duplicated');

  // running it again updates the same form instead of creating another
  E.ss.getSheetByName('軟體清單').appendRow(['Canva 教育版', 3, '帳號', '']);
  E.call('setupForm()');
  assert.strictEqual(E.createdForms.length, 1);
  assert.deepStrictEqual(titlesOf(form), FORM_LAYOUT, 'no duplicated questions');
  assert.strictEqual(sw.choices[sw.choices.length - 1], 'Canva 教育版');
  assert.ok(sw.help.includes('・Canva 教育版：剩餘 3／3'));
});

ok('respondent paths: TeacherSay gate first; 學年/科任 reach Google AI Pro; 行政 submits without ever seeing it', () => {
  const { form } = readyEnv();
  const base = (role, detail) => answersOf({ name: '甲', email: 'a@gmail.com', role, detail, workspace: 'a' + SCHOOL }, 'Kahoot 教師版');
  const T = '【TeacherSay 專業版】', W = '申請軟體授權', R = '【職務】', G = '【學年】', S = '【科任】', O = '【行政】',
    AQ = '【Google AI Pro 教育版】', AW = '【Google AI Pro 學校帳號】';
  assert.deepStrictEqual(walkForm(form, base('學年', '三年級')), { visited: ['基本資料', T, W, R, G, AQ], submitted: true });
  assert.deepStrictEqual(walkForm(form, base('科任', '自然')), { visited: ['基本資料', T, W, R, S, AQ], submitted: true });
  assert.deepStrictEqual(walkForm(form, base('行政', '幼兒園')), { visited: ['基本資料', T, W, R, O], submitted: true });
  assert.deepStrictEqual(walkForm(form, base('行政', '教務處')).visited.indexOf(AQ), -1, '行政 never sees the AI Pro question');
  // TeacherSay unticked: stuck on the TeacherSay page, the software page is never shown
  const noTs = Object.assign(base('學年', '一年級'), { [Q.teacherSay]: '' });
  assert.deepStrictEqual(walkForm(form, noTs), { visited: ['基本資料', T], submitted: false, blockedAt: Q.teacherSay });
  // AI Pro yes: the school-account page is shown and must hold a school account
  const aiYes = answersOf(teacher(1, '二年級', { workspace: 't1' + SCHOOL }), AI);
  assert.deepStrictEqual(walkForm(form, aiYes), { visited: ['基本資料', T, W, R, G, AQ, AW], submitted: true });
  assert.strictEqual(walkForm(form, Object.assign({}, aiYes, { [Q.workspace]: '' })).blockedAt, Q.workspace);
  assert.strictEqual(walkForm(form, Object.assign({}, aiYes, { [Q.workspace]: 't1@gmail.com' })).blockedAt, Q.workspace);
  // AI Pro must be answered by 學年/科任
  assert.strictEqual(walkForm(form, Object.assign({}, aiYes, { [Q.aiPro]: '' })).blockedAt, Q.aiPro);
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
  assert.strictEqual(S.E.props.ENROLL_RESPONSE_SHEET_ID, String(tab.id), 'same response tab');
  assert.deepStrictEqual(titlesOf(f), FORM_LAYOUT);
  const aiPage = f.items.find((i) => i.title === '【Google AI Pro 教育版】');
  assert.strictEqual(pS.goTo, aiPage, '學年 section now goes to AI Pro');
  assert.strictEqual(pO.goTo, aiPage, '科任 section now goes to AI Pro');
  assert.strictEqual(aiPage.goTo, 'SUBMIT', '行政 section now submits');
  assert.strictEqual(pW.goTo, undefined);
  assert.deepStrictEqual(role.nav, [pG, pS, pO], 'role choices still jump to their sections');
  assert.deepStrictEqual(item(f, Q.office).choices.slice(-1), ['幼兒園']);
  assert.deepStrictEqual(item(f, Q.software).choices, ['Wayground (Quizizz)', 'Kahoot 教師版', 'Padlet 學校版', 'Wordwall 專業版', 'HiTeach 專業版']);
  assert.strictEqual(item(f, Q.software).required, false);
  assert.strictEqual(item(f, Q.workspace).required, true);
  assert.ok(S.E.alerts[S.E.alerts.length - 1].includes('已更新申請表單') && S.E.alerts[S.E.alerts.length - 1].includes('新增「' + Q.teacherSay + '」'));
  const walked = walkForm(f, answersOf({ name: '甲', email: 'a@gmail.com', role: '行政', detail: '幼兒園' }, 'Padlet 學校版'));
  assert.deepStrictEqual(walked, { visited: ['基本資料', '【TeacherSay 專業版】', '申請軟體授權', '【職務】', '【行政】'], submitted: true });

  // the pre-upgrade response was registered by the trigger; a new one goes through the new questions
  let rows = regRows(S.E);
  assert.deepStrictEqual(rows.map((r) => r[4]), ['Kahoot 教師版', AI]);
  submitTo(S.E, tab, answersOf({ name: '新', email: 'n@gmail.com', role: '行政', detail: '幼兒園' }, 'Wordwall 專業版'), t(1));
  // a v3.1 response from 行政 that ticked AI Pro in the old checkbox is still refused
  submitTo(S.E, tab, { [Q.name]: '舊行政', [Q.email]: 'o@gmail.com', [Q.role]: '行政', [Q.office]: '教務處',
    [Q.software]: AI, [Q.workspace]: 'o' + SCHOOL }, t(2));
  rows = regRows(S.E);
  assert.deepStrictEqual(rows.slice(2).map((r) => [r[3], r[4], r[5]]),
    [['行政／幼兒園', 'Wordwall 專業版', '正取'], ['行政／幼兒園', TS, '正取'], ['行政／教務處', AI, '不符資格']],
    'order follows 軟體清單, where the upgrade appended TeacherSay at the bottom');
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
ok('registration: TeacherSay for everyone, ranks, waitlist, AI Pro rules, receipts with CC, remaining seats', () => {
  const { E, form, send } = readyEnv();
  for (let i = 1; i <= 5; i++) send(i, teacher(i, '一年級'), 'Kahoot 教師版, HiTeach 專業版');
  send(6, teacher(6, '二年級', { workspace: 't6' + SCHOOL }), 'Kahoot 教師版, ' + AI);
  send(7, teacher(7, '二年級', { workspace: 't7' + SCHOOL }), AI);
  send(8, { name: '科任8', email: 't8@gmail.com', role: '科任', detail: '自然', workspace: 't8' + SCHOOL }, AI);
  send(9, { name: '科任9', email: 't9@gmail.com', role: '科任', detail: '社會', workspace: 't9' + SCHOOL }, AI);
  send(10, { name: '幼10', email: 't10@gmail.com', role: '行政', detail: '幼兒園', workspace: 't10' + SCHOOL }, AI + ', Wordwall 專業版'); // AI Pro never asked
  send(11, teacher(1, '一年級', { email: 'T1@gmail.com' }), 'Kahoot 教師版, Padlet 學校版'); // duplicate Kahoot + TeacherSay
  send(12, teacher(12, '三年級'), ''); // TeacherSay only

  const rows = regRows(E);
  assert.strictEqual(rows.filter((r) => r[4] === TS).length, 11, 'one TeacherSay row per person');
  assert.ok(rows.filter((r) => r[4] === TS).every((r) => r[5] === '正取'), 'TeacherSay is unlimited');
  // TeacherSay 11 + Kahoot 6 (the duplicate is not added) + HiTeach 5 + AI Pro 4 (not 行政) + Wordwall 1 + Padlet 1
  assert.strictEqual(rows.length, 11 + 6 + 5 + 4 + 1 + 1);
  assert.deepStrictEqual(rowOf(rows, 't5@gmail.com', 'Kahoot 教師版').slice(5, 7), ['正取', 5]);
  assert.deepStrictEqual(rowOf(rows, 't6@gmail.com', 'Kahoot 教師版').slice(5, 7), ['備取', 1]);
  assert.strictEqual(rowOf(rows, 't1@gmail.com', 'HiTeach 專業版')[3], '學年／一年級');
  assert.strictEqual(rowOf(rows, 't8@gmail.com', AI)[3], '科任／自然');
  assert.strictEqual(rowOf(rows, 't10@gmail.com', 'Wordwall 專業版')[3], '行政／幼兒園');
  assert.strictEqual(rowOf(rows, 't6@gmail.com', AI)[5], '正取');
  assert.strictEqual(rowOf(rows, 't7@gmail.com', AI)[5], '備取');
  assert.strictEqual(rowOf(rows, 't8@gmail.com', AI)[5], '正取');
  assert.strictEqual(rowOf(rows, 't9@gmail.com', AI)[5], '備取', 'second 科任 waits even with another subject');
  assert.strictEqual(rowOf(rows, 't10@gmail.com', AI), undefined, '行政 cannot apply for AI Pro on the form');
  assert.strictEqual(rowOf(rows, 't10@gmail.com', 'Wordwall 專業版')[11], '', 'no school account recorded for 行政');
  assert.strictEqual(rowOf(rows, 't6@gmail.com', AI)[11], 't6' + SCHOOL);
  assert.strictEqual(rowOf(rows, 't6@gmail.com', AI)[10], '');
  assert.strictEqual(rowOf(rows, 't12@gmail.com', TS)[5], '正取');
  assert.ok(rows.every((r) => r[7] === false && r[8] === ''), 'no approval, no serial at registration');

  const receipts = E.mails.filter((m) => m.subject === '【已登記】教學軟體授權申請');
  assert.strictEqual(receipts.length, 12);
  assert.ok(receipts.every((m) => m.htmlBody.includes('<b>TeacherSay 專業版</b>')));
  assert.deepStrictEqual([receipts[0].to, receipts[0].cc], ['t1@gmail.com', undefined]);
  assert.deepStrictEqual([receipts[5].to, receipts[5].cc], ['t6@gmail.com', 't6' + SCHOOL]);
  assert.ok(receipts[5].htmlBody.includes('備取（備取第 1 位）'));
  assert.ok(!receipts[5].htmlBody.includes('未填寫學校帳號'));
  assert.ok(!receipts[9].htmlBody.includes(AI) && receipts[9].cc === undefined, '行政 receipt has no AI Pro and no CC');
  assert.ok(receipts[10].htmlBody.includes('您先前已登記過') && receipts[10].htmlBody.includes('<b>Padlet 學校版</b>：正取（第 1 位）'));
  const adminMails = E.mails.filter((m) => m.to === 'lres712@mail.lres.tc.edu.tw');
  assert.strictEqual(adminMails.length, 2);

  const help = item(form, Q.software).help;
  assert.ok(help.includes('・Kahoot 教師版：已額滿（備取 1 人）'), help);
  assert.ok(help.includes('・HiTeach 專業版：剩餘 10／15'), help);
  assert.ok(help.includes('・TeacherSay 專業版：不限名額（每位申請者必選）'), help);
  assert.ok(help.includes('・Google AI Pro 教育版：剩餘 5／7｜限各學年、科任各一位，尚可申請：一年級、三年級、四年級、五年級、六年級'), help);
  const stats = E.ss.getSheetByName('申請統計').getRange(1, 1, 8, 10).getValues();
  assert.deepStrictEqual(stats.find((r) => r[0] === 'Kahoot 教師版').slice(1, 8), [5, 6, 5, 1, 0, 0, 0]);
  assert.deepStrictEqual(stats.find((r) => r[0] === TS).slice(1, 8), ['不限', 11, 11, 0, 0, 0, '不限']);
});

// ============ 6. approval + issuing ============
ok('issuing: approved rows only, HiTeach serials, account wording, one mail per person with CC, cancel promotes', () => {
  const { E, lic, ww, send } = readyEnv();
  send(1, { name: '甲', email: 'a@gmail.com', role: '學年', detail: '一年級', workspace: 'a' + SCHOOL }, 'HiTeach 專業版, Wordwall 專業版, ' + AI);
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
  assert.deepStrictEqual(rowOf(rows, 'a@gmail.com', 'HiTeach 專業版').slice(5, 9), ['已發放', '', true, 'HT-A']);
  assert.deepStrictEqual(rowOf(rows, 'b@gmail.com', 'HiTeach 專業版')[8], 'HT-B');
  assert.deepStrictEqual([rowOf(rows, 'c@gmail.com', 'HiTeach 專業版')[5], rowOf(rows, 'c@gmail.com', 'HiTeach 專業版')[8]], ['正取', '']);
  assert.deepStrictEqual([rowOf(rows, 'b@gmail.com', 'Kahoot 教師版')[5], rowOf(rows, 'b@gmail.com', 'Kahoot 教師版')[8]], ['已發放', '']);
  assert.ok(['a', 'b', 'c', 'd'].every((p) => rowOf(rows, p + '@gmail.com', TS)[5] === '已發放'));
  assert.strictEqual(rowOf(rows, 'a@gmail.com', AI)[5], '已發放');
  assert.deepStrictEqual([lic.get(2, 1), lic.get(2, 2)], ['HT-OLD', 'old@x.tw']);
  assert.deepStrictEqual([lic.get(3, 2), lic.get(4, 2)], ['a@gmail.com', 'b@gmail.com']);
  assert.deepStrictEqual(ww.getRange(2, 1, 1, 2).getValues()[0], ['甲', 'a@gmail.com']);

  const issued = E.mails.slice(before).filter((m) => m.subject === '【授權核准】教學軟體授權已發放');
  assert.strictEqual(issued.length, 4, 'one mail per person');
  const mailA = issued.find((m) => m.to === 'a@gmail.com');
  assert.strictEqual(mailA.cc, 'a' + SCHOOL);
  assert.ok(mailA.htmlBody.includes('HT-A') && !mailA.htmlBody.includes('HT-B'));
  assert.ok(mailA.htmlBody.includes('<b>Google AI Pro 教育版</b>－授權將開通於您的學校帳號（a' + SCHOOL + '）'));
  assert.ok(mailA.htmlBody.includes('<b>TeacherSay 專業版</b>－授權將開通於您登記的 Google 帳號（a@gmail.com）'));
  assert.strictEqual(issued.find((m) => m.to === 'b@gmail.com').cc, undefined);
  assert.ok(E.alerts[E.alerts.length - 1].includes('沒有可用序號'));

  const n = E.mails.length;
  E.call('menuIssueApproved()');
  assert.strictEqual(E.mails.length, n, 're-running sends nothing new');

  const reg = E.ss.getSheetByName('申請登記');
  const cRow = rows.findIndex((r) => r[2] === 'c@gmail.com' && r[4] === 'HiTeach 專業版') + 2;
  reg.getRange(cRow, 6).setValue('取消');
  E.call('menuRerank()');
  rows = regRows(E);
  assert.strictEqual(rows[cRow - 2][5], '取消');
  assert.deepStrictEqual(rowOf(rows, 'd@gmail.com', 'HiTeach 專業版').slice(5, 7), ['正取', 1]);
});

ok('mail quota counts CC addresses and blocks before any write', () => {
  const { E, lic, send } = readyEnv({ mailQuota: 1 });
  send(1, { name: '甲', email: 'a@gmail.com', role: '科任', detail: '自然', workspace: 'a' + SCHOOL }, 'HiTeach 專業版, ' + AI);
  E.call('menuApproveAllAccepted()');
  E.call('menuIssueApproved()');
  assert.ok(E.alerts[E.alerts.length - 1].includes('含副本共 2 個收件人'));
  assert.strictEqual(lic.get(3, 2), '');
  assert.ok(regRows(E).every((r) => r[5] === '正取'));
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
  assert.ok(item(form, Q.software).help.includes('・Kahoot 教師版：剩餘 5／5'), 'archived rows no longer use seats');
  submitTo(S.E, sheet, answersOf(teacher(1, '一年級', { workspace: 't1' + SCHOOL }), 'Kahoot 教師版, ' + AI), t(1));
  assert.deepStrictEqual(regRows(S.E).map((r) => [r[4], r[5], r[6]]), [[TS, '正取', 1], ['Kahoot 教師版', '正取', 1], [AI, '正取', 1]]);

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

ok('quota edited in 軟體清單 takes effect on the next submission', () => {
  const { E, send } = readyEnv();
  E.ss.getSheetByName('軟體清單').getRange(4, 2).setValue(1); // Kahoot -> 1 (row 4: header, TeacherSay, Wayground, Kahoot)
  send(1, teacher(1, '一年級'), 'Kahoot 教師版');
  send(2, teacher(2, '一年級'), 'Kahoot 教師版');
  assert.deepStrictEqual(regRows(E).filter((r) => r[4] === 'Kahoot 教師版').map((r) => r[5]), ['正取', '備取']);
});

console.log('\n' + passed + ' scenario groups passed');
