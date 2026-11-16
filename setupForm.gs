/**
 * setupForm.gs — 初次設定：分頁、申請表單、表單剩餘名額、觸發器
 *
 * 使用方式 (試算表選單「授權管理 → 初次設定」，依序執行)：
 *   1. 初始化分頁                          setupSheets()
 *   2. 建立新申請表單（已建立則更新）       setupForm()
 *   3. 建立表單提交觸發                    createFormSubmitTrigger()
 *
 * setupForm() 第一次執行會建立一份「全新」的表單，回應寫入本試算表的新分頁。
 * 舊表單與舊回應分頁完全不動；程式只處理這份新表單的回應，所以不會和舊資料重複。
 * 「申請登記」若已有舊表單留下的登記，會改名封存並另開空白分頁 (執行前會先詢問)。
 * 再次執行 setupForm() 會就地更新同一份表單 (說明、題目、選項、跳題、剩餘名額)，不會再建立新表單。
 * 要重新建立一份表單時，先執行 resetEnrollmentForm() 解除綁定。
 *
 * 表單結構與跳題：
 *   區段 1 基本資料：姓名、開通授權的 Google 帳號、職務
 *     ├ 學年 → 區段 2【學年】任教學年 (一～六年級)
 *     ├ 科任 → 區段 3【科任】任教科目 (簡答)
 *     └ 行政 → 區段 4【行政】所屬處室
 *   區段 2、3、4 填完都到 區段 5【TeacherSay 專業版】：必填核取方塊 (預填連結會預先勾選)
 *     未勾選就無法前往下一頁，也就看不到其他軟體選項
 *   區段 6 申請軟體授權：其他軟體 (複選，選填)
 *     是否申請 Google AI Pro 教育版？ 是 → 區段 7；否 → 送出
 *   區段 7【Google AI Pro 教育版】：學校帳號 (必填，限 @SCHOOL_DOMAIN)
 */

const FORM_TITLE = '教學軟體授權申請表';
const TEACHERSAY_NOTICE = '本次軟體使用申請需協助提交使用成果，同時配合取得並使用TeacherSay。';
const FORM_DESCRIPTION =
  '請勾選要申請的教學軟體授權，每人每項限登記 1 個。\n' +
  TEACHERSAY_NOTICE + '\n' +
  '系統依登記先後排序，名額內為正取，額滿後登記為備取。\n' +
  '登記不代表已取得授權：資訊組審核後，會寄信到您填寫的 Google 帳號，通知授權序號或開通方式。\n' +
  'Google AI Pro 教育版限' + LIMIT_PER_GROUP_TEXT + '可申請，並須填寫學校帳號（@' + SCHOOL_DOMAIN + '）。';

// 區段標題 (setupForm 以標題辨識既有表單上的區段)
const PAGE_SUBJECT = '【科任】';
const PAGE_OFFICE = '【行政】';
const PAGE_TEACHERSAY = '【TeacherSay 專業版】';
const PAGE_SOFTWARE = '申請軟體授權';
const PAGE_AI_PRO = '【Google AI Pro 教育版】';

// =====================================================================
// 1. 初始化分頁
// =====================================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];

  // 軟體清單：已存在則不覆蓋，保留管理者調整過的數量
  let softwareSheet = ss.getSheetByName(SOFTWARE_SHEET_NAME);
  if (!softwareSheet) {
    softwareSheet = ss.insertSheet(SOFTWARE_SHEET_NAME);
    const headers = ['軟體名稱', '可用授權數', '分配方式（序號／帳號）', '申請限制（有填＝每組限一位）'];
    softwareSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    softwareSheet.getRange(2, 1, DEFAULT_SOFTWARE.length, 4).setValues(DEFAULT_SOFTWARE);
    softwareSheet.getRange(2, 3, 50, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList([MODE.SERIAL, MODE.ACCOUNT], true).build());
    softwareSheet.setFrozenRows(1);
    created.push(SOFTWARE_SHEET_NAME);
  }
  const addedSoftware = ensureRequiredSoftwareRow_(softwareSheet);

  // 申請登記
  const hadRegSheet = !!ss.getSheetByName(REG_SHEET_NAME);
  const regSheet = getRegSheet_(ss);
  if (!hadRegSheet) created.push(REG_SHEET_NAME);
  prepareRegSheet_(regSheet);

  // 序號列表：補上 D 欄「軟體」標題；不存在則建立
  let licenseSheet = ss.getSheetByName(LICENSE_SHEET_NAME);
  if (!licenseSheet) {
    licenseSheet = ss.insertSheet(LICENSE_SHEET_NAME);
    licenseSheet.getRange(1, 1, 1, 4).setValues([['序號', '使用者Email', '分配時間', '軟體']]);
    licenseSheet.setFrozenRows(1);
    created.push(LICENSE_SHEET_NAME);
  } else if (String(licenseSheet.getRange(1, 4).getValue()).trim() === '') {
    licenseSheet.getRange(1, 4).setValue('軟體');
  }

  const softwareList = getSoftwareList_(ss);
  refreshStats_(ss, softwareList, readRegistrations_(regSheet));

  notify_('分頁初始化完成。' + (created.length ? '\n新增分頁：' + created.join('、') : '') +
    (addedSoftware ? '\n「' + SOFTWARE_SHEET_NAME + '」已加入「' + REQUIRED_SOFTWARE + '」（可用授權數：' + UNLIMITED_TEXT + '）。' : '') +
    '\n\n請確認「' + SOFTWARE_SHEET_NAME + '」的可用授權數與分配方式。' +
    '分配方式為「序號」的軟體，請在「' + LICENSE_SHEET_NAME + '」貼上序號並於 D 欄標示軟體 (留空視為 ' +
    DEFAULT_SERIAL_SOFTWARE + ')。');
}

/** 既有的「軟體清單」沒有必選軟體 (TeacherSay) 時補上一列。回傳是否有新增。 */
function ensureRequiredSoftwareRow_(softwareSheet) {
  const lastRow = softwareSheet.getLastRow();
  const names = lastRow < 2 ? [] : parseSoftwareRows_(softwareSheet.getRange(2, 1, lastRow - 1, 4).getValues())
    .map(function (sw) { return normalize_(sw.name); });
  if (names.indexOf(normalize_(REQUIRED_SOFTWARE)) >= 0) return false;
  const row = DEFAULT_SOFTWARE.filter(function (r) { return r[0] === REQUIRED_SOFTWARE; })[0];
  softwareSheet.getRange(Math.max(lastRow, 1) + 1, 1, 1, 4).setValues([row]);
  return true;
}

/** 標題列更新為目前版本的欄位，並設定「狀態」欄的下拉選單。不會動到資料列。 */
function prepareRegSheet_(regSheet) {
  regSheet.getRange(1, 1, 1, REG_HEADERS.length).setValues([REG_HEADERS]);
  regSheet.getRange(2, COL.STATUS, regSheet.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([STATUS.OK, STATUS.WAIT, STATUS.INELIGIBLE, STATUS.CANCEL, STATUS.ISSUED], true)
      .setAllowInvalid(true)
      .build());
}

/**
 * 換新表單時，把舊表單留下的登記資料整個分頁改名封存，另開空白的「申請登記」。
 * 沒有資料列時不封存。回傳封存後的分頁名稱 (沒有封存則為空字串)。
 */
function archiveRegistrations_(ss) {
  const sheet = ss.getSheetByName(REG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const name = REG_SHEET_NAME + '_舊_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss');
  sheet.setName(name);
  return name;
}

// =====================================================================
// 2. 建立新申請表單（已建立則更新選項）
// =====================================================================

function setupForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const softwareList = getSoftwareList_(ss);
  const props = PropertiesService.getScriptProperties();

  let form = null;
  try {
    form = getEnrollmentForm_();
  } catch (error) {
    notify_('找不到先前建立的申請表單（可能已被刪除）：' + error +
      '\n\n若要重新建立，請先在編輯器執行 resetEnrollmentForm()，再執行本功能。');
    return;
  }

  let message;
  if (form) {
    // 已建立：就地更新同一份表單
    let changes;
    try {
      changes = syncEnrollmentForm_(form, softwareList);
    } catch (error) {
      notify_('無法更新申請表單：' + (error.message || error) +
        '\n\n請還原表單上的題目或區段名稱，或執行 resetEnrollmentForm() 後重新建立表單。');
      return;
    }
    message = '已更新申請表單。' + (changes.length ? '\n' + changes.join('\n') : '');
  } else {
    // 第一次：建立全新的表單，回應寫入本試算表的新分頁
    const oldRegSheet = ss.getSheetByName(REG_SHEET_NAME);
    const oldRows = oldRegSheet ? Math.max(0, oldRegSheet.getLastRow() - 1) : 0;
    if (oldRows > 0 && !confirm_('建立新申請表單',
      '將建立一份全新的申請表單，回應寫入本試算表的新分頁。\n\n' +
      '目前「' + REG_SHEET_NAME + '」已有 ' + oldRows + ' 筆來自舊表單的登記，會整個分頁改名封存 (資料保留、不再計入名額)，' +
      '並另開空白的「' + REG_SHEET_NAME + '」。\n\n要繼續嗎？')) {
      return;
    }
    const archivedName = archiveRegistrations_(ss);
    const sheetIdsBefore = ss.getSheets().map(function (sheet) { return sheet.getSheetId(); });
    form = FormApp.create(FORM_TITLE);
    buildEnrollmentForm_(form);
    syncEnrollmentForm_(form, softwareList);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    props.setProperty(PROP.FORM_ID, form.getId());
    props.deleteProperty(PROP.RESPONSE_SHEET_ID);

    const responseSheet = findNewResponseSheet_(ss, sheetIdsBefore);
    if (responseSheet) props.setProperty(PROP.RESPONSE_SHEET_ID, String(responseSheet.getSheetId()));
    message = '已建立新的申請表單，回應會寫入' +
      (responseSheet ? '新分頁「' + responseSheet.getName() + '」。' : '本試算表的新分頁。') +
      '\n舊表單與舊回應分頁沒有變動，它們的回應不會被登記；不再使用時請自行關閉舊表單。' +
      (archivedName ? '\n原有的登記資料已封存為分頁「' + archivedName + '」。' : '');
  }

  const regSheet = getRegSheet_(ss);
  prepareRegSheet_(regSheet);
  const rows = readRegistrations_(regSheet);
  const stats = refreshStats_(ss, softwareList, rows);
  findItemByTitle_(form, Q_SOFTWARE).setHelpText(buildRemainingText_(stats, nowText_()));

  const prefillUrl = buildPrefilledUrl_(form);
  props.setProperty(PROP.PREFILL_URL, prefillUrl);

  notify_(message +
    '\n\n請發給老師這個網址（' + REQUIRED_SOFTWARE + ' 已預先勾選）：\n' + prefillUrl +
    '\n\n一般填寫網址：' + form.getPublishedUrl() +
    '\n編輯網址：' + form.getEditUrl() +
    '\n\n之後可從選單「授權管理 → 顯示申請表單網址」再次取得。下一步：執行「3. 建立表單提交觸發」。');
}

/**
 * 建立區段 1～4 與職務跳題。題目順序即回應分頁的欄位順序。
 * TeacherSay、申請軟體授權、Google AI Pro 三個區段與區段間的跳轉由 syncEnrollmentForm_() 建立。
 */
function buildEnrollmentForm_(form) {
  const emailValidation = FormApp.createTextValidation()
    .setHelpText('請輸入有效的 Email 地址').requireTextIsEmail().build();

  // ---------- 區段 1：基本資料 ----------
  form.addTextItem().setTitle(Q_NAME).setRequired(true);
  form.addTextItem().setTitle(Q_EMAIL)
    .setHelpText('多數軟體以 Google 帳號單一登入，請填寫您實際用來登入軟體的 Google 帳號。授權通知也會寄到這個信箱。')
    .setValidation(emailValidation)
    .setRequired(true);
  const roleItem = form.addMultipleChoiceItem().setTitle(Q_ROLE).setRequired(true);

  // ---------- 區段 2：學年 ----------
  const pageGrade = form.addPageBreakItem().setTitle('【學年】');
  form.addMultipleChoiceItem().setTitle(Q_GRADE).setChoiceValues(GRADE_CHOICES).setRequired(true);

  // ---------- 區段 3：科任 ----------
  const pageSubject = form.addPageBreakItem().setTitle(PAGE_SUBJECT);
  form.addTextItem().setTitle(Q_SUBJECT).setHelpText('例：自然、社會、英語').setRequired(true);

  // ---------- 區段 4：行政 ----------
  const pageOffice = form.addPageBreakItem().setTitle(PAGE_OFFICE);
  form.addMultipleChoiceItem().setTitle(Q_OFFICE).setChoiceValues(OFFICE_CHOICES).setRequired(true);

  // ---------- 職務跳題 ----------
  roleItem.setChoices([
    roleItem.createChoice(ROLE.GRADE, pageGrade),
    roleItem.createChoice(ROLE.SUBJECT, pageSubject),
    roleItem.createChoice(ROLE.OFFICE, pageOffice)
  ]);
}

/**
 * 讓表單的說明、行政處室選項，以及 TeacherSay、申請軟體授權、Google AI Pro 三個區段符合目前版本。
 * 新表單與 v3.1 建立的表單都適用；可重複執行。缺少的題目會新增並移到正確位置，
 * 既有題目以標題辨識、就地修改，所以回應分頁的既有欄位與資料不受影響。
 * @returns {string[]} 變更說明
 */
function syncEnrollmentForm_(form, softwareList) {
  const T = FormApp.ItemType;
  const changes = [];
  const find = function (title, type, label) {
    const item = findItemByTitle_(form, title);
    if (item && item.getType() !== type) {
      throw new Error('「' + title + '」題型不符 (' + label + ')');
    }
    return item;
  };
  // 回傳一律是 form.getItems() 的通用 Item (新增的題目也重新取一次)，之後統一用 asXxxItem() 轉型
  const ensure = function (title, type, label, create) {
    if (!find(title, type, label)) {
      create().setTitle(title);
      changes.push('・新增「' + title + '」');
    }
    return find(title, type, label);
  };

  form.setDescription(FORM_DESCRIPTION);

  // 行政處室選項 (加入幼兒園)
  const office = find(Q_OFFICE, T.MULTIPLE_CHOICE, '單選');
  if (!office) throw new Error('找不到「' + Q_OFFICE + '」');
  const officeItem = office.asMultipleChoiceItem();
  if (officeItem.getChoices().length !== OFFICE_CHOICES.length) changes.push('・「' + Q_OFFICE + '」選項更新為：' + OFFICE_CHOICES.join('、'));
  officeItem.setChoiceValues(OFFICE_CHOICES);

  // 區段 5【TeacherSay 專業版】：接在行政區段之後，必填、只有一個選項。
  // 必填題未勾選就不能按「下一頁」，所以沒勾選 TeacherSay 就看不到其他軟體選項。
  const pageTeacherSay = ensure(PAGE_TEACHERSAY, T.PAGE_BREAK, '區段', function () { return form.addPageBreakItem(); });
  placeAfter_(form, pageTeacherSay, office);
  pageTeacherSay.asPageBreakItem().setHelpText(TEACHERSAY_NOTICE);
  const teacherSay = ensure(Q_TEACHERSAY, T.CHECKBOX, '核取方塊', function () { return form.addCheckboxItem(); });
  placeAfter_(form, teacherSay, pageTeacherSay);
  teacherSay.asCheckboxItem()
    .setChoiceValues([REQUIRED_SOFTWARE])
    .setHelpText(TEACHERSAY_NOTICE + '\n請勾選後再按「下一頁」，未勾選無法選擇其他軟體，也無法送出表單。')
    .setRequired(true);

  // 區段 6 申請軟體授權：其他軟體 (TeacherSay 與 Google AI Pro 不在這題)
  const pageSoftware = ensure(PAGE_SOFTWARE, T.PAGE_BREAK, '區段', function () { return form.addPageBreakItem(); });
  placeAfter_(form, pageSoftware, teacherSay);
  const software = ensure(Q_SOFTWARE, T.CHECKBOX, '核取方塊', function () { return form.addCheckboxItem(); });
  placeAfter_(form, software, pageSoftware);
  const otherNames = softwareList
    .map(function (sw) { return sw.name; })
    .filter(function (name) { return name !== REQUIRED_SOFTWARE && name !== WORKSPACE_ACCOUNT_SOFTWARE; });
  if (otherNames.length > 0) software.asCheckboxItem().setChoiceValues(otherNames);
  software.asCheckboxItem().setRequired(false);

  // 學年、科任區段填完都前往 TeacherSay 區段 (行政區段之後依序就是 TeacherSay 區段)。
  // 分頁符號的 setGoToPage：填完「前一個區段」後要前往的區段。
  [PAGE_SUBJECT, PAGE_OFFICE].forEach(function (title) {
    const page = find(title, T.PAGE_BREAK, '區段');
    if (!page) throw new Error('找不到「' + title + '」區段');
    page.asPageBreakItem().setGoToPage(pageTeacherSay.asPageBreakItem());
  });

  // Google AI Pro：是 → 學校帳號頁；否 → 送出
  const aiPro = ensure(Q_AI_PRO, T.MULTIPLE_CHOICE, '單選', function () { return form.addMultipleChoiceItem(); });
  placeAfter_(form, aiPro, software);
  const pageAi = ensure(PAGE_AI_PRO, T.PAGE_BREAK, '區段', function () { return form.addPageBreakItem(); });
  placeAfter_(form, pageAi, aiPro);
  const aiItem = aiPro.asMultipleChoiceItem();
  aiItem.setChoices([
    aiItem.createChoice(AI_YES, pageAi.asPageBreakItem()),
    aiItem.createChoice(AI_NO, FormApp.PageNavigationType.SUBMIT)
  ]);
  aiItem.setHelpText('選填。限' + LIMIT_PER_GROUP_TEXT + '（科任不分科目、行政含幼兒園各算一組），依登記先後排序。' +
    '選「' + AI_YES + '」須在下一頁填寫學校帳號才能送出。')
    .setRequired(true);
  pageAi.asPageBreakItem().setHelpText('Google AI Pro 教育版會開通在學校帳號。');

  // 學校帳號：只有選「是」才會看到，必填且限學校網域
  const workspace = ensure(Q_WORKSPACE, T.TEXT, '簡答', function () { return form.addTextItem(); });
  placeAfter_(form, workspace, pageAi);
  workspace.asTextItem()
    .setHelpText('請填寫學校帳號（@' + SCHOOL_DOMAIN + '）。授權會開通在這個帳號，通知信也會副本到這裡。')
    .setValidation(FormApp.createTextValidation()
      .setHelpText('請填寫學校帳號（@' + SCHOOL_DOMAIN + '）')
      .requireTextMatchesPattern(schoolAccountPattern_())
      .build())
    .setRequired(true);

  return changes;
}

/** 把 item 移到 anchor 的正後方。 */
function placeAfter_(form, item, anchor) {
  const from = item.getIndex();
  const anchorIndex = anchor.getIndex();
  if (from === anchorIndex + 1) return;
  form.moveItem(from, from > anchorIndex ? anchorIndex + 1 : anchorIndex);
}

/** 學校帳號的規則運算式：帳號@學校網域，網域不分大小寫。 */
function schoolAccountPattern_() {
  const domain = SCHOOL_DOMAIN.split('').map(function (ch) {
    if (/[a-z]/i.test(ch)) return '[' + ch.toLowerCase() + ch.toUpperCase() + ']';
    return ch === '.' ? '\\.' : ch;
  }).join('');
  return '^[^\\s@]+@' + domain + '$';
}

/** TeacherSay 已預先勾選的填寫網址。 */
function buildPrefilledUrl_(form) {
  const item = findItemByTitle_(form, Q_TEACHERSAY).asCheckboxItem();
  return form.createResponse()
    .withItemResponse(item.createResponse([REQUIRED_SOFTWARE]))
    .toPrefilledUrl();
}

/** setDestination 之後 Google 會在試算表新增回應分頁；稍候並找出它。找不到回傳 null (第一次收到回應時再辨識)。 */
function findNewResponseSheet_(ss, sheetIdsBefore) {
  for (let attempt = 0; attempt < 8; attempt++) {
    SpreadsheetApp.flush();
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheetIdsBefore.indexOf(sheets[i].getSheetId()) < 0 && sheets[i].getFormUrl()) return sheets[i];
    }
    Utilities.sleep(1500);
  }
  return null;
}

function findItemByTitle_(form, title) {
  const want = normalize_(title);
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (normalize_(items[i].getTitle()) === want) return items[i];
  }
  return null;
}

/** 把剩餘名額寫到表單「軟體」題目下方的說明文字。 */
function refreshFormRemaining_(ss, stats) {
  const form = getEnrollmentForm_();
  if (!form) return false;
  const item = findItemByTitle_(form, Q_SOFTWARE);
  if (!item) return false;
  item.setHelpText(buildRemainingText_(stats, nowText_()));
  return true;
}

/**
 * 解除與目前申請表單的綁定 (不會刪除表單與回應)。
 * 之後再執行 setupForm() 會建立另一份新表單。請從編輯器執行。
 */
function resetEnrollmentForm() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP.FORM_ID);
  props.deleteProperty(PROP.RESPONSE_SHEET_ID);
  Logger.log('已解除申請表單綁定。原表單與回應分頁仍保留；在重新執行 setupForm() 之前，所有表單回應都不會被登記。');
}

// =====================================================================
// 3. 觸發器
// =====================================================================

function createFormSubmitTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'onFormSubmit' &&
      trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT;
  });
  if (exists) {
    notify_('「提交表單時」觸發已存在，沿用原有的觸發。');
    return;
  }
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  notify_('已建立「提交表單時」觸發 onFormSubmit。');
}
