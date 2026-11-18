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
 * 表單結構與跳題 (每次填表最多 2 項：TeacherSay ＋ 另選 1 項)：
 *   區段 1 基本資料：姓名、開通授權的 Google 帳號
 *   區段 2【TeacherSay 專業版】：必填核取方塊 (預填連結會預先勾選)
 *     未勾選就無法前往下一頁，也就看不到其他軟體選項
 *   區段 3【職務】：學年 → 區段 4；科任 → 區段 5；行政 → 區段 6
 *   區段 4【學年】任教學年、區段 5【科任】任教科目 → 填完都到 區段 7
 *   區段 6【行政】所屬處室 → 填完到 區段 8
 *   區段 7 申請軟體授權：下拉選單限選一項 (含 Google AI Pro、不申請其他軟體)
 *     選 Google AI Pro → 區段 9；其他選項 → 送出
 *   區段 8 申請軟體授權（行政）：下拉選單限選一項 (不含 Google AI Pro，由管理者另行分配) → 送出
 *   區段 9【Google AI Pro 學校帳號】：學校帳號 (必填，限 @SCHOOL_DOMAIN)
 */

const FORM_TITLE = '教學軟體授權申請表';
const TEACHERSAY_NOTICE = '本次軟體使用申請需協助提交使用成果，同時配合取得並使用TeacherSay。';
const FORM_DESCRIPTION =
  '每次填表最多申請 2 項軟體：已預設申請 TeacherSay 專業版，只能再選填 1 項（也可以不選）。\n' +
  TEACHERSAY_NOTICE + '\n' +
  '系統依登記先後排序，名額內為正取，額滿後登記為備取。\n' +
  RESUBMIT_RULE_TEXT + '\n' +
  '登記不代表已取得授權：資訊組審核後，會寄信到您填寫的 Google 帳號，通知授權序號或開通方式。\n' +
  'Google AI Pro 教育版限' + LIMIT_PER_GROUP_TEXT + '可申請（算在另選的 1 項內），並須填寫學校帳號（@' + SCHOOL_DOMAIN + '）；' +
  '行政人員的 Google AI Pro 由資訊組另行分配。\n' +
  '平板借用：「申請項目」選「平板借用」，登記借用的平板車與數量（不需勾選 TeacherSay）；' +
  '登記後才能在「數位學習精進方案成果填報」填報平板借用。教學軟體與平板借用都要申請時，請各填一次。';
const EMAIL_HELP = '多數軟體以 Google 帳號單一登入，請填寫您實際用來登入軟體的 Google 帳號。授權通知也會寄到這個信箱。' +
  '填報成果（含平板借用）時也請使用這個帳號。';
const PAGE_TABLET = '【平板借用】';
const PAGE_TABLET_COUNT = '【平板借用數量】';

// 區段標題 (setupForm 以標題辨識既有表單上的區段)
const PAGE_TEACHERSAY = '【TeacherSay 專業版】';
const PAGE_SOFTWARE = '申請軟體授權';
const PAGE_ROLE = '【職務】';
const PAGE_GRADE = '【學年】';
const PAGE_SUBJECT = '【科任】';
const PAGE_OFFICE = '【行政】';
const PAGE_SOFTWARE_OFFICE = '申請軟體授權（行政）';
const PAGE_WORKSPACE = '【Google AI Pro 學校帳號】';
const PAGE_AI_PRO = '【Google AI Pro 教育版】'; // 舊 (v3.2) 的是／否題區段，更新表單時刪除
const OFFICE_AI_NOTICE = '行政人員的 Google AI Pro 教育版授權由資訊組另行分配，本表單不提供申請。';

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

  // 平板車清單：不存在時以預設內容建立
  if (!ss.getSheetByName(TABLET_SHEET_NAME)) created.push(TABLET_SHEET_NAME);
  getTabletCarts_(ss);

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
      changes = syncEnrollmentForm_(form, softwareList, getTabletCarts_(ss));
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
    syncEnrollmentForm_(form, softwareList, getTabletCarts_(ss));
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    props.setProperty(PROP.FORM_ID, form.getId());
    props.deleteProperty(PROP.RESPONSE_SHEET_ID);

    const responseSheet = findNewResponseSheet_(ss, sheetIdsBefore);
    if (responseSheet) props.setProperty(PROP.RESPONSE_SHEET_ID, String(responseSheet.getSheetId()));
    message = '已建立新的申請表單，回應會寫入' +
      (responseSheet ? '新分頁「' + responseSheet.getName() + '」。' : '本試算表的新分頁。') +
      '\n舊表單與舊回應分頁沒有變動，它們的回應不會被登記；不再使用時請自行關閉舊表單。' +
      (archivedName ? '\n原有的登記資料已封存為分頁「' + archivedName + '」。' : '') +
      '\n' + openToAnyoneWithLink_(form);
  }

  const regSheet = getRegSheet_(ss);
  prepareRegSheet_(regSheet);
  const rows = readRegistrations_(regSheet);
  const stats = refreshStats_(ss, softwareList, rows);
  applyRemainingHelp_(form, stats);

  const prefillUrl = buildPrefilledUrl_(form);
  props.setProperty(PROP.PREFILL_URL, prefillUrl);

  notify_(message +
    '\n\n請發給老師這個網址（' + REQUIRED_SOFTWARE + ' 已預先勾選）：\n' + prefillUrl +
    '\n\n一般填寫網址：' + form.getPublishedUrl() +
    '\n編輯網址：' + form.getEditUrl() +
    '\n\n之後可從選單「授權管理 → 顯示申請表單網址」再次取得。下一步：執行「3. 建立表單提交觸發」。');
}

/**
 * 建立基本資料、職務與學年／科任／行政三個區段的題目。題目順序即回應分頁的欄位順序。
 * 區段的排列、TeacherSay、申請軟體授權、Google AI Pro 與所有跳轉由 syncEnrollmentForm_() 處理。
 */
function buildEnrollmentForm_(form) {
  const emailValidation = FormApp.createTextValidation()
    .setHelpText('請輸入有效的 Email 地址').requireTextIsEmail().build();

  form.addTextItem().setTitle(Q_NAME).setRequired(true);
  form.addTextItem().setTitle(Q_EMAIL)
    .setHelpText(EMAIL_HELP)
    .setValidation(emailValidation)
    .setRequired(true);
  form.addMultipleChoiceItem().setTitle(Q_ROLE).setChoiceValues([ROLE.GRADE, ROLE.SUBJECT, ROLE.OFFICE]).setRequired(true);
  form.addPageBreakItem().setTitle(PAGE_GRADE);
  form.addMultipleChoiceItem().setTitle(Q_GRADE).setChoiceValues(GRADE_CHOICES).setRequired(true);
  form.addPageBreakItem().setTitle(PAGE_SUBJECT);
  form.addTextItem().setTitle(Q_SUBJECT).setHelpText('例：自然、社會、英語').setRequired(true);
  form.addPageBreakItem().setTitle(PAGE_OFFICE);
  form.addMultipleChoiceItem().setTitle(Q_OFFICE).setChoiceValues(OFFICE_CHOICES).setRequired(true);
}

/**
 * 讓表單的說明、區段順序、題目與跳轉符合目前版本 (見檔案開頭的表單結構)。
 * 新表單與 v3.1／v3.2 建立的表單都適用；可重複執行。缺少的題目或區段會新增，所有已知題目都移到固定位置，
 * 舊版的軟體核取方塊、Google AI Pro 是／否題與其區段會刪除 (回應分頁的舊欄位與資料保留)。
 * 既有題目以標題辨識、就地修改，所以回應分頁的既有欄位與資料不受影響。
 * @returns {string[]} 變更說明
 */
function syncEnrollmentForm_(form, softwareList, carts) {
  const T = FormApp.ItemType;
  const NAV = FormApp.PageNavigationType;
  const changes = [];
  const find = function (title, type, label) {
    const item = findItemByTitle_(form, title);
    if (item && item.getType() !== type) {
      throw new Error('「' + title + '」題型不符 (' + label + ')');
    }
    return item;
  };
  const need = function (title, type, label) {
    const item = find(title, type, label);
    if (!item) throw new Error('找不到「' + title + '」');
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
  const newPage = function () { return form.addPageBreakItem(); };
  const newList = function () { return form.addListItem(); };

  form.setDescription(FORM_DESCRIPTION);

  // ---------- 取得或建立所有題目與區段 ----------
  const email = need(Q_EMAIL, T.TEXT, '簡答');
  const kind = ensure(Q_KIND, T.MULTIPLE_CHOICE, '單選', function () { return form.addMultipleChoiceItem(); });
  const pageTeacherSay = ensure(PAGE_TEACHERSAY, T.PAGE_BREAK, '區段', newPage);
  const teacherSay = ensure(Q_TEACHERSAY, T.CHECKBOX, '核取方塊', function () { return form.addCheckboxItem(); });
  const pageRole = ensure(PAGE_ROLE, T.PAGE_BREAK, '區段', newPage);
  const role = need(Q_ROLE, T.MULTIPLE_CHOICE, '單選');
  const pageGrade = need(PAGE_GRADE, T.PAGE_BREAK, '區段');
  const grade = need(Q_GRADE, T.MULTIPLE_CHOICE, '單選');
  const pageSubject = need(PAGE_SUBJECT, T.PAGE_BREAK, '區段');
  const subject = need(Q_SUBJECT, T.TEXT, '簡答');
  const pageOffice = need(PAGE_OFFICE, T.PAGE_BREAK, '區段');
  const office = need(Q_OFFICE, T.MULTIPLE_CHOICE, '單選');
  const pageSoftware = ensure(PAGE_SOFTWARE, T.PAGE_BREAK, '區段', newPage);
  const pick = ensure(Q_PICK, T.LIST, '下拉選單', newList);
  const pageSoftwareOffice = ensure(PAGE_SOFTWARE_OFFICE, T.PAGE_BREAK, '區段', newPage);
  const pickOffice = ensure(Q_PICK_OFFICE, T.LIST, '下拉選單', newList);
  const pageWorkspace = ensure(PAGE_WORKSPACE, T.PAGE_BREAK, '區段', newPage);
  const workspace = ensure(Q_WORKSPACE, T.TEXT, '簡答', function () { return form.addTextItem(); });
  const pageTablet = ensure(PAGE_TABLET, T.PAGE_BREAK, '區段', newPage);
  const tabletUnit = ensure(Q_TABLET_UNIT, T.TEXT, '簡答', function () { return form.addTextItem(); });
  const tabletCart = ensure(Q_TABLET_CART, T.LIST, '下拉選單', newList);
  const tabletMode = ensure(Q_TABLET_MODE, T.MULTIPLE_CHOICE, '單選', function () { return form.addMultipleChoiceItem(); });
  const pageTabletCount = ensure(PAGE_TABLET_COUNT, T.PAGE_BREAK, '區段', newPage);
  const tabletCount = ensure(Q_TABLET_COUNT, T.TEXT, '簡答', function () { return form.addTextItem(); });

  // ---------- 排列順序：從「Google 帳號」之後依序排好 ----------
  [kind, pageTeacherSay, teacherSay, pageRole, role, pageGrade, grade, pageSubject, subject, pageOffice, office,
    pageSoftware, pick, pageSoftwareOffice, pickOffice, pageWorkspace, workspace,
    pageTablet, tabletUnit, tabletCart, tabletMode, pageTabletCount, tabletCount].reduce(function (anchor, item) {
    placeAfter_(form, item, anchor);
    return item;
  }, email);

  // ---------- 區段 1：Google 帳號說明、申請項目 (依答案前往教學軟體或平板借用) ----------
  email.asTextItem().setHelpText(EMAIL_HELP);
  const kindItem = kind.asMultipleChoiceItem();
  kindItem.setChoices([
    kindItem.createChoice(KIND.SOFTWARE, pageTeacherSay.asPageBreakItem()),
    kindItem.createChoice(KIND.TABLET, pageTablet.asPageBreakItem())
  ]).setHelpText('教學軟體與平板借用都要申請時，請各填一次表單；兩種登記互不影響。').setRequired(true);

  // ---------- 平板借用：職務與班級、平板車、整車或填數量 ----------
  pageTablet.asPageBreakItem().setHelpText('登記後即可在「數位學習精進方案成果填報」填報平板借用，填報時請使用同一個 Google 帳號。' +
    TABLET_RESUBMIT_TEXT);
  tabletUnit.asTextItem().setHelpText('例：三年二班導師、英語科任、教務處').setRequired(true);
  tabletCart.asListItem().setRequired(true);
  applyTabletCarts_(form, carts || DEFAULT_TABLET_CARTS);
  const modeItem = tabletMode.asMultipleChoiceItem();
  modeItem.setChoices([
    modeItem.createChoice(TABLET_MODE.WHOLE, NAV.SUBMIT),
    modeItem.createChoice(TABLET_MODE.PARTIAL, pageTabletCount.asPageBreakItem())
  ]).setHelpText('整車借用以該車台數登記；部分借用請在下一頁填寫數量。').setRequired(true);
  tabletCount.asTextItem()
    .setHelpText('請填 1 到該平板車的台數；超過該車台數時以整車計。')
    .setRequired(true);

  // ---------- 區段 2【TeacherSay】：必填、只有一個選項；沒勾就不能前往下一頁 ----------
  pageTeacherSay.asPageBreakItem().setHelpText(TEACHERSAY_NOTICE);
  teacherSay.asCheckboxItem()
    .setChoiceValues([REQUIRED_SOFTWARE])
    .setHelpText(TEACHERSAY_NOTICE + '\n請勾選後再按「下一頁」，未勾選無法選擇其他軟體，也無法送出表單。')
    .setRequired(true);

  // ---------- 區段 3【職務】與 4～6 ----------
  const roleItem = role.asMultipleChoiceItem();
  roleItem.setChoices([
    roleItem.createChoice(ROLE.GRADE, pageGrade.asPageBreakItem()),
    roleItem.createChoice(ROLE.SUBJECT, pageSubject.asPageBreakItem()),
    roleItem.createChoice(ROLE.OFFICE, pageOffice.asPageBreakItem())
  ]).setRequired(true);
  const officeItem = office.asMultipleChoiceItem();
  if (officeItem.getChoices().length !== OFFICE_CHOICES.length) changes.push('・「' + Q_OFFICE + '」選項更新為：' + OFFICE_CHOICES.join('、'));
  officeItem.setChoiceValues(OFFICE_CHOICES);
  pageOffice.asPageBreakItem().setHelpText(OFFICE_AI_NOTICE);

  // ---------- 區段 7、8：另選一項 (下拉選單，必填，可選「不申請其他軟體」) ----------
  // TeacherSay 不在選單內；行政的選單不含 Google AI Pro。選 Google AI Pro → 學校帳號頁，其餘 → 送出。
  const others = softwareList
    .map(function (sw) { return sw.name; })
    .filter(function (name) { return name !== REQUIRED_SOFTWARE; });
  const pickItem = pick.asListItem();
  pickItem.setChoices(others.map(function (name) {
    return pickItem.createChoice(name, name === WORKSPACE_ACCOUNT_SOFTWARE ? pageWorkspace.asPageBreakItem() : NAV.SUBMIT);
  }).concat([pickItem.createChoice(NONE_CHOICE, NAV.SUBMIT)])).setRequired(true);
  const pickOfficeItem = pickOffice.asListItem();
  pickOfficeItem.setChoices(others
    .filter(function (name) { return name !== WORKSPACE_ACCOUNT_SOFTWARE; })
    .concat([NONE_CHOICE])
    .map(function (name) { return pickOfficeItem.createChoice(name, NAV.SUBMIT); })).setRequired(true);
  pageSoftwareOffice.asPageBreakItem().setHelpText(OFFICE_AI_NOTICE);

  // ---------- 區段 9 學校帳號：必填且限學校網域 ----------
  pageWorkspace.asPageBreakItem().setHelpText('Google AI Pro 教育版會開通在學校帳號。');
  workspace.asTextItem()
    .setHelpText('請填寫學校帳號（@' + SCHOOL_DOMAIN + '）。授權會開通在這個帳號，通知信也會副本到這裡。')
    .setValidation(FormApp.createTextValidation()
      .setHelpText('請填寫學校帳號（@' + SCHOOL_DOMAIN + '）')
      .requireTextMatchesPattern(schoolAccountPattern_())
      .build())
    .setRequired(true);

  // ---------- 區段之間的跳轉 ----------
  // 分頁符號的 setGoToPage：依序填完「這個分頁符號之前的區段」後要前往哪裡。
  // 區段裡有依答案跳轉的題目時 (職務、學年／科任的下拉選單)，以題目為準。
  pageTeacherSay.asPageBreakItem().setGoToPage(NAV.CONTINUE);                     // (基本資料的申請項目依答案跳轉)
  pageRole.asPageBreakItem().setGoToPage(NAV.CONTINUE);                           // TeacherSay → 職務
  pageGrade.asPageBreakItem().setGoToPage(NAV.CONTINUE);                          // (職務依答案跳轉)
  pageSubject.asPageBreakItem().setGoToPage(pageSoftware.asPageBreakItem());      // 學年填完 → 申請軟體授權
  pageOffice.asPageBreakItem().setGoToPage(pageSoftware.asPageBreakItem());       // 科任填完 → 申請軟體授權
  pageSoftware.asPageBreakItem().setGoToPage(pageSoftwareOffice.asPageBreakItem()); // 行政填完 → 申請軟體授權（行政）
  pageSoftwareOffice.asPageBreakItem().setGoToPage(NAV.CONTINUE);                 // (學年／科任的選單依答案跳轉)
  pageWorkspace.asPageBreakItem().setGoToPage(NAV.SUBMIT);                        // 行政的選單填完 → 送出
  pageTablet.asPageBreakItem().setGoToPage(NAV.SUBMIT);                           // 學校帳號填完 → 送出
  pageTabletCount.asPageBreakItem().setGoToPage(NAV.CONTINUE);                    // (借用方式依答案跳轉)

  // ---------- 刪除舊版題目 (v3.1／v3.2)；先刪題目再刪區段 ----------
  [Q_SOFTWARE, Q_AI_PRO, PAGE_AI_PRO].forEach(function (title) {
    let item = findItemByTitle_(form, title);
    while (item) {
      form.deleteItem(item.getIndex());
      changes.push('・刪除舊題「' + title + '」（回應試算表的舊欄位與資料保留）');
      item = findItemByTitle_(form, title);
    }
  });

  return changes;
}

/** 兩份下拉選單下方的剩餘名額 (各自只列選單裡有的軟體)。 */
function applyRemainingHelp_(form, stats) {
  const now = nowText_();
  const pick = findItemByTitle_(form, Q_PICK);
  if (pick) pick.setHelpText(buildRemainingText_(stats, now, [REQUIRED_SOFTWARE]));
  const pickOffice = findItemByTitle_(form, Q_PICK_OFFICE);
  if (pickOffice) pickOffice.setHelpText(buildRemainingText_(stats, now, [REQUIRED_SOFTWARE, WORKSPACE_ACCOUNT_SOFTWARE]));
  return !!(pick || pickOffice);
}

/**
 * 新表單：發布並把作答者設為「知道連結的任何人」(不必登入就能填寫)。
 * 學校網域建立的表單預設只開放給網域內的使用者，所以要另外加上 Drive 的「發布檢視」權限
 * (type: anyone, view: published)，需要已啟用的 Drive 進階服務 (v3)。
 * 失敗時 (例如管理員禁止對外共用) 不中斷建立流程，只回傳手動設定的說明。
 * 只在建立新表單時執行；更新既有表單時不會變更作答者權限。
 * @returns {string} 結果說明
 */
function openToAnyoneWithLink_(form) {
  try {
    if (typeof form.supportsAdvancedResponderPermissions === 'function' && form.supportsAdvancedResponderPermissions() &&
      !form.isPublished()) {
      form.setPublished(true);
    }
    Drive.Permissions.create({ type: 'anyone', role: 'reader', view: 'published' }, form.getId());
    return '作答者權限：知道連結的任何人（不必登入即可填寫）。';
  } catch (error) {
    Logger.log('設定作答者權限失敗：' + error);
    return '【注意】無法自動開放作答者權限（' + (error.message || error) + '）。' +
      '請在表單編輯器按「已發布」→ 作答者「管理」，把「作答者檢視權限」改為「知道連結的任何人」。';
  }
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

/**
 * 把剩餘名額寫到表單「軟體」題目下方的說明文字，並依「平板車清單」更新平板車選項。
 * @returns {{remaining: boolean, tablet: {carts: Object[], fromSheet: boolean, updated: boolean, error: string, noForm: boolean}}}
 */
function refreshFormRemaining_(ss, stats) {
  const cartState = readTabletCarts_(ss);
  const tablet = { carts: cartState.carts, fromSheet: cartState.fromSheet, updated: false, error: '', noForm: false };
  const form = getEnrollmentForm_();
  if (!form) {
    tablet.noForm = true;
    return { remaining: false, tablet: tablet };
  }
  try {
    tablet.updated = applyTabletCarts_(form, cartState.carts);
  } catch (error) {
    tablet.error = String(error);
    Logger.log('更新平板車選項失敗：' + error);
  }
  return { remaining: applyRemainingHelp_(form, stats), tablet: tablet };
}

/**
 * 依「平板車清單」更新表單的平板車選項與數量上限；選項沒變時不動表單。
 * @returns {boolean} 是否有更新
 */
function applyTabletCarts_(form, carts) {
  const cartItem = findItemByTitle_(form, Q_TABLET_CART);
  if (!cartItem || carts.length === 0) return false;
  const labels = tabletCartLabels_(carts);
  const list = cartItem.asListItem();
  const current = list.getChoices().map(function (choice) { return choice.getValue(); });
  if (current.join('\n') === labels.join('\n')) return false;
  list.setChoiceValues(labels);
  const countItem = findItemByTitle_(form, Q_TABLET_COUNT);
  if (countItem) {
    const maxCart = Math.max.apply(null, carts.map(function (c) { return c.max; }));
    countItem.asTextItem().setValidation(FormApp.createTextValidation()
      .setHelpText('請填 1～' + maxCart + ' 的數字')
      .requireNumberBetween(1, maxCart)
      .build());
  }
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
