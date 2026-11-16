/**
 * setupForm.gs — 初次設定：分頁、申請表單、表單剩餘名額、觸發器
 *
 * 使用方式 (試算表選單「授權管理 → 初次設定」，依序執行)：
 *   1. 初始化分頁                          setupSheets()
 *   2. 建立新申請表單（已建立則更新選項）   setupForm()
 *   3. 建立表單提交觸發                    createFormSubmitTrigger()
 *
 * setupForm() 第一次執行會建立一份「全新」的表單，回應寫入本試算表的新分頁。
 * 舊表單與舊回應分頁完全不動；程式只處理這份新表單的回應，所以不會和舊資料重複。
 * 再次執行 setupForm() 只會同步軟體選項與剩餘名額，不會再建立新表單。
 * 要重新建立一份表單時，先執行 resetEnrollmentForm() 解除綁定。
 *
 * 表單結構與跳題：
 *   區段 1 基本資料：姓名、開通授權的 Google 帳號、職務
 *     ├ 學年 → 區段 2【學年】任教學年 (一～六年級)
 *     ├ 科任 → 區段 3【科任】任教科目 (簡答)
 *     └ 行政 → 區段 4【行政】所屬處室
 *   區段 2、3、4 填完都到 區段 5 申請軟體授權：軟體 (複選)、Google AI Pro 的學校帳號 (選填)
 */

const FORM_TITLE = '教學軟體授權申請表';
const FORM_DESCRIPTION =
  '請勾選要申請的教學軟體授權，每人每項限登記 1 個。\n' +
  '系統依登記先後排序，名額內為正取，額滿後登記為備取。\n' +
  '登記不代表已取得授權：資訊組審核後，會寄信到您填寫的 Google 帳號，通知授權序號或開通方式。\n' +
  'Google AI Pro 教育版限' + LIMIT_PER_GROUP_TEXT + '可申請。';

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

  // 申請登記
  const hadRegSheet = !!ss.getSheetByName(REG_SHEET_NAME);
  const regSheet = getRegSheet_(ss);
  if (!hadRegSheet) created.push(REG_SHEET_NAME);
  regSheet.getRange(2, COL.STATUS, regSheet.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([STATUS.OK, STATUS.WAIT, STATUS.INELIGIBLE, STATUS.CANCEL, STATUS.ISSUED], true)
      .setAllowInvalid(true)
      .build());

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
    '\n\n請確認「' + SOFTWARE_SHEET_NAME + '」的可用授權數與分配方式。' +
    '分配方式為「序號」的軟體，請在「' + LICENSE_SHEET_NAME + '」貼上序號並於 D 欄標示軟體 (留空視為 ' +
    DEFAULT_SERIAL_SOFTWARE + ')。');
}

// =====================================================================
// 2. 建立新申請表單（已建立則更新選項）
// =====================================================================

function setupForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const softwareList = getSoftwareList_(ss);
  const softwareNames = softwareList.map(function (sw) { return sw.name; });
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
    // 已建立：只同步軟體選項
    const softwareItem = findItemByTitle_(form, Q_SOFTWARE);
    if (!softwareItem || softwareItem.getType() !== FormApp.ItemType.CHECKBOX) {
      notify_('申請表單上找不到核取方塊題「' + Q_SOFTWARE + '」，無法同步選項。請還原題目名稱，' +
        '或執行 resetEnrollmentForm() 後重新建立表單。');
      return;
    }
    softwareItem.asCheckboxItem().setChoiceValues(softwareNames);
    message = '已更新申請表單的軟體選項。';
  } else {
    // 第一次：建立全新的表單，回應寫入本試算表的新分頁
    const sheetIdsBefore = ss.getSheets().map(function (sheet) { return sheet.getSheetId(); });
    form = FormApp.create(FORM_TITLE);
    buildEnrollmentForm_(form, softwareNames);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    props.setProperty(PROP.FORM_ID, form.getId());
    props.deleteProperty(PROP.RESPONSE_SHEET_ID);

    const responseSheet = findNewResponseSheet_(ss, sheetIdsBefore);
    if (responseSheet) props.setProperty(PROP.RESPONSE_SHEET_ID, String(responseSheet.getSheetId()));
    message = '已建立新的申請表單，回應會寫入' +
      (responseSheet ? '新分頁「' + responseSheet.getName() + '」。' : '本試算表的新分頁。') +
      '\n舊表單與舊回應分頁沒有變動，它們的回應不會被登記；不再使用時請自行關閉舊表單。';
  }

  const rows = readRegistrations_(getRegSheet_(ss));
  findItemByTitle_(form, Q_SOFTWARE).setHelpText(buildRemainingText_(computeStats_(rows, softwareList), nowText_()));

  notify_(message +
    '\n\n填寫網址：' + form.getPublishedUrl() +
    '\n編輯網址：' + form.getEditUrl() +
    '\n\n下一步：執行「3. 建立表單提交觸發」。');
}

/** 建立表單題目與跳題。題目順序即回應分頁的欄位順序。 */
function buildEnrollmentForm_(form, softwareNames) {
  form.setDescription(FORM_DESCRIPTION);
  const emailValidation = function () {
    return FormApp.createTextValidation().setHelpText('請輸入有效的 Email 地址').requireTextIsEmail().build();
  };

  // ---------- 區段 1：基本資料 ----------
  form.addTextItem().setTitle(Q_NAME).setRequired(true);
  form.addTextItem().setTitle(Q_EMAIL)
    .setHelpText('多數軟體以 Google 帳號單一登入，請填寫您實際用來登入軟體的 Google 帳號。授權通知也會寄到這個信箱。')
    .setValidation(emailValidation())
    .setRequired(true);
  const roleItem = form.addMultipleChoiceItem().setTitle(Q_ROLE).setRequired(true);

  // ---------- 區段 2：學年 ----------
  const pageGrade = form.addPageBreakItem().setTitle('【學年】');
  form.addMultipleChoiceItem().setTitle(Q_GRADE).setChoiceValues(GRADE_CHOICES).setRequired(true);

  // ---------- 區段 3：科任 ----------
  const pageSubject = form.addPageBreakItem().setTitle('【科任】');
  form.addTextItem().setTitle(Q_SUBJECT).setHelpText('例：自然、社會、英語').setRequired(true);

  // ---------- 區段 4：行政 ----------
  const pageOffice = form.addPageBreakItem().setTitle('【行政】');
  form.addMultipleChoiceItem().setTitle(Q_OFFICE).setChoiceValues(OFFICE_CHOICES).setRequired(true);

  // ---------- 區段 5：申請軟體授權 ----------
  const pageSoftware = form.addPageBreakItem().setTitle('申請軟體授權');
  form.addCheckboxItem().setTitle(Q_SOFTWARE).setChoiceValues(softwareNames).setRequired(true);
  form.addTextItem().setTitle(Q_WORKSPACE)
    .setHelpText('有勾選 Google AI Pro 教育版者請填寫學校帳號（@' + SCHOOL_DOMAIN + '），授權會開通在這個帳號，' +
      '通知信也會副本到這裡。未申請 Google AI Pro 可留空。')
    .setValidation(emailValidation())
    .setRequired(false);

  // ---------- 跳題 ----------
  roleItem.setChoices([
    roleItem.createChoice(ROLE.GRADE, pageGrade),
    roleItem.createChoice(ROLE.SUBJECT, pageSubject),
    roleItem.createChoice(ROLE.OFFICE, pageOffice)
  ]);
  // 分頁符號的 setGoToPage：填完「前一個區段」後要前往的區段
  pageSubject.setGoToPage(pageSoftware); // 學年區段填完 → 申請軟體授權
  pageOffice.setGoToPage(pageSoftware);  // 科任區段填完 → 申請軟體授權
  // 行政區段填完會依序進入下一個區段，也就是申請軟體授權
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
