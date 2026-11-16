/**
 * setupForm.gs — 初次設定：分頁、申請表單、表單剩餘名額、觸發器
 *
 * 使用方式 (試算表選單「授權管理 → 初次設定」，依序執行)：
 *   1. 初始化分頁          setupSheets()
 *   2. 建立／更新申請表單   setupForm()
 *   3. 建立表單提交觸發     createFormSubmitTrigger()
 *
 * setupForm() 會沿用已連結到本試算表的表單並就地更新：
 *   - 題目名稱與 code.gs 的 Q_ 設定相同者直接沿用 (選項改為「軟體清單」的內容)。
 *   - 缺少的題目會新增。
 *   - 題型不相容的同名題目 (例如軟體題原本是單選) 會刪除後重建成正確題型；
 *     表單回應試算表中的舊欄位與舊資料仍會保留。
 *   - 表單上其他舊題目不會更動，不需要時請手動刪除。
 * 找不到已連結的表單時，才會建立新表單並連結到本試算表。
 *
 * 修改「軟體清單」的軟體名稱或增減軟體後，請再執行一次 setupForm() 同步表單選項。
 */

const FORM_TITLE = '教學軟體授權申請表';
const FORM_DESCRIPTION =
  '請勾選要申請的教學軟體授權，每人每項限登記 1 個。\n' +
  '系統依登記先後排序，名額內為正取，額滿後登記為備取。\n' +
  '登記不代表已取得授權：資訊組審核後，會另行寄信通知授權序號或開通方式。\n' +
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

  // 登記起算時間：補登只會處理這個時間之後的表單回應，避免把往年的回應算進來
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ENROLL_REG_START')) {
    props.setProperty('ENROLL_REG_START', new Date().toISOString());
  }

  const softwareList = getSoftwareList_(ss);
  const rows = readRegistrations_(regSheet);
  refreshStats_(ss, softwareList, rows);

  notify_('分頁初始化完成。' + (created.length ? '\n新增分頁：' + created.join('、') : '') +
    '\n\n請確認「' + SOFTWARE_SHEET_NAME + '」的可用授權數與分配方式，' +
    '並在「' + LICENSE_SHEET_NAME + '」的 D 欄標示每個序號屬於哪套軟體 (留空視為 ' + DEFAULT_SERIAL_SOFTWARE + ')。');
}

// =====================================================================
// 2. 建立／更新申請表單
// =====================================================================

function setupForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const softwareList = getSoftwareList_(ss);
  const softwareNames = softwareList.map(function (sw) { return sw.name; });

  let form = getLinkedForm_(ss);
  let createdForm = false;
  if (!form) {
    form = FormApp.create(FORM_TITLE);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    createdForm = true;
  }
  form.setDescription(FORM_DESCRIPTION);

  const log = [];
  ensureTextItem_(form, Q_NAME, false, log);
  ensureTextItem_(form, Q_EMAIL, true, log);
  const groupItem = ensureChoiceItem_(form, Q_GROUP, GROUP_CHOICES, false, log);
  const softwareItem = ensureChoiceItem_(form, Q_SOFTWARE, softwareNames, true, log);

  // 身分題排在軟體題之前
  if (groupItem.getIndex() > softwareItem.getIndex()) {
    form.moveItem(groupItem.getIndex(), softwareItem.getIndex());
  }

  const rows = readRegistrations_(getRegSheet_(ss));
  softwareItem.setHelpText(buildRemainingText_(computeStats_(rows, softwareList), nowText_()));

  const message = (createdForm ? '已建立新表單並連結到本試算表。' : '已更新現有表單。') +
    '\n' + log.join('\n') +
    '\n\n填寫網址：' + form.getPublishedUrl() +
    '\n編輯網址：' + form.getEditUrl() +
    '\n\n下一步：執行「3. 建立表單提交觸發」。';
  notify_(message);
}

/** 取得連結到本試算表的表單；沒有則回傳 null。 */
function getLinkedForm_(ss) {
  let url = FORM_URL;
  if (!url) {
    const responseSheet = getResponseSheet_(ss);
    url = responseSheet ? responseSheet.getFormUrl() : null;
  }
  if (!url) url = ss.getFormUrl();
  return url ? FormApp.openByUrl(url) : null;
}

function findItemByTitle_(form, title) {
  const want = normalize_(title);
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (normalize_(items[i].getTitle()) === want) return items[i];
  }
  return null;
}

/** 刪除題型不相容的舊題目並回傳它原本的位置。 */
function removeItem_(form, item, title, log) {
  const index = item.getIndex();
  form.deleteItem(item);
  log.push('・「' + title + '」題型不相容，已刪除並重建 (回應試算表的舊欄位仍保留)');
  return index;
}

function ensureTextItem_(form, title, emailOnly, log) {
  const type = FormApp.ItemType;
  let item = findItemByTitle_(form, title);
  let position = -1;

  if (item && item.getType() === type.PARAGRAPH_TEXT) {
    item.asParagraphTextItem().setRequired(true);
    log.push('・沿用「' + title + '」');
    return item;
  }
  if (item && item.getType() !== type.TEXT) {
    position = removeItem_(form, item, title, log);
    item = null;
  }

  let textItem;
  if (item) {
    textItem = item.asTextItem();
    log.push('・沿用「' + title + '」');
  } else {
    textItem = form.addTextItem().setTitle(title);
    if (position >= 0) form.moveItem(textItem.getIndex(), position);
    else log.push('・新增「' + title + '」');
  }
  textItem.setRequired(true);
  if (emailOnly) {
    textItem.setValidation(FormApp.createTextValidation()
      .setHelpText('請輸入有效的 Email 地址')
      .requireTextIsEmail()
      .build());
  }
  return textItem;
}

/**
 * 確保選擇題存在且選項正確。
 * multiple = true 必須是核取方塊 (可複選)；false 接受單選或下拉。
 */
function ensureChoiceItem_(form, title, choices, multiple, log) {
  const type = FormApp.ItemType;
  let item = findItemByTitle_(form, title);
  let position = -1;

  if (item) {
    const itemType = item.getType();
    if (multiple && itemType === type.CHECKBOX) {
      item.asCheckboxItem().setChoiceValues(choices).setRequired(true);
      log.push('・沿用「' + title + '」並更新選項');
      return item;
    }
    if (!multiple && itemType === type.MULTIPLE_CHOICE) {
      item.asMultipleChoiceItem().setChoiceValues(choices).setRequired(true);
      log.push('・沿用「' + title + '」並更新選項');
      return item;
    }
    if (!multiple && itemType === type.LIST) {
      item.asListItem().setChoiceValues(choices).setRequired(true);
      log.push('・沿用「' + title + '」並更新選項');
      return item;
    }
    position = removeItem_(form, item, title, log);
  }

  const created = multiple
    ? form.addCheckboxItem().setTitle(title).setChoiceValues(choices).setRequired(true)
    : form.addMultipleChoiceItem().setTitle(title).setChoiceValues(choices).setRequired(true);
  if (position >= 0) form.moveItem(created.getIndex(), position);
  else log.push('・新增「' + title + '」');
  return created;
}

/** 把剩餘名額寫到表單「軟體」題目下方的說明文字。 */
function refreshFormRemaining_(ss, stats) {
  const form = getLinkedForm_(ss);
  if (!form) return false;
  const item = findItemByTitle_(form, Q_SOFTWARE);
  if (!item) return false;
  item.setHelpText(buildRemainingText_(stats, nowText_()));
  return true;
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
