/**
 * allocate.gs — 管理者操作：統計、勾選核准、發放授權、補登
 *
 * 發放流程：
 *   1. 表單送出後，「申請登記」分頁依登記先後標示 正取／備取 (不發序號)。
 *   2. 管理者在「核准」欄勾選要發放的名單 (可用選單一次勾選所有正取，再手動增減)。
 *   3. 選單「預覽發放」確認名單 → 「發放已核准授權」從「序號列表」取號並寄信通知。
 *
 * 要取消某筆登記：把「狀態」欄改成「取消」，再執行「重新排序正取／備取」，備取會依序遞補。
 */

// 帳號開通類的軟體，發放時把姓名與帳號同步寫入對應分頁 (分頁不存在則略過)
const ACCOUNT_SYNC_SHEETS = { 'Wordwall 專業版': 'wordwall授權確認' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('授權管理')
    .addItem('顯示申請表單網址', 'menuShowFormUrls')
    .addItem('重新統計並更新表單剩餘名額', 'menuRefreshAll')
    .addItem('重新排序正取／備取', 'menuRerank')
    .addSeparator()
    .addItem('勾選所有正取為核准', 'menuApproveAllAccepted')
    .addItem('預覽發放（不寄信）', 'menuPreviewIssue')
    .addItem('發放已核准授權（寄信通知）', 'menuIssueApproved')
    .addSeparator()
    .addItem('從表單回應補登（漏登時使用）', 'menuBackfill')
    .addItem('補寄登記回條（輸入 Google 帳號）', 'menuResendReceipts')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('初次設定')
      .addItem('1. 初始化分頁', 'setupSheets')
      .addItem('2. 建立新申請表單（已建立則更新）', 'setupForm')
      .addItem('3. 建立表單提交觸發', 'createFormSubmitTrigger'))
    .addToUi();
}

// =====================================================================
// 選單動作
// =====================================================================

function menuShowFormUrls() {
  const form = getEnrollmentForm_();
  if (!form) {
    notify_('尚未建立申請表單，請先執行「初次設定 → 2. 建立新申請表單」。');
    return;
  }
  const prefill = PropertiesService.getScriptProperties().getProperty(PROP.PREFILL_URL);
  notify_('請發給老師這個網址（' + REQUIRED_SOFTWARE + ' 已預先勾選）：\n' + (prefill || '(尚未產生，請再執行一次「2. 建立新申請表單」)') +
    '\n\n一般填寫網址：' + form.getPublishedUrl() +
    '\n編輯網址：' + form.getEditUrl());
}

function menuRefreshAll() {
  const state = recalcAll_(SpreadsheetApp.getActiveSpreadsheet());
  notify_('已重新統計。\n\n' + buildRemainingText_(state.stats));
}

function menuRerank() {
  const state = recalcAll_(SpreadsheetApp.getActiveSpreadsheet());
  notify_('已依登記先後重新排序正取／備取。\n\n' + buildRemainingText_(state.stats));
}

function menuApproveAllAccepted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRegSheet_(ss);
  const state = recalcAll_(ss);
  const targets = state.rows.filter(function (row) { return row.status === STATUS.OK && !row.approved; });
  if (targets.length === 0) {
    notify_('沒有尚未勾選的正取名單。');
    return;
  }
  if (!confirm_('勾選正取名單', '將勾選 ' + targets.length + ' 筆正取為「核准」，此動作不會寄信。要繼續嗎？')) return;
  targets.forEach(function (row) {
    sheet.getRange(row.rowIndex, COL.APPROVED).setValue(true);
  });
  recalcAll_(ss);
  notify_('已勾選 ' + targets.length + ' 筆。確認名單後請執行「預覽發放」。');
}

function menuPreviewIssue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plan = buildIssuePlan_(ss);
  notify_('【預覽，尚未發放】\n\n' + describePlan_(plan));
}

function menuIssueApproved() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preview = buildIssuePlan_(ss);
  if (preview.issue.length === 0) {
    notify_('沒有可發放的名單。\n\n' + describePlan_(preview));
    return;
  }
  const recipients = recipientCount_(preview.issue);
  const mailQuota = MailApp.getRemainingDailyQuota();
  if (recipients.addresses > mailQuota) {
    notify_('今日剩餘寄信額度為 ' + mailQuota + ' 個收件人，不足以通知 ' + recipients.people +
      ' 位申請人（含副本共 ' + recipients.addresses + ' 個收件人）。請明天再發放，或先取消部分核准勾選。');
    return;
  }
  if (!confirm_('發放授權', describePlan_(preview) + '\n\n確定發放並寄信通知 ' + recipients.people + ' 位申請人嗎？')) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let summary;
  try {
    // 取得鎖定後重新讀取，避免確認對話框期間有新的登記或異動；只發放剛才確認過的那些列
    const confirmed = {};
    preview.issue.forEach(function (item) { confirmed[issueKey_(item.row)] = true; });
    const plan = buildIssuePlan_(ss);
    plan.issue = plan.issue.filter(function (item) { return confirmed[issueKey_(item.row)]; });
    summary = executeIssue_(ss, plan);
  } finally {
    lock.releaseLock();
  }
  recalcAll_(ss);
  notify_(summary);
}

function menuBackfill() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!confirm_('從表單回應補登', '將把表單回應中尚未登記的申請補進「申請登記」(不會寄送回條)。要繼續嗎？')) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  try {
    result = backfillFromResponses_(ss);
  } catch (error) {
    notify_(String(error.message || error));
    return;
  } finally {
    lock.releaseLock();
  }
  recalcAll_(ss);
  notify_('補登完成，新增 ' + result.added + ' 筆登記' +
    (result.canceled ? '；依重新填寫（以最後一次為準）取消 ' + result.canceled + ' 筆舊登記' : '') + '。');
}

/**
 * 補寄登記回條 (例如補登後，或申請人沒收到)：輸入 Google 帳號，依目前「申請登記」的有效登記，
 * 寄出與送出表單時相同格式的「已登記」回條，有學校帳號時副本給學校帳號。寄出前會列出收件人確認。
 */
function menuResendReceipts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let input;
  try {
    const ui = SpreadsheetApp.getUi();
    const answer = ui.prompt('補寄登記回條', '請輸入要補寄的 Google 帳號（多個以逗號、空白或換行分隔）：', ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() !== ui.Button.OK) return;
    input = answer.getResponseText();
  } catch (error) {
    Logger.log('此功能需要從試算表選單「授權管理」執行。');
    return;
  }
  const emails = String(input || '').split(/[\s,，;；]+/).filter(function (text) { return text; });
  const plan = buildResendPlan_(readRegistrations_(getRegSheet_(ss)), emails, getSoftwareList_(ss));
  const missingText = plan.missing.length ? '\n\n「申請登記」找不到有效登記，不寄：' + plan.missing.join('、') : '';
  if (plan.send.length === 0) {
    notify_('沒有可補寄的回條。' + missingText);
    return;
  }
  const lines = plan.send.map(function (item) {
    const cc = ccFor_(item.sub.email, item.sub.workspace);
    return '・' + item.sub.name + '（' + item.sub.email + (cc ? '，副本 ' + cc : '') + '）：' +
      item.outcomes.map(function (outcome) { return outcome.software; }).join('、');
  });
  const addresses = plan.send.reduce(function (sum, item) {
    return sum + 1 + (ccFor_(item.sub.email, item.sub.workspace) ? 1 : 0);
  }, 0);
  const quota = MailApp.getRemainingDailyQuota();
  if (addresses > quota) {
    notify_('今日剩餘寄信額度為 ' + quota + ' 個收件人，不足以補寄（共 ' + addresses + ' 個收件人）。請明天再寄。');
    return;
  }
  if (!confirm_('補寄登記回條', '將寄出 ' + plan.send.length + ' 封「已登記」回條：\n' + lines.join('\n') + missingText + '\n\n要寄出嗎？')) return;

  const sent = [];
  const failed = [];
  plan.send.forEach(function (item) {
    try {
      MailApp.sendEmail(withCc_({
        to: item.sub.email,
        subject: '【已登記】教學軟體授權申請',
        htmlBody: buildReceiptHtml_(item.sub, item.outcomes,
          { resentFor: item.since ? Utilities.formatDate(item.since, 'Asia/Taipei', 'MM/dd HH:mm') : '' })
      }, item.sub.email, item.sub.workspace));
      sent.push(item.sub.name);
    } catch (error) {
      failed.push(item.sub.name + '（' + error + '）');
    }
  });
  notify_('已補寄 ' + sent.length + ' 封回條' + (sent.length ? '：' + sent.join('、') : '') + '。' +
    (failed.length ? '\n寄送失敗：' + failed.join('、') : '') + missingText);
}

// =====================================================================
// 發放
// =====================================================================

/** 重新排序、重寫統計、更新表單剩餘名額。 */
function recalcAll_(ss) {
  const softwareList = getSoftwareList_(ss);
  const sheet = getRegSheet_(ss);
  const rows = readRegistrations_(sheet);
  applyRanks_(sheet, rows, softwareList);
  const stats = refreshStats_(ss, softwareList, rows);
  try {
    refreshFormRemaining_(ss, stats);
  } catch (error) {
    Logger.log("更新表單剩餘名額失敗：" + error);
  }
  return { rows: rows, stats: stats, softwareList: softwareList };
}

function buildIssuePlan_(ss) {
  const softwareList = getSoftwareList_(ss);
  const rows = readRegistrations_(getRegSheet_(ss));
  const serials = readSerials_(ss.getSheetByName(LICENSE_SHEET_NAME));
  return planIssue_(rows, softwareList, serials);
}

function readSerials_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(function (v, i) {
      return {
        rowIndex: i + 2,
        serial: String(v[0] == null ? '' : v[0]).trim(),
        email: String(v[1] == null ? '' : v[1]).trim(),
        software: String(v[3] == null ? '' : v[3]).trim()
      };
    })
    .filter(function (item) { return item.serial !== ''; });
}

/** 「序號列表」D 欄是否屬於這套軟體；留空視為 DEFAULT_SERIAL_SOFTWARE，也接受簡寫 (例：HiTeach)。 */
function serialMatches_(cell, softwareName) {
  const text = normalize_(cell);
  if (!text) return softwareName === DEFAULT_SERIAL_SOFTWARE;
  return normalize_(softwareName).indexOf(text) >= 0;
}

/**
 * 決定哪些已勾選核准的登記可以發放 (純邏輯)。
 * 處理對象：核准已勾選，且狀態不是「已發放」「取消」。
 * 略過條件：超過可用授權數、不符每組限一位的規則、序號列表沒有可用序號。
 * @returns {{issue: {row: Object, serial: string, serialRow: number}[], skipped: {row: Object, reason: string}[]}}
 */
function planIssue_(rows, softwareList, serials) {
  const byName = {};
  softwareList.forEach(function (sw) { byName[sw.name] = sw; });

  const issuedCount = {};
  const takenGroups = {};
  rows.forEach(function (row) {
    if (row.status !== STATUS.ISSUED) return;
    issuedCount[row.software] = (issuedCount[row.software] || 0) + 1;
    (takenGroups[row.software] = takenGroups[row.software] || {})[row.group] = true;
  });

  const freeSerials = serials.filter(function (item) { return item.email === ''; });
  const usedSerialRows = {};

  const candidates = rows
    .filter(function (row) {
      return row.approved === true && row.status !== STATUS.ISSUED && row.status !== STATUS.CANCEL;
    })
    .sort(function (a, b) { return (timeValue_(a.time) - timeValue_(b.time)) || (a.rowIndex - b.rowIndex); });

  const issue = [];
  const skipped = [];
  candidates.forEach(function (row) {
    const sw = byName[row.software];
    if (!sw) {
      skipped.push({ row: row, reason: '軟體清單沒有這個項目' });
      return;
    }
    if ((issuedCount[sw.name] || 0) >= sw.quota) {
      skipped.push({ row: row, reason: '超過可用授權數 ' + sw.quota });
      return;
    }
    if (sw.limited) {
      if (LIMITED_GROUPS.indexOf(row.group) < 0) {
        skipped.push({ row: row, reason: '身分不符（限' + LIMIT_PER_GROUP_TEXT + '）' });
        return;
      }
      if (takenGroups[sw.name] && takenGroups[sw.name][row.group]) {
        skipped.push({ row: row, reason: row.group + '已有一位取得授權' });
        return;
      }
    }
    let serial = '';
    let serialRow = 0;
    if (sw.mode === MODE.SERIAL) {
      const found = freeSerials.filter(function (item) {
        return !usedSerialRows[item.rowIndex] && serialMatches_(item.software, sw.name);
      })[0];
      if (!found) {
        skipped.push({ row: row, reason: '「' + LICENSE_SHEET_NAME + '」沒有可用序號' });
        return;
      }
      usedSerialRows[found.rowIndex] = true;
      serial = found.serial;
      serialRow = found.rowIndex;
    }
    issuedCount[sw.name] = (issuedCount[sw.name] || 0) + 1;
    (takenGroups[sw.name] = takenGroups[sw.name] || {})[row.group] = true;
    issue.push({ row: row, serial: serial, serialRow: serialRow, mode: sw.mode });
  });
  return { issue: issue, skipped: skipped };
}

function describePlan_(plan) {
  const lines = [];
  const counts = {};
  plan.issue.forEach(function (item) {
    counts[item.row.software] = (counts[item.row.software] || 0) + 1;
  });
  lines.push('可發放 ' + plan.issue.length + ' 筆：');
  Object.keys(counts).forEach(function (name) { lines.push('・' + name + '：' + counts[name] + ' 筆'); });
  if (plan.issue.length === 0) lines.push('・（無）請先在「' + REG_SHEET_NAME + '」的「核准」欄勾選名單');
  if (plan.skipped.length > 0) {
    lines.push('');
    lines.push('無法發放 ' + plan.skipped.length + ' 筆：');
    plan.skipped.forEach(function (item) {
      lines.push('・第 ' + item.row.rowIndex + ' 列 ' + item.row.name + '／' + item.row.software + '：' + item.reason);
    });
  }
  return lines.join('\n');
}

function issueKey_(row) {
  return row.rowIndex + '|' + normalize_(row.email) + '|' + row.software;
}

/** 申請人數與實際收件人數 (副本也佔寄信額度)。 */
function recipientCount_(issueItems) {
  const people = {};
  const addresses = {};
  issueItems.forEach(function (item) {
    if (!isEmail_(item.row.email)) return;
    people[normalize_(item.row.email)] = true;
    addresses[normalize_(item.row.email)] = true;
    const cc = ccFor_(item.row.email, item.row.workspace);
    if (cc) addresses[normalize_(cc)] = true;
  });
  return { people: Object.keys(people).length, addresses: Object.keys(addresses).length };
}

/** 這位申請人留下的學校帳號 (同一人的任何一筆登記有填即可)。 */
function workspaceOf_(items) {
  for (let i = 0; i < items.length; i++) {
    if (isEmail_(items[i].row.workspace)) return items[i].row.workspace;
  }
  return '';
}

/** 寫入序號與狀態，再依申請人合併寄出一封通知信。 */
function executeIssue_(ss, plan) {
  const regSheet = getRegSheet_(ss);
  const serialSheet = ss.getSheetByName(LICENSE_SHEET_NAME);
  const now = new Date();

  plan.issue.forEach(function (item) {
    if (item.serialRow) {
      serialSheet.getRange(item.serialRow, 2, 1, 2).setValues([[item.row.email, now]]);
    }
    regSheet.getRange(item.row.rowIndex, COL.STATUS, 1, 2).setValues([[STATUS.ISSUED, '']]);
    regSheet.getRange(item.row.rowIndex, COL.SERIAL, 1, 2).setValues([[item.serial, now]]);
    regSheet.getRange(item.row.rowIndex, COL.ISSUED_AT).setNumberFormat('yyyy/MM/dd HH:mm:ss');
    syncAccountSheet_(ss, item.row);
  });
  SpreadsheetApp.flush();

  // 同一位申請人的多套軟體合併成一封信
  const byEmail = {};
  plan.issue.forEach(function (item) {
    const key = normalize_(item.row.email);
    (byEmail[key] = byEmail[key] || []).push(item);
  });

  let sent = 0;
  const failed = [];
  Object.keys(byEmail).forEach(function (key) {
    const items = byEmail[key];
    const first = items[0].row;
    try {
      if (!isEmail_(first.email)) throw new Error('Google 帳號格式有誤');
      const workspace = workspaceOf_(items);
      MailApp.sendEmail(withCc_({
        to: first.email,
        subject: "【授權核准】教學軟體授權已發放",
        htmlBody: buildIssueHtml_(first.name, first.email, workspace, items)
      }, first.email, workspace));
      sent++;
    } catch (error) {
      failed.push(first.name + ' (' + first.email + ')：' + error);
      items.forEach(function (item) {
        regSheet.getRange(item.row.rowIndex, COL.NOTE).setValue('通知信寄送失敗，請人工通知');
      });
    }
  });

  const lines = ['已發放 ' + plan.issue.length + ' 筆授權，寄出 ' + sent + ' 封通知信。'];
  if (failed.length > 0) {
    lines.push('');
    lines.push('通知信寄送失敗 (授權已記錄，請人工通知)：');
    failed.forEach(function (text) { lines.push('・' + text); });
  }
  if (plan.skipped.length > 0) {
    lines.push('');
    lines.push('未發放 ' + plan.skipped.length + ' 筆：');
    plan.skipped.forEach(function (item) {
      lines.push('・第 ' + item.row.rowIndex + ' 列 ' + item.row.name + '／' + item.row.software + '：' + item.reason);
    });
  }
  return lines.join('\n');
}

function buildIssueHtml_(name, email, workspace, items) {
  const list = items.map(function (item) {
    const useWorkspace = item.row.software === WORKSPACE_ACCOUNT_SOFTWARE && isEmail_(workspace);
    const detail = item.mode === MODE.SERIAL
      ? '授權序號：<b>' + escapeHtml_(item.serial) + '</b>'
      : (useWorkspace
        ? '授權將開通於您的學校帳號（' + escapeHtml_(workspace) + '），開通後即可使用'
        : '授權將開通於您登記的 Google 帳號（' + escapeHtml_(email) + '），開通後以該帳號登入即可使用');
    return '<li><b>' + escapeHtml_(item.row.software) + '</b>－' + detail + '</li>';
  }).join('');
  return '<p>親愛的 ' + escapeHtml_(name) + ' 老師，您好：</p>' +
    '<p>您登記的下列教學軟體授權已核准：</p>' +
    '<ul>' + list + '</ul>' +
    '<p>序號請妥善保管，並依照軟體指示進行啟用。如有問題請洽資訊組。</p>' +
    '<br><p>' + SCHOOL_SIGNATURE + '</p>' +
    '<p style="color: grey; font-size: 12px;">此為系統自動發送信件，請勿直接回覆。</p>';
}

/** 帳號開通類軟體：把姓名與帳號補進對應分頁 (A 姓名、B 帳號)，已存在則不重複寫入。 */
function syncAccountSheet_(ss, row) {
  const sheetName = ACCOUNT_SYNC_SHEETS[row.software];
  if (!sheetName) return;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === row.name) return;
    }
  }
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[row.name, row.email]]);
}

// =====================================================================
// 補登
// =====================================================================

/**
 * 把授權申請表單的回應分頁中、尚未登記的申請補進「申請登記」。
 * 只讀 setupForm() 所建立表單的回應分頁，舊表單的回應不會被匯入。
 * 同一帳號＋軟體不會重複登記，可重複執行；不寄送回條。
 * 同一帳號重新填表以最後一次為準，依時間順序重跑時結果相同。
 * @returns {{added: number, canceled: number}} 新增的登記筆數、因重新填寫而取消的舊登記筆數
 */
function backfillFromResponses_(ss) {
  const responseSheet = getResponseSheet_(ss);
  if (!responseSheet) {
    throw new Error('找不到授權申請表單的回應分頁，請先執行選單「初次設定 → 2. 建立新申請表單（已建立則更新）」。');
  }
  if (responseSheet.getLastRow() < 2) return { added: 0, canceled: 0 };

  const softwareList = getSoftwareList_(ss);
  const values = responseSheet.getRange(1, 1, responseSheet.getLastRow(), responseSheet.getLastColumn()).getValues();
  const headers = values[0];
  let added = 0;
  let canceled = 0;

  for (let r = 1; r < values.length; r++) {
    const time = values[r][0];
    const namedValues = {};
    headers.forEach(function (header, c) {
      const key = String(header);
      (namedValues[key] = namedValues[key] || []).push(values[r][c]);
    });
    let sub;
    try {
      sub = parseSubmission_(namedValues, softwareList, time);
    } catch (error) {
      Logger.log('第 ' + (r + 1) + ' 列無法解析：' + error);
      continue;
    }
    if (sub.selections.length === 0) continue;
    const result = processSubmission_(ss, sub, softwareList);
    added += result.added.length;
    canceled += result.canceled.length;
  }
  return { added: added, canceled: canceled };
}

// =====================================================================
// 對話框
// =====================================================================

function notify_(message) {
  Logger.log(message);
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    // 從編輯器或觸發器執行時沒有 UI，只寫入執行紀錄
  }
}

/** 沒有 UI (從編輯器執行) 時一律回傳 false，避免未經確認就寄信。 */
function confirm_(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(title, message, ui.ButtonSet.YES_NO) === ui.Button.YES;
  } catch (error) {
    Logger.log('此功能需要從試算表選單「授權管理」執行。');
    return false;
  }
}
