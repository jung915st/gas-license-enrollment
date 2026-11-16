// 教學軟體授權登記與分配系統
// 版本 v3.1
// 更新：表單送出改為「登記」，依登記先後排定正取／備取，不再自動發放序號。
//       管理者在「申請登記」分頁勾選核准後，再由 allocate.gs 從「序號列表」發放並寄信通知。
// v3.1：改用 setupForm() 建立的新表單 (舊表單的回應一律忽略，避免重複紀錄)；
//       職務題依 學年／科任／行政 分支；通知信寄到開通授權的 Google 帳號，
//       並副本給 Google AI Pro 的學校帳號。
//
// 檔案分工：
//   code.gs             設定、表單送出登記、正取／備取排序、統計
//   allocate.gs         管理者選單、勾選核准後發放、寄送結果通知、補登
//   setupForm.gs        初始化分頁、建立／更新申請表單、表單顯示剩餘名額、觸發器
//   tests_enrollment.gs 離線單元測試 (runEnrollmentTests)

// === 全域設定 ===
const LICENSE_SHEET_NAME = "序號列表";   // A 序號、B 使用者Email、C 分配時間、D 軟體 (留空視為 HiTeach 專業版)
const ADMIN_EMAIL = "lres712@mail.lres.tc.edu.tw"; // 管理者通知信箱
const SCHOOL_DOMAIN = "mail.lres.tc.edu.tw";       // 學校 Google Workspace 網域 (Google AI Pro 以此帳號開通)

const SOFTWARE_SHEET_NAME = "軟體清單"; // A 軟體名稱、B 可用授權數、C 分配方式、D 申請限制
const REG_SHEET_NAME = "申請登記";      // 一列 = 一位申請人的一套軟體
const STATS_SHEET_NAME = "申請統計";    // 各軟體登記人數與剩餘名額 (自動產生，請勿手動編輯)

// 表單題目名稱：由 setupForm() 建立。若在表單上改了題目名稱，這裡要同步修改 (空白與大小寫差異會自動忽略)
const Q_NAME = '您的大名';
const Q_EMAIL = '您要用來開通授權的 Google 帳號';   // 授權通知信的收件信箱
const Q_ROLE = '您的職務（學年／科任／行政）';       // 依選項分支到下面三題之一
const Q_GRADE = '任教學年';
const Q_SUBJECT = '任教科目';
const Q_OFFICE = '所屬處室';
const Q_SOFTWARE = '您要申請的軟體授權';
const Q_WORKSPACE = '申請 Google AI Pro 教育版的學校帳號'; // 選填；授權通知信的副本收件信箱

const ROLE = { GRADE: '學年', SUBJECT: '科任', OFFICE: '行政' };
const GRADE_CHOICES = ['一年級', '二年級', '三年級', '四年級', '五年級', '六年級'];
const OFFICE_CHOICES = ['教務處', '學務處', '總務處', '輔導室'];

// 「每組限一位」軟體可申請的組別：各學年＋科任 (行政不可申請)
const LIMITED_GROUPS = GRADE_CHOICES.concat([ROLE.SUBJECT]);

// 以學校帳號開通的軟體：通知信會寫明開通在學校帳號
const WORKSPACE_ACCOUNT_SOFTWARE = 'Google AI Pro 教育版';

// 指令碼屬性：setupForm() 建立新表單時寫入，用來辨識哪一個回應分頁屬於這份表單
const PROP = { FORM_ID: 'ENROLL_FORM_ID', RESPONSE_SHEET_ID: 'ENROLL_RESPONSE_SHEET_ID' };

const MODE = { SERIAL: '序號', ACCOUNT: '帳號' };
const LIMIT_PER_GROUP_TEXT = '各學年、科任各一位';

// 「軟體清單」分頁的預設內容 (分頁不存在時也以此為準)
// 欄位：軟體名稱、可用授權數、分配方式 (序號 = 從序號列表取號；帳號 = 直接開通在登記的帳號)、申請限制 (有填 = 每組限一位)
const DEFAULT_SOFTWARE = [
  ['Wayground (Quizizz)', 5, MODE.ACCOUNT, ''],
  ['Kahoot 教師版', 5, MODE.ACCOUNT, ''],
  ['Padlet 學校版', 10, MODE.ACCOUNT, ''],
  ['Wordwall 專業版', 10, MODE.ACCOUNT, ''],
  ['HiTeach 專業版', 15, MODE.SERIAL, ''],
  ['Google AI Pro 教育版', 7, MODE.ACCOUNT, LIMIT_PER_GROUP_TEXT]
];

// 「序號列表」D 欄留空時，視為這套軟體的序號 (相容舊版只有 HiTeach 的序號表)
const DEFAULT_SERIAL_SOFTWARE = 'HiTeach 專業版';

const STATUS = { OK: '正取', WAIT: '備取', INELIGIBLE: '不符資格', CANCEL: '取消', ISSUED: '已發放' };

// 「申請登記」分頁欄位 (1 起算)
const COL = {
  TIME: 1, NAME: 2, EMAIL: 3, IDENTITY: 4, SOFTWARE: 5, STATUS: 6,
  RANK: 7, APPROVED: 8, SERIAL: 9, ISSUED_AT: 10, NOTE: 11, WORKSPACE: 12
};
const REG_HEADERS = ['登記時間', '姓名', 'Google 帳號', '職務', '軟體', '狀態', '順位', '核准', '序號', '發放時間', '備註', '學校帳號'];

const SCHOOL_SIGNATURE = '臺中市北區立人國民小學 資訊組 敬上';

// =====================================================================
// 表單送出：登記
// =====================================================================

function onFormSubmit(e) {
  // 檢查事件物件 e 是否存在。不存在代表是手動執行，而非由表單提交觸發。
  if (!e || !e.namedValues) {
    Logger.log("此函式應由 Google 表單提交觸發，請勿手動執行。");
    return;
  }
  // 只處理 setupForm() 建立的那份表單；舊表單或其他表單的回應一律忽略，避免重複紀錄
  if (!isEnrollmentResponseSheet_(e.range ? e.range.getSheet() : null)) {
    Logger.log("此筆回應不是來自授權申請表單 (或尚未執行 setupForm 建立表單)，不處理。");
    return;
  }
  Logger.log("收到新的表單提交，資料如下：");
  Logger.log(JSON.stringify(e.namedValues, null, 2));

  // 鎖定服務，避免多人同時送出時順位互相覆蓋
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(30000);
    locked = true;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const softwareList = getSoftwareList_(ss);
    const sub = parseSubmission_(e.namedValues, softwareList, submissionTime_(e));

    if (sub.selections.length === 0) {
      Logger.log("此筆回應沒有勾選任何清單內的軟體，不登記。");
      return;
    }

    const result = processSubmission_(ss, sub, softwareList);
    const stats = refreshStats_(ss, softwareList, result.rows);

    // 表單上的剩餘名額更新失敗不應影響登記
    try {
      refreshFormRemaining_(ss, stats);
    } catch (formError) {
      Logger.log("更新表單剩餘名額失敗：" + formError);
    }

    // 登記已寫入試算表；回條寄送失敗只通知管理者，不視為登記失敗
    const outcomes = buildOutcomes_(result.rows, sub, result.added);
    let receiptProblem = isEmail_(sub.email) ? '' : 'Google 帳號格式有誤';
    if (!receiptProblem) {
      try {
        MailApp.sendEmail(withCc_({
          to: sub.email,
          subject: "【已登記】教學軟體授權申請",
          htmlBody: buildReceiptHtml_(sub, outcomes)
        }, sub.email, sub.workspace));
        Logger.log("登記回條已寄出：" + sub.email);
      } catch (mailError) {
        receiptProblem = String(mailError);
      }
    }
    if (receiptProblem) {
      MailApp.sendEmail(ADMIN_EMAIL, "【提醒】教學軟體授權登記回條未寄出",
        `${sub.name} (${sub.email}) 已完成登記，但回條未寄出：${receiptProblem}。請與申請人確認帳號。`);
    }

    notifyAdminWhenFull_(outcomes, sub);

  } catch (error) {
    Logger.log("發生未預期的錯誤：" + error);
    try {
      MailApp.sendEmail(ADMIN_EMAIL, "【錯誤】教學軟體授權登記系統發生錯誤",
        error.toString() + "\n\n表單資料：\n" + JSON.stringify(e.namedValues, null, 2) +
        "\n\n修正後可由選單「授權管理 → 從表單回應補登」補回這筆登記。");
    } catch (mailError) {
      Logger.log("錯誤通知信寄送失敗：" + mailError);
    }
  } finally {
    if (locked) lock.releaseLock();
    Logger.log("程式執行完畢。");
  }
}

/**
 * 將一筆申請寫入「申請登記」並重新排定正取／備取。
 * @returns {{rows: Object[], added: string[], duplicates: string[]}}
 */
function processSubmission_(ss, sub, softwareList) {
  const sheet = getRegSheet_(ss);
  let rows = readRegistrations_(sheet);
  const split = splitSelections_(rows, sub);
  if (split.toAdd.length > 0) {
    appendRegistrations_(sheet, sub, split.toAdd);
    rows = readRegistrations_(sheet);
  }
  applyRanks_(sheet, rows, softwareList);
  return { rows: rows, added: split.toAdd, duplicates: split.duplicates };
}

/** 重新排序並寫回狀態與順位；rows 內的 status / rank 也會同步更新。 */
function applyRanks_(sheet, rows, softwareList) {
  if (rows.length === 0) return;
  const ranks = rankRegistrations_(rows, softwareList);
  const values = rows.map(function (row, i) {
    row.status = ranks[i].status;
    row.rank = ranks[i].rank;
    return [row.status, row.rank];
  });
  sheet.getRange(2, COL.STATUS, values.length, 2).setValues(values);
}

// =====================================================================
// 純邏輯 (不碰試算表，可離線測試)
// =====================================================================

function normalize_(text) {
  return String(text == null ? '' : text).toLowerCase().replace(/\s+/g, '');
}

function isEmail_(text) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(text || '').trim());
}

function timeValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return value.getTime();
  const parsed = new Date(value).getTime();
  return isNaN(parsed) ? 0 : parsed;
}

function escapeHtml_(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 將「軟體清單」的列轉成設定物件。 */
function parseSoftwareRows_(values) {
  const list = [];
  values.forEach(function (row) {
    const name = String(row[0] == null ? '' : row[0]).trim();
    if (!name) return;
    list.push({
      name: name,
      quota: Math.max(0, Math.floor(Number(row[1]) || 0)),
      mode: String(row[2] || '').indexOf(MODE.ACCOUNT) >= 0 ? MODE.ACCOUNT : MODE.SERIAL,
      limited: String(row[3] == null ? '' : row[3]).trim() !== ''
    });
  });
  return list;
}

/** 取出某一題的答案；同名題目 (舊題刪除後重建) 會合併非空白的值。找不到題目回傳 null。 */
function answerOf_(namedValues, title) {
  const want = normalize_(title);
  let found = false;
  const parts = [];
  Object.keys(namedValues).forEach(function (key) {
    if (normalize_(key) !== want) return;
    found = true;
    [].concat(namedValues[key]).forEach(function (value) {
      const text = String(value == null ? '' : value).trim();
      if (text) parts.push(text);
    });
  });
  return found ? parts.join(', ') : null;
}

/** 從複選題答案找出勾選了清單中的哪些軟體。 */
function parseSoftwareSelection_(answer, softwareList) {
  const text = normalize_(answer);
  return softwareList
    .filter(function (sw) { return text.indexOf(normalize_(sw.name)) >= 0; })
    .map(function (sw) { return sw.name; });
}

function parseSubmission_(namedValues, softwareList, time) {
  const name = answerOf_(namedValues, Q_NAME);
  const email = answerOf_(namedValues, Q_EMAIL);
  const software = answerOf_(namedValues, Q_SOFTWARE);
  const role = answerOf_(namedValues, Q_ROLE);

  const missing = [];
  if (name === null) missing.push(Q_NAME);
  if (email === null) missing.push(Q_EMAIL);
  if (role === null) missing.push(Q_ROLE);
  if (software === null) missing.push(Q_SOFTWARE);
  if (missing.length > 0) {
    throw new Error('表單找不到題目：「' + missing.join('」、「') + '」。目前表單題目為：' +
      Object.keys(namedValues).join('、') + '。請確認表單題目未被改名，或修改 code.gs 的 Q_ 設定。');
  }
  const identity = buildIdentity_(role,
    answerOf_(namedValues, Q_GRADE), answerOf_(namedValues, Q_SUBJECT), answerOf_(namedValues, Q_OFFICE));
  return {
    time: time || new Date(),
    name: name,
    email: email,
    identity: identity,
    group: groupOf_(identity),
    workspace: answerOf_(namedValues, Q_WORKSPACE) || '',
    selections: parseSoftwareSelection_(software, softwareList)
  };
}

/**
 * 職務文字：「學年／三年級」「科任／自然」「行政／教務處」。
 * 只採用與職務對應的那一題，其他分支留下的舊答案不採用。
 */
function buildIdentity_(role, grade, subject, office) {
  const text = String(role == null ? '' : role).trim();
  let detail = '';
  if (text.indexOf(ROLE.GRADE) >= 0) detail = grade;
  else if (text.indexOf(ROLE.SUBJECT) >= 0) detail = subject;
  else if (text.indexOf(ROLE.OFFICE) >= 0) detail = office;
  const base = [ROLE.GRADE, ROLE.SUBJECT, ROLE.OFFICE].filter(function (r) { return text.indexOf(r) >= 0; })[0] || text;
  detail = String(detail == null ? '' : detail).trim();
  return detail ? base + '／' + detail : base;
}

/** 「每組限一位」用的組別：學年 → 該年級；科任 → 科任；其餘 (行政) → 職務本身。 */
function groupOf_(identity) {
  const parts = String(identity == null ? '' : identity).split('／');
  const base = parts[0].trim();
  if (base === ROLE.GRADE) return (parts[1] || '').trim();
  return base;
}

/** 副本收件人：有填學校帳號、格式正確、且與主要帳號不同時才加。 */
function ccFor_(email, workspace) {
  if (!isEmail_(workspace)) return '';
  return normalize_(workspace) === normalize_(email) ? '' : String(workspace).trim();
}

function withCc_(message, email, workspace) {
  const cc = ccFor_(email, workspace);
  if (cc) message.cc = cc;
  return message;
}

/** 登記時寫入備註欄的提醒。 */
function registrationNote_(sub, software) {
  const notes = [];
  if (!isEmail_(sub.email)) notes.push('Google 帳號格式有誤，無法寄送通知');
  if (software === WORKSPACE_ACCOUNT_SOFTWARE) {
    if (!isEmail_(sub.workspace)) notes.push('未填學校帳號');
    else if (normalize_(sub.workspace).slice(-(SCHOOL_DOMAIN.length + 1)) !== '@' + SCHOOL_DOMAIN) notes.push('學校帳號非 @' + SCHOOL_DOMAIN);
  }
  return notes.join('；');
}

/**
 * 區分本次要新增的軟體與先前已登記過的軟體 (同一 Email＋軟體只算一筆)。
 * 已取消的登記：申請人重新送出表單可再登記；但同一筆回應 (時間不晚於取消列) 不會被補登救回來。
 */
function splitSelections_(rows, sub) {
  const email = normalize_(sub.email);
  const subTime = timeValue_(sub.time);
  const toAdd = [];
  const duplicates = [];
  sub.selections.forEach(function (software) {
    const exists = rows.some(function (row) {
      if (row.software !== software || normalize_(row.email) !== email) return false;
      if (row.status !== STATUS.CANCEL) return true;
      return timeValue_(row.time) >= subTime;
    });
    (exists ? duplicates : toAdd).push(software);
  });
  return { toAdd: toAdd, duplicates: duplicates };
}

/**
 * 依登記先後排定每一列的狀態與順位。
 * - 已發放：維持不動，佔用名額。
 * - 取消：維持不動，不佔名額。
 * - 其餘依時間排序：名額內為正取，額滿後為備取。
 * - 「每組限一位」的軟體：非 LIMITED_GROUPS 的身分為不符資格；同組已有人則為備取。
 * @returns {{status: string, rank: (number|string)}[]} 與 rows 同順序
 */
function rankRegistrations_(rows, softwareList) {
  const result = rows.map(function (row) { return { status: row.status, rank: row.rank === undefined ? '' : row.rank }; });
  const byName = {};
  softwareList.forEach(function (sw) { byName[sw.name] = sw; });

  const indexBySoftware = {};
  rows.forEach(function (row, i) {
    (indexBySoftware[row.software] = indexBySoftware[row.software] || []).push(i);
  });

  Object.keys(indexBySoftware).forEach(function (name) {
    const sw = byName[name];
    if (!sw) return; // 已從軟體清單移除的項目不動
    const indexes = indexBySoftware[name];
    let used = 0;
    const takenGroups = {};

    indexes.forEach(function (i) {
      if (rows[i].status === STATUS.ISSUED) {
        used++;
        takenGroups[rows[i].group] = true;
        result[i] = { status: STATUS.ISSUED, rank: '' };
      } else if (rows[i].status === STATUS.CANCEL) {
        result[i] = { status: STATUS.CANCEL, rank: '' };
      }
    });

    const pending = indexes
      .filter(function (i) { return rows[i].status !== STATUS.ISSUED && rows[i].status !== STATUS.CANCEL; })
      .sort(function (a, b) { return (timeValue_(rows[a].time) - timeValue_(rows[b].time)) || (a - b); });

    let okRank = 0;
    let waitRank = 0;
    pending.forEach(function (i) {
      const group = rows[i].group;
      if (sw.limited && LIMITED_GROUPS.indexOf(group) < 0) {
        result[i] = { status: STATUS.INELIGIBLE, rank: '' };
        return;
      }
      const groupFree = !sw.limited || !takenGroups[group];
      if (groupFree && used < sw.quota) {
        used++;
        takenGroups[group] = true;
        okRank++;
        result[i] = { status: STATUS.OK, rank: okRank };
      } else {
        waitRank++;
        result[i] = { status: STATUS.WAIT, rank: waitRank };
      }
    });
  });
  return result;
}

/** 各軟體的登記統計。剩餘名額 = 可用授權數 − 正取 − 已發放。 */
function computeStats_(rows, softwareList) {
  return softwareList.map(function (sw) {
    const mine = rows.filter(function (row) { return row.software === sw.name; });
    const count = function (status) {
      return mine.filter(function (row) { return row.status === status; }).length;
    };
    const ok = count(STATUS.OK);
    const wait = count(STATUS.WAIT);
    const issued = count(STATUS.ISSUED);
    const stat = {
      name: sw.name,
      quota: sw.quota,
      mode: sw.mode,
      limited: sw.limited,
      registered: ok + wait + issued,
      ok: ok,
      wait: wait,
      ineligible: count(STATUS.INELIGIBLE),
      approved: mine.filter(function (row) {
        return row.approved === true && row.status !== STATUS.ISSUED && row.status !== STATUS.CANCEL;
      }).length,
      issued: issued,
      remaining: Math.max(0, sw.quota - ok - issued),
      openGroups: null
    };
    if (sw.limited) {
      const taken = {};
      mine.forEach(function (row) {
        if (row.status === STATUS.OK || row.status === STATUS.ISSUED) taken[row.group] = true;
      });
      stat.openGroups = LIMITED_GROUPS.filter(function (group) { return !taken[group]; });
      stat.remaining = Math.min(stat.remaining, stat.openGroups.length);
    }
    return stat;
  });
}

/** 顯示在表單題目下方的剩餘名額說明。 */
function buildRemainingText_(stats, updatedAt) {
  const lines = [
    '每人每項限登記 1 個，依登記先後排序；額滿後仍可登記，將列為備取。',
    '目前剩餘名額' + (updatedAt ? '（' + updatedAt + ' 更新）' : '') + '：'
  ];
  stats.forEach(function (stat) {
    let line = '・' + stat.name + '：';
    line += stat.remaining > 0
      ? '剩餘 ' + stat.remaining + '／' + stat.quota
      : '已額滿（備取 ' + stat.wait + ' 人）';
    if (stat.limited) {
      line += '｜限' + LIMIT_PER_GROUP_TEXT;
      if (stat.remaining > 0 && stat.openGroups) line += '，尚可申請：' + stat.openGroups.join('、');
    }
    lines.push(line);
  });
  return lines.join('\n');
}

/** 整理申請人本次勾選的每套軟體的結果，供回條與管理者通知使用。 */
function buildOutcomes_(rows, sub, added) {
  const email = normalize_(sub.email);
  return sub.selections.map(function (software) {
    const row = rows.filter(function (r) {
      return r.software === software && normalize_(r.email) === email && r.status !== STATUS.CANCEL;
    })[0];
    return {
      software: software,
      status: row ? row.status : '',
      rank: row ? row.rank : '',
      isNew: added.indexOf(software) >= 0
    };
  });
}

function describeOutcome_(outcome) {
  let text;
  if (outcome.status === STATUS.OK) {
    text = '正取（第 ' + outcome.rank + ' 位）';
  } else if (outcome.status === STATUS.WAIT) {
    text = '備取（備取第 ' + outcome.rank + ' 位）－目前名額已滿，如有名額釋出將依序遞補';
  } else if (outcome.status === STATUS.INELIGIBLE) {
    text = '不符申請資格（限' + LIMIT_PER_GROUP_TEXT + '）';
  } else if (outcome.status === STATUS.ISSUED) {
    text = '已發放';
  } else {
    text = '已登記';
  }
  return outcome.isNew ? text : '您先前已登記過，維持原登記：' + text;
}

function buildReceiptHtml_(sub, outcomes) {
  const items = outcomes.map(function (outcome) {
    return '<li><b>' + escapeHtml_(outcome.software) + '</b>：' + escapeHtml_(describeOutcome_(outcome)) + '</li>';
  }).join('');
  const needsWorkspace = !isEmail_(sub.workspace) && outcomes.some(function (outcome) {
    return outcome.software === WORKSPACE_ACCOUNT_SOFTWARE && outcome.status !== STATUS.INELIGIBLE;
  });
  return '<p>親愛的 ' + escapeHtml_(sub.name) + ' 老師，您好：</p>' +
    '<p>已收到您的教學軟體授權登記，結果如下：</p>' +
    '<ul>' + items + '</ul>' +
    (needsWorkspace ? '<p>您申請了 ' + escapeHtml_(WORKSPACE_ACCOUNT_SOFTWARE) + '，但未填寫學校帳號（@' + SCHOOL_DOMAIN +
      '）。這項授權需開通在學校帳號，請洽資訊組補填。</p>' : '') +
    '<p>目前<b>尚未發放授權</b>。資訊組審核後，將另行寄信通知授權序號或開通方式。</p>' +
    '<br><p>' + SCHOOL_SIGNATURE + '</p>' +
    '<p style="color: grey; font-size: 12px;">此為系統自動發送信件，請勿直接回覆。</p>';
}

// =====================================================================
// 試算表讀寫
// =====================================================================

function getSoftwareList_(ss) {
  const sheet = ss.getSheetByName(SOFTWARE_SHEET_NAME);
  if (sheet && sheet.getLastRow() >= 2) {
    const list = parseSoftwareRows_(sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues());
    if (list.length > 0) return list;
  }
  return parseSoftwareRows_(DEFAULT_SOFTWARE);
}

/** 取得「申請登記」分頁，不存在時自動建立。 */
function getRegSheet_(ss) {
  let sheet = ss.getSheetByName(REG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(REG_SHEET_NAME);
    sheet.getRange(1, 1, 1, REG_HEADERS.length).setValues([REG_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRegistrations_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, REG_HEADERS.length).getValues().map(function (v, i) {
    return {
      rowIndex: i + 2,
      time: v[COL.TIME - 1],
      name: String(v[COL.NAME - 1]).trim(),
      email: String(v[COL.EMAIL - 1]).trim(),
      identity: String(v[COL.IDENTITY - 1]).trim(),
      group: groupOf_(v[COL.IDENTITY - 1]),
      software: String(v[COL.SOFTWARE - 1]).trim(),
      status: String(v[COL.STATUS - 1]).trim(),
      rank: v[COL.RANK - 1],
      approved: v[COL.APPROVED - 1] === true,
      serial: String(v[COL.SERIAL - 1]).trim(),
      issuedAt: v[COL.ISSUED_AT - 1],
      note: String(v[COL.NOTE - 1]).trim(),
      workspace: String(v[COL.WORKSPACE - 1]).trim()
    };
  });
}

function appendRegistrations_(sheet, sub, softwareNames) {
  const values = softwareNames.map(function (software) {
    return [sub.time, sub.name, sub.email, sub.identity, software, '', '', false, '', '',
      registrationNote_(sub, software), sub.workspace];
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, values.length, REG_HEADERS.length).setValues(values);
  sheet.getRange(startRow, COL.TIME, values.length, 1).setNumberFormat('yyyy/MM/dd HH:mm:ss');
  sheet.getRange(startRow, COL.APPROVED, values.length, 1).insertCheckboxes();
}

/** 重寫「申請統計」分頁並回傳統計結果。 */
function refreshStats_(ss, softwareList, rows) {
  const stats = computeStats_(rows, softwareList);
  let sheet = ss.getSheetByName(STATS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(STATS_SHEET_NAME);
  const headers = ['軟體', '可用授權數', '登記人數', '正取', '備取', '已勾選核准（未發放）', '已發放', '剩餘名額', '分配方式', '備註'];
  const values = [headers].concat(stats.map(function (stat) {
    let note = '';
    if (stat.limited) {
      note = '限' + LIMIT_PER_GROUP_TEXT + '；尚可申請：' + (stat.openGroups.length ? stat.openGroups.join('、') : '無');
      if (stat.ineligible > 0) note += '；不符資格 ' + stat.ineligible + ' 人';
    }
    return [stat.name, stat.quota, stat.registered, stat.ok, stat.wait, stat.approved, stat.issued, stat.remaining, stat.mode, note];
  }));
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  sheet.getRange(values.length + 2, 1).setValue('更新時間：' + nowText_());
  sheet.setFrozenRows(1);
  return stats;
}

/**
 * 這個分頁是否為 setupForm() 所建立表單的回應分頁。
 * 尚未建立表單時一律回傳 false，所以舊表單的回應不會被登記。
 */
function isEnrollmentResponseSheet_(sheet) {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty(PROP.FORM_ID);
  if (!formId || !sheet) return false;
  const cached = props.getProperty(PROP.RESPONSE_SHEET_ID);
  if (cached) return String(sheet.getSheetId()) === cached;
  // 建立表單當下還找不到回應分頁時，改在第一次收到回應時比對並記住
  let match = false;
  try {
    const url = sheet.getFormUrl();
    match = !!url && FormApp.openByUrl(url).getId() === formId;
  } catch (error) {
    Logger.log("比對表單失敗：" + error);
  }
  if (match) props.setProperty(PROP.RESPONSE_SHEET_ID, String(sheet.getSheetId()));
  return match;
}

/** 授權申請表單的回應分頁；找不到回傳 null。 */
function getResponseSheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getFormUrl() && isEnrollmentResponseSheet_(sheets[i])) return sheets[i];
  }
  return null;
}

/** setupForm() 建立的表單；尚未建立回傳 null。 */
function getEnrollmentForm_() {
  const formId = PropertiesService.getScriptProperties().getProperty(PROP.FORM_ID);
  return formId ? FormApp.openById(formId) : null;
}

function nowText_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
}

/** 表單回應那一列的時間戳記；取不到時用現在時間。 */
function submissionTime_(e) {
  try {
    if (e.range) {
      const value = e.range.getSheet().getRange(e.range.getRow(), 1).getValue();
      if (Object.prototype.toString.call(value) === '[object Date]') return value;
    }
  } catch (error) {
    Logger.log("讀取時間戳記失敗，改用現在時間：" + error);
  }
  return new Date();
}

/** 某套軟體第一次出現備取時通知管理者。 */
function notifyAdminWhenFull_(outcomes, sub) {
  const full = outcomes.filter(function (outcome) {
    return outcome.isNew && outcome.status === STATUS.WAIT && outcome.rank === 1;
  });
  if (full.length === 0) return;
  const names = full.map(function (outcome) { return outcome.software; }).join('、');
  MailApp.sendEmail(ADMIN_EMAIL, "【系統通知】教學軟體授權登記已額滿：" + names,
    `下列軟體的登記人數已超過可用授權數，之後的登記將列為備取：${names}\n` +
    `第一位備取：${sub.name} (${sub.email})，時間 ${nowText_()}。`);
}
