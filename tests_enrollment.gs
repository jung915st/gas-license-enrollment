/**
 * tests_enrollment.gs — 離線單元測試
 *
 * 在 Apps Script 編輯器執行 runEnrollmentTests()，結果寫在執行紀錄。
 * 只測純邏輯 (排序、統計、發放計畫、表單答案解析)，不讀寫試算表、不寄信，
 * 可在正式專案安全執行。
 */

function runEnrollmentTests() {
  const results = [];
  const check = function (name, actual, expected) {
    const a = JSON.stringify(actual);
    const b = JSON.stringify(expected);
    results.push({ name: name, ok: a === b, detail: a === b ? '' : '得到 ' + a + '，預期 ' + b });
  };

  const softwareList = parseSoftwareRows_(DEFAULT_SOFTWARE);
  const at = function (minute) { return new Date(2026, 10, 16, 9, minute, 0); };
  let rowIndex = 1;
  const reg = function (minute, email, group, software, extra) {
    rowIndex++;
    const row = {
      rowIndex: rowIndex, time: at(minute), name: email.split('@')[0], email: email, group: group,
      software: software, status: '', rank: '', approved: false, serial: '', issuedAt: '', note: ''
    };
    Object.keys(extra || {}).forEach(function (key) { row[key] = extra[key]; });
    return row;
  };
  const applyRank = function (rows) {
    rankRegistrations_(rows, softwareList).forEach(function (r, i) {
      rows[i].status = r.status;
      rows[i].rank = r.rank;
    });
    return rows;
  };
  const statuses = function (rows) { return rows.map(function (row) { return row.status + row.rank; }); };

  // --- 軟體清單 ---
  check('預設軟體清單共 6 套', softwareList.length, 6);
  check('預設可用授權數', softwareList.map(function (sw) { return sw.quota; }), [5, 5, 10, 10, 15, 7]);
  check('只有 Google AI Pro 有每組限一位', softwareList.filter(function (sw) { return sw.limited; })
    .map(function (sw) { return sw.name; }), ['Google AI Pro 教育版']);
  check('空白列與非數字數量', parseSoftwareRows_([['', 3, '', ''], ['X', 'abc', '帳號開通', '']]),
    [{ name: 'X', quota: 0, mode: '帳號', limited: false }]);

  // --- 表單答案解析 ---
  check('複選答案解析', parseSoftwareSelection_('Kahoot 教師版, HiTeach 專業版', softwareList),
    ['Kahoot 教師版', 'HiTeach 專業版']);
  check('忽略大小寫與空白', parseSoftwareSelection_('hiteach專業版,WAYGROUND (quizizz)', softwareList),
    ['Wayground (Quizizz)', 'HiTeach 專業版']);
  check('清單外的選項不計', parseSoftwareSelection_('Canva', softwareList), []);
  check('同名題目合併非空白值', answerOf_({ '您要申請的軟體授權': [''], ' 您要申請的軟體授權 ': ['Kahoot 教師版'] }, Q_SOFTWARE),
    'Kahoot 教師版');
  check('找不到題目回傳 null', answerOf_({ '其他': ['x'] }, Q_NAME), null);

  let thrown = '';
  try {
    parseSubmission_({ '您的大名': ['王小明'] }, softwareList, at(0));
  } catch (error) {
    thrown = String(error.message).indexOf(Q_EMAIL) >= 0 ? 'missing-email' : 'other';
  }
  check('缺少題目時丟出錯誤並指出題目', thrown, 'missing-email');

  // --- 正取／備取 ---
  const kahoot = [];
  for (let i = 0; i < 7; i++) kahoot.push(reg(i, 't' + i + '@x.tw', '一年級', 'Kahoot 教師版'));
  check('額滿 5 名後為備取', statuses(applyRank(kahoot)), ['正取1', '正取2', '正取3', '正取4', '正取5', '備取1', '備取2']);

  const late = [reg(30, 'late@x.tw', '科任', 'Kahoot 教師版'), reg(5, 'early@x.tw', '科任', 'Kahoot 教師版')];
  check('依時間而非列順序排序', statuses(applyRank(late)), ['正取2', '正取1']);

  kahoot[1].status = STATUS.CANCEL;
  check('取消後備取遞補', statuses(applyRank(kahoot)), ['正取1', '取消', '正取2', '正取3', '正取4', '正取5', '備取1']);

  kahoot[0].status = STATUS.ISSUED;
  check('已發放佔用名額且狀態不變', statuses(applyRank(kahoot)), ['已發放', '取消', '正取1', '正取2', '正取3', '正取4', '備取1']);

  // --- 每組限一位 ---
  const ai = applyRank([
    reg(1, 'a@x.tw', '一年級', 'Google AI Pro 教育版'),
    reg(2, 'b@x.tw', '一年級', 'Google AI Pro 教育版'),
    reg(3, 'c@x.tw', '科任', 'Google AI Pro 教育版'),
    reg(4, 'd@x.tw', '行政／其他', 'Google AI Pro 教育版'),
    reg(5, 'e@x.tw', '', 'Google AI Pro 教育版')
  ]);
  check('同組第二位備取、行政與未填不符資格', statuses(ai), ['正取1', '備取1', '正取2', '不符資格', '不符資格']);

  const aiStats = computeStats_(ai, softwareList).filter(function (s) { return s.name === 'Google AI Pro 教育版'; })[0];
  check('限制型軟體剩餘名額', aiStats.remaining, 5);
  check('限制型軟體尚可申請組別', aiStats.openGroups, ['二年級', '三年級', '四年級', '五年級', '六年級']);

  // --- 統計與剩餘名額文字 ---
  const kahootStats = computeStats_(kahoot, softwareList).filter(function (s) { return s.name === 'Kahoot 教師版'; })[0];
  check('統計：登記／正取／備取／已發放／剩餘',
    [kahootStats.registered, kahootStats.ok, kahootStats.wait, kahootStats.issued, kahootStats.remaining], [6, 4, 1, 1, 0]);
  const text = buildRemainingText_(computeStats_(kahoot.concat(ai), softwareList), '');
  check('剩餘文字：額滿', text.indexOf('・Kahoot 教師版：已額滿（備取 1 人）') >= 0, true);
  check('剩餘文字：未額滿', text.indexOf('・HiTeach 專業版：剩餘 15／15') >= 0, true);
  check('剩餘文字：限制型', text.indexOf('・Google AI Pro 教育版：剩餘 5／7｜限各學年、科任各一位，尚可申請：二年級') >= 0, true);

  // --- 重複登記 ---
  const existing = applyRank([reg(1, 'Dup@x.tw', '科任', 'Kahoot 教師版'), reg(2, 'gone@x.tw', '科任', 'Kahoot 教師版')]);
  existing[1].status = STATUS.CANCEL;
  check('同一 Email 同一軟體不重複 (忽略大小寫)',
    splitSelections_(existing, { email: 'dup@x.tw', time: at(9), selections: ['Kahoot 教師版', 'Padlet 學校版'] }),
    { toAdd: ['Padlet 學校版'], duplicates: ['Kahoot 教師版'] });
  check('取消後重新送出可再登記',
    splitSelections_(existing, { email: 'gone@x.tw', time: at(9), selections: ['Kahoot 教師版'] }).toAdd, ['Kahoot 教師版']);
  check('補登不會救回已取消的同一筆回應',
    splitSelections_(existing, { email: 'gone@x.tw', time: at(2), selections: ['Kahoot 教師版'] }).toAdd, []);

  // --- 發放計畫 ---
  const serials = [
    { rowIndex: 2, serial: 'OLD-1', email: 'used@x.tw', software: '' },
    { rowIndex: 3, serial: 'HT-1', email: '', software: '' },
    { rowIndex: 4, serial: 'KH-1', email: '', software: 'Kahoot' },
    { rowIndex: 5, serial: 'HT-2', email: '', software: 'HiTeach 專業版' }
  ];
  const pool = applyRank([
    reg(1, 'p1@x.tw', '一年級', 'HiTeach 專業版', { approved: true }),
    reg(2, 'p2@x.tw', '一年級', 'HiTeach 專業版', { approved: true }),
    reg(3, 'p3@x.tw', '一年級', 'HiTeach 專業版', { approved: true }),
    reg(4, 'p4@x.tw', '一年級', 'HiTeach 專業版'),
    reg(5, 'p1@x.tw', '一年級', 'Kahoot 教師版', { approved: true }),
    reg(6, 'p1@x.tw', '一年級', 'Wordwall 專業版', { approved: true }),
    reg(7, 'p1@x.tw', '一年級', 'Google AI Pro 教育版', { approved: true }),
    reg(8, 'p2@x.tw', '一年級', 'Google AI Pro 教育版', { approved: true }),
    reg(9, 'p5@x.tw', '行政／其他', 'Google AI Pro 教育版', { approved: true }),
    reg(10, 'p6@x.tw', '科任', 'Padlet 學校版', { approved: true })
  ]);
  const plan = planIssue_(pool, softwareList, serials);
  check('發放：序號依軟體配對，舊序號不重發',
    plan.issue.map(function (item) { return item.row.email + '|' + item.row.software + '|' + item.serial; }),
    ['p1@x.tw|HiTeach 專業版|HT-1', 'p2@x.tw|HiTeach 專業版|HT-2', 'p1@x.tw|Kahoot 教師版|KH-1',
      'p1@x.tw|Wordwall 專業版|', 'p1@x.tw|Google AI Pro 教育版|']);
  check('發放：略過原因',
    plan.skipped.map(function (item) { return item.row.email + '|' + item.row.software + '|' + item.reason; }),
    ['p3@x.tw|HiTeach 專業版|「序號列表」沒有可用序號',
      'p2@x.tw|Google AI Pro 教育版|一年級已有一位取得授權',
      'p5@x.tw|Google AI Pro 教育版|身分不符（限各學年、科任各一位）',
      'p6@x.tw|Padlet 學校版|「序號列表」沒有可用序號']);
  check('發放：未勾選核准者不處理',
    plan.issue.concat(plan.skipped).some(function (item) { return item.row.email === 'p4@x.tw'; }), false);

  const over = [];
  for (let i = 0; i < 6; i++) over.push(reg(i, 'w' + i + '@x.tw', '科任', 'Wordwall 專業版', { approved: true }));
  const smallQuota = [{ name: 'Wordwall 專業版', quota: 4, mode: MODE.ACCOUNT, limited: false }];
  const overPlan = planIssue_(over, smallQuota, []);
  check('發放：不超過可用授權數', [overPlan.issue.length, overPlan.skipped.length], [4, 2]);

  // --- 回條文字 ---
  check('回條：正取', describeOutcome_({ status: STATUS.OK, rank: 3, isNew: true }), '正取（第 3 位）');
  check('回條：重複登記', describeOutcome_({ status: STATUS.WAIT, rank: 2, isNew: false }).indexOf('您先前已登記過') === 0, true);
  check('HTML 跳脫', escapeHtml_('<b>"A&B"</b>'), '&lt;b&gt;&quot;A&amp;B&quot;&lt;/b&gt;');
  check('Email 格式檢查', [isEmail_('a@b.tw'), isEmail_('a@b'), isEmail_('a b@c.tw')], [true, false, false]);

  const failed = results.filter(function (r) { return !r.ok; });
  results.forEach(function (r) { Logger.log((r.ok ? '通過  ' : '失敗  ') + r.name + (r.detail ? '：' + r.detail : '')); });
  Logger.log('合計 ' + results.length + ' 項，通過 ' + (results.length - failed.length) + ' 項，失敗 ' + failed.length + ' 項。');
  if (failed.length > 0) throw new Error('單元測試失敗 ' + failed.length + ' 項，詳見執行紀錄。');
  return results;
}
