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
      software: software, status: '', rank: '', approved: false, serial: '', issuedAt: '', note: '', workspace: ''
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
  check('預設軟體清單共 7 套，TeacherSay 在第一列', [softwareList.length, softwareList[0].name], [7, 'TeacherSay 專業版']);
  check('預設可用授權數 (TeacherSay 不限)', softwareList.map(function (sw) { return quotaText_(sw.quota); }),
    ['不限', 5, 5, 10, 9, 15, 7]);
  check('可用授權數「不限」的寫法', parseSoftwareRows_([['A', '不限', '', ''], ['B', ' 無限 ', '', ''], ['C', 'Unlimited', '', ''], ['D', '', '', '']])
    .map(function (sw) { return isFinite(sw.quota) ? sw.quota : 'inf'; }), ['inf', 'inf', 'inf', 0]);
  check('只有 Google AI Pro 有每組限一位', softwareList.filter(function (sw) { return sw.limited; })
    .map(function (sw) { return sw.name; }), ['Google AI Pro 教育版']);
  check('預設分配方式：只有 HiTeach 取序號', softwareList.filter(function (sw) { return sw.mode === MODE.SERIAL; })
    .map(function (sw) { return sw.name; }), ['HiTeach 專業版']);
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

  // --- 職務分支 ---
  const answers = function (role, extra) {
    const nv = {};
    nv[Q_NAME] = ['王小明'];
    nv[Q_EMAIL] = ['wang@gmail.com'];
    nv[Q_ROLE] = [role];
    nv[Q_GRADE] = [''];
    nv[Q_SUBJECT] = [''];
    nv[Q_OFFICE] = [''];
    nv[Q_SOFTWARE] = ['Google AI Pro 教育版'];
    nv[Q_WORKSPACE] = [''];
    Object.keys(extra || {}).forEach(function (key) { nv[key] = [extra[key]]; });
    return parseSubmission_(nv, softwareList, at(0));
  };
  const gradeExtra = {}; gradeExtra[Q_GRADE] = '三年級'; gradeExtra[Q_SUBJECT] = '自然';
  const gradeSub = answers('學年', gradeExtra);
  check('學年：職務與組別 (忽略其他分支的舊答案)', [gradeSub.identity, gradeSub.group], ['學年／三年級', '三年級']);
  const subjectExtra = {}; subjectExtra[Q_SUBJECT] = ' 自然 '; subjectExtra[Q_WORKSPACE] = 'wang@mail.lres.tc.edu.tw';
  const subjectSub = answers('科任', subjectExtra);
  check('科任：職務、組別、學校帳號', [subjectSub.identity, subjectSub.group, subjectSub.workspace],
    ['科任／自然', '科任', 'wang@mail.lres.tc.edu.tw']);
  const officeExtra = {}; officeExtra[Q_OFFICE] = '教務處';
  const officeSub = answers('行政', officeExtra);
  check('行政：職務與組別', [officeSub.identity, officeSub.group], ['行政／教務處', '行政']);
  check('行政不在可申請組別內', LIMITED_GROUPS.indexOf(officeSub.group), -1);
  const kgExtra = {}; kgExtra[Q_OFFICE] = '幼兒園';
  const kgSub = answers('行政', kgExtra);
  check('行政／幼兒園：職務與組別', [kgSub.identity, kgSub.group], ['行政／幼兒園', '行政']);
  check('行政處室選項含幼兒園', OFFICE_CHOICES.indexOf('幼兒園') >= 0, true);

  // --- 軟體題合併 ---
  const pick = function (texts, aiPro) { return collectSelections_(texts, aiPro, softwareList); };
  check('TeacherSay＋下拉選單一項，依清單順序', pick(['TeacherSay 專業版', 'HiTeach 專業版'], null),
    ['TeacherSay 專業版', 'HiTeach 專業版']);
  check('下拉選單選 Google AI Pro', pick(['TeacherSay 專業版', 'Google AI Pro 教育版'], null), ['TeacherSay 專業版', 'Google AI Pro 教育版']);
  check('「不申請其他軟體」不算一項', pick(['TeacherSay 專業版', NONE_CHOICE], null), ['TeacherSay 專業版']);
  check('v3.2 舊回應：核取方塊多項＋是／否題「是」', pick(['TeacherSay 專業版', 'HiTeach 專業版, Kahoot 教師版'], '是'),
    ['TeacherSay 專業版', 'Kahoot 教師版', 'HiTeach 專業版', 'Google AI Pro 教育版']);
  check('v3.2 舊回應：是／否題「否」不加入', pick(['TeacherSay 專業版', 'Padlet 學校版'], '否'), ['TeacherSay 專業版', 'Padlet 學校版']);
  check('v3.1 舊回應：Google AI Pro 在核取方塊內', pick([null, 'Google AI Pro 教育版, Wordwall 專業版'], null),
    ['Wordwall 專業版', 'Google AI Pro 教育版']);
  check('軟體清單沒有 Google AI Pro 時舊題「是」不加入',
    collectSelections_(['TeacherSay 專業版'], '是', parseSoftwareRows_([['TeacherSay 專業版', '不限', '帳號', '']])), ['TeacherSay 專業版']);

  const nvOf = function (role, detailKey, detail, extra) {
    const nv = {};
    nv[Q_NAME] = ['王']; nv[Q_EMAIL] = ['w@gmail.com']; nv[Q_ROLE] = [role]; nv[detailKey] = [detail];
    nv[Q_TEACHERSAY] = ['TeacherSay 專業版']; nv[Q_PICK] = ['']; nv[Q_PICK_OFFICE] = [''];
    nv[Q_WORKSPACE] = [''];
    Object.keys(extra || {}).forEach(function (key) { nv[key] = [extra[key]]; });
    return nv;
  };
  const aiPickNv = {}; aiPickNv[Q_PICK] = 'Google AI Pro 教育版'; aiPickNv[Q_WORKSPACE] = 'w@mail.lres.tc.edu.tw';
  check('parseSubmission_：學年選 Google AI Pro', parseSubmission_(nvOf('學年', Q_GRADE, '二年級', aiPickNv), softwareList, at(0)).selections,
    ['TeacherSay 專業版', 'Google AI Pro 教育版']);
  const officePickNv = {}; officePickNv[Q_PICK_OFFICE] = 'Wordwall 專業版'; officePickNv[Q_PICK] = 'Kahoot 教師版';
  check('parseSubmission_：行政只採用行政的選單', parseSubmission_(nvOf('行政', Q_OFFICE, '幼兒園', officePickNv), softwareList, at(0)).selections,
    ['TeacherSay 專業版', 'Wordwall 專業版']);
  const teacherPickNv = {}; teacherPickNv[Q_PICK] = 'Kahoot 教師版'; teacherPickNv[Q_PICK_OFFICE] = 'Wordwall 專業版';
  check('parseSubmission_：科任只採用學年／科任的選單', parseSubmission_(nvOf('科任', Q_SUBJECT, '自然', teacherPickNv), softwareList, at(0)).selections,
    ['TeacherSay 專業版', 'Kahoot 教師版']);
  const noneNv = {}; noneNv[Q_PICK] = NONE_CHOICE;
  check('parseSubmission_：不申請其他軟體', parseSubmission_(nvOf('學年', Q_GRADE, '一年級', noneNv), softwareList, at(0)).selections,
    ['TeacherSay 專業版']);
  const legacyNv = {};
  legacyNv[Q_NAME] = ['王']; legacyNv[Q_EMAIL] = ['w@gmail.com']; legacyNv[Q_ROLE] = ['學年']; legacyNv[Q_GRADE] = ['二年級'];
  legacyNv[Q_TEACHERSAY] = ['TeacherSay 專業版']; legacyNv[Q_SOFTWARE] = ['']; legacyNv[Q_AI_PRO] = ['是'];
  check('parseSubmission_：v3.2 回應分頁沒有新題目也能解析', parseSubmission_(legacyNv, softwareList, at(0)).selections,
    ['TeacherSay 專業版', 'Google AI Pro 教育版']);

  // --- 同一帳號另登記其他軟體的提醒 ---
  const held = [
    { email: 'W@gmail.com', software: 'TeacherSay 專業版', status: STATUS.OK },
    { email: 'w@gmail.com', software: 'Kahoot 教師版', status: STATUS.OK },
    { email: 'w@gmail.com', software: 'Padlet 學校版', status: STATUS.CANCEL },
    { email: 'x@gmail.com', software: 'HiTeach 專業版', status: STATUS.OK }
  ];
  check('另登記提醒：已有 Kahoot 再登記 HiTeach', otherSoftwareNotes_(held, { email: 'w@gmail.com' }, ['HiTeach 專業版']),
    { 'HiTeach 專業版': '此帳號另登記了：Kahoot 教師版（每人除 TeacherSay 外限一項，請確認）' });
  check('另登記提醒：取消的與 TeacherSay 不算', otherSoftwareNotes_(held.slice(2), { email: 'w@gmail.com' }, ['TeacherSay 專業版', 'HiTeach 專業版']), {});
  check('另登記提醒：同一筆舊回應勾了兩項', otherSoftwareNotes_([], { email: 'n@gmail.com' }, ['Kahoot 教師版', 'HiTeach 專業版']),
    { 'HiTeach 專業版': '此帳號另登記了：Kahoot 教師版（每人除 TeacherSay 外限一項，請確認）' });

  // --- 學校帳號規則 (表單驗證用) ---
  const schoolRe = new RegExp(schoolAccountPattern_());
  check('學校帳號規則', ['a@mail.lres.tc.edu.tw', 'A@MAIL.LRES.TC.EDU.TW', 'a@gmail.com', 'a@mailxlres.tc.edu.tw', 'a@mail.lres.tc.edu.tw.x.com']
    .map(function (x) { return schoolRe.test(x); }), [true, true, false, false, false]);

  check('groupOf_：從登記表的職務欄還原組別',
    [groupOf_('學年／六年級'), groupOf_('科任／社會'), groupOf_('科任'), groupOf_('行政／輔導室'), groupOf_('')],
    ['六年級', '科任', '科任', '行政', '']);

  // --- 通知信副本與備註 ---
  check('副本：學校帳號與主要帳號不同才加',
    [ccFor_('a@gmail.com', 'a@mail.lres.tc.edu.tw'), ccFor_('A@x.tw', 'a@x.tw'), ccFor_('a@x.tw', ''), ccFor_('a@x.tw', 'oops')],
    ['a@mail.lres.tc.edu.tw', '', '', '']);
  check('withCc_ 沒有副本時不加 cc 欄位', Object.keys(withCc_({ to: 'a@x.tw' }, 'a@x.tw', '')), ['to']);
  check('withCc_ 有學校帳號時加上 cc', withCc_({ to: 'a@gmail.com' }, 'a@gmail.com', 'a@mail.lres.tc.edu.tw').cc, 'a@mail.lres.tc.edu.tw');
  check('備註：Google AI Pro 未填學校帳號', registrationNote_(gradeSub, 'Google AI Pro 教育版'), '未填學校帳號');
  check('備註：其他軟體不需要學校帳號', registrationNote_(gradeSub, 'Kahoot 教師版'), '');
  check('備註：學校帳號非校內網域',
    registrationNote_({ email: 'a@gmail.com', workspace: 'a@gmail.com' }, 'Google AI Pro 教育版'), '學校帳號非 @mail.lres.tc.edu.tw');
  check('備註：校內帳號正常', registrationNote_(subjectSub, 'Google AI Pro 教育版'), '');

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
    reg(4, 'd@x.tw', '行政', 'Google AI Pro 教育版'),
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
  check('剩餘文字：TeacherSay 不限名額', text.indexOf('・TeacherSay 專業版：不限名額（每位申請者必選）') >= 0, true);
  const officeText = buildRemainingText_(computeStats_(kahoot.concat(ai), softwareList), '', ['TeacherSay 專業版', 'Google AI Pro 教育版']);
  check('剩餘文字：行政選單不列 TeacherSay 與 Google AI Pro',
    [officeText.indexOf('TeacherSay 專業版：'), officeText.indexOf('Google AI Pro'), officeText.indexOf('・Kahoot 教師版') >= 0], [-1, -1, true]);
  const unlimited = [];
  for (let i = 0; i < 30; i++) unlimited.push(reg(i, 'u' + i + '@x.tw', '科任', 'TeacherSay 專業版'));
  check('不限名額：30 人全部正取', applyRank(unlimited).every(function (row) { return row.status === STATUS.OK; }), true);
  check('不限名額：統計顯示「不限」',
    quotaText_(computeStats_(unlimited, softwareList).filter(function (st) { return st.name === 'TeacherSay 專業版'; })[0].remaining), '不限');
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

  // --- 重新填表以最後一次為準 ---
  const mine = [
    reg(1, 'R@x.tw', '一年級', 'TeacherSay 專業版', { status: STATUS.OK }),
    reg(1, 'r@x.tw', '一年級', 'Kahoot 教師版', { status: STATUS.OK }),
    reg(1, 'r@x.tw', '一年級', 'Padlet 學校版', { status: STATUS.WAIT }),
    reg(2, 'r@x.tw', '一年級', 'HiTeach 專業版', { status: STATUS.OK, approved: true }),
    reg(2, 'r@x.tw', '一年級', 'Wordwall 專業版', { status: STATUS.ISSUED }),
    reg(2, 'r@x.tw', '一年級', 'Wayground (Quizizz)', { status: STATUS.CANCEL }),
    reg(9, 'r@x.tw', '一年級', 'Google AI Pro 教育版', { status: STATUS.OK }),
    reg(1, 'other@x.tw', '一年級', 'Kahoot 教師版', { status: STATUS.OK })
  ];
  const software = function (list) { return list.map(function (row) { return row.software; }); };
  const replan = planReplacement_(mine, { email: 'r@x.tw', time: at(5), selections: ['TeacherSay 專業版', 'Kahoot 教師版'] });
  check('重新填表：沒再選的取消、已核准只提醒；已發放、已取消、較晚的登記與別人不動',
    [software(replan.cancel), software(replan.flag)], [['Padlet 學校版'], ['HiTeach 專業版']]);
  check('重新填表：沒有 TeacherSay 的舊回應也不會取消 TeacherSay',
    software(planReplacement_(mine, { email: 'r@x.tw', time: at(5), selections: ['Kahoot 教師版'] }).cancel), ['Padlet 學校版']);
  check('重新填表：同一筆回應 (時間相同) 不取消任何登記',
    planReplacement_(mine, { email: 'r@x.tw', time: at(1), selections: ['TeacherSay 專業版'] }).cancel.length, 0);
  check('備註：附加在後面、同樣的文字不重複', [appendNote_('', 'A'), appendNote_('X', 'A'), appendNote_('X；A', 'A')], ['A', 'X；A', 'X；A']);
  const staleNote = '學校帳號非 @mail.lres.tc.edu.tw；此帳號另登記了：Wayground (Quizizz)、Padlet 學校版（每人除 TeacherSay 外限一項，請確認）';
  check('過時的「另登記了」提醒只拿掉那一段', stripOtherSoftwareNote_(staleNote), '學校帳號非 @mail.lres.tc.edu.tw');
  const flaggedRows = [
    reg(1, 's@x.tw', '一年級', 'HiTeach 專業版', { status: STATUS.OK, approved: true }),
    reg(1, 's@x.tw', '一年級', 'Padlet 學校版', { status: STATUS.CANCEL, note: staleNote }),
    reg(1, 's@x.tw', '一年級', 'Kahoot 教師版', { status: STATUS.OK, note: '此帳號另登記了：HiTeach 專業版（每人除 TeacherSay 外限一項，請確認）' })
  ];
  check('仍持有兩項 (已核准的沒取消) 時保留「另登記了」提醒',
    staleOtherSoftwareNotes_(flaggedRows, { email: 's@x.tw', selections: ['TeacherSay 專業版', 'Kahoot 教師版'] }).length, 0);
  check('只剩一項時，有效登記上的「另登記了」提醒視為過時 (已取消的列不動)',
    software(staleOtherSoftwareNotes_(flaggedRows.slice(1), { email: 's@x.tw', selections: ['TeacherSay 專業版', 'Kahoot 教師版'] })),
    ['Kahoot 教師版']);
  const changedHtml = buildReceiptHtml_({ name: '王', workspace: '' }, [{ software: 'Kahoot 教師版', status: STATUS.OK, rank: 1, isNew: false }],
    { canceled: ['Padlet 學校版', 'HiTeach 專業版'], flagged: ['Wordwall 專業版'] });
  check('回條：列出已取消與待確認的軟體',
    [changedHtml.indexOf('<b>已取消</b>：Padlet 學校版、HiTeach 專業版') >= 0, changedHtml.indexOf('確認是否保留：Wordwall 專業版') >= 0], [true, true]);
  check('回條：說明重新填寫以最後一次為準', changedHtml.indexOf(RESUBMIT_RULE_TEXT) >= 0, true);
  check('回條：沒有變更時不出現取消段落',
    buildReceiptHtml_({ name: '王', workspace: '' }, [{ software: 'Kahoot 教師版', status: STATUS.OK, rank: 1, isNew: true }]).indexOf('已取消'), -1);

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
    reg(9, 'p5@x.tw', '行政', 'Google AI Pro 教育版', { approved: true }),
    reg(10, 'p6@x.tw', '科任', 'Padlet 學校版', { approved: true })
  ]);
  // 發放計畫用獨立的清單：Kahoot、Padlet、HiTeach 取序號，Wordwall、Google AI Pro 開通帳號
  const issueList = parseSoftwareRows_([
    ['Kahoot 教師版', 5, MODE.SERIAL, ''], ['Padlet 學校版', 10, MODE.SERIAL, ''], ['HiTeach 專業版', 15, MODE.SERIAL, ''],
    ['Wordwall 專業版', 10, MODE.ACCOUNT, ''], ['Google AI Pro 教育版', 7, MODE.ACCOUNT, LIMIT_PER_GROUP_TEXT]
  ]);
  const plan = planIssue_(pool, issueList, serials);
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
  const p1 = { row: reg(1, 'p1@gmail.com', '一年級', 'HiTeach 專業版', { workspace: 'p1@mail.lres.tc.edu.tw' }) };
  const p1b = { row: reg(2, 'P1@gmail.com', '一年級', 'Kahoot 教師版') };
  const p2 = { row: reg(3, 'p2@gmail.com', '科任', 'Kahoot 教師版') };
  check('寄信額度：副本也算收件人', recipientCount_([p1, p1b, p2]), { people: 2, addresses: 3 });
  check('學校帳號取同一人任一筆登記', workspaceOf_([p1b, p1]), 'p1@mail.lres.tc.edu.tw');
  const aiHtml = buildIssueHtml_('王', 'p1@gmail.com', 'p1@mail.lres.tc.edu.tw', [
    { row: reg(4, 'p1@gmail.com', '一年級', 'Google AI Pro 教育版'), serial: '', mode: MODE.ACCOUNT },
    { row: reg(5, 'p1@gmail.com', '一年級', 'Kahoot 教師版'), serial: '', mode: MODE.ACCOUNT }
  ]);
  check('發放信：Google AI Pro 寫學校帳號，其餘寫 Google 帳號',
    [aiHtml.indexOf('學校帳號（p1@mail.lres.tc.edu.tw）') >= 0, aiHtml.indexOf('Google 帳號（p1@gmail.com）') >= 0], [true, true]);
  check('Email 格式檢查', [isEmail_('a@b.tw'), isEmail_('a@b'), isEmail_('a b@c.tw')], [true, false, false]);

  const failed = results.filter(function (r) { return !r.ok; });
  results.forEach(function (r) { Logger.log((r.ok ? '通過  ' : '失敗  ') + r.name + (r.detail ? '：' + r.detail : '')); });
  Logger.log('合計 ' + results.length + ' 項，通過 ' + (results.length - failed.length) + ' 項，失敗 ' + failed.length + ' 項。');
  if (failed.length > 0) throw new Error('單元測試失敗 ' + failed.length + ' 項，詳見執行紀錄。');
  return results;
}
