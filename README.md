# gas-license-enrollment

教學軟體授權登記與分配系統：Google Apps Script 綁定在表單回應試算表上。
老師在新建立的申請表單送出後只做登記，系統依登記先後排定正取／備取並在表單上顯示剩餘名額；
管理者在試算表勾選核准後，才從「序號列表」發放序號並寄信通知。

功能、分頁欄位與管理者操作見 [部署說明.md](部署說明.md)。

## 檔案

| 檔案 | 用途 |
|---|---|
| `code.gs` | 設定、表單送出登記、正取／備取排序、統計 |
| `allocate.gs` | 「授權管理」選單、發放、通知、補登 |
| `setupForm.gs` | 初始化分頁、建立／就地更新申請表單（職務跳題、TeacherSay 必選、Google AI Pro 分支）、觸發器 |
| `confirm.gs`、`index.html` | Wordwall 帳號確認網頁 |
| `tests_enrollment.gs` | Apps Script 內的離線單元測試 `runEnrollmentTests()` |
| `test/run-tests.js` | 以 Node 模擬 Apps Script 執行完整流程 |
| `tools/deploy.sh` | obsv01 上的部署腳本 |

## 開發與部署流程

```
修改程式 → node test/run-tests.js → git commit / push 到 GitHub
obsv01:  tools/deploy.sh          拉取、跑測試、列出將上傳的檔案
obsv01:  tools/deploy.sh --push   確認後上傳到 Apps Script
```

- clasp 不裝在主機上，透過 `/opt/stacks/netmon/tools/clasp.sh` 在一次性容器內執行，登入憑證在 `~/.clasp-home/`。
- `.clasp.json`（含 Apps Script 專案 ID）不進版控，請由 `.clasp.json.example` 複製後填入。
- `clasp push` 會讓線上專案與本資料夾一致；線上編輯器直接修改的內容要先 `clasp pull` 回來提交，否則會被覆蓋。
- `.claspignore` 與 `skipSubdirectories` 確保只上傳根目錄的 `.gs`、`.html`、`appsscript.json`。
