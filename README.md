# 贈品記數

每年送畀客人嘅贈品：記低產品數量、邊個訂咗幾多、幾時攞、賣出同額外送出。
網頁放喺 Netlify，數據儲喺你自己嘅 Google Sheet。

```
site/            ← Netlify 網站（index.html + Excel 讀取工具）
apps-script/     ← 貼入 Google Sheet「擴充功能 → Apps Script」嘅程式碼
netlify.toml     ← 話畀 Netlify 知網站喺 site 資料夾
```

**呢個 repo 入面冇任何客人資料。** 客人名同數量只會儲喺你嘅 Google Sheet，
匯入 Excel 嘅時候喺你部機讀取，唔會上傳去 GitHub 或 Netlify。

## 第一次設定

1. **Google Sheet**：開一個新試算表 →「擴充功能」→「Apps Script」→ 刪晒原本內容，
   貼入 `apps-script/Code.gs` → 改第 15 行 `PASSCODE` 做你自己嘅密碼 → 儲存。
2. 函式選單揀 `setup` →「執行」→ 授權（「進階」→「前往…」→「允許」）。
3. 「部署」→「新增部署作業」→ 類型「網頁應用程式」→ 執行身分「我」、
   存取權「所有人」→ 部署 → 複製網址（`https://script.google.com/macros/s/…/exec`）。
4. **Netlify**：Add new project → Import from GitHub → 揀呢個 repo。
   設定會自動由 `netlify.toml` 讀取，唔使改，直接 Deploy。
5. 打開網站 →「設定」貼上 Apps Script 網址同密碼 →「連接並儲存」→
   「由 Excel 匯入」揀你個記數 Excel。

## 之後更新

- **網站**：改 `site/` 入面嘅檔案再 push，Netlify 會自動更新。
- **Apps Script**：貼新版 `Code.gs` → 儲存 →「部署」→「管理部署作業」→ 鉛筆 →
  版本揀「新版本」→ 部署（網址唔變）。

## Excel 格式

每年一張工作表，用年份做名（例如 `2026`）：第 1 行由 D 欄開始係產品（`C08a 大福字掛曆`），
第 2 行係總數，C 欄係人名，A 欄係攞貨日期或備註，`sell out` 嗰行當賣出。
收埋咗嘅行同欄（嗰年冇份）會跳過。
