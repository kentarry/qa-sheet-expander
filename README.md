# 品檢檢驗單展表工具 v5

## 功能
- 自動連接 Google Drive 讀取歷史品檢檔案（透過 Apps Script，不需 API Key）
- 自動配對「展表前/展表後」檔案組合
- 上傳待展表 → 自動匹配最像的歷史案件
- AI（Gemini）比對展前→展後差異，精確複製展表模式
- JSZip 層級修改 XLSX，100% 保留字體、顏色、邊框、合併等格式
- 支援多模型切換（2.0 Flash / 2.5 Flash / 2.5 Pro）+ 429 自動重試

## 在 Antigravity 中使用

### Step 1：開啟專案
File → Open Folder → 選擇 `qa-sheet-expander-v5`

### Step 2：安裝 + 啟動
```bash
npm install
npm run dev
```

### Step 3：設定 Gemini Key
介面中點「🔑 設定 Gemini Key」→ 有跳轉按鈕引導取得

## Drive 檔案命名規則
同一案件的兩個版本：
```
w195320_消費活動_金馬開運_展表前.xls
w195320_消費活動_金馬開運_展表後.xlsx
```
系統會自動配對含「展表前」和「展表後」/「品檢」的檔案。

## 專案結構
```
qa-sheet-expander-v5/
├── package.json
├── vite.config.js
├── index.html
├── apps_script_v2.gs    ← Google Apps Script 程式碼
├── README.md
└── src/
    ├── main.jsx          # React 入口
    ├── App.jsx           # 主 UI 元件
    ├── config.js         # 設定（Apps Script URL、Drive ID、模型）
    ├── drive.js          # Drive API（透過 Apps Script）
    ├── gemini.js         # Gemini AI（含 429 重試）
    ├── excel.js          # Excel 解析、摘要、差異比對、JSON 修復
    └── xlBuild.js        # JSZip 層級 XLSX 修改（保留格式）+ 配對邏輯
```

## 技術棧
- React 18 + Vite
- @google/generative-ai（Gemini SDK）
- SheetJS (xlsx) — 讀取解析
- JSZip — ZIP 層級修改 XLSX（保留格式）
- Google Apps Script — 讀取公開 Drive 資料夾
