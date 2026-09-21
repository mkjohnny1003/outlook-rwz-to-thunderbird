# Outlook RWZ to Thunderbird Filter Add-on

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Thunderbird](https://img.shields.io/badge/Thunderbird-MailExtension-orange.svg)](https://www.thunderbird.net/)
[![GitHub](https://img.shields.io/badge/GitHub-mkjohnny1003-181717?logo=github)](https://github.com/mkjohnny1003)

一個為 **Mozilla Thunderbird** 開發的開源附加元件（MailExtension），能夠直接解析微軟 **Outlook 規則精靈檔案（`.rwz` 二進位格式）**，並轉換匯入為 Thunderbird 的郵件篩選器。

特別針對跨用戶端遷移痛點，本工具具備**「若 RWZ 規則內容所參照的資料夾在 Thunderbird 中不存在時，直接自動建立相對應的資料夾」**之功能！

---

## ✨ 核心特色

- 📥 **純前端二進位解析**：以純 JavaScript 原生解析 Outlook 98 ~ 2019+ 之 `.rwz` 二進位封包與 MAPI 屬性，不依賴外部伺服器或原生命令列工具，安全且零隱私疑慮。
- 📁 **資料夾自動建立**：
  - 當 Outlook 規則包含「移動至資料夾」或「複製至資料夾」動作時，系統會自動掃描目標帳號之資料夾階層。
  - **若資料夾不存在，系統將調用 Thunderbird API 依層級自動建立該資料夾**（例如：`財務/2026/發票` 逐層建立）！
- ⚡ **雙軌套用方式**：
  - **一鍵直接套用**：利用 WebExtension Experiment API 直接存取 Thunderbird 內部 `MailServices.filters`，免重啟即可立即在「工具」>「郵件篩選器」中生效。
  - **檔案下載匯出**：亦可一鍵產出並下載標準之 `msgFilterRules.dat` 純文字設定檔或 `outlook_rules.json` 備份檔。
- 🎨 **現代化操作介面**：支援拖曳上傳、即時規則與資料夾狀態預覽（綠色「已存在」/ 藍色「⚡ 自動建立」標籤）、即時操作日誌與深淺色主題。

---

## 📸 介面預覽與使用流程

```
┌───────────────────────────────────────────────────────────┐
│ [1] 上傳 .rwz 規則檔 ➔ [2] 選擇帳號 ➔ [3] 預覽與自動建立 ➔ [4] 一鍵生效  │
└───────────────────────────────────────────────────────────┘
```

1. **上傳規則檔**：將由 Outlook 匯出的 `.rwz` 檔案拖入視窗中。
2. **選擇目標帳號**：從下拉選單挑選要匯入的信箱帳號或「本地資料夾 (Local Folders)」。
3. **確認自動建立設定**：預設勾選「若目標資料夾不存在時直接自動建立」。
4. **規則預覽**：檢視解析出的寄件者、主旨、附件條件與目標資料夾狀態。
5. **套用生效**：點擊「⚡ 直接匯入至 Thunderbird」立即完成遷移！

---

## 📋 支援的規則對應表

| Outlook 規則要素 | Outlook RWZ 代碼 | Thunderbird 篩選器對應 | 說明 |
| :--- | :--- | :--- | :--- |
| **來自指定人員/群組** | `0xcb` | `from,contains` | 自動提取 MAPI SMTP 信箱與顯示名稱 |
| **傳送給指定人員/群組** | `0xcc` | `to,contains` | 提取收件者信箱與群組名稱 |
| **主旨包含特定文字** | `0xcd` | `subject,contains` | 支援多組關鍵字 OR 條件 |
| **內文包含特定文字** | `0xce` | `body,contains` | 支援多組內文關鍵字 |
| **主旨或內文包含** | `0xcf` | `subject/body,contains` | 自動轉換為 OR 複合條件 |
| **含有附件** | `0xde` | `status,is,hasAttachment` | 篩選有附件之郵件 |
| **重要性標記** | `0xd2` | `priority,is` | 高 / 一般 / 低重要性 |
| **移動至指定資料夾** | `0x12c` | `Move to folder` | **若不存在自動建立階層資料夾** |
| **複製至指定資料夾** | `0x139` | `Copy to folder` | **若不存在自動建立階層資料夾** |
| **刪除郵件** | `0x12d` / `0x14a` | `Delete` | 移至垃圾桶或刪除 |
| **標示為已讀** | `0x14c` | `Mark read` | 自動設為已讀狀態 |
| **加上星號/旗標** | `0x150` / `0x151` | `Mark flagged` | 加上星號標籤 |
| **停止處理其他規則** | `0x142` | `Stop execution` | 中止後續規則執行 |

---

## 🛠️ 安裝方式

### 方法 A：使用封裝好的 `.xpi` 檔案（推薦）
1. 執行專案目錄下的打包腳本（或至 Releases 下載最新版 `.xpi`）：
   ```bash
   bash build.sh
   ```
   產生的檔案位於 `dist/outlook-rwz-to-thunderbird.xpi`。
2. 開啟 Mozilla Thunderbird。
3. 點選右上角功能表 ➔ **「附加元件與佈景主題」**（或快捷鍵 `Ctrl+Shift+A` / `Cmd+Shift+A`）。
4. 點擊右上角的**齒輪圖示 ⚙️**，選擇**「從檔案安裝附加元件...」**。
5. 選取 `outlook-rwz-to-thunderbird.xpi`，確認權限後即可完成安裝。
6. 安裝完成後，工具列將出現擴充元件圖示，點擊即可開啟轉換器頁面。

### 方法 B：開發者除錯模式載入
1. 在 Thunderbird 網址列或分頁中開啟 `about:debugging`。
2. 點選左側「This Thunderbird」（此 Thunderbird）。
3. 點擊「Load Temporary Add-on...」（載入暫時性附加元件）。
4. 選取本專案目錄下的 `manifest.json` 即可載入。

---

## 🧪 執行自動化測試

本專案包含完整的單元測試套件，涵蓋二進位解析、資料夾比對與建立、以及篩選器語法生成：

```bash
# 執行所有測試
npm test
# 或直接使用 Node.js
node test/run-all.js
```

---

## 🚀 開源至 GitHub 指引

若您要將此專案推送到您的 GitHub (`https://github.com/mkjohnny1003`)：

1. **在 GitHub 上建立新儲存庫**：
   前往 [GitHub New Repository](https://github.com/new)，建立名為 `outlook-rwz-to-thunderbird` 的公開儲存庫（不需要勾選 Initialize with README）。

2. **在本地端執行 Git 初始化與推送**：
   ```bash
   cd /Users/mkjohnny/.gemini/antigravity/scratch/outlook-rwz-to-thunderbird-addon

   # 初始化 Git 儲存庫
   git init

   # 加入所有檔案並提交
   git add .
   git commit -m "feat: Initial commit of Outlook RWZ to Thunderbird Filter Add-on"

   # 設定 main 分支與遠端倉庫
   git branch -M main
   git remote add origin https://github.com/mkjohnny1003/outlook-rwz-to-thunderbird.git

   # 推送至 GitHub
   git push -u origin main
   ```

---

## 📄 授權條款 (License)

本專案採用 [MIT 授權條款](LICENSE) 開源釋出。歡迎自由 Fork、提交 Pull Request 或回報 Issue！

作者：[@mkjohnny1003](https://github.com/mkjohnny1003)
