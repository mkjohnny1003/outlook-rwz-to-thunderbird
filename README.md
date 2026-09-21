# Outlook RWZ to Thunderbird Filter Add-on

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Thunderbird](https://img.shields.io/badge/Thunderbird-MailExtension-orange.svg)](https://www.thunderbird.net/)
[![Release](https://img.shields.io/github/v/release/mkjohnny1003/outlook-rwz-to-thunderbird?color=green)](https://github.com/mkjohnny1003/outlook-rwz-to-thunderbird/releases)
[![GitHub](https://img.shields.io/badge/GitHub-mkjohnny1003-181717?logo=github)](https://github.com/mkjohnny1003)

[繁體中文 (Traditional Chinese)](#繁體中文說明) | [English (User Guide)](#english-user-guide)

---

<a name="繁體中文說明"></a>
# 繁體中文使用指南

一個專為 **Mozilla Thunderbird** 開發的開源附加元件（MailExtension），可直接解析微軟 **Outlook 規則精靈匯出檔（`.rwz` 二進位格式）**，並轉換匯入為 Thunderbird 的郵件篩選器。

### 🌟 核心特色
- 📁 **不存在之資料夾自動直接建立**：當 Outlook 規則包含「移動/複製到指定資料夾」，若目標帳號中不存在該資料夾，系統將自動調用 Thunderbird API 遞迴建立多層階層資料夾（例如：`專案A/發票`）！
- 📥 **純前端二進位解析**：在瀏覽器環境直接以純 JavaScript 解析 Outlook 98 ~ 2019+ 之二進位 `.rwz` 結構與 MAPI 屬性，不經任何遠端伺服器，資料 100% 安全隱私。
- ⚡ **雙軌套用方式**：
  - **一鍵直接套用**：透過 WebExtension Experiment API 直接存取 Thunderbird 內部 `MailServices.filters`，**無需重啟 Thunderbird 即可立即生效**。
  - **設定檔下載**：亦可下載標準 `msgFilterRules.dat` 純文字檔或 `outlook_rules.json` 備份。
- 🎨 **現代化操作介面**：支援檔案拖曳、即時規則條件與動作預覽、資料夾存在狀態動態標籤（🟢已存在 / 🔵⚡將自動建立）與執行紀錄日誌。

---

## 📌 使用教學：步驟指南

### 第一步：從 Outlook 匯出 `.rwz` 規則檔
1. 開啟微軟 Outlook（傳統桌面版 Outlook 2010 / 2013 / 2016 / 2019 / 2021 / Microsoft 365 傳統版）。
2. 點選左上角 **「檔案」** ➔ **「資訊」** ➔ 點擊 **「管理規則及通知」**。
3. 在「規則及通知」對話方塊上方，點選 **「選項」** 按鈕。
4. 點選 **「匯出規則...」**。
5. 選擇儲存路徑並確認副檔名為 `.rwz`，點擊儲存。

---

### 第二步：安裝附加元件至 Thunderbird
#### 方式 A：使用封裝好的 `.xpi` 檔案（推薦）
1. 至本專案的 [Releases 頁面](https://github.com/mkjohnny1003/outlook-rwz-to-thunderbird/releases/tag/v1.0.0) 下載最新版 **`outlook-rwz-to-thunderbird.xpi`**。
2. 開啟 Mozilla Thunderbird。
3. 點擊右上角功能表 ➔ 選擇 **「附加元件與佈景主題」**（快捷鍵：`Ctrl+Shift+A`，Mac 為 `Cmd+Shift+A`）。
4. 點擊右上角齒輪圖示 ⚙️ ➔ 選擇 **「從檔案安裝附加元件...」**。
5. 選取下載的 `outlook-rwz-to-thunderbird.xpi` 檔案，在權限提示視窗中確認安裝。
6. 安裝後，Thunderbird 工具列將出現轉換器圖示（或自動開啟主分頁）。

#### 方式 B：開發者暫時性載入
1. 在 Thunderbird 網址列輸入 `about:debugging`。
2. 點選左側「This Thunderbird」➔ 點擊「Load Temporary Add-on...」。
3. 選取本專案目錄中的 `manifest.json` 即可完成載入。

---

### 第三步：在附加元件中轉換與匯入
1. **上傳規則檔**：
   - 點擊工具列圖示開啟轉換器頁面。
   - 將第一步匯出的 `.rwz` 檔案拖入「步驟 1」區域，或點擊「選擇檔案」。
   - 系統將自動偵測檔案版本（如 `outlook2019`）並列出規則數量。
2. **選擇目標帳號**：
   - 在「步驟 2」的下拉選單中挑選要套用篩選器的信箱帳號，或選擇「本地資料夾 (Local Folders)」。
   - 確認勾選 **「若目標資料夾在 Thunderbird 不存在時，直接自動建立相對應資料夾」**（預設開啟）。
3. **規則與資料夾預覽**：
   - 在「步驟 3」的表格中檢視每條規則名稱、條件、動作以及目標資料夾狀態：
     - 🟢 **✓ 已存在**：目標帳號已有此資料夾，規則將直接指向它。
     - 🔵 **⚡ 自動建立**：目標帳號目前無此資料夾，匯入時將自動為您建立。
4. **執行套用**：
   - 點擊 **「⚡ 直接匯入至 Thunderbird (免重啟)」**：系統會先自動建立缺少的資料夾，並將篩選規則直接寫入 Thunderbird 核心。
   - 或點擊 **「💾 下載 msgFilterRules.dat 檔案」**：供手動備份或離線設定。

---

## 📋 支援的規則對照表

| Outlook 規則要素 | Outlook RWZ 代碼 | Thunderbird 篩選器語法 | 轉換說明 |
| :--- | :--- | :--- | :--- |
| **來自指定人員/群組** | `0xcb` | `from,contains` | 自動提取 MAPI SMTP 信箱與顯示名稱 |
| **傳送給指定人員/群組** | `0xcc` | `to,contains` | 提取收件者信箱與群組名稱 |
| **主旨包含特定文字** | `0xcd` | `subject,contains` | 支援多組關鍵字 OR 條件 |
| **內文包含特定文字** | `0xce` | `body,contains` | 支援多組內文關鍵字 |
| **主旨或內文包含** | `0xcf` | `subject/body,contains` | 自動轉換為複合搜尋條件 |
| **含有附件** | `0xde` | `status,is,hasAttachment` | 篩選有附件之郵件 |
| **重要性標記** | `0xd2` | `priority,is` | 高 / 一般 / 低重要性 |
| **移動至指定資料夾** | `0x12c` | `Move to folder` | **若資料夾不存在時直接自動建立** |
| **複製至指定資料夾** | `0x139` | `Copy to folder` | **若資料夾不存在時直接自動建立** |
| **刪除郵件** | `0x12d` / `0x14a` | `Delete` | 移至垃圾桶或刪除 |
| **標示為已讀** | `0x14c` | `Mark read` | 自動設為已讀狀態 |
| **加上星號/旗標** | `0x150` / `0x151` | `Mark flagged` | 加上星號標籤 |
| **停止處理其他規則** | `0x142` | `Stop execution` | 中止後續規則執行 |

---

<br/>
<br/>

<a name="english-user-guide"></a>
# English User Guide

An open-source **Mozilla Thunderbird** extension (MailExtension) designed to parse binary **Microsoft Outlook Rules Wizard (`.rwz`)** files and convert them seamlessly into native Thunderbird Message Filters.

### 🌟 Key Highlights
- 📁 **Automatic Folder Creation**: If an Outlook rule directs messages to a folder that does not exist in Thunderbird, the extension **automatically and recursively creates the missing folder hierarchy** (e.g., `Finance/2026/Invoices`)!
- 📥 **Pure Client-Side Binary Parsing**: Parses binary `.rwz` files and MAPI structures from Outlook 98 through 2019+ natively in JavaScript. No remote servers involved — 100% private and secure.
- ⚡ **Dual Import Options**:
  - **Direct Import**: Uses the Thunderbird WebExtension Experiment API to inject filters into `MailServices.filters` directly — **effective immediately without restarting Thunderbird**.
  - **File Download**: Export standard `msgFilterRules.dat` plain-text files or structured `outlook_rules.json` files for backup.
- 🎨 **Modern User Interface**: Drag-and-drop file upload, live folder status badges (🟢 Exists / 🔵⚡ Will Auto-Create), real-time activity log console, and light/dark theme adaptation.

---

## 📌 Step-by-Step Instructions

### Step 1: Export the `.rwz` Rules File from Outlook
1. Open Microsoft Outlook (Classic Desktop Outlook 2010 / 2013 / 2016 / 2019 / 2021 / Microsoft 365 Classic).
2. Go to **File** ➔ **Info** ➔ click **Manage Rules & Alerts**.
3. In the "Rules and Alerts" dialog, click the **Options** button in the toolbar.
4. Click **Export Rules...**.
5. Choose a destination folder, ensure the file format is `.rwz`, and click **Save**.

---

### Step 2: Install the Extension in Thunderbird
#### Option A: Install from `.xpi` File (Recommended)
1. Download the latest **`outlook-rwz-to-thunderbird.xpi`** from the [Releases Page](https://github.com/mkjohnny1003/outlook-rwz-to-thunderbird/releases/tag/v1.0.0).
2. Open Mozilla Thunderbird.
3. Open **Add-ons and Themes** (shortcut: `Ctrl+Shift+A` on Windows/Linux, `Cmd+Shift+A` on macOS).
4. Click the gear icon ⚙️ at the top right ➔ select **"Install Add-on From File..."**.
5. Select the downloaded `.xpi` file and confirm permissions when prompted.
6. Once installed, a new icon will appear in the Thunderbird toolbar. Click it to open the converter tab.

#### Option B: Load as Temporary Extension (Developers)
1. Navigate to `about:debugging` in Thunderbird.
2. Select **This Thunderbird** on the left menu.
3. Click **Load Temporary Add-on...**.
4. Select `manifest.json` inside this repository.

---

### Step 3: Convert and Apply Rules
1. **Upload `.rwz` File**:
   - Drag and drop your `.rwz` file into the upload zone or click "Browse File".
   - The tool detects the Outlook version (e.g., `outlook2019`) and displays the number of rules found.
2. **Select Target Account**:
   - Select the target mail account or "Local Folders" from the dropdown.
   - Keep **"Auto-create missing folders in Thunderbird"** checked (enabled by default).
3. **Review Preview & Folder Badges**:
   - Check the parsed rules list:
     - 🟢 **✓ Exists**: The folder already exists in the selected account.
     - 🔵 **⚡ Auto-Create**: The folder is currently missing and will be created automatically upon import.
4. **Apply / Export**:
   - Click **"⚡ Direct Import to Thunderbird"** to create missing folders and inject filters directly into Thunderbird with immediate effect.
   - Or click **"💾 Download msgFilterRules.dat"** to save the filter configuration file.

---

## 🧪 Automated Testing & Build

```bash
# Run all unit tests
npm test
# or with Node.js directly
node test/run-all.js

# Build and package into .xpi
bash build.sh
```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

Developed and maintained by [@mkjohnny1003](https://github.com/mkjohnny1003).
