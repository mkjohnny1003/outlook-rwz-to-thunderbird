/**
 * Outlook RWZ to Thunderbird Filter Add-on - UI Application Logic
 */

/* global messenger, RwzParser, FolderManager, FilterGenerator */

(function () {
  'use strict';

  // State
  let accounts = [];
  let selectedAccount = null;
  let parsedRwz = null;
  let folderMap = new Map();
  let folderManager = null;
  let isExperimentAvailable = false;

  // DOM Elements
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const browseBtn = document.getElementById('browseBtn');
  const fileSummary = document.getElementById('fileSummary');
  const fileNameEl = document.getElementById('fileName');
  const fileSizeEl = document.getElementById('fileSize');
  const fileVersionEl = document.getElementById('fileVersion');
  const ruleCountEl = document.getElementById('ruleCount');

  const accountSelect = document.getElementById('accountSelect');
  const optAutoCreateFolders = document.getElementById('optAutoCreateFolders');

  const folderStatusSummary = document.getElementById('folderStatusSummary');
  const rulesTableBody = document.getElementById('rulesTableBody');

  const directImportBtn = document.getElementById('directImportBtn');
  const downloadDatBtn = document.getElementById('downloadDatBtn');
  const exportJsonBtn = document.getElementById('exportJsonBtn');

  const logConsole = document.getElementById('logConsole');
  const clearLogBtn = document.getElementById('clearLogBtn');

  // Logger helper
  function log(msg, type = 'info') {
    const time = new Date().toLocaleTimeString();
    const prefix = type === 'error' ? '❌ ' : type === 'success' ? '✅ ' : type === 'warn' ? '⚠️ ' : 'ℹ️ ';
    logConsole.textContent += `[${time}] ${prefix}${msg}\n`;
    logConsole.scrollTop = logConsole.scrollHeight;
  }

  function clearLog() {
    logConsole.textContent = '';
  }

  // Initialize
  async function init() {
    folderManager = new FolderManager(typeof messenger !== 'undefined' ? messenger : null);

    // Check Experiment API availability
    if (typeof messenger !== 'undefined' && messenger.rwzFilters && messenger.rwzFilters.isAvailable) {
      try {
        isExperimentAvailable = await messenger.rwzFilters.isAvailable();
      } catch (_) {
        isExperimentAvailable = false;
      }
    }

    await loadAccounts();
    setupEventListeners();
    log('附加元件已就緒，請上傳 Outlook .rwz 規則檔。');
  }

  // Load configured accounts in Thunderbird
  async function loadAccounts() {
    accountSelect.innerHTML = '';

    if (typeof messenger !== 'undefined' && messenger.accounts && messenger.accounts.list) {
      try {
        accounts = await messenger.accounts.list();
      } catch (err) {
        log(`載入帳號清單失敗: ${err.message}`, 'warn');
        accounts = [];
      }
    }

    // Fallback for preview or browser environment
    if (!accounts || accounts.length === 0) {
      accounts = [
        {
          id: 'account_local',
          name: '本地資料夾 (Local Folders)',
          type: 'none',
          folders: [
            { id: 'f_inbox', name: '收件匣', path: '/INBOX', subFolders: [] },
            { id: 'f_trash', name: '垃圾桶', path: '/Trash', subFolders: [] }
          ]
        },
        {
          id: 'account_imap',
          name: '主要信箱 (IMAP)',
          type: 'imap',
          folders: [
            { id: 'f_imap_inbox', name: 'INBOX', path: '/INBOX', subFolders: [] }
          ]
        }
      ];
      log('目前於展示/預覽模式運作，已載入模擬帳號清單。');
    }

    accounts.forEach((acc, idx) => {
      const opt = document.createElement('option');
      opt.value = acc.id;
      const typeStr = acc.type ? ` (${acc.type.toUpperCase()})` : '';
      opt.textContent = `${acc.name}${typeStr}`;
      accountSelect.appendChild(opt);
      if (idx === 0) selectedAccount = acc;
    });

    if (accounts.length > 0) {
      selectedAccount = accounts[0];
    }
  }

  // Event Listeners
  function setupEventListeners() {
    // Dropzone
    browseBtn.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('click', (e) => {
      if (e.target !== browseBtn) fileInput.click();
    });

    dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
      dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFile(e.target.files[0]);
      }
    });

    // Account change
    accountSelect.addEventListener('change', (e) => {
      selectedAccount = accounts.find(a => a.id === e.target.value) || null;
      log(`已切換目標帳號為: ${selectedAccount ? selectedAccount.name : '無'}`);
      updateRulesTable();
    });

    // Auto create checkbox
    optAutoCreateFolders.addEventListener('change', () => {
      updateRulesTable();
    });

    // Action buttons
    directImportBtn.addEventListener('click', handleDirectImport);
    downloadDatBtn.addEventListener('click', handleDownloadDat);
    exportJsonBtn.addEventListener('click', handleExportJson);
    clearLogBtn.addEventListener('click', clearLog);
  }

  // Handle uploaded file
  function handleFile(file) {
    if (!file.name.toLowerCase().endsWith('.rwz')) {
      log(`警告: 檔案「${file.name}」副檔名不是 .rwz，仍嘗試進行解析...`, 'warn');
    }

    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = `${(file.size / 1024).toFixed(1)} KB`;

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const buffer = e.target.result;
        parsedRwz = RwzParser.parse(buffer);

        fileVersionEl.textContent = parsedRwz.version;
        ruleCountEl.textContent = `${parsedRwz.rules.length} 條規則`;
        fileSummary.style.display = 'grid';

        log(`成功解析 RWZ 檔案！偵測到版本: ${parsedRwz.version}，共 ${parsedRwz.rules.length} 條規則。`, 'success');

        directImportBtn.disabled = parsedRwz.rules.length === 0;
        downloadDatBtn.disabled = parsedRwz.rules.length === 0;
        exportJsonBtn.disabled = parsedRwz.rules.length === 0;

        updateRulesTable();
      } catch (err) {
        log(`解析失敗: ${err.message}`, 'error');
        alert(`解析 .rwz 檔案失敗：\n${err.message}`);
      }
    };

    reader.onerror = function () {
      log('讀取檔案時發生錯誤', 'error');
    };

    reader.readAsArrayBuffer(file);
  }

  // Update preview table & folder status
  function updateRulesTable() {
    if (!parsedRwz || !parsedRwz.rules || parsedRwz.rules.length === 0) {
      rulesTableBody.innerHTML = '<tr><td colspan="5" class="empty-state">尚未載入 .rwz 規則檔</td></tr>';
      folderStatusSummary.innerHTML = '<span class="status-indicator">尚無資料</span>';
      return;
    }

    const { existing, missing, all } = FolderManager.checkRuleFolders(parsedRwz.rules, selectedAccount);

    if (all.length === 0) {
      folderStatusSummary.innerHTML = '<span class="badge badge-info">規則中無指定移動/複製資料夾</span>';
    } else {
      const autoCreate = optAutoCreateFolders.checked;
      folderStatusSummary.innerHTML = `
        <span class="badge badge-success">已存在: ${existing.length}</span>
        <span class="badge ${autoCreate ? 'badge-folder-create' : 'badge-warning'}">
          ${autoCreate ? '⚡ 將自動建立' : '缺少'}: ${missing.length}
        </span>
      `;
    }

    rulesTableBody.innerHTML = '';

    parsedRwz.rules.forEach((rule, idx) => {
      const tr = document.createElement('tr');

      // Status (enabled toggle)
      const tdStatus = document.createElement('td');
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = rule.enabled;
      chk.title = rule.enabled ? '啟用中' : '已停用';
      chk.addEventListener('change', () => {
        rule.enabled = chk.checked;
      });
      tdStatus.appendChild(chk);
      tr.appendChild(tdStatus);

      // Rule Name
      const tdName = document.createElement('td');
      tdName.innerHTML = `<strong>${escapeHtml(rule.name)}</strong>`;
      tr.appendChild(tdName);

      // Conditions
      const tdCond = document.createElement('td');
      if (rule.conditions && rule.conditions.length > 0) {
        const tagList = document.createElement('div');
        tagList.className = 'tag-list';
        rule.conditions.forEach(c => {
          const tag = document.createElement('span');
          tag.className = 'tag';
          tag.textContent = c.desc || `${c.type} ${c.op || ''} ${c.values ? c.values.join(', ') : ''}`;
          tagList.appendChild(tag);
        });
        tdCond.appendChild(tagList);
      } else {
        tdCond.innerHTML = '<span class="tag">套用至所有郵件</span>';
      }
      tr.appendChild(tdCond);

      // Actions
      const tdAct = document.createElement('td');
      if (rule.actions && rule.actions.length > 0) {
        const tagList = document.createElement('div');
        tagList.className = 'tag-list';
        rule.actions.forEach(a => {
          const tag = document.createElement('span');
          tag.className = 'tag';
          tag.textContent = a.desc || a.type;
          tagList.appendChild(tag);
        });
        tdAct.appendChild(tagList);
      } else {
        tdAct.innerHTML = '<span class="tag">無特定動作</span>';
      }
      tr.appendChild(tdAct);

      // Target Folder Status
      const tdFolder = document.createElement('td');
      if (rule.targetFolders && rule.targetFolders.length > 0) {
        const tagList = document.createElement('div');
        tagList.className = 'tag-list';
        rule.targetFolders.forEach(folderName => {
          const isExist = existing.includes(folderName);
          const badge = document.createElement('span');
          if (isExist) {
            badge.className = 'badge badge-folder-exist';
            badge.textContent = `✓ 已存在: ${folderName}`;
          } else if (optAutoCreateFolders.checked) {
            badge.className = 'badge badge-folder-create';
            badge.textContent = `⚡ 自動建立: ${folderName}`;
          } else {
            badge.className = 'badge badge-warning';
            badge.textContent = `⚠️ 不存在: ${folderName}`;
          }
          tagList.appendChild(badge);
        });
        tdFolder.appendChild(tagList);
      } else {
        tdFolder.innerHTML = '<span style="color: var(--text-muted); font-size: 12px;">-</span>';
      }
      tr.appendChild(tdFolder);

      rulesTableBody.appendChild(tr);
    });
  }

  // Ensure folders exist prior to rule creation
  async function prepareFolders() {
    folderMap = new Map();
    if (!parsedRwz || !selectedAccount) return folderMap;

    const allTargetFolders = new Set();
    for (const rule of parsedRwz.rules) {
      if (rule.targetFolders) {
        rule.targetFolders.forEach(f => allTargetFolders.add(f));
      }
    }

    if (allTargetFolders.size === 0) return folderMap;

    if (optAutoCreateFolders.checked) {
      log('開始執行資料夾檢查與自動建立程序...');
      folderMap = await folderManager.ensureFoldersExist(
        Array.from(allTargetFolders),
        selectedAccount,
        (msg) => log(msg)
      );
      log('資料夾準備程序完成！', 'success');
    } else {
      // Just resolve existing ones
      for (const tf of allTargetFolders) {
        const existing = FolderManager.findFolderInAccount(selectedAccount, tf);
        folderMap.set(tf, {
          folder: existing,
          uri: folderManager.constructFolderUri(existing, selectedAccount),
          created: false
        });
      }
    }

    return folderMap;
  }

  // Handle Direct Import
  async function handleDirectImport() {
    if (!parsedRwz || !selectedAccount) return;

    directImportBtn.disabled = true;
    log(`開始匯入規則至帳號「${selectedAccount.name}」...`);

    try {
      // 1. Ensure required folders exist
      const resolvedFolders = await prepareFolders();

      // 2. Direct injection via Experiment API
      if (typeof messenger !== 'undefined' && messenger.rwzFilters && messenger.rwzFilters.importRules) {
        const payload = FilterGenerator.toExperimentPayload(parsedRwz.rules, resolvedFolders);
        log(`正在透過 Thunderbird XPCOM 核心服務寫入 ${payload.length} 條規則...`);
        const result = await messenger.rwzFilters.importRules(selectedAccount.id, payload);
        log(`🎉 匯入成功！共建立並套用 ${result.totalImported} 條篩選器，已直接生效！`, 'success');
        alert(`成功匯入 ${result.totalImported} 條規則至「${selectedAccount.name}」！\n您可至 Thunderbird 的「工具」>「郵件篩選器」中檢視。`);
      } else {
        // Fallback when Experiment API not accessible in this context
        log('未偵測到 Experiment 權限，已為您自動建立所需資料夾，並準備產生 msgFilterRules.dat 檔案。', 'warn');
        handleDownloadDat();
      }
    } catch (err) {
      log(`匯入過程發生錯誤: ${err.message}`, 'error');
      alert(`匯入失敗：\n${err.message}`);
    } finally {
      directImportBtn.disabled = false;
      updateRulesTable();
    }
  }

  // Handle Download msgFilterRules.dat
  async function handleDownloadDat() {
    if (!parsedRwz || !selectedAccount) return;

    try {
      const resolvedFolders = await prepareFolders();
      const content = FilterGenerator.generateMsgFilterRulesDat(parsedRwz.rules, resolvedFolders, {
        logging: true
      });

      const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'msgFilterRules.dat';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      log('已下載 msgFilterRules.dat 檔案。', 'success');
      log(`提示：若手動置放，請將此檔案放入 Thunderbird 個人設定檔中帳號目錄（如 Mail/Local Folders 或 ImapMail/...）。`);
    } catch (err) {
      log(`產生設定檔失敗: ${err.message}`, 'error');
    }
  }

  // Handle Export JSON
  function handleExportJson() {
    if (!parsedRwz) return;
    const jsonStr = JSON.stringify(parsedRwz, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'outlook_rules.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    log('已匯出 outlook_rules.json 結構檔。', 'success');
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
