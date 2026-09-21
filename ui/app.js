/**
 * Outlook RWZ to Thunderbird Filter Add-on - UI Application Logic
 */

/* global messenger, RwzParser, FolderManager, FilterGenerator */

(function () {
  'use strict';

  // Helper to safely resolve library dependencies from global/window scope
  function resolveDeps() {
    return {
      RwzParser: (typeof globalThis !== 'undefined' && globalThis.RwzParser) ||
                 (typeof window !== 'undefined' && window.RwzParser) || null,
      FolderManager: (typeof globalThis !== 'undefined' && globalThis.FolderManager) ||
                     (typeof window !== 'undefined' && window.FolderManager) || null,
      FilterGenerator: (typeof globalThis !== 'undefined' && globalThis.FilterGenerator) ||
                       (typeof window !== 'undefined' && window.FilterGenerator) || null
    };
  }

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
  const exportAccountRulesBtn = document.getElementById('exportAccountRulesBtn');
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
    const { FolderManager } = resolveDeps();
    if (!FolderManager) {
      log('找不到 FolderManager 模組，請重新整理頁面。', 'error');
      return;
    }
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
    if (exportAccountRulesBtn) {
      exportAccountRulesBtn.addEventListener('click', handleExportAccountRules);
    }
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
        const { RwzParser } = resolveDeps();
        if (!RwzParser) {
          throw new Error('找不到 RwzParser 核心模組，請重新整理頁面。');
        }
        parsedRwz = RwzParser.parse(buffer);

        const verDisplay = parsedRwz.versionLabel || parsedRwz.version;
        fileVersionEl.textContent = verDisplay;
        ruleCountEl.textContent = `${parsedRwz.rules.length} 條規則`;
        fileSummary.style.display = 'grid';

        log(`成功解析 RWZ 檔案！偵測到版本: ${verDisplay}，共 ${parsedRwz.rules.length} 條規則。`, 'success');

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
      rulesTableBody.textContent = '';
      const emptyTr = document.createElement('tr');
      const emptyTd = document.createElement('td');
      emptyTd.colSpan = 5;
      emptyTd.className = 'empty-state';
      emptyTd.textContent = '尚未載入 .rwz 規則檔';
      emptyTr.appendChild(emptyTd);
      rulesTableBody.appendChild(emptyTr);

      folderStatusSummary.textContent = '';
      const emptyStatus = document.createElement('span');
      emptyStatus.className = 'status-indicator';
      emptyStatus.textContent = '尚無資料';
      folderStatusSummary.appendChild(emptyStatus);
      return;
    }

    const { FolderManager } = resolveDeps();
    const { existing, missing, all } = FolderManager
      ? FolderManager.checkRuleFolders(parsedRwz.rules, selectedAccount)
      : { existing: [], missing: [], all: [] };

    folderStatusSummary.textContent = '';
    if (all.length === 0) {
      const badge = document.createElement('span');
      badge.className = 'badge badge-info';
      badge.textContent = '規則中無指定移動/複製資料夾';
      folderStatusSummary.appendChild(badge);
    } else {
      const autoCreate = optAutoCreateFolders.checked;
      const bExist = document.createElement('span');
      bExist.className = 'badge badge-success';
      bExist.textContent = `已存在: ${existing.length}`;
      folderStatusSummary.appendChild(bExist);

      const bMiss = document.createElement('span');
      bMiss.className = `badge ${autoCreate ? 'badge-folder-create' : 'badge-warning'}`;
      bMiss.textContent = `${autoCreate ? '⚡ 將自動建立' : '缺少'}: ${missing.length}`;
      folderStatusSummary.appendChild(bMiss);
    }

    rulesTableBody.textContent = '';

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
      const strongName = document.createElement('strong');
      strongName.textContent = rule.name || '';
      tdName.appendChild(strongName);
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
        const tag = document.createElement('span');
        tag.className = 'tag';
        tag.textContent = '套用至所有郵件';
        tdCond.appendChild(tag);
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
        const tag = document.createElement('span');
        tag.className = 'tag';
        tag.textContent = '無特定動作';
        tdAct.appendChild(tag);
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
        const dash = document.createElement('span');
        dash.style.color = 'var(--text-muted)';
        dash.style.fontSize = '12px';
        dash.textContent = '-';
        tdFolder.appendChild(dash);
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
      const { FolderManager } = resolveDeps();
      for (const tf of allTargetFolders) {
        const existing = FolderManager ? FolderManager.findFolderInAccount(selectedAccount, tf) : null;
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

      // 2. Generate msgFilterRules.dat content
      const { FilterGenerator } = resolveDeps();
      if (!FilterGenerator) {
        throw new Error('找不到 FilterGenerator 模組，請重新整理頁面。');
      }
      const datContent = FilterGenerator.generateMsgFilterRulesDat(parsedRwz.rules, resolvedFolders, {
        logging: true
      });

      // 3. Direct injection via Experiment API
      if (typeof messenger !== 'undefined' && messenger.rwzFilters && messenger.rwzFilters.importRules) {
        const result = await messenger.rwzFilters.importRules(selectedAccount.id, datContent);
        if (!result.success) {
          throw new Error(result.error || '核心寫入失敗');
        }
        log(`🎉 匯入成功！共寫入 ${result.totalImported} 條篩選器！`, 'success');
        if (result.targetPath) {
          log(`檔案已寫入: ${result.targetPath}`, 'info');
        }
        if (result.reloadStatus && result.reloadStatus !== 'none') {
          log(`記憶體快取重新載入方式: ${result.reloadStatus}`, 'info');
        }

        // Build post-import message
        let alertMsg = `✅ 成功匯入 ${result.totalImported} 條規則至「${selectedAccount.name}」！\n`;
        alertMsg += `\n檔案路徑: ${result.targetPath}\n`;

        if (result.needsRestart) {
          alertMsg += `\n⚠️ 重要提示：\n`;
          alertMsg += `Thunderbird 的記憶體快取可能尚未更新。\n`;
          alertMsg += `請依照以下步驟確保規則生效：\n\n`;
          alertMsg += `方法 1（推薦）：完全關閉並重新啟動 Thunderbird\n`;
          alertMsg += `方法 2：至「工具」→「郵件篩選器」，關閉後再開啟\n\n`;
          alertMsg += `若重啟後仍未看到規則，系統已同時備份下載 msgFilterRules.dat，\n`;
          alertMsg += `請手動將該檔案放至上述路徑覆蓋即可。`;

          log('⚠️ Thunderbird 記憶體快取可能需要重啟才能刷新。建議完全關閉並重新啟動 Thunderbird。', 'warn');
          log(`📋 若需手動放入，請將 msgFilterRules.dat 放至: ${result.targetPath}`, 'info');

          // Auto-download as backup
          autoDownloadDatBackup(datContent);
        } else {
          alertMsg += `\n已自動重新載入記憶體快取，規則應已立即生效！`;
          alertMsg += `\n請至「工具」→「郵件篩選器」確認。`;
          log('已自動重新載入 Thunderbird 記憶體中的篩選器快取。', 'success');
        }

        alert(alertMsg);
      } else {
        // Fallback when Experiment API not accessible in this context
        log('未偵測到 Experiment 權限，已為您自動建立所需資料夾，並準備產生 msgFilterRules.dat 檔案。', 'warn');
        showManualImportGuide();
        handleDownloadDat();
      }
    } catch (err) {
      log(`核心寫入發生例外: ${err.message}，自動切換為下載 msgFilterRules.dat 規則檔...`, 'warn');

      // Try to get the profile path for guidance
      let profilePath = '';
      try {
        if (typeof messenger !== 'undefined' && messenger.rwzFilters && messenger.rwzFilters.getFilterFilePath) {
          const pathResult = await messenger.rwzFilters.getFilterFilePath(selectedAccount.id);
          if (pathResult.success) {
            profilePath = pathResult.targetPath;
          }
        }
      } catch (_) {}

      handleDownloadDat();

      let fallbackMsg = `核心直接寫入失敗（${err.message}）。\n\n`;
      fallbackMsg += `別擔心！所有資料夾皆已自動建立完成，系統已自動為您下載 msgFilterRules.dat 規則檔！\n\n`;
      if (profilePath) {
        fallbackMsg += `📁 請將下載的 msgFilterRules.dat 放至:\n${profilePath}\n\n`;
        fallbackMsg += `然後重新啟動 Thunderbird 即可。`;
        log(`📋 手動放入路徑: ${profilePath}`, 'info');
      } else {
        fallbackMsg += `請至 Thunderbird：說明 → 疑難排解資訊 → 側寫檔案資料夾 → 開啟資料夾\n`;
        fallbackMsg += `然後進入 Mail/ 下對應的帳號目錄，替換 msgFilterRules.dat。`;
      }
      alert(fallbackMsg);
    } finally {
      directImportBtn.disabled = false;
      updateRulesTable();
    }
  }

  // Auto-download dat file as backup (no user interaction needed)
  function autoDownloadDatBackup(datContent) {
    try {
      const blob = new Blob([datContent], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'msgFilterRules.dat';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      log('📥 已同時備份下載 msgFilterRules.dat（以備手動替換之需）。', 'info');
    } catch (e) {
      console.warn('[autoDownloadDatBackup]', e);
    }
  }

  // Show manual import guide in log
  function showManualImportGuide() {
    log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', 'info');
    log('📖 手動匯入說明：', 'info');
    log('1. 先完全關閉 Thunderbird', 'info');
    log('2. 在 Thunderbird 中點選「說明 → 疑難排解資訊」', 'info');
    log('3. 找到「側寫檔案資料夾」→ 點擊「開啟資料夾」', 'info');
    log('4. 進入 Mail/ 或 ImapMail/ 下對應的帳號目錄', 'info');
    log('5. 將下載的 msgFilterRules.dat 放入該目錄（覆蓋舊檔）', 'info');
    log('6. 重新啟動 Thunderbird → 工具 → 郵件篩選器 即可看到規則', 'info');
    log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', 'info');
  }

  // Handle Download msgFilterRules.dat
  async function handleDownloadDat() {
    if (!parsedRwz || !selectedAccount) return;

    try {
      const { FilterGenerator } = resolveDeps();
      if (!FilterGenerator) {
        throw new Error('找不到 FilterGenerator 模組，請重新整理頁面。');
      }
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

  // Handle Export Existing Thunderbird Rules
  async function handleExportAccountRules() {
    if (!selectedAccount) {
      alert('請先在步驟 2 選擇要匯出規則的 Thunderbird 帳號！');
      return;
    }

    log(`正在擷取帳號「${selectedAccount.name}」目前的 Thunderbird 郵件篩選器...`);

    try {
      if (typeof messenger !== 'undefined' && messenger.rwzFilters && messenger.rwzFilters.exportAccountRules) {
        const result = await messenger.rwzFilters.exportAccountRules(selectedAccount.id);
        const datContent = result.datContent || 'version="9"\nlogging="yes"\n';
        const blob = new Blob([datContent], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const cleanName = String(selectedAccount.name || 'account').replace(/[^a-zA-Z0-9_-]/g, '_');
        a.download = `msgFilterRules_${cleanName}.dat`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        log(`🎉 成功匯出帳號「${selectedAccount.name}」的現有規則，共 ${result.filterCount || 0} 條規則！`, 'success');
        alert(`成功匯出帳號「${selectedAccount.name}」的 Thunderbird 規則！\n共 ${result.filterCount || 0} 條規則，已下載為 msgFilterRules_${cleanName}.dat。`);
      } else {
        // Fallback: prompt user about location or check if already converted rules exist
        log('未偵測到 Experiment 權限，已為您準備導出介面中的規則。', 'warn');
        if (parsedRwz) {
          handleDownloadDat();
        } else {
          alert('尚未載入規則，且未偵測到 Experiment 直接讀取權限。');
        }
      }
    } catch (err) {
      log(`匯出目前帳號規則失敗: ${err.message}`, 'error');
      alert(`匯出失敗：\n${err.message}`);
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

  function escapeHtml(val) {
    if (val === null || val === undefined) return '';
    const str = String(val);
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
