(function () {
  'use strict';

  class FolderManager {
  /**
   * @param {object} [messengerApi] Thunderbird messenger API (or mock)
   */
  constructor(messengerApi) {
    this.messenger = messengerApi || (typeof messenger !== 'undefined' ? messenger : null);
  }

  /**
   * Recursively flatten a folder tree into a list
   * @param {Array} folders
   * @param {string} [prefix='']
   * @returns {Array<{ folder: object, fullPath: string, name: string }>}
   */
  static flattenFolderTree(folders, prefix = '') {
    const list = [];
    if (!folders || !Array.isArray(folders)) return list;

    for (const f of folders) {
      const currentPath = prefix ? `${prefix}/${f.name}` : f.name;
      list.push({
        folder: f,
        fullPath: currentPath,
        name: f.name
      });

      if (f.subFolders && f.subFolders.length > 0) {
        list.push(...FolderManager.flattenFolderTree(f.subFolders, currentPath));
      }
    }
    return list;
  }

  /**
   * Find a folder by name or path in an account
   * @param {object} account
   * @param {string} folderPath e.g. "Work/Invoices" or "Invoices"
   * @returns {object|null} MailFolder
   */
  static findFolderInAccount(account, folderPath) {
    if (!account || !account.folders) return null;
    const flatList = FolderManager.flattenFolderTree(account.folders);

    const cleanPath = folderPath.replace(/\\/g, '/').toLowerCase();
    const leafName = cleanPath.split('/').pop();

    // 1. Exact fullPath match
    const exactMatch = flatList.find(item => item.fullPath.toLowerCase() === cleanPath);
    if (exactMatch) return exactMatch.folder;

    // 2. Leaf name match
    const leafMatch = flatList.find(item => item.name.toLowerCase() === leafName);
    if (leafMatch) return leafMatch.folder;

    return null;
  }

  /**
   * Check which folders referenced by rules exist and which are missing
   * @param {Array} rules
   * @param {object} account
   * @returns {{ existing: Array<string>, missing: Array<string> }}
   */
  static checkRuleFolders(rules, account) {
    const folderSet = new Set();
    for (const rule of rules) {
      if (rule.targetFolders) {
        for (const tf of rule.targetFolders) {
          if (tf) folderSet.add(tf);
        }
      }
    }

    const existing = [];
    const missing = [];

    for (const folderName of folderSet) {
      const found = FolderManager.findFolderInAccount(account, folderName);
      if (found) {
        existing.push(folderName);
      } else {
        missing.push(folderName);
      }
    }

    return { existing, missing, all: Array.from(folderSet) };
  }

  /**
   * Automatically ensure all required folders exist in the target account.
   * If a folder does not exist, it will be created recursively!
   *
   * @param {Array<string>} folderPaths List of folder names/paths
   * @param {object} account Target Thunderbird account
   * @param {Function} [onLog] Callback for progress messages
   * @returns {Promise<Map<string, { folder: object, uri: string, created: boolean }>>}
   */
  async ensureFoldersExist(folderPaths, account, onLog = console.log) {
    const results = new Map();
    if (!account) {
      throw new Error('未指定目標帳號');
    }

    // Determine the root parent folder for new folders in this account
    let rootFolder = null;
    if (account.folders && account.folders.length > 0) {
      // Look for a root folder or the first folder
      rootFolder = account.folders.find(f => f.isRoot || f.path === '/' || !f.path) || account.folders[0];
    }

    if (!rootFolder) {
      throw new Error(`無法在帳號「${account.name}」中找到根資料夾以建立新資料夾`);
    }

    for (const targetPath of folderPaths) {
      if (!targetPath) continue;

      // 1. Check if it already exists
      const existing = FolderManager.findFolderInAccount(account, targetPath);
      if (existing) {
        onLog(`[檢查] 資料夾「${targetPath}」已存在於帳號「${account.name}」`);
        results.set(targetPath, {
          folder: existing,
          uri: this.constructFolderUri(existing, account),
          created: false
        });
        continue;
      }

      // 2. Folder does not exist -> create it!
      onLog(`[建立] 資料夾「${targetPath}」不存在，即將自動建立...`);
      const segments = targetPath.replace(/\\/g, '/').split('/').filter(Boolean);

      let currentParent = rootFolder;
      let currentFolder = null;

      for (let i = 0; i < segments.length; i++) {
        const seg = segments[i];

        // Check if subfolder exists under currentParent
        let sub = null;
        if (currentParent.subFolders) {
          sub = currentParent.subFolders.find(f => f.name.toLowerCase() === seg.toLowerCase());
        }

        if (sub) {
          currentParent = sub;
          currentFolder = sub;
        } else {
          // Create the subfolder using Thunderbird WebExtension API
          onLog(`[建立中] 正在父層「${currentParent.name || '根目錄'}」下建立子資料夾「${seg}」...`);
          try {
            if (this.messenger && this.messenger.folders && this.messenger.folders.create) {
              const newFolder = await this.messenger.folders.create(currentParent.id || currentParent, seg);
              if (!currentParent.subFolders) currentParent.subFolders = [];
              currentParent.subFolders.push(newFolder);
              currentParent = newFolder;
              currentFolder = newFolder;
              onLog(`[成功] 已建立資料夾: ${seg}`);
            } else {
              // Simulated creation for testing or headless environments
              const mockFolder = {
                id: `${currentParent.id || 'folder'}_${seg}`,
                name: seg,
                path: `${currentParent.path || ''}/${seg}`,
                subFolders: []
              };
              if (!currentParent.subFolders) currentParent.subFolders = [];
              currentParent.subFolders.push(mockFolder);
              currentParent = mockFolder;
              currentFolder = mockFolder;
              onLog(`[模擬成功] 已建立虛擬資料夾: ${seg}`);
            }
          } catch (err) {
            onLog(`[錯誤] 建立資料夾「${seg}」失敗: ${err.message}`);
            throw err;
          }
        }
      }

      results.set(targetPath, {
        folder: currentFolder,
        uri: this.constructFolderUri(currentFolder, account),
        created: true
      });
    }

    return results;
  }

  /**
   * Construct a standard Thunderbird folder URI
   * e.g. mailbox://nobody@Local%20Folders/TargetFolder
   * or imap://user@server/TargetFolder
   */
  constructFolderUri(folder, account) {
    if (!folder) return '';

    // If folder already has a native URI
    if (folder.URI) return folder.URI;

    const folderName = folder.name || 'Folder';
    const folderPath = folder.path || `/${folderName}`;
    const cleanPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;

    if (account && (account.type === 'none' || account.name.includes('Local Folders') || account.name.includes('本地資料夾'))) {
      return `mailbox://nobody@Local%20Folders/${encodeURIComponent(cleanPath)}`;
    }

    if (account) {
      const serverType = account.type === 'pop3' ? 'mailbox' : 'imap';
      const cleanServer = (account.name || 'account').replace(/[^a-zA-Z0-9@._-]/g, '_');
      return `${serverType}://${cleanServer}/${encodeURIComponent(cleanPath)}`;
    }

    return `mailbox://nobody@Local%20Folders/${encodeURIComponent(cleanPath)}`;
  }
}

if (typeof globalThis !== 'undefined') {
  globalThis.FolderManager = FolderManager;
}
if (typeof window !== 'undefined') {
  window.FolderManager = FolderManager;
}
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { FolderManager };
}
})();
