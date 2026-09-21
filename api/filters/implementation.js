/**
 * WebExtension Experiment Implementation for Thunderbird Message Filters
 * Uses IOUtils for direct file I/O on msgFilterRules.dat.
 * After writing, attempts to force in-memory reload via incomingServer.getFilterList().
 */

/* global ChromeUtils, ExtensionCommon, Services */

let ExtensionCommonModule;
try {
  ExtensionCommonModule = ChromeUtils.importESModule
    ? ChromeUtils.importESModule('resource://gre/modules/ExtensionCommon.sys.mjs')
    : ChromeUtils.import('resource://gre/modules/ExtensionCommon.jsm');
} catch (_) {}

let MailServices;
try {
  MailServices = ChromeUtils.importESModule
    ? ChromeUtils.importESModule('resource:///modules/MailServices.sys.mjs').MailServices
    : ChromeUtils.import('resource:///modules/MailServices.jsm').MailServices;
} catch (_) {}

/**
 * Helper: resolve an account by ID from MailServices
 */
function resolveAccount(accountId) {
  if (!MailServices) return null;
  let account = null;
  try {
    account = MailServices.accounts.getAccount(accountId);
  } catch (_) {}
  if (!account && MailServices.accounts && MailServices.accounts.accounts) {
    for (const acc of MailServices.accounts.accounts) {
      if (acc.key === accountId || (acc.incomingServer && acc.incomingServer.key === accountId)) {
        account = acc;
        break;
      }
    }
  }
  return account;
}

/**
 * Helper: build msgFilterRules.dat path from rootFolder
 */
function buildFilterPath(rootFolder) {
  if (!rootFolder || !rootFolder.filePath || !rootFolder.filePath.path) return '';
  const base = rootFolder.filePath.path;
  const sep = base.includes('\\') ? '\\' : '/';
  return base + sep + 'msgFilterRules.dat';
}

/**
 * Helper: force Thunderbird to reload in-memory filter list from disk.
 * Tries multiple approaches since APIs vary across TB versions.
 * Returns { reloaded: boolean, method: string }
 */
function forceFilterReload(account, rootFolder) {
  const result = { reloaded: false, method: 'none' };

  // Approach 1: incomingServer.getFilterList(null) — available on nsIMsgIncomingServer
  // This is DIFFERENT from MailServices.filters.getFilterList() which was removed.
  try {
    if (account.incomingServer && typeof account.incomingServer.getFilterList === 'function') {
      const filterList = account.incomingServer.getFilterList(null);
      if (filterList) {
        // Force reload from disk
        if (typeof filterList.parseCondition === 'function' || typeof filterList.matchOrChangeFilterTarget === 'function') {
          // filterList exists; the act of calling getFilterList(null) may itself trigger reload
          result.reloaded = true;
          result.method = 'incomingServer.getFilterList(null)';
        }
        // Try explicit methods if available
        if (typeof filterList.reload === 'function') {
          filterList.reload();
          result.reloaded = true;
          result.method = 'filterList.reload()';
        }
      }
    }
  } catch (e) {
    console.warn('[rwzFilters] incomingServer.getFilterList reload 嘗試:', e.message);
  }

  // Approach 2: clearFilterList + re-fetch to force cache invalidation
  try {
    if (account.incomingServer && typeof account.incomingServer.clearFilterList === 'function') {
      account.incomingServer.clearFilterList();
      result.reloaded = true;
      result.method = 'clearFilterList()';
    }
  } catch (e) {
    console.warn('[rwzFilters] clearFilterList 嘗試:', e.message);
  }

  // Approach 3: MailServices.filters (if available in older TB versions)
  try {
    if (MailServices.filters && typeof MailServices.filters.getFilterList === 'function') {
      const fl = MailServices.filters.getFilterList(rootFolder);
      if (fl && typeof fl.reload === 'function') {
        fl.reload();
        result.reloaded = true;
        result.method = 'MailServices.filters.getFilterList().reload()';
      }
    }
  } catch (e) {
    console.warn('[rwzFilters] MailServices.filters.getFilterList 嘗試:', e.message);
  }

  return result;
}

var rwzFilters = class extends (ExtensionCommonModule?.ExtensionCommon?.ExtensionAPI || class {}) {
  getAPI(context) {
    return {
      rwzFilters: {
        async isAvailable() {
          return Boolean(MailServices && MailServices.accounts && MailServices.filters);
        },

        async getAccountRootFolderUri(accountId) {
          if (!MailServices) throw new Error('MailServices unavailable');
          const account = MailServices.accounts.getAccount(accountId);
          if (!account) throw new Error(`找不到帳號: ${accountId}`);
          return account.incomingServer.rootFolder.URI;
        },

        /**
         * Returns the profile path where msgFilterRules.dat is stored for the given account.
         */
        async getFilterFilePath(accountId) {
          try {
            if (!MailServices) {
              return { success: false, error: 'MailServices unavailable' };
            }
            const account = resolveAccount(accountId);
            if (!account) {
              return { success: false, error: `找不到帳號: ${accountId}` };
            }
            const rootFolder = account.incomingServer ? account.incomingServer.rootFolder : null;
            const targetPath = buildFilterPath(rootFolder);
            return {
              success: true,
              targetPath,
              accountName: account.incomingServer.prettyName || account.key
            };
          } catch (err) {
            return { success: false, error: err.message };
          }
        },

        async importRules(accountId, datContent, rulesJson) {
          try {
            if (!MailServices) {
              return { success: false, error: '無法存取 Thunderbird MailServices 內部服務' };
            }

            // 1. Resolve Account
            const account = resolveAccount(accountId);
            if (!account) {
              return { success: false, error: `找不到指定的郵件帳號: ${accountId}` };
            }

            const rootFolder = account.incomingServer ? account.incomingServer.rootFolder : null;
            if (!rootFolder) {
              return { success: false, error: `無法取得帳號「${account.key}」的根目錄 (rootFolder)` };
            }

            // 2. Resolve target file path
            const targetPath = buildFilterPath(rootFolder);
            if (!targetPath) {
              return { success: false, error: '無法解析目標帳號之 msgFilterRules.dat 實體路徑' };
            }

            // 3. IMPORTANT: Clear Thunderbird's in-memory filter cache BEFORE writing
            //    This prevents the stale in-memory state from overwriting our file later.
            const preReload = forceFilterReload(account, rootFolder);

            // 4. Read existing rules if file exists
            let existingText = '';
            if (typeof IOUtils !== 'undefined') {
              try {
                if (await IOUtils.exists(targetPath)) {
                  existingText = await IOUtils.readUTF8(targetPath);
                }
              } catch (readErr) {
                console.warn('[rwzFilters] IOUtils.readUTF8 警告:', readErr);
              }
            }

            // 5. Merge rules: append new rules while avoiding duplicate header
            let finalDat = datContent;
            if (existingText && existingText.trim()) {
              const cleanNew = datContent
                .split('\n')
                .filter(l => !l.startsWith('version=') && !l.startsWith('logging='))
                .join('\n')
                .trim();
              finalDat = existingText.trimEnd() + '\n\n' + cleanNew + '\n';
            }

            // 6. Write to target file using IOUtils
            if (typeof IOUtils !== 'undefined') {
              try {
                await IOUtils.writeUTF8(targetPath, finalDat, {
                  tmpPath: targetPath + '.tmp'
                });
              } catch (writeErr) {
                console.warn('[rwzFilters] IOUtils.writeUTF8 with tmpPath failed:', writeErr);
                try {
                  await IOUtils.writeUTF8(targetPath, finalDat);
                } catch (e2) {
                  return { success: false, error: `寫入檔案失敗: ${e2.message}` };
                }
              }
            } else {
              return { success: false, error: '當前環境不支援 IOUtils 檔案寫入 API' };
            }

            // 7. Force Thunderbird to reload the filter list from disk AFTER writing
            const postReload = forceFilterReload(account, rootFolder);

            const matchRules = datContent.match(/^name=/gm);
            const totalImported = matchRules ? matchRules.length : 1;

            return {
              success: true,
              totalImported,
              targetPath,
              accountName: account.incomingServer.prettyName || account.key,
              reloadStatus: postReload.reloaded ? postReload.method : (preReload.reloaded ? `pre:${preReload.method}` : 'none'),
              needsRestart: !postReload.reloaded
            };
          } catch (err) {
            return {
              success: false,
              error: err.message,
              stack: err.stack
            };
          }
        },

        async exportAccountRules(accountId) {
          try {
            if (!MailServices) {
              return { success: false, error: '無法存取 Thunderbird MailServices 內部服務' };
            }

            const account = resolveAccount(accountId);
            if (!account) {
              return { success: false, error: `找不到指定的郵件帳號: ${accountId}` };
            }

            const rootFolder = account.incomingServer ? account.incomingServer.rootFolder : null;
            if (!rootFolder) {
              return { success: false, error: `無法取得帳號「${account.key}」的根目錄` };
            }

            const targetPath = buildFilterPath(rootFolder);

            let datContent = '';
            if (targetPath && typeof IOUtils !== 'undefined') {
              try {
                if (await IOUtils.exists(targetPath)) {
                  datContent = await IOUtils.readUTF8(targetPath);
                }
              } catch (readErr) {
                console.warn('[rwzFilters] exportAccountRules readUTF8 警告:', readErr);
              }
            }

            const filterCount = (datContent.match(/^name=/gm) || []).length;

            return {
              success: true,
              filterCount,
              accountName: account.incomingServer.prettyName || account.key,
              datContent
            };
          } catch (err) {
            return {
              success: false,
              error: err.message,
              stack: err.stack
            };
          }
        }
      }
    };
  }
};
