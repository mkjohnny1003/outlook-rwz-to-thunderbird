/**
 * WebExtension Experiment Implementation for Thunderbird Message Filters
 * Uses IOUtils for direct file I/O on msgFilterRules.dat.
 * Does NOT use MailServices.filters.getFilterList (removed in TB 156+).
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

        async importRules(accountId, datContent, rulesJson) {
          try {
            if (!MailServices) {
              return { success: false, error: '無法存取 Thunderbird MailServices 內部服務' };
            }

            // 1. Resolve Account
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

            if (!account) {
              return { success: false, error: `找不到指定的郵件帳號: ${accountId}` };
            }

            const rootFolder = account.incomingServer ? account.incomingServer.rootFolder : null;
            if (!rootFolder) {
              return { success: false, error: `無法取得帳號「${account.key}」的根目錄 (rootFolder)` };
            }

            // 2. Resolve target file path directly (no getFilterList — not available in TB 156+)
            let targetPath = '';
            if (rootFolder.filePath && rootFolder.filePath.path) {
              const sep = rootFolder.filePath.path.includes('\\') ? '\\' : '/';
              targetPath = rootFolder.filePath.path + sep + 'msgFilterRules.dat';
            }

            if (!targetPath) {
              return { success: false, error: '無法解析目標帳號之 msgFilterRules.dat 實體路徑' };
            }

            // 3. Read existing rules if file exists
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

            // 4. Merge rules: append new rules while avoiding duplicate header
            let finalDat = datContent;
            if (existingText && existingText.trim()) {
              const cleanNew = datContent
                .split('\n')
                .filter(l => !l.startsWith('version=') && !l.startsWith('logging='))
                .join('\n')
                .trim();
              finalDat = existingText.trimEnd() + '\n\n' + cleanNew + '\n';
            }

            // 5. Write to target file using IOUtils
            let writeSuccess = false;
            if (typeof IOUtils !== 'undefined') {
              try {
                await IOUtils.writeUTF8(targetPath, finalDat, {
                  tmpPath: targetPath + '.tmp'
                });
                writeSuccess = true;
              } catch (writeErr) {
                console.warn('[rwzFilters] IOUtils.writeUTF8 失敗，嘗試直接寫入:', writeErr);
                try {
                  await IOUtils.writeUTF8(targetPath, finalDat);
                  writeSuccess = true;
                } catch (e2) {
                  return { success: false, error: `寫入檔案失敗: ${e2.message}` };
                }
              }
            } else {
              return { success: false, error: '當前環境不支援 IOUtils 檔案寫入 API' };
            }

            // 6. Reload note: Thunderbird will pick up changes on next filter dialog open or restart

            const matchRules = datContent.match(/^name=/gm);
            const totalImported = matchRules ? matchRules.length : 1;

            return {
              success: true,
              totalImported,
              targetPath,
              accountName: account.incomingServer.prettyName || account.key
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

            if (!account) {
              return { success: false, error: `找不到指定的郵件帳號: ${accountId}` };
            }

            const rootFolder = account.incomingServer ? account.incomingServer.rootFolder : null;
            if (!rootFolder) {
              return { success: false, error: `無法取得帳號「${account.key}」的根目錄` };
            }

            // Resolve target file path directly (no getFilterList — not available in TB 156+)
            let targetPath = '';
            if (rootFolder.filePath && rootFolder.filePath.path) {
              const sep = rootFolder.filePath.path.includes('\\') ? '\\' : '/';
              targetPath = rootFolder.filePath.path + sep + 'msgFilterRules.dat';
            }

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
