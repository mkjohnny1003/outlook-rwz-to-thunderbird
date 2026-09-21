/**
 * WebExtension Experiment Implementation for Thunderbird Message Filters
 * Grants direct, instant access to MailServices.filters without requiring Thunderbird restart.
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
          if (!MailServices) {
            throw new Error('無法存取 Thunderbird MailServices 內部服務');
          }

          // 1. Resolve Account
          let account = null;
          try {
            account = MailServices.accounts.getAccount(accountId);
          } catch (_) {}

          if (!account && MailServices.accounts.accounts) {
            for (const acc of MailServices.accounts.accounts) {
              if (acc.key === accountId || (acc.incomingServer && acc.incomingServer.key === accountId)) {
                account = acc;
                break;
              }
            }
          }

          if (!account) {
            throw new Error(`找不到指定的郵件帳號: ${accountId}`);
          }

          const rootFolder = account.incomingServer ? account.incomingServer.rootFolder : null;
          if (!rootFolder) {
            throw new Error(`無法取得帳號「${account.key}」的根目錄`);
          }

          const filterList = MailServices.filters.getFilterList(rootFolder);

          // 2. Identify target file
          let targetFile = null;
          if (filterList && filterList.defaultFile) {
            targetFile = filterList.defaultFile;
          } else if (rootFolder.filePath) {
            targetFile = rootFolder.filePath.clone();
            targetFile.append('msgFilterRules.dat');
          }

          // 3. Read existing file if present
          let existingText = '';
          if (targetFile && targetFile.exists()) {
            try {
              const fstream = Components.classes['@mozilla.org/network/file-input-stream;1']
                .createInstance(Components.interfaces.nsIFileInputStream);
              const cstream = Components.classes['@mozilla.org/intl/converter-input-stream;1']
                .createInstance(Components.interfaces.nsIConverterInputStream);
              fstream.init(targetFile, -1, 0, 0);
              cstream.init(fstream, 'UTF-8', 1024, Components.interfaces.nsIConverterInputStream.DEFAULT_REPLACEMENT_CHARACTER);
              const readChunk = {};
              while (cstream.readString(4096, readChunk) !== 0) {
                existingText += readChunk.value;
              }
              cstream.close();
            } catch (readErr) {
              console.warn('[rwzFilters] 讀取現存 filter 警告:', readErr);
            }
          }

          // 4. Merge rules
          let finalDat = datContent;
          if (existingText && existingText.trim()) {
            const cleanNew = datContent
              .split('\n')
              .filter(l => !l.startsWith('version=') && !l.startsWith('logging='))
              .join('\n')
              .trim();
            finalDat = existingText.trimEnd() + '\n\n' + cleanNew + '\n';
          }

          // 5. Write to disk
          if (targetFile) {
            const foStream = Components.classes['@mozilla.org/network/file-output-stream;1']
              .createInstance(Components.interfaces.nsIFileOutputStream);
            // 0x02: PR_WRONLY, 0x08: PR_CREATE_FILE, 0x20: PR_TRUNCATE
            foStream.init(targetFile, 0x02 | 0x08 | 0x20, 0o664, 0);
            const converter = Components.classes['@mozilla.org/intl/converter-output-stream;1']
              .createInstance(Components.interfaces.nsIConverterOutputStream);
            converter.init(foStream, 'UTF-8', 0, 0);
            converter.writeString(finalDat);
            converter.close();
            foStream.close();
          }

          // 6. Reload in-memory filter list
          if (filterList) {
            try {
              if (typeof filterList.reload === 'function') {
                filterList.reload();
              } else {
                MailServices.filters.getFilterList(rootFolder);
              }
            } catch (_) {}
          }

          const matchRules = datContent.match(/^name=/gm);
          const totalImported = matchRules ? matchRules.length : 1;

          return {
            success: true,
            totalImported,
            accountName: account.incomingServer.prettyName || account.key
          };
        },

        async exportAccountRules(accountId) {
          if (!MailServices) {
            throw new Error('無法存取 Thunderbird MailServices 內部服務');
          }

          const account = MailServices.accounts.getAccount(accountId);
          if (!account) {
            throw new Error(`找不到指定的郵件帳號: ${accountId}`);
          }

          const rootFolder = account.incomingServer.rootFolder;
          const filterList = MailServices.filters.getFilterList(rootFolder);
          if (!filterList) {
            throw new Error(`無法取得帳號「${account.key}」的篩選器清單`);
          }

          try {
            filterList.saveToDefaultFile();
          } catch (_) {}

          let datContent = '';
          const defaultFile = filterList.defaultFile;

          if (defaultFile && defaultFile.exists()) {
            try {
              const fstream = Components.classes['@mozilla.org/network/file-input-stream;1']
                .createInstance(Components.interfaces.nsIFileInputStream);
              const cstream = Components.classes['@mozilla.org/intl/converter-input-stream;1']
                .createInstance(Components.interfaces.nsIConverterInputStream);
              fstream.init(defaultFile, -1, 0, 0);
              cstream.init(fstream, 'UTF-8', 1024, Components.interfaces.nsIConverterInputStream.DEFAULT_REPLACEMENT_CHARACTER);
              const readChunk = {};
              while (cstream.readString(4096, readChunk) !== 0) {
                datContent += readChunk.value;
              }
              cstream.close();
            } catch (readErr) {
              console.warn('[rwzFilters] 讀取 defaultFile 警告:', readErr);
            }
          }

          return {
            success: true,
            filterCount: filterList.filterCount,
            accountName: account.incomingServer.prettyName || account.key,
            datContent
          };
        }
      }
    };
  }
};
