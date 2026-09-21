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

        async importRules(accountId, rules) {
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

          let addedCount = 0;

          for (const r of rules) {
            try {
              // 1. Create filter
              const filter = filterList.createFilter(r.name || 'Outlook Converted Filter');
              filter.enabled = Boolean(r.enabled);
              filter.filterType = r.filterType || 17; // 1 (Inbox) | 16 (Manual)

              // 2. Add Actions
              if (r.actions && Array.isArray(r.actions)) {
                for (const act of r.actions) {
                  const actionObj = filter.createAction();
                  actionObj.type = act.type; // integer action type

                  if (act.targetFolderUri) {
                    actionObj.targetFolderUri = act.targetFolderUri;
                  }
                  if (act.strValue) {
                    actionObj.strValue = act.strValue;
                  }
                  if (act.priority) {
                    // map priority string
                    actionObj.priority = act.priority === 'Highest' ? 5 :
                                         act.priority === 'High' ? 4 :
                                         act.priority === 'Low' ? 2 :
                                         act.priority === 'Lowest' ? 1 : 3;
                  }
                  filter.appendAction(actionObj);
                }
              }

              // 3. Parse condition
              if (r.condition && r.condition !== 'ALL') {
                filterList.parseCondition(filter, r.condition);
              }

              // 4. Insert into list
              filterList.insertFilterAt(filterList.filterCount, filter);
              addedCount++;
            } catch (ruleErr) {
              console.error(`[rwzFilters] 匯入規則「${r.name}」失敗:`, ruleErr);
            }
          }

          // 5. Persist changes to disk (msgFilterRules.dat)
          try {
            filterList.saveToDefaultFile();
          } catch (saveErr) {
            console.warn('[rwzFilters] saveToDefaultFile 警告:', saveErr);
          }

          return {
            success: true,
            totalImported: addedCount,
            accountName: account.incomingServer.prettyName || accountId
          };
        }
      }
    };
  }
};
