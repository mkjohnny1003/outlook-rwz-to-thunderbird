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


/**
 * Helper: list every descendant nsIMsgFolder of an account root, with
 * the WebExtension-style path ("/INBOX/Sub") and the display-name chain.
 */
function listDescendants(rootFolder) {
  const out = [];
  if (!rootFolder) return out;
  const rootUri = rootFolder.URI;
  let desc = [];
  try {
    desc = rootFolder.descendants;
    if (desc && !Array.isArray(desc) && typeof desc.enumerate === 'function') {
      // very old TB returned nsIArray
      const arr = [];
      const e = desc.enumerate();
      while (e.hasMoreElements()) arr.push(e.getNext());
      desc = arr;
    }
  } catch (_) {
    desc = [];
  }
  for (const f of desc || []) {
    let uriPath = '';
    try {
      uriPath = f.URI.startsWith(rootUri)
        ? f.URI.substring(rootUri.length).split('/').map(s => {
            try { return decodeURIComponent(s); } catch (_) { return s; }
          }).join('/')
        : '';
    } catch (_) {}
    if (uriPath && !uriPath.startsWith('/')) uriPath = '/' + uriPath;
    const names = [];
    let cur = f;
    while (cur && cur !== rootFolder && !cur.isServer) {
      names.unshift(cur.prettyName || cur.name);
      cur = cur.parent;
    }
    out.push({ folder: f, uri: f.URI, uriPath, namePath: '/' + names.join('/'), leaf: (f.prettyName || f.name) });
  }
  return out;
}

/**
 * Helper: resolve one WebExtension folder path (e.g. "/INBOX/客戶") or a
 * display-name path (e.g. "收件匣/客戶") to the real folder URI.
 */
function resolvePathToUri(entries, wanted) {
  const w = '/' + String(wanted || '').replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
  const lw = w.toLowerCase();
  let hit = entries.find(e => e.uriPath === w) ||
            entries.find(e => e.uriPath.toLowerCase() === lw) ||
            entries.find(e => e.namePath.toLowerCase() === lw);
  if (hit) return hit.uri;
  // unique leaf-name match as last resort
  const leaf = lw.split('/').pop();
  const leafHits = entries.filter(e => String(e.leaf).toLowerCase() === leaf);
  if (leafHits.length === 1) return leafHits[0].uri;
  return null;
}

function folderExists(uri) {
  try {
    let MailUtils;
    try {
      MailUtils = ChromeUtils.importESModule
        ? ChromeUtils.importESModule('resource:///modules/MailUtils.sys.mjs').MailUtils
        : ChromeUtils.import('resource:///modules/MailUtils.jsm').MailUtils;
    } catch (_) {}
    if (MailUtils && MailUtils.getExistingFolder) {
      return Boolean(MailUtils.getExistingFolder(uri));
    }
  } catch (_) {}
  return null; // unknown
}

/**
 * Helper: replace filter blocks in an existing msgFilterRules.dat that have the
 * same name as blocks in the new content (so re-importing doesn't duplicate).
 */
function splitDatBlocks(text) {
  const header = [];
  const blocks = [];
  let cur = null;
  for (const line of String(text || '').split(/\r?\n/)) {
    if (line.startsWith('name=')) {
      cur = { name: line, lines: [line] };
      blocks.push(cur);
    } else if (cur) {
      if (line.trim()) cur.lines.push(line);
    } else if (line.trim()) {
      header.push(line);
    }
  }
  return { header, blocks };
}

var rwzFilters = class extends (ExtensionCommonModule?.ExtensionCommon?.ExtensionAPI || class {}) {
  getAPI(context) {
    return {
      rwzFilters: {
        async isAvailable() {
          return Boolean(MailServices && MailServices.accounts);
        },


        /**
         * Resolve folder paths (WebExtension MailFolder.path, e.g. "/INBOX/Sub")
         * to their real Thunderbird folder URIs.
         * Returns an object { [path]: uri|null }.
         */
        async resolveFolderUris(accountId, paths) {
          const out = {};
          const account = resolveAccount(accountId);
          const root = account && account.incomingServer ? account.incomingServer.rootFolder : null;
          const entries = listDescendants(root);
          for (const p of paths || []) {
            out[p] = resolvePathToUri(entries, p);
          }
          return out;
        },

        /**
         * Scan the account's filters for Move/Copy actions pointing at folders that
         * don't exist, and re-point them at the matching real folder.
         * Uses the in-memory filter list + saveToDefaultFile(), so no restart needed.
         */
        async repairFilterTargets(accountId) {
          try {
            const account = resolveAccount(accountId);
            if (!account || !account.incomingServer) {
              return { success: false, error: `找不到帳號: ${accountId}` };
            }
            const server = account.incomingServer;
            const root = server.rootFolder;
            const entries = listDescendants(root);
            const FA = (typeof Ci !== 'undefined' ? Ci : Components.interfaces).nsMsgFilterAction || { MoveToFolder: 1, CopyToFolder: 16 };
            const fl = server.getFilterList(null);
            const fixed = [];
            const unresolved = [];
            for (let i = 0; i < fl.filterCount; i++) {
              const filter = fl.getFilterAt(i);
              const actions = filter.sortedActionList || [];
              for (const a of actions) {
                if (a.type !== FA.MoveToFolder && a.type !== FA.CopyToFolder) continue;
                const uri = a.targetFolderUri;
                if (!uri) continue;
                if (folderExists(uri) === true) continue;
                // Take the part after scheme://host/ and decode it (this also turns %2F back into "/")
                const m = uri.match(/^[a-z-]+:\/\/[^/]*\/(.*)$/i);
                let tail = m ? m[1] : uri;
                try { tail = decodeURIComponent(tail); } catch (_) {}
                const realUri = resolvePathToUri(entries, tail);
                if (realUri) {
                  a.targetFolderUri = realUri;
                  fixed.push({ filter: filter.filterName, from: uri, to: realUri });
                } else {
                  unresolved.push({ filter: filter.filterName, uri });
                }
              }
            }
            if (fixed.length) fl.saveToDefaultFile();
            return { success: true, fixed, unresolved };
          } catch (err) {
            return { success: false, error: err.message, stack: err.stack };
          }
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
            let replacedCount = 0;
            if (existingText && existingText.trim()) {
              const oldParsed = splitDatBlocks(existingText);
              const newParsed = splitDatBlocks(datContent);
              const newNames = new Set(newParsed.blocks.map(b => b.name));
              const kept = oldParsed.blocks.filter(b => !newNames.has(b.name));
              replacedCount = oldParsed.blocks.length - kept.length;
              const header = oldParsed.header.length ? oldParsed.header : newParsed.header;
              finalDat = header.join('\n') + '\n' +
                kept.concat(newParsed.blocks).map(b => b.lines.join('\n')).join('\n') + '\n';
            }
            // Backup the original file once per import
            if (existingText && typeof IOUtils !== 'undefined') {
              try {
                await IOUtils.writeUTF8(targetPath + '.rwz-backup-' + Date.now(), existingText);
              } catch (_) {}
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
              replacedCount,
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
