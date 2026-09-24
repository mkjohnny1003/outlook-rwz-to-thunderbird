/**
 * Regression tests: real Thunderbird folder structure (no isRoot folder in
 * account.folders) and real folder URI resolution inside the Experiment.
 */
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { FolderManager } = require('../lib/folder-manager.js');

async function runTests() {
  console.log('--- 測試真實 Thunderbird 結構與 URI 解析 ---');

  // 1. New folders must be created at account root, not inside Inbox
  const created = [];
  const fakeMessenger = {
    folders: {
      create: async (parent, name) => {
        const parentPath = parent.type ? '' : parent.path; // MailAccount has .type
        const f = { id: `${parentPath}/${name}`, name, path: `${parentPath}/${name}`, subFolders: [] };
        created.push(f.path);
        return f;
      }
    }
  };
  const account = {
    id: 'account1', name: 'mkjohnny@gmail.com', type: 'imap',
    folders: [{ id: 'i', name: 'Inbox', path: '/INBOX', subFolders: [] }]
  };
  const fm = new FolderManager(fakeMessenger);
  await fm.ensureFoldersExist(['客戶/宏齊'], account, () => {});
  assert.deepStrictEqual(created, ['/客戶', '/客戶/宏齊']);
  console.log('✓ 新資料夾建立於帳號根目錄（不再塞進收件匣）');

  // 2. Experiment resolves WebExtension path -> real native URI
  const src = fs.readFileSync(path.join(__dirname, '../api/filters/implementation.js'), 'utf8');
  const ctx = { ChromeUtils: { importESModule: () => { throw new Error('n/a'); } }, console };
  vm.createContext(ctx);
  vm.runInContext(src + '\nthis.listDescendants=listDescendants;this.resolvePathToUri=resolvePathToUri;', ctx);

  const rootUri = 'imap://mkjohnny%40gmail.com@imap.gmail.com';
  const root = { URI: rootUri, isServer: true };
  const mk = (p, name, parent) => ({ URI: rootUri + '/' + p.split('/').map(encodeURIComponent).join('/'), prettyName: name, parent });
  const inbox = mk('INBOX', '收件匣', root);
  const cust = mk('客戶', '客戶', root);
  const hq = mk('客戶/宏齊', '宏齊', cust);
  root.descendants = [inbox, cust, hq];

  const entries = ctx.listDescendants(root);
  assert.strictEqual(ctx.resolvePathToUri(entries, '/客戶/宏齊'), hq.URI);
  assert.strictEqual(ctx.resolvePathToUri(entries, '/INBOX'), inbox.URI);
  assert.strictEqual(ctx.resolvePathToUri(entries, '收件匣'), inbox.URI);
  assert.strictEqual(ctx.resolvePathToUri(entries, '宏齊'), hq.URI);
  assert.strictEqual(ctx.resolvePathToUri(entries, '不存在'), null);
  console.log('✓ 路徑正確解析為原生 URI:', hq.URI);

  // 3. Repair logic: old bogus URI (slashes encoded as %2F) can be mapped back
  const bogus = 'imap://mkjohnny@gmail.com/INBOX%2F%E5%AE%A2%E6%88%B6';
  const tail = decodeURIComponent(bogus.match(/^[a-z-]+:\/\/[^/]*\/(.*)$/i)[1]);
  assert.strictEqual(tail, 'INBOX/客戶');
  assert.strictEqual(ctx.resolvePathToUri(entries, tail), cust.URI); // unique leaf fallback
  console.log('✓ 舊版錯誤 URI 可被修復功能對回實際資料夾');
}

module.exports = runTests;
if (require.main === module) runTests().catch(e => { console.error(e); process.exit(1); });
