/**
 * Unit tests for FolderManager
 * Tests automatic detection and hierarchical folder creation
 */

const assert = require('assert');
const { FolderManager } = require('../lib/folder-manager.js');

async function runTests() {
  console.log('--- 測試 FolderManager 資料夾檢查與自動建立 ---');

  // 1. Mock account structure
  const mockAccount = {
    id: 'account_test',
    name: 'user@example.com',
    type: 'imap',
    folders: [
      {
        id: 'f_root',
        name: 'user@example.com',
        path: '/',
        isRoot: true,
        subFolders: [
          {
            id: 'f_inbox',
            name: 'INBOX',
            path: '/INBOX',
            subFolders: []
          },
          {
            id: 'f_archive',
            name: 'Archive',
            path: '/Archive',
            subFolders: [
              {
                id: 'f_2025',
                name: '2025',
                path: '/Archive/2025',
                subFolders: []
              }
            ]
          }
        ]
      }
    ]
  };

  // Test flattening
  const flat = FolderManager.flattenFolderTree(mockAccount.folders);
  assert.strictEqual(flat.length, 4);
  console.log('✓ 資料夾樹狀結構扁平化正確');

  // Test finding existing folders
  const foundArchive = FolderManager.findFolderInAccount(mockAccount, 'Archive');
  assert.ok(foundArchive);
  assert.strictEqual(foundArchive.id, 'f_archive');

  const foundNested = FolderManager.findFolderInAccount(mockAccount, 'Archive/2025');
  assert.ok(foundNested);
  assert.strictEqual(foundNested.id, 'f_2025');

  // Test finding non-existing folder
  const notFound = FolderManager.findFolderInAccount(mockAccount, '專案A/發票');
  assert.strictEqual(notFound, null);
  console.log('✓ 存在與不存在資料夾比對正確');

  // Test rule folder checking
  const testRules = [
    { name: 'R1', targetFolders: ['Archive'] },
    { name: 'R2', targetFolders: ['專案A/發票', 'Archive/2025'] },
    { name: 'R3', targetFolders: ['客戶信件'] }
  ];

  const checkResult = FolderManager.checkRuleFolders(testRules, mockAccount);
  assert.deepStrictEqual(checkResult.existing.sort(), ['Archive', 'Archive/2025'].sort());
  assert.deepStrictEqual(checkResult.missing.sort(), ['專案A/發票', '客戶信件'].sort());
  console.log('✓ 規則內參照資料夾比對與歸類正確');

  // Test automatic folder creation
  let createdCalls = [];
  const mockMessenger = {
    folders: {
      async create(parentId, name) {
        createdCalls.push({ parentId, name });
        return {
          id: `new_${name}`,
          name: name,
          path: `/${name}`,
          subFolders: []
        };
      }
    }
  };

  const fm = new FolderManager(mockMessenger);
  const logs = [];
  const result = await fm.ensureFoldersExist(
    ['專案A/發票', '客戶信件', 'Archive'],
    mockAccount,
    (msg) => logs.push(msg)
  );

  // 'Archive' already existed -> created: false
  assert.strictEqual(result.get('Archive').created, false);
  assert.ok(result.get('Archive').uri.includes('Archive'));

  // '客戶信件' did not exist -> created: true
  assert.strictEqual(result.get('客戶信件').created, true);

  // '專案A/發票' did not exist -> created 2 levels: '專案A' then '發票'
  assert.strictEqual(result.get('專案A/發票').created, true);

  // Verify createdCalls contains '專案A', '發票', '客戶信件'
  const createdNames = createdCalls.map(c => c.name);
  assert.ok(createdNames.includes('專案A'));
  assert.ok(createdNames.includes('發票'));
  assert.ok(createdNames.includes('客戶信件'));

  console.log('✓ 缺失資料夾之遞迴自動建立與 URI 生成完全符合要求');
}

runTests().catch(err => {
  console.error(err);
  process.exit(1);
});
