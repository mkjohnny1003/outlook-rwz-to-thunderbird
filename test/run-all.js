/**
 * Master test runner for Outlook RWZ to Thunderbird Filter Add-on
 */

const { spawnSync } = require('child_process');
const path = require('path');

const tests = [
  'rwz-parser.test.js',
  'folder-manager.test.js',
  'filter-generator.test.js',
  'real-uri.test.js'
];

let failed = false;

console.log('========================================================');
console.log('🚀 開始執行 Outlook RWZ to Thunderbird Add-on 單元測試');
console.log('========================================================\n');

for (const t of tests) {
  const testPath = path.join(__dirname, t);
  console.log(`▶ 執行測試檔: ${t}`);
  const res = spawnSync(process.execPath, [testPath], { stdio: 'inherit' });
  if (res.status !== 0) {
    failed = true;
    console.error(`❌ 測試失敗: ${t}\n`);
  } else {
    console.log(`✅ 測試成功: ${t}\n`);
  }
}

if (failed) {
  console.error('💥 部分單元測試未通過！');
  process.exit(1);
} else {
  console.log('========================================================');
  console.log('🎉 所有單元測試皆順利通過！');
  console.log('========================================================');
}
