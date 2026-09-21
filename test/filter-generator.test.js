/**
 * Unit tests for FilterGenerator
 */

const assert = require('assert');
const { FilterGenerator } = require('../lib/filter-generator.js');

function runTests() {
  console.log('--- 測試 FilterGenerator 篩選器生成 ---');

  // Test condition string building
  {
    const rule = {
      conditions: [
        { type: 'from', values: ['boss@example.com'] },
        { type: 'subject', values: ['重要', '急件'] },
        { type: 'has_attachment' }
      ]
    };
    const condStr = FilterGenerator.buildConditionString(rule);
    assert.ok(condStr.includes('(from,contains,boss@example.com)'));
    assert.ok(condStr.includes('(subject,contains,重要)'));
    assert.ok(condStr.includes('(status,is,hasAttachment)'));
    console.log('✓ 條件字串語法建構正確:', condStr);
  }

  // Test condition string with spaces
  {
    const rule = {
      conditions: [
        { type: 'subject', values: ['重要 文件 (急件)'] }
      ]
    };
    const condStr = FilterGenerator.buildConditionString(rule);
    assert.ok(condStr.includes('(subject,contains,"重要 文件 (急件)")'));
    console.log('✓ 特殊字元與空白字串引號轉義正確:', condStr);
  }

  // Test msgFilterRules.dat generation
  {
    const rules = [
      {
        name: '發票整理',
        enabled: true,
        conditions: [
          { type: 'subject', values: ['發票'] }
        ],
        actions: [
          { type: 'move_to_folder', folderName: '財務/發票' },
          { type: 'mark_read' }
        ]
      }
    ];

    const folderMap = new Map();
    folderMap.set('財務/發票', {
      uri: 'mailbox://nobody@Local%20Folders/%E8%B2%A1%E5%8B%99%2F%E7%99%BC%E7%A5%A8'
    });

    const datContent = FilterGenerator.generateMsgFilterRulesDat(rules, folderMap);

    assert.ok(datContent.includes('version="9"'));
    assert.ok(datContent.includes('logging="yes"'));
    assert.ok(datContent.includes('name="發票整理"'));
    assert.ok(datContent.includes('enabled="yes"'));
    assert.ok(datContent.includes('type="17"'));
    assert.ok(datContent.includes('action="Move to folder"'));
    assert.ok(datContent.includes('actionValue="mailbox://nobody@Local%20Folders/%E8%B2%A1%E5%8B%99%2F%E7%99%BC%E7%A5%A8"'));
    assert.ok(datContent.includes('action="Mark read"'));
    assert.ok(datContent.includes('condition="AND (subject,contains,發票)"'));

    console.log('✓ msgFilterRules.dat 檔案格式生成完全符合 Thunderbird 規範');
  }

  // Test toExperimentPayload
  {
    const rules = [
      {
        name: '移動歸檔',
        enabled: true,
        conditions: [],
        actions: [
          { type: 'copy_to_folder', folderName: 'Archive' },
          { type: 'stop_execution' }
        ]
      }
    ];
    const folderMap = new Map();
    folderMap.set('Archive', { uri: 'imap://user@server/Archive' });

    const payload = FilterGenerator.toExperimentPayload(rules, folderMap);
    assert.strictEqual(payload.length, 1);
    assert.strictEqual(payload[0].name, '移動歸檔');
    assert.strictEqual(payload[0].condition, 'ALL');
    assert.strictEqual(payload[0].actions.length, 2);
    assert.strictEqual(payload[0].actions[0].type, 16); // CopyToFolder
    assert.strictEqual(payload[0].actions[0].targetFolderUri, 'imap://user@server/Archive');
    assert.strictEqual(payload[0].actions[1].type, 11); // StopExecution

    console.log('✓ Experiment API 載荷轉換正確');
  }
}

runTests();
