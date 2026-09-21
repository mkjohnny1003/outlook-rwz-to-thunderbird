/**
 * Unit tests for BinaryStreamReader and RwzParser
 */

const assert = require('assert');
const { BinaryStreamReader } = require('../lib/rwz-binary-reader.js');
const { RwzParser } = require('../lib/rwz-parser.js');

// Binary builder helpers for synthetic RWZ test generation
function u32(val) {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(val >>> 0, 0);
  return b;
}

function u16(val) {
  const b = Buffer.alloc(2);
  b.writeUInt16LE(val & 0xffff, 0);
  return b;
}

function u8(val) {
  return Buffer.from([val & 0xff]);
}

function buildStringObject(str) {
  const parts = [];
  if (str.length < 0xff) {
    parts.push(u8(str.length));
  } else {
    parts.push(u8(0xff));
    parts.push(u16(str.length));
    parts.push(u16(0)); // 2-byte pad
  }
  for (let i = 0; i < str.length; i++) {
    parts.push(u16(str.charCodeAt(i)));
  }
  return Buffer.concat(parts);
}

function buildOutlook2019Header(nRules) {
  return Buffer.concat([
    u32(1310720),    // signature = outlook2019
    u32(0x06140000), // flags
    u32(0), u32(0), u32(0), // unknown 1..3
    u32(0), u32(0), u32(0), u32(0), // unknown 4..7
    u32(1),          // unknown 8
    u32(0),          // unknown 9
    u16(nRules),     // numberOfRules
    u16(0)           // extra
  ]);
}

function buildRuleHeader(name, nElements, index) {
  const parts = [
    u16(0x0001),              // signature
    buildStringObject(name),  // name
    u32(1),                   // enabled = true
    u32(0), u32(0), u32(0), u32(0), // unknown[0..3]
    u32(0),                   // dataSize
    u16(nElements),           // nRuleElements
    u16(index === 0 ? 0xffff : 0x8001) // separator
  ];

  if (index === 0) {
    parts.push(u16(0)); // padding
    const className = 'CRuleElement';
    parts.push(u16(className.length));
    parts.push(Buffer.from(className, 'ascii'));
  }

  return Buffer.concat(parts);
}

function buildRuleElement(elementId, dataPayload) {
  return Buffer.concat([u32(elementId), dataPayload]);
}

function buildStringsListElement(elementId, words) {
  const parts = [
    u32(words.length)
  ];
  for (const w of words) {
    parts.push(u32(0)); // flags
    parts.push(buildStringObject(w));
  }
  return buildRuleElement(elementId, Buffer.concat(parts));
}

function buildMoveToFolderElement(folderName) {
  const parts = [
    u32(1), // extended
    u32(0), // reserved
    u32(0), // folderEntryId size = 0
    u32(0), // storeEntryId size = 0
    buildStringObject(folderName),
    u32(0)  // secondaryUserStore = false
  ];
  return buildRuleElement(0x12c, Buffer.concat(parts));
}

function runTests() {
  console.log('--- 測試 BinaryStreamReader ---');
  {
    const buf = Buffer.concat([
      u8(42),
      u16(0x1234),
      u32(0xdeadbeef),
      Buffer.from('Hello\0World', 'ascii')
    ]);
    const reader = new BinaryStreamReader(buf);
    assert.strictEqual(reader.readUInt8(), 42);
    assert.strictEqual(reader.readUInt16(), 0x1234);
    assert.strictEqual(reader.readUInt32(), 0xdeadbeef);
    assert.strictEqual(reader.readAsciiUntilNullTerminator(), 'Hello');
    console.log('✓ BinaryStreamReader 基礎數值與 ASCII 讀取通過');
  }

  {
    const str = '測試主旨';
    const strObjBuf = buildStringObject(str);
    const reader = new BinaryStreamReader(strObjBuf);
    assert.strictEqual(reader.readStringObject(), str);
    console.log('✓ BinaryStreamReader UTF-16LE StringObject 讀取通過');
  }

  console.log('--- 測試 RwzParser 規則檔解析 ---');
  {
    // Build a synthetic RWZ with 2 rules
    // Rule 1: Subject contains "發票" -> Move to folder "財務/發票"
    // Rule 2: Sent only to me -> Mark as read
    const rule1Elements = [
      buildRuleElement(0x64, Buffer.concat([u32(1), u32(0), u32(1)])),
      buildRuleElement(0x190, Buffer.concat([u32(1), u32(0), u32(1)])),
      buildStringsListElement(0xcd, ['發票', '帳單']),
      buildMoveToFolderElement('財務/發票')
    ];

    const rule2Elements = [
      buildRuleElement(0x64, Buffer.concat([u32(1), u32(0), u32(1)])),
      buildRuleElement(0x190, Buffer.concat([u32(1), u32(0), u32(1)])),
      buildRuleElement(0xc9, u32(0)), // sent only to me
      buildRuleElement(0x14c, u32(0)) // mark as read
    ];

    const parts = [
      buildOutlook2019Header(2),
      buildRuleHeader('發票自動分類', rule1Elements.length, 0)
    ];

    rule1Elements.forEach((el, idx) => {
      parts.push(el);
      if (idx !== rule1Elements.length - 1) parts.push(u16(0x8001));
    });

    parts.push(u16(0)); // rule separator

    parts.push(buildRuleHeader('重要私信標示已讀', rule2Elements.length, 1));
    rule2Elements.forEach((el, idx) => {
      parts.push(el);
      if (idx !== rule2Elements.length - 1) parts.push(u16(0x8001));
    });

    const rwzBuffer = Buffer.concat(parts);
    const parsed = RwzParser.parse(rwzBuffer);

    assert.strictEqual(parsed.version, 'outlook2019');
    assert.strictEqual(parsed.numberOfRules, 2);
    assert.strictEqual(parsed.rules.length, 2);

    // Verify Rule 1
    const r1 = parsed.rules[0];
    assert.strictEqual(r1.name, '發票自動分類');
    assert.strictEqual(r1.enabled, true);
    assert.strictEqual(r1.conditions.length, 1);
    assert.strictEqual(r1.conditions[0].type, 'subject');
    assert.deepStrictEqual(r1.conditions[0].values, ['發票', '帳單']);
    assert.strictEqual(r1.actions.length, 1);
    assert.strictEqual(r1.actions[0].type, 'move_to_folder');
    assert.strictEqual(r1.actions[0].folderName, '財務/發票');
    assert.deepStrictEqual(r1.targetFolders, ['財務/發票']);

    // Verify Rule 2
    const r2 = parsed.rules[1];
    assert.strictEqual(r2.name, '重要私信標示已讀');
    assert.strictEqual(r2.conditions.length, 1);
    assert.strictEqual(r2.conditions[0].type, 'to_me_only');
    assert.strictEqual(r2.actions.length, 1);
    assert.strictEqual(r2.actions[0].type, 'mark_read');

    console.log('✓ RwzParser 規則解析與規格正規化完全正確');
  }
}

runTests();
