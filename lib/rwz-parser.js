/**
 * Outlook Rules Wizard (.rwz) Parser
 * Pure JavaScript parser capable of extracting rules, conditions, actions, and folder targets
 * from binary .rwz files exported by Microsoft Outlook.
 */

const BinaryStreamReader = (typeof module !== 'undefined' && module.exports)
  ? require('./rwz-binary-reader.js').BinaryStreamReader
  : ((typeof globalThis !== 'undefined' && globalThis.BinaryStreamReader) ||
     (typeof window !== 'undefined' && window.BinaryStreamReader) ||
     null);

const MAPI_PROPERTIES = {
  0x0037: 'Subject',
  0x0042: 'SentRepresentingName',
  0x0c15: 'RecipientType',
  0x3001: 'DisplayName',
  0x3002: 'AddressType',
  0x3003: 'EmailAddress',
  0x39fe: 'SmtpAddress',
  0x3a00: 'Account',
  0x3a20: 'TransmittableDisplayName'
};

const OXCDATA = {
  PtypInteger16: 2,
  PtypInteger32: 3,
  PtypFloating32: 4,
  PtypFloating64: 5,
  PtypBoolean: 0xb,
  PtypInteger64: 0x14,
  PtypString8: 0x1e,
  PtypString: 0x1f,
  PtypTime: 0x40,
  PtypGuid: 0x48,
  PtypBinary: 0x102
};

class PropertyValueArray {
  constructor(reader) {
    this.properties = {};
    const unknown = reader.readUInt32();
    const nProps = reader.readUInt32();
    const propDataSize = reader.readUInt32();
    const startPos = reader.offset;
    const endPos = startPos + propDataSize;

    const headers = [];
    for (let i = 0; i < nProps; i++) {
      const dataType = reader.readUInt16();
      const id = reader.readUInt16();
      const data = [reader.readUInt32(), reader.readUInt32(), reader.readUInt32()];
      headers.push({ dataType, id, data });
    }

    for (const ph of headers) {
      const curPos = reader.offset;
      let val = undefined;
      const offsetInBlob = ph.data[1];

      switch (ph.dataType) {
        case OXCDATA.PtypInteger32:
          val = ph.data[1];
          break;
        case OXCDATA.PtypBoolean:
          val = ph.data[1] !== 0;
          break;
        case OXCDATA.PtypString:
          if (offsetInBlob >= 0 && startPos + offsetInBlob < endPos) {
            reader.offset = startPos + offsetInBlob;
            val = reader.readStringUntilNullTerminator();
            reader.offset = curPos;
          }
          break;
        case OXCDATA.PtypString8:
          if (offsetInBlob >= 0 && startPos + offsetInBlob < endPos) {
            reader.offset = startPos + offsetInBlob;
            val = reader.readAsciiUntilNullTerminator();
            reader.offset = curPos;
          }
          break;
        case OXCDATA.PtypBinary:
          const byteLen = ph.data[2];
          if (offsetInBlob >= 0 && startPos + offsetInBlob + byteLen <= endPos) {
            reader.offset = startPos + offsetInBlob;
            val = Array.from(reader.readBytes(byteLen))
              .map(b => b.toString(16).padStart(2, '0'))
              .join('');
            reader.offset = curPos;
          }
          break;
        default:
          val = ph.data[1];
          break;
      }

      const propName = MAPI_PROPERTIES[ph.id] || `0x${ph.id.toString(16).padStart(4, '0')}`;
      this.properties[propName] = val;
    }

    reader.offset = endPos;
  }

  getBestRecipientString() {
    return this.properties.SmtpAddress ||
           this.properties.EmailAddress ||
           this.properties.DisplayName ||
           this.properties.TransmittableDisplayName ||
           '';
  }
}

class RwzParser {
  /**
   * Parse an Outlook RWZ binary buffer into structured rules.
   * @param {ArrayBuffer|Uint8Array|Buffer} buffer
   * @returns {{ version: string, numberOfRules: number, rules: Array }}
   */
  static parse(buffer) {
    const Reader = BinaryStreamReader ||
      (typeof globalThis !== 'undefined' && globalThis.BinaryStreamReader) ||
      (typeof window !== 'undefined' && window.BinaryStreamReader);
    if (!Reader) {
      throw new Error('找不到 BinaryStreamReader 元件，請確認腳本已正確載入。');
    }
    const reader = new Reader(buffer);
    const result = {
      version: 'unknown',
      numberOfRules: 0,
      rules: []
    };

    if (reader.length < 24) {
      throw new Error('檔案太小，不是有效的 Outlook 規則檔 (.rwz)');
    }

    // 1. Header
    const sig = reader.peekUInt32();
    if (sig === 1310720) result.version = 'outlook2019';
    else if (sig === 1200000) result.version = 'outlook2007';
    else if (sig === 1100000) result.version = 'outlook2003';
    else if (sig === 1000000) result.version = 'outlook2002';
    else if (sig === 980413) result.version = 'outlook2000';
    else if (sig === 970812) result.version = 'outlook98';
    else result.version = 'outlook-generic';

    reader.readUInt32(); // Signature
    reader.readUInt32(); // Flags

    // unknown 1..8
    for (let i = 1; i <= 8; i++) {
      reader.readUInt32();
    }
    // unknown 9
    reader.readUInt32();

    result.numberOfRules = reader.readUInt16();
    reader.readUInt16(); // extra padding

    // 2. Rules
    for (let r = 0; r < result.numberOfRules; r++) {
      const rule = RwzParser.parseRule(reader, r, result.numberOfRules);
      result.rules.push(rule);

      if (r !== result.numberOfRules - 1 && reader.remaining >= 2) {
        const sep = reader.readUInt16();
        if (sep !== 0) {
          // Soft tolerate
        }
      }
    }

    return result;
  }

  static parseRule(reader, index, totalRules) {
    const rawRule = {
      index,
      signature: reader.readUInt16(),
      name: reader.readStringObject() || `未命名規則 ${index + 1}`,
      enabled: reader.readUInt32() !== 0,
      elements: []
    };

    for (let i = 0; i < 4; i++) reader.readUInt32(); // unknown[0..3]
    const dataSize = reader.readUInt32();
    const nElements = reader.readUInt16();
    const separator = reader.readUInt16();

    if (separator === 0xffff) {
      reader.readUInt16(); // padding
      const classNameLen = reader.readUInt16();
      reader.readAsciiString(classNameLen); // CRuleElement
    }

    for (let e = 0; e < nElements; e++) {
      const elem = RwzParser.parseRuleElement(reader);
      rawRule.elements.push(elem);

      if (e !== nElements - 1 && reader.remaining >= 2) {
        reader.readUInt16(); // separator 0x8001
      }
    }

    return RwzParser.normalizeRule(rawRule);
  }

  static parseRuleElement(reader) {
    const elemId = reader.readUInt32();
    const elem = { id: elemId, data: {} };

    switch (elemId) {
      case 0x64: // Unknown mandatory element
      case 0x190: { // Apply rule element
        elem.data.extended = reader.readUInt32();
        elem.data.reserved = reader.readUInt32();
        elem.data.flags = reader.readUInt32();
        break;
      }

      // Simple condition elements
      case 0xc8: // in To box
      case 0xc9: // sent only to me
      case 0xca: // not in To box
      case 0xdc: // automatic reply
      case 0xde: // has attachment
      case 0xe2: // in Cc box
      case 0xe3: // in To or Cc box
      case 0xe9: // exception list
      case 0xeb: // junk email
      case 0xec: // adult content
      case 0xf1: // meeting invitation
      case 0xf2: // from contacts
      case 0xf6: // assigned to any category
      case 0xf7: // from any RSS feed
      case 0x12d: // delete it
      case 0x132: // clear message flag
      case 0x13a: // notify read
      case 0x13b: // notify delivered
      case 0x142: // stop processing more rules
      case 0x143: // do not search for junk
      case 0x148: // print it
      case 0x14a: // permanently delete
      case 0x14c: // mark as read
      case 0x14f: // desktop alert
      case 0x150: // color flag
      case 0x152: // clear categories
      case 0x1f4: // except in To box
      case 0x1f5: // except sent only to me
      case 0x1f6: // except not in To box
      {
        elem.data.extended = reader.readUInt32();
        break;
      }

      // People / Recipient list elements
      case 0xcb: // from people or group
      case 0xcc: // sent to people or group
      case 0x12e: // forward to
      case 0x13c: // Cc to
      case 0x144: // redirect to
      case 0x147: // forward as attachment
      case 0x1f7: // except if from
      case 0x1f8: // except if sent to
      {
        elem.data.extended = reader.readUInt32();
        elem.data.reserved = reader.readUInt32();
        const nValues = reader.readUInt32();
        elem.data.recipients = [];
        for (let i = 0; i < nValues; i++) {
          const propArray = new PropertyValueArray(reader);
          const bestStr = propArray.getBestRecipientString();
          if (bestStr) elem.data.recipients.push(bestStr);
        }
        if (reader.remaining >= 8) {
          reader.readUInt32(); // unknown1 (1)
          reader.readUInt32(); // unknown2 (0)
        }
        break;
      }

      // Strings list elements (Subject, Body, Address, Header)
      case 0xcd: // subject contains
      case 0xce: // body contains
      case 0xcf: // subject or body contains
      case 0xe5: // recipient address contains
      case 0xe6: // sender address contains
      case 0xe8: // message header contains
      case 0xf5: // RSS title contains
      case 0x1f9: // except subject contains
      case 0x1fa: // except body contains
      case 0x1fb: // except subject/body contains
      {
        const nEntries = reader.readUInt32();
        elem.data.words = [];
        for (let i = 0; i < nEntries; i++) {
          const flags = reader.readUInt32();
          const word = reader.readStringObject();
          if (word) elem.data.words.push(word);
        }
        break;
      }

      // Move / Copy to folder
      case 0x12c: // move to folder
      case 0x139: // copy to folder
      {
        elem.data.extended = reader.readUInt32();
        elem.data.reserved = reader.readUInt32();

        // FolderEntryId (FlatEntry: size + payload)
        const folderEntrySize = reader.readUInt32();
        if (folderEntrySize > 0) {
          reader.readBytes(folderEntrySize);
        }

        // StoreEntryId (FlatEntry: size + payload)
        const storeEntrySize = reader.readUInt32();
        if (storeEntrySize > 0) {
          reader.readBytes(storeEntrySize);
        }

        elem.data.folderName = reader.readStringObject();
        elem.data.secondaryUserStore = reader.readUInt32() !== 0;
        break;
      }

      // Importance / Priority
      case 0xd2: // condition: marked as importance
      case 0x137: // action: mark as importance
      {
        elem.data.extended = reader.readUInt32();
        elem.data.reserved = reader.readUInt32();
        elem.data.importance = reader.readUInt32(); // 0: Low, 1: Normal, 2: High
        break;
      }

      // Categories
      case 0xd7:
      case 0x133: {
        elem.data.extended = reader.readUInt32();
        elem.data.reserved = reader.readUInt32();
        elem.data.categories = reader.readStringObject();
        break;
      }

      // Flagged for action
      case 0xd0:
      case 0x151: {
        elem.data.extended = reader.readUInt32();
        elem.data.reserved = reader.readUInt32();
        elem.data.actionName = reader.readStringObject();
        break;
      }

      default: {
        // Fallback for unrecognized elements: try reading common extended header
        try {
          if (reader.remaining >= 4) {
            elem.data.unknown = reader.readUInt32();
          }
        } catch (_) {}
        break;
      }
    }

    return elem;
  }

  /**
   * Convert raw elements into a high-level clean rule model
   */
  static normalizeRule(raw) {
    const rule = {
      name: raw.name,
      enabled: raw.enabled,
      conditions: [],
      actions: [],
      exceptions: [],
      targetFolders: []
    };

    for (const elem of raw.elements) {
      switch (elem.id) {
        // --- Conditions ---
        case 0xc8:
          rule.conditions.push({ type: 'name_in_to', desc: '收件者包含自己' });
          break;
        case 0xc9:
          rule.conditions.push({ type: 'to_me_only', desc: '僅傳送給我' });
          break;
        case 0xca:
          rule.conditions.push({ type: 'name_not_in_to', desc: '收件者不包含自己' });
          break;
        case 0xcb:
          if (elem.data.recipients && elem.data.recipients.length > 0) {
            rule.conditions.push({
              type: 'from',
              op: 'contains',
              values: elem.data.recipients,
              desc: `寄件者為: ${elem.data.recipients.join(', ')}`
            });
          }
          break;
        case 0xcc:
          if (elem.data.recipients && elem.data.recipients.length > 0) {
            rule.conditions.push({
              type: 'to',
              op: 'contains',
              values: elem.data.recipients,
              desc: `收件者包含: ${elem.data.recipients.join(', ')}`
            });
          }
          break;
        case 0xcd:
          if (elem.data.words && elem.data.words.length > 0) {
            rule.conditions.push({
              type: 'subject',
              op: 'contains',
              values: elem.data.words,
              desc: `主旨包含: ${elem.data.words.join(' 或 ')}`
            });
          }
          break;
        case 0xce:
          if (elem.data.words && elem.data.words.length > 0) {
            rule.conditions.push({
              type: 'body',
              op: 'contains',
              values: elem.data.words,
              desc: `內文包含: ${elem.data.words.join(' 或 ')}`
            });
          }
          break;
        case 0xcf:
          if (elem.data.words && elem.data.words.length > 0) {
            rule.conditions.push({
              type: 'subject_or_body',
              op: 'contains',
              values: elem.data.words,
              desc: `主旨或內文包含: ${elem.data.words.join(' 或 ')}`
            });
          }
          break;
        case 0xd2: {
          const impLevels = ['低', '一般', '高'];
          const tbImp = ['Low', 'Normal', 'High'][elem.data.importance] || 'Normal';
          rule.conditions.push({
            type: 'priority',
            value: tbImp,
            desc: `重要性為: ${impLevels[elem.data.importance] || tbImp}`
          });
          break;
        }
        case 0xde:
          rule.conditions.push({ type: 'has_attachment', desc: '含有附件' });
          break;
        case 0xe2:
          rule.conditions.push({ type: 'name_in_cc', desc: '副本 (Cc) 包含自己' });
          break;
        case 0xe3:
          rule.conditions.push({ type: 'name_in_to_or_cc', desc: '收件者或副本包含自己' });
          break;
        case 0xe6:
          if (elem.data.words && elem.data.words.length > 0) {
            rule.conditions.push({
              type: 'from',
              op: 'contains',
              values: elem.data.words,
              desc: `寄件者地址包含: ${elem.data.words.join(' 或 ')}`
            });
          }
          break;
        case 0xe5:
          if (elem.data.words && elem.data.words.length > 0) {
            rule.conditions.push({
              type: 'to',
              op: 'contains',
              values: elem.data.words,
              desc: `收件者地址包含: ${elem.data.words.join(' 或 ')}`
            });
          }
          break;

        // --- Actions ---
        case 0x12c:
          if (elem.data.folderName) {
            rule.actions.push({
              type: 'move_to_folder',
              folderName: elem.data.folderName,
              desc: `移動至資料夾: ${elem.data.folderName}`
            });
            rule.targetFolders.push(elem.data.folderName);
          }
          break;
        case 0x139:
          if (elem.data.folderName) {
            rule.actions.push({
              type: 'copy_to_folder',
              folderName: elem.data.folderName,
              desc: `複製至資料夾: ${elem.data.folderName}`
            });
            rule.targetFolders.push(elem.data.folderName);
          }
          break;
        case 0x12d:
        case 0x14a:
          rule.actions.push({ type: 'delete', desc: '刪除郵件' });
          break;
        case 0x14c:
          rule.actions.push({ type: 'mark_read', desc: '標示為已讀' });
          break;
        case 0x150:
        case 0x151:
          rule.actions.push({ type: 'mark_flagged', desc: '標示星號/旗標' });
          break;
        case 0x142:
          rule.actions.push({ type: 'stop_execution', desc: '停止套用其他規則' });
          break;
        case 0x12e:
          if (elem.data.recipients && elem.data.recipients.length > 0) {
            rule.actions.push({
              type: 'forward',
              recipients: elem.data.recipients,
              desc: `轉寄至: ${elem.data.recipients.join(', ')}`
            });
          }
          break;
        case 0x137: {
          const impLevels = ['低', '一般', '高'];
          const tbImp = ['Lowest', 'Normal', 'Highest'][elem.data.importance] || 'Normal';
          rule.actions.push({
            type: 'change_priority',
            priority: tbImp,
            desc: `變更優先順序為: ${impLevels[elem.data.importance] || tbImp}`
          });
          break;
        }

        // --- Exceptions ---
        case 0x1f7:
          if (elem.data.recipients) {
            rule.exceptions.push({
              type: 'from',
              values: elem.data.recipients,
              desc: `例外：寄件者不是 ${elem.data.recipients.join(', ')}`
            });
          }
          break;
        case 0x1f9:
          if (elem.data.words) {
            rule.exceptions.push({
              type: 'subject',
              values: elem.data.words,
              desc: `例外：主旨不包含 ${elem.data.words.join(', ')}`
            });
          }
          break;
      }
    }

    return rule;
  }
}

if (typeof globalThis !== 'undefined') {
  globalThis.RwzParser = RwzParser;
  globalThis.PropertyValueArray = PropertyValueArray;
}
if (typeof window !== 'undefined') {
  window.RwzParser = RwzParser;
  window.PropertyValueArray = PropertyValueArray;
}
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { RwzParser, PropertyValueArray, MAPI_PROPERTIES, OXCDATA };
}
