/**
 * RWZ Binary Stream Reader
 * Cross-environment (WebExtension / Browser / Node.js) binary reader.
 */

class BinaryStreamReader {
  /**
   * @param {ArrayBuffer|Uint8Array|Buffer} buffer
   */
  constructor(buffer) {
    if (typeof Buffer !== 'undefined' && Buffer.isBuffer(buffer)) {
      this.buffer = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    } else if (buffer instanceof ArrayBuffer) {
      this.buffer = new Uint8Array(buffer);
    } else if (ArrayBuffer.isView(buffer)) {
      this.buffer = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    } else {
      throw new TypeError('Expected ArrayBuffer, Uint8Array, or Buffer');
    }
    this.view = new DataView(this.buffer.buffer, this.buffer.byteOffset, this.buffer.byteLength);
    this.offset = 0;
  }

  get length() {
    return this.buffer.byteLength;
  }

  get remaining() {
    return this.buffer.byteLength - this.offset;
  }

  checkBounds(size) {
    if (this.offset + size > this.buffer.byteLength) {
      throw new Error(`Buffer underrun: requested ${size} bytes at offset ${this.offset}, total length ${this.buffer.byteLength}`);
    }
  }

  readUInt8() {
    this.checkBounds(1);
    const val = this.view.getUint8(this.offset);
    this.offset += 1;
    return val;
  }

  readUInt16() {
    this.checkBounds(2);
    const val = this.view.getUint16(this.offset, true); // Little-Endian
    this.offset += 2;
    return val;
  }

  readUInt32() {
    this.checkBounds(4);
    const val = this.view.getUint32(this.offset, true); // Little-Endian
    this.offset += 4;
    return val;
  }

  readUInt64() {
    this.checkBounds(8);
    const val = this.view.getBigUint64(this.offset, true); // Little-Endian
    this.offset += 8;
    return val;
  }

  readDouble() {
    this.checkBounds(8);
    const val = this.view.getFloat64(this.offset, true); // Little-Endian
    this.offset += 8;
    return val;
  }

  readBytes(len) {
    if (len < 0) throw new Error(`Negative byte length ${len}`);
    this.checkBounds(len);
    const slice = this.buffer.slice(this.offset, this.offset + len);
    this.offset += len;
    return slice;
  }

  readAsciiString(len) {
    const bytes = this.readBytes(len);
    let str = '';
    for (let i = 0; i < bytes.length; i++) {
      if (bytes[i] === 0) break;
      str += String.fromCharCode(bytes[i]);
    }
    return str;
  }

  readAsciiUntilNullTerminator(maxLen = 512) {
    let str = '';
    let count = 0;
    while (this.offset < this.buffer.byteLength && count < maxLen) {
      const b = this.readUInt8();
      count++;
      if (b === 0) break;
      str += String.fromCharCode(b);
    }
    return str;
  }

  readString(charCount) {
    // UTF-16LE characters
    let str = '';
    for (let i = 0; i < charCount; i++) {
      const code = this.readUInt16();
      str += String.fromCharCode(code);
    }
    return str;
  }

  readStringObject() {
    let length = this.readUInt8();
    if (length === 0xff) {
      length = this.readUInt16();
      this.offset += 2; // 2-byte pad
    }
    return this.readString(length);
  }

  readStringUntilNullTerminator(maxChars = 260) {
    let str = '';
    let count = 0;
    while (this.offset + 2 <= this.buffer.byteLength && count < maxChars) {
      const code = this.readUInt16();
      count++;
      if (code === 0) break;
      str += String.fromCharCode(code);
    }
    return str;
  }

  peekUInt32() {
    this.checkBounds(4);
    return this.view.getUint32(this.offset, true);
  }

  peekUInt16() {
    this.checkBounds(2);
    return this.view.getUint16(this.offset, true);
  }
}

if (typeof globalThis !== 'undefined') {
  globalThis.BinaryStreamReader = BinaryStreamReader;
}
if (typeof window !== 'undefined') {
  window.BinaryStreamReader = BinaryStreamReader;
}
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { BinaryStreamReader };
}
