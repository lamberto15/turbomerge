import { crc32, deflateRawSync, inflateRawSync } from 'node:zlib';

/**
 * Minimal ZIP reader/writer on Node's native zlib (C deflate + zlib.crc32).
 * Entries that are not modified are copied in their original compressed form, so a render only
 * pays for compressing the XML parts it actually changed.
 */

export interface ZipEntry {
  readonly name: string;
  readonly method: number; // 0 = stored, 8 = deflate
  readonly crc: number;
  readonly size: number;
  readonly compressed: Uint8Array;
}

const EOCD = 0x06054b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;
const decoder = new TextDecoder();
const encoder = new TextEncoder();

export function readZip(bytes: Uint8Array): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('end of central directory not found');
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  if (count === 0xffff || offset === 0xffffffff) throw new Error('ZIP64 archives are not supported');

  const entries: ZipEntry[] = [];
  for (let n = 0; n < count; n++) {
    if (view.getUint32(offset, true) !== CENTRAL) throw new Error('corrupt central directory');
    const method = view.getUint16(offset + 10, true);
    const crc = view.getUint32(offset + 16, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const size = view.getUint32(offset + 24, true);
    const nameLen = view.getUint16(offset + 28, true);
    const extraLen = view.getUint16(offset + 30, true);
    const commentLen = view.getUint16(offset + 32, true);
    const local = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLen));
    if (view.getUint32(local, true) !== LOCAL) throw new Error(`corrupt local header for ${name}`);
    if (method !== 0 && method !== 8) throw new Error(`unsupported compression method ${method} for ${name}`);
    const dataStart = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    entries.push({ name, method, crc, size, compressed: bytes.subarray(dataStart, dataStart + compressedSize) });
    offset += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

export function inflate(entry: ZipEntry): Uint8Array {
  if (entry.method === 0) return entry.compressed;
  const out = inflateRawSync(entry.compressed);
  return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
}

export function makeEntry(name: string, data: Uint8Array, store: boolean): ZipEntry {
  const crc = crc32(data);
  if (store || data.length < 64) return { name, method: 0, crc, size: data.length, compressed: data };
  const out = deflateRawSync(data, { level: 6 });
  return {
    name,
    method: 8,
    crc,
    size: data.length,
    compressed: new Uint8Array(out.buffer, out.byteOffset, out.byteLength),
  };
}

// DOS date 1980-01-01 00:00:00: fixed so identical input gives byte-identical output.
const DOS_TIME = 0;
const DOS_DATE = (0 << 9) | (1 << 5) | 1;

export function writeZip(entries: readonly ZipEntry[]): Uint8Array {
  const names = entries.map((e) => encoder.encode(e.name));
  let localSize = 0;
  let centralSize = 0;
  entries.forEach((e, i) => {
    localSize += 30 + names[i]!.length + e.compressed.length;
    centralSize += 46 + names[i]!.length;
  });
  if (entries.length > 0xfffe || localSize + centralSize > 0xfffffffe)
    throw new Error('output too large for ZIP (ZIP64 unsupported)');

  const out = new Uint8Array(localSize + centralSize + 22);
  const view = new DataView(out.buffer);
  let pos = 0;
  const offsets: number[] = [];
  entries.forEach((e, i) => {
    const name = names[i]!;
    offsets.push(pos);
    view.setUint32(pos, LOCAL, true);
    view.setUint16(pos + 4, 20, true);
    view.setUint16(pos + 6, 0x0800, true); // UTF-8 names
    view.setUint16(pos + 8, e.method, true);
    view.setUint16(pos + 10, DOS_TIME, true);
    view.setUint16(pos + 12, DOS_DATE, true);
    view.setUint32(pos + 14, e.crc, true);
    view.setUint32(pos + 18, e.compressed.length, true);
    view.setUint32(pos + 22, e.size, true);
    view.setUint16(pos + 26, name.length, true);
    view.setUint16(pos + 28, 0, true);
    out.set(name, pos + 30);
    out.set(e.compressed, pos + 30 + name.length);
    pos += 30 + name.length + e.compressed.length;
  });
  const centralStart = pos;
  entries.forEach((e, i) => {
    const name = names[i]!;
    view.setUint32(pos, CENTRAL, true);
    view.setUint16(pos + 4, 20, true);
    view.setUint16(pos + 6, 20, true);
    view.setUint16(pos + 8, 0x0800, true);
    view.setUint16(pos + 10, e.method, true);
    view.setUint16(pos + 12, DOS_TIME, true);
    view.setUint16(pos + 14, DOS_DATE, true);
    view.setUint32(pos + 16, e.crc, true);
    view.setUint32(pos + 20, e.compressed.length, true);
    view.setUint32(pos + 24, e.size, true);
    view.setUint16(pos + 28, name.length, true);
    // extra, comment, disk, internal attrs = 0
    view.setUint32(pos + 38, 0, true);
    view.setUint32(pos + 42, offsets[i]!, true);
    out.set(name, pos + 46);
    pos += 46 + name.length;
  });
  view.setUint32(pos, EOCD, true);
  view.setUint16(pos + 8, entries.length, true);
  view.setUint16(pos + 10, entries.length, true);
  view.setUint32(pos + 12, pos - centralStart, true);
  view.setUint32(pos + 16, centralStart, true);
  return out;
}
