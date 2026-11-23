/**
 * 汤姆EXCEL万能模块 — 零依赖 ZIP 读写（基于浏览器原生 CompressionStream）
 * xlsm/xlam/docm 本质是 ZIP 容器，这里实现最小可用的 ZIP 解析与打包，
 * 用于读取 customUI.xml / 提取宏列表，以及把编辑后的 Ribbon 写回另存。
 */

const enc = new TextEncoder();
const dec = new TextDecoder();

function crc32(buf) {
  let c, crc = 0xFFFFFFFF;
  const table = crc32.table || (crc32.table = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })());
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function u16(a, o) { return a[o] | (a[o + 1] << 8); }
function u32(a, o) { return (a[o] | (a[o + 1] << 8) | (a[o + 2] << 16) | (a[o + 3] << 24)) >>> 0; }
function set16(a, o, v) { a[o] = v & 0xff; a[o + 1] = (v >>> 8) & 0xff; }
function set32(a, o, v) { a[o] = v & 0xff; a[o + 1] = (v >>> 8) & 0xff; a[o + 2] = (v >>> 16) & 0xff; a[o + 3] = (v >>> 24) & 0xff; }

/** 解压 deflate-raw 数据 */
async function inflateRaw(bytes) {
  if (typeof DecompressionStream === 'undefined') throw new Error('当前浏览器不支持 DecompressionStream，请使用新版 Chrome/Edge');
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** 解压 deflate（带 zlib 头） */
async function inflateZlib(bytes) {
  const ds = new DecompressionStream('deflate');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** 解压单个 ZIP entry 的原始数据 */
async function decompressEntry(entry, zipBytes) {
  const raw = zipBytes.slice(entry.dataOffset, entry.dataOffset + entry.compressedSize);
  if (entry.method === 0) return raw;
  if (entry.method === 8) {
    try { return await inflateRaw(raw); }
    catch {
      try { return await inflateZlib(raw); }
      catch (e) { throw new Error(`解压 ${entry.name} 失败：${e.message}`); }
    }
  }
  throw new Error(`不支持的 ZIP 压缩方法 ${entry.method}（${entry.name}）`);
}

/**
 * 解析 ZIP，返回 entries 数组。每个 entry: {name, method, compressedSize, size, dataOffset, crc}
 */
export async function readZip(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  // 找 EOCD（End of Central Directory）
  let eocd = -1;
  const min = Math.max(0, bytes.length - 22 - 65535);
  for (let i = bytes.length - 22; i >= min; i--) {
    if (u32(bytes, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('不是有效的 ZIP/xlsm 文件（未找到 EOCD）');
  const totalEntries = u16(bytes, eocd + 10);
  let cdOffset = u32(bytes, eocd + 16);
  // Zip64 兜底：若 EOCD 值为 0xFFFFFFFF，读 Zip64 EOCD locator
  if (cdOffset === 0xFFFFFFFF) {
    const z64 = u32(bytes, eocd - 20); // 粗略定位 zip64 eocd
    // 常规 locator 在 eocd-20，zip64 eocd offset 在 locator+8
    const loc = eocd - 20;
    if (u32(bytes, loc) === 0x07064b50) {
      const z64off = u32(bytes, loc + 8);
      cdOffset = u32(bytes, z64off + 48);
    }
  }
  const entries = [];
  let p = cdOffset;
  for (let i = 0; i < totalEntries; i++) {
    if (u32(bytes, p) !== 0x02014b50) break;
    const method = u16(bytes, p + 10);
    const compressedSize = u32(bytes, p + 20);
    const size = u32(bytes, p + 24);
    const nameLen = u16(bytes, p + 28);
    const extraLen = u16(bytes, p + 30);
    const commentLen = u16(bytes, p + 32);
    const localOffset = u32(bytes, p + 42);
    const name = dec.decode(bytes.slice(p + 46, p + 46 + nameLen));
    // 读 local header 拿数据真实起点
    const lhNameLen = u16(bytes, localOffset + 26);
    const lhExtraLen = u16(bytes, localOffset + 28);
    const dataOffset = localOffset + 30 + lhNameLen + lhExtraLen;
    entries.push({ name, method, compressedSize, size, dataOffset, crc: u32(bytes, p + 16) });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return { bytes, entries };
}

/** 读取 entry 为文本 */
export async function readZipText(zip, name) {
  const e = zip.entries.find(x => x.name === name);
  if (!e) return null;
  const data = await decompressEntry(e, zip.bytes);
  return dec.decode(data);
}

/** 读取 entry 为 Uint8Array */
export async function readZipBytes(zip, name) {
  const e = zip.entries.find(x => x.name === name);
  if (!e) return null;
  return decompressEntry(e, zip.bytes);
}

/** 列出所有 entry 名称 */
export function listZipNames(zip) { return zip.entries.map(e => e.name); }

/**
 * 写 ZIP：把 entries [{name, data:Uint8Array}] 打包成 ArrayBuffer。
 * 使用 deflate-raw 压缩（与 xlsm 常规一致），无外部依赖。
 */
export async function writeZip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const e of entries) {
    const nameBytes = enc.encode(e.name);
    const data = e.data instanceof Uint8Array ? e.data : enc.encode(String(e.data ?? ''));
    const useDeflate = e.store !== true && data.length > 0;
    let compData;
    if (useDeflate) {
      const cs = new CompressionStream('deflate-raw');
      const stream = new Blob([data]).stream().pipeThrough(cs);
      compData = new Uint8Array(await new Response(stream).arrayBuffer());
    } else {
      compData = data;
    }
    const method = useDeflate ? 8 : 0;
    const crc = crc32(data);

    // Local file header
    const lh = new Uint8Array(30 + nameBytes.length);
    set32(lh, 0, 0x04034b50);
    set16(lh, 4, 20);
    set16(lh, 6, 0);
    set16(lh, 8, method);
    set16(lh, 10, 0); set16(lh, 12, 0); // time/date
    set32(lh, 14, crc);
    set32(lh, 18, compData.length);
    set32(lh, 22, data.length);
    set16(lh, 26, nameBytes.length);
    set16(lh, 28, 0);
    lh.set(nameBytes, 30);
    locals.push(lh, compData);

    // Central directory
    const cd = new Uint8Array(46 + nameBytes.length);
    set32(cd, 0, 0x02014b50);
    set16(cd, 4, 20); set16(cd, 6, 20);
    set16(cd, 8, 0); set16(cd, 10, method);
    set16(cd, 12, 0); set16(cd, 14, 0);
    set32(cd, 16, crc);
    set32(cd, 20, compData.length);
    set32(cd, 24, data.length);
    set16(cd, 28, nameBytes.length);
    set16(cd, 30, 0); set16(cd, 32, 0);
    set16(cd, 34, 0); set16(cd, 36, 0);
    set32(cd, 38, 0);
    set32(cd, 42, offset);
    cd.set(nameBytes, 46);
    centrals.push(cd);

    offset += lh.length + compData.length;
  }
  const cdSize = centrals.reduce((s, c) => s + c.length, 0);
  const eocd = new Uint8Array(22);
  set32(eocd, 0, 0x06054b50);
  set16(eocd, 4, 0); set16(eocd, 6, 0);
  set16(eocd, 8, entries.length); set16(eocd, 10, entries.length);
  set32(eocd, 12, cdSize);
  set32(eocd, 16, offset);
  set16(eocd, 20, 0);

  const total = offset + cdSize + 22;
  const out = new Uint8Array(total);
  let pos = 0;
  for (const part of [...locals, ...centrals, eocd]) { out.set(part, pos); pos += part.length; }
  return out.buffer;
}

/** 把 Ribbon XML 写入已有 xlsm 的副本，返回新的 ArrayBuffer */
export async function injectRibbonIntoXlsm(originalBuffer, ribbonXml, {
  customUiEntry = 'customUI/customUI14.xml',
  relsEntry = 'customUI/_rels/customUI14.xml.rels',
} = {}) {
  const zip = await readZip(originalBuffer);
  const results = [];
  let hasRel = false;
  for (const e of zip.entries) {
    if (e.name === customUiEntry) {
      results.push({ name: e.name, data: enc.encode(ribbonXml) });
      continue;
    }
    if (e.name === relsEntry) hasRel = true;
    // 其余条目原样复制（解压再重压，保持内容一致）
    const data = await decompressEntry(e, zip.bytes);
    results.push({ name: e.name, data });
  }
  if (!zip.entries.some(e => e.name === customUiEntry)) {
    results.push({ name: customUiEntry, data: enc.encode(ribbonXml) });
  }
  if (!hasRel) {
    results.push({
      name: relsEntry,
      data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI14.xml" /></Relationships>'),
    });
  }
  return writeZip(results);
}

/** 触发浏览器下载 */
export function downloadBlob(buffer, filename, mime = 'application/octet-stream') {
  const blob = new Blob([buffer], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
}
