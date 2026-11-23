/**
 * 汤姆EXCEL万能模块 — 工作簿检查层
 * 从 xlsm/xlam 中提取：现有 Ribbon XML、VBA/JSA 宏过程列表。
 * VBA 项目为 OLE 复合文档且源码经 MS-OVBA 压缩，这里用启发式扫描
 * 提取可读过程名；若无法解析，则以 Ribbon onAction 作为绑定来源兜底。
 */

import { readZip, readZipText, readZipBytes, listZipNames } from './zip.js';
import { parseRibbonXml } from './xml.js';

function detectKind(name) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (['docm', 'dotm'].includes(ext)) return 'word';
  return 'excel';
}

/** 从字节中启发式扫描 VBA 过程名 Sub/Function */
function scanVbaProcedures(bytes) {
  const text = new TextDecoder('latin1').decode(bytes);
  const out = new Set();
  // MS-OVBA 压缩流里过程名常以压缩形式出现，但仍可匹配到部分明文/近似明文
  const re = /(?:Public\s+|Private\s+|Friend\s+|Static\s+)*(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(/g;
  let m;
  while ((m = re.exec(text)) && out.size < 500) out.add(m[1]);
  return [...out];
}

/** 从 JSA 的 JDEData.bin 提取函数名 */
function scanJsaProcedures(bytes) {
  const text = new TextDecoder('utf-8').decode(bytes);
  const out = new Set();
  const re = /function\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*\(/g;
  let m;
  while ((m = re.exec(text)) && out.size < 500) out.add(m[1]);
  return [...out];
}

/**
 * 检查工作簿，返回统一结构：
 * { name, extension, kind, customUiXml, ribbon, macros:[{module,name,signature,scope,kind,bound}], warnings:[] }
 */
export async function inspectWorkbook(file) {
  const buf = await file.arrayBuffer();
  const zip = await readZip(buf);
  const names = listZipNames(zip);
  const kind = detectKind(file.name);

  // 现有 Ribbon
  let customUiXml = null;
  customUiXml = await readZipText(zip, 'customUI/customUI14.xml');
  if (!customUiXml) customUiXml = await readZipText(zip, 'customUI/customUI.xml');

  // 宏列表
  const warnings = [];
  const macros = [];
  const vbaEntry = kind === 'word' ? 'word/vbaProject.bin' : 'xl/vbaProject.bin';
  const jsaEntry = kind === 'word' ? 'word/JDEData.bin' : 'xl/JDEData.bin';
  let foundNames = [];

  if (names.includes(vbaEntry)) {
    const vbaBytes = await readZipBytes(zip, vbaEntry);
    foundNames = scanVbaProcedures(vbaBytes);
    if (!foundNames.length) warnings.push('VBA 工程已加密或过程名为压缩存储，未能自动解析；已根据现有 Ribbon 绑定生成宏清单，可手动补充。');
  } else {
    const jsaBytes = await readZipBytes(zip, jsaEntry);
    if (jsaBytes) foundNames = scanJsaProcedures(jsaBytes);
    if (!foundNames.length) warnings.push('未在文件中发现可解析的宏过程，可在左侧宏面板手动添加。');
  }

  // 从已有 Ribbon onAction 里补充绑定过程
  const boundActions = new Set();
  let ribbon = null;
  if (customUiXml) {
    try {
      ribbon = parseRibbonXml(customUiXml);
      const collect = list => (list || []).forEach(c => {
        if (c.onAction) boundActions.add(c.onAction);
        if (c.items) collect(c.items);
      });
      ribbon.tabs.forEach(t => t.groups.forEach(g => collect(g.controls)));
      collect(ribbon.qat);
    } catch (e) { warnings.push('现有 Ribbon XML 解析失败：' + e.message); }
  }
  boundActions.forEach(a => {
    const parts = a.split('.');
    const name = parts.pop();
    const module = parts.join('.') || 'Module1';
    if (!foundNames.includes(name)) foundNames.push(name);
  });

  // 归类到模块
  foundNames.forEach((n, i) => {
    const fromBinding = [...boundActions].find(a => a.split('.').pop() === n);
    const module = fromBinding ? fromBinding.split('.').slice(0, -1).join('.') || 'Module1' : 'Module' + (Math.floor(i / 8) + 1);
    macros.push({
      module, name: n,
      signature: `${module}.${n}()`,
      scope: 'public', kind: 'Sub',
      bound: fromBinding ? (fromBinding) : '',
    });
  });

  return {
    name: file.name,
    extension: file.name.split('.').pop().toLowerCase(),
    kind,
    customUiXml,
    ribbon,
    macros,
    warnings,
    originalBuffer: buf,
    zipNames: names,
  };
}
