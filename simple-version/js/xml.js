/**
 * 汤姆EXCEL万能模块 — XML 序列化 / 解析 / 校验
 * 负责 Ribbon 树 <-> customUI14.xml（Office RibbonX schema）的可逆转换。
 */

import { uid, newControl } from './model.js';

export function escapeXml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function attrs(c) {
  let a = `id="${escapeXml(c.id)}"`;
  if (c.label !== undefined) a += ` label="${escapeXml(c.label)}"`;
  if (c.imageMso) a += ` imageMso="${escapeXml(c.imageMso)}"`;
  if (c.size === 'large') a += ' size="large"';
  if (c.onAction) a += ` onAction="${escapeXml(c.onAction)}"`;
  if (c.onChange) a += ` onChange="${escapeXml(c.onChange)}"`;
  if (c.screentip) a += ` screentip="${escapeXml(c.screentip)}"`;
  if (c.supertip) a += ` supertip="${escapeXml(c.supertip)}"`;
  if (c.visible === false) a += ' visible="false"';
  if (c.enabled === false) a += ' enabled="false"';
  if (c.description) a += ` description="${escapeXml(c.description)}"`;
  if (c.tag) a += ` tag="${escapeXml(c.tag)}"`;
  return a;
}

function controlToXml(c, indent) {
  const pad = ' '.repeat(indent);
  const voidTypes = ['separator', 'labelControl'];
  if (c.type === 'separator') return `${pad}<separator />`;
  if (voidTypes.includes(c.type)) return `${pad}<labelControl ${attrs(c)} />`;

  if (c.type === 'menu') {
    let s = `${pad}<menu ${attrs(c)}>\n`;
    (c.items || []).forEach(it => { s += controlToXml(it, indent + 2) + '\n'; });
    s += `${pad}</menu>`;
    return s;
  }
  if (c.type === 'gallery') {
    let s = `${pad}<gallery ${attrs(c)}>\n`;
    s += `${pad}  <itemGallery id="${escapeXml(c.id + 'Gallery')}" label="${escapeXml(c.label)}" />\n`;
    s += `${pad}</gallery>`;
    return s;
  }
  if (c.type === 'comboBox') {
    let s = `${pad}<comboBox ${attrs(c)}>\n`;
    (c.items || []).forEach(it => { s += `${pad}  <item id="${escapeXml(it.id || uid('itm'))}" label="${escapeXml(it.label)}" />\n`; });
    s += `${pad}</comboBox>`;
    return s;
  }
  const tag = ['toggleButton','checkBox','editBox','dropDown'].includes(c.type) ? c.type : 'button';
  return `${pad}<${tag} ${attrs(c)} />`;
}

/** 生成 customUI14.xml */
export function buildRibbonXml(ribbon) {
  let s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  s += '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">\n';
  s += `  <ribbon startFromScratch="${ribbon.startFromScratch ? 'true' : 'false'}">\n`;
  if (ribbon.qat && ribbon.qat.length) {
    s += '    <qat>\n';
    ribbon.qat.forEach(c => { s += controlToXml(c, 6) + '\n'; });
    s += '    </qat>\n';
  }
  s += '    <tabs>\n';
  ribbon.tabs.forEach(tab => {
    s += `      <tab id="${escapeXml(tab.id)}" label="${escapeXml(tab.label)}">\n`;
    tab.groups.forEach(group => {
      s += `        <group id="${escapeXml(group.id)}" label="${escapeXml(group.label)}">\n`;
      group.controls.forEach(c => { s += controlToXml(c, 10) + '\n'; });
      s += '        </group>\n';
    });
    s += '      </tab>\n';
  });
  s += '    </tabs>\n';
  s += '  </ribbon>\n';
  s += '</customUI>\n';
  return s;
}

/** 从 XML 文本解析回 Ribbon 树（尽力而为，尽量保留可读属性） */
export function parseRibbonXml(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const err = doc.querySelector('parsererror');
  if (err) throw new Error('XML 解析失败：' + err.textContent.slice(0, 200));
  const ribbonEl = doc.getElementsByTagName('ribbon')[0];
  if (!ribbonEl) throw new Error('未找到 <ribbon> 节点');
  const ribbon = {
    startFromScratch: ribbonEl.getAttribute('startFromScratch') === 'true',
    tabs: [], qat: [],
  };
  const local = el => el.localName || el.nodeName.replace(/^.*:/, '');
  const readAttrs = el => {
    const c = {};
    for (const a of Array.from(el.attributes)) c[a.name] = a.value;
    return c;
  };
  const nodeToControl = el => {
    const tag = local(el);
    if (tag === 'separator') return { type: 'separator', id: el.getAttribute('id') || uid('sep') };
    const raw = readAttrs(el);
    const c = {
      type: ['button','toggleButton','checkBox','editBox','comboBox','dropDown','gallery','menu','labelControl'].includes(tag) ? tag : 'button',
      id: raw.id || uid('ctrl'), label: raw.label || '', imageMso: raw.imageMso || '',
      size: raw.size || 'normal', onAction: raw.onAction || '', onChange: raw.onChange || '',
      screentip: raw.screentip || '', supertip: raw.supertip || '',
      visible: raw.visible !== 'false', enabled: raw.enabled !== 'false',
    };
    if (tag === 'menu' || tag === 'comboBox') {
      c.items = Array.from(el.children).map(ch => nodeToControl(ch)).filter(Boolean);
    }
    return c;
  };
  Array.from(ribbonEl.children).forEach(child => {
    if (local(child) === 'qat') {
      ribbon.qat = Array.from(child.children).map(nodeToControl);
    }
    if (local(child) === 'tabs') {
      Array.from(child.children).forEach(tabEl => {
        if (local(tabEl) !== 'tab') return;
        const tab = { id: tabEl.getAttribute('id') || uid('tab'), label: tabEl.getAttribute('label') || '选项卡', groups: [] };
        Array.from(tabEl.children).forEach(gEl => {
          if (local(gEl) !== 'group') return;
          const group = { id: gEl.getAttribute('id') || uid('grp'), label: gEl.getAttribute('label') || '分组', controls: [] };
          group.controls = Array.from(gEl.children).map(nodeToControl);
          tab.groups.push(group);
        });
        ribbon.tabs.push(tab);
      });
    }
  });
  return ribbon;
}

/** 校验 Ribbon 结构，返回 issue 列表 */
export function validateRibbon(ribbon) {
  const issues = [];
  const ids = new Map();
  const checkId = (c, where) => {
    if (!c.id) issues.push({ level: 'error', where, msg: '控件缺少 id' });
    else if (ids.has(c.id)) issues.push({ level: 'error', where, msg: `id 重复：「${c.id}」已用于 ${ids.get(c.id)}` });
    else ids.set(c.id, where);
    if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(c.id || '')) {
      issues.push({ level: 'error', where, msg: `id「${c.id}」不合法（须以字母开头，仅含字母数字下划线）` });
    }
    if (!['separator','labelControl'].includes(c.type) && !c.label) {
      issues.push({ level: 'warn', where, msg: '控件缺少 label 显示名称' });
    }
  };
  ribbon.tabs.forEach((tab, ti) => {
    const tw = `选项卡 ${ti + 1}「${tab.label}」`;
    checkId(tab, tw);
    if (!tab.label) issues.push({ level: 'error', where: tw, msg: '选项卡缺少 label' });
    if (!tab.groups.length) issues.push({ level: 'warn', where: tw, msg: '选项卡下没有分组' });
    tab.groups.forEach((g, gi) => {
      const gw = `${tw} / 分组 ${gi + 1}「${g.label}」`;
      checkId(g, gw);
      if (!g.label) issues.push({ level: 'warn', where: gw, msg: '分组缺少 label' });
      if (!g.controls.length) issues.push({ level: 'warn', where: gw, msg: '分组下没有控件' });
      const walk = (list) => list.forEach(c => {
        checkId(c, `${gw} / ${c.label || c.type}`);
        if (c.type === 'menu' && (!c.items || c.items.length < 2)) {
          issues.push({ level: 'warn', where: `${gw} / 菜单「${c.label}」`, msg: '菜单建议至少包含 2 个子项' });
        }
        if (['button','toggleButton','checkBox'].includes(c.type) && !c.onAction && !c.onChange) {
          issues.push({ level: 'info', where: `${gw} / ${c.label}`, msg: '按钮未绑定宏 onAction' });
        }
        if (c.items) walk(c.items);
      });
      walk(g.controls);
    });
  });
  (ribbon.qat || []).forEach((c, i) => checkId(c, `快速访问工具栏 ${i + 1}`));
  return issues;
}

/** 计算绑定统计 */
export function bindingStats(ribbon, macros) {
  const boundIds = new Set();
  const collect = list => (list || []).forEach(c => { if (c.onAction) boundIds.add(c.onAction); if (c.items) collect(c.items); });
  ribbon.tabs.forEach(t => t.groups.forEach(g => collect(g.controls)));
  collect(ribbon.qat);
  const macroNames = new Set(macros.map(m => `${m.module}.${m.name}`));
  let bound = 0, unbound = 0;
  macros.forEach(m => (boundIds.has(`${m.module}.${m.name}`) ? bound++ : unbound++));
  return { bound, unbound, total: macros.length };
}
