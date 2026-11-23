/**
 * 汤姆EXCEL万能模块 — 主应用逻辑
 */
import {
  WORKSPACES, CONTROL_TYPES, clone, uid, newControl, newGroup, newTab,
  sampleMacros, walkControls, findControl,
  moveNodeTo, canDrop,
} from './model.js';
import { buildRibbonXml, parseRibbonXml, validateRibbon, bindingStats } from './xml.js';
import { inspectWorkbook } from './workbook.js';
import { injectRibbonIntoXlsm, downloadBlob } from './zip.js';
import {
  loadAiSettings, saveAiSettings, AI_PROVIDERS, testConnection,
  planRibbon, planRibbonStream, planToRibbon,
} from './ai.js';
import { getItem, setItem } from './storage.js';

const $ = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));

/** 全局状态 */
const state = {
  ws: 'excel',
  ribbon: clone(WORKSPACES.excel.ribbon()),
  macros: sampleMacros('excel'),
  selection: null,          // {kind,tabId,groupId,controlId,qat}
  file: null,               // {name, originalBuffer}
  theme: getItem('tom-theme') || 'light',
  aiSettings: loadAiSettings(),
  aiMessages: [],
  aiSelectedMacros: new Set(),
  bottomTab: 'xml',
  inspectorMode: 'inspector',
  filter: '',
  bindFilter: 'all',
  collapsed: new Set(),
  activeTabIndex: 0,
  menuOpen: null,
};

const ICON_BASE = './icons/';
function iconFile(name) {
  if (!name) return '';
  const safe = String(name).replace(/[^A-Za-z0-9_]/g, '');
  const map = window.__ICON_MAP__ || {};
  const file = map[safe] || `${safe}_32.png`;
  return `${ICON_BASE}${file}`;
}
// 一些常见 imageMso 名称到图标文件名的兜底映射
const ICON_ALIAS = {
  RemoveDuplicates: 'RemoveDuplicates', Trim: 'Trim', FillDown: 'FillDown',
  FileSave: 'FileSave', FileSaveAsPdf: 'FileSaveAsPdf', FormatPainter: 'FormatPainter',
  CurrencyFormat: 'CurrencyFormat', PercentStyle: 'PercentStyle', Borders: 'Borders',
  ColumnChartInsert: 'ColumnChartInsert', Bold: 'Bold', Paste: 'Paste',
};

function toast(msg, type = '') {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'toast ' + type;
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.hidden = true; }, 2600);
}
function busy(on, text = '处理中...') {
  $('#busy-text').textContent = text;
  $('#busy').hidden = !on;
}
function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

/** 初始化 */
function init() {
  document.documentElement.dataset.theme = state.theme;
  $('#btn-theme').textContent = state.theme === 'dark' ? '☀️' : '🌙';
  bindEvents();
  renderAll();
  $('#boot-screen').remove();
  $('#app').hidden = false;
}

/** 渲染全部 */
function renderAll() {
  renderFileStrip();
  renderTree();
  renderPreview();
  renderBottom();
  renderInspector();
  renderAi();
}

/** 文件信息条 */
function renderFileStrip() {
  $('#fs-file').textContent = state.file ? state.file.name : '尚未导入 · 使用内置示例设计';
  const issues = validateRibbon(state.ribbon);
  const errs = issues.filter(i => i.level === 'error');
  const st = $('#fs-status');
  if (!state.file) { st.textContent = '示例数据'; st.className = 'fs-status ok'; }
  else if (errs.length) { st.textContent = `${errs.length} 处错误`; st.className = 'fs-status err'; }
  else { st.textContent = '校验通过'; st.className = 'fs-status ok'; }
  const bs = bindingStats(state.ribbon, state.macros);
  $('#fs-macro-count').textContent = `宏 ${bs.total} 个 · 已绑定 ${bs.bound} · 未绑定 ${bs.unbound}`;
  $('#issue-badge').textContent = issues.length;
  $('#macro-badge').textContent = state.macros.length;
}

/* ============ 结构树 ============ */
function renderTree() {
  const root = $('#ribbon-tree');
  const q = state.filter.toLowerCase();
  let html = '';
  const matches = (label, id) => !q || (label || '').toLowerCase().includes(q) || (id || '').toLowerCase().includes(q);
  const isBound = c => c.onAction || c.onChange || (c.items || []).some(isBound);

  const controlHtml = (c, depth) => {
    if (state.bindFilter === 'bound' && !isBound(c)) return '';
    if (state.bindFilter === 'unbound' && isBound(c)) return '';
    const sel = state.selection?.controlId === c.id;
    const icon = iconFile(c.imageMso || ICON_ALIAS[c.imageMso] || c.imageMso);
    const icoHtml = c.type === 'separator' ? '<span class="ico-ph">—</span>'
      : (icon ? `<img class="ico" src="${icon}" onerror="this.style.display='none'" />` : '<span class="ico-ph">▫</span>');
    let children = '';
    if (c.items && c.items.length) {
      const key = 'c:' + c.id;
      const open = !state.collapsed.has(key);
      children = `<span class="twisty">${open ? '▾' : '▸'}</span>`;
      if (open) children += c.items.map(it => controlHtml(it, depth + 1)).join('');
    } else children = '<span class="twisty"></span>';
    return `<div class="tree-item depth-${depth} ${sel ? 'selected' : ''}" draggable="true" data-drag="1" data-ctype="control" data-id="${escapeHtml(c.id)}" title="${escapeHtml(c.label || c.type)} · 可拖动移动">
      ${children}${icoHtml}<span class="label">${escapeHtml(c.label || c.type)}</span>
      <span class="badge-mini">${c.type}</span>
      <span class="bind-dot ${isBound(c) ? 'bound' : 'unbound'}"></span>
    </div>`;
  };

  state.ribbon.tabs.forEach((tab, ti) => {
    if (!matches(tab.label, tab.id) && !tab.groups.some(g => matches(g.label, g.id) || g.controls.some(c => matches(c.label, c.id)))) return;
    const tsel = state.selection?.kind === 'tab' && state.selection.tabId === tab.id;
    const tkey = 't:' + tab.id;
    const open = !state.collapsed.has(tkey);
    let h = `<div class="tree-item depth-0 ${tsel ? 'selected' : ''}" draggable="true" data-drag="1" data-ctype="tab" data-id="${escapeHtml(tab.id)}" title="${escapeHtml(tab.label)} · 可拖动排序">
      <span class="twisty">${open ? '▾' : '▸'}</span><span class="ico-ph">📑</span>
      <span class="label">${escapeHtml(tab.label)}</span><span class="badge-mini">${tab.groups.length} 组</span></div>`;
    if (open) {
      tab.groups.forEach((g, gi) => {
        const gsel = state.selection?.kind === 'group' && state.selection.groupId === g.id;
        const gkey = 'g:' + g.id;
        const gopen = !state.collapsed.has(gkey);
        h += `<div class="tree-item depth-1 ${gsel ? 'selected' : ''}" draggable="true" data-drag="1" data-ctype="group" data-id="${escapeHtml(g.id)}" data-tab="${escapeHtml(tab.id)}" title="${escapeHtml(g.label)} · 可拖动移动">
          <span class="twisty">${gopen ? '▾' : '▸'}</span><span class="ico-ph">📁</span>
          <span class="label">${escapeHtml(g.label)}</span><span class="badge-mini">${g.controls.length}</span></div>`;
        if (gopen) h += g.controls.map(c => controlHtml(c, 2)).join('');
      });
    }
    html += h;
  });

  // QAT
  if (state.ribbon.qat?.length) {
    html += `<div class="tree-item depth-0" style="color:var(--text-3);cursor:default"><span class="ico-ph">⚡</span><span class="label">快速访问工具栏</span></div>`;
    html += state.ribbon.qat.map(c => controlHtml(c, 1)).join('');
  }
  root.innerHTML = html || '<div class="empty-hint">没有匹配的结构</div>';
}

/* ============ Ribbon 预览 ============ */
function renderPreview() {
  const r = state.ribbon;
  const tab = r.tabs[state.activeTabIndex] || r.tabs[0];
  let html = '<div class="office-mock">';
  html += `<div class="office-titlebar"><span>📊 ${escapeHtml(state.ws === 'word' ? 'Word' : state.ws === 'jsa' ? 'WPS 表格' : 'Excel')} · 汤姆EXCEL万能模块</span>
    <span class="qat">${(r.qat || []).map(c => {
      const ic = iconFile(c.imageMso);
      return ic ? `<img src="${ic}" title="${escapeHtml(c.label)}" onerror="this.style.display='none'"/>` : `<span>${escapeHtml(c.label)}</span>`;
    }).join('')}</span></div>`;
  html += '<div class="office-tabs"><span class="office-tab file">文件</span>';
  r.tabs.forEach((t, i) => {
    html += `<span class="office-tab ${i === state.activeTabIndex ? 'active' : ''}" data-tab-index="${i}" data-ctype="tab" data-id="${escapeHtml(t.id)}" title="${escapeHtml(t.label)} · 可把分组拖到这里">${escapeHtml(t.label)}</span>`;
  });
  html += '</div>';
  html += '<div class="office-ribbon">';
  if (tab) {
    tab.groups.forEach(g => {
      html += `<div class="office-group" data-ctype="group" data-id="${escapeHtml(g.id)}" data-tab="${escapeHtml(tab.id)}" title="分组「${escapeHtml(g.label)}」· 可把控件拖到这里"><div class="office-group-body">`;
      g.controls.forEach(c => { html += previewControl(c); });
      html += `</div><div class="office-group-label">${escapeHtml(g.label)}</div></div>`;
    });
  }
  html += '</div></div>';
  $('#ribbon-preview').innerHTML = html;
}

function previewControl(c) {
  if (c.type === 'separator') return '<div class="office-ctrl sep"></div>';
  const sel = state.selection?.controlId === c.id ? 'selected' : '';
  const large = c.size === 'large' ? 'large' : '';
  const icon = iconFile(c.imageMso || ICON_ALIAS[c.imageMso]);
  const ico = icon ? `<img src="${icon}" onerror="this.replaceWith('▫')"/>` : '<span class="ico-ph">▫</span>';
  if (c.type === 'menu') {
    return `<div class="office-ctrl ${large} ${sel} office-menu" draggable="true" data-drag="1" data-ctype="control" data-id="${escapeHtml(c.id)}" title="${escapeHtml(c.screentip || c.label)} · 可拖动移动">
      ${ico}<span class="cl">${escapeHtml(c.label)} ▾</span>
      <div class="menu-popover" hidden>${(c.items || []).map(it => {
        if (it.type === 'separator') return '<div class="mi sep"></div>';
        const iic = iconFile(it.imageMso);
        return `<div class="mi" draggable="true" data-drag="1" data-ctype="control" data-id="${escapeHtml(it.id)}" title="可拖动移动">${iic ? `<img src="${iic}"/>` : '<span>▫</span>'}<span>${escapeHtml(it.label)}</span></div>`;
      }).join('')}</div>
    </div>`;
  }
  return `<div class="office-ctrl ${large} ${sel}" draggable="true" data-drag="1" data-ctype="control" data-id="${escapeHtml(c.id)}" title="${escapeHtml(c.screentip || c.label)} · 可拖动移动">
    ${ico}<span class="cl">${escapeHtml(c.label)}</span></div>`;
}

/* ============ 底部面板 ============ */
function renderBottom() {
  $$('.bottom-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === state.bottomTab));
  const c = $('#bottom-content');
  if (state.bottomTab === 'xml') {
    const xml = buildRibbonXml(state.ribbon);
    c.innerHTML = `<pre class="xml-code">${escapeHtml(xml)}</pre>`;
  } else if (state.bottomTab === 'issues') {
    const issues = validateRibbon(state.ribbon);
    if (!issues.length) { c.innerHTML = '<div class="empty-hint">✅ 校验通过，未发现问题</div>'; return; }
    const lvName = { error: '错误', warn: '警告', info: '提示' };
    c.innerHTML = issues.map(i => `<div class="issue-row ${i.level}"><span class="lv">${lvName[i.level]}</span><span><b>${escapeHtml(i.where)}</b> · ${escapeHtml(i.msg)}</span></div>`).join('');
  } else {
    if (!state.macros.length) { c.innerHTML = '<div class="empty-hint">暂无宏过程</div>'; return; }
    c.innerHTML = `<table class="macro-table"><thead><tr><th>模块</th><th>过程名</th><th>签名</th><th>绑定</th></tr></thead><tbody>
      ${state.macros.map(m => `<tr><td>${escapeHtml(m.module)}</td><td>${escapeHtml(m.name)}</td><td>${escapeHtml(m.signature)}</td>
      <td><span class="bound-tag ${m.bound ? 'yes' : 'no'}">${m.bound ? '已绑定' : '未绑定'}</span></td></tr>`).join('')}
    </tbody></table>`;
  }
}

/* ============ 属性面板 ============ */
function renderInspector() {
  $$('.inspector-tab').forEach(b => b.classList.toggle('active', b.dataset.mode === state.inspectorMode));
  $('#inspector-content').hidden = state.inspectorMode !== 'inspector';
  $('#ai-content').hidden = state.inspectorMode !== 'ai';
  if (state.inspectorMode === 'ai') return;
  const host = $('#inspector-content');
  const sel = state.selection;
  if (!sel) {
    host.innerHTML = `<div class="inspector-title">属性与绑定</div><div class="inspector-sub">未选择</div>
      <div class="hint-box">在左侧结构树或中间预览里选择一个选项卡、分组或控件，即可在这里编辑它的属性与宏绑定。
      <br><br>右键结构树节点可进行增删、复制、上移下移等操作。</div>`;
    return;
  }
  let obj = null, where = '';
  if (sel.kind === 'tab') { obj = state.ribbon.tabs.find(t => t.id === sel.tabId); where = '选项卡'; }
  else if (sel.kind === 'group') { const t = state.ribbon.tabs.find(t => t.id === sel.tabId); obj = t?.groups.find(g => g.id === sel.groupId); where = '分组'; }
  else {
    const f = findControl(state.ribbon, sel.controlId);
    obj = f?.control; where = '控件';
  }
  if (!obj) { host.innerHTML = '<div class="empty-hint">对象已不存在</div>'; return; }

  let h = `<div class="inspector-title">${where}属性</div><div class="inspector-sub">正在编辑：${escapeHtml(obj.label || obj.id)}</div>`;
  h += `<div class="form-group"><label>标识 ID</label><input id="p-id" value="${escapeHtml(obj.id || '')}" ${sel.kind !== 'control' ? '' : ''}/></div>`;
  h += `<div class="form-group"><label>显示名称 (label)</label><input id="p-label" value="${escapeHtml(obj.label || '')}"/></div>`;

  if (sel.kind === 'control') {
    h += `<div class="form-group"><label>控件类型</label><div class="hint-box">${obj.type}</div></div>`;
    if (obj.type !== 'separator' && obj.type !== 'labelControl') {
      h += `<div class="form-group"><label>图标 (imageMso)</label>
        <div class="icon-picker"><img id="p-icon-preview" src="${iconFile(obj.imageMso)}" onerror="this.style.visibility='hidden'"/>
        <input id="p-imageMso" value="${escapeHtml(obj.imageMso || '')}" placeholder="例如 Bold" style="flex:1"/>
        <button id="p-pick-icon" class="mini-btn">选择</button></div></div>`;
      if (['button','toggleButton','checkBox','menu'].includes(obj.type)) {
        h += `<div class="form-group"><label>回调 onAction（模块.过程名）</label><input id="p-onAction" value="${escapeHtml(obj.onAction || '')}" placeholder="例如 CleanData.TrimText"/></div>`;
        h += `<div class="form-group"><label>屏幕提示 screentip</label><input id="p-screentip" value="${escapeHtml(obj.screentip || '')}"/></div>`;
        h += `<div class="form-group"><label>详细说明 supertip</label><textarea id="p-supertip">${escapeHtml(obj.supertip || '')}</textarea></div>`;
      }
      if (['editBox','comboBox','dropDown','gallery'].includes(obj.type)) {
        h += `<div class="form-group"><label>回调 onChange</label><input id="p-onChange" value="${escapeHtml(obj.onChange || '')}"/></div>`;
      }
      h += `<div class="form-row">
        <div class="form-group"><label>尺寸</label><select id="p-size">
          <option value="normal" ${obj.size !== 'large' ? 'selected' : ''}>normal</option>
          <option value="large" ${obj.size === 'large' ? 'selected' : ''}>large</option></select></div>
      </div>`;
      h += `<div class="checkbox-row"><input type="checkbox" id="p-visible" ${obj.visible !== false ? 'checked' : ''}/><label for="p-visible">可见</label></div>`;
      h += `<div class="checkbox-row"><input type="checkbox" id="p-enabled" ${obj.enabled !== false ? 'checked' : ''}/><label for="p-enabled">可用</label></div>`;
    }
    h += `<div class="section-divider"></div><button id="p-delete" class="ghost-btn" style="color:var(--err);width:100%">删除此控件</button>`;
  } else if (sel.kind === 'group') {
    h += `<div class="section-divider"></div>
      <div class="form-group"><label>添加控件</label><select id="p-add-ctrl">
        ${CONTROL_TYPES.map(t => `<option value="${t}">${t}</option>`).join('')}</select></div>
      <button id="p-add-ctrl-btn" class="primary-btn" style="width:100%">+ 添加控件</button>
      <div class="section-divider"></div><button id="p-delete" class="ghost-btn" style="color:var(--err);width:100%">删除此分组</button>`;
  } else if (sel.kind === 'tab') {
    h += `<div class="section-divider"></div><button id="p-add-group" class="primary-btn" style="width:100%">+ 添加分组</button>
      <div style="height:8px"></div><button id="p-delete" class="ghost-btn" style="color:var(--err);width:100%">删除此选项卡</button>`;
  }
  host.innerHTML = h;
  bindInspectorEvents(obj, sel);
}

function bindInspectorEvents(obj, sel) {
  const live = (id, fn) => { const el = $('#' + id); if (el) el.addEventListener('input', fn); };
  live('p-id', e => { obj.id = e.target.value; renderTree(); });
  live('p-label', e => { obj.label = e.target.value; renderTree(); renderPreview(); });
  live('p-imageMso', e => {
    obj.imageMso = e.target.value;
    const img = $('#p-icon-preview'); if (img) { img.src = iconFile(obj.imageMso); img.style.visibility = 'visible'; }
    renderTree(); renderPreview();
  });
  live('p-onAction', e => { obj.onAction = e.target.value; renderTree(); updateMacroBindings(); });
  live('p-onChange', e => { obj.onChange = e.target.value; renderTree(); });
  live('p-screentip', e => { obj.screentip = e.target.value; });
  live('p-supertip', e => { obj.supertip = e.target.value; });
  const size = $('#p-size'); if (size) size.addEventListener('change', e => { obj.size = e.target.value; renderPreview(); });
  const vis = $('#p-visible'); if (vis) vis.addEventListener('change', e => { obj.visible = e.target.checked; renderTree(); renderPreview(); });
  const en = $('#p-enabled'); if (en) en.addEventListener('change', e => { obj.enabled = e.target.checked; });
  const pick = $('#p-pick-icon'); if (pick) pick.addEventListener('click', () => openIconGallery(name => { obj.imageMso = name; renderInspector(); renderTree(); renderPreview(); }));
  const del = $('#p-delete'); if (del) del.addEventListener('click', () => deleteSelection());
  const addG = $('#p-add-group'); if (addG) addG.addEventListener('click', () => {
    const t = state.ribbon.tabs.find(t => t.id === sel.tabId);
    const g = newGroup('新分组'); t.groups.push(g);
    state.selection = { kind: 'group', tabId: t.id, groupId: g.id };
    renderAll();
  });
  const addC = $('#p-add-ctrl-btn'); if (addC) addC.addEventListener('click', () => {
    const t = state.ribbon.tabs.find(t => t.id === sel.tabId);
    const g = t.groups.find(g => g.id === sel.groupId);
    const type = $('#p-add-ctrl').value;
    const c = newControl(type, type === 'separator' ? '' : '新控件');
    g.controls.push(c);
    state.selection = { kind: 'control', tabId: t.id, groupId: g.id, controlId: c.id };
    renderAll();
  });
}

function updateMacroBindings() {
  const used = new Set();
  const collect = list => (list || []).forEach(c => {
    if (c.onAction) used.add(c.onAction);
    if (c.items) collect(c.items);
  });
  state.ribbon.tabs.forEach(t => t.groups.forEach(g => collect(g.controls)));
  (state.ribbon.qat || []).forEach(c => { if (c.onAction) used.add(c.onAction); });
  state.macros.forEach(m => {
    const full = `${m.module}.${m.name}`;
    m.bound = used.has(full) ? full : '';
  });
}

function deleteSelection() {
  const s = state.selection;
  if (!s) return;
  if (s.kind === 'tab') state.ribbon.tabs = state.ribbon.tabs.filter(t => t.id !== s.tabId);
  else if (s.kind === 'group') {
    const t = state.ribbon.tabs.find(t => t.id === s.tabId);
    if (t) t.groups = t.groups.filter(g => g.id !== s.groupId);
  } else {
    const f = findControl(state.ribbon, s.controlId);
    if (f) { f.siblings.splice(f.index, 1); }
  }
  state.selection = null;
  renderAll();
  toast('已删除', 'ok');
}

/* ============ AI 面板 ============ */
function renderAi() {
  const host = $('#ai-content');
  const s = state.aiSettings;
  const provider = AI_PROVIDERS.find(p => p.id === s.providerId);
  let h = `<div class="ai-status">${escapeHtml(provider?.label || s.providerId)} · ${escapeHtml(s.model || '未设置模型')} · ${s.apiKey ? 'API Key 已保存' : '未保存 API Key'}</div>`;
  h += `<div class="ai-messages" id="ai-messages">`;
  if (!state.aiMessages.length) {
    h += `<div class="ai-msg assistant">我可以像设计 Agent 一样，根据当前导入文件里的宏，帮你规划清晰好用的 Ribbon 功能区。描述你的需求，或在下方选择要纳入的宏。</div>`;
  } else {
    state.aiMessages.forEach(m => { h += `<div class="ai-msg ${m.role}">${escapeHtml(m.content)}</div>`; });
  }
  h += `</div>`;
  h += `<div class="ai-macro-chips" id="ai-chips">${state.macros.map(m => {
    const key = `${m.module}.${m.name}`;
    return `<span class="macro-chip ${state.aiSelectedMacros.has(key) ? 'on' : ''}" data-key="${escapeHtml(key)}">@${escapeHtml(m.name)}</span>`;
  }).join('')}</div>`;
  h += `<div class="ai-input-wrap">
    <textarea id="ai-input" placeholder="和 AI 讨论功能区设计，例如：按数据清洗→报表输出规划一个选项卡"></textarea>
    <div class="ai-input-actions">
      <button id="ai-send" class="primary-btn">✨ 发送给 AI</button>
      <button id="ai-clear" class="ghost-btn">清空</button>
    </div></div>`;
  host.innerHTML = h;
  const box = $('#ai-messages'); if (box) box.scrollTop = box.scrollHeight;

  $('#ai-chips')?.addEventListener('click', e => {
    const chip = e.target.closest('.macro-chip'); if (!chip) return;
    const k = chip.dataset.key;
    if (state.aiSelectedMacros.has(k)) state.aiSelectedMacros.delete(k); else state.aiSelectedMacros.add(k);
    chip.classList.toggle('on');
  });
  $('#ai-clear')?.addEventListener('click', () => { state.aiMessages = []; renderAi(); });
  $('#ai-send')?.addEventListener('click', sendAi);
}

async function sendAi() {
  const input = $('#ai-input');
  const text = input.value.trim();
  if (!text) { toast('请输入设计需求', 'err'); return; }
  if (!state.aiSettings.apiKey) { toast('请先在 AI 配置中填写 API Key', 'err'); openSettings(); return; }
  const selected = [...state.aiSelectedMacros];
  const macros = selected.length ? state.macros.filter(m => selected.includes(`${m.module}.${m.name}`)) : state.macros;
  state.aiMessages.push({ role: 'user', content: text });
  input.value = '';
  renderAi();
  busy(true, 'AI 正在规划 Ribbon...');
  try {
    const imageMsos = listIconNames().slice(0, 200);
    const result = await planRibbon(state.aiSettings, text, macros, imageMsos, 'replace');
    const ribbon = planToRibbon(result);
    state.ribbon = ribbon;
    state.activeTabIndex = 0;
    state.selection = null;
    state.aiMessages.push({ role: 'assistant', content: (result.summary || '已生成 Ribbon 方案') + '\n（已应用到设计画布，可继续微调）' });
    renderAll();
    toast('AI 方案已应用', 'ok');
  } catch (e) {
    state.aiMessages.push({ role: 'assistant', content: '生成失败：' + e.message });
    renderAi();
    toast('AI 生成失败：' + e.message, 'err');
  } finally { busy(false); }
}

async function openSettings() {
  const s = state.aiSettings;
  const opts = AI_PROVIDERS.map(p => `<option value="${p.id}" ${p.id === s.providerId ? 'selected' : ''}>${p.label}</option>`).join('');
  openModal('AI 配置', `
    <div class="form-group"><label>服务商</label><select id="set-provider">${opts}</select></div>
    <div class="form-group"><label>Base URL（OpenAI 兼容）</label><input id="set-base" value="${escapeHtml(s.baseUrl)}"/></div>
    <div class="form-group"><label>模型名称</label><input id="set-model" value="${escapeHtml(s.model)}"/></div>
    <div class="form-group"><label>API Key（仅保存在本机浏览器）</label><input id="set-key" type="password" value="${escapeHtml(s.apiKey)}" placeholder="sk-..."/></div>
    <div class="hint-box">配置仅存于本机浏览器 localStorage，不会写入工作簿文件。浏览器直连大模型需服务端允许跨域（CORS）；若连接失败，可点「复制方案提示词」到支持的工具里生成 XML 后粘贴导入。</div>
    <div id="set-test-result" style="margin-top:10px"></div>
  `, [
    { label: '测试连接', cls: 'ghost-btn', action: async () => {
      const ns = readSettingsForm();
      $('#set-test-result').textContent = '测试中...';
      const r = await testConnection(ns);
      $('#set-test-result').innerHTML = `<div class="issue-row ${r.ok ? 'info' : 'error'}">${r.ok ? '✅' : '❌'} ${escapeHtml(r.message)}</div>`;
    }, keepOpen: true },
    { label: '保存配置', cls: 'primary-btn', action: () => {
      state.aiSettings = readSettingsForm();
      saveAiSettings(state.aiSettings);
      renderAi();
      closeModal();
      toast('AI 配置已保存', 'ok');
    }},
  ], 'wide');
  $('#set-provider').addEventListener('change', e => {
    const p = AI_PROVIDERS.find(x => x.id === e.target.value);
    if (p && p.baseUrl) { $('#set-base').value = p.baseUrl; $('#set-model').value = p.model; }
  });
}
function readSettingsForm() {
  return {
    providerId: $('#set-provider').value,
    baseUrl: $('#set-base').value.trim(),
    model: $('#set-model').value.trim(),
    apiKey: $('#set-key').value.trim(),
  };
}

/* ============ 图标库 ============ */
let _iconNames = null;
function listIconNames() {
  if (_iconNames) return _iconNames;
  // 从已加载的 32px 图标文件名生成候选 imageMso 名称
  _iconNames = [];
  return _iconNames;
}

function openIconGallery(onPick) {
  const names = getIconNames();
  const body = `
    <input id="icon-search" class="search-input" placeholder="搜索 imageMso 名称，例如 Bold / Font / Table" style="width:100%"/>
    <div id="icon-grid" class="icon-gallery"></div>`;
  openModal('选择 Office imageMso 图标', body, [], 'wide');
  const grid = $('#icon-grid');
  const draw = (q = '') => {
    const list = names.filter(n => !q || n.toLowerCase().includes(q.toLowerCase())).slice(0, 300);
    grid.innerHTML = list.map(n => `<div class="icon-cell" data-name="${escapeHtml(n)}">
      <img src="${iconFile(n)}" onerror="this.style.display='none'"/><span>${escapeHtml(n)}</span></div>`).join('')
      || '<div class="empty-hint">没有匹配图标</div>';
  };
  draw();
  $('#icon-search').addEventListener('input', e => draw(e.target.value));
  grid.addEventListener('click', e => {
    const cell = e.target.closest('.icon-cell'); if (!cell) return;
    onPick(cell.dataset.name); closeModal();
  });
}

let _iconNameCache = null;
function getIconNames() {
  if (_iconNameCache) return _iconNameCache;
  const fromIcons = window.__ICON_NAMES__ || [];
  const fromIdmso = (window.__IDMSO__ || []).map(d => d.imageMso).filter(Boolean);
  _iconNameCache = [...new Set([...fromIcons, ...fromIdmso])]
    .filter(n => /^[A-Za-z][A-Za-z0-9_]*$/.test(n)).sort();
  return _iconNameCache;
}

/* ============ 弹窗 ============ */
function openModal(title, bodyHtml, buttons = [], size = '') {
  const root = $('#modal-root');
  const foot = buttons.map((b, i) => `<button class="${b.cls || 'ghost-btn'}" data-mbtn="${i}">${b.label}</button>`).join('');
  root.innerHTML = `<div class="modal-overlay"><div class="modal ${size}">
    <div class="modal-head"><h3>${escapeHtml(title)}</h3><button class="close" id="modal-close">×</button></div>
    <div class="modal-body">${bodyHtml}</div>
    ${foot ? `<div class="modal-foot">${foot}</div>` : ''}
  </div></div>`;
  $('#modal-close').addEventListener('click', closeModal);
  root.querySelector('.modal-overlay').addEventListener('click', e => { if (e.target.classList.contains('modal-overlay')) closeModal(); });
  buttons.forEach((b, i) => {
    const el = root.querySelector(`[data-mbtn="${i}"]`);
    if (el) el.addEventListener('click', async () => {
      const r = await b.action();
      if (!b.keepOpen && r !== false) closeModal();
    });
  });
}
function closeModal() { $('#modal-root').innerHTML = ''; }

/* ============ 导入 / 导出 ============ */
async function doImport(file) {
  if (!file) return;
  busy(true, '正在检查工作簿...');
  try {
    const info = await inspectWorkbook(file);
    state.file = { name: info.name, originalBuffer: info.originalBuffer };
    state.macros = info.macros;
    if (info.ribbon && info.ribbon.tabs.length) {
      state.ribbon = info.ribbon;
    } else {
      state.ribbon = clone(WORKSPACES[state.ws].ribbon());
      // 用真实宏名重建示例绑定的提示
    }
    state.ws = info.kind === 'word' ? 'word' : state.ws;
    $$('.ws-btn').forEach(b => b.classList.toggle('active', b.dataset.ws === state.ws));
    state.selection = null;
    state.activeTabIndex = 0;
    renderAll();
    if (info.warnings.length) toast(info.warnings[0], '');
    else toast(`已导入 ${info.name} · ${info.macros.length} 个宏`, 'ok');
  } catch (e) {
    toast('导入失败：' + e.message, 'err');
  } finally { busy(false); $('#file-input').value = ''; }
}

async function doSaveWorkbook() {
  if (!state.file) { toast('请先导入工作簿文件', 'err'); return; }
  busy(true, '正在写入 Ribbon 并另存...');
  try {
    const xml = buildRibbonXml(state.ribbon);
    const isWord = state.file.name.toLowerCase().endsWith('.docm') || state.file.name.toLowerCase().endsWith('.dotm');
    const out = await injectRibbonIntoXlsm(state.file.originalBuffer, xml, {
      customUiEntry: 'customUI/customUI14.xml',
      relsEntry: 'customUI/_rels/customUI14.xml.rels',
    });
    const name = state.file.name.replace(/\.(xlsm|xlam|docm|dotm)$/i, '') + '-汤姆模块.xlsm';
    downloadBlob(out, name, 'application/vnd.ms-excel.sheet.macroEnabled.12');
    toast('已另存工作簿（Ribbon 已写入，宏保持不变）', 'ok');
  } catch (e) {
    toast('另存失败：' + e.message, 'err');
  } finally { busy(false); }
}

function doExportXml() {
  const xml = buildRibbonXml(state.ribbon);
  downloadBlob(new TextEncoder().encode(xml), 'customUI14.xml', 'application/xml');
  toast('已导出 customUI14.xml', 'ok');
}

/* ============ 结构树右键菜单 ============ */
function showTreeContextMenu(e, el) {
  e.preventDefault();
  const id = el.dataset.id;
  const kind = el.dataset.ctype;
  let menu = [
    { label: '✏️ 编辑属性', act: () => selectNode(el) },
  ];
  if (kind !== 'tab') menu.push({ label: '⧉ 复制', act: () => duplicateNode(id, kind) });
  menu.push({ sep: true });
  menu.push({ label: '↑ 上移', act: () => moveNode(id, -1) });
  menu.push({ label: '↓ 下移', act: () => moveNode(id, 1) });
  menu.push({ sep: true });
  menu.push({ label: '🗑 删除', danger: true, act: () => { selectNode(el); deleteSelection(); } });

  const existing = $('.ctx-menu'); if (existing) existing.remove();
  const m = document.createElement('div');
  m.className = 'ctx-menu';
  m.innerHTML = menu.map((it, i) => it.sep ? '<div class="sep"></div>' : `<button data-i="${i}" class="${it.danger ? 'danger' : ''}">${it.label}</button>`).join('');
  m.style.left = Math.min(e.clientX, window.innerWidth - 180) + 'px';
  m.style.top = Math.min(e.clientY, window.innerHeight - 240) + 'px';
  document.body.appendChild(m);
  m.addEventListener('click', ev => {
    const b = ev.target.closest('button'); if (!b) return;
    const items = menu.filter(x => !x.sep);
    const idx = +b.dataset.i;
    // 重新映射（sep 不占按钮 index，但 data-i 用的是原始下标）
    const real = menu[idx];
    if (real && real.act) real.act();
    m.remove();
  });
  setTimeout(() => document.addEventListener('click', () => m.remove(), { once: true }), 0);
}

function selectNode(el) {
  const id = el.dataset.id, kind = el.dataset.ctype;
  if (kind === 'tab') state.selection = { kind: 'tab', tabId: id };
  else if (kind === 'group') state.selection = { kind: 'group', tabId: el.dataset.tab, groupId: id };
  else {
    const ti = state.ribbon.tabs.findIndex(t => t.groups.some(g => (g.controls || []).some(c => c.id === id) || (g.controls || []).some(c => (c.items||[]).some(i=>i.id===id))));
    if (ti >= 0) state.activeTabIndex = ti;
    state.selection = { kind: 'control', controlId: id };
  }
  state.inspectorMode = 'inspector';
  renderTree(); renderPreview(); renderBottom(); renderInspector();
  $$('.inspector-tab').forEach(b => b.classList.toggle('active', b.dataset.mode === 'inspector'));
  $('#inspector-content').hidden = false; $('#ai-content').hidden = true;
}

function duplicateNode(id, kind) {
  const f = findControl(state.ribbon, id);
  if (f) {
    const copy = clone(f.control);
    copy.id = uid(copy.type);
    (copy.items || []).forEach(it => it.id = uid(it.type));
    f.siblings.splice(f.index + 1, 0, copy);
    toast('已复制控件', 'ok');
  } else if (kind === 'group') {
    for (const t of state.ribbon.tabs) {
      const i = t.groups.findIndex(g => g.id === id);
      if (i >= 0) { const g = clone(t.groups[i]); g.id = uid('grp'); g.label += ' (副本)'; g.controls.forEach(c => c.id = uid(c.type)); t.groups.splice(i+1, 0, g); toast('已复制分组','ok'); break; }
    }
  }
  renderAll();
}

function moveNode(id, dir) {
  const f = findControl(state.ribbon, id);
  if (f) {
    const j = f.index + dir;
    if (j < 0 || j >= f.siblings.length) return;
    [f.siblings[f.index], f.siblings[j]] = [f.siblings[j], f.siblings[f.index]];
  } else {
    const tabs = state.ribbon.tabs;
    let ti = tabs.findIndex(t => t.id === id);
    if (ti >= 0) { const j = ti + dir; if (j < 0 || j >= tabs.length) return; [tabs[ti], tabs[j]] = [tabs[j], tabs[ti]]; }
    else for (const t of tabs) {
      const gi = t.groups.findIndex(g => g.id === id);
      if (gi >= 0) { const j = gi + dir; if (j < 0 || j >= t.groups.length) return; [t.groups[gi], t.groups[j]] = [t.groups[j], t.groups[gi]]; break; }
    }
  }
  renderAll();
}

/* ============ 拖放移动 ============ */
const dragCtx = { src: null, el: null, target: null, position: null };

/** 计算拖放目标位置：before / after / inside */
function calcDropPosition(el, src, dst, e) {
  const rect = el.getBoundingClientRect();
  // 树节点竖排 -> 用垂直方向；预览区 tab/控件横排 -> 用水平方向
  const inTree = !!el.closest('#ribbon-tree');
  const horizontal = !inTree && rect.width > rect.height * 1.2;
  let ratio = horizontal ? (e.clientX - rect.left) / rect.width : (e.clientY - rect.top) / rect.height;
  if (!isFinite(ratio) || (rect.width === 0 && rect.height === 0)) ratio = 0.5; // 兜底（如无布局环境）

  // 支持 inside 的目标：group（放控件/tab）、menu 控件（放控件）、tab（放分组）
  let allowInside = false;
  if (src.kind === 'control' && dst.kind === 'group') allowInside = true;
  if (src.kind === 'control' && dst.kind === 'control') {
    const c = findControl(state.ribbon, dst.id)?.control;
    if (c?.type === 'menu') allowInside = true;
  }
  if (src.kind === 'group' && dst.kind === 'tab') allowInside = true;

  if (allowInside) {
    // 中间 40% 区域 -> inside
    if (ratio > 0.3 && ratio < 0.7) return 'inside';
  }
  return ratio < 0.5 ? 'before' : 'after';
}

function clearDropHints() {
  $$('.drop-target, .drop-before, .drop-after').forEach(el => el.classList.remove('drop-target', 'drop-before', 'drop-after'));
}

function onDragStart(e) {
  const el = e.target.closest('[data-drag="1"]');
  if (!el || !el.dataset.id) return;
  // 预览菜单弹层里的子项也允许拖动
  dragCtx.src = { kind: el.dataset.ctype, id: el.dataset.id };
  dragCtx.el = el;
  el.classList.add('dragging');
  try {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', el.dataset.id);
    e.dataTransfer.setDragImage(el, 14, 12);
  } catch {}
}

function onDragOver(e) {
  if (!dragCtx.src) return;
  const el = e.target.closest('[data-ctype][data-id]');
  if (!el || !el.dataset.id) return;
  const dst = { kind: el.dataset.ctype, id: el.dataset.id };
  if (!canDrop(dragCtx.src, dst)) return;
  e.preventDefault();
  try { e.dataTransfer.dropEffect = 'move'; } catch {}

  const position = calcDropPosition(el, dragCtx.src, dst, e);
  clearDropHints();
  if (position === 'inside') el.classList.add('drop-target');
  else el.classList.add(position === 'before' ? 'drop-before' : 'drop-after');
  dragCtx.target = dst;
  dragCtx.position = position;
}

function onDrop(e) {
  if (!dragCtx.src || !dragCtx.target) return;
  e.preventDefault();
  e.stopPropagation();
  const result = moveNodeTo(state.ribbon, dragCtx.src, dragCtx.target, dragCtx.position);
  if (result.ok) {
    // 若移动后选中的节点仍在，更新选择；否则清空
    renderAll();
    toast(result.message, 'ok');
  } else {
    toast(result.message, 'err');
  }
  cleanupDrag();
}

function cleanupDrag() {
  if (dragCtx.el) dragCtx.el.classList.remove('dragging');
  dragCtx.src = null; dragCtx.el = null; dragCtx.target = null; dragCtx.position = null;
  clearDropHints();
}

function bindDragAndDrop() {
  const tree = $('#ribbon-tree');
  const preview = $('#ribbon-preview');
  [tree, preview].forEach(root => {
    if (!root) return;
    root.addEventListener('dragstart', onDragStart);
    root.addEventListener('dragover', onDragOver);
    root.addEventListener('dragleave', e => {
      // 离开根节点时清除提示
      if (!root.contains(e.relatedTarget)) clearDropHints();
    });
    root.addEventListener('drop', onDrop);
  });
  document.addEventListener('dragend', cleanupDrag);
}

/* ============ 事件绑定 ============ */
function bindEvents() {
  bindDragAndDrop();
  // 主题
  $('#btn-theme').addEventListener('click', () => {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = state.theme;
    setItem('tom-theme', state.theme);
    $('#btn-theme').textContent = state.theme === 'dark' ? '☀️' : '🌙';
  });

  // 工作区切换
  $$('.ws-btn').forEach(b => b.addEventListener('click', () => {
    state.ws = b.dataset.ws;
    $$('.ws-btn').forEach(x => x.classList.toggle('active', x === b));
    state.ribbon = clone(WORKSPACES[state.ws].ribbon());
    state.macros = sampleMacros(state.ws);
    state.file = null;
    state.selection = null;
    state.activeTabIndex = 0;
    renderAll();
    toast(`已切换到 ${WORKSPACES[state.ws].label} 工作区`, 'ok');
  }));

  // 顶部操作
  $('#btn-import').addEventListener('click', () => $('#file-input').click());
  $('#file-input').addEventListener('change', e => doImport(e.target.files[0]));
  $('#btn-save-xml').addEventListener('click', doExportXml);
  $('#btn-save-workbook').addEventListener('click', doSaveWorkbook);
  $('#btn-settings').addEventListener('click', openSettings);
  $('#btn-reset').addEventListener('click', () => {
    state.ribbon = clone(WORKSPACES[state.ws].ribbon());
    state.macros = sampleMacros(state.ws);
    state.selection = null; state.activeTabIndex = 0;
    renderAll(); toast('已重置为示例设计', 'ok');
  });
  $('#btn-refresh').addEventListener('click', () => {
    updateMacroBindings(); renderAll(); toast('已刷新宏绑定信息', 'ok');
  });
  $('#btn-fill-icons').addEventListener('click', () => {
    const names = getIconNames();
    let n = 0;
    const fill = list => (list || []).forEach(c => {
      if (!c.imageMso && c.type !== 'separator' && c.type !== 'labelControl') {
        c.imageMso = names[Math.floor(Math.random() * Math.min(names.length, 150))] || 'Bold'; n++;
      }
      if (c.items) fill(c.items);
    });
    state.ribbon.tabs.forEach(t => t.groups.forEach(g => fill(g.controls)));
    renderAll(); toast(n ? `已补全 ${n} 个图标` : '没有需要补全的图标', n ? 'ok' : '');
  });
  $('#btn-add-tab').addEventListener('click', () => {
    const t = newTab('新选项卡'); state.ribbon.tabs.push(t);
    state.selection = { kind: 'tab', tabId: t.id };
    state.activeTabIndex = state.ribbon.tabs.length - 1;
    renderAll();
  });
  $('#btn-expand-all').addEventListener('click', () => { state.collapsed.clear(); renderTree(); });
  $('#btn-collapse-all').addEventListener('click', () => {
    state.collapsed.clear();
    state.ribbon.tabs.forEach(t => { state.collapsed.add('t:' + t.id); t.groups.forEach(g => state.collapsed.add('g:' + g.id)); });
    renderTree();
  });

  // 筛选
  $('#tree-filter').addEventListener('input', e => { state.filter = e.target.value; renderTree(); });
  $('#bind-filter').addEventListener('change', e => { state.bindFilter = e.target.value; renderTree(); });

  // 结构树点击 / 右键 / 展开折叠
  $('#ribbon-tree').addEventListener('click', e => {
    const item = e.target.closest('.tree-item'); if (!item) return;
    const tw = e.target.closest('.twisty');
    if (tw && item.querySelector('.twisty')) {
      const key = item.dataset.ctype === 'tab' ? 't:' + item.dataset.id
        : item.dataset.ctype === 'group' ? 'g:' + item.dataset.id : 'c:' + item.dataset.id;
      if (state.collapsed.has(key)) state.collapsed.delete(key); else state.collapsed.add(key);
      renderTree(); return;
    }
    selectNode(item);
  });
  $('#ribbon-tree').addEventListener('contextmenu', e => {
    const item = e.target.closest('.tree-item'); if (item) showTreeContextMenu(e, item);
  });

  // 预览点击
  $('#ribbon-preview').addEventListener('click', e => {
    const tabEl = e.target.closest('.office-tab[data-tab-index]');
    if (tabEl) { state.activeTabIndex = +tabEl.dataset.tabIndex; renderPreview(); return; }
    const ctrl = e.target.closest('.office-ctrl[data-id]');
    if (ctrl) {
      if (ctrl.classList.contains('office-menu')) {
        const pop = ctrl.querySelector('.menu-popover');
        if (pop) pop.hidden = !pop.hidden;
      }
      selectNode({ dataset: { id: ctrl.dataset.id, ctype: 'control' } });
    }
  });

  // 底部标签
  $$('.bottom-tab').forEach(b => b.addEventListener('click', () => { state.bottomTab = b.dataset.tab; renderBottom(); }));
  $('#btn-copy-xml').addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(buildRibbonXml(state.ribbon)); toast('已复制 XML 到剪贴板', 'ok'); }
    catch { toast('复制失败，请手动选择', 'err'); }
  });

  // 属性 / AI 切换
  $$('.inspector-tab').forEach(b => b.addEventListener('click', () => {
    state.inspectorMode = b.dataset.mode;
    $$('.inspector-tab').forEach(x => x.classList.toggle('active', x === b));
    $('#inspector-content').hidden = state.inspectorMode !== 'inspector';
    $('#ai-content').hidden = state.inspectorMode !== 'ai';
    if (state.inspectorMode === 'ai') renderAi();
  }));

  // 拖拽导入（仅当拖入的是文件时；内部节点拖放已由 bindDragAndDrop 处理）
  window.addEventListener('dragover', e => { e.preventDefault(); });
  window.addEventListener('drop', e => {
    e.preventDefault();
    const f = e.dataTransfer.files?.[0];
    if (f && /\.(xlsm|xlam|docm|dotm)$/i.test(f.name)) doImport(f);
    else if (!dragCtx.src && f) toast('请拖入 .xlsm / .xlam / .docm 文件', 'err');
  });

  // 分栏拖拽
  $$('.splitter').forEach(sp => sp.addEventListener('mousedown', e => {
    e.preventDefault();
    const left = sp.dataset.split === 'left';
    const panel = left ? $('.tree-panel') : $('.inspector-panel');
    const startX = e.clientX, startW = panel.offsetWidth;
    const move = ev => {
      const dx = ev.clientX - startX;
      const w = left ? startW + dx : startW - dx;
      panel.style.width = Math.max(180, Math.min(520, w)) + 'px';
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  }));

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') { closeModal(); const m = $('.ctx-menu'); if (m) m.remove(); }
  });
}

/** 载入 idMso 与图标名称数据后启动 */
Promise.all([
  fetch('./js/idmso-data.json').then(r => r.json()).catch(() => []),
  fetch('./js/icon-names.json').then(r => r.json()).catch(() => []),
  fetch('./js/icon-map.json').then(r => r.json()).catch(() => ({})),
]).then(([idmso, icons, map]) => {
  window.__IDMSO__ = idmso;
  window.__ICON_NAMES__ = icons;
  window.__ICON_MAP__ = map;
  init();
}).catch(() => init());
