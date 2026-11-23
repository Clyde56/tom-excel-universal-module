/**
 * 汤姆EXCEL万能模块 — 数据模型层
 * 定义 Ribbon 的 Tab / Group / Control 树结构，以及工作区状态与示例数据。
 * 结构与原应用导出的 customUI14.xml 一一对应，保证可逆转换。
 */

export const CONTROL_TYPES = [
  'button', 'toggleButton', 'checkBox', 'editBox', 'comboBox',
  'dropDown', 'gallery', 'menu', 'labelControl', 'separator',
];

/** 生成一个唯一 id */
let _seq = 0;
export function uid(prefix = 'id') {
  _seq += 1;
  return `${prefix}${Date.now().toString(36)}${_seq.toString(36)}`;
}

export function newControl(type = 'button', label = '新按钮') {
  const base = { type, id: uid('ctrl'), label, visible: true, enabled: true };
  if (type === 'button' || type === 'toggleButton' || type === 'checkBox' || type === 'menu') {
    base.onAction = '';
    base.imageMso = '';
    base.size = 'normal';
    base.screentip = '';
    base.supertip = '';
  }
  if (type === 'menu') base.items = [];
  if (type === 'editBox' || type === 'comboBox' || type === 'dropDown' || type === 'gallery') {
    base.onChange = '';
  }
  return base;
}

export function newGroup(label = '新分组') {
  return { id: uid('grp'), label, controls: [] };
}

export function newTab(label = '新选项卡') {
  return { id: uid('tab'), label, groups: [] };
}

/** 深拷贝 */
export function clone(o) {
  return JSON.parse(JSON.stringify(o));
}

/** 默认示例 Ribbon（Excel 工作区） */
export function sampleRibbonExcel() {
  return {
    startFromScratch: false,
    tabs: [
      {
        id: 'tabTomTools', label: '汤姆工具箱', groups: [
          {
            id: 'grpDataClean', label: '数据清洗', controls: [
              { type: 'button', id: 'btnRemoveDup', label: '删除重复项', imageMso: 'RemoveDuplicates', size: 'large', onAction: 'CleanData.RemoveDuplicates', screentip: '删除重复项', supertip: '基于当前选区一键移除重复行', visible: true, enabled: true },
              { type: 'button', id: 'btnTrim', label: '去空格', imageMso: 'WrapText', size: 'normal', onAction: 'CleanData.TrimText', screentip: '去除多余空格', supertip: '清理文本首尾及连续空格', visible: true, enabled: true },
              { type: 'button', id: 'btnBlank', label: '填空值', imageMso: 'FillDown', size: 'normal', onAction: 'CleanData.FillBlanks', screentip: '填充空白单元格', supertip: '用上方或默认值填充选中区域的空单元格', visible: true, enabled: true },
              { type: 'menu', id: 'mnuFormat', label: '格式工具', imageMso: 'FormatPainter', size: 'large', onAction: '', items: [
                { type: 'button', id: 'btnCurrency', label: '货币格式', imageMso: 'AccountingFormat', onAction: 'Format.ApplyCurrency', visible: true, enabled: true },
                { type: 'button', id: 'btnPercent', label: '百分比', imageMso: 'PercentStyle', onAction: 'Format.ApplyPercent', visible: true, enabled: true },
                { type: 'separator', id: 'sep1' },
                { type: 'button', id: 'btnBorder', label: '边框', imageMso: 'BordersGallery', onAction: 'Format.ApplyBorder', visible: true, enabled: true },
              ], visible: true, enabled: true },
            ],
          },
          {
            id: 'grpReport', label: '报表输出', controls: [
              { type: 'button', id: 'btnExportPdf', label: '导出 PDF', imageMso: 'FileSaveAsPdfOrXps', size: 'large', onAction: 'Report.ExportPdf', screentip: '导出为 PDF', supertip: '将当前工作簿导出为 PDF 文件', visible: true, enabled: true },
              { type: 'button', id: 'btnChart', label: '生成图表', imageMso: 'ChartStylesGallery', size: 'normal', onAction: 'Report.BuildChart', screentip: '生成图表', supertip: '根据当前数据区域自动生成柱形图', visible: true, enabled: true },
            ],
          },
        ],
      },
    ],
    qat: [
      { type: 'button', id: 'qatSave', label: '快速保存', imageMso: 'FileSave', onAction: 'Quick.Save', visible: true, enabled: true },
    ],
  };
}

export function sampleRibbonWord() {
  return {
    startFromScratch: false,
    tabs: [
      {
        id: 'tabTomWord', label: '汤姆文档工具', groups: [
          {
            id: 'grpDocFmt', label: '文档排版', controls: [
              { type: 'button', id: 'btnToc', label: '生成目录', imageMso: 'NumberingGallery', size: 'large', onAction: 'Doc.BuildToc', screentip: '生成目录', supertip: '按标题样式自动生成文档目录', visible: true, enabled: true },
              { type: 'button', id: 'btnPageNum', label: '插入页码', imageMso: 'HeaderFooterPageNumberInsertExcel', size: 'normal', onAction: 'Doc.InsertPageNumber', visible: true, enabled: true },
            ],
          },
        ],
      },
    ],
    qat: [],
  };
}

export function sampleRibbonJsa() {
  return {
    startFromScratch: false,
    tabs: [
      {
        id: 'tabTomJsa', label: '汤姆JS宏工具', groups: [
          {
            id: 'grpJsaClean', label: '数据处理', controls: [
              { type: 'button', id: 'btnJsaClean', label: '一键清洗', imageMso: 'RemoveDuplicates', size: 'large', onAction: 'jsaCleanData', visible: true, enabled: true },
              { type: 'button', id: 'btnJsaSum', label: '汇总统计', imageMso: 'AutoSum', size: 'normal', onAction: 'jsaSummarize', visible: true, enabled: true },
            ],
          },
        ],
      },
    ],
    qat: [],
  };
}

/** 示例宏过程列表（未导入文件时的演示数据） */
export function sampleMacros(kind = 'excel') {
  if (kind === 'word') {
    return [
      { module: 'DocTools', name: 'BuildToc', signature: 'Public Sub BuildToc()', scope: 'public', kind: 'Sub', bound: 'btnToc' },
      { module: 'DocTools', name: 'InsertPageNumber', signature: 'Public Sub InsertPageNumber()', scope: 'public', kind: 'Sub', bound: 'btnPageNum' },
      { module: 'Format', name: 'NormalizeStyle', signature: 'Public Sub NormalizeStyle()', scope: 'public', kind: 'Sub', bound: '' },
    ];
  }
  if (kind === 'jsa') {
    return [
      { module: 'main', name: 'jsaCleanData', signature: 'function jsaCleanData()', scope: 'public', kind: 'Function', bound: 'btnJsaClean' },
      { module: 'main', name: 'jsaSummarize', signature: 'function jsaSummarize()', scope: 'public', kind: 'Function', bound: 'btnJsaSum' },
      { module: 'utils', name: 'jsaExport', signature: 'function jsaExport()', scope: 'public', kind: 'Function', bound: '' },
    ];
  }
  return [
    { module: 'CleanData', name: 'RemoveDuplicates', signature: 'Public Sub RemoveDuplicates()', scope: 'public', kind: 'Sub', bound: 'btnRemoveDup' },
    { module: 'CleanData', name: 'TrimText', signature: 'Public Sub TrimText()', scope: 'public', kind: 'Sub', bound: 'btnTrim' },
    { module: 'CleanData', name: 'FillBlanks', signature: 'Public Sub FillBlanks()', scope: 'public', kind: 'Sub', bound: 'btnBlank' },
    { module: 'Format', name: 'ApplyCurrency', signature: 'Public Sub ApplyCurrency()', scope: 'public', kind: 'Sub', bound: 'btnCurrency' },
    { module: 'Format', name: 'ApplyPercent', signature: 'Public Sub ApplyPercent()', scope: 'public', kind: 'Sub', bound: 'btnPercent' },
    { module: 'Format', name: 'ApplyBorder', signature: 'Public Sub ApplyBorder()', scope: 'public', kind: 'Sub', bound: 'btnBorder' },
    { module: 'Report', name: 'ExportPdf', signature: 'Public Sub ExportPdf()', scope: 'public', kind: 'Sub', bound: 'btnExportPdf' },
    { module: 'Report', name: 'BuildChart', signature: 'Public Sub BuildChart()', scope: 'public', kind: 'Sub', bound: 'btnChart' },
    { module: 'Quick', name: 'Save', signature: 'Public Sub Save()', scope: 'public', kind: 'Sub', bound: 'qatSave' },
    { module: 'Utils', name: 'UnusedProc', signature: 'Public Sub UnusedProc()', scope: 'public', kind: 'Sub', bound: '' },
  ];
}

export const WORKSPACES = {
  excel: { label: 'Excel', extensions: ['xlsm', 'xlam'], appName: 'Excel', ribbon: sampleRibbonExcel },
  word:  { label: 'Word', extensions: ['docm', 'dotm'], appName: 'Word', ribbon: sampleRibbonWord },
  jsa:   { label: 'JS宏', extensions: ['xlsm', 'xlam', 'docm', 'dotm'], appName: 'WPS', ribbon: sampleRibbonJsa },
};

/** 遍历所有控件（含 menu 子项），带回调 */
export function walkControls(controls, cb, parent = null) {
  (controls || []).forEach((c, i) => {
    cb(c, i, controls, parent);
    if (c.items) walkControls(c.items, cb, c);
  });
}

/** 根据 id 查找控件，返回 {control, siblings, index, tab, group, parent} */
export function findControl(ribbon, id) {
  let found = null;
  for (const tab of ribbon.tabs) {
    for (const group of tab.groups) {
      walkControls(group.controls, (c, i, siblings, parent) => {
        if (c.id === id) found = { control: c, siblings, index: i, tab, group, parent };
      });
    }
  }
  (ribbon.qat || []).forEach((c, i, arr) => {
    if (c.id === id) found = { control: c, siblings: arr, index: i, tab: null, group: null, parent: null, qat: true };
  });
  return found;
}

/* ============ 拖放移动（模型层） ============ */

/** 按 kind+id 在 Ribbon 树中定位任意节点，返回完整上下文 */
export function locateNode(ribbon, kind, id) {
  if (kind === 'tab') {
    const i = ribbon.tabs.findIndex(t => t.id === id);
    return i < 0 ? null : { kind: 'tab', node: ribbon.tabs[i], siblings: ribbon.tabs, index: i };
  }
  if (kind === 'group') {
    for (const tab of ribbon.tabs) {
      const i = tab.groups.findIndex(g => g.id === id);
      if (i >= 0) return { kind: 'group', node: tab.groups[i], siblings: tab.groups, index: i, tab };
    }
    return null;
  }
  if (kind === 'control') {
    const f = findControl(ribbon, id);
    if (!f) return null;
    return {
      kind: 'control', node: f.control, siblings: f.siblings, index: f.index,
      tab: f.tab, group: f.group, parent: f.parent, qat: f.qat,
    };
  }
  return null;
}

/** 判断 node 的子树里是否含有 id（用于防止把节点拖进自己的后代） */
export function containsId(node, id) {
  if (node.id === id) return true;
  for (const c of node.items || []) if (containsId(c, id)) return true;
  return false;
}

/**
 * 把来源节点移动到目标位置。
 * 参数均为 {kind, id}，position 为 'before' | 'after' | 'inside'。
 * 返回 {ok, message}。
 *
 * 合法组合：
 *   tab    -> tab(before/after)
 *   group  -> group(before/after) | tab(inside)
 *   control-> control(before/after) | group(inside) | menu(inside，成为子项) | qat(inside)
 */
export function moveNodeTo(ribbon, src, dst, position = 'after') {
  if (!src || !dst) return { ok: false, message: '无效的拖放对象' };
  if (src.kind === dst.kind && src.id === dst.id) return { ok: false, message: '不能拖到自己身上' };

  const from = locateNode(ribbon, src.kind, src.id);
  const to = locateNode(ribbon, dst.kind, dst.id);
  if (!from || !to) return { ok: false, message: '节点不存在' };

  // 防止把节点拖入自己的子树
  if (src.kind === 'control' && dst.kind === 'control' && containsId(from.node, dst.id)) {
    return { ok: false, message: '不能把控件拖入它自己的子项' };
  }
  if (src.kind === 'group' && dst.kind === 'group' && from.node.id === to.node.id) {
    return { ok: false, message: '不能拖到自己身上' };
  }

  let targetList, targetIndex;

  // ---- Tab ----
  if (src.kind === 'tab') {
    if (dst.kind !== 'tab') return { ok: false, message: '选项卡只能在选项卡之间排序' };
    targetList = ribbon.tabs;
    const di = ribbon.tabs.findIndex(t => t.id === dst.id);
    targetIndex = position === 'before' ? di : di + 1;
    from.siblings.splice(from.index, 1);
    if (targetIndex > from.index) targetIndex--;
    ribbon.tabs.splice(targetIndex, 0, from.node);
    return { ok: true, message: '已移动选项卡' };
  }

  // ---- Group ----
  if (src.kind === 'group') {
    if (dst.kind === 'tab') {
      const tab = to.node;
      from.siblings.splice(from.index, 1);
      tab.groups.push(from.node);
      return { ok: true, message: `已把分组移入「${tab.label}」` };
    }
    if (dst.kind === 'group') {
      if (position === 'inside') return { ok: false, message: '分组不能嵌套在分组里' };
      targetList = to.siblings;
      const di = to.siblings.findIndex(g => g.id === dst.id);
      targetIndex = position === 'before' ? di : di + 1;
      if (from.siblings === to.siblings) {
        from.siblings.splice(from.index, 1);
        if (targetIndex > from.index) targetIndex--;
      } else {
        from.siblings.splice(from.index, 1);
      }
      targetList.splice(targetIndex, 0, from.node);
      return { ok: true, message: '已移动分组' };
    }
    return { ok: false, message: '分组只能拖到其他分组或选项卡中' };
  }

  // ---- Control ----
  if (src.kind === 'control') {
    if (dst.kind === 'tab') return { ok: false, message: '请把控件拖到分组或菜单里' };
    if (dst.kind === 'group') {
      const group = to.node;
      from.siblings.splice(from.index, 1);
      group.controls.push(from.node);
      return { ok: true, message: `已把控件移入「${group.label}」` };
    }
    if (dst.kind === 'control') {
      const dstCtl = to.node;
      // 拖入 menu 内部成为子项
      if (position === 'inside' && dstCtl.type === 'menu') {
        from.siblings.splice(from.index, 1);
        dstCtl.items = dstCtl.items || [];
        dstCtl.items.push(from.node);
        return { ok: true, message: `已把控件移入菜单「${dstCtl.label}」` };
      }
      // 普通排序：before / after
      const usePosition = (position === 'inside') ? 'after' : position;
      targetList = to.siblings;
      const di = to.siblings.findIndex(c => c.id === dst.id);
      targetIndex = usePosition === 'before' ? di : di + 1;
      if (from.siblings === to.siblings) {
        from.siblings.splice(from.index, 1);
        if (targetIndex > from.index) targetIndex--;
      } else {
        from.siblings.splice(from.index, 1);
      }
      targetList.splice(targetIndex, 0, from.node);
      return { ok: true, message: '已移动控件' };
    }
  }

  return { ok: false, message: '不支持这种拖放组合' };
}

/** 判断某种拖放组合是否合法（用于拖拽时给出视觉提示） */
export function canDrop(src, dst) {
  if (!src || !dst) return false;
  if (src.kind === dst.kind && src.id === dst.id) return false;
  const pairs = {
    tab: ['tab'],
    group: ['group', 'tab'],
    control: ['control', 'group', 'tab'], // tab 仅用于位置提示，实际细查在 moveNodeTo
  };
  return (pairs[src.kind] || []).includes(dst.kind);
}
