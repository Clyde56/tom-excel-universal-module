import { sampleRibbonExcel, clone, moveNodeTo, locateNode, containsId, findControl } from './js/model.js';

let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log('  PASS', n); } else { fail++; console.log('  FAIL', n); } };
const fresh = () => sampleRibbonExcel();

console.log('=== locateNode ===');
let r = fresh();
ok(locateNode(r, 'tab', 'tabTomTools')?.kind === 'tab', 'locate tab');
ok(locateNode(r, 'group', 'grpDataClean')?.tab.id === 'tabTomTools', 'locate group with parent tab');
ok(locateNode(r, 'control', 'btnTrim')?.group.id === 'grpDataClean', 'locate control with parent group');
ok(locateNode(r, 'control', 'qatSave')?.qat === true, 'locate QAT control');
console.log('=== containsId (防循环) ===');
const menu = findControl(r, 'mnuFormat').control;
ok(containsId(menu, 'btnCurrency') === true, 'menu contains its child');
ok(containsId(menu, 'btnTrim') === false, 'menu does not contain unrelated');
ok(moveNodeTo(r, {kind:'control',id:'mnuFormat'}, {kind:'control',id:'btnCurrency'}, 'inside').ok === false, 'block dragging menu into its own child');

console.log('=== 同列表控件排序 (before/after) ===');
r = fresh();
let g = locateNode(r, 'group', 'grpDataClean').node;
ok(g.controls[0].id === 'btnRemoveDup', 'initial order');
ok(moveNodeTo(r, {kind:'control',id:'btnBlank'}, {kind:'control',id:'btnRemoveDup'}, 'before').ok, 'move btnBlank before first');
ok(g.controls[0].id === 'btnBlank' && g.controls[1].id === 'btnRemoveDup', 'order after before-move');
r = fresh();
g = locateNode(r, 'group', 'grpDataClean').node;
moveNodeTo(r, {kind:'control',id:'btnRemoveDup'}, {kind:'control',id:'btnBlank'}, 'after');
ok(g.controls[2].id === 'btnRemoveDup', 'order after after-move');
ok(g.controls.length === 4, 'no item lost in reorder');

console.log('=== 跨分组移动 / 拖入菜单 ===');
r = fresh();
ok(moveNodeTo(r, {kind:'control',id:'btnTrim'}, {kind:'group',id:'grpReport'}, 'inside').ok, 'control into another group');
ok(locateNode(r, 'control', 'btnTrim').group.id === 'grpReport', 'control now in target group');
ok(!findControl(r, 'btnTrim') === false, 'control still findable');
r = fresh();
ok(moveNodeTo(r, {kind:'control',id:'btnChart'}, {kind:'control',id:'mnuFormat'}, 'inside').ok, 'control into menu');
ok(findControl(r, 'mnuFormat').control.items.some(i => i.id === 'btnChart'), 'menu has new child');
ok(locateNode(r, 'control', 'btnChart').parent.id === 'mnuFormat', 'parent pointer correct');

console.log('=== Group 移动 ===');
r = fresh();
ok(moveNodeTo(r, {kind:'group',id:'grpReport'}, {kind:'group',id:'grpDataClean'}, 'before').ok, 'group before another');
ok(r.tabs[0].groups[0].id === 'grpReport', 'group order updated');
// 新建第二个 tab 测跨 tab
r.tabs.push({ id:'tab2', label:'另一个', groups:[] });
ok(moveNodeTo(r, {kind:'group',id:'grpReport'}, {kind:'tab',id:'tab2'}, 'inside').ok, 'group into another tab');
ok(r.tabs[1].groups[0].id === 'grpReport', 'group moved across tabs');

console.log('=== Tab 排序 ===');
r = fresh();
r.tabs.push({ id:'tab2', label:'另一个', groups:[] });
ok(moveNodeTo(r, {kind:'tab',id:'tab2'}, {kind:'tab',id:'tabTomTools'}, 'before').ok, 'tab reorder');
ok(r.tabs[0].id === 'tab2', 'tab order updated');

console.log('=== 非法组合 ===');
r = fresh();
ok(moveNodeTo(r, {kind:'tab',id:'tabTomTools'}, {kind:'group',id:'grpReport'}, 'inside').ok === false, 'tab cannot drop into group');
ok(moveNodeTo(r, {kind:'group',id:'grpDataClean'}, {kind:'control',id:'btnTrim'}, 'inside').ok === false, 'group cannot drop into control');
ok(moveNodeTo(r, {kind:'control',id:'btnTrim'}, {kind:'control',id:'btnTrim'}, 'after').ok === false, 'self drop rejected');

console.log(`\n=== RESULT: ${pass} passed, ${fail} failed ===`);
process.exit(fail ? 1 : 0);
