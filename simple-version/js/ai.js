/**
 * 汤姆EXCEL万能模块 — AI 规划层
 * OpenAI 兼容 /chat/completions 接口，默认 DeepSeek。
 * 还原原应用的设计准则与两种模式：整份方案(JSON) / 流式操作(NDJSON)。
 */
import { getItem, setItem } from './storage.js';

export const AI_DEFAULTS = {
  providerId: 'deepseek',
  baseUrl: 'https://api.deepseek.com',
  model: 'deepseek-chat',
  apiKey: '',
};

export const AI_PROVIDERS = [
  { id: 'deepseek', label: 'DeepSeek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-chat' },
  { id: 'qwen', label: '通义千问', baseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1', model: 'qwen-plus' },
  { id: 'kimi', label: 'Kimi', baseUrl: 'https://api.moonshot.cn/v1', model: 'moonshot-v1-8k' },
  { id: 'glm', label: '智谱 GLM', baseUrl: 'https://open.bigmodel.cn/api/paas/v4', model: 'glm-4-flash' },
  { id: 'custom', label: '自定义 (OpenAI 兼容)', baseUrl: '', model: '' },
];

const SETTINGS_KEY = 'tom-excel-ai-settings';

export function loadAiSettings() {
  try {
    const raw = getItem(SETTINGS_KEY);
    if (raw) return { ...AI_DEFAULTS, ...JSON.parse(raw) };
  } catch {}
  return { ...AI_DEFAULTS };
}
export function saveAiSettings(s) {
  setItem(SETTINGS_KEY, JSON.stringify(s));
}

/** 共享设计准则（改写自原应用，品牌改为汤姆EXCEL万能模块） */
function designGuidelines() {
  return [
    '设计目标：像资深 Office 产品设计师一样规划 Ribbon，而不是按宏清单机械铺按钮。',
    '先在内部完成信息架构：识别用户任务、使用频率、风险等级和宏之间的先后关系，再输出结果。',
    '默认只创建 1 个自定义选项卡，把不同任务阶段拆成多个逻辑分组；只有用户明确要求多个选项卡时才创建多个。',
    '分组代表任务阶段或对象类别，例如导入、清洗、分析、格式、输出、维护，禁止使用「功能1/其他」这类弱分组。',
    '每个分组建议 2-5 个直接可见控件；超过 5 个相关动作时应使用 menu 收纳低频项。',
    '每个分组默认最多 1 个 large 主入口按钮，其余用 normal，禁止所有按钮都是 large。',
    '禁止创建只有 1 个子控件的 menu，menu 至少包含 2 个相关子动作。',
    '控件顺序遵循工作流：准备/导入 → 处理/计算 → 审核/预览 → 输出/维护。',
    'Label 要短、可执行、面向用户，不要照搬英文宏名；screentip 说明动作，supertip 说明结果与注意事项。',
    '所有面向用户的中文文本不要以句号结尾，尽量不用句号分隔。',
    '只能绑定 macros 中真实存在的宏，targetMacroName 必须逐字等于宏 name，targetMacroModule 必须等于宏 module。',
    'output 只返回 JSON，不要 Markdown 代码块或解释性散文。',
  ].map((x, i) => `${i + 1}. ${x}`).join('\n');
}

export function buildPlanPrompt(userRequest, macros, availableImageMsos, applyMode = 'replace') {
  const system = [
    '你是「汤姆EXCEL万能模块」内置的 Office Ribbon 信息架构设计助手。',
    designGuidelines(),
    '根据宏列表规划 tab、group、control、按钮名称、imageMso 与宏绑定。',
    '返回格式：{"summary":string,"tabs":[{"label":string,"groups":[{"label":string,"controls":[{"type":"button","label":string,"imageMso":string,"size":"large|normal","onAction":"模块.过程名","screentip":string,"supertip":string}]}]}]}',
    'applyMode=replace 时重新规划整套 Ribbon；append 时只追加；modify 时基于 currentRibbon 做最小改动。',
  ].join('\n');
  const user = JSON.stringify({
    userRequest, applyMode,
    macros: macros.map(m => ({ module: m.module, name: m.name })),
    availableImageMsos: availableImageMsos.slice(0, 200),
  });
  return { system, user };
}

export function endpointOf(baseUrl) {
  const b = (baseUrl || '').replace(/\/+$/, '');
  return b.endsWith('/chat/completions') ? b : `${b}/chat/completions`;
}

function headers(settings) {
  return { 'Content-Type': 'application/json', Authorization: `Bearer ${settings.apiKey}` };
}

/** 连通性测试 */
export async function testConnection(settings) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), 15000);
  try {
    const res = await fetch(endpointOf(settings.baseUrl), {
      method: 'POST', headers: headers(settings), signal: controller.signal,
      body: JSON.stringify({ model: settings.model, messages: [{ role: 'user', content: 'ping' }], max_tokens: 8, temperature: 0 }),
    });
    const text = await res.text();
    return { ok: res.ok, status: res.status, message: res.ok ? '连接成功' : text.slice(0, 200) };
  } catch (e) {
    return { ok: false, status: 0, message: e.name === 'AbortError' ? '连接超时' : (e.message || '连接失败') };
  } finally { clearTimeout(t); }
}

/** 非流式规划，返回 {summary, tabs} */
export async function planRibbon(settings, userRequest, macros, imageMsos, applyMode = 'replace') {
  const { system, user } = buildPlanPrompt(userRequest, macros, imageMsos, applyMode);
  let body = {
    model: settings.model, temperature: 0.2,
    messages: [{ role: 'system', content: system }, { role: 'user', content: user }],
    response_format: { type: 'json_object' },
  };
  if (/^deepseek-v\d/i.test(settings.model)) body.thinking = { type: 'disabled' };
  let res = await fetch(endpointOf(settings.baseUrl), { method: 'POST', headers: headers(settings), body: JSON.stringify(body) });
  if (res.status === 400) {
    delete body.response_format;
    res = await fetch(endpointOf(settings.baseUrl), { method: 'POST', headers: headers(settings), body: JSON.stringify(body) });
  }
  if (!res.ok) throw new Error(`AI 请求失败 ${res.status}: ${(await res.text()).slice(0, 200)}`);
  const json = await res.json();
  const content = json.choices?.[0]?.message?.content || '';
  return parsePlan(content);
}

function parsePlan(text) {
  let clean = String(text).trim();
  clean = clean.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  const s = clean.indexOf('{'), e = clean.lastIndexOf('}');
  if (s < 0 || e < 0) throw new Error('AI 未返回有效 JSON 方案');
  const obj = JSON.parse(clean.slice(s, e + 1));
  if (!Array.isArray(obj.tabs)) throw new Error('AI 方案缺少 tabs 数组');
  return obj;
}

/** 流式规划：逐行回调 onDelta(text) / onOperation(op)，返回完整文本 */
export async function planRibbonStream(settings, userRequest, macros, imageMsos, applyMode, onDelta, signal) {
  const { system, user } = buildPlanPrompt(userRequest, macros, imageMsos, applyMode);
  const streamSystem = system + '\n请用 NDJSON 输出：每行一个完整 JSON 操作对象。操作类型 insertControl/insertControls/createTab/createGroup/message/finish。优先用 createTab、createGroup、insertControls 一次插入多个控件，最后输出 {"type":"finish","summary":"..."}。';
  const body = {
    model: settings.model, temperature: 0.2, stream: true,
    messages: [{ role: 'system', content: streamSystem }, { role: 'user', content: user }],
  };
  if (/^deepseek-v\d/i.test(settings.model)) body.thinking = { type: 'disabled' };
  const res = await fetch(endpointOf(settings.baseUrl), { method: 'POST', headers: headers(settings), body: JSON.stringify(body), signal });
  if (!res.ok || !res.body) throw new Error(`AI 流式请求失败 ${res.status}`);
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '', full = '';
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop();
    for (const line of lines) {
      if (!line.startsWith('data:')) continue;
      const data = line.slice(5).trim();
      if (data === '[DONE]') continue;
      try {
        const j = JSON.parse(data);
        const d = j.choices?.[0]?.delta?.content || j.choices?.[0]?.delta?.reasoning_content || '';
        if (d) { full += d; onDelta?.(d); }
      } catch {}
    }
  }
  return full;
}

/** 把 AI 方案对象转成 Ribbon 结构 */
export function planToRibbon(plan) {
  const ribbon = { startFromScratch: false, tabs: [], qat: [] };
  (plan.tabs || []).forEach((t, ti) => {
    const tab = { id: `tabAi${Date.now().toString(36)}${ti}`, label: t.label || `选项卡${ti + 1}`, groups: [] };
    (t.groups || []).forEach((g, gi) => {
      const group = { id: `grpAi${ti}_${gi}_${uidNum()}`, label: g.label || `分组${gi + 1}`, controls: [] };
      (g.controls || []).forEach(c => {
        const ctl = {
          type: c.type || 'button',
          id: `ctrlAi${uidNum()}`,
          label: c.label || '按钮',
          imageMso: c.imageMso || '',
          size: c.size === 'large' ? 'large' : 'normal',
          onAction: c.onAction || c.targetMacroModule + '.' + c.targetMacroName || '',
          onChange: c.onChange || '',
          screentip: c.screentip || '',
          supertip: c.supertip || '',
          visible: true, enabled: true,
        };
        if (ctl.type === 'menu') {
          ctl.items = (c.items || []).map(it => ({
            type: it.type || 'button', id: `ctrlAi${uidNum()}`,
            label: it.label || '项目', imageMso: it.imageMso || '',
            onAction: it.onAction || `${it.targetMacroModule || ''}.${it.targetMacroName || ''}`.replace(/^\./, ''),
            visible: true, enabled: true,
          }));
        }
        group.controls.push(ctl);
      });
      tab.groups.push(group);
    });
    ribbon.tabs.push(tab);
  });
  return ribbon;
}
let _u = 0;
function uidNum() { return (Date.now() % 100000).toString(36) + (_u++).toString(36) + Math.floor(Math.random() * 1000).toString(36); }
