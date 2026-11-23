/**
 * 安全的本地存储封装。
 * 在 file://、隐私模式或某些沙箱环境里 localStorage 可能不可用，
 * 这里降级到内存存储，保证应用始终能启动。
 */
const mem = new Map();
let available = null;

function check() {
  if (available !== null) return available;
  try {
    const k = '__tom_test__';
    window.localStorage.setItem(k, '1');
    window.localStorage.removeItem(k);
    available = true;
  } catch {
    available = false;
  }
  return available;
}

export function getItem(key) {
  if (check()) { try { return window.localStorage.getItem(key); } catch { /* fallthrough */ } }
  return mem.has(key) ? mem.get(key) : null;
}
export function setItem(key, value) {
  mem.set(key, String(value));
  if (check()) { try { window.localStorage.setItem(key, String(value)); } catch { /* ignore */ } }
}
export function removeItem(key) {
  mem.delete(key);
  if (check()) { try { window.localStorage.removeItem(key); } catch { /* ignore */ } }
}
