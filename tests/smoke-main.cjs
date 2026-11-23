// 冒烟测试：mock electron，加载改名后的 main.cjs，验证无语法/引用错误
const Module = require('module');
const path = require('path');
const origLoad = Module._load;

// 最小 Electron stub
const browserWindows = [];
class BrowserWindow {
  constructor(opts) { this.opts = opts; this.webContents = {
    on() {}, send() {}, session: { setPermissionRequestHandler() {} },
    setWindowOpenHandler() {}, openDevTools() {},
  }; browserWindows.push(this); }
  static getAllWindows() { return browserWindows; }
  once() {} on() {} loadFile() {} show() {}
  setMenu() {} isMinimized() { return false; } restore() {}
  isVisible() { return true; } focus() {}
}
const ipcHandlers = {};
const stub = {
  app: {
    getPath: (n) => path.join(require('os').homedir(), '.tom-test', n),
    getAppPath: () => path.join(__dirname, 'electron'),
    getVersion: () => '1.0.9',
    getName: () => 'tom-excel-universal-module',
    whenReady: () => Promise.resolve(),
    on: () => {}, quit: () => {},
    setName: () => {}, setPath: () => {}, requestSingleInstanceLock: () => true,
    getLocale: () => 'zh-CN', getSystemLocale: () => 'zh-CN',
    isPackaged: false,
    commandLine: { appendSwitch() {} },
  },
  BrowserWindow,
  ipcMain: {
    handle: (ch, fn) => { ipcHandlers[ch] = fn; },
    on: () => {}, removeHandler: () => {},
  },
  dialog: {
    showOpenDialog: async () => ({ canceled: true, filePaths: [] }),
    showSaveDialog: async () => ({ canceled: true, filePath: '' }),
    showMessageBox: async () => ({ response: 0 }),
  },
  shell: { openExternal: async () => {}, openPath: async () => {} },
  contextBridge: { exposeInMainWorld: () => {} },
  webUtils: { getPathForFile: () => '' },
  safeStorage: { decryptString: (b) => b.toString(), encryptString: (s) => Buffer.from(s) },
  Menu: { buildFromTemplate: () => ({}), setApplicationMenu: () => {} },
  nativeImage: { createFromPath: () => ({}) },
  screen: { getPrimaryDisplay: () => ({ workArea: { x:0,y:0,width:1280,height:800 } }) },
};

Module._load = function(request, parent, isMain) {
  if (request === 'electron') return stub;
  if (request === 'electron/main') return stub;
  if (request === 'electron/common') return stub;
  return origLoad.apply(this, arguments);
};

const mainPath = path.join(__dirname, 'electron', 'dist', 'main', 'main.cjs');
let failed = false;
try {
  require(mainPath);
  console.log('  PASS main.cjs loaded without errors');
} catch (e) {
  failed = true;
  console.log('  FAIL main.cjs load:', e.message);
  console.log(e.stack.split('\n').slice(0, 4).join('\n'));
}
console.log('  IPC handlers registered:', Object.keys(ipcHandlers).length);
const expected = ['ai:get-settings','ai:plan-ribbon-stream','workbook:inspect','workbook:save-copy-to-path','icon:list-mso','workbook:create-installer-package','vba-access:respond'];
for (const ch of expected) {
  if (ipcHandlers[ch]) console.log('  PASS handler:', ch);
  else { failed = true; console.log('  FAIL missing handler:', ch); }
}
process.exit(failed ? 1 : 0);
