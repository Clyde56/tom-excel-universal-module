# 汤姆EXCEL万能模块（Tom Excel Universal Module）

一个用于设计 Office Ribbon 功能区（customUI.xml）的桌面应用，可为 Excel / Word / WPS 工作簿设计工具栏、绑定 VBA / JSA 宏，并支持 AI 辅助设计与加载项安装包制作。

本项目是一个 **fork / 衍生作品**，基于 `jerry-office/jr-ribbon-xml-editor` v1.0.9 改名而来，遵循原项目的 Apache-2.0 协议。详见 [NOTICE](NOTICE) 与 [来源与授权说明](#来源与授权说明)。

> ⚠️ **合规提示**：本仓库保留了原项目 Apache-2.0 许可声明与原作者归属。图标资源（ICO 目录）中的 Office 图标源自 Microsoft Office 的 imageMso 资源，其权利归 Microsoft 所有；如用于商业分发，请自行评估并获取相应授权。

---

## 功能特性

- **Ribbon 可视化设计**：以树形结构组织 Tab / Group / Control（按钮、菜单、下拉框、库等），支持拖放排序与跨容器移动
- **宏识别与绑定**：从 `.xlsm` / `.xlam` 的 `vbaProject.bin` 提取 VBA 过程，或从 WPS 的 `JDEData.bin` 提取 JSA 函数；自动检测回调绑定关系
- **真实写回工作簿**：通过 adm-zip 改写 OOXML part（customUI14.xml、rels、[Content_Types].xml），完整保留原 VBA 工程；通过 PowerShell COM 向工作簿写入 `TOMExcelCallbacks` 回调模块
- **图标库**：内置 6300+ 个 Office 2019 imageMso 图标（Excel / Word 两套），支持自定义图片、语义匹配、随机补全
- **AI 辅助设计**：OpenAI 兼容接口（默认 DeepSeek），Agent 模式多轮 function calling（10 个工具）自动规划 Ribbon；支持流式输出、缺失宏绑定修复
- **安装包制作**：生成 Windows EXE（IExpress）/ VBS 安装脚本，以及面向国产 Linux（统信/麒麟 + WPS）的 DEB 包
- **VBA 访问授权**：检测到 `AccessVBOM` 未启用时，弹出确认并通过注册表自动开启 Office 的 VBA 工程对象模型访问

---

## 源码构成

| 路径 | 说明 |
|---|---|
| `electron/dist/main/main.cjs` | 主进程完整实现（esbuild bundle，保留函数名与源码注释，约 1.3 万行）。包含 openxml-package、vba-inspector、vba-callback-runtime、installer-windows 等全部业务模块 |
| `electron/dist/main/preload.cjs` | Electron contextBridge 桥接层（暴露 `window.tomExcel`，27 个方法） |
| `electron/dist/renderer/` | 渲染进程（React 19 生产构建，**minified**，功能 1:1 但原始 TSX 源码不可逆） |
| `electron/ICO/` | 内置 Office imageMso 图标（6304 个 PNG）——**因版权未随本仓库分发**，运行前需自行准备，见下方说明 |
| `simple-version/` | 另附的零依赖 Web 重写版（干净模块化 ES modules 源码，功能为核心子集，可直接打开 HTML 运行） |
| `rebrand.py` | 改名脚本：把原项目的品牌字符串（含 Unicode 转义）系统性替换为本项目品牌 |

**关于"1:1 复刻"的说明**：主进程 bundle 由 esbuild 生成、未压缩，保留了完整业务逻辑与注释，可直接阅读与维护；渲染进程为 Vite/React 生产构建，已 minified，功能与原始版本一致，但原始 JSX 源码无法通过静态手段还原。若需要完全可维护的渲染层源码，请基于 `simple-version/` 的模块化实现继续开发。

---

## 运行与构建

```bash
cd electron
npm install            # 安装运行依赖（清单见 package.json）
npm start              # 启动应用（需本机安装 Electron）
```

打包可执行文件：

```bash
npm install -D electron-builder
npx electron-builder --win   # 或 --linux
```

> 注：依赖需通过 `npm install` 安装。图标资源（`electron/ICO/`、`simple-version/icons/`）因源自 Microsoft Office、**版权原因未随仓库分发**，应用在没有图标时仍可启动，但图标选择器会显示占位符。若需完整图标体验，请从本机安装的 Office/WPS 中获取对应 imageMso 图标，或自行准备同名 PNG 放入该目录（图标文件命名格式：`{数字前缀}_{imageMso名}_32.png`）。

---

## 已打包的可运行版本

`build/汤姆EXCEL万能模块/` 是已经完成打包、可直接运行的 Windows 桌面版（约 363 MB）：

- `汤姆EXCEL万能模块.exe` — Electron 主程序（已改名）
- `resources/app.asar` — 由 `electron/` 目录打包出的应用代码（14 MB，含全部改名后的主进程/渲染逻辑与依赖）
- `启动应用.bat` — Windows 启动脚本

**运行**：双击 `启动应用.bat` 或 `汤姆EXCEL万能模块.exe`。

> 注意：若在某些 AI/开发工具的内置终端里运行，`ELECTRON_RUN_AS_NODE=1` 环境变量会导致 Electron 被当成 Node 启动而闪退；`启动应用.bat` 已自动清除该变量。普通 Windows 资源管理器双击不受影响。

**重新打包**：

```bash
npx @electron/asar pack electron build/app.asar
copy build/app.asar build/汤姆EXCEL万能模块/resources/app.asar
```

---

## 来源与授权说明

- **原项目**：`jerry-office/jr-ribbon-xml-editor`，版本 1.0.9，原作者 jerry-office，许可证 Apache-2.0
- **本项目修改内容**：全面改名（品牌、标识符、存储键、回调模块名、VBA 失败标记等）；移除原作者个人站点与打赏信息；改为开源仓库结构
- **授权**：本项目沿用 Apache License 2.0，完整条款见 [LICENSE](LICENSE)，归属声明见 [NOTICE](NOTICE)

本项目为技术研究与开源学习目的发布。原项目的所有权利归原作者所有。若原作者不希望本项目继续发布，请联系维护者，我们将及时下架。
