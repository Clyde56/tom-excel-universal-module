# -*- coding: utf-8 -*-
"""
构建脚本：
1) 生成 js/bundle.js —— 把所有 ES module 合并成普通 script，内联 JSON 数据，
   供 index.html 在 file:// 协议下直接加载（不再用 ES module + fetch，绕开 CORS）
2) 生成 汤姆EXCEL万能模块-独立版.html —— 把 CSS 与 bundle.js 内联进 HTML，
   双击即用，连 bundle.js 文件都不依赖
"""
import os, re

BASE = os.path.dirname(os.path.abspath(__file__))
JSDIR = os.path.join(BASE, 'js')

def read(p):
    with open(p, encoding='utf-8') as f: return f.read()

# 按依赖顺序拼接（storage -> model -> xml -> zip -> workbook -> ai -> app）
order = ['storage.js', 'model.js', 'xml.js', 'zip.js', 'workbook.js', 'ai.js', 'app.js']
parts = []
for name in order:
    src = read(os.path.join(JSDIR, name))
    src = re.sub(r'^import\s+[\s\S]*?from\s+[\'"][^\'"]+[\'"]\s*;?\s*$', '', src, flags=re.M)
    src = re.sub(r'^import\s+[\'"].*?[\'"]\s*;?\s*$', '', src, flags=re.M)
    src = re.sub(r'^export\s+', '', src, flags=re.M)
    parts.append(f'\n/* ===== {name} ===== */\n' + src)

idmso = read(os.path.join(JSDIR, 'idmso-data.json'))
icons = read(os.path.join(JSDIR, 'icon-names.json'))
iconmap = read(os.path.join(JSDIR, 'icon-map.json'))
data_js = f'window.__IDMSO__ = {idmso};\nwindow.__ICON_NAMES__ = {icons};\nwindow.__ICON_MAP__ = {iconmap};\n'

merged = '\n'.join(parts)
fetch_block = re.compile(r'/\*\* 载入 idMso 与图标名称数据后启动 \*/[\s\S]*?\.catch\(\(\) => init\(\)\);?')
merged, n = fetch_block.subn('init();', merged)
print('replaced startup block:', n)

# 用 IIFE 包裹，避免 $ 等变量名与宿主环境（浏览器扩展/测试环境）全局冲突
bundle = '/* 汤姆EXCEL万能模块 bundle — file:// 可直接运行 */\n(function(){\n' + data_js + '\n' + merged + '\n})();\n'
bundle_path = os.path.join(JSDIR, 'bundle.js')
with open(bundle_path, 'w', encoding='utf-8') as f:
    f.write(bundle)
print('built:', bundle_path, os.path.getsize(bundle_path), 'bytes')

# 独立版：内联 CSS 与 bundle.js
html = read(os.path.join(BASE, 'index.html'))
css = read(os.path.join(BASE, 'styles.css'))
html = html.replace('<link rel="stylesheet" href="styles.css" />', f'<style>\n{css}\n</style>')
html = html.replace('<script src="js/bundle.js"></script>',
                    f'<script>\ntry {{\n{bundle}\n}} catch (e) {{ console.error(e); document.body.innerHTML = "<pre style=\\"color:red;padding:20px\\">启动失败: " + e.message + "</pre>"; }}\n</script>')

out = os.path.join(BASE, '汤姆EXCEL万能模块-独立版.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print('built:', out, os.path.getsize(out), 'bytes')
