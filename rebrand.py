# -*- coding: utf-8 -*-
r"""
把解包出的 Electron 应用从「杰瑞Ribbon设计器」彻底改名为「汤姆EXCEL万能模块」。
处理 Unicode 转义（bundle 里的中文是 \\uXXXX 字面量）与明文两套。
"""
import os, shutil

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'electron')
MAIN = os.path.join(BASE, 'dist', 'main', 'main.cjs')
PRELOAD = os.path.join(BASE, 'dist', 'main', 'preload.cjs')
RENDER_JS = os.path.join(BASE, 'dist', 'renderer', 'assets', 'index-B5DZ-YNN.js')
RENDER_HTML = os.path.join(BASE, 'dist', 'renderer', 'index.html')
PKG = os.path.join(BASE, 'package.json')

# 先备份原文件
for p in (MAIN, PRELOAD, RENDER_JS, RENDER_HTML, PKG):
    if not os.path.exists(p + '.orig'):
        shutil.copy(p, p + '.orig')

# 通用替换：先长后短，避免前缀冲突
COMMON = [
    # ---- Unicode 转义的中文（main.cjs 里是 \\uXXXX 字面量）----
    (r'\u6770\u745ERibbon\u8BBE\u8BA1\u5668', r'\u6C64\u59C6EXCEL\u4E07\u80FD\u6A21\u5757'),  # 杰瑞Ribbon设计器
    (r'\u6770\u745E', r'\u6C64\u59C6'),  # 杰瑞 -> 汤姆
    # ---- 明文品牌（renderer / html）----
    ('杰瑞Ribbon设计器', '汤姆EXCEL万能模块'),
    ('杰瑞办公开发', '汤姆办公开发'),
    ('杰瑞', '汤姆'),
    # ---- ASCII 标识符：长的在前 ----
    ('JRRibbonCallbacks', 'TOMExcelCallbacks'),
    ('JR_RIBBON_DEVTOOLS', 'TOM_EXCEL_DEVTOOLS'),
    ('JR-VBA-FAIL', 'TOM-VBA-FAIL'),
    ('JR_VBA_FAIL', 'TOM_VBA_FAIL'),
    ('JR-Ribbon-Installer', 'TOM-Excel-Installer'),
    ('JR-Ribbon-Preview', 'TOM-Excel-Preview'),
    ('JR-Ribbon-xml', 'TOM-Excel-Module'),
    ('JR Ribbon Designer', 'TOM EXCEL Module'),
    ('JR Ribbon Installer', 'TOM EXCEL Installer'),
    ('JR Ribbon', 'TOM EXCEL'),
    ('jrbghome.cn', 'tom-excel.example'),  # 个人站点占位（开源仓库不应保留私人站点）
    ('jrRibbon', 'tomExcel'),  # window 桥接对象
]


def apply(path):
    with open(path, encoding='utf-8') as f:
        s = f.read()
    n = 0
    for a, b in COMMON:
        c = s.count(a)
        if c:
            s = s.replace(a, b)
            n += c
    with open(path, 'w', encoding='utf-8') as f:
        f.write(s)
    print(f'{os.path.basename(path)}: {n} replacements')


for p in (MAIN, PRELOAD, RENDER_JS, RENDER_HTML):
    apply(p)

# package.json 用 JSON 结构改，避免误伤
import json
with open(PKG, encoding='utf-8') as f:
    pkg = json.load(f)
pkg['name'] = 'tom-excel-universal-module'
pkg['displayName'] = '汤姆EXCEL万能模块'
pkg['description'] = 'A desktop Ribbon XML designer for Excel .xlsm and .xlam macro files (open-source fork).'
pkg['author'] = 'Tom Excel Module contributors (fork of jerry-office/jr-ribbon-xml-editor)'
pkg['productName'] = '汤姆EXCEL万能模块'
with open(PKG, 'w', encoding='utf-8') as f:
    json.dump(pkg, f, ensure_ascii=False, indent=2)
print('package.json updated')
