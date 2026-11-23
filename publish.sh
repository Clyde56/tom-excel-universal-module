#!/usr/bin/env bash
# 一键发布到 GitHub（不含图标资源，见 README）
# 用法：先去 https://github.com/new 创建名为 tom-excel-universal-module 的公开空仓库，
#       然后在本目录执行： ./publish.sh [你的GitHub用户名]
set -e
OWNER="${1:-Clyde56}"
REPO="${2:-tom-excel-universal-module}"
ROOT="$(cd "$(dirname "$0")" && pwd)"

echo "==> 初始化本地 Git 仓库"
cd "$ROOT"
git init -q
git config user.name "Tom Excel Module contributors"
git config user.email "noreply@example.invalid"
git add README.md LICENSE NOTICE .gitignore rebrand.py publish.sh
git add electron/package.json
git add electron/dist/main/main.cjs electron/dist/main/preload.cjs
git add electron/dist/renderer/index.html
git add electron/dist/renderer/assets/index-B5DZ-YNN.js
git add electron/dist/renderer/assets/index-D239V13M.css
git add simple-version/index.html simple-version/styles.css simple-version/build-standalone.py
git add simple-version/js/*.js
git add simple-version/js/*.json
git commit -m "Initial commit: 汤姆EXCEL万能模块 (fork of jr-ribbon-xml-editor)" -q

echo "==> 关联远程仓库"
git remote remove origin 2>/dev/null || true
git remote add origin "https://github.com/${OWNER}/${REPO}.git"
git branch -M main

echo "==> 推送"
git push -u origin main

echo "完成： https://github.com/${OWNER}/${REPO}"
