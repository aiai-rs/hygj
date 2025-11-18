#!/usr/bin/env bash
set -e

npm install

# 强制下载 Chrome 到 Render 能访问的路径
npx puppeteer browsers install chrome --with-deps
