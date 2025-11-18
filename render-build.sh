#!/usr/bin/env bash
set -o errexit

# 先正常安装依赖 + 下载Chrome
npm install
npx puppeteer browsers install chrome

# 关键：把下载的Chrome复制到Render运行时能访问的路径
PUPPETEER_CACHE_DIR="/opt/render/.cache/puppeteer"

if [[ ! -d $PUPPETEER_CACHE_DIR ]]; then
  echo "首次复制Chrome到缓存..."
  mkdir -p $PUPPETEER_CACHE_DIR
  cp -R ~/.cache/puppeteer/* $PUPPETEER_CACHE_DIR/
fi
