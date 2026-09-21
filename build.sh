#!/usr/bin/env bash
set -e

# Package script for Thunderbird Add-on (.xpi)
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DIST_DIR="${PROJECT_DIR}/dist"
XPI_NAME="outlook-rwz-to-thunderbird.xpi"

echo "📦 開始封裝 Thunderbird 附加元件..."

mkdir -p "${DIST_DIR}"
rm -f "${DIST_DIR}/${XPI_NAME}"

cd "${PROJECT_DIR}"

# Ensure tests pass before packaging
echo "🧪 執行單元測試..."
node test/run-all.js

echo "🗜️ 建立 .xpi 封裝檔..."
zip -r -q "${DIST_DIR}/${XPI_NAME}" \
  manifest.json \
  background.js \
  lib/ \
  api/ \
  ui/ \
  icons/ \
  README.md \
  LICENSE

echo "========================================================"
echo "🎉 附加元件封裝成功！"
echo "檔案路徑: ${DIST_DIR}/${XPI_NAME}"
echo "大小: $(du -h "${DIST_DIR}/${XPI_NAME}" | cut -f1)"
echo "========================================================"
echo ""
echo "👉 安裝方法："
echo "1. 開啟 Mozilla Thunderbird"
echo "2. 前往「工具」>「附加元件與佈景主題」(或按下 Ctrl+Shift+A / Cmd+Shift+A)"
echo "3. 點選齒輪圖示 ⚙️，選擇「從檔案安裝附加元件...」"
echo "4. 選擇此 .xpi 檔案即可完成安裝！"
