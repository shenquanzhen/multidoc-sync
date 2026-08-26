#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_ROOT"

python3 scripts/generate_icons.py
ICONSET="$PROJECT_ROOT/build/MultiDocSync.iconset"
case "$ICONSET" in
  "$PROJECT_ROOT"/build/*) rm -rf -- "$ICONSET" ;;
  *) echo "Refusing to clean unexpected path: $ICONSET" >&2; exit 1 ;;
esac
mkdir -p "$ICONSET"
for size in 16 32 128 256 512; do
  sips -z "$size" "$size" assets/MultiDocSync.png --out "$ICONSET/icon_${size}x${size}.png" >/dev/null
  double=$((size * 2))
  sips -z "$double" "$double" assets/MultiDocSync.png --out "$ICONSET/icon_${size}x${size}@2x.png" >/dev/null
done
iconutil -c icns "$ICONSET" -o assets/MultiDocSync.icns

python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --onefile \
  --windowed \
  --name MultiDocSync \
  --osx-bundle-identifier io.github.shenquanzhen.multidoc-sync \
  --icon assets/MultiDocSync.icns \
  --add-data "assets/MultiDocSync.png:assets" \
  --exclude-module numpy \
  --exclude-module PIL \
  --exclude-module pandas \
  --exclude-module lxml \
  --exclude-module pytz \
  --exclude-module dateutil \
  --exclude-module pymupdf4llm \
  --exclude-module bs4 \
  --exclude-module tkinter \
  --exclude-module matplotlib \
  --exclude-module scipy \
  --exclude-module PySide6.QtQml \
  --exclude-module PySide6.QtQuick \
  --exclude-module PySide6.QtNetwork \
  --exclude-module PySide6.QtPdf \
  --exclude-module PySide6.QtPdfWidgets \
  --exclude-module PySide6.QtOpenGL \
  --exclude-module PySide6.QtOpenGLWidgets \
  --exclude-module PySide6.QtWebEngineCore \
  --exclude-module PySide6.QtWebEngineWidgets \
  multi_document_viewer.py

ditto -c -k --sequesterRsrc --keepParent dist/MultiDocSync.app dist/MultiDocSync-macOS-arm64.zip
echo "Built: $PROJECT_ROOT/dist/MultiDocSync-macOS-arm64.zip"
