#!/bin/bash
set -e
cd "$(dirname "$0")"

BASE_PYTHON=/Library/Frameworks/Python.framework/Versions/3.13/bin/python3
TARGET_ARCH=universal2
VENV_DIR="$(pwd)/buildenv"

# ── 1. 创建干净的虚拟环境 ──────────────────────────────────────
echo "🧹 创建干净虚拟环境..."
rm -rf "$VENV_DIR"
"$BASE_PYTHON" -m venv "$VENV_DIR"

PYTHON="$VENV_DIR/bin/python"
PIP="$VENV_DIR/bin/pip"

# ── 2. 只装必要的包 ────────────────────────────────────────────
echo "📦 安装依赖（仅必要包）..."
"$PIP" install --quiet pandas pyqt6 icalendar openpyxl pyinstaller

# ── 3. 找 PyQt6 插件路径 ──────────────────────────────────────
PYQT6_PLUGINS=$("$PYTHON" -c "import PyQt6, os; print(os.path.join(os.path.dirname(PyQt6.__file__), 'Qt6', 'plugins'))")
echo "📂 PyQt6 插件路径: $PYQT6_PLUGINS"

# ── 4. 清理旧包 ───────────────────────────────────────────────
echo "🗑  清理旧 build/dist..."
rm -rf dist build "ICS Visits 管理.spec"

# ── 5. 打包 ───────────────────────────────────────────────────
echo "🔨 开始打包..."
"$PYTHON" -m PyInstaller \
    --noconfirm \
    --windowed \
    --target-arch "$TARGET_ARCH" \
    --name "ICS Visits 管理" \
    --add-data "$PYQT6_PLUGINS:PyQt6/Qt6/plugins" \
    --runtime-hook rthook_qt.py \
    --hidden-import pandas \
    --hidden-import icalendar \
    --hidden-import openpyxl \
    --hidden-import openpyxl.styles \
    --hidden-import openpyxl.styles.fills \
    --exclude-module PyQt5 \
    --exclude-module PySide2 \
    --exclude-module PySide6 \
    nurseapp_qt.py

# ── 6. 修复 cocoa 插件（确保是 Qt6 版本）─────────────────────
echo "🔧 修复 cocoa 平台插件..."
CORRECT_COCOA="$PYQT6_PLUGINS/platforms/libqcocoa.dylib"
BUNDLE_PLATFORMS="dist/ICS Visits 管理.app/Contents/Frameworks/PyQt6/Qt6/plugins/platforms"

if [ -d "$BUNDLE_PLATFORMS" ]; then
    cp "$CORRECT_COCOA" "$BUNDLE_PLATFORMS/libqcocoa.dylib"
    echo "✅ cocoa 插件已替换（Qt6）"
else
    echo "⚠️  未找到 bundle platforms 目录，跳过替换"
fi

# ── 7. 打 zip ─────────────────────────────────────────────────
echo "📦 打包成 zip..."
cd dist
zip -qr "ICS Visits 管理.zip" "ICS Visits 管理.app"
cd ..

APP_SIZE=$(du -sh "dist/ICS Visits 管理.app" | cut -f1)
ZIP_SIZE=$(du -sh "dist/ICS Visits 管理.zip" | cut -f1)

echo ""
echo "✅ 完成！"
echo "   .app 大小: $APP_SIZE"
echo "   .zip 大小: $ZIP_SIZE"
echo "👉 发送 dist/ICS Visits 管理.zip 给对方即可"
