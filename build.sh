#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "▸ Installing dependencies..."
pip3 install -q -r requirements.txt

echo "▸ Building .app with PyInstaller..."
pyinstaller \
  --windowed \
  --name "PDF to Excel" \
  --osx-bundle-identifier com.pdftable.excel \
  --collect-all tkinterdnd2 \
  --collect-all pdfplumber \
  --collect-all pdfminer \
  --hidden-import openpyxl \
  --noconfirm \
  app.py

echo ""
echo "✓ Done!  App is at: dist/PDF to Excel.app"
echo "  Drag it to /Applications or double-click to run."
open dist/
