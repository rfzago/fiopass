#!/usr/bin/env bash
# Gera o FioPass.app no macOS com todos os arquivos de dados necessários
# (template .xlsx e VERSION). Ver README.md para detalhes.
set -euo pipefail
cd "$(dirname "$0")"

python3 -m PyInstaller -y --windowed \
  --name "FioPass" \
  --add-data "Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx:." \
  --add-data "VERSION:." \
  fiopass_gui.py

echo "Build concluído: dist/FioPass.app"
