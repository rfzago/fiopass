@echo off
REM Gera o FioPass.exe no Windows com todos os arquivos de dados necessarios
REM (template .xlsx e VERSION). Ver README.md para detalhes.
cd /d "%~dp0"

python -m PyInstaller -y --windowed ^
  --name "FioPass" ^
  --add-data "Anexo I_PLANILHA PASSAGENS E DIÁRIAS 2026.xlsx;." ^
  --add-data "VERSION;." ^
  fiopass_gui.py

echo Build concluido: dist\FioPass\FioPass.exe
