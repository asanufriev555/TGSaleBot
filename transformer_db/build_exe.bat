@echo off
rem Сборка одного exe (запускать на машине с Python 3.8 — последней версией для Windows 7)
py -3.8 -m pip install "pyinstaller==4.10"
py -3.8 -m PyInstaller --noconsole --onefile --name TransformerCatalog main.py
echo Готово: dist\TransformerCatalog.exe
