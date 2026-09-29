@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".\.venv\Scripts\python.exe" (
  echo Нужен .venv. Смотрите раздел «Установка и запуск» в README.md.
  exit /b 1
)
if not exist ".\.venv\Scripts\pyinstaller.exe" (
  ".\.venv\Scripts\python.exe" -m pip install pyinstaller
)
".\.venv\Scripts\pyinstaller.exe" --noconfirm --onefile --noconsole --name pdfcutter --hidden-import windnd pdfcutter.py
if errorlevel 1 exit /b 1
echo Готово: dist\pdfcutter.exe
