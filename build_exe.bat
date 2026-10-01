@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run run_app.bat once before building the EXE.
  pause
  exit /b 1
)
call .venv\Scripts\activate.bat
pyinstaller --noconfirm --clean --windowed --name "HiyazaReportGenerator" ^
  --add-data "assets\DST_Agricultural_Association_Report_Template.docx;assets" ^
  main.py
echo.
echo EXE created at: dist\HiyazaReportGenerator\HiyazaReportGenerator.exe
pause

