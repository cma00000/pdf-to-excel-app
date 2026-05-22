@echo off
setlocal
cd /d "%~dp0"

echo Installing dependencies...
pip install -q -r requirements.txt
if errorlevel 1 goto :error

echo Building PDF to Excel.exe with PyInstaller...
pyinstaller ^
  --windowed ^
  --onefile ^
  --name "PDF to Excel" ^
  --collect-all tkinterdnd2 ^
  --collect-all pdfplumber ^
  --collect-all pdfminer ^
  --hidden-import openpyxl ^
  --noconfirm ^
  app.py
if errorlevel 1 goto :error

echo.
echo Done!  Executable is at: dist\PDF to Excel.exe
explorer dist
goto :eof

:error
echo.
echo Build failed. See errors above.
exit /b 1
