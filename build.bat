@echo off
cd /d "%~dp0"

echo Building executable...
.venv\Scripts\pyinstaller.exe --onefile --noconsole --add-data "templates;templates" --name "绩效考核工具" app.py

echo.
echo Done! Executable is in the "dist" folder.
pause
