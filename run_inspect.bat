@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Running inspector...
py "%~dp0inspect_docx.py" 2>&1
echo.
echo If you see "Done" above, check template_structure.txt
pause
