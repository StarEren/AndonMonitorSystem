@echo off
setlocal

set "script=AndonReport.py"
set "script_path=%~dp0\%script%"

set "excel_file=AndonReport.xlsx"
set "excel_path=%~dp0\%excel_file%"

echo "Closing Excel..."
taskkill /f /im excel.exe >nul

echo "Starting refresh..."
python "%script_path%"

echo "Opening Excel...
start "" "%excel_path%"

endlocal