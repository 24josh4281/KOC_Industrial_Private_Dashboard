@echo off
setlocal
if "%~1"=="" (echo Drag the reviewed workbook or CSV onto this file.& pause & exit /b 1)
set "PY=C:\Users\24jos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PY%" ("%PY%" "%~dp0import_workflow.py" "%~1" --commit) else (py -3 "%~dp0import_workflow.py" "%~1" --commit)
pause
