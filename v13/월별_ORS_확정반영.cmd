@echo off
setlocal
if "%~1"=="" (echo Drag a reviewed ORS ZIP onto this file.& pause & exit /b 1)
set "PY=C:\Users\24jos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PY%" ("%PY%" "%~dp0monthly_ors.py" "%~1" --commit) else (py -3 "%~dp0monthly_ors.py" "%~1" --commit)
pause
