@echo off
setlocal
set "PY=C:\Users\24jos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "DB=%~dp0..\v13\state\working.sqlite3"
if exist "%PY%" (
  "%PY%" "%~dp0build_v14.py" --db "%DB%"
) else (
  py -3 "%~dp0build_v14.py" --db "%DB%"
)
if errorlevel 1 echo Build failed. Review the message above.
pause
