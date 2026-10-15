@echo off
setlocal
set "PY=C:\Users\24jos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PY%" ("%PY%" "%~dp0refresh_outputs.py") else (py -3 "%~dp0refresh_outputs.py")
pause
