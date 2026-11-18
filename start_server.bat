@echo off
setlocal
cd /d "%~dp0"

if not exist "dist\index.html" (
  echo A interface ainda nao foi compilada.
  echo Em um computador com Node.js, execute: npm install ^&^& npm run build
  echo Depois copie a pasta dist para este servidor.
  pause
  exit /b 1
)

set PYTHON_EXE=\\britfs2\gdv\SistemaDaQualidade\RPAs\RPA\py3_12_venv\Scripts\python.exe

if not exist "%PYTHON_EXE%" (
  echo Python nao encontrado em:
  echo %PYTHON_EXE%
  pause
  exit /b 1
)

"%PYTHON_EXE%" server.py

pause