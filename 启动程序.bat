@echo off
setlocal
set "APP_SCRIPT=%~dp0app.py"
set "PY_EXE=C:\Users\Administrator\AppData\Local\Programs\Python\Python314\python.exe"

if exist "%PY_EXE%" (
  goto run_app
)

where python
if not errorlevel 1 (
  set "PY_EXE=python"
  goto run_app
)

where py
if not errorlevel 1 (
  set "PY_EXE=py"
  goto run_app
)

echo Python was not found. Please install Python 3.11 or later.
pause
exit /b 1

:run_app
"%PY_EXE%" -c "import openpyxl"
if errorlevel 1 (
  echo Installing required Python libraries...
  "%PY_EXE%" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo Failed to install required Python libraries.
    pause
    exit /b 1
  )
)
"%PY_EXE%" "%APP_SCRIPT%"
if errorlevel 1 pause
endlocal
