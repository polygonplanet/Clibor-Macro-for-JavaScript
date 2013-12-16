@echo off

rem ---Set your python directory---
set PYTHON_DIR=I:\Python
rem -------------------------------

set PYTHONHOME=%PYTHON_DIR%

setlocal
set APP_DIR=%CD%

cd %PYTHON_DIR%
call python "%APP_DIR%\jsmacro.clip.py" "%APP_DIR%\stmp\jsmacro.tmpclip.txt" "%1"
endlocal

exit /b

