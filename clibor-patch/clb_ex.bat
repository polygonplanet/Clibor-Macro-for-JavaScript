@echo OFF

setlocal
set APP_DIR=%CD%
set MACRO_DIR=%APP_DIR%\stmp
set MACRO_FILE=%MACRO_DIR%\jsmacro.wsf

copy /y %MACRO_DIR%\*.py %MACRO_FILE%
call %APP_DIR%\jsmacro.trigger.wsf %MACRO_FILE%
endlocal

exit /b 0

