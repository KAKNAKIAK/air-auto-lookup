@echo off
setlocal
set "PYTHONUTF8=1"
set "APP_DIR=%~dp0"
where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" /D "%APP_DIR%" pythonw.exe "%APP_DIR%air_auto_lookup_mvp.py"
) else (
    start "" /D "%APP_DIR%" pyw.exe "%APP_DIR%air_auto_lookup_mvp.py"
)
endlocal
exit /b 0
