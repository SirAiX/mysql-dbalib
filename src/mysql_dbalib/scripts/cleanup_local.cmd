@echo off
setlocal EnableExtensions

set "PACKAGE_NAME=mysql-dbalib"

if defined VIRTUAL_ENV (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -m pip uninstall -y %PACKAGE_NAME%
  )
)

if defined LOCALAPPDATA call :safe_remove "%LOCALAPPDATA%\%PACKAGE_NAME%"
if errorlevel 1 exit /b 1

if defined APPDATA call :safe_remove "%APPDATA%\%PACKAGE_NAME%"
if errorlevel 1 exit /b 1

if defined USERPROFILE call :safe_remove "%USERPROFILE%\Documents\%PACKAGE_NAME%"
if errorlevel 1 exit /b 1

echo Cleanup completed.
exit /b 0

:safe_remove
set "TARGET=%~1"
if "%TARGET%"=="" exit /b 0

for %%I in ("%TARGET%") do set "TARGET_FULL=%%~fI"
for %%I in ("%TARGET_FULL%") do set "TARGET_NAME=%%~nxI"

if /I not "%TARGET_NAME%"=="%PACKAGE_NAME%" (
  echo ERROR: refusing to remove unexpected path "%TARGET_FULL%".
  exit /b 1
)

if exist "%TARGET_FULL%\" (
  rmdir /s /q "%TARGET_FULL%"
  if exist "%TARGET_FULL%\" (
    echo ERROR: failed to remove "%TARGET_FULL%".
    exit /b 1
  )
  echo Removed: "%TARGET_FULL%"
) else (
  echo Not found: "%TARGET_FULL%"
)

exit /b 0
