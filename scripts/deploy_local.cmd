@echo off
setlocal EnableExtensions

set "PACKAGE_NAME=mysql-dbalib"
set "PACKAGE_VERSION=0.1.1"
set "ARCHIVE_NAME=%PACKAGE_NAME%-%PACKAGE_VERSION%.zip"
set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"

set "RELEASE_DIR=%PROJECT_ROOT%\release"
set "ARCHIVE_PATH=%RELEASE_DIR%\%ARCHIVE_NAME%"

if not exist "%PROJECT_ROOT%\pyproject.toml" (
  echo ERROR: pyproject.toml was not found in "%PROJECT_ROOT%".
  exit /b 1
)

if not exist "%PROJECT_ROOT%\src\mysql_dbalib\" (
  echo ERROR: package directory was not found in "%PROJECT_ROOT%\src\mysql_dbalib".
  exit /b 1
)

if not exist "%RELEASE_DIR%\" mkdir "%RELEASE_DIR%"
if errorlevel 1 exit /b 1

if exist "%ARCHIVE_PATH%" del /q "%ARCHIVE_PATH%"
if errorlevel 1 exit /b 1

pushd "%PROJECT_ROOT%" >nul

where python >nul 2>nul
if not errorlevel 1 (
  python -c "from pathlib import Path; import zipfile; root=Path.cwd(); out=Path(r'%ARCHIVE_PATH%'); sources=[Path('pyproject.toml'),Path('README.md'),Path('MANIFEST.in'),Path('src/mysql_dbalib')]; z=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED); [z.write(p, p.relative_to(root).as_posix()) for s in sources for p in ((root/s).rglob('*') if (root/s).is_dir() else [root/s]) if p.is_file()]; z.close()"
) else (
  where powershell >nul 2>nul
  if errorlevel 1 (
    popd >nul
    echo ERROR: neither python nor powershell was found.
    exit /b 1
  )
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'pyproject.toml','README.md','MANIFEST.in','src\mysql_dbalib' -DestinationPath '%ARCHIVE_PATH%' -Force"
)

set "PACK_RESULT=%ERRORLEVEL%"
popd >nul
if not "%PACK_RESULT%"=="0" (
  echo ERROR: archive build failed.
  exit /b %PACK_RESULT%
)

if not exist "%ARCHIVE_PATH%" (
  echo ERROR: archive was not created at "%ARCHIVE_PATH%".
  exit /b 1
)

set "COPIED_PATHS="

if defined LOCALAPPDATA call :copy_archive "%LOCALAPPDATA%\%PACKAGE_NAME%\archive"
if errorlevel 1 exit /b 1

if defined APPDATA call :copy_archive "%APPDATA%\%PACKAGE_NAME%\archive"
if errorlevel 1 exit /b 1

if defined USERPROFILE call :copy_archive "%USERPROFILE%\Documents\%PACKAGE_NAME%\archive"
if errorlevel 1 exit /b 1

if not defined LOCALAPPDATA (
  echo WARNING: LOCALAPPDATA is not defined; install manifest was not written.
  echo Archive: "%ARCHIVE_PATH%"
  exit /b 0
)

set "MANIFEST_DIR=%LOCALAPPDATA%\%PACKAGE_NAME%"
set "MANIFEST_PATH=%MANIFEST_DIR%\install-manifest.txt"
if not exist "%MANIFEST_DIR%\" mkdir "%MANIFEST_DIR%"
if errorlevel 1 exit /b 1

(
  echo package=%PACKAGE_NAME%
  echo version=%PACKAGE_VERSION%
  echo source=%PROJECT_ROOT%
  echo archive=%ARCHIVE_PATH%
  echo created=%DATE% %TIME%
  echo copies=%COPIED_PATHS%
) > "%MANIFEST_PATH%"

echo Archive created: "%ARCHIVE_PATH%"
echo Manifest written: "%MANIFEST_PATH%"
exit /b 0

:copy_archive
set "DEST_DIR=%~1"
if "%DEST_DIR%"=="" exit /b 0
if not exist "%DEST_DIR%\" mkdir "%DEST_DIR%"
if errorlevel 1 exit /b 1

copy /y "%ARCHIVE_PATH%" "%DEST_DIR%\%ARCHIVE_NAME%" >nul
if errorlevel 1 (
  echo ERROR: failed to copy archive to "%DEST_DIR%".
  exit /b 1
)

if defined COPIED_PATHS (
  set "COPIED_PATHS=%COPIED_PATHS%;%DEST_DIR%\%ARCHIVE_NAME%"
) else (
  set "COPIED_PATHS=%DEST_DIR%\%ARCHIVE_NAME%"
)
echo Copied: "%DEST_DIR%\%ARCHIVE_NAME%"
exit /b 0
