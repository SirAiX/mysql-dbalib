@echo off
setlocal EnableExtensions

set "PYCHARM_TEMPLATE_SELF=%~f0"
set "PYCHARM_TEMPLATE_CONFIG_ARG=%~1"

where powershell >nul 2>nul
if errorlevel 1 (
  echo ERROR: powershell was not found.
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$marker = '# POWERSHELL_PAYLOAD'; $content = Get-Content -Raw -LiteralPath $env:PYCHARM_TEMPLATE_SELF; $idx = $content.LastIndexOf($marker); if ($idx -lt 0) { throw 'PowerShell payload marker was not found.' }; $payload = $content.Substring($idx + $marker.Length); Invoke-Expression $payload"
set "RESULT=%ERRORLEVEL%"
exit /b %RESULT%

# POWERSHELL_PAYLOAD
$ErrorActionPreference = 'Stop'

function Resolve-PyCharmConfig {
    param([string] $ConfigArg)

    if (-not [string]::IsNullOrWhiteSpace($ConfigArg)) {
        $expanded = [Environment]::ExpandEnvironmentVariables($ConfigArg)
        return [IO.Path]::GetFullPath($expanded)
    }

    if ([string]::IsNullOrWhiteSpace($env:APPDATA)) {
        throw 'APPDATA is not defined. Pass the PyCharm config path as the first argument.'
    }

    $jetBrainsRoot = Join-Path $env:APPDATA 'JetBrains'
    $candidate = Get-ChildItem -LiteralPath $jetBrainsRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like 'PyCharm*' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $candidate) {
        throw "PyCharm config was not found under '$jetBrainsRoot'. Pass it explicitly as the first argument."
    }

    return $candidate.FullName
}

$configPath = Resolve-PyCharmConfig $env:PYCHARM_TEMPLATE_CONFIG_ARG
$templatePath = Join-Path (Join-Path $configPath 'templates') 'mysql-dbalib.xml'

if (Test-Path -LiteralPath $templatePath -PathType Leaf) {
    Remove-Item -LiteralPath $templatePath -Force
    Write-Host "Removed: $templatePath"
} else {
    Write-Host "Not found: $templatePath"
}
