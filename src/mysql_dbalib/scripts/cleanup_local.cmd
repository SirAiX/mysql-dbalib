@echo off
setlocal EnableExtensions

set "CLEANUP_LOCAL_SELF=%~f0"

where powershell >nul 2>nul
if errorlevel 1 (
  echo ERROR: powershell was not found.
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$marker = '# POWERSHELL_PAYLOAD'; $content = Get-Content -Raw -LiteralPath $env:CLEANUP_LOCAL_SELF; $idx = $content.LastIndexOf($marker); if ($idx -lt 0) { throw 'PowerShell payload marker was not found.' }; $payload = $content.Substring($idx + $marker.Length); Invoke-Expression $payload"
set "RESULT=%ERRORLEVEL%"
exit /b %RESULT%

# POWERSHELL_PAYLOAD
$ErrorActionPreference = 'Stop'

$packageName = 'mysql-dbalib'

function Resolve-FullPath {
    param([string] $Path)
    return [IO.Path]::GetFullPath($Path)
}

function Find-SitePackagesDir {
    param([string] $StartDir)

    $current = Get-Item -LiteralPath (Resolve-FullPath $StartDir) -ErrorAction Stop
    while ($current) {
        if ($current.Name -ieq 'site-packages') {
            return $current.FullName
        }

        $current = $current.Parent
    }

    return $null
}

function Resolve-VenvPython {
    $candidates = New-Object 'System.Collections.Generic.List[string]'

    if (-not [string]::IsNullOrWhiteSpace($env:VIRTUAL_ENV)) {
        [void] $candidates.Add((Join-Path (Join-Path $env:VIRTUAL_ENV 'Scripts') 'python.exe'))
    }

    $scriptPath = (Resolve-Path -LiteralPath $env:CLEANUP_LOCAL_SELF).Path
    $scriptDir = Split-Path -Parent $scriptPath
    $sitePackagesDir = Find-SitePackagesDir $scriptDir

    if ($sitePackagesDir) {
        $libDir = Split-Path -Parent $sitePackagesDir
        $venvRoot = Split-Path -Parent $libDir
        if (-not [string]::IsNullOrWhiteSpace($venvRoot)) {
            [void] $candidates.Add((Join-Path (Join-Path $venvRoot 'Scripts') 'python.exe'))
        }
    }

    foreach ($candidate in $candidates) {
        if ((-not [string]::IsNullOrWhiteSpace($candidate)) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Remove-SafePackageDir {
    param([string] $Target)

    if ([string]::IsNullOrWhiteSpace($Target)) {
        return
    }

    $targetFull = Resolve-FullPath $Target
    $targetName = [IO.Path]::GetFileName($targetFull.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar))

    if ($targetName -ine $packageName) {
        throw "Refusing to remove unexpected path '$targetFull'."
    }

    if (Test-Path -LiteralPath $targetFull -PathType Container) {
        Remove-Item -LiteralPath $targetFull -Recurse -Force

        if (Test-Path -LiteralPath $targetFull -PathType Container) {
            throw "Failed to remove '$targetFull'."
        }

        Write-Host "Removed: `"$targetFull`""
    }
    else {
        Write-Host "Not found: `"$targetFull`""
    }
}

function Invoke-PipUninstall {
    param([string] $PythonPath)

    if ([string]::IsNullOrWhiteSpace($PythonPath)) {
        Write-Host 'Python uninstall skipped: virtual environment python was not detected.'
        return
    }

    if ($env:MYSQL_DBALIB_SKIP_PIP_UNINSTALL -eq '1') {
        Write-Host "Python uninstall skipped by MYSQL_DBALIB_SKIP_PIP_UNINSTALL: `"$PythonPath`""
        return
    }

    Write-Host "Uninstalling $packageName with `"$PythonPath`"..."
    & $PythonPath -m pip uninstall -y $packageName

    if ($LASTEXITCODE -ne 0) {
        throw "pip uninstall failed with exit code $LASTEXITCODE."
    }
}

if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Remove-SafePackageDir (Join-Path $env:LOCALAPPDATA $packageName)
}

if (-not [string]::IsNullOrWhiteSpace($env:APPDATA)) {
    Remove-SafePackageDir (Join-Path $env:APPDATA $packageName)
}

if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
    Remove-SafePackageDir (Join-Path (Join-Path $env:USERPROFILE 'Documents') $packageName)
}

Invoke-PipUninstall (Resolve-VenvPython)

Write-Host 'Cleanup completed.'
