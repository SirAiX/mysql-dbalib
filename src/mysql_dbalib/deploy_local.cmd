@echo off
setlocal EnableExtensions

set "DEPLOY_LOCAL_SELF=%~f0"

where powershell >nul 2>nul
if errorlevel 1 (
  echo ERROR: powershell was not found.
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$marker = '# POWERSHELL_PAYLOAD'; $content = Get-Content -Raw -LiteralPath $env:DEPLOY_LOCAL_SELF; $idx = $content.LastIndexOf($marker); if ($idx -lt 0) { throw 'PowerShell payload marker was not found.' }; $payload = $content.Substring($idx + $marker.Length); Invoke-Expression $payload"
set "RESULT=%ERRORLEVEL%"
exit /b %RESULT%

# POWERSHELL_PAYLOAD
$ErrorActionPreference = 'Stop'

$packageName = 'mysql-dbalib'
$packageDirName = 'mysql_dbalib'
$defaultVersion = '0.1.1'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Resolve-FullPath {
    param([string] $Path)
    return [IO.Path]::GetFullPath($Path)
}

function Test-PackageDir {
    param([string] $Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    return (Test-Path -LiteralPath (Join-Path $Path '__init__.py') -PathType Leaf)
}

function Get-RelativePath {
    param(
        [string] $BasePath,
        [string] $Path
    )

    $baseFull = Resolve-FullPath $BasePath
    if (-not $baseFull.EndsWith([IO.Path]::DirectorySeparatorChar)) {
        $baseFull += [IO.Path]::DirectorySeparatorChar
    }

    $pathFull = Resolve-FullPath $Path
    $baseUri = New-Object Uri($baseFull)
    $pathUri = New-Object Uri($pathFull)
    $relativeUri = $baseUri.MakeRelativeUri($pathUri)
    return [Uri]::UnescapeDataString($relativeUri.ToString()).Replace('/', [IO.Path]::DirectorySeparatorChar)
}

function Test-IsSameOrChildPath {
    param(
        [string] $Path,
        [string] $BasePath
    )

    $fullPath = Resolve-FullPath $Path
    $fullBasePath = Resolve-FullPath $BasePath

    if ($fullPath.Equals($fullBasePath, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    if (-not $fullBasePath.EndsWith([IO.Path]::DirectorySeparatorChar)) {
        $fullBasePath += [IO.Path]::DirectorySeparatorChar
    }

    return $fullPath.StartsWith($fullBasePath, [StringComparison]::OrdinalIgnoreCase)
}

function Find-SourceRoot {
    param([string] $StartDir)

    $current = Resolve-FullPath $StartDir
    $startFull = Resolve-FullPath $StartDir
    for ($i = 0; $i -lt 8; $i++) {
        $projectFile = Join-Path $current 'pyproject.toml'
        $packageDir = Join-Path (Join-Path $current 'src') $packageDirName
        $rootScriptsDir = Join-Path $current 'scripts'

        if ((Test-Path -LiteralPath $projectFile -PathType Leaf) -and (Test-PackageDir $packageDir)) {
            $scriptIsInSourcePackage = Test-IsSameOrChildPath $startFull $packageDir
            $scriptIsInRootScripts = Test-IsSameOrChildPath $startFull $rootScriptsDir
            $scriptIsAtProjectRoot = $startFull.Equals($current, [StringComparison]::OrdinalIgnoreCase)

            if ($scriptIsInSourcePackage -or $scriptIsInRootScripts -or $scriptIsAtProjectRoot) {
                return $current
            }
        }

        $parent = Split-Path -Parent $current
        if ([string]::IsNullOrWhiteSpace($parent) -or ($parent -eq $current)) {
            break
        }

        $current = $parent
    }

    return $null
}

function New-Layout {
    param(
        [string] $Kind,
        [string] $WorkRoot,
        [string] $PackageDir,
        [string] $SourceRoot
    )

    return [pscustomobject] @{
        Kind = $Kind
        WorkRoot = Resolve-FullPath $WorkRoot
        PackageDir = Resolve-FullPath $PackageDir
        SourceRoot = if ([string]::IsNullOrWhiteSpace($SourceRoot)) { $null } else { Resolve-FullPath $SourceRoot }
    }
}

function Resolve-DeployLayout {
    $scriptPath = (Resolve-Path -LiteralPath $env:DEPLOY_LOCAL_SELF).Path
    $scriptDir = Split-Path -Parent $scriptPath

    $sourceRoot = Find-SourceRoot $scriptDir
    if ($sourceRoot) {
        return New-Layout 'source' $sourceRoot (Join-Path (Join-Path $sourceRoot 'src') $packageDirName) $sourceRoot
    }

    if (Test-PackageDir $scriptDir) {
        return New-Layout 'installed-package' (Split-Path -Parent $scriptDir) $scriptDir $null
    }

    $parentDir = Split-Path -Parent $scriptDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and (Test-PackageDir $parentDir)) {
        return New-Layout 'installed-package-scripts' (Split-Path -Parent $parentDir) $parentDir $null
    }

    $siblingPackageDir = Join-Path $scriptDir $packageDirName
    if (Test-PackageDir $siblingPackageDir) {
        return New-Layout 'installed-site-packages' $scriptDir $siblingPackageDir $null
    }

    throw "Could not find source project or installed '$packageDirName' package near '$scriptDir'."
}

function Get-PackageVersion {
    param($Layout)

    if ($Layout.SourceRoot) {
        $pyprojectPath = Join-Path $Layout.SourceRoot 'pyproject.toml'
        if (Test-Path -LiteralPath $pyprojectPath -PathType Leaf) {
            $match = [regex]::Match([IO.File]::ReadAllText($pyprojectPath), '(?m)^\s*version\s*=\s*"([^"]+)"')
            if ($match.Success) {
                return $match.Groups[1].Value
            }
        }
    }

    $initPath = Join-Path $Layout.PackageDir '__init__.py'
    if (Test-Path -LiteralPath $initPath -PathType Leaf) {
        $match = [regex]::Match([IO.File]::ReadAllText($initPath), '(?m)^\s*__version__\s*=\s*"([^"]+)"')
        if ($match.Success) {
            return $match.Groups[1].Value
        }
    }

    return $defaultVersion
}

function Add-ZipFile {
    param(
        [IO.Compression.ZipArchive] $Zip,
        [string] $SourcePath,
        [string] $EntryName
    )

    $normalizedEntryName = ($EntryName -replace '\\', '/').TrimStart('/')
    [IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
        $Zip,
        $SourcePath,
        $normalizedEntryName,
        [IO.Compression.CompressionLevel]::Optimal
    ) | Out-Null
}

function Add-ZipText {
    param(
        [IO.Compression.ZipArchive] $Zip,
        [string] $EntryName,
        [string] $Text
    )

    $normalizedEntryName = ($EntryName -replace '\\', '/').TrimStart('/')
    $entry = $Zip.CreateEntry($normalizedEntryName, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $utf8NoBom = New-Object Text.UTF8Encoding($false)
    $writer = New-Object IO.StreamWriter -ArgumentList $stream, $utf8NoBom

    try {
        $writer.Write($Text)
    }
    finally {
        $writer.Dispose()
    }
}

function Test-SkippedPackageFile {
    param(
        [string] $PackageDir,
        [IO.FileInfo] $File
    )

    $relativePath = Get-RelativePath $PackageDir $File.FullName
    $parts = $relativePath -split '[\\/]'

    if ($parts -contains '__pycache__') {
        return $true
    }

    if (($File.Extension -eq '.pyc') -or ($File.Extension -eq '.pyo')) {
        return $true
    }

    return $false
}

function New-Archive {
    param(
        $Layout,
        [string] $PackageVersion,
        [string] $ArchivePath
    )

    $archiveDir = Split-Path -Parent $ArchivePath
    New-Item -ItemType Directory -Force -Path $archiveDir | Out-Null

    if (Test-Path -LiteralPath $ArchivePath -PathType Leaf) {
        Remove-Item -LiteralPath $ArchivePath -Force
    }

    $zip = [IO.Compression.ZipFile]::Open($ArchivePath, [IO.Compression.ZipArchiveMode]::Create)

    try {
        if ($Layout.SourceRoot) {
            foreach ($metadataName in @('pyproject.toml', 'README.md', 'MANIFEST.in')) {
                $metadataPath = Join-Path $Layout.SourceRoot $metadataName
                if (Test-Path -LiteralPath $metadataPath -PathType Leaf) {
                    Add-ZipFile $zip $metadataPath $metadataName
                }
            }
        }
        else {
            $generatedPyproject = @"
[build-system]
requires = ["setuptools>=69", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "$packageName"
version = "$PackageVersion"
description = "Local study package with bundled course documents."
requires-python = ">=3.10"
dependencies = []

[tool.setuptools]
package-dir = { "" = "src" }
include-package-data = true

[tool.setuptools.packages.find]
where = ["src"]

[tool.setuptools.package-data]
mysql_dbalib = [
  "*",
  ".*",
  "**/*",
  "**/.*"
]
"@

            Add-ZipText $zip 'pyproject.toml' $generatedPyproject
            Add-ZipText $zip 'README.md' "# mysql-dbalib`r`n`r`nArchive generated from an installed mysql-dbalib package.`r`n"
            Add-ZipText $zip 'MANIFEST.in' "recursive-include src/mysql_dbalib *`r`nglobal-exclude __pycache__ *.py[cod]`r`n"
        }

        $packageFiles = Get-ChildItem -LiteralPath $Layout.PackageDir -Recurse -Force -File
        foreach ($file in $packageFiles) {
            if (Test-SkippedPackageFile $Layout.PackageDir $file) {
                continue
            }

            $relativePath = Get-RelativePath $Layout.PackageDir $file.FullName
            $entryName = Join-Path (Join-Path 'src' $packageDirName) $relativePath
            Add-ZipFile $zip $file.FullName $entryName
        }
    }
    finally {
        $zip.Dispose()
    }

    if (-not (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) {
        throw "Archive was not created at '$ArchivePath'."
    }
}

function Copy-Archive {
    param(
        [string] $ArchivePath,
        [string] $ArchiveName,
        [string] $DestinationDir,
        [System.Collections.Generic.List[string]] $CopiedPaths
    )

    if ([string]::IsNullOrWhiteSpace($DestinationDir)) {
        return
    }

    New-Item -ItemType Directory -Force -Path $DestinationDir | Out-Null
    $destinationPath = Join-Path $DestinationDir $ArchiveName
    Copy-Item -LiteralPath $ArchivePath -Destination $destinationPath -Force
    [void] $CopiedPaths.Add($destinationPath)
    Write-Host "Copied: `"$destinationPath`""
}

$layout = Resolve-DeployLayout
$packageVersion = Get-PackageVersion $layout
$archiveName = "$packageName-$packageVersion.zip"
$releaseDir = Join-Path $layout.WorkRoot 'release'
$archivePath = Join-Path $releaseDir $archiveName

New-Archive $layout $packageVersion $archivePath

$copiedPaths = New-Object 'System.Collections.Generic.List[string]'

if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Copy-Archive $archivePath $archiveName (Join-Path (Join-Path $env:LOCALAPPDATA $packageName) 'archive') $copiedPaths
}

if (-not [string]::IsNullOrWhiteSpace($env:APPDATA)) {
    Copy-Archive $archivePath $archiveName (Join-Path (Join-Path $env:APPDATA $packageName) 'archive') $copiedPaths
}

if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
    Copy-Archive $archivePath $archiveName (Join-Path (Join-Path (Join-Path $env:USERPROFILE 'Documents') $packageName) 'archive') $copiedPaths
}

if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Write-Host 'WARNING: LOCALAPPDATA is not defined; install manifest was not written.'
    Write-Host "Archive: `"$archivePath`""
    exit 0
}

$manifestDir = Join-Path $env:LOCALAPPDATA $packageName
$manifestPath = Join-Path $manifestDir 'install-manifest.txt'
New-Item -ItemType Directory -Force -Path $manifestDir | Out-Null

$utf8NoBom = New-Object Text.UTF8Encoding($false)
$manifestLines = @(
    "package=$packageName",
    "version=$packageVersion",
    "layout=$($layout.Kind)",
    "source=$($layout.PackageDir)",
    "archive=$archivePath",
    "created=$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))",
    "copies=$([string]::Join(';', $copiedPaths.ToArray()))"
)

[IO.File]::WriteAllText($manifestPath, (($manifestLines -join [Environment]::NewLine) + [Environment]::NewLine), $utf8NoBom)

Write-Host "Archive created: `"$archivePath`""
Write-Host "Manifest written: `"$manifestPath`""
Write-Host "Layout: $($layout.Kind)"
