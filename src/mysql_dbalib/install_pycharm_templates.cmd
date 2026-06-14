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

$packageDirName = 'mysql_dbalib'

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

function New-TemplateLayout {
    param(
        [string] $Kind,
        [string] $PackageDir
    )

    return [pscustomobject] @{
        Kind = $Kind
        PackageDir = Resolve-FullPath $PackageDir
    }
}

function Resolve-TemplateLayout {
    $scriptPath = (Resolve-Path -LiteralPath $env:PYCHARM_TEMPLATE_SELF).Path
    $scriptDir = Split-Path -Parent $scriptPath

    $sourceRoot = Find-SourceRoot $scriptDir
    if ($sourceRoot) {
        return New-TemplateLayout 'source' (Join-Path (Join-Path $sourceRoot 'src') $packageDirName)
    }

    if (Test-PackageDir $scriptDir) {
        return New-TemplateLayout 'installed-package' $scriptDir
    }

    $parentDir = Split-Path -Parent $scriptDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and (Test-PackageDir $parentDir)) {
        return New-TemplateLayout 'installed-package-scripts' $parentDir
    }

    $siblingPackageDir = Join-Path $scriptDir $packageDirName
    if (Test-PackageDir $siblingPackageDir) {
        return New-TemplateLayout 'installed-site-packages' $siblingPackageDir
    }

    throw "Could not find source project or installed '$packageDirName' package near '$scriptDir'."
}

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

function Resolve-SourceFile {
    param(
        $Layout,
        [string] $RelativePackagePath
    )

    $localPath = Join-Path $Layout.PackageDir $RelativePackagePath
    if (Test-Path -LiteralPath $localPath -PathType Leaf) {
        return (Resolve-Path -LiteralPath $localPath).Path
    }

    throw "Source file was not found: '$localPath'."
}

function Convert-ToLiveTemplateValue {
    param([string] $Text)

    $normalized = $Text -replace "`r`n", "`n" -replace "`r", "`n"
    $literalDollars = $normalized -replace '\$', '$$$$'
    $escaped = [Security.SecurityElement]::Escape($literalDollars)
    return $escaped -replace "`n", '&#10;'
}

$layout = Resolve-TemplateLayout
$configPath = Resolve-PyCharmConfig $env:PYCHARM_TEMPLATE_CONFIG_ARG
$templatesDir = Join-Path $configPath 'templates'
$templatePath = Join-Path $templatesDir 'mysql-dbalib.xml'

New-Item -ItemType Directory -Force -Path $templatesDir | Out-Null

$definitions = @(
    @{ Name = 'mdb_v2_pack'; Version = 'v2'; Description = 'mysql-dbalib v2 pack.py' },
    @{ Name = 'mdb_v3_pack'; Version = 'v3'; Description = 'mysql-dbalib v3 pack.py' },
    @{ Name = 'mdb_v5_pack'; Version = 'v5'; Description = 'mysql-dbalib v5 pack.py' }
)

$builder = New-Object Text.StringBuilder
[void] $builder.AppendLine('<templateSet group="mysql-dbalib">')

foreach ($definition in $definitions) {
    $relativePath = "resources\assets\$($definition.Version)\pack.py"
    $sourcePath = Resolve-SourceFile $layout $relativePath
    $sourceText = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::UTF8)
    $value = Convert-ToLiveTemplateValue $sourceText
    $description = [Security.SecurityElement]::Escape($definition.Description)

    [void] $builder.AppendLine("  <template name=""$($definition.Name)"" value=""$value"" description=""$description"" toReformat=""false"" toShortenFQNames=""false"">")
    [void] $builder.AppendLine('    <context>')
    [void] $builder.AppendLine('      <option name="Python" value="true" />')
    [void] $builder.AppendLine('    </context>')
    [void] $builder.AppendLine('  </template>')

    Write-Host "Added template $($definition.Name) from $sourcePath"
}

[void] $builder.AppendLine('</templateSet>')

$utf8NoBom = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText($templatePath, $builder.ToString(), $utf8NoBom)

Write-Host "PyCharm templates written: $templatePath"
Write-Host "Layout: $($layout.Kind)"
