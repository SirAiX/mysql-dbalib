@echo off
setlocal EnableExtensions

set "PYCHARM_TEMPLATE_SELF=%~f0"
set "PYCHARM_TEMPLATE_PROJECT_ROOT=%~dp0.."
for %%I in ("%PYCHARM_TEMPLATE_PROJECT_ROOT%") do set "PYCHARM_TEMPLATE_PROJECT_ROOT=%%~fI"
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

function Get-FallbackProjectRoot {
    $fallbackDirName = -join @([char]0x041A, ' ', [char]0x0443, [char]0x0441, [char]0x0442, [char]0x0430, [char]0x043D, [char]0x043E, [char]0x0432, [char]0x043A, [char]0x0435)
    return Join-Path (Join-Path 'E:\' $fallbackDirName) 'mysql-dbalib'
}

function Resolve-SourceFile {
    param([string] $RelativePath)

    $localPath = Join-Path $env:PYCHARM_TEMPLATE_PROJECT_ROOT $RelativePath
    if (Test-Path -LiteralPath $localPath -PathType Leaf) {
        return (Resolve-Path -LiteralPath $localPath).Path
    }

    $fallbackPath = Join-Path (Get-FallbackProjectRoot) $RelativePath
    if (Test-Path -LiteralPath $fallbackPath -PathType Leaf) {
        return (Resolve-Path -LiteralPath $fallbackPath).Path
    }

    throw "Source file was not found: '$localPath' or '$fallbackPath'."
}

function Convert-ToLiveTemplateValue {
    param([string] $Text)

    $normalized = $Text -replace "`r`n", "`n" -replace "`r", "`n"
    $literalDollars = $normalized -replace '\$', '$$$$'
    $escaped = [Security.SecurityElement]::Escape($literalDollars)
    return $escaped -replace "`n", '&#10;'
}

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
    $relativePath = "src\mysql_dbalib\resources\assets\$($definition.Version)\pack.py"
    $sourcePath = Resolve-SourceFile $relativePath
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
