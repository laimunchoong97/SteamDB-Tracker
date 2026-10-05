param(
    [switch]$SkipInstaller
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$distRoot = Join-Path $root 'dist'
$portableRoot = Join-Path $distRoot 'SteamDB-Tracker'
$runtimeRoot = Join-Path $portableRoot 'runtime'
$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$node = (Get-Command node -ErrorAction Stop).Source

if (-not (Test-Path -LiteralPath $csc)) {
    throw "C# compiler not found: $csc"
}

if (Test-Path -LiteralPath $distRoot) {
    Remove-Item -LiteralPath $distRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $portableRoot | Out-Null
New-Item -ItemType Directory -Path $runtimeRoot | Out-Null

& $csc /nologo /target:winexe /platform:anycpu /optimize+ `
    /reference:System.dll /reference:System.Core.dll `
    /reference:System.Drawing.dll /reference:System.Windows.Forms.dll `
    /out:"$(Join-Path $portableRoot 'SteamDB Tracker.exe')" `
    "$(Join-Path $root 'launcher\SteamDbTrackerLauncher.cs')"
if ($LASTEXITCODE -ne 0) {
    throw 'Launcher compilation failed.'
}

Copy-Item -LiteralPath $node -Destination (Join-Path $runtimeRoot 'node.exe')
Copy-Item -LiteralPath (Join-Path $root 'node_modules') -Destination $portableRoot -Recurse

$files = @(
    'steamdb_tracker.js',
    'package.json',
    'package-lock.json',
    '.env.example',
    'README.md',
    'README.zh-CN.md'
)
foreach ($file in $files) {
    Copy-Item -LiteralPath (Join-Path $root $file) -Destination $portableRoot
}
Copy-Item -LiteralPath (Join-Path $root '.env.example') -Destination (Join-Path $portableRoot '.env')

$version = (& 'C:\Program Files\Git\cmd\git.exe' -C $root rev-parse --short HEAD).Trim()
Set-Content -LiteralPath (Join-Path $portableRoot 'version.txt') -Value $version -Encoding ASCII

$launcherTest = Start-Process -FilePath (Join-Path $portableRoot 'SteamDB Tracker.exe') `
    -ArgumentList '--self-test' -Wait -PassThru
if ($launcherTest.ExitCode -ne 0) {
    throw 'Packaged launcher self-test failed.'
}

Push-Location $portableRoot
try {
    & (Join-Path $runtimeRoot 'node.exe') -e "require('./steamdb_tracker.js'); console.log('Packaged Node runtime OK')"
    if ($LASTEXITCODE -ne 0) {
        throw 'Packaged Node runtime validation failed.'
    }
} finally {
    Pop-Location
}

$zipFile = Join-Path $distRoot 'SteamDB-Tracker-Windows-x64.zip'
Compress-Archive -Path (Join-Path $portableRoot '*') -DestinationPath $zipFile -CompressionLevel Optimal

$isccCandidates = @(
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe',
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe')
)
$iscc = $isccCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

if (-not $SkipInstaller -and $iscc) {
    & $iscc (Join-Path $root 'installer\SteamDBTracker.iss')
    if ($LASTEXITCODE -ne 0) {
        throw 'Installer compilation failed.'
    }
} elseif (-not $SkipInstaller) {
    Write-Warning 'Inno Setup was not found. Portable ZIP was created; Setup.exe was skipped.'
}

Write-Host "Portable folder: $portableRoot"
Write-Host "Portable ZIP:    $zipFile"
if (Test-Path -LiteralPath (Join-Path $distRoot 'SteamDB-Tracker-Setup.exe')) {
    Write-Host "Installer:       $(Join-Path $distRoot 'SteamDB-Tracker-Setup.exe')"
}

$checksumFiles = @($zipFile, (Join-Path $distRoot 'SteamDB-Tracker-Setup.exe')) `
    | Where-Object { Test-Path -LiteralPath $_ }
$checksumLines = foreach ($file in $checksumFiles) {
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $file).Hash.ToLowerInvariant()
    "$hash *$(Split-Path -Leaf $file)"
}
$checksumFile = Join-Path $distRoot 'SHA256SUMS.txt'
Set-Content -LiteralPath $checksumFile -Value $checksumLines -Encoding ASCII
Write-Host "Checksums:       $checksumFile"
