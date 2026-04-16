[CmdletBinding()]
param(
    [ValidateSet("Share", "SingleExe")]
    [string]$Mode = "Share",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$AppName = "SakuraLoadMonitor"
$VenvPython = Join-Path $ProjectRoot ".venv311\Scripts\python.exe"
$BuildDir = Join-Path $ProjectRoot "build"
$DistDir = Join-Path $ProjectRoot "dist"
$OutputName = if ($Mode -eq "SingleExe") { "$AppName-SingleExe" } else { $AppName }
$ReleaseDir = Join-Path $DistDir $OutputName
$ResDir = Join-Path $ProjectRoot "res"
$LibDir = Join-Path $ProjectRoot "lib"

function Invoke-PythonStep {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $Python @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Python command failed: $($Arguments -join ' ')"
    }
}

function Resolve-BuildPython {
    if (Test-Path $VenvPython) {
        return $VenvPython
    }

    foreach ($version in @("3.11", "3.12", "3.13")) {
        try {
            & py "-$version" -m venv ".venv311"
            if (Test-Path $VenvPython) {
                return $VenvPython
            }
        }
        catch {
        }
    }

    throw "Unable to create .venv311. Install Python 3.11, 3.12, or 3.13 with the py launcher."
}

$Python = Resolve-BuildPython

if ($Clean) {
    $PathsToClean = @($BuildDir, $ReleaseDir)
    if ($Mode -eq "Share") {
        $PathsToClean += Join-Path $DistDir "$AppName-SingleExe.exe"
    }
    else {
        $PathsToClean += Join-Path $DistDir $AppName
    }

    foreach ($path in $PathsToClean) {
        if (Test-Path $path) {
            Remove-Item -LiteralPath $path -Recurse -Force
        }
    }
}

Invoke-PythonStep -Arguments @("-m", "pip", "install", "--upgrade", "pip")
Invoke-PythonStep -Arguments @("-m", "pip", "install", "-r", "requirements.txt")
Invoke-PythonStep -Arguments @("-m", "pip", "install", "pyinstaller>=6.0")

$PyInstallerArgs = @(
    "-m"
    "PyInstaller"
    "--noconfirm"
    "--clean"
    "--windowed"
    if ($Mode -eq "SingleExe") { "--onefile" } else { "--onedir" }
    "--name"
    $OutputName
    "--distpath"
    $DistDir
    "--workpath"
    $BuildDir
    "--specpath"
    $BuildDir
    "--hidden-import"
    "clr"
    "--collect-submodules"
    "pythonnet"
    "--add-data"
    "${ResDir};res"
    "--add-data"
    "${LibDir};lib"
    "main.py"
)

Invoke-PythonStep -Arguments $PyInstallerArgs

Write-Host ""
Write-Host "Build complete:" -ForegroundColor Green
Write-Host "  $ReleaseDir"
Write-Host ""
Write-Host "Main executable:" -ForegroundColor Green
if ($Mode -eq "SingleExe") {
    Write-Host "  $ReleaseDir.exe"
}
else {
    Write-Host "  $(Join-Path $ReleaseDir "$OutputName.exe")"
}
Write-Host ""
if ($Mode -eq "SingleExe") {
    Write-Host "This build is a single EXE. PyInstaller still extracts runtime files to a temp folder at launch."
}
else {
    Write-Host "This build is one-dir so it can be copied to and run from a network share."
}
