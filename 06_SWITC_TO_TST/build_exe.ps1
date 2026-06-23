[CmdletBinding()]
param(
    [string]$PythonExe = $env:PYTHON_EXE,
    [switch]$Clean,
    [switch]$NoVenv
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppScript = Join-Path $ProjectRoot "SWITC_TO_TST_v1_7.py"
$Requirements = Join-Path $ProjectRoot "requirements.txt"
$VenvDir = Join-Path $ProjectRoot ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$ExeName = "SWITC_TO_TST_v1_7"

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE`: $Command $($Arguments -join ' ')"
    }
}

function Test-PythonCandidate {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [string[]]$Prefix = @()
    )

    $cmd = Get-Command $Command -ErrorAction SilentlyContinue
    if (-not $cmd) {
        return $null
    }

    & $cmd.Source @Prefix "--version" *> $null
    if ($LASTEXITCODE -ne 0) {
        return $null
    }

    return [pscustomobject]@{
        Command = $cmd.Source
        Prefix = $Prefix
    }
}

function Resolve-Python {
    if ($PythonExe) {
        if (-not (Test-Path $PythonExe)) {
            throw "PYTHON_EXE was set, but the file does not exist: $PythonExe"
        }

        Invoke-Checked -Command $PythonExe -Arguments @("--version")
        return [pscustomobject]@{
            Command = (Resolve-Path $PythonExe).Path
            Prefix = @()
        }
    }

    $candidate = Test-PythonCandidate -Command "py" -Prefix @("-3")
    if ($candidate) {
        return $candidate
    }

    $candidate = Test-PythonCandidate -Command "python"
    if ($candidate) {
        return $candidate
    }

    throw "Python 3 was not found. Install Python 3 or set PYTHON_EXE=C:\Path\To\python.exe before running this script."
}

if (-not (Test-Path $AppScript)) {
    throw "Application script not found: $AppScript"
}

if (-not (Test-Path $Requirements)) {
    throw "Requirements file not found: $Requirements"
}

Push-Location $ProjectRoot
try {
    if ($Clean) {
        foreach ($path in @("build", "dist")) {
            if (Test-Path $path) {
                Remove-Item -Recurse -Force $path
            }
        }
    }

    if ($NoVenv) {
        $python = Resolve-Python
        $buildPython = $python.Command
        $buildPrefix = $python.Prefix
    }
    else {
        if (-not (Test-Path $VenvPython)) {
            $python = Resolve-Python
            Invoke-Checked -Command $python.Command -Arguments @($python.Prefix + @("-m", "venv", $VenvDir))
        }

        $buildPython = $VenvPython
        $buildPrefix = @()
    }

    Invoke-Checked -Command $buildPython -Arguments @($buildPrefix + @("-m", "pip", "install", "--upgrade", "pip"))
    Invoke-Checked -Command $buildPython -Arguments @($buildPrefix + @("-m", "pip", "install", "-r", $Requirements))
    Invoke-Checked -Command $buildPython -Arguments @(
        $buildPrefix + @(
            "-m", "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--windowed",
            "--name", $ExeName,
            "--distpath", "dist",
            "--workpath", "build",
            "--specpath", "build",
            $AppScript
        )
    )

    $outputExe = Join-Path $ProjectRoot "dist\$ExeName.exe"
    if (-not (Test-Path $outputExe)) {
        throw "Build finished, but the executable was not found: $outputExe"
    }

    Write-Host "Built executable: $outputExe"
}
finally {
    Pop-Location
}
