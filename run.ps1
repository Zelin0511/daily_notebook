$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir ".venv"
$PythonExe = Join-Path $VenvDir "Scripts\python.exe"
$PythonwExe = Join-Path $VenvDir "Scripts\pythonw.exe"
$ReadyMarker = Join-Path $VenvDir ".deps-ready"

if (-not (Test-Path $PythonExe)) {
    python -m venv $VenvDir
}

if (-not (Test-Path $ReadyMarker)) {
    & $PythonExe -m pip install --upgrade pip --quiet
    & $PythonExe -m pip install -r (Join-Path $ProjectDir "requirements.txt") --quiet
    New-Item -ItemType File -Force -Path $ReadyMarker | Out-Null
}

$Runner = if (Test-Path $PythonwExe) { $PythonwExe } else { $PythonExe }
& $Runner (Join-Path $ProjectDir "app.py")
