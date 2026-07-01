$ErrorActionPreference = "Stop"
$bundledPython = "C:\Users\Yucon\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if (-not $env:PUBLIC_BASE_URL) {
    $env:PUBLIC_BASE_URL = "https://center.gxajb.site"
}

if (Test-Path "$PSScriptRoot\.venv\Scripts\python.exe") {
    $python = "$PSScriptRoot\.venv\Scripts\python.exe"
} elseif (Test-Path $bundledPython) {
    $python = $bundledPython
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $python = "python"
} else {
    throw "Python 3 was not found. Install Python 3.11 or newer."
}

& $python "$PSScriptRoot\app.py"
