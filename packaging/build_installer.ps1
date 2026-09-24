$ErrorActionPreference = "Stop"

$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "No se encontró .venv. Crea el entorno e instala requirements.txt y requirements-build.txt."
}

Push-Location $root
try {
    & $python -m PyInstaller --noconfirm --clean "packaging\RevisaBitacoras.spec"
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller terminó con error."
    }

    $isccCandidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )
    $iscc = $isccCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $iscc) {
        $command = Get-Command "iscc.exe" -ErrorAction SilentlyContinue
        if ($command) {
            $iscc = $command.Source
        }
    }
    if (-not $iscc) {
        throw "No se encontró Inno Setup 6 (ISCC.exe)."
    }

    & $iscc "packaging\RevisaBitacoras.iss"
    if ($LASTEXITCODE -ne 0) {
        throw "Inno Setup terminó con error."
    }

    Write-Host "Instalador creado en release\RevisorBitacoras-Setup-1.0.0.exe"
}
finally {
    Pop-Location
}
