$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

python -m pip install -r requirements.txt

python -m PyInstaller `
  --noconfirm `
  --windowed `
  --name "ScoreTool" `
  --paths "$ProjectRoot\src" `
  "$ProjectRoot\src\score_tool\app.py"

Write-Host ""
Write-Host "Build completed: $ProjectRoot\dist\ScoreTool\ScoreTool.exe"