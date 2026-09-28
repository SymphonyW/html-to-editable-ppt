$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")
npm install
npx playwright install chromium
Write-Host "`nSetup complete. Example:"
Write-Host "  npm run convert -- .\examples\demo.html .\demo.pptx --report .\demo-report.json"
