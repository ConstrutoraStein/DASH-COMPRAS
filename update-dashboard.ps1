Set-Location "C:\Users\gabriel.abel\.openclaw\workspace\dash-compras"
git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Output "Dashboard atual mantida a partir da versão do GitHub."
