Set-Location "C:\Users\gabriel.abel\.openclaw\workspace\dash-compras"
git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
node "C:\Users\gabriel.abel\.openclaw\workspace\jorgin\scripts\atualiza-stamp-data.js"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git add index.html update-dashboard.ps1
$changes = git diff --cached --name-only
if (-not $changes) { Write-Output "Sem mudanças para publicar: a dashboard foi mantida na versão atual do GitHub."; exit 0 }
git commit -m "chore: atualiza carimbo de data da dashboard"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git push origin main
