Set-Location "C:\Users\gabriel.abel\.openclaw\workspace\dash-compras"
git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
node .\scripts\gerar-dashboard-notas-sem-pedido.js
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git add index.html scripts\gerar-dashboard-notas-sem-pedido.js update-dashboard.ps1
$changes = git diff --cached --name-only
if (-not $changes) { Write-Output "Sem mudanças para publicar: a dashboard já estava atualizada."; exit 0 }
git commit -m "chore: atualiza dashboard automaticamente"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git push origin main
