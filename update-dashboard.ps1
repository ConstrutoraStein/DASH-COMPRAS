Set-Location "C:\Users\gabriel.abel\.openclaw\workspace\dash-compras"
node .\scripts\gerar-dashboard-notas-sem-pedido.js
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git add index.html scripts\gerar-dashboard-notas-sem-pedido.js update-dashboard.ps1 README.md .gitignore
$changes = git diff --cached --name-only
if (-not $changes) { Write-Output "Sem mudanças para publicar."; exit 0 }
git commit -m "chore: atualiza dashboard automaticamente"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git pull --rebase origin main
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
git push origin main
