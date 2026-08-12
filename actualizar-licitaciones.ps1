# ============================================================
# CINSA - Actualizador de Licitaciones (Monitor de Proyectos)
# Lee una copia local cacheada (CSV) de la pestaña "Listado licitaciones"
# del Google Sheet "Dashboard Proyectos" (descargada por separado antes de
# correr este script), recalcula KPIs, evolución por año, desempeño por
# rubro y ranking de competidores (incluyendo el puesto de CINSA), y
# actualiza la página web.
# ============================================================

$CsvPath    = "C:\Users\tira1\Listado Licitaciones (cache).csv"
$PyScript   = "C:\Users\tira1\actualizar-licitaciones.py"
$RepoPath   = "C:\Users\tira1"

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   CINSA - Actualizador de Licitaciones   ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 1. Verificar que el CSV cacheado existe
if (-not (Test-Path $CsvPath)) {
    Write-Host "❌ No se encontró la copia local cacheada del CSV." -ForegroundColor Red
    Write-Host "   Descarga la pestaña 'Listado licitaciones' de Dashboard Proyectos como CSV" -ForegroundColor Yellow
    Write-Host "   a esa ruta antes de correr este script." -ForegroundColor Yellow
    Read-Host "Presiona Enter para salir"
    exit
}
Write-Host "✅ CSV encontrado: $CsvPath" -ForegroundColor Green

# ── 2. Leer, recalcular y actualizar index.html (delegado a Python/pandas)
Write-Host "📖 Leyendo y recalculando datos de licitaciones..." -ForegroundColor Yellow

$output = & python $PyScript 2>&1
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    Write-Host "❌ Error al procesar el Excel:" -ForegroundColor Red
    Write-Host $output -ForegroundColor Yellow
    Read-Host "Presiona Enter para salir"
    exit
}
Write-Host "✅ $output" -ForegroundColor Green
Write-Host "📝 index.html actualizado" -ForegroundColor Green

# ── 3. Git: add, commit, push
Write-Host "🚀 Subiendo cambios a GitHub..." -ForegroundColor Yellow

Set-Location $RepoPath
$today = Get-Date -Format "yyyy-MM-dd HH:mm"

$gitAdd    = & git add index.html 2>&1
$gitCommit = & git commit -m "Actualizar datos licitaciones - $today" 2>&1
$gitPush   = & git push 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Página web actualizada exitosamente" -ForegroundColor Green
} else {
    Write-Host $gitPush -ForegroundColor Yellow
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  ✅ Listo! La web se actualizó en ~2 min ║" -ForegroundColor Green
Write-Host "║  dashboardcinsa.netlify.app              ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Read-Host "Presiona Enter para cerrar"
