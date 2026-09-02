# ============================================================
# CINSA - Actualizador automático de Monitor de Proyectos
# Lee una copia local cacheada (CSV) del Google Sheet "Dashboard
# Proyectos" (descargada por separado vía el conector de Google
# Drive antes de correr este script) y actualiza la página web.
# ============================================================

$CsvPath  = "C:\Users\tira1\Dashboard Proyectos (cache).csv"
$HtmlPath = "C:\Users\tira1\index.html"

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   CINSA - Actualizador de Proyectos      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 1. Verificar que el CSV cacheado existe
if (-not (Test-Path $CsvPath)) {
    Write-Host "❌ No se encontró la copia local cacheada del CSV." -ForegroundColor Red
    Write-Host "   Debe descargarse primero desde Google Drive antes de correr este script." -ForegroundColor Yellow
    Read-Host "Presiona Enter para salir"
    exit
}

Write-Host "✅ CSV encontrado: $CsvPath" -ForegroundColor Green

# ── 2. Leer y parsear el CSV
# Columnas (fila de encabezado real, fila 2 del CSV):
# 1 Fecha Adjudicada, 2 ID(ncontrato), 3 CC, 4 Proyecto, 5 Tipo, 6 Unidades,
# 7 Monto Contrato, 8 Plazo, 9 Inicio, 10 Termino, 11 Avance Programado,
# 12 % Avance real, 13-16 Facturado Real (por año), 17 Total Facturado, 18 % Facturado
Write-Host "📖 Leyendo datos del CSV..." -ForegroundColor Yellow

try {
    $headers = 1..18 | ForEach-Object { "c$_" }
    $dataLines = Get-Content -Path $CsvPath -Encoding UTF8 | Select-Object -Skip 2
    $csvRows = $dataLines | ConvertFrom-Csv -Header $headers

    $proyectos = @()
    $id = 1

    foreach ($r in $csvRows) {
        $nombre = ($r.c4 -as [string])
        if ($nombre) { $nombre = $nombre.Trim() }
        if ([string]::IsNullOrEmpty($nombre) -or $nombre -eq "Total Neto" -or $nombre -eq "Total Bruto") { break }

        $anioRaw     = ($r.c1  -as [string]).Trim()
        $ncontrato   = ($r.c2  -as [string]).Trim()
        $cc          = ($r.c3  -as [string]).Trim()
        $tipoRaw     = ($r.c5  -as [string]).Trim()
        $contratoRaw = (($r.c7  -as [string]).Trim()) -replace '[\.\$ ,]',''
        $plazoRaw    = ($r.c8  -as [string]).Trim()
        $inicioRaw   = ($r.c9  -as [string]).Trim()
        $terminoRaw  = ($r.c10 -as [string]).Trim()
        $avanceRaw   = ((($r.c12 -as [string]).Trim()) -replace '[%\s]','')
        $facturadoRaw= (($r.c17 -as [string]).Trim()) -replace '[\.\$ ,]',''

        # Mapear tipo
        $tipo = switch -Wildcard ($tipoRaw) {
            "*Luminaria*"   { "luminaria" }
            "*Televigilancia*" { "camara" }
            "*Postes*"      { "poste" }
            default         { "luminaria" }
        }

        # Parsear valores numéricos
        $contrato  = if ($contratoRaw  -match '^\d+$') { [long]$contratoRaw }  else { 0 }
        $facturado = if ($facturadoRaw -match '^\d+$') { [long]$facturadoRaw } else { 0 }
        $avance    = if ($avanceRaw    -match '^\d+$') { [int]$avanceRaw }     else { 0 }
        $plazo     = if ($plazoRaw     -match '^\d+$') { $plazoRaw }           else { "null" }
        $anio      = if ($anioRaw -match '^\d{4}') { [int]$anioRaw } else { 2025 }

        # Estado según avance
        $estado = if ($avance -ge 100) { "completado" } else { "activo" }

        # Ubicación = nombre del proyecto (se puede ajustar)
        $ubicacion = $nombre

        # Fechas: el CSV las exporta como M/D/YYYY; la web espera DD-MM-YYYY
        function ConvertFecha($raw) {
            if ($raw -match '^(\d{1,2})/(\d{1,2})/(\d{4})$') {
                $mm = [int]$matches[1]; $dd = [int]$matches[2]; $yyyy = $matches[3]
                return "{0:D2}-{1:D2}-{2}" -f $dd, $mm, $yyyy
            }
            return $raw
        }
        $inicioFmt  = if ($inicioRaw)  { ConvertFecha $inicioRaw }  else { "" }
        $terminoFmt = if ($terminoRaw) { ConvertFecha $terminoRaw } else { "" }

        $proyectos += [PSCustomObject]@{
            id         = $id
            anio       = $anio
            ncontrato  = if ($ncontrato) { $ncontrato } else { "—" }
            cc         = $cc
            nombre     = $nombre
            tipo       = $tipo
            estado     = $estado
            ubicacion  = $ubicacion
            contrato   = $contrato
            facturado  = $facturado
            avance     = $avance
            plazo      = $plazo
            inicio     = if ($inicioFmt) { $inicioFmt } else { "—" }
            termino    = if ($terminoFmt) { $terminoFmt } else { "—" }
        }

        $id++
    }

    Write-Host "✅ Leídos $($proyectos.Count) proyectos del CSV" -ForegroundColor Green

} catch {
    Write-Host "❌ Error al leer el CSV: $_" -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit
}

if ($proyectos.Count -eq 0) {
    Write-Host "❌ No se leyó ningún proyecto. Abortando para no sobrescribir la web con datos vacíos." -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit
}

# ── 3. Generar el array JavaScript
Write-Host "🔧 Generando datos actualizados..." -ForegroundColor Yellow

$jsItems = @()
foreach ($p in $proyectos) {
    $plazoJs = if ($p.plazo -match '^\d') { $p.plazo } else { "null" }

    $ccJs = if ($p.cc) { "cc:'$($p.cc)', " } else { "" }
    $item = "  { id:$($p.id), $ccJs" + "anio:$($p.anio), ncontrato:'$($p.ncontrato -replace "'","\\'")', nombre:'$($p.nombre -replace "'","\\'")', tipo:'$($p.tipo)', estado:'$($p.estado)', ubicacion:'$($p.ubicacion -replace "'","\\'")', contrato:$($p.contrato), facturado:$($p.facturado), avance:$($p.avance), inicio:'$($p.inicio)', termino:'$($p.termino)', plazo:$plazoJs }"
    $jsItems += $item
}

$today = Get-Date -Format "yyyy-MM-dd HH:mm"
$newArray = "// Actualizado automáticamente el $today`nlet proyectos = [`n" + ($jsItems -join ",`n") + "`n];"

# ── 4. Reemplazar en index.html
Write-Host "📝 Actualizando index.html..." -ForegroundColor Yellow

$html = Get-Content $HtmlPath -Raw -Encoding UTF8

# Reemplazar el bloque de proyectos (desde el comentario hasta el cierre del array)
$pattern = '(?s)// ══.*?DATOS REALES.*?let proyectos = \[.*?\];'
$newHtml = $html -replace $pattern, $newArray

if ($newHtml -eq $html) {
    # Intentar reemplazo alternativo
    $pattern2 = '(?s)(// Actualizado automáticamente.*?|)let proyectos = \[.*?\];'
    $newHtml = $html -replace $pattern2, $newArray
}

$newHtml | Set-Content $HtmlPath -Encoding UTF8
Write-Host "✅ index.html actualizado" -ForegroundColor Green

# ── 5. Git: add, commit, push
Write-Host "🚀 Subiendo cambios a GitHub..." -ForegroundColor Yellow

Set-Location "C:\Users\tira1"

$gitAdd    = & git add index.html 2>&1
$gitCommit = & git commit -m "Actualizar datos proyectos - $today" 2>&1
$gitPush   = & git push 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Página web actualizada exitosamente" -ForegroundColor Green
} else {
    Write-Host $gitPush -ForegroundColor Yellow
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  ✅ Listo! La web se actualizó en ~2 min ║" -ForegroundColor Green
Write-Host "║  cinsajav.github.io/Monitor-proyectos    ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Read-Host "Presiona Enter para cerrar"
