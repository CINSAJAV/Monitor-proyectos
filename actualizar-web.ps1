# ============================================================
# CINSA - Actualizador automático de Monitor de Proyectos
# Lee el Excel de Google Drive y actualiza la página web
# ============================================================

$ExcelPath = "G:\Mi unidad\5. Estudios\Dashboard Proyectos.xlsx"
$HtmlPath  = "C:\Users\tira1\index.html"

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   CINSA - Actualizador de Proyectos      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ── 1. Verificar que el archivo Excel existe
if (-not (Test-Path $ExcelPath)) {
    Write-Host "❌ No se encontró el archivo Excel en Google Drive." -ForegroundColor Red
    Write-Host "   Verifica que Google Drive esté sincronizado." -ForegroundColor Yellow
    Read-Host "Presiona Enter para salir"
    exit
}

Write-Host "✅ Excel encontrado: $ExcelPath" -ForegroundColor Green

# ── 2. Abrir Excel y leer datos
Write-Host "📖 Leyendo datos del Excel..." -ForegroundColor Yellow

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath)
    $ws = $wb.Sheets.Item(1)

    $proyectos = @()
    $row = 3  # Los datos empiezan en la fila 3 (fila 1 es título, fila 2 es encabezados)
    $id = 1

    while ($true) {
        $nombre = $ws.Cells($row, 3).Text.Trim()
        if ([string]::IsNullOrEmpty($nombre) -or $nombre -eq "Total Neto" -or $nombre -eq "Total Bruto") { break }

        $anioRaw  = $ws.Cells($row, 1).Text.Trim()
        $ncontrato= $ws.Cells($row, 2).Text.Trim()
        $tipoRaw  = $ws.Cells($row, 4).Text.Trim()
        $contratoRaw = ($ws.Cells($row, 6).Text.Trim()) -replace '[\.\$ ,]',''
        $plazoRaw = $ws.Cells($row, 7).Text.Trim()
        $inicioRaw= $ws.Cells($row, 8).Text.Trim()
        $terminoRaw=$ws.Cells($row, 9).Text.Trim()
        $avanceRaw= ($ws.Cells($row, 11).Text.Trim()) -replace '[%\s]',''
        $facturadoRaw = ($ws.Cells($row, 16).Text.Trim()) -replace '[\.\$ ,]',''

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

        $proyectos += [PSCustomObject]@{
            id         = $id
            anio       = $anio
            ncontrato  = if ($ncontrato) { $ncontrato } else { "—" }
            nombre     = $nombre
            tipo       = $tipo
            estado     = $estado
            ubicacion  = $ubicacion
            contrato   = $contrato
            facturado  = $facturado
            avance     = $avance
            plazo      = $plazo
            inicio     = if ($inicioRaw) { $inicioRaw } else { "—" }
            termino    = if ($terminoRaw) { $terminoRaw } else { "—" }
        }

        $id++
        $row++
    }

    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Write-Host "✅ Leídos $($proyectos.Count) proyectos del Excel" -ForegroundColor Green

} catch {
    Write-Host "❌ Error al leer el Excel: $_" -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit
}

# ── 3. Generar el array JavaScript
Write-Host "🔧 Generando datos actualizados..." -ForegroundColor Yellow

$jsItems = @()
foreach ($p in $proyectos) {
    $plazoJs = if ($p.plazo -match '^\d') { $p.plazo } else { "null" }

    $item = "  { id:$($p.id), anio:$($p.anio), ncontrato:'$($p.ncontrato -replace "'","\\'")', nombre:'$($p.nombre -replace "'","\\'")', tipo:'$($p.tipo)', estado:'$($p.estado)', ubicacion:'$($p.ubicacion -replace "'","\\'")', contrato:$($p.contrato), facturado:$($p.facturado), avance:$($p.avance), inicio:'$($p.inicio)', termino:'$($p.termino)', plazo:$plazoJs }"
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
