$ErrorActionPreference = "Stop"

$xlsxPath  = "R:\Power BI Selección\Selecion Control de Gestion\Requerimientos\Req Postula Aqui.xlsx"
$pendPath  = "R:\Power BI Selección\Selecion Control de Gestion\Requerimientos\Solicitudes Pendiente.xlsx"
$repoDir   = "C:\Users\constanza.cesped\panel-requerimientos"
$buildDir  = Join-Path $repoDir "build"
$outJson   = Join-Path $buildDir "kpi_data.json"
$part1     = Join-Path $buildDir "dash_part1.html"
$part2     = Join-Path $buildDir "dash_part2.html"
$indexHtml = Join-Path $repoDir "index.html"
$rCopy     = "R:\Power BI Selección\Selecion Control de Gestion\Requerimientos\Panel de Requerimientos.html"

Write-Host "=================================================="
Write-Host " Actualizando Panel de Requerimientos"
Write-Host "=================================================="
Write-Host ""

if (-not (Test-Path $xlsxPath)) {
  Write-Host "ERROR: no se encuentra el archivo:" -ForegroundColor Red
  Write-Host "  $xlsxPath" -ForegroundColor Red
  Write-Host "Verifica que el disco R: esté conectado (red de Cygnus)." -ForegroundColor Red
  exit 1
}

function JEsc([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s.Replace('\', '\\').Replace('"', '\"')
  $s = $s.Replace("`r`n", ' ').Replace("`n", ' ').Replace("`r", ' ').Replace("`t", ' ')
  return $s
}

Write-Host "Paso 1/5: Leyendo Excel (esto puede tardar un minuto)..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wbPend = $excel.Workbooks.Open($pendPath, $false, $true)
$wsPend = $wbPend.Sheets.Item("Sheet1")
$usedPend = $wsPend.UsedRange
$dataPend = $usedPend.Value2
$rowsPend = $dataPend.GetLength(0)
$pendingSet = New-Object 'System.Collections.Generic.HashSet[string]'
for ($r = 2; $r -le $rowsPend; $r++) {
  $f = $dataPend[$r,4]
  if ($null -ne $f -and [string]$f -ne "") { [void]$pendingSet.Add([string]$f) }
}
$wbPend.Close($false)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsPend) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbPend) | Out-Null

$wb = $excel.Workbooks.Open($xlsxPath, $false, $true)
$ws = $wb.Sheets.Item("REQ")
$used = $ws.UsedRange
$data = $used.Value2
$rows = $data.GetLength(0)

$fields = @("fc_key","empresa","agrupador","referido","fecha_creacion","fecha_firma","contratado","region","cargo","motivo","fecha_envio","asignado","fecha_asignacion","fecha_termino_asignacion","causa_cierre","pendiente","folio","correlativo","nombre_supervisor","fecha_emision_contrato","nombre_persona_contratada","rut_contrato","legajo_contrato","fecha_inicio_labores","fecha_termino_contrato","fecha_inicio_servicio","fecha_termino_servicio","fecha_anulacion_seleccion","observacion_cierre")
$fieldsJson = '["' + ([string]::Join('","', $fields)) + '"]'

$sw = New-Object System.IO.StreamWriter($outJson, $false, (New-Object System.Text.UTF8Encoding($true)))
$sw.Write('{"meta":{"fields":' + $fieldsJson + ',"source":"Req Postula Aqui.xlsx (hoja REQ) + Solicitudes Pendiente.xlsx (Sheet1)","total_rows":')

$rowsJson = New-Object System.Collections.Generic.List[string]
$count = 0
for ($r = 2; $r -le $rows; $r++) {
  $fc = $data[$r,41]
  if ($null -eq $fc -or [string]$fc -eq "") { continue }

  $empresa    = [string]$data[$r,2]
  $agrupador  = [string]$data[$r,42]
  $referido   = [string]$data[$r,17]
  $fcreacion  = [string]$data[$r,5]
  $ffirma     = [string]$data[$r,29]
  $legajo     = $data[$r,27]
  $contratado = if ($null -ne $legajo -and [string]$legajo -ne "") { 1 } else { 0 }
  $region     = [string]$data[$r,40]
  $cargo      = [string]$data[$r,4]
  $motivo     = [string]$data[$r,22]
  $fenvio     = [string]$data[$r,18]
  $asignado   = [string]$data[$r,19]
  $fasign     = [string]$data[$r,20]
  $ftermasign = [string]$data[$r,21]
  $causa      = [string]$data[$r,38]

  $folioRaw = [string]$data[$r,6]
  $pendiente = if ($pendingSet.Contains($folioRaw)) { 1 } else { 0 }

  $correlativo   = [string]$data[$r,7]
  $supervisor    = [string]$data[$r,9]
  $femision      = [string]$data[$r,24]
  $nombreContrat = [string]$data[$r,25]
  $rutContrato   = [string]$data[$r,26]
  $legajoStr     = [string]$legajo
  $finicioLab    = [string]$data[$r,30]
  $ftermContrato = [string]$data[$r,31]
  $finicioServ   = [string]$data[$r,35]
  $ftermServ     = [string]$data[$r,36]
  $fanulacion    = [string]$data[$r,37]
  $obsCierre     = [string]$data[$r,39]

  $line = '["' + (JEsc([string]$fc)) + '","' + (JEsc($empresa)) + '","' + (JEsc($agrupador)) + '","' + (JEsc($referido)) + '","' + (JEsc($fcreacion)) + '","' + (JEsc($ffirma)) + '",' + $contratado + ',"' + (JEsc($region)) + '","' + (JEsc($cargo)) + '","' + (JEsc($motivo)) + '","' + (JEsc($fenvio)) + '","' + (JEsc($asignado)) + '","' + (JEsc($fasign)) + '","' + (JEsc($ftermasign)) + '","' + (JEsc($causa)) + '",' + $pendiente + ',"' + (JEsc($folioRaw)) + '","' + (JEsc($correlativo)) + '","' + (JEsc($supervisor)) + '","' + (JEsc($femision)) + '","' + (JEsc($nombreContrat)) + '","' + (JEsc($rutContrato)) + '","' + (JEsc($legajoStr)) + '","' + (JEsc($finicioLab)) + '","' + (JEsc($ftermContrato)) + '","' + (JEsc($finicioServ)) + '","' + (JEsc($ftermServ)) + '","' + (JEsc($fanulacion)) + '","' + (JEsc($obsCierre)) + '"]'

  $rowsJson.Add($line)
  $count++
}

$sw.Write($count)
$sw.Write('},"rows":[')
$sw.Write([string]::Join(",", $rowsJson))
$sw.Write(']}')
$sw.Close()

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "  -> $count filas exportadas."
Write-Host ""

Write-Host "Paso 2/5: Reconstruyendo el panel..."
$encBOM = New-Object System.Text.UTF8Encoding($true)
$c1 = [System.IO.File]::ReadAllText($part1, [System.Text.Encoding]::UTF8)
$cd = [System.IO.File]::ReadAllText($outJson, [System.Text.Encoding]::UTF8)
$c2 = [System.IO.File]::ReadAllText($part2, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($indexHtml, ($c1 + $cd + $c2), $encBOM)
Write-Host "  -> index.html reconstruido."
Write-Host ""

Write-Host "Paso 3/5: Copiando a R:\ ..."
Copy-Item -LiteralPath $indexHtml -Destination $rCopy -Force
Write-Host "  -> Copiado."
Write-Host ""

Write-Host "Paso 4/5: Subiendo a GitHub..."
$gitExe = "C:\Users\constanza.cesped\AppData\Local\Programs\Git\cmd\git.exe"
if (-not (Test-Path $gitExe)) {
  # Por si Git se reinstala en otra ubicación más adelante, intenta el PATH normal como respaldo.
  $gitExe = "git"
}
Set-Location $repoDir
& $gitExe add index.html
& $gitExe diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
  Write-Host "  -> Sin cambios respecto a la última publicación, no hay nada nuevo que subir."
} else {
  $fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
  & $gitExe commit -m "Actualizacion $fecha" | Out-Null
  & $gitExe push origin main
  if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host " No se pudo subir a GitHub." -ForegroundColor Yellow
    Write-Host " Lo más probable es que el token de acceso haya vencido." -ForegroundColor Yellow
    Write-Host ""
    Write-Host " Para arreglarlo:" -ForegroundColor Yellow
    Write-Host " 1. Ve a https://github.com/settings/tokens/new" -ForegroundColor Yellow
    Write-Host " 2. Genera un token nuevo (mismos pasos de siempre, marca 'repo')." -ForegroundColor Yellow
    Write-Host " 3. Pásale el token nuevo a Claude para que lo reconfigure." -ForegroundColor Yellow
    Write-Host ""
    Write-Host " (El panel de R:\ SÍ se actualizó bien - solo el link" -ForegroundColor Yellow
    Write-Host " público se quedó con los datos anteriores hasta que" -ForegroundColor Yellow
    Write-Host " arregles esto.)" -ForegroundColor Yellow
    Write-Host "==================================================" -ForegroundColor Yellow
  } else {
    Write-Host "  -> Publicado."
  }
}
Write-Host ""

Write-Host "Paso 5/5: Listo."
Write-Host ""
Write-Host "=================================================="
Write-Host " El link publico va a mostrar los datos nuevos"
Write-Host " en 1-2 minutos:"
Write-Host " https://requerimientos-seleccion.constanza-cesped.workers.dev/"
Write-Host "=================================================="
