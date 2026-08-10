$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot\..
$skipExport = $args -contains '-SkipExport'

# ---------- 1) Export Excel → CSV ----------
$outDir = Join-Path (Get-Location) 'extracted\xlsx_v2'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$csvPath = Join-Path $outDir 'sheet_Corte Loja 18.csv'

if (-not $skipExport) {
$xlsx = Get-ChildItem -Filter '*.xlsx' | Select-Object -First 1
if (-not $xlsx) { throw 'Nenhum .xlsx encontrado na pasta do projeto.' }
Write-Host "XLSX: $($xlsx.FullName)  ($($xlsx.LastWriteTime))"

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 800

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
  $wb = $excel.Workbooks.Open($xlsx.FullName)
  $names = @($wb.Worksheets | ForEach-Object { $_.Name })
  Write-Host ("Abas: " + ($names -join ' | '))
  # Preferir estritamente o layout "Corte" (totalizador em colunas de hardware).
  # A aba LJ01 tem layout diferente (GP-01 etc.) e NÃO deve alimentar este parser.
  $preferred = @('Corte Loja 18', 'Corte Loja18', 'Corte')
  $ws = $null
  foreach ($want in $preferred) {
    foreach ($s in $wb.Worksheets) {
      if ($s.Name.Trim() -eq $want -or $s.Name.Trim() -like 'Corte*Loja*18*') { $ws = $s; break }
    }
    if ($ws) { break }
  }
  if (-not $ws) {
    # fallback: aba com cabeçalho de hardware do Corte (Módulos 1,3x0,40)
    foreach ($s in $wb.Worksheets) {
      $u = $s.UsedRange
      $maxC = [Math]::Min([int]$u.Columns.Count, 40)
      $hdrBits = @()
      for ($c = 1; $c -le $maxC; $c++) { $hdrBits += [string]$u.Cells.Item(2, $c).Text }
      $hdrJoin = ($hdrBits -join '|')
      if ($hdrJoin -match 'Módulos 1,3x0,40|Prateleiras 1,30x40|Braços 40cm') { $ws = $s; break }
    }
  }
  if (-not $ws) { throw 'Nenhuma aba no layout Corte Loja 18 encontrada no Excel.' }
  Write-Host "Usando aba: $($ws.Name)"

  $used = $ws.UsedRange
  $rows = [int]$used.Rows.Count
  $cols = [int]$used.Columns.Count
  Write-Host "Dimensao: ${rows}x${cols}"

  $sb = New-Object System.Text.StringBuilder
  for ($r = 1; $r -le $rows; $r++) {
    $cells = @()
    for ($c = 1; $c -le $cols; $c++) {
      $v = $used.Cells.Item($r, $c).Text
      if ($null -eq $v) { $v = '' }
      $v = [string]$v
      $v = $v -replace '[\r\n]+', ' '
      $v = $v -replace ';', ','
      $cells += $v
    }
    [void]$sb.AppendLine(($cells -join ';'))
  }
  [IO.File]::WriteAllText($csvPath, $sb.ToString(), [Text.UTF8Encoding]::new($false))
  Write-Host "CSV: $csvPath ($((Get-Item $csvPath).Length) bytes)"
}
finally {
  if ($wb) { $wb.Close($false) }
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
} else {
  Write-Host "Reusando CSV: $csvPath"
  if (-not (Test-Path -LiteralPath $csvPath)) { throw "CSV nao encontrado: $csvPath" }
}

# ---------- 2) Parse CSV → data ----------
function ToNum($x) {
  if ([string]::IsNullOrWhiteSpace($x)) { return 0 }
  try { return [double](($x -replace ',', '.') -replace '[^\d\.\-]', '') } catch { return 0 }
}
function ToInt($x) { return [int][Math]::Round((ToNum $x)) }

$lines = [System.Collections.Generic.List[string]]::new()
$lines.AddRange([string[]](Get-Content -LiteralPath $csvPath -Encoding UTF8))
Write-Host "Linhas CSV: $($lines.Count)"
if ($lines.Count -lt 4) { throw 'CSV muito curto / vazio' }

# GP-01 existe na planilha LJ01 mas costuma faltar no layout Corte — completar se ausente
if (-not ($lines | Where-Object { $_ -match ';GP-01([.;]|$)' })) {
  $setorPad = 'PADARIA / PAES'
  $bpSample = $lines | Where-Object { $_ -match ';BP-03\.B\.6;' } | Select-Object -First 1
  if ($bpSample -match '^[^;]*;([^;]+);') { $setorPad = $Matches[1] }
  $insertAt = -1
  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match ';BP-03\.B\.6;') { $insertAt = $i + 1; break }
  }
  if ($insertAt -ge 0) {
    $ocirc = [char]0x00F4; $oac = [char]0x00F3; $ced = [char]0x00E7
    $gp = New-Object System.Collections.Generic.List[string]
    $gp.Add((";{0};GP-01;G{1}ndola; 17 M{2}dulos  ;Padaria;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" -f $setorPad, $ocirc, $oac))
    for ($n = 1; $n -le 17; $n++) {
      $spec = "base + 5 prateleiras + 10 bra${ced}os + 6 aparadores"
      $hw = '1,30x0,35;;;;1;;;;1;;;;5;;10;;;;;;;;6;;;;;;;;;;;'
      $gp.Add((";{0};GP-01.{1};---;{2};Padaria;{3}" -f $setorPad, $n, $spec, $hw))
    }
    $lines.InsertRange($insertAt, $gp)
    Write-Host ("AVISO: GP-01 ausente no Corte — injetados {0} modulos a partir da planilha LJ01" -f 17)
  }
}

$hdr = $lines[1] -split ';', -1
$tr  = $lines[2] -split ';', -1

Write-Host '--- Totalizador (linha oficial) ---'
for ($i = 8; $i -lt [Math]::Min($hdr.Count, $tr.Count); $i++) {
  if ($hdr[$i] -or $tr[$i]) {
    '{0,2} | {1,-40} | {2}' -f $i, $hdr[$i], $tr[$i]
  }
}

$faces = [ordered]@{}
for ($i = 3; $i -lt $lines.Count; $i++) {
  $p = $lines[$i] -split ';', -1
  $end = ($p[2] + '').Trim()
  if ($end -notmatch '^((?:GO|BR|BP|GB|GC|GP|EX)-\d+(?:\.[AB])?)(?:\.(\d+))?$') { continue }
  $face = $Matches[1]; $modN = $Matches[2]
  if (-not $faces.Contains($face)) {
    $faces[$face] = @{
      id = $face; setor = $p[1]; cat = $p[5]; spec = $p[4]
      linear = $p[6]; headerMods = 0; mods = @()
    }
  }
  if (-not $modN) {
    if ($p[4] -match '(\d+)\s*M') { $faces[$face].headerMods = [int]$Matches[1] }
    if ($p[5]) { $faces[$face].cat = $p[5] }
    if ($p[1]) { $faces[$face].setor = $p[1] }
    if ($p[4]) { $faces[$face].spec = $p[4] }
    if ($p[6]) { $faces[$face].linear = $p[6] }
  } else {
    # Cols 16-21 = mercearia (35/40 cm); cols 37-38 = bebidas profundas (0,60 cm) — BR-05
    $prat = [int]((ToNum $p[16]) + (ToNum $p[17]) + (ToNum $p[18]) + (ToNum $p[19]))
    $bracos = [int]((ToNum $p[20]) + (ToNum $p[21]))
    if ($p.Count -gt 38) {
      $prat += [int](ToNum $p[38])   # Prateleiras 1,30x0,60
      $bracos += [int](ToNum $p[37]) # Braços 0,60 cm
    }
    $faces[$face].mods += ,@( @{
      id = $end; n = [int]$modN; cat = ($p[5] + ''); spec = ($p[4] + '')
      obs = ($p[7] + ''); dim = ($p[6] + '')
      prat = $prat
      bracos = $bracos
      barras = [int]((ToNum $p[22]) + (ToNum $p[23]))
      ganchos = [int]((ToNum $p[24]) + (ToNum $p[25]) + (ToNum $p[26]))
      apar = [int]((ToNum $p[27]) + (ToNum $p[28]))
      cestos = [int]((ToNum $p[29]) + (ToNum $p[30]))
      sintetico = $false
    })
  }
}

foreach ($k in @($faces.Keys)) {
  $f = $faces[$k]
  if ($f.mods.Count -eq 0 -and $f.headerMods -gt 0) {
    $syn = @()
    for ($n = 1; $n -le $f.headerMods; $n++) {
      $syn += @{
        id = ("$k.$n"); n = $n; cat = $f.cat
        spec = 'Módulo (sem detalhe de hardware)'; obs = ''; dim = ''
        prat = 0; bracos = 0; barras = 0; ganchos = 0; apar = 0; cestos = 0
        sintetico = $true
      }
    }
    $f.mods = $syn
  }
}

$typeLabel = @{
  GO = 'Gôndola Mercearia'
  GB = 'Gôndola Perfumaria'
  BR = 'Rack Bebidas'
  BP = 'Gôndola Pães'
  GP = 'Gôndola Padaria'
  GC = 'Expositor'
  EX = 'Expositor Ovos'
}

$islands = [ordered]@{}
foreach ($k in $faces.Keys) {
  $isle = ($k -replace '\.[AB]$', '')
  if (-not $islands.Contains($isle)) { $islands[$isle] = @() }
  $islands[$isle] += $faces[$k]
}

$gondolas = @()
foreach ($isle in $islands.Keys) {
  $prefix = $isle.Substring(0, 2)
  $fs = $islands[$isle]
  $faceArr = @()
  $totMods = 0; $totPrat = 0; $totBracos = 0
  foreach ($f in $fs) {
    $modArr = @()
    $prat = 0; $bracos = 0
    foreach ($m in $f.mods) {
      $prat += $m.prat; $bracos += $m.bracos
      $niveis = if ($m.prat -gt 0) { $m.prat } else { 4 }
      $modArr += @{
        id = $m.id; n = $m.n; categoria = $m.cat; especificacao = $m.spec
        observacao = $m.obs; dim = $m.dim
        prateleiras = $m.prat; bracos = $m.bracos; barras = $m.barras
        ganchos = $m.ganchos; aparadores = $m.apar; cestos = $m.cestos
        niveis = $niveis; sintetico = $m.sintetico
      }
    }
    $hm = if ($f.headerMods -gt 0) { $f.headerMods } else { $modArr.Count }
    $totMods += $hm; $totPrat += $prat; $totBracos += $bracos
    $faceArr += @{
      id = $f.id; categoria = $f.cat; especificacao = $f.spec; linear = $f.linear
      headerMods = $hm; prateleiras = $prat; bracos = $bracos; modulos = $modArr
    }
  }
  $setor = ($fs | ForEach-Object { $_.setor } | Where-Object { $_ } | Select-Object -First 1)
  $cats = ($faceArr | ForEach-Object { $_.categoria } | Where-Object { $_ } | Select-Object -Unique) -join ' · '
  $fill = if ($totPrat -gt 0) { 'completo' } elseif ($totMods -gt 0) { 'parcial' } else { 'vazio' }
  $gondolas += @{
    id = $isle; tipo = $prefix; tipoLabel = $typeLabel[$prefix]; setor = $setor
    categorias = $cats; faces = $faceArr
    totalMods = $totMods; totalPrat = $totPrat; totalBracos = $totBracos
    preenchimento = $fill
  }
}

$hw = 0; $all = 0
foreach ($g in $gondolas) {
  foreach ($f in $g.faces) {
    foreach ($m in $f.modulos) {
      $all++
      if (-not $m.sintetico) { $hw++ }
    }
  }
}

# Totals from official row (indices from original Corte Loja 18 layout)
$m100x35 = ToInt $tr[8]
$m100x40 = ToInt $tr[9]
$m13x35  = ToInt $tr[10]
$m13x40  = ToInt $tr[11]
$b100x35 = ToInt $tr[12]; if ($b100x35 -eq 0) { $b100x35 = $m100x35 }
$b100x40 = ToInt $tr[13]; if ($b100x40 -eq 0) { $b100x40 = $m100x40 }
$b13x35  = ToInt $tr[14]; if ($b13x35 -eq 0) { $b13x35 = $m13x35 }
$b13x40  = ToInt $tr[15]; if ($b13x40 -eq 0) { $b13x40 = $m13x40 }
# Ordem real no Corte Loja 18: 1,00x35 · 1,30x40 · 1,30x35 · 1,00x40
$p100x35 = ToInt $tr[16]
$p13x40  = ToInt $tr[17]
$p13x35  = ToInt $tr[18]
$p100x40 = ToInt $tr[19]
$br35    = ToInt $tr[20]
$br40    = ToInt $tr[21]
$bar100  = ToInt $tr[22]
$bar130  = ToInt $tr[23]
$g4035   = ToInt $tr[24]
$g3530   = ToInt $tr[25]
$g3025   = ToInt $tr[26]
$ap100   = ToInt $tr[27]
$ap130   = ToInt $tr[28]
$ce100   = ToInt $tr[29]
$ce130   = ToInt $tr[30]
$gvass   = 0
if ($tr.Count -gt 31) { $gvass += ToInt $tr[31] }
if ($tr.Count -gt 32) { $gvass += ToInt $tr[32] }
# Colunas novas (bebidas / vinhos) — presentes no LJ01 atual
$m150x130 = 0; $b050 = 0; $cestoVinhos = 0; $b060 = 0; $br60 = 0; $p13x60 = 0
if ($tr.Count -gt 33) { $m150x130 = ToInt $tr[33] }
if ($tr.Count -gt 34) { $b050 = ToInt $tr[34] }
if ($tr.Count -gt 35) { $cestoVinhos = ToInt $tr[35] }
if ($tr.Count -gt 36) { $b060 = ToInt $tr[36] }
if ($tr.Count -gt 37) { $br60 = ToInt $tr[37] }
if ($tr.Count -gt 38) { $p13x60 = ToInt $tr[38] }

$modulosTot = $m100x35 + $m100x40 + $m13x35 + $m13x40 + $m150x130
$pratTot = $p100x35 + $p100x40 + $p13x35 + $p13x40 + $p13x60
$bracosTot = $br35 + $br40 + $br60
$ganchosTot = $g4035 + $g3530 + $g3025 + $gvass
$aparTot = $ap100 + $ap130
$cestosTot = $ce100 + $ce130 + $cestoVinhos

# Fallback if official totals row empty: sum from body
if ($modulosTot -eq 0) {
  Write-Host 'AVISO: totalizador oficial zerado — somando corpo'
  $modulosTot = ($gondolas | Measure-Object totalMods -Sum).Sum
  $pratTot = ($gondolas | Measure-Object totalPrat -Sum).Sum
  $bracosTot = ($gondolas | Measure-Object totalBracos -Sum).Sum
}

$pecas = @(
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,00 × 0,35 m'; qtd = $m100x35; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,00 × 0,40 m'; qtd = $m100x40; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,30 × 0,35 m'; qtd = $m13x35; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,30 × 0,40 m'; qtd = $m13x40; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,50 × 1,30 m'; qtd = $m150x130; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,00 × 0,35 m'; qtd = $b100x35; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,00 × 0,40 m'; qtd = $b100x40; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,30 × 0,35 m'; qtd = $b13x35; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,30 × 0,40 m'; qtd = $b13x40; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '0,50 m'; qtd = $b050; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '0,60 m'; qtd = $b060; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,00 m × 35 cm'; qtd = $p100x35; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,00 m × 40 cm'; qtd = $p100x40; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 35 cm'; qtd = $p13x35; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 40 cm'; qtd = $p13x40; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 60 cm'; qtd = $p13x60; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '35 cm'; qtd = $br35; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '40 cm'; qtd = $br40; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '60 cm'; qtd = $br60; unidade = 'un' }
  @{ grupo = 'Barras'; peca = 'Barra'; tamanho = '1,00 m'; qtd = $bar100; unidade = 'un' }
  @{ grupo = 'Barras'; peca = 'Barra'; tamanho = '1,30 m'; qtd = $bar130; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho horizontal'; tamanho = '40 × 35 cm'; qtd = $g4035; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho horizontal'; tamanho = '35 × 30 cm'; qtd = $g3530; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho horizontal'; tamanho = '30 × 25 cm'; qtd = $g3025; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho vassoura'; tamanho = '40 × 35 cm'; qtd = $gvass; unidade = 'un' }
  @{ grupo = 'Aparadores'; peca = 'Aparador'; tamanho = '1,00 m'; qtd = $ap100; unidade = 'un' }
  @{ grupo = 'Aparadores'; peca = 'Aparador'; tamanho = '1,30 m'; qtd = $ap130; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = '1,00 m'; qtd = $ce100; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = '1,30 m'; qtd = $ce130; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = 'Vinhos'; qtd = $cestoVinhos; unidade = 'un' }
)

$summary = @{
  loja = 'Loja 18 — Supermarket Santa Cruz'
  pavimento = '1º Pavimento'
  fonte = 'Mercadológico-LJ01-24.07 / Corte Loja 18'
  totalizador = @{
    modulos = [int]$modulosTot
    prateleiras = [int]$pratTot
    bracos = [int]$bracosTot
    m100x35 = $m100x35; m100x40 = $m100x40; m13x35 = $m13x35; m13x40 = $m13x40
    barras130 = $bar130
    ganchos = [int]$ganchosTot
    aparadores = [int]$aparTot
    cestos = [int]$cestosTot
    porTamanho = @{
      modulos = @(
        @{ qtd = $m100x35; tamanho = '1,00x0,35' }
        @{ qtd = $m100x40; tamanho = '1,00x0,40' }
        @{ qtd = $m13x35; tamanho = '1,30x0,35' }
        @{ qtd = $m13x40; tamanho = '1,30x0,40' }
        @{ qtd = $m150x130; tamanho = '1,50x1,30' }
      )
      prateleiras = @(
        @{ qtd = $p100x35; tamanho = '1,00x35cm' }
        @{ qtd = $p100x40; tamanho = '1,00x40cm' }
        @{ qtd = $p13x35; tamanho = '1,30x35cm' }
        @{ qtd = $p13x40; tamanho = '1,30x40cm' }
        @{ qtd = $p13x60; tamanho = '1,30x60cm' }
      )
      bracos = @(
        @{ qtd = $br35; tamanho = '35cm' }
        @{ qtd = $br40; tamanho = '40cm' }
        @{ qtd = $br60; tamanho = '60cm' }
      )
      barras = @(
        @{ qtd = $bar100; tamanho = '1,00m' }
        @{ qtd = $bar130; tamanho = '1,30m' }
      )
      ganchos = @(
        @{ qtd = $g4035; tamanho = '40x35cm' }
        @{ qtd = $g3530; tamanho = '35x30cm' }
        @{ qtd = $g3025; tamanho = '30x25cm' }
        @{ qtd = $gvass; tamanho = 'vassoura' }
      )
      aparadores = @(
        @{ qtd = $ap100; tamanho = '1,00m' }
        @{ qtd = $ap130; tamanho = '1,30m' }
      )
      cestos = @(
        @{ qtd = $ce100; tamanho = '1,00m' }
        @{ qtd = $ce130; tamanho = '1,30m' }
        @{ qtd = $cestoVinhos; tamanho = 'vinhos' }
      )
    }
  }
  pecas = $pecas
  specs = @{
    moduloMenor = 1.10; moduloMaior = 1.30; altura = 1.85
    corredor = 1.80; corredorCentral = 2.80
  }
  contagem = @{
    gondolas = $gondolas.Count
    faces = $faces.Count
    modulesHardware = $hw
    modulesTotal = $all
    go = @($gondolas | Where-Object { $_.tipo -eq 'GO' }).Count
    br = @($gondolas | Where-Object { $_.tipo -eq 'BR' }).Count
    bp = @($gondolas | Where-Object { $_.tipo -eq 'BP' }).Count
    gp = @($gondolas | Where-Object { $_.tipo -eq 'GP' }).Count
    gb = @($gondolas | Where-Object { $_.tipo -eq 'GB' }).Count
    ex = @($gondolas | Where-Object { $_.tipo -eq 'EX' }).Count
    gc = @($gondolas | Where-Object { $_.tipo -eq 'GC' }).Count
  }
}

$data = @{ summary = $summary; gondolas = $gondolas }
$planDir = Join-Path (Get-Location) 'planograma'
New-Item -ItemType Directory -Force -Path $planDir | Out-Null

# Prefer .NET serializer for unicode stability
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = 50MB
# Convert nested hashtables: ConvertTo-Json then parse back as object for serializer? 
# Simpler: use ConvertTo-Json -Depth
$jsonPretty = $data | ConvertTo-Json -Depth 12
$jsonCompact = $data | ConvertTo-Json -Depth 12 -Compress

[IO.File]::WriteAllText((Join-Path $planDir 'data.json'), $jsonPretty, [Text.UTF8Encoding]::new($false))
$jsPayload = "window.PLANOGRAMA_DATA=$jsonCompact;"
[IO.File]::WriteAllText((Join-Path $planDir 'data.js'), $jsPayload, [Text.UTF8Encoding]::new($false))
# index.html carrega data.test.js — manter espelhado
[IO.File]::WriteAllText((Join-Path $planDir 'data.test.js'), $jsPayload, [Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host '=== REGEN OK ==='
Write-Host ("Equipamentos={0} Faces={1} ModsHW={2} ModsTot={3}" -f $gondolas.Count, $faces.Count, $hw, $all)
Write-Host ("Totalizador: mods={0} prat={1} bracos={2} apar={3} ganchos={4}" -f $modulosTot, $pratTot, $bracosTot, $aparTot, $ganchosTot)
Write-Host ("data.js = {0} bytes" -f (Get-Item (Join-Path $planDir 'data.js')).Length)
$gondolas | ForEach-Object {
  '{0,-7} mods={1,3} prat={2,4} {3}' -f $_.id, $_.totalMods, $_.totalPrat, $_.preenchimento
}
