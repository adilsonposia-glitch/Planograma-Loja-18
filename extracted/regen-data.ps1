$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot\..
$skipExport = $args -contains '-SkipExport'

# ---------- helpers ----------
function ToNum($x) {
  if ([string]::IsNullOrWhiteSpace($x)) { return 0 }
  try { return [double](($x -replace ',', '.') -replace '[^\d\.\-]', '') } catch { return 0 }
}
function ToInt($x) { return [int][Math]::Round((ToNum $x)) }
function NormDim($x) {
  $d = (($x + '') -replace '\s+', '' -replace '[xX×]', 'x' -replace ',', ',').ToLowerInvariant()
  if ($d -match '^1[,.]?0*0?x0?[,.]?35') { return '1,00x0,35' }
  if ($d -match '^1[,.]?0*0?x0?[,.]?40') { return '1,00x0,40' }
  if ($d -match '^1[,.]?30?x0?[,.]?35') { return '1,30x0,35' }
  if ($d -match '^1[,.]?30?x0?[,.]?40') { return '1,30x0,40' }
  if ($d -match '^1[,.]?50') { return '1,50x1,30' }
  if ($d -match '0[,.]?60') { return '0,60' }
  if ($d -match '0[,.]?50') { return '0,50' }
  return ($x + '').Trim()
}

# ---------- 1) Export Excel → CSV (aba LJ01, NÃO Cortes) ----------
$outDir = Join-Path (Get-Location) 'extracted\xlsx_v2'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$csvPath = Join-Path $outDir 'sheet_Mercadologico-LJ01.csv'

# Bloqueio: nunca usar artefatos de Cortes como fonte
if ($csvPath -match '(?i)corte') { throw 'Caminho CSV invalido (Cortes).' }

if (-not $skipExport) {
  $xlsx = Get-ChildItem -Filter '*.xlsx' | Where-Object { $_.Name -match '(?i)loja\s*18|planograma|layout' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $xlsx) { $xlsx = Get-ChildItem -Filter '*.xlsx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
  if (-not $xlsx) { throw 'Nenhum .xlsx Loja 18 encontrado na pasta do projeto.' }
  if ($xlsx.Name -match '(?i)^cortes?') { throw "Recusado workbook de Cortes: $($xlsx.Name)" }
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

    $preferred = @(
      'Mercadológico-LJ01-24.07'
      'Mercadologico-LJ01-24.07'
      'Mercadológico-LJ01'
      'Mercadologico-LJ01'
      'LJ01'
    )
    $ws = $null
    foreach ($want in $preferred) {
      foreach ($s in $wb.Worksheets) {
        $n = $s.Name.Trim()
        if ($n -eq $want -or $n -like '*LJ01*') { $ws = $s; break }
      }
      if ($ws) { break }
    }
    if (-not $ws) {
      foreach ($s in $wb.Worksheets) {
        if ($s.Name -match '(?i)corte') { continue }
        $u = $s.UsedRange
        $maxC = [Math]::Min([int]$u.Columns.Count, 25)
        $hdrBits = @()
        for ($c = 1; $c -le $maxC; $c++) { $hdrBits += [string]$u.Cells.Item(1, $c).Text }
        $hdrJoin = ($hdrBits -join '|')
        if ($hdrJoin -match 'Endereço Completo|Endereco Completo|Endereço Gôndola|Numero\s+Módulo|Número\s+Módulo') {
          $ws = $s; break
        }
      }
    }
    if (-not $ws) { throw 'Nenhuma aba LJ01 (Mercadológico) encontrada no Excel. Cortes nao sera usado.' }
    if ($ws.Name -match '(?i)corte') { throw "Recusado usar aba de Cortes: $($ws.Name)" }
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
    # espelho para analise
    $analiseDir = Join-Path (Get-Location) 'extracted\xlsx_analise_nova'
    New-Item -ItemType Directory -Force -Path $analiseDir | Out-Null
    [IO.File]::WriteAllText((Join-Path $analiseDir 'lj01_full_keycols.csv'), $sb.ToString(), [Text.UTF8Encoding]::new($false))
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
  if (-not (Test-Path -LiteralPath $csvPath)) {
    $fallback = Join-Path (Get-Location) 'extracted\xlsx_analise_nova\lj01_full_keycols.csv'
    if ((Test-Path -LiteralPath $fallback) -and ($fallback -notmatch '(?i)corte')) {
      Write-Host "Fallback LJ01: $fallback"
      $csvPath = $fallback
    } else {
      throw "CSV LJ01 nao encontrado: $csvPath (Cortes nao e fonte)."
    }
  }
  if ($csvPath -match '(?i)corte') { throw "Recusado CSV de Cortes: $csvPath" }
}

# ---------- 2) Parse layout LJ01 ----------
$lines = Get-Content -LiteralPath $csvPath -Encoding UTF8
Write-Host "Linhas CSV: $($lines.Count)"
if ($lines.Count -lt 2) { throw 'CSV LJ01 muito curto / vazio' }

$hdr = $lines[0] -split ';', -1
Write-Host ("Header: " + (($hdr | Select-Object -First 16) -join ' | '))
if (($hdr -join '|') -notmatch 'Endere|Gôndola|Gondola|Prateleiras|Bra') {
  throw 'CSV nao parece layout LJ01 (esperado Endereço Completo / Prateleiras / Braços).'
}

$typeLabel = @{
  GO = 'Gôndola Mercearia'
  GB = 'Gôndola Perfumaria'
  BR = 'Rack Bebidas'
  BP = 'Gôndola Pães'
  GP = 'Gôndola Padaria'
  GC = 'Expositor Carvão'
  EX = 'Expositor Ovos'
  IC = 'Ilha Congelados'
  MH = 'Mesa Hortifruti'
  MM = 'Mesa Mercearia'
}
$allowed = 'GO|BR|BP|GB|GC|GP|EX|IC|MH|MM'

$faces = [ordered]@{}
$dimCounts = @{}
$pratByDim = @{}
$bracoByProf = @{ '35cm' = 0; '40cm' = 0; '60cm' = 0 }
$sumBarras = 0; $sumGanchos = 0; $sumApar = 0; $sumCestos = 0; $sumCestoVinhos = 0

for ($i = 1; $i -lt $lines.Count; $i++) {
  $p = $lines[$i] -split ';', -1
  if ($p.Count -lt 12) { continue }

  $setor = ($p[1] + '').Trim()
  $endCompleto = ($p[2] + '').Trim()
  $endGondola = ($p[3] + '').Trim()
  $lado = ($p[4] + '').Trim().ToUpperInvariant()
  $modNraw = ($p[5] + '').Trim()
  $base = ($p[6] + '').Trim()
  $prat = ToInt $p[7]
  $bracos = ToInt $p[8]
  $barras = ToInt $p[9]
  $ganchos = ToInt $p[10]
  $apar = ToInt $p[11]
  $cestos = ToInt $p[12]
  $tp = ($p[13] + '').Trim()
  $spec = if ($p.Count -gt 14) { ($p[14] + '').Trim() } else { '' }
  $cat = if ($p.Count -gt 15) { ($p[15] + '').Trim() } else { '' }
  $dim = if ($p.Count -gt 16 -and ($p[16] + '').Trim()) { ($p[16] + '').Trim() } else { $base }
  $obs = if ($p.Count -gt 20) { ($p[20] + '').Trim() } else { '' }
  $cestoVinhos = if ($p.Count -gt 21) { ToInt $p[21] } else { 0 }

  $modId = $endCompleto
  $face = $null
  $modN = $null

  if ($modId -match "^((?:$allowed)-\d+(?:\.[AB])?)(?:\.(\d+))?$") {
    $face = $Matches[1]
    $modN = $Matches[2]
  } elseif ($endGondola -match "^((?:$allowed)-\d+)$") {
    $isle = $Matches[1]
    if ($lado -match '^[AB]$') { $face = "$isle.$lado" } else { $face = $isle }
    if ($modNraw -match '^\d+$') {
      $modN = $modNraw
      $modId = "$face.$modN"
    }
  } else {
    continue
  }

  # MH/MM/GC (e similares): endereço = equipamento sem .N → 1 módulo
  if (-not $modN -and $face -match '^(MH|MM|GC)-\d+$') {
    $modN = '1'
    if ($modId -eq $face) { $modId = "$face.1" }
  }

  # Linha só de cabeçalho de face (sem número de módulo)
  if (-not $modN) {
    if (-not $faces.Contains($face)) {
      $faces[$face] = @{
        id = $face; setor = $setor; cat = $cat; spec = $spec
        linear = $dim; headerMods = 0
        mods = [System.Collections.ArrayList]@()
      }
    } else {
      if ($cat) { $faces[$face].cat = $cat }
      if ($setor) { $faces[$face].setor = $setor }
      if ($spec) { $faces[$face].spec = $spec }
    }
    continue
  }

  if (-not $faces.Contains($face)) {
    $faces[$face] = @{
      id = $face; setor = $setor; cat = $cat; spec = $spec
      linear = $dim; headerMods = 0
      mods = [System.Collections.ArrayList]@()
    }
  }
  if ($cat -and -not $faces[$face].cat) { $faces[$face].cat = $cat }
  if ($setor) { $faces[$face].setor = $setor }

  [void]$faces[$face].mods.Add([ordered]@{
    id = $modId; n = [int]$modN; cat = $cat; spec = $spec
    obs = $obs; dim = $dim
    prat = $prat; bracos = $bracos; barras = $barras
    ganchos = $ganchos; apar = $apar; cestos = $cestos
    sintetico = $false
  })

  $nd = NormDim $dim
  if ($nd) {
    if (-not $dimCounts.ContainsKey($nd)) { $dimCounts[$nd] = 0 }
    $dimCounts[$nd]++
    if (-not $pratByDim.ContainsKey($nd)) { $pratByDim[$nd] = 0 }
    $pratByDim[$nd] += $prat
  }
  if ($nd -match '0,35|1,00x0,35|1,30x0,35') { $bracoByProf['35cm'] += $bracos }
  elseif ($nd -match '0,60') { $bracoByProf['60cm'] += $bracos }
  else { $bracoByProf['40cm'] += $bracos }

  $sumBarras += $barras
  $sumGanchos += $ganchos
  $sumApar += $apar
  $sumCestos += $cestos
  $sumCestoVinhos += $cestoVinhos
}

$islands = [ordered]@{}
foreach ($k in $faces.Keys) {
  $isle = ($k -replace '\.[AB]$', '')
  if (-not $islands.Contains($isle)) {
    $islands[$isle] = [System.Collections.ArrayList]@()
  }
  [void]$islands[$isle].Add($faces[$k])
}

$gondolas = [System.Collections.ArrayList]@()
foreach ($isle in $islands.Keys) {
  $prefix = $isle.Substring(0, 2)
  if (-not $typeLabel.ContainsKey($prefix)) { continue }
  $fs = @($islands[$isle])
  $faceArr = [System.Collections.ArrayList]@()
  $totMods = 0; $totPrat = 0; $totBracos = 0
  foreach ($f in $fs) {
    $modArr = [System.Collections.ArrayList]@()
    $prat = 0; $bracos = 0
    foreach ($m in @($f.mods)) {
      $prat += [int]$m.prat; $bracos += [int]$m.bracos
      $niveis = if ([int]$m.prat -gt 0) { [int]$m.prat } else { 4 }
      [void]$modArr.Add([ordered]@{
        id = $m.id; n = $m.n; categoria = $m.cat; especificacao = $m.spec
        observacao = $m.obs; dim = $m.dim
        prateleiras = [int]$m.prat; bracos = [int]$m.bracos; barras = [int]$m.barras
        ganchos = [int]$m.ganchos; aparadores = [int]$m.apar; cestos = [int]$m.cestos
        niveis = $niveis; sintetico = [bool]$m.sintetico
      })
    }
    if ($modArr.Count -eq 0) { continue }
    $hm = $modArr.Count
    $totMods += $hm; $totPrat += $prat; $totBracos += $bracos
    $faceCat = $f.cat
    if (-not $faceCat) {
      $faceCat = (($modArr | ForEach-Object { $_.categoria } | Where-Object { $_ } | Select-Object -Unique) -join ' · ')
    }
    [void]$faceArr.Add([ordered]@{
      id = $f.id; categoria = $faceCat; especificacao = $f.spec; linear = $f.linear
      headerMods = $hm; prateleiras = $prat; bracos = $bracos; modulos = @($modArr.ToArray())
    })
  }
  if ($faceArr.Count -eq 0) { continue }
  $setor = ($fs | ForEach-Object { $_.setor } | Where-Object { $_ } | Select-Object -First 1)
  $cats = ($faceArr | ForEach-Object { $_.categoria } | Where-Object { $_ } | Select-Object -Unique) -join ' · '
  $fill = if ($totPrat -gt 0) { 'completo' }
    elseif (($prefix -match '^(IC|MH|MM|GC)$') -and $totMods -gt 0) { 'completo' }
    elseif ($totMods -gt 0) { 'parcial' }
    else { 'vazio' }
  [void]$gondolas.Add([ordered]@{
    id = $isle; tipo = $prefix; tipoLabel = $typeLabel[$prefix]; setor = $setor
    categorias = $cats; faces = @($faceArr.ToArray())
    totalMods = $totMods; totalPrat = $totPrat; totalBracos = $totBracos
    preenchimento = $fill
  })
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

# Totais somados do corpo LJ01 (não há linha de totalizador no layout Cortes)
$m100x35 = if ($dimCounts['1,00x0,35']) { [int]$dimCounts['1,00x0,35'] } else { 0 }
$m100x40 = if ($dimCounts['1,00x0,40']) { [int]$dimCounts['1,00x0,40'] } else { 0 }
$m13x35  = if ($dimCounts['1,30x0,35']) { [int]$dimCounts['1,30x0,35'] } else { 0 }
$m13x40  = if ($dimCounts['1,30x0,40']) { [int]$dimCounts['1,30x0,40'] } else { 0 }
$m150x130 = if ($dimCounts['1,50x1,30']) { [int]$dimCounts['1,50x1,30'] } else { 0 }

$p100x35 = if ($pratByDim['1,00x0,35']) { [int]$pratByDim['1,00x0,35'] } else { 0 }
$p100x40 = if ($pratByDim['1,00x0,40']) { [int]$pratByDim['1,00x0,40'] } else { 0 }
$p13x35  = if ($pratByDim['1,30x0,35']) { [int]$pratByDim['1,30x0,35'] } else { 0 }
$p13x40  = if ($pratByDim['1,30x0,40']) { [int]$pratByDim['1,30x0,40'] } else { 0 }
$p13x60  = 0
$br35 = [int]$bracoByProf['35cm']
$br40 = [int]$bracoByProf['40cm']
$br60 = [int]$bracoByProf['60cm']

$modulosTot = $m100x35 + $m100x40 + $m13x35 + $m13x40 + $m150x130
if ($modulosTot -eq 0) { $modulosTot = $all }
$pratTot = 0; $bracosTot = 0
foreach ($g in $gondolas) { $pratTot += [int]$g.totalPrat; $bracosTot += [int]$g.totalBracos }
$ganchosTot = $sumGanchos
$aparTot = $sumApar
$cestosTot = $sumCestos + $sumCestoVinhos
$bar100 = 0; $bar130 = $sumBarras
$b100x35 = $m100x35; $b100x40 = $m100x40; $b13x35 = $m13x35; $b13x40 = $m13x40
$b050 = 0; $b060 = 0; $g4035 = 0; $g3530 = $sumGanchos; $g3025 = 0; $gvass = 0
$ap100 = 0; $ap130 = $sumApar; $ce100 = 0; $ce130 = $sumCestos; $cestoVinhos = $sumCestoVinhos

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
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,00 m × 35 cm'; qtd = $p100x35; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,00 m × 40 cm'; qtd = $p100x40; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 35 cm'; qtd = $p13x35; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 40 cm'; qtd = $p13x40; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '35 cm'; qtd = $br35; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '40 cm'; qtd = $br40; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '60 cm'; qtd = $br60; unidade = 'un' }
  @{ grupo = 'Barras'; peca = 'Barra'; tamanho = '1,30 m'; qtd = $bar130; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho'; tamanho = 'conforme módulo'; qtd = $ganchosTot; unidade = 'un' }
  @{ grupo = 'Aparadores'; peca = 'Aparador'; tamanho = '1,30 m'; qtd = $ap130; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = '1,30 m'; qtd = $ce130; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = 'Vinhos'; qtd = $cestoVinhos; unidade = 'un' }
)

$summary = @{
  loja = 'Loja 18 — Supermarket Santa Cruz'
  pavimento = '1º Pavimento'
  fonte = 'Mercadológico-LJ01 — Excel Loja 18 (Cortes ignorado)'
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
      )
      bracos = @(
        @{ qtd = $br35; tamanho = '35cm' }
        @{ qtd = $br40; tamanho = '40cm' }
        @{ qtd = $br60; tamanho = '60cm' }
      )
      barras = @(
        @{ qtd = $bar130; tamanho = '1,30m' }
      )
      ganchos = @(
        @{ qtd = $ganchosTot; tamanho = 'conforme módulo' }
      )
      aparadores = @(
        @{ qtd = $ap130; tamanho = '1,30m' }
      )
      cestos = @(
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
    faces = @($gondolas | ForEach-Object { $_.faces } | ForEach-Object { $_ }).Count
    modulesHardware = $hw
    modulesTotal = $all
    go = @($gondolas | Where-Object { $_.tipo -eq 'GO' }).Count
    br = @($gondolas | Where-Object { $_.tipo -eq 'BR' }).Count
    bp = @($gondolas | Where-Object { $_.tipo -eq 'BP' }).Count
    gp = @($gondolas | Where-Object { $_.tipo -eq 'GP' }).Count
    gb = @($gondolas | Where-Object { $_.tipo -eq 'GB' }).Count
    ex = @($gondolas | Where-Object { $_.tipo -eq 'EX' }).Count
    gc = @($gondolas | Where-Object { $_.tipo -eq 'GC' }).Count
    ic = @($gondolas | Where-Object { $_.tipo -eq 'IC' }).Count
    mh = @($gondolas | Where-Object { $_.tipo -eq 'MH' }).Count
    mm = @($gondolas | Where-Object { $_.tipo -eq 'MM' }).Count
  }
}

$data = @{ summary = $summary; gondolas = @($gondolas.ToArray()) }
$planDir = Join-Path (Get-Location) 'planograma'
New-Item -ItemType Directory -Force -Path $planDir | Out-Null

$jsonPretty = $data | ConvertTo-Json -Depth 12
$jsonCompact = $data | ConvertTo-Json -Depth 12 -Compress
$jsPayload = "window.PLANOGRAMA_DATA=$jsonCompact;"

[IO.File]::WriteAllText((Join-Path $planDir 'data.json'), $jsonPretty, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $planDir 'data.js'), $jsPayload, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $planDir 'data.test.js'), $jsPayload, [Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host '=== REGEN OK (fonte: LJ01 / Excel) ==='
Write-Host ("Equipamentos={0} ModsHW={1} ModsTot={2}" -f $gondolas.Count, $hw, $all)
Write-Host ("Totalizador: mods={0} prat={1} bracos={2} apar={3} ganchos={4}" -f $modulosTot, $pratTot, $bracosTot, $aparTot, $ganchosTot)
Write-Host ("data.test.js = {0} bytes" -f (Get-Item (Join-Path $planDir 'data.test.js')).Length)
$gondolas | ForEach-Object {
  '{0,-7} mods={1,3} prat={2,4} {3}' -f $_.id, $_.totalMods, $_.totalPrat, $_.preenchimento
}
