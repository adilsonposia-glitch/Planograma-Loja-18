$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot\..
$skipExport = $args -contains '-SkipExport'

# ---------- helpers ----------
function ToNum($x) {
  if ([string]::IsNullOrWhiteSpace($x)) { return 0 }
  $t = (($x + '') -replace '\s+', '' -replace ',', '.')
  if ($t -eq '.' -or $t -eq '-' -or $t -eq ',') { return 0 }
  try { return [double]($t -replace '[^\d\.\-]', '') } catch { return 0 }
}
function ToInt($x) { return [int][Math]::Round((ToNum $x)) }
function NormHdr([string]$s) {
  $t = ($s + '').Trim().ToLowerInvariant()
  if ($t.Length -eq 0) { return '' }
  $t = $t.Normalize([Text.NormalizationForm]::FormD)
  $t = [regex]::Replace($t, '\p{Mn}', '')
  $t = $t -replace '\s+', ' '
  $t = $t -replace '[^a-z0-9,x ]', ''
  return $t.Trim()
}
function NormBase($x) {
  $d = (($x + '') -replace '\s+', '' -replace '[xX×]', 'x' -replace '\.', ',').ToLowerInvariant()
  if (-not $d) { return '' }
  # Dimensão 3 partes (ex.: 1,85x1,30x0,40) → largura x profundidade
  if ($d -match '^[\d,]+\s*x\s*([\d,]+)\s*x\s*([\d,]+)$') {
    $d = $Matches[1] + 'x' + $Matches[2]
  }
  $d = $d -replace '^1,3x', '1,30x' -replace '^1,0x', '1,00x' -replace '^1x', '1,00x'
  $d = $d -replace 'x0,4$', 'x0,40' -replace 'x1$', 'x1,00'
  switch -Regex ($d) {
    '^1,00x0,35' { return '1,00x0,35' }
    '^1,00x0,40' { return '1,00x0,40' }
    '^1,30x0,35' { return '1,30x0,35' }
    '^1,30x0,40' { return '1,30x0,40' }
    '^1,30x0,60' { return '1,30x0,60' }
    '^1,00x1,00' { return '1,00x1,00' }
    '^1,20x1,20' { return '1,20x1,20' }
    '^2,00x1,00' { return '2,00x1,00' }
    '^5,00x2,00' { return '5,00x2,00' }
    default { return $d }
  }
}

# ---------- 1) Export Excel → CSV (aba LJ01, NÃO Cortes) ----------
$outDir = Join-Path (Get-Location) 'extracted\xlsx_v2'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$csvPath = Join-Path $outDir 'sheet_Mercadologico-LJ01.csv'

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
        $maxC = [Math]::Min([int]$u.Columns.Count, 40)
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
    $lastHdr = 0
    for ($c = 1; $c -le $cols; $c++) {
      $h = ([string]$used.Cells.Item(1, $c).Text).Trim()
      if ($h) { $lastHdr = $c }
    }
    $exportCols = if ($lastHdr -gt 0) { $lastHdr } else { [Math]::Min($cols, 40) }
    Write-Host "Dimensao used=${rows}x${cols}  export=${rows}x${exportCols}"

    $sb = New-Object System.Text.StringBuilder
    for ($r = 1; $r -le $rows; $r++) {
      $cells = @()
      for ($c = 1; $c -le $exportCols; $c++) {
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

# ---------- 2) Parse layout LJ01 (colunas pelo nome) ----------
$lines = [IO.File]::ReadAllLines($csvPath, [Text.UTF8Encoding]::new($false))
Write-Host "Linhas CSV: $($lines.Count)"
if ($lines.Count -lt 2) { throw 'CSV LJ01 muito curto / vazio' }

$hdr = $lines[0] -split ';', -1
$hdrNorm = @($hdr | ForEach-Object { NormHdr $_ })
Write-Host ("Header ($($hdr.Count) cols): " + (($hdr | ForEach-Object { $_.Trim() }) -join ' | '))

function FindCol([string[]]$aliases) {
  foreach ($a in $aliases) {
    $na = NormHdr $a
    if (-not $na) { continue }
    for ($i = 0; $i -lt $hdrNorm.Count; $i++) {
      if ($hdrNorm[$i] -eq $na) { return $i }
    }
  }
  return -1
}
function Cell($parts, $idx) {
  if ($idx -lt 0) { return '' }
  if ($idx -ge $parts.Count) { return '' }
  return ($parts[$idx] + '').Trim()
}

$col = @{
  nro = FindCol @('Nº', 'N', 'No')
  setor = FindCol @('Setor / Área', 'Setor / Area', 'Setor')
  endCompleto = FindCol @('Endereço Completo', 'Endereco Completo')
  endGondola = FindCol @('Endereço Gôndola', 'Endereco Gondola')
  lado = FindCol @('Lado')
  modN = FindCol @('Número  Módulo', 'Número Módulo', 'Numero Modulo')
  base = FindCol @('Base')
  prat = FindCol @('Prateleiras')
  br40 = FindCol @('Braços 0,40', 'Bracos 0,40')
  br35 = FindCol @('Braços 0,35', 'Bracos 0,35')
  br60 = FindCol @('Braços 0,60', 'Bracos 0,60')
  bracosLegacy = FindCol @('Braços', 'Bracos')
  barra100 = FindCol @('Barra 1,00')
  barras130 = FindCol @('Barras 1,30')
  barrasLegacy = FindCol @('Barras')
  g3530 = FindCol @('Ganchos 35x30cm')
  g3025 = FindCol @('Ganchos 30x25cm')
  g4035 = FindCol @('Ganchos 40x35cm')
  gVass = FindCol @('Ganchos Grosso Vassoura')
  ganchosLegacy = FindCol @('Ganchos')
  anteparoAramado = FindCol @('Anteparo Aramado 0,05cm', 'Aparadores')
  cestoCarvao = FindCol @('Cesto carvão', 'Cesto carvao')
  cestoVinhos = FindCol @('Cesto vinhos', 'Cesto Vinhos')
  anteparoAcrilico = FindCol @('Anteparo Acrílico 1,00', 'Anteparo Acrilico 1,00')
  cestos100 = FindCol @('Cestos 1,00')
  cestos130 = FindCol @('Cestos 1,30')
  cestosLegacy = FindCol @('Cestos')
  div130 = FindCol @('Divisor cesto 1,30')
  div100 = FindCol @('Divisor cesto 1,00')
  cabeceira = FindCol @('Cabeceira Gôndola', 'Cabeceira Gondola')
  portaEtiqueta = FindCol @('Porta Etiqueta')
  tp = FindCol @('Tp Mobiliário', 'Tp Mobiliario')
  spec = FindCol @('Especificação', 'Especificacao')
  cat = FindCol @('Categoria(s) Mercadológica(s) Proposta(s)', 'Categoria(s) Mercadologica(s) Proposta(s)')
  dim = FindCol @('Dimensão', 'Dimensao')
  largura = FindCol @('Largura Gôndola', 'Largura Gondola', 'Largura')
  prof = FindCol @('Profundidade Gôndola', 'Profundidade Gondola', 'Profundidade')
  altura = FindCol @('Altura  Gôndola', 'Altura Gôndola', 'Altura Gondola')
  obs = FindCol @('Observações', 'Observacoes')
}

Write-Host 'Mapa de colunas:'
$col.GetEnumerator() | Sort-Object Name | ForEach-Object {
  $h = if ($_.Value -ge 0) { $hdr[$_.Value].Trim() } else { '(ausente)' }
  '  {0,-18} -> [{1,2}] {2}' -f $_.Name, $_.Value, $h
}

if ($col.endCompleto -lt 0 -and $col.endGondola -lt 0) {
  throw 'CSV nao parece layout LJ01 (esperado Endereço Completo / Endereço Gôndola).'
}

$typeLabel = @{
  GO = 'Gôndola Mercearia'
  GB = 'Gôndola Perfumaria'
  BR = 'Gôndola Bebidas'
  BA = 'Gôndola Parede Bebidas'
  BP = 'Expositor Pães'
  GP = 'Gôndola Padaria'
  GS = 'Expositor Salgados'
  GC = 'Expositor Carvão'
  EC = 'Expositor Carvão'
  EX = 'Expositor Ovos'
  IC = 'Ilha Congelados'
  MH = 'Mesa Hortifruti'
  MM = 'Mesa Mercearia'
  MQ = 'Mesa Quente'
  MS = 'Mesa Salgados'
}
$allowed = 'GO|GB|BR|BA|BP|GP|GS|GC|EC|EX|IC|MH|MM|MQ|MS'
$mesaLike = '^(IC|MH|MM|GC|EC|MQ|MS|GS)$'

$faces = [ordered]@{}
$dimCounts = @{}
$pratByDim = @{}
$sum = @{
  prat = 0; br40 = 0; br35 = 0; br60 = 0
  barra100 = 0; barras130 = 0
  g3530 = 0; g3025 = 0; g4035 = 0; gVass = 0
  anteparoAramado = 0; anteparoAcrilico = 0
  cestoCarvao = 0; cestoVinhos = 0; cestos100 = 0; cestos130 = 0
  div100 = 0; div130 = 0; cabeceira = 0; portaEtiqueta = 0
}

for ($i = 1; $i -lt $lines.Count; $i++) {
  $p = $lines[$i] -split ';', -1
  $setor = Cell $p $col.setor
  $endCompleto = Cell $p $col.endCompleto
  $endGondola = Cell $p $col.endGondola
  $lado = (Cell $p $col.lado).ToUpperInvariant()
  $modNraw = Cell $p $col.modN
  $base = Cell $p $col.base
  $prat = ToInt (Cell $p $col.prat)
  $br40 = if ($col.br40 -ge 0) { ToInt (Cell $p $col.br40) } else { 0 }
  $br35 = if ($col.br35 -ge 0) { ToInt (Cell $p $col.br35) } else { 0 }
  $br60 = if ($col.br60 -ge 0) { ToInt (Cell $p $col.br60) } else { 0 }
  if ($col.br40 -lt 0 -and $col.bracosLegacy -ge 0) { $br40 = ToInt (Cell $p $col.bracosLegacy) }
  $bracos = $br40 + $br35 + $br60
  $barra100 = ToInt (Cell $p $col.barra100)
  $barras130 = if ($col.barras130 -ge 0) { ToInt (Cell $p $col.barras130) } else { ToInt (Cell $p $col.barrasLegacy) }
  $barras = $barra100 + $barras130
  $g3530 = ToInt (Cell $p $col.g3530)
  $g3025 = ToInt (Cell $p $col.g3025)
  $g4035 = ToInt (Cell $p $col.g4035)
  $gVass = ToInt (Cell $p $col.gVass)
  if ($col.g3530 -lt 0 -and $col.ganchosLegacy -ge 0) { $g3530 = ToInt (Cell $p $col.ganchosLegacy) }
  $ganchos = $g3530 + $g3025 + $g4035 + $gVass
  $anteparoAramado = ToInt (Cell $p $col.anteparoAramado)
  $anteparoAcrilico = ToInt (Cell $p $col.anteparoAcrilico)
  $cestoCarvao = ToInt (Cell $p $col.cestoCarvao)
  $cestoVinhos = ToInt (Cell $p $col.cestoVinhos)
  $cestos100 = ToInt (Cell $p $col.cestos100)
  $cestos130 = if ($col.cestos130 -ge 0) { ToInt (Cell $p $col.cestos130) } else { ToInt (Cell $p $col.cestosLegacy) }
  $cestos = $cestos100 + $cestos130 + $cestoVinhos + $cestoCarvao
  $div100 = ToInt (Cell $p $col.div100)
  $div130 = ToInt (Cell $p $col.div130)
  $cabeceira = ToInt (Cell $p $col.cabeceira)
  $portaEtiqueta = ToInt (Cell $p $col.portaEtiqueta)
  $tp = Cell $p $col.tp
  $spec = Cell $p $col.spec
  $cat = Cell $p $col.cat
  $dimRaw = Cell $p $col.dim
  $largura = Cell $p $col.largura
  $prof = Cell $p $col.prof
  $altura = Cell $p $col.altura
  $obs = Cell $p $col.obs
  $dim = if ($dimRaw) { $dimRaw } else { $base }

  $modId = $endCompleto
  $face = $null
  $modN = $null

  if ($modId -match "^((?:$allowed)-\d+(?:\.[ABPCabpc])?)(?:\.(\d+))?$") {
    $face = $Matches[1]
    $modN = $Matches[2]
  } elseif ($endGondola -match "^((?:$allowed)-\d+)$") {
    $isle = $Matches[1]
    if ($lado -match '^[ABPC]$') { $face = "$isle.$lado" } else { $face = $isle }
    if ($modNraw -match '^\d+$') {
      $modN = $modNraw
      $modId = "$face.$modN"
    }
  } else {
    continue
  }

  if (-not $modN -and $modNraw -match '^\d+$') { $modN = $modNraw }
  $prefFace = ($face.Substring(0, [Math]::Min(2, $face.Length)))
  if (-not $modN -and $prefFace -match $mesaLike) {
    $modN = '1'
    if ($modId -eq $face) { $modId = "$face.1" }
  }

  if (-not $modN) {
    if (-not $faces.Contains($face)) {
      $faces[$face] = @{
        id = $face; setor = $setor; cat = $cat; spec = $spec; lado = $lado
        linear = $dim; tp = $tp; headerMods = 0
        mods = [System.Collections.ArrayList]@()
      }
    } else {
      if ($cat) { $faces[$face].cat = $cat }
      if ($setor) { $faces[$face].setor = $setor }
      if ($spec) { $faces[$face].spec = $spec }
      if ($lado) { $faces[$face].lado = $lado }
    }
    continue
  }

  if (-not $faces.Contains($face)) {
    $faces[$face] = @{
      id = $face; setor = $setor; cat = $cat; spec = $spec; lado = $lado
      linear = $dim; tp = $tp; headerMods = 0
      mods = [System.Collections.ArrayList]@()
    }
  }
  if ($cat -and -not $faces[$face].cat) { $faces[$face].cat = $cat }
  if ($setor) { $faces[$face].setor = $setor }
  if ($lado) { $faces[$face].lado = $lado }
  if ($tp -and -not $faces[$face].tp) { $faces[$face].tp = $tp }

  $nd = NormBase $(if ($base) { $base } else { $dim })

  [void]$faces[$face].mods.Add([ordered]@{
    id = $modId; n = [int]$modN; cat = $cat; spec = $spec
    obs = $obs; dim = $dim; base = $base; lado = $lado; tp = $tp
    largura = $largura; profundidade = $prof; altura = $altura
    prat = $prat; bracos = $bracos; barras = $barras
    ganchos = $ganchos; apar = $anteparoAramado; cestos = $cestos
    bom = [ordered]@{
      prateleiras = $prat
      bracos040 = $br40; bracos035 = $br35; bracos060 = $br60
      barra100 = $barra100; barras130 = $barras130
      ganchos3530 = $g3530; ganchos3025 = $g3025
      ganchos4035 = $g4035; ganchosVassoura = $gVass
      anteparoAramado = $anteparoAramado; anteparoAcrilico = $anteparoAcrilico
      cestos100 = $cestos100; cestos130 = $cestos130
      cestoVinhos = $cestoVinhos; cestoCarvao = $cestoCarvao
      divisorCesto100 = $div100; divisorCesto130 = $div130
      cabeceira = $cabeceira; portaEtiqueta = $portaEtiqueta
    }
    sintetico = $false
  })

  if ($nd) {
    if (-not $dimCounts.ContainsKey($nd)) { $dimCounts[$nd] = 0 }
    $dimCounts[$nd]++
    if (-not $pratByDim.ContainsKey($nd)) { $pratByDim[$nd] = 0 }
    $pratByDim[$nd] += $prat
  }

  $sum.prat += $prat
  $sum.br40 += $br40; $sum.br35 += $br35; $sum.br60 += $br60
  $sum.barra100 += $barra100; $sum.barras130 += $barras130
  $sum.g3530 += $g3530; $sum.g3025 += $g3025; $sum.g4035 += $g4035; $sum.gVass += $gVass
  $sum.anteparoAramado += $anteparoAramado; $sum.anteparoAcrilico += $anteparoAcrilico
  $sum.cestoCarvao += $cestoCarvao; $sum.cestoVinhos += $cestoVinhos
  $sum.cestos100 += $cestos100; $sum.cestos130 += $cestos130
  $sum.div100 += $div100; $sum.div130 += $div130
  $sum.cabeceira += $cabeceira; $sum.portaEtiqueta += $portaEtiqueta
}

$islands = [ordered]@{}
foreach ($k in $faces.Keys) {
  $isle = ($k -replace '\.[ABPCabpc]$', '')
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
  $totBarras = 0; $totGanchos = 0; $totAnteparos = 0; $totCestos = 0
  foreach ($f in $fs) {
    $modArr = [System.Collections.ArrayList]@()
    $prat = 0; $bracos = 0
    foreach ($m in @($f.mods)) {
      $prat += [int]$m.prat; $bracos += [int]$m.bracos
      $totBarras += [int]$m.barras; $totGanchos += [int]$m.ganchos
      $totAnteparos += [int]$m.apar + [int]$m.bom.anteparoAcrilico
      $totCestos += [int]$m.cestos
      $niveis = if ([int]$m.prat -gt 0) { [int]$m.prat } else { 4 }
      [void]$modArr.Add([ordered]@{
        id = $m.id; n = $m.n; categoria = $m.cat; especificacao = $m.spec
        observacao = $m.obs; dim = $m.dim; base = $m.base; lado = $m.lado
        tpMobiliario = $m.tp
        largura = $m.largura; profundidade = $m.profundidade; altura = $m.altura
        prateleiras = [int]$m.prat; bracos = [int]$m.bracos; barras = [int]$m.barras
        ganchos = [int]$m.ganchos; aparadores = [int]$m.apar; cestos = [int]$m.cestos
        bom = $m.bom
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
      id = $f.id; lado = $f.lado; categoria = $faceCat; especificacao = $f.spec
      tpMobiliario = $f.tp; linear = $f.linear
      headerMods = $hm; prateleiras = $prat; bracos = $bracos; modulos = @($modArr.ToArray())
    })
  }
  if ($faceArr.Count -eq 0) { continue }
  $setor = ($fs | ForEach-Object { $_.setor } | Where-Object { $_ } | Select-Object -First 1)
  $cats = ($faceArr | ForEach-Object { $_.categoria } | Where-Object { $_ } | Select-Object -Unique) -join ' · '
  $fill = if ($totPrat -gt 0) { 'completo' }
    elseif (($prefix -match $mesaLike) -and $totMods -gt 0) { 'completo' }
    elseif ($totMods -gt 0) { 'parcial' }
    else { 'vazio' }
  [void]$gondolas.Add([ordered]@{
    id = $isle; tipo = $prefix; tipoLabel = $typeLabel[$prefix]; setor = $setor
    categorias = $cats; faces = @($faceArr.ToArray())
    totalMods = $totMods; totalPrat = $totPrat; totalBracos = $totBracos
    totalBarras = $totBarras; totalGanchos = $totGanchos
    totalAnteparos = $totAnteparos; totalCestos = $totCestos
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

function DimQ($key) { if ($dimCounts.ContainsKey($key)) { [int]$dimCounts[$key] } else { 0 } }
function PratQ($key) { if ($pratByDim.ContainsKey($key)) { [int]$pratByDim[$key] } else { 0 } }

$m100x35 = DimQ '1,00x0,35'
$m100x40 = DimQ '1,00x0,40'
$m13x35  = DimQ '1,30x0,35'
$m13x40  = DimQ '1,30x0,40'
$m13x60  = DimQ '1,30x0,60'
$m100x100 = DimQ '1,00x1,00'
$m120x120 = DimQ '1,20x1,20'
$m200x100 = DimQ '2,00x1,00'
$m500x200 = DimQ '5,00x2,00'

$p100x35 = PratQ '1,00x0,35'
$p100x40 = PratQ '1,00x0,40'
$p13x35  = PratQ '1,30x0,35'
$p13x40  = PratQ '1,30x0,40'
$p13x60  = PratQ '1,30x0,60'

$modulosTot = 0
foreach ($k in $dimCounts.Keys) { $modulosTot += [int]$dimCounts[$k] }
if ($modulosTot -eq 0) { $modulosTot = $all }
$pratTot = $sum.prat
$bracosTot = $sum.br40 + $sum.br35 + $sum.br60
$ganchosTot = $sum.g3530 + $sum.g3025 + $sum.g4035 + $sum.gVass
$aparTot = $sum.anteparoAramado
$cestosTot = $sum.cestos100 + $sum.cestos130 + $sum.cestoVinhos + $sum.cestoCarvao

$pecas = @(
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,00 × 0,35 m'; qtd = $m100x35; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,00 × 0,40 m'; qtd = $m100x40; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,30 × 0,35 m'; qtd = $m13x35; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,30 × 0,40 m'; qtd = $m13x40; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,30 × 0,60 m'; qtd = $m13x60; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,00 × 1,00 m'; qtd = $m100x100; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '1,20 × 1,20 m'; qtd = $m120x120; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '2,00 × 1,00 m'; qtd = $m200x100; unidade = 'un' }
  @{ grupo = 'Módulos'; peca = 'Módulo'; tamanho = '5,00 × 2,00 m'; qtd = $m500x200; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,00 × 0,35 m'; qtd = $m100x35; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,00 × 0,40 m'; qtd = $m100x40; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,30 × 0,35 m'; qtd = $m13x35; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,30 × 0,40 m'; qtd = $m13x40; unidade = 'un' }
  @{ grupo = 'Bases'; peca = 'Base'; tamanho = '1,30 × 0,60 m'; qtd = $m13x60; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,00 m × 35 cm'; qtd = $p100x35; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,00 m × 40 cm'; qtd = $p100x40; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 35 cm'; qtd = $p13x35; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 40 cm'; qtd = $p13x40; unidade = 'un' }
  @{ grupo = 'Prateleiras'; peca = 'Prateleira'; tamanho = '1,30 m × 60 cm'; qtd = $p13x60; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '35 cm'; qtd = [int]$sum.br35; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '40 cm'; qtd = [int]$sum.br40; unidade = 'un' }
  @{ grupo = 'Braços'; peca = 'Braço'; tamanho = '60 cm'; qtd = [int]$sum.br60; unidade = 'un' }
  @{ grupo = 'Barras'; peca = 'Barra'; tamanho = '1,00 m'; qtd = [int]$sum.barra100; unidade = 'un' }
  @{ grupo = 'Barras'; peca = 'Barra'; tamanho = '1,30 m'; qtd = [int]$sum.barras130; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho'; tamanho = '35 × 30 cm'; qtd = [int]$sum.g3530; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho'; tamanho = '30 × 25 cm'; qtd = [int]$sum.g3025; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho'; tamanho = '40 × 35 cm'; qtd = [int]$sum.g4035; unidade = 'un' }
  @{ grupo = 'Ganchos'; peca = 'Gancho vassoura'; tamanho = 'grosso'; qtd = [int]$sum.gVass; unidade = 'un' }
  @{ grupo = 'Anteparos'; peca = 'Anteparo aramado'; tamanho = '0,05 cm'; qtd = [int]$sum.anteparoAramado; unidade = 'un' }
  @{ grupo = 'Anteparos'; peca = 'Anteparo acrílico'; tamanho = '1,00 m'; qtd = [int]$sum.anteparoAcrilico; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = '1,00 m'; qtd = [int]$sum.cestos100; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = '1,30 m'; qtd = [int]$sum.cestos130; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto'; tamanho = 'Vinhos'; qtd = [int]$sum.cestoVinhos; unidade = 'un' }
  @{ grupo = 'Cestos'; peca = 'Cesto carvão'; tamanho = 'carvão'; qtd = [int]$sum.cestoCarvao; unidade = 'un' }
  @{ grupo = 'Divisores'; peca = 'Divisor de cesto'; tamanho = '1,00 m'; qtd = [int]$sum.div100; unidade = 'un' }
  @{ grupo = 'Divisores'; peca = 'Divisor de cesto'; tamanho = '1,30 m'; qtd = [int]$sum.div130; unidade = 'un' }
  @{ grupo = 'Estrutura'; peca = 'Cabeceira'; tamanho = 'gôndola'; qtd = [int]$sum.cabeceira; unidade = 'un' }
  @{ grupo = 'Estrutura'; peca = 'Porta etiqueta'; tamanho = 'un'; qtd = [int]$sum.portaEtiqueta; unidade = 'un' }
)

$colunasNome = @($hdr | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$summary = @{
  loja = 'Loja 18 — Supermarket Santa Cruz'
  pavimento = '1º Pavimento'
  fonte = 'Mercadológico-LJ01 — Excel Loja 18 (Cortes ignorado)'
  colunas = $colunasNome
  totalizador = @{
    modulos = [int]$modulosTot
    prateleiras = [int]$pratTot
    bracos = [int]$bracosTot
    m100x35 = $m100x35; m100x40 = $m100x40; m13x35 = $m13x35; m13x40 = $m13x40
    m13x60 = $m13x60
    barras100 = [int]$sum.barra100
    barras130 = [int]$sum.barras130
    ganchos = [int]$ganchosTot
    aparadores = [int]$aparTot
    anteparoAcrilico = [int]$sum.anteparoAcrilico
    cestos = [int]$cestosTot
    divisores = [int]($sum.div100 + $sum.div130)
    cabeceiras = [int]$sum.cabeceira
    porTamanho = @{
      modulos = @(
        @{ qtd = $m100x35; tamanho = '1,00x0,35' }
        @{ qtd = $m100x40; tamanho = '1,00x0,40' }
        @{ qtd = $m13x35; tamanho = '1,30x0,35' }
        @{ qtd = $m13x40; tamanho = '1,30x0,40' }
        @{ qtd = $m13x60; tamanho = '1,30x0,60' }
        @{ qtd = $m100x100; tamanho = '1,00x1,00' }
        @{ qtd = $m120x120; tamanho = '1,20x1,20' }
        @{ qtd = $m200x100; tamanho = '2,00x1,00' }
        @{ qtd = $m500x200; tamanho = '5,00x2,00' }
      )
      prateleiras = @(
        @{ qtd = $p100x35; tamanho = '1,00x35cm' }
        @{ qtd = $p100x40; tamanho = '1,00x40cm' }
        @{ qtd = $p13x35; tamanho = '1,30x35cm' }
        @{ qtd = $p13x40; tamanho = '1,30x40cm' }
        @{ qtd = $p13x60; tamanho = '1,30x60cm' }
      )
      bracos = @(
        @{ qtd = [int]$sum.br35; tamanho = '35cm' }
        @{ qtd = [int]$sum.br40; tamanho = '40cm' }
        @{ qtd = [int]$sum.br60; tamanho = '60cm' }
      )
      barras = @(
        @{ qtd = [int]$sum.barra100; tamanho = '1,00m' }
        @{ qtd = [int]$sum.barras130; tamanho = '1,30m' }
      )
      ganchos = @(
        @{ qtd = [int]$sum.g3530; tamanho = '35x30cm' }
        @{ qtd = [int]$sum.g3025; tamanho = '30x25cm' }
        @{ qtd = [int]$sum.g4035; tamanho = '40x35cm' }
        @{ qtd = [int]$sum.gVass; tamanho = 'vassoura' }
      )
      anteparos = @(
        @{ qtd = [int]$sum.anteparoAramado; tamanho = 'aramado 0,05cm' }
        @{ qtd = [int]$sum.anteparoAcrilico; tamanho = 'acrílico 1,00m' }
      )
      cestos = @(
        @{ qtd = [int]$sum.cestos100; tamanho = '1,00m' }
        @{ qtd = [int]$sum.cestos130; tamanho = '1,30m' }
        @{ qtd = [int]$sum.cestoVinhos; tamanho = 'vinhos' }
        @{ qtd = [int]$sum.cestoCarvao; tamanho = 'carvão' }
      )
      divisores = @(
        @{ qtd = [int]$sum.div100; tamanho = '1,00m' }
        @{ qtd = [int]$sum.div130; tamanho = '1,30m' }
      )
      estrutura = @(
        @{ qtd = [int]$sum.cabeceira; tamanho = 'cabeceira' }
        @{ qtd = [int]$sum.portaEtiqueta; tamanho = 'porta etiqueta' }
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
    gb = @($gondolas | Where-Object { $_.tipo -eq 'GB' }).Count
    br = @($gondolas | Where-Object { $_.tipo -eq 'BR' }).Count
    ba = @($gondolas | Where-Object { $_.tipo -eq 'BA' }).Count
    bp = @($gondolas | Where-Object { $_.tipo -eq 'BP' }).Count
    gp = @($gondolas | Where-Object { $_.tipo -eq 'GP' }).Count
    gs = @($gondolas | Where-Object { $_.tipo -eq 'GS' }).Count
    gc = @($gondolas | Where-Object { $_.tipo -eq 'GC' }).Count
    ec = @($gondolas | Where-Object { $_.tipo -eq 'EC' }).Count
    ex = @($gondolas | Where-Object { $_.tipo -eq 'EX' }).Count
    ic = @($gondolas | Where-Object { $_.tipo -eq 'IC' }).Count
    mh = @($gondolas | Where-Object { $_.tipo -eq 'MH' }).Count
    mm = @($gondolas | Where-Object { $_.tipo -eq 'MM' }).Count
    mq = @($gondolas | Where-Object { $_.tipo -eq 'MQ' }).Count
    ms = @($gondolas | Where-Object { $_.tipo -eq 'MS' }).Count
  }
}

$data = @{ summary = $summary; gondolas = @($gondolas.ToArray()) }
$planDir = Join-Path (Get-Location) 'planograma'
New-Item -ItemType Directory -Force -Path $planDir | Out-Null

$jsonPretty = $data | ConvertTo-Json -Depth 14
$jsonCompact = $data | ConvertTo-Json -Depth 14 -Compress
$jsPayload = "window.PLANOGRAMA_DATA=$jsonCompact;"

[IO.File]::WriteAllText((Join-Path $planDir 'data.json'), $jsonPretty, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $planDir 'data.js'), $jsPayload, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $planDir 'data.test.js'), $jsPayload, [Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host '=== REGEN OK (fonte: LJ01 / Excel) ==='
Write-Host ("Equipamentos={0} ModsHW={1} ModsTot={2}" -f $gondolas.Count, $hw, $all)
Write-Host ("BOM: prat={0} bracos={1} ganchos={2} anteparoAramado={3} cestos={4}" -f $pratTot, $bracosTot, $ganchosTot, $aparTot, $cestosTot)
$c = $summary.contagem
Write-Host ("Tipos: GO={0} GB={1} BR={2} BA={3} BP={4} GP={5} GS={6} IC={7} MH={8} MM={9} MS={10} MQ={11} EX={12} EC={13} GC={14}" -f $c.go,$c.gb,$c.br,$c.ba,$c.bp,$c.gp,$c.gs,$c.ic,$c.mh,$c.mm,$c.ms,$c.mq,$c.ex,$c.ec,$c.gc)
Write-Host ("data.test.js = {0} bytes" -f (Get-Item (Join-Path $planDir 'data.test.js')).Length)
$gondolas | ForEach-Object {
  '{0,-7} mods={1,3} prat={2,4} {3}' -f $_.id, $_.totalMods, $_.totalPrat, $_.preenchimento
}
