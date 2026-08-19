# ============================================================
# FURIA DRE - Script de Atualizacao Mensal
# Execute sempre que o Excel for atualizado.
# Pre-requisito: Microsoft Excel instalado + internet para PTAX BACEN.
# ============================================================

param(
    [string]$ExcelPath = "C:\Users\PC\Downloads\Informacoes para Dashboard.xlsx"
)

$RepoPath   = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputJson = Join-Path $RepoPath "data.json"

Write-Host ""
Write-Host "=== FURIA DRE - Atualizacao ===" -ForegroundColor Yellow
Write-Host "Excel: $ExcelPath"
Write-Host ""

if (-not (Test-Path $ExcelPath)) {
    Write-Host "ERRO: Arquivo nao encontrado: $ExcelPath" -ForegroundColor Red
    pause; exit 1
}

# Sheets cujos valores estao em USD (serao convertidas via PTAX automatico)
$USD_SHEETS = @('FURIAGG', 'FURIA LLC', 'GGCORP', 'FURIA GG CORP')

# ── Funcoes auxiliares ───────────────────────────────────────

function ParseDate([object]$val, [string]$txt) {
    if ($null -ne $val -and $val -is [double]) {
        try { return [DateTime]::FromOADate($val) } catch {}
    }
    if ($txt -match '^\d\d/\d\d/\d\d\d\d$') {
        try { return [DateTime]::ParseExact($txt, 'dd/MM/yyyy', $null) } catch {}
    }
    return $null
}

function ParseNumeric([object]$val) {
    if ($null -eq $val) { return 0.0 }
    if ($val -is [double]) { return $val }
    if ($val -is [int] -or $val -is [long]) { return [double]$val }
    $s = $val.ToString().Trim() -replace '\$', '' -replace ' ', ''
    if ($s -eq '' -or $s -eq '-') { return 0.0 }
    $hasDot   = $s.Contains('.')
    $hasComma = $s.Contains(',')
    if ($hasDot -and $hasComma) {
        if ($s.LastIndexOf('.') -gt $s.LastIndexOf(',')) {
            $s = $s -replace ',', ''                         # US: 138,888.89
        } else {
            $s = ($s -replace '\.', '') -replace ',', '.'   # BR: 138.888,89
        }
    } elseif ($hasComma) {
        $s = $s -replace ',', '.'
    }
    try { return [double]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture) } catch { return 0.0 }
}

function EscJson([string]$s) { $s -replace '\\', '\\' -replace '"', '\"' }

# Busca PTAX venda BACEN para o ultimo dia util do mes MM/yyyy
function GetPTAX([string]$month) {
    $parts   = $month -split '/'
    $mo      = [int]$parts[0]
    $yr      = [int]$parts[1]
    $lastDay = [DateTime]::new($yr, $mo, [DateTime]::DaysInMonth($yr, $mo))

    # Tenta os ultimos 5 dias (ignora fins de semana e feriados sem cotacao)
    for ($d = 0; $d -lt 5; $d++) {
        $date    = $lastDay.AddDays(-$d)
        $dateStr = $date.ToString("MM-dd-yyyy")
        $url     = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='$dateStr'&`$top=1&`$format=json&`$select=cotacaoVenda"
        try {
            $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15
            $json = $resp.Content | ConvertFrom-Json
            if ($json.value -and $json.value.Count -gt 0) {
                $rate = $json.value[0].cotacaoVenda
                Write-Host "  PTAX venda $($date.ToString('dd/MM/yyyy')): R$ $rate" -ForegroundColor Green
                return $rate
            }
        } catch {}
    }
    throw "Nao foi possivel obter PTAX do BACEN para $month. Verifique a conexao com a internet."
}

# ── Abre Excel ───────────────────────────────────────────────
Write-Host "Abrindo Excel..." -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.Visible       = $false
$xl.DisplayAlerts = $false

$aggBRL    = @{}   # registros em BRL (direto)
$aggUSD    = @{}   # registros em USD bruto (conversao aplicada depois)
$totalRows = 0
$skipped   = 0

try {
    $wb = $xl.Workbooks.Open($ExcelPath, 0, $true)

    # Detecta sheets com dados contabeis (col 4 com codigo contabil nas primeiras 30 linhas)
    $sheetsToProcess = @()
    foreach ($ws in $wb.Sheets) {
        if ($ws.Name -match 'MODELO') { continue }
        $found    = $false
        $maxCheck = [Math]::Min(30, $ws.UsedRange.Rows.Count)
        for ($r = 2; $r -le $maxCheck; $r++) {
            if ($ws.Cells($r, 4).Text.Trim() -match '^\d+\.\d+') { $found = $true; break }
        }
        if ($found) {
            $sheetsToProcess += $ws
            $tipo = if ($USD_SHEETS -contains $ws.Name) { "USD" } else { "BRL" }
            Write-Host "  Encontrada: '$($ws.Name)' [$tipo]" -ForegroundColor Green
        }
    }

    if ($sheetsToProcess.Count -eq 0) {
        throw "Nenhuma sheet com dados contabeis encontrada."
    }

    Write-Host ""

    # ── Leitura ──────────────────────────────────────────────
    foreach ($ws in $sheetsToProcess) {
        $lastRow = $ws.UsedRange.Rows.Count
        $isUSD   = $USD_SHEETS -contains $ws.Name
        $label   = if ($isUSD) { "USD" } else { "BRL" }
        Write-Host "Lendo '$($ws.Name)': $lastRow linhas [$label]..." -ForegroundColor Cyan
        $sheetRows = 0

        for ($i = 2; $i -le $lastRow; $i++) {
            $accRaw = $ws.Cells($i, 4).Text.Trim()
            if (-not $accRaw -or $accRaw -notmatch '^\d+\.\d+') { $skipped++; continue }

            $dt = ParseDate $ws.Cells($i, 1).Value2 $ws.Cells($i, 1).Text.Trim()
            if (-not $dt) { $skipped++; continue }
            $month = $dt.ToString("MM/yyyy")

            $deb = ParseNumeric $ws.Cells($i, 7).Value2
            $cre = ParseNumeric $ws.Cells($i, 8).Value2
            if ($deb -eq 0 -and $cre -eq 0) { $skipped++; continue }

            $cc     = $ws.Cells($i, 11).Text.Trim()
            $filial = $ws.Cells($i, 12).Text.Trim()
            if (-not $filial) { $filial = $ws.Name }
            if (-not $cc)     { $cc = "Sem CC" }

            $parts = $accRaw -split '\.'
            if ($parts.Count -lt 4) { $skipped++; continue }
            $acc4 = ($parts[0..3]) -join '.'

            $key = "$month|$acc4|$cc|$filial"
            $target = if ($isUSD) { $aggUSD } else { $aggBRL }

            if ($target.ContainsKey($key)) {
                $target[$key].d  += $deb
                $target[$key].cr += $cre
            } else {
                $target[$key] = @{ m=$month; c=$acc4; cc=$cc; f=$filial; d=$deb; cr=$cre }
            }
            $totalRows++; $sheetRows++
        }
        Write-Host "  -> $sheetRows linhas lidas" -ForegroundColor Gray
    }

} finally {
    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

Write-Host ""
Write-Host "Total: $totalRows linhas | Ignoradas: $skipped" -ForegroundColor Green
Write-Host "BRL: $($aggBRL.Count) registros | USD (bruto): $($aggUSD.Count) registros" -ForegroundColor Green

# ── PTAX automatico: ultimo mes do razao ─────────────────────
if ($aggUSD.Count -gt 0) {
    # Encontra o ultimo mes presente em TODOS os dados
    $allMonths = @()
    $aggBRL.Values | ForEach-Object { $allMonths += $_.m }
    $aggUSD.Values | ForEach-Object { $allMonths += $_.m }
    $lastMonth = $allMonths | Sort-Object {
        $p = $_ -split '/'; [int]$p[1] * 100 + [int]$p[0]
    } | Select-Object -Last 1

    Write-Host ""
    Write-Host "Ultimo mes detectado: $lastMonth" -ForegroundColor Cyan
    Write-Host "Buscando PTAX BACEN..." -ForegroundColor Cyan
    $ptax = GetPTAX $lastMonth

    # Converte USD -> BRL e mescla com aggBRL
    foreach ($key in $aggUSD.Keys) {
        $rec = $aggUSD[$key]
        $dBRL  = $rec.d  * $ptax
        $crBRL = $rec.cr * $ptax
        if ($aggBRL.ContainsKey($key)) {
            $aggBRL[$key].d  += $dBRL
            $aggBRL[$key].cr += $crBRL
        } else {
            $aggBRL[$key] = @{ m=$rec.m; c=$rec.c; cc=$rec.cc; f=$rec.f; d=$dBRL; cr=$crBRL }
        }
    }
    Write-Host "Conversao aplicada: 1 USD = R$ $ptax" -ForegroundColor Green
}

# ── Gera JSON ────────────────────────────────────────────────
Write-Host ""
Write-Host "Gerando data.json ($($aggBRL.Count) registros)..." -ForegroundColor Cyan
$records = $aggBRL.Values | ForEach-Object {
    '{"f":"' + (EscJson $_.f) + '","cr":' + [math]::Round($_.cr, 2) + ',"m":"' + $_.m + '","d":' + [math]::Round($_.d, 2) + ',"c":"' + $_.c + '","cc":"' + (EscJson $_.cc) + '"}'
}
$json = '[' + ($records -join ',') + ']'
[System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.Encoding]::UTF8)

$kb = [math]::Round((Get-Item $OutputJson).Length / 1KB, 1)
Write-Host "data.json salvo: $kb KB" -ForegroundColor Green

# ── Git push ─────────────────────────────────────────────────
Write-Host ""
Write-Host "Enviando para GitHub..." -ForegroundColor Cyan
Push-Location $RepoPath

$today = Get-Date -Format "dd/MM/yyyy"
git add data.json
git commit -m "Atualizar dados DRE - $today"
git push

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "=== CONCLUIDO ===" -ForegroundColor Green
    Write-Host "Dashboard atualizado! Aguarde ~1 min e recarregue:" -ForegroundColor Green
    Write-Host "https://alexsandercampina-sketch.github.io/FURIA-ESPORTS-DRE-2026/" -ForegroundColor Cyan
} else {
    Write-Host "ERRO no git push. Verifique conexao e credenciais do GitHub." -ForegroundColor Red
}

Pop-Location
Write-Host ""
pause
