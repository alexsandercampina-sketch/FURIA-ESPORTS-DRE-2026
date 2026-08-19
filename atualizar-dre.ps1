# ============================================================
# FURIA DRE - Script de Atualização Mensal
# Execute sempre que o Excel for atualizado.
# Pré-requisito: Microsoft Excel instalado.
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
    Write-Host "Uso: .\atualizar-dre.ps1 -ExcelPath 'C:\caminho\arquivo.xlsx'"
    pause; exit 1
}

# ── Detecta data a partir do valor da célula (serial ou texto) ──
function ParseDate([object]$val, [string]$txt) {
    if ($null -ne $val -and $val -is [double]) {
        try { return [DateTime]::FromOADate($val) } catch {}
    }
    if ($txt -match '^(\d{2})/(\d{2})/(\d{4})$') {
        try { return [DateTime]::ParseExact($txt, 'dd/MM/yyyy', $null) } catch {}
    }
    return $null
}

# ── Escapa aspas para JSON ───────────────────────────────────
function EscJson([string]$s) { $s -replace '\\','\\' -replace '"','\"' }

# ── Sheets em USD: valores convertidos para BRL via PTAX venda 31/07/2026 ──
# Adicione aqui o nome exato de cada sheet em dólar.
$USD_SHEETS = @('FURIAGG', 'FURIA LLC')
$PTAX_VENDA = 5.0773   # PTAX venda BACEN 31/07/2026

# ── Abre Excel ───────────────────────────────────────────────
Write-Host "Abrindo Excel..." -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.Visible        = $false
$xl.DisplayAlerts  = $false

$agg       = @{}
$totalRows = 0
$skipped   = 0

try {
    $wb = $xl.Workbooks.Open($ExcelPath, 0, $true)

    # ── Detecta sheets com dados contábeis ───────────────────
    # Ignora MODELO DRE e sheets vazias; aceita qualquer sheet
    # onde coluna 4 tenha um código no formato X.X.XX.XX em alguma das primeiras 30 linhas.
    $sheetsToProcess = @()
    foreach ($ws in $wb.Sheets) {
        if ($ws.Name -match 'MODELO|model') { continue }
        $found = $false
        $maxCheck = [Math]::Min(30, $ws.UsedRange.Rows.Count)
        for ($r = 2; $r -le $maxCheck; $r++) {
            if ($ws.Cells($r, 4).Text.Trim() -match '^\d+\.\d+') { $found = $true; break }
        }
        if ($found) {
            $sheetsToProcess += $ws
            Write-Host "  Encontrada: '$($ws.Name)'" -ForegroundColor Green
        }
    }

    if ($sheetsToProcess.Count -eq 0) {
        throw "Nenhuma sheet com dados contabeis encontrada. Coluna 4 deve ter codigo no formato 3.1.01.01.0001"
    }

    Write-Host ""

    foreach ($ws in $sheetsToProcess) {
        $lastRow   = $ws.UsedRange.Rows.Count
        $isUSD     = $USD_SHEETS -contains $ws.Name
        $fxRate    = if ($isUSD) { $PTAX_VENDA } else { 1.0 }
        $currency  = if ($isUSD) { "USD→BRL ($PTAX_VENDA)" } else { "BRL" }
        Write-Host "Lendo '$($ws.Name)': $lastRow linhas [$currency]..." -ForegroundColor Cyan
        $sheetRows = 0

        for ($i = 2; $i -le $lastRow; $i++) {
            # Conta contábil obrigatória na coluna 4
            $accRaw = $ws.Cells($i, 4).Text.Trim()
            if (-not $accRaw -or $accRaw -notmatch '^\d+\.\d+') { $skipped++; continue }

            # Data (col 1) — aceita serial Excel ou texto DD/MM/YYYY
            $dt = ParseDate $ws.Cells($i, 1).Value2 $ws.Cells($i, 1).Text.Trim()
            if (-not $dt) { $skipped++; continue }
            $month = $dt.ToString("MM/yyyy")

            # Débito / Crédito — usa Value2 para pegar negativos corretamente
            $deb = $ws.Cells($i, 7).Value2; if (-not $deb) { $deb = 0.0 }
            $cre = $ws.Cells($i, 8).Value2; if (-not $cre) { $cre = 0.0 }
            if ($deb -eq 0 -and $cre -eq 0) { $skipped++; continue }

            # Converte USD → BRL se necessário (PTAX venda 31/07/2026)
            $deb = $deb * $fxRate
            $cre = $cre * $fxRate

            # Centro de custo (col 11) e filial (col 12)
            $cc     = $ws.Cells($i, 11).Text.Trim()
            $filial = $ws.Cells($i, 12).Text.Trim()
            if (-not $filial) { $filial = $ws.Name }   # fallback: nome da sheet
            if (-not $cc)     { $cc = "Sem CC" }

            # Trunca código para 4 níveis: 3.1.01.01.0001 → 3.1.01.01
            $parts = $accRaw -split '\.'
            if ($parts.Count -lt 4) { $skipped++; continue }
            $acc4 = ($parts[0..3]) -join '.'

            $key = "$month|$acc4|$cc|$filial"
            if ($agg.ContainsKey($key)) {
                $agg[$key].d  += $deb
                $agg[$key].cr += $cre
            } else {
                $agg[$key] = @{ m=$month; c=$acc4; cc=$cc; f=$filial; d=$deb; cr=$cre }
            }
            $totalRows++; $sheetRows++
        }
        Write-Host "  -> $sheetRows linhas lidas" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "Total: $totalRows linhas | Ignoradas: $skipped | Registros: $($agg.Count)" -ForegroundColor Green

    # ── Gera JSON ────────────────────────────────────────────
    Write-Host "Gerando data.json..." -ForegroundColor Cyan
    $records = $agg.Values | ForEach-Object {
        '{"f":"'+(EscJson $_.f)+'","cr":'+[math]::Round($_.cr,2)+',"m":"'+$_.m+'","d":'+[math]::Round($_.d,2)+',"c":"'+$_.c+'","cc":"'+(EscJson $_.cc)+'"}'
    }
    $json = '[' + ($records -join ',') + ']'
    [System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.Encoding]::UTF8)

    $kb = [math]::Round((Get-Item $OutputJson).Length / 1KB, 1)
    Write-Host "data.json salvo: $kb KB" -ForegroundColor Green

} finally {
    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

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
