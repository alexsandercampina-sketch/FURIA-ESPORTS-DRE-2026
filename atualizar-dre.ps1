# ============================================================
# FURIA DRE - Script de Atualização Mensal
# Execute este arquivo toda vez que o Razão for atualizado.
# Lê todas as sheets do Excel que contenham dados de RAZÃO.
# Pré-requisito: Microsoft Excel instalado.
# ============================================================

param(
    [string]$ExcelPath = "C:\Users\PC\Downloads\Informacoes para Dashboard.xlsx"
)

$RepoPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputJson = Join-Path $RepoPath "data.json"

Write-Host ""
Write-Host "=== FURIA DRE - Atualizacao ===" -ForegroundColor Yellow
Write-Host "Excel: $ExcelPath"
Write-Host ""

if (-not (Test-Path $ExcelPath)) {
    Write-Host "ERRO: Arquivo Excel nao encontrado: $ExcelPath" -ForegroundColor Red
    Write-Host "Passe o caminho: .\atualizar-dre.ps1 -ExcelPath 'C:\caminho\arquivo.xlsx'"
    pause; exit 1
}

# ── Abre Excel ───────────────────────────────────────────────
Write-Host "Abrindo Excel..." -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$agg = @{}
$totalRows = 0
$skipped = 0

try {
    $wb = $xl.Workbooks.Open($ExcelPath, 0, $true)

    # ── Detecta sheets com dados de RAZÃO ───────────────────
    # Aceita qualquer sheet que tenha conteúdo nas colunas 1,4,7,8,11,12
    $sheetsToProcess = @()
    foreach ($ws in $wb.Sheets) {
        $testCell = $ws.Cells(2, 4).Text.Trim()  # coluna conta contábil, linha 2
        if ($testCell -match '^\d+\.\d+') {
            $sheetsToProcess += $ws
            Write-Host "  Sheet encontrada: '$($ws.Name)'" -ForegroundColor Green
        }
    }

    if ($sheetsToProcess.Count -eq 0) {
        throw "Nenhuma sheet com dados de RAZAO encontrada. Verifique a estrutura do Excel."
    }

    Write-Host ""

    foreach ($ws in $sheetsToProcess) {
        $lastRow = $ws.UsedRange.Rows.Count
        Write-Host "Lendo '$($ws.Name)': $lastRow linhas..." -ForegroundColor Cyan

        for ($i = 2; $i -le $lastRow; $i++) {
            $dateRaw = $ws.Cells($i, 1).Value2
            $accRaw  = $ws.Cells($i, 4).Text.Trim()
            $deb     = $ws.Cells($i, 7).Value2
            $cre     = $ws.Cells($i, 8).Value2
            $cc      = $ws.Cells($i, 11).Text.Trim()
            $filial  = $ws.Cells($i, 12).Text.Trim()

            if (-not $accRaw -or $accRaw -eq "") { $skipped++; continue }
            if ($null -eq $dateRaw)              { $skipped++; continue }

            # Usa o nome da sheet como filial se coluna 12 estiver vazia
            if (-not $filial -or $filial -eq "") { $filial = $ws.Name }

            # Converte data serial Excel → MM/YYYY
            try {
                $dt = [DateTime]::FromOADate($dateRaw)
                $month = $dt.ToString("MM/yyyy")
            } catch { $skipped++; continue }

            # Trunca código para 4 níveis (ex: 3.1.01.01.0001 → 3.1.01.01)
            $parts = $accRaw -split '\.'
            if ($parts.Count -lt 4) { $skipped++; continue }
            $acc4 = ($parts[0..3]) -join '.'

            if (-not $deb) { $deb = 0.0 }
            if (-not $cre) { $cre = 0.0 }

            # Chave agora inclui filial
            $key = "$month|$acc4|$cc|$filial"
            if ($agg.ContainsKey($key)) {
                $agg[$key].d  += $deb
                $agg[$key].cr += $cre
            } else {
                $agg[$key] = @{ m=$month; c=$acc4; cc=$cc; f=$filial; d=$deb; cr=$cre }
            }
            $totalRows++
        }
    }

    Write-Host ""
    Write-Host "Total linhas lidas: $totalRows | Ignoradas: $skipped | Agregados: $($agg.Count)" -ForegroundColor Green

    # ── Gera JSON ────────────────────────────────────────────
    Write-Host "Gerando data.json..." -ForegroundColor Cyan

    function EscapeJson([string]$s) { $s -replace '\\','\\' -replace '"','\"' }

    $records = $agg.Values | ForEach-Object {
        '{"f":"' + (EscapeJson $_.f) + '","cr":' + [math]::Round($_.cr,2) + ',"m":"' + $_.m + '","d":' + [math]::Round($_.d,2) + ',"c":"' + $_.c + '","cc":"' + (EscapeJson $_.cc) + '"}'
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
    Write-Host "Dashboard atualizado! Aguarde ~1 minuto e recarregue:" -ForegroundColor Green
    Write-Host "https://alexsandercampina-sketch.github.io/FURIA-ESPORTS-DRE-2026/" -ForegroundColor Cyan
} else {
    Write-Host "ERRO no git push. Verifique sua conexao e credenciais do GitHub." -ForegroundColor Red
}

Pop-Location
Write-Host ""
pause
