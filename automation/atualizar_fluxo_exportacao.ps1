<#
Atualizacao diaria do painel de Fluxo de Exportacao: extrai as abas de programacao (Excel COM),
recalcula pedidos/SKUs/resumo, detecta o que mudou desde a ultima rodada (vendedor alterou
quantidade ou incluiu item sem avisar) e gera o HTML final autocontido.

Tarefa agendada separada de AtualizarPainelEstoques (projeto painel-estoques) de proposito —
esse painel nao deve arriscar a atualizacao do outro. Mesmo padrao de seguranca do Excel COM
(mata sobras invisiveis, aborta se houver janela real aberta) e de log (automation\logs).
#>

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$nodeExe = "C:\Program Files\nodejs\node.exe"
$logDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir ("atualizacao-{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line
}

function Invoke-Step {
    param([string]$Name, [scriptblock]$Action)
    Write-Log "INICIO: $Name"
    $start = Get-Date
    try {
        & $Action
        $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
        Write-Log "OK: $Name (${elapsed}s)"
    }
    catch {
        $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
        Write-Log "FALHA: $Name (${elapsed}s) - $($_.Exception.Message)"
        Write-Log "ROTINA INTERROMPIDA. Painel anterior preservado; nada parcial foi publicado."
        exit 1
    }
}

# Mata sobras invisiveis do Excel (sem janela real) e diz se e seguro seguir. Retorna $false
# (sem matar nada) se houver uma janela REAL aberta - nunca fecha trabalho de verdade do usuario.
function Clear-StrayExcel {
    $excelProcs = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($excelProcs.Count -eq 0) { return $true }
    $realWindows = @($excelProcs | Where-Object { $_.MainWindowTitle })
    if ($realWindows.Count -gt 0) {
        $titles = ($realWindows | ForEach-Object { $_.MainWindowTitle }) -join ", "
        Write-Log "FALHA: ha janela(s) real(is) do Excel aberta(s) ($titles). Feche antes de rodar a atualizacao."
        return $false
    }
    foreach ($proc in $excelProcs) {
        Write-Log "Limpando instancia invisivel de Excel sobrada (PID $($proc.Id))"
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
    return $true
}

# A extracao (Excel COM, ~3-6min por rodada com milhares de celulas lidas uma a uma) e o passo
# mais sujeito a falha transitoria - ja vimos RPC_E_DISCONNECTED (18/08/2026) sem causa clara
# (Excel travou/perdeu conexao no meio do caminho). Como o build inteiro depende só de arquivo
# local depois disso, vale tentar de novo automaticamente em vez de deixar o painel desatualizado
# até alguém notar. Só desiste de vez se todas as tentativas falharem, ou se aparecer uma janela
# REAL do Excel no meio do caminho (nesse caso é o usuário trabalhando, não mexe).
function Invoke-StepWithRetry {
    param([string]$Name, [scriptblock]$Action, [int]$MaxAttempts = 3)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if (-not (Clear-StrayExcel)) { exit 1 }
        Write-Log "INICIO: $Name (tentativa $attempt de $MaxAttempts)"
        $start = Get-Date
        try {
            & $Action
            $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
            Write-Log "OK: $Name (${elapsed}s)"
            return
        }
        catch {
            $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
            Write-Log "FALHA: $Name na tentativa $attempt (${elapsed}s) - $($_.Exception.Message)"
            if ($attempt -eq $MaxAttempts) {
                Write-Log "ROTINA INTERROMPIDA apos $MaxAttempts tentativas. Painel anterior preservado; nada parcial foi publicado."
                exit 1
            }
            Write-Log "Tentando de novo em 15s..."
            Start-Sleep -Seconds 15
        }
    }
}

if (-not (Test-Path -LiteralPath $nodeExe)) {
    Write-Log "FALHA: Node nao encontrado em $nodeExe"
    exit 1
}

Write-Log "=== Atualizacao do Fluxo de Exportacao iniciada ==="

Invoke-StepWithRetry "Extrair abas de programacao de exportacao (Excel COM)" {
    # Rodado como Job com timeout de propósito (21/08/2026): já vimos essa extração travar de
    # verdade por mais de 1h sem lançar erro nenhum (Excel preso, sem RPC_E_DISCONNECTED nem
    # nada pro catch pegar) - nesse caso o retry de 3 tentativas nunca entrava em ação, porque
    # só reage a excecao, nao a travamento silencioso. O timeout forca uma excecao mesmo quando
    # o processo trava sem erro, pra dar chance de tentar de novo em vez de ficar parado pra sempre.
    $extractScript = Join-Path $PSScriptRoot "extract_export_data.ps1"
    $job = Start-Job -ScriptBlock {
        param($scriptPath)
        $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath 2>&1
        [PSCustomObject]@{ Output = $out; ExitCode = $LASTEXITCODE }
    } -ArgumentList $extractScript

    $completed = Wait-Job $job -Timeout 900
    if (-not $completed) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Clear-StrayExcel | Out-Null
        throw "extract_export_data.ps1 nao terminou em 15 minutos (travado) - processo interrompido"
    }

    $jobResult = Receive-Job $job
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    if ($jobResult.ExitCode -ne 0) { throw "extract_export_data.ps1 saiu com codigo $($jobResult.ExitCode)`: $($jobResult.Output -join ' | ')" }
    Write-Log ($jobResult.Output -join " ")
}

Invoke-Step "Gerar pedidos/SKUs/resumo, detectar alteracoes e montar o painel (Node)" {
    Push-Location $projectRoot
    try {
        $result = & $nodeExe (Join-Path $PSScriptRoot "build.js") 2>&1
        if ($LASTEXITCODE -ne 0) { throw "build.js saiu com codigo $LASTEXITCODE`: $($result -join ' | ')" }
        Write-Log ($result -join " ")
    }
    finally { Pop-Location }
}

# Envio pro GitHub tratado como nao-fatal de proposito: nessa altura o painel local ja foi
# atualizado com sucesso (unica coisa que realmente importa se a internet cair ou o push falhar
# por qualquer motivo transitorio) - um problema so no envio nao deveria fazer a rotina inteira
# "falhar" quando o trabalho de verdade (extracao + build) ja terminou bem.
Write-Log "INICIO: Enviar painel atualizado pro GitHub (privado)"
Push-Location $projectRoot
try {
    # De proposito SEM "2>&1" em nenhuma chamada de git aqui: no PowerShell 5.1, redirecionar o
    # stderr de um executavel nativo pra dentro do pipeline embrulha cada linha (mesmo avisos
    # inofensivos, como o de conversao de fim de linha CRLF/LF) num ErrorRecord - com
    # $ErrorActionPreference="Stop" (topo do script) isso derruba o try/catch como se fosse erro
    # de verdade, mesmo quando o git terminou com sucesso (exit code 0). Ja aconteceu de verdade
    # em 24/08/2026: o aviso de CRLF interrompeu o envio bem antes do "git push" rodar.
    & git add -A
    $statusOutput = & git status --porcelain
    if (-not $statusOutput) {
        Write-Log "OK: nada mudou no painel - nada pra enviar"
    }
    else {
        $commitMsg = "Atualizacao automatica {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm")
        & git commit -m $commitMsg | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Log "AVISO: git commit falhou (painel local esta ok, so nao sincronizou com o GitHub agora)"
        }
        else {
            & git push origin main
            if ($LASTEXITCODE -ne 0) {
                Write-Log "AVISO: git push falhou (painel local esta ok, so nao sincronizou com o GitHub agora)"
            }
            else {
                Write-Log "OK: Enviado pro GitHub - $commitMsg"
            }
        }
    }
}
catch {
    Write-Log "AVISO: erro ao enviar pro GitHub (painel local esta ok) - $($_.Exception.Message)"
}
finally { Pop-Location }

Write-Log "=== Atualizacao do Fluxo de Exportacao concluida com sucesso ==="
