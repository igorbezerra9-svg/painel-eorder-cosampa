$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host "=== Painel eOrder — instalador da automação ===" -ForegroundColor Cyan
Write-Host "Pasta: $scriptDir"
Write-Host ""

# ── 1. Python ────────────────────────────────────────────────
$py = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $py) {
    Write-Host "❌ Python não encontrado no PATH." -ForegroundColor Red
    Write-Host "Instale o Python (python.org) marcando 'Add to PATH' na instalação, e rode este script de novo."
    Read-Host "Pressione Enter para fechar"
    exit 1
}
Write-Host "✔ Python encontrado: $py"

# ── 2. Chrome (checagem best-effort) ───────────────────────────
$programFilesX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
$chromePaths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$programFilesX86\Google\Chrome\Application\chrome.exe",
    "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
)
$chromeOk = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($chromeOk) {
    Write-Host "✔ Google Chrome encontrado: $chromeOk"
} else {
    Write-Host "⚠️  Não encontrei o Google Chrome nos caminhos usuais. Instale o Chrome antes de usar a automação." -ForegroundColor Yellow
}

# ── 3. Bibliotecas Python ──────────────────────────────────────
Write-Host ""
Write-Host "=== Instalando bibliotecas Python (selenium, openpyxl) ===" -ForegroundColor Cyan
& $py -m pip install --upgrade pip --quiet
& $py -m pip install selenium openpyxl --quiet
Write-Host "✔ Bibliotecas instaladas"

# ── 4. Credenciais ──────────────────────────────────────────────
Write-Host ""
Write-Host "=== Credenciais do eOrder ===" -ForegroundColor Cyan
$credFile = Join-Path $scriptDir "credenciais_eorder.json"
$credExample = Join-Path $scriptDir "credenciais_eorder.example.json"
if (-not (Test-Path $credFile)) {
    Copy-Item $credExample $credFile
    Write-Host "Criei credenciais_eorder.json a partir do exemplo." -ForegroundColor Yellow
    Write-Host "Vou abrir o arquivo no Bloco de Notas — preencha usuário/senha reais e o caminho da pasta de Downloads DESTA máquina, salve e feche."
    Start-Process notepad.exe $credFile -Wait
} else {
    Write-Host "✔ credenciais_eorder.json já existe — não mexi nele."
}

# ── 5. Tarefa agendada ──────────────────────────────────────────
Write-Host ""
Write-Host "=== Criando tarefa agendada (roda de hora em hora, 7h-17h) ===" -ForegroundColor Cyan
$pyDir = Split-Path $py
$pyw = Join-Path $pyDir "pythonw.exe"
if (-not (Test-Path $pyw)) { $pyw = $py }

$scriptPath = Join-Path $scriptDir "rodar_automatico.py"
$taskName = "PainelEOrder_Automatico"

$action = New-ScheduledTaskAction -Execute $pyw -Argument "`"$scriptPath`"" -WorkingDirectory $scriptDir
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00AM
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At 7:00AM -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 10)).Repetition
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes 25) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Roda o bot eOrder (Busca Execucao + TdC) de hora em hora (07h-17h), todo dia, e publica automaticamente no painel Supabase" `
    -Force | Out-Null

Write-Host "✔ Tarefa '$taskName' criada (dispara de hora em hora, 7h às 17h)"

Write-Host ""
Write-Host "=== Tudo pronto! ===" -ForegroundColor Green
Write-Host "Pra testar agora mesmo, rode no PowerShell:"
Write-Host "  Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor White
Write-Host "O log de cada rodada fica em: $scriptDir\automatico.log"
Read-Host "Pressione Enter para fechar"
