# weekly_report_task.ps1
# Executado pelo Task Scheduler toda sexta-feira às 00:30 BRT
# Gera o relatório semanal de reuniões e envia por email e Teams

$ScriptsDir = "C:\Users\sidne\.claude\skills\resumo-reunioes-teams\scripts"
$ReportFile = "C:\Users\sidne\AppData\Local\Temp\relatorio_semanal.html"
$LogFile    = "C:\Users\sidne\AppData\Local\Temp\weekly_report_task.log"

function Log($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp  $msg" | Tee-Object -FilePath $LogFile -Append
}

Log "=== Iniciando relatório semanal ==="

# Passo 1: gerar relatório HTML a partir dos .docx do OneDrive
Log "Gerando relatório..."
$result = python "$ScriptsDir\weekly_consolidate.py" --output $ReportFile 2>&1
Log $result

if (-not (Test-Path $ReportFile)) {
    Log "ERRO: relatório não gerado. Abortando envio."
    exit 1
}

Log "Relatório gerado: $ReportFile"

# Passo 2: enviar por email e Teams
Log "Enviando por email e Teams..."
$result2 = python "$ScriptsDir\send_report.py" `
    --email sidneydearaujosilva@gmail.com `
    --teams-team "Cowork Thunders" `
    --teams-channel "Notificações Importantes" `
    --report-file $ReportFile 2>&1
Log $result2

Log "=== Concluído ==="
