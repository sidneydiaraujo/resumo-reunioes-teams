# weekly_report_task.ps1
# Executado pelo Task Scheduler toda sexta-feira
# Gera o relatório semanal de reuniões e envia por email e Teams
#
# Configuração: edite skill_config.json na raiz da skill antes de agendar.

$SkillDir   = "$env:USERPROFILE\.claude\skills\resumo-reunioes-teams"
$ScriptsDir = "$SkillDir\scripts"
$ConfigFile = "$SkillDir\skill_config.json"
$ReportFile = "$env:TEMP\relatorio_semanal.html"
$LogFile    = "$env:TEMP\weekly_report_task.log"

function Log($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp  $msg" | Tee-Object -FilePath $LogFile -Append
}

Log "=== Iniciando relatório semanal ==="

# Ler configurações
if (-not (Test-Path $ConfigFile)) {
    Log "ERRO: skill_config.json não encontrado em $SkillDir. Crie o arquivo a partir de skill_config.example.json."
    exit 1
}
$config = Get-Content $ConfigFile | ConvertFrom-Json
$WorkEmail    = $config.work_email
$TeamsTeam    = $config.teams_team
$TeamsChannel = $config.teams_channel

if (-not $WorkEmail) {
    Log "ERRO: work_email não configurado em skill_config.json."
    exit 1
}

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
$sendArgs = @("$ScriptsDir\send_report.py", "--email", $WorkEmail, "--report-file", $ReportFile)
if ($TeamsTeam)    { $sendArgs += @("--teams-team", $TeamsTeam) }
if ($TeamsChannel) { $sendArgs += @("--teams-channel", $TeamsChannel) }

$result2 = python @sendArgs 2>&1
Log $result2

Log "=== Concluído ==="
