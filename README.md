# resumo-reunioes-teams

Skill do Claude Code que processa transcrições de reuniões do Microsoft Teams e gera um resumo estruturado salvo automaticamente no OneDrive.

---

## O que essa skill faz

1. Busca reuniões no seu calendário via Microsoft 365 (ou aceita arquivo `.vtt`/`.docx` exportado manualmente)
2. Identifica o **teor da reunião** (daily, diagnóstico, operacional, estratégico, alinhamento, review, planning, workshop) e estrutura o resumo de forma inteligente — sem template fixo
3. Rastreia **contextos entre reuniões**: problemas abertos numa daily aparecem como pendentes nas seguintes e são marcados como resolvidos quando encerrados
4. Gera um documento `.docx` formatado e salva no OneDrive em `Reuniões dos Times/<time>/`
5. Evita duplicidade: antes de processar, verifica se o resumo daquela reunião já existe no OneDrive

---

## Pré-requisitos

- [Claude Code](https://claude.ai/code) instalado
- Python 3.8 ou superior
- Integração **Microsoft 365** habilitada no Claude Code

> As dependências Python (`msal`, `requests`, `python-docx`) são instaladas automaticamente na primeira execução — não é necessário rodar `pip install` manualmente.

---

## Instalação

### 1. Clone o repositório na pasta de skills do Claude

**macOS / Linux:**
```bash
mkdir -p ~/.claude/skills
cd ~/.claude/skills
git clone https://github.com/sidneydiaraujo/resumo-reunioes-teams
```

**Windows (PowerShell):**
```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills"
Set-Location "$env:USERPROFILE\.claude\skills"
git clone https://github.com/sidneydiaraujo/resumo-reunioes-teams
```

### 2. Habilite a integração Microsoft 365 no Claude Code

No Claude Code, acesse as configurações de integrações e habilite **Microsoft 365**. Na primeira vez, o Claude solicitará autenticação com sua conta corporativa.

### 3. Reinicie o Claude Code

Feche e reabra o Claude Code para carregar a skill.

---

## Como usar

Basta pedir ao Claude Code:

```
"Resume as reuniões de hoje"
"Faz o resumo da reunião de alinhamento do time Projetos"
"Processa essa transcrição: C:\Users\...\reuniao.vtt"
"Salva o resumo da reunião de ontem no OneDrive"
```

O Claude vai:
1. Buscar a reunião no seu calendário (ou processar o arquivo fornecido)
2. Verificar se já existe resumo salvo para evitar duplicidade
3. Identificar o time se não estiver claro
4. Gerar o `.docx` com resumo estruturado conforme o teor da reunião
5. Salvar no OneDrive e retornar o link

---

## Filtros de reuniões

Na primeira execução, a skill pergunta como você quer filtrar as reuniões e salva a configuração em `skill_config.json` (arquivo local, não versionado). As opções incluem:

- **Anfitriões autorizados** — processar apenas reuniões criadas por pessoas específicas
- **Palavras-chave no título** — ignorar reuniões cujo título contenha determinados termos
- **Times com rotina própria** — ignorar dailys, reviews e plannings de times com seus próprios Scrum Masters

Consulte `skill_config.example.json` para ver a estrutura completa de configuração.

### Configuração pronta para uso

Se você trabalha em um contexto similar ao da Thunders (Scrum Master em empresa com múltiplos times), existe uma configuração personalizada pronta:

> **[resumo-reunioes-teams-thunders](https://github.com/sidneydiaraujo/resumo-reunioes-teams-thunders)** — filtros por anfitrião autorizado, cerimônias ignoradas por time e títulos específicos.

---

## Estrutura do projeto

```
resumo-reunioes-teams/
├── SKILL.md                        # Definição da skill (lida pelo Claude)
├── skill_config.example.json       # Modelo de configuração de filtros
├── scripts/
│   ├── create_docx_smart.py        # Gera .docx no formato inteligente (padrão)
│   ├── create_docx_diagnostico.py  # Gera .docx no formato diagnóstico
│   ├── create_docx.py              # Gera .docx no formato legado
│   ├── onedrive_upload.py          # Faz upload para o OneDrive
│   ├── parse_transcript.py         # Processa arquivos .vtt e .docx locais
│   ├── send_report.py              # Envia relatório semanal consolidado
│   ├── weekly_consolidate.py       # Consolida reuniões da semana em JSON/HTML
│   └── weekly_report_task.ps1      # Script para agendamento semanal automático
└── README.md
```

---

## Integração com reunioes-consulta

Após salvar resumos, você pode consultar o conteúdo das reuniões em linguagem natural ("o que foi decidido sobre X?", "quais são os próximos passos do time B2B?"). Isso usa a skill complementar **[reunioes-consulta](https://github.com/sidneydiaraujo/reunioes-consulta)** — instale-a também para ter o fluxo completo.
