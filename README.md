# resumo-reunioes-teams

Skill do Claude Code que processa transcrições de reuniões do Microsoft Teams e gera um resumo estruturado salvo automaticamente no OneDrive.

---

## O que essa skill faz

1. Busca transcrições de reuniões diretamente via Microsoft 365 (ou aceita arquivo .vtt/.docx exportado)
2. Analisa o conteúdo e extrai automaticamente:
   - Regras de negócio discutidas
   - Combinados e responsáveis
   - Próximos passos com prazos
   - Decisões tomadas
   - Riscos e impedimentos
   - Pendências em aberto
3. Gera um documento `.docx` formatado
4. Salva no OneDrive em `Reuniões dos Times/<nome-do-time>/`

---

## Pré-requisitos

- [Claude Code](https://claude.ai/code) instalado
- Python 3.8 ou superior
- Integração **Microsoft 365** habilitada no Claude Code (para buscar reuniões e salvar no OneDrive)
- Bibliotecas Python:
  ```bash
  pip install python-docx requests
  ```

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
cd "$env:USERPROFILE\.claude\skills"
git clone https://github.com/sidneydiaraujo/resumo-reunioes-teams
```

### 2. Instale as dependências Python

```bash
pip install python-docx requests
```

### 3. Habilite a integração Microsoft 365 no Claude Code

Esta skill usa o MCP do Microsoft 365 para acessar calendário, transcrições de reuniões e OneDrive.

No Claude Code, acesse as configurações de integrações e habilite **Microsoft 365**. Na primeira vez que a skill precisar acessar sua conta, o Claude vai solicitar autenticação.

### 4. Reinicie o Claude Code

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
2. Extrair o conteúdo da transcrição
3. Perguntar o time (B2B, Projetos, etc.) se não estiver claro
4. Gerar o documento `.docx` com o resumo estruturado
5. Salvar no OneDrive e retornar o link

---

## Filtros automáticos

Na primeira execução, a skill pergunta como você quer filtrar as reuniões e salva a configuração em `skill_config.json`. As opções incluem:

- **Anfitriões autorizados** — processar apenas reuniões criadas por pessoas específicas
- **Palavras-chave no título** — ignorar reuniões cujo título contenha determinados termos
- **Times com rotina própria** — ignorar dailys, reviews e plannings de times com seus próprios Scrum Masters

Consulte `skill_config.example.json` para ver a estrutura completa.

### Configuração pronta para uso

Se você trabalha em um contexto similar ao da Thunders (Scrum Master em empresa com múltiplos times), existe uma configuração personalizada pronta:

> **[resumo-reunioes-teams-thunders](https://github.com/sidneydiaraujo/resumo-reunioes-teams-thunders)** — filtros por anfitrião autorizado, cerimônias ignoradas por time e títulos específicos.

---

## Estrutura do projeto

```
resumo-reunioes-teams/
├── SKILL.md                     # Definição da skill (lida pelo Claude)
├── scripts/
│   ├── create_docx.py           # Gera o arquivo .docx a partir do JSON
│   ├── onedrive_upload.py       # Faz upload para o OneDrive
│   ├── parse_transcript.py      # Processa arquivos .vtt e .docx locais
│   ├── send_report.py           # Envia relatório consolidado
│   └── weekly_consolidate.py   # Consolida reuniões da semana
└── README.md
```

---

## Integração com reunioes-consulta

Após salvar um resumo, o Claude vai oferecer consultar informações das reuniões salvas. Isso usa a skill **reunioes-consulta** — instale-a também para ter o fluxo completo.
