---
name: resumo-reunioes-teams
description: Processa transcrições de reuniões do Microsoft Teams e gera um resumo estruturado salvo automaticamente no OneDrive do usuário. Use esta skill SEMPRE que o usuário mencionar: resumir reunião, processar transcrição, arquivo de reunião, VTT do Teams, exportar reunião, salvar resumo, transcrição Teams. Acione também para frases como "faz o resumo dessa reunião", "resume as reuniões de hoje", "processa essa transcrição", "salva essa reunião no OneDrive", "quero resumir minha reunião", ou quando o usuário fornecer o caminho de um arquivo .vtt ou .docx de transcrição do Teams.
---

# Resumo de Reuniões Teams

Processa transcrições do Teams e gera resumo estruturado salvo no OneDrive.

---

## Configuração de filtros

Os filtros determinam quais reuniões do calendário serão processadas. São definidos pelo usuário na primeira execução e salvos em `skill_config.json`.

### Primeira execução (sem config ou sem a chave `filters`)

Pergunte ao usuário:

> "Como você quer filtrar as reuniões para resumo?"
> 1. **Processar tudo** — todas as reuniões com transcrição disponível
> 2. **Configurar filtros** — escolher padrões a ignorar

**Se escolher "Processar tudo":** salve `"filters": { "mode": "all" }` e siga em frente.

**Se escolher "Configurar filtros":** faça as três perguntas abaixo, salve em `skill_config.json` e use nos processamentos seguintes.

---

#### Pergunta 1 — Palavras-chave no título

> "Tem alguma palavra-chave no título que deve fazer a reunião ser **sempre ignorada**?"  
> *(ex: `[Refinamento]`, `Standup`, `Diálogo de Inovação` — deixe em branco para não filtrar por título)*

Salva em: `filters.ignore_title_contains` (lista de strings)

#### Pergunta 2 — Recorrentes de outros organizadores

> "Reuniões **recorrentes** onde você **não é o organizador** (dailys de outros times, standups, etc.) devem ser ignoradas?"  
> *(sim/não)*

Salva em: `filters.ignore_recurring_not_organizer` (true/false)

#### Pergunta 3 — Times ou grupos com rotina irrelevante

> "Tem algum **time ou grupo** cujas reuniões de rotina (daily, review, planning, refinamento) você **não quer resumir**?"  
> Para cada time informado, pergunte quais tipos de reunião ignorar.  
> *(deixe em branco para não filtrar por time)*

Salva em: `filters.ignore_team_patterns` — lista de objetos `{ "team_keywords": [...], "meeting_types": [...] }`

---

### Aplicar os filtros a cada evento do calendário

```
1. filters.mode == "all"?                                          → processar
2. Título contém algum item de filters.ignore_title_contains?      → ignorar
3. ignore_recurring_not_organizer: true
   E recurrence != null  E isOrganizer: false?                     → ignorar
4. Para cada padrão em filters.ignore_team_patterns:
     organizador ou título contém team_keywords
     E tipo da reunião está em meeting_types?                       → ignorar
5. Qualquer outro caso                                             → processar
```

> **Nota:** se `skill_config.json` já existir com a chave `filters` configurada, pule o wizard e aplique os filtros diretamente.

---

## Fluxo de execução

### 1. Identificar a fonte da transcrição

Há dois modos de obter a transcrição — prefira sempre o Modo A:

**Modo A — Direto pelo MCP (sem arquivo)** ← preferencial
O usuário pode pedir "resume as reuniões de hoje" ou "faz o resumo da reunião X" sem fornecer nenhum arquivo.

1. Busque as reuniões no calendário com `mcp__claude_ai_Microsoft_365__outlook_calendar_search`
2. Leia o evento completo com `mcp__claude_ai_Microsoft_365__read_resource` usando o URI do evento
3. Use o campo `meetingTranscriptUrl` do evento — esse é o URI da transcrição no formato:
   `meeting-transcript:///events/{base64encodedJoinWebUrl}`
4. Leia a transcrição com `mcp__claude_ai_Microsoft_365__read_resource` usando esse URI

**Limitações conhecidas do Modo A:**
- Reuniões recorrentes diárias (Daily, Stand-up) normalmente não têm transcrição disponível via API
- Reuniões onde o usuário não é organizador podem retornar 403 — sem acesso à transcrição
- Reuniões sem gravação ativada não têm transcrição
- Se a transcrição retornar `"content": ""`, não há transcrição disponível para aquela ocorrência

Se a reunião não tiver transcrição disponível, informe o usuário e ofereça o Modo B.

**Modo B — Arquivo local (VTT ou DOCX)**
Pergunte ao usuário o caminho do arquivo de transcrição exportado manualmente do Teams. Aceite `.vtt` ou `.docx`.

---

### 2. Extrair o texto da transcrição

**Via MCP (Modo A):** o conteúdo já vem estruturado no campo `content` da resposta. Se for muito grande (>200KB), spawne um subagente para processar e retornar apenas o JSON de resumo.

**VTT (Modo B):** leia o arquivo com a ferramenta Read. Formato:
```
00:01:23.000 --> 00:01:28.000
<v João Silva>Precisamos definir o fluxo de aprovação.</v>
```
Extraia nomes dos speakers das tags `<v Nome>` ou do padrão `Nome: texto`.

**DOCX (Modo B):** execute:
```powershell
python "$env:USERPROFILE\.claude\skills\resumo-reunioes-teams\scripts\parse_transcript.py" "<caminho_arquivo>"
```
Retorna JSON com texto estruturado.

---

### 3. Identificar o time

Se o usuário não informar o time, analise o conteúdo e pergunte:
> "Esse resumo é de qual time ou área?"

Para um novo time, a pasta será criada automaticamente no OneDrive durante o upload.

---

### 4. Gerar o resumo (sua principal responsabilidade)

Leia o conteúdo da transcrição e **entenda o teor da reunião antes de estruturar qualquer coisa**. Não aplique formato fixo — escolha as seções que fazem sentido para o que foi discutido.

#### Como identificar o teor

Pergunte-se: *"O que as pessoas estavam tentando resolver ou decidir nessa reunião?"*

| Teor | O que buscar | Seções naturais |
|---|---|---|
| **Diagnóstico** | debate sobre o que funciona/não funciona, saúde do processo | O que funciona, O que não funciona, Encaminhamentos |
| **Operacional** | tarefas, prazos, bloqueios, quem faz o quê | Combinados, Impedimentos, Próximos Passos |
| **Estratégico** | direção, prioridades, posicionamento, decisões de produto | Contexto, Decisões, Implicações, Próximos Passos |
| **Alinhamento** | dúvidas resolvidas, acordos entre times ou pessoas | Pontos Alinhados, Decisões, Combinados |
| **Review / Demo** | apresentação de entregas, feedback, aceite | O que foi entregue, Feedbacks, Pendências de aceite |
| **Planning** | histórias, estimativas, capacidade, metas de sprint | Meta, Histórias comprometidas, Riscos, Capacidade |
| **Workshop / Levantamento** | regras de negócio, fluxos, requisitos | Regras de Negócio, Fluxos, Dúvidas em aberto |

Se a reunião mistura dois teores, priorize o que gerou mais discussão.

#### JSON de saída (estrutura inteligente)

```json
{
  "titulo": "...",
  "data": "YYYY-MM-DD",
  "time": "...",
  "participantes": ["Nome1", "Nome2"],
  "tipo_reuniao": "diagnóstico|operacional|estratégico|alinhamento|review|planning|workshop",
  "contexto": "1-2 frases: por que essa reunião aconteceu e qual era o objetivo central",
  "secoes": [
    {
      "titulo": "Título relevante ao conteúdo",
      "itens": [
        "item como texto simples",
        {"texto": "item com contexto", "detalhe": "explicação ou nuance"},
        {"acao": "o que fazer", "responsavel": "quem", "prazo": "quando"}
      ]
    }
  ],
  "acoes_pendentes": [
    {"item": "...", "responsavel": "...", "prazo": "..."}
  ]
}
```

**Regras:**
- `secoes` é livre — crie os títulos que melhor descrevem o que foi discutido
- Use `{"texto", "detalhe"}` quando um item precisa de contexto para ser compreendido
- Use `{"acao", "responsavel", "prazo"}` para compromissos e próximos passos
- `acoes_pendentes` é obrigatório e sempre preenchido — é o que alimenta o rastreamento semanal. Se não há prazo, use `"A definir"`
- Se um campo não se aplica, omita — não force seções vazias

#### Boas práticas para fácil consulta

- **Título de seção** = o que o leitor vai procurar, não o que o facilitador usou ("Bloqueios do Sprint" é melhor que "Riscos")
- **Itens curtos** — a ideia principal em uma frase; use `detalhe` para contexto extra
- **Capture intenção**, não transcrição — "Time vai pausar novas features até resolver o débito técnico" é melhor que "Fulano disse que precisa parar de fazer feature"

---

### 5. Criar o arquivo .docx

Use a ferramenta Write para criar o JSON em `$env:TEMP\resumo_reuniao.json`. Em seguida:

```powershell
cd "$env:USERPROFILE\.claude\skills\resumo-reunioes-teams"
python scripts/create_docx_smart.py --file "$env:TEMP\resumo_reuniao.json" --output "$env:TEMP\resumo_reuniao.docx"
```

O script retorna `{"success": true, "path": "...", "filename": "..."}`.

Antes de fazer upload, renomeie o arquivo com o padrão correto:
```powershell
Rename-Item "$env:TEMP\resumo_reuniao.docx" "YYYY-MM-DD_titulo_time.docx"
```

---

### 7. Fazer upload para o OneDrive

```powershell
cd "$env:USERPROFILE\.claude\skills\resumo-reunioes-teams"
python scripts/onedrive_upload.py "$env:TEMP\YYYY-MM-DD_titulo_time.docx" --team "<nome_do_time>"
```

O token de autenticação fica em cache após a primeira execução — nas próximas vezes o upload é silencioso.

O script cuida de:
- Criar a pasta `Reuniões dos Times` se não existir
- Criar a subpasta do time se não existir
- Retornar o link direto para o arquivo no OneDrive

---

### 8. Confirmar ao usuário

Informe:
- Nome do arquivo salvo
- Caminho no OneDrive (`Reuniões dos Times/<time>/`)
- Link direto (se disponível)

Pergunte: *"Quer consultar alguma informação desta reunião ou de reuniões anteriores?"*
Se sim → acione a skill `reunioes-consulta`.

---

## Relatório semanal por email

O relatório é um **estudo da semana** — não apenas uma lista de itens, mas uma análise dos temas recorrentes e do que não avançou entre as reuniões.

### Fluxo completo

**Passo 1 — Verificar configuração:**

Leia `skill_config.json` na raiz da skill. Se o arquivo não existir ou estiver incompleto, pergunte ao usuário:
- Email corporativo (usado no envio via Outlook): `work_email`
- Email pessoal para rascunho (Gmail draft): `personal_draft_email`
- Nome do time no Teams: `teams_team`
- Nome do canal no Teams: `teams_channel`

Salve as respostas em `skill_config.json` (não commitado no git).

**Passo 2 — Coletar dados brutos das reuniões da semana:**
```powershell
cd "$env:USERPROFILE\.claude\skills\resumo-reunioes-teams"
python scripts/weekly_consolidate.py --json-output "$env:TEMP\reunioes_semana.json"
```

**Passo 3 — Analisar e escrever o HTML do relatório (sua responsabilidade):**

Leia o JSON gerado. Para cada reunião, você terá `tipo_reuniao`, `contexto`, `secoes` e `acoes_pendentes`.

Escreva o HTML do relatório com as seguintes seções:

1. **Estudo da Semana** — 3-5 parágrafos analisando os temas que atravessaram as reuniões. Identifique padrões: o que o time está tentando resolver? Quais tensões apareceram em mais de uma reunião? O que ainda não tem resposta?

2. **Pontos que não avançaram** — itens de `acoes_pendentes` que aparecem em mais de uma reunião, ou que eram pendência de semanas anteriores (compare títulos e responsáveis). Destaque o tempo que estão em aberto.

3. **Combinados e Próximos Passos** — todos os `acoes_pendentes` da semana, agrupados por responsável, com prazo.

4. **Reuniões da semana** — lista com título, time, data e `contexto` de cada reunião processada.

Salve o HTML em `$env:TEMP\relatorio_semanal.html` usando a ferramenta Write.

**Passo 4 — Enviar via Outlook:**
```powershell
cd "$env:USERPROFILE\.claude\skills\resumo-reunioes-teams"
python scripts/send_report.py --email "<work_email>" --teams-team "<teams_team>" --teams-channel "<teams_channel>" --report-file "$env:TEMP\relatorio_semanal.html"
```
*(substitua os valores pelo que está em `skill_config.json`)*

**Passo 5 — Criar rascunho no Gmail:**

Use `mcp__claude_ai_Gmail__create_draft` com:
- `to`: `["<personal_draft_email>"]`
- `subject`: `"Relatório Semanal de Reuniões — DD/MM/YYYY"`
- `body`: conteúdo HTML do relatório

Confirme ao usuário os destinatários usados.

### Notas

- `send_report.py` envia via Graph API (Outlook) — aceita qualquer destinatário no parâmetro `--email`
- O rascunho Gmail fica disponível para revisão e reencaminhamento antes de enviar
- Para alterar destinatários sem reconfigurar tudo, edite `skill_config.json` diretamente

---

## Como extrair contexto com inteligência

**Regras de negócio** — afirmações sobre como o sistema ou processo DEVE funcionar.
Sinais: "precisa sempre", "a regra é que", "não pode ser feito sem", "obrigatoriamente", "por padrão".

**Combinados** — compromissos assumidos entre pessoas, explicitamente ou implicitamente.
Sinais: "você consegue fazer isso?", "fica com você", "combinamos que", "fulano vai verificar", "pode ficar responsável por".

**Próximos passos** — ações com dono e/ou prazo definido ou sugerido.
Sinais: "até sexta", "na próxima sprint", "vou fazer isso", "precisamos entregar", "precisa estar pronto".

**Decisões** — resoluções de dúvidas ou escolhas que estavam em aberto.
Sinais: "ficou decidido", "vamos usar X", "descartamos a opção Y", "escolhemos", "então vai ser assim".

**Riscos/Impedimentos** — preocupações com impacto potencial no projeto ou entrega.
Sinais: "pode travar", "tem um risco de", "se não resolvermos", "está bloqueado por", "dependência de", "preocupa".

**Pendências** — itens em aberto aguardando resposta ou ação externa.
Sinais: "ainda não temos resposta de", "precisa de aprovação", "estamos esperando", "falta definir", "depende de".

---

## Templates dos documentos finais

### Formato A — Padrão (`create_docx.py`)

Para referência:

```
RESUMO DE REUNIÃO
═══════════════════════════════════════════════

Título:        [titulo]
Data:          [data formatada em pt-BR]
Time:          [time]
Participantes: [Nome1, Nome2, Nome3]

───────────────────────────────────────────────
REGRAS DE NEGÓCIO

• [regra]
(se vazio: "Nenhuma regra de negócio mencionada.")

───────────────────────────────────────────────
COMBINADOS DA REUNIÃO

• [combinado] — Responsável: [nome] | Prazo: [prazo]
(se vazio: "Nenhum combinado registrado.")

───────────────────────────────────────────────
PRÓXIMOS PASSOS

• [ação] — Responsável: [nome] | Prazo: [prazo]
(se vazio: "Nenhum próximo passo registrado.")

───────────────────────────────────────────────
DECISÕES TOMADAS

• [decisão]
(se vazio: "Nenhuma decisão formal registrada.")

───────────────────────────────────────────────
RISCOS E IMPEDIMENTOS

• [risco] — Impacto: [Alto/Médio/Baixo]
(se vazio: "Nenhum risco ou impedimento levantado.")

───────────────────────────────────────────────
PENDÊNCIAS

• [pendência] — Responsável: [nome] | Prazo: [prazo]
(se vazio: "Nenhuma pendência registrada.")

═══════════════════════════════════════════════
Resumo gerado em: [data e hora atual]
```

---

### Formato B — Diagnóstico (`create_docx_diagnostico.py`)

```
RESUMO DE REUNIÃO
═══════════════════════════════════════════════

Título:        [titulo]
Data:          [data formatada em pt-BR]
Time:          [time]
Participantes: [Nome1, Nome2, Nome3]

───────────────────────────────────────────────
O QUE FUNCIONA

• [titulo]
  [detalhe]

───────────────────────────────────────────────
O QUE NÃO FUNCIONA

• [titulo]
  [detalhe]

───────────────────────────────────────────────
ENCAMINHAMENTOS                    (omitir seção se vazio)

• [item] — Responsável: [nome] | Prazo: [prazo]

═══════════════════════════════════════════════
Resumo gerado em: [data e hora atual]
```
