---
name: resumo-reunioes-teams
description: Processa transcrições de reuniões do Microsoft Teams e gera um resumo estruturado salvo automaticamente no OneDrive do usuário. Use esta skill SEMPRE que o usuário mencionar: resumir reunião, processar transcrição, arquivo de reunião, VTT do Teams, exportar reunião, salvar resumo, transcrição Teams. Acione também para frases como "faz o resumo dessa reunião", "resume as reuniões de hoje", "processa essa transcrição", "salva essa reunião no OneDrive", "quero resumir minha reunião", ou quando o usuário fornecer o caminho de um arquivo .vtt ou .docx de transcrição do Teams.
---

# Resumo de Reuniões Teams

Processa transcrições do Teams e gera resumo estruturado salvo no OneDrive.

---

## Critérios de filtragem de reuniões

Antes de processar qualquer reunião, aplique os filtros abaixo. Reuniões fora do critério são ignoradas silenciosamente.

### Ignorar sempre

| Condição | Exemplos |
|---|---|
| Título contém `[Refinamento]` | [Refinamento] Projetos Sprint #36 |
| Título contém `Diálogo de Inovação` | Diálogo de Inovação |
| Reunião **recorrente** (`recurrence != null`) onde o usuário **não é organizador** | Dailys, Standups de outros times |
| Qualquer reunião dos times **B2B2C** ou **Evolução e Regulatório** dos tipos: Daily, Review, Refinamento, Alinhamento Técnico, Planning | [DAILY] B2B2C, [DAILY] Evolução e Regulatório, [Alinhamento Técnico] Evolução... |

Organizadores típicos dessas reuniões ignoradas: `henrique.oliveira@thunders.com.br`, `DEVEvoluoThunders2@thunders.com.br`.

### Processar

- Reuniões onde `isOrganizer: true` (o usuário criou a reunião)
- Reuniões não recorrentes com participação ativa, independente do time (avaliações, workshops, demos, alinhamentos pontuais)
- Reuniões do time B2B ou Projetos que o usuário organizou

### Como identificar na prática

Ao varrer o calendário, para cada evento:
1. Título tem `[Refinamento]` ou `Diálogo de Inovação`? → **ignorar**
2. `recurrence != null` E `isOrganizer: false`? → **ignorar**
3. Organizador é de B2B2C ou Evolução e Regulatório E título tem Daily/Review/Alinhamento/Planning/Refinamento? → **ignorar**
4. Qualquer outro caso → **processar** (tentar obter transcrição)

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
<v Sidney Silva>Precisamos definir o fluxo de aprovação.</v>
```
Extraia nomes dos speakers das tags `<v Nome>` ou do padrão `Nome: texto`.

**DOCX (Modo B):** execute:
```powershell
python "C:\Users\sidne\.claude\skills\resumo-reunioes-teams\scripts\parse_transcript.py" "<caminho_arquivo>"
```
Retorna JSON com texto estruturado.

---

### 3. Identificar o time

Se o usuário não informar o time, analise o conteúdo e pergunte:
> "Esse resumo é do time **B2B**, **Projetos** ou de outro time?"

Para um novo time, a pasta será criada automaticamente no OneDrive durante o upload.

---

### 4. Gerar o resumo (sua principal responsabilidade)

Analise profundamente o texto e preencha o JSON abaixo. Seja criterioso — vá além das palavras literais, capture a intenção real do que foi dito.

```json
{
  "titulo": "...",
  "data": "YYYY-MM-DD",
  "time": "...",
  "participantes": ["Nome1", "Nome2"],
  "regras_negocio": ["..."],
  "combinados": [{"item": "...", "responsavel": "...", "prazo": "..."}],
  "proximos_passos": [{"acao": "...", "responsavel": "...", "prazo": "..."}],
  "decisoes": ["..."],
  "riscos": [{"item": "...", "impacto": "Alto|Médio|Baixo"}],
  "pendencias": [{"item": "...", "responsavel": "...", "prazo": "..."}]
}
```

Se um campo não tiver conteúdo, use lista vazia `[]`.

---

### 5. Criar o arquivo .docx

Use a ferramenta Write para criar o JSON em `C:\Users\sidne\AppData\Local\Temp\resumo_reuniao.json`. Em seguida:

```powershell
cd "C:\Users\sidne\.claude\skills\resumo-reunioes-teams"
python scripts/create_docx.py --file "C:\Users\sidne\AppData\Local\Temp\resumo_reuniao.json" --output "C:\Users\sidne\AppData\Local\Temp\resumo_reuniao.docx"
```

O script retorna `{"success": true, "path": "...", "filename": "..."}`.

Antes de fazer upload, renomeie o arquivo com o padrão correto:
```powershell
Rename-Item "C:\Users\sidne\AppData\Local\Temp\resumo_reuniao.docx" "YYYY-MM-DD_titulo_time.docx"
```

---

### 6. Fazer upload para o OneDrive

```powershell
python scripts/onedrive_upload.py "C:\Users\sidne\AppData\Local\Temp\YYYY-MM-DD_titulo_time.docx" --team "<nome_do_time>"
```

O token de autenticação fica em cache após a primeira execução — nas próximas vezes o upload é silencioso.

O script cuida de:
- Criar a pasta `Reuniões dos Times` se não existir
- Criar a subpasta do time se não existir
- Retornar o link direto para o arquivo no OneDrive

---

### 7. Confirmar ao usuário

Informe:
- Nome do arquivo salvo
- Caminho no OneDrive (`Reuniões dos Times/<time>/`)
- Link direto (se disponível)

Pergunte: *"Quer consultar alguma informação desta reunião ou de reuniões anteriores?"*
Se sim → acione a skill `reunioes-consulta`.

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

## Template do documento final

O `create_docx.py` usa este template automaticamente. Para referência:

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
