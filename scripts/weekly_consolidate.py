# -*- coding: utf-8 -*-
"""
weekly_consolidate.py
Lê os resumos .docx da semana atual no OneDrive e gera um relatório HTML
com todos os combinados e pendências consolidados.

Uso:
  python weekly_consolidate.py --output relatorio_semanal.html
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime, timedelta

TOKEN_CACHE_FILE = Path.home() / ".claude" / "onedrive_auth_cache.bin"
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.ReadWrite", "User.Read"]
ROOT_FOLDER = "Reuniões dos Times"


def ensure_deps():
    try:
        import msal, requests
        from docx import Document
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "msal", "requests", "python-docx", "-q"])


def get_token() -> str:
    import msal

    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result or "access_token" not in result:
        raise RuntimeError("Token expirado. Execute onedrive_upload.py para re-autenticar.")

    TOKEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def get_week_range():
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return monday.replace(hour=0, minute=0, second=0), sunday.replace(hour=23, minute=59, second=59)


def list_week_files(token: str) -> list:
    import requests

    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    week_start, week_end = get_week_range()
    files = []

    # Lista subpastas de Reuniões dos Times
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/root:/{ROOT_FOLDER}:/children",
        headers=headers
    )
    if resp.status_code != 200:
        print(f"Aviso: não encontrou {ROOT_FOLDER}: {resp.status_code}")
        return files

    team_folders = resp.json().get("value", [])
    for folder in team_folders:
        if "folder" not in folder:
            continue
        folder_id = folder["id"]
        team_name = folder["name"]

        files_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children",
            headers=headers
        )
        for item in files_resp.json().get("value", []):
            if not item["name"].endswith(".docx"):
                continue
            modified = datetime.fromisoformat(item["lastModifiedDateTime"].replace("Z", "+00:00")).replace(tzinfo=None)
            if week_start <= modified <= week_end:
                files.append({
                    "name": item["name"],
                    "team": team_name,
                    "id": item["id"],
                    "modified": modified.strftime("%d/%m/%Y %H:%M")
                })

    return files


def download_and_parse(token: str, file_id: str) -> dict:
    import requests, io
    from docx import Document

    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
        headers=headers
    )
    resp.raise_for_status()

    doc = Document(io.BytesIO(resp.content))
    result = {
        "titulo": "", "data": "", "time": "", "participantes": [],
        "combinados": [], "proximos_passos": [], "decisoes": [],
        "riscos": [], "pendencias": [], "regras_negocio": []
    }

    current_section = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if "RESUMO DE REUNIÃO" in text or "═" in text:
            continue
        if text.startswith("Título:"):
            result["titulo"] = text.replace("Título:", "").strip()
        elif text.startswith("Data:"):
            result["data"] = text.replace("Data:", "").strip()
        elif text.startswith("Time:"):
            result["time"] = text.replace("Time:", "").strip()
        elif text.startswith("Participantes:"):
            result["participantes"] = [p.strip() for p in text.replace("Participantes:", "").split(",")]
        elif "COMBINADOS" in text:
            current_section = "combinados"
        elif "PRÓXIMOS PASSOS" in text:
            current_section = "proximos_passos"
        elif "DECISÕES" in text:
            current_section = "decisoes"
        elif "RISCOS" in text:
            current_section = "riscos"
        elif "PENDÊNCIAS" in text:
            current_section = "pendencias"
        elif "REGRAS DE NEGÓCIO" in text:
            current_section = "regras_negocio"
        elif text.startswith("•") and current_section:
            item = text.lstrip("•").strip()
            if item and "Nenhum" not in item and "Nenhuma" not in item:
                result[current_section].append(item)
        elif text.startswith("Resumo gerado"):
            current_section = None

    return result


def format_html_report(meetings: list, week_start: datetime, week_end: datetime) -> str:
    semana = f"{week_start.strftime('%d/%m')} a {week_end.strftime('%d/%m/%Y')}"

    combinados_html = ""
    pendencias_html = ""
    proximos_html = ""
    decisoes_html = ""

    for m in meetings:
        titulo = m.get("titulo") or m.get("name", "")
        team = m.get("team", "")
        data = m.get("data", "")

        def section_block(items, label_key=None):
            if not items:
                return ""
            rows = ""
            for item in items:
                rows += f"<li>{item}</li>"
            return rows

        if m.get("combinados"):
            combinados_html += f'<h4>📋 {titulo} <span style="color:#888;font-weight:normal">({team} — {data})</span></h4><ul>{section_block(m["combinados"])}</ul>'

        if m.get("pendencias"):
            pendencias_html += f'<h4>⏳ {titulo} <span style="color:#888;font-weight:normal">({team} — {data})</span></h4><ul>{section_block(m["pendencias"])}</ul>'

        if m.get("proximos_passos"):
            proximos_html += f'<h4>▶ {titulo} <span style="color:#888;font-weight:normal">({team} — {data})</span></h4><ul>{section_block(m["proximos_passos"])}</ul>'

        if m.get("decisoes"):
            decisoes_html += f'<h4>✅ {titulo} <span style="color:#888;font-weight:normal">({team} — {data})</span></h4><ul>{section_block(m["decisoes"])}</ul>'

    reunioes_lista = "".join(
        f'<li><strong>{m.get("titulo") or m.get("name","")}</strong> — {m.get("team","")} ({m.get("data","")})</li>'
        for m in meetings
    )

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 24px; color: #242424; }}
  h1 {{ color: #1f497d; border-bottom: 2px solid #1f497d; padding-bottom: 8px; }}
  h2 {{ color: #1f497d; margin-top: 32px; border-left: 4px solid #1f497d; padding-left: 12px; }}
  h4 {{ margin: 16px 0 4px; color: #333; }}
  ul {{ margin: 4px 0 12px 0; padding-left: 20px; }}
  li {{ margin-bottom: 4px; line-height: 1.5; }}
  .meta {{ color: #666; font-size: 14px; margin-bottom: 24px; }}
  .section {{ background: #f9f9f9; border-radius: 8px; padding: 16px; margin-bottom: 16px; }}
  .empty {{ color: #999; font-style: italic; }}
</style>
</head>
<body>
<h1>Relatório Semanal de Reuniões</h1>
<p class="meta">Semana de {semana} &nbsp;|&nbsp; Gerado em {datetime.now().strftime("%d/%m/%Y às %H:%M")}</p>

<h2>Reuniões processadas ({len(meetings)})</h2>
<div class="section"><ul>{reunioes_lista or '<li class="empty">Nenhuma reunião processada.</li>'}</ul></div>

<h2>Combinados</h2>
<div class="section">{combinados_html or '<p class="empty">Nenhum combinado registrado.</p>'}</div>

<h2>Próximos Passos</h2>
<div class="section">{proximos_html or '<p class="empty">Nenhum próximo passo registrado.</p>'}</div>

<h2>Pendências em Aberto</h2>
<div class="section">{pendencias_html or '<p class="empty">Nenhuma pendência registrada.</p>'}</div>

<h2>Decisões Tomadas</h2>
<div class="section">{decisoes_html or '<p class="empty">Nenhuma decisão registrada.</p>'}</div>

</body>
</html>"""
    return html


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default="C:\\Users\\sidne\\AppData\\Local\\Temp\\relatorio_semanal.html")
    args = parser.parse_args()

    ensure_deps()
    token = get_token()

    week_start, week_end = get_week_range()
    print(f"Buscando arquivos da semana: {week_start.strftime('%d/%m')} a {week_end.strftime('%d/%m/%Y')}")

    files = list_week_files(token)
    print(f"Encontrados {len(files)} arquivos")

    meetings = []
    for f in files:
        print(f"  Lendo: {f['name']}")
        try:
            parsed = download_and_parse(token, f["id"])
            parsed["name"] = f["name"]
            parsed["team"] = f["team"]
            meetings.append(parsed)
        except Exception as e:
            print(f"  Aviso: erro ao ler {f['name']}: {e}")

    html = format_html_report(meetings, week_start, week_end)
    Path(args.output).write_text(html, encoding="utf-8")
    print(f"Relatório salvo em: {args.output}")
    print(json.dumps({"success": True, "meetings": len(meetings), "output": args.output}))


if __name__ == "__main__":
    main()
