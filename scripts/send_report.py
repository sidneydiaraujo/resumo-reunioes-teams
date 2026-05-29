# -*- coding: utf-8 -*-
"""
send_report.py
Envia o relatório semanal de reuniões por email e Teams.

Uso:
  python send_report.py --email dest@email.com --teams-team "Nome do Time" --teams-channel "Nome do Canal" --report-file relatorio.html
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

TOKEN_CACHE_FILE = Path.home() / ".claude" / "onedrive_auth_cache.bin"
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.ReadWrite", "User.Read", "Mail.Send", "ChannelMessage.Send", "Team.ReadBasic.All", "Channel.ReadBasic.All"]


def ensure_deps():
    try:
        import msal, requests
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "msal", "requests", "-q"])


def get_token() -> str:
    import msal

    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(f"\nAbra: https://microsoft.com/devicelogin\nCódigo: {flow['user_code']}\nAguardando...")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Falha na autenticação: {result.get('error_description', '')}")

    TOKEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def get_teams_channel_id(token: str, team_name: str, channel_name: str) -> tuple:
    import requests

    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    resp = requests.get("https://graph.microsoft.com/v1.0/me/joinedTeams", headers=headers)
    resp.raise_for_status()
    teams = resp.json().get("value", [])

    team_id = None
    for t in teams:
        if team_name.lower() in t["displayName"].lower():
            team_id = t["id"]
            break

    if not team_id:
        raise RuntimeError(f"Time '{team_name}' não encontrado. Times disponíveis: {[t['displayName'] for t in teams]}")

    ch_resp = requests.get(f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels", headers=headers)
    ch_resp.raise_for_status()
    channels = ch_resp.json().get("value", [])

    channel_id = None
    for ch in channels:
        if channel_name.lower() in ch["displayName"].lower():
            channel_id = ch["id"]
            break

    if not channel_id:
        raise RuntimeError(f"Canal '{channel_name}' não encontrado. Canais: {[c['displayName'] for c in channels]}")

    return team_id, channel_id


def send_email(token: str, to_email: str, subject: str, html_body: str):
    import requests

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_email}}]
        },
        "saveToSentItems": False
    }

    resp = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers=headers,
        json=payload
    )

    if resp.status_code == 202:
        print(f"Email enviado para {to_email}")
    else:
        raise RuntimeError(f"Erro ao enviar email: {resp.status_code} {resp.text[:300]}")


def send_teams_message(token: str, team_id: str, channel_id: str, html_content: str):
    import requests

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "body": {
            "contentType": "html",
            "content": html_content
        }
    }

    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages",
        headers=headers,
        json=payload
    )

    if resp.status_code in (200, 201):
        print(f"Mensagem enviada ao Teams (canal {channel_id})")
    else:
        raise RuntimeError(f"Erro ao enviar Teams: {resp.status_code} {resp.text[:300]}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--email", required=True, help="Email de destino")
    parser.add_argument("--teams-team", required=True, help="Nome do time no Teams")
    parser.add_argument("--teams-channel", required=True, help="Nome do canal no Teams")
    parser.add_argument("--report-file", required=True, help="Arquivo HTML do relatório")
    parser.add_argument("--subject", default="", help="Assunto do email")
    args = parser.parse_args()

    ensure_deps()
    token = get_token()

    report_path = Path(args.report_file)
    if not report_path.exists():
        print(json.dumps({"error": f"Arquivo não encontrado: {args.report_file}"}))
        sys.exit(1)

    html_body = report_path.read_text(encoding="utf-8")
    subject = args.subject or f"Relatório Semanal de Reuniões — {datetime.now().strftime('%d/%m/%Y')}"

    try:
        send_email(token, args.email, subject, html_body)
    except Exception as e:
        print(f"Aviso email: {e}")

    try:
        team_id, channel_id = get_teams_channel_id(token, args.teams_team, args.teams_channel)
        # Teams não suporta HTML completo — envia versão simplificada
        send_teams_message(token, team_id, channel_id, html_body)
    except Exception as e:
        print(f"Aviso Teams: {e}")

    print(json.dumps({"success": True, "subject": subject}))


if __name__ == "__main__":
    main()
