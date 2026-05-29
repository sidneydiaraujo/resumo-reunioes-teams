"""
onedrive_upload.py
Faz upload de um arquivo .docx para o OneDrive via Microsoft Graph API.
Na primeira execução, pede autenticação via código (microsoft.com/devicelogin).
O token fica em cache — próximas execuções são silenciosas.

Uso: python onedrive_upload.py <caminho_arquivo.docx> --team <nome_do_time>
     python onedrive_upload.py <caminho_arquivo.docx> --team B2B
"""

import sys
import json
import argparse
import subprocess
from pathlib import Path


TOKEN_CACHE_FILE = Path.home() / ".claude" / "onedrive_auth_cache.bin"
ROOT_FOLDER = "Reuniões dos Times"

# App ID público do Microsoft Graph Explorer (sem necessidade de registro)
# Se o seu tenant bloquear este app, registre um app no Azure AD e substitua o CLIENT_ID
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.ReadWrite", "User.Read"]


def ensure_deps():
    try:
        import msal
        import requests
    except ImportError:
        print("Instalando dependências (msal, requests)...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "msal", "requests", "-q"])


def get_token() -> str:
    import msal

    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        # Tenta Integrated Windows Auth (funciona em máquinas domain-joined)
        try:
            result = app.acquire_token_by_integrated_windows_auth(scopes=SCOPES)
        except Exception:
            result = None

    if not result or "access_token" not in result:
        # Device code flow (abre autenticação no browser)
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Erro ao iniciar autenticação: {flow.get('error_description','desconhecido')}")

        print("\n" + "="*60)
        print("AUTENTICAÇÃO NECESSÁRIA (apenas na primeira vez)")
        print("="*60)
        print(f"\n1. Abra: https://microsoft.com/devicelogin")
        print(f"2. Digite o código: {flow['user_code']}")
        print(f"\nAguardando autenticação...")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Falha na autenticação: {result.get('error_description', result.get('error',''))}")

    # salva cache
    TOKEN_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")

    return result["access_token"]


def graph_request(token: str, method: str, url: str, **kwargs):
    import requests

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }
    headers.update(kwargs.pop("headers", {}))
    resp = getattr(requests, method)(url, headers=headers, **kwargs)

    if resp.status_code in (401, 403):
        raise RuntimeError(f"Sem permissão para acessar o OneDrive ({resp.status_code}). "
                           "Verifique se o app tem permissão no seu tenant.")
    return resp


def ensure_folder(token: str, path: str) -> str:
    """Garante que a pasta existe no OneDrive. Cria se necessário. Retorna o ID da pasta."""
    base_url = "https://graph.microsoft.com/v1.0/me/drive"
    encoded_path = path.strip("/")

    # Tenta acessar o caminho completo diretamente
    resp = graph_request(token, "get", f"{base_url}/root:/{encoded_path}")
    if resp.status_code == 200:
        return resp.json()["id"]

    # Cria pasta a pasta recursivamente
    parts = encoded_path.split("/")
    current_path = ""
    current_id = None

    for part in parts:
        current_path = f"{current_path}/{part}".lstrip("/")
        resp = graph_request(token, "get", f"{base_url}/root:/{current_path}")
        if resp.status_code == 200:
            current_id = resp.json()["id"]
        else:
            # Pasta não existe — cria
            parent_url = (
                f"{base_url}/root/children"
                if not current_id
                else f"{base_url}/items/{current_id}/children"
            )
            # Se já temos current_id, o pai é o nível anterior
            if current_id:
                # Sobe um nível para criar dentro do pai correto
                parent_path = "/".join(current_path.split("/")[:-1])
                if parent_path:
                    parent_resp = graph_request(token, "get", f"{base_url}/root:/{parent_path}")
                    parent_id = parent_resp.json()["id"]
                    parent_url = f"{base_url}/items/{parent_id}/children"
                else:
                    parent_url = f"{base_url}/root/children"

            create_resp = graph_request(
                token, "post", parent_url,
                headers={"Content-Type": "application/json"},
                json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
            )
            if create_resp.status_code not in (200, 201):
                raise RuntimeError(f"Erro ao criar pasta '{part}': {create_resp.text}")
            current_id = create_resp.json()["id"]
            print(f"  Pasta criada: {part}")

    return current_id


def upload_file(token: str, folder_id: str, filename: str, content: bytes) -> dict:
    """Faz upload do arquivo para a pasta especificada."""
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{filename}:/content"
    resp = graph_request(
        token, "put", url,
        headers={"Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
        data=content
    )

    if resp.status_code not in (200, 201):
        raise RuntimeError(f"Erro no upload: {resp.text}")

    item = resp.json()
    return {
        "id": item.get("id"),
        "name": item.get("name"),
        "webUrl": item.get("webUrl", ""),
        "size": item.get("size", 0)
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="Caminho do arquivo .docx para upload")
    parser.add_argument("--team", required=True, help="Nome do time (pasta de destino)")
    args = parser.parse_args()

    file_path = Path(args.file)
    if not file_path.exists():
        print(json.dumps({"error": f"Arquivo não encontrado: {args.file}"}))
        sys.exit(1)

    ensure_deps()

    print(f"Fazendo upload de: {file_path.name}")
    print(f"Destino: {ROOT_FOLDER}/{args.team}/")

    try:
        token = get_token()

        # garante estrutura de pastas
        folder_path = f"{ROOT_FOLDER}/{args.team}"
        print(f"Verificando/criando pastas em OneDrive...")
        folder_id = ensure_folder(token, folder_path)

        # faz upload
        content = file_path.read_bytes()
        result = upload_file(token, folder_id, file_path.name, content)

        print(json.dumps({
            "success": True,
            "filename": result["name"],
            "path": f"{ROOT_FOLDER}/{args.team}/{result['name']}",
            "webUrl": result["webUrl"],
            "size_kb": round(result["size"] / 1024, 1)
        }, ensure_ascii=False, indent=2))

    except RuntimeError as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
