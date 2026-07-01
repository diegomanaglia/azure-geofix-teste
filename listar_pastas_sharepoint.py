"""
Lista as pastas de primeiro nivel (dentro de '/', sem entrar em subpastas)
de uma pasta do SharePoint / OneDrive corporativo, a partir do link de
compartilhamento dessa pasta.

Usa o mesmo mecanismo de autenticacao (MSAL + cache de token) do script
'hospedagem_conciliado_csv.py', reaproveitando as credenciais do .env.
"""

import base64
import os
import sys
import time
import warnings

import msal
import requests

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Credenciais (mesmas do hospedagem_conciliado_csv.py) ──
CLIENT_ID  = os.environ["AZURE_CLIENT_ID"]
TENANT_ID  = os.environ["AZURE_TENANT_ID"]
CACHE_JSON = os.environ["MSAL_TOKEN_CACHE"]
SCOPES     = ["https://graph.microsoft.com/Files.Read.All",
              "https://graph.microsoft.com/Sites.Read.All"]

# Link de compartilhamento da pasta a ser explorada
SHARE_LINK_PASTA = os.environ["SHAREPOINT_SHARE_LINK_PASTA"]

GRAPH_URL = "https://graph.microsoft.com/v1.0"
# ────────────────────────────────────────────────────────────


def get_token() -> str:
    cache = msal.SerializableTokenCache()
    cache.deserialize(CACHE_JSON)

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    contas = app.get_accounts()
    if not contas:
        raise Exception("Nenhuma conta encontrada no cache. Refaca o login Device Code localmente.")

    result = app.acquire_token_silent(SCOPES, account=contas[0])
    if not result:
        raise Exception("Token expirado e nao foi possivel renovar. Refaca o login Device Code localmente.")

    if "access_token" not in result:
        raise Exception(f"Erro ao obter token: {result.get('error_description')}")

    if cache.has_state_changed:
        with open("token_cache.json", "w", encoding="utf-8") as f:
            f.write(cache.serialize())
        print("Cache do token atualizado em disco.")

    return result["access_token"]


def encode_share_url(url: str) -> str:
    b64 = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")


def resolver_pasta(token: str, share_link: str) -> dict:
    """Resolve o link de compartilhamento para o driveItem da pasta raiz."""
    headers = {"Authorization": f"Bearer {token}"}
    share_id = encode_share_url(share_link)
    resp = requests.get(
        f"{GRAPH_URL}/shares/{share_id}/driveItem",
        headers=headers,
        params={"$select": "id,name,folder,parentReference"},
    )
    resp.raise_for_status()
    item = resp.json()
    if "folder" not in item:
        raise ValueError(f"O link informado nao aponta para uma pasta (item: '{item.get('name')}').")
    return item


def listar_pastas_primeiro_nivel(token: str, drive_id: str, item_id: str) -> list:
    """Lista apenas as pastas filhas diretas de 'item_id' (sem entrar em
    subpastas), retornando uma lista de dicionarios {nome, qtd_itens}."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/children"
    params = {"$select": "id,name,folder", "$top": 200}
    resultado = []

    while url:
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        dados = resp.json()

        for filho in dados.get("value", []):
            if "folder" not in filho:
                continue  # ignora arquivos, so nos interessam pastas
            resultado.append({
                "nome": filho["name"],
                "qtd_itens": filho["folder"].get("childCount", 0),
            })

        url = dados.get("@odata.nextLink")
        params = None  # o nextLink ja vem com os parametros embutidos

    return resultado


def main():
    inicio = time.perf_counter()

    print("Obtendo token...")
    token = get_token()

    print("\nResolvendo pasta a partir do link de compartilhamento...")
    pasta_raiz = resolver_pasta(token, SHARE_LINK_PASTA)
    drive_id = pasta_raiz["parentReference"]["driveId"]
    print(f"Pasta raiz: {pasta_raiz['name']}")

    print("\nListando pastas de primeiro nivel...")
    pastas = listar_pastas_primeiro_nivel(token, drive_id, pasta_raiz["id"])

    print(f"\n--- {len(pastas)} pasta(s) encontrada(s) em '{pasta_raiz['name']}/' ---")
    for p in pastas:
        print(f"{p['nome']}  ({p['qtd_itens']} item(ns))")

    if not pastas:
        print("Nenhuma pasta encontrada (pode conter apenas arquivos, ou estar vazia).")

    duracao = time.perf_counter() - inicio
    print(f"\nTempo total da operacao: {duracao:.2f} segundos")


if __name__ == "__main__":
    main()
