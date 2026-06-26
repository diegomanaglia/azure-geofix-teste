import base64
import io
import json
import os
import sys
import warnings
from datetime import datetime, timezone, timedelta

import msal
import requests
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore", message="Print area cannot be set")

# ── Credenciais via secrets do GitHub ─────────────────────
CLIENT_ID   = os.environ["AZURE_CLIENT_ID"]
TENANT_ID   = os.environ["AZURE_TENANT_ID"]
SHARE_LINK  = os.environ["SHAREPOINT_SHARE_LINK"]
CACHE_JSON  = os.environ["MSAL_TOKEN_CACHE"]
ABA         = os.environ.get("SHEET_NAME", "BASE SISTEMA")
SCOPES      = ["https://graph.microsoft.com/Files.Read.All",
               "https://graph.microsoft.com/Sites.Read.All"]
# ──────────────────────────────────────────────────────────

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

    # Salva o cache atualizado em disco para o update_secret.py ler
    if cache.has_state_changed:
        with open("token_cache.json", "w", encoding="utf-8") as f:
            f.write(cache.serialize())
        print("Cache do token atualizado em disco.")

    return result["access_token"]


def encode_share_url(url: str) -> str:
    b64 = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")


def df_para_json(df: pd.DataFrame) -> list:
    registros = []
    for _, row in df.iterrows():
        registro = {}
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                registro[str(col)] = None
            elif isinstance(val, pd.Timestamp):
                registro[str(col)] = val.isoformat()
            elif hasattr(val, "item"):
                registro[str(col)] = val.item()
            else:
                registro[str(col)] = val
        registros.append(registro)
    return registros


def main():
    print("Obtendo token...")
    token   = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    print("Resolvendo link de compartilhamento...")
    share_id  = encode_share_url(SHARE_LINK)
    item_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
        headers=headers,
    )
    item_resp.raise_for_status()
    item = item_resp.json()
    print(f"Arquivo: {item['name']}")

    print("Baixando planilha...")
    conteudo = requests.get(item["@microsoft.graph.downloadUrl"]).content

    print(f"Lendo aba '{ABA}'...")
    df = pd.read_excel(io.BytesIO(conteudo), sheet_name=ABA)
    print(f"{len(df)} linhas x {len(df.columns)} colunas")

    fuso_brasilia = timezone(timedelta(hours=-3))
    agora = datetime.now(fuso_brasilia).strftime("%Y-%m-%dT%H:%M:%S-03:00")

    output = {
        "ultima_atualizacao": agora,
        "aba": ABA,
        "total": len(df),
        "dados": df_para_json(df)
    }

    os.makedirs("docs", exist_ok=True)
    with open("docs/dados.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"Concluido: docs/dados.json gerado com {len(df)} registros.")


if __name__ == "__main__":
    main()
