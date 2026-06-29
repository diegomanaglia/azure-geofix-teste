import base64
import json
import os
import sys

import requests
from nacl import encoding, public

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

GH_PAT = os.environ["GH_PAT"]
REPO   = os.environ["REPO"]

CACHE_FILE    = "token_cache.json"
SECRET_NAME   = "MSAL_TOKEN_CACHE"

def encrypt_secret(public_key_b64: str, secret_value: str) -> str:
    """Criptografa o valor usando a chave publica do repositorio."""
    key    = public.PublicKey(public_key_b64.encode("utf-8"), encoding.Base64Encoder)
    box    = public.SealedBox(key)
    encrypted = box.encrypt(secret_value.encode("utf-8"))
    return base64.b64encode(encrypted).decode("utf-8")

def main():
    if not os.path.exists(CACHE_FILE):
        print(f"Arquivo {CACHE_FILE} nao encontrado — cache nao foi atualizado.")
        return

    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        novo_cache = f.read()

    headers = {
        "Authorization": f"Bearer {GH_PAT}",
        "Accept":        "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    # Busca a chave publica do repositorio para criptografar o secret
    url_key = f"https://api.github.com/repos/{REPO}/actions/secrets/public-key"
    resp_key = requests.get(url_key, headers=headers)
    resp_key.raise_for_status()
    key_data = resp_key.json()

    encrypted = encrypt_secret(key_data["key"], novo_cache)

    # Atualiza o secret
    url_secret = f"https://api.github.com/repos/{REPO}/actions/secrets/{SECRET_NAME}"
    resp_secret = requests.put(
        url_secret,
        headers=headers,
        json={
            "encrypted_value": encrypted,
            "key_id":          key_data["key_id"],
        }
    )

    if resp_secret.status_code in (201, 204):
        print("Cache do token atualizado no GitHub Secrets com sucesso.")
    else:
        print(f"Erro ao atualizar secret: {resp_secret.status_code} - {resp_secret.text}")
        sys.exit(1)

if __name__ == "__main__":
    main()
