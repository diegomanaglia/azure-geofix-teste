"""
Login interativo (Device Code) para gerar o cache de token MSAL usado pelos
demais scripts (hospedagem_conciliado_csv.py, listar_pastas_sharepoint.py etc.).

Rode este script UMA VEZ, localmente, a partir da raiz do repositorio
(pasta 'azure-geofix-teste'), sempre que precisar (re)criar o MSAL_TOKEN_CACHE
- por exemplo, quando o valor salvo nos GitHub Secrets nao pode ser recuperado.

O que ele faz:
  1. Inicia um fluxo de Device Code: mostra um codigo e um link no terminal.
  2. Voce abre o link em um navegador, informa o codigo e faz login com a
     conta corporativa (a mesma que tem acesso ao SharePoint/OneDrive).
  3. Salva 'token_cache.json' na raiz do repositorio.
  4. Imprime o conteudo (em uma linha) para copiar no .env -> MSAL_TOKEN_CACHE
     (e, se for o caso, atualizar tambem o secret MSAL_TOKEN_CACHE no GitHub).
"""

import os
import sys

import msal

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
TENANT_ID = os.environ["AZURE_TENANT_ID"]
SCOPES    = ["https://graph.microsoft.com/Files.Read.All",
             "https://graph.microsoft.com/Sites.Read.All"]

# Salva na raiz do repositorio (mesmo local usado pelos outros scripts)
CACHE_FILE = "token_cache.json"


def main():
    cache = msal.SerializableTokenCache()

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"Falha ao iniciar o Device Code Flow: {flow}")

    print(flow["message"])
    sys.stdout.flush()

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception(f"Erro ao obter token: {result.get('error_description')}")

    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        f.write(cache.serialize())

    print(f"\nLogin concluido com sucesso. Cache salvo em: {os.path.abspath(CACHE_FILE)}")
    print("\nCopie a linha abaixo para o .env, na variavel MSAL_TOKEN_CACHE")
    print("(e, se necessario, atualize tambem o secret MSAL_TOKEN_CACHE no GitHub):\n")
    print(cache.serialize())


if __name__ == "__main__":
    main()
