import base64
import io
import os
import re
import sys
import warnings

import msal
import requests
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore", message="Print area cannot be set")

# Carrega variaveis do .env em execucoes locais (ignorado se nao instalado)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Credenciais via secrets do GitHub ─────────────────────
CLIENT_ID   = os.environ["AZURE_CLIENT_ID"]
TENANT_ID   = os.environ["AZURE_TENANT_ID"]
CACHE_JSON  = os.environ["MSAL_TOKEN_CACHE"]
SCOPES      = ["https://graph.microsoft.com/Files.Read.All",
               "https://graph.microsoft.com/Sites.Read.All"]

# Planilha BASE (NFs de entrada)
SHARE_LINK_BASE = os.environ["SHAREPOINT_SHARE_LINK"]
ABA_BASE        = os.environ.get("SHEET_NAME", "BASE NFS DE ENTRADA").strip() or "BASE NFS DE ENTRADA"

# Planilha COMPLEMENTO (itens do pedido / sistema)
SHARE_LINK_COMP = os.environ["SHAREPOINT_SHARE_LINK_COMPLEMENTO"]
ABA_COMP        = os.environ.get("SHEET_NAME_COMPLEMENTO", "BASE SISTEMA").strip() or "BASE SISTEMA"

# Saida final, gravada na propria pasta do script
SAIDA_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "hospedagem_conciliado.csv")
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

    if cache.has_state_changed:
        with open("token_cache.json", "w", encoding="utf-8") as f:
            f.write(cache.serialize())
        print("Cache do token atualizado em disco.")

    return result["access_token"]


def encode_share_url(url: str) -> str:
    b64 = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")


def baixar_planilha(token: str, share_link: str, aba: str) -> pd.DataFrame:
    """Resolve o link de compartilhamento, baixa o arquivo e le a aba indicada.
    Le tudo como texto (dtype=str) para preservar o tratamento de codigos."""
    headers   = {"Authorization": f"Bearer {token}"}
    share_id  = encode_share_url(share_link)
    item_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
        headers=headers,
    )
    item_resp.raise_for_status()
    item = item_resp.json()
    print(f"  Arquivo: {item['name']}")

    conteudo = requests.get(item["@microsoft.graph.downloadUrl"]).content
    df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, dtype=str)
    print(f"  Aba '{aba}': {len(df)} linhas x {len(df.columns)} colunas")
    return df


# ── Helpers de conciliacao (de gerar_itens_pedido_conciliados.py) ──
def normaliza_chave(valor):
    """Normaliza codigos para o cruzamento: remove '.0' de floats exportados
    e zeros a esquerda (base usa '0006', sistema usa '6')."""
    if pd.isna(valor):
        return None
    texto = re.sub(r"\.0+$", "", str(valor).strip())
    if texto == "":
        return None
    return texto.lstrip("0") or "0"


def limpa_artefato_float(valor):
    """Remove o sufixo '.0' deixado pela leitura Excel em inteiros,
    preservando decimais reais e codigos com zeros a esquerda."""
    if pd.isna(valor):
        return ""
    texto = str(valor)
    if re.fullmatch(r"-?\d+\.0+", texto):
        return re.sub(r"\.0+$", "", texto)
    return texto


def conciliar(base: pd.DataFrame, comp: pd.DataFrame) -> pd.DataFrame:
    """Cruza os itens do pedido (espinha) com a base de NFs de entrada."""
    # Chaves normalizadas (nao alteram as colunas de saida)
    base["_oc"]  = base["ORDEM DE COMPRA"].map(normaliza_chave)
    base["_cod"] = base["CÓD PROD/SERV"].map(normaliza_chave)
    comp["_oc"]  = comp["Nº Ordem Compra"].map(normaliza_chave)
    comp["_cod"] = comp["Serviço"].map(normaliza_chave)

    # ── Enriquecimento 1: nome da categoria do servico (por codigo) ──
    nomes = (
        base.dropna(subset=["DESCRIÇÃO PRODUTO OU SERVIÇO"])
        .groupby("_cod")["DESCRIÇÃO PRODUTO OU SERVIÇO"]
        .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else s.iloc[0])
    )
    mapa_categoria = nomes.to_dict()

    # ── Enriquecimento 2: conciliacao com a base de NFs (nivel OC+item) ──
    base_val = pd.to_numeric(base["VALOR TOTAL"], errors="coerce")
    nf_valor = base_val.groupby([base["_oc"], base["_cod"]]).sum().to_dict()
    nf_qtde  = base.groupby(["_oc", "_cod"]).size().to_dict()

    rat_val      = pd.to_numeric(comp["Valor do Rateio"], errors="coerce")
    rateio_total = rat_val.groupby([comp["_oc"], comp["_cod"]]).sum().to_dict()

    # ── Monta a planilha (espinha = itens do pedido) ──────────
    sai = pd.DataFrame()
    sai["Nº Ordem Compra"]          = comp["Nº Ordem Compra"].map(limpa_artefato_float)
    sai["Seq."]                     = comp["Seq."].map(limpa_artefato_float)
    sai["Cód Serviço"]              = comp["Serviço"].map(limpa_artefato_float)
    sai["Categoria do Serviço"]     = comp["_cod"].map(lambda c: mapa_categoria.get(c, ""))
    sai["Complemento da Descrição"] = comp["Complemento da Descrição"].fillna("")
    sai["Quantidade Pedida"]        = pd.to_numeric(comp["Quantidade Pedida"], errors="coerce")
    sai["U.M. O.C."]                = comp["U.M. O.C."].map(limpa_artefato_float)
    sai["Valor do Rateio"]          = rat_val
    sai["Fornecedor"]               = comp["Fornecedor"].map(limpa_artefato_float)
    sai["Fantasia Fornecedor"]      = comp["Fantasia Fornecedor"].map(limpa_artefato_float)
    sai["Projeto"]                  = comp["Projeto"].map(limpa_artefato_float)
    sai["Descr. Projeto"]           = comp["Descr. Projeto"].map(limpa_artefato_float)
    sai["Fase"]                     = comp["Fase"].map(limpa_artefato_float)
    sai["Descr. Fase"]              = comp["Descr. Fase"].map(limpa_artefato_float)
    sai["Conta Financeira"]         = comp["Conta Financeira"].map(limpa_artefato_float)
    sai["Conta Contábil"]           = comp["Conta Contábil"].map(limpa_artefato_float)
    sai["Usuário Comprador"]        = comp["Usuário Comprador (Nome)"].map(limpa_artefato_float)
    sai["Setor"]                    = comp["Setor"].map(limpa_artefato_float)
    sai["Emissão"]                  = comp["Emissão"].map(limpa_artefato_float)

    # Conciliacao planejado (rateio) x realizado (NF), no nivel OC+item
    chaves = list(zip(comp["_oc"], comp["_cod"]))
    consta = [k in nf_qtde for k in chaves]
    sai["Consta na base de NF"]              = ["Sim" if c else "Não" for c in consta]
    sai["Lançamentos NF na base (OC+item)"]  = [nf_qtde.get(k, 0) for k in chaves]
    sai["Valor Rateio total (OC+item)"]      = [rateio_total.get(k) for k in chaves]
    sai["Valor NF na base (OC+item)"]        = [nf_valor.get(k) for k in chaves]
    sai["Diferença NF - Rateio (OC+item)"]   = [
        (nf_valor.get(k) - rateio_total.get(k))
        if (nf_valor.get(k) is not None and rateio_total.get(k) is not None)
        else None
        for k in chaves
    ]

    # ── Relatorio ─────────────────────────────────────────────
    n_consta = sum(consta)
    print("\n--- Resumo da conciliacao ---")
    print(f"  Itens do pedido (linhas)        : {len(sai)}")
    print(f"  Categoria do servico preenchida : {(sai['Categoria do Serviço']!='').sum()}")
    print(f"  Itens que constam na base de NF : {n_consta} ({100*n_consta/len(sai):.1f}%)")

    return sai


def main():
    print("Obtendo token...")
    token = get_token()

    print("\nBaixando planilha BASE (NFs de entrada)...")
    base = baixar_planilha(token, SHARE_LINK_BASE, ABA_BASE)

    print("\nBaixando planilha COMPLEMENTO (itens do pedido)...")
    comp = baixar_planilha(token, SHARE_LINK_COMP, ABA_COMP)

    print("\nConciliando...")
    sai = conciliar(base, comp)

    sai.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig")
    print(f"\nConcluido: '{SAIDA_CSV}' gerado com {len(sai)} linhas "
          f"x {len(sai.columns)} colunas.")


if __name__ == "__main__":
    main()
