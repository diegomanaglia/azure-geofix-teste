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

# Planilhas BASE (NFs de entrada) — 2026 e 2025, mesma aba e mesmo layout
SHARE_LINK_BASE      = os.environ["SHAREPOINT_SHARE_LINK"]
SHARE_LINK_BASE_2025 = os.environ["SHAREPOINT_SHARE_LINK_BASE_2025"]
ABA_BASE             = os.environ.get("SHEET_NAME", "BASE NFS DE ENTRADA").strip() or "BASE NFS DE ENTRADA"
# Nas duas bases consideramos apenas as colunas A ate AH (1 a 34)
COLS_BASE            = "A:AH"

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


def baixar_planilha(token: str, share_link: str, aba: str, usecols=None) -> pd.DataFrame:
    """Resolve o link de compartilhamento, baixa o arquivo e le a aba indicada.
    Le tudo como texto (dtype=str) para preservar o tratamento de codigos.
    usecols: intervalo de colunas no estilo Excel (ex.: 'A:AH') ou None para todas."""
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
    df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, dtype=str, usecols=usecols)
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


def normaliza_codigo(valor):
    """Normaliza o codigo do produto/servico para o cruzamento como TEXTO:
    remove apenas o artefato '.0' da exportacao, mas PRESERVA zeros a esquerda
    (ex.: '26' e '0026' sao produtos diferentes e nao podem ser tratados iguais)."""
    if pd.isna(valor):
        return None
    texto = re.sub(r"\.0+$", "", str(valor).strip())
    if texto == "":
        return None
    return texto


def limpa_artefato_float(valor):
    """Remove o sufixo '.0' deixado pela leitura Excel em inteiros,
    preservando decimais reais e codigos com zeros a esquerda."""
    if pd.isna(valor):
        return ""
    texto = str(valor)
    if re.fullmatch(r"-?\d+\.0+", texto):
        return re.sub(r"\.0+$", "", texto)
    return texto


def _descricao_repr(serie):
    """Descricao representativa de um grupo: a mais frequente (mode)."""
    s = serie.dropna()
    if s.empty:
        return ""
    m = s.mode()
    return m.iloc[0] if not m.empty else s.iloc[0]


# Colunas numericas da saida (mantidas como numero p/ exportar com decimal ',')
COLS_NUMERICAS = [
    "Quantidade Pedida", "Valor do Rateio", "Lançamentos NF na base (OC+item)",
    "Valor Rateio total (OC+item)", "Valor NF na base (OC+item)",
    "Diferença NF - Rateio (OC+item)",
]


def conciliar(base: pd.DataFrame, comp: pd.DataFrame) -> pd.DataFrame:
    """Full outer entre itens do pedido (complemento) e NFs da base (nivel OC+item).
    Mantem TODOS os itens de pedido e TODAS as NFs da base; o que nao cruza fica em branco.
    - Parte 1: itens do pedido (com agregados da NF quando houver match).
    - Parte 2: NFs da base sem item de pedido (agregadas por OC+item)."""
    base = base.copy()
    comp = comp.copy()
    # Chaves: OC normalizada numericamente; CODIGO comparado como texto (preserva zeros)
    base["_oc"]  = base["ORDEM DE COMPRA"].map(normaliza_chave)
    base["_cod"] = base["CÓD PROD/SERV"].map(normaliza_codigo)
    comp["_oc"]  = comp["Nº Ordem Compra"].map(normaliza_chave)
    comp["_cod"] = comp["Serviço"].map(normaliza_codigo)

    # Categoria do servico = descricao mais frequente por CODIGO (na base)
    mapa_categoria = (
        base.dropna(subset=["DESCRIÇÃO PRODUTO OU SERVIÇO"])
        .groupby("_cod")["DESCRIÇÃO PRODUTO OU SERVIÇO"]
        .agg(_descricao_repr)
        .to_dict()
    )

    # Agregado da base por (OC+item): valor de NF, qtde de lancamentos e descricao
    base["_valnf"] = pd.to_numeric(base["VALOR TOTAL"], errors="coerce")
    sem_chave = base["_oc"].isna() | base["_cod"].isna()
    n_sem_chave = int(sem_chave.sum())
    agg = (
        base[~sem_chave]
        .groupby(["_oc", "_cod"])
        .agg(oc_orig=("ORDEM DE COMPRA", "first"),
             cod_orig=("CÓD PROD/SERV", "first"),
             descricao=("DESCRIÇÃO PRODUTO OU SERVIÇO", _descricao_repr),
             nf_valor=("_valnf", "sum"),
             nf_qtde=("_valnf", "size"))
        .reset_index()
    )
    agg["_key"] = list(zip(agg["_oc"], agg["_cod"]))
    nf_valor = dict(zip(agg["_key"], agg["nf_valor"]))
    nf_qtde  = dict(zip(agg["_key"], agg["nf_qtde"]))
    desc_map = dict(zip(agg["_key"], agg["descricao"]))

    rat_val      = pd.to_numeric(comp["Valor do Rateio"], errors="coerce")
    rateio_total = rat_val.groupby([comp["_oc"], comp["_cod"]]).sum().to_dict()

    # ── Parte 1: itens do pedido (espinha) ────────────────────
    chaves = list(zip(comp["_oc"], comp["_cod"]))
    sai = pd.DataFrame()
    sai["Nº Ordem Compra"]              = comp["Nº Ordem Compra"].map(limpa_artefato_float)
    sai["Seq."]                         = comp["Seq."].map(limpa_artefato_float)
    sai["Cód Serviço"]                  = comp["Serviço"].map(limpa_artefato_float)
    sai["DESCRIÇÃO PRODUTO OU SERVIÇO"] = [desc_map.get(k, "") for k in chaves]
    sai["Categoria do Serviço"]         = comp["_cod"].map(lambda c: mapa_categoria.get(c, ""))
    sai["Complemento da Descrição"]     = comp["Complemento da Descrição"].fillna("")
    sai["Quantidade Pedida"]            = pd.to_numeric(comp["Quantidade Pedida"], errors="coerce")
    sai["U.M. O.C."]                    = comp["U.M. O.C."].map(limpa_artefato_float)
    sai["Valor do Rateio"]              = rat_val.values
    sai["Fornecedor"]                   = comp["Fornecedor"].map(limpa_artefato_float)
    sai["Fantasia Fornecedor"]          = comp["Fantasia Fornecedor"].map(limpa_artefato_float)
    sai["Projeto"]                      = comp["Projeto"].map(limpa_artefato_float)
    sai["Descr. Projeto"]               = comp["Descr. Projeto"].map(limpa_artefato_float)
    sai["Fase"]                         = comp["Fase"].map(limpa_artefato_float)
    sai["Descr. Fase"]                  = comp["Descr. Fase"].map(limpa_artefato_float)
    sai["Conta Financeira"]             = comp["Conta Financeira"].map(limpa_artefato_float)
    sai["Conta Contábil"]               = comp["Conta Contábil"].map(limpa_artefato_float)
    sai["Usuário Comprador"]            = comp["Usuário Comprador (Nome)"].map(limpa_artefato_float)
    sai["Setor"]                        = comp["Setor"].map(limpa_artefato_float)
    sai["Emissão"]                      = comp["Emissão"].map(limpa_artefato_float)
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

    # ── Parte 2: NFs da base SEM item de pedido (agregado OC+item) ──
    chaves_comp = set(chaves)
    falt = agg[~agg["_key"].isin(chaves_comp)].copy()
    extra = pd.DataFrame(index=falt.index, columns=sai.columns)
    if len(falt):
        extra["Nº Ordem Compra"]              = falt["oc_orig"].map(limpa_artefato_float)
        extra["Seq."]                         = ""
        extra["Cód Serviço"]                  = falt["cod_orig"].map(limpa_artefato_float)
        extra["DESCRIÇÃO PRODUTO OU SERVIÇO"] = falt["descricao"].fillna("")
        extra["Categoria do Serviço"]         = falt["_cod"].map(lambda c: mapa_categoria.get(c, ""))
        for col in ["Complemento da Descrição", "U.M. O.C.", "Fornecedor",
                    "Fantasia Fornecedor", "Projeto", "Descr. Projeto", "Fase",
                    "Descr. Fase", "Conta Financeira", "Conta Contábil",
                    "Usuário Comprador", "Setor", "Emissão"]:
            extra[col] = ""
        extra["Consta na base de NF"]             = "NF sem pedido"
        extra["Lançamentos NF na base (OC+item)"] = falt["nf_qtde"].values
        extra["Valor NF na base (OC+item)"]       = falt["nf_valor"].values
        # Quantidade Pedida / Valor do Rateio / Valor Rateio total / Diferenca: vazios

    saida = pd.concat([sai, extra], ignore_index=True)

    # Garante dtype numerico nas colunas de valor (p/ exportar com decimal ',')
    for col in COLS_NUMERICAS:
        saida[col] = pd.to_numeric(saida[col], errors="coerce")

    # ── Relatorio ─────────────────────────────────────────────
    n_consta = sum(consta)
    print("\n--- Resumo da conciliacao (full outer) ---")
    print(f"  Itens do pedido               : {len(sai)}")
    print(f"    com NF na base (Sim)        : {n_consta}")
    print(f"    sem NF na base (Não)        : {len(sai) - n_consta}")
    print(f"  NFs da base sem pedido        : {len(falt)} chaves OC+item")
    print(f"  TOTAL de linhas               : {len(saida)}")
    if n_sem_chave:
        print(f"  [aviso] {n_sem_chave} linhas de NF na base sem OC ou codigo nao entraram (sem chave)")

    return saida


def main():
    print("Obtendo token...")
    token = get_token()

    print("\nBaixando planilha BASE 2026 (NFs de entrada)...")
    base_2026 = baixar_planilha(token, SHARE_LINK_BASE, ABA_BASE, usecols=COLS_BASE)

    print("\nBaixando planilha BASE 2025 (NFs de entrada)...")
    base_2025 = baixar_planilha(token, SHARE_LINK_BASE_2025, ABA_BASE, usecols=COLS_BASE)

    # As duas bases tem layout identico (colunas A:AH) -> empilha em uma so
    base = pd.concat([base_2026, base_2025], ignore_index=True)
    print(f"\nBase consolidada (2026 + 2025): {len(base)} linhas "
          f"({len(base_2026)} + {len(base_2025)}) x {len(base.columns)} colunas")

    print("\nBaixando planilha COMPLEMENTO (itens do pedido)...")
    comp = baixar_planilha(token, SHARE_LINK_COMP, ABA_COMP)

    print("\nConciliando...")
    sai = conciliar(base, comp)

    # Padrao brasileiro para abrir corretamente no Excel pt-BR:
    # separador de colunas ';' e separador decimal ',' (evita que 1739.4997
    # seja lido como 17394997 quando o ponto e' tratado como separador de milhar).
    sai.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig", sep=";", decimal=",")
    print(f"\nConcluido: '{SAIDA_CSV}' gerado com {len(sai)} linhas "
          f"x {len(sai.columns)} colunas.")


if __name__ == "__main__":
    main()
