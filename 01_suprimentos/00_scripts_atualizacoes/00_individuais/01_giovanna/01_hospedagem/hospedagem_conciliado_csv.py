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


_RE_NUM = r"^-?\d+(\.\d+)?$"        # numero inteiro ou decimal (ponto)
_RE_DECIMAL = r"\.\d*[1-9]"          # tem parte decimal diferente de zero


def coluna_numerica(serie: pd.Series) -> pd.Series:
    """Converte a coluna para numero (float) somente se ela contiver valores
    decimais reais (ex.: '1739.4997'); assim o CSV sai com decimal ',' e o Excel
    pt-BR le certo. Caso contrario mantem como TEXTO (limpando o artefato '.0'),
    preservando codigos com zeros a esquerda, OCs e datas."""
    s = serie.astype("string").str.strip()
    nao_vazio = s[s.notna() & (s != "")]
    if len(nao_vazio) == 0:
        return serie.map(limpa_artefato_float)
    todos_num   = bool(nao_vazio.str.match(_RE_NUM).all())
    tem_decimal = bool(nao_vazio.str.contains(_RE_DECIMAL).any())
    if todos_num and tem_decimal:
        return pd.to_numeric(serie, errors="coerce")
    return serie.map(limpa_artefato_float)


# Tolerancia em reais para considerar VALOR TOTAL == Valor do Rateio
TOLERANCIA_VALOR = 0.01

# Niveis de correspondencia, do mais estrito ao mais flexivel:
# (rotulo, usa_codigo, usa_quantidade) -- o valor (VALOR TOTAL == Valor do Rateio)
# e exigido em todos os niveis.
NIVEIS_MATCH = [
    ("EXATO (cod+valor+qtd)", True,  True),
    ("ALTO (cod+valor)",      True,  False),
    ("MEDIO (valor+qtd)",     False, True),
    ("MEDIO (somente valor)", False, False),
]


def conciliar(base: pd.DataFrame, comp: pd.DataFrame) -> pd.DataFrame:
    """Mantem a base de NFs (2025+2026) COMPLETA como espinha (uma linha por NF,
    todas as colunas) e traz 'Complemento da Descrição' do complemento (alpha)
    para cada linha, por correspondencia em cascata dentro de cada Ordem de Compra:
      1) EXATO : codigo + valor (VALOR TOTAL == Valor do Rateio) + quantidade
      2) ALTO  : codigo + valor
      3) MEDIO : valor + quantidade
      4) MEDIO : somente valor
    Cada complemento e consumido no maximo uma vez por Ordem de Compra (evita que
    duas linhas reivindiquem a mesma origem). O que nao casar fica em branco."""
    base = base.copy()
    comp = comp.copy()
    # Chaves: OC normalizada numericamente; CODIGO comparado como texto (preserva zeros)
    base["_oc"]  = base["ORDEM DE COMPRA"].map(normaliza_chave)
    base["_cod"] = base["CÓD PROD/SERV"].map(normaliza_codigo)
    comp["_oc"]  = comp["Nº Ordem Compra"].map(normaliza_chave)
    comp["_cod"] = comp["Serviço"].map(normaliza_codigo)

    # Valores numericos usados na correspondencia
    base["_total"]  = pd.to_numeric(base["VALOR TOTAL"], errors="coerce")
    base["_qtd"]    = pd.to_numeric(base["QUANTIDADE"], errors="coerce")
    comp["_rateio"] = pd.to_numeric(comp["Valor do Rateio"], errors="coerce")
    comp["_qtdped"] = pd.to_numeric(comp["Quantidade Pedida"], errors="coerce")
    comp["_seq"]    = pd.to_numeric(comp["Seq."], errors="coerce")

    # Complemento escolhido na ordem de Seq. (o primeiro candidato = menor Seq.)
    comp = comp.sort_values(["_oc", "_seq"], kind="stable")
    comp_por_oc = {oc: grp for oc, grp in comp.groupby("_oc")}

    complemento = {}   # indice da linha do base -> texto do complemento (ou None)
    status = {}        # indice da linha do base -> status do vinculo

    for oc, base_grp in base.groupby("_oc", dropna=False):
        alpha_grp = comp_por_oc.get(oc)
        if alpha_grp is None or alpha_grp.empty:
            for idx in base_grp.index:
                complemento[idx] = None
                status[idx] = "OC AUSENTE NO ALPHA"
            continue

        a_idx   = list(alpha_grp.index)              # ja em ordem de Seq.
        a_cod   = alpha_grp["_cod"].to_dict()
        a_rat   = alpha_grp["_rateio"].to_dict()
        a_qtd   = alpha_grp["_qtdped"].to_dict()
        a_compl = alpha_grp["Complemento da Descrição"].to_dict()
        b_cod   = base_grp["_cod"].to_dict()
        b_total = base_grp["_total"].to_dict()
        b_qtd   = base_grp["_qtd"].to_dict()

        usados = set()                               # alpha ja consumidos nesta OC
        pendentes = list(base_grp.index)

        for rotulo, usar_cod, usar_qtd in NIVEIS_MATCH:
            ainda = []
            for bidx in pendentes:
                cod, total, qtd = b_cod[bidx], b_total[bidx], b_qtd[bidx]
                candidatos = []
                for aidx in a_idx:
                    if aidx in usados:
                        continue
                    if usar_cod and a_cod[aidx] != cod:
                        continue
                    rat = a_rat[aidx]
                    if pd.isna(rat) or pd.isna(total) or abs(rat - total) > TOLERANCIA_VALOR:
                        continue
                    if usar_qtd:
                        aq = a_qtd[aidx]
                        if pd.isna(aq) or pd.isna(qtd) or abs(aq - qtd) > 1e-9:
                            continue
                    candidatos.append(aidx)
                if not candidatos:
                    ainda.append(bidx)
                    continue
                distintos = {a_compl[c] for c in candidatos}
                escolhido = candidatos[0]            # primeiro pela ordem de Seq.
                usados.add(escolhido)
                obs = "" if len(distintos) == 1 else " [multiplos candidatos - revisar]"
                complemento[bidx] = a_compl[escolhido]
                status[bidx] = rotulo + obs
            pendentes = ainda
            if not pendentes:
                break

        for bidx in pendentes:
            complemento[bidx] = None
            status[bidx] = "SEM CORRESPONDENTE"

    # ── Saida: base COMPLETA + Complemento da Descrição + Status ──
    sai = pd.DataFrame()
    for col in base.columns:
        if col.startswith("_"):
            continue
        sai[col] = coluna_numerica(base[col])
    sai["Complemento da Descrição"] = [complemento.get(i) for i in base.index]
    sai["Status do Vínculo"]        = [status.get(i) for i in base.index]

    # ── Relatorio ─────────────────────────────────────────────
    serie_status = pd.Series([status.get(i) for i in base.index])
    com_match = int(sai["Complemento da Descrição"].notna().sum())
    print("\n--- Resumo (base enriquecida / cascata por Ordem de Compra) ---")
    print(f"  Linhas de NF (base 2025+2026): {len(sai)}")
    for st, qt in serie_status.value_counts().items():
        print(f"    {st}: {qt}")
    print(f"  Com complemento preenchido   : {com_match} ({com_match/len(sai)*100:.1f}%)")

    return sai


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
