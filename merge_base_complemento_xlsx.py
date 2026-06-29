import os
import re
import sys

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# ── Caminhos ──────────────────────────────────────────────
DOCS_DIR   = "docs"
BASE_CSV   = os.path.join(DOCS_DIR, "dados_base.csv")
COMP_CSV   = os.path.join(DOCS_DIR, "dados_complemento.csv")
SAIDA_XLSX = os.path.join(DOCS_DIR, "dados_base_com_complemento.xlsx")

# Colunas-chave de cada planilha
OC_BASE    = "ORDEM DE COMPRA"        # base: numero da ordem de compra
COD_BASE   = "CÓD PROD/SERV"          # base: codigo do produto/servico
OC_COMP    = "Nº Ordem Compra"        # complemento: numero da ordem de compra
COD_COMP   = "Serviço"                # complemento: codigo do servico
SEQ_COMP   = "Seq."                   # complemento: sequencia do item dentro da OC
COL_COMPL  = "Complemento da Descrição"  # coluna que sera trazida para a base
# ──────────────────────────────────────────────────────────


def normaliza_chave(valor):
    """Normaliza codigos para comparacao: remove '.0' de floats exportados
    e zeros a esquerda (base usa '0025', sistema usa '25')."""
    if pd.isna(valor):
        return None
    texto = re.sub(r"\.0+$", "", str(valor).strip())
    if texto == "":
        return None
    return texto.lstrip("0") or "0"


def limpa_artefato_float(valor):
    """Remove o sufixo '.0' deixado pela exportacao Excel->CSV em valores
    inteiros (ex.: '145113.0' -> '145113'), preservando decimais reais
    e codigos com zeros a esquerda (ex.: '0025' permanece '0025')."""
    if pd.isna(valor):
        return ""
    texto = str(valor)
    if re.fullmatch(r"-?\d+\.0+", texto):
        return re.sub(r"\.0+$", "", texto)
    return texto


def agrega_complementos(serie):
    """Concatena os complementos distintos de uma mesma chave (OC + codigo),
    preservando a ordem de sequencia e sem repetir textos iguais."""
    vistos = []
    for valor in serie:
        if pd.isna(valor):
            continue
        texto = str(valor).strip()
        if texto and texto not in vistos:
            vistos.append(texto)
    return " | ".join(vistos)


def main():
    print(f"Lendo base: {BASE_CSV}")
    base = pd.read_csv(BASE_CSV, encoding="utf-8-sig", dtype=str)
    print(f"  {len(base)} linhas x {len(base.columns)} colunas")

    print(f"Lendo complemento: {COMP_CSV}")
    comp = pd.read_csv(COMP_CSV, encoding="utf-8-sig", dtype=str)
    print(f"  {len(comp)} linhas x {len(comp.columns)} colunas")

    # Chaves normalizadas para o cruzamento (nao alteram as colunas de saida)
    base["_oc"]  = base[OC_BASE].map(normaliza_chave)
    base["_cod"] = base[COD_BASE].map(normaliza_chave)
    comp["_oc"]  = comp[OC_COMP].map(normaliza_chave)
    comp["_cod"] = comp[COD_COMP].map(normaliza_chave)

    # Ordena o complemento por sequencia para preservar a ordem dos itens
    comp["_seq"] = pd.to_numeric(comp[SEQ_COMP], errors="coerce")
    comp = comp.sort_values(["_oc", "_cod", "_seq"], kind="stable")

    # Agrega o complemento por (OC + codigo do item)
    mapa = (
        comp.groupby(["_oc", "_cod"])[COL_COMPL]
        .apply(agrega_complementos)
    )
    mapa = mapa[mapa != ""]
    dic_complemento = mapa.to_dict()
    print(f"Chaves (OC + codigo) com complemento: {len(dic_complemento)}")

    # Cruza com a base (LEFT join: mantem todas as linhas da base)
    chaves = list(zip(base["_oc"], base["_cod"]))
    base[COL_COMPL] = [dic_complemento.get(chave, "") for chave in chaves]

    # Limpa artefato '.0' das colunas originais (mantem codigos com zeros)
    colunas_originais = [c for c in base.columns if not c.startswith("_") and c != COL_COMPL]
    for coluna in colunas_originais:
        base[coluna] = base[coluna].map(limpa_artefato_float)

    # Monta a saida: colunas originais + a nova coluna ao final
    saida = base[colunas_originais + [COL_COMPL]]

    # Relatorio de cobertura
    preenchidas = (saida[COL_COMPL].astype(str).str.strip() != "").sum()
    chaves_base = set(zip(base["_oc"], base["_cod"]))
    sem_corresp = sum(1 for k in dic_complemento if k not in chaves_base)
    print("\n--- Cobertura do cruzamento ---")
    print(f"  Linhas da base com complemento : {preenchidas} de {len(saida)} "
          f"({100*preenchidas/len(saida):.1f}%)")
    print(f"  Chaves do complemento sem correspondencia na base: {sem_corresp}")

    os.makedirs(DOCS_DIR, exist_ok=True)
    saida.to_excel(SAIDA_XLSX, index=False, engine="openpyxl")
    print(f"\nConcluido: '{SAIDA_XLSX}' gerado com {len(saida)} linhas "
          f"x {len(saida.columns)} colunas.")


if __name__ == "__main__":
    main()
