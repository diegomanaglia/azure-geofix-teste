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
SAIDA_XLSX = os.path.join(DOCS_DIR, "itens_pedido_conciliados.xlsx")
# ──────────────────────────────────────────────────────────


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
    """Remove o sufixo '.0' deixado pela exportacao Excel->CSV em inteiros,
    preservando decimais reais e codigos com zeros a esquerda."""
    if pd.isna(valor):
        return ""
    texto = str(valor)
    if re.fullmatch(r"-?\d+\.0+", texto):
        return re.sub(r"\.0+$", "", texto)
    return texto


def main():
    print(f"Lendo base de NFs: {BASE_CSV}")
    base = pd.read_csv(BASE_CSV, encoding="utf-8-sig", dtype=str)
    print(f"  {len(base)} linhas x {len(base.columns)} colunas")

    print(f"Lendo itens do pedido (complemento): {COMP_CSV}")
    comp = pd.read_csv(COMP_CSV, encoding="utf-8-sig", dtype=str)
    print(f"  {len(comp)} linhas x {len(comp.columns)} colunas")

    # Chaves normalizadas (nao alteram as colunas de saida)
    base["_oc"]  = base["ORDEM DE COMPRA"].map(normaliza_chave)
    base["_cod"] = base["CÓD PROD/SERV"].map(normaliza_chave)
    comp["_oc"]  = comp["Nº Ordem Compra"].map(normaliza_chave)
    comp["_cod"] = comp["Serviço"].map(normaliza_chave)

    # ── Enriquecimento 1: nome da categoria do servico (por codigo) ──
    # Usa a descricao mais frequente de cada codigo na base de NFs.
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

    rat_val  = pd.to_numeric(comp["Valor do Rateio"], errors="coerce")
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
    print("\n--- Resumo da planilha gerada ---")
    print(f"  Itens do pedido (linhas)            : {len(sai)}")
    print(f"  Categoria do serviço preenchida     : {(sai['Categoria do Serviço']!='').sum()}")
    print(f"  Itens que constam na base de NF     : {n_consta} ({100*n_consta/len(sai):.1f}%)")

    os.makedirs(DOCS_DIR, exist_ok=True)
    sai.to_excel(SAIDA_XLSX, index=False, engine="openpyxl")
    print(f"\nConcluido: '{SAIDA_XLSX}' gerado com {len(sai)} linhas x {len(sai.columns)} colunas.")


if __name__ == "__main__":
    main()
