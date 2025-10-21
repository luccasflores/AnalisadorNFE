from __future__ import annotations
from pathlib import Path
import pandas as pd

def conciliar(notas_xlsx: Path = Path("Notas.xlsx"), extratos_xlsx: Path = Path("extratos_consolidados_tratados.xlsx")) -> Path:
    notas = pd.read_excel(notas_xlsx)
    extr = pd.read_excel(extratos_xlsx)

    notas["Valor Total"] = pd.to_numeric(notas["Valor Total da Nota"], errors="coerce")
    notas["Data Emissão"] = pd.to_datetime(notas["Data Emissão"], errors="coerce", dayfirst=True)
    notas["Conciliado"] = False
    notas["Data Pagamento"] = ""
    notas["Valor Pago"] = ""
    notas["Lançamento"] = ""
    notas["Diferença"] = ""
    notas["Justificativa"] = ""

    credito_col = next((c for c in extr.columns if "créd" in c.lower() or "cred" in c.lower()), "Crédito (R$)")
    extr = extr[extr[credito_col].notna() & (extr[credito_col] > 0)]
    if "Data" in extr.columns:
        extr["Data"] = pd.to_datetime(extr["Data"], errors="coerce", dayfirst=True)

    pagamentos = extr.copy().sort_values(by="Data", kind="stable")

    # regra 1: matching 1:1 com tolerância
    for i, row in notas.iterrows():
        candidatos = pagamentos[pagamentos["Data"] >= row["Data Emissão"]]
        ok = False
        for _, pg in candidatos.iterrows():
            diff = abs(float(pg[credito_col]) - float(row["Valor Total"]))
            if diff <= 0.04:
                notas.loc[i, ["Conciliado","Data Pagamento","Valor Pago","Lançamento","Diferença","Justificativa"]] = [
                    True,
                    pg["Data"].strftime("%d/%m/%Y") if pd.notna(pg["Data"]) else "",
                    pg[credito_col],
                    str(pg.get("Lançamento","")),
                    diff,
                    "Match 1:1 com tolerância ≤ 0,04",
                ]
                ok = True
                break
        if ok:
            continue

    # (extensões futuras: agrupamento por beneficiário, soma de lançamentos, etc.)

    out = Path("resultado_conciliacao_bancaria.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl", date_format="DD/MM/YYYY") as w:
        notas[notas["Conciliado"]].to_excel(w, sheet_name="Notas Conciliadas", index=False)
        notas[~notas["Conciliado"]].to_excel(w, sheet_name="Notas Não Conciliadas", index=False)
        extr.to_excel(w, sheet_name="Extrato Tratado", index=False)
    return out
