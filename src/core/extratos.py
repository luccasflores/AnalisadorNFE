from __future__ import annotations
from pathlib import Path
from typing import Iterable, List
import pandas as pd

def _ler_arquivo(caminho: Path) -> pd.DataFrame:
    ext = caminho.suffix.lower()
    if ext == ".xls":
        return pd.read_excel(caminho, engine="xlrd", header=None)
    if ext == ".xlsx":
        return pd.read_excel(caminho, engine="openpyxl", header=None)
    if ext == ".csv":
        return pd.read_csv(caminho, sep=";", header=None, encoding="utf-8")
    raise ValueError(f"Formato não suportado: {ext}")

def carregar_e_tratar(caminhos: Iterable[str | Path]) -> pd.DataFrame:
    """Une extratos e retorna somente créditos normalizados."""
    frames: List[pd.DataFrame] = []
    for c in caminhos:
        raw = _ler_arquivo(Path(c))
        # detecta linha de cabeçalho pela presença de 'Data'
        header_idx = None
        for i, row in raw.iterrows():
            if row.astype(str).str.contains("Data", case=False).any():
                header_idx = i
                break
        if header_idx is None:
            continue
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = raw.iloc[header_idx].astype(str).str.strip().tolist()

        # pega colunas principais se existirem
        possiveis = ["Data", "Lançamento", "Dcto.", "Crédito (R$)", "Credito (R$)", "Crédito", "Credito"]
        cols = [c for c in possiveis if c in df.columns]
        df = df[cols].copy()

        # normaliza col crédito
        credito_col = next((c for c in df.columns if "créd" in c.lower() or "cred" in c.lower()), None)
        if not credito_col:
            continue
        df = df[df[credito_col].notna()]
        df[credito_col] = (
            df[credito_col]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df[credito_col] = pd.to_numeric(df[credito_col], errors="coerce")
        df = df[df[credito_col] > 0]

        # normaliza Data
        if "Data" in df.columns:
            df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)

        frames.append(df)

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    out.to_excel("extratos_consolidados_tratados.xlsx", index=False)
    return out
