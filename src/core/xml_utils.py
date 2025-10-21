from __future__ import annotations
import re
from pathlib import Path
from typing import Dict, List, Optional
import xml.etree.ElementTree as ET
from dateutil.parser import parse

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

def _f(v: Optional[str]) -> float:
    if not v:
        return 0.0
    v = str(v).replace(".", "").replace(",", ".")
    try:
        return float(v)
    except Exception:
        return 0.0

def extrair_itens_xml(xml_path: Path) -> List[dict]:
    tree = ET.parse(xml_path)
    root = tree.getroot()

    dest = root.find(".//nfe:dest", NS)
    cnpj_dest = dest.findtext("nfe:CNPJ", default="", namespaces=NS) if dest is not None else ""
    nome_dest = dest.findtext("nfe:xNome", default="", namespaces=NS) if dest is not None else ""
    inf_adic = root.findtext(".//nfe:infAdic/nfe:infCpl", default="", namespaces=NS) or ""

    # secretaria
    secretaria = ""
    m = re.search(r"SECRETARIA\s+(?:MUNICIPAL|ESTADUAL)?\s*DE\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
    if not m:
        m = re.search(r"SEC\.\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
    if m:
        secretaria = m.group(1).strip().upper()

    # empenho / ordem
    empenho = ""
    m = re.search(r"(?:EMP(?:ENHO)?\s*[:\-]?\s*)([A-Z]{0,3}\s*\d{2,5}/?\d{0,4}[A-Z0-9]*)", inf_adic, re.IGNORECASE)
    if m:
        empenho = m.group(1).strip()
    else:
        mo = re.search(r"ORDEM\s+DE\s+COMPRA\s*[:\-]?\s*(\d{1,5}/\d{4})", inf_adic, re.IGNORECASE)
        if mo:
            empenho = mo.group(1).strip()

    ide = root.find(".//nfe:ide", NS)
    numero_nota = ide.findtext("nfe:nNF", default="", namespaces=NS) if ide is not None else ""
    data_emissao_raw = ide.findtext("nfe:dhEmi", default="", namespaces=NS) if ide is not None else ""
    data_emissao = parse(data_emissao_raw).strftime("%d/%m/%Y") if data_emissao_raw else ""

    itens: List[dict] = []
    for det in root.findall(".//nfe:det", NS):
        prod = det.find("nfe:prod", NS)
        if prod is None:
            continue
        q = _f(prod.findtext("nfe:qCom", default="", namespaces=NS))
        vu = _f(prod.findtext("nfe:vUnCom", default="", namespaces=NS))
        vt = round(q * vu, 2)
        itens.append({
            "Arquivo XML": xml_path.name,
            "CNPJ Destinatário": cnpj_dest,
            "Nome Destinatário": nome_dest,
            "Empenho": empenho,
            "Secretaria": secretaria,
            "Número NF": numero_nota,
            "Data Emissão": data_emissao,
            "Código do Produto": prod.findtext("nfe:cProd", default="", namespaces=NS),
            "Descrição": prod.findtext("nfe:xProd", default="", namespaces=NS),
            "Quantidade": q,
            "Valor Unitário": vu,
            "Valor Total": vt,
            "Unidade": prod.findtext("nfe:uCom", default="", namespaces=NS),
            "NCM": prod.findtext("nfe:NCM", default="", namespaces=NS),
            "CFOP": prod.findtext("nfe:CFOP", default="", namespaces=NS)
        })
    return itens

def extrair_resumo_xml(xml_path: Path) -> Optional[dict]:
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        dest = root.find(".//nfe:dest", NS)
        cnpj_dest = dest.findtext("nfe:CNPJ", default="", namespaces=NS) if dest is not None else ""
        nome_dest = dest.findtext("nfe:xNome", default="", namespaces=NS) if dest is not None else ""
        inf_adic = root.findtext(".//nfe:infAdic/nfe:infCpl", default="", namespaces=NS) or ""

        secretaria = ""
        m = re.search(r"SECRETARIA\s+(?:MUNICIPAL|ESTADUAL)?\s*DE\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
        if not m:
            m = re.search(r"SEC\.\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
        if m:
            secretaria = m.group(1).strip().upper()

        empenho = ""
        m = re.search(r"(?:EMP(?:ENHO)?\s*[:\-]?\s*)([A-Z]{0,3}\s*\d{2,5}/?\d{0,4}[A-Z0-9]*)", inf_adic, re.IGNORECASE)
        if m:
            empenho = m.group(1).strip()
        else:
            mo = re.search(r"ORDEM\s+DE\s+COMPRA\s*[:\-]?\s*(\d{1,5}/\d{4})", inf_adic, re.IGNORECASE)
            if mo:
                empenho = mo.group(1).strip()

        ide = root.find(".//nfe:ide", NS)
        numero_nota = ide.findtext("nfe:nNF", default="", namespaces=NS) if ide is not None else ""
        data_emissao_raw = ide.findtext("nfe:dhEmi", default="", namespaces=NS) if ide is not None else ""
        data_emissao = parse(data_emissao_raw).strftime("%d/%m/%Y") if data_emissao_raw else ""

        total_nf = 0.0
        for det in root.findall(".//nfe:det", NS):
            prod = det.find("nfe:prod", NS)
            if prod is None:
                continue
            q = _f(prod.findtext("nfe:qCom", default="", namespaces=NS))
            vu = _f(prod.findtext("nfe:vUnCom", default="", namespaces=NS))
            total_nf += round(q * vu, 2)

        return {
            "Arquivo XML": xml_path.name,
            "CNPJ Destinatário": cnpj_dest,
            "Nome Destinatário": nome_dest,
            "Empenho": empenho,
            "Secretaria": secretaria,
            "Número NF": numero_nota,
            "Data Emissão": data_emissao,
            "Valor Total da Nota": round(total_nf, 2),
        }
    except Exception:
        return None
