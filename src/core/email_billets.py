from __future__ import annotations
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
import imaplib
import email
from email.header import decode_header
import openpyxl
import pdfplumber
from pypdf import PdfReader
from .config import settings

def _limpar_nome(nome: str) -> str:
    nome = nome.replace("\r","").replace("\n","").strip()
    return re.sub(r'[\\/*?:"<>|]', "_", nome)

def _pdf_protegido(path: Path) -> bool:
    try:
        return PdfReader(str(path)).is_encrypted
    except Exception:
        return True

def _extrair_vencimento(texto: str) -> str:
    for m in re.findall(r"\d{2}/\d{2}/\d{4}", texto):
        try:
            dt = datetime.strptime(m, "%d/%m/%Y")
            if datetime(2020,1,1) < dt < datetime(2100,1,1):
                return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
    return ""

def baixar_boletos_por_email(dt_ini: datetime, dt_fim: datetime, *, save_xlsx: Path = Path("boletos_email.xlsx")) -> Path:
    if not settings.email_user or not settings.email_pass or not settings.imap_server:
        raise RuntimeError("Credenciais IMAP ausentes no .env")

    mail = imaplib.IMAP4_SSL(settings.imap_server, settings.imap_port)
    mail.login(settings.email_user, settings.email_pass)
    mail.select("inbox")

    since = dt_ini.strftime("%d-%b-%Y")
    before = (dt_fim + timedelta(days=1)).strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'SINCE {since} BEFORE {before}')
    ids = messages[0].split()

    pdf_dir = Path("pdf_email"); pdf_dir.mkdir(exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Boletos"
    ws.append(["Nome do Arquivo", "Tipo", "Valor", "Vencimento"])

    for num in ids:
        _, msg_data = mail.fetch(num, "(RFC822)")
        for resp in msg_data:
            if not isinstance(resp, tuple):
                continue
            msg = email.message_from_bytes(resp[1])
            for part in msg.walk():
                cdisp = (part.get("Content-Disposition") or "")
                if "attachment" not in cdisp.lower():
                    continue
                filename = part.get_filename()
                if not filename or not filename.lower().endswith(".pdf"):
                    continue
                filename = _limpar_nome(filename)
                path = pdf_dir / filename
                path.write_bytes(part.get_payload(decode=True))
                if _pdf_protegido(path):
                    continue
                with pdfplumber.open(path) as pdf:
                    texto = "\n".join((p.extract_text() or "") for p in pdf.pages)
                tipo = "Boleto" if "boleto" in texto.lower() else "Outro"
                valores = re.findall(r"R\$\s?([0-9.,]+)", texto)
                venc = _extrair_vencimento(texto)
                ws.append([filename, tipo, valores[0] if valores else "", venc])

    wb.save(save_xlsx)
    mail.logout()
    return save_xlsx
