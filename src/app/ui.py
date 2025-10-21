from __future__ import annotations
import asyncio
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog
from tkcalendar import DateEntry
from datetime import datetime
from threading import Thread
from PIL import Image
import pandas as pd

from src.core.config import settings
from src.core.egestor import obter_token, listar_codigos_nfe, baixar_lote_nfes
from src.core.xml_utils import extrair_itens_xml, extrair_resumo_xml
from src.core.extratos import carregar_e_tratar
from src.core.reconcile import conciliar
from src.core.email_billets import baixar_boletos_por_email

# -------------- APP LOOK --------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

janela = ctk.CTk()
janela.geometry("1000x800")
janela.title("Analisador Financeiro - M&H Soluções")
try:
    janela.iconbitmap("logo.ico")
except Exception:
    pass

# logo
try:
    logo_img = Image.open("LOGOMEH_convertido.png").resize((150,150))
    logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(150,150))
    ctk.CTkLabel(janela, image=logo_ctk, text="").pack(pady=(10,0))
except Exception:
    ctk.CTkLabel(janela, text="M&H Soluções", font=("Segoe UI", 22, "bold")).pack(pady=(10,0))
ctk.CTkLabel(janela, text="Analisador de NF-e, Boletos e Conciliação", font=("Segoe UI", 22, "bold")).pack(pady=(0, 10))

# log
frame_log = ctk.CTkFrame(janela); frame_log.pack(pady=5, padx=10, fill="both")
ctk.CTkLabel(frame_log, text="📜 Log:", font=("Segoe UI", 14)).pack(anchor="w", padx=10, pady=5)
log_box = ctk.CTkTextbox(frame_log, height=120); log_box.pack(padx=10, pady=5, fill="both")
log_box.configure(state="disabled")

def log(msg: str):
    def _append():
        log_box.configure(state="normal")
        log_box.insert("end", msg + "\n")
        log_box.see("end"); log_box.configure(state="disabled")
    janela.after(0, _append)

tabs = ctk.CTkTabview(janela, width=950, height=600); tabs.pack(padx=10, pady=10, fill="both", expand=True)
tab_nfe = tabs.add("NF-e eGestor")
tab_email = tabs.add("Boletos por Email")
tab_conc = tabs.add("Conciliação Bancária")

# -------------- NF-e --------------
frame_nfe = ctk.CTkFrame(tab_nfe); frame_nfe.pack(pady=30, padx=20, fill="x")
ctk.CTkLabel(frame_nfe, text="📅 Datas NF-e eGestor:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 10))
row = ctk.CTkFrame(frame_nfe); row.pack(pady=5)
ctk.CTkLabel(row, text="Início:").pack(side="left", padx=(0,5))
d_ini_nfe = DateEntry(row, locale="pt_BR", date_pattern="dd/mm/yyyy"); d_ini_nfe.pack(side="left", padx=10)
ctk.CTkLabel(row, text="Fim:").pack(side="left", padx=(10,5))
d_fim_nfe = DateEntry(row, locale="pt_BR", date_pattern="dd/mm/yyyy"); d_fim_nfe.pack(side="left", padx=10)
pb_nfe = ctk.CTkProgressBar(frame_nfe, width=400); pb_nfe.set(0); pb_nfe.pack(pady=(5,10))

def _run_async(coro):
    return Thread(target=lambda: asyncio.run(coro)).start()

def _processar_nfe():
    async def _inner():
        try:
            headers = obter_token()
            dt_ini = d_ini_nfe.get_date().strftime("%Y-%m-%d")
            dt_fim = d_fim_nfe.get_date().strftime("%Y-%m-%d")
            codigos = listar_codigos_nfe(headers, dt_ini, dt_fim)
            if not codigos:
                log("⚠️ Nenhuma NF-e encontrada no período.")
                return
            log(f"🔎 {len(codigos)} NF-e para baixar")
            xml_dir = Path("xmls"); xml_dir.mkdir(exist_ok=True)
            detalhes = await baixar_lote_nfes(headers, codigos, xml_dir)
            log(f"📥 Detalhes baixados: {len(detalhes)} | XMLs salvos em {xml_dir.resolve()}")

            # gerar Notas.xlsx (resumo) e itens_nfe_egestor.xlsx (itens)
            resumos, itens = [], []
            for p in xml_dir.glob("*.xml"):
                r = extrair_resumo_xml(p);
                if r: resumos.append(r)
                itens.extend(extrair_itens_xml(p))

            if resumos:
                pd.DataFrame(resumos).to_excel("Notas.xlsx", index=False)
                log("✅ 'Notas.xlsx' gerado.")
            if itens:
                pd.DataFrame(itens).to_excel("itens_nfe_egestor.xlsx", index=False)
                log("✅ 'itens_nfe_egestor.xlsx' gerado.")
        except Exception as e:
            log(f"❌ Erro NF-e: {e}")
    _run_async(_inner())

ctk.CTkButton(frame_nfe, text="🚀 Processar NF-e eGestor", command=_processar_nfe).pack(pady=10)

# -------------- E-mail --------------
frame_email = ctk.CTkFrame(tab_email); frame_email.pack(pady=30, padx=20, fill="x")
ctk.CTkLabel(frame_email, text="📧 Datas dos Emails:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 10))
rowe = ctk.CTkFrame(frame_email); rowe.pack(pady=5)
ctk.CTkLabel(rowe, text="Início:").pack(side="left", padx=(0,5))
d_ini_mail = DateEntry(rowe, locale="pt_BR", date_pattern="dd/mm/yyyy"); d_ini_mail.pack(side="left", padx=10)
ctk.CTkLabel(rowe, text="Fim:").pack(side="left", padx=(10,5))
d_fim_mail = DateEntry(rowe, locale="pt_BR", date_pattern="dd/mm/yyyy"); d_fim_mail.pack(side="left", padx=10)
pb_mail = ctk.CTkProgressBar(frame_email, width=400); pb_mail.set(0); pb_mail.pack(pady=(5,10))

def _baixar_boletos():
    try:
        p = baixar_boletos_por_email(d_ini_mail.get_date(), d_fim_mail.get_date())
        log(f"📥 Planilha de boletos salva em: {p.resolve()}")
    except Exception as e:
        log(f"❌ Email: {e}")

ctk.CTkButton(frame_email, text="📥 Baixar Boletos por Email", command=lambda: Thread(target=_baixar_boletos).start()).pack(pady=10)

# -------------- Conciliação --------------
frame_conc = ctk.CTkFrame(tab_conc); frame_conc.pack(pady=30, padx=20, fill="x")
ctk.CTkLabel(frame_conc, text="💰 Datas para Conciliação Bancária:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 10))
rowc = ctk.CTkFrame(frame_conc); rowc.pack(pady=5)
ctk.CTkLabel(rowc, text="Início:").pack(side="left", padx=(0,5))
d_ini_conc = DateEntry(rowc, locale="pt_BR", date_pattern="dd/mm/yyyy"); d_ini_conc.pack(side="left", padx=10)
ctk.CTkLabel(rowc, text="Fim:").pack(side="left", padx=(10,5))
d_fim_conc = DateEntry(rowc, locale="pt_BR", date_pattern="dd/mm/yyyy"); d_fim_conc.pack(side="left", padx=10)
pb_conc = ctk.CTkProgressBar(frame_conc, width=400); pb_conc.set(0); pb_conc.pack(pady=(5,10))

caminhos_extratos: list[str] = []

def _anexar_extratos():
    global caminhos_extratos
    caminhos_extratos = filedialog.askopenfilenames(
        title="Selecione os extratos bancários",
        filetypes=[("Arquivos Excel", "*.xls *.xlsx *.csv")]
    )
    for p in caminhos_extratos:
        log(f"📎 Extrato anexado: {Path(p).name}")

def _iniciar_conciliacao():
    async def _inner():
        try:
            headers = obter_token()
            # Baixar NF-e no período selecionado
            dt_ini = d_ini_conc.get_date().strftime("%Y-%m-%d")
            dt_fim = d_fim_conc.get_date().strftime("%Y-%m-%d")
            codigos = listar_codigos_nfe(headers, dt_ini, dt_fim)
            if not codigos:
                log("⚠️ Nenhuma NF-e no período para conciliação.")
                return
            xml_dir = Path("xmls")
            await baixar_lote_nfes(headers, codigos, xml_dir)

            # Gerar Notas.xlsx (resumo) a partir dos XMLs
            resumos = []
            for p in xml_dir.glob("*.xml"):
                r = extrair_resumo_xml(p)
                if r:
                    resumos.append(r)
            if not resumos:
                log("⚠️ Sem XML válido para conciliação.")
                return
            pd.DataFrame(resumos).to_excel("Notas.xlsx", index=False)

            # Tratar extratos anexados
            if not caminhos_extratos:
                log("⚠️ Anexe pelo menos um extrato.")
                return
            carregar_e_tratar(caminhos_extratos)

            # Conciliação
            out = conciliar()
            log(f"✅ Conciliação concluída. Arquivo: {out.resolve()}")
        except Exception as e:
            log(f"❌ Conciliação: {e}")
    _run_async(_inner())

btns = ctk.CTkFrame(frame_conc); btns.pack(pady=(10, 0))
ctk.CTkButton(btns, text="📎 Anexar Extrato", command=_anexar_extratos).pack(side="left", padx=10)
ctk.CTkButton(btns, text="✅ Iniciar", command=_iniciar_conciliacao).pack(side="left", padx=10)

# -------------- BOOT LOG --------------
janela.after(100, lambda: log("🔍 Sistema iniciado com sucesso!"))
janela.after(200, lambda: log("💡 Use as abas para cada processo."))

janela.mainloop()
