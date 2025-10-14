
import os
import re
from decimal import Decimal
import time
import asyncio
import aiohttp
import pandas as pd
import requests
import xml.etree.ElementTree as ET
import customtkinter as ctk
from tkinter import filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import imaplib
import email
from email.header import decode_header
import openpyxl
import pdfplumber
from pypdf import PdfReader
from email.utils import parsedate_to_datetime
from dateutil.parser import parse
from PIL import Image, ImageTk
from threading import Thread
from itertools import combinations

total_codigos = 0
progresso_atual = 0

# ========== CONFIGURAÇÃO ==========
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

janela = ctk.CTk()
janela.geometry("1000x800")
janela.title("Analisador Financeiro - M&H Soluções")
janela.iconbitmap("logo.ico")  # Adiciona ícone na janela

# ========== LOGO ==========
logo_img = Image.open("LOGOMEH_convertido.png")
logo_img = logo_img.resize((150, 150))
logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(150, 150))
ctk.CTkLabel(janela, image=logo_ctk, text="").pack(pady=(10, 0))
ctk.CTkLabel(janela, text="Analisador de NF-e, Boletos e Conciliação", font=("Segoe UI", 22, "bold")).pack(pady=(0, 10))

# ========== LOG FINAL (POSICIONADO NO TOPO) ==========
frame_log = ctk.CTkFrame(janela)
frame_log.pack(pady=5, padx=10, fill="both", expand=False)
ctk.CTkLabel(frame_log, text="📜 Log:", font=("Segoe UI", 14)).pack(anchor="w", padx=10, pady=5)
log_texto = ctk.CTkTextbox(frame_log, height=120)
log_texto.pack(padx=10, pady=5, fill="both", expand=True)
log_texto.configure(state="disabled")

# ========== FUNÇÃO DE LOG CORRIGIDA ==========
def log(msg):
    try:
        log_texto.configure(state="normal")
        log_texto.insert("end", msg + "\n")
        log_texto.see("end")
        log_texto.configure(state="disabled")
    except Exception as e:
        print(f"Erro ao escrever no log: {e}")

# ========== GLOBAIS ==========
personal_token = "token"
headers = {}
access_token = ""
caminho_extrato = ""
codigos_nfe = []
itens_extraidos_global = []

ultima_requisicao = time.monotonic()
ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

# ========== ABAS ==========
tabs = ctk.CTkTabview(janela, width=950, height=600)
tabs.pack(padx=10, pady=10, fill="both", expand=True)
tab_nfe = tabs.add("NF-e eGestor")
tab_email = tabs.add("Boletos por Email")
tab_conc = tabs.add("Conciliação Bancária")
# ========== CONCILIAÇÃO BANCÁRIA ==========
frame_conc = ctk.CTkFrame(tab_conc)
frame_conc.pack(pady=30, padx=20, fill="x")

ctk.CTkLabel(frame_conc, text="💰 Datas para Conciliação Bancária:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 10))

data_frame_conc = ctk.CTkFrame(frame_conc)
data_frame_conc.pack(pady=5)

ctk.CTkLabel(data_frame_conc, text="Início:").pack(side="left", padx=(0, 5))
data_inicio_conc = DateEntry(data_frame_conc, locale="pt_BR", date_pattern="dd/mm/yyyy")
data_inicio_conc.pack(side="left", padx=10)

ctk.CTkLabel(data_frame_conc, text="Fim:").pack(side="left", padx=(10, 5))
data_fim_conc = DateEntry(data_frame_conc, locale="pt_BR", date_pattern="dd/mm/yyyy")
data_fim_conc.pack(side="left", padx=10)

# Barra de progresso para conciliação
progress_conc = ctk.CTkProgressBar(frame_conc, width=400)
progress_conc.set(0)
progress_conc.pack(pady=(5, 10))

# Botão de conciliação
# ctk.CTkButton(
#     frame_conc,
#     text="🔍 Iniciar Conciliação",
#     command=lambda: log("⚠️ Funcionalidade de conciliação desativada.")
# ).pack(pady=10)

# Botões adicionais
botoes_frame_conc = ctk.CTkFrame(frame_conc)
botoes_frame_conc.pack(pady=(10, 0))
# Substitua esta função:
def anexar_extratos():
    global caminhos_extratos  # ← isso aqui é essencial
    caminhos_extratos = filedialog.askopenfilenames(
        title="Selecione os extratos bancários",
        filetypes=[("Arquivos Excel", "*.xls *.xlsx *.csv")]
    )
    for caminho in caminhos_extratos:
        log(f"📎 Extrato anexado: {os.path.basename(caminho)}")
def ler_e_tratar_extratos(lista_caminhos):
    todos_dfs = []

    for caminho in lista_caminhos:
        try:
            extensao = os.path.splitext(caminho)[-1].lower()
            if extensao == ".xls":
                df = pd.read_excel(caminho, engine="xlrd", skiprows=8)
            elif extensao == ".xlsx":
                df = pd.read_excel(caminho, engine="openpyxl", skiprows=8)
            elif extensao == ".csv":
                df = pd.read_csv(caminho, sep=";", skiprows=8)
            else:
                log(f"❌ Formato não suportado: {caminho}")
                continue

            df.columns = [col.strip() for col in df.columns]

            log(f"🔎 Colunas em {os.path.basename(caminho)}: {df.columns.tolist()}")

            # Identifica colunas principais automaticamente
            col_data = next((c for c in df.columns if 'data' in c.lower()), None)
            col_lanc = next((c for c in df.columns if 'lançamento' in c.lower()), None)
            col_cred = next((c for c in df.columns if 'crédito' in c.lower()), None)

            if not col_data or not col_lanc or not col_cred:
                log(f"❌ Colunas essenciais não encontradas em {os.path.basename(caminho)}.")
                continue

            df_filtrado = df[[col_data, col_lanc, col_cred]].copy()
            df_filtrado = df_filtrado[df_filtrado[col_cred].notna()]

            # Remove linhas com textos não numéricos
            def eh_valor_valido(valor):
                try:
                    # Remove R$, pontos e espaços e tenta converter
                    Decimal(str(valor).replace("R$", "").replace(".", "").replace(",", ".").strip())
                    return True
                except:
                    return False

            df_filtrado = df_filtrado[df_filtrado[col_cred].apply(eh_valor_valido)]

            todos_dfs.append(df_filtrado)

        except Exception as e:
            log(f"❌ Erro ao processar {caminho}: {e}")

    if todos_dfs:
        df_geral = pd.concat(todos_dfs, ignore_index=True)
        log(f"✅ {len(df_geral)} lançamentos com crédito foram encontrados.")
        return df_geral
    else:
        log("⚠️ Nenhum dado válido foi encontrado nos extratos.")
        return pd.DataFrame()


ctk.CTkButton(
    botoes_frame_conc,
    text="📎 Anexar Extrato",
    command=anexar_extratos
).pack(side="left", padx=10)

ctk.CTkButton(
    botoes_frame_conc,
    text="✅ Iniciar",
    command=lambda: Thread(target=lambda: asyncio.run(iniciar_conciliacao())).start()
).pack(side="left", padx=10)


# ========== NF-e E-GESTOR ==========
frame_nfe = ctk.CTkFrame(tab_nfe)
frame_nfe.pack(pady=30, padx=20, fill="x")

ctk.CTkLabel(frame_nfe, text="📅 Datas NF-e eGestor:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 10))

data_frame_nfe = ctk.CTkFrame(frame_nfe)
data_frame_nfe.pack(pady=5)

ctk.CTkLabel(data_frame_nfe, text="Início:").pack(side="left", padx=(0, 5))
data_inicio = DateEntry(data_frame_nfe, locale="pt_BR", date_pattern="dd/mm/yyyy")
data_inicio.pack(side="left", padx=10)

ctk.CTkLabel(data_frame_nfe, text="Fim:").pack(side="left", padx=(10, 5))
data_fim = DateEntry(data_frame_nfe, locale="pt_BR", date_pattern="dd/mm/yyyy")
data_fim.pack(side="left", padx=10)

# Barra de progresso para NF-e
tq_progress = ctk.CTkProgressBar(frame_nfe, width=400)
tq_progress.set(0)
tq_progress.pack(pady=(5, 10))

ctk.CTkButton(
    frame_nfe,
    text="🚀 Processar NF-e eGestor",
    command=lambda: Thread(target=lambda: asyncio.run(iniciar_processo_egestor())).start()
).pack(pady=10)

# ========== BOLETOS POR EMAIL ==========
frame_email = ctk.CTkFrame(tab_email)
frame_email.pack(pady=30, padx=20, fill="x")

ctk.CTkLabel(frame_email, text="📧 Datas dos Emails:", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(0, 10))

data_frame_email = ctk.CTkFrame(frame_email)
data_frame_email.pack(pady=5)

ctk.CTkLabel(data_frame_email, text="Início:").pack(side="left", padx=(0, 5))
data_inicio_email = DateEntry(data_frame_email, locale="pt_BR", date_pattern="dd/mm/yyyy")
data_inicio_email.pack(side="left", padx=10)

ctk.CTkLabel(data_frame_email, text="Fim:").pack(side="left", padx=(10, 5))
data_fim_email = DateEntry(data_frame_email, locale="pt_BR", date_pattern="dd/mm/yyyy")
data_fim_email.pack(side="left", padx=10)

# Barra de progresso para email
progress_email = ctk.CTkProgressBar(frame_email, width=400)
progress_email.set(0)
progress_email.pack(pady=(5, 10))

ctk.CTkButton(frame_email, text="📥 Baixar Boletos por Email", command=lambda: Thread(target=baixar_boletos_email).start()).pack(pady=10)


def limpar_nome_arquivo(nome):
    nome = nome.replace("\r", "").replace("\n", "").strip()
    return re.sub(r'[\\/*?:"<>|]', "_", nome)

def verificar_protecao_pdf(caminho):
    try:
        reader = PdfReader(caminho)
        return reader.is_encrypted
    except Exception:
        return True

def extrair_data_vencimento(texto):
    datas = re.findall(r"\d{2}/\d{2}/\d{4}", texto)
    for d in datas:
        try:
            dt = datetime.strptime(d, "%d/%m/%Y")
            if datetime(2020, 1, 1) < dt < datetime(2100, 1, 1):
                return dt.strftime("%d/%m/%Y")
        except:
            continue
    return ""

def baixar_boletos_email():
    IMAP_SERVER = "imap.mail.yahoo.com"
    IMAP_PORT = 993
    EMAIL_USER = "usuario@example.com"
    EMAIL_PASS = "senha"

    dt_ini = data_inicio_email.get_date()
    dt_fim = data_fim_email.get_date()
    data_inicio_str = dt_ini.strftime("%d-%b-%Y")
    data_fim_str = (dt_fim + timedelta(days=1)).strftime("%d-%b-%Y")

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, messages = mail.search(None, f'SINCE {data_inicio_str} BEFORE {data_fim_str}')
    msg_nums = messages[0].split()
    total_msgs = len(msg_nums)

    if not os.path.exists("pdf_email"):
        os.makedirs("pdf_email")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Boletos"
    ws.append(["Nome do Arquivo", "Tipo", "Valor", "Vencimento"])

    for i, num in enumerate(msg_nums):
        _, msg_data = mail.fetch(num, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                for part in msg.walk():
                    content_disposition = part.get("Content-Disposition", "")
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename and filename.lower().endswith(".pdf"):
                            filename = limpar_nome_arquivo(filename)
                            path = os.path.join("pdf_email", filename)
                            with open(path, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            if verificar_protecao_pdf(path):
                                continue
                            with pdfplumber.open(path) as pdf:
                                texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
                                tipo = "Boleto" if "boleto" in texto.lower() else "Outro"
                                valor = re.findall(r"R\$\s?([0-9.,]+)", texto)
                                venc = extrair_data_vencimento(texto)
                                ws.append([filename, tipo, valor[0] if valor else "", venc])
        progress_email.set((i + 1) / total_msgs)
        progress_email.update()

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salvar planilha de boletos",
        initialfile="boletos_email.xlsx"
    )
    if save_path:
        wb.save(save_path)
        log(f"📥 Planilha salva com sucesso em: {save_path}")
    else:
        log("⚠️ Salvamento cancelado pelo usuário.")
    mail.logout()

# (As funções de NF-e e conciliação seguem abaixo no código original, mantidas.)




def obter_token():
    global headers, access_token
    response = requests.post(
        "https://api.egestor.com.br/api/oauth/access_token",
        headers={"Content-Type": "application/json"},
        json={"grant_type": "personal", "personal_token": personal_token}
    )
    if response.status_code == 200:
        access_token = response.json()['access_token']
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        log("✅ Token obtido com sucesso")
    else:
        log("❌ Falha ao obter token")


def buscar_codigos(dt_ini, dt_fim):
    global codigos_nfe
    params_base = {
        "dtIni": dt_ini,
        "dtFim": dt_fim,
        "fields": "codigo",
        "orderBy": "codigo,asc",
        "limit": 100
    }
    codigos_nfe.clear()
    pagina = 1
    while True:
        params = params_base.copy()
        params["page"] = pagina
        response = requests.get("https://api.egestor.com.br/api/v1/nfe", headers=headers, params=params)
        if response.status_code != 200:
            log(f"❌ Erro na página {pagina}: {response.text}")
            break
        data = response.json()
        registros = data.get("data", [])
        if not registros:
            break
        codigos_nfe.extend([nfe["codigo"] for nfe in registros])
        pagina += 1
    log(f"📄 Total de códigos encontrados: {len(codigos_nfe)}")


async def fetch_detalhes(session, codigo):
    # await asyncio.sleep(0.3)  # Respeita limite de requisições

    url_detalhes = f"https://api.egestor.com.br/api/v1/nfe/{codigo}"
    url_xml = f"https://api.egestor.com.br/api/v1/nfe/{codigo}/xml"

    try:
        async with session.get(url_detalhes) as resp:
            if resp.status != 200:
                log(f"❌ Erro ao buscar detalhes da NF {codigo}: {resp.status}")
                return None
            dados = await resp.json()
    except Exception as e:
        log(f"❌ Erro ao obter detalhes da NF {codigo}: {e}")
        return None

    try:
        async with session.get(url_xml) as r:
            if r.status == 200:
                xml_bytes = await r.read()
                dados["xml"] = xml_bytes.decode("utf-8", errors="ignore")  # Adiciona o XML como string
                # --- Salvar XML como arquivo .xml ---
                nome_arquivo = f"xmls/NF_{codigo}.xml"
                os.makedirs("xmls", exist_ok=True)
                with open(nome_arquivo, "w", encoding="utf-8") as f:
                    f.write(dados["xml"])

                log(f"📄 XML carregado para NF {codigo}")
            else:
                log(f"⚠️ Não foi possível obter o XML da NF {codigo}")
                dados["xml"] = ""
    except Exception as e:
        log(f"❌ Erro ao baixar XML da NF {codigo}: {e}")
        dados["xml"] = ""

    return dados
def salvar_xml_local(codigo_nf, conteudo_xml):
    pasta = os.path.join(os.getcwd(), "xmls")
    os.makedirs(pasta, exist_ok=True)
    caminho = os.path.join(pasta, f"{codigo_nf}.xml")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(conteudo_xml)
    log(f"📁 XML salvo: {caminho}")

def extrair_itens_nfe(nfe_json):
    itens_extraidos = []
    xml_str = nfe_json.get("xml", "")
    if not xml_str:
        return itens_extraidos

    try:
        root = ET.fromstring(xml_str)

        inf_adic = root.findtext(".//nfe:infAdic/nfe:infCpl", default="", namespaces=ns)
        secretaria = ""

        match_sec = re.search(r"SEC\\.\\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\\s]+)", inf_adic)
        if not match_sec:
            match_sec = re.search(r"SECRETARIA\\s+(?:MUNICIPAL|ESTADUAL)?\\s*DE\\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\\s]+)", inf_adic)
        if match_sec:
            secretaria = match_sec.group(1).strip().upper()

        numero_empenho = ""

        # Captura tudo que vem após EMP: até encontrar espaço ou ; ou nova chave
        match = re.search(r"EMP(?:ENHO)?\s*[:\-]?\s*([A-Z]*\s*\d{2,5}/?\d{0,4}[A-Z0-9]*)", inf_adic, re.IGNORECASE)

        if match:
            numero_empenho = match.group(1).strip()
        else:
            # Captura ORDEM DE COMPRA como fallback
            match_ordem = re.search(r"ORDEM\s+DE\s+COMPRA\s*[:\-]?\s*(\d{1,5}/\d{4})", inf_adic, re.IGNORECASE)
            if match_ordem:
                numero_empenho = match_ordem.group(1).strip()

        ide = root.find(".//nfe:ide", ns)
        numero_nota = ide.findtext("nfe:nNF", default="", namespaces=ns) if ide is not None else ""
        data_emissao_raw = ide.findtext("nfe:dhEmi", default="", namespaces=ns) if ide is not None else ""
        data_emissao = ""
        if data_emissao_raw:
            try:
                data_emissao = parse(data_emissao_raw).strftime("%d/%m/%Y")
            except:
                data_emissao = data_emissao_raw

        for det in root.findall(".//{*}det"):
            prod = det.find("{*}prod")
            if prod is None:
                continue

            quantidade = prod.findtext("nfe:qCom", default="0", namespaces=ns)
            valor_unitario = prod.findtext("nfe:vUnCom", default="0", namespaces=ns)
            valor_total = prod.findtext("nfe:vProd", default="0", namespaces=ns)

            item = {
                "Arquivo XML": nfe_json.get("codigo", ""),
                "CNPJ Destinatário": nfe_json.get("destinatario", {}).get("cnpj", ""),
                "Nome Destinatário": nfe_json.get("destinatario", {}).get("nome", ""),
                "Empenho": numero_empenho,
                "Secretaria": secretaria,
                "Número NF": numero_nota,
                "Data Emissão": data_emissao,
                "Código do Produto": prod.findtext("{*}cProd") or "",
                "Descrição": prod.findtext("{*}xProd") or "",
                "Quantidade": prod.findtext("{*}qCom") or "0",
                "Valor Unitário": prod.findtext("{*}vUnCom") or "0",
                "Valor Total": prod.findtext("{*}vProd") or "0",
                "Unidade": prod.findtext("{*}uCom") or "",
                "NCM": prod.findtext("{*}NCM") or "",
                "CFOP": prod.findtext("{*}CFOP") or ""
            }
            itens_extraidos.append(item)
    except Exception as e:
        log(f"❌ Erro ao extrair itens: {e}")

    return itens_extraidos

async def fetch_detalhes(session, codigo):
    await asyncio.sleep(1.1)  # Respeita o limite de requisições da API

    url_detalhes = f"https://api.egestor.com.br/api/v1/nfe/{codigo}"
    url_xml = f"https://api.egestor.com.br/api/v1/nfe/{codigo}/xml"

    try:
        async with session.get(url_detalhes) as resp:
            if resp.status != 200:
                log(f"❌ Erro ao buscar detalhes da NF {codigo}: {resp.status}")
                return None
            dados = await resp.json()
    except Exception as e:
        log(f"❌ Erro ao obter detalhes da NF {codigo}: {e}")
        return None

    try:
        async with session.get(url_xml) as r:
            if r.status == 200:
                xml_bytes = await r.read()
                xml_str = xml_bytes.decode("utf-8", errors="ignore")
                dados["xml"] = xml_str
                salvar_xml_local(codigo, xml_str)
            else:
                log(f"⚠️ Não foi possível obter o XML da NF {codigo}")
                dados["xml"] = ""
    except Exception as e:
        log(f"❌ Erro ao baixar XML da NF {codigo}: {e}")
        dados["xml"] = ""

    return dados

# ========== CONCILIAÇÃO BANCÁRIA - LÓGICA COMPLETA ==========

async def iniciar_conciliacao():
    try:
        dt_ini = data_inicio_conc.get_date().strftime("%Y-%m-%d")
        dt_fim = data_fim_conc.get_date().strftime("%Y-%m-%d")
        obter_token()
        buscar_codigos(dt_ini, dt_fim)

        global total_codigos, progresso_atual
        total_codigos = len(codigos_nfe)
        progresso_atual = 0
        progress_conc.set(0)

        await processar_notas()

        # NOVO: tratamento automático dos extratos
        tratar_e_consolidar_extratos(caminhos_extratos)

        # Executa a conciliação com o extrato tratado
        realizar_conciliacao_bancaria()

        log("🎯 Conciliação finalizada com sucesso.")
    except Exception as e:
        log(f"❌ Erro na conciliação: {e}")
def tratar_e_consolidar_extratos(caminhos):
    todos_dfs = []

    for caminho in caminhos:
        try:
            # Lê sem cabeçalho para detectar a linha correta
            df_raw = pd.read_excel(caminho, sheet_name=0, header=None)
            linha_header = 8

            for i, row in df_raw.iterrows():
                if row.astype(str).str.contains("Data", case=False).any():
                    linha_header = i
                    break

            if linha_header is None:
                raise Exception("Cabeçalho com 'Data' não encontrado.")

            # Releitura com o cabeçalho real
            df = pd.read_excel(caminho, sheet_name=0, header=linha_header)

            # Remove linhas com "Total"
            df = df[df['Data'].astype(str).str.strip().str.lower() != 'total']

            # Seleciona colunas válidas
            colunas_desejadas = ['Data', 'Lançamento', 'Dcto.', 'Crédito (R$)', 'Débito (R$)', 'Saldo (R$)']
            df = df[[col for col in colunas_desejadas if col in df.columns]]

            # Conversão da coluna de crédito para número
            df['Crédito (R$)'] = (
                df['Crédito (R$)']
                .astype(str)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False)
                .apply(pd.to_numeric, errors='coerce')
            )

            # Mantém apenas linhas com crédito positivo
            df = df[df['Crédito (R$)'].notnull() & (df['Crédito (R$)'] > 0)]

            todos_dfs.append(df)

        except Exception as e:
            log(f"⚠️ Erro ao tratar extrato {caminho}: {e}")

    if not todos_dfs:
        raise Exception("Nenhum extrato foi carregado ou tratado corretamente.")

    df_consolidado = pd.concat(todos_dfs, ignore_index=True)
    df_consolidado.to_excel("extratos_consolidados_tratados.xlsx", index=False)
    return df_consolidado

async def fetch_detalhes_e_xml(session, codigo):
    async with asyncio.Semaphore(5):
        # await respeitar_limite_requisicoes()
        url_detalhes = f"https://api.egestor.com.br/api/v1/nfe/{codigo}"
        url_xml = f"https://api.egestor.com.br/api/v1/nfe/{codigo}/xml"

        nota = None
        try:
            async with session.get(url_detalhes) as resp:
                if resp.status == 200:
                    nota = await resp.json()
                else:
                    log(f"❌ Erro ao buscar detalhes {codigo}: {resp.status}")

                    return None
        except Exception as e:
            log(f"❌ Exceção em detalhes {codigo}: {e}")
            return None

        # await respeitar_limite_requisicoes()
        try:
            async with session.get(url_xml) as r:
                if r.status == 200:
                    os.makedirs("xmls", exist_ok=True)
                    content = await r.read()
                    with open(f"xmls/{codigo}.xml", "wb") as f:
                        f.write(content)
                    log(f"✅ XML salvo: {codigo}.xml")
        except Exception as e:
            log(f"❌ Erro ao baixar XML {codigo}: {e}")
        global progresso_atual
        progresso_atual += 1
        progress_conc.set(progresso_atual / total_codigos)
        return nota
def extrair_itens_xml(caminho_xml):
    def extrair_float(valor_str):
        try:
            return float(valor_str.replace(",", "."))
        except:
            return 0.0

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        itens = []

        dest = root.find(".//nfe:dest", ns)
        cnpj_dest = dest.findtext("nfe:CNPJ", default="", namespaces=ns) if dest is not None else ""
        nome_dest = dest.findtext("nfe:xNome", default="", namespaces=ns) if dest is not None else ""
        inf_adic = root.findtext(".//nfe:infAdic/nfe:infCpl", default="", namespaces=ns)

        secretaria = ""
        match_sec = re.search(r"SEC\.\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
        if not match_sec:
            match_sec = re.search(r"SECRETARIA\s+(?:MUNICIPAL|ESTADUAL)?\s*DE\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
        if match_sec:
            secretaria = match_sec.group(1).strip().upper()

        numero_empenho = ""
        match_empenho_1 = re.search(r"EMPENHO:.*?(\d{1,5}/\d{4})", inf_adic, re.IGNORECASE)
        match_empenho_2 = re.search(r"EMPENHO:\s*([A-Z0-9]{6,})", inf_adic, re.IGNORECASE)
        match_ordem = re.search(r"ORDEM\s+DE\s+COMPRA:\s*(\d{1,5}/\d{4})", inf_adic, re.IGNORECASE)
        if match_empenho_1:
            numero_empenho = match_empenho_1.group(1).strip()
        elif match_empenho_2:
            numero_empenho = match_empenho_2.group(1).strip()
        elif match_ordem:
            numero_empenho = match_ordem.group(1).strip()

        ide = root.find(".//nfe:ide", ns)
        numero_nota = ide.findtext("nfe:nNF", default="", namespaces=ns) if ide is not None else ""
        data_emissao_raw = ide.findtext("nfe:dhEmi", default="", namespaces=ns) if ide is not None else ""
        data_emissao = parse(data_emissao_raw).strftime("%d/%m/%Y") if data_emissao_raw else ""

        for det in root.findall(".//nfe:det", ns):
            prod = det.find("nfe:prod", ns)
            if prod is None:
                continue

            quantidade = extrair_float(prod.findtext("nfe:qCom", default="0", namespaces=ns))
            valor_unitario = extrair_float(prod.findtext("nfe:vUnCom", default="0", namespaces=ns))
            valor_total = round(quantidade * valor_unitario, 2)

            itens.append({
                "Arquivo XML": os.path.basename(caminho_xml),
                "CNPJ Destinatário": cnpj_dest,
                "Nome Destinatário": nome_dest,
                "Empenho": numero_empenho,
                "Secretaria": secretaria,
                "Número NF": numero_nota,
                "Data Emissão": data_emissao,
                "Código do Produto": prod.findtext("nfe:cProd", default="", namespaces=ns),
                "Descrição": prod.findtext("nfe:xProd", default="", namespaces=ns),
                "Quantidade": quantidade,
                "Valor Unitário": valor_unitario,
                "Valor Total": valor_total,
                "Unidade": prod.findtext("nfe:uCom", default="", namespaces=ns),
                "NCM": prod.findtext("nfe:NCM", default="", namespaces=ns),
                "CFOP": prod.findtext("nfe:CFOP", default="", namespaces=ns)
            })

        return itens
    except Exception as e:
        log(f"Erro ao processar {caminho_xml}: {e}")
        return []

def extrair_dados_consolidados_xml(caminho_xml):
    def extrair_float(valor_str):
        try:
            return float(valor_str.replace(",", "."))
        except:
            return 0.0

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        dest = root.find(".//nfe:dest", ns)
        cnpj_dest = dest.findtext("nfe:CNPJ", default="", namespaces=ns) if dest is not None else ""
        nome_dest = dest.findtext("nfe:xNome", default="", namespaces=ns) if dest is not None else ""
        inf_adic = root.findtext(".//nfe:infAdic/nfe:infCpl", default="", namespaces=ns)

        secretaria = ""
        match_sec = re.search(r"SEC\.\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
        if not match_sec:
            match_sec = re.search(r"SECRETARIA\s+(?:MUNICIPAL|ESTADUAL)?\s*DE\s*([A-ZÇÁÉÍÓÚÂÊÔÃÕa-zçáéíóúâêôãõ\s]+)", inf_adic)
        if match_sec:
            secretaria = match_sec.group(1).strip().upper()

        numero_empenho = ""

        # Verifica todos os formatos possíveis de empenho
        match_empenho = re.search(
            r"(?:EMP(?:ENHO)?\s*[:\-]?\s*)([A-Z]{0,3}\s*\d{2,5}/?\d{0,4}[A-Z0-9]*)",
            inf_adic, re.IGNORECASE
        )

        if match_empenho:
            numero_empenho = match_empenho.group(1).strip()
        else:
            match_ordem = re.search(r"ORDEM\s+DE\s+COMPRA\s*[:\-]?\s*(\d{1,5}/\d{4})", inf_adic, re.IGNORECASE)
            if match_ordem:
                numero_empenho = match_ordem.group(1).strip()

        ide = root.find(".//nfe:ide", ns)
        numero_nota = ide.findtext("nfe:nNF", default="", namespaces=ns) if ide is not None else ""
        data_emissao_raw = ide.findtext("nfe:dhEmi", default="", namespaces=ns) if ide is not None else ""
        data_emissao = parse(data_emissao_raw).strftime("%d/%m/%Y") if data_emissao_raw else ""

        valor_total_nf = 0.0
        for det in root.findall(".//nfe:det", ns):
            prod = det.find("nfe:prod", ns)
            if prod is None:
                continue
            quantidade = extrair_float(prod.findtext("nfe:qCom", default="0", namespaces=ns))
            valor_unitario = extrair_float(prod.findtext("nfe:vUnCom", default="0", namespaces=ns))
            valor_total_nf += round(quantidade * valor_unitario, 2)

        return {
            "Arquivo XML": os.path.basename(caminho_xml),
            "CNPJ Destinatário": cnpj_dest,
            "Nome Destinatário": nome_dest,
            "Empenho": numero_empenho,
            "Secretaria": secretaria,
            "Número NF": numero_nota,
            "Data Emissão": data_emissao,
            "Valor Total da Nota": round(valor_total_nf, 2)
        }

    except Exception as e:
        log(f"❌ Erro ao processar {caminho_xml}: {e}")
        return None

async def processar_notas():
    connector = aiohttp.TCPConnector(limit=10)
    timeout = aiohttp.ClientTimeout(total=30)  # até 30 segundos por requisição
    async with aiohttp.ClientSession(headers=headers, connector=connector, timeout=timeout) as session:
        tasks = [fetch_detalhes_e_xml(session, codigo) for codigo in codigos_nfe]
        resultados = await asyncio.gather(*tasks)
        detalhes = [r for r in resultados if r]
        if detalhes:
            df = pd.json_normalize(detalhes)


    todos_itens = []
    for nome_arquivo in os.listdir("xmls"):
        if nome_arquivo.endswith(".xml"):
            caminho = os.path.join("xmls", nome_arquivo)
            dados_nf = extrair_dados_consolidados_xml(caminho)
            if dados_nf:
                todos_itens.append(dados_nf)

    if todos_itens:
        df_itens = pd.DataFrame(todos_itens)
        df_itens.to_excel("Notas.xlsx", index=False)
        log("✅ Itens extraídos e salvos em 'Notas.xlsx'")
    else:
        log("⚠️ Nenhum item extraído dos XMLs.")

import os
import pandas as pd
from decimal import Decimal

def tratar_extratos_para_conciliacao(lista_arquivos):
    extratos_tratados = []

    for caminho in lista_arquivos:
        extensao = os.path.splitext(caminho)[-1].lower()
        try:
            if extensao == ".xls":
                df = pd.read_excel(caminho, engine="xlrd", skiprows=8)
            elif extensao == ".xlsx":
                df_raw = pd.read_excel(caminho, engine="openpyxl", header=None)
                header_index = None
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains("Data", case=False).any():
                        header_index = i
                        break
                if header_index is not None:
                    df = pd.read_excel(caminho, engine="openpyxl", skiprows=header_index)
                else:
                    continue
            elif extensao == ".csv":
                df = pd.read_csv(caminho, sep=";", encoding="utf-8")
            else:
                continue

            df.columns = [str(c).strip() for c in df.columns]

            # Remover linhas com valores como "Total", cabeçalhos repetidos ou sem crédito
            df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains("data|total|últimos lançamentos", na=False).any(), axis=1)]

            col_credito = next((col for col in df.columns if "crédito" in col.lower()), None)
            if col_credito:
                df = df[df[col_credito].notna()]  # Manter apenas onde há crédito
                df[col_credito] = df[col_credito].apply(lambda x: str(x).replace(".", "").replace(",", "."))
                df[col_credito] = df[col_credito].apply(lambda x: float(x) if x.replace(".", "", 1).isdigit() else None)

            # Manter colunas principais, se existirem
            colunas_desejadas = ["Data", "Lançamento", "Dcto.", col_credito]
            df = df[[col for col in colunas_desejadas if col in df.columns]]

            extratos_tratados.append(df)

        except Exception as e:
            print(f"Erro ao processar {caminho}: {e}")

    if extratos_tratados:
        df_geral = pd.concat(extratos_tratados, ignore_index=True)
        df_geral.to_excel("extratos_consolidados_tratados.xlsx", index=False)
        return df_geral

    return pd.DataFrame()
import pandas as pd
from datetime import datetime
import os

def realizar_conciliacao_bancaria():
    try:
        log("🔎 Iniciando conciliação bancária...")

        # Carrega os arquivos
        notas_df = pd.read_excel("Notas.xlsx")
        extratos_df = pd.read_excel("extratos_consolidados_tratados.xlsx")

        # Prepara dados
        notas_df["Valor Total"] = pd.to_numeric(notas_df["Valor Total da Nota"], errors="coerce")
        notas_df["Data Emissão"] = pd.to_datetime(notas_df["Data Emissão"], errors="coerce", dayfirst=True)
        notas_df["Conciliado"] = False
        notas_df["Data Pagamento"] = ""
        notas_df["Valor Pago"] = ""
        notas_df["Lançamento"] = ""
        notas_df["Diferença"] = ""
        notas_df["Justificativa"] = ""

        extratos_df = extratos_df[extratos_df["Crédito (R$)"].notna()]
        extratos_df = extratos_df[extratos_df["Crédito (R$)"] > 0]
        extratos_df["Data"] = pd.to_datetime(extratos_df["Data"], errors="coerce", dayfirst=True)

        pagamentos = extratos_df[["Data", "Lançamento", "Dcto.", "Crédito (R$)"]].dropna()
        pagamentos = pagamentos.sort_values("Data")

        # Conciliação
        for i, nota in notas_df.iterrows():
            data_emissao = nota["Data Emissão"]
            valor_nota = nota["Valor Total"]
            nome_dest = nota["Nome Destinatário"]
            candidatos = pagamentos[pagamentos["Data"] >= data_emissao]

            conciliado = False
            for j, pg in candidatos.iterrows():
                valor_pgto = pg["Crédito (R$)"]
                diff = abs(valor_pgto - valor_nota)
                if diff <= 0.04:
                    notas_df.at[i, "Conciliado"] = True
                    notas_df.at[i, "Data Pagamento"] = pg["Data"].strftime("%d/%m/%Y")
                    notas_df.at[i, "Valor Pago"] = valor_pgto
                    notas_df.at[i, "Lançamento"] = pg["Lançamento"]
                    notas_df.at[i, "Diferença"] = diff
                    notas_df.at[i, "Justificativa"] = "Valor total conciliado com diferença ≤ 0,04"
                    conciliado = True
                    break

            if not conciliado:
                subset = notas_df[
                    (notas_df["Nome Destinatário"] == nome_dest) &
                    (~notas_df["Conciliado"])
                ]
                total_a_conciliar = subset["Valor Total"].sum()
                for k in range(len(candidatos)):
                    grupo = candidatos.iloc[k:k + len(subset)]
                    if len(grupo) < 1:
                        continue
                    soma_pg = grupo["Crédito (R$)"].sum()
                    if abs(soma_pg - total_a_conciliar) <= 0.04:
                        for idx in subset.index:
                            notas_df.at[idx, "Conciliado"] = True
                            notas_df.at[idx, "Data Pagamento"] = grupo.iloc[0]["Data"].strftime("%d/%m/%Y")
                            notas_df.at[idx, "Valor Pago"] = soma_pg
                            notas_df.at[idx, "Lançamento"] = ", ".join(grupo["Lançamento"].astype(str))
                            notas_df.at[idx, "Diferença"] = abs(soma_pg - total_a_conciliar)
                            notas_df.at[idx, "Justificativa"] = "Notas agrupadas e conciliadas por nome de destinatário"
                        break

        conciliadas = notas_df[notas_df["Conciliado"]]
        nao_conciliadas = notas_df[~notas_df["Conciliado"]]

        # Salva tudo em um único arquivo com 3 abas
        with pd.ExcelWriter("resultado_conciliacao_bancaria.xlsx", engine='openpyxl', date_format='DD/MM/YYYY') as writer:
            conciliadas.to_excel(writer, sheet_name="Notas Conciliadas", index=False)
            nao_conciliadas.to_excel(writer, sheet_name="Notas Não Conciliadas", index=False)
            extratos_df.to_excel(writer, sheet_name="Extrato Tratado", index=False)

        log("✅ Conciliação concluída e arquivo 'resultado_conciliacao_bancaria.xlsx' gerado.")
    except Exception as e:
        log(f"❌ Erro na conciliação: {e}")

# ========== FUNÇÕES DE NF-E ==========
async def iniciar_processo_egestor():
    try:
        dt_ini = data_inicio.get_date().strftime("%Y-%m-%d")
        dt_fim = data_fim.get_date().strftime("%Y-%m-%d")

        obter_token()
        buscar_codigos(dt_ini, dt_fim)

        if not codigos_nfe:
            log("⚠️ Nenhuma NF-e encontrada no período selecionado.")
            return

        connector = aiohttp.TCPConnector(limit=5)
        timeout = aiohttp.ClientTimeout(total=30)  # até 30 segundos por requisição
        async with aiohttp.ClientSession(headers=headers, connector=connector, timeout=timeout) as session:
            tasks = [fetch_detalhes_e_xml(session, codigo) for codigo in codigos_nfe]
            resultados = await asyncio.gather(*tasks, return_exceptions=True)
            detalhes = [r for r in resultados if isinstance(r, dict)]
            if detalhes:
                df = pd.json_normalize(detalhes)
                df.to_excel("detalhes_nfe_egestor.xlsx", index=False)
                log("📥 Detalhes salvos em 'detalhes_nfe_egestor.xlsx'")

        todos_itens = []
        for nome_arquivo in os.listdir("xmls"):
            if nome_arquivo.endswith(".xml"):
                caminho = os.path.join("xmls", nome_arquivo)
                todos_itens.extend(extrair_itens_xml(caminho))

        if todos_itens:
            df_itens = pd.DataFrame(todos_itens)
            df_itens.to_excel("itens_nfe_egestor.xlsx", index=False)
            log("✅ Itens extraídos e salvos em 'itens_nfe_egestor.xlsx'")
        else:
            log("⚠️ Nenhum item extraído dos XMLs.")

    except Exception as e:
        log(f"❌ Erro no processamento NF-e eGestor: {e}")
# ========== INICIAR ==========
janela.after(100, lambda: log("🔍 Sistema iniciado com sucesso!"))
janela.after(200, lambda: log("💡 Utilize os botões nas abas para iniciar os processos."))

janela.mainloop()
