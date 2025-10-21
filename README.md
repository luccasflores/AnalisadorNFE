# 🧠 Analisador NFE – M&H Soluções

Sistema de automação inteligente para **análise de NF-e, boletos e conciliação bancária**, desenvolvido pela **M&H Soluções**.  
Integra diferentes fontes de dados (API, e-mail e planilhas) em uma interface moderna e intuitiva feita em **Python + CustomTkinter**.

> Projeto desenvolvido e mantido por **Luccas Flores (M&H Soluções)** como parte da suíte de automações fiscais.
---
**Stack:** `Python 3.11` · `CustomTkinter` · `Playwright` · `Pandas` · `OpenPyXL` · `fdb` · `dotenv`
---
## 📁 Estrutura do Projeto

```bash
AnalisadorNFE/
├── src/
│   ├── app/
│   │   └── ui.py                 # Interface principal (CustomTkinter)
│   ├── core/
│   │   ├── egestor.py            # Integração com API eGestor
│   │   ├── xml_utils.py          # Extração e leitura de XMLs
│   │   ├── extratos.py           # Tratamento de extratos bancários
│   │   ├── reconcile.py          # Lógica de conciliação
│   │   └── email_billets.py      # Download e leitura de boletos por e-mail
│   └── __init__.py
│
├── docs/                         # Capturas de tela e manuais
│   ├── interface.png
│   └── exemploderesultadoparte1.png
│
├── xmls/                         # XMLs baixados da API
├── pdf_email/                    # Boletos em PDF baixados via IMAP
├── logo.ico                      # Ícone da janela
├── LOGOMEH_convertido.png         # Logo M&H
├── .env                          # Configurações locais
├── .gitignore
├── pyproject.toml
├── requirements.txt
└── README.md
```


---

### 🧮 4. Seção “Instalação e Execução”

Deixe mais clara e igual à dos outros repositórios:

```markdown
## 🚀 Instalação e Execução

1️⃣ Criar ambiente virtual e instalar dependências

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt


2️⃣ Rodar o sistema
python -m src.app.ui

---

### 🔑 5. Incluir exemplo de `.env`

```markdown
## 🔐 Exemplo de arquivo `.env`

```ini
EGESTOR_PERSONAL_TOKEN=seu_token_pessoal
IMAP_SERVER=imap.mail.yahoo.com
IMAP_PORT=993
EMAIL_USER=usuario@example.com
EMAIL_PASS=senha_de_app
TZ=America/Sao_Paulo

```
**Luccas Flores**  
Desenvolvedor Python | Especialista em RPA e Automação Fiscal  
**M&H Soluções**

📧 luccasflores.dev@gmail.com  
🌐 [LinkedIn](https://www.linkedin.com/in/luccas-flores-038757231/) | 🐙 [GitHub](https://github.com/luccasflores)

---

## ⚖️ Licença

Projeto sob a licença MIT – consulte o arquivo `LICENSE` para mais detalhes.
