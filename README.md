# 🧠 Analisador NFE - M&H Soluções

Sistema de automação inteligente para **análise de NF-e, boletos e conciliação bancária**, desenvolvido pela **M&H Soluções**.  
Integra diferentes fontes de dados (API, e-mail e planilhas) em uma interface visual moderna e intuitiva construída com **CustomTkinter**.
![Interface do sistema](docs/interface.png)

---

## ⚙️ Funcionalidades Principais

### 🧾 NF-e (API eGestor)
- Autenticação via **token pessoal** da API eGestor  
- Download automático dos **XMLs** das notas fiscais  
- Extração de informações completas (CNPJ, secretaria, empenho, valores, itens, NCM, CFOP)  
- Geração automática de planilhas Excel (`Notas.xlsx` e `itens_nfe_egestor.xlsx`)

### 💌 Boletos por E-mail (IMAP)
- Conexão com caixa de entrada (Yahoo, Gmail, etc.)  
- Download automático de **PDFs anexos** dentro de um intervalo de datas  
- Identificação de boletos e extração de **valor** e **vencimento**  
- Exportação de resultados em planilha (`boletos_email.xlsx`)

### 💰 Conciliação Bancária
- Importação de **extratos bancários Excel ou CSV**  
- Tratamento e padronização automáticos dos dados  
- Comparação de valores com notas fiscais — individual ou agrupada por fornecedor  
- Geração de **relatório final consolidado** (`resultado_conciliacao_bancaria.xlsx`)

---

## 🖥️ Interface (GUI)

A interface foi construída com **CustomTkinter**, oferecendo modo escuro e organização por abas:

- **NF-e eGestor** → busca e download de XMLs  
- **Boletos por E-mail** → leitura e extração de PDFs  
- **Conciliação Bancária** → cruzamento entre notas e extratos  

![Interface do sistema](docs/interface.png)
![Exemplo de resultado](docs/exemploderesultadoparte1.png)

---

## 📁 Estrutura do Projeto
AnalisadorNFE/
│
├── Analisador.py # Código principal da aplicação
├── requirements.txt # Dependências principais
├── requirements.lock.txt # Versões fixadas
├── logo.ico # Ícone da janela
│
├── docs/ # Prints para documentação
│ ├── interface.png
│ └── exemploderesultadoparte1.png
│
├── xmls/ # (gerado) XMLs baixados da API
├── pdf_email/ # (gerado) Boletos em PDF baixados por IMAP
├── extratos_consolidados.xlsx # (gerado) Extratos tratados
└── resultado_conciliacao_bancaria.xlsx # (gerado) Conciliação final

---

## 🧩 Requisitos

- **Python 3.11+**
- Instalar dependências:

```bash
pip install -r requirements.txt
    