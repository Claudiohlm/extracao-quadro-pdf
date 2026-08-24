# 📄 Extração de Dados de PDF para Excel

Automação desenvolvida em **Python** para extrair informações de um quadro padronizado presente em documentos PDF e consolidar os dados em uma **planilha Excel**.

O sistema foi desenvolvido para processar documentos que possuem uma estrutura previamente definida, identificando informações como número do ofício, número do processo, destinatário, endereço, bairro, CEP, cidade e UF.

## 🛠️ Tecnologias

* Python
* pdfplumber
* Regex
* OpenPyXL

## 📦 Instalação

Instale as dependências necessárias:

```bash
pip install pdfplumber openpyxl
```

## ▶️ Como usar

### 1. Configure os caminhos

No início do código, informe a pasta onde estão os PDFs e o caminho do arquivo Excel:

```python
PASTA_PDFS = r"C:\caminho\dos\arquivos"
ARQUIVO_EXCEL = r"C:\caminho\resultado.xlsx"
```

### 2. Execute o script

```bash
python extracao_dados.py
```

O programa percorre os arquivos PDF da pasta configurada, extrai o texto dos documentos, identifica as informações necessárias e adiciona os resultados à planilha Excel.

## 📊 Dados extraídos

A automação identifica informações como:

* Número do ofício SEI
* Número do processo SEI
* Nome do destinatário
* Endereço
* Número
* Complemento
* Bairro
* CEP
* Cidade
* UF

## 📁 Resultado

Os dados extraídos são organizados automaticamente em uma planilha Excel com as seguintes colunas:

| DOCUMENTO | NOME | ENDEREÇO | CEP |
| --------- | ---- | -------- | --- |

Cada documento PDF processado gera uma nova linha na planilha.

## 🎯 Objetivo

Automatizar a extração de informações de documentos PDF padronizados, reduzindo a necessidade de preenchimento manual e facilitando a organização dos dados em formato estruturado.

## ⚠️ Observação

A automação foi desenvolvida considerando a estrutura específica de certos documentos. 

