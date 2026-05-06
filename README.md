# patrimar_dependencies

![Versão](https://img.shields.io/badge/versão-2.19.0-blue)
![Python](https://img.shields.io/badge/python-3.x-blue?logo=python&logoColor=white)
![Licença](https://img.shields.io/badge/licença-MIT-green)
![Patrimar Engenharia](https://img.shields.io/badge/Patrimar-Engenharia-orange)

Biblioteca de dependências e utilitários para projetos de automação RPA da **Patrimar Engenharia**.  
Centraliza funcionalidades como automação web, envio de e-mails, integração com IA, logging, acesso ao SharePoint, automação SAP e muito mais.

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Instalação](#instalação)
- [Uso](#uso)
  - [Credenciais](#credenciais)
  - [Logs](#logs)
  - [SendEmail](#sendemail)
  - [Navegador (Chrome / Edge / Firefox)](#navegador-chrome--edge--firefox)
  - [SAP](#sap)
  - [SharePoint](#sharepoint)
  - [IA — Gemini e GPT](#ia--gemini-e-gpt)
  - [Screenshot](#screenshot)
  - [Functions](#functions)
- [Dependências](#dependências)
- [Licença](#licença)

---

## Funcionalidades

| Módulo | Descrição |
|---|---|
| `Credential` | Gerenciamento seguro de credenciais em arquivos `.json` |
| `Logs` | Sistema de logging com integração ao servidor Central RPA |
| `SendEmail` | Envio de e-mails com suporte a anexos via SMTP |
| `NavegadorChrome` | Automação web com Google Chrome (Selenium) |
| `NavegadorEdge` | Automação web com Microsoft Edge (Selenium) |
| `NavegadorFirefox` | Automação web com Mozilla Firefox (Selenium) |
| `SAP` | Automação de processos no sistema SAP |
| `SharePointGraph` | Integração com SharePoint via Microsoft Graph API |
| `SharePointFolder` | Acesso e manipulação de pastas do SharePoint |
| `GeminiIA` | Integração com a IA generativa Google Gemini |
| `GPTIA` | Integração com a IA generativa OpenAI GPT |
| `screenshot` | Captura de tela da área de trabalho |
| `Functions` | Utilitários gerais (Excel, formatação de caminhos, colorama) |
| `Config` | Gerenciamento de configurações do projeto |
| `Arguments` | Gerenciamento de argumentos de execução |
| `TaskBotCity` | Integração com tarefas do BotCity |
| `Informativo` | Exibição de informativos padronizados |

---

## Instalação

Instale diretamente pelo repositório do GitHub:

```bash
pip install git+https://github.com/RenanMGX/dependencies_patrimar.git
```

Ou clone e instale localmente:

```bash
git clone https://github.com/RenanMGX/dependencies_patrimar.git
cd dependencies_patrimar
pip install .
```

---

## Uso

### Credenciais

Leitura e criação de credenciais armazenadas em arquivos `.json`.

```python
from patrimar_dependencies import Credential

# Carregar credenciais existentes
cred = Credential(path_raiz="C:/credenciais", name_file="sistema_x")
dados = cred.load()
usuario = dados["usuario"]
senha   = dados["senha"]

# Criar novo arquivo de credenciais
Credential.create(path_raiz="C:/credenciais", name_file="sistema_x")
```

---

### Logs

Registra eventos de execução com integração ao servidor Central RPA.

```python
from patrimar_dependencies import Logs

log = Logs(name="meu_processo")

log.register("Processo iniciado")
log.register("Etapa 1 concluída")
log.error("Falha ao abrir arquivo")
```

---

### SendEmail

Envio de e-mails com suporte a cópia (CC), anexos e conteúdo HTML.

```python
from patrimar_dependencies import SendEmail

email = SendEmail(
    email="rpa@patrimar.com.br",
    password="senha_aqui",
    smtp_server="smtp-mail.outlook.com",
    smtp_port=587
)

email.mensagem(
    Destino="destinatario@exemplo.com",
    Assunto="Relatório diário",
    Corpo="<b>Processo finalizado com sucesso.</b>",
    tipo="html"
)
```

---

### Navegador (Chrome / Edge / Firefox)

Automação web baseada em Selenium com métodos auxiliares prontos.

```python
from patrimar_dependencies import NavegadorChrome

nav = NavegadorChrome(headless=False)

nav.get("https://www.exemplo.com")

# Clicar em elemento por XPath
nav.clicar("//button[@id='entrar']")

# Preencher campo
nav.preencher("//input[@name='usuario']", "meu_usuario")

nav.quit()
```

> Substitua `NavegadorChrome` por `NavegadorEdge` ou `NavegadorFirefox` conforme necessário.

---

### SAP

Automação de transações no sistema SAP via scripting.

```python
from patrimar_dependencies import SAP

sap = SAP()
sap.abrir_transacao("SE16")
```

---

### SharePoint

Acesse arquivos e pastas do SharePoint via Microsoft Graph API.

```python
from patrimar_dependencies import SharePointGraph

sp = SharePointGraph(
    client_id="...",
    client_secret="...",
    tenant_id="..."
)

# Listar arquivos de uma pasta
arquivos = sp.listar_arquivos(caminho="/sites/RPA/Documentos")

# Fazer download de um arquivo
sp.download(caminho="/sites/RPA/Documentos/relatorio.xlsx", destino="C:/Downloads")
```

---

### IA — Gemini e GPT

Integração com modelos de linguagem para processamento inteligente de dados.

```python
from patrimar_dependencies import GeminiIA

ia = GeminiIA(api_key="sua_api_key")
resposta = ia.perguntar("Resuma o seguinte texto: ...")
print(resposta)
```

```python
from patrimar_dependencies import GPTIA

ia = GPTIA(api_key="sua_api_key")
resposta = ia.perguntar("Classifique este documento: ...")
print(resposta)
```

---

### Screenshot

Captura a tela atual e salva como arquivo `.png`.

```python
from patrimar_dependencies import screenshot

caminho = screenshot()          # salva em ./screenshot/screenshot.png
caminho = screenshot("C:/tmp/captura")  # salva em C:/tmp/captura.png
print(f"Imagem salva em: {caminho}")
```

---

### Functions

Utilitários gerais para manipulação de arquivos Excel e formatação de caminhos.

```python
from patrimar_dependencies import Functions

# Fechar um arquivo Excel aberto
Functions.fechar_excel("relatorio.xlsx")

# Listar arquivos Excel abertos
abertos = Functions.excel_open()
print(abertos)

# Normalizar caminho (remover barra final)
caminho = Functions.tratar_caminho("C:/pasta/subpasta/")
# Resultado: "C:/pasta/subpasta"
```

---

## Dependências

| Pacote | Finalidade |
|---|---|
| `selenium` | Automação web |
| `pandas` | Manipulação de dados |
| `openpyxl` | Leitura/escrita de arquivos Excel |
| `xlwings` | Integração avançada com Excel |
| `requests` | Requisições HTTP |
| `google-generativeai` | Integração com Google Gemini |
| `openai` | Integração com OpenAI GPT |
| `pyautogui` | Automação de interface gráfica |
| `pillow` | Processamento de imagens |
| `python-dotenv` | Gerenciamento de variáveis de ambiente |
| `colorama` | Saída colorida no terminal |
| `psutil` | Monitoramento de processos do sistema |
| `pywin32` | APIs do Windows |

Instale todas as dependências com:

```bash
pip install -r requirements.txt
```

---

## Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE).  
Copyright © 2026 [Renan Oliveira](mailto:renanmgx@hotmail.com)
