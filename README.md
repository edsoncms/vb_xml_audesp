# 📄 Gerador de XML para AUDESP

<div align="center">

![Status](https://img.shields.io/badge/status-legado%20/%20arquivo-yellow?style=for-the-badge)
![Linguagem](https://img.shields.io/badge/linguagem-Visual%20Basic%206-blue?style=for-the-badge&logo=windows)
![Plataforma](https://img.shields.io/badge/plataforma-Windows-0078D6?style=for-the-badge&logo=windows)
![Órgão](https://img.shields.io/badge/órgão-TCE--SP-darkgreen?style=for-the-badge)
![Estado](https://img.shields.io/badge/estado-São%20Paulo-brightgreen?style=for-the-badge)
![Licença](https://img.shields.io/badge/licença-Proprietária-red?style=for-the-badge)

</div>

---

## 📌 Sobre o Projeto

O **vb_xml_audesp** é um sistema desktop desenvolvido em **Visual Basic 6** pela [MS3TI](http://www.ms3ti.com.br), com o objetivo de automatizar a geração de arquivos **XML** no padrão exigido pelo sistema **AUDESP** do **Tribunal de Contas do Estado de São Paulo (TCE-SP)**.

O sistema lê arquivos de dados no formato **CSV (separado por ponto e vírgula)** e os converte nos XMLs padronizados pelo TCE-SP para prestação de contas do quadro funcional de entidades públicas municipais.

---

## 🏛️ O que é a AUDESP?

A **AUDESP** é o sistema de **Auditoria Digital do Estado de São Paulo**, mantido pelo **Tribunal de Contas do Estado de São Paulo (TCE-SP)**. Por meio dela, os municípios e entidades públicas paulistas enviam periodicamente informações sobre sua gestão — incluindo dados sobre servidores públicos — em formato XML com schemas validados pelo próprio Tribunal.

---

## ⚙️ O que o sistema faz?

O sistema processa dois tipos de exportação de dados para a AUDESP, selecionáveis pelo usuário na própria interface:

### 👤 Agente Público (`AgentePublico.xml`)

Lê um arquivo CSV contendo os dados cadastrais dos servidores e gera o XML no schema:
`AUDESP_AGENTEPUBLICO_2016_A.XSD`

**Campos processados do CSV:**

| Posição | Campo |
|-----------|-------|
| 0 | Nome do Agente |
| 1 | CPF |
| 2 | PIS/PASEP |
| 3 | Data de Nascimento |
| 4 | Sexo |
| 5 | Nacionalidade |
| 6 | Escolaridade |
| 7 | Especialidade |

---

### 🏢 Lotação do Agente Público (`Lotacao.xml`)

Lê **dois arquivos CSV** — um com os dados de lotação atual e outro com o histórico — e gera o XML no schema:
`AUDESP_LOTACAOAGENTEPUBLICO_2016_A.XSD`

**Campos do arquivo de lotação:**

| Posição | Campo |
|-----------|-------|
| 0 | CPF |
| 1 | Data de Lotação |
| 2 | Nome do Agente |
| 3 | Exercício de Atividade |
| 4 | Código do Cargo |
| 5 | Código da Função |
| 6 | Cargo/Função |
| 7 | Unidade de Lotação |
| 8 | Função de Governo |
| 9 | Forma de Provimento |

**Campos do arquivo histórico:**

| Posição | Campo |
|-----------|-------|
| 0 | CPF |
| 1 | Cargo/Função (histórico) |
| 2 | Data de Lotação |
| 3 | Nome do Agente |
| 4 | Data das Situações |
| 5 | Status/Situação |

---

## 🗂️ Estrutura do Repositório

```
vb_xml_audesp/
├── audesp.frm        # Formulário principal (lógica da aplicação)
├── audesp.frx        # Recursos binários do formulário (imagens, ícones)
├── frmAbout.frm      # Tela "Sobre"
├── frmAbout.frx      # Recursos binários da tela "Sobre"
├── audesp.vbp        # Arquivo de projeto Visual Basic
├── audesp.vbg        # Grupo de projetos VB
├── audesp.vbw        # Configurações da janela do IDE
├── audesp.exe        # Executável compilado
├── audesp2.exe       # Executável compilado (versão 2)
├── audesp.OBJ        # Arquivo objeto compilado
├── audesp1.OBJ       # Arquivo objeto compilado (módulo 1)
├── audesp.log        # Log de compilação
├── MSSCCPRJ.SCC      # Controle de versão do SourceSafe
└── loading-geral1.gif # Animação de carregamento
```

---

## 🖥️ Requisitos para Execução

> ⚠️ **Sistema legado** — requer ambiente Windows 32-bit com as dependências do Visual Basic 6.

- **Sistema Operacional:** Windows (recomendado Windows XP / 7 32-bit para maior compatibilidade)
- **Runtime VB6:** `MSVBVM60.DLL` instalado
- **OCX necessário:** `COMDLG32.OCX` (diálogo de arquivo padrão do Windows)

---

## 🚀 Como Usar

1. Execute o arquivo `audesp.exe`
2. Selecione o **tipo de arquivo** que deseja gerar:
   - **Agente Público** — dados cadastrais dos servidores
   - **Lotação** — dados de lotação e histórico
3. Clique em **"Selecionar Arquivo"** e escolha o CSV de entrada (separado por `;`)
4. Para **Lotação**, clique também em **"Selecionar Arquivo Histórico"** com o CSV de histórico
5. Clique em **"Gerar XML"**
6. O arquivo gerado será salvo na **mesma pasta do executável**:
   - `AgentePublico.xml` ou `Lotacao.xml`

---

## 📋 Formato dos Arquivos CSV de Entrada

Os arquivos de entrada devem ser **texto puro**, com campos separados por **ponto e vírgula (`;`)**, sem cabeçalho, uma linha por registro.

**Exemplo — Agente Público:**
```
JOAO DA SILVA;12345678901;12345678901;19800101;M;1;5;1
```

**Exemplo — Lotação:**
```
12345678901;20160101;JOAO DA SILVA;1;001;002;1;SEC ADMINISTRACAO;1;1
```

---

## 📤 Arquivos XML Gerados

Os XMLs seguem o padrão de namespaces do TCE-SP:

- `http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-agentepublico`
- `http://www.tce.sp.gov.br/audesp/xml/quadrofuncional-lotacaoagentepublico`

A codificação utilizada é **ISO-8859-1** (Latin-1), conforme exigido pelo TCE-SP para compatibilidade com acentuação do português.

---

## 🏷️ Contexto e Observações

- O sistema foi desenvolvido com referência ao exercício de **2016**, com os schemas de validação `*_2016_A.XSD`
- Os campos **Entidade (`10063`)** e **Município (`7107`)** estão fixos no código-fonte e devem ser ajustados conforme a entidade municipal que utilizar o sistema
- O desenvolvimento foi realizado pela empresa **[MS3TI](http://www.ms3ti.com.br)**

---

## 👨‍💻 Autor

Desenvolvido por **MS3TI** — Soluções em Tecnologia da Informação.

---

*Este repositório tem caráter histórico e de referência para municípios paulistas que utilizaram ou ainda utilizam esta solução para envio de dados ao TCE-SP via AUDESP.*
