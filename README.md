# VBA_PROJECT

Sistema desenvolvido em **VBA (Visual Basic for Applications)** para automatizar processos administrativos utilizando o **Microsoft Excel**.

O projeto utiliza **UserForms**, manipulação de planilhas, validação de dados e integração com o **Microsoft Outlook** para facilitar o processo de solicitação de compras.

## 📋 Sobre o projeto

O `VBA_PROJECT` foi desenvolvido para transformar o Excel em uma interface de interação mais simples para processos que normalmente exigiriam o preenchimento e organização manual de informações.

A aplicação utiliza formulários personalizados para que o usuário possa informar os dados necessários para uma solicitação de compra. Após o preenchimento, as informações são registradas automaticamente na planilha correspondente e podem ser encaminhadas por e-mail como uma planilha Excel.

O projeto também possui funcionalidades relacionadas ao cadastro e gerenciamento de itens, utilizando diferentes planilhas como base de dados.

## ⚙️ Funcionalidades

### 🛒 Solicitação de compra

O formulário de solicitação de compra permite informar:

* Solicitante;
* Data da solicitação;
* Data de entrega;
* Item solicitado;
* Quantidade;
* Justificativa da solicitação.

A data da solicitação é preenchida automaticamente pelo sistema.

Os itens disponíveis para seleção são carregados a partir da planilha `ITEM_CADASTRADOS`.

### 📦 Controle de itens

Os dados dos itens são utilizados como fonte para os formulários do sistema.

Durante a seleção de um item, o sistema consulta a planilha `ITEM_CADASTRADOS` para obter informações relacionadas ao produto selecionado.

Itens com quantidade disponível igual ou inferior a **10 unidades** são apresentados no formulário de solicitação de compra, facilitando a identificação de itens que precisam de reposição.

### 📝 Preenchimento automático

Após o preenchimento do formulário, os dados são transferidos automaticamente para a planilha `SOLICITACAO_COMPRA`.

Entre as informações preenchidas estão:

| Informação          | Localização |
| ------------------- | ----------- |
| Solicitante         | D4          |
| Data da solicitação | D5          |
| Data de entrega     | D6          |
| Justificativa       | D7          |
| Item                | C11         |
| Quantidade          | C12         |

Isso reduz a necessidade de inserir manualmente os mesmos dados diretamente na planilha.

### ✉️ Envio por e-mail

Após uma solicitação ser preenchida corretamente, o sistema pode gerar uma cópia da planilha de solicitação em formato `.xlsx`.

Essa cópia é anexada automaticamente a um novo e-mail utilizando o **Microsoft Outlook**.

O processo envolve:

1. Cópia da planilha `SOLICITACAO_COMPRA`;
2. Criação de um arquivo temporário `.xlsx`;
3. Criação de uma mensagem utilizando o Outlook;
4. Anexação da planilha;
5. Envio do e-mail;
6. Exclusão do arquivo temporário.

## 🖥️ Interface

A aplicação utiliza **UserForms do VBA** para fornecer uma interface gráfica ao usuário.

Entre os formulários presentes no projeto estão:

* `telaInicial`
* `cadastrarItem`
* `retirarItem`
* `fazerPedido`

Esses formulários permitem centralizar as principais operações do sistema em uma interface mais simples do que a manipulação direta das planilhas.

## 🗂️ Estrutura do projeto

```text
VBA_PROJECT/
│
├── EstaPastaDeTrabalho.cls
│
├── MOD/
│   └── Módulos VBA
│
├── Planilha1.cls
├── Planilha3.cls
├── Planilha4.cls
│
├── cadastrarItem.frm
├── cadastrarItem.frx
│
├── fazerPedido.frm
├── fazerPedido.frx
│
├── retirarItem.frm
├── retirarItem.frx
│
├── telaInicial.frm
├── telaInicial.frx
│
└── README.md
```

### Principais tipos de arquivos

| Extensão | Função                                             |
| -------- | -------------------------------------------------- |
| `.cls`   | Classes e módulos associados a objetos do VBA      |
| `.frm`   | Código dos UserForms                               |
| `.frx`   | Informações gráficas associadas aos UserForms      |
| `.cls`   | Código dos objetos de planilha e pasta de trabalho |

## 📊 Planilhas utilizadas

Entre as planilhas utilizadas pela aplicação estão:

### `ITEM_CADASTRADOS`

Utilizada como fonte de dados dos itens cadastrados.

O sistema consulta essa planilha para:

* Identificar os itens disponíveis;
* Obter informações associadas aos itens;
* Verificar a quantidade disponível;
* Exibir itens com estoque igual ou inferior a 10 unidades no formulário de solicitação.

### `SOLICITACAO_COMPRA`

Utilizada como modelo para registrar as informações de uma solicitação de compra.

Os dados preenchidos pelo usuário através do UserForm são inseridos automaticamente nessa planilha antes do envio.

## 🔄 Fluxo da solicitação

O funcionamento básico do processo pode ser representado da seguinte forma:

```text
        ┌─────────────────┐
        │   Usuário       │
        └────────┬────────┘
                 │
                 ▼
        ┌─────────────────┐
        │    telaInicial  │
        └────────┬────────┘
                 │
                 ▼
        ┌─────────────────┐
        │   fazerPedido   │
        └────────┬────────┘
                 │
                 ▼
        ┌─────────────────┐
        │ ITEM_CADASTRADOS│
        └────────┬────────┘
                 │
                 ▼
        ┌─────────────────┐
        │ Preenchimento   │
        │ da solicitação  │
        └────────┬────────┘
                 │
                 ▼
        ┌─────────────────┐
        │SOLICITACAO_COMPRA│
        └────────┬────────┘
                 │
                 ▼
        ┌─────────────────┐
        │     Outlook     │
        │  + anexo .xlsx  │
        └─────────────────┘
```

## 🛠️ Tecnologias utilizadas

* **Microsoft Excel**
* **VBA (Visual Basic for Applications)**
* **UserForms**
* **Microsoft Outlook**
* Manipulação de objetos do Excel
* Manipulação de arquivos temporários
* Automação de envio de e-mails

## ▶️ Como utilizar

Para utilizar o projeto:

1. Abra o arquivo do Excel que contém o projeto VBA.
2. Habilite a execução de macros, caso solicitado.
3. Abra a interface principal do sistema.
4. Utilize as opções disponíveis para cadastrar, retirar ou solicitar itens.
5. Para realizar uma solicitação de compra, preencha os campos apresentados no formulário.
6. Selecione o item desejado.
7. Informe a quantidade, data de entrega e justificativa.
8. Após a validação dos campos, o sistema preencherá automaticamente a planilha de solicitação.
9. O Outlook será utilizado para enviar a solicitação como um arquivo Excel anexado.

> **Observação:** o funcionamento do envio automático depende da disponibilidade do Microsoft Outlook no ambiente em que o projeto for executado.

## ⚠️ Dependências e observações

O projeto foi desenvolvido para utilização em um ambiente com:

* Microsoft Excel com suporte a VBA;
* Macros habilitadas;
* Microsoft Outlook instalado e configurado para envio de e-mails.

O código atualmente utiliza uma configuração de destinatário diretamente na implementação do formulário de solicitação. Para utilizar o projeto em outro ambiente, recomenda-se adaptar essa configuração conforme a necessidade.

## 🎯 Objetivos do projeto

O projeto foi desenvolvido com foco na **automação de processos utilizando Excel e VBA**, explorando conceitos como:

* Programação orientada a eventos;
* Desenvolvimento de interfaces com UserForms;
* Manipulação de planilhas;
* Leitura e escrita de células;
* Validação de entradas;
* Estruturas condicionais e de repetição;
* Manipulação de arquivos;
* Automação do Outlook;
* Integração entre diferentes aplicações do Microsoft Office.

Além de automatizar tarefas repetitivas, o projeto serviu como uma aplicação prática dos conhecimentos de programação em um processo administrativo real.

## 👨‍💻 Autor

**Rafael Farias de Lima**

Projeto desenvolvido para estudo, prática e aplicação de **VBA, automação de processos e integração entre Microsoft Excel e Outlook**.
