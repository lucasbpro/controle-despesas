# 💰 Controle de Despesas

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Plataforma](https://img.shields.io/badge/plataforma-Windows-0078D6?logo=windows&logoColor=white)
![SQLite](https://img.shields.io/badge/banco-SQLite-003B57?logo=sqlite&logoColor=white)
![Build](https://img.shields.io/badge/build-GitHub%20Actions-2088FF?logo=githubactions&logoColor=white)

Aplicativo desktop para Windows que registra despesas **fixas**, **variáveis** e **adicionais** por mês/ano,
guarda tudo num banco SQLite local (`financas.db`) e mostra um resumo com gráfico da evolução,
salário líquido e saldo remanescente.

---

## 📑 Sumário

- [Recursos](#-recursos)
- [Requisitos](#-requisitos)
- [Instalação e execução](#-instalação-e-execução)
- [Gerando o executável (.exe)](#-gerando-o-executável-exe)
- [Como usar](#-como-usar)
- [Importar despesas de uma planilha](#-importar-despesas-de-uma-planilha)
- [Exportar o histórico](#-exportar-o-histórico)
- [Onde ficam os dados](#-onde-ficam-os-dados)
- [Estrutura do projeto](#-estrutura-do-projeto)

---

## ✨ Recursos

- Lançamento de despesas em três categorias: **Fixas**, **Variáveis** e **Adicionais**.
- Navegação por **mês e ano**, com cópia dos lançamentos do mês anterior.
- **Resumo mensal** com totais por categoria, salário líquido e saldo (em vermelho se negativo).
- **Gráfico** da evolução de cada categoria (últimos 12 meses, ano selecionado ou todo o histórico).
- **Importação** de despesas a partir de planilhas `.xlsx` ou `.csv`, com prévia e tratamento de conflitos.
- **Exportação** do histórico completo para `.xlsx`.
- Valores no formato brasileiro (`1.234,56` ou `R$ 1.234,56`).
- Dados 100% locais em um único arquivo SQLite, fácil de copiar para backup.

## 📋 Requisitos

- Windows 10 ou superior
- [Python 3.10+](https://www.python.org/downloads/) (apenas para rodar a partir do código ou gerar o `.exe`)
- Dependências listadas em [`requirements.txt`](requirements.txt):
  - `matplotlib>=3.8`
  - `openpyxl>=3.1`
  - `pyinstaller>=6.0` (somente para o build)

## 🚀 Instalação e execução

Para rodar direto do código-fonte, sem gerar o `.exe`:

```bash
git clone https://github.com/<seu-usuario>/controle-despesas.git
cd controle-despesas
pip install matplotlib openpyxl
python controle_despesas.py
```

## 📦 Gerando o executável (.exe)

O ícone do app (`icone.ico`) é aplicado automaticamente ao `.exe` e às janelas.

### Opção 1: no seu PC com `build.bat`

1. Instale o [Python 3.10 ou superior](https://www.python.org/downloads/) e marque **Add python.exe to PATH**.
2. Dê dois cliques em [`build.bat`](build.bat).
   O script cria um ambiente virtual (`.venv`), instala as dependências e roda o PyInstaller.
3. O executável fica em `dist\ControleDespesas.exe`. Copie-o para a pasta que quiser.

### Opção 2: GitHub Actions (sem instalar nada)

O workflow [`.github/workflows/build-windows.yml`](.github/workflows/build-windows.yml) gera o `.exe` num runner
Windows com Python 3.12:

- automaticamente a cada **push na branch `main`**; ou
- manualmente em **Actions → Build Windows .exe → Run workflow**.

Ao final, baixe o executável na seção **Artifacts → ControleDespesas** da execução.

## 🖥️ Como usar

- No topo, escolha **Mês** e **Ano** (ou use ◀ ▶). Tudo o que você lança fica ligado a esse mês.
- Nas abas **Fixas / Variáveis / Adicionais**, escolha ou digite a descrição, informe o valor
  (aceita `1.234,56`) e clique em **Adicionar** (ou tecle Enter).
  Clique numa linha para editar; **Excluir selecionada** remove o lançamento.
- **Copiar do mês anterior** traz os lançamentos do mês passado para o mês atual (útil para as fixas);
  depois é só ajustar os valores.
- Na aba **Resumo**, informe o **salário líquido** do mês e veja os totais por categoria, o total geral e o
  **saldo** (salário − despesas). O gráfico mostra a evolução de cada categoria.
  Dê duplo clique numa linha do histórico para ir àquele mês.

## 📥 Importar despesas de uma planilha

Clique em **Importar dados** no topo (o botão **ⓘ** ao lado explica o formato) e escolha um arquivo
`.xlsx` ou `.csv` com as colunas:

| Data       | Categoria | Descrição         | Valor    |
|------------|-----------|-------------------|----------|
| 05/01/2026 | Fixa      | Condomínio        | 506,59   |
| 12/01/2026 | Variável  | Cartão de Crédito | 4.045,60 |
| 20/01/2026 | Adicional | Doações           | 68,00    |

> 💡 Um modelo pronto para preencher está em [`modelo-importacao-despesas.xlsx`](modelo-importacao-despesas.xlsx).
> Também é possível gerá-lo pelo botão **Salvar planilha modelo...** na janela do **ⓘ**.

**Regras de leitura**

- **Colunas**: com cabeçalho, a ordem não importa (também aceita "Data da despesa", "Valor da despesa" etc.).
  Sem cabeçalho, vale a ordem acima.
- **Data**: data do Excel ou texto como `15/03/2026`, `2026-03-15`, `03/2026` ou `mar/2026`.
  Só o mês e o ano são usados.
- **Categoria**: Fixa(s), Variável(is) ou Adicional(is) (Extra também vira Adicional).
  Maiúsculas e acentos não importam.
- **Valor**: número ou texto como `1.234,56` ou `R$ 1.234,56`.
- Lançamentos com a mesma descrição no mesmo mês e categoria são **somados** (ex.: várias corridas de Uber).

Antes de gravar, uma janela mostra a **prévia**, as linhas com problema (que serão ignoradas) e, se já houver
lançamentos iguais no banco, deixa escolher entre **substituir** ou **somar** os valores.

## 📤 Exportar o histórico

**Exportar histórico** salva um `.xlsx` com duas abas:

| Aba               | Conteúdo                                                                 |
|-------------------|--------------------------------------------------------------------------|
| **Despesas**      | Todas as despesas, no mesmo formato da importação, mais a observação     |
| **Resumo mensal** | Totais por categoria, salário e saldo de cada mês                        |

O arquivo pode ser importado de volta, por exemplo para levar os dados para outro computador.

## 💾 Onde ficam os dados

O arquivo `financas.db` é criado na mesma pasta do `.exe`. Se a pasta não permitir escrita, ele vai para
`%APPDATA%\ControleDespesas`. O caminho em uso aparece no canto superior direito do app.

**Backup:** basta copiar o arquivo `financas.db`.

## 🗂️ Estrutura do projeto

```
controle-despesas/
├── .github/
│   └── workflows/
│       └── build-windows.yml          # Build do .exe via GitHub Actions
├── build.bat                          # Build local do .exe (Windows)
├── controle_despesas.py               # Código-fonte do aplicativo (Tkinter + SQLite)
├── icone.ico / icone.png              # Ícones do aplicativo
├── modelo-importacao-despesas.xlsx    # Planilha modelo para importação
├── requirements.txt                   # Dependências Python
├── LEIA-ME.md                         # Instruções resumidas
└── README.md
```
