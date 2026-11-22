# Controle de Despesas

Aplicativo para Windows que registra despesas **fixas**, **variáveis** e **adicionais** por mês/ano,
guarda tudo num banco SQLite local (`financas.db`) e mostra um resumo com gráfico da evolução,
salário líquido e saldo remanescente.

## Como gerar o .exe

**Opção 1 – no seu PC (mais simples)**
1. Instale o Python 3.10 ou superior em https://www.python.org/downloads/ (marque *Add python.exe to PATH*).
2. Dê dois cliques em `build.bat` nesta pasta.
3. O executável fica em `dist\ControleDespesas.exe`. Copie-o para a pasta que quiser.

**Opção 2 – GitHub Actions (sem instalar nada)**
Suba esta pasta para um repositório no GitHub. O workflow `.github/workflows/build-windows.yml`
gera o `.exe` a cada push na `main` (ou manualmente em *Actions → Build Windows .exe → Run workflow*).
Baixe em *Artifacts → ControleDespesas*.

O ícone do app (`icone.ico`) é aplicado automaticamente ao `.exe` e às janelas.

**Rodar sem gerar o .exe:** `pip install matplotlib openpyxl` e depois `python controle_despesas.py`.

## Como usar

- No topo, escolha **Mês** e **Ano** (ou use ◀ ▶). Tudo o que você lança fica ligado a esse mês.
- Abas **Fixas / Variáveis / Adicionais**: escolha ou digite a descrição, informe o valor
  (aceita `1.234,56`) e clique **Adicionar** (ou Enter). Clique numa linha para editar; **Excluir selecionada** remove.
- **Copiar do mês anterior** traz os lançamentos do mês passado para o mês atual (útil para as fixas); depois é só ajustar os valores.
- Aba **Resumo**: informe o **salário líquido** do mês e veja totais por categoria, total geral e **saldo**
  (salário − despesas, em vermelho se negativo). O gráfico mostra a evolução de cada categoria
  (últimos 12 meses, ano selecionado ou todo o histórico). Dê duplo clique numa linha do histórico para ir àquele mês.

## Importar despesas de uma planilha

Clique em **Importar dados** no topo (o botão **ⓘ** ao lado explica o formato) e escolha um arquivo `.xlsx` ou `.csv` com as colunas:

| Data | Categoria | Descrição | Valor |
|---|---|---|---|
| 05/01/2026 | Fixa | Condomínio | 506,59 |
| 12/01/2026 | Variável | Cartão de Crédito | 4.045,60 |
| 20/01/2026 | Adicional | Doações | 68,00 |

- A ordem das colunas não importa se houver cabeçalho (também aceita "Data da despesa", "Valor da despesa" etc.). Sem cabeçalho, a ordem acima é usada.
- **Data**: data do Excel ou texto como `15/03/2026`, `2026-03-15`, `03/2026` ou `mar/2026`. Só o mês e o ano são usados.
- **Categoria**: Fixa(s), Variável(is) ou Adicional(is) (Extra também vira Adicional). Maiúsculas e acentos não importam.
- **Valor**: número ou texto como `1.234,56` ou `R$ 1.234,56`.
- Lançamentos com a mesma descrição no mesmo mês e categoria são somados (ex.: várias corridas de Uber).
- Antes de gravar, uma janela mostra a prévia, as linhas com problema (que serão ignoradas) e, se já houver
  lançamentos iguais no banco, deixa escolher entre **substituir** ou **somar** os valores.
- Na janela do **ⓘ**, o botão **Salvar planilha modelo...** gera um arquivo de exemplo pronto para preencher
  (também incluído aqui como `modelo-importacao-despesas.xlsx`).

## Exportar o histórico

**Exportar histórico** salva um `.xlsx` com duas abas: **Despesas** (todas as despesas, no mesmo formato da
importação, mais a observação) e **Resumo mensal** (totais por categoria, salário e saldo de cada mês).
O arquivo pode ser importado de volta, por exemplo para levar os dados para outro computador.

## Onde ficam os dados

O arquivo `financas.db` é criado na mesma pasta do `.exe` (se a pasta não permitir escrita, em
`%APPDATA%\ControleDespesas`). O caminho aparece no canto superior direito do app.
Para fazer backup basta copiar esse arquivo.
