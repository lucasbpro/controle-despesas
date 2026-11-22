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

**Rodar sem gerar o .exe:** `pip install matplotlib` e depois `python controle_despesas.py`.

## Como usar

- No topo, escolha **Mês** e **Ano** (ou use ◀ ▶). Tudo o que você lança fica ligado a esse mês.
- Abas **Fixas / Variáveis / Adicionais**: escolha ou digite a descrição, informe o valor
  (aceita `1.234,56`) e clique **Adicionar** (ou Enter). Clique numa linha para editar; **Excluir selecionada** remove.
- **Copiar do mês anterior** traz os lançamentos do mês passado para o mês atual (útil para as fixas); depois é só ajustar os valores.
- Aba **Resumo**: informe o **salário líquido** do mês e veja totais por categoria, total geral e **saldo**
  (salário − despesas, em vermelho se negativo). O gráfico mostra a evolução de cada categoria
  (últimos 12 meses, ano selecionado ou todo o histórico). Dê duplo clique numa linha do histórico para ir àquele mês.

## Onde ficam os dados

O arquivo `financas.db` é criado na mesma pasta do `.exe` (se a pasta não permitir escrita, em
`%APPDATA%\ControleDespesas`). O caminho aparece no canto superior direito do app.
Para fazer backup basta copiar esse arquivo.
