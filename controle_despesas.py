# -*- coding: utf-8 -*-
"""
Controle de Despesas
Aplicativo desktop (Tkinter + SQLite + Matplotlib) para registrar despesas
fixas, variáveis e adicionais por mês/ano e acompanhar o saldo mensal.
"""
import os
import sys
import sqlite3
import datetime as dt
import csv
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import FuncFormatter

APP_NAME = "Controle de Despesas"
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
         "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MESES_ABREV = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set",
               "out", "nov", "dez"]

# (chave no banco, título da aba, cor no gráfico)
CATEGORIAS = [
    ("fixas", "Despesas Fixas", "#2a78d6"),
    ("variaveis", "Despesas Variáveis", "#e8862a"),
    ("adicionais", "Despesas Adicionais/Extras", "#8a4fd0"),
]
COR_TOTAL = "#c43c3c"
COR_SALARIO = "#2e9e5b"

# Itens sugeridos, baseados na planilha "Finanças".
ITENS_PADRAO = {
    "fixas": ["Aluguel / Financiamento", "Condomínio", "Gás", "Luz",
              "Celular e Internet", "Faxina", "Cursos (idiomas, forró, etc)"],
    "variaveis": ["Cartão de Crédito", "Supermercado", "Transporte (Uber, combustível)",
                  "Saúde, cuidado pessoal", "Lazer"],
    "adicionais": ["Doações", "Transferências Pix", "Outros boletos (Ex.: Impostos)",
                   "Manutenção do carro", "Médico"],
}


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def pasta_dados():
    """Pasta onde o banco é guardado: ao lado do .exe (ou do script).
    Se não for possível escrever ali, usa %APPDATA%\\ControleDespesas."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    teste = os.path.join(base, ".teste_escrita")
    try:
        with open(teste, "w") as f:
            f.write("ok")
        os.remove(teste)
        return base
    except OSError:
        alt = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                           "ControleDespesas")
        os.makedirs(alt, exist_ok=True)
        return alt


def formatar_brl(valor):
    if valor is None:
        valor = 0.0
    s = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-R$ " if valor < 0 else "R$ ") + s


def ler_valor(texto):
    """Aceita '1.234,56', '1234,56', '1234.56', 'R$ 1.234,56'."""
    t = (texto or "").strip().replace("R$", "").replace(" ", "")
    if not t:
        raise ValueError("vazio")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1 or (t.count(".") == 1 and len(t.split(".")[1]) == 3):
        # "1.234.567" ou "12.800" -> ponto como separador de milhar
        t = t.replace(".", "")
    return round(float(t), 2)



def caminho_recurso(nome):
    """Caminho de um arquivo empacotado junto com o app (ícone etc.)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome)


# --------------------------------------------------------------------------
# Importação de planilha
# --------------------------------------------------------------------------
def _normalizar(texto):
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join(t.lower().replace("_", " ").split())


CABECALHOS = {
    "data": ("data", "data da despesa", "dia", "date", "mes", "mes/ano", "competencia"),
    "categoria": ("categoria", "tipo", "tipo de despesa", "grupo", "category"),
    "descricao": ("descricao", "descricao da despesa", "despesa", "item", "historico", "description"),
    "valor": ("valor", "valor da despesa", "valor (r$)", "montante", "quantia", "amount", "value"),
}


def mapear_categoria(texto):
    t = _normalizar(texto)
    if t.startswith("fix"):
        return "fixas"
    if t.startswith("vari"):
        return "variaveis"
    if t.startswith(("adic", "extra", "esporad", "eventua")):
        return "adicionais"
    return None


def ler_data(v):
    """Retorna (ano, mes) a partir de datetime/date, número de série do Excel
    ou textos como 15/03/2025, 2025-03-15, 03/2025, mar/2025."""
    if isinstance(v, dt.datetime):
        return v.year, v.month
    if isinstance(v, dt.date):
        return v.year, v.month
    if isinstance(v, (int, float)) and 20000 < v < 80000:   # série do Excel
        d = dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        return d.year, d.month
    t = str(v or "").strip()
    if not t:
        raise ValueError("data vazia")
    t = t.split(" ")[0].split("T")[0]
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y",
                "%m/%Y", "%Y-%m", "%m-%Y"):
        try:
            d = dt.datetime.strptime(t, fmt)
            return d.year, d.month
        except ValueError:
            pass
    partes = _normalizar(t).replace("-", "/").split("/")
    if len(partes) == 2 and partes[0][:3] in MESES_ABREV:
        ano = int(partes[1])
        return (ano + 2000 if ano < 100 else ano), MESES_ABREV.index(partes[0][:3]) + 1
    raise ValueError(f"data não reconhecida: {v}")


def ler_linhas_arquivo(caminho):
    """Lê .xlsx/.xlsm ou .csv e devolve lista de linhas (listas de células)."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
        ws = wb.active
        linhas = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return linhas
    if ext in (".csv", ".txt"):
        with open(caminho, "rb") as f:
            bruto = f.read()
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                texto = bruto.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        amostra = texto[:4096]
        try:
            dialeto = csv.Sniffer().sniff(amostra, delimiters=";,\t")
        except csv.Error:
            dialeto = csv.excel
            dialeto.delimiter = ";" if amostra.count(";") >= amostra.count(",") else ","
        return [r for r in csv.reader(texto.splitlines(), dialeto)]
    raise ValueError("Formato não suportado. Use .xlsx ou .csv")


def interpretar_planilha(linhas):
    """Converte as linhas em lançamentos.
    Retorna (lancamentos, erros). Cada lançamento: dict(ano, mes, categoria, descricao, valor, linha)."""
    linhas = [r for r in linhas if r and any(c not in (None, "") for c in r)]
    if not linhas:
        return [], ["A planilha está vazia."]

    # Localiza as colunas pelo cabeçalho; sem cabeçalho, usa a ordem
    # Data | Categoria | Descrição | Valor
    cab = [_normalizar(c) for c in linhas[0]]
    idx = {}
    for chave, nomes in CABECALHOS.items():
        for i, c in enumerate(cab):
            if c in nomes and i not in idx.values():
                idx[chave] = i
                break
    if len(idx) == 4:
        corpo, primeira = linhas[1:], 2
    elif len(idx) == 0:
        idx = {"data": 0, "categoria": 1, "descricao": 2, "valor": 3}
        corpo, primeira = linhas, 1
    else:
        faltam = [k for k in CABECALHOS if k not in idx]
        nomes = {"data": "Data", "categoria": "Categoria", "descricao": "Descrição", "valor": "Valor"}
        return [], ["Coluna(s) não encontrada(s) no cabeçalho: " + ", ".join(nomes[k] for k in faltam)]

    lancs, erros = [], []
    for n, r in enumerate(corpo, start=primeira):
        r = list(r) + [None] * 4

        def cel(k):
            return r[idx[k]]
        try:
            ano, mes = ler_data(cel("data"))
            cat = mapear_categoria(cel("categoria"))
            if cat is None:
                raise ValueError(f"categoria \"{cel('categoria')}\" inválida (use Fixa, Variável ou Adicional)")
            desc = str(cel("descricao") or "").strip()
            if not desc:
                raise ValueError("descrição vazia")
            v = cel("valor")
            valor = round(float(v), 2) if isinstance(v, (int, float)) else ler_valor(str(v))
            lancs.append({"ano": ano, "mes": mes, "categoria": cat, "descricao": desc,
                          "valor": valor, "linha": n})
        except (ValueError, TypeError) as e:
            erros.append(f"Linha {n}: {e}")
    return lancs, erros


def agrupar_lancamentos(lancs):
    """Soma lançamentos com a mesma descrição no mesmo mês/categoria."""
    grupos = {}
    for l in lancs:
        k = (l["ano"], l["mes"], l["categoria"], l["descricao"])
        grupos[k] = round(grupos.get(k, 0.0) + l["valor"], 2)
    return grupos


NOME_CATEGORIA_PLANILHA = {"fixas": "Fixa", "variaveis": "Variável", "adicionais": "Adicional"}

TEXTO_FORMATO = """A planilha pode ser um arquivo Excel (.xlsx) ou CSV (.csv).

Cada linha é uma despesa, com 4 colunas:

  Data        Data da despesa. Ex.: 15/03/2026, 2026-03-15,
              03/2026 ou mar/2026. Só o mês e o ano são usados.
  Categoria   Fixa, Variável ou Adicional (o plural e "Extra"
              também valem; maiúsculas e acentos não importam).
  Descrição   Nome da despesa. Ex.: Condomínio, Uber, Doações.
  Valor       Valor em reais. Ex.: 1234,56 ou R$ 1.234,56.

Exemplo:

  Data         Categoria   Descrição           Valor
  05/01/2026   Fixa        Condomínio          506,59
  12/01/2026   Variável    Cartão de Crédito   4.045,60
  20/01/2026   Adicional   Doações             68,00

Regras:
  • A primeira linha deve ser o cabeçalho com os nomes das colunas
    (em qualquer ordem). Sem cabeçalho, use a ordem acima.
  • Colunas a mais (ex.: Observação) são ignoradas.
  • Despesas com a mesma descrição no mesmo mês e categoria são
    somadas.
  • Linhas com erro são listadas e ignoradas; antes de gravar você
    vê uma prévia e decide se substitui ou soma valores que já
    existem no histórico.
  • Um arquivo gerado em "Exportar histórico" pode ser importado
    de volta."""


def exportar_historico(db, caminho):
    """Exporta todas as despesas (formato de importação) e um resumo mensal."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    cab_font, cab_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2E7D5B")

    ws = wb.active
    ws.title = "Despesas"
    ws.append(["Data", "Categoria", "Descrição", "Valor", "Observação"])
    ordem = {c: i for i, (c, _, _) in enumerate(CATEGORIAS)}
    linhas = db.con.execute(
        "SELECT ano, mes, categoria, descricao, valor, observacao FROM despesas").fetchall()
    linhas.sort(key=lambda r: (r[0], r[1], ordem.get(r[2], 9), r[3].lower()))
    for ano, mes, cat, desc, valor, obs in linhas:
        ws.append([dt.date(ano, mes, 1), NOME_CATEGORIA_PLANILHA.get(cat, cat), desc, valor, obs or ""])
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"
        row[3].number_format = "#,##0.00"
    for col, w in zip("ABCDE", (14, 14, 34, 14, 30)):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet("Resumo mensal")
    ws2.append(["Mês/Ano", "Fixas", "Variáveis", "Adicionais", "Total", "Salário líquido", "Saldo"])
    for h in db.historico():
        ws2.append([dt.date(h["ano"], h["mes"], 1), h["fixas"], h["variaveis"], h["adicionais"],
                    h["total"], h["salario"], h["saldo"]])
    for row in ws2.iter_rows(min_row=2):
        row[0].number_format = "MM/YYYY"
        for c in row[1:]:
            c.number_format = "#,##0.00"
    for col in "ABCDEFG":
        ws2.column_dimensions[col].width = 16

    for w in (ws, ws2):
        for c in w[1]:
            c.font, c.fill = cab_font, cab_fill
        w.freeze_panes = "A2"
    wb.save(caminho)
    return len(linhas)


def salvar_modelo(caminho):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Despesas"
    ws.append(["Data", "Categoria", "Descrição", "Valor"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E7D5B")
    exemplos = [
        (dt.date(2026, 1, 5), "Fixa", "Condomínio", 506.59),
        (dt.date(2026, 1, 10), "Fixa", "Luz", 147.06),
        (dt.date(2026, 1, 12), "Variável", "Cartão de Crédito", 4045.60),
        (dt.date(2026, 1, 20), "Adicional", "Doações", 68.00),
    ]
    for e in exemplos:
        ws.append(list(e))
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "DD/MM/YYYY"
        row[3].number_format = '#,##0.00'
    for col, w in zip("ABCD", (14, 14, 34, 14)):
        ws.column_dimensions[col].width = w
    wb.save(caminho)


# --------------------------------------------------------------------------
# Banco de dados
# --------------------------------------------------------------------------
class Banco:
    def __init__(self, caminho):
        self.caminho = caminho
        self.con = sqlite3.connect(caminho)
        self.con.execute("PRAGMA foreign_keys = ON")
        self._criar()

    def _criar(self):
        c = self.con
        c.executescript("""
        CREATE TABLE IF NOT EXISTS despesas (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            ano       INTEGER NOT NULL,
            mes       INTEGER NOT NULL,
            categoria TEXT    NOT NULL,
            descricao TEXT    NOT NULL,
            valor     REAL    NOT NULL,
            observacao TEXT   DEFAULT '',
            UNIQUE(ano, mes, categoria, descricao)
        );
        CREATE TABLE IF NOT EXISTS salarios (
            ano   INTEGER NOT NULL,
            mes   INTEGER NOT NULL,
            valor REAL    NOT NULL,
            PRIMARY KEY (ano, mes)
        );
        CREATE TABLE IF NOT EXISTS itens (
            categoria TEXT NOT NULL,
            descricao TEXT NOT NULL,
            PRIMARY KEY (categoria, descricao)
        );
        """)
        if c.execute("SELECT COUNT(*) FROM itens").fetchone()[0] == 0:
            for cat, itens in ITENS_PADRAO.items():
                c.executemany("INSERT OR IGNORE INTO itens VALUES (?, ?)",
                              [(cat, i) for i in itens])
        c.commit()

    # --- itens sugeridos
    def itens(self, categoria):
        rows = self.con.execute(
            "SELECT descricao FROM itens WHERE categoria=? "
            "UNION SELECT DISTINCT descricao FROM despesas WHERE categoria=? "
            "ORDER BY 1 COLLATE NOCASE", (categoria, categoria)).fetchall()
        return [r[0] for r in rows]

    # --- despesas
    def despesas(self, ano, mes, categoria):
        return self.con.execute(
            "SELECT id, descricao, valor, observacao FROM despesas "
            "WHERE ano=? AND mes=? AND categoria=? ORDER BY descricao COLLATE NOCASE",
            (ano, mes, categoria)).fetchall()

    def salvar_despesa(self, ano, mes, categoria, descricao, valor, obs=""):
        self.con.execute(
            "INSERT INTO despesas (ano, mes, categoria, descricao, valor, observacao) "
            "VALUES (?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(ano, mes, categoria, descricao) "
            "DO UPDATE SET valor=excluded.valor, observacao=excluded.observacao",
            (ano, mes, categoria, descricao, valor, obs))
        self.con.execute("INSERT OR IGNORE INTO itens VALUES (?, ?)", (categoria, descricao))
        self.con.commit()

    def excluir_despesa(self, id_):
        self.con.execute("DELETE FROM despesas WHERE id=?", (id_,))
        self.con.commit()

    def copiar_mes(self, ano_o, mes_o, ano_d, mes_d, categoria):
        cur = self.con.execute(
            "INSERT OR IGNORE INTO despesas (ano, mes, categoria, descricao, valor, observacao) "
            "SELECT ?, ?, categoria, descricao, valor, observacao FROM despesas "
            "WHERE ano=? AND mes=? AND categoria=?",
            (ano_d, mes_d, ano_o, mes_o, categoria))
        self.con.commit()
        return cur.rowcount

    def total(self, ano, mes, categoria):
        return self.con.execute(
            "SELECT COALESCE(SUM(valor), 0) FROM despesas WHERE ano=? AND mes=? AND categoria=?",
            (ano, mes, categoria)).fetchone()[0]

    def importar(self, grupos, modo):
        """modo: 'somar' (soma ao valor já existente) ou 'substituir'.
        Retorna (novos, alterados)."""
        novos = alterados = 0
        for (ano, mes, cat, desc), valor in grupos.items():
            atual = self.con.execute(
                "SELECT valor FROM despesas WHERE ano=? AND mes=? AND categoria=? AND descricao=?",
                (ano, mes, cat, desc)).fetchone()
            if atual is None:
                novos += 1
            else:
                alterados += 1
                if modo == "somar":
                    valor = round(atual[0] + valor, 2)
            self.con.execute(
                "INSERT INTO despesas (ano, mes, categoria, descricao, valor) VALUES (?, ?, ?, ?, ?) "
                "ON CONFLICT(ano, mes, categoria, descricao) DO UPDATE SET valor=excluded.valor",
                (ano, mes, cat, desc, valor))
            self.con.execute("INSERT OR IGNORE INTO itens VALUES (?, ?)", (cat, desc))
        self.con.commit()
        return novos, alterados

    def conflitos(self, grupos):
        n = 0
        for (ano, mes, cat, desc) in grupos:
            if self.con.execute("SELECT 1 FROM despesas WHERE ano=? AND mes=? AND categoria=? "
                                "AND descricao=?", (ano, mes, cat, desc)).fetchone():
                n += 1
        return n

    # --- salário
    def salario(self, ano, mes):
        r = self.con.execute("SELECT valor FROM salarios WHERE ano=? AND mes=?",
                             (ano, mes)).fetchone()
        return r[0] if r else None

    def salvar_salario(self, ano, mes, valor):
        self.con.execute("INSERT INTO salarios VALUES (?, ?, ?) "
                         "ON CONFLICT(ano, mes) DO UPDATE SET valor=excluded.valor",
                         (ano, mes, valor))
        self.con.commit()

    # --- histórico
    def historico(self):
        """Lista de dicts por mês (ordem cronológica) com totais e salário."""
        dados = {}
        for ano, mes, cat, soma in self.con.execute(
                "SELECT ano, mes, categoria, SUM(valor) FROM despesas GROUP BY ano, mes, categoria"):
            dados.setdefault((ano, mes), {})[cat] = soma
        for ano, mes, valor in self.con.execute("SELECT ano, mes, valor FROM salarios"):
            dados.setdefault((ano, mes), {})["salario"] = valor
        saida = []
        for (ano, mes) in sorted(dados):
            d = dados[(ano, mes)]
            linha = {"ano": ano, "mes": mes}
            for cat, _, _ in CATEGORIAS:
                linha[cat] = d.get(cat, 0.0)
            linha["total"] = sum(linha[c] for c, _, _ in CATEGORIAS)
            linha["salario"] = d.get("salario")
            linha["saldo"] = (linha["salario"] - linha["total"]) if linha["salario"] is not None else None
            saida.append(linha)
        return saida


# --------------------------------------------------------------------------
# Aba de categoria
# --------------------------------------------------------------------------
class AbaCategoria(ttk.Frame):
    def __init__(self, master, app, categoria, titulo):
        super().__init__(master, padding=10)
        self.app = app
        self.categoria = categoria
        self.titulo = titulo
        self.id_editando = None

        ttk.Label(self, text=titulo, style="Titulo.TLabel").pack(anchor="w")

        # Formulário
        form = ttk.LabelFrame(self, text="Lançar despesa", padding=8)
        form.pack(fill="x", pady=(8, 8))
        ttk.Label(form, text="Descrição:").grid(row=0, column=0, sticky="w")
        self.cb_desc = ttk.Combobox(form, width=38)
        self.cb_desc.grid(row=0, column=1, sticky="we", padx=(4, 12))
        ttk.Label(form, text="Valor (R$):").grid(row=0, column=2, sticky="w")
        self.en_valor = ttk.Entry(form, width=14, justify="right")
        self.en_valor.grid(row=0, column=3, padx=(4, 12))
        ttk.Label(form, text="Observação:").grid(row=0, column=4, sticky="w")
        self.en_obs = ttk.Entry(form, width=24)
        self.en_obs.grid(row=0, column=5, sticky="we", padx=(4, 12))
        self.bt_salvar = ttk.Button(form, text="Adicionar", command=self.salvar,
                                    style="Acao.TButton")
        self.bt_salvar.grid(row=0, column=6)
        ttk.Button(form, text="Limpar", command=self.limpar).grid(row=0, column=7, padx=(6, 0))
        form.columnconfigure(1, weight=2)
        form.columnconfigure(5, weight=1)
        for w in (self.cb_desc, self.en_valor, self.en_obs):
            w.bind("<Return>", lambda e: self.salvar())

        # Tabela
        quadro = ttk.Frame(self)
        quadro.pack(fill="both", expand=True)
        self.tabela = ttk.Treeview(quadro, columns=("desc", "valor", "obs"),
                                   show="headings", selectmode="browse")
        self.tabela.heading("desc", text="Descrição")
        self.tabela.heading("valor", text="Valor")
        self.tabela.heading("obs", text="Observação")
        self.tabela.column("desc", width=320)
        self.tabela.column("valor", width=130, anchor="e")
        self.tabela.column("obs", width=260)
        sb = ttk.Scrollbar(quadro, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=sb.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tabela.bind("<<TreeviewSelect>>", self.selecionar)
        self.tabela.bind("<Delete>", lambda e: self.excluir())

        # Rodapé
        rod = ttk.Frame(self)
        rod.pack(fill="x", pady=(8, 0))
        ttk.Button(rod, text="Excluir selecionada", command=self.excluir).pack(side="left")
        ttk.Button(rod, text="Copiar do mês anterior", command=self.copiar_anterior).pack(
            side="left", padx=6)
        self.lb_total = ttk.Label(rod, text="Total: R$ 0,00", style="Total.TLabel")
        self.lb_total.pack(side="right")

    def atualizar(self):
        ano, mes = self.app.periodo()
        self.cb_desc["values"] = self.app.db.itens(self.categoria)
        self.tabela.delete(*self.tabela.get_children())
        for id_, desc, valor, obs in self.app.db.despesas(ano, mes, self.categoria):
            self.tabela.insert("", "end", iid=str(id_), values=(desc, formatar_brl(valor), obs or ""))
        total = self.app.db.total(ano, mes, self.categoria)
        self.lb_total.config(text=f"Total de {MESES[mes - 1]}/{ano}: {formatar_brl(total)}")
        self.limpar()

    def limpar(self):
        self.id_editando = None
        self.cb_desc.set("")
        self.en_valor.delete(0, "end")
        self.en_obs.delete(0, "end")
        self.bt_salvar.config(text="Adicionar")
        if self.tabela.selection():
            self.tabela.selection_remove(self.tabela.selection())

    def selecionar(self, _evt=None):
        sel = self.tabela.selection()
        if not sel:
            return
        desc, valor, obs = self.tabela.item(sel[0], "values")
        self.id_editando = int(sel[0])
        self.cb_desc.set(desc)
        self.en_valor.delete(0, "end")
        self.en_valor.insert(0, valor.replace("R$ ", "").replace("-R$ ", "-"))
        self.en_obs.delete(0, "end")
        self.en_obs.insert(0, obs)
        self.bt_salvar.config(text="Atualizar")

    def salvar(self):
        desc = self.cb_desc.get().strip()
        if not desc:
            messagebox.showwarning(APP_NAME, "Informe a descrição da despesa.")
            return
        try:
            valor = ler_valor(self.en_valor.get())
        except ValueError:
            messagebox.showwarning(APP_NAME, "Valor inválido. Use, por exemplo, 1.234,56")
            return
        ano, mes = self.app.periodo()
        if self.id_editando is not None:
            self.app.db.excluir_despesa(self.id_editando)
        self.app.db.salvar_despesa(ano, mes, self.categoria, desc, valor, self.en_obs.get().strip())
        self.app.atualizar_tudo()
        self.cb_desc.focus_set()

    def excluir(self):
        sel = self.tabela.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Selecione uma despesa na tabela.")
            return
        desc = self.tabela.item(sel[0], "values")[0]
        if messagebox.askyesno(APP_NAME, f"Excluir \"{desc}\" deste mês?"):
            self.app.db.excluir_despesa(int(sel[0]))
            self.app.atualizar_tudo()

    def copiar_anterior(self):
        ano, mes = self.app.periodo()
        ano_o, mes_o = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
        n = self.app.db.copiar_mes(ano_o, mes_o, ano, mes, self.categoria)
        if n:
            messagebox.showinfo(APP_NAME, f"{n} despesa(s) copiada(s) de "
                                          f"{MESES[mes_o - 1]}/{ano_o}. Ajuste os valores se necessário.")
        else:
            messagebox.showinfo(APP_NAME, f"Nada novo para copiar de {MESES[mes_o - 1]}/{ano_o}.")
        self.app.atualizar_tudo()


# --------------------------------------------------------------------------
# Aba de resumo
# --------------------------------------------------------------------------
class AbaResumo(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=10)
        self.app = app

        topo = ttk.Frame(self)
        topo.pack(fill="x")

        # Quadro do mês
        mesbox = ttk.LabelFrame(topo, text="Mês selecionado", padding=10)
        mesbox.pack(side="left", fill="y")
        ttk.Label(mesbox, text="Salário líquido (R$):").grid(row=0, column=0, sticky="w")
        self.en_sal = ttk.Entry(mesbox, width=14, justify="right")
        self.en_sal.grid(row=0, column=1, padx=4)
        self.en_sal.bind("<Return>", lambda e: self.salvar_salario())
        ttk.Button(mesbox, text="Salvar salário", command=self.salvar_salario,
                   style="Acao.TButton").grid(row=0, column=2)
        ttk.Button(mesbox, text="Usar do mês anterior",
                   command=self.salario_anterior).grid(row=0, column=3, padx=(6, 0))

        self.labels = {}
        linhas = [(c, t) for c, t, _ in CATEGORIAS] + [("total", "Total de despesas"),
                                                         ("salario", "Salário líquido"),
                                                         ("saldo", "Saldo remanescente")]
        for i, (chave, texto) in enumerate(linhas, start=1):
            estilo = "Saldo.TLabel" if chave == "saldo" else (
                "Total.TLabel" if chave == "total" else "TLabel")
            ttk.Label(mesbox, text=texto + ":", style=estilo).grid(
                row=i, column=0, sticky="w", pady=2)
            lb = ttk.Label(mesbox, text="-", style=estilo, anchor="e")
            lb.grid(row=i, column=1, columnspan=2, sticky="e", pady=2)
            self.labels[chave] = lb
            if chave == "adicionais":
                ttk.Separator(mesbox).grid(row=i, column=0, columnspan=4, sticky="sew")

        # Gráfico
        grafbox = ttk.LabelFrame(topo, text="Evolução das despesas", padding=6)
        grafbox.pack(side="left", fill="both", expand=True, padx=(10, 0))
        ctrl = ttk.Frame(grafbox)
        ctrl.pack(fill="x")
        ttk.Label(ctrl, text="Período:").pack(side="left")
        self.cb_periodo = ttk.Combobox(ctrl, state="readonly", width=18,
                                       values=["Últimos 12 meses", "Ano selecionado", "Todo o histórico"])
        self.cb_periodo.current(0)
        self.cb_periodo.pack(side="left", padx=4)
        self.cb_periodo.bind("<<ComboboxSelected>>", lambda e: self.atualizar())
        self.var_sal = tk.BooleanVar(value=True)
        ttk.Checkbutton(ctrl, text="Mostrar salário", variable=self.var_sal,
                        command=self.atualizar).pack(side="left", padx=8)

        self.fig = Figure(figsize=(7, 3.6), dpi=96)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=grafbox)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Tabela do histórico
        hist = ttk.LabelFrame(self, text="Histórico mensal", padding=6)
        hist.pack(fill="both", expand=True, pady=(10, 0))
        cols = ("periodo", "fixas", "variaveis", "adicionais", "total", "salario", "saldo")
        nomes = ("Mês/Ano", "Fixas", "Variáveis", "Adicionais/Extras", "Total", "Salário", "Saldo")
        self.tabela = ttk.Treeview(hist, columns=cols, show="headings", height=7)
        for c, n in zip(cols, nomes):
            self.tabela.heading(c, text=n)
            self.tabela.column(c, width=120, anchor="w" if c == "periodo" else "e")
        sb = ttk.Scrollbar(hist, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=sb.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tabela.tag_configure("negativo", foreground=COR_TOTAL)
        self.tabela.bind("<Double-1>", self.ir_para_mes)

    def salvar_salario(self):
        try:
            valor = ler_valor(self.en_sal.get())
        except ValueError:
            messagebox.showwarning(APP_NAME, "Salário inválido. Use, por exemplo, 12.800,00")
            return
        ano, mes = self.app.periodo()
        self.app.db.salvar_salario(ano, mes, valor)
        self.atualizar()

    def salario_anterior(self):
        ano, mes = self.app.periodo()
        ano_o, mes_o = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
        v = self.app.db.salario(ano_o, mes_o)
        if v is None:
            messagebox.showinfo(APP_NAME, f"Não há salário salvo em {MESES[mes_o - 1]}/{ano_o}.")
            return
        self.app.db.salvar_salario(ano, mes, v)
        self.atualizar()

    def ir_para_mes(self, _evt=None):
        sel = self.tabela.selection()
        if sel:
            ano, mes = map(int, sel[0].split("-"))
            self.app.definir_periodo(ano, mes)

    def atualizar(self):
        db = self.app.db
        ano, mes = self.app.periodo()

        # Quadro do mês
        totais = {c: db.total(ano, mes, c) for c, _, _ in CATEGORIAS}
        total = sum(totais.values())
        sal = db.salario(ano, mes)
        for c in totais:
            self.labels[c].config(text=formatar_brl(totais[c]))
        self.labels["total"].config(text=formatar_brl(total))
        self.en_sal.delete(0, "end")
        if sal is not None:
            self.en_sal.insert(0, formatar_brl(sal).replace("R$ ", ""))
            self.labels["salario"].config(text=formatar_brl(sal))
            saldo = sal - total
            self.labels["saldo"].config(text=formatar_brl(saldo),
                                        foreground=COR_SALARIO if saldo >= 0 else COR_TOTAL)
        else:
            self.labels["salario"].config(text="não informado")
            self.labels["saldo"].config(text="-", foreground="")

        # Histórico
        hist = db.historico()
        self.tabela.delete(*self.tabela.get_children())
        for h in reversed(hist):
            neg = h["saldo"] is not None and h["saldo"] < 0
            self.tabela.insert("", "end", iid=f"{h['ano']}-{h['mes']}", tags=("negativo",) if neg else (),
                               values=(f"{MESES[h['mes'] - 1]}/{h['ano']}",
                                       formatar_brl(h["fixas"]), formatar_brl(h["variaveis"]),
                                       formatar_brl(h["adicionais"]), formatar_brl(h["total"]),
                                       formatar_brl(h["salario"]) if h["salario"] is not None else "-",
                                       formatar_brl(h["saldo"]) if h["saldo"] is not None else "-"))
        self.desenhar(hist, ano, mes)

    def desenhar(self, hist, ano, mes):
        escolha = self.cb_periodo.get()
        if escolha == "Ano selecionado":
            pontos = [(ano, m) for m in range(1, 13)]
        elif escolha == "Últimos 12 meses":
            pontos = []
            a, m = ano, mes
            for _ in range(12):
                pontos.append((a, m))
                a, m = (a, m - 1) if m > 1 else (a - 1, 12)
            pontos.reverse()
        else:
            if hist:
                a, m = hist[0]["ano"], hist[0]["mes"]
                fim = max((hist[-1]["ano"], hist[-1]["mes"]), (ano, mes))
                pontos = []
                while (a, m) <= fim:
                    pontos.append((a, m))
                    a, m = (a, m + 1) if m < 12 else (a + 1, 1)
            else:
                pontos = []

        por_mes = {(h["ano"], h["mes"]): h for h in hist}
        ax = self.ax
        ax.clear()
        if not pontos or not any(p in por_mes for p in pontos):
            ax.text(0.5, 0.5, "Sem dados para o período", ha="center", va="center",
                    transform=ax.transAxes, color="#777777")
            ax.set_xticks([])
            ax.set_yticks([])
            for s in ax.spines.values():
                s.set_visible(False)
            self.canvas.draw_idle()
            return

        x = list(range(len(pontos)))
        rot = [f"{MESES_ABREV[m - 1]}/{str(a)[2:]}" for a, m in pontos]

        def serie(chave):
            return [por_mes[p][chave] if p in por_mes and por_mes[p][chave] is not None else None
                    for p in pontos]

        for chave, titulo, cor in CATEGORIAS:
            ys = serie(chave)
            ax.plot(x, [y if y is not None else float("nan") for y in ys], marker="o", ms=3.5,
                    lw=2, color=cor, label=titulo.replace("Despesas ", ""))
        ys = serie("total")
        ax.plot(x, [y if y is not None else float("nan") for y in ys], lw=2.4, color=COR_TOTAL,
                label="Total", marker="o", ms=3.5)
        if self.var_sal.get():
            ys = serie("salario")
            ax.plot(x, [y if y is not None else float("nan") for y in ys], lw=1.6, ls="--",
                    color=COR_SALARIO, label="Salário líquido")
        if (ano, mes) in pontos:
            ax.axvline(pontos.index((ano, mes)), color="#999999", lw=0.8, ls=":")

        ax.set_xticks(x)
        ax.set_xticklabels(rot, rotation=45 if len(x) > 12 else 0, fontsize=8)
        if len(x) > 24:
            for i, lbl in enumerate(ax.get_xticklabels()):
                lbl.set_visible(i % 3 == 0)
        ax.yaxis.set_major_formatter(FuncFormatter(
            lambda v, _: f"{v:,.0f}".replace(",", ".")))
        ax.tick_params(axis="y", labelsize=8)
        ax.grid(axis="y", color="#e5e5e5", lw=0.8)
        ax.set_axisbelow(True)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        ax.set_ylim(bottom=0)
        ax.legend(fontsize=8, frameon=False, ncol=3, loc="lower center",
                  bbox_to_anchor=(0.5, 1.0))
        self.fig.tight_layout(rect=(0, 0, 1, 1))
        self.canvas.draw_idle()



# --------------------------------------------------------------------------
# Janela de importação
# --------------------------------------------------------------------------
class JanelaImportacao(tk.Toplevel):
    def __init__(self, app, caminho, grupos, lancs, erros):
        super().__init__(app)
        self.app = app
        self.grupos = grupos
        self.title("Importar planilha")
        self.transient(app)
        self.resizable(True, True)
        self.geometry("720x520")
        app.aplicar_icone(self)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text=os.path.basename(caminho), style="Titulo.TLabel").pack(anchor="w")
        meses = sorted({(a, m) for (a, m, _, _) in grupos})
        periodo = (f"{MESES[meses[0][1]-1]}/{meses[0][0]} a {MESES[meses[-1][1]-1]}/{meses[-1][0]}"
                   if meses else "-")
        total = sum(grupos.values())
        por_cat = {c: sum(v for (_, _, cc, _), v in grupos.items() if cc == c) for c, _, _ in CATEGORIAS}
        txt = (f"{len(lancs)} linha(s) válida(s), {len(grupos)} lançamento(s) após somar descrições "
               f"repetidas no mesmo mês.\nPeríodo: {periodo} ({len(meses)} mês(es))\n"
               + "   ".join(f"{t.replace('Despesas ', '')}: {formatar_brl(por_cat[c])}"
                          for c, t, _ in CATEGORIAS)
               + f"   Total: {formatar_brl(total)}")
        ttk.Label(f, text=txt, justify="left").pack(anchor="w", pady=(6, 6))

        # Prévia
        cols = ("periodo", "categoria", "descricao", "valor")
        tv = ttk.Treeview(f, columns=cols, show="headings", height=9)
        for c, n, w, a in zip(cols, ("Mês/Ano", "Categoria", "Descrição", "Valor"),
                              (110, 110, 300, 120), ("w", "w", "w", "e")):
            tv.heading(c, text=n)
            tv.column(c, width=w, anchor=a)
        nome_cat = {c: t.replace("Despesas ", "") for c, t, _ in CATEGORIAS}
        for (a, m, c, d), v in sorted(grupos.items()):
            tv.insert("", "end", values=(f"{MESES_ABREV[m-1]}/{a}", nome_cat[c], d, formatar_brl(v)))
        tv.pack(fill="both", expand=True)

        if erros:
            ttk.Label(f, text=f"{len(erros)} linha(s) ignorada(s):", foreground=COR_TOTAL).pack(
                anchor="w", pady=(8, 0))
            caixa = tk.Text(f, height=4, wrap="word")
            caixa.insert("1.0", "\n".join(erros[:200]))
            caixa.config(state="disabled")
            caixa.pack(fill="x")

        n_conf = app.db.conflitos(grupos)
        self.modo = tk.StringVar(value="substituir")
        if n_conf:
            box = ttk.LabelFrame(f, text=f"{n_conf} lançamento(s) já existem no banco (mesmo mês, "
                                         f"categoria e descrição)", padding=6)
            box.pack(fill="x", pady=(8, 0))
            ttk.Radiobutton(box, text="Substituir pelo valor da planilha", value="substituir",
                            variable=self.modo).pack(anchor="w")
            ttk.Radiobutton(box, text="Somar ao valor já existente", value="somar",
                            variable=self.modo).pack(anchor="w")

        bts = ttk.Frame(f)
        bts.pack(fill="x", pady=(10, 0))
        ttk.Button(bts, text="Cancelar", command=self.destroy).pack(side="right")
        b = ttk.Button(bts, text="Importar", style="Acao.TButton", command=self.importar)
        b.pack(side="right", padx=6)
        if not grupos:
            b.state(["disabled"])
        self.grab_set()

    def importar(self):
        novos, alterados = self.app.db.importar(self.grupos, self.modo.get())
        self.destroy()
        ultimo = max((a, m) for (a, m, _, _) in self.grupos)
        self.app.definir_periodo(*ultimo)
        messagebox.showinfo(APP_NAME, f"Importação concluída.\n{novos} lançamento(s) novo(s), "
                                      f"{alterados} atualizado(s).")


# --------------------------------------------------------------------------
# Aplicativo
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, caminho_db=None):
        super().__init__()
        self.title(APP_NAME)
        self.aplicar_icone(self)
        self.geometry("1150x760")
        self.minsize(950, 620)
        self.db = Banco(caminho_db or os.path.join(pasta_dados(), "financas.db"))
        self._estilos()

        # Barra de período
        barra = ttk.Frame(self, padding=(10, 8))
        barra.pack(fill="x")
        ttk.Label(barra, text="Mês:").pack(side="left")
        hoje = dt.date.today()
        self.cb_mes = ttk.Combobox(barra, values=MESES, state="readonly", width=12)
        self.cb_mes.current(hoje.month - 1)
        self.cb_mes.pack(side="left", padx=(4, 12))
        ttk.Label(barra, text="Ano:").pack(side="left")
        self.var_ano = tk.IntVar(value=hoje.year)
        self.sp_ano = ttk.Spinbox(barra, from_=2000, to=2100, textvariable=self.var_ano, width=6,
                                  command=self.atualizar_tudo)
        self.sp_ano.pack(side="left", padx=4)
        ttk.Button(barra, text="◀", width=3, command=lambda: self.mover(-1)).pack(side="left", padx=(12, 2))
        ttk.Button(barra, text="▶", width=3, command=lambda: self.mover(1)).pack(side="left")
        ttk.Button(barra, text="Hoje", command=lambda: self.definir_periodo(hoje.year, hoje.month)).pack(
            side="left", padx=6)
        ttk.Button(barra, text="Importar dados", command=self.importar_planilha).pack(
            side="left", padx=(18, 0))
        ttk.Button(barra, text="ⓘ", width=3, command=self.ajuda_importacao).pack(side="left", padx=(2, 8))
        ttk.Button(barra, text="Exportar histórico", command=self.exportar_historico).pack(side="left")
        ttk.Label(barra, text=f"Banco: {self.db.caminho}", foreground="#888888").pack(side="right")
        self.cb_mes.bind("<<ComboboxSelected>>", lambda e: self.atualizar_tudo())
        self.sp_ano.bind("<Return>", lambda e: self.atualizar_tudo())
        self.sp_ano.bind("<FocusOut>", lambda e: self.atualizar_tudo())

        # Abas
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.abas = []
        for chave, titulo, _ in CATEGORIAS:
            aba = AbaCategoria(self.nb, self, chave, titulo)
            self.nb.add(aba, text=f"  {titulo.replace('Despesas ', '')}  ")
            self.abas.append(aba)
        self.resumo = AbaResumo(self.nb, self)
        self.nb.add(self.resumo, text="  Resumo  ")
        self.atualizar_tudo()

    def _estilos(self):
        st = ttk.Style(self)
        if "vista" in st.theme_names():
            st.theme_use("vista")
        elif "clam" in st.theme_names():
            st.theme_use("clam")
        st.configure("Titulo.TLabel", font=("Segoe UI", 14, "bold"))
        st.configure("Total.TLabel", font=("Segoe UI", 11, "bold"))
        st.configure("Saldo.TLabel", font=("Segoe UI", 13, "bold"))
        st.configure("Treeview", rowheight=24)

    def aplicar_icone(self, janela):
        try:
            ico = caminho_recurso("icone.ico")
            if sys.platform.startswith("win") and os.path.exists(ico):
                janela.iconbitmap(ico)
            else:
                png = caminho_recurso("icone.png")
                if os.path.exists(png):
                    self._icone_img = getattr(self, "_icone_img", None) or tk.PhotoImage(file=png)
                    janela.iconphoto(False, self._icone_img)
        except tk.TclError:
            pass

    def importar_planilha(self):
        caminho = filedialog.askopenfilename(
            parent=self, title="Importar despesas de planilha",
            filetypes=[("Planilhas", "*.xlsx *.xlsm *.csv"), ("Excel", "*.xlsx *.xlsm"),
                       ("CSV", "*.csv"), ("Todos os arquivos", "*.*")])
        if not caminho:
            return
        try:
            lancs, erros = interpretar_planilha(ler_linhas_arquivo(caminho))
        except Exception as e:  # arquivo corrompido, aberto no Excel etc.
            messagebox.showerror(APP_NAME, f"Não foi possível ler a planilha:\n{e}")
            return
        if not lancs:
            messagebox.showwarning(APP_NAME, "Nenhuma despesa válida encontrada.\n\n"
                                   + "\n".join(erros[:10]) + "\n\nA planilha deve ter as colunas "
                                   "Data, Categoria, Descrição e Valor.")
            return
        JanelaImportacao(self, caminho, agrupar_lancamentos(lancs), lancs, erros)

    def ajuda_importacao(self):
        janela = tk.Toplevel(self)
        janela.title("Formato da planilha de importação")
        janela.transient(self)
        janela.resizable(False, False)
        self.aplicar_icone(janela)
        f = ttk.Frame(janela, padding=14)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Como montar a planilha para importar", style="Titulo.TLabel").pack(anchor="w")
        caixa = tk.Text(f, width=70, height=30, wrap="none", relief="flat",
                        font=("Consolas", 10) if sys.platform.startswith("win") else ("DejaVu Sans Mono", 9),
                        background=janela.cget("background"))
        caixa.insert("1.0", TEXTO_FORMATO)
        caixa.config(state="disabled")
        caixa.pack(pady=(8, 8))
        bts = ttk.Frame(f)
        bts.pack(fill="x")
        ttk.Button(bts, text="Fechar", command=janela.destroy).pack(side="right")
        ttk.Button(bts, text="Salvar planilha modelo...",
                   command=lambda: self.salvar_modelo(janela)).pack(side="right", padx=6)
        janela.bind("<Escape>", lambda e: janela.destroy())
        janela.grab_set()
        janela.focus_set()

    def exportar_historico(self):
        if not self.db.con.execute("SELECT 1 FROM despesas LIMIT 1").fetchone():
            messagebox.showinfo(APP_NAME, "Ainda não há despesas no histórico para exportar.")
            return
        caminho = filedialog.asksaveasfilename(
            parent=self, title="Exportar histórico", defaultextension=".xlsx",
            initialfile=f"historico-despesas-{dt.date.today():%Y-%m-%d}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return
        try:
            n = exportar_historico(self.db, caminho)
        except OSError as e:
            messagebox.showerror(APP_NAME, f"Não foi possível salvar (o arquivo está aberto no Excel?):\n{e}")
            return
        messagebox.showinfo(APP_NAME, f"{n} despesa(s) exportada(s) para:\n{caminho}\n\n"
                                      "A aba \"Resumo mensal\" traz os totais, o salário e o saldo de cada mês.")

    def salvar_modelo(self, parent=None):
        caminho = filedialog.asksaveasfilename(
            parent=parent or self, title="Salvar planilha modelo", defaultextension=".xlsx",
            initialfile="modelo-importacao-despesas.xlsx", filetypes=[("Excel", "*.xlsx")])
        if caminho:
            try:
                salvar_modelo(caminho)
                messagebox.showinfo(APP_NAME, f"Modelo salvo em:\n{caminho}")
            except OSError as e:
                messagebox.showerror(APP_NAME, f"Não foi possível salvar:\n{e}")

    def periodo(self):
        try:
            ano = int(self.var_ano.get())
        except (tk.TclError, ValueError):
            ano = dt.date.today().year
            self.var_ano.set(ano)
        return ano, self.cb_mes.current() + 1

    def definir_periodo(self, ano, mes):
        self.var_ano.set(ano)
        self.cb_mes.current(mes - 1)
        self.atualizar_tudo()

    def mover(self, delta):
        ano, mes = self.periodo()
        mes += delta
        if mes < 1:
            ano, mes = ano - 1, 12
        elif mes > 12:
            ano, mes = ano + 1, 1
        self.definir_periodo(ano, mes)

    def atualizar_tudo(self):
        for aba in self.abas:
            aba.atualizar()
        self.resumo.atualizar()


if __name__ == "__main__":
    App().mainloop()
