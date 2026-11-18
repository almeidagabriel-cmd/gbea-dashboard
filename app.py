"""
GBEA · Dashboard Gerencial de Fluxo de Caixa e DRE
===================================================
Executar:
    pip install -r requirements.txt
    streamlit run app.py

Entrada: extratos em .csv / .xlsx / .xls (um por unidade ou um consolidado com
coluna de empresa). Colunas reconhecidas automaticamente (ajustáveis na tela):
    Data | Empresa/Unidade (opcional) | Categoria/Plano de contas | Descrição (opcional)
    Valor (com sinal)  -ou-  Valor + Tipo (D/C)  -ou-  Entrada + Saída
"""
from __future__ import annotations

import io
import re
import unicodedata

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ════════════════════════════════════════════════════════════════════════════
# 1. CADASTROS
# ════════════════════════════════════════════════════════════════════════════
# Distribuidora = atacado da Serra Carioca · CD Produção = produção da padaria
UNIDADES = [
    "Feitoria Cozinha",
    "Casa Cônego",
    "Serra Carioca",
    "Distribuidora",
    "CD Produção",
    "Savá Padaria",
    "Rua Burguer",
    "Feito Brasa",
]
CONSOLIDADO = "Grupo GBEA"

# Palavras-chave para reconhecer a unidade a partir de nome de arquivo / coluna empresa
UNIDADE_CHAVES = {
    "Feitoria Cozinha": ["feitoria"],
    "Casa Cônego": ["conego", "casa conego"],
    "Serra Carioca": ["serra"],
    "Distribuidora": ["distribuidora", "atacado"],
    "CD Produção": ["cd", "central", "producao", "cd producao"],
    "Savá Padaria": ["sava", "padaria"],
    "Rua Burguer": ["rua burguer", "burguer", "burger"],
    "Feito Brasa": ["brasa"],
}

# Paleta (validada para daltonismo – ordem fixa, não reciclar)
COR = {
    "s1": "#2a78d6", "s2": "#eb6834", "s3": "#1baf7a", "s4": "#eda100",
    "good": "#0ca30c", "critical": "#d03b3b", "neutral": "#8a8983",
}

# ════════════════════════════════════════════════════════════════════════════
# 2. DE-PARA OFICIAL GBEA
#    (termos do extrato) -> (grupo gerencial, bloco da DRE, natureza)
#    natureza: "D" saída | "C" entrada | "±" depende do sinal do extrato
# ════════════════════════════════════════════════════════════════════════════
REGRAS = [
    (["custo dos produtos", "custo produtos b&a"], "(-) CMV (Custo de Mercadoria)", "CMV", "D"),
    (["custo embalagem"], "(-) CMV / Insumos Operacionais", "CMV", "D"),
    (["impostos"], "(-) Impostos e Tributos (DAS, etc.)", "IMPOSTOS", "D"),
    (["locacao de equipamentos"], "(-) Ocupação / Estrutura (Locação de equipamentos)", "OCUPACAO", "D"),
    (["gas", "agua", "energia", "telefone", "internet"], "(-) Ocupação (Utilidades)", "OCUPACAO", "D"),
    (["aluguel"], "(-) Ocupação (Estrutura)", "OCUPACAO", "D"),
    (["contabilidade", "sistema", "alarme", "seguro", "assinaturas", "outras"],
     "(-) Despesas Administrativas", "ADMIN", "D"),
    (["custo gestao"], "(-) Despesas Administrativas / Gestão", "ADMIN", "D"),
    (["apoios e convenios"], "(-) Despesas Institucionais", "INSTITUCIONAL", "D"),
    (["mensalidade bancaria", "despesa financeira"], "(-) Despesas Financeiras", "FINANCEIRO", "D"),
    (["logistica"], "(-) Despesas Operacionais / Logística", "OPERACIONAIS", "D"),
    (["manutencao"], "(-) Manutenção de Equipamentos e Imóvel", "MANUTENCAO", "D"),
    (["compra de publicidade", "marketing", "despesa artistas"],
     "(-) Marketing e Entretenimento", "MARKETING", "D"),
    (["extra"], "(-) Pessoal (Freelancers)", "PESSOAL", "D"),
    (["salario", "prolabore", "pro labore"], "(-) Pessoal (Folha)", "PESSOAL", "D"),
    (["encargos trabalhistas"], "(-) Pessoal (Encargos)", "PESSOAL", "D"),
    (["ferias", "rescisao"], "(-) Pessoal (Provisões)", "PESSOAL", "D"),
    (["vale transporte", "vale refeicao"], "(-) Pessoal (Benefícios)", "PESSOAL", "D"),
    (["uniforme", "comissao (10%)", "comissao 10%", "comissao"],
     "(-) Pessoal / Operacionais (Uniforme, Comissão 10%)", "PESSOAL", "D"),
    (["receitas", "receita", "receitas b&a", "dinheiro", "pix", "boleto", "credito", "debito"],
     "(+) Receita Bruta Operacional", "RECEITA", "C"),
    (["receitas com antecipacao de credito"], "(+) Receita Operacional (Antecipações)", "RECEITA", "C"),
    (["receiras financeiras", "receitas financeiras"], "(+) Receitas Financeiras", "FINANCEIRO", "C"),
    (["transf entrada", "entrada b&a"], "(±) Trânsito / Contas Internas", "TRANSITO", "C"),
    (["transf saida", "saida b&a", "pagamentos b&a", "desp nao operacional"],
     "(±) Trânsito / Contas Internas", "TRANSITO", "D"),
    (["rec nao operacional", "receita nao identificada"],
     "(±) Movimentações Não Operacionais", "NAO_OP", "C"),
    (["emprestimo", "emprestimo d"], "(±) Financiamentos e Empréstimos", "EMPRESTIMO", "±"),
    (["despesa nao identificada"], "(-) Pendências de Conciliação", "PENDENCIA", "D"),
    (["ativos", "obras"], "(•) Investimentos (CAPEX)", "CAPEX", "D"),
    (["dividendos", "consumo socios"], "(-) Retiradas de Sócios (Distribuição)", "SOCIOS", "D"),
]
SEM_CLASSIF = ("(?) Sem classificação", "SEM_CLASSIF", "±")

# Blocos usados na DRE / fluxo
BLOCOS_DESP_OP = [
    ("PESSOAL", "(-) Despesas com Pessoal"),
    ("OCUPACAO", "(-) Despesas de Ocupação"),
    ("ADMIN", "(-) Despesas Administrativas"),
    ("MARKETING", "(-) Marketing e Entretenimento"),
    ("OPERACIONAIS", "(-) Operacionais / Logística"),
    ("MANUTENCAO", "(-) Manutenção"),
    ("INSTITUCIONAL", "(-) Despesas Institucionais"),
]
BLOCOS_ABAIXO = [
    ("NAO_OP", "(±) Movimentações Não Operacionais"),
    ("PENDENCIA", "(-) Pendências de Conciliação"),
    ("SEM_CLASSIF", "(?) Sem classificação (revisar)"),
    ("CAPEX", "(•) Investimentos (CAPEX)"),
    ("EMPRESTIMO", "(±) Financiamentos e Empréstimos"),
    ("SOCIOS", "(-) Retiradas de Sócios"),
]
NOME_BLOCO = dict(
    BLOCOS_DESP_OP + BLOCOS_ABAIXO
    + [("RECEITA", "Receita Bruta Operacional"), ("IMPOSTOS", "(-) Impostos e Tributos"),
       ("CMV", "(-) CMV"), ("FINANCEIRO", "(±) Resultado Financeiro"),
       ("TRANSITO", "(±) Trânsito entre empresas GBEA")]
)
MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


# ════════════════════════════════════════════════════════════════════════════
# 3. UTILITÁRIOS
# ════════════════════════════════════════════════════════════════════════════
def norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"[^a-z0-9&%() ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


INDICE = {norm(t): (g, b, n) for termos, g, b, n in REGRAS for t in termos}
GRUPOS_OPCOES = list(dict.fromkeys(g for _, g, _, _ in REGRAS))
GRUPO_INFO = {}
for _, g, b, n in REGRAS:
    GRUPO_INFO[g] = (b, "±" if g in GRUPO_INFO and GRUPO_INFO[g][1] != n else n)


def classificar(cat, extra: dict) -> tuple[str, str, str]:
    k = norm(cat)
    if k in extra:
        g = extra[k]
        return (g, *GRUPO_INFO[g])
    if k in INDICE:
        return INDICE[k]
    k2 = re.sub(r"\(.*?\)", "", k).strip()
    return INDICE.get(k2, SEM_CLASSIF)


def parse_valor(x) -> float:
    if x is None:
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip()
    if not s or s.lower() in {"nan", "none", "-"}:
        return np.nan
    neg = "-" in s or (s.startswith("(") and s.endswith(")")) or s.upper().endswith(" D")
    s = re.sub(r"[^0-9,.]", "", s)
    if not s:
        return np.nan
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        v = float(s)
    except ValueError:
        return np.nan
    return -v if neg else v


def brl(v: float, sinal=False) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "–"
    s = f"{abs(v):,.0f}".replace(",", ".")
    if v < 0:
        return f"({s})"
    return f"+{s}" if sinal and v > 0 else s


def brl_full(v: float) -> str:
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R$ {s}" if v < 0 else f"R$ {s}"


def brl_curto(v: float) -> str:
    a = abs(v)
    if a >= 1e6:
        t = f"{a / 1e6:.2f}".replace(".", ",") + " mi"
    elif a >= 1e4:
        t = f"{a / 1e3:.1f}".replace(".", ",") + " mil"
    else:
        t = f"{a:,.0f}".replace(",", ".")
    return f"-R$ {t}" if v < 0 else f"R$ {t}"


def pct(v: float) -> str:
    return "–" if v is None or not np.isfinite(v) else f"{v:.1%}".replace(".", ",")


def rotulo_mes(p: str) -> str:
    y, m = p.split("-")
    return f"{MESES_PT[int(m) - 1]}/{y[2:]}"


def match_unidade(texto) -> str | None:
    t = f" {norm(texto)} "
    t = re.sub(r"[_\-.]", " ", t)
    ordem = ["Distribuidora", "CD Produção"]
    for u in ordem + [x for x in UNIDADE_CHAVES if x not in ordem]:
        chaves = UNIDADE_CHAVES[u]
        for c in chaves:
            if f" {c} " in t or (len(c) > 3 and c in t):
                return u
    return None


# ════════════════════════════════════════════════════════════════════════════
# 4. LEITURA
