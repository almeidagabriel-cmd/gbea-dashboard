"""
GBEA · Dashboard Gerencial de Fluxo de Caixa e DRE
===================================================
Executar:
    pip install -r requirements.txt
    streamlit run app.py

Entrada: extratos em .csv / .xlsx / .xls / .pdf (um por unidade ou um consolidado com
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
from pypdf import PdfReader
import plotly.graph_objects as go
import streamlit as st

# ════════════════════════════════════════════════════════════════════════════
# 1. CADASTROS
# ════════════════════════════════════════════════════════════════════════════
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

COR = {
    "s1": "#2a78d6", "s2": "#eb6834", "s3": "#1baf7a", "s4": "#eda100",
    "good": "#0ca30c", "critical": "#d03b3b", "neutral": "#8a8983",
}

# ════════════════════════════════════════════════════════════════════════════
# 2. DE-PARA OFICIAL GBEA
# ════════════════════════════════════════════════════════════════════════════
REGRAS = [
    (["custo dos produtos", "custo produtos b&a"], "(-) CMV (Custo de Mercadoria)", "CMV", "D"),
    (["custo embalagem", "consumiveis"], "(-) CMV / Insumos Operacionais", "CMV", "D"),
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
        t = f"{a:,.0f}".replace(".", ".")
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
# 4. LEITURA DOS EXTRATOS (EXCEL, CSV E PDF)
# ════════════════════════════════════════════════════════════════════════════
CANDIDATOS = {
    "data": ["data", "dt", "venc", "data movimento", "data pagamento", "data lancamento", "competencia",
             "vencimento", "data de pagamento", "data do lancamento"],
    "empresa": ["empresa", "unidade", "loja", "filial", "estabelecimento", "cnpj empresa"],
    "categoria": ["categoria", "plano de contas", "conta", "classificacao", "categoria financeira",
                  "conta gerencial", "natureza"],
    "descricao": ["descricao", "historico", "descr", "observacao", "lancamento", "fornecedor",
                  "favorecido", "cliente fornecedor"],
    "valor": ["valor", "valor (r$)", "valor r$", "montante", "total", "valor pago", "valor liquido"],
    "tipo": ["tipo", "d c", "operacao", "tipo lancamento", "entrada saida"],
    "entrada": ["entrada", "entradas", "credito", "creditos", "recebimento"],
    "saida": ["saida", "saidas", "debito", "debitos", "pagamento"],
}


def _ler_bruto(nome: str, conteudo: bytes) -> pd.DataFrame:
    if nome.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(io.BytesIO(conteudo), header=None, dtype=object)
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(conteudo), header=None, dtype=str, sep=None,
                               engine="python", encoding=enc)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise ValueError(f"Não consegui ler {nome}")


def ler_pdf(conteudo: bytes) -> pd.DataFrame:
    """Extrai transações do novo modelo de extrato em PDF do Sicredi."""
    reader = PdfReader(io.BytesIO(conteudo))
    linhas_texto = []
    
    for pagina in reader.pages:
        txt = pagina.extract_text()
        if txt:
            for linha in txt.split("\n"):
                linha = linha.strip()
                if linha:
                    linhas_texto.append(linha)

    padrao_data = re.compile(r"^(\d{4}-\d{2}-\d{2})")
    transacoes = []
    transacao_atual = []

    for linha in linhas_texto:
        if "Data" in linha and "Lançamento" in linha:
            continue
            
        if padrao_data.match(linha):
            if transacao_atual:
                transacoes.append(" ".join(transacao_atual))
            transacao_atual = [linha]
        else:
            if transacao_atual:
                transacao_atual.append(linha)
                
    if transacao_atual:
        transacoes.append(" ".join(transacao_atual))

    parsed_rows = []
    for t in transacoes:
        match_dt = padrao_data.search(t)
        if not match_dt:
            continue
        data = match_dt.group(1)
        resto = t[len(data):].strip()

        valores_encontrados = re.findall(r"(-?R?\$?\s*-?\d{1,3}(?:\.\d{3})*,\d{2})", resto)
        
        valor_final = 0.0
        is_saida = "-" in resto or any(op in resto.upper() for op in ["DEB", "LIQUIDACAO", "ESTORNO", "SAQUE", "CESTA", "TARIF"])
        
        if valores_encontrados:
            val_str = valores_encontrados[-1]
            val_num = parse_valor(val_str)
            if not np.isnan(val_num):
                if is_saida or "-" in val_str:
                    valor_final = -abs(val_num)
                else:
                    valor_final = abs(val_num)
            descricao = resto.replace(val_str, "").strip()
        else:
            descricao = resto

        categoria_atribuida = "Outros"
        lower_desc = descricao.lower()
        if "pix" in lower_desc and "recebimento" in lower_desc:
            categoria_atribuida = "Receitas"
        elif "pix" in lower_desc and "pagamento" in lower_desc:
            categoria_atribuida = "Custo dos produtos"
        elif "boleto" in lower_desc or "liquidação" in lower_desc:
            categoria_atribuida = "Custo dos produtos"
        elif "tarifa" in lower_desc or "cesta" in lower_desc or "iof" in lower_desc:
            categoria_atribuida = "Mensalidade bancaria"
        elif "energia" in lower_desc:
            categoria_atribuida = "Energia"
        elif "aluguel" in lower_desc or "arp" in lower_desc:
            categoria_atribuida = "Aluguel"
        elif "salario" in lower_desc or "folha" in lower_desc:
            categoria_atribuida = "Salario"
        else:
            categoria_atribuida = descricao[:30]

        parsed_rows.append({
            "data": data,
            "descricao": descricao,
            "categoria": categoria_atribuida,
            "valor": valor_final
        })

    df = pd.DataFrame(parsed_rows)
    if df.empty:
        raise ValueError("Nenhuma transação válida encontrada no PDF.")
    return df


@st.cache_data(show_spinner=False)
def ler_arquivo(nome: str, conteudo: bytes) -> pd.DataFrame:
    """Lê o arquivo. Se for PDF, processa direto; se for Excel/CSV, localiza o cabeçalho."""
    if nome.lower().endswith(".pdf"):
        return ler_pdf(conteudo)
        
    bruto = _ler_bruto(nome, conteudo)
    todos = {c for v in CANDIDATOS.values() for c in v}
    linha_cab = 0
    for i in range(min(20, len(bruto))):
        hits = sum(norm(x) in todos for x in bruto.iloc[i].tolist())
        if hits >= 2:
            linha_cab = i
            break
    df = bruto.iloc[linha_cab + 1:].copy()
    cab, vistos = [], {}
    for c in bruto.iloc[linha_cab].tolist():
        c = str(c).strip() if pd.notna(c) else "coluna"
        vistos[c] = vistos.get(c, 0) + 1
        cab.append(c if vistos[c] == 1 else f"{c}_{vistos[c]}")
    df.columns = cab
    return df.dropna(how="all").reset_index(drop=True)


def detectar_colunas(cols: list[str]) -> dict:
    achado = {}
    normalizadas = {c: norm(c) for c in cols}
    for campo, cands in CANDIDATOS.items():
        achado[campo] = next((c for c in cols if normalizadas[c] in cands), None)
        if achado[campo] is None:
            achado[campo] = next((c for c in cols if any(normalizadas[c].startswith(x) for x in cands
                                                           if len(x) > 3)), None)
    if achado["valor"]:
        achado["entrada"] = achado["saida"] = None
    return achado


def padronizar(df: pd.DataFrame, cols: dict, unidade_padrao: str, extra_map: dict, origem: str):
    out = pd.DataFrame()
    s_data = df[cols["data"]].astype(str).str.strip()
    out["data"] = pd.to_datetime(s_data, dayfirst=True, errors="coerce")
    if out["data"].isna().all():
        out["data"] = pd.to_datetime(s_data.str.extract(r"(\d{2}/\d{2}/\d{4})")[0], dayfirst=True, errors="coerce")

    out["categoria"] = df[cols["categoria"]].astype(str).str.strip()
    out["descricao"] = df[cols["descricao"]].astype(str) if cols.get("descricao") else ""

    if cols.get("valor"):
        v = df[cols["valor"]].map(parse_valor)
        if cols.get("tipo"):
            t = df[cols["tipo"]].map(norm)
            saida = t.str.match(r"^(d|s|p|deb|saida|pag)")
            v = v.abs() * np.where(saida, -1, 1)
    else:
        v = df[cols["entrada"]].map(parse_valor).fillna(0) - df[cols["saida"]].map(parse_valor).fillna(0).abs()
    out["valor"] = v
    assinado = bool((v < 0).any())

    if cols.get("empresa"):
        out["unidade"] = df[cols["empresa"]].map(match_unidade).fillna(unidade_padrao)
    else:
        out["unidade"] = unidade_padrao

    cls = out["categoria"].map(lambda c: classificar(c, extra_map))
    out["grupo"] = cls.str[0]
    out["bloco"] = cls.str[1]
    nat = cls.str[2]
    if not assinado:
        out["valor"] = np.where(nat == "D", -out["valor"].abs(), out["valor"].abs())
    out["origem"] = origem
    out = out.dropna(subset=["data", "valor"])
    return out[out["valor"] != 0]


# ════════════════════════════════════════════════════════════════════════════
# 5. DADOS DE EXEMPLO
# ════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def dados_exemplo(extra: tuple = (), seed: int = 7) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    meses = pd.period_range("2026-06", "2026-11", freq="M")
    perfil = {
        "Feitoria Cozinha": (420_000, .31, .24, 28_000, False, True),
        "Casa Cônego": (360_000, .29, .26, 32_000, True, True),
        "Serra Carioca": (300_000, .33, .25, 22_000, True, True),
        "Distribuidora": (230_000, .68, .08, 15_000, False, False),
        "CD Produção": (90_000, .45, .30, 16_000, False, False),
        "Savá Padaria": (180_000, .36, .22, 12_000, False, False),
        "Rua Burguer": (240_000, .35, .21, 14_000, False, False),
        "Feito Brasa": (280_000, .38, .23, 18_000, True, True),
    }
    linhas = []

    def add(u, d, cat, val, desc=""):
        linhas.append((d, u, cat, desc or cat, round(float(val), 2)))

    for i, m in enumerate(meses):
        ini = m.to_timestamp()
        sazonal = 1 + 0.06 * np.sin(i / 1.6) + (0.12 if m.month in (11, 12) else 0)
        for u, (rec, cmv, folha, alug, artistas, dez) in perfil.items():
            r = rec * sazonal * rng.normal(1, .04)
            for semana in range(4):
                d = ini + pd.Timedelta(days=3 + 7 * semana)
                for canal, w in [("Pix", .30), ("Credito", .33), ("Debito", .20), ("Dinheiro", .05),
                                 ("Receitas", .09)]:
                    add(u, d, canal, r * w / 4 * rng.normal(1, .05), f"Vendas {canal}")
                add(u, d + pd.Timedelta(days=1), "Custo dos produtos", -r * cmv * .93 / 4 * rng.normal(1, .06),
                    "Fornecedor insumos")
            add(u, ini + pd.Timedelta(days=15), "Receitas com antecipação de crédito", r * .03, "Antecipação adquirente")
            add(u, ini + pd.Timedelta(days=9), "Custo embalagem", -r * .025 * rng.normal(1, .1))
            add(u, ini + pd.Timedelta(days=19), "Impostos", -r * .065, "DAS")
            fl = r * folha
            add(u, ini + pd.Timedelta(days=4), "Salario", -fl * .66, "Folha")
            add(u, ini + pd.Timedelta(days=4), "Prolabore", -fl * .05)
            add(u, ini + pd.Timedelta(days=19), "Encargos trabalhistas", -fl * .14, "FGTS/INSS")
            add(u, ini + pd.Timedelta(days=1), "Vale transporte", -fl * .05)
            add(u, ini + pd.Timedelta(days=1), "Vale refeição", -fl * .04)
            add(u, ini + pd.Timedelta(days=26), "Extra", -fl * .06 * rng.normal(1, .3), "Freelancers")
            if rng.random() < .3:
                add(u, ini + pd.Timedelta(days=12), "Rescisão", -fl * .10)
            if dez:
                add(u, ini + pd.Timedelta(days=5), "Comissão (10%)", -r * .055, "Repasse 10%")
            add(u, ini + pd.Timedelta(days=7), "Uniforme", -600 * rng.random())
            add(u, ini + pd.Timedelta(days=5), "Aluguel", -alug)
            add(u, ini + pd.Timedelta(days=10), "Energia", -r * .032 * rng.normal(1, .1))
            add(u, ini + pd.Timedelta(days=10), "Agua", -r * .008)
            add(u, ini + pd.Timedelta(days=10), "Gas", -r * .012)
            add(u, ini + pd.Timedelta(days=10), "Internet", -480)
            add(u, ini + pd.Timedelta(days=10), "Telefone", -260)
            add(u, ini + pd.Timedelta(days=8), "Locação de equipamentos", -1_900)
            add(u, ini + pd.Timedelta(days=6), "Contabilidade", -3_200)
            add(u, ini + pd.Timedelta(days=6), "Sistema", -890)
            add(u, ini + pd.Timedelta(days=6), "Alarme", -220)
            add(u, ini + pd.Timedelta(days=6), "Seguro", -640)
            add(u, ini + pd.Timedelta(days=6), "Assinaturas", -310)
            add(u, ini + pd.Timedelta(days=20), "Custo gestao", -r * .02, "Rateio holding")
            add(u, ini + pd.Timedelta(days=14), "Marketing", -r * .012)
            add(u, ini + pd.Timedelta(days=14), "Compra de Publicidade", -r * .006, "Meta Ads")
            if artistas:
                add(u, ini + pd.Timedelta(days=22), "Despesa artistas", -9_000 * rng.normal(1, .2), "Música ao vivo")
            add(u, ini + pd.Timedelta(days=17), "Manutencao", -r * .01 * rng.normal(1, .4))
            add(u, ini + pd.Timedelta(days=17), "Logistica", -r * (.045 if "Distribuidora" in u else .02 if "CD" in u else .005))
            add(u, ini + pd.Timedelta(days=2), "Mensalidade bancaria", -189)
            add(u, ini + pd.Timedelta(days=15), "Despesa financeira", -r * .03 * .025, "Taxa antecipação")
            add(u, ini + pd.Timedelta(days=11), "Apoios e convenios", -450)
            if u in ("Feitoria Cozinha", "Casa Cônego", "Feito Brasa"):
                add(u, ini + pd.Timedelta(days=27), "Dividendos", -15_000)
            add(u, ini + pd.Timedelta(days=28), "Consumo sócios", -900 * rng.random())
            if rng.random() < .35:
                add(u, ini + pd.Timedelta(days=21), "Despesa não identificada", -rng.uniform(300, 2_500))
            if u in ("Feito Brasa", "Savá Padaria"):
                add(u, ini + pd.Timedelta(days=13), "Taxa iFood", -r * .018, "Categoria fora do DE-PARA")
        for origem, destino, val in [("Distribuidora", "Serra Carioca", 22_000),
                                     ("Savá Padaria", "CD Produção", 35_000),
                                     ("Feitoria Cozinha", "Savá Padaria", 15_000),
                                     ("Casa Cônego", "Rua Burguer", 30_000)]:
            d = ini + pd.Timedelta(days=16)
            v = val * rng.normal(1, .1)
            add(origem, d, "Transf saida", -v, f"Transf. para {destino}")
            add(destino, d, "Transf entrada", v, f"Transf. de {origem}")
        add("Distribuidora", ini + pd.Timedelta(days=12), "Receitas financeiras", 1_150, "Rendimento aplicação")
    add("Rua Burguer", pd.Timestamp("2026-06-05"), "Emprestimo", 80_000, "Capital de giro – entrada")
    for m in meses[1:]:
        add("Rua Burguer", m.to_timestamp() + pd.Timedelta(days=9), "Emprestimo D", -4_350, "Parcela empréstimo")
    add("Rua Burguer", pd.Timestamp("2026-07-14"), "Obras", -46_000, "Reforma salão")
    add("Serra Carioca", pd.Timestamp("2026-09-03"), "Ativos", -18_500, "Forno combinado")
    add("Casa Cônego", pd.Timestamp("2026-10-18"), "Rec não operacional", 6_200, "Venda equipamento usado")

    df = pd.DataFrame(linhas, columns=["data", "unidade", "categoria", "descricao", "valor"])
    cls = df["categoria"].map(lambda c: classificar(c, dict(extra)))
    df["grupo"], df["bloco"] = cls.str[0], cls.str[1]
    df["origem"] = "exemplo"
    return df


# ════════════════════════════════════════════════════════════════════════════
# 6. CÁLCULOS GERENCIAIS
# ════════════════════════════════════════════════════════════════════════════
def soma_bloco(piv: pd.DataFrame, bloco: str) -> pd.Series:
    return piv.loc[bloco] if bloco in piv.index else pd.Series(0.0, index=piv.columns)


def montar_dre(df: pd.DataFrame, detalhe: bool) -> pd.DataFrame:
    meses = sorted(df["mes"].unique())
    piv = df.pivot_table(index="bloco", columns="mes", values="valor", aggfunc="sum").reindex(columns=meses).fillna(0)
    pivg = df.pivot_table(index=["bloco", "grupo"], columns="mes", values="valor", aggfunc="sum") \
        .reindex(columns=meses).fillna(0)
    linhas: list[tuple[str, str, pd.Series]] = []

    def bloco(b, rot=None, estilo="conta"):
        s = soma_bloco(piv, b)
        linhas.append((rot or NOME_BLOCO[b], estilo, s))
        if detalhe and b in pivg.index.get_level_values(0):
            sub = pivg.loc[b]
            if len(sub) > 1:
                for g, vals in sub.iterrows():
                    linhas.append((g, "sub", vals))
        return s

    rec = bloco("RECEITA", "Receita Bruta Operacional", "conta-forte")
    imp = bloco("IMPOSTOS")
    rl = rec + imp
    linhas.append(("(=) Receita Líquida", "total", rl))
    cmv = bloco("CMV")
    mc = rl + cmv
    linhas.append(("(=) Margem de Contribuição", "total", mc))
    desp = sum((bloco(b) for b, _ in BLOCOS_DESP_OP), pd.Series(0.0, index=meses))
    ebitda = mc + desp
    linhas.append(("(=) Resultado Operacional da Unidade (EBITDA Gerencial)", "destaque", ebitda))
    fin = bloco("FINANCEIRO")
    rf = ebitda + fin
    linhas.append(("(=) Resultado após Financeiro", "total", rf))
    linhas.append(("Movimentos de caixa abaixo do resultado", "secao", pd.Series(np.nan, index=meses)))
    abaixo = sum((bloco(b) for b, _ in BLOCOS_ABAIXO), pd.Series(0.0, index=meses))
    fcl = rf + abaixo
    linhas.append(("(=) Fluxo de Caixa Livre da Operação (sem trânsitos)", "destaque", fcl))
    tr = bloco("TRANSITO")
    linhas.append(("(=) Variação Total de Caixa (extrato)", "total", fcl + tr))

    out = pd.DataFrame([vals.reindex(meses).values for _, _, vals in linhas],
                       index=[rot for rot, _, _ in linhas], columns=meses)
    out.columns = [rotulo_mes(m) for m in meses]
    out["Total"] = out.sum(axis=1, min_count=1)
    receita_total = rec.sum()
    out["AV %"] = out["Total"] / receita_total if receita_total else np.nan
    out.insert(0, "_estilo", [e for _, e, _ in linhas])
    return out


def indicadores(df: pd.DataFrame) -> dict:
    b = df.groupby("bloco")["valor"].sum()
    g = lambda k: float(b.get(k, 0.0))  # noqa: E731
    rec = g("RECEITA")
    rl = rec + g("IMPOSTOS")
    mc = rl + g("CMV")
    ebitda = mc + sum(g(k) for k, _ in BLOCOS_DESP_OP)
    rf = ebitda + g("FINANCEIRO")
    fcl = rf + sum(g(k) for k, _ in BLOCOS_ABAIXO)
    tr = g("TRANSITO")
    folha = -g("PESSOAL")
    return dict(receita=rec, receita_liq=rl, cmv=-g("CMV"), mc=mc, ebitda=ebitda, res_fin=rf,
                fcl=fcl, transito=tr, var_total=fcl + tr, folha=folha,
                cmv_pct=-g("CMV") / rec if rec else np.nan, folha_pct=folha / rec if rec else np.nan,
                mc_pct=mc / rec if rec else np.nan, ebitda_pct=ebitda / rec if rec else np.nan,
                ocup_pct=-g("OCUPACAO") / rec if rec else np.nan,
                sem_classif=g("SEM_CLASSIF"), pendencias=g("PENDENCIA"),
                tr_entradas=float(df.loc[(df.bloco == "TRANSITO") & (df.valor > 0), "valor"].sum()),
                tr_saidas=float(df.loc[(df.bloco == "TRANSITO") & (df.valor < 0), "valor"].sum()))


# ════════════════════════════════════════════════════════════════════════════
# 7. VISUAL
# ════════════════════════════════════════════════════════════════════════════
CSS = """
<style>
.block-container {padding-top: 3.2rem; max-width: 1500px;}
.gbea-brand {display:flex; align-items:center; gap:0.8rem; margin-bottom:0.4rem}
.gbea-brand .dia {width:28px; height:28px; transform:rotate(45deg); display:grid;
  grid-template-columns:1fr 1fr; gap:2px; border-radius:4px; overflow:hidden; flex-shrink:0;}
.gbea-brand .dia span:nth-child(1){background:#2a78d6}.gbea-brand .dia span:nth-child(2){background:#eb6834}
.gbea-brand .dia span:nth-child(3){background:#1baf7a}.gbea-brand .dia span:nth-child(4){background:#eda100}
.gbea-brand h1 {font-size:1.55rem; margin:0; padding:0; line-height:1.2;}
.gbea-sub {opacity:.65; font-size:.9rem; margin-bottom:1rem}
.kpi {border:1px solid rgba(128,128,128,.22); border-radius:10px; padding:.65rem .75rem; height:100%; display:flex; flex-direction:column; justify-content:space-between; box-sizing:border-box; overflow:hidden;}
.kpi .l {font-size:.68rem; opacity:.75; text-transform:uppercase; letter-spacing:.02em; line-height:1.15; min-height:2.4em;}
.kpi .v {font-size:1.1rem; white-space:nowrap; font-weight:700; margin-top:.2rem; font-variant-numeric:tabular-nums; letter-spacing:-0.03em;}
.kpi .s {font-size:.72rem; opacity:.75; margin-top:.15rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.kpi.neg .v {color:#d03b3b}
table.dre {border-collapse:collapse; width:100%; font-size:.84rem; font-variant-numeric:tabular-nums}
table.dre th {text-align:right; padding:.45rem .6rem; border-bottom:2px solid rgba(128,128,128,.35);
  font-weight:600; white-space:nowrap; position:sticky; top:0; background:var(--background-color, inherit)}
table.dre th:first-child {text-align:left}
table.dre td {padding:.32rem .6rem; text-align:right; white-space:nowrap; border-bottom:1px solid rgba(128,128,128,.12)}
table.dre td:first-child {text-align:left; white-space:normal; min-width:280px}
table.dre tr.sub td {opacity:.72; font-size:.8rem}
table.dre tr.sub td:first-child {padding-left:1.6rem}
table.dre tr.conta-forte td {font-weight:600}
table.dre tr.total td {font-weight:650; border-top:1px solid rgba(128,128,128,.4)}
table.dre tr.destaque td {font-weight:700; background:rgba(42,120,214,.10); border-top:1px solid rgba(42,120,214,.5)}
table.dre tr.secao td {font-size:.72rem; text-transform:uppercase; letter-spacing:.05em; opacity:.6; padding-top:.9rem}
table.dre td.neg {color:#d03b3b}
table.dre td.col-total {border-left:1px solid rgba(128,128,128,.3)}
.dre-wrap {max-height:640px; overflow:auto; border:1px solid rgba(128,128,128,.18); border-radius:10px}
</style>
"""


def kpi(col, rotulo, valor, sub="", negativo=False):
    col.markdown(f'<div class="kpi{" neg" if negativo else ""}"><div class="l">{rotulo}</div>'
                 f'<div class="v">{valor}</div><div class="s">{sub}</div></div>', unsafe_allow_html=True)


def tabela_dre_html(dre: pd.DataFrame) -> str:
    cols = [c for c in dre.columns if c != "_estilo"]
    h = ['<div class="dre-wrap"><table class="dre"><thead><tr><th>Conta gerencial (R$)</th>']
    h += [f'<th>{c}</th>' for c in cols] + ["</tr></thead><tbody>"]
    for rot, row in dre.iterrows():
        est = row["_estilo"]
        h.append(f'<tr class="{est}"><td>{rot}</td>')
        for c in cols:
            v = row[c]
            if est == "secao":
                h.append("<td></td>")
                continue
            txt = pct(v) if c == "AV %" else brl(v)
            cls = ("neg " if (isinstance(v, float) and v < 0 and c != "AV %") else "")
            cls += "col-total" if c == "Total" else ""
            h.append(f'<td class="{cls}">{txt}</td>')
        h.append("</tr>")
    h.append("</tbody></table></div>")
    return "".join(h)


def layout(fig, h=360, legenda=True):
    fig.update_layout(
        height=h, margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=None) if legenda else None,
        showlegend=legenda, font=dict(size=12), separators=",.",
    )
    fig.update_xaxes(showgrid=False, linecolor="rgba(128,128,128,.4)")
    fig.update_yaxes(gridcolor="rgba(128,128,128,.15)", zeroline=True, zerolinecolor="rgba(128,128,128,.4)")
    return fig


def plot(fig):
    st.plotly_chart(fig, width="stretch", theme="streamlit", config={"displaylogo": False})


# ════════════════════════════════════════════════════════════════════════════
# 8. APP
# ════════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="GBEA · Fluxo de Caixa & DRE", page_icon="◆", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.session_state.setdefault("extra_map", {})

    # ── Sidebar: fonte de dados ───────────────────────────────────────────
    with st.sidebar:
        st.markdown('<div class="gbea-brand"><div class="dia"><span></span><span></span><span></span>'
                    '<span></span></div><h1>GBEA</h1></div>', unsafe_allow_html=True)
        st.caption("Grupo Bebidas, Entretenimento e Alimentos")
        fonte = st.radio("Fonte de dados", ["Extratos enviados", "Dados de exemplo"], index=1,
                         help="Use os dados de exemplo para validar o modelo antes de carregar os extratos reais.")
        arquivos = []
        if fonte == "Extratos enviados":
            arquivos = st.file_uploader("Extratos (.csv, .xlsx, .pdf)", type=["csv", "xlsx", "xls", "xlsm", "pdf"],
                                        accept_multiple_files=True)

    frames, avisos = [], []
    if fonte == "Dados de exemplo":
        frames.append(dados_exemplo(tuple(sorted(st.session_state["extra_map"].items()))))
    else:
        for arq in arquivos or []:
            try:
                if arq.name.lower().endswith(".pdf"):
                    df_pdf = ler_pdf(arq.getvalue())
                    u_auto = match_unidade(arq.name) or UNIDADES[0]
                    df_pdf["unidade"] = u_auto
                    df_pdf["origem"] = arq.name
                    
                    cls = df_pdf["categoria"].map(lambda c: classificar(c, st.session_state["extra_map"]))
                    df_pdf["grupo"], df_pdf["bloco"] = cls.str[0], cls.str[1]
                    df_pdf["data"] = pd.to_datetime(df_pdf["data"], errors="coerce")
                    df_pdf["mes"] = df_pdf["data"].dt.strftime("%Y-%m")
                    df_pdf = df_pdf.dropna(subset=["data", "valor"])
                    frames.append(df_pdf[df_pdf["valor"] != 0])
                    continue

                bruto = ler_arquivo(arq.name, arq.getvalue())
            except Exception as e:
                avisos.append(f"❌ {arq.name}: {e}")
                continue
            
            auto = detectar_colunas(list(bruto.columns))
            with st.sidebar.expander(f"📄 {arq.name}", expanded=not (auto["data"] and auto["categoria"]
                                                                     and (auto["valor"] or auto["entrada"]))):
                opcoes = ["—"] + list(bruto.columns)
                sel = {}
                for campo, rot in [("data", "Data"), ("categoria", "Categoria"), ("valor", "Valor"),
                                   ("tipo", "Tipo D/C (opcional)"), ("entrada", "Entrada (se sem Valor)"),
                                   ("saida", "Saída (se sem Valor)"), ("empresa", "Empresa (opcional)"),
                                   ("descricao", "Descrição (opcional)")]:
                    pad = auto.get(campo)
                    escolha = st.selectbox(rot, opcoes, index=opcoes.index(pad) if pad in opcoes else 0,
                                           key=f"{arq.name}-{campo}")
                    sel[campo] = None if escolha == "—" else escolha
                u_auto = match_unidade(arq.name) or UNIDADES[0]
                unid = st.selectbox("Unidade do arquivo" + (" (quando não houver coluna Empresa)" if sel["empresa"] else ""),
                                    UNIDADES, index=UNIDADES.index(u_auto), key=f"{arq.name}-unid")
            if not (sel["data"] and sel["categoria"] and (sel["valor"] or (sel["entrada"] and sel["saida"]))):
                avisos.append(f"⚠️ {arq.name}: indique as colunas Data, Categoria e Valor (ou Entrada/Saída).")
                continue
            frames.append(padronizar(bruto, sel, unid, st.session_state["extra_map"], arq.name))

    for a in avisos:
        st.sidebar.warning(a)

    if not frames:
        st.markdown('<div class="gbea-brand"><div class="dia"><span></span><span></span><span></span>'
                    '<span></span></div><h1>Dashboard Gerencial · Fluxo de Caixa & DRE</h1></div>',
                    unsafe_allow_html=True)
        st.info("Envie os extratos na barra lateral (ou selecione **Dados de exemplo**).")
        st.markdown(
            "**Formato esperado** – uma linha por lançamento com, no mínimo: `Data`, `Categoria` e `Valor` "
            "(negativo para saídas). Também aceita `Valor` + `Tipo (D/C)` ou colunas `Entrada`/`Saída`, "
            "e uma coluna `Empresa` para extratos consolidados.")
        return

    base = pd.concat(frames, ignore_index=True)
    base["mes"] = base["data"].dt.strftime("%Y-%m")

    # ── Sidebar: filtros ─────────────────────────────────────────────────
    with st.sidebar:
        st.divider()
        meses = sorted(base["mes"].unique()) if not base.empty else []
        
        if not meses:
            st.error("Nenhum mês/dado encontrado. Verifique os arquivos enviados.")
            return

        if len(meses) > 1:
            ini, fim = st.select_slider("Período", options=meses, value=(meses[0], meses[-1]),
                                        format_func=rotulo_mes)
        else:
            ini = fim = meses[0]
        detalhe = st.toggle("Detalhar contas na DRE", value=True)
        saldo_ini = st.number_input("Saldo inicial de caixa (R$)", value=0.0, step=1000.0, format="%.2f",
                                    help="Saldo bancário no início do período para a visão selecionada.")
        st.caption("Trânsitos entre empresas GBEA são isolados: não entram no resultado operacional.")

    st.markdown('<div class="gbea-brand"><div class="dia"><span></span><span></span><span></span>'
                '<span></span></div><h1>Dashboard Gerencial · Fluxo de Caixa & DRE</h1></div>',
                unsafe_allow_html=True)
    visao = st.segmented_control("Unidade", [CONSOLIDADO] + UNIDADES, default=CONSOLIDADO,
                                 label_visibility="collapsed") or CONSOLIDADO

    periodo = base[(base["mes"] >= ini) & (base["mes"] <= fim)]
    df = periodo if visao == CONSOLIDADO else periodo[periodo["unidade"] == visao]
    
    if df.empty or not df["mes"].nunique():
        st.warning("⚠️ Nenhum lançamento encontrado para esta combinação de unidade e período. Ajuste os filtros na barra lateral.")
        return

    st.markdown(f'<div class="gbea-sub">{visao} · {rotulo_mes(ini)} a {rotulo_mes(fim)} · '
                f'{len(df):,} lançamentos</div>'.replace(",", "."), unsafe_allow_html=True)

    k = indicadores(df)

    # ── KPIs ─────────────────────────────────────────────────────────────
    c = st.columns(6)
    kpi(c[0], "Receita bruta", brl_curto(k["receita"]), f"líquida {brl_curto(k['receita_liq'])}")
    kpi(c[1], "Margem de contribuição", brl_curto(k["mc"]), f"{pct(k['mc_pct'])} da receita", k["mc"] < 0)
    kpi(c[2], "EBITDA gerencial", brl_curto(k["ebitda"]), f"margem {pct(k['ebitda_pct'])}", k["ebitda"] < 0)
    kpi(c[3], "CMV / faturamento", pct(k["cmv_pct"]), brl_curto(k["cmv"]))
    kpi(c[4], "Folha / faturamento", pct(k["folha_pct"]), brl_curto(k["folha"]))
    kpi(c[5], "Caixa livre s/ trânsitos", brl_curto(k["fcl"]),
        f"trânsito líquido {brl_curto(k['transito'])}", k["fcl"] < 0)
    if k["sem_classif"] or k["pendencias"]:
        st.warning(f"Há {brl_full(abs(k['sem_classif']))} em categorias **sem classificação** e "
                   f"{brl_full(abs(k['pendencias']))} em **pendências de conciliação** – veja a aba Conciliação.")
    st.write("")

    abas = st.tabs(["📊 Visão geral", "📑 DRE gerencial", "💸 Fluxo de caixa × Trânsitos",
                    "🏢 Comparativo de unidades", "🧾 Conciliação", "📚 DE-PARA"])

    meses_sel = sorted(df["mes"].unique())
    rot_meses = [rotulo_mes(m) for m in meses_sel]

    # ── Aba 1: visão geral ───────────────────────────────────────────────
    with abas[0]:
        col1, col2 = st.columns([1.15, 1])
        with col1:
            st.markdown("**Receita bruta × EBITDA gerencial por mês**")
            por_mes = df.groupby(["mes", "bloco"])["valor"].sum().unstack(fill_value=0).reindex(meses_sel)
            rec_m = por_mes["RECEITA"] if "RECEITA" in por_mes.columns else pd.Series(0.0, index=meses_sel)
            ebit_m = pd.Series([indicadores(df[df.mes == m])["ebitda"] for m in meses_sel], index=meses_sel)
            fig = go.Figure()
            fig.add_bar(x=rot_meses, y=rec_m, name="Receita bruta", marker_color=COR["s1"],
                        hovertemplate="R$ %{y:,.0f}")
            fig.add_bar(x=rot_meses, y=ebit_m, name="EBITDA gerencial", marker_color=COR["s2"],
                        hovertemplate="R$ %{y:,.0f}")
            fig.update_layout(barmode="group", bargap=.3, bargroupgap=.08)
            fig.update_traces(marker_line_width=0, marker_cornerradius=4)
            plot(layout(fig))
        with col2:
            st.markdown("**Peso do CMV e da folha sobre o faturamento**")
            ind_m = [indicadores(df[df.mes == m]) for m in meses_sel]
            fig = go.Figure()
            fig.add_scatter(x=rot_meses, y=[i["cmv_pct"] for i in ind_m], name="CMV %", mode="lines+markers",
                            line=dict(width=2, color=COR["s1"]), marker=dict(size=8))
            fig.add_scatter(x=rot_meses, y=[i["folha_pct"] for i in ind_m], name="Folha %", mode="lines+markers",
                            line=dict(width=2, color=COR["s2"]), marker=dict(size=8))
            fig.add_scatter(x=rot_meses, y=[i["ocup_pct"] for i in ind_m], name="Ocupação %",
                            mode="lines+markers", line=dict(width=2, color=COR["s3"]), marker=dict(size=8))
            fig.update_yaxes(tickformat=".0%", rangemode="tozero")
            fig.update_traces(hovertemplate="%{y:.1%}")
            plot(layout(fig))

        col3, col4 = st.columns([1, 1.15])
        with col3:
            st.markdown("**Composição dos custos e despesas operacionais**")
            blocos_custo = ["CMV", "IMPOSTOS"] + [b for b, _ in BLOCOS_DESP_OP]
            comp = (-df[df.bloco.isin(blocos_custo)].groupby("bloco")["valor"].sum()).sort_values()
            comp = comp[comp > 0]
            nomes = [NOME_BLOCO[b].replace("(-) ", "") for b in comp.index]
            fig = go.Figure(go.Bar(
                x=comp.values, y=nomes, orientation="h", marker_color=COR["s1"], marker_cornerradius=4,
                text=[f"{v / k['receita']:.1%}".replace(".", ",") if k["receita"] else "" for v in comp.values],
                textposition="outside", cliponaxis=False,
                customdata=[v / k["receita"] if k["receita"] else 0 for v in comp.values],
                hovertemplate="%{y}<br>R$ %{x:,.0f} · %{customdata:.1%} da receita<extra></extra>"))
            fig.update_layout(hovermode="closest")
            if len(comp):
                fig.update_xaxes(range=[0, comp.max() * 1.15])
            plot(layout(fig, h=380, legenda=False))
        with col4:
            st.markdown("**Evolução do saldo de caixa**")
            diario = df.assign(dia=df["data"].dt.normalize())
            tot = diario.groupby("dia")["valor"].sum()
            op = diario[diario.bloco != "TRANSITO"].groupby("dia")["valor"].sum().reindex(tot.index, fill_value=0)
            fig = go.Figure()
            fig.add_scatter(x=tot.index, y=saldo_ini + tot.cumsum(), name="Saldo real (extrato)",
                            line=dict(width=2, color=COR["s1"]), hovertemplate="R$ %{y:,.0f}")
            fig.add_scatter(x=op.index, y=saldo_ini + op.cumsum(), name="Saldo sem trânsitos GBEA",
                            line=dict(width=2, color=COR["s2"], dash="dot"), hovertemplate="R$ %{y:,.0f}")
            fig.update_xaxes(tickformat="%d/%m")
            plot(layout(fig, h=380))

    # ── Aba 2: DRE ───────────────────────────────────────────────────────
    with abas[1]:
        dre = montar_dre(df, detalhe)
        st.markdown(tabela_dre_html(dre), unsafe_allow_html=True)
        st.caption("Valores em R$ (negativos entre parênteses). AV % = análise vertical sobre a receita bruta. "
                   "Trânsitos entre empresas GBEA ficam fora do resultado e aparecem só na conciliação de caixa.")
        exp = dre.drop(columns="_estilo")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            exp.to_excel(xw, sheet_name="DRE")
            df.drop(columns=["mes"]).to_excel(xw, sheet_name="Lançamentos", index=False)
        st.download_button("⬇️ Exportar DRE + lançamentos (.xlsx)", buf.getvalue(),
                           file_name=f"DRE_{norm(visao).replace(' ', '_')}_{ini}_{fim}.xlsx")

    # ── Aba 3: fluxo de caixa × trânsitos ────────────────────────────────
    with abas[2]:
        col1, col2 = st.columns([1.2, 1])
        with col1:
            st.markdown("**Do resultado operacional à variação de caixa**")
            b = df.groupby("bloco")["valor"].sum()
            curto = {"NAO_OP": "Não operac.", "PENDENCIA": "Pendências", "SEM_CLASSIF": "Sem classif.",
                     "CAPEX": "CAPEX", "EMPRESTIMO": "Empréstimos", "SOCIOS": "Sócios"}
            passos = [("EBITDA", k["ebitda"], "absolute"),
                      ("Financeiro", float(b.get("FINANCEIRO", 0)), "relative")]
            passos += [(curto[x], float(b.get(x, 0)), "relative") for x, _ in BLOCOS_ABAIXO
                       if abs(b.get(x, 0)) > 0.5]
            passos += [("Caixa livre", k["fcl"], "total"),
                       ("Trânsitos", k["transito"], "relative"),
                       ("Var. caixa", k["var_total"], "total")]
            fig = go.Figure(go.Waterfall(
                x=[p[0] for p in passos], y=[p[1] if p[2] != "total" else 0 for p in passos],
                measure=[p[2] for p in passos],
                text=[brl_curto(p[1]).replace("R$ ", "") for p in passos], textposition="outside",
                textfont=dict(size=11), cliponaxis=False,
                increasing=dict(marker=dict(color=COR["good"])),
                decreasing=dict(marker=dict(color=COR["critical"])),
                totals=dict(marker=dict(color=COR["s1"])),
                connector=dict(line=dict(color="rgba(128,128,128,.4)", width=1)),
                hovertemplate="%{x}<br>R$ %{y:,.0f}<extra></extra>"))
            fig.update_layout(hovermode="closest")
            plot(layout(fig, h=420, legenda=False))
            st.caption("Verde = entrada de caixa · vermelho = saída · azul = subtotais.")
        with col2:
            st.markdown("**Resumo do caixa no período**")
            resumo = pd.DataFrame({
                "R$": [k["ebitda"], k["res_fin"] - k["ebitda"], k["fcl"] - k["res_fin"], k["fcl"],
                       k["tr_entradas"], k["tr_saidas"], k["transito"], k["var_total"]]},
                index=["EBITDA gerencial", "Resultado financeiro", "Mov. abaixo do resultado (CAPEX, sócios, etc.)",
                       "= Caixa livre da operação", "Trânsitos recebidos de empresas GBEA",
                       "Trânsitos enviados a empresas GBEA", "= Trânsito líquido", "= Variação total de caixa"])
            st.dataframe(resumo.style.format(lambda v: brl_full(v)), width="stretch")
            if visao == CONSOLIDADO:
                if abs(k["transito"]) < 1:
                    st.success("Trânsitos entre empresas zeram no consolidado ✔")
                else:
                    st.error(f"Trânsitos não zeram no consolidado: sobra {brl_full(k['transito'])}. "
                             "Algum lado da transferência está faltando ou classificado em outra conta.")

        st.markdown("**Trânsitos por unidade** (quem financia e quem é financiado dentro do grupo)")
        trp = periodo[periodo.bloco == "TRANSITO"]
        if trp.empty:
            st.caption("Sem trânsitos no período.")
        else:
            ent = trp[trp.valor > 0].groupby("unidade")["valor"].sum().reindex(UNIDADES, fill_value=0)
            sai = -trp[trp.valor < 0].groupby("unidade")["valor"].sum().reindex(UNIDADES, fill_value=0)
            fig = go.Figure()
            fig.add_bar(x=UNIDADES, y=ent, name="Recebido de outras empresas", marker_color=COR["s1"],
                        hovertemplate="R$ %{y:,.0f}")
            fig.add_bar(x=UNIDADES, y=sai, name="Enviado a outras empresas", marker_color=COR["s2"],
                        hovertemplate="R$ %{y:,.0f}")
            fig.update_layout(barmode="group", bargap=.35, bargroupgap=.08)
            fig.update_traces(marker_cornerradius=4)
            plot(layout(fig, h=330))
            with st.expander("Lançamentos de trânsito"):
                t = (trp if visao == CONSOLIDADO else trp[trp.unidade == visao])
                st.dataframe(t[["data", "unidade", "categoria", "descricao", "valor"]].sort_values("data"),
                             width="stretch", hide_index=True,
                             column_config={"valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                                            "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY")})

    # ── Aba 4: comparativo ───────────────────────────────────────────────
    with abas[3]:
        linhas = []
        for u in UNIDADES:
            du = periodo[periodo.unidade == u]
            if du.empty:
                continue
            i = indicadores(du)
            linhas.append({"Unidade": u, "Receita bruta": i["receita"], "CMV %": i["cmv_pct"],
                           "Folha %": i["folha_pct"], "Ocupação %": i["ocup_pct"],
                           "Margem contrib. %": i["mc_pct"], "EBITDA": i["ebitda"], "Margem EBITDA": i["ebitda_pct"],
                           "Caixa livre": i["fcl"], "Trânsito líquido": i["transito"]})
        comp = pd.DataFrame(linhas)
        if not comp.empty:
            g = indicadores(periodo)
            comp.loc[len(comp)] = {"Unidade": CONSOLIDADO, "Receita bruta": g["receita"], "CMV %": g["cmv_pct"],
                                   "Folha %": g["folha_pct"], "Ocupação %": g["ocup_pct"],
                                   "Margem contrib. %": g["mc_pct"], "EBITDA": g["ebitda"],
                                   "Margem EBITDA": g["ebitda_pct"], "Caixa livre": g["fcl"],
                                   "Trânsito líquido": g["transito"]}
            col1, col2 = st.columns([1, 1])
            with col1:
                st.markdown("**Margem EBITDA por unidade**")
                cu = comp[comp.Unidade != CONSOLIDADO].sort_values("Margem EBITDA")
                fig = go.Figure(go.Bar(
                    x=cu["Margem EBITDA"], y=cu["Unidade"], orientation="h", marker_cornerradius=4,
                    marker_color=[COR["s1"] if (visao in (CONSOLIDADO, u)) else "rgba(128,128,128,.35)"
                                  for u in cu["Unidade"]],
                    text=[pct(v) for v in cu["Margem EBITDA"]], textposition="outside", cliponaxis=False,
                    hovertemplate="%{y}: %{x:.1%}<extra></extra>"))
                fig.add_vline(x=g["ebitda_pct"], line_dash="dot", line_color="rgba(128,128,128,.7)",
                              annotation_text=f"Grupo {pct(g['ebitda_pct'])}", annotation_position="top")
                fig.update_xaxes(tickformat=".0%", range=[min(0, cu["Margem EBITDA"].min() * 1.45),
                                                          cu["Margem EBITDA"].max() * 1.18])
                fig.update_layout(hovermode="closest")
                plot(layout(fig, h=380, legenda=False))
            with col2:
                st.markdown("**Receita bruta × EBITDA por unidade**")
                cu2 = comp[comp.Unidade != CONSOLIDADO]
                fig = go.Figure()
                fig.add_bar(x=cu2["Unidade"], y=cu2["Receita bruta"], name="Receita bruta", marker_color=COR["s1"],
                            hovertemplate="R$ %{y:,.0f}")
                fig.add_bar(x=cu2["Unidade"], y=cu2["EBITDA"], name="EBITDA gerencial", marker_color=COR["s2"],
                            hovertemplate="R$ %{y:,.0f}")
                fig.update_layout(barmode="group", bargap=.3, bargroupgap=.08)
                fig.update_traces(marker_cornerradius=4)
                plot(layout(fig, h=380))
            if True:
                st.markdown("**Indicadores por unidade**")
                money = st.column_config.NumberColumn(format="localized")
                perc = st.column_config.NumberColumn(format="percent")
                st.dataframe(comp, hide_index=True, width="stretch",
                             column_config={"Receita bruta": money, "EBITDA": money, "Caixa livre": money,
                                            "Trânsito líquido": money, "CMV %": perc, "Folha %": perc,
                                            "Ocupação %": perc, "Margem contrib. %": perc,
                                            "Margem EBITDA": perc})

    # ── Aba 5: conciliação ───────────────────────────────────────────────
    with abas[4]:
        sc = periodo[periodo.bloco == "SEM_CLASSIF"]
        st.markdown("**Categorias fora do DE-PARA**")
        if sc.empty:
            st.success("Todas as categorias do período foram reconhecidas pelo DE-PARA ✔")
        else:
            st.caption("Escolha a conta gerencial e clique em aplicar. A regra vale para esta sessão; "
                       "para torná-la permanente, inclua o termo na lista `REGRAS` do código.")
            resumo = sc.groupby("categoria").agg(lancamentos=("valor", "size"), total=("valor", "sum")).reset_index()
            resumo["Classificar como"] = None
            with st.form("form_map"):
                ed = st.data_editor(
                    resumo, hide_index=True, width="stretch", disabled=["categoria", "lancamentos", "total"],
                    column_config={"total": st.column_config.NumberColumn("Total", format="R$ %.2f"),
                                   "Classificar como": st.column_config.SelectboxColumn(options=GRUPOS_OPCOES)})
                if st.form_submit_button("Aplicar classificação"):
                    for _, r in ed.dropna(subset=["Classificar como"]).iterrows():
                        st.session_state["extra_map"][norm(r["categoria"])] = r["Classificar como"]
                    st.cache_data.clear()
                    st.rerun()
        if st.session_state["extra_map"]:
            with st.expander("Classificações manuais desta sessão"):
                st.json(st.session_state["extra_map"])
                if st.button("Limpar classificações manuais"):
                    st.session_state["extra_map"] = {}
                    st.cache_data.clear()
                    st.rerun()

        st.markdown("**Pendências de conciliação** (Despesa não identificada)")
        pend = periodo[periodo.bloco == "PENDENCIA"]
        if pend.empty:
            st.success("Sem pendências no período ✔")
        else:
            st.dataframe(pend[["data", "unidade", "categoria", "descricao", "valor"]].sort_values("data"),
                         hide_index=True, width="stretch",
                         column_config={"valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                                        "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY")})
        st.markdown("**Todos os lançamentos classificados**")
        st.dataframe(df[["data", "unidade", "categoria", "grupo", "descricao", "valor", "origem"]].sort_values("data"),
                     hide_index=True, width="stretch", height=320,
                     column_config={"valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY")})

    # ── Aba 6: regras ────────────────────────────────────────────────────
    with abas[5]:
        st.markdown("**DE-PARA oficial GBEA** – termos do extrato (sem acento/maiúscula) → conta gerencial")
        regras = pd.DataFrame([{"Categoria no extrato": ", ".join(t), "Conta gerencial": g,
                                "Bloco": NOME_BLOCO.get(b, b),
                                "Entra no EBITDA": "Sim" if b in {"RECEITA", "IMPOSTOS", "CMV"} | {x for x, _ in BLOCOS_DESP_OP}
                                else "Não"} for t, g, b, _ in REGRAS])
        st.dataframe(regras, hide_index=True, width="stretch", height=620)


if __name__ == "__main__":
    main()
