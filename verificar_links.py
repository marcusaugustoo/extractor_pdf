#!/usr/bin/env python3
"""
verificar_links.py

Confere se a URL de cada notícia (coluna `url`, gerada pelo extrair_links.py)
realmente corresponde à notícia do PDF (coluna `titulo`).

Não altera o extrair_links.py: lê a planilha que ele gerou e cria uma nova,
com colunas extras de verificação.

Como funciona, para cada linha:
  1. Abre a URL e lê o título da página (og:title, <title> e <h1>).
  2. Compara com o título do PDF (similaridade de 0 a 1).
  3. Compara também com o "slug" da própria URL (o trecho final do endereço,
     que quase sempre é o título da matéria).
  4. Cruza a data da planilha com a data que aparece na URL (/2025/01/15/)
     ou nos metadados da página.
  5. Marca URLs repetidas e links que levam só à página inicial do site.

Resultado em `verificacao`:
  confere              título bate e nada suspeito
  revisar              parece certo, mas há dúvida (título parcial, data
                       diferente, URL repetida)
  não confere          o título da página/URL é de outra notícia
  sem link da matéria  a URL é só a home do site (ou está vazia)
  não abriu            a página não carregou (HTTP 404, timeout etc.)

Uso:
  python verificar_links.py janeiro.xlsx
  python verificar_links.py janeiro.xlsx janeiro_verificado.xlsx
  python verificar_links.py janeiro.xlsx --sem-rede     # só compara com o slug

Dependências:
  pip install pandas openpyxl requests beautifulsoup4
"""

import argparse
import re
import sys
import time
import unicodedata
from collections import Counter
from difflib import SequenceMatcher
from urllib.parse import urlparse, unquote

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:  # só é obrigatório quando se usa a rede
    requests = None
    BeautifulSoup = None

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; verificar_links/1.0; projeto academico)"
}

STOPWORDS = {
    "uma", "que", "sem", "com", "para", "por", "nas", "nos", "das", "dos",
    "aos", "como", "mais", "foi", "são", "sao", "ser", "seu", "sua", "diz",
}

CORES = {
    "confere": "C6EFCE",
    "revisar": "FFEB9C",
    "não confere": "FFC7CE",
    "sem link da matéria": "FFC7CE",
    "não abriu": "D9D9D9",
}


# --------------------------------------------------------------------------
# Texto e similaridade
# --------------------------------------------------------------------------
def normalizar(txt):
    txt = unicodedata.normalize("NFKD", str(txt or ""))
    txt = txt.encode("ascii", "ignore").decode().lower()
    txt = re.sub(r"[^a-z0-9\s]", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


def tokens(txt):
    return [t for t in normalizar(txt).split() if len(t) > 2 and t not in STOPWORDS]


def similaridade(a, b):
    """0 a 1. Usa o maior entre: sequência de caracteres, Jaccard de palavras e
    'contenção' (quanto do título menor está dentro do maior). A contenção
    ajuda quando o OCR do PDF traz texto a mais ou duplicado."""
    ta, tb = set(tokens(a)), set(tokens(b))
    if not ta or not tb:
        return 0.0
    inter = len(ta & tb)
    jaccard = inter / len(ta | tb)
    contencao = inter / min(len(ta), len(tb)) if min(len(ta), len(tb)) >= 4 else 0.0
    seq = SequenceMatcher(None, normalizar(a), normalizar(b)).ratio()
    return round(max(seq, jaccard, 0.9 * contencao), 2)


def limpar_titulo_pagina(t):
    """Tira o nome do site do título ('Título | Brasil de Fato')."""
    partes = re.split(r"\s[\|\-–—]\s", t)
    return max(partes, key=len).strip() if partes else t.strip()


def slug_da_url(url):
    caminho = unquote(urlparse(url).path).strip("/")
    if not caminho:
        return ""
    ultimo = caminho.split("/")[-1]
    ultimo = re.sub(r"\.(s?html?|php|aspx?)$", "", ultimo, flags=re.I)
    ultimo = re.sub(r"^\d+-", "", ultimo)  # '7033-familias-sem-terra...'
    ultimo = re.sub(r"[-_]+", " ", ultimo)
    return "" if ultimo.replace(" ", "").isdigit() else ultimo.strip()


def eh_home(url):
    p = urlparse(url)
    return not p.path.strip("/") and not p.query


# --------------------------------------------------------------------------
# Datas
# --------------------------------------------------------------------------
def data_da_url(url):
    """Retorna (Timestamp, 'dia') ou (Timestamp, 'mes') ou (None, None)."""
    m = re.search(r"/(20\d\d)/(\d{1,2})/(\d{1,2})(?:/|$)", urlparse(url).path)
    if m:
        try:
            return pd.Timestamp(int(m[1]), int(m[2]), int(m[3])), "dia"
        except ValueError:
            pass
    m = re.search(r"/(20\d\d)/(\d{1,2})(?:/|$)", urlparse(url).path)
    if m:
        try:
            return pd.Timestamp(int(m[1]), int(m[2]), 1), "mes"
        except ValueError:
            pass
    return None, None


def datas_divergem(data_planilha, data_ref, precisao, tolerancia_dias):
    if data_planilha is None or data_ref is None:
        return False
    if precisao == "mes":
        return (data_planilha.year, data_planilha.month) != (data_ref.year, data_ref.month)
    return abs((data_planilha - data_ref).days) > tolerancia_dias


# --------------------------------------------------------------------------
# Rede
# --------------------------------------------------------------------------
def abrir(url, sessao, timeout, ultimos_acessos, pausa, tentativas=2):
    """Retorna (conteudo_bytes, url_final, erro)."""
    host = urlparse(url).netloc
    erro = None
    for i in range(tentativas):
        espera = pausa - (time.time() - ultimos_acessos.get(host, 0))
        if espera > 0:
            time.sleep(espera)
        ultimos_acessos[host] = time.time()
        try:
            r = sessao.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
            if r.status_code == 200:
                return r.content, r.url, None
            erro = f"HTTP {r.status_code}"
            if r.status_code in (401, 403, 404, 410):
                break
        except Exception as e:  # inclui erros do urllib3 que o requests não encapsula
            nome = type(e).__name__
            if nome == "LocationParseError":
                erro = "redirecionamento malformado do site (LocationParseError)"
                break  # tentar de novo não adianta
            erro = nome
        time.sleep(1.5 * (i + 1))
    return None, None, erro


def extrair_pagina(conteudo):
    """Retorna (lista_de_titulos_candidatos, data_publicacao_str)."""
    soup = BeautifulSoup(conteudo, "html.parser")

    def meta(**attrs):
        tag = soup.find("meta", attrs=attrs)
        return (tag.get("content") or "").strip() if tag else ""

    candidatos = []
    for t in (
        meta(property="og:title"),
        soup.title.get_text(" ", strip=True) if soup.title else "",
        soup.find("h1").get_text(" ", strip=True) if soup.find("h1") else "",
    ):
        t = limpar_titulo_pagina(t) if t else ""
        if t and t not in candidatos:
            candidatos.append(t)

    data = (
        meta(property="article:published_time")
        or meta(name="date")
        or meta(itemprop="datePublished")
    )
    if not data:
        tag = soup.find("time", attrs={"datetime": True})
        data = tag["datetime"] if tag else ""
    return candidatos, data


# --------------------------------------------------------------------------
# Verificação de uma linha
# --------------------------------------------------------------------------
def verificar_linha(row, cfg, sessao, cache, ultimos_acessos, urls_repetidas):
    url = str(row.get("url") or "").strip()
    titulo_pdf = str(row.get("titulo") or "")
    res = {
        "verificacao": "",
        "similaridade_titulo": None,
        "comparado_com": "",
        "titulo_comparado": "",
        "data_url_ou_pagina": "",
        "motivo": "",
    }

    if not url.startswith("http"):
        res.update(verificacao="sem link da matéria", motivo="URL vazia ou inválida")
        return res
    if eh_home(url):
        res.update(
            verificacao="sem link da matéria",
            motivo="a URL é só a página inicial do site, não a matéria",
        )
        return res

    motivos = []
    slug = slug_da_url(url)
    melhor = (similaridade(titulo_pdf, slug) if slug else 0.0, slug, "slug da URL")
    data_pagina_str = ""

    if not cfg.sem_rede:
        if url not in cache:
            cache[url] = abrir(url, sessao, cfg.timeout, ultimos_acessos, cfg.pausa)
        conteudo, url_final, erro = cache[url]
        if conteudo is None:
            res.update(
                verificacao="não abriu",
                similaridade_titulo=melhor[0],
                comparado_com="slug da URL",
                titulo_comparado=slug,
                motivo=f"{erro}; similaridade só com o slug da URL: {melhor[0]}",
            )
            return res
        if eh_home(url_final):
            res.update(
                verificacao="não confere",
                motivo="a URL redireciona para a página inicial (matéria removida?)",
            )
            return res
        candidatos, data_pagina_str = extrair_pagina(conteudo)
        for c in candidatos:
            s = similaridade(titulo_pdf, c)
            if s > melhor[0]:
                melhor = (s, c, "página")
    else:
        motivos.append("página não aberta (--sem-rede): comparado só com o slug da URL")

    score, comparado, fonte = melhor
    if score >= cfg.limiar_ok:
        status = "confere"
    elif score >= cfg.limiar_revisar:
        status = "revisar"
        motivos.append("título parecido, mas só em parte")
    else:
        status = "não confere"
        motivos.append("o título da URL/página parece ser de outra notícia")

    # Data
    data_planilha = pd.to_datetime(row.get("data_publicacao"), dayfirst=True, errors="coerce")
    data_planilha = None if pd.isna(data_planilha) else data_planilha
    data_ref, precisao = data_da_url(url)
    origem_data = "URL"
    if data_ref is None and data_pagina_str:
        d = pd.to_datetime(data_pagina_str[:10], errors="coerce")
        if not pd.isna(d):
            data_ref, precisao, origem_data = d, "dia", "página"
    if data_ref is not None:
        res["data_url_ou_pagina"] = f"{data_ref:%d/%m/%Y}" if precisao == "dia" else f"{data_ref:%m/%Y}"
    if datas_divergem(data_planilha, data_ref, precisao, cfg.tolerancia_dias):
        motivos.append(
            f"data na {origem_data} ({res['data_url_ou_pagina']}) difere da data da "
            f"planilha ({data_planilha:%d/%m/%Y})"
        )
        if status == "confere":
            status = "revisar"

    # URL repetida
    if urls_repetidas.get(url, 0) > 1:
        motivos.append(f"a mesma URL aparece em {urls_repetidas[url]} linhas")
        if status == "confere":
            status = "revisar"

    res.update(
        verificacao=status,
        similaridade_titulo=score,
        comparado_com=fonte,
        titulo_comparado=comparado,
        motivo="; ".join(motivos),
    )
    return res


# --------------------------------------------------------------------------
# Planilha
# --------------------------------------------------------------------------
def formatar(ws, df):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True)
    col_status = list(df.columns).index("verificacao") + 1
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_status)
        cor = CORES.get(cell.value)
        if cor:
            cell.fill = PatternFill("solid", fgColor=cor)
    larguras = {"titulo": 60, "titulo_comparado": 60, "url": 55, "motivo": 70}
    for i, nome in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = larguras.get(nome, 18)


def main():
    ap = argparse.ArgumentParser(description="Verifica se as URLs das notícias batem com os PDFs.")
    ap.add_argument("entrada", help="planilha .xlsx gerada pelo extrair_links.py")
    ap.add_argument("saida", nargs="?", help="planilha de saída (padrão: <entrada>_verificado.xlsx)")
    ap.add_argument("--sem-rede", action="store_true", help="não abre as páginas; compara só com o slug da URL")
    ap.add_argument("--timeout", type=float, default=15, help="timeout por página, em segundos")
    ap.add_argument("--pausa", type=float, default=1.0, help="pausa mínima entre acessos ao mesmo site")
    ap.add_argument("--limiar-ok", type=float, default=0.60, help="similaridade mínima para 'confere'")
    ap.add_argument("--limiar-revisar", type=float, default=0.35, help="similaridade mínima para 'revisar'")
    ap.add_argument("--tolerancia-dias", type=int, default=1, help="diferença de dias aceita entre a data da planilha e a da URL")
    cfg = ap.parse_args()

    if not cfg.sem_rede and (requests is None or BeautifulSoup is None):
        sys.exit("Instale as dependências: pip install requests beautifulsoup4 (ou use --sem-rede)")

    saida = cfg.saida or re.sub(r"\.xlsx$", "", cfg.entrada, flags=re.I) + "_verificado.xlsx"

    xls = pd.ExcelFile(cfg.entrada)
    aba = "Notícias" if "Notícias" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(cfg.entrada, sheet_name=aba)
    for col in ("titulo", "url"):
        if col not in df.columns:
            sys.exit(f"A aba '{aba}' não tem a coluna '{col}'.")

    urls_repetidas = Counter(
        u for u in df["url"].dropna().astype(str) if u.startswith("http") and not eh_home(u)
    )

    sessao = requests.Session() if not cfg.sem_rede else None
    cache, ultimos_acessos, resultados = {}, {}, []
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        print(f"[{i}/{len(df)}] {row.get('codigo', '')}", flush=True)
        try:
            resultados.append(verificar_linha(row, cfg, sessao, cache, ultimos_acessos, urls_repetidas))
        except Exception as e:  # uma linha ruim não derruba as outras
            resultados.append({
                "verificacao": "não abriu",
                "similaridade_titulo": None,
                "comparado_com": "",
                "titulo_comparado": "",
                "data_url_ou_pagina": "",
                "motivo": f"erro inesperado: {type(e).__name__}: {e}",
            })

    df_out = pd.concat([df.reset_index(drop=True), pd.DataFrame(resultados)], axis=1)

    resumo = df_out["verificacao"].value_counts().rename_axis("verificacao").reset_index(name="quantidade")
    por_origem = None
    if "url_origem" in df_out.columns:
        origem = df_out["url_origem"].astype(str).str.replace(r"\s*\(.*\)$", "", regex=True)
        por_origem = pd.crosstab(origem, df_out["verificacao"]).reset_index().rename(columns={"url_origem": "url_origem"})

    with pd.ExcelWriter(saida, engine="openpyxl") as w:
        df_out.to_excel(w, sheet_name="Notícias", index=False)
        resumo.to_excel(w, sheet_name="Resumo_verificacao", index=False)
        if por_origem is not None:
            por_origem.to_excel(w, sheet_name="Por_url_origem", index=False)
        formatar(w.sheets["Notícias"], df_out)
        for nome in ("Resumo_verificacao", "Por_url_origem"):
            if nome in w.sheets:
                for cell in w.sheets[nome][1]:
                    cell.font = Font(bold=True)
                for col in w.sheets[nome].columns:
                    w.sheets[nome].column_dimensions[col[0].column_letter].width = 28

    print("\nResumo:")
    print(resumo.to_string(index=False))
    print(f"\nPlanilha gerada: {saida}")


if __name__ == "__main__":
    main()
