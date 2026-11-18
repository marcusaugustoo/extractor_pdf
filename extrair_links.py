"""
Extrai metadados de PDFs de notícias e gera uma planilha Excel.

Instalação:
    pip install pymupdf pandas openpyxl pytesseract pillow
    + Tesseract OCR (com idioma "Portuguese") para leitura de PDFs em imagem.

Uso:
    python extrair_links.py PASTA_COM_PDFS noticias.xlsx
    python extrair_links.py UM_ARQUIVO.pdf teste.xlsx
    python extrair_links.py PASTA noticias.xlsx --debug-ocr dbg
    python extrair_links.py PASTA noticias.xlsx --modelo-slm qwen2.5:3b
"""
import argparse
import difflib
import json
import os
import re
import shutil
import sys
import unicodedata
import urllib.error
import urllib.request
from collections import Counter
from datetime import date
from pathlib import Path
from urllib.parse import unquote, urlparse

import pandas as pd
import pymupdf

NOMES_FONTE = {
    "agenciabrasil.ebc.com.br": "Agência Brasil",
    "cienciadoleite.com.br": "Ciência do Leite",
    "folha.uol.com.br": "Folha de S.Paulo",
    "poder360.com.br": "Poder360",
    "soundcloud.com": "SoundCloud",
    "ifpa.edu.br": "IFPA",
    "anovademocracia.com.br": "A Nova Democracia",
    "mab.org.br": "MAB – Movimento dos Atingidos por Barragens",
    "correiodecarajas.com.br": "Correio de Carajás",
    "brasildefato.com.br": "Brasil de Fato",
    "cptnacional.org.br": "CPT – Comissão Pastoral da Terra",
    "metropoles.com": "Metrópoles",
    "fup.org.br": "FUP – Federação Única dos Petroleiros",
    "seculodiario.com.br": "Século Diário",
    "saibamais.jor.br": "Saiba Mais",
    "dw.com": "Deutsche Welle",
    "g1.globo.com": "G1",
    "apublica.org": "Agência Pública",
    "saulovale.com.br": "Saulo Vale",
    "vgnoticias.com.br": "VGN",
    "campograndenews.com.br": "Campo Grande News",
    "amazoniabrasilradioweb.com.br": "Amazônia Brasil Rádio Web",
    "sorriso.mt.gov.br": "Prefeitura de Sorriso",
    "mpf.mp.br": "MPF – Ministério Público Federal",
    "didigalvao.com.br": "Blog do Didi Galvão",
    "progressive.international": "Progressive International",
    "noticiapreta.com.br": "Notícia Preta",
    "racismoambiental.net.br": "Combate Racismo Ambiental",
    "causaoperaria.org.br": "Diário Causa Operária",
    "reporterbrasil.org.br": "Repórter Brasil",
}

MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5,
    "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10,
    "novembro": 11, "dezembro": 12,
}

IGNORAR_PASTAS = {"__macosx", "$recycle.bin", "system volume information", "node_modules"}
NIVEIS_PASTA = 3

OCR_DPI = 300
OCR_DPI_MAX_IMAGEM = 450
OCR_MAX_PAGINAS = 8
OCR_MAX_REGIOES = 30
OCR_DPI_MOLDURA = 400
MOLDURA_PT = 36
FAIXA_URL_PT = 60
OCR_VARIANTES = (("cinza", "--psm 6"), ("binaria", "--psm 6"), ("cinza", "--psm 11"))
CONF_BOA = 85

DEBUG = {"dir": None, "pdf": ""}
SLM = {
    "modelo": "",
    "url": os.environ.get("OLLAMA_URL", "http://localhost:11434/api/generate"),
    "timeout": 20,
}

RE_PUB_NUM = re.compile(
    r"(?:Publicad[oa]|Postad[oa])(?:\s+em)?:?\s*(\d{1,2})/(\d{1,2})/(\d{2,4})"
    r"(?:\s*(?:-|–|,|às|as)?\s*(\d{1,2})[:h](\d{2}))?",
    re.IGNORECASE,
)
RE_PUB_EXT = re.compile(
    r"(?:Publicad[oa]|Postad[oa])(?:\s+em)?:?\s*(\d{1,2})\s+de\s+([a-zç]+)\s+de\s+(\d{4})"
    r"(?:\s*(?:-|–|,|às|as)?\s*(\d{1,2})[:h](\d{2}))?",
    re.IGNORECASE,
)
RE_IMPRESSAO = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4}),?\s+(\d{1,2}):(\d{2})")
RE_URL_ANO_MES = re.compile(r"/(20\d{2})[-/](\d{2})(?:[-/](\d{2}))?(?:/|$)")
RE_NOME_ARQUIVO = re.compile(r"^(\d{2})(\d{2})(\d{2})[A-Za-z]")  # ddmmaa + letras

_TLDS = r"(?:com|org|net|gov|edu|mp|leg|jus|jor|ind|agr|inf|art|mil|coop|imb|b|tv|fm|news|info|gg|me|io|international)"
RE_URL_OCR = re.compile(
    rf"^(?:https?:/{{1,2}}\S+|www\.\S+"
    rf"|[a-z0-9-]+(?:\.[a-z0-9-]+)*\.(?:{_TLDS}(?:\.br)?|br)(?:/\S*)?)$",
    re.IGNORECASE,
)
RE_PARECE_LINK = re.compile(rf":/|/\bwww\.|\.(?:{_TLDS}|br)\b", re.IGNORECASE)
RE_INICIO_URL = re.compile(r"^(?:https?:/{1,2}|www\.)", re.IGNORECASE)
RE_RESTO_URL = re.compile(r"^[A-Za-z0-9\-._~%/?=&#+:@]{3,}$")

_DOMINIOS_IGNORAR = (
    "doubleclick.net", "googleadservices.com", "adssettings.google.com", "google.com",
    "googleapis.com", "gstatic.com", "googleusercontent.com", "googlesyndication.com",
    "googletagmanager.com", "google-analytics.com", "g.co", "catarse.me", "livepix.gg",
    "paypal.com", "whatsapp.com", "t.me", "twitter.com", "x.com", "facebook.com",
    "fb.com", "youtube.com", "youtu.be", "linkedin.com", "bsky.app", "pinterest.com",
    "tiktok.com", "instagram.com", "creativecommons.org", "lojadoand.com.br",
    "mobister.com.br", "ajor.org.br",
)

_PADROES_URL_INVALIDA = re.compile(
    r"(?:^https?:/{1,2}(?:accounts|mail|drive|docs)\.google\."
    r"|/(?:signin|login|logout|auth|oauth|sharer|share|intent/tweet|wp-login|wp-admin|feed|rss)(?:\.php\b|/|\?|$|#))",
    re.IGNORECASE,
)

_IGNORAR_CAMINHOS_PADRAO = re.compile(
    r"^/(?:sobre|quem-somos|contato|expediente|assine|apoie|doe|contribua|"
    r"categoria|categorias|category|tag|tags|editoria|editorias|autor|autores|author|"
    r"colunista|colunistas|topico|tudo-sobre|ultimas-noticias|noticias-do-rn-e-nordeste|"
    r"radioagencia|programas|especial|tipo|editorial|organizacoes-|personalidades-|"
    r"brasil/[a-z-]+/?$|busca|search|ouvidoria/?$|politica-de-privacidade|politica-privacidade|"
    r"transparencia|denuncie-e-seja-uma-fonte|republique|"
    r"(?:ba|ce|df|mg|pa|pb|pe|pr|rj|rn|rs|sp)/?$)(?:/|\?|$)",
    re.IGNORECASE,
)


def _perfil(padrao: str, reconstruir_slug: bool = True, ignorar: re.Pattern = _IGNORAR_CAMINHOS_PADRAO) -> dict:
    return {
        "padrao_artigo": re.compile(padrao, re.IGNORECASE),
        "reconstruir_slug": reconstruir_slug,
        "ignorar_caminhos": ignorar,
    }


PERFIS_FONTES = {
    "soundcloud.com": _perfil(
        r"^/(?!discover|stream|signin|login|upload|search|settings|you|pages|charts)[a-z0-9_-]+/[a-z0-9_-]+",
        reconstruir_slug=False,
    ),
    "ifpa.edu.br": _perfil(r"/(?:noticias|comunicacao|ultimas-noticias)/.+|/[a-z0-9-]{20,}"),
    "anovademocracia.com.br": _perfil(r"/(?:20\d{2}/\d{2}/|[a-z0-9-]{15,}/?$)"),
    "mab.org.br": _perfil(r"/(?:noticias/|20\d{2}/\d{2}/|[a-z0-9-]{15,}/?$)"),
    "correiodecarajas.com.br": _perfil(r"/(?:[a-z0-9-]+/)?[a-z0-9-]{15,}/?$"),
    "brasildefato.com.br": _perfil(r"/(?:20\d{2}/\d{2}/\d{2}/|podcast/.+/\d{4}/\d{2}/\d{2}/)[a-z0-9-]{10,}"),
    "cptnacional.org.br": _perfil(r"/(?:publicacoes|noticias|mural|pastoral)/.+|/\d+-[a-z0-9-]{10,}"),
    "metropoles.com": _perfil(r"/(?:20\d{2}/\d{2}/|[a-z0-9-]+/[a-z0-9-]{15,})"),
    "fup.org.br": _perfil(r"/[a-z0-9-]{15,}/?$"),
    "seculodiario.com.br": _perfil(r"/(?:[a-z0-9-]+/)?[a-z0-9-]{15,}/?$"),
    "saibamais.jor.br": _perfil(r"/(?:20\d{2}/\d{2}/|[a-z0-9-]{15,}/?$)"),
    "dw.com": _perfil(r"/pt-br/.+/a-\d+", reconstruir_slug=False),
    "g1.globo.com": _perfil(r"/(?:noticia|post)/20\d{2}/\d{2}/\d{2}/.+\.ghtml", reconstruir_slug=False),
    "apublica.org": _perfil(r"/20\d{2}/\d{2}/[a-z0-9-]{10,}"),
    "saulovale.com.br": _perfil(r"/(?:20\d{2}/\d{2}/(?:\d{2}/)?[a-z0-9-]{10,}|[a-z0-9-]{15,}/?$)"),
    "vgnoticias.com.br": _perfil(r"/(?:[a-z0-9-]+/)?[a-z0-9-]{15,}"),
    "campograndenews.com.br": _perfil(r"/(?:noticia|[a-z0-9-]+)/[a-z0-9-]{15,}"),
    "amazoniabrasilradioweb.com.br": _perfil(r"/(?:noticias/|[a-z0-9-]{15,}/?$)"),
    "sorriso.mt.gov.br": _perfil(r"/(?:noticia|noticias|comunicacao|imprensa)/.+|/[a-z0-9-]{20,}"),
    "mpf.mp.br": _perfil(r"/(?:noticias|sala-de-imprensa|comunicacao)/.+|/[a-z0-9-]{20,}"),
    "didigalvao.com.br": _perfil(r"/(?:20\d{2}/\d{2}/|[a-z0-9-]{15,}/?$)"),
    "progressive.international": _perfil(r"/(?:wire|briefing|article)/.+"),
    "noticiapreta.com.br": _perfil(r"/(?:20\d{2}/\d{2}/|[a-z0-9-]{15,}/?$)"),
    "racismoambiental.net.br": _perfil(r"/20\d{2}/\d{2}/(?:\d{2}/)?[a-z0-9-]{10,}"),
    "causaoperaria.org.br": _perfil(r"/20\d{2}/[a-z0-9-]{10,}"),
    "reporterbrasil.org.br": _perfil(r"/20\d{2}/\d{2}/[a-z0-9-]{10,}"),
}


def obter_perfil_fonte(dom: str) -> dict | None:
    if not dom:
        return None
    for d, p in PERFIS_FONTES.items():
        if dom == d or dom.endswith("." + d):
            return p
    return None


def nome_fonte_dominio(dom: str) -> str:
    if not dom:
        return ""
    return next((n for d, n in NOMES_FONTE.items() if dom == d or dom.endswith("." + d)), "")


def natural(s: str):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", sem_acento(s).lower())]


def niveis_pasta(rel: str) -> dict:
    partes = Path(rel).parts if rel else ()
    n = NIVEIS_PASTA
    vals = list(partes[:n - 1]) + ([" / ".join(partes[n - 1:])] if len(partes) >= n else [])
    return {f"pasta_{i + 1}": (vals[i] if i < len(vals) else "") for i in range(n)}


def pasta_relativa(pdf_path: Path, raiz: Path) -> str:
    rel = pdf_path.parent.relative_to(raiz)
    return "" if str(rel) == "." else str(rel)


def linha_base(pdf_path: Path, raiz: Path) -> dict:
    pasta = pasta_relativa(pdf_path, raiz)
    return {"codigo": pdf_path.stem, "pasta": pasta, **niveis_pasta(pasta)}


RE_CONTROLE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def norm(s: str) -> str:
    return " ".join(RE_CONTROLE.sub("", s).split())


# "Microsoft Print to PDF" grava apenas o byte baixo de alguns caracteres U+20xx nos metadados
CONSERTO_META = str.maketrans({
    "\x13": "–", "\x14": "—", "\x18": "‘", "\x19": "’", "\x1c": "“", "\x1d": "”",
    "˘": "‘", "ˇ": "’", "˝": "“", "˛": "”",
})


def titulo_metadados(doc) -> str:
    return norm(((doc.metadata or {}).get("title") or "").translate(CONSERTO_META))


def sem_acento(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def mk_date(d, m, a):
    a = int(a)
    if a < 100:
        a += 2000
    try:
        return date(a, int(m), int(d))
    except ValueError:
        return None


def chave_pasta(parte: str):
    s = sem_acento(parte).lower().strip()
    for nome, i in MESES.items():
        if re.search(rf"\b{sem_acento(nome)}\b", s):
            return (0, i, natural(s))
    n = re.match(r"\s*(\d{1,2})\b", s)
    if n:
        return (0, int(n[1]), natural(s))
    if re.fullmatch(r"20\d{2}", s):
        return (0, int(s), natural(s))
    return (1, 99, natural(s))


def chave_pdf(p: Path, raiz: Path):
    partes = p.parent.relative_to(raiz).parts
    mn = RE_NOME_ARQUIVO.match(p.name)
    cron = (int(mn[3]), int(mn[2]), int(mn[1])) if mn else (99, 99, 99)
    return (tuple(chave_pasta(x) for x in partes), cron, natural(p.name))


def listar_pdfs(raiz: Path):
    achados, erros = [], []
    for dirpath, dirnames, filenames in os.walk(raiz, onerror=erros.append):
        dirnames[:] = [d for d in dirnames
                       if not d.startswith((".", "~$")) and d.lower() not in IGNORAR_PASTAS]
        for f in filenames:
            if f.lower().endswith(".pdf") and not f.startswith(("._", "~$")):
                achados.append(Path(dirpath) / f)
    for e in erros:
        print(f"AVISO: não consegui abrir a pasta {getattr(e, 'filename', '?')}: {e}", file=sys.stderr)
    return sorted(achados, key=lambda p: chave_pdf(p, raiz))


def extrair_dominio(url: str) -> str:
    if not url:
        return ""
    u = re.sub(r"^(https?):/+", r"\1://", url.strip())
    if "://" not in u:
        u = "https://" + u
    return urlparse(u).netloc.lower().removeprefix("www.")


def url_valida(u: str, permitir_instagram: bool = False) -> bool:
    if not u:
        return False
    limpa = limpar_token(u.strip())
    if not limpa or _PADROES_URL_INVALIDA.search(limpa):
        return False
    dom = extrair_dominio(limpa)
    if not dom or "." not in dom:
        return False
    if dom == "soundcloud.com" or dom.endswith(".soundcloud.com"):
        u_parse = urlparse(limpa if "://" in limpa else "https://" + limpa)
        return bool(PERFIS_FONTES["soundcloud.com"]["padrao_artigo"].search(u_parse.path or ""))
    if permitir_instagram and (dom == "instagram.com" or dom.endswith(".instagram.com")):
        return bool(re.search(r"/(?:p|reel)/[A-Za-z0-9_-]+", limpa))
    if any(dom == d or dom.endswith("." + d) for d in _DOMINIOS_IGNORAR):
        return False
    return True


def ler_cabecalho(page):
    blocos = [b for b in page.get_text("blocks") if b[3] < 40]
    linhas = [norm(l) for b in blocos for l in b[4].splitlines() if l.strip()]
    data_imp, resto = None, []
    for l in linhas:
        m = RE_IMPRESSAO.match(l)
        if m and data_imp is None:
            data_imp = mk_date(m[1], m[2], m[3])
            sobra = norm(l[m.end():])
            if sobra:
                resto.append(sobra)
        else:
            resto.append(l)
    return data_imp, " ".join(resto)


def urls_do_texto(doc):
    por_url = {}
    for n, page in enumerate(doc):
        h = page.rect.height
        for w in page.get_text("words"):
            u = limpar_token(w[4])
            if re.match(r"https?://", u) and url_valida(u):
                d = por_url.setdefault(u, {"pags": set(), "y": 0.0})
                d["pags"].add(n)
                d["y"] = max(d["y"], w[1] / h)
    return por_url


def ler_url_rodape(doc, ctx):
    por_url = urls_do_texto(doc)
    if not por_url:
        return ""
    if ctx["n"] >= 2:
        url, d = max(por_url.items(), key=lambda kv: (len(kv[1]["pags"]), kv[1]["y"]))
        if len(d["pags"]) >= max(2, (ctx["n"] + 1) // 2):
            return url
    for url, d in sorted(por_url.items(), key=lambda kv: -kv[1]["y"]):
        if d["y"] >= 0.94:
            return url
    return ""


def url_do_corpo(doc):
    por_url = urls_do_texto(doc)
    if not por_url:
        return ""
    return max(por_url.items(), key=lambda kv: (len(kv[1]["pags"]), kv[1]["y"]))[0]


def recuperar_url_completa(doc, url):
    prefixo = url.rstrip("…").rstrip(".")
    padrao = re.compile(re.escape(prefixo) + r"[^\s&|\"'<>]+")
    for page in doc:
        for l in page.get_links():
            uri_decodificada = unquote(l.get("uri") or "")
            m = padrao.search(uri_decodificada)
            if m:
                cand = limpar_token(m.group(0))
                if not cand.endswith(("…", "...")) and url_valida(cand):
                    return cand
    return url


def slugificar_titulo(titulo: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", sem_acento(titulo or "").lower()).strip("-")


def reconstruir_url_por_titulo(url: str, titulo: str) -> str:
    if not url or not titulo:
        return url
    perfil = obter_perfil_fonte(extrair_dominio(url))
    if not perfil or not perfil.get("reconstruir_slug"):
        return url
    limpa = url.rstrip("…").rstrip(".")
    if "/" not in limpa.split("://", 1)[-1]:
        return url
    base_url, slug_parcial = limpa.rsplit("/", 1)
    if not slug_parcial:
        return url
    slug_completo = slugificar_titulo(titulo)
    if len(slug_completo) <= len(slug_parcial):
        return url
    if slug_completo.startswith(slug_parcial):
        return f"{base_url}/{slug_completo}/"
    sp_limpo = slug_parcial.strip("-")
    sem_stop_p = "-".join(w for w in sp_limpo.split("-") if w and w not in _STOP_SLUG)
    sem_stop_c = "-".join(w for w in slug_completo.split("-") if w and w not in _STOP_SLUG)
    if len(sem_stop_p) >= 6 and sem_stop_c.startswith(sem_stop_p):
        return f"{base_url}/{slug_completo}/"
    janela = slug_completo[: len(sp_limpo) + 12]
    if len(sp_limpo) >= 8 and (
        sp_limpo in janela
        or difflib.SequenceMatcher(None, sp_limpo, slug_completo[: len(sp_limpo)]).ratio() >= 0.85
    ):
        return f"{base_url}/{slug_completo}/"
    return url


def recuperar_url_compartilhamento(doc):
    padrao_share = re.compile(
        r"(?:[?&](?:url|u)=|api\.whatsapp\.com/send\?text=.*?)(https?://[^\s&|\"'<>]+)",
        re.IGNORECASE,
    )
    for page in doc:
        for l in page.get_links():
            m = padrao_share.search(unquote(l.get("uri") or ""))
            if m:
                cand = limpar_token(m.group(1))
                if url_valida(cand):
                    return cand
    return ""


_STOP_SLUG = {
    "de", "da", "do", "das", "dos", "em", "no", "na", "nos", "nas", "para", "por",
    "com", "sem", "sob", "sobre", "um", "uma", "ao", "aos", "as", "os", "que", "se",
}


def _palavras_slug(s: str) -> set:
    return {w for w in re.findall(r"[a-z0-9]{3,}", sem_acento(s or "").lower()) if w not in _STOP_SLUG}


def recuperar_url_links_pagina(doc, titulo: str, titulo_aba: str = ""):
    pals = _palavras_slug(titulo) | _palavras_slug(titulo_aba)
    eh_insta = "instagram" in (titulo_aba or "").lower()
    melhores = []
    artigos_perfil = []
    por_dom = Counter()
    base_dom = {}

    for page in doc:
        for l in page.get_links():
            u = limpar_token(unquote(l.get("uri") or "").strip())
            if not u.startswith(("http://", "https://")):
                continue
            if not url_valida(u, permitir_instagram=eh_insta):
                continue
            dom = extrair_dominio(u)
            if eh_insta and "instagram.com" in dom and re.search(r"/(?:p|reel)/[A-Za-z0-9_-]+", u):
                return u, f"https://{dom}/"
            por_dom[dom] += 1
            esquema = urlparse(u).scheme or "https"
            base_dom.setdefault(dom, f"{esquema}://{urlparse(u).netloc}/")

            u_limpa = re.sub(
                r"/(?:politica-de-privacidade|politica-privacidade|respond|comments)/?$",
                "/",
                u.split("#")[0],
            )
            if not url_valida(u_limpa, permitir_instagram=eh_insta):
                continue
            parsed = urlparse(u_limpa)
            caminho = parsed.path or "/"
            if caminho == "/" or parsed.query.startswith("s="):
                continue

            perfil = obter_perfil_fonte(dom)
            ignorar_re = (perfil or {}).get("ignorar_caminhos") or _IGNORAR_CAMINHOS_PADRAO
            if ignorar_re.search(caminho):
                continue

            partes_path = [p for p in caminho.split("/") if p]
            slug = partes_path[-1] if partes_path else ""
            if re.fullmatch(r"a-\d+|\d+", slug) and len(partes_path) >= 2:
                slug = f"{partes_path[-2]}-{slug}"
            pals_slug = _palavras_slug(slug)
            inter = pals & pals_slug
            overlap = (len(inter) / len(pals)) if pals else 0.0

            pontos = 0
            eh_artigo_perfil = bool(perfil and perfil["padrao_artigo"].search(caminho))
            if eh_artigo_perfil:
                pontos += 10
                if u_limpa not in artigos_perfil:
                    artigos_perfil.append(u_limpa)

            bate_titulo = bool(pals) and (len(inter) >= 2 or overlap >= 0.35)
            if bate_titulo or (eh_artigo_perfil and len(inter) >= 1):
                pontos += len(inter) * 5 + int(overlap * 10)
                melhores.append((pontos, len(inter), len(slug), u_limpa))

    if melhores:
        melhores.sort(reverse=True)
        return melhores[0][3], ""
    if len(artigos_perfil) == 1:
        return artigos_perfil[0], ""
    if por_dom:
        dom_principal = por_dom.most_common(1)[0][0]
        return "", base_dom[dom_principal]
    return "", ""


def coletar_urls_candidatas(doc, permitir_instagram: bool = False) -> list:
    vistas, prioritarias, com_caminho, raizes = set(), [], [], []
    for page in doc:
        for l in page.get_links():
            u = limpar_token(unquote(l.get("uri") or "").strip())
            if not u.startswith(("http://", "https://")):
                continue
            u_limpa = re.sub(
                r"/(?:politica-de-privacidade|politica-privacidade|respond|comments)/?$",
                "/",
                u.split("#")[0],
            )
            if not url_valida(u_limpa, permitir_instagram=permitir_instagram) or u_limpa in vistas:
                continue
            vistas.add(u_limpa)
            parsed = urlparse(u_limpa)
            caminho = parsed.path or "/"
            if caminho.strip("/") == "" or parsed.query.startswith("s="):
                raizes.append(u_limpa)
                continue
            perfil = obter_perfil_fonte(extrair_dominio(u_limpa))
            ignorar_re = (perfil or {}).get("ignorar_caminhos") or _IGNORAR_CAMINHOS_PADRAO
            if ignorar_re.search(caminho):
                continue
            if perfil and perfil["padrao_artigo"].search(caminho):
                prioritarias.append(u_limpa)
            else:
                com_caminho.append(u_limpa)
    for u in urls_do_texto(doc):
        if u not in vistas and url_valida(u, permitir_instagram=permitir_instagram):
            vistas.add(u)
            com_caminho.append(u)
    nao_home = prioritarias + com_caminho
    return (nao_home if nao_home else raizes)[:25]


def _chamar_ollama(prompt: str, formato_json: bool = False, modelo: str = "") -> str | None:
    mod = (modelo or SLM.get("modelo") or "").strip()
    if not mod:
        return None
    payload = {
        "model": mod,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0},
    }
    if formato_json:
        payload["format"] = "json"
    dados = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        SLM.get("url", "http://localhost:11434/api/generate"),
        data=dados,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=SLM.get("timeout", 20)) as resp:
            corpo = json.loads(resp.read().decode("utf-8", errors="replace"))
            if isinstance(corpo, dict):
                return str(corpo.get("response") or "").strip()
    except (urllib.error.URLError, TimeoutError, OSError, ValueError, json.JSONDecodeError):
        return None
    return None


def escolher_url_slm(candidatas: list, titulo: str, texto: str = "", modelo: str = "") -> str:
    validas = []
    for u in candidatas or []:
        u_limpa = limpar_token(str(u).strip())
        if u_limpa and url_valida(u_limpa) and u_limpa not in validas:
            validas.append(u_limpa)
    if not validas:
        return ""
    mod = (modelo or SLM.get("modelo") or "qwen2.5:3b").strip()
    lista_num = "\n".join(f"{i}. {u}" for i, u in enumerate(validas, start=1))
    trecho = norm(texto)[:600] if texto else ""
    prompt = (
        "Você é um classificador de URLs de notícias.\n"
        f"Título da notícia: {titulo}\n"
        f"Trecho inicial: {trecho}\n\n"
        "Escolha na lista abaixo APENAS o número da URL que corresponde à matéria original.\n"
        "Se nenhuma URL corresponder à matéria específica, responda 0.\n"
        "Responda SOMENTE com o número inteiro (ex: 1 ou 0), sem texto adicional.\n\n"
        f"{lista_num}\n"
    )
    resp = _chamar_ollama(prompt, formato_json=False, modelo=mod)
    if not resp:
        return ""
    m = re.search(r"\b(\d+)\b", resp)
    if not m:
        return ""
    idx = int(m.group(1))
    if 1 <= idx <= len(validas):
        return validas[idx - 1]
    return ""


def extrair_metadados_slm(texto: str, modelo: str = "") -> dict:
    if not texto or not texto.strip():
        return {}
    mod = (modelo or SLM.get("modelo") or "qwen2.5:3b").strip()
    trecho = texto.strip()[:3000]
    prompt = (
        "Extraia os metadados da notícia abaixo e responda ESTRITAMENTE em JSON com as chaves: "
        '"titulo", "fonte", "data_publicacao" (formato DD/MM/AAAA), '
        '"hora_publicacao" (formato HH:MM) e "autor". '
        'Use "" para campos não encontrados.\n\n'
        f"Texto:\n{trecho}"
    )
    resp = _chamar_ollama(prompt, formato_json=True, modelo=mod)
    if not resp:
        return {}
    limpo = re.sub(r"^```(?:json)?\s*|\s*```$", "", resp.strip(), flags=re.IGNORECASE)
    try:
        dados = json.loads(limpo)
    except (ValueError, json.JSONDecodeError):
        m_json = re.search(r"\{.*\}", limpo, flags=re.DOTALL)
        if not m_json:
            return {}
        try:
            dados = json.loads(m_json.group(0))
        except (ValueError, json.JSONDecodeError):
            return {}
    if not isinstance(dados, dict):
        return {}

    out = {}
    for k in ("titulo", "fonte", "autor"):
        val = dados.get(k)
        if isinstance(val, str) and norm(val):
            out[k] = norm(val)

    raw_data = str(dados.get("data_publicacao") or "").strip()
    m_d = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$", raw_data)
    if m_d:
        dt = mk_date(m_d[1], m_d[2], m_d[3])
        if dt:
            out["data_publicacao"] = dt
    else:
        m_iso = re.match(r"^(20\d{2})-(\d{1,2})-(\d{1,2})$", raw_data)
        if m_iso:
            dt = mk_date(m_iso[3], m_iso[2], m_iso[1])
            if dt:
                out["data_publicacao"] = dt

    raw_hora = str(dados.get("hora_publicacao") or "").strip()
    m_h = re.match(r"^(\d{1,2})[:h](\d{2})$", raw_hora, flags=re.IGNORECASE)
    if m_h and 0 <= int(m_h[1]) <= 23 and 0 <= int(m_h[2]) <= 59:
        out["hora_publicacao"] = f"{int(m_h[1]):02d}:{m_h[2]}"

    return out


def _aplicar_fallback_metadados_slm(linha: dict, texto_p1: str, avisos: list):
    if not SLM.get("modelo") or not texto_p1.strip():
        return
    divergencia = any("difere" in a for a in avisos)
    faltando = any(not linha.get(c) for c in ("titulo", "data_publicacao", "fonte"))
    if not (faltando or divergencia):
        return
    meta_slm = extrair_metadados_slm(texto_p1)
    if not meta_slm:
        return
    preenchidos = []
    for campo in ("titulo", "fonte", "data_publicacao", "hora_publicacao"):
        if not linha.get(campo) and meta_slm.get(campo):
            linha[campo] = meta_slm[campo]
            preenchidos.append(campo)
    if "titulo" in preenchidos:
        avisos[:] = [a for a in avisos if a != "título não encontrado"]
    if "data_publicacao" in preenchidos:
        avisos[:] = [a for a in avisos if a != "data de publicação não encontrada"]
    if preenchidos:
        avisos.append(f"metadados preenchidos via SLM ({', '.join(preenchidos)}): conferir")


def dividir_titulo_fonte(titulo_completo):
    if " | " in titulo_completo:
        t, f = titulo_completo.rsplit(" | ", 1)
        return t.strip(), f.strip()
    m = re.match(r"^(.*\S)\s+[-–—]\s+([^-–—]{2,40})$", titulo_completo)
    if m:
        return m[1].strip(), m[2].strip()
    return titulo_completo.strip(), ""


MESES_ABREV = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
               "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
RE_DATA_ABREV = re.compile(
    r"\b(\d{1,2})[.\s/-]+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-zç]*[.\s/-]+(20\d{2})"
    r"(?:[^\n\d]{0,30}?(\d{1,2})[:h](\d{2}))?", re.IGNORECASE)


def ler_data_publicacao(texto_pag1):
    m = RE_PUB_NUM.search(texto_pag1)
    if m:
        d = mk_date(m[1], m[2], m[3])
        hora = f"{int(m[4]):02d}:{m[5]}" if m[4] else ""
        return d, hora
    m = RE_PUB_EXT.search(texto_pag1)
    if m and m[2].lower() in MESES:
        hora = f"{int(m[4]):02d}:{m[5]}" if m[4] else ""
        return mk_date(m[1], MESES[m[2].lower()], m[3]), hora
    m = RE_DATA_ABREV.search(texto_pag1)
    if m:
        hora = f"{int(m[4]):02d}:{m[5]}" if m[4] else ""
        return mk_date(m[1], MESES_ABREV[m[2].lower()], m[3]), hora
    return None, ""


RE_DATA_SOLTA = re.compile(
    r"(\d{1,2})/(\d{1,2})/(\d{2,4})(?:\s*(?:às|as|-|,)?\s*(\d{1,2})[:h](\d{2}))?", re.IGNORECASE)


def ler_data_solta(texto):
    for m in RE_DATA_SOLTA.finditer(texto):
        antes = sem_acento(texto[max(0, m.start() - 20):m.start()]).lower()
        if re.search(r"atualizad|modificad|editad|revisad", antes):
            continue
        d = mk_date(m[1], m[2], m[3])
        if d:
            return d, (f"{int(m[4]):02d}:{m[5]}" if m[4] else "")
    return None, ""


def ler_tags(doc):
    tags = []
    for page in doc:
        for l in page.get_links():
            if re.search(r"/tags?/", l.get("uri") or ""):
                nome = norm(page.get_textbox(l["from"]))
                if nome and nome not in tags:
                    tags.append(nome)
    return tags


def chave_linha(top, texto):
    return (round(top / 4), re.sub(r"\d+", "#", texto)[:80])


def linhas_dict(page):
    pr, saida = page.rect, []
    for b in page.get_text("dict", sort=True)["blocks"]:
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            spans = [s for s in l["spans"] if s["text"].strip()]
            x0, y0, x1, y1 = l["bbox"]
            if not spans or x1 < 0 or y1 < 0 or x0 > pr.width or y0 > pr.height:
                continue
            dom = max(spans, key=lambda s: len(s["text"].strip()))
            negrito = bool(dom["flags"] & 16) or bool(
                re.search(r"bold|black|heavy|semibold|demi", dom["font"], re.IGNORECASE))
            texto_linha = norm("".join(s["text"] for s in l["spans"]))
            if not texto_linha:
                continue
            saida.append({"top": y0, "bot": y1, "x0": x0, "size": dom["size"], "bold": negrito,
                          "texto": texto_linha})
    return saida


def contexto(doc):
    n, repetidas = len(doc), set()
    if n >= 2:
        cont = {}
        for page in doc:
            for k in {chave_linha(l["top"], l["texto"]) for l in linhas_dict(page)}:
                cont[k] = cont.get(k, 0) + 1
        repetidas = {k for k, c in cont.items() if c >= max(2, (n + 1) // 2)}
    return {"n": n, "repetidas": repetidas}


def em_moldura(l, page, ctx):
    if ctx["n"] >= 2:
        return chave_linha(l["top"], l["texto"]) in ctx["repetidas"]
    return l["top"] < 40 or l["bot"] > page.rect.height - 45


def linhas_corpo(page, ctx):
    return [l["texto"] for l in linhas_dict(page) if not em_moldura(l, page, ctx)]


def achar_titulo(linhas, titulo):
    alvo = norm(titulo)
    for s in range(len(linhas)):
        acc = ""
        for e in range(s, min(s + 6, len(linhas))):
            acc = norm(acc + " " + linhas[e])
            if acc == alvo:
                return s, e
            if not alvo.startswith(acc):
                break
    return None


def titulos_parecidos(a, b):
    f = lambda s: re.sub(r"[\W_]+", " ", sem_acento(s).lower()).strip()
    a, b = f(a), f(b)
    return bool(a and b) and (a in b or b in a or difflib.SequenceMatcher(None, a, b).ratio() >= 0.6)


def titulo_visual(page, ctx, referencia=""):
    h = page.rect.height
    linhas = [l for l in linhas_dict(page) if not em_moldura(l, page, ctx)]
    if not linhas:
        return ""
    tamanhos = []
    for l in linhas:
        tamanhos += [l["size"]] * len(l["texto"])
    if not tamanhos:
        return ""
    corpo = sorted(tamanhos)[len(tamanhos) // 2]

    linhas.sort(key=lambda l: (round(l["top"]), l["x0"]))
    grupos, atual = [], []
    for l in linhas:
        if (atual and abs(l["size"] - atual[-1]["size"]) <= 0.6
                and l["top"] - atual[-1]["bot"] <= 0.8 * atual[-1]["size"]):
            atual.append(l)
        else:
            if atual:
                grupos.append(atual)
            atual = [l]
    if atual:
        grupos.append(atual)

    cands = []
    for g in grupos:
        texto = norm(" ".join(l["texto"] for l in g))
        palavras = texto.lower().split()
        size = sum(l["size"] for l in g) / len(g)
        negrito = sum(l["bold"] for l in g) > len(g) / 2
        if len(palavras) < 2 or len(set(palavras)) < 2 or len(texto) < 10 or g[0]["top"] > 0.65 * h:
            continue
        if not (size >= 1.15 * corpo or (negrito and size >= corpo)):
            continue
        cands.append((size * (1.15 if negrito else 1.0), texto))
    if not cands:
        return ""
    if referencia:
        confirmados = [c for c in cands if titulos_parecidos(c[1], referencia)]
        if confirmados:
            cands = confirmados
    return max(cands, key=lambda c: c[0])[1]


def sem_camada_de_texto(doc):
    por_pagina = [len(p.get_text().strip()) for p in doc]
    if sum(por_pagina) < 50:
        return True
    return bool(por_pagina) and sum(1 for t in por_pagina if t < 30) / len(por_pagina) >= 0.6


_OCR_ESTADO = {}


def ocr_pronto():
    if "r" in _OCR_ESTADO:
        return _OCR_ESTADO["r"]
    try:
        import pytesseract
        from PIL import Image  # noqa: F401

        cmd = os.environ.get("TESSERACT_CMD") or shutil.which("tesseract")
        if not cmd:
            for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                      r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                      r"C:\Users\%s\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
                      % os.environ.get("USERNAME", "")):
                if os.path.exists(c):
                    cmd = c
                    break
        if cmd:
            pytesseract.pytesseract.tesseract_cmd = cmd
        langs = pytesseract.get_languages(config="")
        if "por" in langs and "eng" in langs:
            lang = "por+eng"
        else:
            lang = "por" if "por" in langs else "eng"
        res = (pytesseract, lang, "")
    except Exception as e:  # noqa: BLE001
        res = (None, "", f"{type(e).__name__}: {e}")
    _OCR_ESTADO["r"] = res
    return res


def preparar_imagem(img, variante):
    if not variante:
        return img
    from PIL import ImageOps, ImageStat

    g = ImageOps.autocontrast(img.convert("L"), cutoff=1)
    media = ImageStat.Stat(g).mean[0]
    if media < 110:
        g = ImageOps.invert(g)
        media = 255 - media
    if variante == "binaria":
        limiar = media * 0.85
        g = g.point(lambda p: 255 if p > limiar else 0)
    return g


def regioes_ocr(page, escaneado):
    if escaneado:
        return [(page.rect, OCR_DPI)]
    regioes, vistos = [], set()
    for info in page.get_image_info():
        bbox = pymupdf.Rect(info["bbox"])
        r = bbox & page.rect
        if r.is_empty or r.width < 80 or r.height < 8:
            continue
        chave = tuple(round(v) for v in r)
        if chave in vistos:
            continue
        vistos.add(chave)
        nativo = info.get("width", 0) / (bbox.width / 72) if bbox.width else 0
        regioes.append((r, int(min(OCR_DPI_MAX_IMAGEM, max(OCR_DPI, nativo)))))
    return regioes


def ocr_regiao(page, clip, n_pag, pt, lang, dpi=OCR_DPI, config="", variante=None, rotulo=""):
    from PIL import Image

    pix = page.get_pixmap(dpi=dpi, clip=clip)
    img = preparar_imagem(Image.frombytes("RGB", (pix.width, pix.height), pix.samples), variante)
    if DEBUG["dir"] and rotulo:
        pasta = Path(DEBUG["dir"])
        pasta.mkdir(parents=True, exist_ok=True)
        img.save(pasta / f"{DEBUG['pdf']}_{rotulo}.png")
    d = pt.image_to_data(img, lang=lang, config=config, output_type=pt.Output.DICT)
    esc = 72 / dpi
    por_linha = {}
    for i, txt in enumerate(d["text"]):
        txt = txt.strip()
        try:
            conf = float(d["conf"][i])
        except (TypeError, ValueError):
            conf = -1
        if not txt or conf < 0:
            continue
        y0 = clip.y0 + d["top"][i] * esc
        w = {"texto": txt, "conf": conf, "y0": y0, "y1": y0 + d["height"][i] * esc,
             "x1": clip.x0 + (d["left"][i] + d["width"][i]) * esc}
        chave = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        por_linha.setdefault(chave, []).append(w)

    saida = []
    for ws in por_linha.values():
        alturas = sorted(w["y1"] - w["y0"] for w in ws if len(w["texto"]) >= 2 and w["conf"] >= 50)
        saida.append({
            "pag": n_pag, "h_pag": page.rect.height, "clip_x1": clip.x1, "palavras": ws,
            "texto": norm(" ".join(w["texto"] for w in ws)), "extra": False,
            "top": min(w["y0"] for w in ws), "bot": max(w["y1"] for w in ws),
            "altura": alturas[len(alturas) * 3 // 4] if alturas else 0,
            "conf": sum(w["conf"] for w in ws) / len(ws),
        })
    return sorted(saida, key=lambda l: l["top"])


def limpar_token(t):
    s = t.strip("([<“\"'").rstrip(",;:)]}>”\"'»")
    if s.endswith(("...", "…")):
        return s
    return s.rstrip(".")


def tokens_url(linhas):
    for l in linhas:
        for w in l["palavras"]:
            tok = limpar_token(w["texto"])
            if len(tok) >= 12 and RE_URL_OCR.match(tok) and url_valida(tok):
                yield tok, w, l


def ocr_com_variantes(page, clip, n_pag, pt, lang, dpi, rotulo, config_base="--psm 6"):
    todas = []
    primeira = ocr_regiao(page, clip, n_pag, pt, lang, dpi=dpi, config=config_base,
                          variante="cinza", rotulo=f"{rotulo}_cinza")
    todas += primeira
    if any(w["conf"] >= CONF_BOA for _, w, _ in tokens_url(primeira)):
        return todas
    for variante, cfg in OCR_VARIANTES:
        if variante == "cinza" and cfg == config_base:
            continue
        ls = ocr_regiao(page, clip, n_pag, pt, lang, dpi=dpi, config=cfg, variante=variante,
                        rotulo=f"{rotulo}_{variante}_{cfg.split()[-1]}")
        for l in ls:
            l["extra"] = True
        todas += ls
        if any(w["conf"] >= CONF_BOA for _, w, _ in tokens_url(ls)):
            break
    return todas


def ler_ocr(doc, escaneado):
    pt, lang, erro = ocr_pronto()
    if pt is None:
        return None, [f"OCR indisponível ({erro}); instale Tesseract + pytesseract + pillow"]
    linhas, avisos, lidas = [], [], 0
    try:
        for n, page in enumerate(doc):
            if n >= OCR_MAX_PAGINAS:
                avisos.append(f"OCR limitado às {OCR_MAX_PAGINAS} primeiras páginas")
                break
            for k, (clip, dpi) in enumerate(regioes_ocr(page, escaneado)):
                if lidas >= OCR_MAX_REGIOES:
                    avisos.append(f"OCR limitado às primeiras {OCR_MAX_REGIOES} imagens")
                    return linhas, avisos
                lidas += 1
                base = ocr_regiao(page, clip, n, pt, lang, dpi=dpi, rotulo=f"p{n}_reg{k}_orig")
                linhas += base
                if not any(w["conf"] >= CONF_BOA for _, w, _ in tokens_url(base)):
                    linhas += [dict(l, extra=True) for l in
                               ocr_com_variantes(page, clip, n, pt, lang, dpi, f"p{n}_reg{k}",
                                                 config_base="--psm 3")]
    except Exception as e:  # noqa: BLE001
        avisos.append(f"falha no OCR: {type(e).__name__}: {e}")
    return linhas, avisos


def url_do_ocr(linhas):
    cands, suspeitos = {}, []
    for i, l in enumerate(linhas):
        for w in l["palavras"]:
            tok = limpar_token(w["texto"])
            if len(tok) >= 12 and RE_URL_OCR.match(tok) and url_valida(tok):
                borda = l["top"] < 0.1 * l["h_pag"] or l["bot"] > 0.9 * l["h_pag"]
                cortada = w["x1"] >= l["clip_x1"] - 5
                c = cands.setdefault(tok, {"url": tok, "conf": 0, "votos": 0, "borda": False,
                                           "cortada": False, "continua": ""})
                c["votos"] += 1
                c["conf"] = max(c["conf"], w["conf"])
                c["borda"] |= borda
                c["cortada"] |= cortada
                if (not c["continua"] and w is l["palavras"][-1] and RE_INICIO_URL.match(tok)
                        and (cortada or len(tok) >= 45) and i + 1 < len(linhas)):
                    prox = linhas[i + 1]
                    p0 = limpar_token(prox["palavras"][0]["texto"]) if prox["palavras"] else ""
                    if (prox["pag"] == l["pag"] and prox["top"] - l["bot"] < 2 * max(l["altura"], 6)
                            and RE_RESTO_URL.match(p0) and not RE_URL_OCR.match(p0)):
                        c["continua"] = p0
            elif RE_PARECE_LINK.search(tok) and url_valida(tok) and tok not in suspeitos:
                suspeitos.append(tok)

    ordem = sorted(cands.values(), key=lambda c: (not c["borda"], -c["votos"], -c["conf"], -len(c["url"])))
    avisos = []
    if ordem:
        c = ordem[0]
        av = f"URL lida por OCR em imagem (confiança {c['conf']:.0f}%): conferir no PDF"
        if c["conf"] < CONF_BOA:
            av += "; confiança baixa"
        if c["cortada"]:
            av += "; o link encosta na borda da imagem e pode estar cortado ou continuar na linha de baixo"
        avisos.append(av)
        if c["continua"]:
            avisos.append(f"o link pode continuar na linha seguinte da imagem: {c['url']}{c['continua']}")
        if len(ordem) > 1:
            avisos.append("outros links na imagem: " + ", ".join(o["url"] for o in ordem[1:4]))
        return c["url"], avisos
    if suspeitos:
        avisos.append("URL não reconhecida; texto parecido com link na imagem: "
                      + ", ".join(suspeitos[:3]))
    else:
        avisos.append("URL não encontrada (nem no texto, nem nas imagens)")
    return "", avisos


def titulo_do_ocr(linhas_p1, h_pag):
    cand = [l for l in linhas_p1
            if l["conf"] >= 50 and l["altura"] > 0 and l["top"] > 0.03 * h_pag and l["bot"] < 0.85 * h_pag]
    grupos, atual = [], []
    for l in cand:
        if (atual and abs(l["altura"] - atual[-1]["altura"]) <= 0.25 * atual[-1]["altura"]
                and l["top"] - atual[-1]["bot"] <= atual[-1]["altura"]):
            atual.append(l)
        else:
            if atual:
                grupos.append(atual)
            atual = [l]
    if atual:
        grupos.append(atual)
    grupos = [g for g in grupos
              if len(norm(" ".join(l["texto"] for l in g)).split()) >= 2
              and len(" ".join(l["texto"] for l in g)) >= 10
              and g[0]["top"] < 0.65 * h_pag]
    if not grupos:
        return ""
    g = max(grupos, key=lambda g: sum(l["altura"] for l in g) / len(g))
    return norm(" ".join(l["texto"] for l in g))


def faixa_moldura(page, onde, altura=MOLDURA_PT):
    r = page.rect
    if onde == "topo":
        return pymupdf.Rect(0, 0, r.width, altura)
    return pymupdf.Rect(0, r.height - altura, r.width, r.height)


def url_moldura_ocr(doc, pt, lang, onde):
    leituras, cortadas = [], set()
    for n, page in enumerate(doc):
        if n >= OCR_MAX_PAGINAS:
            break
        linhas = ocr_com_variantes(page, faixa_moldura(page, onde, FAIXA_URL_PT), n, pt, lang,
                                   OCR_DPI_MOLDURA, f"p{n}_{onde}")
        for tok, w, _ in tokens_url(linhas):
            leituras.append((tok, w["conf"]))
            if w["texto"].rstrip().endswith(("...", "…")):
                cortadas.add(tok)
    return leituras, cortadas


def url_rodape_ocr(doc, pt, lang):
    onde = "rodapé"
    leituras, cortadas = url_moldura_ocr(doc, pt, lang, "rodape")
    if not leituras:
        onde = "cabeçalho"
        leituras, cortadas = url_moldura_ocr(doc, pt, lang, "topo")
    if not leituras:
        return "", []
    votos = Counter(u for u, _ in leituras)
    conf = {u: max(c for x, c in leituras if x == u) for u in votos}
    url = max(votos, key=lambda u: (votos[u], conf[u], len(u)))
    av = (f"URL lida por OCR no {onde} ({votos[url]} de {len(leituras)} leituras iguais, "
          f"confiança {conf[url]:.0f}%): conferir no PDF")
    if conf[url] < CONF_BOA and not (votos[url] >= 3 and votos[url] / len(leituras) >= 0.6):
        av += "; confiança baixa"
    if len(votos) > 1:
        av += "; outras leituras: " + ", ".join(u for u in votos if u != url)[:300]
    if url in cortadas:
        av += "; URL cortada com '...' pelo navegador: completar manualmente"
    return url, [av]


RE_IMPRESSAO_OCR = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4}),?\s+(\d{1,2})[:.](\d{2})")


def cabecalho_ocr(page, pt, lang):
    linhas = ocr_regiao(page, faixa_moldura(page, "topo"), 0, pt, lang,
                        dpi=OCR_DPI_MOLDURA, config="--psm 6")
    texto = norm(" ".join(l["texto"] for l in linhas))
    m = RE_IMPRESSAO_OCR.search(texto)
    if not m:
        return None, texto
    return mk_date(m[1], m[2], m[3]), norm(texto[:m.start()] + " " + texto[m.end():])


def chave_titulo(s):
    return re.sub(r"[\W_]+", " ", sem_acento(s).lower()).strip()


def combina_com_meta(t, meta):
    if not t or not meta:
        return False
    alvo = chave_titulo(t)
    partes = [p for p in (chave_titulo(x) for x in re.split(r"\.{3}|…", meta)) if len(p) >= 5]
    if partes and all(p in alvo for p in partes):
        return True
    return titulos_parecidos(t, meta)


def extrair_escaneado(doc, linha, pdf_path):
    avisos = ["PDF sem texto (imagem ou letras vetorizadas): campos lidos por OCR, conferir"]
    meta = titulo_metadados(doc)
    titulo_meta, sufixo = dividir_titulo_fonte(meta.replace(" _ ", " | "))
    meta_cortado = bool(re.search(r"\.{3}|…", titulo_meta))
    meta_ok = bool(titulo_meta) and not meta_cortado and "_" not in titulo_meta

    pt, lang, erro = ocr_pronto()
    if pt is None:
        avisos = ["PDF sem texto (imagem ou letras vetorizadas)",
                  f"OCR indisponível ({erro}); instale Tesseract + pytesseract + pillow"]
        if titulo_meta:
            avisos.append("título vindo dos metadados do PDF"
                          + ("" if meta_ok else " (cortado/alterado pelo navegador)"))
        else:
            avisos.append("título não encontrado")
        linha.update({"titulo": titulo_meta, "fonte": sufixo, "avisos": "; ".join(avisos)})
        return linha

    h = doc[0].rect.height
    data_imp, titulo_cab = None, ""
    try:
        url, av = url_rodape_ocr(doc, pt, lang)
        avisos += av
        url_origem = "rodapé (OCR)" if url else ""
        data_imp, titulo_cab = cabecalho_ocr(doc[0], pt, lang)
        if url:
            p1 = ocr_regiao(doc[0], doc[0].rect, 0, pt, lang)
        else:
            ocr, av = ler_ocr(doc, escaneado=True)
            avisos += av
            url, av = url_do_ocr(ocr or [])
            avisos += av
            url_origem = "imagem (OCR)" if url else ""
            p1 = [l for l in (ocr or []) if l["pag"] == 0 and not l.get("extra")]
    except Exception as e:  # noqa: BLE001
        avisos.append(f"falha no OCR: {type(e).__name__}: {e}")
        linha.update({"titulo": titulo_meta, "fonte": sufixo, "avisos": "; ".join(avisos)})
        return linha

    corpo = [l for l in p1 if l["top"] >= MOLDURA_PT and l["bot"] <= h - MOLDURA_PT]

    t_ocr = titulo_do_ocr(corpo, h)
    if meta_ok:
        titulo = titulo_meta
    elif t_ocr and combina_com_meta(t_ocr, titulo_meta or titulo_cab):
        titulo = t_ocr
        avisos.append("título lido por OCR (o dos metadados está cortado/alterado): conferir")
    elif titulo_meta:
        titulo = titulo_meta
        avisos.append("título dos metadados, cortado/alterado pelo navegador"
                      + (f"; OCR leu «{t_ocr}»" if t_ocr else ""))
    else:
        titulo = t_ocr or titulo_cab
        if titulo:
            avisos.append("título lido por OCR: conferir")
    if not titulo:
        avisos.append("título não encontrado")

    if url.endswith(("…", "...")):
        reconstruida = reconstruir_url_por_titulo(url, titulo or titulo_meta)
        if reconstruida != url:
            nome_f = nome_fonte_dominio(extrair_dominio(reconstruida)) or extrair_dominio(reconstruida)
            url = reconstruida
            url_origem = f"rodapé + título ({nome_f})"
            avisos.append(f"URL cortada no rodapé; reconstruída pelo título ({nome_f}): conferir")

    texto = "\n".join(l["texto"] for l in corpo)
    data_pub, hora_pub = ler_data_publicacao(texto)
    if not data_pub and len(doc) > 1:
        try:
            p2 = [l for l in ocr_regiao(doc[1], doc[1].rect, 1, pt, lang) if l["bot"] <= h - MOLDURA_PT]
            data_pub, hora_pub = ler_data_publicacao("\n".join(l["texto"] for l in p2[:12]))
            if data_pub:
                avisos.append("data de publicação lida na 2ª página")
        except Exception as e:  # noqa: BLE001
            avisos.append(f"falha no OCR da 2ª página: {type(e).__name__}: {e}")
    if not data_pub:
        data_pub, hora_pub = ler_data_solta(texto)
        if data_pub:
            avisos.append("data lida sem 'Publicado em' (primeira data da página): conferir")
    if not data_pub:
        avisos.append("data de publicação não encontrada")
    mn = RE_NOME_ARQUIVO.match(pdf_path.name)
    if mn and data_pub:
        data_nome = mk_date(mn[1], mn[2], mn[3])
        if data_nome and data_nome != data_pub:
            avisos.append(f"data do nome do arquivo ({data_nome:%d/%m/%Y}) difere da data publicada")

    dominio = extrair_dominio(url)
    linha.update({
        "titulo": titulo,
        "fonte": nome_fonte_dominio(dominio) or sufixo or dominio,
        "data_publicacao": data_pub,
        "hora_publicacao": hora_pub,
        "url": url,
        "url_origem": url_origem,
    })
    _aplicar_fallback_metadados_slm(linha, texto, avisos)
    linha["avisos"] = "; ".join(avisos)
    return linha


def abrir_pdf(p: Path):
    try:
        return pymupdf.open(p)
    except Exception:
        if os.name != "nt":
            raise
        longo = Path("\\\\?\\" + str(p.resolve()))  # caminhos > 260 caracteres no Windows
        return pymupdf.open(stream=longo.read_bytes(), filetype="pdf")


def extrair(pdf_path: Path, raiz: Path) -> dict:
    DEBUG["pdf"] = re.sub(r"[^\w.-]+", "_", f"{pasta_relativa(pdf_path, raiz)}_{pdf_path.stem}").strip("_")
    with abrir_pdf(pdf_path) as doc:
        avisos = []
        linha = linha_base(pdf_path, raiz)

        if sem_camada_de_texto(doc):
            return extrair_escaneado(doc, linha, pdf_path)

        p1 = doc[0]
        ctx = contexto(doc)
        _, titulo_cab = ler_cabecalho(p1)
        titulo_completo = titulo_metadados(doc) or titulo_cab
        titulo_aba, sufixo = dividir_titulo_fonte(titulo_completo)

        titulo = titulo_visual(p1, ctx, titulo_aba)
        if titulo:
            if titulo_aba and not titulos_parecidos(titulo, titulo_aba):
                avisos.append(f"título em destaque difere do título da aba do navegador («{titulo_aba}»)")
        else:
            titulo = titulo_aba
            if titulo:
                avisos.append("título vindo da aba do navegador (nenhum destaque detectado na página)")

        corpo_p1 = linhas_corpo(p1, ctx)
        texto_p1 = "\n".join(corpo_p1)

        url, url_origem = ler_url_rodape(doc, ctx), "rodapé do PDF"
        if url.endswith(("…", "...")):
            completa = recuperar_url_completa(doc, url)
            if completa != url:
                url = completa
                avisos.append("URL cortada no rodapé; recuperada pelos links")
            else:
                reconstruida = reconstruir_url_por_titulo(url, titulo or titulo_aba)
                if reconstruida != url:
                    nome_f = nome_fonte_dominio(extrair_dominio(reconstruida)) or extrair_dominio(reconstruida)
                    url = reconstruida
                    url_origem = f"rodapé + título ({nome_f})"
                    avisos.append(f"URL cortada no rodapé; reconstruída pelo título ({nome_f}): conferir")
                else:
                    avisos.append("URL cortada no rodapé; conferir manualmente")
        if not url:
            url_share = recuperar_url_compartilhamento(doc)
            if url_share:
                url, url_origem = url_share, "link de compartilhamento"
                avisos.append("URL sem rodapé; recuperada dos links de compartilhamento")
        url_link_mat, url_dom_links = ("", "") if url else recuperar_url_links_pagina(doc, titulo, titulo_aba)
        if not url and url_link_mat:
            url, url_origem = url_link_mat, "link da página"
            avisos.append("URL sem rodapé; recuperada dos links clicáveis da página: conferir")
        if not url and SLM.get("modelo"):
            cands = coletar_urls_candidatas(doc, permitir_instagram="instagram" in (titulo_aba or "").lower())
            eh_fonte_mapeada = bool(obter_perfil_fonte(extrair_dominio(url_dom_links)))
            if len(cands) >= 2 or (eh_fonte_mapeada and len(cands) >= 1):
                url_slm = escolher_url_slm(cands, titulo or titulo_aba, texto_p1)
                if url_slm:
                    url, url_origem = url_slm, "link da página (SLM)"
                    avisos.append("URL sem rodapé; escolhida entre os links da página via SLM: conferir")
        if not url:
            url = url_do_corpo(doc)
            if url:
                url_origem = "texto do PDF (conferir)"
                avisos.append("URL tirada do corpo do texto, não do rodapé: pode não ser a do artigo")
        if not url and url_dom_links:
            url = url_dom_links
            url_origem = "home do site (link da matéria não encontrado no PDF)"
            avisos.append("PDF impresso sem rodapé; link específico da matéria não encontrado, preenchido com a home do site")
        if not url:
            pt, lang, _ = ocr_pronto()
            try:
                if pt is not None:
                    url, av = url_rodape_ocr(doc, pt, lang)
                    avisos += av
                    url_origem = "rodapé (OCR)" if url else ""
                if not url:
                    ocr, av = ler_ocr(doc, escaneado=False)
                    avisos += av
                    if ocr is not None:
                        url, av_url = url_do_ocr(ocr)
                        avisos += av_url
                        url_origem = "imagem (OCR)" if url else ""
                    elif pt is None:
                        avisos.append("URL não encontrada no texto")
            except Exception as e:  # noqa: BLE001
                avisos.append(f"falha no OCR: {type(e).__name__}: {e}")
        if not url:
            url_origem = ""

        dominio = extrair_dominio(url)
        fonte = nome_fonte_dominio(dominio) or sufixo or dominio

        data_pub, hora_pub = ler_data_publicacao(texto_p1)
        if not data_pub and len(doc) > 1:
            topo2 = [l["texto"] for l in linhas_dict(doc[1])
                     if l["bot"] <= doc[1].rect.height - MOLDURA_PT][:12]
            data_pub, hora_pub = ler_data_publicacao("\n".join(topo2))
            if data_pub:
                avisos.append("data de publicação lida na 2ª página")
        if not data_pub:
            data_pub, hora_pub = ler_data_solta(texto_p1)
            if data_pub:
                avisos.append("data lida sem 'Publicado em' (primeira data da página): conferir")
        if not data_pub:
            avisos.append("data de publicação não encontrada")
        m = RE_URL_ANO_MES.search(url)
        if data_pub and m and (data_pub.year, data_pub.month) != (int(m[1]), int(m[2])):
            avisos.append("data difere do ano/mês da URL")
        mn = RE_NOME_ARQUIVO.match(pdf_path.name)
        if mn and data_pub:
            data_nome = mk_date(mn[1], mn[2], mn[3])
            if data_nome and data_nome != data_pub:
                avisos.append(f"data do nome do arquivo ({data_nome:%d/%m/%Y}) difere da data publicada")

        if not titulo:
            avisos.append("título não encontrado")

        linha.update({
            "titulo": titulo,
            "fonte": fonte,
            "data_publicacao": data_pub,
            "hora_publicacao": hora_pub,
            "url": url,
            "url_origem": url_origem,
        })

        _aplicar_fallback_metadados_slm(linha, texto_p1, avisos)
        linha["avisos"] = "; ".join(avisos)
        return linha


COLUNAS = [
    "codigo", "pasta", "pasta_1", "pasta_2", "pasta_3", "titulo", "fonte", "data_publicacao",
    "hora_publicacao", "url", "url_origem",
]


def salvar_planilha(linhas, saida):
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Font, PatternFill

    df = pd.DataFrame(linhas)
    for c in COLUNAS:
        if c not in df:
            df[c] = ""
        elif c != "data_publicacao":
            df[c] = df[c].fillna("")

    df["data_publicacao"] = pd.to_datetime(df["data_publicacao"], errors="coerce")
    df = df[COLUNAS]
    df = df.apply(lambda col: col.map(
        lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v))

    resumo = (df.assign(com_url=df["url"].ne(""), sem_url=df["url"].eq(""))
              .groupby("pasta", sort=False)
              .agg(pdfs=("codigo", "size"), com_url=("com_url", "sum"),
                   sem_url=("sem_url", "sum"))
              .reset_index())
    resumo["pasta"] = resumo["pasta"].replace("", "(raiz)")

    with pd.ExcelWriter(saida, engine="openpyxl", datetime_format="DD/MM/YYYY") as xw:
        df.to_excel(xw, index=False, sheet_name="Notícias")
        resumo.to_excel(xw, index=False, sheet_name="Resumo")
        ws = xw.sheets["Notícias"]
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        larguras = {"codigo": 16, "pasta": 24, "pasta_1": 16, "pasta_2": 16, "pasta_3": 16,
                    "titulo": 60, "fonte": 18, "data_publicacao": 16,
                    "hora_publicacao": 12, "url": 50, "url_origem": 22}
        for i, nome in enumerate(COLUNAS, start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = larguras.get(nome, 16)
            ws.cell(1, i).font = Font(bold=True)
        wr = xw.sheets["Resumo"]
        for col, w in zip("ABCD", (40, 10, 10, 10)):
            wr.column_dimensions[col].width = w
        for c in wr[1]:
            c.font = Font(bold=True)
        i_data = COLUNAS.index("data_publicacao") + 1
        i_url = COLUNAS.index("url") + 1
        vermelho = PatternFill("solid", fgColor="FFC7CE")
        for r in range(2, ws.max_row + 1):
            ws.cell(r, i_data).number_format = "DD/MM/YYYY"
            c = ws.cell(r, i_url)
            if c.value:
                val = str(c.value).strip()
                c.hyperlink = val if val.lower().startswith(("http://", "https://")) else f"https://{val}"
                c.font = Font(color="0563C1", underline="single")
            else:
                c.fill = vermelho
    return df


def main():
    ap = argparse.ArgumentParser(description="Extrai título, fonte, data e URL de PDFs de notícias.")
    ap.add_argument("entrada", help="pasta com PDFs (busca em todas as subpastas) ou um arquivo .pdf")
    ap.add_argument("saida", help="planilha de saída (.xlsx)")
    ap.add_argument("--debug-ocr", metavar="PASTA",
                    help="salva nesta pasta os recortes (já pré-processados) que o OCR leu")
    ap.add_argument("--modelo-slm", metavar="MODELO", default="",
                    help="modelo Ollama local usado como fallback (ex.: qwen2.5:3b, llama3.2:3b)")
    args = ap.parse_args()
    DEBUG["dir"] = args.debug_ocr
    SLM["modelo"] = args.modelo_slm.strip()

    pasta, saida = Path(args.entrada), args.saida
    if pasta.is_file():
        pdfs, pasta = [pasta], pasta.parent
    elif pasta.is_dir():
        pdfs = listar_pdfs(pasta)
    else:
        sys.exit(f"Não encontrei: {pasta}")
    if not pdfs:
        sys.exit(f"Nenhum PDF em {pasta}")

    pt, lang, erro = ocr_pronto()
    if pt is None:
        print(f"ATENÇÃO: OCR indisponível ({erro}).\n"
              "  PDFs em imagem NÃO terão URL. Instale o Tesseract (veja o topo do arquivo).",
              file=sys.stderr)
    else:
        print(f"OCR pronto (idioma: {lang})")
        if "por" not in lang:
            print("  Aviso: idioma português não instalado no Tesseract; só inglês (títulos com acento "
                  "podem sair errados; URLs não são afetadas).", file=sys.stderr)

    linhas = []
    for i, p in enumerate(pdfs, start=1):
        rel = p.relative_to(pasta)
        print(f"[{i}/{len(pdfs)}] {rel}", flush=True)
        try:
            linhas.append(extrair(p, pasta))
        except Exception as e:
            linhas.append({**linha_base(p, pasta), "avisos": f"erro ao ler: {e}"})
    try:
        df = salvar_planilha(linhas, saida)
    except PermissionError:
        alt = str(Path(saida).with_name(Path(saida).stem + "_novo.xlsx"))
        print(f"A planilha {saida} está aberta (feche no Excel). Salvando como {alt}", file=sys.stderr)
        saida = alt
        df = salvar_planilha(linhas, saida)

    sem_url = int(df["url"].eq("").sum())
    print(f"{len(df)} PDFs processados -> {saida} ({sem_url} sem URL)")


if __name__ == "__main__":
    main()
