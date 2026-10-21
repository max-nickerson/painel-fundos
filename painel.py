"""
PAINEL DE FUNDOS - um arquivo so (dados abertos CVM + ANBIMA + Banco Central).

Como usar em qualquer computador:
  1) Instale o Python 3.11+ (python.org; no Windows marque "Add python.exe to PATH").
  2) No terminal:   pip install pandas numpy requests pyarrow openpyxl plotly fastapi uvicorn
  3) Salve este arquivo como painel.py e rode:   python painel.py
     -> baixa da CVM o que mudou, processa NOSSOS_FUNDOS + PEERS (listas abaixo, desde DESDE) e grava
        painel_fundos.html ao lado deste arquivo (um arquivo so, abre com dois cliques, sem servidor).
        Abas: Visao geral (retorno em % do CDI, vol, spread, duration) | Fundo (alocacao, spread/duration por categoria,
        retorno x CDI com atribuicao, carteira, melhores ativos em retorno e em bps, mudancas na carteira) | Performance de
        ativos (todas as carteiras e meses: retorno, vol, Sharpe, fundos, % PL; clique = peso em cada fundo, spread, PU,
        ANBIMA) | Exposicoes vs. peers (papeis e grupos que os peers tem e nos nao).
        Retorno do ativo = retorno total do mes (preco + juros/amortizacao pagos), comparavel ao CDI: debentures pelo PU par do SND,
        FIDC pela rentabilidade oficial, CRI/CRA/NC sem fluxo publico pelo carrego (queda forte de preco fica marcada para conferir).
        So entra carteira publica: meses com menos de 5% confidencial (a CVM permite esconder por ate ~90 dias).
  Painel antigo no navegador (localhost):   python painel.py servidor
  Excel:   python painel.py excel

Fontes: CVM (CDA mensal, informe diario, cadastro), ANBIMA (debentures), Banco Central (CDI). Tudo publico e gratuito.
Os dados ficam em %LOCALAPPDATA%/cvm_cache no Windows (fora do OneDrive; a 1a vez baixa ~2-3 GB da CVM) ou em dados_painel/
ao lado deste arquivo nos outros sistemas; PAINEL_DADOS=<pasta> escolhe outra.
Depois so baixa o que mudou; o painel processado do dia e reaproveitado (abre em segundos).
Arquivos mensais republicados pela CVM (fundos tem ate 3 meses para abrir a carteira toda) sao rebaixados 1x por dia.
"""
import sys
import io
import json
import os
import re
import time
import zipfile
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import requests

# ------------------------------------------------------------------ CONFIGURACAO (edite aqui)
NOSSOS_FUNDOS = {               # nome curto: CNPJ (qualquer formato)
    "GD MM": "14.416.823/0001-07",
    "GD MM Flexprev": "40.209.105/0001-70",
    "GD MM Ultra": "42.332.169/0001-99",
    "Precision": "32.292.528/0001-78",
    "Dual": "34.803.938/0001-61",
    "Prev": "49.803.664/0001-88",
    "Dual Prev": "58.156.912/0001-37",
    "Dual Global": "62.917.953/0001-76",
    "Dual Prev Dist": "65.983.811/0001-03",
    "Fuji": "67.269.289/0001-10",
    "GD RF": "32.973.123/0001-03",
    "Prev BB": "58.013.738/0001-73",
    "GD RF Flexprev": "39.566.756/0001-38",
}
PEERS = {                       # se um CNPJ tambem estiver em NOSSOS_FUNDOS, vale como nosso
    "XP Corporate Top Credito": "04.621.721/0001-70",
    "Valora Absolute": "10.326.625/0001-00",
    "XP Corporate Light": "11.046.179/0001-34",
    "Principal Claritas": "11.447.136/0001-60",
    "Sparta Top": "14.188.162/0001-00",
    "Plural Credito Corporativo": "18.316.558/0001-46",
    "Sparta Max": "26.773.148/0001-52",
    "Itau Multimercado Credito Privado": "28.840.420/0001-03",
    "AF Invest Geraes 30": "29.044.189/0001-04",
    "JGP Corporate Plus": "32.892.264/0001-93",
    "Dual Advanced": "34.803.938/0001-61",
    "Compass Yield 30": "36.318.479/0001-56",
    "Itau High Yield All": "42.263.927/0001-64",
    "Icatu Vanguarda": "42.501.967/0001-05",
    "ARX Vinson": "42.698.327/0001-29",
    "Itau Sinfonia Multimercado": "42.717.960/0001-17",
    "Occam Credito Corporativo": "47.586.648/0001-55",
    "Absolute Creta Selecao": "48.094.354/0001-79",
    "Mag High Grade Plus 30": "50.697.486/0001-37",
    "Vinland Credito Selecao": "50.980.211/0001-06",
    "Itau Sinfonia All": "54.278.393/0001-29",
    "Solis Capital Antares": "13.054.728/0001-48",
    "Capitania Premium 45": "20.146.294/0001-71",
    "AZ Quest Altro": "22.100.009/0001-07",
    "XP Corporate Plus": "23.999.611/0001-90",
    "ARX Everest": "32.102.131/0001-76",
    "Precision Advanced": "32.292.528/0001-78",
    "JGP Select": "32.892.615/0001-66",
    "AZ Quest Supra": "36.352.498/0001-07",
    "Root Capital Credito HG": "42.405.028/0001-59",
    "SPX Seahawk": "42.431.531/0001-89",
    "AF Horizonte": "44.025.131/0001-07",
    "Sparta Max 60": "44.643.192/0001-20",
    "Absolute Atenas Itau": "49.645.368/0001-04",
    "ARX RF": "41.575.611/0001-45",
    "ARX K2 IVP Inflacao Curta": "46.997.356/0001-42",
    "BTG Pactual Credito Corporativo": "42.827.247/0001-26",
    "Capitania Itau": "41.709.507/0001-04",
    "Capitania Idence Itau": "42.827.631/0001-29",
    "Icatu Vanguarda Absoluto II": "47.212.476/0001-50",
    "Icatu Vanguarda A": "34.781.249/0001-01",
    "Itau Active Fix": "41.301.133/0001-85",
    "Itau High Yield II": "42.860.483/0001-44",
    "JGP Credito Itau": "41.955.494/0001-45",
    "Kinea FI RF": "26.491.419/0001-87",
    "Porto Credito RF": "54.974.029/0001-01",
    "Schroder Idencia Itau I": "42.535.150/0001-40",
    "Sparta Inflacao IU": "43.737.649/0001-00",
    "Sparta IU": "46.997.384/0001-60",
    "SPX Seahawk Itau": "42.014.260/0001-66",
    "SulAmerica Cred ESG Itau": "45.615.787/0001-34",
    "Vinland Credito A T1 MM": "49.456.416/0001-08",
    "Absolute Delfos Itau": "49.645.692/0001-14",
    "Capitania Reit Itau": "42.934.005/0001-31",
    "Ibiuna I Credit Itau": "45.644.076/0001-98",
    "Icatu Vanguarda A Qualificado": "41.867.248/0001-31",
    "Itau BTG Pactual CorpPlus": "41.735.237/0001-06",
    "Itau Flexprev Advanced": "49.803.664/0001-88",
    "Itau Dual Prev": "58.156.912/0001-37",
    "Itau Sinfonia MM": "42.380.882/0001-08",
    "Vinland Credito IQ": "61.735.868/0001-24",
}
DESDE = "2019-01"               # primeiro mes de carteira/cota (CDA existe desde 2005; mais antigo = mais download na 1a vez)
# grupo economico: a CVM so informa o emissor. Regra = trecho do nome do emissor (sem acento, maiusculo) -> grupo.
# Securitizadoras (Opea, Vert, True...) ficam como estao: o devedor do lastro nao vem na CDA.
GRUPOS = {
    "LOCALIZA|UNIDAS": "Localiza", "REDE D.?OR|DOR SAO LUIZ": "Rede D'Or", "AEGEA|AGUAS DO RIO|AGUAS DO PARA|EQUIPAV|CORSAN|RIOGRANDENSE DE SANEAMENTO": "Aegea",
    "BRK AMBIENTAL": "BRK", "EQUATORIAL|CEEE": "Equatorial", "ENERGISA": "Energisa", "CEMIG": "Cemig", "COPEL": "Copel",
    "SABESP|SANEAMENTO BASICO DO ESTADO DE SAO PAULO": "Sabesp", "COPASA|SANEAMENTO DE MINAS": "Copasa", "PETROBRAS|PETROLEO BRASILEIRO": "Petrobras",
    "PRIO|PETRO RIO": "Prio", "PETRORECONCAVO": "PetroReconcavo", "COSAN|RAIZEN|RUMO|COMPASS|MOOVE": "Cosan", "VALE S": "Vale",
    "SUZANO": "Suzano", "KLABIN": "Klabin", "SENDAS|ASSAI": "Assaí", "HAPVIDA|NOTRE DAME": "Hapvida", "VAMOS|SIMPAR|JSL|MOVIDA|CS INFRA": "Simpar",
    "MILLS": "Mills", "EUROFARMA": "Eurofarma", "IOCHPE": "Iochpe-Maxion", "CCR|MOTIVA": "Motiva (CCR)", "ECORODOVIAS": "EcoRodovias",
    "ELETROBRAS|CENTRAIS ELETRICAS BRASILEIRAS|FURNAS|CHESF|ELETRONORTE": "Eletrobras", "TAESA|TRANSMISSORA ALIANCA": "Taesa",
    "ISA ENERGIA|CTEEP": "ISA Energia", "AUREN|CESP|AES BRASIL": "Auren", "NOVA TRANSPORTADORA DO SUDESTE": "NTS", "VLI ": "VLI",
    "ITAU": "Itaú", "BRADESCO": "Bradesco", "SANTANDER": "Santander", "BANCO DO BRASIL": "Banco do Brasil", "CAIXA ECONOMICA": "Caixa Econômica",
    "BTG": "BTG Pactual", "VOTORANTIM|BANCO BV": "BV", "XP ": "XP", "SAFRA": "Safra", "DAYCOVAL": "Daycoval",
    "TESOURO NACIONAL|SECRETARIA DO TESOURO": "Tesouro Nacional",
}
OUT = "lookthrough_cvm.xlsx"
WORKERS = 4

URL = "https://dados.cvm.gov.br/dados/FI/DOC"
CACHE = Path(os.environ.get("PAINEL_DADOS") or (Path(os.environ["LOCALAPPDATA"]) / "cvm_cache" if os.environ.get("LOCALAPPDATA")
             else Path(__file__).resolve().parent / "dados_painel"))   # Windows: fora do OneDrive (GBs da CVM); senao ao lado do .py
(CACHE / "rec").mkdir(parents=True, exist_ok=True)
VERSAO_DIARIO = 4               # recorte das cotas diarias (3: fundos que so informam cota por subclasse)
VERSAO = 3                      # muda quando o recorte guarda colunas novas (forca reprocessar); 3: CNPJ do emissor e descricao
REN = {"CNPJ_FUNDO": "CNPJ_FUNDO_CLASSE", "CNPJ_FUNDO_COTA": "CNPJ_FUNDO_CLASSE_COTA",
       "NM_FUNDO_COTA": "NM_FUNDO_CLASSE_SUBCLASSE_COTA"}             # nomes antigos (ate 2022) -> novos
NUM = ["QT_POS_FINAL", "VL_CUSTO_POS_FINAL", "VL_AQUIS_NEGOC", "VL_VENDA_NEGOC", "PR_INDEXADOR_POSFX",
       "PR_CUPOM_POSFX", "PR_TAXA_PREFX"]
COLS = {"CNPJ_FUNDO_CLASSE", "DT_COMPTC", "TP_APLIC", "TP_ATIVO", "VL_MERC_POS_FINAL", "CNPJ_FUNDO_CLASSE_COTA",
        "NM_FUNDO_CLASSE_SUBCLASSE_COTA", "DS_ATIVO", "CD_ATIVO", "CD_ISIN", "TP_TITPUB", "DT_VENC", "EMISSOR",
        "CNPJ_EMISSOR", "CPF_CNPJ_EMISSOR", "DS_SWAP", "CD_SWAP", "DS_ATIVO_EXTERIOR", "CD_ATIVO_BV_MERC", "PAIS",
        "DS_INDEXADOR_POSFX", "GRAU_RISCO", "AG_RISCO", *NUM} | set(REN)
PASSIVO = r"(?i)a pagar|obriga|exigibilidade"                        # entram negativos (PL = ativos - passivos)
DERIV = r"(?i)mercado futuro|opç|swap"   # derivativos (nocional): fora dos 100%; termo de acoes e posicao real e entra
AJUSTE = "Ajuste carteira x PL"


def fmt(c):
    d = re.sub(r"\D", "", c)
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


def cnpj_of(x):
    if "maisretorno" in x:
        m = re.search(r"id=(\d{14})", x) or re.search(
            r"CNPJ\D{0,30}(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", requests.get(x, timeout=60, headers={"User-Agent": "Mozilla/5.0"}).text)
        x = m.group(1)
    return fmt(x)


def baixa(url):
    f = CACHE / url.rsplit("/", 1)[1]
    if not f.exists():
        tmp = f.with_suffix(f".{os.getpid()}.part")
        with requests.get(url, stream=True, timeout=600) as r:
            r.raise_for_status()
            with open(tmp, "wb") as o:
                for ch in r.iter_content(1 << 20):
                    o.write(ch)
        tmp.replace(f)
    return f


def lista(pasta, padrao):        # arquivos disponiveis no diretorio da CVM
    return sorted(set(re.findall(padrao, requests.get(f"{URL}/{pasta}", timeout=60).text)))


def arquivos(tipo, desde):       # {arquivo: [meses]} - mensal recente + anual (HIST) antigo
    mens = lista(f"{tipo}/DADOS/", rf"{tipo.lower()}_fi_(\d{{6}})\.zip")
    anos = lista(f"{tipo}/DADOS/HIST/", rf"{tipo.lower()}_fi_(\d{{4}})\.zip")
    out = {}
    for a in anos:
        ms = [f"{a}-{m:02d}" for m in range(1, 13) if f"{a}-{m:02d}" >= desde and f"{a}{m:02d}" not in mens]
        if ms:
            out[f"{URL}/{tipo}/DADOS/HIST/{tipo.lower()}_fi_{a}.zip"] = ms
    for m in mens:
        if f"{m[:4]}-{m[4:]}" >= desde:
            out[f"{URL}/{tipo}/DADOS/{tipo.lower()}_fi_{m}.zip"] = [f"{m[:4]}-{m[4:]}"]
    return out


def le(z, nome, cols):
    df = pd.read_csv(z.open(nome), sep=";", encoding="latin1", dtype=str, low_memory=False, quoting=3,
                     usecols=lambda c: c in cols)                      # quoting=3: CVM nao usa aspas (e tem " soltas)
    return df.rename(columns=REN)


def marca(kind, m):
    f = CACHE / "rec" / f"{kind}_{m}.json"
    try:
        j = json.loads(f.read_text())
        return set(j["fundos"]) if isinstance(j, dict) and j.get("v") == (VERSAO_DIARIO if kind == "diario" else VERSAO) else None
    except (OSError, ValueError):
        return None


def fecho(lk, peers, base=()):   # tudo que os peers alcancam via cotas de fundos
    alvo, novo = set(peers) | set(base), set(peers)
    while novo:
        novo = set(lk[lk.CNPJ_FUNDO_CLASSE.isin(novo)].CNPJ_FUNDO_CLASSE_COTA) - alvo
        alvo |= novo
    return alvo


def processa_cda(url, meses, peers):
    """Le um zip CDA: links fundo->cota e PL de todos; posicoes so dos fundos ligados aos peers."""
    z = zipfile.ZipFile(baixa(url))
    nomes = z.namelist()
    if not any("_PL_" in n for n in nomes) or not any("_BLC_2_" in n for n in nomes):
        return url + " (CVM ainda publicando este mes: pulado)"
    pl = le(z, next(n for n in nomes if "_PL_" in n), {"CNPJ_FUNDO_CLASSE", "CNPJ_FUNDO", "DT_COMPTC", "VL_PATRIM_LIQ"})
    blc = pd.concat([le(z, n, COLS).assign(conf="CONFID" in n) for n in nomes if "_BLC_" in n or "CONFID" in n],
                    ignore_index=True)
    for df in (pl, blc):
        df["mes"] = df.DT_COMPTC.str[:7]
    links = blc.dropna(subset=["CNPJ_FUNDO_CLASSE_COTA"])[["mes", "CNPJ_FUNDO_CLASSE", "CNPJ_FUNDO_CLASSE_COTA"]]
    for m in meses:
        lk = links[links.mes == m]
        alvo = fecho(lk, peers, marca("cda", m) or ())
        b = blc[(blc.mes == m) & blc.CNPJ_FUNDO_CLASSE.isin(alvo)].copy()
        b["v"] = pd.to_numeric(b.VL_MERC_POS_FINAL, errors="coerce").fillna(0)
        b.loc[b.TP_APLIC.fillna("").str.contains(PASSIVO), "v"] *= -1

        def first(*cs):                           # primeira coluna preenchida entre as candidatas
            out = pd.Series(None, index=b.index, dtype=object)
            for c in cs:
                if c in b:
                    out = out.where(out.notna(), b[c])
            return out
        b["ativo"] = first("DS_ATIVO", "NM_FUNDO_CLASSE_SUBCLASSE_COTA", "DS_SWAP", "DS_ATIVO_EXTERIOR", "TP_TITPUB", "TP_ATIVO")
        b["emissor"] = first("EMISSOR", "CNPJ_EMISSOR", "CPF_CNPJ_EMISSOR")
        b["cnpj_emissor"] = first("CPF_CNPJ_EMISSOR", "CNPJ_EMISSOR")
        b["descricao"] = first("DS_ATIVO")
        b["codigo"] = first("CD_ATIVO", "CD_ISIN", "CNPJ_FUNDO_CLASSE_COTA", "CD_SWAP", "CD_ATIVO_BV_MERC")
        b.loc[b.conf, "ativo"] = "(confidencial) " + b.loc[b.conf, "TP_APLIC"].fillna("")
        b["ativo"] = b.ativo.fillna(b.TP_ATIVO).fillna(b.TP_APLIC)
        venc = b.DT_VENC.fillna("").str[:7] if "DT_VENC" in b else ""
        b["ativo"] = (b.ativo.fillna("") + " " + b.emissor.fillna("").where(b.ativo.fillna("") != b.emissor.fillna(""), "")
                      + " " + venc).str.replace(r"\s+", " ", regex=True).str.strip()
        for c in NUM:
            b[c] = pd.to_numeric(b[c], errors="coerce") if c in b else None
        keep = ["CNPJ_FUNDO_CLASSE", "TP_APLIC", "TP_ATIVO", "ativo", "codigo", "emissor", "cnpj_emissor", "descricao", "CNPJ_FUNDO_CLASSE_COTA", "v",
                "DT_VENC", "DS_INDEXADOR_POSFX", "GRAU_RISCO", "AG_RISCO", "PAIS", *NUM]
        destino = CACHE / "rec" / f"cda_{m}.parquet"
        if destino.exists():                     # point in time: guarda a versao anterior do recorte (republicacao da CVM)
            (CACHE / "rec" / "versoes").mkdir(exist_ok=True)
            destino.replace(CACHE / "rec" / "versoes" / f"cda_{m}_{date.fromtimestamp(destino.stat().st_mtime)}.parquet")
        b.reindex(columns=keep).assign(visto_em=str(date.today())).to_parquet(destino, index=False)
        lk.drop(columns="mes").to_parquet(CACHE / "rec" / f"links_{m}.parquet", index=False)
        pl[pl.mes == m][["CNPJ_FUNDO_CLASSE", "VL_PATRIM_LIQ"]].to_parquet(CACHE / "rec" / f"pl_{m}.parquet", index=False)
        (CACHE / "rec" / f"cda_{m}.json").write_text(json.dumps({"v": VERSAO, "fundos": sorted(alvo)}))
    return url


def processa_diario(url, meses, peers):
    z = zipfile.ZipFile(baixa(url))
    cols = {"CNPJ_FUNDO_CLASSE", "CNPJ_FUNDO", "ID_SUBCLASSE", "DT_COMPTC", "VL_QUOTA", "VL_PATRIM_LIQ", "CAPTC_DIA",
            "RESG_DIA", "NR_COTST"}
    d = pd.concat([le(z, n, cols) for n in z.namelist()], ignore_index=True)
    alvo = set(peers).union(*[marca("diario", m) or set() for m in meses])   # mantem quem ja estava no recorte
    d = d[d.CNPJ_FUNDO_CLASSE.isin(alvo)]
    if "ID_SUBCLASSE" in d:                       # cota da classe; se a classe so informa por subclasse, a de maior PL
        cls = d[d.ID_SUBCLASSE.isna()]
        tem = set(zip(cls.CNPJ_FUNDO_CLASSE, cls.DT_COMPTC))                 # decide dia a dia (fundo migrando no mes)
        sub = d[d.ID_SUBCLASSE.notna()]
        sub = sub[[k not in tem for k in zip(sub.CNPJ_FUNDO_CLASSE, sub.DT_COMPTC)]]
        if len(sub):
            pl = pd.to_numeric(sub.VL_PATRIM_LIQ, errors="coerce").groupby([sub.CNPJ_FUNDO_CLASSE, sub.ID_SUBCLASSE]).sum()
            maior = pl.groupby(level=0).idxmax().map(lambda t: t[1])
            sub = sub[sub.ID_SUBCLASSE == sub.CNPJ_FUNDO_CLASSE.map(maior)]
        d = pd.concat([cls, sub], ignore_index=True)
    for m in meses:
        d[d.DT_COMPTC.str[:7] == m].to_parquet(CACHE / "rec" / f"diario_{m}.parquet", index=False)
        (CACHE / "rec" / f"diario_{m}.json").write_text(json.dumps({"v": VERSAO_DIARIO, "fundos": sorted(alvo)}))
    return url


def precisa(kind, m, peers):     # recorte em cache cobre os fundos pedidos?
    tem = marca(kind, m)
    if tem is None:
        return True
    if kind == "diario":
        return not set(peers) <= tem
    return not fecho(pd.read_parquet(CACHE / "rec" / f"links_{m}.parquet"), peers) <= tem


def _modificado(u):
    try:
        return u, pd.Timestamp(requests.head(u, timeout=60).headers["Last-Modified"]).timestamp()
    except (requests.RequestException, KeyError, ValueError):
        return u, None


def atualiza_publicacoes(cda, dia, log=print):
    """A CVM republica o arquivo do mes conforme os fundos entregam ou abrem a carteira (ha meses de prazo).
    Uma vez por dia confere a data de TODOS os arquivos mensais e rebaixa os que mudaram."""
    marca_dia = CACHE / f"verificado_{date.today()}.txt"
    if marca_dia.exists():
        return
    urls = [u for u in list(cda) + list(dia) if "HIST" not in u and (CACHE / u.rsplit("/", 1)[1]).exists()]
    with ThreadPoolExecutor(8) as ex:
        mudou = [u for u, t in ex.map(_modificado, urls) if t and t > (CACHE / u.rsplit("/", 1)[1]).stat().st_mtime]
    for u in mudou:
        (CACHE / u.rsplit("/", 1)[1]).unlink(missing_ok=True)
        for m in (cda.get(u) or dia.get(u)):
            for j in CACHE.glob(f"rec/*_{m}.json"):
                j.unlink(missing_ok=True)
    if mudou:
        log(f"{len(mudou)} arquivo(s) da CVM republicado(s) desde a ultima carga: baixando de novo")
    for velho in CACHE.glob("verificado_*.txt"):
        velho.unlink(missing_ok=True)
    marca_dia.write_text(f"{len(urls)} arquivos conferidos, {len(mudou)} atualizados")


def prepara(cnpjs, desde, workers=WORKERS, log=print):
    """Baixa/recorta da CVM so o que falta. Devolve (meses de carteira, {arquivo diario: meses})."""
    cda, dia = arquivos("CDA", desde), arquivos("INF_DIARIO", desde)
    atualiza_publicacoes(cda, dia, log)
    tarefas = ([(processa_cda, u, ms) for u, ms in cda.items() if any(precisa("cda", m, cnpjs) for m in ms)]
               + [(processa_diario, u, ms) for u, ms in dia.items() if any(precisa("diario", m, cnpjs) for m in ms)])

    def roda(i, fn):
        try:
            log(f"{i}/{len(tarefas)} {fn().rsplit('/', 1)[1]}")
        except Exception as e:                    # um arquivo ruim da CVM nao derruba o resto
            log(f"{i}/{len(tarefas)} ATENCAO: arquivo pulado ({type(e).__name__}: {e})")
    if workers > 1 and len(tarefas) > 1:
        with ProcessPoolExecutor(workers) as ex:
            fut = [ex.submit(f, u, ms, cnpjs) for f, u, ms in tarefas]
            for i, f in enumerate(fut, 1):
                roda(i, f.result)
    else:
        for i, (f, u, ms) in enumerate(tarefas, 1):
            roda(i, lambda: f(u, ms, cnpjs))
    meses = sorted(m for ms in cda.values() for m in ms if (CACHE / "rec" / f"cda_{m}.parquet").exists())
    return meses, dia


def fund_cat(name):
    n = (name or "").upper()
    return ("FIDC" if "FIDC" in n or "DIREITOS CREDIT" in n else "FIAGRO" if "FIAGRO" in n or "CADEIAS PRODUTIVAS" in n
            else "Cotas de Fundos")


def carrega_pos(meses):
    """Posicoes (so dos fundos recortados) + PL de todos, todos os meses, com a linha de ajuste carteira x PL."""
    pos = pd.concat([pd.read_parquet(CACHE / "rec" / f"cda_{m}.parquet").assign(mes=m) for m in meses], ignore_index=True)
    pl = pd.concat([pd.read_parquet(CACHE / "rec" / f"pl_{m}.parquet").assign(mes=m) for m in meses], ignore_index=True)
    pl = pl.drop_duplicates(["mes", "CNPJ_FUNDO_CLASSE"]).set_index(["mes", "CNPJ_FUNDO_CLASSE"]).VL_PATRIM_LIQ.astype(float)
    pos = pos.join(pl.rename("pl_fundo"), on=["mes", "CNPJ_FUNDO_CLASSE"])
    # a CDA nem sempre permite dar sinal certo a toda linha (ex.: compromissada tomada x doada): a diferenca
    # entre PL e soma da carteira vira uma linha explicita, em vez de inflar/encolher as categorias
    der = pos.TP_APLIC.fillna("").str.contains(DERIV)
    aj = (pos[~der].groupby(["mes", "CNPJ_FUNDO_CLASSE"]).agg(soma=("v", "sum"), pl_fundo=("pl_fundo", "first"))
          .assign(v=lambda a: a.pl_fundo - a.soma))
    aj = aj[(aj.v.abs() > 0.005 * aj.pl_fundo.abs())].reset_index().assign(
        TP_APLIC=AJUSTE, ativo="Diferença entre PL informado e soma da carteira (fonte CVM)")
    pos = pd.concat([pos, aj.drop(columns="soma")], ignore_index=True)
    return pos, pl


def lookthrough(nomes, meses, carga=None):
    """nomes: {nome: cnpj}. Uma linha por ativo final x caminho, com % do PL do fundo e movimentacao do mes."""
    pos, pl = carga or carrega_pos(meses)
    tem_cart = set(zip(pos.mes, pos.CNPJ_FUNDO_CLASSE))
    front = pd.DataFrame([dict(fundo=n, mes=m, veic=c, peso=1.0, caminho=n, cadeia=c)
                          for n, c in nomes.items() for m in meses if (m, c) in tem_cart])
    folhas = []
    for _ in range(12):
        if front.empty:
            break
        r = front.merge(pos, left_on=["mes", "veic"], right_on=["mes", "CNPJ_FUNDO_CLASSE"])
        f = r.peso / r.pl_fundo * 100              # R$ no veiculo -> % do PL do peer
        r["perc_pl"], r["compra_pct"], r["venda_pct"] = r.v * f, r.VL_AQUIS_NEGOC * f, r.VL_VENDA_NEGOC * f
        r["fator"] = f
        cota = r.CNPJ_FUNDO_CLASSE_COTA
        abre = cota.notna() & pd.Series([(m, c) in tem_cart and c not in cad for m, c, cad in zip(r.mes, cota, r.cadeia)],
                                        index=r.index, dtype=bool)
        nxt = r[abre]
        front = pd.DataFrame(dict(fundo=nxt.fundo, mes=nxt.mes, veic=nxt.CNPJ_FUNDO_CLASSE_COTA, peso=nxt.perc_pl / 100,
                                  caminho=nxt.caminho + " > " + nxt.ativo, cadeia=nxt.cadeia + "|" + nxt.CNPJ_FUNDO_CLASSE_COTA))
        folha = r[~abre].copy()
        eh_cota = folha.TP_APLIC.eq("Cotas de Fundos")
        folha["categoria"] = folha.TP_APLIC.where(~eh_cota, folha.ativo.map(fund_cat))
        folha["veiculo"] = folha.caminho.str.split(" > ").str[-1]
        folhas.append(folha)
    if not folhas:                              # CVM sem nenhuma carteira desse CNPJ no periodo
        return pd.DataFrame()
    df = pd.concat(folhas, ignore_index=True)
    df = df.rename(columns={"veic": "cnpj_veiculo", "TP_ATIVO": "tipo_ativo", "DT_VENC": "vencimento",
                            "DS_INDEXADOR_POSFX": "indexador", "PR_INDEXADOR_POSFX": "pct_indexador",
                            "PR_CUPOM_POSFX": "cupom", "PR_TAXA_PREFX": "taxa_pre", "GRAU_RISCO": "rating",
                            "AG_RISCO": "agencia", "PAIS": "pais", "v": "valor_veiculo", "QT_POS_FINAL": "quantidade",
                            "VL_CUSTO_POS_FINAL": "custo", "VL_AQUIS_NEGOC": "compras_veiculo",
                            "VL_VENDA_NEGOC": "vendas_veiculo", "pl_fundo": "pl_veiculo"})
    df = df.reindex(columns=["fundo", "mes", "categoria", "ativo", "codigo", "emissor", "cnpj_emissor", "descricao", "perc_pl", "compra_pct", "venda_pct", "veiculo",
             "cnpj_veiculo", "caminho", "tipo_ativo", "vencimento", "indexador", "pct_indexador", "cupom", "taxa_pre",
             "rating", "agencia", "pais", "quantidade", "valor_veiculo", "custo", "compras_veiculo", "vendas_veiculo",
             "pl_veiculo", "fator"])
    df["categoria"] = df.categoria.fillna("Outros")
    df["codigo"] = df.codigo.str.strip()                  # CVM manda codigos com espacos (ex.: "ASER12  ")
    curto = df.ativo.fillna("").str.len() <= 4             # descricao generica ("DEB"): usa codigo + emissor
    df.loc[curto, "ativo"] = (df.codigo.fillna("") + " " + df.emissor.fillna("")).str.strip()[curto]
    df["derivativo"] = df.categoria.str.contains(DERIV)
    df["confidencial"] = df.ativo.str.startswith("(confidencial)", na=False)
    df["chave"] = df.cnpj_veiculo + "|" + df.codigo.fillna(df.ativo)
    return df


def mes_aberto(df, limite=5):
    """Por fundo, o mes mais recente em que a parte confidencial da carteira fica abaixo de `limite`% do PL."""
    t = df.groupby(["fundo", "mes"]).agg(conf=("perc_pl", lambda s: s[df.loc[s.index, "confidencial"]].sum()))
    ok = t[t.conf < limite].reset_index()
    return ok.groupby("fundo").mes.max()


def contribuicao(df):
    """Estimativa mensal de resultado por ativo pela variacao de preco das posicoes mantidas: quantidade do mes anterior
    x (preco unitario do mes - preco do mes anterior), em % do PL do peer. Preco unitario = valor / quantidade (CDA).
    Fica de fora: ganho de negociacao dentro do mes, carrego de caixa/compromissadas (sem quantidade) e ativos
    confidenciais. Cupons/amortizacoes pagos aparecem como queda de preco."""
    x = df[~df.derivativo & ~df.confidencial & ~df.categoria.eq(AJUSTE) & df.quantidade.gt(0)]
    k = x.groupby(["fundo", "mes", "chave"]).agg(v=("valor_veiculo", "sum"), qt=("quantidade", "sum"),
                                                  fator=("fator", "first"), categoria=("categoria", "first"),
                                                  ativo=("ativo", "first"), emissor=("emissor", "first")).reset_index()
    meses = sorted(df.mes.unique())
    seguinte = dict(zip(meses[:-1], meses[1:]))
    prev = k[["fundo", "mes", "chave", "v", "qt"]].assign(mes=k.mes.map(seguinte)).dropna(subset=["mes"])
    k = k.merge(prev.rename(columns={"v": "v_ant", "qt": "qt_ant"}), on=["fundo", "mes", "chave"])
    k["resultado_pct"] = k.qt_ant * (k.v / k.qt - k.v_ant / k.qt_ant) * k.fator
    return k


def _le_anbima(f):
    linhas = f.read_bytes().decode("latin1").splitlines()
    i = next(j for j, l in enumerate(linhas) if l.startswith("Código@"))
    a = pd.read_csv(io.StringIO("\n".join(linhas[i:])), sep="@", dtype=str)
    a.columns = ["codigo", "nome", "venc", "indice", "taxa_compra", "taxa_venda", "taxa_ind", "desvio", "int_min",
                 "int_max", "pu", "pct_par", "duration_du", "pct_reune", "ref_ntnb"][:a.shape[1]]
    for c in ["taxa_ind", "pu", "pct_par", "duration_du"]:
        a[c] = pd.to_numeric(a[c].str.replace(",", ".", regex=False), errors="coerce")
    ind = a.indice.fillna("").str.upper()
    a["tipo"] = pd.Series("Outros", index=a.index).mask(ind.str.contains("IGP"), "IGP-M +").mask(
        ind.str.contains("IPCA"), "IPCA +").mask(ind.str.contains("%") & ind.str.contains("DI"), "% do DI").mask(
        ind.str.match(r"^DI\s*\+"), "DI +")
    a["duration_anos"] = a.duration_du / 252
    return a.assign(data=pd.Timestamp(f"20{f.stem[-6:-4]}-{f.stem[-4:-2]}-{f.stem[-2:]}"))


def anbima_debentures():
    """Marcacao ANBIMA gratuita de debentures (taxa indicativa, PU, duration). O site so guarda ~15 dias uteis:
    cada execucao salva os arquivos da janela, entao o historico local cresce a partir de agora.
    Devolve (ultimo dia, historico de todos os dias guardados)."""
    hoje = pd.Timestamp.today().normalize()
    for dia in pd.bdate_range(hoje - pd.Timedelta(days=25), hoje):
        f = CACHE / f"anbima_db{dia:%y%m%d}.txt"
        if f.exists():
            continue
        try:
            r = requests.get(f"https://www.anbima.com.br/informacoes/merc-sec-debentures/arqs/db{dia:%y%m%d}.txt",
                             headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
        except requests.RequestException:
            continue
        if r.status_code == 200 and "Código@".encode("latin1") in r.content:     # feriado/ainda nao publicado: pula
            f.write_bytes(r.content)
    fs = sorted(CACHE.glob("anbima_db*.txt"))
    if not fs:
        return pd.DataFrame(), pd.DataFrame()
    hist = pd.concat([_le_anbima(f) for f in fs], ignore_index=True)
    return hist[hist.data == hist.data.max()], hist


# ------------------------------------------------------------------ CRI/CRA: informe mensal das securitizadoras (CVM)
SECURIT = "https://dados.cvm.gov.br/dados/SECURIT/DOC/INF_MENSAL_{t}/DADOS/inf_mensal_{tl}_{a}.zip"


def securitizadoras(log=print, asof=None):
    """Series de CRI e CRA (informe mensal das securitizadoras, CVM): por serie e mes, valor por certificado (VNA) e o que foi
    pago por certificado no mes (juros + amortizacao). Conferido: (VNA + pago) / VNA anterior = CDI nas series 100% CDI
    (erro < 0,05 p.p. em 89% dos meses). Devolve (series, devedores/cedentes por certificado). Ultima versao de cada documento;
    `entrega` = data em que a CVM recebeu o documento (point in time). asof="AAAA-MM-DD": so o que a CVM tinha recebido ate a data.
    A CVM guarda so a ultima versao de cada informe (8% sao reapresentados, mediana 162 dias depois): o arquivo anterior vai para
    securit/versoes quando muda, para refazer o passado como era visto."""
    pasta = CACHE / "securit"
    pasta.mkdir(exist_ok=True)
    num = lambda x: pd.to_numeric(x.astype(str).str.replace(",", ".", regex=False), errors="coerce")
    ser, dev = [], []
    for t in ("CRA", "CRI"):
        for a in range(int(DESDE[:4]), date.today().year + 1):
            f = pasta / f"inf_mensal_{t.lower()}_{a}.zip"
            if not f.exists() or (a >= date.today().year - 1 and date.fromtimestamp(f.stat().st_mtime) < date.today()):
                try:
                    r = requests.get(SECURIT.format(t=t, tl=t.lower(), a=a), timeout=300)
                    if r.status_code == 200:
                        if f.exists() and f.stat().st_size != len(r.content):      # point in time: guarda a versao anterior
                            (pasta / "versoes").mkdir(exist_ok=True)
                            f.replace(pasta / "versoes" / f"{f.stem}_{date.fromtimestamp(f.stat().st_mtime)}.zip")
                        f.write_bytes(r.content)
                except requests.RequestException:
                    log(f"   securitizadoras: {f.name} indisponivel")
            if not f.exists():
                continue
            z = zipfile.ZipFile(f)
            def le_(chave):
                n = next((x for x in z.namelist() if chave in x), None)       # cedente_devedor nao existe em 2019 (CRI)
                return pd.read_csv(z.open(n), sep=";", encoding="latin1", dtype=str, on_bad_lines="skip") if n else pd.DataFrame(
                    columns=["Codigo_Identificacao_Certificado", "CNPJ", "Tipo"])
            c, g, d = le_("classe"), le_("geral"), le_("cedente_devedor")
            c = c.merge(g.reindex(columns=["Codigo_Identificacao_Certificado", "Data_Referencia", "Versao", "Data_Entrega", "CNPJ_Emissora",
                                           "Data_Emissao"]).drop_duplicates(),
                        on=["Codigo_Identificacao_Certificado", "Data_Referencia", "Versao"], how="left", suffixes=("", "_g"))
            ser.append(c.assign(tipo=t))
            dev.append(d[["Codigo_Identificacao_Certificado", "CNPJ", "Tipo"]])
    if not ser:
        return pd.DataFrame(), pd.DataFrame()
    s = pd.concat(ser, ignore_index=True)
    for k in ("Quantidade_Certificados", "Valor_Certificados", "Rendimentos", "Amortizacoes", "Versao"):
        s[k] = num(s[k])
    s["mes"] = s.Data_Referencia.str[:7]
    if asof:
        s = s[s.Data_Entrega.str[:10] <= str(asof)]
    s = s.sort_values("Versao").drop_duplicates(["Codigo_CETIP", "Classe", "mes"], keep="last")
    s = s[s.Quantidade_Certificados > 0]
    cn = lambda x: x.astype(str).str.replace(r"\D", "", regex=True)
    out = pd.DataFrame({"serie": s.Codigo_CETIP.str.strip().str.upper(), "isin": s.Codigo_ISIN.str.strip().str.upper(), "tipo": s.tipo,
                        "certificado": s.Codigo_Identificacao_Certificado, "classe": s.Classe, "mes": s.mes,
                        "venc": s.Data_Vencimento.str[:10], "taxa": s.Taxa_Juros.fillna(""), "cnpj_sec": cn(s.CNPJ_Emissora.fillna("")),
                        "V": s.Valor_Certificados, "Q": s.Quantidade_Certificados, "P": s.Rendimentos.fillna(0) + s.Amortizacoes.fillna(0),
                        "entrega": s.Data_Entrega.str[:10], "emissao": s.Data_Emissao.astype(str).str[:10]}).sort_values(["serie", "classe", "mes"])
    # cada securitizadora preenche de um jeito: "Valor_Certificados" ora e o total da serie, ora o valor POR certificado; os
    # pagamentos idem. Por serie, fica a leitura (4 combinacoes) em que o retorno ao par mes a mes e coerente (mediana perto de
    # 1% a.m. e pouca dispersao); serie sem leitura coerente fica sem VNA (nao usada).
    mi = out.mes.str[:4].astype(int) * 12 + out.mes.str[5:7].astype(int)
    seg = (out.serie == out.serie.shift()) & (out.classe == out.classe.shift()) & (mi - mi.shift() == 1)
    melhor, nota = None, None
    for vu in (False, True):
        for pu_ in (False, True):
            vna = out.V if vu else out.V / out.Q
            pago = out.P if pu_ else out.P / out.Q
            r = ((vna + pago) / vna.shift() - 1).where(seg)
            sc = r.sub(0.01).abs().groupby([out.serie, out.classe]).median()
            if melhor is None:
                melhor, nota = pd.DataFrame({"sc": sc}), {}
            melhor[(vu, pu_)] = sc
    esc = melhor.drop(columns="sc").idxmin(axis=1)
    sc_min = melhor.drop(columns="sc").min(axis=1)
    k = pd.MultiIndex.from_arrays([out.serie, out.classe])
    e = esc.reindex(k).values
    vu = np.array([x[0] if isinstance(x, tuple) else False for x in e])
    pu_ = np.array([x[1] if isinstance(x, tuple) else False for x in e])
    ok = sc_min.reindex(k).values < 0.01                                  # mediana a menos de 1 p.p. de 1% a.m.
    out["vna"] = np.where(ok, np.where(vu, out.V, out.V / out.Q), np.nan)
    out["pago"] = np.where(ok, np.where(pu_, out.P, out.P / out.Q), np.nan)
    # V, Q, P ficam: no motor a leitura e escolhida mes a mes pelo PU que o proprio fundo informa (ancora)
    d = pd.concat(dev, ignore_index=True).drop_duplicates()
    d = pd.DataFrame({"certificado": d.Codigo_Identificacao_Certificado, "cnpj": cn(d.CNPJ), "papel": d.Tipo})
    return out.reset_index(drop=True), d


# ------------------------------------------------------------------ SND (debentures.com.br), FIDC (CVM) e Treasury: cobertura total
SND = "https://www.debentures.com.br/exploreosnd/consultaadados/emissoesdedebentures/"
UA = {"User-Agent": "Mozilla/5.0"}


def _get(url, tent=4, **kw):
    for t in range(tent):
        try:
            r = requests.get(url, headers=UA, timeout=300, **kw)
            r.raise_for_status()
            return r.content
        except requests.RequestException:
            if t == tent - 1:
                raise
            time.sleep(3 * (t + 1))


def _tabela_snd(conteudo, inicio):
    t = conteudo.decode("latin1").splitlines()
    i = next((j for j, l in enumerate(t) if l.startswith(inicio)), None)
    if i is None:
        return pd.DataFrame()
    a = pd.read_csv(io.StringIO("\n".join(t[i:])), sep="\t", dtype=str, quoting=3)
    a.columns = [c.strip() for c in a.columns]
    return a.apply(lambda s: s.str.strip())


def _diario(f):                    # arquivo que muda: rebaixa 1x por dia
    return not f.exists() or date.fromtimestamp(f.stat().st_mtime) < date.today()


def snd_caracteristicas():
    """Cadastro de TODAS as debentures do SND (ativas e ja vencidas/excluidas): indice, taxa de emissao, vencimento,
    periodicidade de juros, amortizacao e se e incentivada (Lei 12.431)."""
    partes = []
    for exc in ("False", "True"):
        f = CACHE / f"snd_caracteristicas_{exc}.txt"
        if _diario(f):
            try:
                f.write_bytes(_get(SND + f"caracteristicas_e.asp?tip_deb=publicas&op_exc={exc}&ativo="))
            except requests.RequestException:
                pass
        if f.exists():
            partes.append(_tabela_snd(f.read_bytes(), "Codigo do Ativo"))
    a = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    if not len(a):
        return pd.DataFrame()
    dt = lambda c: pd.to_datetime(a[c], format="%d/%m/%Y", errors="coerce")
    meses = lambda c, u: pd.to_numeric(a[c], errors="coerce") * np.where(a[u].eq("DIA"), 1 / 30, 1)
    out = pd.DataFrame({"codigo": a["Codigo do Ativo"], "empresa": a["Empresa"], "cnpj": a["CNPJ"],
                        "venc": dt("Data de Saida / Novo Vencimento").fillna(dt("Data de Vencimento")),
                        "indice": a["indice"].map(_sem_acento),
                        "pct": pd.to_numeric(a["Percentual Multiplicador/Rentabilidade"], errors="coerce"),
                        "juros": pd.to_numeric(a["Juros Criterio Novo - Taxa"].str.replace(",", "."), errors="coerce"),
                        "cada_juros": meses("Juros Criterio Novo - Cada", "Juros Criterio Novo - Unidade"),
                        "cada_amort": meses("Amortizacao - Cada", "Amortizacao - Unidade"),
                        "carencia": dt("Amortizacao - Carencia"),
                        "incentivada": a["Deb. Incent. (Lei 12.431)"].eq("S")})
    return out.drop_duplicates("codigo").set_index("codigo")


def snd_pu(datas, workers=4, log=print):
    """PU par (curva de emissao, com juros acumulados e amortizacoes) de todas as debentures em cada data.
    Um arquivo pequeno por data, guardado: so baixa datas novas."""
    vazio = lambda d: CACHE / f"snd_pu_{d:%Y%m%d}_vazio"        # feriado: marcado para nao pedir de novo

    def um(d):
        if vazio(d).exists():
            return None
        partes = []
        for e in ("False", "True"):             # ativas + ja vencidas/excluidas
            f = CACHE / f"snd_pu_{d:%Y%m%d}_{e}.txt"
            if not f.exists():
                try:
                    b = _get(SND + f"puhistorico_e.asp?op_exc={e}&ativo=&dt_ini={d:%d/%m/%Y}&dt_fim={d:%d/%m/%Y}")
                except requests.RequestException:
                    return None
                if e == "False" and b.count(b"\n") < 20:     # feriado / sem dados
                    if d < pd.Timestamp.today().normalize() - pd.Timedelta(days=7):
                        vazio(d).touch()
                    return None
                f.write_bytes(b)
            partes.append(_tabela_snd(f.read_bytes(), "Data do PU"))
        a = pd.concat(partes, ignore_index=True)
        a = a[a.iloc[:, 0].str.match(r"\d\d/\d\d/\d{4}", na=False)]
        num = lambda s: pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
        return pd.DataFrame({"data": d, "codigo": a.iloc[:, 1].values, "pu_par": num(a.iloc[:, 5]).values})
    novas = [d for d in datas if not (CACHE / f"snd_pu_{d:%Y%m%d}_True.txt").exists() and not vazio(d).exists()]
    if novas:
        log(f"SND: PU par das debêntures em {len(novas)} fins de mês (só na 1a vez)...")
    with ThreadPoolExecutor(workers) as ex:
        out = [x for x in ex.map(um, datas) if x is not None]
    return pd.concat(out, ignore_index=True).drop_duplicates(["data", "codigo"]) if out else pd.DataFrame(columns=["data", "codigo", "pu_par"])


FIDC_URL = "https://dados.cvm.gov.br/dados/FIDC/DOC/INF_MENSAL/DADOS/"
_MEIO = [15, 45, 75, 105, 135, 165, 270, 540, 900, 1440]       # meio de cada faixa de prazo (dias) da tabela V


def _prazo_fidc(z, nome):
    t = pd.read_csv(z.open(nome), sep=";", encoding="latin1", dtype=str, quoting=3)
    c = "CNPJ_FUNDO_CLASSE" if "CNPJ_FUNDO_CLASSE" in t.columns else "CNPJ_FUNDO"
    b = t[[k for k in t.columns if re.match(r"TAB_V_A\d+_VL_PRAZO_VENC", k)]].apply(pd.to_numeric, errors="coerce").fillna(0).values
    tot = b.sum(1)
    return pd.DataFrame({"cnpj": t[c].values, "mes": t.DT_COMPTC.str[:7].values,
                         "anos": np.where(tot > 0, (b * _MEIO).sum(1) / np.where(tot > 0, tot, 1) / 365.25, np.nan)}).dropna()


def _series_fidc(z, n2, n3, n6):
    """Valor e quantidade de cotas (tab X_2), rentabilidade do mes (X_3), desempenho esperado (X_6) e amortizacoes
    pagas no mes (X_4) de cada classe/serie de cada FIDC."""
    le = lambda n: pd.read_csv(z.open(n), sep=";", encoding="latin1", dtype=str, quoting=3)
    a, b = le(n2), le(n3)
    n4 = n2.replace("_X_2_", "_X_4_")
    am = le(n4) if n4 in z.namelist() else pd.DataFrame(columns=list(a.columns[:4]) + ["TAB_X_TP_OPER", "TAB_X_CLASSE_SERIE", "TAB_X_VL_TOTAL"])
    e = le(n6) if n6 in z.namelist() else pd.DataFrame(columns=list(a.columns[:4]) + ["TAB_X_CLASSE_SERIE", "TAB_X_PR_DESEMP_ESPERADO"])
    c = "CNPJ_FUNDO_CLASSE" if "CNPJ_FUNDO_CLASSE" in a.columns else "CNPJ_FUNDO"
    e = e.rename(columns={"CNPJ_FUNDO": c}) if c not in e.columns else e
    am = am.rename(columns={"CNPJ_FUNDO": c}) if c not in am.columns else am
    am = am[am.TAB_X_TP_OPER.fillna("").str.contains("Amortiza")].copy()
    for t in (a, b, e, am):
        t["n"] = t.groupby([c, "DT_COMPTC", "TAB_X_CLASSE_SERIE"]).cumcount()
    k = [c, "DT_COMPTC", "TAB_X_CLASSE_SERIE", "n"]
    x = a.merge(b, on=k, how="outer").merge(e[k + ["TAB_X_PR_DESEMP_ESPERADO"]], on=k, how="left")         .merge(am[k + ["TAB_X_VL_TOTAL"]], on=k, how="left")
    num = lambda s: pd.to_numeric(s.str.replace(",", ".", regex=False), errors="coerce")
    return pd.DataFrame({"cnpj": x[c].values, "mes": x.DT_COMPTC.str[:7].values, "serie": x.TAB_X_CLASSE_SERIE.str.strip().values,
                         "cota": num(x.TAB_X_VL_COTA).values, "qt": num(x.TAB_X_QT_COTA).values, "rent": num(x.TAB_X_VL_RENTAB_MES).values,
                         "esperado": num(x.TAB_X_PR_DESEMP_ESPERADO).values, "amort": num(x.TAB_X_VL_TOTAL).fillna(0).values})


def fidc_mensal(desde, log=print):
    """Do informe mensal de FIDC da CVM: (1) prazo medio (anos) dos recebiveis a vencer de cada FIDC (tabela V) =
    duration estimada da cota; (2) valor da cota e rentabilidade de cada classe/serie (tabelas X_2/X_3) = spread da
    serie que o fundo tem. Guarda so o resultado (pequeno), nao os zips."""
    f, fs = CACHE / "fidc_prazo4.parquet", CACHE / "fidc_series4.parquet"
    tem = pd.read_parquet(f) if f.exists() else pd.DataFrame(columns=["cnpj", "mes", "anos"])
    ser = [pd.read_parquet(fs)] if fs.exists() else []
    feitos = set(tem.mes)
    hoje = pd.Timestamp.today()
    novos = []
    for a in range(int(desde[:4]), hoje.year + 1):
        ms = [f"{a}-{m:02d}" for m in range(1, 13) if pd.Timestamp(a, m, 1) <= hoje and f"{a}-{m:02d}" >= desde]
        falta = [m for m in ms if m not in feitos]
        if not falta:
            continue
        urls = [FIDC_URL + f"inf_mensal_fidc_{m.replace('-', '')}.zip" for m in falta] if a >= 2025 else [FIDC_URL + f"HIST/inf_mensal_fidc_{a}.zip"]
        for u in urls:
            try:
                z = zipfile.ZipFile(io.BytesIO(_get(u, tent=2)))
            except (requests.RequestException, zipfile.BadZipFile):
                continue
            log(f"FIDC: prazo e rentabilidade das séries ({u.rsplit('/', 1)[1]})...")
            nomes = z.namelist()
            for n in nomes:
                if "_tab_V_" in n:
                    novos.append(_prazo_fidc(z, n))
                if "_tab_X_2_" in n and n.replace("_X_2_", "_X_3_") in nomes:
                    ser.append(_series_fidc(z, n, n.replace("_X_2_", "_X_3_"), n.replace("_X_2_", "_X_6_")))
    if novos:
        tem = pd.concat([tem] + novos, ignore_index=True).drop_duplicates(["cnpj", "mes"], keep="last")
        tem.to_parquet(f, index=False)
        pd.concat(ser, ignore_index=True).drop_duplicates(["cnpj", "mes", "serie", "cota"], keep="last").to_parquet(fs, index=False)
    return tem, (pd.concat(ser, ignore_index=True) if ser else pd.DataFrame(columns=["cnpj", "mes", "serie", "cota", "qt", "rent", "esperado", "amort"]))


def taxas():
    """Taxa de administracao e de performance (com o benchmark) de cada fundo: extrato anual da CVM, ultimo informado."""
    partes = []
    for a in range(max(int(DESDE[:4]), 2019), date.today().year + 1):
        f = CACHE / f"extrato_fi_{a}.csv"
        if not f.exists() or (a == date.today().year and _diario(f)):
            try:
                f.write_bytes(_get(f"{URL}/EXTRATO/DADOS/extrato_fi_{a}.csv", tent=2))
            except requests.RequestException:
                pass
        if f.exists():
            t = pd.read_csv(f, sep=";", encoding="latin1", dtype=str, quoting=3)
            t = t.rename(columns={"CNPJ_FUNDO": "CNPJ_FUNDO_CLASSE"})
            partes.append(t[[c for c in ["CNPJ_FUNDO_CLASSE", "DT_COMPTC", "TAXA_ADM", "TAXA_PERFM", "PARAM_TAXA_PERFM",
                                         "PR_INDICE_REFER_TAXA_PERFM"] if c in t.columns]])
    if not partes:
        return {}
    t = pd.concat(partes, ignore_index=True).sort_values("DT_COMPTC").drop_duplicates("CNPJ_FUNDO_CLASSE", keep="last")
    num = lambda c: pd.to_numeric(t[c].str.replace(",", ".", regex=False), errors="coerce") if c in t.columns else np.nan
    t = t.assign(adm=num("TAXA_ADM"), perf=num("TAXA_PERFM"), pct=num("PR_INDICE_REFER_TAXA_PERFM"))
    return {c: (a if a == a else 0.0, p if p == p else 0.0, str(par or ""), q if q == q and q > 0 else 100.0)   # 0 = nao informado
            for c, a, p, par, q in zip(t.CNPJ_FUNDO_CLASSE, t.adm, t.perf, t.get("PARAM_TAXA_PERFM", ""), t.pct)}


def treasury():
    """Curva de juros do Tesouro americano (treasury.gov, gratuito), no formato de tesouro() para usar em curva()."""
    tenor = {"1 Mo": 1 / 12, "3 Mo": .25, "6 Mo": .5, "1 Yr": 1, "2 Yr": 2, "3 Yr": 3, "5 Yr": 5, "7 Yr": 7, "10 Yr": 10, "20 Yr": 20, "30 Yr": 30}
    partes = []
    for a in range(int(DESDE[:4]), date.today().year + 1):
        f = CACHE / f"treasury_{a}.csv"
        if not f.exists() or (a == date.today().year and _diario(f)):
            try:
                f.write_bytes(_get("https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/"
                                   f"{a}/all?type=daily_treasury_yield_curve&field_tdr_date_value={a}&page&_format=csv", tent=2))
            except requests.RequestException:
                continue
        if f.exists():
            partes.append(pd.read_csv(f))
    if not partes:
        return {}
    d = pd.concat(partes, ignore_index=True)
    d["data"] = pd.to_datetime(d.Date, format="%m/%d/%Y")
    d = d.drop_duplicates("data").sort_values("data")
    cols = [c for c in tenor if c in d.columns]
    x = np.array([tenor[c] for c in cols])
    por_dia = {r.data: (x[~np.isnan(v)], v[~np.isnan(v)]) for r, v in zip(d.itertuples(), d[cols].to_numpy(float)) if (~np.isnan(v)).sum() > 2}
    datas = sorted(por_dia)
    return {"UST": (np.array(datas, dtype="datetime64[ns]"), datas, por_dia)}


def sgs(codigo, desde):
    """Serie do Banco Central (SGS). Diarias: no maximo 10 anos por consulta."""
    out, hoje = [], pd.Timestamp.today()
    for a in range(int(desde[:4]), hoje.year + 1, 9):
        fim = min(pd.Timestamp(a + 8, 12, 31), hoje).strftime("%d/%m/%Y")
        for tent in range(5):
            try:
                out += requests.get(f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados", timeout=60,
                                    params={"formato": "json", "dataInicial": f"01/01/{a}", "dataFinal": fim}).json()
                break
            except (requests.RequestException, ValueError):
                time.sleep(5 * (tent + 1))
    d = pd.DataFrame(out)
    return pd.Series(d.valor.astype(float).values, index=pd.to_datetime(d.data, dayfirst=True)).sort_index() if len(d) else pd.Series(dtype=float)


def cdi_mensal(desde, diario=False):
    s = sgs(12, desde) / 100
    return s if diario else (1 + s).groupby(s.index.strftime("%Y-%m")).prod() - 1


def ipca_indice():
    """Numero-indice do IPCA, mensal (indice "AAAA-MM"). Fonte: IBGE/SIDRA; se vier vazio ou falhar, monta pela
    variacao mensal do Banco Central (SGS 433) ancorada no numero-indice oficial de dez/2019 (5320,25)."""
    f = CACHE / "ipca_indice.json"
    if not f.exists() or f.stat().st_mtime < time.time() - 86400:
        try:
            r = requests.get("https://apisidra.ibge.gov.br/values/t/1737/n1/all/v/2266/p/all", timeout=120).json()
            d = {x["D3C"]: x["V"] for x in r[1:]}
            if len(d) > 12:
                f.write_text(json.dumps(d))
        except (requests.RequestException, ValueError, KeyError, TypeError, IndexError):
            pass
    s = pd.Series(dtype=float, index=pd.Index([], dtype=object))
    if f.exists():
        d = json.loads(f.read_text())
        s = pd.Series({f"{k[:4]}-{k[4:]}": float(v) for k, v in d.items() if v not in ("...", "-", "")}, dtype=float)
    if len(s) < 12:
        v = sgs(433, "1994-01") / 100
        if len(v):
            v.index = v.index.strftime("%Y-%m")
            nivel = (1 + v).cumprod()
            s = nivel / nivel.get("2019-12", nivel.iloc[-1]) * 5320.25
    s.index = s.index.astype(object)
    return s.sort_index()


def tesouro():
    """Curvas do Tesouro Direto (gratuito, diario desde 2004): prefixado (LTN) e IPCA+ (NTN-B principal), taxa x prazo."""
    f = CACHE / "tesouro_direto.csv"
    if not f.exists() or f.stat().st_mtime < time.time() - 86400:
        try:
            r = requests.get("https://www.tesourotransparente.gov.br/ckan/dataset/df56aa42-484a-4a59-8184-7676580c81e3/"
                             "resource/796d2059-14e9-44e3-80c9-2d9e30b405c1/download/PrecoTaxaTesouroDireto.csv", timeout=300)
            r.raise_for_status()
            f.write_bytes(r.content)
        except requests.RequestException:
            if not f.exists():
                return {}
    d = pd.read_csv(f, sep=";", decimal=",")
    d = d[d["Tipo Titulo"].isin(["Tesouro Prefixado", "Tesouro IPCA+"])].copy()
    d["data"] = pd.to_datetime(d["Data Base"], dayfirst=True)
    d["anos"] = (pd.to_datetime(d["Data Vencimento"], dayfirst=True) - d.data).dt.days / 365.25
    d["taxa"] = d[["Taxa Compra Manha", "Taxa Venda Manha"]].replace(0, np.nan).mean(axis=1)
    d["tipo"] = np.where(d["Tipo Titulo"].eq("Tesouro Prefixado"), "PRE", "IPCA")
    d = d.dropna(subset=["taxa"])
    d = d[d.anos >= 0.5].sort_values(["tipo", "data", "anos"])      # titulo vencendo em meses distorce a taxa (projecao do IPCA)
    out = {}
    for t, g in d.groupby("tipo"):
        por_dia = {k: (x.anos.values, x.taxa.values) for k, x in g.groupby("data")}
        datas = sorted(por_dia)
        out[t] = (np.array(datas, dtype="datetime64[ns]"), datas, por_dia)
    return out


def curva(curvas, tipo, dia, anos):
    """Taxa (% a.a.) da curva `tipo` ("PRE" ou "IPCA") no ultimo dia util <= `dia`, interpolada no prazo `anos`."""
    if tipo not in curvas:
        return np.full(np.shape(anos), np.nan)
    arr, datas, por_dia = curvas[tipo]
    i = np.searchsorted(arr, np.datetime64(pd.Timestamp(dia)), side="right") - 1
    if i < 0:
        return np.full(np.shape(anos), np.nan)
    x, y = por_dia[datas[i]]
    return np.interp(np.clip(anos, x.min(), x.max()), x, y)


def _sem_acento(t):
    import unicodedata
    return unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode().upper()


_GRUPOS_RE = [(re.compile(k), v) for k, v in GRUPOS.items()]
_SUFIXOS = re.compile(r"(?<![\w.])(S\.A\.?|S/A|SA|LTDA|SPE|PARTICIPACOES|HOLDING|COMPANHIA|CIA|EMPREENDIMENTOS|E|DE|DO|DA|DOS|DAS|EM)(?![\w/])")


def grupo(emissor, categoria=""):
    """Grupo economico pelo nome do emissor (mapa GRUPOS); sem mapa, o proprio emissor sem sufixos."""
    if categoria in ("Operações Compromissadas",):
        return "Compromissadas"
    if categoria == "Títulos Públicos":
        return "Tesouro Nacional"
    if categoria in ("Disponibilidades", "Valores a receber", "Valores a pagar", AJUSTE):
        return "Caixa e outros"
    if not isinstance(emissor, str) or not emissor.strip():
        return None
    n = _sem_acento(emissor)
    for rx, g in _GRUPOS_RE:
        if rx.search(n):
            return g
    base = " ".join(w for w in _SUFIXOS.sub(" ", n).split() if not w.isdigit())
    return base.title()[:40] or emissor


def diario(nomes, dia):
    nome = {c: n for n, c in nomes.items()}
    d = pd.concat([pd.read_parquet(CACHE / "rec" / f"diario_{m}.parquet") for ms in dia.values() for m in ms
                   if (CACHE / "rec" / f"diario_{m}.parquet").exists()], ignore_index=True)
    d = d[d.CNPJ_FUNDO_CLASSE.isin(nome)]
    for c in ["VL_QUOTA", "VL_PATRIM_LIQ", "CAPTC_DIA", "RESG_DIA", "NR_COTST"]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d = d.drop_duplicates(["CNPJ_FUNDO_CLASSE", "DT_COMPTC"], keep="last").sort_values("DT_COMPTC")
    return d.assign(data=pd.to_datetime(d.DT_COMPTC), mes=d.DT_COMPTC.str[:7], fundo=d.CNPJ_FUNDO_CLASSE.map(nome))


def distribuicoes(d):
    """Fundos que distribuem rendimentos (ex.: infra "renda"): a cota cai todo mes no dia da distribuicao e nao mostra
    o retorno total. Detecta queda > 0,4% (bem abaixo do normal do fundo) no inicio do mes, em >= 60% dos meses, e
    reconstroi a cota de retorno total somando a distribuicao de volta. Devolve (d ajustado, {fundo: (n, media %)})."""
    partes, info = [], {}
    for n, g in d.groupby("fundo", sort=False):
        g = g.sort_values("data").copy()
        q = g.VL_QUOTA.values.astype(float)
        if len(q) < 60:
            partes.append(g); continue
        r = np.r_[0.0, q[1:] / q[:-1] - 1]
        base = pd.Series(r).rolling(21, min_periods=5, center=True).median().values
        cand = (r < -0.004) & (r < base - 0.004) & (g.data.dt.day.values <= 12)
        meses_cand = sorted(set(g.mes[cand].tolist()))
        if len(meses_cand) >= 6:
            ini, fim = pd.Period(meses_cand[0]), pd.Period(meses_cand[-1])
            if len(meses_cand) >= 0.6 * ((fim - ini).n + 1):
                info[n] = (int(cand.sum()), float(-np.mean(r[cand]) * 100))
                r = np.where(cand, base, r)
                g["VL_QUOTA"] = q[0] * np.cumprod(1 + np.nan_to_num(r))
        partes.append(g)
    return (pd.concat(partes, ignore_index=True) if partes else d), info


def rentabilidade(d, cdi):
    g = d.groupby(["fundo", "mes"])
    rent = pd.DataFrame({"cota_fim": g.VL_QUOTA.last(), "pl_fim": g.VL_PATRIM_LIQ.last(),
                         "captacao_liquida": g.CAPTC_DIA.sum() - g.RESG_DIA.sum(), "cotistas": g.NR_COTST.last()})
    rent["retorno"] = rent.groupby(level=0).cota_fim.pct_change()
    rent["cdi"] = rent.index.get_level_values(1).map(cdi)
    rent["pct_cdi"] = rent.retorno / rent.cdi
    return rent.reset_index()


def excel(df, rent, ordem, out):
    from openpyxl import Workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.chart import AreaChart, Reference
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils.indexed_list import IndexedList
    limpa = lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v
    df = df.copy()
    for c in df.select_dtypes("object"):
        df[c] = df[c].map(limpa)
    B, N = Font(name="Arial", size=10, bold=True), Font(name="Arial", size=10)
    HF = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    FILL, MFILL = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="D9E1F2")
    PCT, PCT2 = '0.0000"%";(0.0000"%")', '0.00"%";(0.00"%")'

    def put(ws, vals, font=None, fill=None, fmt=None):
        ws.append(vals)
        r = ws._current_row                       # O(1); ws.max_row recontaria a planilha toda
        for j in range(1, len(vals) + 1):
            c = ws.cell(r, j)
            if font:
                c.font = font
            if fill:
                c.fill = fill
            if fmt and j in fmt:
                c.number_format = fmt[j]
        return r

    def header(ws, cols, widths):
        put(ws, cols, HF, FILL)
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w

    wb = Workbook()
    wb._fonts = IndexedList([N])                  # Arial 10 como fonte padrao
    wb.remove(wb.active)
    for nm, g in df.groupby("fundo", sort=False):
        ws = wb.create_sheet(nm[:31])
        header(ws, ["Categoria / Ativo", "Codigo", "% PL " + nm, "% da categoria", "Veiculo", "Emissor"], (70, 18, 14, 14, 50, 40))
        ws.sheet_properties.outlinePr.summaryBelow = False
        ws.freeze_panes = "A2"
        for ym, gm in sorted(g.groupby("mes"), reverse=True):
            put(ws, [ym, None, gm.perc_pl.sum()], B, MFILL, {3: PCT})
            ats = gm.groupby(["categoria", "ativo", "codigo"], dropna=False).agg(
                perc_pl=("perc_pl", "sum"), veiculo=("veiculo", "first"), emissor=("emissor", "first"))
            for cat, tot in gm.groupby("categoria").perc_pl.sum().sort_values(ascending=False).items():
                r = put(ws, ["   " + cat, None, tot], B, None, {3: PCT})
                ws.row_dimensions[r].outline_level = 1
                for (_, ativo, cod), a in ats.loc[[cat]].sort_values("perc_pl", ascending=False).iterrows():
                    last = put(ws, ["      " + ativo, None if pd.isna(cod) else cod, a.perc_pl,
                                    a.perc_pl / tot * 100 if tot else None, a.veiculo, a.emissor], fmt={3: PCT, 4: PCT})
                ws.row_dimensions.group(r + 1, last, outline_level=2, hidden=True)

    CURTO = {"Depósitos a prazo e outros títulos de IF": "Depósitos a prazo / IF",
             "Outros valores mobiliários registrados na CVM objeto de oferta pública": "Outros VM (CVM)",
             "Títulos ligados ao agronegócio": "Títulos agro", "Títulos de Crédito Privado": "Crédito privado (CCB/NC)"}
    CORES, OUTROS_COR = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7"], "898781"
    cat_ = df.categoria.map(lambda c: CURTO.get(c, c))
    CONTABIL = r"(?i)pagar|receber|obriga|termo|disponibilidade|exigibilidade|swap|confidencial|ajuste"
    TOP = (df.assign(c=cat_)[~cat_.str.contains(CONTABIL) & ~df.derivativo].groupby(["fundo", "mes", "c"]).perc_pl.sum()
           .groupby("c").mean().sort_values(ascending=False).index[:7].tolist())
    SERIES = TOP + ["Outros"]
    ev = df[~df.derivativo].assign(c=cat_.where(cat_.isin(TOP), "Outros")).pivot_table(
        index=["fundo", "mes"], columns="c", values="perc_pl", aggfunc="sum", fill_value=0).reindex(columns=SERIES, fill_value=0)
    ws, gr = wb.create_sheet("evolucao", 0), wb.create_sheet("graficos", 0)
    put(gr, ["Composicao por categoria (% do PL, look-through, dados CVM). Numeros em 'evolucao'."], B)
    for k, nm in enumerate(n for n in ordem if n in ev.index.get_level_values(0)):
        row0 = put(ws, [nm], B)
        put(ws, ["mes"] + SERIES, HF, FILL)
        for ym, r in ev.loc[nm].iterrows():
            last = put(ws, [ym] + [float(v) for v in r], fmt={j: PCT2 for j in range(2, len(SERIES) + 2)})
        ws.append([])
        ws.append([])
        ch = AreaChart()
        ch.grouping, ch.title, ch.y_axis.title = "stacked", f"{nm} — % do PL por categoria", "% do PL"
        ch.y_axis.scaling.min, ch.y_axis.scaling.max, ch.y_axis.majorGridlines = 0, 100, None
        ch.x_axis.tickLblSkip, ch.x_axis.delete, ch.y_axis.delete = 12, False, False
        ch.add_data(Reference(ws, min_col=2, max_col=1 + len(SERIES), min_row=row0 + 1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=row0 + 2, max_row=last))
        for s, cor in zip(ch.series, CORES + [OUTROS_COR]):
            s.graphicalProperties.solidFill, s.graphicalProperties.line.solidFill = cor, "FCFCFB"
        ch.legend.position, ch.width, ch.height = "r", 26, 9
        gr.add_chart(ch, f"A{3 + k * 19}")
    ws.column_dimensions["A"].width = 12

    ws = wb.create_sheet("rentabilidade", 2)
    header(ws, ["fundo", "mes", "retorno", "CDI", "% CDI", "cota fim", "PL fim", "captacao liquida", "cotistas"],
           (16, 9, 11, 11, 11, 14, 18, 18, 10))
    for r in rent.itertuples(index=False):
        put(ws, [r.fundo, r.mes, r.retorno, r.cdi, r.pct_cdi, r.cota_fim, r.pl_fim, r.captacao_liquida, r.cotistas],
            fmt={3: "0.00%", 4: "0.00%", 5: "0.0%", 7: "#,##0", 8: "#,##0;(#,##0)"})

    ult = df[df.mes == df.fundo.map(df.groupby("fundo").mes.max())]
    top10 = (ult.groupby(["categoria", "fundo", "mes", "ativo", "codigo"], dropna=False).perc_pl.sum().reset_index()
             .sort_values(["categoria", "perc_pl"], ascending=[True, False]).groupby("categoria").head(10))
    top10.insert(1, "rank", top10.groupby("categoria").cumcount() + 1)
    ws = wb.create_sheet("top10", 3)
    header(ws, ["categoria", "rank", "fundo", "mes", "ativo", "codigo", "% PL do fundo"], (34, 6, 16, 9, 70, 18, 14))
    for r in top10.itertuples(index=False):
        put(ws, [None if pd.isna(x) else x for x in r], fmt={7: PCT})
    wb.save(out)
    df.drop(columns=["chave", "fator"]).to_csv(out.replace(".xlsx", "_dados.csv"), index=False, sep=";", decimal=",",
                                               encoding="utf-8-sig")


lt = sys.modules[__name__]          # o servidor chama as funcoes de dados como lt.funcao

# ====================================================================== servidor + pagina

import json
import io
import math
import os
import re
import threading
from contextlib import asynccontextmanager
from datetime import date
from pathlib import Path

import numpy as np
import requests
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response


@asynccontextmanager
async def _vida(_app):                       # partida do servidor (lifespan: vale para FastAPI/Starlette novos e antigos)
    _inicio()
    yield


app = FastAPI(title="Painel de fundos", lifespan=_vida)

CAT_CORES = ["#2a78d6", "#1baf7a", "#eda100", "#e87ba4", "#4a3aa7", "#008300", "#e34948", "#7ba7d9"]
NOSSAS_CORES = ["#eb6834", "#c24f1d", "#f29a6b", "#a8401a", "#f5b38d"]
GENERICO = (r"(?i)\b(FUNDO DE INVESTIMENTO|FUNDO|EM DIREITOS CREDIT[OÓ]RIOS|DE INVESTIMENTO|FIDC|FIC|FI|COTAS|MULTIMERCADO|"
            r"RENDA FIXA|CR[EÉ]DITO PRIVADO|RESPONSABILIDADE LIMITADA|RESP LTDA|N[AÃ]O PADRONIZADO|NP|LONGO PRAZO|LP)\b")
METODO = 30                                     # suba quando mudar o calculo -> refaz o cache do dia
CAIXA = ("Caixa",)
CAIXA_CVM = ("Disponibilidades", "Operações Compromissadas", "Valores a receber", "Valores a pagar", lt.AJUSTE)
TIPO_CASA = {"CDB/ RDB": "CDB", "CDB Vinculado": "CDB", "DPGE": "DPGE", "FI Imobiliário": "FII", "FI Participações": "FIP",
             "Fundos de Índice": "ETF", "Fundos Offshore": "Fundos offshore", "CCB": "CCB", "CCCB": "CCB", "LCA": "LCA", "CPR": "CPR",
             "CDCA": "CDCA", "CCI": "CCI", "NCA": "NCA", "Letra de Câmbio/ Letra Hipotecária/ Letra Imobiliária": "LCI/LH"}
CURTO_CVM = {"Depósitos a prazo e outros títulos de IF": "Outros títulos de IF", "Títulos ligados ao agronegócio": "Agro (outros)",
             "Outros valores mobiliários registrados na CVM objeto de oferta pública": "Outros valores mobiliários",
             "Investimento no Exterior": "Exterior (outros)", "Títulos de Crédito Privado": "Crédito privado (outros)",
             "Cotas de Fundos": "Cotas de fundos"}
SEM_SPREAD = ("Ações", "FII", "FIP", "ETF", "Confidencial")
MERCADO = ("ANBIMA", "SND + preço do fundo")            # renda variavel: spread e duration nao se aplicam
FUNDOS = ("FIDC", "FII", "FIP", "ETF", "FIAGRO", "Cotas de fundos", "Fundos offshore")
USD = ("Bonds", "Fundos offshore", "Exterior (outros)")
TP = [("LFT", r"BRSTNCLF|FINANCEIRAS DO TESOURO|\bLFT\b"), ("NTN-B", r"BRSTNCNTB|SERIE B|NTN-B"),
      ("NTN-F", r"BRSTNCNTF|SERIE F|NTN-F"), ("LTN", r"BRSTNCLTN|LETRAS DO TESOURO NACIONAL|\bLTN\b")]
MESES_COD = "FGHJKMNQUVXZ"
CAPTURA = None                                  # auditoria: dict que recebe as linhas calculadas (None = desligado)
PRE = {"status": "parado", "log": [], "erro": None, "pasta": None}


# ------------------------------------------------------------------ utilidades
def limpo(o):
    if isinstance(o, dict):
        return {str(k): limpo(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [limpo(v) for v in o]
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, (np.floating, float)):
        return None if (math.isnan(o) or math.isinf(o)) else round(float(o), 6)
    if isinstance(o, pd.Timestamp):
        return None if pd.isna(o) else o.strftime("%Y-%m-%d")
    if o is None or isinstance(o, (str, bool)):
        return o
    try:
        return None if pd.isna(o) else o
    except (TypeError, ValueError):
        return o


def _txt(x):
    """Texto como array numpy de objetos str (funciona com pandas 2 e 3/Arrow)."""
    return np.array(["" if v is None or (isinstance(v, float) and v != v) else str(v) for v in x], dtype=object)


def _tem(arr, sub):
    return np.array([sub in v for v in arr], dtype=bool)


def wavg(v, w):
    v, w = np.asarray(v, float), np.asarray(w, float)
    ok = ~np.isnan(v) & (w > 0)
    return float(np.average(v[ok], weights=w[ok])) if ok.any() else None


def fim_mes(m):
    return pd.Timestamp(m) + pd.offsets.MonthEnd(0)


def contrato(texto):
    """Tipo, vencimento (ano, mes) de um futuro a partir do codigo/descricao da CDA (ex.: DI1FUTF29, FUT DAP/Q30)."""
    t = lt._sem_acento(texto)
    tipo = ("DAP" if re.search(r"DAP|CUPOM DE IPCA|DI X IPCA", t) else "DI1" if re.search(r"DI1|DI DE 1 DIA", t)
            else "WDO" if "WDO" in t else "DOL" if re.search(r"DOL|DOLAR", t) else "DDI" if re.search(r"DDI|FRC|CUPOM CAMBIAL", t)
            else "IND" if re.search(r"\bIND|WIN|IBOV", t) else "OUTRO")
    m = re.search(r"(?:DI1|DAP|DOL|WDO|DDI|FRC|IND|WIN)\s*(?:FUT)?\s*/?\s*([FGHJKMNQUVXZ])(\d{2})\b", t)
    return tipo, ((2000 + int(m.group(2)), MESES_COD.index(m.group(1)) + 1) if m else None)


# ------------------------------------------------------------------ fluxo de caixa: duration e taxa implicita pelo preco
def _grade(T, p_jur, p_amo, t_car):
    """Fluxos por unidade de principal numa grade mensal contada a partir do vencimento (vetorizado).
    T: anos ate o vencimento; juros a cada p_jur anos; amortizacao constante a cada p_amo anos (0 = no vencimento)
    a partir de t_car anos. Devolve (prazos, peso_juros, principal, periodo_juros)."""
    T = np.clip(np.nan_to_num(np.atleast_1d(np.asarray(T, float)), nan=1.0), 1 / 12, 40)
    col = lambda v, pad: np.broadcast_to(np.nan_to_num(np.asarray(v, float), nan=pad), T.shape)[:, None]
    j = np.arange(int(np.ceil(T.max() * 12)) + 1)[None, :]
    tau = T[:, None] - j / 12
    ok = tau > 1e-6
    kj = np.clip(np.round(col(p_jur, 0.5) * 12), 1, 600)
    ka = np.round(col(p_amo, 0) * 12)
    am = np.where(ka > 0, (j % np.maximum(ka, 1) == 0) & (tau >= col(t_car, 0) - 1e-6), j == 0) & ok
    am |= (j == 0) & ~am.any(1, keepdims=True)
    princ = am / am.sum(1, keepdims=True)
    saldo = np.cumsum(princ, 1)                       # saldo devedor antes do pagamento em tau
    return np.where(ok, tau, 1.0), ((j % kj == 0) & ok) * saldo, princ, kj / 12


def pv(tau, wj, princ, per, y_cup, y):
    """Valor presente por unidade de principal (e soma de prazo x VP): juros a taxa do titulo y_cup, desconto a y (% a.a.)."""
    fl = wj * ((1 + np.asarray(y_cup, float)[:, None] / 100) ** per - 1) + princ
    df = (1 + np.asarray(y, float)[:, None] / 100) ** -tau
    return (fl * df).sum(1), (fl * df * tau).sum(1)


def taxa_e_duration(T, p_jur, p_amo, t_car, base, s0, pct_par):
    """Taxa de mercado s (mesma unidade de s0) que leva o preco a pct_par do PU par, e a duration (anos) nessa taxa.
    y = (1+base)(1+s)-1: base = CDI para DI+, 0 para IPCA+/pre (s ja e a taxa). pct_par NaN -> fica s0."""
    tau, wj, princ, per = _grade(T, p_jur, p_amo, t_car)
    n = len(tau)
    base, s0 = np.broadcast_to(np.asarray(base, float), (n,)), np.broadcast_to(np.asarray(s0, float), (n,)).copy()
    alvo = np.broadcast_to(np.asarray(pct_par, float), (n,))
    y = lambda s, b: ((1 + b / 100) * (1 + s / 100) - 1) * 100
    s = s0.copy()
    r = np.where(np.isfinite(alvo) & np.isfinite(s0))[0]
    if len(r):                                       # bissecao: o preco cai quando a taxa sobe
        a = (tau[r], wj[r], princ[r], per[r])
        v0, _ = pv(*a, y(s0[r], base[r]), y(s0[r], base[r]))
        lo, hi = np.full(len(r), -10.0), np.full(len(r), 80.0)
        for _ in range(30):
            mid = (lo + hi) / 2
            v, _ = pv(*a, y(s0[r], base[r]), y(mid, base[r]))
            caro = v / v0 > alvo[r]
            lo, hi = np.where(caro, mid, lo), np.where(caro, hi, mid)
        s[r] = (lo + hi) / 2
    v, vt = pv(tau, wj, princ, per, y(np.nan_to_num(s0), base), y(np.nan_to_num(s), base))
    return s, vt / v


# ------------------------------------------------------------------ calculo por fundo
class Contexto:
    """Dados de mercado compartilhados por todos os fundos."""

    def __init__(self, log, meses):
        log("Baixando curvas (Tesouro Direto), IPCA (IBGE), PTAX e CDI (Banco Central) e marcação ANBIMA...")
        self.curvas = lt.tesouro()
        self.ipca = lt.ipca_indice()
        self.ptax = lt.sgs(1, lt.DESDE)
        self.cdi_d = lt.cdi_mensal(lt.DESDE, diario=True)
        self.cdi_m = (1 + self.cdi_d).groupby(self.cdi_d.index.strftime("%Y-%m")).prod() - 1
        self.an_u, self.an_h = lt.anbima_debentures()
        self.an_dias = sorted(self.an_h.data.unique()) if len(self.an_h) else []
        log("Cadastro e PU par das debêntures (SND), prazo dos FIDCs (CVM) e Treasury (EUA)...")
        self.snd = lt.snd_caracteristicas()
        hoje = pd.Timestamp.today().normalize() - pd.offsets.BDay(1)
        alvo = {m: min(pd.Timestamp(m) + pd.offsets.BMonthEnd(0), hoje) for m in meses}
        pu = lt.snd_pu(sorted(set(alvo.values())), log=log)
        falta = {m: d - pd.offsets.BDay(1) for m, d in alvo.items() if d not in set(pu.data)}      # feriado no fim do mes
        if falta:
            pu = pd.concat([pu, lt.snd_pu(sorted(set(falta.values())), log=log)], ignore_index=True)
            alvo.update(falta)
        self.pu_dia = alvo
        self.pu = {d: x.drop_duplicates("codigo").set_index("codigo").pu_par for d, x in pu.groupby("data")}
        log("Informe mensal das securitizadoras (CRI/CRA, CVM)...")
        self.sec, self.sec_dev = lt.securitizadoras(log)
        if len(self.sec):
            self.sec_m = self.sec.drop_duplicates(["serie", "mes"]).set_index(["serie", "mes"])[["V", "Q", "P", "vna", "pago"]]
            self.sec_isin = self.sec[self.sec["isin"].str.len() == 12].drop_duplicates("isin").set_index("isin").serie
            ids = set(self.sec.serie)
            por = self.sec.drop_duplicates(["serie", "certificado"])[["serie", "certificado", "venc", "taxa", "classe", "cnpj_sec", "emissao"]]
            self.sec_prefixo = self.sec.serie.drop_duplicates()
            ligados = pd.concat([por.assign(cnpj=por.cnpj_sec), por.merge(self.sec_dev, on="certificado")[por.columns.tolist() + ["cnpj"]]])
            self.sec_chave = ligados.drop_duplicates(["serie", "cnpj"])
        f, ser = lt.fidc_mensal(lt.DESDE, log)
        self.fidc = {c: (x.mes.values, x.anos.values) for c, x in f.sort_values("mes").groupby("cnpj")}
        ser = ser[(ser.cota > 0) & ser.rent.notna()].sort_values(["cnpj", "serie", "mes"])
        # retorno total da serie no mes = (cota + amortizacao paga por cota) / cota do mes anterior - 1  (tabs X_2 e X_4)
        g = ser.groupby(["cnpj", "serie"])
        ant = g.cota.shift()
        seguido = pd.to_datetime(ser.mes) == (pd.to_datetime(g.mes.shift()) + pd.DateOffset(months=1))
        am_pc = (ser.amort / ser.qt).where(ser.qt > 0, 0).fillna(0)
        rt = ((ser.cota + am_pc) / ant - 1).where(seguido & (ant > 0), ser.rent / 100)
        oficial = (ser.rent / 100).where(ser.rent != 0)   # oficial (X_3) tem prioridade; 0,00 = nao informado
        rt = oficial.where(oficial.notna() & (oficial > -0.2) & (oficial < 0.1), rt)
        ser = ser.assign(cota_ant=ant.where(seguido), rt=rt)
        e = (1 + rt) / (1 + ser.mes.map(self.cdi_m)) - 1
        primeiro = g.cumcount() == 0                                         # 1o mes da serie e parcial
        ser = ser.assign(e=e.where(~primeiro & (e > -0.5) & (e < 0.2)))     # perda real conta; so descarta absurdo
        ser["spread"] = ((1 + ser.groupby(["cnpj", "serie"]).e.transform(lambda x: x.rolling(6, min_periods=3).median())) ** 12 - 1) * 100
        esp = ((1 + ser.esperado / 100) / (1 + ser.mes.map(self.cdi_m))) ** 12 * 100 - 100        # benchmark da serie (X_6)
        ruim = ~((ser.spread > -30) & (ser.spread < 40))                 # perda recorrente aparece; absurdo vira benchmark
        ser["spread"] = ser.spread.mask(ruim, esp.where((ser.esperado > 0) & (esp > -5) & (esp < 30)))
        ser = ser.dropna(subset=["spread"])
        sr = ser.serie.str.contains("Senior|Sênior", case=False, na=False)
        self.fidc_ser = {k: (x.cota.values, x.spread.values, float(x.spread[sr[x.index]].median()) if sr[x.index].any() else float(x.spread.median()),
                             x.cota_ant.values, x.rt.values, x.rent.values / 100) for k, x in ser.groupby(["cnpj", "mes"])}
        self.taxas = lt.taxas()
        self.ust = lt.treasury()
        px = self.ptax.groupby(self.ptax.index.strftime("%Y-%m")).last()
        self.fx_m = px / px.shift() - 1
        u = pd.Series({m: float(lt.curva(self.ust, "UST", fim_mes(m), np.array([1.0]))[0]) / 1200 for m in px.index}) if self.ust else pd.Series(dtype=float)
        self.ust_m = u.shift()
        self._taxas = {}
        self.emissores = {}                       # (mes, grupo) -> soma spread*peso, duration*peso, peso
        self.ref = pd.Series(dtype=float)        # (mes, codigo) -> preco mediano do mesmo ativo em 2+ fundos

    def precos_ref(self, pos):
        """Preco unitario mediano de cada ativo com codigo, por mes, entre todos os fundos/veiculos que o tem."""
        q = pos.assign(cod=pos.codigo.fillna("").astype(str).str.strip(), pu=pos.v / pd.to_numeric(pos.QT_POS_FINAL, errors="coerce"))
        q = q[(q.cod != "") & (q.pu > 0)]
        g = q.groupby(["mes", "cod"]).agg(ref=("pu", "median"), n=("CNPJ_FUNDO_CLASSE", "nunique"))
        self.ref = g.ref[g.n >= 2]

    def corrige(self, m, cods, pu):
        """Preco informado que foge > 10% da mediana dos outros fundos no mes = erro de reporte: usa a mediana."""
        pu = np.asarray(pu, float)
        if not len(self.ref) or not len(pu):
            return pu
        r = self.ref.reindex(pd.MultiIndex.from_arrays([[m] * len(pu), pd.Series(cods).fillna("").astype(str).str.strip()])).values
        return np.where(np.isfinite(r) & np.isfinite(pu) & (np.abs(pu / r - 1) > 0.10), r, pu)

    def pu_par(self, m):
        return self.pu.get(self.pu_dia.get(m), pd.Series(dtype=float))

    def proj_ipca(self, m):
        """IPCA mensal projetado (media dos 12 meses anteriores): o PU par do SND so usa o indice ja divulgado."""
        x = self.ipca[self.ipca.index < m]
        return float((x.iloc[-1] / x.iloc[-13]) ** (1 / 12) - 1) if len(x) > 13 else 0.004

    def fidc_anos(self, cnpj, m):
        x = self.fidc.get(cnpj)
        if x is None:
            return np.nan
        i = np.searchsorted(x[0], m, side="right") - 1
        return float(x[1][max(i, 0)])

    def fidc_spread(self, cnpj, m, pu):
        """Spread da serie do FIDC que o fundo tem: acha a serie pelo valor da cota (tab X_2) e usa a mediana de 6 meses
        da rentabilidade acima do CDI (tab X_3; benchmark X_6 se ela for anomala). Sem serie com a mesma cota: seniores."""
        for k in range(4):                                  # informe do mes ou de ate 3 meses antes
            x = self.fidc_ser.get((cnpj, (pd.Timestamp(m) - pd.DateOffset(months=k)).strftime("%Y-%m")))
            if x is not None:
                cota, spr, senior = x[:3]
                if np.isfinite(pu) and pu > 0:
                    dif = np.abs(cota / pu - 1)
                    if dif.min() < 0.02:
                        return float(spr[dif.argmin()])
                return senior
        return np.nan

    def fidc_ret(self, cnpj, m, pu_ant):
        """Retorno total do mes (cota + amortizacao) da serie que o fundo tinha no fim do mes anterior (achada pela cota)."""
        x = self.fidc_ser.get((cnpj, m))
        if x is None or not np.isfinite(pu_ant) or pu_ant <= 0:
            return np.nan
        ant, rt, rent = x[3], x[4], x[5]
        dif = np.abs(ant / pu_ant - 1)
        if not np.isfinite(dif).any() or np.nanmin(dif) >= 0.02:
            return np.nan
        i = np.nanargmin(dif)                   # rentabilidade oficial do mes (X_3); cota + amortizacao so se ela faltar/absurda
        return float(rent[i]) if np.isfinite(rent[i]) and -0.2 < rent[i] < 0.1 else float(rt[i])

    def serie_securit(self, x):
        """CRI/CRA da carteira -> serie no informe das securitizadoras. 1) ISIN ou codigo CETIP na descricao; 2) CNPJ (securitizadora,
        devedor ou cedente) + vencimento, desempate pelo indexador e pelo PU do fundo mais perto do valor da serie no mes. So casa
        quando sobra uma serie; ambiguo fica sem serie (continua a regra do carrego)."""
        if not len(getattr(self, "sec", ())):
            return pd.Series(None, index=x.index, dtype=object)
        txt = (x.descricao.fillna("") + " " + x.ativo.fillna("")).str.upper()
        isin = txt.str.extract(r"\b(BR[A-Z0-9]{4}CR[AI][A-Z0-9]{3})\b")[0]
        out = isin.map(self.sec_isin).where(isin.notna())
        cet = txt.str.extract(r"(?<![0-9A-Z])(\d{2}[A-L]\d{7}|CR[AI]\d{3}[0-9A-Z]{4,6})(?![0-9A-Z])")[0]
        out = out.where(out.notna(), cet.where(cet.isin(set(self.sec.serie))))
        ddmmyy = lambda t: pd.to_datetime(t, format="%d%m%y", errors="coerce").dt.strftime("%Y-%m-%d")
        venc = pd.to_datetime(x.vencimento, errors="coerce").dt.strftime("%Y-%m-%d")
        venc = venc.fillna(pd.to_datetime(txt.str.extract(r"(\d{2}/\d{2}/\d{4})")[0], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d"))
        barras = txt.str.extract(r"CR[IA]\s*/[^/]+/\s*(\d{6})\s*/\s*(\d{6})\s*/\s*(\d{14})")   # CRI/SECURIT/emissao/vencimento/CNPJ
        venc = venc.fillna(ddmmyy(barras[1])).fillna(ddmmyy(txt.str.extract(r"CR[IA]:[^:]+:(\d{6})\b")[0]))   # CRI:SIGLA:vencimento
        emis = ddmmyy(barras[0])
        pref = txt.str.extract(r"CR[IA]-(\d{2}[A-L]\d{3,6})\b")[0]                                        # codigo CETIP truncado
        for i in x.index[out.isna() & pref.notna()]:
            c_ = self.sec_prefixo[self.sec_prefixo.str.startswith(pref[i], na=False)]
            if len(c_) == 1:
                out[i] = c_.iloc[0]
        cn = x.cnpj_emissor.fillna("").astype(str).str.replace(r"\D", "", regex=True)
        cn = cn.where(cn.ne(""), barras[2].fillna(""))
        idx = x.indexador.fillna("").astype(str).str.upper()
        pu = (x.valor_veiculo / x.quantidade).where(x.quantidade > 0)
        falta = out.isna() & cn.ne("") & venc.notna()
        if falta.any():
            k, self.motivo = self.sec_chave, {}
            # mesmo papel ao longo dos meses (veiculo + CNPJ + vencimento + descricao): decide a serie com o historico inteiro
            grupo = (x.get("cnpj_veiculo", pd.Series("", index=x.index)).fillna("").astype(str) + "|" + cn + "|" + venc.fillna("") + "|"
                     + x.descricao.fillna("").astype(str))
            for _, ii in pd.Series(x.index[falta]).groupby(grupo[falta].values):
                ii = list(ii)
                i0 = ii[0]
                kc = k[k.cnpj == cn[i0]]
                c, perto = kc[kc.venc == venc[i0]], False
                if c.empty:                            # vencimento informado diferente do informe (dias uteis, prorrogacao) ou data = emissao
                    dv = (pd.to_datetime(kc.venc, errors="coerce") - pd.Timestamp(venc[i0])).dt.days.abs()
                    c, perto = kc[(dv <= 40) | (kc.emissao == venc[i0])], True
                if len(c) > 1 and idx[i0]:
                    di = "DI" in idx[i0] or "CDI" in idx[i0]
                    tx = c.taxa.str.upper()
                    c = c[tx.str.contains("DI") == di] if "IPCA" not in idx[i0] else c[tx.str.contains("IPCA")]
                if c.serie.nunique() > 1 and isinstance(emis[i0], str):        # data de emissao no texto
                    c2 = c[c.emissao == emis[i0]]
                    c = c2 if len(c2) else c
                cands = list(c.serie.unique())
                self.motivo.update(dict.fromkeys(ii, len(cands)))
                if len(cands) > 1 or (perto and cands):  # distancia mediana |PU do fundo / VNA - 1| em todos os meses com o papel
                    d = pd.Series({s_: np.nanmedian([self.ancora(s_, x.mes[i], pu[i])[2] for i in ii if np.isfinite(pu[i])] or [np.inf])
                                   for s_ in cands}).sort_values()
                    ok_ = d.iloc[0] < 0.05 and (len(d) == 1 or d.iloc[1] > 2 * d.iloc[0] + 0.01)
                    if not ok_ and d.iloc[0] < 0.05:
                        iguais = list(d[d < d.iloc[0] + 0.01].index)
                        # 1) series economicamente iguais (programas pulverizados): mesmo pago/VNA todo mes
                        taxa_pg = np.array([[(lambda a: a[1] / a[0] if a[0] else np.nan)(self.ancora(s_, x.mes[i], pu[i])) for i in ii
                                             if np.isfinite(pu[i])] for s_ in iguais], float)
                        ok_ = np.isfinite(taxa_pg).any() and np.nanmax(np.nanmax(taxa_pg, 0) - np.nanmin(taxa_pg, 0)) < 0.003
                        if not ok_:                    # 2) series irmas: o retorno oficial ao par (o fluxo) e o mesmo mes a mes
                            viz = d[d < min(0.05, d.iloc[0] + 0.02)].index
                            rp = np.array([self.erro_trajeto(s_, x.mes[ii].values, pu[ii].values, par=True) for s_ in viz], float)
                            ok_ = rp.shape[1] > 0 and np.isfinite(rp).all(0).sum() >= 2 and np.nanmax(np.nanmax(rp, 0) - np.nanmin(rp, 0)) < 0.002
                    self.motivo.update(dict.fromkeys(ii, d.iloc[:4].round(4).to_dict()))
                    cands = [d.index[0]] if ok_ else []
                if len(cands) == 1:
                    for i in ii:
                        out[i] = cands[0]
        return out

    def erro_trajeto(self, serie, meses, pus, par=False):
        """Erro mediano entre o retorno mensal do PU do fundo e o da serie ao par (VNA + pago) nos meses seguidos;
        par=True devolve o retorno ao par de cada mes seguido (NaN quando falta a serie no mes)."""
        o = np.argsort(meses)
        meses, pus = np.asarray(meses)[o], np.asarray(pus, float)[o]
        er = []
        for j in range(1, len(meses)):
            if meses[j - 1] != (pd.Period(meses[j]) - 1).strftime("%Y-%m") or not (pus[j] > 0 and pus[j - 1] > 0):
                continue
            v1, pg, _ = self.ancora(serie, meses[j], pus[j])
            v0 = self.ancora(serie, meses[j - 1], pus[j - 1])[0]
            ok = np.isfinite(v1) and np.isfinite(v0) and v0 > 0
            er.append(((v1 + pg) / v0 - 1 if ok else np.nan) if par else abs((pus[j] + pg) / pus[j - 1] - (v1 + pg) / v0) if ok else np.nan)
        if par:
            return er
        er = [e_ for e_ in er if np.isfinite(e_)]
        return float(np.median(er)) if er else np.inf

    def ancora(self, serie, m, pu):
        """(VNA, pago) da serie no mes na leitura que bate com o PU informado pelo fundo: cada securitizadora preenche o
        informe de um jeito (valor total da serie ou por certificado; pagamento total ou por certificado). Devolve tambem a
        distancia |PU/VNA - 1|; sem leitura a menos de 25% do PU -> NaN (serie nao usada no mes)."""
        try:
            V, Q, P, vna_c, pago_c = self.sec_m.loc[(serie, m)]
        except KeyError:
            return np.nan, np.nan, np.inf
        if not (np.isfinite(pu) and pu > 0 and Q > 0):
            return vna_c, pago_c, (abs(pu / vna_c - 1) if np.isfinite(pu) and vna_c > 0 else np.inf)
        cands = [(V / Q, P / Q), (V, P / Q), (V / Q, P), (V, P)]
        dist = [abs(pu / v - 1) if v > 0 else np.inf for v, _ in cands]
        k = int(np.argmin(dist))
        v, pg = cands[k]
        if pg > 0.6 * v:                                 # pagamento maior que 60% do valor: a outra escala do pagamento
            pg = P / Q if pg == P else P
        return (v, pg, dist[k]) if dist[k] < 0.25 else (np.nan, np.nan, dist[k])

    def pago_securit(self, series, m, pu1, pu0):
        """(pago por certificado no mes, VNA do mes, VNA do mes anterior) de cada serie, na escala do PU do fundo."""
        p = (pd.Timestamp(m) - pd.DateOffset(months=1)).strftime("%Y-%m")
        a = [self.ancora(s_, m, u) for s_, u in zip(series, pu1)]
        b = [self.ancora(s_, p, u) for s_, u in zip(series, pu0)]
        v1, v0 = np.array([x[0] for x in a], float), np.array([x[0] for x in b], float)
        pg = np.array([x[1] for x in a], float)
        for j, s_ in enumerate(series):                 # escala do pagamento (total ou por certificado): a que da o retorno ao par mais normal
            try:
                _, Q, P, _, _ = self.sec_m.loc[(s_, m)]
            except KeyError:
                continue
            if np.isfinite(v0[j]) and v0[j] > 0 and Q > 0 and np.isfinite(P):
                pg[j] = min((P, P / Q), key=lambda g: abs((v1[j] + g) / v0[j] - 1.01))
        return pg, v1, v0

    def pagto_snd(self, cods, p, m, pu_ant):
        """Debenture do SND: o que foi pago no mes por unidade (amortizacao + juros) = PU par do mes anterior corrigido
        pela remuneracao - PU par do mes. Devolve o pagamento (R$) ou NaN quando o SND nao calcula o PU (nao padrao)."""
        c = self.snd.reindex(cods)
        p0, p1 = self.pu_par(p).reindex(cods).values.astype(float), self.pu_par(m).reindex(cods).values.astype(float)
        idx, jur, pct = _txt(c.indice), c.juros.values.astype(float), c.pct.values.astype(float)
        cdi = float(self.cdi_m.get(m, 0.0))
        i0, i1 = self.ipca[self.ipca.index <= p], self.ipca[self.ipca.index <= m]
        inf = float(i1.iloc[-1] / i0.iloc[-1] - 1) if len(i0) and len(i1) else 0.0
        di = idx == "DI"
        s = np.where(di & (pct > 100) & (np.nan_to_num(jur) == 0), (pct - 100) / 100 * self.cdi_aa(m), jur)
        base = np.where(di, cdi, np.where(np.isin(idx, ["IPCA", "IGP-M", "INPC"]), inf, np.where(idx == "PRE", 0.0, np.nan)))
        padrao = np.isfinite(s) & (np.nan_to_num(jur) > 0) | (di & (pct > 100))
        pag = np.clip(p0 * (1 + base) * (1 + np.nan_to_num(s) / 100) ** (1 / 12) - p1, 0, None)
        pag = np.where((p1 < p0) & (pag > 0.005 * p0), pag, 0.0)   # evento de verdade derruba o PU par; mes sem evento so acumula
        ok = padrao & np.isfinite(pag) & (p0 > 0) & (p1 > 0) & np.isfinite(pu_ant)
        return np.where(ok, pag, np.nan)

    def taxa(self, m, cods, T, pj, pa, car, base, s0, pp_):
        """taxa_e_duration com memoria: o mesmo ativo no mesmo mes e preco aparece em varios fundos."""
        ch = [(m, c, round(float(p), 5) if np.isfinite(p) else None, round(float(b), 4)) for c, p, b in zip(cods, pp_, s0)]
        novos = {k: i for i, k in enumerate(ch) if k not in self._taxas}
        if novos:
            i = np.array(list(novos.values()))
            s, d = taxa_e_duration(T[i], pj[i], pa[i], car[i], base[i], s0[i], pp_[i])
            self._taxas.update(zip(novos, zip(s, d)))
        r = np.array([self._taxas[k] for k in ch], dtype=float).reshape(-1, 2)
        return r[:, 0], r[:, 1]

    def cdi_aa(self, m):
        x = self.cdi_d[self.cdi_d.index <= fim_mes(m)]
        return float(((1 + x.iloc[-1]) ** 252 - 1) * 100) if len(x) else np.nan

    def ptax_em(self, m):
        x = self.ptax[self.ptax.index <= fim_mes(m)]
        return float(x.iloc[-1]) if len(x) else np.nan

    def anbima_mes(self, m, recentes):
        """Marcacao ANBIMA para a carteira do mes: o dia guardado mais proximo do fim do mes (ate 10 dias);
        nos meses mais recentes, o ultimo dia disponivel."""
        if not self.an_dias:
            return None, None
        alvo = fim_mes(m)
        perto = [d for d in self.an_dias if abs((pd.Timestamp(d) - alvo).days) <= 10]
        d = min(perto, key=lambda d: abs((pd.Timestamp(d) - alvo).days)) if perto else (self.an_dias[-1] if m in recentes else None)
        return (d, self.an_h[self.an_h.data == d].drop_duplicates("codigo").set_index("codigo")) if d is not None else (None, None)


def tipo_tp(g):
    s = (g.codigo.fillna("") + " " + g.ativo.fillna("")).map(lt._sem_acento)
    out = pd.Series(None, index=g.index, dtype=object)
    for nome, rx in TP:
        out = out.mask(out.isna() & s.str.contains(rx), nome)
    return out


def vencimento(g):
    """Vencimento informado ou lido do nome do ativo (CRI: 'CRI:APCS:150629', '... / 25/09/2029 / ...', '... 2029-05')."""
    a = g.ativo.fillna("")
    v = pd.to_datetime(g.vencimento, errors="coerce")
    for rx, fmt in [(r"(\d{2}/\d{2}/\d{4})", "%d/%m/%Y"), (r"CR[AI]:[A-Z0-9]+:(\d{6})\b", "%d%m%y"), (r"\b((?:19|20)\d{2}-\d{2})$", "%Y-%m")]:
        v = v.fillna(pd.to_datetime(a.str.extract(rx)[0], format=fmt, errors="coerce"))
    return v


def casa(df, ctx):
    """Categorias da casa: Caixa, LF, LFSN, LFSC, DEB, DEBIN, Bonds, CRA, CRI, NC, FIDC. O que nao se encaixa fica com o
    nome do tipo/categoria da CVM (CDB, Títulos Públicos, FII, ...)."""
    t, c = df.tipo_ativo.fillna(""), df.categoria
    a = df.ativo.fillna("").map(lt._sem_acento)
    inc = df.codigo.map(ctx.snd.incentivada) if len(ctx.snd) else pd.Series(np.nan, index=df.index)
    out = t.map(TIPO_CASA).fillna(c.map(CURTO_CVM)).fillna(c)
    fundo = t.isin(["FIDC", "FI Imobiliário", "Fundos de Índice", "FI Participações", "Fundos Offshore"]) | c.isin(["FIDC", "Cotas de Fundos", "FIAGRO"])
    out = out.mask(c.eq("FIAGRO"), "FIAGRO").mask(t.str.startswith("Aç") | t.str.contains("subscri"), "Ações")
    out = out.mask(t.isin(["Bonds e Treasury", "Título da Dívida Externa"]), "Bonds")
    out = out.mask(t.eq("FIDC") | c.eq("FIDC"), "FIDC")
    fiagro = a.str.contains(r"FIAGRO|CADEIAS PRODUTIVAS DO AGRO")                     # FIAGRO (inclusive de direitos creditorios)
    out = out.mask(fiagro, "FIAGRO")
    fundo = fundo | fiagro
    ext = c.eq("Investimento no Exterior") & t.isin(["Outros", ""])                   # exterior sem tipo: bond com vencimento ou caixa
    venc_ = pd.to_datetime(df.vencimento, errors="coerce")
    out = out.mask(ext & venc_.notna(), "Bonds").mask(ext & venc_.isna() & a.str.contains(r"USD|BNY|CASH|CAIXA|MARGIN|CONTA"), "Caixa")
    out = out.mask(c.str.contains("termo", case=False, na=False) & ~c.str.contains(lt.PASSIVO, na=False), "Ações")
    acao = c.str.contains(r"Depository Receipt|BDR|recibo de dep|cedidos em empr|^Ações", case=False, na=False, regex=True)         | t.str.contains(r"BDR|Depository|recibo de dep", case=False, na=False, regex=True)
    out = out.mask(acao & ~c.str.contains(lt.PASSIVO, na=False), "Ações")               # BDR, recibos, acoes emprestadas
    out = out.mask(~fundo & (t.str.contains("Nota Promiss|Nota Comercial") | a.str.match(r"NOTA COMERCIAL")), "NC")
    out = out.mask(~fundo & (t.str.contains("imobili.rios") | a.str.contains(r"\bCRI\b")), "CRI")
    out = out.mask(~fundo & (t.eq("CRA") | a.str.match(r"CRA\b")), "CRA")
    deb = ~fundo & (t.str.startswith("Debênture") | a.str.match(r"DEBENTURE") | inc.notna())
    out = out.mask(deb, np.where(inc.eq(True), "DEBIN", "DEB"))
    lf = t.eq("Letra Financeira")
    venc = pd.to_datetime(df.vencimento, errors="coerce")
    chave = df.emissor.fillna("") + "|" + df.vencimento.astype(str)
    primeiro = pd.to_datetime(df.mes.where(lf)).groupby(chave).transform("min")
    longo = (venc - primeiro).dt.days / 365.25 > 6          # LF subordinada nivel II: prazo longo (> 6 anos), nao perpetua
    out = out.mask(lf, np.where(venc.dt.year >= 2049, "LFSC", np.where(longo, "LFSN", "LF")))
    out = out.mask(c.eq("Debêntures") & ~out.isin(["DEB", "DEBIN"]), "DEB")
    contabil = c.isin(CAIXA_CVM) | (c.str.contains(lt.PASSIVO, na=False) & ~df.derivativo)
    out = out.mask(contabil, "Caixa").mask(c.eq("Títulos Públicos"), "Títulos Públicos")
    lft = c.eq("Títulos Públicos") & (df.codigo.fillna("").str.contains("BRSTNCLF") | a.str.contains(r"FINANCEIRAS DO TESOURO|\bLFT\b"))
    out = out.mask(lft, "Caixa")                                                       # LFT: pos-fixado na Selic = caixa
    return out.mask(df.confidencial & ~df.derivativo, "Confidencial")       # CVM so informa o tipo, sem o ativo


def rotulo(g, venc, tp):
    """Nome curto 'codigo - mm/aaaa': debenture/acao pelo ticker, CRI/CRA pelo codigo, titulo publico pelo tipo,
    LF/CDB pelo tipo + grupo, fundos pelo nome sem palavras genericas."""
    a = g.ativo.fillna("")
    cod = g.codigo.fillna("").str.strip().str.split("_").str[0]
    cod = cod.where(~cod.str.fullmatch(r"\d+"), "")                                     # codigo so numerico nao identifica
    lote = a.str.extract(r"^\w+?_[-\d.e]+_[-\d.e]+_\d+\s+(.*?)(?:\s+-\s+(?:KY|US|LU|IE|BR)[A-Z0-9]{9,}.*)?$")[0]
    a = a.where(lote.isna(), lote.str.replace(r"\s+-\s+", " - ", regex=True).str.strip(" -"))      # offshore em "lotes" 
    fundo = g.casa.isin(FUNDOS)
    ticker = cod.str.fullmatch(r"[A-Z]{4}[A-Z0-9]{1,4}") & ~fundo
    cri = a.str.extract(r"\b(CR[AI]):([A-Z0-9]+)")
    isin = a.str.extract(r"\b(BR[A-Z0-9]{10})\b")[0]
    curto = a.str.replace(GENERICO, " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip(" -").str.slice(0, 38)
    base = np.select([ticker, tp.notna(), cri[0].notna(), isin.notna(), fundo],
                     [cod, tp.fillna(""), cri[0].fillna("") + " " + cri[1].fillna(""), isin.fillna(""), curto],
                     g.casa + " " + g.grupo.fillna("").str.slice(0, 26))
    cx = g.casa.eq("Caixa")
    base = np.where(cx & tp.isna(), np.where(g.categoria.eq("Operações Compromissadas"), "Compromissada", g.categoria), base)
    base = np.where(cx & tp.notna() & g.categoria.eq("Operações Compromissadas"), "Compromissada " + tp.fillna(""), base)
    v = venc.dt.strftime("%m/%Y")
    return pd.Series(base, index=g.index).str.strip() + np.where(v.notna() & ~fundo & (~cx | (tp.notna() & ~g.categoria.eq("Operações Compromissadas"))),
                                                                  " - " + v.fillna(""), "")


def implicito(df, ctx):
    """Spread implicito pelo preco (ativos sem taxa publica: cotas de FIDC, CRI sem taxa, bonds...): mediana dos ultimos
    6 retornos mensais do preco unitario acima do CDI, anualizada. Em dolar: tira o cambio e compara com o Treasury 1 ano."""
    k = lt.contribuicao(df)
    if not len(k):
        return pd.Series(dtype=float)
    k = k[(k.qt > 0) & (k.qt_ant > 0) & (k.v_ant > 0)].copy()
    rp = (k.v / k.qt) / (k.v_ant / k.qt_ant) - 1
    usd = k.chave.map(df.drop_duplicates("chave").set_index("chave").casa).isin(USD)
    e_brl = (1 + rp) / (1 + k.mes.map(ctx.cdi_m)) - 1
    e_usd = (1 + rp) / ((1 + k.mes.map(ctx.fx_m).fillna(0)) * (1 + k.mes.map(ctx.ust_m).fillna(0))) - 1
    # ativo em dolar com hedge (cota em reais): a serie em reais e a mais estavel; sem hedge, a em dolar
    vol = pd.DataFrame({"c": k.chave, "b": e_brl, "u": e_usd}).groupby("c")[["b", "u"]].std()
    em_usd = k.chave.map(vol.u < vol.b).fillna(True).astype(bool)
    e = e_brl.where(~(usd & em_usd), e_usd)
    k["e"] = e.where(e.abs() < 0.2)
    k = k.sort_values(["chave", "mes"])
    k["med"] = k.groupby("chave").e.transform(lambda s: s.rolling(6, min_periods=1).median())
    return ((1 + k.set_index(["mes", "chave"]).med) ** 12 - 1) * 100


def mede(g, m, ctx, recentes, impl, ant=None):
    """Spread CDI (% a.a., equivalente), duration (anos), PU, indexador e fonte de TODAS as linhas da carteira do mes.
    Preferencia: ANBIMA (mercado) > SND + preco do fundo (mercado) > taxa contratada (CVM) > implicito pelo preco >
    media da categoria. Renda variavel (acoes, FII, FIP, ETF, FIAGRO): nao se aplica."""
    g = g.copy()
    fm, cdi, n = fim_mes(m), ctx.cdi_aa(m), len(g)
    sp, du = np.full(n, np.nan), np.full(n, np.nan)
    fonte, ind = np.full(n, "", dtype=object), np.full(n, "", dtype=object)
    cod = g.codigo.fillna("").values
    snd = ctx.snd.reindex(cod)
    venc = g.venc.fillna(pd.Series(snd.venc.values, index=g.index))
    anos = ((venc - fm).dt.days / 365.25).clip(lower=1 / 365).values
    tem_venc = venc.notna().values
    pu_f = ctx.corrige(m, cod, (g.valor_veiculo / g.quantidade).where(g.quantidade > 0).values)
    ca = _txt(g.casa)
    nome = _txt(g.ativo.fillna("").map(lt._sem_acento))
    curva = lambda t, a: lt.curva(ctx.curvas, t, fm, np.nan_to_num(a, nan=3))
    valido = lambda v: np.isfinite(v) & (v > -5) & (v < 30)          # spread fora disso = dado ruim: tenta a proxima fonte

    def poe(msk, s, d, f, i):                    # preenche so onde ainda nao ha spread
        msk = msk & np.isnan(sp)
        b = lambda v: np.broadcast_to(np.asarray(v, dtype=object if isinstance(v, str) or (np.ndim(v) and np.asarray(v).dtype == object) else float), (n,))
        sp[msk], du[msk], fonte[msk], ind[msk] = b(s)[msk], b(d)[msk], b(f)[msk], b(i)[msk]

    # 1) caixa e titulos publicos
    poe((ca == "Caixa") | ((ca == "ETF") & _tem(nome, "SELIC")), 0.0, 0.0, "caixa", "DI")
    tp = _txt(g.tp)
    m_tp = ca == "Títulos Públicos"
    if m_tp.any():
        y = np.where(tp == "NTN-B", curva("IPCA", anos), curva("PRE", anos))
        _, d_cup = taxa_e_duration(anos, 0.5, 0, 0, 0, y, np.nan)
        poe(m_tp, 0.0, np.where(np.isin(tp, ["NTN-B", "NTN-F"]), d_cup, anos), "Tesouro (curva)",
            np.where(tp == "NTN-B", "IPCA", np.where(np.isin(tp, ["LTN", "NTN-F"]), "PRE", "DI")).astype(object))
    # 2) debentures: ANBIMA do dia mais proximo do fim do mes
    dia, an = ctx.anbima_mes(m, recentes)
    deb = np.isin(ca, ["DEB", "DEBIN"])
    if an is not None and deb.any():
        a = an.reindex(cod)
        tx, tipo, dur = a.taxa_ind.values.astype(float), _txt(a.tipo), a.duration_anos.values.astype(float)
        real = np.isin(tipo, ["IPCA +", "IGP-M +"])
        s = np.where(tipo == "DI +", tx, np.where(tipo == "% do DI", (tx - 100) / 100 * cdi, np.nan))
        s = np.where(real, tx - lt.curva(ctx.curvas, "IPCA", pd.Timestamp(dia), np.nan_to_num(dur, nan=3)), s)
        s = np.where(tipo == "Outros", tx - lt.curva(ctx.curvas, "PRE", pd.Timestamp(dia), np.nan_to_num(dur, nan=2)), s)
        poe(deb & np.isfinite(s), s, dur, "ANBIMA", np.where(real, "IPCA", np.where(tipo == "Outros", "PRE", "DI")).astype(object))
    # 3) debentures fora da ANBIMA: taxa de emissao (SND) ajustada pelo preco do fundo / PU par do SND
    idx = _txt(snd.indice)
    di, real, pre = idx == "DI", np.isin(idx, ["IPCA", "IGP-M", "INPC"]), idx == "PRE"
    m3 = deb & np.isnan(sp) & (di | real | pre) & tem_venc
    if m3.any():
        pct, jur = snd.pct.values.astype(float), snd.juros.values.astype(float)
        s0 = np.where(di & (pct > 100) & (np.nan_to_num(jur) == 0), (pct - 100) / 100 * cdi, jur)
        s0 = np.where((np.nan_to_num(jur) == 0) & ~(di & (pct > 100)), np.nan, s0)     # "nao padrao" no SND: sem taxa util
        par = ctx.pu_par(m).reindex(cod).values.astype(float) * np.where(real, 1 + ctx.proj_ipca(m), 1)
        pp_ = pu_f / par
        pp_ = np.where((pp_ > 0.7) & (pp_ < 1.15), pp_, np.nan)
        car = ((snd.carencia - fm).dt.days / 365.25).values.astype(float)
        r = np.where(m3 & np.isfinite(s0))[0]
        s, d = np.full(n, np.nan), np.full(n, np.nan)
        s[r], d[r] = ctx.taxa(m, cod[r], anos[r], snd.cada_juros.values.astype(float)[r] / 12,
                              np.nan_to_num(snd.cada_amort.values.astype(float)[r]) / 12, car[r], np.where(di[r], cdi, 0), s0[r], pp_[r])
        s = np.where(np.abs(s - s0) > 15, s0, s)              # preco sem sentido (evento, erro de unidade): fica a taxa de emissao
        spr = np.where(di, s, np.where(real, s - curva("IPCA", d), s - curva("PRE", d)))
        poe(m3 & valido(spr), spr, d, np.where(np.isfinite(pp_), "SND + preço do fundo", "SND (taxa de emissão)").astype(object),
            np.where(di, "DI", np.where(real, "IPCA", "PRE")).astype(object))
    # 4) taxa contratada informada a CVM (LF, CDB, CRA, NC, debenture sem codigo...)
    ix = _txt(g.indexador)
    cup = g.cupom.values.astype(float)
    cup = np.where(np.nan_to_num(cup) != 0, cup, np.nan)
    pct_i, tpre = g.pct_indexador.values.astype(float), g.taxa_pre.values.astype(float)
    i_di = _tem(ix, "DI de um dia") | _tem(ix, "Selic")
    outros = (_tem(ix, "OUTROS") | _tem(ix, "Outros")) & np.isfinite(g.cupom.values.astype(float)) & (g.cupom.values.astype(float) > 0) & (g.cupom.values.astype(float) <= 5.5)
    i_di = i_di | outros                                   # indexador "outros" com cupom baixo (ate 5,5): CDI + cupom
    i_ip, i_pr = _tem(ix, "IPCA"), _tem(ix, "prefixada")
    m4 = np.isnan(sp) & (i_di | i_ip | i_pr) & ~np.isin(ca, SEM_SPREAD) & tem_venc
    if m4.any():
        cup_di = np.where(cup >= 50, (cup - 100) / 100 * cdi, cup)                  # % do CDI informado no campo de cupom
        taxa = np.where(i_di, np.where(np.isfinite(cup), cup_di, (pct_i - 100) / 100 * cdi), np.where(i_ip, cup, np.where(np.isfinite(tpre), tpre, cup)))
        zero = np.isin(ca, ["LF", "LFSN", "CDB", "DPGE", "LCA", "NC", "CCB", "LCI/LH"])   # pagam tudo no vencimento
        _, d = taxa_e_duration(anos, np.where(zero, anos + 1, 0.5), 0, 0, np.where(i_di, cdi, 0), np.nan_to_num(taxa), np.nan)
        spr = np.where(i_di, taxa, np.where(i_ip, taxa - curva("IPCA", d), taxa - curva("PRE", d)))
        poe(m4 & valido(spr), spr, d, "taxa contratada (CVM)", np.where(i_di, "DI", np.where(i_ip, "IPCA", "PRE")).astype(object))
    # 5) sem taxa publica: implicito pelo preco (cotas de FIDC, CRI sem taxa, bonds, outros)
    fid = np.isin(ca, ["FIDC", "FIAGRO"])                 # FIAGRO de direitos creditorios tambem entrega o informe de FIDC
    d_fidc = np.full(n, np.nan)
    d_fidc[fid] = [ctx.fidc_anos(c, m) for c in cod[fid]]
    m_f = fid & np.isnan(sp)
    if m_f.any():
        v = np.full(n, np.nan)
        v[m_f] = [ctx.fidc_spread(c, m, p) for c, p in zip(cod[m_f], pu_f[m_f])]
        poe(m_f & np.isfinite(v) & (v > -30) & (v < 40), v, d_fidc, "FIDC: rentabilidade da série (CVM)", "DI")
    m5 = np.isnan(sp) & ~np.isin(ca, SEM_SPREAD)
    if m5.any() and len(impl):
        v = impl.reindex(pd.MultiIndex.from_arrays([[m] * n, g.chave.values])).values.astype(float)
        usd = np.isin(ca, USD)
        d = d_fidc.copy()
        _, d_cri = taxa_e_duration(anos, 1 / 12, 1 / 12, 0, cdi, np.nan_to_num(v), np.nan)       # CRI: amortiza todo mes
        d = np.where((ca == "CRI") & tem_venc, d_cri, d)
        ust = lt.curva(ctx.ust, "UST", fm, np.nan_to_num(anos, nan=5)) if ctx.ust else np.full(n, 4.5)
        _, d_b = taxa_e_duration(anos, 0.5, 0, 0, ust, np.nan_to_num(v), np.nan)
        d = np.where((ca == "Bonds") & tem_venc, d_b, d)
        zero = np.isin(ca, ["LF", "LFSN", "CDB", "DPGE", "LCA", "NC", "CCB", "LCI/LH"])
        _, d_g = taxa_e_duration(anos, np.where(zero, anos + 1, 0.5), 0, 0, cdi, np.nan_to_num(v), np.nan)
        d = np.where(np.isnan(d) & tem_venc, d_g, d)
        poe(m5 & valido(v) & (v > -2) & ~(np.isin(ca, USD) & (v > 12)), v, d, np.where(usd, "implícito (preço em US$)", np.where(fid & np.isfinite(d), "implícito (cota) + prazo FIDC",
                                                                                        "implícito (preço)")).astype(object),
            np.where(usd, "USD", "DI").astype(object))
    # 6) renda variavel: nao se aplica; fundo offshore sem preco utilizavel: carteira nao e publica
    off = (ca == "Fundos offshore") & np.isnan(sp)
    na = (np.isin(ca, SEM_SPREAD) & np.isnan(sp)) | off
    fonte[na] = np.where(ca[na] == "Confidencial", "confidencial (CVM)",
                         np.where(off[na], "não se aplica (fundo offshore: carteira não é pública)", "não se aplica"))
    # 7) o que faltar: mesmo emissor (neste fundo ou em outro fundo no mes); senao media da categoria; senao do fundo
    w = g.perc_pl.values.astype(float)
    gr = _txt(g.grupo)
    ok = np.isfinite(sp) & np.isfinite(du) & ~na & (w > 0) & ~np.isin(ca, ["Caixa", "Títulos Públicos"])         & ~np.array([str(f).startswith("estimado") for f in fonte], dtype=bool)
    for nome_g in np.unique(gr[ok]):                                    # alimenta o mapa de emissores do mes
        x = ok & (gr == nome_g)
        a = ctx.emissores.setdefault((m, nome_g), [0.0, 0.0, 0.0])
        a[0] += (sp[x] * w[x]).sum(); a[1] += (du[x] * w[x]).sum(); a[2] += w[x].sum()
    sem = np.isnan(sp) & ~na
    for nome_g in np.unique(gr[sem]):
        a = ctx.emissores.get((m, nome_g)) if nome_g else None
        if a and a[2] > 0:
            sel = sem & (gr == nome_g)
            sp[sel], fonte[sel], ind[sel] = a[0] / a[2], "estimado (mesmo emissor)", "DI"
            du[sel] = np.where(np.isnan(du[sel]), a[1] / a[2], du[sel])
    falta_s, falta_d = np.isnan(sp) & ~na, np.isnan(du) & ~na
    if falta_s.any() or falta_d.any():
        cred = ~np.isin(ca, ["Caixa", "Títulos Públicos"]) & ~na & (w > 0)

        def media(v, msk):                         # mediana ponderada: um ativo em default nao contamina a estimativa
            ok = msk & np.isfinite(v) & (w > 0)
            if not ok.any():
                return np.nan
            o = np.argsort(v[ok])
            acum = np.cumsum(w[ok][o])
            return float(v[ok][o][np.searchsorted(acum, acum[-1] / 2)])
        for c in np.unique(ca[falta_s | falta_d]):
            mc = ca == c
            for v, falta, rotulo_ in ((sp, falta_s, None), (du, falta_d, " · duration estimada")):
                sel = mc & falta
                if not sel.any():
                    continue
                x, orig = media(v, mc & ~falta), "estimado (média da categoria)"
                if np.isnan(x) or (v is sp and x < -2):          # nao espalha perda de outro ativo para quem nao tem dado
                    x, orig = media(v, cred & ~falta), "estimado (média do fundo)"
                if v is sp and np.isfinite(x) and x < -2:
                    x = np.nan
                if np.isnan(x) and ant:                  # FIC sem a carteira do fundo investido no mes
                    x, orig = ant[0 if v is sp else 1], "estimado (carteira do mês anterior)"
                v[sel] = x
                if rotulo_ is None:
                    fonte[sel], ind[sel] = orig, "DI"
                else:
                    fonte[sel & ~falta_s] = fonte[sel & ~falta_s] + rotulo_
    g["spread"], g["duration"], g["pu"], g["fonte"] = sp, du, pu_f, fonte
    g["ipca"], g["pre"] = ind == "IPCA", ind == "PRE"
    g["usd"] = np.isin(ca, USD) | (ind == "USD")
    return g, dia


def derivativos_mes(gd, m, ctx, dur_ipca, dur_pre):
    """Contratos futuros do mes: tipo, lado, contratos, prazo e nocional estimado (R$)."""
    out = []
    for r in gd.itertuples():
        tipo, venc = contrato(f"{r.codigo or ''} {r.ativo or ''} {r.tipo_ativo or ''}")
        if "mercado futuro" not in str(r.categoria).lower():   # termo/opcao/swap: sem P&L de futuro
            tipo, venc = "OUTRO", None
        if tipo in ("DI1", "DAP") and not venc:                 # sem vencimento nao da para saber o prazo (ex.: trava de curva)
            tipo = "OUTRO"
        lado = -1 if "vendida" in str(r.categoria).lower() else 1
        n = abs(r.quantidade) if pd.notna(r.quantidade) else np.nan
        if venc:
            anos = max((pd.Timestamp(venc[0], venc[1], 15 if tipo == "DAP" else 1) - fim_mes(m)).days / 365.25, 1 / 252)
        else:
            anos = dur_ipca if tipo == "DAP" else dur_pre if tipo == "DI1" else 1.0
        out.append(dict(i=r.Index, tipo=tipo, lado=lado, n=n, anos=anos, fator=r.fator, categoria=r.categoria,
                        ativo=r.ativo, perc_cda=r.perc_pl))
    return out


def valor_ponto(tipo, m, ctx):
    if tipo == "DI1":
        return 1.0
    if tipo == "DAP":
        x = ctx.ipca[ctx.ipca.index <= m]
        return 0.00025 * float(x.iloc[-1]) if len(x) else np.nan
    return np.nan


def pu(tipo, m, anos, ctx):
    t = "PRE" if tipo == "DI1" else "IPCA"
    y = float(lt.curva(ctx.curvas, t, fim_mes(m), np.array([anos]))[0])
    return 1e5 / (1 + y / 100) ** anos


def nocional(c, m, ctx, pl_fundo):
    if c["tipo"] in ("DI1", "DAP") and pd.notna(c["n"]):
        return c["lado"] * c["n"] * pu(c["tipo"], m, c["anos"], ctx) * valor_ponto(c["tipo"], m, ctx)
    if c["tipo"] in ("DOL", "WDO") and pd.notna(c["n"]):
        return c["lado"] * c["n"] * (50000 if c["tipo"] == "DOL" else 10000) * ctx.ptax_em(m)
    return c["perc_cda"] / 100 * pl_fundo if pd.notna(c["perc_cda"]) else np.nan


def pnl_futuro(c, p, m, ctx):
    """P&L estimado (R$) do contrato entre o fim do mes p e o fim do mes m, com a posicao de p."""
    if pd.isna(c["n"]):
        return 0.0
    cdi = float(ctx.cdi_m.get(m, 0.0))
    if c["tipo"] in ("DI1", "DAP"):
        a1 = max(c["anos"] - (fim_mes(m) - fim_mes(p)).days / 365.25, 1 / 252)
        pu0, pu1 = pu(c["tipo"], p, c["anos"], ctx), pu(c["tipo"], m, a1, ctx)
        if c["tipo"] == "DI1":
            return c["lado"] * c["n"] * (pu1 - pu0 * (1 + cdi))
        i0, i1 = ctx.ipca[ctx.ipca.index <= p], ctx.ipca[ctx.ipca.index <= m]
        inf = float(i1.iloc[-1] / i0.iloc[-1] - 1) if len(i0) and len(i1) else 0.0
        return c["lado"] * c["n"] * valor_ponto("DAP", m, ctx) * (pu1 - pu0 * (1 + cdi) / (1 + inf))
    if c["tipo"] in ("DOL", "WDO"):
        return c["lado"] * c["n"] * (50000 if c["tipo"] == "DOL" else 10000) * (ctx.ptax_em(m) - ctx.ptax_em(p))
    return 0.0


def detalhe(nome, cnpj, df, ctx, pl, rent):
    """Tudo que as abas Carteira e Retorno mostram para um fundo + tabela de ativos (busca/top/grupos) de cada mes."""
    df = df.copy()
    # fecha 100% do PL em todo mes: a diferenca vira uma linha explicita (confidencial se o mes tem parte relevante assim)
    nd = df[~df.derivativo]
    falta = (100 - nd.groupby("mes").perc_pl.sum()).loc[lambda x: x.abs() > 0.05]
    if len(falta):
        conf = (nd[nd.confidencial].groupby("mes").perc_pl.sum().reindex(falta.index).fillna(0) >= 5) & (falta.values > 0)   # negativo = ajuste
        df = pd.concat([df, pd.DataFrame({"fundo": nome, "mes": falta.index, "categoria": lt.AJUSTE, "perc_pl": falta.values,
                                          "ativo": "Diferença entre PL e carteira informada (CVM)", "derivativo": False,
                                          "confidencial": conf.values, "chave": "ajuste|" + falta.index})], ignore_index=True)
    for c in ("emissor", "codigo", "ativo", "tipo_ativo", "indexador", "categoria"):     # coluna toda vazia vira texto
        if c in df.columns:
            df[c] = df[c].astype(object)
    if len(ctx.snd):                                    # emissor faltando: nome do emissor pelo SND / ANBIMA
        df["emissor"] = df.emissor.fillna(df.codigo.map(ctx.snd.empresa))
    if ctx.an_h is not None and len(ctx.an_h):
        df["emissor"] = df.emissor.fillna(df.codigo.map(ctx.an_h.drop_duplicates("codigo", keep="last").set_index("codigo").nome))
    lixo = r"(?i)\bBONDS?\s*\d*\s*DAYS?\s*SETTLE\b|\s[\d.,]{5,}\b"           # bonds: 'BONDS 2 DAYS SETTLE - emissor - qtd'
    df["emissor"] = df.emissor.astype(object).str.replace(lixo, " ", regex=True).str.strip(" -").replace("", "Emissor não informado")
    emis = df[["emissor", "categoria"]].drop_duplicates()
    k = lambda e: e if isinstance(e, str) else None                            # NaN nao serve de chave
    gmap = {(k(e), c): lt.grupo(e, c) for e, c in zip(emis.emissor, emis.categoria)}
    df["grupo"] = [gmap[(k(e), c)] for e, c in zip(df.emissor, df.categoria)]
    curto = df.ativo.str.replace(GENERICO, " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip(" -").str.slice(0, 40)
    df["grupo"] = df.grupo.fillna(curto)
    tick = df.ativo.fillna("").str.strip()
    df["codigo"] = df.codigo.where(df.codigo.notna(), tick.where(tick.str.fullmatch(r"[A-Z]{4}[A-Z0-9]{1,4}")))
    df["casa"] = casa(df, ctx)
    sec_ = df.casa.isin(["CRI", "CRA"])                 # CRI/CRA: serie do informe da securitizadora (CVM) vira o codigo do ativo
    df["serie"] = None
    if sec_.any():
        df.loc[sec_, "serie"] = ctx.serie_securit(df[sec_]).values
        df["codigo"] = df.codigo.where(df.serie.isna(), df.serie)
    df["venc"] = vencimento(df)
    df["tp"] = tipo_tp(df)
    df["rotulo"] = rotulo(df, df.venc.fillna(df.codigo.map(ctx.snd.venc) if len(ctx.snd) else df.venc), df.tp)
    impl = implicito(df, ctx)
    meses = sorted(df.mes.unique())
    recentes = set(meses[-4:])
    pl_f = {m: float(pl.get((m, cnpj), np.nan)) for m in meses}
    cart_meses, ricos, tab_ativos, ant = [], {}, [], None
    for m in meses:
        g = df[df.mes == m]
        e, dia = mede(g[~g.derivativo], m, ctx, recentes, impl, ant)
        if CAPTURA is not None:
            CAPTURA[(nome, m, "carteira")] = e.copy()
        ricos[m] = (e, g[g.derivativo])
        e["fin"] = e.perc_pl / 100 * pl_f[m] / 1e6
        cats = [{"categoria": c, "perc": x.perc_pl.sum(), "fin": x.fin.sum(), "spread": wavg(x.spread, x.perc_pl),
                 "duration": wavg(x.duration, x.perc_pl)} for c, x in e.groupby("casa")]
        cats.sort(key=lambda r: -r["perc"])
        total = {"categoria": "Total", "perc": e.perc_pl.sum(), "fin": e.fin.sum(), "spread": wavg(e.spread, e.perc_pl),
                 "duration": wavg(e.duration, e.perc_pl)}
        pos = e[e.perc_pl > 0]
        fontes = (pos.groupby(pos.fonte.str.split(" · ").str[0]).perc_pl.sum() / max(pos.perc_pl.sum(), 1e-9) * 100).round(1)
        if total["spread"] is not None and total["duration"] is not None:
            ant = (total["spread"], total["duration"])
        cart_meses.append({"mes": m, "pl_mm": pl_f[m] / 1e6, "anbima": dia, "conf": e[e.confidencial].perc_pl.sum(),
                           "cats": cats, "total": total, "fontes": fontes.sort_values(ascending=False).to_dict()})
        # tabela de ativos do mes (rotulo curto), maior posicao primeiro
        x = e[e.perc_pl != 0].sort_values("perc_pl", ascending=False)
        x = x.assign(sw=(x.spread * x.perc_pl).fillna(0), sp_=x.perc_pl.where(x.spread.notna(), 0),
                     dw=(x.duration * x.perc_pl).fillna(0), dp=x.perc_pl.where(x.duration.notna(), 0))
        t = x.groupby("rotulo", sort=False).agg(categoria=("casa", "first"), grupo=("grupo", "first"), emissor=("emissor", "first"),
                                                codigo=("codigo", "first"), perc=("perc_pl", "sum"), sw=("sw", "sum"), sp_=("sp_", "sum"),
                                                dw=("dw", "sum"), dp=("dp", "sum"), pu=("pu", "first"), fonte=("fonte", "first")).reset_index()
        t["spread"] = (t.sw / t.sp_).where(t.sp_ != 0)
        t["pct_par"] = (t.pu / ctx.pu_par(m).reindex(t.codigo.fillna("").values).values * 100).where(t.categoria.isin(["DEB", "DEBIN"]))
        t["pct_par"] = t.pct_par.where(t.pct_par.between(30, 130))
        t["duration"] = (t.dw / t.dp).where(t.dp != 0)
        tab_ativos.append(t.drop(columns=["sw", "sp_", "dw", "dp"]).assign(mes=m))

    # ---- atribuicao mensal (retorno por categoria, hedges alocados nos ativos protegidos)
    real = rent[rent.fundo == nome].set_index("mes").retorno
    k = lt.contribuicao(df)
    if len(k):                                   # variacao do preco unitario de cada posicao mantida (v/qt)
        k = k[(k.qt > 0) & (k.qt_ant > 0) & (k.v_ant > 0)]
        k = k.assign(rp=(k.v / k.qt) / (k.v_ant / k.qt_ant) - 1, pu1=k.v / k.qt, pu0=k.v_ant / k.qt_ant)
        preco = k.groupby(["mes", "chave"]).rp.first()
        nivel = k.groupby(["mes", "chave"])[["pu0", "pu1"]].first()
    else:
        preco = pd.Series(dtype=float)
    atr_meses, cats_atr, ativos, contrib = [], {}, {}, []
    id_ativo = lambda a, c, gr, cod, em: ativos.setdefault((a, c), {"i": len(ativos), "ativo": a, "categoria": c, "grupo": gr,
                                                                     "codigo": cod, "emissor": em})["i"]
    for j in range(1, len(meses)):
        p, m = meses[j - 1], meses[j]
        if m not in real.index or pd.isna(real.get(m)):
            continue
        e0, gd0 = ricos[p]
        w = e0[e0.perc_pl.abs() > 0].copy()            # confidenciais entram (rendem CDI + residuo): a soma sempre fecha
        conf_p, preench = float(w[w.casa.eq("Confidencial")].perc_pl.sum()), 0.0
        if conf_p >= 5:                                 # carteira escondida: a parte confidencial vira a ultima carteira aberta (estimativa)
            q = next((x for x in reversed(meses[:j]) if ricos[x][0][ricos[x][0].casa.eq("Confidencial")].perc_pl.sum() < 5), None)
            if q is not None:
                eq = ricos[q][0]
                pub = eq[~eq.casa.isin(["Confidencial", *CAIXA]) & eq.perc_pl.gt(0)]
                if len(pub):
                    w = pd.concat([w[~w.casa.eq("Confidencial")], pub.assign(perc_pl=pub.perc_pl / pub.perc_pl.sum() * conf_p)],
                                  ignore_index=True)
                    preench = conf_p
        cdi_m = float(ctx.cdi_m.get(m, 0.0))
        # 1) retorno de cada ativo: esperado = CDI + spread proprio; usa a variacao real de preco quando ela e critica
        #    (queda maior que 1,5 p.p. abaixo do esperado = pagamento de cupom/amortizacao -> fica o esperado)
        esperado = (1 + cdi_m) * (1 + w.spread.fillna(0).clip(-5, 30) / 100) ** (1 / 12) - 1
        # + marcacao a mercado: -duration x (variacao da curva IPCA/pre no prazo + variacao do spread de mercado)
        D = w.duration.fillna(0).clip(0, 30).values
        dc = np.zeros(len(w))
        for flag, t in ((w.ipca.values, "IPCA"), (w.pre.values, "PRE")):
            if flag.any():
                dc[flag] = lt.curva(ctx.curvas, t, fim_mes(m), D[flag]) - lt.curva(ctx.curvas, t, fim_mes(p), D[flag])
        e1 = ricos[m][0].drop_duplicates("chave").set_index("chave")
        ds = (w.chave.map(e1.spread) - w.spread).where(w.fonte.isin(MERCADO) & w.chave.map(e1.fonte).isin(MERCADO))
        y = ctx.cdi_aa(p) / 100 + w.spread.fillna(0).values / 100
        esperado = esperado - D / (1 + y) * (np.nan_to_num(dc) + ds.fillna(0).clip(-5, 5).values) / 100
        rp = w.chave.map(preco.loc[m]) if len(preco) and m in preco.index.get_level_values(0) else pd.Series(np.nan, index=w.index)
        r = rp.where(rp.notna() & (rp >= esperado - 0.015) & (rp <= esperado + 0.03), esperado)
        # retorno total sem filtro onde da para separar pagamento de perda:
        #   debenture do SND: (preco do mes + pagamento do mes pelo PU par) / preco anterior - 1
        #   FIDC: retorno total oficial da serie (cota + amortizacao, informe mensal da CVM)
        nv = nivel.loc[m].reindex(w.chave.values) if len(preco) and m in nivel.index.get_level_values(0) else pd.DataFrame(index=w.chave.values, columns=["pu0", "pu1"], dtype=float)
        cods_w = w.codigo.fillna("").values
        pu0, pu1 = ctx.corrige(p, cods_w, nv.pu0.values), ctx.corrige(m, cods_w, nv.pu1.values)
        rp = rp.mask(np.isfinite(pu0) & np.isfinite(pu1), pu1 / pu0 - 1)                 # preco corrigido entre fundos
        # CRI/CRA/NC/CCB pagam juros e amortizam todo mes e a CVM nao publica o fluxo: preco abaixo do esperado = pagamento
        # (fica o carrego); so perda que persiste no mes seguinte (calote/remarcacao) entra como queda
        amort = w.casa.isin(["CRI", "CRA", "NC", "CCB"]).values
        piso = np.where(amort, esperado - 0.003, esperado - 0.015)
        faixa = rp.notna() & (rp >= piso) & (rp <= esperado + 0.03)
        r = rp.where(faixa, esperado)
        src = np.where(faixa, "preço do fundo", "carrego + curva (sem preço)").astype(object)
        # queda forte so vale se persistir no mes seguinte (calote/remarcacao fica; erro de dado da CVM volta)
        m2 = meses[j + 1] if j + 1 < len(meses) else None
        pu2 = (ctx.corrige(m2, cods_w, nivel.loc[m2].pu1.reindex(w.chave.values).values)
               if m2 and len(preco) and m2 in nivel.index.get_level_values(0) else np.full(len(w), np.nan))
        persiste = lambda queda: np.isfinite(pu2) & (pu2 / pu0 - 1 < queda / 2)
        perda = amort & np.isfinite(rp.values) & (rp.values < esperado.values - 0.10) & persiste(rp.values)   # queda > 10% que fica
        r = r.mask(perda, rp)
        src[perda] = "preço do fundo (perda que persiste)"
        exp_ = esperado.values
        sc = w.serie.notna().values if "serie" in w else np.zeros(len(w), bool)
        if sc.any():                                    # CRI/CRA casados: (preco + pago por certificado) / preco anterior - 1
            pg, v1, v0 = ctx.pago_securit(w.serie.values[sc], m, pu1[sc], pu0[sc])
            rs = (pu1[sc] + pg) / pu0[sc] - 1
            ofi = (v1 + pg) / v0 - 1                    # retorno oficial da serie (ao par) quando falta o preco do fundo
            parado = np.isclose(pu1[sc], pu0[sc], rtol=1e-6)     # PU do fundo nao atualizado no mes: vale o par oficial
            sem_pg = (v1 / v0 - 1 < -0.01) & ~(pg > 0.001 * v0)  # VNA caiu sem pagamento informado: fluxo incompleto, fica a regra do carrego
            okp = np.isfinite(rs) & (rs > -0.25) & (rs < 0.3) & ~parado & ~sem_pg
            oko = ~okp & ~sem_pg & np.isfinite(ofi) & (ofi > -0.5) & (ofi < 0.3)
            rr, ss = np.asarray(r, float).copy(), src.copy()
            ix = np.flatnonzero(sc)
            rr[ix[okp]], ss[ix[okp]] = rs[okp], "preço + pagamento (securitizadora, CVM)"
            rr[ix[oko]], ss[ix[oko]] = ofi[oko], "série oficial ao par (securitizadora, CVM)"
            r, src = pd.Series(rr, index=w.index), ss
        deb = w.casa.isin(["DEB", "DEBIN"]).values
        if deb.any():
            pag = np.full(len(w), np.nan)
            pag[deb] = ctx.pagto_snd(w.codigo.fillna("").values[deb], p, m, pu0[deb])
            rt = (pu1 + pag) / pu0 - 1
            # preco ja corrigido entre fundos: marcacao a mercado entra direto; queda > 25% no mes so se persistir
            ok = np.isfinite(rt) & (rt < 0.3) & ((rt > -0.25) | ((rt > -0.8) & persiste(rt)))
            r = r.mask(ok, rt)
            src[np.asarray(ok)] = "preço + pagamento (SND)"
        # titulos publicos: preco do Tesouro + cupom semestral no mes de pagamento (NTN-B 6% a.a., NTN-F 10% a.a.)
        tpm = w.casa.eq("Títulos Públicos").values & np.isfinite(pu0) & np.isfinite(pu1)
        if tpm.any():
            mv = pd.to_datetime(w.venc).dt.month.values
            mm = int(m[5:])
            cup = np.where((w.tp.values == "NTN-B") & ((mm == mv) | (mm == (mv + 5) % 12 + 1)), 1.06 ** 0.5 - 1,
                           np.where((w.tp.values == "NTN-F") & np.isin(mm, [1, 7]), 1.10 ** 0.5 - 1, 0.0))
            rtp = (pu1 * (1 + 0) + pu0 * cup) / pu0 - 1
            okt = tpm & np.isfinite(rtp) & (rtp > -0.2) & (rtp < 0.2)
            r = r.mask(okt, rtp)
            src[np.asarray(okt)] = "Tesouro + cupom"
        fid = w.casa.eq("FIDC").values
        if fid.any():
            rf = np.full(len(w), np.nan)
            rf[fid] = [ctx.fidc_ret(c, m, pu) for c, pu in zip(w.codigo.fillna("").values[fid], w.pu.values[fid].astype(float))]
            okf = np.isfinite(rf) & (rf > -0.8) & (rf < 0.3)
            r = r.mask(okf, rf)          # retorno oficial da serie (cota + amortizacao)
            src[okf] = "FIDC: rentabilidade oficial (CVM)"
        # renda variavel (acoes, FII, FIAGRO, ETF, FIP): preco de mercado sem filtro; FII/FIAGRO + dividendo estimado
        usd_r = w.casa.isin(USD).values & np.isfinite(rp.values) & (rp.values > -0.1) & (rp.values < 0.1)   # > 10% no mes = erro de unidade
        r = r.mask(usd_r, rp.values)                     # em dolar: preco real em reais; o futuro de dolar compensa o cambio
        src[usd_r] = "preço em R$"
        rv = w.casa.isin(SEM_SPREAD).values & ~w.casa.eq("Confidencial").values & np.isfinite(rp.values)
        if rv.any():
            div = np.where(w.casa.isin(["FII", "FIAGRO"]).values, 0.75 * cdi_m, 0.0)
            okv = rv & (rp.values > -0.6) & (rp.values < 0.6)
            r = r.mask(okv, rp.values + div)
            src[okv] = "preço de mercado"
        cx = w.casa.isin(CAIXA)
        r[cx] = cdi_m
        src[cx.values] = "CDI"
        if CAPTURA is not None:
            CAPTURA[(nome, m, "retorno")] = w.assign(r_=np.asarray(r, float), rp_=rp.values, pu0_=pu0, pu1_=pu1, esp_=esperado.values)
        ativo_c = w.perc_pl * r
        # 2) derivativos -> ativos protegidos
        ip = w[w.ipca & w.perc_pl.gt(0)]
        dur_ipca = wavg(ip.duration, ip.perc_pl) or 3.0
        pr = w[w.pre & w.perc_pl.gt(0)]
        dur_pre = wavg(pr.duration, pr.perc_pl) or 1.0
        hedge = 0.0
        for c in derivativos_mes(gd0, p, ctx, dur_ipca, dur_pre):
            v = pnl_futuro(c, p, m, ctx) * c["fator"] / 100 * 100 if pd.notna(c["fator"]) else 0.0  # % do PL do fundo
            if not v or np.isnan(v):
                continue
            alvo = (w.ipca if c["tipo"] == "DAP" else w.pre if c["tipo"] == "DI1" and w.pre.any() else w.ipca
                    if c["tipo"] == "DI1" else w.usd if c["tipo"] in ("DOL", "WDO", "DDI") else pd.Series(False, index=w.index))
            alvo = alvo & w.perc_pl.gt(0)
            if not alvo.any():
                alvo = w.perc_pl.gt(0) & ~cx
            ativo_c[alvo] += v * w.perc_pl[alvo] / w.perc_pl[alvo].sum()
            hedge += v
        # 3) taxas do fundo (extrato CVM): administracao + performance estimada sobre o benchmark
        adm, perf, par, pct = ctx.taxas.get(cnpj, (0.0, 0.0, "", 100.0))
        tx = adm / 12
        if perf and "CDI" in par.upper():                # performance sobre o retorno liquido da cota acima do benchmark
            f = perf / 100
            tx += max(0.0, float(real[m]) * 100 - pct / 100 * cdi_m * 100) * f / (1 - f)
        # 4) resto (negociacao, marcacao fora do esperado, taxas de fundos investidos) -> categorias que nao sao caixa
        resid = float(real[m]) * 100 - ativo_c.sum() + tx
        cr = w.perc_pl.gt(0) & ~cx
        if not cr.any():
            cr = w.perc_pl.gt(0)
        ativo_c[cr] += resid * w.perc_pl[cr] / w.perc_pl[cr].sum()
        atr_meses.append({"mes": m, "real": float(real[m]) * 100, "cdi": float(ctx.cdi_m.get(m, np.nan)) * 100,
                          "hedge": hedge, "resid": resid, "taxas": tx, "preench": preench})
        jm = len(atr_meses) - 1
        por_cat = ativo_c.groupby(w.casa).sum()
        if tx:
            por_cat["Taxas do fundo"] = -tx
        peso = w.groupby("casa").perc_pl.sum()
        for c in set(por_cat.index) | set(cats_atr):
            cats_atr.setdefault(c, {"contrib": [None] * jm, "peso": [None] * jm})
            cats_atr[c]["contrib"].append(float(por_cat.get(c, 0.0)))
            cats_atr[c]["peso"].append(float(peso.get(c, 0.0)))
        sel_ = (~cx & ~w.casa.eq("Confidencial")).values
        nc = w[sel_].assign(c_=ativo_c[sel_], rw_=np.asarray(r, float)[sel_] * w.perc_pl.values[sel_], src_=src[sel_])   # caixa fora do ranking
        pa = nc.groupby(["rotulo", "casa"]).agg(v=("c_", "sum"), p=("perc_pl", "sum"), rw=("rw_", "sum"), src=("src_", "first"),
                                                gr=("grupo", "first"), cod=("codigo", "first"), em=("emissor", "first"))
        for (a, c), x in pa.iterrows():
            if abs(x.v) > 1e-5:
                contrib.append([id_ativo(a, c, x.gr, x.cod, x.em), jm, round(float(x.v), 6), round(float(x.p), 4),
                                round(float(x.rw / x.p) * 100, 4) if x.p > 0 else None, x.src])
    for c in cats_atr.values():
        for key in ("contrib", "peso"):
            c[key] += [None] * (len(atr_meses) - len(c[key]))
    tab = pd.concat(tab_ativos, ignore_index=True) if tab_ativos else pd.DataFrame()
    return limpo({"nome": nome, "cnpj": cnpj, "carteira": cart_meses,
                  "retorno": {"meses": atr_meses, "categorias": cats_atr,
                              "ativos": sorted(ativos.values(), key=lambda x: x["i"]), "contrib": contrib}}), tab


# ------------------------------------------------------------------ pre-carga de todos os fundos
def pasta_dia():
    return lt.CACHE / "painel" / str(date.today())


def precarrega(log=None):
    escreve = log or PRE["log"].append
    PRE["status"] = "rodando"
    try:
        fundos = [(n, lt.cnpj_of(x), True) for n, x in lt.NOSSOS_FUNDOS.items()]
        vistos = {c for _, c, _ in fundos}
        for n, x in lt.PEERS.items():                  # peer repetido ou que e nosso: entra uma vez so, como nosso
            c = lt.cnpj_of(x)
            if c not in vistos:
                fundos.append((n, c, False))
                vistos.add(c)
        dest = pasta_dia()
        chave = "|".join(sorted(c for _, c, _ in fundos)) + "|" + lt.DESDE + f"|m{METODO}"
        if (dest / "visao.json").exists() and json.loads((dest / "visao.json").read_text(encoding="utf-8")).get("chave") == chave:
            escreve("Usando o painel já processado hoje.")
            PRE.update(status="ok", pasta=str(dest))
            return
        dest.mkdir(parents=True, exist_ok=True)
        for velho in (lt.CACHE / "painel").glob("*"):
            if velho != dest and velho.is_dir():
                for f in velho.glob("*"):
                    f.unlink(missing_ok=True)
                velho.rmdir()
        escreve(f"Baixando/atualizando dados da CVM de {len(fundos)} fundos desde {lt.DESDE}...")
        workers = max(1, min(4, (os.cpu_count() or 2)))
        meses, dia = lt.prepara(sorted(c for _, c, _ in fundos), lt.DESDE, workers=workers, log=escreve)
        ctx = Contexto(escreve, meses)
        escreve("Carregando carteiras...")
        carga = lt.carrega_pos(meses)
        ctx.precos_ref(carga[0])
        pl = carga[1]
        nomes = {n: c for n, c, _ in fundos}
        d, distrib = lt.distribuicoes(lt.diario(nomes, dia))     # cota de retorno total (soma rendimentos distribuidos)
        for n_, (k_, med_) in distrib.items():
            escreve(f"   {n_}: distribui rendimentos ({k_} vezes, ~{med_:.2f}% por vez) - cota ajustada para retorno total")
        rent = lt.rentabilidade(d, ctx.cdi_m) if len(d) else pd.DataFrame(columns=["fundo", "mes", "retorno"])
        for i, (n, c, _) in enumerate(fundos, 1):
            escreve(f"[{i}/{len(fundos)}] {n}: look-through, carteira e atribuição...")
            fid = re.sub(r"\D", "", c)
            try:                                 # um fundo com problema nao derruba os outros
                df = lt.lookthrough({n: c}, meses, carga)
                if not len(df):
                    escreve(f"   {n}: a CVM não tem carteira deste CNPJ no período (só cota/retorno).")
                    continue
                det, tab = detalhe(n, c, df, ctx, pl, rent)
            except Exception as e:
                escreve(f"   {n}: pulado ({type(e).__name__}: {e})")
                continue
            (dest / (fid + ".json")).write_text(json.dumps(det), encoding="utf-8")
            tab.to_parquet(dest / (fid + "_ativos.parquet"), index=False)
            df.drop(columns=["chave", "fator"], errors="ignore").to_parquet(dest / (fid + ".parquet"), index=False)
        # visao geral: cotas diarias de todos (para retorno em qualquer periodo, consolidado e risco x retorno)
        datas = sorted(d.data.unique())
        idx = pd.DatetimeIndex(datas)
        q = d.pivot_table(index="data", columns="fundo", values="VL_QUOTA").reindex(idx)
        plf = d.pivot_table(index="data", columns="fundo", values="VL_PATRIM_LIQ").reindex(idx)
        cdi_idx = (1 + ctx.cdi_d.reindex(idx).fillna(0)).cumprod()
        cores, k_nosso = {}, 0
        for n, c, nosso in fundos:
            cores[n] = NOSSAS_CORES[k_nosso % len(NOSSAS_CORES)] if nosso else "#9aa0a6"
            k_nosso += nosso
        visao = {"chave": chave, "gerado": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"), "desde": lt.DESDE,
                 "ult_carteira": max(meses) if meses else None, "anbima": ctx.an_dias[-1] if ctx.an_dias else None,
                 "datas": [t.strftime("%Y-%m-%d") for t in idx], "cdi": cdi_idx.tolist(),
                 "fundos": [{"nome": n, "cnpj": c, "id": re.sub(r"\D", "", c), "nosso": nosso, "cor": cores[n],
                             "tem_carteira": (dest / (re.sub(r"\D", "", c) + ".json")).exists(),
                             "distribui": distrib.get(n),
                             "cota": q[n].tolist() if n in q else [], "pl": float(plf[n].dropna().iloc[-1]) / 1e6 if n in plf and plf[n].notna().any() else None}
                            for n, c, nosso in fundos],
                 "cat_cores": CAT_CORES}
        (dest / "visao.json").write_text(json.dumps(limpo(visao)), encoding="utf-8")
        escreve("Pronto.")
        PRE.update(status="ok", pasta=str(dest))
    except Exception as e:
        import traceback
        PRE.update(status="erro", erro=f"{type(e).__name__}: {e}")
        escreve("ERRO: " + PRE["erro"])
        escreve(traceback.format_exc()[-1500:])


# ====================================================================== painel em HTML (arquivo unico, sem servidor)
LIMITE_CONF = 5.0                          # % do PL confidencial acima do qual o mes nao conta como "carteira aberta"
MESES_ATIVOS = 3                           # meses abertos com a lista de ativos no HTML (por fundo)
FORA_PERF = {"Caixa", "Títulos Públicos", "Ajuste", "Confidencial", "Taxas do fundo", lt.AJUSTE}
N_DETALHE = 1800                           # ativos com historico completo no HTML (clique)
SRC = {"preço + pagamento (securitizadora, CVM)": "Q", "série oficial ao par (securitizadora, CVM)": "O", "preço + pagamento (SND)": "S", "preço do fundo": "P", "carrego + curva (sem preço)": "E", "FIDC: rentabilidade oficial (CVM)": "F",
       "Tesouro + cupom": "T", "preço em R$": "U", "preço de mercado": "M", "preço do fundo (perda que persiste)": "L", "CDI": "C"}


def _r(x, n=2):
    try:
        x = float(x)
    except (TypeError, ValueError):
        return None
    return round(x, n) if math.isfinite(x) else None


def _chave(codigo, rotulo):
    """Mesmo ativo em fundos diferentes: codigo (debenture/CRI/ISIN/CNPJ de fundo) ou, sem codigo, o rotulo."""
    c = str(codigo or "").strip()
    return c.upper() if c and c.lower() not in ("none", "nan") else re.sub(r"\s+", " ", str(rotulo)).upper().strip()


def _ret_mensal(datas, serie):
    s = pd.Series(serie, index=pd.DatetimeIndex(datas), dtype=float).dropna()
    if s.empty:
        return {}
    m = s.groupby(s.index.to_period("M")).last()
    return {str(p): float(v) for p, v in (m / m.shift(1) - 1).dropna().items()}


def _metricas(r, cdi):
    """r, cdi: % ao mes alinhados (NaN = sem dado). Acumulado, % do CDI, vol anual, Sharpe (excesso sobre o CDI)."""
    ok = ~np.isnan(r)
    n = int(ok.sum())
    if n == 0:
        return [None] * 5
    a = (np.prod(1 + r[ok] / 100) - 1) * 100
    c = (np.prod(1 + cdi[ok] / 100) - 1) * 100
    vol = np.std(r[ok], ddof=1) * np.sqrt(12) if n >= 6 else np.nan
    ex = r[ok] - cdi[ok]
    sh = ex.mean() / np.std(r[ok], ddof=1) * np.sqrt(12) if n >= 6 and np.std(r[ok], ddof=1) > 1e-9 else np.nan
    return [_r(a), _r(a / c * 100 if c > 0 else np.nan, 0), _r(vol), _r(sh), n]


def exporta_html(pasta=None, saida=None, log=print):
    """Le o painel processado (pasta do dia) e grava um HTML unico com tudo embutido (abre com dois cliques, sem servidor)."""
    pasta = Path(pasta or PRE.get("pasta") or pasta_dia())
    saida = Path(saida or Path.cwd() / "painel_fundos.html")
    saida.parent.mkdir(parents=True, exist_ok=True)
    V = json.loads((pasta / "visao.json").read_text(encoding="utf-8"))
    cdi_m = _ret_mensal(V["datas"], V["cdi"])
    ult_dia = pd.Timestamp(V["datas"][-1])                               # mes corrente so conta quando fechar
    ult_completo = str(ult_dia.to_period("M") if ult_dia == ult_dia + pd.offsets.BMonthEnd(0) else ult_dia.to_period("M") - 1)
    fundos, ativos, contrib, pos = [], [], [], []
    for f in V["fundos"]:
        jf = pasta / (f["id"] + ".json")
        F = json.loads(jf.read_text(encoding="utf-8")) if jf.exists() else None
        carteira, abertos = [], []
        for c in (F or {}).get("carteira", []):
            aberto = (c.get("conf") or 0) < LIMITE_CONF
            carteira.append({"m": c["mes"], "pl": _r(c.get("pl_mm"), 1), "conf": _r(c.get("conf"), 1), "ab": aberto,
                             "sp": _r(c["total"].get("spread")), "du": _r(c["total"].get("duration")),
                             "cats": {k["categoria"]: [_r(k["perc"], 2), _r(k.get("spread")), _r(k.get("duration"))] for k in c["cats"]}})
            if aberto:
                abertos.append(c["mes"])
        ret = (F or {}).get("retorno", {})
        meses_ret = [m["mes"] for m in ret.get("meses", [])]
        lista = ret.get("ativos", [])
        for c_ in ret.get("contrib", []):          # [ativo, mes, contribuicao (pp do PL), peso (% PL), retorno do ativo (%), fonte]
            i_a, jm, v, p = c_[:4]
            a = lista[i_a] if i_a < len(lista) else None
            if a and a.get("categoria") not in FORA_PERF and jm < len(meses_ret):
                contrib.append((f["id"], _chave(a.get("codigo"), a.get("ativo")), a.get("ativo"), a.get("categoria"), a.get("grupo") or "",
                                a.get("emissor") or "", meses_ret[jm], v, p, c_[4] if len(c_) > 4 else None, SRC.get(c_[5], "?") if len(c_) > 5 else "?"))
        cota = f.get("cota", [])
        rm = _ret_mensal(V["datas"], cota) if cota else {}
        ult_cota = next((V["datas"][i] for i in range(len(cota) - 1, -1, -1) if cota[i] is not None), None) if cota else None
        fundos.append({"id": f["id"], "nome": f["nome"], "cnpj": f["cnpj"], "nosso": f["nosso"], "pl": _r(f.get("pl"), 0),
                       "distribui": bool(f.get("distribui")), "ult_cota": ult_cota, "ret": {k: _r(v * 100, 3) for k, v in rm.items()},
                       "carteira": carteira, "abertos": abertos, "ret_meses": meses_ret,
                       "ret_cats": {k: [_r(x, 3) for x in v.get("contrib", [])] for k, v in ret.get("categorias", {}).items()}})
        ja = pasta / (f["id"] + "_ativos.parquet")
        if ja.exists():
            t = pd.read_parquet(ja)
            t = t[~t.categoria.isin(["Confidencial"]) & (t.perc.abs() >= .005)]
            t = t.assign(chave=[_chave(c, r) for c, r in zip(t.codigo, t.rotulo)], id=f["id"])
            pos.append(t.reindex(columns=["id", "mes", "chave", "rotulo", "categoria", "grupo", "emissor", "perc", "pct_par", "spread", "duration"]))
            pls = {c["m"]: c["pl"] for c in carteira}
            for r in t[t.mes.isin(abertos[-MESES_ATIVOS:])].itertuples():
                ativos.append([f["id"], r.mes, r.chave, str(r.rotulo), r.categoria, str(r.grupo or ""), str(r.emissor or ""),
                               _r(r.perc, 3), _r((r.perc or 0) / 100 * (pls.get(r.mes) or 0), 2), _r(r.spread), _r(r.duration), ""])
    nosso = {f["id"]: f["nosso"] for f in fundos}
    ult = {f["id"]: f["abertos"][-1] for f in fundos if f["abertos"]}
    pls_mes = {(f["id"], c["m"]): c["pl"] for f in fundos for c in f["carteira"]}
    P = pd.concat(pos, ignore_index=True) if pos else pd.DataFrame(columns=["id", "mes", "chave", "rotulo", "categoria", "grupo", "emissor", "perc"])
    P = P.groupby(["id", "mes", "chave"], as_index=False).agg(rotulo=("rotulo", "first"), categoria=("categoria", "first"),
                                                              grupo=("grupo", "first"), emissor=("emissor", "first"), perc=("perc", "sum"),
                                                              pct_par=("pct_par", "median"), spread=("spread", "median"), duration=("duration", "median"))
    P["nosso"] = P.id.map(nosso)
    P["rs"] = [(p or 0) / 100 * (pls_mes.get((i, m)) or 0) for i, m, p in zip(P.id, P.mes, P.perc)]
    U = P[[ult.get(i) == m for i, m in zip(P.id, P.mes)] & ~P.categoria.isin(FORA_PERF)]     # ultima carteira aberta de cada fundo
    temos = set(U[U.nosso].chave)
    n_nossos_ = U[U.nosso].id.nunique() or 1
    nossos_grupos = U[U.nosso].groupby("grupo").perc.sum() / n_nossos_
    # ---------------- exposicoes dos peers fora das nossas carteiras
    overlap = [{"chave": ch, "rotulo": g.rotulo.iloc[0], "categoria": g.categoria.iloc[0], "grupo": g.grupo.iloc[0], "emissor": g.emissor.iloc[0],
                "n": int(g.id.nunique()), "peers": sorted(g.id.unique().tolist()), "perc_med": _r(g.groupby("id").perc.sum().mean()),
                "rs": _r(g.rs.sum(), 1), "spread": _r(g.spread.median()), "duration": _r(g.duration.median()), "temos": ch in temos,
                "grupo_nosso": _r(nossos_grupos.get(g.grupo.iloc[0], 0.0))} for ch, g in U[~U.nosso].groupby("chave")]
    n_peers = U[~U.nosso].id.nunique() or 1
    grupos = [{"grupo": gr, "peers_perc": _r(g[~g.nosso].perc.sum() / n_peers), "nossos_perc": _r(g[g.nosso].perc.sum() / n_nossos_),
               "n_peers": int(g[~g.nosso].id.nunique()), "n_nossos": int(g[g.nosso].id.nunique()), "rs_peers": _r(g[~g.nosso].rs.sum(), 1),
               "categorias": sorted(g.categoria.unique().tolist())} for gr, g in U.groupby("grupo") if gr]
    # ---------------- performance de ativos (todos os fundos, todos os meses): retorno mensal = contribuicao / peso
    meses = sorted(m for m in cdi_m if m <= ult_completo and m >= lt.DESDE)
    mi = {m: i for i, m in enumerate(meses)}
    cdi_v = np.array([cdi_m[m] * 100 for m in meses])
    C = pd.DataFrame(contrib, columns=["id", "chave", "rotulo", "categoria", "grupo", "emissor", "mes", "v", "p", "r", "src"])
    C = C[C.mes.isin(mi) & (C.p > 0.01) & C.r.notna()]
    # retorno do ativo no mes = retorno total do proprio ativo (preco + pagamentos; sem hedge/residuo do fundo), medio ponderado entre fundos
    agg = C.assign(rp=C.r * C.p).sort_values("p").groupby(["chave", "mes"]).agg(rp=("rp", "sum"), p=("p", "sum"), src=("src", "last")).reset_index()
    agg["r"] = agg.rp / agg.p
    # CRI/CRA/NC/CCB: queda grande de preco que fica pode ser amortizacao ou perda (fluxo nao publico): no retorno fica o carrego
    # (CDI + spread do papel) e o mes e marcado "A" com a queda observada, para o usuario conferir
    sp_m = P.groupby(["chave", "mes"]).spread.median()
    ev = agg.src.eq("L")
    agg["queda"] = agg.r.where(ev)
    car = [((1 + cdi_m.get(m, 0)) * (1 + (sp_m.get((c, m), np.nan) if np.isfinite(sp_m.get((c, m), np.nan)) else 2.0) / 100) ** (1 / 12) - 1) * 100
           for c, m in zip(agg.chave[ev], agg.mes[ev])]
    agg.loc[ev, "r"] = car
    agg.loc[ev, "src"] = "A"
    agg = agg[agg.r.between(-60, 60)]
    info = C.groupby("chave").agg(rotulo=("rotulo", "last"), categoria=("categoria", "last"), grupo=("grupo", "last"),
                                  emissor=("emissor", "last"), n_ever=("id", "nunique"))
    ult_pos = U.groupby("chave").agg(n_hoje=("id", "nunique"), perc_med=("perc", "mean"), rs=("rs", "sum"),
                                     spread=("spread", "median"), duration=("duration", "median"),
                                     nossos=("nosso", "sum"))
    perf, series, fontes, eventos = [], {}, {}, {}
    for ch, g in agg.groupby("chave"):
        if g.queda.notna().any():
            eventos[ch] = {mi[m]: _r(q) for m, q in zip(g.mes, g.queda) if np.isfinite(q)}
        r = np.full(len(meses), np.nan)
        r[[mi[m] for m in g.mes]] = g.r.values
        fs = [" "] * len(meses)
        for m_, s_ in zip(g.mes, g.src):
            fs[mi[m_]] = s_
        fontes[ch] = "".join(fs)
        if np.isfinite(r).sum() < 3:
            continue
        i0 = info.loc[ch]
        u = ult_pos.loc[ch] if ch in ult_pos.index else None
        w12, w36 = _metricas(r[-12:], cdi_v[-12:]), _metricas(r[-36:], cdi_v[-36:])
        tudo = _metricas(r, cdi_v)
        ult_r = [_r(x) for x in r[-3:]]
        perf.append({"chave": ch, "rotulo": i0.rotulo, "categoria": i0.categoria, "grupo": i0.grupo, "emissor": i0.emissor,
                     "w12": w12, "w36": w36, "wt": tudo, "ult3": ult_r, "n_ever": int(i0.n_ever),
                     "n_hoje": int(u.n_hoje) if u is not None else 0, "perc_med": _r(u.perc_med) if u is not None else None,
                     "rs": _r(u.rs, 1) if u is not None else None, "spread": _r(u.spread) if u is not None else None,
                     "duration": _r(u.duration) if u is not None else None, "temos": ch in temos})
        series[ch] = r
    perf.sort(key=lambda x: -(x["n_ever"] * (x["wt"][4] or 0)))
    det_ch = ({a[2] for a in ativos} | {p["chave"] for p in perf[:N_DETALHE]}       # todo ativo de carteira aberta + os mais relevantes
              | {o["chave"] for o in sorted(overlap, key=lambda o: -o["n"])[:500]})
    # historico do clique: retorno mensal, % PL em cada fundo, spread e PU informados pelos fundos, ANBIMA, ultimos negocios B3
    Pd = P[P.chave.isin(det_ch) & P.mes.isin(set(meses[-36:]))]          # % PL e spread: ultimos 36 meses
    pl_hist = {ch: {fid: [[mi[m], _r(p, 3)] for m, p in zip(gg.mes, gg.perc)] for fid, gg in g.groupby("id")} for ch, g in Pd.groupby("chave")}
    med = Pd.groupby(["chave", "mes"]).agg(sp=("spread", "median"), par=("pct_par", "median"), du=("duration", "median")).reset_index()
    med_h = {ch: {"m": [mi[m] for m in g.mes], "sp": [_r(x) for x in g.sp], "par": [_r(x, 2) for x in g.par], "du": [_r(x) for x in g.du]}
             for ch, g in med.groupby("chave")}
    try:
        _, an = lt.anbima_debentures()
        an = an[an.codigo.str.strip().str.upper().isin(det_ch)] if len(an) else an
        tipo_tx = {"DI +": "CDI+ (%)", "% do DI": "% do CDI", "IPCA +": "IPCA+ (%)", "IGP-M +": "IGP-M+ (%)", "Outros": "pré (%)"}
        anb = {ch: {"d": g.data.dt.strftime("%Y-%m-%d").tolist(), "tx": [_r(x, 4) for x in g.taxa_ind], "par": [_r(x, 3) for x in g.pct_par],
                    "ind": tipo_tx.get(str(g.tipo.iloc[-1]), str(g.indice.iloc[-1]))}
               for ch, g in an.sort_values("data").groupby(an.codigo.str.strip().str.upper())} if len(an) else {}
    except Exception as e:
        log(f"   ANBIMA: sem histórico ({type(e).__name__})")
        anb = {}
    det = {ch: {"r": [_r(x, 3) for x in series[ch]] if ch in series else None, "s": fontes.get(ch), "ev": eventos.get(ch), "pl": pl_hist.get(ch, {}),
                "h": med_h.get(ch), "anb": anb.get(ch)} for ch in det_ch}
    # ---------------- melhores ativos de cada fundo (12 meses): retorno e contribuicao em bps
    f_top = {}
    m12 = set(meses[-12:])
    for fid, g in C[C.mes.isin(m12)].groupby("id"):
        x = g.groupby("chave").agg(rotulo=("rotulo", "last"), categoria=("categoria", "last"), v=("v", "sum"), p=("p", "mean"), n=("mes", "nunique"))
        g = g.assign(r=g.r.where(g.src.ne("L"), np.nan))
        g = g.assign(r=g.r.fillna(g.chave.map(lambda c: np.nan)).fillna(g.mes.map(lambda m: ((1 + cdi_m.get(m, 0)) * 1.02 ** (1 / 12) - 1) * 100)))
        rr = g.groupby("chave").r.apply(lambda s: (np.prod(1 + (s / 100).clip(-.6, .6)) - 1) * 100)      # retorno total do ativo, 12m
        x = x.join(rr.rename("ret"))
        x = x.loc[x.v.abs().sort_values(ascending=False).index[:60].union(x.ret.sort_values(ascending=False).index[:40])]
        f_top[fid] = [[ch, r.rotulo, r.categoria, _r(r.ret), _r(r.v * 100, 1), _r(r.p), int(r.n)] for ch, r in x.iterrows()]
    dados = {"gerado": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"), "cdi": {k: _r(v * 100, 3) for k, v in cdi_m.items()},
             "ult_mes_completo": ult_completo, "meses": meses, "fundos": fundos, "ativos": ativos, "overlap": overlap, "grupos": grupos,
             "perf": perf, "det": det, "f_top": f_top, "limite_conf": LIMITE_CONF, "anb_ini": min((v["d"][0] for v in anb.values()), default=None),
             "preench": {f["id"]: [m["mes"] for m in json.loads((pasta / (f["id"] + ".json")).read_text(encoding="utf-8"))["retorno"]["meses"]
                                   if m.get("preench")] for f in fundos if (pasta / (f["id"] + ".json")).exists()}}
    pagina = PAGINA_HTML.replace("/*PLOTLY*/", _plotly_js()).replace("/*DADOS*/", json.dumps(limpo(dados), ensure_ascii=False))
    saida.write_text(pagina, encoding="utf-8")
    log(f"HTML: {saida} ({saida.stat().st_size / 1e6:.1f} MB; {len(fundos)} fundos, {len(perf)} ativos com desempenho, "
        f"{len(det)} com histórico)")
    return saida


def _plotly_js():
    """Plotly embutido (o HTML funciona offline). Usa o pacote plotly se houver; senao baixa uma vez e guarda no cache."""
    try:
        import plotly.offline
        return plotly.offline.get_plotlyjs()
    except ImportError:
        f = lt.CACHE / "plotly.min.js"
        if not f.exists():
            f.write_bytes(requests.get("https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js", timeout=120).content)
        return f.read_text(encoding="utf-8")


PAGINA_HTML = r'''<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Painel de Fundos</title>
<style>
:root{--t:#16191f;--s:#5d6470;--b:#e4e6ea;--f:#f5f6f8;--nosso:#eb6834;--peer:#9aa0a6;--pos:#1f7a4d;--neg:#b42318}
*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:var(--t);background:var(--f)}
header{background:#111;color:#fff;padding:14px 22px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;align-items:center}
header h1{margin:0;font-size:18px}header .s{color:#aab;font-size:11px}
nav{background:#fff;border-bottom:1px solid var(--b);padding:0 22px;display:flex;gap:2px;position:sticky;top:0;z-index:5;flex-wrap:wrap}
nav a{padding:12px 14px;cursor:pointer;font-weight:bold;color:var(--s);border-bottom:3px solid transparent}nav a.on{color:var(--t);border-color:var(--nosso)}
main{padding:16px 22px;max-width:1520px;margin:0 auto}.aba{display:none}.aba.on{display:block}
.cx{background:#fff;border:1px solid var(--b);border-radius:10px;padding:12px 16px;margin-bottom:12px}
.cx h2{font-size:12px;letter-spacing:1.3px;text-transform:uppercase;margin:0 0 4px}.sub{color:var(--s);font-size:11.5px;margin-bottom:8px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}@media(max-width:1000px){.g2{grid-template-columns:1fr}}
.ctl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px}
select,input{font:inherit;padding:6px 8px;border:1px solid var(--b);border-radius:6px;background:#fff}
.chip{border:1px solid var(--b);border-radius:14px;padding:4px 11px;cursor:pointer;font-weight:bold;color:var(--s);background:#fff;display:inline-block}.chip.on{background:var(--t);color:#fff;border-color:var(--t)}
table.t{width:100%;border-collapse:collapse;font-size:12px}table.t th{background:#111;color:#fff;padding:7px 6px;text-align:right;cursor:pointer;white-space:nowrap;font-size:10.5px;text-transform:uppercase}
table.t td{padding:6px;border-bottom:1px solid #eef0f2;text-align:right;white-space:nowrap}table.t th.l,table.t td.l{text-align:left}
table.t tr.nosso td:first-child{border-left:4px solid var(--nosso)}table.t tr.cl:hover td{background:#f3f8fb;cursor:pointer}table.t tr.det>td{background:#fafbfc;white-space:normal;text-align:left;padding:12px}
.ponto{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px}.tag{background:var(--nosso);color:#fff;border-radius:4px;padding:1px 5px;font-size:10px;font-weight:bold}
.pos{color:var(--pos)}.neg{color:var(--neg)}.mut{color:#999}.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:12px}
.k{background:#fff;border:1px solid var(--b);border-radius:10px;padding:10px 14px}.k .r{font-size:10px;letter-spacing:1px;color:var(--s);text-transform:uppercase}.k .v{font-size:19px;font-weight:bold;margin-top:3px}.k .x{font-size:11px;color:var(--s)}
.lk{color:#1f5f87;cursor:pointer;text-decoration:underline dotted}.scroll{overflow-x:auto}ul.nv{margin:0;padding-left:18px;line-height:1.7}
details.ms{position:relative;display:inline-block}details.ms summary{list-style:none;cursor:pointer;border:1px solid var(--b);border-radius:6px;padding:6px 10px;background:#fff}
details.ms .op{position:absolute;z-index:9;background:#fff;border:1px solid var(--b);border-radius:8px;padding:8px 12px;max-height:280px;overflow:auto;min-width:220px;box-shadow:0 6px 18px rgba(0,0,0,.12)}
details.ms label{display:block;padding:2px 0;white-space:nowrap}.mini{font-size:11px;color:var(--s);margin:4px 0 0}
</style></head><body>
<header><div><h1>Painel de Fundos — nossos × peers</h1><div class="s" id="hsub"></div></div>
</header>
<nav id="nav"></nav>
<main>
<section class="aba" id="a_geral">
  <div class="ctl"><span id="g_chips"></span><input id="g_busca" placeholder="buscar fundo"></div>
  <div class="cx"><h2>Todos os fundos</h2><div class="scroll"><table class="t" id="g_tab"></table></div></div>
  <div class="g2"><div class="cx"><h2>Retorno × risco (12 meses)</h2><div id="g_rr" style="height:380px"></div></div>
    <div class="cx"><h2>Spread × duration (última carteira aberta)</h2><div id="g_sd" style="height:380px"></div></div></div>
</section>
<section class="aba" id="a_fundo">
  <div class="ctl"><select id="f_sel" style="min-width:320px"></select><span id="f_info" class="sub"></span></div>
  <div class="kpis" id="f_kpis"></div>
  <div class="cx"><h2>Alocação por categoria (% do PL) — só meses abertos</h2><div id="f_aloc" style="height:360px"></div></div>
  <div class="cx"><h2>Spread CDI+ e duration da carteira</h2><div class="ctl"><details class="ms" id="f_sdcat"><summary>Categorias ▾</summary><div class="op"></div></details></div><div id="f_sd" style="height:360px"></div></div>
  <div class="cx"><h2>Retorno mensal × CDI e de onde veio (por categoria)</h2><div id="f_ret" style="height:380px"></div></div>
  <div class="cx"><h2>Carteira</h2><div class="ctl"><select id="f_mes"></select><select id="f_cat"></select><input id="f_busca" placeholder="buscar ativo/emissor"></div><div class="scroll"><table class="t" id="f_tab"></table></div></div>
  <div class="cx"><h2>Melhores ativos do fundo — últimos 12 meses</h2><div class="ctl"><span class="chip on" data-m="ret">Retorno do ativo</span><span class="chip" data-m="bps">Contribuição para o fundo (bps)</span><select id="f_topcat"></select></div><div class="scroll"><table class="t" id="f_top"></table></div></div>
  <div class="cx"><h2>Mudanças na carteira</h2><div class="sub" id="f_mud_sub"></div><div class="g2"><div><b>Entradas e aumentos</b><table class="t" id="f_mud_up"></table></div><div><b>Saídas e reduções</b><table class="t" id="f_mud_dn"></table></div></div></div>
</section>
<section class="aba" id="a_perf">
  <div class="ctl"><select id="v_jan"><option value="w12">Últimos 12 meses</option><option value="w36">Últimos 36 meses</option><option value="wt">Desde o início</option></select>
    <select id="v_cat"></select><label>em pelo menos <select id="v_min"><option>1</option><option selected>2</option><option>3</option><option>5</option><option>10</option></select> fundos</label>
    <label><input type="checkbox" id="v_nt"> só os que não temos</label><input id="v_busca" placeholder="buscar ativo/emissor/grupo"></div>
  <div class="cx"><h2>Performance de ativos — todas as carteiras, todos os meses</h2>
    <div class="scroll"><table class="t" id="v_tab"></table></div></div>
</section>
<section class="aba" id="a_peers">
  <div class="ctl"><select id="p_cat"></select><label>em pelo menos <select id="p_min"><option>1</option><option selected>2</option><option>3</option><option>5</option></select> peers</label>
    <label><input type="checkbox" id="p_so"> incluir o que já temos</label><label><input type="checkbox" id="p_gr"> só grupos sem exposição nossa</label><input id="p_busca" placeholder="buscar ativo/emissor/grupo"></div>
  <div class="cx"><h2>Exposições dos peers ausentes nas nossas carteiras</h2><div class="scroll"><table class="t" id="p_tab"></table></div></div>
  <div class="cx"><h2>Grupos econômicos: peso médio dos peers × nosso</h2><div class="scroll"><table class="t" id="p_gru"></table></div></div>
</section>

</main>
<script>/*PLOTLY*/</script>
<script>
const D=/*DADOS*/;
const $=s=>document.querySelector(s), F={}; D.fundos.forEach(f=>F[f.id]=f);
const PERF={};D.perf.forEach(p=>PERF[p.chave]=p);const OVL={};D.overlap.forEach(o=>OVL[o.chave]=o);
const br=(v,c=2)=>v==null||isNaN(v)?"—":Number(v).toLocaleString("pt-BR",{minimumFractionDigits:c,maximumFractionDigits:c});
const MS=["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"],mesBR=m=>m?MS[+m.slice(5,7)-1]+"/"+m.slice(2,4):"—",diaBR=d=>d?d.slice(8,10)+"/"+d.slice(5,7)+"/"+d.slice(2,4):"—";
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const cor=f=>f.nosso?"var(--nosso)":"var(--peer)", CC={};let ci=0;
const CORES={"DEB":"#2a78d6","DEBIN":"#7ba7d9","FIDC":"#e87ba4","LF":"#4a3aa7","LFSN":"#a61e4d","LFSC":"#8a5a00","CDB":"#9775fa","DPGE":"#5f3dc4","CRA":"#0b7285","CRI":"#e34948",
 "NC":"#1baf7a","CCB":"#74b816","FIAGRO":"#a9e34b","Caixa":"#adb5bd","Títulos Públicos":"#343a40","Bonds":"#eda100","Ações":"#f76707","FII":"#ffd43b","FIP":"#d9480f","ETF":"#fab005",
 "Fundos offshore":"#99e9f2","Cotas de Fundos":"#66d9e8","Confidencial":"#dee2e6","Outras aplicações":"#868e96","Taxas do fundo":"#495057","Outros":"#ced4da"};
const corCat=c=>CORES[c]||CC[c]||(CC[c]=["#3bc9db","#b197fc","#ff8787","#63e6be","#ffa94d","#748ffc","#c0eb75","#f783ac"][ci++%8]);
const LAY=x=>Object.assign({margin:{l:56,r:16,t:10,b:40},font:{size:11},legend:{orientation:"h",x:0,y:1.12,font:{size:10}},hovermode:"closest"},x||{});
const LAYR=x=>LAY(Object.assign({margin:{l:56,r:170,t:10,b:40},legend:{x:1.02,y:1,xanchor:"left",font:{size:10}}},x||{}));    // legenda ao lado: nunca em cima do grafico
const plot=(id,d,l)=>Plotly.newPlot(id,d,l,{displayModeBar:false,responsive:true});
const cls=v=>v==null?"":v>=0?"pos":"neg";
$("#hsub").textContent=`Gerado em ${D.gerado} • ${D.fundos.length} fundos (${D.fundos.filter(f=>f.nosso).length} nossos) • retorno até ${mesBR(D.ult_mes_completo)}`;
const ABAS=[["a_geral","Visão geral"],["a_fundo","Fundo"],["a_perf","Performance de ativos"],["a_peers","Exposições vs. peers"]];
$("#nav").innerHTML=ABAS.map(a=>`<a data-a="${a[0]}">${a[1]}</a>`).join("");
const vai=a=>{document.querySelectorAll("nav a").forEach(x=>x.classList.toggle("on",x.dataset.a==a));document.querySelectorAll(".aba").forEach(x=>x.classList.toggle("on",x.id==a));history.replaceState(null,"","#"+a);window.dispatchEvent(new Event("resize"))};
document.querySelectorAll("nav a").forEach(x=>x.onclick=()=>vai(x.dataset.a));
const abreFundo=id=>{$("#f_sel").value=id;fundo();vai("a_fundo");scrollTo(0,0)};
// ---------- retorno em janelas
const meses=Object.keys(D.cdi).sort().filter(m=>m<=D.ult_mes_completo), ultM=meses[meses.length-1];
function acum(f,n){const ms=meses.slice(-n);let r=1,c=1,k=0;for(const m of ms){if(f.ret[m]==null)continue;r*=1+f.ret[m]/100;c*=1+D.cdi[m]/100;k++}return k<ms.length*.8?null:[(r-1)*100,(c-1)*100]}
function vol(f,n){const v=meses.slice(-n).map(m=>f.ret[m]).filter(x=>x!=null);if(v.length<n*.8)return null;const mu=v.reduce((a,b)=>a+b)/v.length;return Math.sqrt(v.reduce((a,b)=>a+(b-mu)**2,0)/(v.length-1))*Math.sqrt(12)}
const pcdi=(f,n)=>{const a=acum(f,n);return a&&a[1]?a[0]/a[1]*100:null};
const ytdN=meses.filter(m=>m.slice(0,4)==ultM.slice(0,4)).length;
const ultAb=f=>{const m=f.abertos[f.abertos.length-1];return f.carteira.find(c=>c.m==m)};
function topCats(c){if(!c)return"—";return Object.entries(c.cats).filter(([k])=>!["Caixa","Ajuste"].includes(k)).sort((a,b)=>b[1][0]-a[1][0]).slice(0,3).map(([k,v])=>`${k} ${br(v[0],0)}%`).join(" · ")}
function ordena(L,o){return L.sort((a,b)=>{const x=a[o.k],y=b[o.k];if(x==null)return 1;if(y==null)return-1;return(x>y?1:x<y?-1:0)*(o.asc?1:-1)})}
function cabec(t,C,o,cb){t.querySelectorAll("th").forEach((h,i)=>h.onclick=()=>{const k=C[i][0];o.asc=o.k==k?!o.asc:false;o.k=k;cb()})}
const th=(C,o)=>"<tr>"+C.map(c=>`<th class="${c[2]||""}">${c[1]}${o&&o.k==c[0]?(o.asc?" ▲":" ▼"):""}</th>`).join("")+"</tr>";
// ---------- visão geral
let gFil="Todos";const gOrd={k:"r12",asc:false};
$("#g_chips").innerHTML=["Todos","Nossos","Peers"].map(c=>`<span class="chip${c==gFil?" on":""}" data-c="${c}">${c}</span>`).join(" ");
document.querySelectorAll("#g_chips .chip").forEach(e=>e.onclick=()=>{gFil=e.dataset.c;document.querySelectorAll("#g_chips .chip").forEach(x=>x.classList.toggle("on",x==e));geral()});
$("#g_busca").oninput=geral;
function geral(){const b=$("#g_busca").value.toLowerCase();
  const L=ordena(D.fundos.filter(f=>(gFil=="Todos"||(gFil=="Nossos")==f.nosso)&&f.nome.toLowerCase().includes(b)).map(f=>{const c=ultAb(f);
    return{f,nome:f.nome,pl:f.pl,ab:f.abertos[f.abertos.length-1]||"",m1:pcdi(f,1),m3:pcdi(f,3),ytd:pcdi(f,ytdN),r12:pcdi(f,12),r24:pcdi(f,24),vol:vol(f,12),sp:c?c.sp:null,du:c?c.du:null,cats:topCats(c)}}),gOrd);
  const C=[["nome","Fundo","l"],["pl","PL (R$ mm)"],["ab","Carteira aberta"],["m1",`Mês (${mesBR(ultM)})`],["m3","3 meses"],["ytd","No ano"],["r12","12 meses"],["r24","24 meses"],["vol","Vol 12m"],["sp","Spread CDI+"],["du","Duration"],["cats","Maiores categorias","l"]];
  $("#g_tab").innerHTML=th(C,gOrd)+L.map(r=>`<tr class="${r.f.nosso?"nosso":""}">`+C.map(c=>{const v=r[c[0]];
    if(c[0]=="nome")return`<td class="l"><span class="ponto" style="background:${cor(r.f)}"></span><span class="lk" data-id="${r.f.id}">${esc(v)}</span>${r.f.nosso?' <span class="tag">NOSSO</span>':""}</td>`;
    if(c[0]=="ab")return`<td>${mesBR(v)}</td>`;if(c[0]=="cats")return`<td class="l mut">${v}</td>`;if(c[0]=="pl")return`<td>${br(v,0)}</td>`;
    if(["m1","m3","ytd","r12","r24"].includes(c[0]))return`<td class="${v==null?"":v>=100?"pos":"neg"}">${v==null?"—":br(v,0)+"%"}</td>`;return`<td>${br(v)}</td>`}).join("")+"</tr>").join("");
  cabec($("#g_tab"),C,gOrd,geral);$("#g_tab").querySelectorAll(".lk").forEach(e=>e.onclick=()=>abreFundo(e.dataset.id));
  const pts=L.filter(r=>r.vol!=null&&acum(r.f,12)), cdi12=meses.slice(-12).reduce((a,m)=>a*(1+D.cdi[m]/100),1)*100-100;
  plot("g_rr",[false,true].map(n=>{const p=pts.filter(r=>r.f.nosso==n);return{x:p.map(r=>r.vol),y:p.map(r=>acum(r.f,12)[0]),text:p.map(r=>r.nome),mode:"markers",name:n?"nossos":"peers",
    marker:{color:n?"#eb6834":"#9aa0a6",size:n?12:9,line:{color:"#fff",width:1}},hovertemplate:"%{text}<br>retorno %{y:.2f}%<br>vol %{x:.2f}%<extra></extra>"}}),
    LAY({xaxis:{title:{text:"volatilidade anual (%)"}},yaxis:{title:{text:"retorno 12 meses (%)"}},shapes:[{type:"line",xref:"paper",x0:0,x1:1,y0:cdi12,y1:cdi12,line:{dash:"dash",color:"#333"}}]}));
  const q=L.filter(r=>r.sp!=null&&r.du!=null);
  plot("g_sd",[false,true].map(n=>{const p=q.filter(r=>r.f.nosso==n);return{x:p.map(r=>r.du),y:p.map(r=>r.sp),text:p.map(r=>r.nome),mode:"markers",name:n?"nossos":"peers",
    marker:{color:n?"#eb6834":"#9aa0a6",size:n?12:9,line:{color:"#fff",width:1}},hovertemplate:"%{text}<br>CDI+ %{y:.2f}%<br>duration %{x:.2f}<extra></extra>"}}),LAY({xaxis:{title:{text:"duration (anos)"}},yaxis:{title:{text:"spread CDI+ (% a.a.)"}}}))}
// ---------- detalhe de um ativo (clique em qualquer tabela de ativos)
let detN=0;
const FONTE={Q:"preço do fundo + juros/amortização pagos (informe da securitizadora, CVM)",O:"retorno oficial da série ao par (informe da securitizadora, CVM)",S:"preço do fundo + juros/amortização pagos (PU par do SND)",P:"preço informado pelo fundo",E:"carrego + curva (sem preço observável no mês)",
 F:"FIDC: rentabilidade oficial da série (CVM)",T:"Tesouro + cupom",U:"preço em R$ (sem hedge)",M:"preço de mercado",C:"CDI",
 A:"carrego (CDI + spread): o preço caiu forte — amortização ou perda? o fluxo deste papel não é público, confira"};
function detalheAtivo(tr,ch,ncol,ctx){const prox=tr.nextElementSibling;if(prox&&prox.classList.contains("det")){prox.remove();return}
  document.querySelectorAll("tr.det").forEach(x=>x.remove());const d=D.det[ch],p=PERF[ch]||{},o=OVL[ch]||{},id="dt"+(++detN);
  const n=document.createElement("tr");n.className="det";tr.after(n);
  if(!d){n.innerHTML=`<td colspan="${ncol}">Sem histórico detalhado para este ativo no arquivo (fica só para os ${Object.keys(D.det).length} mais relevantes).</td>`;return}
  const w=p.w12||[],u3=p.ult3||[];
  const k=[["Último mês",u3[2]==null?"—":br(u3[2])+"%",mesBR(D.meses[D.meses.length-1])],["3 meses",u3.length&&u3.every(x=>x!=null)?br((u3.reduce((a,x)=>a*(1+x/100),1)-1)*100)+"%":"—",""],
    ["12 meses",w[0]==null?"—":br(w[0])+"%",w[1]==null?"":br(w[1],0)+"% do CDI"],["Vol / Sharpe 12m",(w[2]==null?"—":br(w[2])+"%")+" / "+(w[3]==null?"—":br(w[3])),""],
    ["Spread / duration",(p.spread??o.spread)==null?"—":"CDI+ "+br(p.spread??o.spread),(p.duration??o.duration)==null?"":br(p.duration??o.duration)+" anos"],
    ["Fundos com o ativo",(p.n_hoje??o.n??"—")+"",`já esteve em ${p.n_ever??"—"} • ${(p.temos??o.temos)?"nós temos":"não temos"}`]];
  n.innerHTML=`<td colspan="${ncol}"><div class="kpis">${k.map(x=>`<div class="k"><div class="r">${x[0]}</div><div class="v">${x[1]}</div><div class="x">${x[2]}</div></div>`).join("")}</div>
   <div class="g2"><div><b>Retorno total mensal do ativo × CDI</b><div id="${id}a" style="height:270px"></div></div><div><b>Curva acumulada: ativo × CDI</b><div id="${id}b" style="height:270px"></div></div>
   <div><b>% do PL</b><div id="${id}c" style="height:260px"></div></div><div><b>Spread CDI+ e % do PU par</b><div id="${id}d" style="height:260px"></div></div>
   <div><b>ANBIMA: taxa indicativa e % do PU par (diário)</b><div id="${id}e" style="height:260px"></div></div><div class="mini" id="${id}f" style="align-self:center;line-height:1.6"></div></div></td>`;
  const M=D.meses,R=d.r||[],S=d.s||"",idx=M.map((m,i)=>i).filter(i=>R[i]!=null);if(!idx.length){$("#"+id+"a").innerHTML=$("#"+id+"b").innerHTML="<div class=\"mini\">Sem retorno mensal medido (ativo novo ou só em meses sem carteira anterior).</div>"}
  const EV=d.ev||{},fonte=i=>(FONTE[S[i]]||"—")+(EV[i]!=null?` (queda de preço observada: ${br(EV[i])}%)`:""),est=idx.filter(i=>S[i]=="E"||S[i]=="A").length,nev=Object.keys(EV).length;
  if(idx.length)plot(id+"a",[{x:idx.map(i=>mesBR(M[i])),y:idx.map(i=>D.cdi[M[i]]),type:"bar",name:"CDI",marker:{color:"#c9cdd3"},hovertemplate:"CDI %{x}: %{y:.2f}%<extra></extra>"},
    {x:idx.map(i=>mesBR(M[i])),y:idx.map(i=>R[i]),type:"bar",name:"ativo",marker:{color:idx.map(i=>S[i]=="A"?"#eda100":S[i]=="E"?"#9ec5e8":R[i]>=0?"#1f5f87":"#b42318")},
     text:idx.map(i=>fonte(i)),textposition:"none",hovertemplate:"ativo %{x}: %{y:.2f}%<br>%{text}<extra></extra>"}],LAY({barmode:"group",bargap:.15,yaxis:{title:{text:"% no mês"}},xaxis:{type:"category"}}));
  let a=100,c=100;const ca=[],cc=[];idx.forEach(i=>{a*=1+R[i]/100;c*=1+D.cdi[M[i]]/100;ca.push(a);cc.push(c)});
  if(idx.length)plot(id+"b",[{x:idx.map(i=>mesBR(M[i])),y:ca,name:"ativo",mode:"lines",line:{color:"#1f5f87",width:2.5},hovertemplate:"ativo: %{y:.1f}<extra>%{x}</extra>"},
    {x:idx.map(i=>mesBR(M[i])),y:cc,name:"CDI",mode:"lines",line:{color:"#111",dash:"dash"},hovertemplate:"CDI: %{y:.1f}<extra>%{x}</extra>"}],LAY({yaxis:{title:{text:"base 100"}},xaxis:{type:"category"}}));
  const ent=Object.entries(d.pl),meu=ent.filter(([fid])=>fid==ctx||F[fid].nosso),tl=meu.map(([fid,arr])=>({x:arr.map(q=>mesBR(M[q[0]])),y:arr.map(q=>q[1]),name:F[fid].nome+(fid==ctx&&!F[fid].nosso?" (este fundo)":""),
    mode:"lines+markers",line:{color:fid==ctx&&!F[fid].nosso?"#111":"#eb6834",width:2.5},marker:{size:4},hovertemplate:`${esc(F[fid].nome)}: %{y:.2f}%<extra>%{x}</extra>`}));
  if(!ctx||F[ctx].nosso){const pe=ent.filter(([fid])=>!F[fid].nosso),mm={};pe.forEach(([,arr])=>arr.forEach(q=>{(mm[q[0]]=mm[q[0]]||[]).push(q[1])}));
    const ks=Object.keys(mm).map(Number).sort((x,y)=>x-y);if(ks.length)tl.push({x:ks.map(q=>mesBR(M[q])),y:ks.map(q=>mm[q].reduce((x,y)=>x+y)/mm[q].length),name:"média dos peers que têm",mode:"lines",line:{color:"#9aa0a6",width:2,dash:"dot"},
      text:ks.map(q=>mm[q].length+" peers"),hovertemplate:"média peers: %{y:.2f}% (%{text})<extra>%{x}</extra>"})}
  if(tl.length)plot(id+"c",tl,LAY({yaxis:{title:{text:"% do PL"}},xaxis:{type:"category",categoryorder:"array",categoryarray:M.map(mesBR)}}));else $("#"+id+"c").innerHTML='<div class="mini">Nem este fundo nem os nossos têm o ativo.</div>';
  const h=d.h,temPar=h&&h.par.some(x=>x!=null);
  if(h)plot(id+"d",[{x:h.m.map(i=>mesBR(M[i])),y:h.sp,name:"spread CDI+ (esq.)",mode:"lines+markers",line:{color:"#eb6834",width:2}},
    ...(temPar?[{x:h.m.map(i=>mesBR(M[i])),y:h.par,name:"% do PU par (dir.)",mode:"lines+markers",line:{color:"#2a78d6",width:2,dash:"dot"},yaxis:"y2"}]:[])],
    LAY({yaxis:{title:{text:"spread CDI+ (% a.a.)"}},yaxis2:{title:{text:"% do PU par"},overlaying:"y",side:"right"},xaxis:{type:"category"},margin:{l:56,r:60,t:10,b:40}}));
  const an=d.anb,temTx=an&&an.tx.some(x=>x!=null),temAp=an&&an.par.some(x=>x!=null);
  if(an&&(temTx||temAp))plot(id+"e",[...(temTx?[{x:an.d,y:an.tx,name:`taxa indicativa (${an.ind}) (esq.)`,mode:"lines+markers",line:{color:"#1f5f87",width:2}}]:[]),
    ...(temAp?[{x:an.d,y:an.par,name:"% do PU par (dir.)",mode:"lines",line:{color:"#eb6834",width:2,dash:"dot"},yaxis:"y2"}]:[])],
    LAY({yaxis:{title:{text:"taxa (%)"}},yaxis2:{title:{text:"% PU par"},overlaying:"y",side:"right"},margin:{l:56,r:60,t:10,b:40}}));
  else $("#"+id+"e").innerHTML=`<div class="mini">Sem marcação ANBIMA.</div>`;
  $("#"+id+"f").innerHTML=nev?`<b style="color:#b26b00">${nev} mês(es) com queda forte de preço (barra amarela)</b>`:""}
// ---------- fundo
$("#f_sel").innerHTML=[...D.fundos].sort((a,b)=>(b.nosso-a.nosso)||a.nome.localeCompare(b.nome)).map(f=>`<option value="${f.id}">${f.nosso?"★ ":""}${esc(f.nome)}</option>`).join("");
$("#f_sel").onchange=fundo;["#f_mes","#f_cat","#f_busca"].forEach(s=>$(s).oninput=tabFundo);
const tOrd={k:"ret",asc:false};document.querySelectorAll("#a_fundo .chip[data-m]").forEach(e=>e.onclick=()=>{tOrd.k=e.dataset.m;tOrd.asc=false;document.querySelectorAll("#a_fundo .chip[data-m]").forEach(x=>x.classList.toggle("on",x==e));topFundo()});
$("#f_topcat").oninput=topFundo;
function fundo(){const f=F[$("#f_sel").value],c=ultAb(f);
  $("#f_info").innerHTML=`${f.cnpj} • última CDA ${mesBR(f.carteira.length?f.carteira[f.carteira.length-1].m:null)} • última carteira aberta ${mesBR(f.abertos[f.abertos.length-1])} • cota até ${diaBR(f.ult_cota)}${f.distribui?" • distribui rendimentos (retorno total)":""}`;
  const k=[["PL","R$ "+br(f.pl,0)+" mm",""],["12 meses",pcdi(f,12)==null?"—":br(pcdi(f,12),0)+"% do CDI",acum(f,12)?br(acum(f,12)[0])+"%":""],["No ano",pcdi(f,ytdN)==null?"—":br(pcdi(f,ytdN),0)+"% do CDI",""],
    ["Spread carteira",c?"CDI+ "+br(c.sp):"—",c?mesBR(c.m):""],["Duration",c?br(c.du)+" anos":"—",""],["Vol 12m",vol(f,12)==null?"—":br(vol(f,12))+"%",""]];
  $("#f_kpis").innerHTML=k.map(x=>`<div class="k"><div class="r">${x[0]}</div><div class="v">${x[1]}</div><div class="x">${x[2]}</div></div>`).join("");
  const ab=f.carteira.filter(x=>x.ab), cats=[...new Set(ab.flatMap(x=>Object.keys(x.cats)))].filter(k=>k!="Ajuste");
  cats.sort((a,b)=>Math.max(...ab.map(x=>(x.cats[b]||[0])[0]))-Math.max(...ab.map(x=>(x.cats[a]||[0])[0])));
  plot("f_aloc",cats.map(k=>({x:ab.map(x=>mesBR(x.m)),y:ab.map(x=>(x.cats[k]||[0])[0]),name:k,type:"bar",marker:{color:corCat(k),line:{width:0}},hovertemplate:`${k}: %{y:.1f}%<extra>%{x}</extra>`})),
    LAYR({barmode:"stack",bargap:0,yaxis:{ticksuffix:"%",title:{text:"% do PL"}},xaxis:{type:"category"}}));
  const op=$("#f_sdcat .op");op.innerHTML=["Total",...cats].map((k,i)=>`<label><input type="checkbox" value="${esc(k)}"${i==0?" checked":""}> ${k=="Total"?"<b>Carteira total</b>":esc(k)}</label>`).join("");
  op.querySelectorAll("input").forEach(x=>x.onchange=spdur);spdur();
  const ms=meses.slice(-24), pre=new Set((D.preench||{})[f.id]||[]), rot=m=>mesBR(m)+(pre.has(m)?"*":""), cr=Object.keys(f.ret_cats).filter(k=>k!="Confidencial"||ms.some(m=>{const i=f.ret_meses.indexOf(m);return i>=0&&Math.abs(f.ret_cats[k][i]||0)>.001}));
  const tr=cr.map(k=>({x:ms.map(rot),y:ms.map(m=>{const i=f.ret_meses.indexOf(m);return i<0?null:f.ret_cats[k][i]}),name:k,type:"bar",marker:{color:corCat(k)},hovertemplate:`${k}: %{y:.3f} pp<extra>%{x}</extra>`}));
  tr.push({x:ms.map(rot),y:ms.map(m=>f.ret[m]),name:"retorno da cota",mode:"lines+markers",line:{color:"#eb6834",width:3},hovertemplate:"cota: %{y:.2f}%<extra>%{x}</extra>"},
    {x:ms.map(rot),y:ms.map(m=>D.cdi[m]),name:"CDI",mode:"lines",line:{color:"#111",dash:"dash"},hovertemplate:"CDI: %{y:.2f}%<extra>%{x}</extra>"});
  plot("f_ret",tr,LAYR({barmode:"relative",yaxis:{ticksuffix:"%",title:{text:"% no mês"}}}));
  const mm=[...new Set(D.ativos.filter(a=>a[0]==f.id).map(a=>a[1]))].sort().reverse();
  $("#f_mes").innerHTML=mm.map(m=>`<option value="${m}">${mesBR(m)}</option>`).join("")||"<option>sem carteira aberta</option>";
  $("#f_cat").innerHTML="";tabFundo();topFundo();mudancas()}
function spdur(){const f=F[$("#f_sel").value],ab=f.carteira.filter(x=>x.ab),sel=[...document.querySelectorAll("#f_sdcat input:checked")].map(x=>x.value);
  $("#f_sdcat summary").textContent=`Categorias (${sel.length}) ▾`;const tr=[],SPC=["#eb6834","#2a78d6","#1baf7a","#e34948","#9775fa","#eda100"],DUC=["#111111","#0b7285","#5d6470","#8a5a00","#a61e4d","#3d348b"];
  sel.forEach((k,i)=>{const get=(x,j)=>k=="Total"?(j==1?x.sp:x.du):(x.cats[k]||[])[j],nm=k=="Total"?"Carteira":k;
    tr.push({x:ab.map(x=>mesBR(x.m)),y:ab.map(x=>get(x,1)),name:`${nm} — spread (esq.)`,mode:"lines+markers",line:{color:SPC[i%6],width:2.2},marker:{size:4},hovertemplate:`${nm} spread: CDI+ %{y:.2f}%<extra>%{x}</extra>`},
     {x:ab.map(x=>mesBR(x.m)),y:ab.map(x=>get(x,2)),name:`${nm} — duration (dir.)`,mode:"lines",line:{color:DUC[i%6],width:2,dash:"dash"},yaxis:"y2",hovertemplate:`${nm} duration: %{y:.2f} anos<extra>%{x}</extra>`})});
  plot("f_sd",tr,LAYR({yaxis:{title:{text:"Spread CDI+ (% a.a.) — linhas cheias"}},yaxis2:{title:{text:"Duration (anos) — tracejadas"},overlaying:"y",side:"right"},xaxis:{type:"category"},margin:{l:56,r:250,t:10,b:40},legend:{x:1.08,y:1,xanchor:"left",font:{size:10}}}))}
function tabFundo(){const f=F[$("#f_sel").value],m=$("#f_mes").value,b=$("#f_busca").value.toLowerCase();
  const rs=D.ativos.filter(a=>a[0]==f.id&&a[1]==m), cats=[...new Set(rs.map(a=>a[4]))].sort();
  if(!$("#f_cat").options.length)$("#f_cat").innerHTML='<option value="">todas as categorias</option>'+cats.map(c=>`<option>${c}</option>`).join("");
  const c=$("#f_cat").value, temos=new Set(D.ativos.filter(a=>F[a[0]].nosso&&a[1]==F[a[0]].abertos[F[a[0]].abertos.length-1]).map(a=>a[2]));
  const L=rs.filter(a=>(!c||a[4]==c)&&(a[3]+" "+a[5]+" "+a[6]).toLowerCase().includes(b)).sort((x,y)=>y[7]-x[7]), nc=f.nosso?8:9;
  $("#f_tab").innerHTML="<tr><th class='l'>Ativo</th><th class='l'>Categoria</th><th class='l'>Grupo</th><th>% PL</th><th>R$ mm</th><th>Spread CDI+</th><th>Duration</th><th>Retorno 12m</th>"+(f.nosso?"":"<th>Nós temos?</th>")+"</tr>"+
   L.map(a=>{const p=PERF[a[2]];return`<tr class="cl" data-ch="${esc(a[2])}"><td class="l">${esc(a[3])}</td><td class="l"><span class="ponto" style="background:${corCat(a[4])}"></span>${esc(a[4])}</td><td class="l mut">${esc(a[5])}</td><td>${br(a[7])}</td><td>${br(a[8],1)}</td><td>${br(a[9])}</td><td>${br(a[10])}</td><td class="${cls(p&&p.w12[0])}">${p&&p.w12[0]!=null?br(p.w12[0])+"%":"—"}</td>${f.nosso?"":`<td>${temos.has(a[2])?"sim":'<b class="neg">não</b>'}</td>`}</tr>`}).join("");
  $("#f_tab").querySelectorAll("tr.cl").forEach(t=>t.onclick=()=>detalheAtivo(t,t.dataset.ch,nc,f.id))}
function topFundo(){const f=F[$("#f_sel").value],todos=D.f_top[f.id]||[],cats=[...new Set(todos.map(a=>a[2]))].sort();
  if($("#f_topcat").dataset.f!=f.id){$("#f_topcat").innerHTML='<option value="">todas as categorias</option>'+cats.map(c=>`<option>${esc(c)}</option>`).join("");$("#f_topcat").dataset.f=f.id}
  const c=$("#f_topcat").value;
  const L=ordena(todos.filter(a=>!c||a[2]==c).map(a=>({ch:a[0],ativo:a[1],cat:a[2],ret:a[3],bps:a[4],perc:a[5],n:a[6]})),tOrd).slice(0,40);
  const C=[["ativo","Ativo","l"],["cat","Categoria","l"],["ret","Retorno 12m"],["bps","Contribuição (bps)"],["perc","% PL médio"],["n","Meses na carteira"]];
  $("#f_top").innerHTML=th(C,tOrd)+L.map(a=>`<tr class="cl" data-ch="${esc(a.ch)}"><td class="l">${esc(a.ativo)}</td><td class="l"><span class="ponto" style="background:${corCat(a.cat)}"></span>${esc(a.cat)}</td><td class="${cls(a.ret)}">${a.ret==null?"—":br(a.ret)+"%"}</td><td class="${cls(a.bps)}">${br(a.bps,1)}</td><td>${br(a.perc)}</td><td>${a.n}</td></tr>`).join("");
  cabec($("#f_top"),C,tOrd,()=>{document.querySelectorAll("#a_fundo .chip[data-m]").forEach(x=>x.classList.toggle("on",x.dataset.m==tOrd.k));topFundo()});
  $("#f_top").querySelectorAll("tr.cl").forEach(t=>t.onclick=()=>detalheAtivo(t,t.dataset.ch,C.length,f.id))}
function mudancas(){const f=F[$("#f_sel").value],ms=[...new Set(D.ativos.filter(a=>a[0]==f.id).map(a=>a[1]))].sort();
  if(ms.length<2){$("#f_mud_sub").textContent="Precisa de duas carteiras abertas.";$("#f_mud_up").innerHTML=$("#f_mud_dn").innerHTML="";return}
  const [m0,m1]=ms.slice(-2),pega=m=>{const o={};D.ativos.filter(a=>a[0]==f.id&&a[1]==m&&!["Caixa","Títulos Públicos","Ajuste"].includes(a[4])).forEach(a=>{o[a[2]]=o[a[2]]||{r:a[3],c:a[4],p:0};o[a[2]].p+=a[7]});return o};
  const A=pega(m0),B=pega(m1),ks=new Set([...Object.keys(A),...Object.keys(B)]),L=[...ks].map(k=>({k,r:(B[k]||A[k]).r,c:(B[k]||A[k]).c,a:A[k]?A[k].p:0,b:B[k]?B[k].p:0})).map(x=>({...x,d:x.b-x.a}));
  $("#f_mud_sub").textContent=`${mesBR(m0)} → ${mesBR(m1)}`;
  const tb=(arr)=>"<tr><th class='l'>Ativo</th><th class='l'>Categoria</th><th>Antes</th><th>Depois</th><th>Δ pp</th></tr>"+arr.map(x=>`<tr class="cl" data-ch="${esc(x.k)}"><td class="l">${esc(x.r)}${!x.a?' <span class="tag" style="background:#1f7a4d">NOVO</span>':!x.b?' <span class="tag" style="background:#b42318">SAIU</span>':""}</td><td class="l mut">${esc(x.c)}</td><td>${br(x.a)}</td><td>${br(x.b)}</td><td class="${cls(x.d)}"><b>${x.d>0?"+":""}${br(x.d)}</b></td></tr>`).join("");
  $("#f_mud_up").innerHTML=tb(L.filter(x=>x.d>0.005).sort((a,b)=>b.d-a.d).slice(0,15));$("#f_mud_dn").innerHTML=tb(L.filter(x=>x.d<-0.005).sort((a,b)=>a.d-b.d).slice(0,15));
  ["#f_mud_up","#f_mud_dn"].forEach(s=>$(s).querySelectorAll("tr.cl").forEach(t=>t.onclick=()=>detalheAtivo(t,t.dataset.ch,5,f.id)))}
// ---------- performance de ativos
const RV=["Ações","FII","FIP","ETF","Fundos offshore","Cotas de Fundos"],vc=[...new Set(D.perf.map(p=>p.categoria))].sort();
$("#v_cat").innerHTML='<option value="_cred">crédito (sem renda variável/offshore)</option><option value="">todas as categorias</option>'+vc.map(c=>`<option>${esc(c)}</option>`).join("");
["#v_jan","#v_cat","#v_min","#v_nt","#v_busca"].forEach(s=>$(s).oninput=perf);const vOrd={k:"ret",asc:false};
function perf(){const j=$("#v_jan").value,c=$("#v_cat").value,mn=+$("#v_min").value,nt=$("#v_nt").checked,b=$("#v_busca").value.toLowerCase();
  const L=ordena(D.perf.filter(p=>(!c||(c=="_cred"?!RV.includes(p.categoria):p.categoria==c))&&p.n_ever>=mn&&(!nt||!p.temos)&&p[j][4]>=3&&(p.rotulo+" "+p.emissor+" "+p.grupo).toLowerCase().includes(b))
    .map(p=>({p,rotulo:p.rotulo,categoria:p.categoria,grupo:p.grupo,ret:p[j][0],pcdi:p[j][1],vol:p[j][2],sh:p[j][3],n:p[j][4],u1:p.ult3[2],n_hoje:p.n_hoje,n_ever:p.n_ever,perc:p.perc_med,sp:p.spread,du:p.duration,temos:p.temos?1:0})),vOrd);
  const C=[["rotulo","Ativo","l"],["categoria","Categoria","l"],["grupo","Grupo","l"],["ret","Retorno acum."],["pcdi","% do CDI"],["vol","Vol a.a."],["sh","Sharpe"],["n","Meses"],["u1",`Último mês`],["n_hoje","Fundos hoje"],["n_ever","Fundos (já)"],["perc","% PL médio"],["sp","Spread CDI+"],["du","Duration"],["temos","Nós temos"]];
  $("#v_tab").innerHTML=th(C,vOrd)+L.slice(0,300).map(r=>`<tr class="cl" data-ch="${esc(r.p.chave)}">`+C.map(c=>{const v=r[c[0]];
    if(c[2])return`<td class="l${c[0]=="grupo"?" mut":""}">${c[0]=="categoria"?`<span class="ponto" style="background:${corCat(v)}"></span>`:""}${esc(v)}</td>`;
    if(c[0]=="temos")return`<td>${v?"sim":'<b class="neg">não</b>'}</td>`;if(["ret","u1"].includes(c[0]))return`<td class="${cls(v)}">${v==null?"—":br(v)+"%"}</td>`;
    if(c[0]=="pcdi")return`<td class="${v==null?"":v>=100?"pos":"neg"}">${v==null?"—":br(v,0)+"%"}</td>`;if(["n","n_hoje","n_ever"].includes(c[0]))return`<td>${v??"—"}</td>`;return`<td>${br(v)}</td>`}).join("")+"</tr>").join("")+
    (L.length>300?`<tr><td class="l mut" colspan="${C.length}">Mostrando 300 de ${L.length}: use os filtros ou ordene por outra coluna.</td></tr>`:"");
  cabec($("#v_tab"),C,vOrd,perf);$("#v_tab").querySelectorAll("tr.cl").forEach(t=>t.onclick=()=>detalheAtivo(t,t.dataset.ch,C.length))}
// ---------- exposições vs peers
const pc=[...new Set(D.overlap.map(o=>o.categoria))].sort();$("#p_cat").innerHTML='<option value="">todas as categorias</option>'+pc.map(c=>`<option>${esc(c)}</option>`).join("");
["#p_cat","#p_min","#p_so","#p_gr","#p_busca"].forEach(s=>$(s).oninput=peers);const pOrd={k:"n",asc:false};
function peers(){const c=$("#p_cat").value,mn=+$("#p_min").value,so=$("#p_so").checked,gr=$("#p_gr").checked,b=$("#p_busca").value.toLowerCase();
  const L=ordena(D.overlap.filter(o=>(!c||o.categoria==c)&&o.n>=mn&&(so||!o.temos)&&(!gr||!o.grupo_nosso)&&(o.rotulo+" "+o.emissor+" "+o.grupo).toLowerCase().includes(b)).map(o=>({...o,r12:PERF[o.chave]?PERF[o.chave].w12[0]:null})),pOrd);
  const C=[["rotulo","Ativo","l"],["categoria","Categoria","l"],["grupo","Grupo","l"],["n","Nº peers"],["perc_med","% PL médio"],["rs","R$ mm (peers)"],["spread","Spread CDI+"],["duration","Duration"],["r12","Retorno 12m"],["temos","Nós temos o papel"],["grupo_nosso","Nosso % no grupo"],["peers","Quem tem","l"]];
  $("#p_tab").innerHTML=th(C,pOrd)+L.slice(0,400).map(o=>`<tr class="cl" data-ch="${esc(o.chave)}">`+C.map(x=>{const v=o[x[0]];
    if(x[0]=="peers")return`<td class="l mut" title="${esc(v.map(i=>F[i].nome).join(", "))}">${esc(v.map(i=>F[i].nome).slice(0,3).join(", "))}${v.length>3?" +"+(v.length-3):""}</td>`;
    if(x[0]=="temos")return`<td>${v?"sim":'<b class="neg">não</b>'}</td>`;if(x[0]=="grupo_nosso")return`<td>${v?br(v):'<b class="neg">0</b>'}</td>`;if(x[0]=="r12")return`<td class="${cls(v)}">${v==null?"—":br(v)+"%"}</td>`;
    if(x[2])return`<td class="l">${esc(v)}</td>`;if(x[0]=="n")return`<td><b>${v}</b></td>`;return`<td>${br(v,x[0]=="rs"?1:2)}</td>`}).join("")+"</tr>").join("");
  cabec($("#p_tab"),C,pOrd,peers);$("#p_tab").querySelectorAll("tr.cl").forEach(t=>t.onclick=()=>detalheAtivo(t,t.dataset.ch,C.length));
  const G=D.grupos.filter(g=>(!c||g.categorias.includes(c))&&g.grupo.toLowerCase().includes(b)).map(g=>({...g,dif:(g.peers_perc||0)-(g.nossos_perc||0)})).sort((a,x)=>x.dif-a.dif).slice(0,60);
  $("#p_gru").innerHTML="<tr><th class='l'>Grupo</th><th class='l'>Categorias</th><th>Peers: % PL médio</th><th>Nossos: % PL médio</th><th>Diferença</th><th>Nº peers</th><th>Nº nossos</th><th>R$ mm (peers)</th></tr>"+
   G.map(g=>`<tr><td class="l">${esc(g.grupo)}</td><td class="l mut">${esc(g.categorias.join(", "))}</td><td>${br(g.peers_perc)}</td><td>${br(g.nossos_perc)}</td><td class="${g.dif>0?"neg":"pos"}"><b>${g.dif>0?"+":""}${br(g.dif)}</b></td><td>${g.n_peers}</td><td>${g.n_nossos}</td><td>${br(g.rs_peers,1)}</td></tr>`).join("")}
geral();fundo();perf();peers();vai(ABAS.some(a=>"#"+a[0]==location.hash)?location.hash.slice(1):"a_geral");
</script></body></html>
'''   # painel em HTML (arquivo unico)


def _inicio():                              # servidor (pasta/Docker): pre-carga em segundo plano
    if PRE["status"] == "parado":
        threading.Thread(target=precarrega, daemon=True).start()




# ------------------------------------------------------------------ API
@app.middleware("http")
async def sem_cache(req, call_next):         # pagina/JS/CSS novos aparecem logo depois de atualizar o arquivo
    r = await call_next(req)
    if not req.url.path.startswith("/api/"):
        r.headers["Cache-Control"] = "no-cache"
    return r


@app.get("/api/inicial")
def inicial():
    return {"status": PRE["status"], "log": PRE["log"][-8:], "erro": PRE["erro"]}


def _arquivo(nome):
    if PRE["status"] != "ok":
        raise HTTPException(503, "painel ainda carregando")
    f = os.path.join(PRE["pasta"], nome)
    if not os.path.exists(f):
        raise HTTPException(404, "sem dados")
    return f


@app.get("/api/visao")
def visao():
    return FileResponse(_arquivo("visao.json"), media_type="application/json")


@app.get("/api/fundo/{fid}")
def fundo(fid: str):
    return FileResponse(_arquivo(re.sub(r"\D", "", fid) + ".json"), media_type="application/json")


_ATIVOS = {}


@app.get("/api/ativos/{fid}")
def ativos_mes(fid: str, mes: str):
    """Todos os ativos do fundo no mes (rotulo, categoria, grupo, % PL, spread, duration, PU, fonte)."""
    f = _arquivo(re.sub(r"\D", "", fid) + "_ativos.parquet")
    if f not in _ATIVOS:
        if len(_ATIVOS) > 30:
            _ATIVOS.clear()
        _ATIVOS[f] = pd.read_parquet(f)
    t = _ATIVOS[f]
    return JSONResponse(limpo(t[t.mes == mes].drop(columns="mes").to_dict("records")))


@app.get("/api/csv/{fid}")
def csv(fid: str):
    df = pd.read_parquet(_arquivo(re.sub(r"\D", "", fid) + ".parquet"))
    return Response(df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"), media_type="text/csv",
                    headers={"Content-Disposition": f'attachment; filename="carteira_{fid}.csv"'})


@app.get("/")
def index():
    return HTMLResponse(PAGINA)


PAGINA = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Painel de fundos</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect x='4' y='16' width='5' height='12' rx='1' fill='%232a78d6'/><rect x='13' y='9' width='5' height='19' rx='1' fill='%231baf7a'/><rect x='22' y='4' width='5' height='24' rx='1' fill='%23eb6834'/></svg>">
<style>
:root {
  --bg: #f5f6f8; --surface: #ffffff; --surface-2: #f0f1f4; --line: #e3e5ea; --line-2: #d0d3da;
  --text-1: #14161a; --text-2: #4d525c; --muted: #8a8f99;
  --accent: #2a78d6; --accent-soft: rgba(42,120,214,.10); --nosso: #eb6834; --nosso-soft: rgba(235,104,52,.12);
  --good: #0f7b3f; --bad: #c43d3d; --radius: 12px;
  --shadow: 0 1px 2px rgba(16,24,40,.05), 0 1px 3px rgba(16,24,40,.06);
  --font: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", Arial, sans-serif;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --bg: #0f1012; --surface: #17191c; --surface-2: #1f2226; --line: #2a2d32; --line-2: #3a3e45;
    --text-1: #f2f3f5; --text-2: #b7bcc6; --muted: #858b96; --accent: #3987e5; --accent-soft: rgba(57,135,229,.16);
    --good: #2fbf6b; --bad: #ef6b6b; --shadow: none; --nosso-soft: rgba(235,104,52,.18);
  }
}
:root[data-theme="dark"] {
  --bg: #0f1012; --surface: #17191c; --surface-2: #1f2226; --line: #2a2d32; --line-2: #3a3e45;
  --text-1: #f2f3f5; --text-2: #b7bcc6; --muted: #858b96; --accent: #3987e5; --accent-soft: rgba(57,135,229,.16);
  --good: #2fbf6b; --bad: #ef6b6b; --shadow: none; --nosso-soft: rgba(235,104,52,.18);
}
* { box-sizing: border-box; }
html, body { margin: 0; background: var(--bg); color: var(--text-1); font: 14px/1.5 var(--font); }
h1 { font-size: 16px; margin: 0; }
h2 { font-size: 20px; margin: 0; letter-spacing: -.01em; }
h3 { font-size: 14px; margin: 0 0 2px; }
button { font: inherit; cursor: pointer; }
input, select { font: inherit; color: var(--text-1); background: var(--surface); border: 1px solid var(--line-2); border-radius: 8px; padding: 8px 10px; }

.barra { position: sticky; top: 0; z-index: 20; display: flex; align-items: center; gap: 18px; padding: 8px 22px; background: var(--surface); border-bottom: 1px solid var(--line); }
.marca { display: flex; align-items: center; gap: 10px; min-width: 0; }
.logo { width: 26px; height: 26px; flex: none; } .logo.grande { width: 48px; height: 48px; }
.datas { margin: 0; font-size: 11.5px; color: var(--text-2); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.abas { display: flex; gap: 4px; margin-left: auto; }
.abas button { border: 0; background: transparent; color: var(--text-2); padding: 8px 14px; border-radius: 8px; font-weight: 600; }
.abas button:hover { background: var(--surface-2); }
.abas button[aria-selected="true"] { background: var(--accent-soft); color: var(--accent); }
.botao { border: 1px solid var(--line-2); background: var(--surface); color: var(--text-1); border-radius: 8px; padding: 7px 12px; }
.botao:hover { background: var(--surface-2); }
.botao.icone { width: 36px; padding: 7px 0; }
.botao.sel { border-color: var(--accent); color: var(--accent); background: var(--accent-soft); font-weight: 600; }

main { max-width: 1380px; margin: 0 auto; padding: 18px 24px 60px; }
.cab { display: flex; flex-wrap: wrap; gap: 14px; align-items: flex-end; justify-content: space-between; margin-bottom: 14px; }
.titulo-fundo { min-width: 0; }
.titulo-fundo h2 { display: flex; align-items: center; gap: 8px; }
.tag { font-size: 11px; font-weight: 700; color: var(--nosso); background: var(--nosso-soft); border-radius: 6px; padding: 2px 7px; letter-spacing: .02em; }
.sub { color: var(--muted); font-size: 12px; margin: 2px 0 0; }
.controles { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
.controles label { font-size: 12px; color: var(--text-2); display: flex; gap: 6px; align-items: center; }

/* busca de fundo por digitacao */
.busca { position: relative; width: min(420px, 100%); }
.busca input { width: 100%; padding-left: 34px; }
.busca::before { content: "⌕"; position: absolute; left: 11px; top: 6px; font-size: 17px; color: var(--muted); }
.sugestoes { position: absolute; z-index: 30; left: 0; right: 0; top: 100%; margin: 4px 0 0; padding: 4px; list-style: none; max-height: 340px; overflow: auto;
  background: var(--surface); border: 1px solid var(--line-2); border-radius: 10px; box-shadow: 0 10px 28px rgba(0,0,0,.16); }
.sugestoes li { padding: 7px 9px; border-radius: 7px; cursor: pointer; display: flex; justify-content: space-between; gap: 10px; font-size: 13px; }
.sugestoes li small { color: var(--muted); }
.sugestoes li:hover, .sugestoes li.ativo { background: var(--surface-2); }

.grade { display: grid; gap: 16px; margin-bottom: 16px; }
.g2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.g-tabela { grid-template-columns: minmax(0, 3fr) minmax(0, 2fr); }
.cartao { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); box-shadow: var(--shadow); padding: 14px 16px; min-width: 0; margin-bottom: 16px; }
.grade > .cartao { margin-bottom: 0; }
.cartao .sub { margin-bottom: 8px; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 16px; }
.kpi { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); box-shadow: var(--shadow); padding: 12px 14px; }
.kpi .r { font-size: 12px; color: var(--text-2); }
.kpi .v { font-size: 24px; font-weight: 700; letter-spacing: -.02em; margin-top: 2px; font-variant-numeric: tabular-nums; }
.kpi .s { font-size: 11.5px; color: var(--muted); }
.pos { color: var(--good); } .neg { color: var(--bad); }

.tabela-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 10px; }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th, td { padding: 7px 11px; border-bottom: 1px solid var(--line); text-align: left; white-space: nowrap; }
th { position: sticky; top: 0; background: var(--surface-2); color: var(--text-2); font-weight: 600; font-size: 12px; z-index: 1; }
th.ord { cursor: pointer; } th.ord:hover { color: var(--text-1); }
tbody tr:hover td { background: var(--accent-soft); }
td.n, th.n { text-align: right; font-variant-numeric: tabular-nums; }
td.texto { white-space: normal; min-width: 200px; }
tr.total td { font-weight: 700; border-top: 2px solid var(--line-2); background: var(--surface-2); }
tr.secao td { font-size: 11.5px; text-transform: uppercase; letter-spacing: .05em; color: var(--muted); background: var(--surface); padding-top: 14px; }
tr.nosso td:first-child { box-shadow: inset 3px 0 0 var(--nosso); }
.ponto { display: inline-block; width: 9px; height: 9px; border-radius: 3px; margin-right: 6px; vertical-align: 0; }
.nota { color: var(--muted); font-size: 11.5px; margin: 8px 0 0; }

/* filtros de ativos: categorias liga/desliga + busca + quantos */
.filtros { display: flex; flex-wrap: wrap; gap: 10px 16px; align-items: center; justify-content: space-between; }
.chips { display: flex; flex-wrap: wrap; gap: 6px; }
.chip-t { border: 1px solid var(--line-2); background: var(--surface); color: var(--text-2); border-radius: 999px; padding: 4px 11px; font-size: 12.5px; }
.chip-t[aria-pressed="true"] { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); font-weight: 600; }
.chip-t.mudo { border-style: dashed; }
.filtros .busca { width: min(320px, 100%); }
.muted { color: var(--muted); font-size: 12px; }

/* selecao multipla (consolidado) */
.multi { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; padding: 6px; border: 1px solid var(--line-2); border-radius: 10px; background: var(--surface); min-height: 44px; }
.multi .chip { display: inline-flex; align-items: center; gap: 5px; border-radius: 999px; padding: 3px 4px 3px 10px; font-size: 12.5px; background: var(--surface-2); }
.multi .chip.nosso { background: var(--nosso-soft); }
.multi .chip button { border: 0; background: none; color: var(--muted); font-size: 15px; line-height: 1; padding: 0 4px; }
.multi .busca { width: 220px; flex: 1; }
.multi .busca input { border: 0; padding: 5px 5px 5px 30px; }
.multi .busca::before { top: 2px; }

.carregando { position: fixed; inset: 0; z-index: 50; display: grid; place-items: center; background: var(--bg); }
.carregando[hidden] { display: none; }
.carregando .caixa { width: min(560px, 92vw); text-align: center; }
.carregando h2 { margin-top: 12px; }
.descricao { color: var(--text-2); font-size: 13px; }
.barra-prog { height: 4px; background: var(--line); border-radius: 4px; overflow: hidden; margin: 18px 0 10px; }
.barra-prog div { height: 100%; width: 35%; background: var(--accent); border-radius: 4px; animation: vai 1.4s ease-in-out infinite; }
@keyframes vai { 0% { transform: translateX(-100%); } 100% { transform: translateX(300%); } }
pre { text-align: left; white-space: pre-wrap; font: 12px/1.5 ui-monospace, Consolas, monospace; color: var(--text-2); background: var(--surface);
  border: 1px solid var(--line); border-radius: 8px; padding: 10px; max-height: 220px; overflow: auto; margin: 0; }
.vazio { padding: 60px 0; text-align: center; color: var(--text-2); }

@media (max-width: 980px) {
  html, body { overflow-x: hidden; }
  .barra { flex-wrap: wrap; padding: 8px 14px; gap: 8px; }
  .abas { margin-left: 0; order: 3; width: 100%; }
  main { padding: 14px 14px 50px; }
  .g2, .g-tabela { grid-template-columns: minmax(0, 1fr); }
  .kpi .v { font-size: 20px; }
}

/* metodologia */
.metodo { display: grid; grid-template-columns: repeat(auto-fill, minmax(min(100%, 340px), 1fr)); gap: 10px 22px; margin-top: 10px; }
.metodo b { font-size: 13px; }
.metodo p { margin: 2px 0 0; font-size: 12.5px; color: var(--text-2); line-height: 1.45; }
.marca-sel { display: inline-block; width: 16px; color: var(--accent); font-weight: 700; }
.chip-t .ponto { margin-right: 5px; }
#busca_cons .busca { width: min(340px, 100%); }

</style>
<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js"></script>
</head>
<body>
<header class="barra">
  <div class="marca">
    <svg class="logo" viewBox="0 0 32 32" aria-hidden="true"><rect x="4" y="16" width="5" height="12" rx="1.5" fill="#2a78d6"/><rect x="13.5" y="9" width="5" height="19" rx="1.5" fill="#1baf7a"/><rect x="23" y="4" width="5" height="24" rx="1.5" fill="#eb6834"/></svg>
    <div><h1>Painel de fundos</h1><p id="datas" class="datas"></p></div>
  </div>
  <nav class="abas" role="tablist">
    <button role="tab" data-aba="carteira" aria-selected="true">Carteira</button>
    <button role="tab" data-aba="retorno" aria-selected="false">Retorno</button>
    <button role="tab" data-aba="consolidado" aria-selected="false">Consolidado</button>
  </nav>
  <button id="tema" class="botao icone" title="Tema claro/escuro" aria-label="Alternar tema">◐</button>
</header>
<main id="conteudo"></main>

<div id="carregando" class="carregando">
  <div class="caixa">
    <svg class="logo grande" viewBox="0 0 32 32" aria-hidden="true"><rect x="4" y="16" width="5" height="12" rx="1.5" fill="#2a78d6"/><rect x="13.5" y="9" width="5" height="19" rx="1.5" fill="#1baf7a"/><rect x="23" y="4" width="5" height="24" rx="1.5" fill="#eb6834"/></svg>
    <h2>Preparando o painel</h2>
    <p class="descricao">Baixando e processando os dados de todos os fundos. Na primeira vez do dia pode levar alguns minutos.</p>
    <div class="barra-prog"><div></div></div>
    <pre id="log-carga"></pre>
  </div>
</div>
<script>
// Painel de fundos - 3 abas: Carteira, Retorno, Consolidado. Dados processados pelo servidor (/api).
const $ = (s, el = document) => el.querySelector(s);
const nf = (d) => new Intl.NumberFormat("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const num = (v, d = 2) => (v == null || !isFinite(v) ? "–" : nf(d).format(v));
const pct = (v, d = 2) => (v == null || !isFinite(v) ? "–" : nf(d).format(v) + "%");
const pp = (v, d = 2) => (v == null || !isFinite(v) ? "–" : (v > 0 ? "+" : "") + nf(d).format(v) + " pp");
const cls = (v) => (v == null ? "" : v >= 0 ? "pos" : "neg");
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
const mesBR = (m) => (m ? MESES[+m.slice(5, 7) - 1] + "/" + m.slice(0, 4) : "–");
const dataBR = (d) => (d ? d.slice(8, 10) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : "–");
const semAcento = (t) => String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
const espera = (ms) => new Promise((ok) => setTimeout(ok, ms));
const st = { V: null, F: {}, aba: "carteira", fundo: null, mesCart: {}, ret: { de: null, ate: null }, cons: { sel: null, de: null, ate: null, ord: "ret", asc: false },
  A: {}, gf: { fora: new Set(), de: null, ate: null, fundo: null }, cf: { fora: new Set(["Caixa", "Confidencial"]), n: 15, q: "" }, rf: { fora: new Set(["Caixa", "Confidencial"]), n: 15, q: "" } };

// ------------------------------------------------------------------ tema e abas
try { const t = localStorage.getItem("tema"); if (t) document.documentElement.dataset.theme = t; } catch (e) {}
$("#tema").onclick = () => {
  const escuro = document.documentElement.dataset.theme === "dark" || (!document.documentElement.dataset.theme && matchMedia("(prefers-color-scheme: dark)").matches);
  document.documentElement.dataset.theme = escuro ? "light" : "dark";
  try { localStorage.setItem("tema", document.documentElement.dataset.theme); } catch (e) {}
  if (st.V) desenha();
};
document.querySelectorAll(".abas button").forEach((b) => (b.onclick = () => {
  st.aba = b.dataset.aba;
  document.querySelectorAll(".abas button").forEach((x) => x.setAttribute("aria-selected", x === b));
  desenha();
}));

// ------------------------------------------------------------------ carga
(async function inicia() {
  for (;;) {
    let s;
    try { s = await (await fetch("/api/inicial")).json(); } catch (e) { await espera(2000); continue; }
    $("#log-carga").textContent = (s.log || []).join("\n") || "Iniciando...";
    if (s.status === "erro") return ($("#log-carga").textContent += "\n\nErro: " + s.erro);
    if (s.status === "ok") break;
    await espera(2000);
  }
  st.V = await (await fetch("/api/visao")).json();
  const V = st.V;
  st.fundo = (V.fundos.find((f) => f.nosso && f.tem_carteira) || V.fundos.find((f) => f.tem_carteira) || V.fundos[0]).id;
  st.cons.sel = new Set(V.fundos.map((f) => f.id));
  $("#datas").textContent = `${V.fundos.length} fundos (${V.fundos.filter((f) => f.nosso).length} nossos) · carteiras até ${mesBR(V.ult_carteira)} · cotas até ${dataBR(V.datas[V.datas.length - 1])} · ANBIMA ${dataBR(V.anbima)}`;
  $("#carregando").hidden = true;
  desenha();
})();
async function detalhe(id) {
  if (!st.F[id]) st.F[id] = await (await fetch("/api/fundo/" + id)).json();
  return st.F[id];
}
const fundoPor = (id) => st.V.fundos.find((f) => f.id === id);

// ------------------------------------------------------------------ busca de fundo (digitando, sugestoes estilo Google)
function pontua(q, f) {
  // quanto o fundo combina com o que foi digitado: comeco do nome > comeco de palavra > trecho > palavras > parecido (erros de digitacao)
  const n = semAcento(f.nome), qq = semAcento(q).trim(), dig = q.replace(/\D/g, "");
  if (!qq) return 1;
  if (dig.length > 2 && f.id.includes(dig)) return 100;
  if (n.startsWith(qq)) return 95;
  const pal = n.split(/[\s_\-.]+/);
  if (pal.some((p) => p.startsWith(qq))) return 85;
  if (n.includes(qq)) return 75;
  const toks = qq.split(/\s+/).filter(Boolean);
  if (toks.every((t) => pal.some((p) => p.startsWith(t)))) return 65;
  const big = (s) => { const b = new Set(); for (let i = 0; i < s.length - 1; i++) b.add(s.slice(i, i + 2)); return b; };
  const a = big(qq.replace(/\s+/g, "")), b = big(n.replace(/\s+/g, ""));
  let c = 0; a.forEach((x) => b.has(x) && c++);
  const dice = (2 * c) / Math.max(a.size + b.size, 1);
  return dice >= 0.3 ? 50 * dice : 0;
}
function campoBusca(onPick, placeholder = "Digite o nome ou CNPJ do fundo...", soComCarteira = false, marcado = null) {
  // marcado(id) -> true/false: modo selecao multipla (mostra ✓ e nao fecha ao escolher)
  const wrap = document.createElement("div");
  wrap.className = "busca";
  wrap.innerHTML = `<input type="search" placeholder="${placeholder}" autocomplete="off"><ul class="sugestoes" hidden></ul>`;
  const inp = $("input", wrap), ul = $("ul", wrap);
  let sel = 0, lista = [];
  const filtra = () => {
    lista = st.V.fundos.filter((f) => !soComCarteira || f.tem_carteira).map((f) => ({ f, p: pontua(inp.value, f) })).filter((x) => x.p > 0)
      .sort((a, b) => b.p - a.p || b.f.nosso - a.f.nosso || a.f.nome.localeCompare(b.f.nome)).slice(0, 12).map((x) => x.f);
    sel = 0;
    ul.innerHTML = lista.map((f, i) => `<li data-i="${i}" class="${i === 0 ? "ativo" : ""}"><span>${marcado ? `<span class="marca-sel">${marcado(f.id) ? "✓" : ""}</span>` : ""}${esc(f.nome)}${f.nosso ? ' <span class="tag">NOSSO</span>' : ""}</span><small>${esc(f.cnpj)}</small></li>`).join("") || "<li>Nada encontrado</li>";
    ul.hidden = false;
  };
  const escolhe = (i) => {
    if (!lista[i]) return;
    if (!marcado) { ul.hidden = true; inp.value = ""; }
    onPick(lista[i]);
    if (marcado) filtra();
  };
  inp.addEventListener("input", filtra);
  inp.addEventListener("focus", filtra);
  inp.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      sel = (sel + (e.key === "ArrowDown" ? 1 : -1) + lista.length) % Math.max(lista.length, 1);
      ul.querySelectorAll("li").forEach((x, i) => x.classList.toggle("ativo", i === sel)); e.preventDefault();
    } else if (e.key === "Enter") { e.preventDefault(); escolhe(sel); } else if (e.key === "Escape") ul.hidden = true;
  });
  ul.addEventListener("mousedown", (e) => { const li = e.target.closest("li[data-i]"); if (li) { e.preventDefault(); escolhe(+li.dataset.i); } });
  inp.addEventListener("blur", () => setTimeout(() => (ul.hidden = true), 150));
  return wrap;
}

// ------------------------------------------------------------------ graficos e tabelas
function layout(extra = {}) {
  return Object.assign({
    paper_bgcolor: "rgba(0,0,0,0)", plot_bgcolor: "rgba(0,0,0,0)", font: { family: css("--font"), size: 12, color: css("--text-2") },
    margin: { l: 50, r: 14, t: 10, b: 40 }, hovermode: "x unified", legend: { orientation: "h", y: -0.18, font: { color: css("--text-2") } },
    xaxis: { gridcolor: "rgba(0,0,0,0)", linecolor: css("--line-2"), automargin: true },
    yaxis: { gridcolor: css("--line"), zerolinecolor: css("--line-2"), automargin: true },
    hoverlabel: { bgcolor: css("--surface"), bordercolor: css("--line-2"), font: { color: css("--text-1") } },
  }, extra);
}
function plota(el, dados, lay, h = 340) {
  el.style.height = h + "px";
  Plotly.newPlot(el, dados, lay, { displaylogo: false, responsive: true, displayModeBar: innerWidth > 700 ? "hover" : false,
    modeBarButtonsToRemove: ["lasso2d", "select2d", "autoScale2d"] });
}
const cartao = (titulo, sub, corpo) => `<div class="cartao"><h3>${titulo}</h3>${sub ? `<p class="sub">${sub}</p>` : ""}${corpo}</div>`;
function tabela(cols, linhas, extra = {}) {
  const th = cols.map((c) => `<th class="${c.n ? "n" : ""} ${c.ord ? "ord" : ""}" ${c.ord ? `data-ord="${c.ord}"` : ""}>${c.t}</th>`).join("");
  const tr = linhas.map((r) => `<tr class="${r._cls || ""}">${cols.map((c) => `<td class="${c.n ? "n" : c.w ? "texto" : ""}">${c.f ? c.f(r[c.k], r) : esc(r[c.k] ?? "–")}</td>`).join("")}</tr>`).join("");
  return `<div class="tabela-wrap" style="max-height:${extra.alt || 620}px"><table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table></div>`;
}
const kpi = (r, v, s = "", c = "") => `<div class="kpi"><div class="r">${r}</div><div class="v ${c}">${v}</div><div class="s">${s}</div></div>`;
function categoriasPrincipais(series, n = 7) {
  return Object.entries(series).map(([c, ys]) => [c, ys.reduce((s, y) => s + (y || 0), 0)]).sort((a, b) => b[1] - a[1]).slice(0, n).map((x) => x[0]);
}

function desenha() {
  const el = $("#conteudo");
  try {
    if (st.aba === "carteira") carteira(el);
    else if (st.aba === "retorno") retorno(el);
    else consolidado(el);
  } catch (e) { el.innerHTML = `<div class="vazio">Erro ao desenhar: ${esc(e.message)}</div>`; console.error(e); }
}
function cabecalhoFundo(el, f, extra = "") {
  el.innerHTML = `<div class="cab"><div class="titulo-fundo"><h2>${esc(f.nome)} ${f.nosso ? '<span class="tag">NOSSO</span>' : ""}</h2>
    <p class="sub">${esc(f.cnpj)}${f.pl ? ` · PL R$ ${num(f.pl, 0)} mm` : ""}</p></div><div class="controles" id="ctl">${extra}</div></div><div id="corpo"></div>`;
  $("#ctl").prepend(campoBusca((x) => { st.fundo = x.id; desenha(); }, "Trocar de fundo: digite nome ou CNPJ...", st.aba === "carteira"));
}

// ------------------------------------------------------------------ filtros de ativos (categorias liga/desliga, quantos, busca)
function filtros(box, todas, f, redesenha, placeholder) {
  box.innerHTML = `<div class="filtros"><div class="chips">${todas.map((c) => `<button class="chip-t" data-c="${esc(c)}" aria-pressed="${!f.fora.has(c)}">${esc(c)}</button>`).join("")}
      <button class="chip-t mudo" data-t="1">Todas</button><button class="chip-t mudo" data-t="0">Nenhuma</button></div>
    <div class="controles"><div class="busca"><input type="search" class="q" placeholder="${placeholder}" value="${esc(f.q)}"></div>
      <label>Mostrar <select class="n">${[5, 10, 15, 20, 30, 0].map((n) => `<option value="${n}" ${n === f.n ? "selected" : ""}>${n || "Todos"}</option>`).join("")}</select></label></div></div>`;
  const marca = () => box.querySelectorAll("[data-c]").forEach((x) => x.setAttribute("aria-pressed", !f.fora.has(x.dataset.c)));
  box.querySelectorAll("[data-c]").forEach((b) => (b.onclick = () => { const c = b.dataset.c; f.fora.has(c) ? f.fora.delete(c) : f.fora.add(c); marca(); redesenha(); }));
  box.querySelectorAll("[data-t]").forEach((b) => (b.onclick = () => { f.fora = new Set(b.dataset.t === "1" ? [] : todas); marca(); redesenha(); }));
  $(".q", box).oninput = (e) => { f.q = e.target.value; redesenha(); };
  $(".n", box).onchange = (e) => { f.n = +e.target.value; redesenha(); };
}
const passa = (f, x) => !f.fora.has(x.categoria) && (!f.q.trim() || semAcento([x.ativo, x.grupo, x.emissor, x.codigo].join(" ")).includes(semAcento(f.q.trim())));
const corta = (arr, n) => (n ? arr.slice(0, n) : arr);
const fonteCurta = (v) => `<span class="muted">${esc(v || "–")}</span>`;

// ------------------------------------------------------------------ graficos da carteira: categorias liga/desliga + periodo
function graficosCarteira(D, ms, aberto) {
  const G = st.gf;
  const iAb = ms.indexOf(aberto.mes);
  if (!ms.includes(G.de) || G.fundo !== st.fundo) G.de = ms[0];
  if (!ms.includes(G.ate) || G.fundo !== st.fundo) G.ate = ms[iAb >= 0 ? iAb : ms.length - 1];
  G.fundo = st.fundo;
  const peso = {};
  D.carteira.forEach((x) => x.cats.forEach((c) => (peso[c.categoria] = (peso[c.categoria] || 0) + Math.abs(c.perc))));
  const ordem = Object.keys(peso).sort((a, b) => peso[b] - peso[a]);
  const pal = st.V.cat_cores.concat(["#7d5ba6", "#c2185b", "#00838f", "#8d6e63"]);
  const cor = (c) => (ordem.indexOf(c) < pal.length ? pal[ordem.indexOf(c)] : "#9aa0a6");
  const opt = (v) => ms.map((m) => `<option value="${m}" ${m === v ? "selected" : ""}>${mesBR(m)}${D.carteira.find((x) => x.mes === m).conf >= 5 ? " (confid.)" : ""}</option>`).join("");
  $("#ctl_graf").innerHTML = `<label>De <select id="gde">${opt(G.de)}</select></label><label>até <select id="gate">${opt(G.ate)}</select></label>
    <button class="botao" data-gp="aberto">Até a última aberta</button><button class="botao" data-gp="tudo">Tudo</button>`;
  $("#f_graf").innerHTML = `<div class="chips">${ordem.map((c) => `<button class="chip-t" data-gc="${esc(c)}" aria-pressed="${!G.fora.has(c)}"><span class="ponto" style="background:${cor(c)}"></span>${esc(c)}</button>`).join("")}
    <button class="chip-t mudo" data-gt="1">Todas</button><button class="chip-t mudo" data-gt="0">Nenhuma</button></div>`;
  const marca = () => document.querySelectorAll("[data-gc]").forEach((x) => x.setAttribute("aria-pressed", !G.fora.has(x.dataset.gc)));
  document.querySelectorAll("[data-gc]").forEach((b) => (b.onclick = () => { const c = b.dataset.gc; G.fora.has(c) ? G.fora.delete(c) : G.fora.add(c); marca(); desenhaG(); }));
  document.querySelectorAll("[data-gt]").forEach((b) => (b.onclick = () => { G.fora = new Set(b.dataset.gt === "1" ? [] : ordem); marca(); desenhaG(); }));
  $("#gde").onchange = (e) => { G.de = e.target.value; desenhaG(); };
  $("#gate").onchange = (e) => { G.ate = e.target.value; desenhaG(); };
  document.querySelectorAll("[data-gp]").forEach((b) => (b.onclick = () => {
    G.de = ms[0]; G.ate = b.dataset.gp === "tudo" ? ms[ms.length - 1] : ms[iAb >= 0 ? iAb : ms.length - 1];
    $("#gde").value = G.de; $("#gate").value = G.ate; desenhaG();
  }));
  function desenhaG() {
    let a = ms.indexOf(G.de), b = ms.indexOf(G.ate);
    if (a > b) [a, b] = [b, a];
    const M = ms.slice(a, b + 1), X = M.map((m) => m + "-15");
    const eixoX = { ...layout().xaxis, tickformat: "%m/%y", dtick: M.length <= 12 ? "M1" : M.length <= 36 ? "M3" : "M12" };
    const cats = ordem.filter((c) => !G.fora.has(c));
    const val = (c, k) => M.map((m) => { const x = D.carteira.find((y) => y.mes === m).cats.find((y) => y.categoria === c); return x ? x[k] : (k === "perc" ? 0 : null); });
    plota($("#g_aloc"), cats.map((c) => ({ x: X, y: val(c, "perc"), name: c, stackgroup: "a", line: { width: 0.5, color: css("--surface") }, fillcolor: cor(c),
      hovertemplate: `${esc(c)}: %{y:.1f}%<extra></extra>` })), layout({ xaxis: eixoX, yaxis: { ticksuffix: "%", gridcolor: css("--line") } }), 380);
    const cs = cats.filter((c) => !["Caixa", "Confidencial"].includes(c) && !SEM_SPREAD.includes(c));
    plota($("#g_spread"), cs.map((c) => ({ x: X, y: val(c, "spread").map((v, i) => (val(c, "perc")[i] > 0.05 ? v : null)), name: c, mode: "lines", connectgaps: false,
      line: { color: cor(c), width: 2 }, hovertemplate: `${esc(c)}: %{y:.2f}%<extra></extra>` })),
      layout({ xaxis: eixoX, yaxis: { ticksuffix: "%", gridcolor: css("--line") } }), 380);
  }
  desenhaG();
}
const SEM_SPREAD = ["Ações", "FII", "FIP", "ETF"];

// ------------------------------------------------------------------ metodologia (com exemplos deste fundo)
const METODO = [
  ["caixa", "Caixa, compromissadas, LFT e provisões: spread 0 e duration 0 (rendem CDI/Selic)."],
  ["ANBIMA", "Debêntures com taxa indicativa ANBIMA no fim do mês. DI+: spread = taxa. %DI: (x − 100%) × CDI. IPCA+ ou pré: taxa − curva do Tesouro (NTN-B ou prefixado) na mesma duration. Duration = a da ANBIMA."],
  ["SND + preço do fundo", "Debêntures fora da ANBIMA: fluxo de juros/amortização e taxa de emissão do SND (debentures.com.br). Acha-se a taxa que leva o fluxo ao preço que o fundo informou à CVM (preço ÷ PU par) e converte-se em spread como acima. Duration = do fluxo nessa taxa. Testado contra a ANBIMA: erro mediano 0,01 pp (DI+) e 0,09 pp (IPCA+)."],
  ["SND (taxa de emissão)", "Debênture sem preço confiável no mês: taxa de emissão do SND."],
  ["taxa contratada (CVM)", "LF, CDB, CRA, NC e outros com taxa na carteira da CVM: spread pela taxa informada (DI+ direto; IPCA+/pré contra a curva do Tesouro). Duration: paga tudo no vencimento (LF, CDB, NC) = prazo; com cupom semestral nos demais."],
  ["Tesouro (curva)", "NTN-B, LTN, NTN-F: spread 0 (são a própria curva); duration do fluxo com cupom semestral."],
  ["FIDC: rentabilidade da série (CVM)", "Acha-se a série que o fundo tem (mesmo valor de cota) no informe mensal de FIDC da CVM. Spread = mediana de 6 meses da rentabilidade oficial acima do CDI (sênior/mezanino ≈ benchmark contratado; subordinada = retorno realizado, não há taxa fixa). Duration = prazo médio dos recebíveis do FIDC."],
  ["implícito (preço)", "Sem taxa pública (CRI sem taxa, cotas de fundos): mediana de 6 meses da variação mensal do preço acima do CDI, anualizada."],
  ["implícito (preço em US$)", "Bonds e fundos offshore: variação do preço em dólar (sem o câmbio) acima da Treasury americana do mesmo prazo — aproximação do cupom cambial (custo do hedge para CDI)."],
  ["estimado (mesmo emissor)", "Sem dado do próprio ativo: média de outros papéis do mesmo grupo econômico no mês (neste ou em outros fundos)."],
  ["estimado (média da categoria)", "Sem dado do ativo nem do emissor: média ponderada dos ativos da mesma categoria no fundo."],
  ["estimado (média do fundo)", "Categoria inteira sem dado: média do crédito do fundo."],
  ["estimado (carteira do mês anterior)", "Fundo investido ainda sem carteira na CVM no mês: usa o spread/duration do mês anterior."],
  ["duration estimada", "Spread veio de uma fonte acima, mas o ativo não tem vencimento/fluxo: duration = média da categoria."],
  ["não se aplica", "Ações, FII, FIP, ETF: renda variável, sem spread nem duration. Fundo offshore sem preço utilizável: a carteira lá fora não é pública (não se estima)."],
  ["confidencial (CVM)", "Parte que a CVM ainda não divulgou (o fundo tem até 3 meses para abrir a carteira)."],
];
function metodologia(A) {
  const ex = (chave) => {
    const x = A.filter((a) => (a.fonte || "").includes(chave) && a.perc > 0).sort((a, b) => b.perc - a.perc)[0];
    return x ? `<br><span class="muted">Neste fundo: ${esc(x.ativo)} — ${pct(x.perc)} do PL, spread ${num(x.spread)}, duration ${num(x.duration)}</span>` : "";
  };
  return `<h3>Como os números são calculados</h3>
    <p class="sub">Spread CDI+ = quanto o ativo rende acima do CDI, em % a.a. (equivalente em CDI). Duration = prazo médio dos fluxos do ativo ponderado pelo valor presente (anos). Totais = médias ponderadas pelo % do PL. Cada ativo mostra a fonte na tabela acima; abaixo, o que cada fonte significa, em ordem de preferência.</p>
    <div class="metodo">${METODO.map(([k, t]) => `<div><b>${esc(k)}</b><p>${t}${ex(k)}</p></div>`).join("")}</div>
    <p class="nota">Retorno (aba Retorno): cada ativo rende pelo preço + pagamentos do mês (debêntures pelo PU par do SND; FIDC pela rentabilidade oficial; títulos públicos com cupom), ou CDI + spread − duration × variação da curva quando o preço não é confiável; futuros DI1/DAP/dólar vão para os ativos que protegem; taxas de administração e performance (CVM) em linha própria; o que sobra (negociação no mês, swaps) é distribuído pelo peso. A soma fecha com o retorno da cota.</p>`;
}

// ------------------------------------------------------------------ CARTEIRA
async function carteira(el) {
  const f = fundoPor(st.fundo);
  if (!f.tem_carteira) { cabecalhoFundo(el, f); $("#corpo").innerHTML = `<div class="vazio">A CVM não tem carteira deste fundo no período.</div>`; return; }
  cabecalhoFundo(el, f);
  $("#corpo").innerHTML = `<div class="vazio">Carregando ${esc(f.nome)}...</div>`;
  const D = await detalhe(f.id);
  if (st.fundo !== f.id || st.aba !== "carteira") return;
  const ms = D.carteira.map((x) => x.mes);
  const aberto = [...D.carteira].reverse().find((x) => x.conf < 5) || D.carteira[D.carteira.length - 1];
  const mes = ms.includes(st.mesCart[f.id]) ? st.mesCart[f.id] : aberto.mes;
  const C = D.carteira.find((x) => x.mes === mes);
  $("#ctl").insertAdjacentHTML("beforeend", `<label>Carteira <select id="mes">${ms.slice().reverse().map((m) => `<option value="${m}" ${m === mes ? "selected" : ""}>${mesBR(m)}${D.carteira.find((x) => x.mes === m).conf >= 5 ? " (confid.)" : ""}</option>`).join("")}</select></label>
    <a class="botao" href="/api/csv/${f.id}" title="Carteira look-through completa (CSV)">CSV</a>`);
  $("#mes").onchange = (e) => { st.mesCart[f.id] = e.target.value; desenha(); };

  const linhas = C.cats.map((c) => ({ ...c })).concat([{ ...C.total, _cls: "total", categoria: "Total ponderado" }]);
  const tab = tabela([
    { t: "Categoria", k: "categoria", w: 1 },
    { t: "% PL", k: "perc", n: 1, f: (v) => pct(v) },
    { t: "Financeiro (R$ mm)", k: "fin", n: 1, f: (v) => num(v) },
    { t: "Spread CDI+ (% a.a.)", k: "spread", n: 1, f: (v) => num(v) },
    { t: "Duration (anos)", k: "duration", n: 1, f: (v) => num(v) },
  ], linhas);
  const conf = C.conf >= 0.5 ? ` · ${pct(C.conf, 1)} do PL ainda confidencial` : "";
  const fontes = Object.entries(C.fontes || {}).filter(([, v]) => v >= 0.1).map(([k, v]) => `${esc(k)} ${pct(v, 1)}`).join(" · ");
  $("#corpo").innerHTML = cartao(`Carteira de ${mesBR(mes)} por categoria`, `PL R$ ${num(C.pl_mm, 2)} mm · look-through (cotas de fundos abertas até o ativo)${conf}`, tab +
    `<p class="nota">Spread CDI+ = equivalente em CDI (IPCA+ e prefixados contra a curva do Tesouro de mesma duration). Cobertura de 100% dos ativos de crédito; ações, FII, FIP e ETF não têm spread. Como cada número é calculado: no fim da página. Fonte, em % do PL: ${fontes}.</p>`) +
    `<div class="cartao"><div class="cab" style="margin-bottom:8px"><div><h3>Evolução por categoria</h3><p class="sub">Ligue/desligue categorias e escolha o período (padrão: do início até a última carteira aberta)</p></div><div class="controles" id="ctl_graf"></div></div>
      <div id="f_graf"></div><div class="grade g2" style="margin:12px 0 0"><div><h3>% do PL por categoria</h3><div id="g_aloc"></div></div><div><h3>Spread CDI+ por categoria</h3><p class="sub">% a.a., média ponderada</p><div id="g_spread"></div></div></div></div>
    <div class="cartao"><h3>Ativos de ${mesBR(mes)}</h3><p class="sub">Maiores posições por % do PL · filtre por categoria ou busque por ativo, emissor, grupo econômico ou código</p>
      <div id="f_at"></div><div class="grade g-tabela" style="margin:12px 0 0"><div id="t_top"></div><div><h3>Maiores grupos econômicos</h3><p class="sub">% do PL nas categorias escolhidas</p><div id="g_grupos"></div></div></div></div>
    <div class="cartao" id="metodo"></div>`;

  graficosCarteira(D, ms, aberto);
  const cores = st.V.cat_cores;

  $("#t_top").innerHTML = `<div class="vazio">Carregando ativos...</div>`;
  const k = f.id + mes;
  if (!st.A[k]) st.A[k] = await (await fetch(`/api/ativos/${f.id}?mes=${mes}`)).json();
  if (st.fundo !== f.id || st.aba !== "carteira" || !$("#t_top")) return;
  const A = st.A[k].map((x) => ({ ...x, ativo: x.rotulo }));
  const peso = {};
  A.forEach((x) => (peso[x.categoria] = (peso[x.categoria] || 0) + x.perc));
  const todas = Object.keys(peso).sort((a, b) => peso[b] - peso[a]);
  const cols = [{ t: "Ativo", k: "ativo", w: 1 }, { t: "Categoria", k: "categoria" }, { t: "Grupo econômico", k: "grupo", w: 1 },
    { t: "% PL", k: "perc", n: 1, f: (v) => pct(v) }, { t: "Spread CDI+", k: "spread", n: 1, f: (v) => num(v) },
    { t: "Duration", k: "duration", n: 1, f: (v) => num(v) }, { t: "PU (R$)", k: "pu", n: 1, f: (v) => num(v) }, { t: "Fonte", k: "fonte", f: fonteCurta }];
  const redesenha = () => {
    const F = st.cf, lista = A.filter((x) => passa(F, x)).sort((a, b) => b.perc - a.perc);
    const soma = lista.reduce((s, x) => s + x.perc, 0);
    $("#t_top").innerHTML = tabela(cols, corta(lista, F.n), { alt: 640 }) +
      `<p class="nota">${lista.length} ativo(s) · ${pct(soma)} do PL nas categorias escolhidas${F.n && lista.length > F.n ? ` · mostrando ${F.n}` : ""}</p>`;
    const gp = {};
    lista.forEach((x) => (gp[x.grupo] = (gp[x.grupo] || 0) + x.perc));
    const g = corta(Object.entries(gp).sort((a, b) => b[1] - a[1]), F.n || 40).reverse();
    plota($("#g_grupos"), [{ type: "bar", orientation: "h", x: g.map((x) => x[1]), y: g.map((x) => x[0]), marker: { color: cores[0] }, hovertemplate: "%{y}: %{x:.2f}% do PL<extra></extra>" }],
      layout({ hovermode: "closest", margin: { l: 8, r: 14, t: 6, b: 30 }, yaxis: { automargin: true, gridcolor: "rgba(0,0,0,0)" }, xaxis: { ticksuffix: "%", gridcolor: css("--line") } }), Math.max(220, 24 * g.length + 50));
  };
  filtros($("#f_at"), todas, st.cf, redesenha, "Buscar ativo, emissor, grupo ou código...");
  redesenha();
  $("#metodo").innerHTML = metodologia(A);
}

// ------------------------------------------------------------------ RETORNO
function periodoMeses(ms, chave) {
  if (!ms.includes(st[chave].de)) st[chave].de = ms[Math.max(0, ms.length - 12)];
  if (!ms.includes(st[chave].ate)) st[chave].ate = ms[ms.length - 1];
  const opt = (v) => ms.slice().reverse().map((m) => `<option value="${m}" ${m === v ? "selected" : ""}>${mesBR(m)}</option>`).join("");
  return `<label>De <select id="de">${opt(st[chave].de)}</select></label><label>até <select id="ate">${opt(st[chave].ate)}</select></label>
    ${[["6m", 6], ["12m", 12], ["24m", 24], ["No ano", "ano"], ["Tudo", 999]].map(([t, n]) => `<button class="botao" data-p="${n}">${t}</button>`).join("")}`;
}
function ligaPeriodo(ms, chave) {
  $("#de").onchange = (e) => { st[chave].de = e.target.value; desenha(); };
  $("#ate").onchange = (e) => { st[chave].ate = e.target.value; desenha(); };
  document.querySelectorAll("[data-p]").forEach((b) => (b.onclick = () => {
    const ult = ms[ms.length - 1], n = b.dataset.p;
    st[chave].ate = ult;
    st[chave].de = n === "ano" ? (ms.find((m) => m.startsWith(ult.slice(0, 4))) || ms[0]) : ms[Math.max(0, ms.length - +n)];
    desenha();
  }));
}
function estatPeriodo(f, i0, i1) {
  // retorno, vol anual e drawdown da cota entre os indices i0..i1 de V.datas (precisa existir no inicio e no fim)
  const c = f.cota; if (!c || !c.length) return null;
  let a = i0; while (a <= i1 && c[a] == null) a++;
  if (a > i0 + 6 || a > i1) return null;
  let b = i1; while (b > a && c[b] == null) b--;
  if (b < i1 - 6) return null;
  const rets = []; let pico = c[a], dd = 0, ant = c[a];
  for (let k = a + 1; k <= b; k++) { if (c[k] == null) continue; rets.push(Math.log(c[k] / ant)); ant = c[k]; pico = Math.max(pico, c[k]); dd = Math.min(dd, c[k] / pico - 1); }
  const m = rets.reduce((s, x) => s + x, 0) / Math.max(rets.length, 1);
  const vol = Math.sqrt(rets.reduce((s, x) => s + (x - m) ** 2, 0) / Math.max(rets.length - 1, 1) * 252);
  const cdi = st.V.cdi[b] / st.V.cdi[a] - 1, ret = c[b] / c[a] - 1;
  return { ret: ret * 100, cdi: cdi * 100, pcdi: cdi ? (ret / cdi) * 100 : null, vol: vol * 100, dd: dd * 100 };
}
function indicesDatas(de, ate) {
  const D = st.V.datas;
  let i0 = D.findIndex((d) => d >= de); if (i0 < 0) i0 = D.length - 1;
  let i1 = D.length - 1; while (i1 > 0 && D[i1] > ate) i1--;
  return [Math.max(0, i0 - 1), i1];   // inclui a cota do dia anterior ao inicio (retorno do 1o dia)
}
async function retorno(el) {
  const f = fundoPor(st.fundo);
  cabecalhoFundo(el, f);
  if (!f.tem_carteira) { $("#corpo").innerHTML = `<div class="vazio">A CVM não tem carteira deste fundo no período.</div>`; return; }
  $("#corpo").innerHTML = `<div class="vazio">Carregando ${esc(f.nome)}...</div>`;
  const D = await detalhe(f.id);
  if (st.fundo !== f.id || st.aba !== "retorno") return;
  const R = D.retorno, ms = R.meses.map((x) => x.mes);
  if (!ms.length) { $("#corpo").innerHTML = `<div class="vazio">Sem meses suficientes para calcular retorno.</div>`; return; }
  $("#ctl").insertAdjacentHTML("beforeend", periodoMeses(ms, "ret"));
  ligaPeriodo(ms, "ret");
  const j0 = ms.indexOf(st.ret.de), j1 = ms.indexOf(st.ret.ate);
  const [a, b] = j0 <= j1 ? [j0, j1] : [j1, j0];
  const J = Array.from({ length: b - a + 1 }, (_, k) => a + k);
  const tot = J.reduce((p, j) => p * (1 + R.meses[j].real / 100), 1) - 1, cdi = J.reduce((p, j) => p * (1 + R.meses[j].cdi / 100), 1) - 1;
  const linhas = Object.entries(R.categorias).map(([c, v]) => {
    let contrib = 0, ret = 1, nP = 0, sP = 0, nR = 0;
    J.forEach((j) => { const cc = v.contrib[j] || 0, p = v.peso[j] || 0; contrib += cc; if (p > 0.05) { ret *= 1 + cc / p; nR++; } sP += p; nP++; });
    return { categoria: c, peso: sP / Math.max(nP, 1), contrib, ret: nR ? (ret - 1) * 100 : null, pcdi: nR && cdi ? ((ret - 1) / cdi) * 100 : null };
  }).filter((r) => Math.abs(r.peso) > 0.05 || Math.abs(r.contrib) > 0.005).sort((x, y) => y.contrib - x.contrib);
  linhas.push({ _cls: "total", categoria: "Fundo (retorno da cota)", peso: 100, contrib: tot * 100, ret: tot * 100, pcdi: cdi ? (tot / cdi) * 100 : null });
  const soma = {}, pesoAt = {};
  R.contrib.forEach(([ia, jm, v, w]) => { if (jm >= a && jm <= b) { soma[ia] = (soma[ia] || 0) + v; pesoAt[ia] = (pesoAt[ia] || 0) + (w || 0); } });
  const ativos = Object.keys(soma).map((ia) => ({ ...R.ativos[ia], contrib: soma[ia], perc: pesoAt[ia] / J.length }));
  const hedge = J.reduce((s, j) => s + (R.meses[j].hedge || 0), 0);
  const per = `${mesBR(ms[a])} a ${mesBR(ms[b])}`;

  $("#corpo").innerHTML = `<div class="kpis">${kpi("Retorno do fundo", pct(tot * 100), per, cls(tot))}${kpi("CDI", pct(cdi * 100), "mesmo período")}
      ${kpi("% do CDI", pct(cdi ? (tot / cdi) * 100 : null, 1), "")}${kpi("Excesso sobre o CDI", pp((tot - cdi) * 100), "", cls(tot - cdi))}${kpi("Derivativos", pp(hedge), "já alocado nos ativos protegidos", cls(hedge))}</div>` +
    cartao("Retorno por categoria", "Contribuição em pontos percentuais do PL; retorno e % do CDI da própria categoria no período", tabela([
      { t: "Categoria", k: "categoria", w: 1 }, { t: "Peso médio", k: "peso", n: 1, f: (v) => pct(v, 1) }, { t: "Contribuição", k: "contrib", n: 1, f: (v) => `<span class="${cls(v)}">${pp(v)}</span>` },
      { t: "Retorno", k: "ret", n: 1, f: (v) => pct(v) }, { t: "% do CDI", k: "pcdi", n: 1, f: (v) => pct(v, 0) }], linhas) +
      `<p class="nota">Estimado com as carteiras mensais da CVM: cada ativo rende CDI + seu spread, ou a variação real de preço quando ela é crível; caixa rende CDI; derivativos (DAP → ativos IPCA+, DI1 → prefixados, dólar → ativos no exterior) entram nos ativos que protegem; a diferença para o retorno real da cota (taxas, negociação, marcação) é distribuída pelo peso. A soma das categorias fecha com o retorno real da cota.</p>`) +
    `<div class="cartao"><h3>Ativos por contribuição</h3><p class="sub">${per} · % PL = peso médio no período · filtre por categoria ou busque por ativo, emissor, grupo ou código</p>
      <div id="f_ret"></div><div class="grade g2" style="margin:12px 0 0"><div><h3>Melhores</h3><div id="t_mel"></div></div><div><h3>Piores</h3><div id="t_pio"></div></div></div></div>`;
  const peso = {};
  ativos.forEach((x) => (peso[x.categoria] = (peso[x.categoria] || 0) + Math.abs(x.perc)));
  const todas = Object.keys(peso).sort((x, y) => peso[y] - peso[x]);
  const cols = [{ t: "Ativo", k: "ativo", w: 1 }, { t: "Categoria", k: "categoria" }, { t: "Grupo econômico", k: "grupo", w: 1 },
    { t: "% PL", k: "perc", n: 1, f: (v) => pct(v) }, { t: "Contribuição", k: "contrib", n: 1, f: (v) => `<span class="${cls(v)}">${pp(v, 3)}</span>` }];
  const redesenha = () => {
    const F = st.rf, lista = ativos.filter((x) => passa(F, x));
    $("#t_mel").innerHTML = tabela(cols, corta(lista.slice().sort((x, y) => y.contrib - x.contrib), F.n), { alt: 560 });
    $("#t_pio").innerHTML = tabela(cols, corta(lista.slice().sort((x, y) => x.contrib - y.contrib), F.n), { alt: 560 });
  };
  filtros($("#f_ret"), todas, st.rf, redesenha, "Buscar ativo, emissor, grupo ou código...");
  redesenha();
}
function riscoRetorno(el, linhas) {
  const grupo = (arr, nome, cor) => ({ type: "scatter", mode: "markers+text", name: nome, x: arr.map((r) => r.vol), y: arr.map((r) => r.pcdi),
    text: arr.map((r) => (r.f.nosso || arr.length <= 12 ? r.f.nome : "")), textposition: "top center", textfont: { size: 11, color: css("--text-2") },
    customdata: arr.map((r) => [r.f.nome, r.ret]), marker: { size: 10, color: cor, opacity: 0.9, line: { width: 1, color: css("--surface") } },
    hovertemplate: "%{customdata[0]}<br>retorno %{customdata[1]:.2f}% · %{y:.1f}% do CDI · vol %{x:.2f}%<extra></extra>" });
  plota(el, [grupo(linhas.filter((r) => !r.f.nosso), "Peers", "#9aa0a6"), grupo(linhas.filter((r) => r.f.nosso), "Nossos", "#eb6834")],
    layout({ hovermode: "closest", legend: { orientation: "h", y: -0.2 }, xaxis: { title: "volatilidade anual (%)", gridcolor: css("--line"), zeroline: false },
      yaxis: { title: "% do CDI", gridcolor: css("--line"), zeroline: false } }), 440);
}

// ------------------------------------------------------------------ CONSOLIDADO
function consolidado(el) {
  const V = st.V, C = st.cons, ult = V.datas[V.datas.length - 1];
  if (!C.ate) C.ate = ult;
  if (!C.de) { const d = new Date(ult + "T12:00"); d.setFullYear(d.getFullYear() - 1); C.de = d.toISOString().slice(0, 10); }
  el.innerHTML = `<div class="cab"><div class="titulo-fundo"><h2>Consolidado</h2><p class="sub">Retorno de cada fundo no período · só fundos ativos do início ao fim do período</p></div>
    <div class="controles"><label>De <input type="date" id="cde" value="${C.de}" min="${V.datas[0]}" max="${ult}"></label><label>até <input type="date" id="cate" value="${C.ate}" min="${V.datas[0]}" max="${ult}"></label>
    ${[["12m", 12], ["24m", 24], ["No ano", "ano"], ["Tudo", "tudo"]].map(([t, n]) => `<button class="botao" data-q="${n}">${t}</button>`).join("")}</div></div>
    <div class="cartao"><div class="controles" style="margin-bottom:8px"><span class="sub" style="margin:0">Fundos:</span>
      <button class="botao" data-s="todos">Todos</button><button class="botao" data-s="nossos">Só os nossos</button><button class="botao" data-s="peers">Só peers</button><button class="botao" data-s="nenhum">Limpar</button>
      <div id="busca_cons"></div><span class="sub" id="n_sel" style="margin:0"></span></div>
      <div class="multi" id="multi"></div></div>
    <div id="cc"></div>`;
  $("#cde").onchange = (e) => { C.de = e.target.value; desenha(); };
  $("#cate").onchange = (e) => { C.ate = e.target.value; desenha(); };
  document.querySelectorAll("[data-q]").forEach((b) => (b.onclick = () => {
    const d = new Date(ult + "T12:00"), q = b.dataset.q;
    C.ate = ult;
    C.de = q === "tudo" ? V.datas[0] : q === "ano" ? ult.slice(0, 4) + "-01-01" : (d.setMonth(d.getMonth() - +q), d.toISOString().slice(0, 10));
    desenha();
  }));
  document.querySelectorAll("[data-s]").forEach((b) => (b.onclick = () => {
    const s = b.dataset.s;
    C.sel = new Set(V.fundos.filter((f) => s === "todos" || (s === "nossos" && f.nosso) || (s === "peers" && !f.nosso)).map((f) => f.id));
    atualiza();
  }));
  // busca com selecao multipla: digitar, clicar liga/desliga o fundo (✓) sem fechar a lista
  $("#busca_cons").appendChild(campoBusca((x) => { C.sel.has(x.id) ? C.sel.delete(x.id) : C.sel.add(x.id); atualiza(); },
    "Adicionar/remover fundo: digite...", false, (id) => C.sel.has(id)));

  function atualiza() {
    const chips = V.fundos.filter((f) => C.sel.has(f.id)).sort((a, b) => b.nosso - a.nosso || a.nome.localeCompare(b.nome));
    $("#n_sel").textContent = `${chips.filter((f) => f.nosso).length} nossos + ${chips.filter((f) => !f.nosso).length} peers selecionados`;
    $("#multi").innerHTML = chips.map((f) => `<span class="chip ${f.nosso ? "nosso" : ""}">${esc(f.nome)}<button data-x="${f.id}" aria-label="Remover">×</button></span>`).join("") || '<span class="sub">Nenhum fundo selecionado</span>';
    $("#multi").querySelectorAll("[data-x]").forEach((b) => (b.onclick = () => { C.sel.delete(b.dataset.x); atualiza(); }));
    resultados();
  }

  function resultados() {
    const [i0, i1] = indicesDatas(C.de, C.ate);
    const linhas = [], fora = [];
    V.fundos.filter((f) => C.sel.has(f.id)).forEach((f) => { const e = estatPeriodo(f, i0, i1); (e ? linhas : fora).push({ f, ...(e || {}) }); });
    const cdiP = (V.cdi[i1] / V.cdi[i0] - 1) * 100;
    const ord = C.ord, sgn = C.asc ? 1 : -1;
    linhas.sort((x, y) => sgn * ((x[ord] ?? -1e9) - (y[ord] ?? -1e9)));
    $("#cc").innerHTML = cartao("Retorno acumulado", `${dataBR(V.datas[i0 + 1] || V.datas[i0])} a ${dataBR(V.datas[i1])} · CDI ${pct(cdiP)} · laranja = nossos · cinza = peers`, '<div id="g_cons"></div>') +
      cartao("Risco x retorno", "Mesmo período e mesmos fundos · laranja = nossos · cinza = peers", '<div id="g_rr"></div>') +
      cartao("Ranking no período", fora.length ? `${fora.length} fundo(s) selecionado(s) fora por não existirem no período inteiro: ${fora.map((x) => esc(x.f.nome)).join(", ")}` : "Clique no título da coluna para ordenar",
        tabela([{ t: "#", k: "_i" }, { t: "Fundo", k: "nome", w: 1, f: (v, r) => `<span class="ponto" style="background:${r.f.nosso ? "var(--nosso)" : "#9aa0a6"}"></span>${esc(r.f.nome)}${r.f.nosso ? ' <span class="tag">NOSSO</span>' : ""}${r.f.distribui ? ' <span class="muted" title="Distribui rendimentos: retorno total, somando as distribuições">· distribui</span>' : ""}` },
          { t: "Retorno", k: "ret", n: 1, ord: "ret", f: (v) => `<span class="${cls(v)}">${pct(v)}</span>` }, { t: "% do CDI", k: "pcdi", n: 1, ord: "pcdi", f: (v) => pct(v, 1) },
          { t: "Vol. anual", k: "vol", n: 1, ord: "vol", f: (v) => pct(v) }, { t: "Pior drawdown", k: "dd", n: 1, ord: "dd", f: (v) => pct(v) },
          { t: "PL (R$ mm)", k: "pl", n: 1, ord: "pl", f: (v) => num(v, 0) }],
          linhas.map((r, i) => ({ ...r, _i: i + 1, nome: r.f.nome, pl: r.f.pl, _cls: r.f.nosso ? "nosso" : "" }))));
    document.querySelectorAll("th[data-ord]").forEach((th) => (th.onclick = () => { if (C.ord === th.dataset.ord) C.asc = !C.asc; else { C.ord = th.dataset.ord; C.asc = false; } resultados(); }));
    const xs = V.datas.slice(i0, i1 + 1);
    const serie = (f) => { const c = f.cota.slice(i0, i1 + 1); const b = c.find((v) => v != null); return c.map((v) => (v == null ? null : (v / b - 1) * 100)); };
    const dados = linhas.filter((r) => !r.f.nosso).map((r) => ({ x: xs, y: serie(r.f), name: r.f.nome, mode: "lines", line: { color: "#9aa0a6", width: 1.2 }, opacity: 0.55, showlegend: false,
      hovertemplate: `${esc(r.f.nome)}: %{y:.2f}%<extra></extra>` }));
    linhas.filter((r) => r.f.nosso).forEach((r) => dados.push({ x: xs, y: serie(r.f), name: r.f.nome, mode: "lines", line: { color: r.f.cor || "#eb6834", width: 2.6 }, hovertemplate: `${esc(r.f.nome)}: %{y:.2f}%<extra></extra>` }));
    const c0 = V.cdi[i0];
    dados.push({ x: xs, y: V.cdi.slice(i0, i1 + 1).map((v) => (v / c0 - 1) * 100), name: "CDI", mode: "lines", line: { color: css("--text-1"), width: 2, dash: "dash" }, hovertemplate: "CDI: %{y:.2f}%<extra></extra>" });
    plota($("#g_cons"), dados, layout({ hovermode: "closest", yaxis: { ticksuffix: "%", gridcolor: css("--line") } }), 460);
    riscoRetorno($("#g_rr"), linhas);
  }
  atualiza();
}

</script>
</body>
</html>
"""


def gera_excel():
    CNPJ = {n: cnpj_of(x) for n, x in {**NOSSOS_FUNDOS, **PEERS}.items()}
    meses, dia = prepara(sorted(CNPJ.values()), DESDE)
    print(f"{len(meses)} meses de carteira ({meses[0]} a {meses[-1]})")
    df = lookthrough(CNPJ, meses)
    soma = df[~df.derivativo].groupby(["fundo", "mes"]).perc_pl.sum()
    print("soma do %PL por carteira (sem derivativos): min", round(soma.min(), 3), "max", round(soma.max(), 3))
    grandes = df[df.categoria.eq(AJUSTE) & (df.perc_pl.abs() > 1)]
    if len(grandes):
        print("ajustes carteira x PL acima de 1% do peer (veiculo onde a CDA nao fecha com o PL):")
        print(grandes.groupby(["fundo", "veiculo"]).perc_pl.agg(["count", "min", "max"]).round(2).to_string())
    cruz = df[df.cnpj_veiculo.isin(CNPJ.values()) & (df.cnpj_veiculo != df.fundo.map(CNPJ))]
    print("peers investindo em outro peer:", "nenhum" if cruz.empty else sorted(set(zip(cruz.fundo, cruz.veiculo))))
    rent = rentabilidade(diario(CNPJ, dia), cdi_mensal(DESDE))
    excel(df, rent, list(CNPJ), OUT)
    print("salvo", OUT, "e", OUT.replace(".xlsx", "_dados.csv"), len(df), "linhas")


if __name__ == "__main__":                 # importante: os processos auxiliares do Windows reimportam este arquivo
    if len(sys.argv) > 1 and sys.argv[1].lower() == "excel":
        gera_excel()
    elif len(sys.argv) == 1 or sys.argv[1].lower() == "html":
        import webbrowser
        print(f"Atualizando: {len({cnpj_of(c) for c in list(NOSSOS_FUNDOS.values()) + list(PEERS.values())})} fundos desde {DESDE}. Na 1a vez pode levar bastante tempo...")
        precarrega(log=lambda m: print("  " + m, flush=True))
        if PRE["status"] != "ok":
            sys.exit("Nao foi possivel preparar o painel: " + str(PRE["erro"]))
        arq = exporta_html(PRE["pasta"], os.path.join(os.path.dirname(os.path.abspath(__file__)), "painel_fundos.html"),
                           log=lambda m: print("  " + m, flush=True))
        webbrowser.open(arq.as_uri())
    else:
        import webbrowser
        import uvicorn
        print(f"Preparando o painel: {len({cnpj_of(c) for c in list(NOSSOS_FUNDOS.values()) + list(PEERS.values())})} fundos desde {DESDE}. Na 1a vez pode levar bastante tempo...")
        precarrega(log=lambda m: print("  " + m, flush=True))
        if PRE["status"] != "ok":
            sys.exit("Nao foi possivel preparar o painel: " + str(PRE["erro"]))
        porta = int(os.environ.get("PORT", 7860))
        threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{porta}")).start()
        print(f"Pronto. Painel em http://localhost:{porta}  (deixe esta janela aberta; Ctrl+C para fechar)")
        uvicorn.run(app, host="127.0.0.1", port=porta, log_level="warning")
