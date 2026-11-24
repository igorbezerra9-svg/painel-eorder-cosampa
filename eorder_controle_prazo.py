# -*- coding: utf-8 -*-
"""Controle de Ordens no Prazo -- Gerência Sul (Ligações Novas e Religações por Centro de Serviço).

Replica a planilha "PLANILHA DA COMERCIAL - Centro sul.xlsm" (aba ORDENS NO PRAZO), mas com o "fora do prazo" CALCULADO em vez
de digitado:

  * Linhas = export do TdC (Busca TdCs, busca salva CONTROLE_PRAZO_SUL: Estado Finalizado/Encerrado, Data Resultado do dia 1 até
    o fim do período). Cada linha é uma OS finalizada.
  * Cada ORDEM conta uma vez só (como na planilha, que remove duplicadas): o export traz a mesma OS em mais de uma linha quando ela
    teve mais de um TdC; vale o TdC mais recente. Conferido: com isso os totais de out/2026 reproduzem a planilha (3193 e 4936).
  * Grupo: "Ligação Nova" = Tipo TdC Ligação Nova + Verificação Comercial + Alteração de Ligação; "Religação" = Religação.
    (Corte não entra no controle.)
  * Centro de Serviço: pelo município (mesma tabela da aba Planilha2 da planilha).
  * Fora do prazo: a PRIMEIRA baixa da OS (produtiva ou improdutiva; devolvida não conta -- mesma regra da aba Cumprimento de
    Prazo; se só houver devolvida, vale a primeira de qualquer tipo) aconteceu DEPOIS da "Data Prevista Finalização Trabalhos".
    Ex.: vencia 03/10 10:00 e foi baixada 03/10 11:00 = fora do prazo. Conferido contra a planilha: 0/0/0 em Ligações Novas e
    0/1/0 em Religações (outubro/2026).
  * Realizado % = (executadas - fora do prazo) / executadas; Avanço % = realizado / meta (99,5%).

Só biblioteca padrão: roda no robô da VPS (eorder_execucao_bot.py) e nos testes sem rede.
"""
import unicodedata
from datetime import datetime, date, timedelta

META = 0.995

GRUPO_POR_TIPO_TDC = {
    "Ligação Nova": "ligacao_nova",
    "Verificação Comercial": "ligacao_nova",
    "Alteração de Ligação": "ligacao_nova",
    "Religação": "religacao",
}
GRUPOS = ("ligacao_nova", "religacao")
CENTROS = ("CAMPOS SALES", "JUAZEIRO", "MILAGRES")

# Município -> Centro de Serviço (aba Planilha2 da planilha da Comercial, 29 municípios).
CS_POR_MUNICIPIO = {
    "ASSARE": "CAMPOS SALES", "SANTANA DO CARIRI": "JUAZEIRO", "JUAZEIRO DO NORTE": "JUAZEIRO", "JARDIM": "JUAZEIRO",
    "MAURITI": "MILAGRES", "PENAFORTE": "MILAGRES", "CRATO": "JUAZEIRO", "BARBALHA": "JUAZEIRO", "MILAGRES": "MILAGRES",
    "TARRAFAS": "CAMPOS SALES", "MISSAO VELHA": "MILAGRES", "BARRO": "MILAGRES", "BREJO SANTO": "MILAGRES",
    "AURORA": "MILAGRES", "SABOEIRO": "CAMPOS SALES", "CAMPOS SALES": "CAMPOS SALES", "SALITRE": "CAMPOS SALES",
    "FARIAS BRITO": "JUAZEIRO", "PORTEIRAS": "MILAGRES", "CARIRIACU": "JUAZEIRO", "NOVA OLINDA": "JUAZEIRO",
    "GRANJEIRO": "JUAZEIRO", "AIUABA": "CAMPOS SALES", "ARARIPE": "CAMPOS SALES", "POTENGI": "CAMPOS SALES",
    "ANTONINA DO NORTE": "CAMPOS SALES", "ALTANEIRA": "JUAZEIRO", "JATI": "MILAGRES", "ABAIARA": "MILAGRES",
}

DIAS_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]


def _sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s or "")) if unicodedata.category(c) != "Mn")


def chave_municipio(s):
    return " ".join(_sem_acento(s).upper().split())


def centro_servico(municipio):
    return CS_POR_MUNICIPIO.get(chave_municipio(municipio), "")


def parse_dt(s):
    """'dd/mm/aaaa hh:mm' (formato do eOrder) -> datetime; None se não der. Aceita também datetime (célula de data do xlsx)."""
    if isinstance(s, datetime):
        return s.replace(second=0, microsecond=0)
    try:
        return datetime.strptime(str(s).strip()[:16], "%d/%m/%Y %H:%M")
    except Exception:
        return None


def primeira_baixa(baixas):
    """baixas = [(datetime, classe)], classe em PRODUTIVO/IMPRODUTIVO/DEVOLVIDO/None. Devolvida só vale se não houver outra."""
    if not baixas:
        return None
    uteis = [t for t, c in baixas if c in ("PRODUTIVO", "IMPRODUTIVO")]
    return min(uteis) if uteis else min(t for t, _ in baixas)


def dia_do_resultado(baixas, ate):
    """Dia (date) em que a OS foi finalizada, pra filtrar por calendário. O export do TdC não traz a Data Resultado, então vem das
    baixas: a última produtiva até o fim do período; sem produtiva, a última produtiva/improdutiva; sem nenhuma, a última de qualquer
    tipo. Baixas depois do fim do período são ignoradas (OS finalizada no período que já voltou a ser atendida depois)."""
    if not baixas:
        return None
    ate_dt = datetime.combine(ate, datetime.max.time())
    dentro = [(t, c) for t, c in baixas if t <= ate_dt] or list(baixas)
    prod = [t for t, c in dentro if c == "PRODUTIVO"]
    uteis = [t for t, c in dentro if c in ("PRODUTIVO", "IMPRODUTIVO")]
    return (max(prod) if prod else max(uteis) if uteis else max(t for t, _ in dentro)).date()


def periodos(hoje):
    """(mes_atual, mes_anterior) como (de, ate) em date. O mês atual vai do dia 1 até ONTEM; no dia 1 ele ainda não existe (None).
    O mês anterior é sempre o mês fechado inteiro."""
    ontem = hoje - timedelta(days=1)
    primeiro_atual = hoje.replace(day=1)
    ultimo_anterior = primeiro_atual - timedelta(days=1)
    anterior = (ultimo_anterior.replace(day=1), ultimo_anterior)
    atual = (primeiro_atual, ontem) if ontem >= primeiro_atual else None
    return atual, anterior


def titulo_data(ref):
    """02/11/segunda-feira (como na planilha)."""
    return f"{ref:%d/%m}/{DIAS_SEMANA[ref.weekday()]}"


def _recencia(r):
    return (parse_dt(r.get("Data e Hora de Criação TdC")) or datetime.min, str(r.get("Código TdC") or ""))


def uma_linha_por_ordem(linhas):
    """Mantém 1 linha por (ordem, grupo): a do TdC mais recente. Linhas sem número de ordem ficam como estão. Devolve (linhas, repetidas)."""
    escolhidas, sem_ordem = {}, []
    for r in linhas:
        os_num = str(r.get("Numero de Serviço") or "").strip()
        if not os_num:
            sem_ordem.append(r)
            continue
        chave = (os_num, GRUPO_POR_TIPO_TDC.get(str(r.get("Tipo TdC") or "").strip()))
        atual = escolhidas.get(chave)
        if atual is None or _recencia(r) > _recencia(atual):
            escolhidas[chave] = r
    return list(escolhidas.values()) + sem_ordem, len(linhas) - len(escolhidas) - len(sem_ordem)


def calcular(linhas, baixas_por_os, de, ate, referencia=None, agora=None, uma_por_ordem=True):
    """Monta o payload que o painel mostra.

    linhas: dicts do export do TdC (chaves 'Numero de Serviço', 'Tipo TdC', 'Município', 'Data Prevista Finalização Trabalhos').
    baixas_por_os: {'A0104...': [(datetime, classe), ...]} -- baixas da Execução (inclui meses anteriores ao período, pra não
    perder a primeira tentativa de uma OS que só foi concluída agora).
    """
    por_cs = {g: {cs: {"executadas": 0, "fora_do_prazo": 0} for cs in CENTROS} for g in GRUPOS}
    repetidas = 0
    if uma_por_ordem:
        linhas, repetidas = uma_linha_por_ordem(linhas)
    fora, avisos = [], {"municipios_sem_cs": {}, "sem_baixa": 0, "sem_prazo": 0, "tipos_ignorados": 0, "linhas_repetidas_ignoradas": repetidas}
    por_dia = {}          # 'aaaa-mm-dd' (ou 'sem_data') -> grupo -> CS -> [executadas, fora do prazo]  (alimenta o calendário do painel)
    for r in linhas:
        grupo = GRUPO_POR_TIPO_TDC.get(str(r.get("Tipo TdC") or "").strip())
        if not grupo:
            avisos["tipos_ignorados"] += 1          # Corte e qualquer tipo novo: fora do controle
            continue
        mun = str(r.get("Município") or "").strip()
        cs = centro_servico(mun)
        if not cs:
            k = chave_municipio(mun) or "(vazio)"
            avisos["municipios_sem_cs"][k] = avisos["municipios_sem_cs"].get(k, 0) + 1
            continue
        por_cs[grupo][cs]["executadas"] += 1
        os_num = str(r.get("Numero de Serviço") or "").strip()
        prazo = parse_dt(r.get("Data Prevista Finalização Trabalhos"))
        hist = baixas_por_os.get(os_num)
        baixa = primeira_baixa(hist)
        dia = dia_do_resultado(hist, ate)
        celula = por_dia.setdefault(dia.isoformat() if dia else "sem_data", {g: {c: [0, 0] for c in CENTROS} for g in GRUPOS})[grupo][cs]
        celula[0] += 1
        if not prazo:
            avisos["sem_prazo"] += 1
            continue
        if not baixa:
            avisos["sem_baixa"] += 1                # sem baixa na Execução: não dá pra julgar, conta como no prazo
            continue
        if baixa > prazo:
            por_cs[grupo][cs]["fora_do_prazo"] += 1
            celula[1] += 1
            fora.append({"os": os_num, "grupo": grupo, "cs": cs, "municipio": mun, "prazo": prazo.strftime("%d/%m/%Y %H:%M"),
                         "baixa": baixa.strftime("%d/%m/%Y %H:%M"), "atraso_h": round((baixa - prazo).total_seconds() / 3600, 1),
                         "dia": dia.isoformat() if dia else ""})
    fora.sort(key=lambda x: -x["atraso_h"])
    ref = referencia or ate
    return {
        "periodo": {"de": de.isoformat(), "ate": ate.isoformat(), "mes": de.strftime("%Y-%m"), "fechado": ate.month != (ate + timedelta(days=1)).month},
        "referencia": ref.isoformat(),
        "titulo_data": titulo_data(ref),
        "meta": META,
        "grupos": por_cs,
        "por_dia": dict(sorted(por_dia.items())),
        "fora_do_prazo": fora[:200],
        "avisos": avisos,
        "gerado_em": (agora or datetime.now()).isoformat(timespec="seconds"),
    }


def resumo(grupo_cs):
    """{'CAMPOS SALES': {...}, ...} -> linhas + total, com realizado e avanço (o painel calcula igual; aqui serve pros testes)."""
    out = {}
    tot_e = tot_f = 0
    for cs in CENTROS:
        e, f = grupo_cs[cs]["executadas"], grupo_cs[cs]["fora_do_prazo"]
        tot_e += e; tot_f += f
        out[cs] = (e, f, (e - f) / e if e else 0.0)
    out["TOTAL"] = (tot_e, tot_f, (tot_e - tot_f) / tot_e if tot_e else 0.0)
    return out
