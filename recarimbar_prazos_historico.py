# -*- coding: utf-8 -*-
"""Reescreve o prazo nas linhas da Execução que já estão no Supabase.

O robô carimba o prazo (`Data Prevista Finalização Trabalhos`) na linha da
Execução na hora de publicar, usando o índice local prazos_os_idx.json. Linha
já publicada nunca é revisitada -- então tudo que entrou no índice DEPOIS
(inclusive o que o baixar_prazos_finalizados.py acabou de trazer) não aparece
no histórico, e a aba "Ordens no Prazo" continua sem conseguir julgar as
ordens antigas.

Este script faz a mesma anexação, mas em cima do que já está gravado em
snapshots_historico (regiao='Execucao'), dia por dia.

REGRA (a mesma do baixar_prazos_finalizados.py): só escreve quando a linha
não tem prazo nenhum, ou quando o prazo do índice é MAIS NOVO que o que está
na linha. Prazo não anda pra trás -- reabertura e reprogramação empurram pra
frente, e sobrescrever com um prazo de ciclo anterior faz ordem executada no
prazo virar "fora do prazo".

Egress: baixa o `dados` de cada dia pedido (com gzip, ~65 KB/dia na Execução).
Rode com --desde pra não varrer mais do que precisa. Só regrava o dia que
mudou de verdade. Com --cache DIR o dia baixado fica guardado em disco e o run
seguinte (a simulação e depois a gravação de verdade) não paga egress outra
vez -- a cota do Supabase conta o que sai, e já estourou uma vez por causa de
download repetido (ver CLAUDE.md).

    python recarimbar_prazos_historico.py --desde 2026-09-01 --cache /tmp/hist
    python recarimbar_prazos_historico.py --desde 2026-09-01 --cache /tmp/hist --aplicar
"""
import argparse
import gzip
import io
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime

from eorder_execucao_bot import (
    SB_URL, SB_KEY, PRAZOS_OS_PATH, PRAZO_COL,
    _parse_data_fim, _carregar_json_idx,
)

REGIAO = "Execucao"
COL_NUM = "Numero de Serviço"
HDR = {"apikey": SB_KEY, "Authorization": "Bearer " + SB_KEY}


def log(msg):
    linha = f"[{datetime.now():%H:%M:%S}] {msg}"
    try:
        print(linha, flush=True)
    except UnicodeEncodeError:
        # console do Windows em cp1252 engasga com emoji; na VPS (UTF-8) não cai aqui
        print(linha.encode("ascii", "replace").decode("ascii"), flush=True)


def _get(caminho):
    """GET no PostgREST pedindo gzip -- o `dados` de um dia da Execução passa
    de 500 KB cru e cai pra ~80 KB comprimido; a cota de egress do Supabase
    conta o que sai, então sempre comprimido (ver CLAUDE.md)."""
    req = urllib.request.Request(SB_URL + "/rest/v1/" + caminho,
                                 headers={**HDR, "Accept-Encoding": "gzip"})
    with urllib.request.urlopen(req, timeout=180) as resp:
        bruto = resp.read()
        # na rede trafegou o comprimido -- é esse tamanho que a cota conta
        tam = len(bruto)
        if (resp.headers.get("Content-Encoding") or "").lower() == "gzip":
            bruto = gzip.GzipFile(fileobj=io.BytesIO(bruto)).read()
    return json.loads(bruto), tam


def _dados_do_dia(data, cache_dir):
    """O `dados` de um dia, do cache se já estiver em disco. Devolve
    (linhas, bytes_baixados) -- 0 byte quando veio do cache."""
    arq = os.path.join(cache_dir, data + ".json.gz") if cache_dir else None
    if arq and os.path.exists(arq):
        with gzip.open(arq, "rt", encoding="utf-8") as f:
            return json.load(f), 0
    reg, tam = _get(f"snapshots_historico?regiao=eq.{REGIAO}&data=eq.{data}&select=dados")
    linhas = (reg[0].get("dados") or []) if reg else []
    if arq:
        os.makedirs(cache_dir, exist_ok=True)
        with gzip.open(arq, "wt", encoding="utf-8") as f:
            json.dump(linhas, f, ensure_ascii=False)
    return linhas, tam


def _patch_dia(data, linhas):
    corpo = json.dumps({"dados": linhas}, default=str).encode("utf-8")
    req = urllib.request.Request(
        f"{SB_URL}/rest/v1/snapshots_historico?regiao=eq.{REGIAO}&data=eq.{data}",
        data=corpo, method="PATCH",
        headers={**HDR, "Content-Type": "application/json", "Prefer": "return=minimal"},
    )
    with urllib.request.urlopen(req, timeout=180) as resp:
        return resp.status


def recarimbar(desde, ate, aplicar, cache_dir=None):
    idx = _carregar_json_idx(PRAZOS_OS_PATH) or {}
    if not idx:
        sys.exit(f"❌ Índice vazio ou não encontrado: {PRAZOS_OS_PATH}")
    log(f"📦 Índice: {len(idx)} ordens com prazo")

    filtro = f"snapshots_historico?regiao=eq.{REGIAO}&select=data&order=data"
    if desde:
        filtro += f"&data=gte.{desde}"
    if ate:
        filtro += f"&data=lte.{ate}"
    dias = [d["data"] for d in _get(filtro)[0]]
    if not dias:
        sys.exit("❌ Nenhum dia de histórico da Execução nesse período.")
    log(f"🗓  {len(dias)} dia(s) no histórico: {dias[0]} a {dias[-1]}")

    # Cache: o prazo parseado é reusado em todo dia que tem a mesma ordem.
    visto = {}

    def prazo_idx(n):
        if n not in visto:
            p = idx.get(n)
            visto[n] = (p, _parse_data_fim(p)) if p else (None, None)
        return visto[n]

    t_preenchidos = t_trocados = t_dias = 0
    baixado = 0
    for data in dias:
        linhas, tam = _dados_do_dia(data, cache_dir)
        baixado += tam
        preenchidos = trocados = 0
        for r in linhas:
            n = (r.get(COL_NUM) or "").strip()
            if not n:
                continue
            p, pd = prazo_idx(n)
            if not p or not pd:
                continue
            atual = (r.get(PRAZO_COL) or "").strip()
            ad = _parse_data_fim(atual) if atual else None
            if not ad:
                # sem prazo, ou com um prazo que não dá pra ler
                if atual == p:
                    continue
                r[PRAZO_COL] = p
                preenchidos += 1
            elif ad < pd:
                r[PRAZO_COL] = p
                trocados += 1
        if not (preenchidos or trocados):
            continue
        t_dias += 1
        t_preenchidos += preenchidos
        t_trocados += trocados
        marca = "" if aplicar else " (teste)"
        log(f"   {data}: {len(linhas)} linhas → +{preenchidos} com prazo, "
            f"{trocados} trocado(s) por mais novo{marca}")
        if aplicar:
            try:
                _patch_dia(data, linhas)
            except urllib.error.HTTPError as e:
                log(f"   ⚠️  {data} NÃO regravado: {e.code} "
                    f"{e.read().decode('utf-8', 'replace')[:200]}")
            except Exception as e:
                log(f"   ⚠️  {data} NÃO regravado: {e}")

    log(f"✅ {t_dias} dia(s) com mudança — {t_preenchidos} linha(s) ganharam prazo, "
        f"{t_trocados} tiveram o prazo atualizado")
    log(f"📉 Egress deste run: {baixado / 1048576:.1f} MB (comprimido)")
    if not aplicar:
        log("🔍 MODO TESTE -- nada foi gravado. Rode com --aplicar pra valer.")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--desde", help="Primeiro dia (aaaa-mm-dd)")
    ap.add_argument("--ate", help="Último dia (aaaa-mm-dd)")
    ap.add_argument("--cache", metavar="DIR",
                    help="Guarda/reaproveita o histórico baixado (economiza egress)")
    ap.add_argument("--aplicar", action="store_true",
                    help="Regrava no Supabase (sem isso é só simulação)")
    a = ap.parse_args()
    recarimbar(a.desde, a.ate, a.aplicar, a.cache)


if __name__ == "__main__":
    main()
