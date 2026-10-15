# -*- coding: utf-8 -*-
"""Carga inicial do rastreio PLC a partir do histórico já publicado.

O índice do bot (plc_rastreio_idx.json) começa vazio e se constrói rodada a
rodada, olhando só o dia corrente -- então, sem esta carga, as vistorias com
PLC que já aconteceram ficariam de fora e a aba levaria meses pra ter conteúdo.

Este script lê o snapshots_historico da Execução no Supabase, dia a dia em
ordem cronológica, e alimenta o MESMO _atualizar_indice_plc() que o bot usa em
produção -- o resultado é idêntico ao que o bot teria produzido se já
estivesse rodando desde o começo do histórico.

Rodar UMA vez, na VPS:

    cd /opt/painel-eorder/painel-eorder-cosampa
    /opt/painel-eorder/venv/bin/python3 carregar_historico_plc.py

É seguro rodar de novo: _atualizar_indice_plc é idempotente (compara a
'Data fim Execução' exata), então reprocessar não duplica evento nenhum.

Opções:
    --desde AAAA-MM   primeiro mês a ler (padrão: 2026-01)
    --publicar        publica o PLC_rastreio no fim, sem esperar a próxima
                      rodada do bot
"""
import argparse
import calendar
import json
import sys
import urllib.error
import urllib.request
from datetime import date, datetime, timezone

import eorder_execucao_bot as bot


def _buscar_mes(ano, mes):
    """Um mês por vez: o período inteiro de uma vez estoura o statement
    timeout do Supabase (a coluna 'dados' é jsonb e pesa)."""
    ultimo = calendar.monthrange(ano, mes)[1]
    url = (bot.SB_URL + "/rest/v1/snapshots_historico"
           "?regiao=eq.Execucao"
           "&data=gte.%04d-%02d-01&data=lte.%04d-%02d-%02d"
           "&select=dados,data" % (ano, mes, ano, mes, ultimo))
    req = urllib.request.Request(url, headers={
        "apikey": bot.SB_KEY,
        "Authorization": "Bearer " + bot.SB_KEY,
        "Accept-Encoding": "gzip",
    })
    with urllib.request.urlopen(req, timeout=180) as resp:
        bruto = resp.read()
        if resp.headers.get("Content-Encoding") == "gzip":
            import gzip
            bruto = gzip.decompress(bruto)
    return json.loads(bruto.decode("utf-8"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", default="2026-01", help="primeiro mês (AAAA-MM)")
    ap.add_argument("--publicar", action="store_true",
                    help="publica o PLC_rastreio ao final")
    args = ap.parse_args()

    ano0, mes0 = (int(x) for x in args.desde.split("-"))
    hoje = date.today()

    print("Carga inicial do rastreio PLC")
    print("  índice: %s" % bot.PLC_IDX_PATH)
    print("  período: %s até %s" % (args.desde, hoje.strftime("%Y-%m")))
    print()

    # Junta tudo por dia antes de aplicar: a ordem cronológica importa, porque
    # o índice só aceita eventos POSTERIORES à baixa do PLC.
    por_dia = {}
    ano, mes = ano0, mes0
    while (ano, mes) <= (hoje.year, hoje.month):
        try:
            linhas = _buscar_mes(ano, mes)
            n = sum(len(l.get("dados") or []) for l in linhas)
            for l in linhas:
                por_dia.setdefault(l["data"], []).extend(l.get("dados") or [])
            print("  %04d-%02d: %d dia(s), %d registro(s)" % (ano, mes, len(linhas), n))
        except urllib.error.HTTPError as e:
            print("  %04d-%02d: ERRO %s %s" % (ano, mes, e.code,
                                               e.read().decode("utf-8", "replace")[:120]))
        except Exception as e:
            print("  %04d-%02d: ERRO %s" % (ano, mes, e))
        mes += 1
        if mes > 12:
            ano, mes = ano + 1, 1

    if not por_dia:
        print("\nNada lido -- nada a fazer.")
        return 1

    print()
    print("Aplicando %d dia(s) em ordem cronológica..." % len(por_dia))
    indice = {}
    for dia in sorted(por_dia):
        indice = bot._atualizar_indice_plc(por_dia[dia])

    abertas = [v for v in indice.values() if not v.get("executada_em")]
    execs = [v for v in indice.values() if v.get("executada_em")]
    mesma = sum(1 for v in execs if v.get("eq_exec") == v.get("eq_plc"))

    print()
    print("Resultado:")
    print("  ordens com vistoria PLC .... %d" % len(indice))
    print("  em aberto .................. %d" % len(abertas))
    print("  executadas ................. %d" % len(execs))
    print("     na mesma equipe ......... %d" % mesma)
    print("     por outra equipe ........ %d" % (len(execs) - mesma))

    if args.publicar:
        print()
        print("Publicando PLC_rastreio...")
        # Mesma função que o bot usa -- a lista de campos mora num lugar só.
        dados = bot.montar_dados_plc(indice)
        corpo = json.dumps({
            "regiao": "PLC_rastreio",
            "dados": dados,
            "atualizado_em": datetime.now(timezone.utc).isoformat(),
        }, default=str).encode("utf-8")
        req = urllib.request.Request(
            bot.SB_URL + "/rest/v1/snapshots", data=corpo, method="POST",
            headers={"apikey": bot.SB_KEY,
                     "Authorization": "Bearer " + bot.SB_KEY,
                     "Content-Type": "application/json",
                     "Prefer": "resolution=merge-duplicates"})
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                print("  publicado -- status %s, %d KB" % (resp.status, len(corpo) // 1024))
        except urllib.error.HTTPError as e:
            print("  ERRO %s %s" % (e.code, e.read().decode("utf-8", "replace")[:200]))
            return 1
    else:
        print()
        print("Índice gravado. O bot publica na próxima rodada.")
        print("(use --publicar pra publicar agora)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
