# -*- coding: utf-8 -*-
"""Puxa do eOrder o prazo das ordens que o robô nunca viu abertas.

O prazo ("Data Prevista Finalização Trabalhos") só existe no TdC -- o export
da Execução não traz (ele tem "Respeito ANS Legal/Interno/Parceira", que
seriam a resposta pronta, mas conferido em 30/10/2026 vêm 'false' em 100% das
linhas, inclusive nas executadas dentro do prazo: campo não alimentado nesse
contrato). Por isso o robô guarda o prazo de cada ordem enquanto
ela ainda aparece no TdC (prazos_os_idx.json) e carimba esse prazo nas linhas
da Execução na hora de publicar.

Ordem que foi aberta e baixada entre duas leituras do robô nunca passou pelo
TdC: fica sem prazo nenhum, e a aba "Ordens no Prazo" não consegue dizer se
ela estourou. Medido em 30/10/2026: 1.336 ordens assim, quase todas de antes
de 05/10 (foi quando o robô passou a anotar o prazo a cada rodada em vez de
só no fechamento do dia).

A busca de TdC do robô DESMARCA Finalizado/Anulado/Encerrado/Suspenso, porque
o painel só quer o que está aberto. Este script faz a MESMA busca com todos os
estados marcados -- as ordens que faltam são justamente as finalizadas -- e
funde o resultado no índice.

FATIADO POR SEMANA (importante): pedir o período todo de uma vez não funciona.
A primeira tentativa (01/08 a 30/10, 5 estados) morreu no "Exportar em xls":
a grade ainda estava carregando. Medido depois, 1 semana = 3 MB de export e
~6 min de ponta a ponta -- quase o mesmo que a rodada normal de TdC do robô
(1,9 MB / 5 min). Então o caminho é fatiar: um login só, e dentro dele uma
busca por semana. Cada fatia pesa como uma rodada normal, e o flock garante
que nunca disputa o navegador com a Busca Execução do cron.

REGRA DA FUSÃO (importante): só grava quando a ordem não tem prazo nenhum, ou
quando o prazo do export é MAIS NOVO que o guardado. Prazo de ordem não anda
pra trás: reabertura e reprogramação empurram pra frente. Um prazo mais VELHO
vindo do export é de um ciclo anterior da mesma ordem -- gravar ele faria
ordem executada no prazo aparecer como atrasada. Conferido em out/2026: sem
essa regra, 82 ligações novas viravam "fora do prazo" sem estar.

Depois de rodar aqui, rode recarimbar_prazos_historico.py pra reescrever o
prazo nas linhas da Execução que já estão no Supabase.

    python baixar_prazos_finalizados.py                       # 60 dias, de 7 em 7
    python baixar_prazos_finalizados.py --de 01/09/2026 --ate 30/10/2026
    python baixar_prazos_finalizados.py --passo 14            # fatias de 2 semanas
    python baixar_prazos_finalizados.py --so-ler A.xlsx B.xlsx  # sem abrir o eOrder
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta

from selenium.webdriver.support import expected_conditions as EC

from eorder_execucao_bot import (
    EOrderExecucaoBot, PRAZOS_OS_PATH, PRAZO_COL, NOME_EXPORT_TDC,
    XP_TRES_PONTOS_ESTADO, XP_ESTADO_FINALIZADO, XP_ESTADO_ANULADO,
    XP_ESTADO_ENCERRADO, XP_ESTADO_SUSPENSO, XP_ESTADO_ABERTO,
    XP_DATAS_REF_SPAN, XP_DATA_LANC_INI, XP_DATA_LANC_FIM,
    XP_LISTA_EXPORT_TDC, XP_TRES_PONTOS_LISTA_TDC,
    _parse_data_fim, _carregar_json_idx, _salvar_json_idx,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CRED_FILE = os.path.join(BASE_DIR, "credenciais_eorder.json")
COL_NUM = "Numero de Serviço"
# Os 5 estados da busca. O fluxo normal deixa só os "abertos"; aqui todos.
ESTADOS = [
    ("Finalizado", XP_ESTADO_FINALIZADO),
    ("Anulado", XP_ESTADO_ANULADO),
    ("Encerrado", XP_ESTADO_ENCERRADO),
    ("Suspenso", XP_ESTADO_SUSPENSO),
    ("Aberto", XP_ESTADO_ABERTO),
]


def log(msg):
    linha = f"[{datetime.now():%H:%M:%S}] {msg}"
    try:
        print(linha, flush=True)
    except UnicodeEncodeError:
        # console do Windows em cp1252 engasga com emoji; na VPS (UTF-8) nao cai aqui
        print(linha.encode("ascii", "replace").decode("ascii"), flush=True)


def carregar_credenciais():
    if not os.path.exists(CRED_FILE):
        sys.exit(f"❌ Credenciais não encontradas: {CRED_FILE}")
    with open(CRED_FILE, encoding="utf-8") as f:
        return json.load(f)


def fatiar(de, ate, passo):
    """Quebra o período em pedaços de `passo` dias (dd/mm/aaaa nas duas pontas)."""
    d = datetime.strptime(de, "%d/%m/%Y")
    f = datetime.strptime(ate, "%d/%m/%Y")
    fatias = []
    while d <= f:
        fim = min(d + timedelta(days=passo - 1), f)
        fatias.append((d.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")))
        d = fim + timedelta(days=1)
    return fatias


# ── eOrder ───────────────────────────────────────────────────────────
def _marcado(bot, xp):
    el = bot._find(xp, EC.presence_of_element_located, timeout=15)
    return "tvRowSelected" in (el.get_attribute("class") or "")


def _marcar_todos_os_estados(bot):
    """Deixa os 5 estados marcados. Clica só no que está desmarcado (o clique
    é toggle) e confere no fim -- um clique que não pega inverte a busca
    inteira em silêncio, que é o mesmo cuidado que o _estado_tdc() original
    tem no sentido contrário. Como o eOrder lembra o filtro entre sessões, da
    segunda fatia em diante isso normalmente só confere e segue."""
    for tentativa in range(1, 4):
        log(f"⚙️  Abrindo Estado de TdC (marcando todos) [tentativa {tentativa}/3]...")
        bot._click(XP_TRES_PONTOS_ESTADO, timeout=20)
        for nome, xp in ESTADOS:
            if _marcado(bot, xp):
                log(f"   ✔ '{nome}' já marcado")
            else:
                log(f"   ➕ Marcando '{nome}'...")
                bot._click(xp, timeout=20)
        if all(_marcado(bot, xp) for _, xp in ESTADOS):
            log("✅ Todos os estados marcados.")
            return
        log("⚠️  Algum estado não ficou marcado -- tentando de novo...")
    raise RuntimeError("Não consegui deixar os 5 estados marcados após 3 tentativas")


def _datas(bot, de, ate):
    log("📅 Abrindo Datas Referência...")
    bot._click(XP_DATAS_REF_SPAN, timeout=20)
    log(f"📅 Data Lançamento: {de} até {ate}...")
    bot._type(XP_DATA_LANC_INI, de, timeout=20)
    bot._type(XP_DATA_LANC_FIM, ate, timeout=20)


def _esperar_e_exportar(bot, minutos=20):
    """Abre o menu de exportação, com paciência.

    A busca normal do robô (só o que está aberto, 60 dias) volta em segundos,
    então o fluxo original clica em Buscar e exporta na sequência. Com os 5
    estados marcados a grade volta muito maior e leva minutos -- e enquanto
    ela está carregando o menu dos 3 pontinhos não abre. Sem esta espera o
    primeiro run (01/08 a 30/10) morreu em 'Item Exportar em xls não
    encontrado no menu' 22 segundos depois do Buscar.
    """
    limite = time.time() + minutos * 60
    tentativa = 0
    while True:
        tentativa += 1
        try:
            bot._exportar_tdc()
            return
        except Exception as e:
            if time.time() >= limite:
                raise
            log(f"⏳ Grade ainda carregando (tentativa {tentativa}) -- espero 30s. [{e}]")
            time.sleep(30)


def baixar_fatias(fatias, download_dir, usuario, senha):
    """Um login só, uma busca+export por fatia. Devolve os arquivos baixados.

    Uma fatia que falhar não derruba as outras: o que já baixou é fundido no
    fim, e o período que faltou pode ser pedido de novo com --de/--ate.
    """
    from eorder_execucao_bot import XP_BTN_BUSCAR_TDC
    bot = EOrderExecucaoBot(lambda m: log(m), download_dir, minimizado=False)
    bot._start_driver()
    caminhos = []
    try:
        bot._login(usuario, senha)
        for i, (de, ate) in enumerate(fatias, 1):
            log(f"===== Fatia {i}/{len(fatias)}: {de} a {ate} =====")
            try:
                bot._navegar_busca_tdcs()
                bot._marcar_centro_operativo_tdc()
                bot._processos_subprocessos()
                _marcar_todos_os_estados(bot)
                _datas(bot, de, ate)
                log("🔎 Buscando...")
                bot._click(XP_BTN_BUSCAR_TDC, timeout=20)
                _esperar_e_exportar(bot)
                # Base com finalizadas é maior que a do fluxo normal; 30 min.
                ok = bot._baixar_exportacao(XP_LISTA_EXPORT_TDC, XP_TRES_PONTOS_LISTA_TDC,
                                            NOME_EXPORT_TDC, espera_max=1800)
                if not ok:
                    raise RuntimeError("Export não ficou pronto")
                caminho = bot._achar_export_mais_recente(
                    NOME_EXPORT_TDC, timestamp_esperado=ok if isinstance(ok, str) else None)
                if not caminho:
                    raise RuntimeError("Não achei o arquivo baixado")
                log(f"💾 Export: {caminho}")
                caminhos.append(caminho)
            except Exception as e:
                log(f"⚠️  Fatia {de} a {ate} FALHOU: {e}")
    finally:
        try:
            if bot.driver:
                bot.driver.quit()
        except Exception:
            pass
    return caminhos


# ── Fusão no índice ──────────────────────────────────────────────────
def _ler_export(caminho):
    """Cabeçalho + linhas do xlsx do eOrder.

    SEM read_only: o arquivo que o eOrder gera não declara estilo padrão e,
    em modo read_only, o openpyxl devolve o cabeçalho truncado (uma coluna
    só, tipo ['CO']) -- foi o que derrubou a primeira versão disso aqui. O
    robô usa load_workbook(data_only=True) e é esse caminho que funciona.
    """
    import openpyxl
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active
    return [c.value for c in ws[1]], ws.iter_rows(min_row=2, values_only=True)


def fundir(caminhos, aplicar):
    idx = _carregar_json_idx(PRAZOS_OS_PATH) or {}
    antes = len(idx)
    novas = atualizadas = ignoradas = sem_prazo = lidas = 0
    for caminho in caminhos:
        cab, linhas = _ler_export(caminho)
        if COL_NUM not in cab or PRAZO_COL not in cab:
            log(f"⚠️  {os.path.basename(caminho)} sem as colunas esperadas, pulando.")
            log(f"   Tem: {cab}")
            continue
        i_num, i_pz = cab.index(COL_NUM), cab.index(PRAZO_COL)
        i_legal = cab.index("Prazo ANS Legal") if "Prazo ANS Legal" in cab else None
        for row in linhas:
            n = (row[i_num] or "")
            n = n.strip() if isinstance(n, str) else str(n)
            if not n:
                continue
            lidas += 1
            p = row[i_pz] or (row[i_legal] if i_legal is not None else None)
            p = p.strip() if isinstance(p, str) else (p.strftime("%d/%m/%Y %H:%M") if hasattr(p, "strftime") else None)
            if not p:
                sem_prazo += 1
                continue
            pnovo = _parse_data_fim(p)
            if not pnovo:
                sem_prazo += 1
                continue
            atual = idx.get(n)
            if atual is None:
                idx[n] = p
                novas += 1
            elif (_parse_data_fim(atual) or datetime.min) < pnovo:
                idx[n] = p
                atualizadas += 1
            else:
                ignoradas += 1

    log(f"📄 Linhas lidas: {lidas} (sem prazo no export: {sem_prazo})")
    log(f"➕ Prazos novos (ordem que o robô nunca viu): {novas}")
    log(f"🔄 Prazos trocados por um mais novo (reabertura/reprogramação): {atualizadas}")
    log(f"⏭  Ignorados por serem mais velhos que o guardado: {ignoradas}")
    log(f"📦 Índice: {antes} → {antes + novas} ordens")
    if not aplicar:
        log("🔍 MODO TESTE -- nada foi gravado. Rode com --aplicar pra valer.")
        return
    _salvar_json_idx(PRAZOS_OS_PATH, idx)
    log(f"✅ Gravado em {PRAZOS_OS_PATH}")


def main():
    hoje = datetime.now()
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--de", default=(hoje - timedelta(days=60)).strftime("%d/%m/%Y"),
                    help="Data Lançamento inicial (dd/mm/aaaa)")
    ap.add_argument("--ate", default=hoje.strftime("%d/%m/%Y"),
                    help="Data Lançamento final (dd/mm/aaaa)")
    ap.add_argument("--passo", type=int, default=7,
                    help="Tamanho de cada fatia em dias (padrão 7)")
    ap.add_argument("--so-ler", metavar="ARQUIVO", nargs="+",
                    help="Pula o eOrder e funde xlsx já baixados")
    ap.add_argument("--aplicar", action="store_true",
                    help="Grava no índice (sem isso é só simulação)")
    a = ap.parse_args()

    caminhos = a.so_ler
    if not caminhos:
        cred = carregar_credenciais()
        # acesso2 e o que o cron usa pro TdC -- mesma tela, mesmo login.
        acesso = cred.get("acesso2") or cred.get("acesso1") or cred
        download_dir = cred.get("download_dir") or os.path.join(os.path.expanduser("~"), "Downloads")
        fatias = fatiar(a.de, a.ate, a.passo)
        log(f"🗓  Período (Data Lançamento): {a.de} a {a.ate} "
            f"— {len(fatias)} fatia(s) de até {a.passo} dia(s)")
        caminhos = baixar_fatias(fatias, download_dir, acesso["usuario"], acesso["senha"])
        if not caminhos:
            sys.exit("❌ Nenhuma fatia baixou.")
        log(f"📥 {len(caminhos)} arquivo(s) baixado(s), fundindo...")
    fundir(caminhos, a.aplicar)


if __name__ == "__main__":
    main()
