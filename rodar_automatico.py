import json
import logging
import os
import sys
import threading
from datetime import datetime

from eorder_execucao_bot import EOrderExecucaoBot

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CRED_FILE = os.path.join(BASE_DIR, "credenciais_eorder.json")
LOG_FILE = os.path.join(BASE_DIR, "automatico.log")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
)


def log(msg):
    try:
        print(msg)
    except Exception:
        pass
    logging.info(msg)


def carregar_credenciais():
    if not os.path.exists(CRED_FILE):
        log(
            f"❌ Arquivo de credenciais não encontrado: {CRED_FILE} — "
            "copie credenciais_eorder.example.json para credenciais_eorder.json e preencha "
            "com usuário/senha reais do eOrder."
        )
        raise SystemExit(1)
    with open(CRED_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _fechar(bot):
    try:
        if bot.driver:
            bot.driver.quit()
    except Exception:
        pass


def main():
    cred = carregar_credenciais()
    download_dir = cred.get("download_dir") or os.path.join(os.path.expanduser("~"), "Downloads")
    modo = str(cred.get("modo", "1"))
    data_str = datetime.now().strftime("%d/%m/%Y")

    log(f"=== Rodada automática iniciada — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} — data busca: {data_str} — modo {modo} ===")

    if modo == "2":
        acesso1 = cred["acesso1"]
        acesso2 = cred["acesso2"]
        bot1 = EOrderExecucaoBot(lambda m: log(f"[Execução] {m}"), download_dir, minimizado=False)
        bot2 = EOrderExecucaoBot(lambda m: log(f"[TdC] {m}"), download_dir, minimizado=False)

        t1 = threading.Thread(target=bot1.executar, args=(acesso1["usuario"], acesso1["senha"], data_str))
        t2 = threading.Thread(target=bot2.executar_tdc, args=(acesso2["usuario"], acesso2["senha"]))
        t1.start()
        t2.start()
        t1.join()
        t2.join()
        _fechar(bot1)
        _fechar(bot2)
    else:
        acesso1 = cred["acesso1"]
        bot = EOrderExecucaoBot(lambda m: log(f"[Único] {m}"), download_dir, minimizado=False)
        try:
            bot._start_driver()
            bot._login(acesso1["usuario"], acesso1["senha"])
            bot.fazer_busca_execucao(data_str)
            bot.fazer_tdc()
        except Exception as e:
            log(f"❌ Erro: {e}")
        finally:
            _fechar(bot)

    log("=== Rodada automática concluída ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"❌ ERRO FATAL: {e}")
        sys.exit(1)
