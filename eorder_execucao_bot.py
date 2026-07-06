import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
import os
import re
import glob
import json
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException

import openpyxl

EORDER_URL = "https://eorder-ceara.enel.com/geocallcoe/w/index.htm"
NOME_EXPORT = "EXECUCAO"

# ── Publicação automática no painel (Supabase) ────────────────────────
SB_URL = "https://xnkvpxireoosrnrfwcws.supabase.co"
SB_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6"
          "Inhua3ZweGlyZW9vc3JucmZ3Y3dzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODMxMDA3"
          "NzIsImV4cCI6MjA5ODY3Njc3Mn0.BaCa1dUZAEHhwcqx9Es-U1oXrICk08J14e4mUkieH9g")

# Só publica as colunas que o painel web realmente usa — export completo
# tem ~40 colunas e o payload inteiro estoura o timeout de escrita do Supabase.
COLS_TDC = [
    "Numero de Serviço", "Tipo Remessa WIN", " Tipo de Serviço", "Estado TdC",
    "Código Equipe", "Chefe/Responsável de Equipe", "Município", "Endereço Completo",
    "Data Prevista Finalização Trabalhos", "Latitude", "Longitude", "Dica Localização",
    "Rota de Leitura", "Código Cliente", "Nome e Sobrenome Cliente",
]
COLS_EXECUCAO = [
    "Numero de Serviço", "Tipo Remessa WIN", " Tipo de Serviço",
    "Recurso/Equipe", "Município", "Data fim Execução", "Código Resultado",
]

# ── XPaths ───────────────────────────────────────────────────────────
XP_USER          = "/html/body/table/tbody/tr/td/div/div[2]/div/div/form/div/div[2]/table/tbody/tr[1]/td[2]/input"
XP_PASS          = "/html/body/table/tbody/tr/td/div/div[2]/div/div/form/div/div[2]/table/tbody/tr[2]/td[2]/input"

XP_PLANEJAMENTO  = '//*[@id="TBB_tbm2"]/div[6]'
XP_BUSCA_EXEC    = '//*[@id="TBB_tbm2"]/div[2]'
XP_CENTRO_OP     = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[1]/tbody[1]/tr[2]/td[1]/table/tbody/tr[2]/td[2]/select'
CENTRO_OP_VALOR  = "Cosampa - Sul"
XP_TRES_PONTOS   = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[1]/tbody[1]/tr[2]/td[1]/table/tbody/tr[4]/td[2]/table/tbody/tr/td/div/div[1]/div/div[1]'
XP_CHK_EMERG     = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[1]/tbody[1]/tr[2]/td[1]/table/tbody/tr[4]/td[2]/table/tbody/tr/td/div/div[2]/div/div/div/table/tbody/tr[2]/td[2]'
XP_DATA_EXEC_ROW = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[1]/tbody[1]/tr[2]/td[1]/table/tbody/tr[12]/td[1]'
XP_DATA_INI      = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[1]/tbody[1]/tr[2]/td[1]/table/tbody/tr[12]/td[2]/table/tbody/tr/td[1]/table//input'
XP_DATA_FIM      = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[1]/tbody[1]/tr[2]/td[1]/table/tbody/tr[12]/td[2]/table/tbody/tr/td[3]/table/tbody/tr/td[1]//input'
XP_BTN_BUSCAR    = '/html/body/div[2]/div/div[2]/div/div[2]/div/div[2]/div/div/div[1]/table/tbody/tr/td/div/div[2]/div/form/table[3]/tbody/tr/td/table/tbody/tr/td[3]/button'

XP_BTN_EXPORTAR    = '/html/body/div[2]/div/div[2]/div/div[3]/div[1]/div/div[2]'
XP_CAMPO_NOME_EXP  = '//input[contains(@placeholder, "default")]'
XP_BTN_OK_EXP      = '/html/body/div[2]/div/div[2]/div/div[5]/div/div[2]/div/div/div/div/div[2]/div/form/div[2]/table/tbody/tr/td[1]/button'
XP_FECHAR_MSG_EXP  = '/html/body/div[3]/div[2]/div[1]/div/div[2]'
XP_LISTA_EXPORT    = '/html/body/div[1]/div[2]/table/tbody/tr[3]/td/div/div/table/tbody/tr[1]/td/div[3]/div[2]/div'
XP_TRES_PONTOS_LISTA = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div/form/div/div[1]/div/div[1]'

# ── XPaths — Exportação TdC ─────────────────────────────────────────
NOME_EXPORT_TDC = "CosampaCDU"

XP_LISTA_TDC      = "/html/body/div[1]/div[2]/table/tbody/tr[3]/td/div/div/table/tbody/tr[1]/td/div/div[2]/div[5]"
XP_BUSCA_TDC_MENU = "/html/body/div[1]/div[2]/table/tbody/tr[3]/td/div/div/table/tbody/tr[1]/td/div[2]/div[2]/div[1]"

XP_CENTRO_OP_TDC  = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/select'

XP_TRES_PONTOS_PROC = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/div[1]/div/div[1]'
XP_CHK_EMERG_TDC    = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[2]'

XP_TRES_PONTOS_ESTADO = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/div[1]/div/div[1]'
XP_ESTADO_FINALIZADO = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[2]'
XP_ESTADO_ANULADO    = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[3]'
XP_ESTADO_ENCERRADO  = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[4]'
XP_ESTADO_SUSPENSO   = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[3]/tr[6]/td[1]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/table/tbody/tr/td[1]/div/div[2]/div/div/div/div[2]/div/table/tbody/tr[6]'

XP_DATAS_REF_SPAN = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[9]/tr[1]/td/span'
XP_DATA_LANC_INI  = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[9]/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/fieldset/table/tbody/tr[1]/td/table/tbody/tr/td[1]//input'
XP_DATA_LANC_FIM  = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/table/tbody[9]/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr/td/fieldset/table/tbody/tr[3]/td/table/tbody/tr/td[1]//input'

XP_BTN_BUSCAR_TDC = '/html/body/div[2]/div/div[2]/div/div[1]/div[2]/div[1]/div/form/div[2]/table/tbody/tr/td[2]/button'

XP_TRES_PONTOS_EXPORT_TDC = '/html/body/div[2]/div/div[2]/div/div[3]/div[1]/div/div[1]'
XP_BTN_OK_EXP_TDC         = '/html/body/div[2]/div/div[2]/div/div[16]/div/div[2]/div/div/div/div/div[2]/div/form/div[2]/table/tbody/tr/td[1]/button'
XP_FECHAR_MSG_TDC         = '/html/body/div[7]/div[2]/div[1]/div/div[2]'
XP_LISTA_EXPORT_TDC       = '/html/body/div[1]/div[2]/table/tbody/tr[3]/td/div/div/table/tbody/tr[1]/td/div[2]/div[2]/div[4]'
XP_TRES_PONTOS_LISTA_TDC  = '/html/body/div[2]/div/div[2]/div/div[1]/div/div[1]'


def _xp_linha_arquivo(nome_export):
    return f'//td[starts-with(normalize-space(text()), "{nome_export}_")]'


class EOrderExecucaoBot:
    def __init__(self, log_cb, download_dir, minimizado=False):
        self.log = log_cb
        self.download_dir = download_dir
        self.minimizado = minimizado
        self.driver = None
        self.stop_flag = False
        self._frame_cache = {}

    def _plog(self, msg):
        self.log(msg)

    def _start_driver(self):
        opts = webdriver.ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-infobars")
        opts.page_load_strategy = "eager"
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.add_experimental_option("prefs", {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.password_manager_leak_detection": False,
        })
        opts.add_argument("--disable-save-password-bubble")
        self.driver = webdriver.Chrome(options=opts)
        if self.minimizado:
            self.driver.minimize_window()

    def _find(self, xpath, condition=EC.visibility_of_element_located, timeout=20):
        driver = self.driver
        wait_fast = WebDriverWait(driver, min(timeout, 4), poll_frequency=0.2)

        cached = self._frame_cache.get(xpath)
        if cached is not None:
            try:
                driver.switch_to.default_content()
                if cached != "default":
                    driver.switch_to.frame(cached)
                return wait_fast.until(condition((By.XPATH, xpath)))
            except Exception:
                pass

        driver.switch_to.default_content()
        try:
            el = wait_fast.until(condition((By.XPATH, xpath)))
            self._frame_cache[xpath] = "default"
            return el
        except Exception:
            pass

        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for idx, frame in enumerate(frames):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                el = wait_fast.until(condition((By.XPATH, xpath)))
                self._frame_cache[xpath] = idx
                return el
            except Exception:
                pass

        driver.switch_to.default_content()
        if cached is not None:
            try:
                if cached != "default":
                    driver.switch_to.frame(cached)
                return WebDriverWait(driver, timeout, poll_frequency=0.2).until(
                    condition((By.XPATH, xpath)))
            except Exception:
                pass

        driver.switch_to.default_content()
        raise TimeoutException(f"Elemento não encontrado: {xpath}")

    def _click(self, xpath, timeout=20):
        el = self._find(xpath, EC.element_to_be_clickable, timeout)
        try:
            el.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", el)
        return el

    def _double_click(self, xpath, timeout=20):
        el = self._find(xpath, EC.element_to_be_clickable, timeout)
        try:
            webdriver.ActionChains(self.driver).double_click(el).perform()
        except Exception:
            self.driver.execute_script(
                "var ev = new MouseEvent('dblclick', {bubbles: true}); arguments[0].dispatchEvent(ev);", el)
        return el

    def _click_por_texto(self, texto):
        """
        Clica no elemento mais específico (mais aninhado) cujo texto exato é
        `texto`, usando JS — evita depender da estrutura exata do DOM de um
        item de menu (ex.: "Atualizar" dentro do menu de 3 pontinhos).
        Retorna True se encontrou e clicou, False caso contrário.
        """
        script = """
            var texto = arguments[0];
            var exatos = document.evaluate(
                "//*[normalize-space(text())=" + JSON.stringify(texto) + "]",
                document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
            if (exatos.snapshotLength > 0) {
                exatos.snapshotItem(exatos.snapshotLength - 1).click();
                return true;
            }
            var parciais = document.evaluate(
                "//*[contains(normalize-space(text())," + JSON.stringify(texto) + ")]",
                document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
            if (parciais.snapshotLength === 0) return false;
            var melhor = parciais.snapshotItem(0);
            for (var i = 1; i < parciais.snapshotLength; i++) {
                var el = parciais.snapshotItem(i);
                if (el.textContent.length < melhor.textContent.length) melhor = el;
            }
            melhor.click();
            return true;
        """
        return bool(self.driver.execute_script(script, texto))

    def _type(self, xpath, texto, timeout=20):
        el = self._find(xpath, EC.element_to_be_clickable, timeout)
        el.click(); el.clear(); el.send_keys(str(texto))
        return el

    def _login(self, usuario, senha):
        self._plog("🌐 Abrindo eOrder...")
        self.driver.get(EORDER_URL)
        self._plog("🔑 Fazendo login...")
        campo_user = self._find(XP_USER, EC.visibility_of_element_located, timeout=20)
        campo_user.click(); campo_user.clear(); campo_user.send_keys(usuario)
        campo_pwd = self.driver.find_element(By.XPATH, XP_PASS)
        campo_pwd.click(); campo_pwd.clear(); campo_pwd.send_keys(senha)
        campo_pwd.send_keys(Keys.RETURN)
        self.driver.switch_to.default_content()
        time.sleep(3)
        self._plog("✅ Login efetuado.")

    def _navegar_busca_execucao(self):
        self._plog("📂 Abrindo Planejamento...")
        self._click(XP_PLANEJAMENTO, timeout=20)
        self._plog("📂 Abrindo Busca Execução...")
        self._click(XP_BUSCA_EXEC, timeout=20)

    def _marcar_centro_operativo(self):
        self._plog(f"📍 Marcando Centro Operativo: {CENTRO_OP_VALOR}...")
        el = self._find(XP_CENTRO_OP, EC.presence_of_element_located, timeout=20)
        Select(el).select_by_visible_text(CENTRO_OP_VALOR)

    def _desmarcar_emergencia(self):
        self._plog("⚙️  Abrindo opções de filtro...")
        self._click(XP_TRES_PONTOS, timeout=20)
        self._plog("🚫 Desmarcando 'Atendimento de Emergência'...")
        self._click(XP_CHK_EMERG, timeout=20)

    def _preencher_data(self, data_str):
        self._plog(f"📅 Preenchendo data de execução: {data_str}...")
        self._click(XP_DATA_EXEC_ROW, timeout=20)
        self._type(XP_DATA_INI, data_str, timeout=20)
        self._type(XP_DATA_FIM, data_str, timeout=20)

    def _buscar(self):
        self._plog("🔎 Buscando...")
        self._click(XP_BTN_BUSCAR, timeout=20)

    def _exportar(self):
        self._plog("📤 Exportando resultados...")
        self._click(XP_BTN_EXPORTAR, timeout=20)
        self._plog("   ...clicou no botão de exportar")
        self._exportar_generico(NOME_EXPORT, XP_BTN_OK_EXP, XP_FECHAR_MSG_EXP)

    def _exportar_generico(self, nome_export, xp_btn_ok, xp_fechar_msg):
        """
        Corpo comum de qualquer exportação: digita o nome do arquivo no
        popup "Nome Arquivo", confirma (Enter, com fallback no botão Ok)
        e fecha o popup de confirmação. Quem abre o popup é responsabilidade
        de cada fluxo (botão direto na execução, menu de 3 pontinhos no TdC).
        """
        campo = self._type(XP_CAMPO_NOME_EXP, nome_export, timeout=20)
        self._plog("   ...digitou o nome do arquivo")
        try:
            campo.send_keys(Keys.RETURN)
            self._plog("   ...confirmou com Enter")
        except Exception:
            self._click(xp_btn_ok, timeout=10)
            self._plog("   ...clicou em OK")
        try:
            self._click(xp_fechar_msg, timeout=15)
            self._plog("   ...fechou popup de confirmação")
        except TimeoutException:
            self._plog("⚠️  Popup de confirmação não encontrado, seguindo...")

    @staticmethod
    def _timestamp_de(el):
        texto = (el.text or "").strip()
        m = re.search(r'(\d{8}_\d{6})', texto)
        return m.group(1) if m else ""

    def _tamanho_kb(self, elemento):
        """
        Lê a coluna "Dimensão (KB)" da linha do elemento — o arquivo aparece
        na lista com 0 KB enquanto o servidor ainda está gerando o conteúdo;
        só deve ser baixado depois que esse valor for > 0.
        """
        script = """
            var tr = arguments[0].closest('tr');
            if (!tr) return null;
            var tds = tr.querySelectorAll('td');
            for (var i = 0; i < tds.length; i++) {
                var t = tds[i].textContent.trim();
                if (/^[0-9]+(\\.[0-9]+)?$/.test(t)) { return parseFloat(t); }
            }
            return null;
        """
        return self.driver.execute_script(script, elemento)

    def _elemento_mais_recente(self, xpath, timestamp_limite):
        """
        Entre todos os elementos que casam com xpath (vários arquivos
        EXECUCAO_AAAAMMDD_HHMMSS na lista — outros usuários também exportam
        com esse nome), retorna o de timestamp mais recente, desde que seja
        ESTRITAMENTE maior que `timestamp_limite`.

        `timestamp_limite` é o maior timestamp já presente na lista no
        momento em que abrimos "Lista Exportações" (logo após clicar em
        Exportar) — usamos o relógio do PRÓPRIO SERVIDOR (embutido no nome
        do arquivo) como referência, e não o relógio do PC local, pra não
        sofrer com diferença de horário entre as duas máquinas.
        """
        try:
            self._find(xpath, EC.presence_of_element_located, timeout=3)  # garante frame certo no cache
        except TimeoutException:
            return None  # lista ainda vazia — nenhum arquivo desse nome existe ainda
        elementos = self.driver.find_elements(By.XPATH, xpath)
        novos = [el for el in elementos if self._timestamp_de(el) > timestamp_limite]
        if not novos:
            return None
        return max(novos, key=self._timestamp_de)

    def _baixar_exportacao(self, xp_lista_export, xp_tres_pontos_lista, nome_export,
                            espera_max=600, intervalo=5):
        xp_linha = _xp_linha_arquivo(nome_export)

        self._plog("📋 Abrindo Lista de Exportações...")
        self._click(xp_lista_export, timeout=20)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.invisibility_of_element_located((By.ID, "darkdiv")))
        except TimeoutException:
            pass
        time.sleep(1)

        existentes = self.driver.find_elements(By.XPATH, xp_linha)
        timestamp_limite = max((self._timestamp_de(el) for el in existentes), default="")
        self._plog(f"   (ignorando arquivos com timestamp <= {timestamp_limite or '(nenhum existente)'})")

        self._plog("⏳ Aguardando arquivo aparecer na lista (clicando em Atualizar)...")
        decorrido = 0
        elemento = None
        while decorrido < espera_max:
            if self.stop_flag:
                self._plog("🛑 Parado.")
                return False
            elemento = self._elemento_mais_recente(xp_linha, timestamp_limite)
            if elemento is not None:
                tamanho = self._tamanho_kb(elemento)
                if tamanho and tamanho > 0:
                    self._plog(f"✅ Arquivo encontrado na lista! ({tamanho:.0f} KB)")
                    break
                self._plog("   ...arquivo apareceu mas ainda está sendo gerado (0 KB), aguardando...")
                elemento = None
            try:
                self._click(xp_tres_pontos_lista, timeout=5)
                time.sleep(1.5)
                if not self._click_por_texto("Atualizar"):
                    # Site pode estar lento pra renderizar o menu — tenta mais
                    # uma vez antes de desistir dessa rodada
                    time.sleep(1.0)
                    if not self._click_por_texto("Atualizar"):
                        self._plog("   ⚠️  Item 'Atualizar' não encontrado no menu.")
            except TimeoutException:
                pass
            time.sleep(intervalo)
            decorrido += intervalo
            # Se a exportação foi rápida demais, o arquivo já podia estar na
            # lista no instante em que tiramos a "foto" do que já existia
            # (timestamp_limite) — nesse caso nunca vamos achar algo
            # ESTRITAMENTE mais novo. Se já tentamos algumas vezes e o
            # "limite" é muito recente (poucos minutos), assumimos que é o
            # nosso arquivo mesmo.
            if elemento is None and decorrido >= 15 and timestamp_limite:
                try:
                    ts = datetime.strptime(timestamp_limite, "%Y%m%d_%H%M%S")
                    if abs((datetime.now() - ts).total_seconds()) <= 300:
                        candidatos = self.driver.find_elements(By.XPATH, xp_linha)
                        candidato = next(
                            (el for el in candidatos if self._timestamp_de(el) == timestamp_limite), None)
                        if candidato is not None:
                            tamanho = self._tamanho_kb(candidato)
                            if tamanho and tamanho > 0:
                                elemento = candidato
                                self._plog("   (nenhum arquivo mais novo apareceu — assumindo que o mais recente já é o nosso)")
                                self._plog(f"✅ Arquivo encontrado na lista! ({tamanho:.0f} KB)")
                                break
                except ValueError:
                    pass
        else:
            self._plog("❌ Tempo esgotado esperando o arquivo de exportação.")
            return False

        self._plog(f"⬇️  Baixando arquivo mais recente: {elemento.text.strip()}...")
        try:
            webdriver.ActionChains(self.driver).double_click(elemento).perform()
        except Exception:
            self.driver.execute_script(
                "var ev = new MouseEvent('dblclick', {bubbles: true}); arguments[0].dispatchEvent(ev);", elemento)
        return True

    # ── Publicação automática no painel (Supabase) ───────────────────
    def _achar_export_mais_recente(self, prefixo, espera_max=60, intervalo=2):
        """
        Espera o arquivo baixado aparecer na pasta de downloads (o navegador
        pode levar alguns segundos pra terminar de gravar em disco) e retorna
        o caminho do mais recente que casa com o prefixo.
        """
        decorrido = 0
        while decorrido < espera_max:
            candidatos = glob.glob(os.path.join(self.download_dir, f"{prefixo}_*.xlsx"))
            candidatos = [c for c in candidatos if not c.endswith(".crdownload")]
            if candidatos:
                return max(candidatos, key=os.path.getmtime)
            time.sleep(intervalo)
            decorrido += intervalo
        return None

    def _xlsx_para_linhas(self, caminho, colunas):
        wb = openpyxl.load_workbook(caminho, data_only=True)
        ws = wb.active
        cabecalho = [c.value for c in ws[1]]
        linhas = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            registro = {}
            for h, v in zip(cabecalho, row):
                if h is None or h not in colunas:
                    continue
                registro[h] = v
            linhas.append(registro)
        return linhas

    def _publicar_nuvem(self, regiao, prefixo_arquivo, colunas):
        try:
            caminho = self._achar_export_mais_recente(prefixo_arquivo)
            if not caminho:
                self._plog(f"⚠️  Não encontrei o arquivo {prefixo_arquivo}*.xlsx pra publicar.")
                return
            linhas = self._xlsx_para_linhas(caminho, colunas)
            corpo = json.dumps({
                "regiao": regiao,
                "dados": linhas,
                "atualizado_em": datetime.now(timezone.utc).isoformat(),
            }, default=str).encode("utf-8")
            req = urllib.request.Request(
                SB_URL + "/rest/v1/snapshots",
                data=corpo,
                method="POST",
                headers={
                    "apikey": SB_KEY,
                    "Authorization": "Bearer " + SB_KEY,
                    "Content-Type": "application/json",
                    "Prefer": "resolution=merge-duplicates",
                },
            )
            with urllib.request.urlopen(req) as resp:
                self._plog(f"☁ Painel atualizado ({regiao}): {len(linhas)} registros — status {resp.status}")
        except urllib.error.HTTPError as e:
            self._plog(f"⚠️  Falha ao publicar no painel ({regiao}): {e.code} {e.read().decode('utf-8', 'replace')[:200]}")
        except Exception as e:
            self._plog(f"⚠️  Falha ao publicar no painel ({regiao}): {e}")

    def executar(self, usuario, senha, data_str):
        try:
            self._start_driver()
            self._login(usuario, senha)
            self.fazer_busca_execucao(data_str)
        except Exception as e:
            self._plog(f"❌ Erro: {e}")
        finally:
            self._plog("🏁 Finalizado.")

    def fazer_busca_execucao(self, data_str):
        self._navegar_busca_execucao()
        self._marcar_centro_operativo()
        self._desmarcar_emergencia()
        self._preencher_data(data_str)
        self._buscar()
        self._exportar()
        ok = self._baixar_exportacao(XP_LISTA_EXPORT, XP_TRES_PONTOS_LISTA, NOME_EXPORT)
        if ok:
            self._plog("✅ Fluxo Busca Execução concluído. Aguardando download finalizar...")
            time.sleep(5)
            self._plog(f"💾 Verifique a pasta de downloads: {self.download_dir}")
            self._publicar_nuvem("Execucao", NOME_EXPORT, COLS_EXECUCAO)
        return ok

    # ── Fluxo TdC ────────────────────────────────────────────────────
    def _navegar_busca_tdcs(self):
        self.driver.switch_to.default_content()
        self._plog("📂 Abrindo Lista TdC...")
        self._click(XP_LISTA_TDC, timeout=20)
        self._plog("📂 Abrindo Busca TdCs...")
        self._click(XP_BUSCA_TDC_MENU, timeout=20)

    def _marcar_centro_operativo_tdc(self):
        self._plog(f"📍 Marcando Centro Operativo: {CENTRO_OP_VALOR}...")
        el = self._find(XP_CENTRO_OP_TDC, EC.presence_of_element_located, timeout=20)
        Select(el).select_by_visible_text(CENTRO_OP_VALOR)

    def _processos_subprocessos(self):
        self._plog("⚙️  Abrindo Processos/Subprocessos...")
        self._click(XP_TRES_PONTOS_PROC, timeout=20)
        self._plog("🚫 Desmarcando 'Atendimento de Emergência'...")
        self._click(XP_CHK_EMERG_TDC, timeout=20)

    def _estado_tdc(self):
        self._plog("⚙️  Abrindo Estado de TdC (marcando todos)...")
        self._click(XP_TRES_PONTOS_ESTADO, timeout=20)
        for nome, xp in [("Finalizado", XP_ESTADO_FINALIZADO),
                          ("Anulado", XP_ESTADO_ANULADO),
                          ("Encerrado", XP_ESTADO_ENCERRADO),
                          ("Suspenso", XP_ESTADO_SUSPENSO)]:
            self._plog(f"🚫 Desmarcando '{nome}'...")
            self._click(xp, timeout=20)

    def _datas_referencia_tdc(self):
        self._plog("📅 Abrindo Datas Referência...")
        self._click(XP_DATAS_REF_SPAN, timeout=20)
        hoje = datetime.now().strftime("%d/%m/%Y")
        ha_59_dias = (datetime.now() - timedelta(days=59)).strftime("%d/%m/%Y")
        self._plog(f"📅 Data Lançamento: {ha_59_dias} até {hoje}...")
        self._type(XP_DATA_LANC_INI, ha_59_dias, timeout=20)
        self._type(XP_DATA_LANC_FIM, hoje, timeout=20)

    def _buscar_tdc(self):
        self._plog("🔎 Buscando...")
        self._click(XP_BTN_BUSCAR_TDC, timeout=20)

    def _exportar_tdc(self):
        self._plog("📤 Exportando TdCs...")
        self._click(XP_TRES_PONTOS_EXPORT_TDC, timeout=20)
        time.sleep(0.8)
        if not self._click_por_texto("Exportar em xls"):
            self._plog("⚠️  Item 'Exportar em xls' não encontrado no menu.")
        self._plog("   ...abriu menu de exportação")
        self._exportar_generico(NOME_EXPORT_TDC, XP_BTN_OK_EXP_TDC, XP_FECHAR_MSG_TDC)

    def executar_tdc(self, usuario=None, senha=None):
        try:
            if usuario is not None:
                self._start_driver()
                self._login(usuario, senha)
            self.fazer_tdc()
        except Exception as e:
            self._plog(f"❌ Erro: {e}")
        finally:
            self._plog("🏁 Finalizado.")

    def fazer_tdc(self):
        self._navegar_busca_tdcs()
        self._marcar_centro_operativo_tdc()
        self._processos_subprocessos()
        self._estado_tdc()
        self._datas_referencia_tdc()
        self._buscar_tdc()
        self._exportar_tdc()
        ok = self._baixar_exportacao(XP_LISTA_EXPORT_TDC, XP_TRES_PONTOS_LISTA_TDC, NOME_EXPORT_TDC)
        if ok:
            self._plog("✅ Fluxo TdC concluído. Aguardando download finalizar...")
            time.sleep(5)
            self._plog(f"💾 Verifique a pasta de downloads: {self.download_dir}")
            self._publicar_nuvem("Sul", NOME_EXPORT_TDC, COLS_TDC)
        return ok


class App(tk.Tk):
    BG       = "#1e1e2e"
    FG       = "#cdd6f4"
    ENTRY_BG = "#313244"
    ACCENT   = "#89b4fa"
    GREEN    = "#a6e3a1"
    RED      = "#f38ba8"

    def __init__(self):
        super().__init__()
        self.title("🤖 eOrder Bot — Execução + TdC — Cosampa Sul")
        self.geometry("560x720")
        self.resizable(False, False)
        self.configure(bg=self.BG)
        self.bots = []
        self._ui()

    def _ui(self):
        BG, FG, ENTRY_BG, ACCENT, GREEN, RED = (
            self.BG, self.FG, self.ENTRY_BG, self.ACCENT, self.GREEN, self.RED)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("G.TButton", background=GREEN, foreground="#1e1e2e",
                        font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("R.TButton", background=RED, foreground="#1e1e2e",
                        font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("TRadiobutton", background=BG, foreground=FG, font=("Segoe UI", 9))
        style.map("TRadiobutton", background=[("active", BG)])

        tk.Label(self, text="eOrder Bot — Execução + TdC", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=ACCENT).pack(pady=(16, 2))
        tk.Label(self, text="Busca Execução (EXECUCAO) + Busca TdCs (CosampaCDU) → Cosampa Sul",
                 font=("Segoe UI", 9), bg=BG, fg="#6c7086").pack(pady=(0, 12))

        frm = tk.Frame(self, bg=BG)
        frm.pack(padx=24, fill="x")

        def lbl(t):
            tk.Label(frm, text=t, bg=BG, fg=FG,
                     font=("Segoe UI", 9, "bold"), anchor="w").pack(fill="x", pady=(8, 2))

        def entry(parent, var, show=None):
            e = tk.Entry(parent, textvariable=var, bg=ENTRY_BG, fg=FG,
                         insertbackground=FG, relief="flat",
                         font=("Segoe UI", 9), show=show or "")
            e.pack(fill="x", ipady=5)
            return e

        lbl("⚙️  Modo de execução")
        self.modo_var = tk.StringVar(value="1")
        modo_frame = tk.Frame(frm, bg=BG)
        modo_frame.pack(fill="x")
        ttk.Radiobutton(modo_frame, text="1 acesso (sequencial: Execução depois TdC)",
                        variable=self.modo_var, value="1",
                        command=self._atualizar_modo).pack(anchor="w")
        ttk.Radiobutton(modo_frame, text="2 acessos (paralelo: cada um abre seu eOrder)",
                        variable=self.modo_var, value="2",
                        command=self._atualizar_modo).pack(anchor="w")

        lbl("👤 Usuário (acesso 1 — Busca Execução)")
        self.usuario1_var = tk.StringVar()
        entry(frm, self.usuario1_var)

        lbl("🔑 Senha (acesso 1)")
        self.senha1_var = tk.StringVar()
        entry(frm, self.senha1_var, show="●")

        self.lbl_user2 = tk.Label(frm, text="👤 Usuário (acesso 2 — Busca TdCs)", bg=BG, fg=FG,
                                   font=("Segoe UI", 9, "bold"), anchor="w")
        self.usuario2_var = tk.StringVar()
        self.entry_user2 = tk.Entry(frm, textvariable=self.usuario2_var, bg=ENTRY_BG, fg=FG,
                                     insertbackground=FG, relief="flat", font=("Segoe UI", 9))

        self.lbl_senha2 = tk.Label(frm, text="🔑 Senha (acesso 2)", bg=BG, fg=FG,
                                    font=("Segoe UI", 9, "bold"), anchor="w")
        self.senha2_var = tk.StringVar()
        self.entry_senha2 = tk.Entry(frm, textvariable=self.senha2_var, bg=ENTRY_BG, fg=FG,
                                      insertbackground=FG, relief="flat", font=("Segoe UI", 9), show="●")

        lbl("📅 Data de execução (dd/mm/aaaa) — só para Busca Execução")
        self.data_var = tk.StringVar()
        entry(frm, self.data_var)

        lbl("📁 Pasta de download")
        row_d = tk.Frame(frm, bg=BG)
        row_d.pack(fill="x")
        self.dir_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Downloads"))
        tk.Entry(row_d, textvariable=self.dir_var, bg=ENTRY_BG, fg=FG,
                 insertbackground=FG, relief="flat",
                 font=("Segoe UI", 9)).pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))
        tk.Button(row_d, text="Procurar", bg=ACCENT, fg="#1e1e2e",
                  font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
                  command=self._escolher_pasta).pack(side="right", ipadx=10, ipady=4)

        bf = tk.Frame(self, bg=BG)
        bf.pack(pady=12, padx=24, fill="x")
        self.btn_ini = ttk.Button(bf, text="▶  INICIAR", style="G.TButton", command=self._iniciar)
        self.btn_ini.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.btn_par = ttk.Button(bf, text="■  PARAR", style="R.TButton",
                                  command=self._parar, state="disabled")
        self.btn_par.pack(side="left", fill="x", expand=True)

        tk.Label(self, text="📋 Log", bg=BG, fg=FG,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=24, pady=(8, 2))
        self.log_box = scrolledtext.ScrolledText(
            self, height=16, bg=ENTRY_BG, fg=FG,
            font=("Consolas", 9), relief="flat", state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        self._atualizar_modo()

    def _atualizar_modo(self):
        if self.modo_var.get() == "2":
            self.lbl_user2.pack(fill="x", pady=(8, 2))
            self.entry_user2.pack(fill="x", ipady=5)
            self.lbl_senha2.pack(fill="x", pady=(8, 2))
            self.entry_senha2.pack(fill="x", ipady=5)
        else:
            self.lbl_user2.pack_forget()
            self.entry_user2.pack_forget()
            self.lbl_senha2.pack_forget()
            self.entry_senha2.pack_forget()

    def _escolher_pasta(self):
        from tkinter import filedialog
        p = filedialog.askdirectory(title="Selecione a pasta de download")
        if p:
            self.dir_var.set(p)

    def log(self, msg):
        def _do():
            self.log_box.config(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _do)

    def _iniciar(self):
        modo = self.modo_var.get()
        usuario1 = self.usuario1_var.get().strip()
        senha1 = self.senha1_var.get().strip()
        data_str = self.data_var.get().strip()
        download_dir = self.dir_var.get().strip()

        if not usuario1 or not senha1:
            messagebox.showerror("Erro", "Preencha usuário e senha do acesso 1.")
            return
        if not data_str:
            messagebox.showerror("Erro", "Preencha a data de execução (dd/mm/aaaa).")
            return
        if not os.path.isdir(download_dir):
            messagebox.showerror("Erro", "Pasta de download inválida.")
            return

        if modo == "2":
            usuario2 = self.usuario2_var.get().strip()
            senha2 = self.senha2_var.get().strip()
            if not usuario2 or not senha2:
                messagebox.showerror("Erro", "Preencha usuário e senha do acesso 2.")
                return

        self.btn_ini.config(state="disabled")
        self.btn_par.config(state="normal")
        self.bots = []

        if modo == "1":
            self.log(f"🚀 Iniciando (1 acesso, sequencial)... data={data_str}")
            bot = EOrderExecucaoBot(lambda m: self.log(f"[Único] {m}"), download_dir)
            self.bots.append(bot)

            def _run():
                try:
                    bot._start_driver()
                    bot._login(usuario1, senha1)
                    bot.fazer_busca_execucao(data_str)
                    bot.fazer_tdc()
                except Exception as e:
                    bot._plog(f"❌ Erro: {e}")
                finally:
                    bot._plog("🏁 Finalizado.")
                    self.after(0, self._on_fim)

            threading.Thread(target=_run, daemon=True).start()

        else:
            self.log(f"🚀 Iniciando (2 acessos, paralelo)... data={data_str}")
            bot1 = EOrderExecucaoBot(lambda m: self.log(f"[Acesso1-Execução] {m}"), download_dir)
            bot2 = EOrderExecucaoBot(lambda m: self.log(f"[Acesso2-TdC] {m}"), download_dir)
            self.bots = [bot1, bot2]

            done = {"n": 0}

            def _marcar_fim():
                done["n"] += 1
                if done["n"] >= 2:
                    self.after(0, self._on_fim)

            def _run1():
                bot1.executar(usuario1, senha1, data_str)
                _marcar_fim()

            def _run2():
                bot2.executar_tdc(usuario2, senha2)
                _marcar_fim()

            threading.Thread(target=_run1, daemon=True).start()
            threading.Thread(target=_run2, daemon=True).start()

    def _on_fim(self):
        self.btn_ini.config(state="normal")
        self.btn_par.config(state="disabled")

    def _parar(self):
        for bot in self.bots:
            bot.stop_flag = True
        self.log("🛑 Sinal de parada enviado...")
        self.btn_par.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
