# -*- coding: utf-8 -*-
"""Robô do "Controle de Ordens no Prazo" (Gerência Sul): 1 download por dia, às 05:00, e publica a região Controle_Prazo no Supabase.

Fluxo (mesmo do TdC do robô da Execução, só que com a BUSCA SALVA do eOrder):
  Busca TdCs -> busca salva CONTROLE_PRAZO_SUL (Processos: Gestão Inadimplências/Suspensão, Corte e Religação + Novas Conexões;
  Estado TdC: Finalizado e Encerrado) -> Data Resultado (dia 1 do mês até ontem) -> Buscar -> Exportar -> baixar.
  No começo do mês (e quando o "mês anterior" ainda não estiver fechado no painel) baixa também o mês anterior inteiro.

Depois de baixar, o cálculo (eorder_controle_prazo.py) cruza com as baixas da Execução (Supabase, com cache local) e publica só os
números (poucos KB). Nada é publicado se o arquivo baixado não bater com os filtros esperados (guarda contra a busca salva não ter
aplicado direito).

Uso:
    python eorder_controle_robo.py --testar                  # no seu PC: pede usuário/senha, baixa o MÊS ANTERIOR, imprime as tabelas e NÃO publica
    python eorder_controle_robo.py --testar --de 01/10/2026 --ate 31/10/2026
    python eorder_controle_robo.py --vps-teste               # na VPS: usa as credenciais que já estão lá (sem digitar nada), NÃO publica
    python eorder_controle_robo.py --auto                    # VPS (cron 05:00): credenciais de credenciais_eorder.json, publica
"""
import argparse
import getpass
import glob
import gzip
import json
import os
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta, timezone

import eorder_controle_prazo as cp

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BUSCA_SALVA = "CONTROLE_PRAZO_SUL"
NOME_EXPORT_CONTROLE = "ControlePrazoSul"
REGIAO = "Controle_Prazo"
ESTADOS_OK = {"Finalizado", "Encerrado"}
PROCESSOS_OK = {"Gestão Inadimplências", "Novas Conexões"}
COLUNAS_OBRIGATORIAS = ("Numero de Serviço", "Tipo TdC", "Município", "Data Prevista Finalização Trabalhos", "Estado TdC", "Processo")
LOOKBACK_BAIXAS_DIAS = 45        # baixas anteriores ao período entram (a 1ª tentativa de uma OS pode ser do mês passado)
CACHE_DIAS_FECHADOS = 2          # dias mais novos que isso são rebaixados a cada rodada (a reconsulta de ontem ainda pode mexer)
CACHE_BAIXAS = os.path.join(BASE_DIR, "controle_baixas_idx.json")
LOG_FILE = os.path.join(BASE_DIR, "controle_prazo.log")


def log(msg):
    linha = f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    try:
        print(linha)
    except Exception:
        pass
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass


def xlsx_integro(caminho):
    """True se o .xlsx baixado abre inteiro (zip válido e sem entrada corrompida)."""
    import zipfile
    try:
        if not zipfile.is_zipfile(caminho):
            return False
        with zipfile.ZipFile(caminho) as z:
            return z.testzip() is None
    except Exception:
        return False


def mesma_data(lido, esperado):
    """'01/10/26' == '01/10/2026' (o campo do eOrder tem máscara dd/MM/yy e pode devolver o ano com 2 dígitos)."""
    def norm(s):
        partes = str(s or "").strip().split("/")
        if len(partes) != 3 or not all(p.isdigit() for p in partes):
            return None
        d, m, a = (int(x) for x in partes)
        return d, m, a % 100
    return norm(lido) is not None and norm(lido) == norm(esperado)


DIAGNOSTICO_SEMPRE = False        # ligado nos testes: grava um retrato da tela depois de escolher a busca salva


# ══ Validação do arquivo baixado ═════════════════════════════════════════════════════════════════════════════════════════
def validar_export(linhas):
    """(ok, problemas). Só publica se TODAS as linhas forem de OS finalizadas/encerradas dos dois processos esperados: é o que
    garante que a busca salva aplicou os filtros (uma busca sem filtro traria o backlog aberto e estragaria os números)."""
    problemas = []
    if not linhas:
        return False, ["arquivo vazio (nenhuma OS)"]
    faltam = [c for c in COLUNAS_OBRIGATORIAS if c not in linhas[0]]
    if faltam:
        return False, ["faltam colunas no arquivo: " + ", ".join(faltam)]
    estados = {str(r.get("Estado TdC") or "").strip() for r in linhas}
    if not estados <= ESTADOS_OK:
        problemas.append("Estado TdC fora do esperado (só Finalizado/Encerrado): " + ", ".join(sorted(estados - ESTADOS_OK)))
    procs = {str(r.get("Processo") or "").strip() for r in linhas}
    if not procs <= PROCESSOS_OK:
        problemas.append("Processo fora do esperado: " + ", ".join(sorted(procs - PROCESSOS_OK)))
    return (not problemas), problemas


# ══ Baixas da Execução (Supabase) com cache local ════════════════════════════════════════════════════════════════════════
def _sb_headers():
    from eorder_execucao_bot import SB_KEY
    return {"apikey": SB_KEY, "Authorization": "Bearer " + SB_KEY, "Accept-Encoding": "gzip", "User-Agent": "ControlePrazo/1.0"}


def _sb_get(caminho):
    from eorder_execucao_bot import SB_URL
    req = urllib.request.Request(SB_URL + "/rest/v1/" + caminho, headers=_sb_headers())
    with urllib.request.urlopen(req, timeout=180) as r:
        corpo = r.read()
        if r.headers.get("Content-Encoding") == "gzip":
            corpo = gzip.decompress(corpo)
    return json.loads(corpo.decode("utf-8"))


def buscar_dia_supabase(iso):
    """Linhas da Execução de um dia fechado (snapshots_historico); None se esse dia não existe."""
    j = _sb_get(f"snapshots_historico?regiao=eq.Execucao&data=eq.{iso}&select=dados")
    return j[0]["dados"] if j else None


def carregar_baixas(de, ate, hoje, buscar_dia, classificar, cache_path=CACHE_BAIXAS, logf=log):
    """{'A0104...': [(datetime, classe), ...]} das baixas de `de` a `ate` (inclusive). Dias já fechados ficam em cache local, então
    depois da primeira carga só os 2 últimos dias são baixados do Supabase (egress mínimo)."""
    try:
        with open(cache_path, encoding="utf-8") as f:
            cache = json.load(f).get("dias", {})
    except Exception:
        cache = {}
    baixas, novos = {}, 0
    d = de
    while d <= ate:
        iso = d.isoformat()
        fechado = (hoje - d).days > CACHE_DIAS_FECHADOS
        if fechado and iso in cache:
            itens = cache[iso]
        else:
            linhas = buscar_dia(iso)
            if linhas is None:
                d += timedelta(days=1)
                continue
            itens = []
            for r in linhas:
                os_num = str(r.get("Numero de Serviço") or "").strip()
                quando = cp.parse_dt(r.get("Data fim Execução"))
                if os_num and quando:
                    itens.append([os_num, quando.strftime("%d/%m/%Y %H:%M"), classificar(r)])
            if fechado:
                cache[iso] = itens
                novos += 1
        for os_num, quando, classe in itens:
            baixas.setdefault(os_num, []).append((cp.parse_dt(quando), classe))
        d += timedelta(days=1)
    corte = (hoje - timedelta(days=LOOKBACK_BAIXAS_DIAS + 70)).isoformat()      # poda o cache (arquivo não cresce sem limite)
    cache = {k: v for k, v in cache.items() if k >= corte}
    try:
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump({"dias": cache}, f, ensure_ascii=False)
    except Exception as e:
        logf(f"⚠️  Não consegui gravar o cache de baixas: {e}")
    logf(f"📚 Baixas da Execução: {len(baixas)} OS em {(ate - de).days + 1} dia(s) ({novos} dia(s) novos baixados do Supabase)")
    return baixas


# ══ Montagem e publicação ════════════════════════════════════════════════════════════════════════════════════════════════
def montar_publicacao(linhas_atual, linhas_anterior, baixas, hoje, publicado=None, agora=None):
    """dados = {'atual': payload|None, 'anterior': payload|None}. Sem linhas novas pro mês anterior, mantém o que já estava publicado
    (se for do mesmo mês e fechado)."""
    atual_p, ant_p = cp.periodos(hoje)
    ontem = hoje - timedelta(days=1)
    atual = cp.calcular(linhas_atual, baixas, atual_p[0], atual_p[1], referencia=ontem, agora=agora) if (atual_p and linhas_atual is not None) else None
    if linhas_anterior is not None:
        anterior = cp.calcular(linhas_anterior, baixas, ant_p[0], ant_p[1], referencia=ant_p[1], agora=agora)
    else:
        antigo = (publicado or {}).get("anterior")
        anterior = antigo if antigo and antigo.get("periodo", {}).get("mes") == ant_p[0].strftime("%Y-%m") and antigo["periodo"].get("fechado") else None
    return {"atual": atual, "anterior": anterior}


def precisa_baixar_anterior(hoje, publicado):
    _, ant_p = cp.periodos(hoje)
    antigo = (publicado or {}).get("anterior")
    return not (antigo and antigo.get("periodo", {}).get("mes") == ant_p[0].strftime("%Y-%m") and antigo["periodo"].get("fechado"))


def ler_publicado():
    try:
        j = _sb_get(f"snapshots?regiao=eq.{REGIAO}&select=dados")
        return j[0]["dados"] if j else None
    except Exception as e:
        log(f"⚠️  Não consegui ler o que já está publicado: {e}")
        return None


def publicar(dados, post=None):
    from eorder_execucao_bot import SB_URL, SB_KEY
    corpo = json.dumps({"regiao": REGIAO, "dados": dados, "atualizado_em": datetime.now(timezone.utc).isoformat()}, ensure_ascii=False).encode("utf-8")
    if post is not None:
        return post(corpo)
    req = urllib.request.Request(SB_URL + "/rest/v1/snapshots", data=corpo, method="POST", headers={
        "apikey": SB_KEY, "Authorization": "Bearer " + SB_KEY, "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates"})
    ultimo = None
    for tentativa in range(3):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return r.status
        except urllib.error.HTTPError:
            raise
        except Exception as e:
            ultimo = e
            time.sleep(3)
    raise ultimo


def imprimir_tabelas(payload):
    """Mesma leitura da planilha, no console (usado no --testar pra você comparar com a ABA ORDENS NO PRAZO)."""
    p = payload["periodo"]
    print(f"\n=== {payload['titulo_data']} | período {p['de']} a {p['ate']} | meta {payload['meta']*100:.2f}% ===")
    for nome, g in (("LIGAÇÕES NOVAS", "ligacao_nova"), ("RELIGAÇÕES", "religacao")):
        print(f"\n{nome:16s} {'EXECUTADAS':>11s} {'FORA PRAZO':>11s} {'REALIZADO':>10s} {'AVANÇO':>9s}")
        for cs, (e, f, real) in cp.resumo(payload["grupos"][g]).items():
            print(f"  {cs:14s} {e:11d} {f:11d} {real*100:9.2f}% {real/payload['meta']*100:8.2f}%")
    if payload["fora_do_prazo"]:
        print("\nFora do prazo (maiores atrasos):")
        for x in payload["fora_do_prazo"][:10]:
            print(f"  {x['os']}  {x['cs']:12s} prazo {x['prazo']}  1ª baixa {x['baixa']}  ({x['atraso_h']} h)")
    print("\navisos:", payload["avisos"])


# ══ Navegador (reaproveita login/navegação/exportação do robô da Execução SEM alterar aquele arquivo) ═════════════════════
def _criar_bot_controle(*args, **kw):
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import Select, WebDriverWait
    import eorder_execucao_bot as base

    class EOrderControleBot(base.EOrderExecucaoBot):
        # Tudo por TEXTO do rótulo (não por XPath absoluto): a tela da busca salva e das datas de resultado não tem XPath conhecido.
        XP_SELECT_BUSCA_SALVA = '(//*[normalize-space(text())="Buscas salvas"]/following::select)[1]'
        XP_FIELDSET_DATA_RESULTADO = ('//fieldset[.//legend[contains(normalize-space(.), "Data Resultado")]]'
                                      ' | //legend[contains(normalize-space(.), "Data Resultado")]/ancestor::fieldset[1]')
        XP_INPUTS_APOS_ROTULO = '(//*[normalize-space(text())="Data Resultado" or normalize-space(text())="— Data Resultado —"])[1]/following::input[not(@type="hidden")]'

        def _diagnostico(self, motivo):
            """Grava um retrato da tela (txt + foto PNG): linhas marcadas nos filtros, selects, campos de data e os textos de
            'exportar' de cada frame -- pra eu enxergar o que o portal está mostrando quando algo não acha."""
            carimbo = f"{datetime.now():%Y%m%d_%H%M%S}"
            base_nome = os.path.join(self.download_dir, f"controle_diagnostico_{carimbo}")
            try:
                self.driver.switch_to.default_content()
                info = self.driver.execute_script("""
                    var sel = Array.from(document.querySelectorAll('.tvRowSelected')).slice(0, 80).map(function(r){return (r.innerText||'').trim().slice(0,60);});
                    var datas = Array.from(document.querySelectorAll('fieldset')).map(function(f){var l=f.querySelector('legend'); return {legenda:(l?l.innerText:(f.innerText||'').slice(0,40)).trim(), inputs:Array.from(f.querySelectorAll('input')).map(function(i){return i.value;})};}).slice(0,30);
                    return {titulo: document.title, linhasMarcadas: sel, fieldsets: datas, iframes: document.querySelectorAll('iframe').length};
                """)
            except Exception as e:
                info = {"erro_ao_ler_tela": str(e)}
            try:                              # textos de exportação/aviso em cada frame (o menu pode estar em outro frame)
                textos = {}
                frames = [None] + self.driver.find_elements(By.TAG_NAME, "iframe")
                for i, fr in enumerate(frames):
                    try:
                        self.driver.switch_to.default_content()
                        if fr is not None:
                            self.driver.switch_to.frame(fr)
                        textos[f"frame{i}"] = {"export": self._dump_texto_visivel("xport", 15), "avisos": self._dump_texto_visivel("limite", 5) + self._dump_texto_visivel("erro", 5)}
                        html = self.driver.execute_script("""
                            var e = Array.from(document.querySelectorAll('body *')).find(function(x){return x.children.length===0 && (x.textContent||'').indexOf('superior ao limite')>=0;});
                            if (!e) return ''; var b = e; for (var k = 0; k < 4 && b.parentElement; k++) b = b.parentElement; return b.outerHTML.slice(0, 3500);
                        """)
                        if html:
                            textos[f"frame{i}"]["janela_de_aviso_html"] = html
                    except Exception as e:
                        textos[f"frame{i}"] = f"erro: {e}"
                self.driver.switch_to.default_content()
                info["textos_por_frame"] = textos
            except Exception as e:
                info["textos_por_frame"] = f"erro: {e}"
            try:
                self.driver.switch_to.default_content()
                self.driver.save_screenshot(base_nome + ".png")
                info["foto"] = base_nome + ".png"
            except Exception as e:
                info["foto"] = f"não consegui: {e}"
            try:
                with open(base_nome + ".txt", "w", encoding="utf-8") as f:
                    f.write(f"motivo: {motivo}\n" + json.dumps(info, ensure_ascii=False, indent=1))
                self._plog(f"🩺 Diagnóstico gravado em {base_nome}.txt (+ .png)")
            except Exception as e:
                self._plog(f"🩺 Não consegui gravar o diagnóstico: {e}")

        def _esperar_sem_carregando(self, max_s=180):
            """Depois do Buscar a tela mostra o 'carregando' (darkdiv) enquanto monta a lista (12 mil linhas). Esperar ele sumir evita
            clicar no menu de exportar com a tela ainda se remontando."""
            self.driver.switch_to.default_content()
            try:
                WebDriverWait(self.driver, max_s).until(EC.invisibility_of_element_located((By.ID, "darkdiv")))
            except TimeoutException:
                self._plog(f"⚠️  A tela continuou 'carregando' por mais de {max_s}s; seguindo mesmo assim...")
            time.sleep(3)

        def _click_texto_em_frames(self, texto):
            """_click_por_texto só olha o frame atual; o menu de exportar pode estar em outro. Tenta o topo e cada iframe."""
            self.driver.switch_to.default_content()
            if self._click_por_texto(texto):
                return True
            for fr in self.driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame(fr)
                    if self._click_por_texto(texto):
                        return True
                except Exception:
                    continue
            self.driver.switch_to.default_content()
            return False

        def _selecionar_busca_salva(self, nome):
            self._plog(f"🔖 Escolhendo a busca salva '{nome}'...")
            el = self._find(self.XP_SELECT_BUSCA_SALVA, EC.presence_of_element_located, timeout=20)
            opcoes = [o.text.strip() for o in Select(el).options]
            self._plog(f"   buscas salvas disponíveis: {opcoes}")
            if nome not in opcoes:
                self._diagnostico(f"busca salva '{nome}' não aparece na lista (usuário diferente do que criou?)")
                raise RuntimeError(f"A busca salva '{nome}' não aparece para este usuário. Disponíveis: {opcoes}")
            Select(el).select_by_visible_text(nome)
            time.sleep(4)               # a tela reaplica os filtros salvos
            if DIAGNOSTICO_SEMPRE:
                self._diagnostico("depois de escolher a busca salva (confira os filtros marcados)")

        def _achar_em_frames(self, xpath):
            """Elementos do XPath no topo da página ou no 1º iframe onde existirem. Deixa o driver DENTRO desse frame (os elementos só
            valem nele). Não depende de onde o driver estava antes -- o diagnóstico, por exemplo, troca de frame e estragava a busca."""
            self.driver.switch_to.default_content()
            achou = self.driver.find_elements(By.XPATH, xpath)
            if achou:
                return achou
            for fr in self.driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame(fr)
                    achou = self.driver.find_elements(By.XPATH, xpath)
                    if achou:
                        return achou
                except Exception:
                    continue
            self.driver.switch_to.default_content()
            return []

        def _campos_data_resultado(self):
            """[campo De, campo A] da Data Resultado, ou [] se a seção não está aberta. Cada data tem DOIS inputs: o que se digita
            (vazio) e outro que só mostra o formato (dd/MM/yy) -- por isso pulamos os que mostram formato."""
            achou = self._achar_em_frames(self.XP_FIELDSET_DATA_RESULTADO)      # deixa o driver DENTRO do frame onde achou
            if not achou:
                return []
            inputs = [i for i in achou[0].find_elements(By.XPATH, ".//input[not(@type='hidden')]") if i.is_displayed()]
            digitaveis = [i for i in inputs if (i.get_attribute("value") or "").strip() not in ("dd/MM/yy", "dd/MM/yyyy", "HH:mm")]
            return digitaveis[:2] if len(digitaveis) >= 2 else []

        def _datas_resultado(self, de_str, ate_str):
            # Com a busca salva a seção "Datas Referência" costuma JÁ estar aberta: clicar nela de novo FECHA (foi o erro do 1º teste).
            campos = self._campos_data_resultado()
            if not campos:
                self._plog("📅 Abrindo Datas Referência...")
                self._click(base.XP_DATAS_REF_SPAN, timeout=20)
                time.sleep(1.5)
                campos = self._campos_data_resultado()
            if len(campos) < 2:
                self._diagnostico("campos de Data Resultado não encontrados")
                raise RuntimeError("Não achei os dois campos de 'Data Resultado' (De/A) na tela.")
            for campo, valor in ((campos[0], de_str), (campos[1], ate_str)):
                campo.click()
                campo.clear()
                campo.send_keys(valor)
            time.sleep(0.5)
            lidos = [c.get_attribute("value") for c in campos[:2]]
            self._plog(f"📅 Data Resultado: {lidos[0]} até {lidos[1]}")
            if not (mesma_data(lidos[0], de_str) and mesma_data(lidos[1], ate_str)):
                self._diagnostico(f"datas não ficaram como pedido ({lidos} != {[de_str, ate_str]})")
                raise RuntimeError(f"As datas digitadas não ficaram como pedido: {lidos} (esperado {[de_str, ate_str]})")

        JS_FECHAR_AVISO = """
            var els = Array.from(document.querySelectorAll('body *')).filter(function(e){ return e.children.length === 0 && (e.textContent||'').indexOf('superior ao limite') >= 0; });
            if (!els.length) return 'ausente';
            var cand = [], box = els[0];
            for (var i = 0; i < 6 && box; i++, box = box.parentElement) {
              Array.from(box.querySelectorAll('*')).forEach(function(e){
                var t = (e.textContent||'').trim(), r = e.getBoundingClientRect();
                var cls = String((e.className && e.className.baseVal !== undefined) ? e.className.baseVal : (e.className||''));
                var tit = (e.getAttribute('title')||'') + ' ' + (e.getAttribute('alt')||'');
                if (r.width > 0 && (/clos|fech/i.test(cls) || /clos|fech/i.test(tit) || t === '\u00d7' || t === 'x' || t === 'X' ||
                    (t === '' && r.width > 8 && r.width < 50 && r.height > 8 && r.height < 50))) cand.push(e);
              });
              if (cand.length) break;
            }
            if (!cand.length) return 'sem-botao';
            var alvo = cand[cand.length - 1];
            alvo.click();
            return 'clicou:' + (alvo.className || alvo.tagName);
        """

        def _aviso_limite_presente(self):
            """True se o aviso 'Número de TdC superior ao limite' está na tela (em qualquer frame)."""
            return bool(self._texto_em_frames("superior ao limite"))

        def _texto_em_frames(self, trecho):
            achados = []
            self.driver.switch_to.default_content()
            frames = [None] + self.driver.find_elements(By.TAG_NAME, "iframe")
            for fr in frames:
                try:
                    self.driver.switch_to.default_content()
                    if fr is not None:
                        self.driver.switch_to.frame(fr)
                    achados += self._dump_texto_visivel(trecho, 5)
                except Exception:
                    continue
            self.driver.switch_to.default_content()
            return achados

        def _fechar_aviso_limite(self):
            """Depois do Buscar com mais de 5.000 resultados o eOrder mostra 'Serão visualizados somente os primeiros 5.000 TdCs' e o
            aviso BLOQUEIA a página até fechar (era isso que impedia o menu Exportar de abrir). O limite é só da tela: o export traz tudo."""
            if not self._aviso_limite_presente():
                return
            self._plog("⚠️  Aviso 'Número de TdC superior ao limite' na tela — fechando...")
            for tentativa in range(1, 5):
                self.driver.switch_to.default_content()
                frames = [None] + self.driver.find_elements(By.TAG_NAME, "iframe")
                for fr in frames:
                    try:
                        self.driver.switch_to.default_content()
                        if fr is not None:
                            self.driver.switch_to.frame(fr)
                        r = self.driver.execute_script(self.JS_FECHAR_AVISO)
                        if r and r != "ausente":
                            self._plog(f"   aviso: {r}")
                    except Exception:
                        continue
                self.driver.switch_to.default_content()
                time.sleep(1.5)
                if not self._aviso_limite_presente():
                    self._plog("   ✅ Aviso fechado.")
                    return
                try:                                        # 2º recurso: tecla ESC
                    from selenium.webdriver.common.keys import Keys
                    from selenium.webdriver.common.action_chains import ActionChains
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(1)
                except Exception:
                    pass
                if not self._aviso_limite_presente():
                    self._plog("   ✅ Aviso fechado (ESC).")
                    return
            self._diagnostico("não consegui fechar o aviso 'Número de TdC superior ao limite'")
            raise RuntimeError("Não consegui fechar o aviso 'Número de TdC superior ao limite' (o menu Exportar fica bloqueado).")

        def _total_da_lista(self):
            """N de 'Lista TdC - 10/12292' (total de resultados da busca) ou None. Lê o texto inteiro de cada frame (o rótulo pode
            estar partido em vários elementos, então procurar por elemento não achava)."""
            import re
            self.driver.switch_to.default_content()
            frames = [None] + self.driver.find_elements(By.TAG_NAME, "iframe")
            total = None
            for fr in frames:
                try:
                    self.driver.switch_to.default_content()
                    if fr is not None:
                        self.driver.switch_to.frame(fr)
                    texto = self.driver.execute_script("return document.body ? document.body.innerText : ''") or ""
                    m = re.search(r"Lista TdC\s*-\s*\d+\s*/\s*(\d+)", texto)
                    if m:
                        total = int(m.group(1))
                        break
                except Exception:
                    continue
            self.driver.switch_to.default_content()
            return total

        def _exportar_controle(self, nome_export):
            self._plog("📤 Exportando TdCs (controle)...")
            achou = False
            for tentativa in range(1, 5):
                self._click(base.XP_TRES_PONTOS_EXPORT_TDC, timeout=20)
                for _ in range(12):                      # a lista grande demora pra renderizar o menu: espera até ~12s
                    time.sleep(1)
                    if self._click_texto_em_frames("Exportar em xls"):
                        achou = True
                        break
                if achou:
                    break
                self._plog(f"   🔁 Menu de exportar não apareceu (tentativa {tentativa}/4) — reclicando nos 3 pontinhos...")
            if not achou:
                self._diagnostico("item 'Exportar em xls' não apareceu no menu")
                raise TimeoutException("Item 'Exportar em xls' não encontrado no menu do TdC.")
            self._plog("   ...abriu menu de exportação")
            self._exportar_generico(nome_export, base.XP_BTN_OK_EXP_TDC, base.XP_FECHAR_MSG_TDC)

        def baixar_periodo(self, de, ate):
            """Roda a busca salva para [de, ate] e devolve as linhas do export (lista de dicts)."""
            self._navegar_busca_tdcs()
            self._selecionar_busca_salva(BUSCA_SALVA)
            self._datas_resultado(de.strftime("%d/%m/%Y"), ate.strftime("%d/%m/%Y"))
            self._buscar_tdc()
            self._esperar_sem_carregando()
            self._fechar_aviso_limite()
            total_lista = self._total_da_lista()
            self._plog(f"📋 A busca encontrou {total_lista if total_lista is not None else '?'} TdCs")
            if DIAGNOSTICO_SEMPRE:
                self._diagnostico("depois do Buscar (aviso fechado)")
            if total_lista == 0:
                # Sem nenhuma linha o eOrder nem oferece "exportar" (início de mês, domingo/feriado). Só aceita o zero depois de
                # reler a lista: um 0 lido cedo demais não pode virar uma tabela zerada publicada.
                time.sleep(20)
                self._esperar_sem_carregando()
                if self._total_da_lista() == 0:
                    self._plog("📭 Nenhum TdC finalizado nesse período — nada a exportar (tabela zerada).")
                    return []
                self._plog("⚠️  A lista mudou na segunda leitura — seguindo com a exportação normal.")
            self._exportar_controle(NOME_EXPORT_CONTROLE)
            caminho = None
            for tentativa_arq in range(1, 4):
                ok = self._baixar_exportacao(base.XP_LISTA_EXPORT_TDC, base.XP_TRES_PONTOS_LISTA_TDC, NOME_EXPORT_CONTROLE, espera_max=1200)
                if not ok:
                    raise RuntimeError("O export não ficou pronto a tempo.")
                time.sleep(5)
                caminho = self._achar_export_mais_recente(NOME_EXPORT_CONTROLE, timestamp_esperado=ok if isinstance(ok, str) else None)
                if not caminho:
                    raise RuntimeError("Não achei o arquivo baixado na pasta de downloads.")
                self._plog(f"💾 Arquivo: {caminho}")
                if xlsx_integro(caminho):
                    break
                # O eOrder lista o arquivo já com algum tamanho enquanto AINDA gera (apareceu com 517 KB e terminou com 13 MB):
                # baixado nessa hora ele vem cortado. Apaga e baixa de novo daqui a pouco, já completo.
                self._plog(f"⚠️  O arquivo baixado veio incompleto (o eOrder ainda estava gerando) — apagando e baixando de novo em 60s ({tentativa_arq}/3)...")
                try:
                    os.remove(caminho)
                except OSError:
                    pass
                caminho = None
                time.sleep(60)
            if not caminho:
                raise RuntimeError("O arquivo exportado continuou incompleto depois de 3 tentativas.")
            linhas = self._xlsx_para_linhas(caminho, colunas=None)
            self._limpar_exports_antigos(NOME_EXPORT_CONTROLE)
            self._plog(f"📄 O arquivo tem {len(linhas)} linhas")
            if total_lista and len(linhas) < total_lista * 0.995:      # export cortado (o limite de 5.000 é só da tela; aqui confere)
                raise RuntimeError(f"O arquivo veio com {len(linhas)} linhas, mas a busca tinha {total_lista}: export incompleto, nada será publicado.")
            return linhas

    return EOrderControleBot(*args, **kw)


def baixar_periodos(usuario, senha, periodos, download_dir, minimizado=False, tentativas=3):
    """Um login só, uma busca por período. Devolve {nome: linhas}."""
    bot = _criar_bot_controle(lambda m: log(m), download_dir, minimizado=minimizado)
    ultimo = None
    for t in range(1, tentativas + 1):
        try:
            bot._start_driver()
            bot._login(usuario, senha)
            saida = {}
            for nome, (de, ate) in periodos.items():
                log(f"▶ Baixando '{nome}': {de:%d/%m/%Y} a {ate:%d/%m/%Y}")
                saida[nome] = bot.baixar_periodo(de, ate)
            return saida
        except Exception as e:
            ultimo = e
            log(f"❌ Erro (tentativa {t}/{tentativas}): {e}")
            if t < tentativas:
                time.sleep(5)
        finally:                      # sempre fecha o navegador (sucesso ou erro), pra não deixar Chrome pendurado na VPS
            try:
                if bot.driver:
                    bot.driver.quit()
            except Exception:
                pass
            bot.driver = None
    raise ultimo


# ══ Execução ═════════════════════════════════════════════════════════════════════════════════════════════════════════════
def _classificador():
    from eorder_execucao_bot import _classificar_resultado
    return _classificar_resultado


def _credenciais_arquivo():
    """Usuário/senha do MESMO arquivo que o robô da Execução usa na VPS (credenciais_eorder.json). Nunca imprime nada disso."""
    with open(os.path.join(BASE_DIR, "credenciais_eorder.json"), encoding="utf-8") as f:
        cred = json.load(f)
    # Mesma regra do rodar_automatico.py: modo 2 = o TdC usa o acesso 2; modo 1 = um acesso só (acesso1). "controle" (opcional) manda em tudo.
    acesso = cred.get("controle") or (cred["acesso2"] if str(cred.get("modo", "1")) == "2" else cred["acesso1"])
    return acesso["usuario"], acesso["senha"], (cred.get("download_dir") or os.path.join(os.path.expanduser("~"), "Downloads"))


def principal():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--testar", action="store_true", help="pede usuário/senha, baixa, imprime as tabelas e NÃO publica")
    ap.add_argument("--auto", action="store_true", help="modo cron da VPS (credenciais_eorder.json) e publica")
    ap.add_argument("--vps-teste", action="store_true", help="na VPS, com as credenciais já salvas lá: baixa o mês anterior, imprime as tabelas e NÃO publica")
    ap.add_argument("--de", help="dd/mm/aaaa (só nos testes)")
    ap.add_argument("--ate", help="dd/mm/aaaa (só nos testes)")
    ap.add_argument("--publicar", action="store_true", help="no --testar, publica também")
    ap.add_argument("--sem-publicar", action="store_true", help="com --auto: faz tudo (mês atual + mês anterior), imprime as tabelas e NÃO publica")
    args = ap.parse_args()
    if not (args.testar or args.auto or args.vps_teste):
        ap.print_help()
        return 2
    hoje = date.today()
    teste = args.testar or args.vps_teste
    global DIAGNOSTICO_SEMPRE
    DIAGNOSTICO_SEMPRE = teste

    if teste:
        if args.testar:
            usuario = input("Usuário do eOrder: ").strip()
            senha = getpass.getpass("Senha (não aparece na tela): ")
            download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        else:
            usuario, senha, download_dir = _credenciais_arquivo()
        atual_p, ant_p = cp.periodos(hoje)
        if args.de and args.ate:
            de = datetime.strptime(args.de, "%d/%m/%Y").date(); ate = datetime.strptime(args.ate, "%d/%m/%Y").date()
        else:
            de, ate = ant_p          # padrão do teste: o mês anterior fechado -- dá pra comparar com a sua planilha
        periodos = {"teste": (de, ate)}
        minimizado = False
    else:
        usuario, senha, download_dir = _credenciais_arquivo()
        publicado = ler_publicado()
        atual_p, ant_p = cp.periodos(hoje)
        periodos = {}
        if atual_p:
            periodos["atual"] = atual_p
        if precisa_baixar_anterior(hoje, publicado):
            periodos["anterior"] = ant_p
        minimizado = False

    if not periodos:
        log("Nada a baixar hoje (mês anterior já fechado no painel e o mês atual ainda não começou).")
        return 0
    log(f"=== Controle no Prazo iniciado ({'teste' if teste else 'automático'}) — períodos: " + ", ".join(f"{k} {v[0]:%d/%m}-{v[1]:%d/%m}" for k, v in periodos.items()) + " ===")
    baixados = baixar_periodos(usuario, senha, periodos, download_dir, minimizado=minimizado)
    for nome, linhas in baixados.items():
        if not linhas:                 # período confirmado sem TdCs (ver baixar_periodo): nada a validar, a tabela sai zerada
            log(f"🔎 '{nome}': 0 linhas — período sem TdCs finalizados")
            continue
        ok, problemas = validar_export(linhas)
        log(f"🔎 '{nome}': {len(linhas)} linhas — " + ("conferido" if ok else "PROBLEMAS: " + "; ".join(problemas)))
        if not ok:
            log("🛑 Nada foi publicado: o arquivo não bate com os filtros esperados (confira a busca salva).")
            return 1

    todos = [d for k, (de, ate) in periodos.items() for d in (de, ate)]
    inicio = min(todos) - timedelta(days=LOOKBACK_BAIXAS_DIAS)
    baixas = carregar_baixas(inicio, hoje - timedelta(days=1), hoje, buscar_dia_supabase, _classificador())

    if teste and "teste" in baixados:
        de, ate = periodos["teste"]
        payload = cp.calcular(baixados["teste"], baixas, de, ate, referencia=ate)
        imprimir_tabelas(payload)
        arq = os.path.join(download_dir, "controle_prazo_teste.json")
        with open(arq, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
        log(f"📝 Resultado completo gravado em {arq} (NÃO publicado)")
        if args.publicar:
            publicar({"atual": None, "anterior": payload})
            log("☁ Publicado (modo teste com --publicar).")
        return 0

    dados = montar_publicacao(baixados.get("atual"), baixados.get("anterior"), baixas, hoje, publicado=publicado)
    if args.sem_publicar:
        for nome in ("atual", "anterior"):
            if dados.get(nome):
                print(f"\n########## {nome.upper()} ##########")
                imprimir_tabelas(dados[nome])
        arq = os.path.join(download_dir, "controle_prazo_auto_simulado.json")
        with open(arq, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=1)
        log(f"📝 Simulação gravada em {arq} (NÃO publicado)")
        return 0
    status = publicar(dados)
    log(f"☁ Controle publicado — status {status} | atual: {bool(dados['atual'])} | anterior: {bool(dados['anterior'])}")
    return 0


if __name__ == "__main__":
    sys.exit(principal())
