# Painel eOrder — Cosampa Sul

Painel de acompanhamento (TdC + Execução) do eOrder pra Cosampa Sul, feito
pra ficar numa TV de sala atualizando sozinho o dia todo. Site publicado:

**https://igorbezerra9-svg.github.io/painel-eorder-cosampa/**

Repositório: https://github.com/igorbezerra9-svg/painel-eorder-cosampa

---

## Como funciona (visão geral)

```
[Windows Task Scheduler]  →  [rodar_automatico.py]  →  [Selenium / eOrder]
   dispara de hora em hora        loga com 2 contas         baixa 2 exports
   (7h–17h, todo dia)             (Execução + TdC)          (.xlsx)
                                          │
                                          ▼
                                  publica sozinho no
                                  Supabase (tabela snapshots)
                                          │
                                          ▼
                              index.html (GitHub Pages)
                          lê do Supabase e atualiza sozinho
                          (realtime + poll a cada 2 min)
```

Ninguém precisa clicar em nada no dia a dia: a tarefa agendada roda o
bot, o bot publica, e o painel (aberto na TV) se atualiza sozinho.

---

## Arquivos do projeto

| Arquivo | Pra que serve |
|---|---|
| `index.html` | O painel em si (Painel Gerencial + Execução). Publicado via GitHub Pages. Lê/escreve no Supabase direto do navegador. |
| `eorder_execucao_bot.py` | Bot com interface gráfica (Tkinter) — login manual, roda Busca Execução + Busca TdCs, publica no Supabase automaticamente ao final de cada fluxo. |
| `rodar_automatico.py` | Versão sem interface — lê credenciais de `credenciais_eorder.json`, roda os dois fluxos em paralelo (2 Chromes) e sai. É isso que a tarefa agendada dispara. |
| `credenciais_eorder.example.json` | Modelo do arquivo de credenciais (esse **vai** pro Git). |
| `credenciais_eorder.json` | Credenciais reais (usuário/senha do eOrder) — **não vai pro Git** (`.gitignore`). Precisa ser criado em cada máquina que rodar a automação. |
| `instalar.bat` / `instalar.ps1` | Instalador de clique duplo — confere Python/Chrome, instala bibliotecas, cria `credenciais_eorder.json` se não existir, registra a tarefa agendada. Usar ao configurar numa máquina nova. |
| `publicar_teste.py` | Script avulso pra publicar manualmente um `.xlsx` já baixado no Supabase (útil pra testar sem rodar o Selenium de novo). |
| `eorder_tdc_bot.py` | Bot **separado**, sem relação com o painel — extrai telefone de observações de TdC por número de OS. |
| `DOC_REGRAS_GERAIS.md` | Todas as regras de classificação (Regulada/Não Regulada, Estado TdC, Prioridade, fim de semana, etc.) — ver seção abaixo. |
| `eOrderExecucaoBot.exe` / `eOrderTdCBot.exe` / `*.spec` | Versões compiladas dos bots com interface (não vão pro Git — são grandes e recompiláveis via `pyinstaller <nome>.spec`). |

---

## Arquitetura de dados (Supabase)

Projeto Supabase dedicado (não é o mesmo de outros projetos do usuário).
Tabela única:

```sql
create table snapshots (
  regiao text primary key,
  dados jsonb not null,
  atualizado_em timestamptz not null default now()
);
-- RLS liberando select/insert/update públicos (anon key)
```

- `regiao = 'Sul'` → export de TdC (o que alimenta o Painel Gerencial)
- `regiao = 'Execucao'` → export do dia da Busca Execução (aba Execução)

Cada publicação faz **upsert** (`Prefer: resolution=merge-duplicates`) —
sempre sobrescreve o snapshot anterior daquela região. Não é
cumulativo entre dias: representa "o que a busca trouxe na última vez
que o bot rodou".

⚠️ Só um subconjunto de colunas de cada export é publicado (não a
planilha inteira) — publicar todas as ~40 colunas do export completo
estourava o timeout de escrita do Supabase. As colunas usadas estão
em `COLS_TDC` / `COLS_EXECUCAO` no início de `eorder_execucao_bot.py`
e `publicar_teste.py` (mantenha os dois arquivos em sincronia se
adicionar uma coluna nova).

---

## Automação (tarefa agendada)

Tarefa do Windows: **`PainelEOrder_Automatico`**

- Gatilho **diário**, às 7h, repetindo de hora em hora por 10h (cobre
  7h–17h).
- `StartWhenAvailable`: se o PC ligar depois das 7h, ele recupera a
  rodada perdida assim que o Windows abrir — os horários seguintes
  continuam fixos (8h, 9h... não deslizam pro novo horário de ligar).
- Roda `rodar_automatico.py` com Chrome **visível** (não minimizado —
  ver histórico de bugs abaixo).
- Log de cada rodada em `automatico.log` (não versionado).

Comandos úteis (PowerShell):
```powershell
# Disparar uma rodada agora, fora do horário
Start-ScheduledTask -TaskName "PainelEOrder_Automatico"

# Ver status da última rodada
Get-ScheduledTaskInfo -TaskName "PainelEOrder_Automatico" | Format-List LastRunTime, LastTaskResult, NextRunTime

# Pausar/reativar (ex: rodando em 2 máquinas, só uma deve ficar ativa)
Disable-ScheduledTask -TaskName "PainelEOrder_Automatico"
Enable-ScheduledTask -TaskName "PainelEOrder_Automatico"
```

### Rodar em outra máquina

1. `git clone https://github.com/igorbezerra9-svg/painel-eorder-cosampa.git`
   (ou baixar o ZIP pelo GitHub, se não tiver Git instalado)
2. Clique duplo em `instalar.bat`
3. Preencher `credenciais_eorder.json` quando o Bloco de Notas abrir
   (usuário/senha reais do eOrder + pasta de Downloads **dessa
   máquina**)
4. Só uma máquina deve ficar com a tarefa ativa ao mesmo tempo — duas
   rodando junto com os mesmos logins do eOrder pode derrubar sessão
   uma da outra.

Pra atualizar o código depois (nova versão do painel/bot):
```powershell
git pull
```

---

## Regras de negócio

Todas as regras de classificação (Regulada/Não Regulada, Estado TdC,
Prioridade Estourada/Vence Hoje/No Prazo, regra de fim de semana,
Tipo de Equipe, Centro de Serviço/Supervisor) estão documentadas em
detalhe em **[`DOC_REGRAS_GERAIS.md`](DOC_REGRAS_GERAIS.md)**. Esse é
o documento de referência — qualquer mudança de regra deve ser
refletida lá também.

Resumo rápido:
- **Vence Hoje**: até 09:00 do dia seguinte — exceto no sábado, quando
  se estende até 09:00 de segunda (cobre sábado + domingo).
- **Vence Hoje** é a **única** métrica que não exclui os estados TdC
  `programáveis`/`a reprogramar` — as demais (Estouradas, No Prazo,
  totais, gráficos, tabela por equipe) excluem normalmente.
- **Produtivo/Improdutivo/Devolvido** (aba Execução): classificação
  baseada no `Código Resultado` de cada serviço, mapa
  `RESULTADO_PRODUTIVIDADE` dentro do `index.html`.

---

## Histórico de bugs/decisões importantes

Vale ler antes de mexer de novo nesse projeto:

- **Chrome minimizado quebra o eOrder.** Testado com login real: o
  site não renderiza corretamente alguns elementos (menus, campos de
  data) quando a janela está minimizada — falha em pontos diferentes
  a cada tentativa. `rodar_automatico.py` roda com `minimizado=False`
  de propósito. Não reativar sem testar de novo com cuidado.
- **`print()` com emoji quebra em `pythonw`/console Windows.** O
  console padrão do Windows usa a codepage `cp1252`, que não tem os
  emojis que o bot usa nas mensagens de log — isso derrubava a thread
  de login silenciosamente, antes até de preencher usuário/senha. O
  `log()` de `rodar_automatico.py` envolve o `print()` num try/except
  por causa disso; o log de arquivo (`automatico.log`, UTF-8) sempre
  funciona.
- **Gatilho "Once" não repete todo dia.** `New-ScheduledTaskTrigger
  -Once -At ... -RepetitionInterval ...` só repete dentro do próprio
  dia em que foi criado — no dia seguinte, `NextRunTime` fica vazio e
  a tarefa nunca mais dispara sozinha. O gatilho certo é `-Daily -At
  7:00AM` com a mesma `Repetition` aplicada por cima (é assim que
  `instalar.ps1` e a tarefa atual estão configurados agora).
- **Backslash em JSON precisa ser duplicado.** Usuários no formato
  `ENELINT\SEU_USUARIO` em `credenciais_eorder.json` têm que ser
  escritos como `ENELINT\\SEU_USUARIO` (JSON trata `\` como caractere
  de escape) — erro comum ao editar esse arquivo manualmente.
- **Scripts `.ps1` precisam de BOM UTF-8** pra rodar no Windows
  PowerShell 5.1 sem quebrar por causa de acentos/emoji nas mensagens
  — sem BOM, o parser lê como `cp1252` e o script quebra com erros de
  sintaxe sem relação nenhuma com o problema real.
- **GitHub Pages grátis é sempre público**, mesmo com o repositório
  marcado como privado (Pages em repo privado só existe em plano
  pago). Repositório e Pages deste projeto são públicos por decisão
  consciente do usuário — os dados de clientes (nome, endereço) ficam
  visíveis pra quem tiver o link. A chave anon do Supabase embutida no
  `index.html` também é pública por design (protegida por RLS, não
  por sigilo).
- **GitHub Pages fica com cache do navegador** depois de um push novo
  — se a página não parecer atualizada, force reload
  (`Ctrl+Shift+R`) ou adicione `?v=123` na URL.

---

## Verificar se está tudo funcionando

```powershell
# A tarefa existe e está pronta?
Get-ScheduledTask -TaskName "PainelEOrder_Automatico" | Format-List TaskName, State

# Quando foi a última vez que rodou, e quando roda de novo?
Get-ScheduledTaskInfo -TaskName "PainelEOrder_Automatico" | Format-List LastRunTime, LastTaskResult, NextRunTime

# O que aconteceu na última rodada?
Get-Content .\automatico.log -Tail 30
```

No painel, o indicador **● Atualizado HH:MM** no topo mostra a última
vez que qualquer uma das duas bases (Sul/Execução) foi publicada.
