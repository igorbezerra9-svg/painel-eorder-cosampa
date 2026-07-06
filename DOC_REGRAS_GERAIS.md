# Regras Gerais do Sistema — Gerador de Capa de Rota (DESUL Sul)

Resumo de **todas** as regras de classificação que o `index.html` usa —
Regulada/Não Regulada, Estado TdC, Prioridade (estourada/vence hoje/no
prazo), Tipo de Equipe, Centro de Serviço e Supervisor. Serve como
referência única pra não perder nenhuma regra de vista.

> Documentação complementar e mais detalhada sobre Centro de Serviço e
> Supervisor por região: `DOC_centros_de_servico.md`.

---

## 1. Regulada (REG) x Não Regulada (NREG)

Classificação baseada no **código do tipo de serviço** (colunas `Tipo
Remessa WIN` / `Tipo de Serviço`), função `getTipoOrdem(r)`:

```js
const TIPO_REGULADA = {
  'OUTMML':true,'OUTMGD':true,'OUTVPC':true,'OUTIRD':true,'OUTTCP':true,
  'OUTNCU':true,'OUTAMC':true,'OUTSEM':true,'OUTIRM':true,'NOVLBT':true,
  'OUTAQM':true,'OUTSDJ':true,'OUTSRL':true,'OUTTME':true,'OUTNCR':true,
  'OUTDCM':true,'OUTACT':true,'OUTIMB':true,'OUTRAR':true,'OUTRNR':true,
  'OUTRNU':true,'OUTRIM':true,'OUTRAU':true,'OUTCP0':true,'OUTRNL':true,
  'OUTVMT':true,'NOVPRO':true,'OUTOTC':true,'OUTIMR':true,'OUTSTB':true,
  'OUTIRC':true,'OUTMLD':true,'OUTMMI':true
};
const TIPO_NAO_REGULADA = {
  'OUTCDN':true,'OUTRNN':true,'OUTERM':true,'OUTRFS':true,'OUTCSC':true,
  'OUTCRT':true
};

function getTipoOrdem(r) {
  const t = (r['Tipo Remessa WIN'] || r[' Tipo de Serviço'] || r['Tipo de Serviço'] || '').trim().toUpperCase();
  const cod = t.split(' - ')[0].trim().split(' ')[0]; // pega só o código, ex: OUTCDN
  if (TIPO_REGULADA[cod]) return 'REG';
  if (TIPO_NAO_REGULADA[cod]) return 'NREG';
  return 'OUTRO';
}
```

- **REG** — 32 códigos cadastrados (manutenções, ordens técnicas
  regulares, etc.)
- **NREG** — 6 códigos cadastrados: `OUTCDN` (corte por Dunning/
  inadimplência), `OUTRNN`, `OUTERM`, `OUTRFS`, `OUTCSC`, `OUTCRT`
- **OUTRO** — qualquer código que não esteja em nenhuma das duas listas
  (aparece sem badge REG/NR na interface)

Usado em: filtro "Reguladas / Não reguladas" da aba Gerador e do Painel
Gerencial (`pgTipo`), badges `[REG]` / `[NÃO REG]` nas listagens, gráfico
"Por tipo" do Painel Gerencial, e coluna REG/NREG na tabela por equipe.

⚠️ Se aparecer um código de tipo de serviço novo que não esteja em
nenhuma das duas listas, ele cai em "OUTRO" silenciosamente — vale
revisar essa lista de tempos em tempos.

---

## 2. Estado TdC

Estados observados na base (`Estado TdC`):

| Estado | O que significa | Tem Código Equipe? |
|---|---|---|
| `Designado` | Equipe já definida, ainda não confirmada em campo | Sim (100%) |
| `Enviado ao Campo` | Equipe recebeu e está executando | Sim (100%) |
| `programáveis` *(note o espaço antes, na base vem como `' programáveis'`)* | Backlog, normalmente sem equipe ainda | Raramente (~2%) |
| `A Reprogramar` | Já foi designado, mas caiu pra reprogramação | Sim (100%) |

### Onde cada estado é considerado

- **Aba Gerador** (`getFiltered()`) — **exclui** `programáveis` e
  `A Reprogramar` sempre. Faz sentido: só se gera capa de rota pra quem
  tem equipe de fato em campo.
- **Painel Gerencial** (`pgGetFiltrado()` / `pgGetFiltradoTodos()`) — ver
  seção 3 abaixo, tem regra própria e mais fina.

---

## 3. Prioridade (Estourada / Vence Hoje / No Prazo)

Baseada na coluna **`Data Prevista Finalização Trabalhos`**, função
`getPrioridade(r)`:

```js
function getPrioridade(r) {
  const dt = parseDateTimeStr(r['Data Prevista Finalização Trabalhos'] || '');
  if (!dt) return 'ok';
  const agora = new Date();
  if (dt < agora) return 'est';                       // Estourada: prazo já passou
  const amanha9h = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()+1, 9, 0, 0);
  if (dt <= amanha9h) return 'hj';                     // Vence Hoje: até 09h de amanhã
  return 'ok';                                          // No Prazo
}
```

- **Estourada (`est`)** — prazo já passou do momento atual
- **Vence Hoje (`hj`)** — vence entre agora e 09:00 do dia seguinte
  (dá uma margem pra cobrir o começo do turno seguinte) — **exceto no
  sábado, ver regra abaixo**
- **No Prazo (`ok`)** — vence depois disso, ou não tem data preenchida

### Regra do fim de semana (desde 05/07/2026)

As equipes não trabalham domingo. Então, se hoje é **sábado**, tudo que
vence sábado **ou** domingo entra como "Vence Hoje" — a janela se
estende até 09:00 de **segunda-feira** em vez de até 09:00 de domingo.
Isso evita que serviços que vencem domingo fiquem escondidos como "no
prazo" no sábado e virem "estourados" direto na segunda, sem chance de
ninguém ter resolvido no fim de semana.

```js
function getPrioridade(r) {
  const dt = parseDateTimeStr(r['Data Prevista Finalização Trabalhos'] || '');
  if (!dt) return 'ok';
  const agora = new Date();
  if (dt < agora) return 'est';

  const ehSabado = hoje.getDay() === 6; // 0=domingo ... 6=sábado
  const diasJanela = ehSabado ? 2 : 1;
  const limite9h = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + diasJanela, 9, 0, 0);
  if (dt <= limite9h) return 'hj';

  return 'ok';
}
```

| Hoje é... | Janela de "Vence Hoje" |
|---|---|
| Qualquer dia da semana (dom–sex) | agora até 09:00 do dia seguinte |
| **Sábado** | agora até **09:00 de segunda-feira** (cobre sábado + domingo) |

Essa regra vale tanto pra aba Gerador quanto pro Painel Gerencial, porque
as duas usam a mesma função `getPrioridade()`. No Painel Gerencial, o
texto do card "Vence Hoje" também muda no sábado pra deixar claro que a
janela está estendida.

### Regra especial no Painel Gerencial (desde 05/07/2026)

O Painel Gerencial usa **duas bases diferentes** dependendo da métrica —
essa é a regra mais importante de todo o sistema, então repetindo aqui
além do `DOC_centros_de_servico.md`:

| Métrica | Base usada | Exclui programáveis/a reprogramar? |
|---|---|---|
| Estouradas | `pgGetFiltrado()` | **Sim** |
| No Prazo | `pgGetFiltrado()` | **Sim** |
| Total filtrado / nº de equipes | `pgGetFiltrado()` | **Sim** |
| Gráficos (tipo, centro de serviço), tabela por equipe | `pgGetFiltrado()` | **Sim** |
| **Vence Hoje** | `pgGetFiltradoTodos()` | **Não — olha todos os estados** |

Motivo: contar `programáveis`/`A Reprogramar` em Estouradas/No Prazo/
equipes inflava os números com backlog sem equipe e deixava o painel
"irreal" (chegou a mostrar 1.142 estouradas assim, quase 8x mais que as
145 reais). Mas excluir esses estados do "Vence Hoje" corria o risco de
esconder um serviço que uma equipe deixou cair pra reprogramação e que
vence justamente hoje — por isso esse card específico ignora o filtro de
estado.

---

## 4. Tipo de Equipe (Carro x Moto)

Baseado no **Código Equipe**, função `getTipoEquipe(cod)`:

```js
function getTipoEquipe(cod) {
  if (!cod) return '';
  if (/-LB-/.test(cod)) return 'LB'; // Carro (Leve Básico?)
  if (/-MM-/.test(cod)) return 'MM'; // Moto
  return '';
}
```

- **LB** → 🚗 Carro
- **MM** → 🏍 Moto
- Usado no filtro "Tipo de Equipe" da aba Gerador, no ícone ao lado do
  código da equipe (tabela e ranking de equipes críticas), e nas regras
  de supervisor da Leste (RSL-moto → Valdete, RSL-carro → Wilder).

---

## 5. Região → Centro de Serviço → Supervisor

Visão resumida (detalhes completos em `DOC_centros_de_servico.md`):

### Sul
| Centro | Sigla(s) | Supervisor | Como é ligado |
|---|---|---|---|
| Juazeiro do Norte | JZ | Cirllek Nepomuceno, Cicera Maria, Jorge Luis, Amanda Vasques, Helenilton Silva (varia por código) | `SUPERVISOR_MAP` — código exato de equipe |
| Milagres | ML | Helenilton Silva | `SUPERVISOR_MAP` |
| Campos Sales | CP | Cirllek Nepomuceno | `SUPERVISOR_MAP` |

### Leste
| Centro | Sigla(s) | Supervisor | Como é ligado |
|---|---|---|---|
| Russas | RSL (moto, `-MM-`) | Valdete | regra por prefixo/tipo |
| Russas | RSL (carro, `-LB-`) | Wilder | regra por prefixo/tipo |
| Aracati | ART | Airton | regra por prefixo |
| Jaguaribe | JGB | Mailton | regra por prefixo |

### Centro Sul
| Centro | Sigla(s) | Supervisor | Como é ligado |
|---|---|---|---|
| Ico | ICO | Barbara | regra por prefixo |
| Iguatu | IGT | Fernando | regra por prefixo |
| Taua | TAA | Riverson | regra por prefixo |
| Senador (Senador Pompeu) | TAM / TAS / SNP | Maycon | regra por prefixo |

⚠️ Sul tem mapeamento **exato por código de equipe** (mais preciso);
Leste e Centro Sul usam **regra por prefixo/tipo** (mais simples, mas
assume que não há exceção dentro do mesmo prefixo).

---

## 6. Onde cada regra é aplicada na tela

| Tela | REG/NREG | Estado TdC | Prioridade | Tipo Equipe | Centro Serviço | Supervisor |
|---|---|---|---|---|---|---|
| Aba Gerador (geração de capas) | ✅ filtro | ✅ exclui programáveis/a reprogramar | ✅ stats do topo | ✅ filtro | ✅ filtro (`regiaoAtual`) | ✅ filtro |
| Painel Gerencial — KPIs gerais | ✅ filtro | ✅ exclui (exceto Vence Hoje) | ✅ 3 cards | — | ✅ filtro (`pgRegiao`) | ✅ filtro |
| Painel Gerencial — Vence Hoje | ✅ filtro | ❌ **não exclui nada** | ✅ | — | ✅ | ✅ |
| Painel Gerencial — tabela por equipe / equipes críticas | ✅ | ✅ exclui | ✅ | ✅ ícone | — | ✅ |

---

## Histórico

- **04/07/2026** — Criadas as configs `CS_POR_REGIAO` e
  `SUPERVISORES_POR_REGIAO` (documentado em detalhe em
  `DOC_centros_de_servico.md`).
- **05/07/2026** — Ajustada a regra de Estado TdC no Painel Gerencial:
  Estouradas/No Prazo/equipes voltaram a excluir programáveis/a
  reprogramar (como sempre foi); só o card "Vence Hoje" passou a olhar
  todos os estados, pra não perder nada que vença no dia.
- **05/07/2026** — Criado este documento consolidando REG/NREG, Estado
  TdC, Prioridade e Tipo de Equipe, que antes só existiam no código sem
  registro central.
- **05/07/2026** — Adicionada a regra de fim de semana: no sábado, a
  janela de "Vence Hoje" se estende até 09:00 de segunda-feira (cobre
  sábado + domingo), já que não tem equipe trabalhando domingo.
