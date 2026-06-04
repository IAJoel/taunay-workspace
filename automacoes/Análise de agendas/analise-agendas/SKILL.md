---
name: analise-agendas
description: >
  Gera a planilha MATRIZ diária cruzando os três arquivos da pasta "Análise de agendas":
  Aprovados em lote, Base GED e Análise das petições. Executa filtros, merges, análise
  jurídica de equivalência e identificação de inconsistências. Salva MATRIZ_YYYY-MM-DD.xlsx.
  Use quando o usuário disser "analise as agendas", "gerar matriz", "análise de agendas",
  "rodar análise", "análise de agenda", "realize a análise das agendas".
---

# /analise-agendas — Geração da MATRIZ diária

## Contexto

Processo diário do escritório Taunay Advogados. Três planilhas chegam na pasta `Análise de agendas/`:
- `Aprovados em lote*.xlsx` — tarefas aprovadas (eventos, prazos, responsáveis)
- `Base GED*.xlsx` — petições protocoladas no GED
- `Análise das petições*.xlsx` — análise de protocolos e teor das petições

## Fluxo

### PASSO 1 — Verificar arquivos de entrada

Confirmar que os três arquivos existem em `Análise de agendas/`. Se faltar algum:
> "O arquivo [nome] não está na pasta 'Análise de agendas'. Coloque o arquivo e rode novamente."

### PASSO 2 — Gerar MATRIZ

Rodar a partir do diretório `ccos-ratos/`:

```bash
python "Análise de agendas/gerar_matriz.py"
```

O script vai:
- Ler os três arquivos (acumulando todos os históricos de GED e Petições)
- Filtrar responsável RAMON VINICIUS BAPTISTA DA SILVA
- Ordenar por Prazo Interno (ascendente)
- Fazer merges via número CNJ (`Número que está na petição`)
- Salvar `MATRIZ_YYYY-MM-DD.xlsx` com Equivalência = "ANALISAR INDIVIDUALMENTE" para todos
- Salvar `pares_analise_YYYY-MM-DD.json` com os pares que têm GED e Teor preenchidos

### PASSO 3 — Analisar equivalências

Ler `Análise de agendas/pares_analise_YYYY-MM-DD.json`.

Se o JSON estiver vazio (0 pares), pular este passo e reportar que não houve correspondências GED + petição.

Para cada par, atuar como advogado experiente em Processo Civil e determinar:

**O conteúdo da petição (campo "teor") equivale ao evento processual registrado (campo "evento")?**

O campo "ged" (tipo de documento no GED) é contexto adicional — confirma o que foi efetivamente protocolado.

---

#### Regra geral

- **SIM**: o teor da petição corresponde ao tipo de ato processual descrito no evento
- **NÃO**: o teor descreve um ato processual distinto do evento registrado

#### Padrões de equivalência conhecidos

**COMPROVANTE PAGTO no GED:**
- Teor sobre pagamento, depósito, obrigação de pagar/fazer + evento PETICIONAR PAGAMENTO / PETICIONAR OBRIG DE FAZER / PETICIONAR PGTO HON PERICIAIS → **SIM**
- Teor sobre pagamento + evento IMPUGNAÇÃO AO CUMPRIMENTO DE SENTENÇA → **NÃO** (pagamento ≠ impugnação)

**CONTESTAÇÃO CÍVEL no GED:**
- Qualquer variante de contestação + evento CONTESTAÇÃO LIGHT → **SIM**
- Qualquer variante de contestação + evento que NÃO seja contestação (CUMPRIMENTO LIMINAR, AGRAVO DE INSTRUMENTO, PETIÇÃO 1018, VERIFICAR CUMPRIMENTO LIMINAR, etc.) → **NÃO**

**PETIÇÃO - PERÍCIA – QUESITOS no GED:**
- Teor sobre quesitos periciais, juntada de quesitos, justificação de provas + evento QUESITOS / ASSISTENTE → **SIM**
- Teor sobre quesitos + evento COMUNICAR PERÍCIA → **NÃO** (juntar quesitos ≠ comunicar perícia)

**PETIÇÃO INTERCORRENTE no GED:**
- Manifestação de cumprimento de tutela/liminar + evento CUMPRIMENTO LIMINAR → **SIM**
- Manifestação genérica (sobre depósitos, habilitação, documentos, petição do perito, laudo, etc.) + evento MANIFESTAÇÃO - NCPC (qualquer variante) → **SIM**
- Manifestação sobre esclarecimentos periciais, resposta ao perito, concordância com laudo + evento SOBRE PET DO PERITO → **SIM**
- Juntada de comprovante de custas + evento CUSTAS - NCPC → **SIM**
- Aditamento à contestação + evento MANIFESTAÇÃO - NCPC → **SIM**

**COMPROVAÇÃO DE RECOLHIMENTO DE CUSTAS no GED:**
- Qualquer teor sobre custas/recolhimento + eventos CUSTAS - NCPC / MANIFESTAÇÃO - CPC73 / ACOMPANHAR / DISPENSA / GRERJ AGRAVO DE INSTRUMENTO → **SIM**

**EMBARGOS À EXECUÇÃO no GED:**
- Impugnação ao cumprimento de sentença (qualquer variante) + evento IMPUGNAÇÃO AO CUMPRIMENTO DE SENTENÇA → **SIM**

**IMPUGNAÇÃO AO LAUDO PERICIAL no GED:**
- Impugnação ao laudo pericial + evento IMPUGNAR LAUDO / JUNTAR PARECER - NCPC → **SIM**

**CONTRARRAZÕES DE RECURSO no GED:**
- Contrarrazões de apelação + evento CONTRARRAZÕES EM APELAÇÃO → **SIM**
- Contrarrazões aos embargos de declaração + evento CONTRARRAZÕES EM ED → **SIM**

#### Critérios para casos não listados

1. **Contestação em evento não-contestação**: sempre NÃO, independentemente do nome do evento.
2. **Obrigação híbrida (fazer e pagar)**: petição que menciona ambas as obrigações equivale a cada evento separado — SIM em ambos.
3. **Substabelecimento combinado**: classificar pelo ato principal (ex: "pedido de sessão + substabelecimento" → ato principal é o pedido de sessão).
4. **Evento SOBRE PET DO PERITO**: engloba manifestações sobre esclarecimentos periciais, resposta à impugnação de laudo, juntada de docs ao perito, concordância com laudo — SIM quando o teor indica resposta a ato do perito.
5. **Manifestação genérica**: teor que descreve "manifestação" sem especificar o ato → verificar se o evento é igualmente genérico (MANIFESTAÇÃO - NCPC); se sim, SIM.

---

Montar JSON com a chave sendo o `row_excel` e o valor "SIM" ou "NÃO":

```json
{"37": "SIM", "42": "NÃO"}
```

### PASSO 4 — Atualizar a MATRIZ

```bash
python "Análise de agendas/gerar_matriz.py" --update "{\"37\":\"SIM\",\"42\":\"NÃO\",...}"
```

O script preenche a coluna H (Equivalência) e recalcula a coluna L (Inconsistências).
Inconsistências NÃO: `Conteúdo da petição ("X") não corresponde ao evento ("Y")`.
Inconsistências de protocolo: `Número do protocolo não corresponde ao número da petição`.

### PASSO 5 — Reportar resultado

> **MATRIZ gerada:** `Análise de agendas/MATRIZ_[data].xlsx`
>
> - Linhas totais: [N]
> - Equivalência SIM: [N]
> - Equivalência NÃO: [N]
> - Sem correspondência GED/petição (ANALISAR INDIVIDUALMENTE): [N]
> - Inconsistências detectadas: [N]

Se houver NÃO > 0, perguntar:
> "Quer que eu detalhe as inconsistências encontradas?"

## Regras

- Não apagar os arquivos de entrada
- Se o script falhar, mostrar o erro completo sem tentar corrigir automaticamente
- Se o JSON de pares estiver vazio (0 pares), pular o PASSO 3 e reportar
- Executar sempre a partir do diretório `ccos-ratos/`, não de dentro de `Análise de agendas/`
