# Análise de Sentenças — Light S.A.

## Propósito

Automação diária de análise de sentenças judiciais para construção da matriz de decisões da Light S.A. Filtra o relatório do Autojur, exclui embargos de declaração, extrai dados jurídicos via Claude API e gera um Excel estruturado para controle processual.

## Fluxo diário

1. Nova planilha **"Follow Up Diario*.xlsx"** é depositada nesta pasta (exportada do Autojur)
2. Invocar `/analisar-sentencas`
3. O script processa os dados e gera **MATRIZ_SENTENCAS_YYYY-MM-DD.xlsx** na mesma pasta

## Scripts

- **`gerar_matriz_sentencas.py`** — script canônico. Lê o Follow Up Diario mais recente, filtra eventos, exclui embargos, chama Claude API por sentença e gera o Excel com 9 colunas. **Usar sempre este para novas execuções.**
- **`gerar_matriz_direto.py`** — auxiliar. Gera o Excel a partir de dados hardcoded. Usar apenas quando a API não estiver disponível ou para correções pontuais de linhas já analisadas.

## Colunas da matriz (A–I)

| Col | Nome |
|-----|------|
| A | Número do processo |
| B | Fase atual |
| C | 1ª Inst. Dano Moral |
| D | 1ª Inst. Dano Material |
| E | Situação 1ª instância |
| F | Data situação 1ª inst. |
| G | Juiz 1ª instância |
| H | Quantidade Dias Sem Energia |
| I | Matéria |

## Regras críticas de negócio

1. **Coluna H (Dias Sem Energia):** só pode ser preenchida se a coluna "Descrição dos Objetos > Pasta > Pasta Tarefa" do arquivo de entrada contiver a palavra **"Interrupção"**. Para todos os demais casos, manter em branco — independentemente do conteúdo da sentença.
2. **Dano moral / dano material:** valor em centavos, inteiro sem pontos ou vírgulas (ex: R$ 3.000,00 → `300000`). Retornar `"000"` para improcedente em dano moral; `"0"` para improcedente ou ilíquido em dano material.
3. **Embargos de declaração:** excluídos via busca full-text no campo "Conteúdo" (não apenas no cabeçalho).
4. **DATA_SITUACAO:** sempre hoje − 2 dias.
5. **Coluna I (Matéria):** valor exato de "Descrição dos Objetos > Pasta > Pasta Tarefa" do arquivo de entrada.

## Eventos filtrados do Follow Up Diario

- `Sentença - Extinção sem julgamento de Mérito`
- `Sentença - Improcedente`
- `Sentença - Procedente em Parte`

## Pré-requisito para execução automática

A variável de ambiente `ANTHROPIC_API_KEY` deve estar definida no shell antes de executar o script. Sem ela, usar o fallback de análise direta na conversa via `/analisar-sentencas`.
