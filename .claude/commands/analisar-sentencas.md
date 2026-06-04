---
name: analisar-sentencas
description: >
  Analisa as sentenças judiciais do arquivo "Follow Up Diario" mais recente na pasta
  automacoes/Análise de sentenças/, filtra os eventos relevantes, exclui embargos de
  declaração, extrai dados jurídicos via Claude API e gera MATRIZ_SENTENCAS_YYYY-MM-DD.xlsx.
  Use quando o usuário disser "analise as sentenças", "gera a matriz de sentenças",
  "roda a análise de hoje", "processa as sentenças" ou invocar /analisar-sentencas.
---

# /analisar-sentencas — Matriz Diária de Sentenças

## O que fazer

Orquestrar a análise diária de sentenças judiciais da Light S.A. e gerar a planilha MATRIZ_SENTENCAS_YYYY-MM-DD.xlsx.

## Passo 1: Verificar arquivo de entrada

Confirmar que existe exatamente um arquivo "Follow Up Diario*.xlsx" na pasta `automacoes/Análise de sentenças/` que **não** começa com `~$` (arquivos temporários do Excel devem ser ignorados).

Se não houver arquivo:
> "Nenhum arquivo 'Follow Up Diario*.xlsx' encontrado na pasta. Deposite o arquivo exportado do Autojur e tente novamente."

## Passo 2: Verificar API key

Testar se `ANTHROPIC_API_KEY` está disponível no ambiente executando:

```
python -c "import anthropic; anthropic.Anthropic().models.list()" 2>&1
```

## Passo 3a: Execução automática (API key disponível)

Executar o script principal:

```
cd "automacoes/Análise de sentenças" && python gerar_matriz_sentencas.py
```

Aguardar conclusão e reportar:
- Caminho completo do arquivo Excel gerado
- Número de linhas de dados processadas
- Data situação usada (DD/MM/YYYY)
- Quantas sentenças foram excluídas como embargos de declaração

## Passo 3b: Fallback — análise direta na conversa (API key indisponível)

Se a API key não estiver configurada, executar o seguinte fluxo alternativo:

1. Avisar o usuário que a `ANTHROPIC_API_KEY` não está no ambiente e que a análise será feita diretamente na conversa.
2. Rodar um script auxiliar para extrair as linhas filtradas e não-embargo:

```python
import pandas as pd, glob, os, re, json

PASTA = r"automacoes/Análise de sentenças"
EVENTOS = {
    "Sentença - Extinção sem julgamento de Mérito",
    "Sentença - Improcedente",
    "Sentença - Procedente em Parte",
}
PADROES_EMBARGO = [
    r"trata-se de embargos de declara",
    r"julgamento dos embargos",
    r"recebo os embargos de declara",
    r"rejeito os embargos de declara",
    r"conhece-se dos embargos de declara",
    r"conhe[çc]o dos embargos",
    r"embargos de declara[çc][ãa]o opostos",
    r"nego provimento aos embargos de declara",
]

matches = [m for m in glob.glob(os.path.join(PASTA, "Follow Up Diario*.xlsx"))
           if not os.path.basename(m).startswith("~$")]
arq = sorted(matches)[-1]
df = pd.read_excel(arq)
df = df[df["Nome > Evento"].isin(EVENTOS)].copy()

def nao_embargo(row):
    texto = str(row["Conteúdo"]).lower()
    return not any(re.search(p, texto) for p in PADROES_EMBARGO)

df = df[df.apply(nao_embargo, axis=1)].reset_index(drop=True)

for _, linha in df.iterrows():
    cnj = str(linha.get("Número CNJ > Pasta > Pasta Tarefa", "")).strip()
    desc = str(linha.get("Descrição dos Objetos > Pasta > Pasta Tarefa", "")).strip()
    conteudo = str(linha.get("Conteúdo", "")).strip()
    print(json.dumps({"cnj": cnj, "descricao": desc, "conteudo": conteudo}, ensure_ascii=False))
```

3. Para cada sentença retornada, analisar diretamente na conversa extraindo:
   - **dano_moral**: valor em centavos (inteiro) ou `"000"` se improcedente
   - **dano_material**: valor em centavos (inteiro) ou `"0"` se improcedente/ilíquido
   - **situacao**: exatamente uma de: `Acordo Após Julgamento` / `Acordo Judicial` / `Extinto S/ Julg. Mérito – Demais Hipóteses` / `Improcedente` / `Procedente`
   - **juiz**: nome completo ao final da sentença ou `"Não identificado"`
   - **dias_sem_energia**: quantidade de dias de interrupção mencionada na fundamentação, **somente se** a coluna "Descrição dos Objetos" contiver a palavra "Interrupção"; caso contrário, `""`

4. Ao concluir a análise de todas as sentenças, atualizar `gerar_matriz_direto.py` com os novos dados no array `DADOS` e executar:

```
cd "automacoes/Análise de sentenças" && python gerar_matriz_direto.py
```

## Regras que devem ser respeitadas em qualquer caminho

- **Coluna H (Dias Sem Energia):** preencher **somente** se "Descrição dos Objetos" contém "Interrupção" (case-insensitive). Em todos os outros casos, valor = `""`.
- **Coluna I (Matéria):** valor exato de "Descrição dos Objetos > Pasta > Pasta Tarefa".
- **DATA_SITUACAO:** hoje − 2 dias, formato DD/MM/YYYY na planilha.
- **Embargos:** qualquer sentença que trate de julgamento de embargos de declaração deve ser excluída, mesmo que o trecho identificador apareça após um cabeçalho DJE longo.
