# Skill: Dashboard de Prazos Pendentes

## Quando usar
Ativar sempre que o usuário pedir o dashboard de prazos pendentes, semáforo de prazos, relatório de atraso por advogado, ou variações como "gera o dashboard", "atualiza os prazos", "quero o painel de prazos".

## O que esta skill faz
Gera um dashboard HTML executivo e interativo com:
- 7 KPIs (total, preclusivos, regulares, taxa de risco, maior volume, situação crítica, atraso médio)
- Tabela por advogado com semáforo de risco (🔴 Crítico / 🟡 Médio / 🟢 Baixo)
- Painel de prazos preclusivos com base legal e impacto processual
- Gráfico de barras empilhadas (preclusivos × regulares por advogado)

## Arquivos envolvidos
| Arquivo | Papel |
|---|---|
| `ccos-joel/automacoes/prazos pendentes dashboard/Prazos pendentes*.xlsx` | Fonte de dados (detectada automaticamente) |
| `ccos-joel/automacoes/prazos pendentes dashboard/gerar_dashboard.py` | Script gerador |
| `ccos-joel/automacoes/prazos pendentes dashboard/dashboard_prazos.html` | Output gerado |

## Estrutura do xlsx (colunas fixas, dados mudam a cada rodada)
| Coluna | Conteúdo |
|---|---|
| `Prazo Interno` | Data do prazo (datetime) |
| `Número CNJ > Pasta > Pasta Tarefa` | Número do processo (não usado no dashboard) |
| `Nome > Responsável` | Advogado responsável |
| `Nome > Evento` | Tipo de prazo/evento |

## Prazos classificados como PRECLUSIVOS (hardcoded no script)
- CONTESTAÇÃO LIGHT → Art. 335 CPC → Revelia
- CONTRARRAZÕES EM APELAÇÃO → Art. 1.010 CPC → Trânsito em Julgado
- CONTRARRAZÕES EM ED → Art. 1.023 CPC → Trânsito em Julgado
- IMPUGNAÇÃO AO CUMPRIMENTO DE SENTENÇA → Art. 525 CPC → Penhora sem defesa
- IMPUGNAR LAUDO / JUNTAR PARECER - NCPC → Art. 477 CPC → Preclusão do laudo
- MEMORIAIS DE 2ª INSTÂNCIA → Art. 364 CPC → Infração normas processuais
- AGRAVO DE INSTRUMENTO → Art. 1.015 CPC → Preclusão do recurso
- EMBARGOS DE DECLARAÇÃO - NCPC → Art. 1.022 CPC → Preclusão
- IMPUG ART. 854 NCPC → Art. 854 CPC → Penhora on-line sem defesa

## Lógica de semáforo de risco
- 🔴 CRÍTICO: ≥ 3 preclusivos OU atraso médio > 30 dias
- 🟡 MÉDIO: 1–2 preclusivos OU atraso médio entre 15–30 dias
- 🟢 BAIXO: 0 preclusivos E atraso médio < 15 dias

## Passos de execução

### 1. Verificar xlsx disponível
Checar se há arquivo `Prazos pendentes*.xlsx` em `ccos-joel/automacoes/prazos pendentes dashboard/`.
O script detecta automaticamente o mais recente por data de modificação.
- Se não houver arquivo: avisar o usuário e pedir que coloque o xlsx na própria pasta `prazos pendentes dashboard/`.
- Se houver: prosseguir.

### 2. Executar o script
```
python "ccos-joel/automacoes/prazos pendentes dashboard/gerar_dashboard.py"
```
Executar da raiz do projeto (`ccos-joel/`). O script resolve todos os caminhos relativos ao próprio `__file__`, então funciona de qualquer diretório.

### 3. Abrir o dashboard
Abrir `ccos-joel/automacoes/prazos pendentes dashboard/dashboard_prazos.html` no navegador padrão.

### 4. Reportar ao usuário
Informar:
- Nome do arquivo xlsx detectado e usado
- Data-base utilizada (data de hoje)
- Totais: prazos, preclusivos, regulares, taxa de risco
- Advogado com maior volume e advogado em situação crítica
- Caminho do HTML gerado

## Quando o xlsx muda (nova rodada periódica)
O usuário simplesmente coloca o novo arquivo em `ccos-joel/automacoes/prazos pendentes dashboard/` e pede o dashboard novamente.
O script sempre pega o xlsx mais recente por data de modificação — não é necessário alterar nada no script.

## Quando adicionar um novo tipo de prazo preclusivo
Editar o dicionário `PRECLUSIVOS_MAP` em `gerar_dashboard.py`, adicionando a chave (nome exato do evento em MAIÚSCULAS, como aparece na coluna `Nome > Evento`) e os campos `label`, `base` e `impacto`.

## Possíveis erros e soluções
| Erro | Causa provável | Solução |
|---|---|---|
| `FileNotFoundError` | Nenhum xlsx em `dados/` | Colocar o arquivo exportado do Autojur em `ccos-joel/dados/` |
| Dashboard abre vazio | JavaScript quebrado | Verificar se há placeholders inválidos no HTML (não deveria ocorrer nesta versão) |
| Advogado aparece com nome estranho | Espaços extras no xlsx | O script normaliza com `.strip()` — verificar se o nome tem caracteres especiais |
| Prazo preclusivo não classificado | Evento com nome diferente do esperado | Adicionar o novo nome ao `PRECLUSIVOS_MAP` no script |
