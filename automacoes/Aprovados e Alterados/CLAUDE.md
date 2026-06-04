# Aprovados e Alterados — Light S.A.

## Propósito

Automação diária de extração e aprovação em lote de prazos processuais da Light S.A. no sistema Autojur. Exporta planilha com tarefas pendentes de aprovação e executa a conferência em massa via Painel de Tarefas.

## Fluxo diário

1. Acessar o Autojur via Chrome (MCP Playwright ou Claude in Chrome)
2. Invocar `/aprovados-em-lote`
3. A skill exporta a planilha e salva nesta pasta como `Aprovados em lote - DD.MM.AAAA.xlsx`
4. Com permissão do usuário, executa a aprovação em massa das tarefas filtradas

## Arquivos gerados

- **`Aprovados em lote - DD.MM.AAAA.xlsx`** — planilha exportada do Autojur com as tarefas a aprovar, usando o modelo "Follow-up Diário_Ampla" (Usuário: MAURA DE OLIVEIRA FARIAS NEVES)

## Credenciais Autojur

- **URL:** `https://taunayadv.autojur.com.br/login.jsf?url=%2Fsistema%2Ftarefa%2Ftarefa.jsf`
- **Usuário:** joel.costa
- **Senha:** Joel@123

## Filtros do Painel de Tarefas

| Campo | Valor |
|-------|-------|
| Primeiro dropdown | Grupo Empresarial Principal |
| Contém | Light |
| Datas | Vazias (limpar ambos os campos) |
| Tarefas de Todos | Marcado |
| Visão | Tarefas a Aprovar |

## Modelo de exportação

**Follow-up Diário_Ampla** (Público: Sim - Usuário: MAURA DE OLIVEIRA FARIAS NEVES - Limitação Perfil: Não)

## Aprovação em massa (Parte 2)

Após exportar a planilha, abrir "Conferir em Massa" (ícone de olho na barra de ferramentas) com:
- **Conferência:** Aprovar
- **Motivo:** `APROVADOS EM LOTE NO DIA DD/MM/AAAA`
- **Aplicar Sobre:** Filtrados
- **SEMPRE pedir permissão ao usuário antes de clicar em Salvar**

## Regras

- Sempre pedir permissão antes de aprovar em lote
- Nomear a planilha com a data do dia: `Aprovados em lote - DD.MM.AAAA.xlsx`
- Não apagar planilhas anteriores
- Se a sessão do Autojur expirar, refazer o login
- O modal de exportação permanece aberto após o clique em Exportar; fechar manualmente clicando no X
