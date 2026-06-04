---
name: aprovados-em-lote
description: >
  Extrai planilha de prazos a aprovar da Light no Autojur e executa aprovação em massa.
  Acessa o Painel de Tarefas, filtra por Grupo Empresarial Principal "Light" + Tarefas a Aprovar,
  exporta via modelo Follow-up Diário_Ampla e salva em automacoes/Aprovados e Alterados/.
  Após exportação, abre Conferir em Massa para aprovar com motivo padronizado.
  Use quando o usuário disser "aprovar em lote", "aprovados em lote", "extrair planilha de aprovação",
  "exportar aprovados", "aprovar prazos", "aprovação em massa", "rodar aprovação".
---

# /aprovados-em-lote — Extração e aprovação em lote de prazos Light

## Contexto

Processo diário do escritório Taunay Advogados. Prazos processuais da Light S.A. precisam ser exportados do Autojur e aprovados em lote no Painel de Tarefas.

## Pré-requisitos

- Chrome com extensão Claude in Chrome conectada OU MCP Playwright disponível
- Acesso ao Autojur (credenciais em CLAUDE.md desta pasta)

## Fluxo

### PASSO 1 — Acessar o Autojur

Navegar para:
```
https://taunayadv.autojur.com.br/login.jsf?url=%2Fsistema%2Ftarefa%2Ftarefa.jsf
```

Se não estiver logado, usar as credenciais:
- Usuário: `joel.costa`
- Senha: `Joel@123`
- Clicar em "ENTRAR"

Se já estiver logado, prosseguir direto.

### PASSO 2 — Configurar filtros no Painel de Tarefas

Executar na seguinte ordem:

1. **Primeiro dropdown** (padrão "Evento"): mudar para `Grupo Empresarial Principal`
2. **Campo de texto** ao lado de "Contém": digitar `Light`
3. **Checkbox "Tarefas de Todos"**: marcar
4. **Campos de data**: limpar ambos (usar JavaScript se necessário para limpar os inputs `data-desde_input` e `data-ate_input`)
5. **Dropdown de visão** (padrão "Minhas Tarefas"): mudar para `Tarefas a Aprovar`
6. **Clicar na lupa azul** para pesquisar

### PASSO 3 — Exportar planilha

1. Clicar no ícone verde **"Exportar para Planilha"** (na barra de ícones coloridos à direita, tooltip "Exportar para Planilha")
2. No modal "Exportar para Planilha", selecionar o modelo:
   `Follow-up Diário_Ampla (Público: Sim - Usuário: MAURA DE OLIVEIRA FARIAS NEVES - Limitação Perfil: Não)`
3. Clicar no botão **"Exportar"**
   - **ATENÇÃO:** O modal NÃO fecha automaticamente após a exportação. Verificar a pasta Downloads para confirmar que o arquivo foi gerado.
   - O botão de exportar tem o ID: `j_idt634:form-btn-exportar:btn-exportar-planilha`
4. Fechar o modal clicando no **X**

### PASSO 4 — Salvar planilha na pasta correta

1. Localizar o arquivo mais recente `Follow-up Diario_Ampla-*.xlsx` na pasta Downloads
2. Copiar para `automacoes/Aprovados e Alterados/`
3. Renomear para: `Aprovados em lote - DD.MM.AAAA.xlsx` (data do dia)

### PASSO 5 — Aprovar em massa

1. No Painel de Tarefas (com os filtros ainda ativos), clicar no ícone **"Conferir em Massa"** (ícone de olho na barra de ferramentas azul)
   - Localizar via: `find("conferir em massa eye icon button")`
2. No modal "Conferir em Massa", configurar:
   - **Conferência:** `Aprovar` (radio button)
   - **Motivo:** `APROVADOS EM LOTE NO DIA DD/MM/AAAA`
   - **Aplicar Sobre:** `Filtrados` (radio button)
3. **PARAR e pedir permissão ao usuário antes de clicar em "Salvar"**

> "O modal está preenchido com [N] tarefas filtradas para aprovar em lote. Posso clicar em Salvar?"

4. Se autorizado, clicar em **"Salvar"**
5. Se não autorizado, clicar em **"Fechar"**

### PASSO 6 — Reportar resultado

> **Planilha exportada:** `Aprovados em lote - DD.MM.AAAA.xlsx`
>
> - Total de registros: [N]
> - Aprovação em lote: [executada/não executada a pedido do usuário]

## Dicas técnicas

- Os campos de data do Autojur usam datepicker; para limpá-los com segurança, usar JavaScript:
  ```javascript
  document.getElementById('form-pesquisa-tarefa:data-desde_input').value = '';
  document.getElementById('form-pesquisa-tarefa:data-ate_input').value = '';
  ```
- O botão "Exportar" submete um formulário PrimeFaces. Se o clique visual não funcionar, usar:
  ```javascript
  document.getElementById('j_idt634:form-btn-exportar:btn-exportar-planilha').click();
  ```
- O modal de exportação permanece aberto após o download; sempre fechar manualmente

## Regras

- SEMPRE pedir permissão antes de aprovar em lote (Passo 5)
- Não apagar planilhas de dias anteriores
- Se a sessão expirar, refazer login pelo Passo 1
- Se o modal de exportação não fechar, verificar Downloads e fechar manualmente
