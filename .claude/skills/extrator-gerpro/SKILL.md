# Skill: Extrator de Relatório Gerpro

## Gatilho
Ativado quando o usuário diz `/extrator-gerpro`, "extrair relatório do Gerpro", "relatório de encerramentos Gerpro" ou variações.

## O que faz
Automatiza a extração do Relatório de Encerramentos do sistema Gerpro usando o MCP Playwright (`@playwright/mcp`). O Claude controla o navegador passo a passo, marca todos os checkboxes necessários na Pesquisa Avançada e exporta o relatório para a pasta destino.

## Pasta destino
```
c:\Users\Joel\Desktop\Claude Code - Pastas\Claude Code OS\ccos-joel\automacoes\Relatório de Encerramentos Gerpro\
```

---

## Execução

### PASSO 0 — Coletar datas do usuário
Antes de abrir o browser, perguntar ao usuário:
- **Data DE** (início do período — formato dd/mm/aaaa)
- **Data ATÉ** (fim do período — formato dd/mm/aaaa)

Guardar os valores para usar no PASSO 3, campo "Data da Fase".

---

### PASSO 1 — Login no Gerpro
1. Navegar para: `https://gerpro.com.br/`
2. Localizar o campo de login e preencher: `JOELSOUZA.LIGHT`
3. Localizar o campo de senha e preencher: `JOE123`
4. Submeter o formulário (clicar no botão de login ou pressionar Enter)
5. Aguardar carregamento completo da tela inicial

---

### PASSO 2 — Navegar até Pesquisa Avançada
1. Localizar o menu **"Pesquisa"** no canto superior esquerdo da tela
2. Clicar em "Pesquisa"
3. Aguardar carregamento completo da tela de **Pesquisa Avançada**

---

### PASSO 3 — Configurar os filtros

> **Regra geral:** Antes de clicar em qualquer checkbox, ler o estado atual da página. Marcar apenas os campos que ainda não estiverem selecionados. **Nunca desmarcar** campos que já estejam ativos.

#### 3.1 — Seção principal (topo da página)

Marcar os seguintes checkboxes:
- [ ] **Processo**
- [ ] **Número na Justiça**
- [ ] **Estado do Processo** → após marcar, selecionar **"Ativo"** no dropdown ao lado
- [ ] **Data Cadastramento**
- [ ] **Comarca**
- [ ] **Número**
- [ ] **Vara**
- [ ] **Classe**

*Fazer scroll para baixo para visualizar os próximos campos.*

#### 3.2 — Campo "Matéria"
- [ ] **Matéria** → após marcar, selecionar **"Cível"** no dropdown ao lado
- ⚠️ **Esta ação recarrega a página. Aguardar o reload completo antes de continuar.**

#### 3.3 — Campo "Fase"
- [ ] **Fase** → marcar o checkbox
- Na lista de fases exibida abaixo, localizar e clicar sobre o item **"Encerramento"** (deve ficar destacado/selecionado)

#### 3.3.1 — Campo "Data da Fase"
- [ ] **Data da Fase** → marcar o checkbox
- No campo **"de"**, inserir a data coletada no PASSO 0
- No campo **"até"**, inserir a data coletada no PASSO 0

#### 3.4 — Campo "Motivo da Ação"
- [ ] **Motivo da Ação**

*Fazer scroll para baixo.*

#### 3.5 — Seção "Origem do Processo"
- [ ] **Motivo Geral**

*Fazer scroll para baixo.*

#### 3.6 — Seção "Específicos Perícia"
- [ ] **Laudo Técnico**
- [ ] **Prova Pericial**
- [ ] **Resultado Técnico**
- [ ] **Requerida pela Light?**
- [ ] **Tipo de Perícia**
- [ ] **Nome do Perito**
- [ ] **Honorários Homologados**
- [ ] **Data da Homologação**
- [ ] **Solicitante Prova Pericial**
- [ ] **Resultado Perícia**

*Fazer scroll para baixo.*

#### 3.7 — Seção "Resultado"
- [ ] **Dano Material Apurado**
- [ ] **Data Dano Material**
- [ ] **Recurso Interposto**
- [ ] **Valor Total Pedido**
- [ ] **1ª Inst. Dano Moral**
- [ ] **1ª Inst. Dano Material**
- [ ] **Situação 1ª Instância**
- [ ] **Data Situação 1ª Instância**
- [ ] **2ª Inst. Dano Moral**
- [ ] **2ª Inst. Dano Material**
- [ ] **Situação 2ª Instância**
- [ ] **Data Situação 2ª Instância**
- [ ] **Camara Civel/Turma Recursal**
- [ ] **Juiz 1ª Instância**

*Fazer scroll para baixo.*

#### 3.8 — Seção "Avaliação do Processo"
- [ ] **Valor Pago (R$)**
- [ ] **Situação**
- [ ] **Quantidade de Dias Sem Energia**
- [ ] **Justificativa**

---

### PASSO 4 — Executar a pesquisa
1. Rolar a página até o final
2. Localizar o botão **"Pesquisar"** no rodapé
3. Clicar em **"Pesquisar"**
4. Aguardar o carregamento completo dos resultados antes de prosseguir

---

### PASSO 5 — Exportar o relatório
1. Na tela de resultados, localizar o botão **"Exportar"** ou **"Excel"**
2. Clicar no botão de exportação
3. Aguardar o download completar na pasta Downloads (`C:\Users\Joel\Downloads\`)

---

### PASSO 6 — Mover o arquivo para a pasta destino
1. Identificar o arquivo mais recente na pasta `C:\Users\Joel\Downloads\` (o relatório gerado pelo Gerpro)
2. Usar PowerShell para movê-lo para a pasta destino:
   ```powershell
   $origem = Get-ChildItem "C:\Users\Joel\Downloads\" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
   $destino = "c:\Users\Joel\Desktop\Claude Code - Pastas\Claude Code OS\ccos-joel\automacoes\Relatório de Encerramentos Gerpro\"
   Move-Item $origem.FullName $destino
   ```
3. Confirmar ao usuário o nome e caminho completo do arquivo salvo

---

## Confirmação final

Ao concluir, informar ao usuário:
```
Relatório extraído com sucesso.
Arquivo: <nome do arquivo>
Salvo em: ccos-joel\automacoes\Relatório de Encerramentos Gerpro\
```
