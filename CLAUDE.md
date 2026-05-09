# Taunay Advogados — Claude Code OS

## O que é esse workspace

Workspace operacional de Joel Costa para o escritório Taunay Advogados. Centraliza automações jurídicas, produção de peças processuais, relatórios Visual Law, conteúdo para redes sociais e prospecção de clientes.

**Estrutura de pastas:**
- `_contexto/` — memória persistente do negócio (empresa, preferências, estratégia)
- `clientes/` — pasta por cliente com briefing e histórico
- `pecas/` — peças processuais organizadas por tipo (contestações, recursos, memoriais)
- `automacoes/` — scripts RPA e fluxos automatizados
- `conteudo/` — conteúdo para Instagram e redes sociais
- `relatorios/` — relatórios gerenciais e Visual Law
- `dados/` — CSVs, planilhas e arquivos de entrada
- `marca/` — logo, design-guide e identidade visual
- `templates/skills/` — templates de skills prontos pra personalizar com /mapear
- `templates/ferramentas/catalogo.md` — APIs e ferramentas disponíveis pra usar em skills

## Sobre o negócio

L.A. Taunay Advogados Associados, Rio de Janeiro. Atuação jurídica sênior em Processo Civil, Direito do Consumidor e Direito de Energia Elétrica (ANEEL Resolução 1.000), com foco na defesa de concessionárias de serviços públicos, principalmente a Light S.A. Combina expertise jurídica com desenvolvimento de software para automação de litígio de alto volume.

## O que mais fazemos aqui

- Automação de processos jurídicos repetitivos (RPA) com Python e Claude
- Elaboração de peças processuais de alta performance
- Relatórios gerenciais com Visual Law
- Extração e consolidação de dados processuais
- Produção de conteúdo e prospecção via Instagram e redes sociais

## Clientes e contexto

Grandes clientes corporativos externos (concessionárias de serviços públicos) + uso interno da equipe. Equipe: Joel Costa, Leonardo Ferreira Loffler, Luciano Bogado.

## Tom de voz

Profissional, técnico, persuasivo. Precisão de mestre em Direito Processual Civil. Sem "robusta", "crucial" ou travessão "—". Sem teses ou precedentes alucinados. Argumentação baseada em lógica jurídica e semiótica. Regras completas em `_contexto/preferencias.md`.

## Ferramentas conectadas

- Autojur (gestão processual)
- Python: Pandas, Selenium, Playwright, PyAutoGUI, Pytesseract
- Gemini, ChatGPT
- *(MCPs instalados aparecerão aqui conforme forem adicionados)*

---

## Comandos disponíveis

Todos os comandos ficam em `.claude/commands/`:

- `/setup` — Wizard de onboarding que entrevista o usuário, detecta o perfil (freelancer, agência, empresa, solopreneur, criador, profissional), preenche os arquivos de contexto, cria a estrutura de pastas e recomenda MCPs conforme as ferramentas mencionadas
- `/iniciar` — Carrega o contexto do negócio no início de cada sessão e pergunta o que fazer
- `/mapear` — Descobre processos repetitivos do usuário e cria skills personalizadas ou adapta templates existentes
- `/novo-projeto` — Cria subpastas de projeto (clientes, produtos, séries de conteúdo) com CLAUDE.md próprio
- `/atualizar` — Varre o estado atual do projeto vs. o contexto documentado e propõe atualizações nos arquivos de memória
- `/syncar` — Faz commit e push das mudanças pro GitHub (também configurado como hook automático ao parar a sessão)

Skills pré-instaladas ficam em `.claude/skills/` ou `~/.claude/skills/` (globais). Templates prontos para adaptação estão em `templates/skills/`.

---

## Arquitetura

Este projeto não tem build, testes ou código compilado — é um sistema de templates em markdown que ensina o Claude a trabalhar com o contexto específico de um negócio.

**Fluxo central:**
1. `/setup` lê o perfil do usuário e escreve os arquivos de contexto (`_contexto/`) e a estrutura de pastas
2. Em cada sessão, o Claude lê `_contexto/` automaticamente e usa como base para todas as respostas
3. Skills em `.claude/skills/` ou `.claude/commands/` têm prioridade sobre comportamento genérico

**Diretórios principais:**
- `_contexto/` — Memória persistente do negócio (`empresa.md`, `preferencias.md`, `estrategia.md`)
- `.claude/commands/` — Os 6 comandos principais do sistema
- `.claude/skills/` — Skills específicas desse workspace (criadas pelo usuário via `/mapear`)
- `templates/` — Base para criar novas skills e perfis; contém `templates/skills/`, `templates/perfis/`, e `templates/ferramentas/catalogo.md`
- `marca/` — Guia de identidade visual consultado por todas as skills visuais (carrossel, proposta, slide)
- `dados/` — Drop zone para arquivos CSV/XLSX/PDF usados com `/analisar-dados`

**Dependências de runtime:**
- Git + GitHub (para `/syncar`)
- Node.js / npx (para skills de renderização visual como `/carrossel`)
- MCPs opcionais instalados durante o `/setup` conforme as ferramentas do usuário (Notion, Google Drive, Gmail, etc.)

---

<!-- Este arquivo será atualizado pelo /setup com o contexto do seu negócio. -->

## Contexto do negócio

No início de toda conversa, ler os seguintes arquivos (se existirem e estiverem configurados):

1. `_contexto/empresa.md` — quem é o usuário, o que faz, como funciona o negócio
2. `_contexto/preferencias.md` — tom de voz, estilo de escrita, o que evitar
3. `_contexto/estrategia.md` — foco atual, prioridades, o que pode esperar

Usar essas informações como base pra qualquer resposta ou decisão. Ao sugerir prioridades, formatos ou abordagens, considerar o foco atual descrito em `estrategia.md`.

Para qualquer tarefa visual (carrossel, proposta, slide, landing page), consultar `marca/design-guide.md` como referência de estilo.

Não é necessário listar o que foi lido nem confirmar a leitura. Apenas usar o contexto naturalmente.

---

## Fluxo de trabalho

Antes de executar qualquer tarefa, verificar se existe uma skill relevante em `.claude/skills/` ou `.claude/commands/`.
Se encontrar, seguir as instruções da skill.
Se não encontrar, executar a tarefa normalmente.

Ao concluir uma tarefa que não tinha skill mas parece repetível (o usuário provavelmente vai pedir de novo no futuro), perguntar:

> "Isso pode virar uma skill pra próxima vez. Quer que eu crie?"

Não perguntar pra tarefas pontuais ou perguntas simples. Só quando o padrão de repetição for claro.

---

## Aprender com correções

Quando o usuário corrigir algo, melhorar uma resposta ou dar uma instrução que parece permanente (frases como "na verdade é assim", "não faça mais isso", "prefiro assim", "sempre que...", "evita...", "da próxima vez..."), perguntar:

> "Quer que eu salve isso pra não precisar repetir?"

Se sim, identificar onde faz mais sentido salvar:

- **Sobre o negócio** (quem são os clientes, como funciona a empresa, serviços, mercado) → adicionar em `_contexto/empresa.md`
- **Sobre preferências e estilo** (tom de voz, formato de resposta, o que evitar, como estruturar textos) → adicionar em `_contexto/preferencias.md`
- **Sobre prioridades e foco atual** (projetos em andamento, metas do momento, prazos importantes, o que é prioridade agora) → adicionar em `_contexto/estrategia.md`
- **Regra de comportamento nessa pasta** (onde salvar arquivos, como nomear, fluxos específicos) → adicionar no próprio `CLAUDE.md`

Salvar com uma linha nova clara, sem reformatar o arquivo inteiro. Confirmar o que foi salvo mostrando a linha adicionada.

Não perguntar se a correção for óbvia de contexto imediato (ex: "na verdade o arquivo se chama X"). Só perguntar quando a informação tiver valor duradouro.

---

## Manter contexto atualizado

Ao terminar uma tarefa que mudou algo relevante no projeto (novo cliente, nova skill, mudança de foco, novo processo, ferramenta instalada, estrutura de pastas alterada), perguntar:

> "Isso mudou algo no teu contexto. Quer que eu atualize os arquivos de memória?"

Se sim, identificar o que precisa atualizar:

- **Novo cliente, serviço, ferramenta, equipe** → `_contexto/empresa.md`
- **Mudança de prioridade ou foco** → `_contexto/estrategia.md`
- **Correção de tom ou estilo** → `_contexto/preferencias.md`
- **Nova pasta, regra de organização, skill criada** → `CLAUDE.md`
- **Mudança visual (cores, fontes, logo)** → `marca/design-guide.md`

Mostrar o que vai mudar antes de salvar. Não reformatar o arquivo inteiro, só adicionar ou editar a linha relevante.

**Quando NÃO perguntar:**
- Tarefas pontuais que não mudam o contexto (ex: escrever um email, criar um post avulso)
- Perguntas simples ou conversas sem ação
- Mudanças que já foram salvas pelo bloco "Aprender com correções"

**Dica:** se não sabe se algo mudou, rode `/atualizar` pra uma varredura completa.

---

## Criação de skills

Quando o usuário pedir pra criar uma nova skill:

1. Verificar se existe um template relevante em `templates/skills/`. Se existir, usar como base e adaptar pro contexto do usuário
2. Perguntar: "Essa skill é específica pra esse projeto ou vai ser útil em qualquer projeto?"
   - Específica desse negócio → salvar em `.claude/skills/nome-da-skill/SKILL.md` (local)
   - Útil em qualquer projeto → salvar em `~/.claude/skills/nome-da-skill/SKILL.md` (global)
3. Ler `_contexto/empresa.md` e `_contexto/preferencias.md` pra calibrar o conteúdo da skill ao contexto do negócio
4. Se a skill precisar de arquivos de apoio (templates, referências, exemplos), criar dentro da pasta da skill
5. Seguir o fluxo da skill-creator nativa do Claude Code
