# 📦 Sistema de Controle de Devoluções — Transben

> **Plataforma integrada ao Google Sheets + Apps Script para gestão completa do ciclo de vida das devoluções por fornecedor, com automação de e-mails, relatórios em PDF e painel de gestão em tempo real.**

---

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Estrutura da Planilha](#estrutura-da-planilha)
- [Menu Principal](#menu-principal)
- [Funcionalidades](#funcionalidades)
  - [Lançar / Excluir Devolução](#1-lançar--excluir-devolução)
  - [Buscar NF / Fornecedor](#2-buscar-nf--fornecedor)
  - [Enviar E-mail de Devolução](#3-enviar-e-mail-de-devolução)
  - [Dar Baixa para Devolução (PDF)](#4-dar-baixa-para-devolução-pdf)
  - [Dar Baixa para Venda](#5-dar-baixa-para-venda)
  - [Reabrir Devoluções](#6-reabrir-devoluções)
  - [Relatórios](#7-relatórios)
  - [Verificar Atrasos](#8-verificar-atrasos)
  - [Arquivamento Manual](#9-arquivamento-manual)
  - [Backup e Restauração](#10-backup-e-restauração)
  - [Auditoria e Histórico](#11-auditoria-e-histórico)
  - [Configurações do Sistema](#12-configurações-do-sistema)
- [Status e Cores](#status-e-cores)
- [Automações e Triggers](#automações-e-triggers)
- [Instalação / Reinstalação](#instalação--reinstalação)
- [Perguntas Frequentes](#perguntas-frequentes)

---

## Visão Geral

O sistema foi desenvolvido inteiramente dentro do Google Sheets com Google Apps Script. Ele centraliza o controle de notas fiscais devolvidas por fornecedor, desde o lançamento inicial até a conclusão (devolução ou venda), com rastreabilidade completa de todas as ações via log interno.

**Fornecedores operacionais padrão:**
- Britania
- Unilever
- Fornecedores Variados

Novos fornecedores podem ser adicionados pelo menu, sem precisar editar código.

---

## Estrutura da Planilha

| Aba | Descrição |
|---|---|
| `Dashboard` | Painel de gestão com KPIs, gráficos e filtro por período |
| `Britania` | Devoluções do fornecedor Britania |
| `Unilever` | Devoluções do fornecedor Unilever |
| `Fornecedores Variados` | Devoluções de outros fornecedores |
| `Historico_Arquivo` | Itens concluídos arquivados automaticamente |
| `_Log` | Registro auditável de todas as ações do sistema |
| `_EmailsEnviados` | Histórico de todos os e-mails enviados |
| `_Backup_Snapshot` | Snapshot para backup/restauração (oculta) |

---

## Menu Principal

Ao abrir a planilha, o menu **📦 Devoluções** aparece automaticamente na barra superior do Google Sheets:

```
📦 Devoluções
├── ➕ Lançar / Excluir Devolução
├── 🔍 Buscar NF / Fornecedor
├── 📨 Enviar E-mail de Devolução
├── ─────────────────────────────
├── 📄 Dar baixa para Devolução
├── 🛒 Dar Baixa para Venda
├── 🔓 Reabrir Devoluções
├── ─────────────────────────────
├── 📊 Relatórios (Mensal / Semanal / Diário / Fornecedor)
├── 🔔 Verificar Atrasos Agora
├── ─────────────────────────────
├── 📦 Forçar Arquivamento Manual
├── 💾 Backup e Restauração
├── ─────────────────────────────
├── 🔍 Auditoria e Histórico
├── ⚙️ Configurações do Sistema
└── 🔧 Configurar/Reinstalar Sistema
```

---

## Funcionalidades

### 1. Lançar / Excluir Devolução

**Menu:** `📦 Devoluções → ➕ Lançar / Excluir Devolução`

Três modos de operação:

**Lançamento Individual**
- Selecione o fornecedor/aba de destino
- Preencha: NFD, Nº NF, Data de Entrada, Tipo, Motivo, Descrição, Quantidade, Valor Unitário e Responsável
- Opcionalmente anexe o arquivo da nota fiscal (foto ou PDF) — armazenado no Google Drive
- Clique em **Salvar**

**Lançamento em Lote**
- Selecione o fornecedor
- Adicione múltiplas NFs de uma vez (mesmo fornecedor)
- Ideal para registrar várias notas de uma chegada em uma única operação

**Excluir Lançamento**
- Busque a NF ou NFD pelo número
- O sistema localiza e exibe o item para confirmação
- Informe o motivo da exclusão (obrigatório — registrado no log)
- Apenas itens com status **Pendente** podem ser excluídos

> ⚠️ O sistema impede lançamento de NFs duplicadas dentro da mesma aba.

---

### 2. Buscar NF / Fornecedor

**Menu:** `📦 Devoluções → 🔍 Buscar NF / Fornecedor`

- Busca por número de NF ou NFD em todas as abas operacionais simultaneamente
- Exibe: fornecedor, status, data de entrada, descrição, valor total e responsável
- Mostra também se o item está arquivado no `Historico_Arquivo`

---

### 3. Enviar E-mail de Devolução

**Menu:** `📦 Devoluções → 📨 Enviar E-mail de Devolução`

- Informe uma ou mais NFDs (separadas por vírgula ou quebra de linha)
- O sistema localiza os dados de cada nota e monta um e-mail profissional em HTML
- Permite anexar: PDFs/fotos das notas fiscais e comunicado adicional
- Destinatários configuráveis via **⚙️ Configurações**
- Itens do tipo **Falta** são arquivados automaticamente após o envio
- Histórico de todos os e-mails enviados é salvo na aba `_EmailsEnviados`

---

### 4. Dar Baixa para Devolução (PDF)

**Menu:** `📦 Devoluções → 📄 Dar baixa para Devolução`

- Informe os números das NFs (devem ser do mesmo fornecedor)
- O sistema gera um **PDF de relação de mercadorias** com layout formal
- O PDF é salvo automaticamente na pasta configurada no Google Drive
- O status de cada NF é alterado para **Devolvido**
- Link para abertura do PDF é exibido ao final

---

### 5. Dar Baixa para Venda

**Menu:** `📦 Devoluções → 🛒 Dar Baixa para Venda`

- Selecione as NFs que serão encaminhadas para venda
- Gera PDF com a relação de mercadorias (modelo próprio para venda)
- O PDF é salvo na pasta de vendas no Google Drive
- O status é alterado para **Venda**
- Registro automático no log

---

### 6. Reabrir Devoluções

**Menu:** `📦 Devoluções → 🔓 Reabrir Devoluções`

- Informe os números das NFs/NFDs concluídas que precisam ser reabertas
- O sistema localiza os itens (incluindo no `Historico_Arquivo`) e restaura o status para **Pendente**
- Remove a proteção de linha aplicada na conclusão
- Toda reabertura é registrada no log com data e hora

---

### 7. Relatórios

**Menu:** `📦 Devoluções → 📊 Relatórios`

Quatro tipos de relatório, todos gerados em **PDF** e enviados por e-mail:

| Tipo | Período | Conteúdo |
|---|---|---|
| **Mensal** | Mês selecionado | KPIs gerais, resumo por fornecedor, listagem completa |
| **Semanal** | Semana selecionada | Movimentações da semana por fornecedor |
| **Diário** | Dia selecionado | Lançamentos e baixas do dia |
| **Por Fornecedor** | Período livre | Visão detalhada de um fornecedor específico |

Todos os relatórios incluem:
- KPIs: total pendente, devolvido, venda e taxa de resolução
- Resumo por fornecedor com valores
- Listagem detalhada com NF, data, tipo, descrição, quantidade e valor
- PDF salvo no Drive + enviado por e-mail automaticamente

---

### 8. Verificar Atrasos

**Menu:** `📦 Devoluções → 🔔 Verificar Atrasos Agora`

- Varre todas as abas em busca de itens com status **Pendente** há mais de **30 dias**
- Gera relatório PDF de atraso crítico com os itens identificados
- Envia alerta automático por e-mail para os destinatários configurados
- Itens em atraso também são destacados em vermelho na planilha

> 💡 Este processo também pode ser disparado automaticamente por trigger diário.

---

### 9. Arquivamento Manual

**Menu:** `📦 Devoluções → 📦 Forçar Arquivamento Manual`

- Move todos os itens com status **Devolvido** ou **Venda** para a aba `Historico_Arquivo`
- Libera espaço nas abas operacionais
- O arquivamento automático também ocorre periodicamente via trigger

---

### 10. Backup e Restauração

**Menu:** `📦 Devoluções → 💾 Backup e Restauração`

**Fazer Backup:**
- Copia todos os dados de todas as abas operacionais para a aba oculta `_Backup_Snapshot`
- Exibe resumo por aba (quantidade de linhas salvas) e data/hora do backup

**Restaurar Backup:**
- Repõe os dados do snapshot nas abas operacionais
- Deve ser usado após executar **Configurar/Reinstalar Sistema**
- Requer que as abas já existam (rodar o instalador primeiro)

**Fluxo recomendado ao reinstalar:**
```
1. Fazer Backup  →  2. Configurar/Reinstalar  →  3. Restaurar Backup
```

---

### 11. Auditoria e Histórico

**Menu:** `📦 Devoluções → 🔍 Auditoria e Histórico`

Painel unificado com três abas:

**Histórico de NF**
- Busque qualquer NF ou NFD e veja todo o histórico de alterações
- Exibe: campo alterado, valor anterior, valor novo, data/hora e ação

**E-mails Enviados**
- Lista todos os e-mails de devolução enviados pelo sistema
- Filtro por período, fornecedor ou assunto
- Exibe: data, destinatários, NFDs incluídas, valor total e quantidade de anexos

**Log do Sistema**
- Registro completo de todas as operações (lançamentos, exclusões, exportações, configurações)
- Filtro por período

---

### 12. Configurações do Sistema

**Menu:** `📦 Devoluções → ⚙️ Configurações do Sistema`

Seis áreas de configuração:

**📧 E-mails e Destinatários**
- Gerencie a lista de e-mails que recebem relatórios e resumos
- Configure destinatários separados (CC) para alertas de atraso
- Opções: "Todos da lista geral" ou "Somente destinatários CC"

**🎨 Cores da Planilha**
- Personalize as cores de fundo por status:
  - 🔵 Pendente
  - 🟢 Devolvido
  - 🟠 Venda
  - 🔴 Alerta (+30 dias)
  - Cor do cabeçalho
- Cores são reaplicadas automaticamente em todas as abas ao salvar

**🏭 Adicionar Novo Fornecedor**
- Cria uma nova aba operacional com o layout padrão do sistema
- Opção de fixar o nome do fornecedor automaticamente em cada linha
- Sem necessidade de editar código

**🔧 Diagnóstico e Status**
- Exibe versão do sistema, data do último backup
- Contadores: itens concluídos, proteções ativas
- Uso de linhas por aba operacional

**🗂️ Limpeza do Log Antigo**
- Arquiva ou apaga registros do `_Log` com mais de N meses
- Opções: 1, 3, 6 ou 12 meses

**🗑️ Limpeza de Arquivos no Drive**
- Remove PDFs gerados (relatórios, vendas, devoluções) das pastas do Drive
- Filtros por tipo de arquivo e data de criação
- Arquivos vão para a lixeira (recuperáveis por até 30 dias)
- Os anexos de NF originais **nunca são apagados**

---

## Status e Cores

| Status | Cor padrão | Significado |
|---|---|---|
| **Pendente** | 🔵 Azul claro `#DDEEFF` | Item lançado, aguardando resolução |
| **Pendente +30 dias** | 🔴 Vermelho `#FFD5D5` | Crítico — mais de 30 dias sem conclusão |
| **Devolvido** | 🟢 Verde claro `#DDFFDD` | Mercadoria devolvida ao fornecedor com PDF |
| **Venda** | 🟠 Laranja claro `#FFE5CC` | Mercadoria enviada para venda |

> As cores podem ser totalmente personalizadas em **⚙️ Configurações → Cores da Planilha**.

---

## Automações e Triggers

O sistema utiliza triggers automáticos do Google Apps Script:

| Trigger | Frequência | Ação |
|---|---|---|
| `onOpen` | A cada abertura | Reaplicação de cores (cache de 1h) + criação do menu |
| Alerta de atraso | Diário (automático) | Verifica itens com +30 dias e envia e-mail |
| Resumo semanal | Segunda-feira | Envia resumo semanal por e-mail para todos os destinatários |
| Arquivamento | Periódico | Move itens concluídos para `Historico_Arquivo` |
| Dashboard | Após cada operação | Atualiza KPIs e gráficos (debounce de 8 segundos) |

---

## Instalação / Reinstalação

1. No Google Sheets, acesse **📦 Devoluções → 🔧 Configurar/Reinstalar Sistema**
2. O sistema criará automaticamente:
   - Todas as abas operacionais com layout e formatação padrão
   - Aba Dashboard com gráficos e KPIs
   - Abas internas (`_Log`, `_EmailsEnviados`, `Historico_Arquivo`)
   - Proteções de coluna e linha
   - Triggers automáticos
3. Se houver dados existentes, use **💾 Backup** antes de reinstalar e **Restaurar** depois

> Para adicionar um novo fornecedor sem reinstalar: **⚙️ Configurações → 🏭 Adicionar Novo Fornecedor**

---

## Perguntas Frequentes

**Posso excluir uma NF já devolvida ou vendida?**
Não diretamente. Devoluções concluídas são protegidas. Use **🔓 Reabrir Devoluções** para retornar ao status Pendente e então faça a exclusão.

**O sistema funciona offline?**
Não. Por ser baseado em Google Sheets + Apps Script, requer conexão com a internet e conta Google com permissões na planilha.

**Como adicionar um novo fornecedor?**
Via menu **⚙️ Configurações → 🏭 Adicionar Novo Fornecedor**. Não é necessário editar o código.

**Os PDFs gerados ficam onde?**
Em pastas do Google Drive configuradas no código (`ID_PASTA_DESTINO_VENDA`, `ID_PASTA_DESTINO_DEV`, `ID_PASTA_RELATORIOS`). Os links são exibidos ao final de cada operação.

**O que acontece se dois usuários usarem ao mesmo tempo?**
O sistema utiliza `LockService` do Apps Script para evitar conflitos. Se o sistema estiver ocupado, uma mensagem de "tente novamente" será exibida.

**Como alterar os e-mails que recebem relatórios?**
Via menu **⚙️ Configurações → 📧 E-mails e Destinatários**. As alterações são salvas nas propriedades do script e persistem mesmo após reinstalação.

**O backup salva os anexos das NFs?**
O backup salva os dados da planilha (incluindo os links para os anexos no Drive). Os arquivos físicos no Drive não são movidos e permanecem acessíveis pelos links.

---

## Versão

**v6.0** — Sistema de Controle de Devoluções · Transben

> Desenvolvido com Google Apps Script · Planilha Google Sheets
