# FRT — Dashboard de Frete

> **Aplicação web 100% client-side** para análise e acompanhamento de faturamento de fretes por filial, construída em HTML + CSS + JavaScript puro, sem servidor nem banco de dados.

---

## Sumário

- [Visão Geral](#visão-geral)
- [Tecnologias](#tecnologias)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Estrutura de Dados Esperada](#estrutura-de-dados-esperada)
- [Como Usar](#como-usar)
- [Funcionalidades](#funcionalidades)
  - [Visão Geral (Overview)](#visão-geral-overview)
  - [Abas por Filial](#abas-por-filial)
  - [Análise Anual (Intel)](#análise-anual-intel)
  - [Negociação](#negociação)
  - [Busca AWB](#busca-awb)
  - [Comissão](#comissão)
- [Conceito SC — Sem Comissionado](#conceito-sc--sem-comissionado)
- [Fluxo de Dados Interno](#fluxo-de-dados-interno)
- [Fórmulas e Métricas](#fórmulas-e-métricas)
- [Persistência](#persistência)
- [Limitações e Requisitos](#limitações-e-requisitos)

---

## Visão Geral

O **Dashboard de Frete** é uma SPA (Single Page Application) que lê arquivos `.xls` / `.xlsx` diretamente do sistema de arquivos do usuário — sem upload para nenhum servidor — e gera análises completas de faturamento de fretes, divididos por filial, período e tipo.

A aplicação detecta automaticamente as filiais pelos nomes dos arquivos (`FILIAL-Carteira.xls`, `FILIAL-Extra.xls`, `FILIAL-SC.xls`) e monta os dashboards dinamicamente. Tudo roda localmente no navegador via **File System Access API** (Chrome/Edge).

---

## Tecnologias

| Tecnologia | Uso |
|---|---|
| `HTML5` | Estrutura e interface |
| `CSS3` (Vanilla) | Estilização dark-mode, glassmorphism, responsividade |
| `JavaScript ES2022+` | Toda lógica de negócio |
| [`Chart.js 4.4`](https://www.chartjs.org/) | Gráficos (barras, linhas, doughnut) |
| [`SheetJS (xlsx 0.18)`](https://sheetjs.com/) | Leitura de arquivos `.xls` / `.xlsx` |
| **File System Access API** | Acesso à pasta local sem upload |
| **IndexedDB** | Persistência do handle da pasta entre sessões |
| **localStorage** | Persistência da lista de CNPJs de negociação e período ativo |

---

## Estrutura do Projeto

```
Faturamento/
├── index.html      ← Estrutura HTML completa (592 linhas)
├── styles.css      ← Estilo dark-mode, componentes, responsivo (159 linhas)
├── app.js          ← Toda a lógica JS (~1.438 linhas)
├── Tabelas.xlsx    ← Planilha auxiliar de referência
├── 2020/           ← Dados históricos por filial (sem subpasta mensal)
│   └── CSU/
│       ├── CSU-Carteira.xls
│       └── CSU-Extra.xls
├── 2021/ ... 2025/ ← Idem
└── 2026/           ← Ano atual: subpastas por mês
    ├── 01/
    │   └── CSU/
    │       ├── CSU-Carteira.xls
    │       ├── CSU-Extra.xls
    │       └── CSU-SC.xls   ← Sem Comissionado (opcional)
    ├── 02/ ... 04/
```

---

## Estrutura de Dados Esperada

### Anos anteriores (histórico)
```
AAAA/
└── SIGLA/
    ├── SIGLA-Carteira.xls
    └── SIGLA-Extra.xls
```

### Ano atual
```
AAAA/
└── MM/
    └── SIGLA/
        ├── SIGLA-Carteira.xls   ← Faturamento comissionado (Carteira)
        ├── SIGLA-Extra.xls      ← Faturamento comissionado (Extra)
        └── SIGLA-SC.xls         ← Sem Comissionado (opcional)
```

> A sigla da filial é extraída automaticamente do nome do arquivo (ex: `CSU-Carteira.xls` → filial `CSU`).

### Colunas reconhecidas nos XLS

O parser é tolerante a variações de nome de coluna, buscando por aliases:

| Campo | Aliases aceitos |
|---|---|
| AWB | `Nro. AWB`, `AWB`, `Numero`, etc. |
| Data | `Data Emissão`, `Dt. Emissão`, `Data`, etc. |
| Remetente | `Remetente`, `Nome Remetente`, `Razão Social Rem` |
| Destinatário | `Destinatário`, `Destinatario`, etc. |
| Cidade/UF | `Cidade Destino`, `Município`, etc. |
| Volumes | `Volumes`, `Qtd Volumes`, `Qtd` |
| Peso | `Peso Real`, `Peso Bruto`, `Peso (kg)` |
| Valor Mercantil | `Valor Mercantil`, `Vl Mercantil` |
| **Valor Frete** | `Valor do Frete`, `Valor Frete Total`, `Vl. Frete` |
| Tipo de Frete | `Tipo de Frete`, `Tp. Frete` — valores: `CIF` / `FOB` |
| Modal | `Modal de Transporte` — valores: `Aéreo` / `Rodoviário` |
| Ramo de Atividade | `Ramo de Atividade`, `Segmento` |
| CNPJ Remetente | `CNPJ Remetente`, `CNPJ/CPF Remetente` |
| CNPJ Destinatário | `CNPJ Destinatário`, `CNPJ/CPF Destinatário` |

---

## Como Usar

1. **Abra `index.html`** no Chrome ou Edge (requer suporte à File System Access API).
2. Clique em **"Selecionar pasta raiz"** e escolha a pasta raiz que contém os subdiretórios de anos.
3. A aplicação escaneia automaticamente todas as subpastas, detecta filiais e períodos, e carrega os dados.
4. Use as **setas ← →** no header para navegar entre os meses carregados.
5. Clique nas abas para alternar entre as visões disponíveis.

> A pasta escolhida é salva no IndexedDB e reaberta automaticamente na próxima visita.

---

## Funcionalidades

### Visão Geral (Overview)

Painel consolidado do mês selecionado com todas as filiais juntas:

- **Banner de KPIs**: Total atual, Total ano anterior (mês completo), % vs ano anterior pró-rata, Projeção mês completo, Gap vs meta.
- **Barra de dias úteis**: Dias úteis totais, faturados e restantes no mês atual vs anterior.
- **KPIs por filial**: Faturamento de cada filial com badge de variação %.
- **KPIs globais**: AWBs totais, Carteira total, Extra total, Média diária.
- **Gráficos**:
  - Faturamento diário empilhado por filial + linha do ano anterior + overlay SC.
  - Distribuição por filial (barras horizontais).
  - Projeção × Meta (linha de run-rate vs meta).
  - Top ramos de atividade por filial.
  - Tipo de frete consolidado (CIF/FOB/etc).

---

### Abas por Filial

Cada filial detectada ganha uma aba dedicada com:

- **KPIs individuais**: Faturado (Carteira + Extra), SC Potencial, Média Diária, Projeção Mês, AWBs.
- **Barra de progresso**: Realizado atual vs total do mesmo mês no ano anterior (meta).
- **Gráfico diário**: Barras empilhadas (atual vs SC) + linha do ano anterior.
- **Gráficos**: Tipo de frete (CIF/FOB) e Modal (Aéreo/Rodoviário).
- **Tabela diária**: Dia a dia com Carteira, Extra, SC do dia, Total, AWBs, Volumes, equivalente ano anterior e Δ%.
  - Linha clicável: abre **modal de clientes novos/retorno** do dia.
  - Coluna SC clicável: abre **modal com AWBs SC do dia**.
- **Top clientes** (gráfico de barras horizontais).
- **Exportação CSV**: faturamento diário e SC separados.

---

### Análise Anual (Intel)

Projeção e inteligência para o ano completo:

- **Projeção anual consolidada** com cenário (🟢 Realista / 🟡 Conservador / 🔴 Agressivo).
- **Realizado até ontem**, SC potencial do ano, Dias úteis realizados.
- **Cards por filial**: Projeção, realizado, média diária, comparativo com melhor ano histórico.
- **SC Anual**: Totais SC, migrados para comissão, a recuperar, barra de conversão, top dias com mais SC.
- **Top Meses Históricos** por filial (melhores meses de todos os anos anteriores).
- **Índice de Sazonalidade**: gráfico de linha mostrando o coeficiente histórico por mês (1.0 = média).
- **Análise por Tipo de Frete**: CIF, FOB, Aéreo, Rodoviário — valor, %, AWBs e gráfico doughnut por filial.
- **Top 10 Clientes** (acumulado anual por filial): nome, CNPJ, frete, AWBs, % do total.

---

### Negociação

Carteira de CNPJs em negociação com cruzamento automático:

- **Carregamento**: via planilha XLS (CNPJ, Razão Social, Data Negociação, Tabela, Vendedor, Contato, Obs) ou por adição manual.
- **Cruzamento automático**: CNPJs da lista são verificados contra todos os AWBs carregados.
- **Alertas**: badge de notificação na aba quando algum CNPJ em negociação movimentar frete.
- **Tabela com filtros**: status (Movimentou / Aguardando), vendedor, busca por CNPJ/nome/tabela.
- **Ordenação** por qualquer coluna.
- **Modal de AWBs**: clicando na linha de um CNPJ que movimentou, exibe os embarques detalhados.
- **Persistência**: lista salva no `localStorage` entre sessões.

---

### Busca AWB

Motor de busca cross-período para todos os AWBs carregados:

- **Filtros**: AWB, Remetente, Destinatário, Cidade/UF, Filial, Período, Tipo (Carteira/Extra/SC), Modal, intervalo de datas, frete mínimo.
- **Ordenação** por AWB, Data, Remetente, Peso, Frete.
- **Paginação** (50 por página).
- **Exportação CSV** do resultado filtrado.

---

### Comissão

Calculadora de comissão pessoal baseada nos AWBs carregados:

- **Taxas fixas**: FOB = 1,3% · CIF = 0,2%.
- **Banner de KPIs**: Comissão do ano, do mês, FOB/CIF base faturado.
- **SC Potencial**: Se os AWBs SC migrarem para comissão este mês, qual seria a comissão adicional.
- **Tabela por filial** (acumulado anual): FOB base, CIF base, Comissão, SC a recuperar, Comissão potencial SC.
- **Gráficos**: Evolução mensal da comissão (comissionado + potencial SC) e distribuição por tipo de frete.
- **Tabela diária** do mês atual: FOB/CIF faturado por dia, comissão gerada, SC do dia e comissão potencial SC.
- **Exportação CSV**.

---

## Conceito SC — Sem Comissionado

O arquivo `FILIAL-SC.xls` contém AWBs **não comissionados** (embarques que o vendedor fez mas que ainda não entraram na carteira comissionada). O sistema trata esses AWBs como:

1. **Potencial a recuperar**: somados como potencial de comissão do mês.
2. **Acompanhamento diário**: visível na tabela da filial, coluna SC, clicável para ver os AWBs do dia.
3. **Projeção**: SC do mês atual entra na projeção anual como potencial adicional.
4. **Análise anual**: seção dedicada mostra a conversão (migrado → comissionado / a recuperar).

> Anos anteriores não têm SC — apenas o ano atual pode ter arquivo SC.

---

## Fluxo de Dados Interno

```
openFolder()
    └── scanAllPeriods()
            ├── Detecta anos, meses, filiais, arquivos
            ├── loadPeriodData() → parseXLS() → {cur[], prev[], sc[]}
            ├── loadHistory() → readAnnualAllMonths() → computeFilialHistory()
            ├── discoverFiliaisFromData()
            ├── buildDynamicTabs() + buildDynamicPanes()
            ├── rebuildAllAWBIndex()
            ├── computeAnnualProjection()
            └── loadAndRender()
                    └── renderPane(activeTab)
                            ├── renderOverview()
                            ├── renderFilial(f)
                            ├── renderIntel()
                            ├── renderNeg()
                            ├── initAWB()
                            └── renderComissao()
```

**Monitoramento de modificações**: a cada 5 segundos, o sistema verifica se algum arquivo XLS foi modificado e recarrega automaticamente os dados.

---

## Fórmulas e Métricas

| Métrica | Fórmula |
|---|---|
| **Projeção do mês** | `(Total atual ÷ Dias úteis decorridos) × Dias úteis totais do mês` |
| **% vs Ano Anterior pró-rata** | `(Total atual − (Total prev ÷ DU_prev × DU_decorridos)) ÷ (Total prev ÷ DU_prev × DU_decorridos) × 100` |
| **Projeção anual** | Run-rate × dias úteis restantes, com ajuste de sazonalidade histórica (cap 0.5× a 2.0×), blended com histórico quando < 15 dias realizados |
| **Comissão FOB** | `Valor Frete FOB × 1,3%` |
| **Comissão CIF** | `Valor Frete CIF × 0,2%` |
| **Sazonalidade** | `Média histórica do mês ÷ Média histórica geral` |
| **Cliente Novo** | CNPJ/nome que nunca transportou antes do dia analisado |
| **Cliente Retorno** | CNPJ/nome que não aparecia há mais de 90 dias |

---

## Persistência

| Dado | Mecanismo | Descrição |
|---|---|---|
| Handle da pasta raiz | IndexedDB | Reaberta automaticamente na próxima sessão |
| Último período ativo | localStorage | Restaura o mês selecionado |
| Lista de negociação | localStorage | CNPJs em negociação persistem entre sessões |

---

## Limitações e Requisitos

- **Navegador**: Chrome ou Edge com suporte à [File System Access API](https://caniuse.com/native-filesystem-api). Não funciona no Firefox.
- **Conexão**: Necessária apenas para carregar as fontes Google Fonts e as libs CDN (Chart.js, SheetJS) na primeira carga. Após isso, funciona offline.
- **Performance**: Arquivos muito grandes (> 50k linhas por XLS) podem tornar o carregamento lento — toda a leitura é feita no thread principal do browser.
- **Segurança**: Nenhum dado sai do computador. Tudo é processado localmente.
- **Formato de data**: O parser suporta datas como objeto Date, serial Excel (number), `dd/mm/yyyy` e `yyyy-mm-dd`.

---

## Deploy no GitHub Pages

Por ser um projeto **100% estático** (HTML + CSS + JS), pode ser publicado diretamente no GitHub Pages:

1. Faça o push dos arquivos `index.html`, `styles.css` e `app.js` para o repositório.
2. Ative o GitHub Pages nas configurações do repositório (branch `main`, pasta `/`).
3. Acesse via `https://zvitorsantos.github.io/faturamento/`.

> ⚠️ Os dados (arquivos XLS) **nunca devem ser commitados** — adicione as pastas de anos ao `.gitignore`.

```gitignore
# .gitignore
2020/
2021/
2022/
2023/
2024/
2025/
2026/
Tabelas.xlsx
```
