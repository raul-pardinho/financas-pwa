# Dashboard V2 — arquitetura financeira

## Objetivo

A interface deixa de ser organizada por **Competência × Caixa** e passa a ser organizada pelas perguntas de gestão:

1. Como estamos no mês?
2. Onde o orçamento está desviando?
3. Quanto temos?
4. Quanto devemos?
5. Quanto do nosso crédito está exposto a terceiros?
6. O que acontece operacionalmente em cartões e contas?

Competência e caixa continuam existindo, mas como camadas de cálculo.

## Navegação

- **Visão geral** — resumo gerencial do mês.
- **Orçamento** — Orçado × Meta × Realizado MTD × Forecast por categoria.
- **Patrimônio** — ativos, passivos, liquidez e evolução.
- **Dívidas** — saldos devedores e parcelas.
- **Terceiros** — exposição de Heitor e outros terceiros.
- **Cartões & Caixa** — reconciliação operacional.

## Quatro referências mensais

### Orçado
Valor realista aprovado para a categoria.

### Meta
Valor de desafio. Em despesas, normalmente menor ou igual ao orçamento. Não substitui o orçamento: mostra quanto queremos tentar economizar além do cenário realista.

### Realizado MTD
Consumo reconhecido por competência até o momento.

### Forecast
Estimativa de fechamento do mês.

Regras iniciais:
- **FIXO / FIXO_ESCALONADO / FIXO_TEMPORARIO / FIXO_SAZONAL**: usa o valor conhecido, respeitando realizado acima do previsto.
- **EVENTO**: usa o valor conhecido do evento; não extrapola por pacing.
- **PACING_CALENDARIO**: extrapola o realizado pelos dias corridos.
- **PACING_DIAS_UTEIS**: extrapola o realizado pelos dias úteis.
- Nos primeiros 15% do mês, o forecast não reduz automaticamente abaixo do orçamento por falta de movimentos.
- Em meses passados, o forecast converge para o realizado.
- Em meses futuros, o forecast começa igual ao orçamento.

## Regra de cartões

Cartão não é categoria de despesa.

Uma compra pode aparecer em três perspectivas sem ser duplicada:
- **Competência**: categoria econômica no mês da decisão de consumo.
- **Cartão**: obrigação na fatura.
- **Caixa**: saída quando a fatura for paga.

Pagamentos de fatura nunca devem duplicar o consumo por competência.

## Terceiros

Compras de Heitor e outros terceiros:
- não compõem consumo familiar;
- aumentam a obrigação do cartão;
- geram exposição/valor a receber;
- reembolsos aparecem no caixa operacional.

## Alimentação

Estrutura pacificada:
- **Supermercado**
- **Alimentação dia a dia — trabalho Raul**
- **Restaurantes** — presencial e delivery
- **Conveniência** — padaria, OXXO e compras equivalentes

## Fonte de orçamento

Enquanto a revisão estiver em andamento, o branch V2 reconhece **orcamento_v2** como fonte de orçamento e a identifica visualmente como **rascunho**.

A estrutura também suporta cenário **META** na aba `planejamento`. Enquanto não houver linhas META, a interface mostra **Meta: A definir**.

## Regra de publicação

O branch `dashboard-v2` é experimental. O `main` continua sendo a versão estável até revisão funcional, reconciliação dos números e aprovação da interface.
