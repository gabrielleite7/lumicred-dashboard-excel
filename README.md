# LUMICRED - Dashboard de Credito Consignado

![LUMICRED](logo.png)

Projeto de dashboard executivo em Excel para acompanhamento de uma operacao de credito consignado INSS. O arquivo consolida contratos digitados, status da carteira, volume financeiro, produtividade por operador e concentracao por banco em um painel de leitura gerencial.

## Objetivo

Criar uma visao unica para a diretoria acompanhar producao, conversao, cancelamentos, mix bancario e desempenho operacional do time comercial.

O projeto foi desenvolvido como estudo pratico de Excel para analise de dados, com foco em:

- tratamento e modelagem de base operacional;
- indicadores executivos de credito consignado;
- tabelas dinamicas e graficos dinamicos;
- formulas de KPI, ranking e diagnostico;
- segmentacao visual para acompanhamento do funil.

## Arquivo principal

O dashboard final esta em:

[`Dados/LUMICRED - Credito Consignado.xlsx`](Dados/LUMICRED%20-%20Cr%C3%A9dito%20Consignado.xlsx)

## Estrutura do workbook

- `Dashboard`: painel visual com KPIs e graficos da operacao.
- `Analises`: indicadores, tabelas auxiliares e rankings.
- `Insights`: leitura executiva dos resultados e plano de acao sugerido.
- `fContratos`: tabela fato com os contratos tratados.
- `dCalendario`: dimensao de calendario.
- `dOperador`: dimensao de operadores.

## Processo de tratamento

A base foi preparada no Power Query com as seguintes etapas:

- importacao da planilha original;
- promocao da primeira linha como cabecalho;
- remocao de colunas vazias ou sem uso analitico;
- substituicao de valores `---` por nulos;
- aplicacao de tipos corretos para datas, numeros e moedas;
- remocao de registros sem data de lancamento;
- remocao de duplicidades por CPF, numero de contrato e data;
- criacao da flag `AVERBADO`;
- criacao de `Ano_Mes`;
- simplificacao do tipo de contrato em `Tipo Simples`.

## Indicadores acompanhados

- quantidade de contratos digitados;
- quantidade de contratos averbados;
- taxa de conversao;
- volume bruto;
- volume bruto averbado;
- volume liquido;
- ticket medio;
- prazo medio;
- taxa de cancelamento;
- ranking por operador;
- volume por banco;
- distribuicao por status.

## Principais formulas usadas

```excel
=COUNTA(fContratos[Nº CONTRATO])
=SUMIF(fContratos[AVERBADO],1,fContratos[AVERBADO])
=SUM(fContratos[VLR BRUTO])
=C6/C4
=AVERAGE(fContratos[PRAZO])
=COUNTIF(fContratos!K:K,"CANCELADO")/COUNTA(fContratos!K:K)
=COUNTIF(fContratos!Y:Y,1)/COUNTA(fContratos!A:A)
=SUMIF(fContratos!Y:Y,1,fContratos!R:R)
```

Tambem foram usadas formulas de analise executiva na aba `Insights`, incluindo `COUNTIFS`, `SUMIFS`, `INDEX`, `MATCH`, `MAX`, `MIN`, `LARGE` e `TEXT`.

## Leitura dos resultados

A analise aponta uma operacao com volume bruto digitado relevante, mas com gargalos claros no funil:

- baixa conversao de contratos digitados em contratos averbados;
- cancelamento elevado em relacao ao total da carteira;
- concentracao relevante de volume nos principais bancos parceiros;
- forte concentracao de producao nos operadores de maior volume;
- oportunidade de acompanhamento dos contratos em assinatura digital e etapas pendentes.

## Plano de acao sugerido

- investigar causas de cancelamento;
- acompanhar contratos parados em assinatura digital;
- diversificar a producao entre bancos parceiros;
- replicar boas praticas dos operadores com maior produtividade;
- reforcar campanhas nos periodos de menor producao;
- avaliar treinamento ou redistribuicao operacional para operadores com baixa producao.

## Tecnologias

- Microsoft Excel
- Power Query
- Formulas Excel
- Tabelas Dinamicas
- Graficos Dinamicos
- Segmentacao de Dados

## Observacao

Este projeto foi estruturado para portfolio e demonstra competencias em tratamento de dados, modelagem analitica, construcao de KPIs e comunicacao executiva em Excel.
