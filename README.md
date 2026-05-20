# LUMICRED - Dashboard de Crédito Consignado

![Logo LUMICRED](logo.png)

Projeto de dashboard em Excel para análise de uma operação de crédito consignado INSS. A proposta foi transformar uma base de contratos em uma visão simples para acompanhar produção, conversão, volume financeiro, bancos parceiros e desempenho dos operadores.

Os dados são fictícios e foram usados apenas para estudo e portfólio.

## Dashboard

![Dashboard LUMICRED](imagens/Dashboard.png)

## Arquivo principal

[`LUMICRED - Crédito Consignado.xlsx`](LUMICRED%20%20ANALYTICS/LUMICRED%20-%20Cr%C3%A9dito%20Consignado.xlsx)

## Pontos-chave do projeto

- Tratamento da base no Power Query.
- Criação da tabela fato `fContratos`.
- Criação de tabelas auxiliares para calendário e operadores.
- Relacionamento entre tabelas para facilitar a análise.
- Construção de KPIs principais da operação.
- Ranking de operadores por produção e conversão.
- Análise de volume bruto por banco.
- Segmentações por operador, status e banco.
- Aba `Insights` com leitura dos principais pontos de atenção.

## KPIs acompanhados

![KPIs do dashboard](imagens/5%20KPIs.png)

- Quantidade de contratos digitados.
- Quantidade de contratos averbados.
- Percentual de conversão.
- Valor bruto total.
- Ticket médio.

## Modelo e relacionamento

![Relacionamento de tabelas](imagens/Relacionamento%20de%20tabelas.png)

O modelo foi organizado com a base principal de contratos e tabelas de apoio para análise por data e operador.


As visualizações ajudam a identificar:

- meses com maior ou menor volume de contratos;
- bancos com maior participação no volume bruto;
- operadores com maior produção;
- concentração de contratos por status;
- gargalos na conversão.

## Fórmulas usadas

Algumas fórmulas aplicadas no projeto:

```excel
=CONT.VALORES(fContratos[Nº CONTRATO])
=CONT.SES(fContratos[AVERBADO];1)
=SOMA(fContratos[VLR BRUTO])
=SOMASES(fContratos[VLR BRUTO];fContratos[AVERBADO];1)
=CONT.SE(fContratos[STATUS];"CANCELADO")/CONT.VALORES(fContratos[STATUS])
=MÉDIA(fContratos[PRAZO])
=ÍNDICE(Analises!B21:B27;CORRESP(MÁXIMO(Analises!C21:C27);Analises!C21:C27;0))
```

## Principais aprendizados

- O dashboard facilitou a leitura da operação em poucos indicadores.
- A conversão foi um dos pontos mais importantes da análise.
- O cancelamento apareceu como ponto de atenção.
- A produção ficou concentrada em alguns operadores.
- Alguns bancos concentraram grande parte do volume bruto.

## Ferramentas utilizadas

- Microsoft Excel
- Power Query
- Fórmulas Excel
- Tabelas Dinâmicas
- Gráficos Dinâmicos
- Segmentação de Dados
