# LUMICRED - Dashboard de Crédito Consignado

![LUMICRED](logo.png)

Projeto de dashboard em Excel para acompanhar uma operação de crédito consignado INSS. A ideia foi transformar uma base de contratos em uma visão mais clara da produção, conversão, cancelamentos, bancos parceiros e desempenho dos operadores.

Os dados usados são fictícios e fazem parte de um estudo prático de análise de dados no Excel.

## Objetivo

Montar um painel que ajude a responder perguntas simples da operação:

- quantos contratos foram digitados;
- quantos contratos foram averbados;
- qual foi a taxa de conversão;
- quanto foi movimentado em volume bruto;
- quais bancos concentram mais produção;
- quais operadores mais produziram;
- onde estão os principais gargalos do funil.

## Arquivo principal

O dashboard está no arquivo:

[`Dados/LUMICRED - Crédito Consignado.xlsx`](Dados/LUMICRED%20-%20Cr%C3%A9dito%20Consignado.xlsx)

## Estrutura do arquivo

- `Dashboard`: painel principal com KPIs e gráficos.
- `Insights`: resumo dos principais pontos de atenção da operação.
- `Analises`: cálculos, tabelas auxiliares e rankings.
- `fContratos`: base tratada dos contratos.
- `dCalendario`: calendário usado nas análises por período.
- `dOperador`: lista de operadores.

## Tratamento da base

A base foi tratada no Power Query antes da criação do dashboard. As principais etapas foram:

- importação da planilha original;
- ajuste dos cabeçalhos;
- remoção de colunas vazias ou sem uso;
- substituição de valores `---` por nulo;
- correção dos tipos de dados, como datas, números e valores financeiros;
- remoção de linhas sem data de lançamento;
- remoção de duplicidades;
- criação da coluna `AVERBADO`;
- criação do campo `Ano_Mes`;
- criação do campo `Tipo Simples` para facilitar a análise do tipo de contrato.

## Indicadores criados

- Contratos digitados
- Contratos averbados
- Taxa de conversão
- Volume bruto digitado
- Volume bruto averbado
- Ticket médio
- Prazo médio
- Taxa de cancelamento
- Produção por operador
- Volume por banco
- Distribuição por status

## Fórmulas principais

Algumas das fórmulas usadas no projeto:

```excel
=CONT.VALORES(fContratos[Nº CONTRATO])
=CONT.SES(fContratos[AVERBADO];1)
=SOMA(fContratos[VLR BRUTO])
=SOMASES(fContratos[VLR BRUTO];fContratos[AVERBADO];1)
=CONT.SE(fContratos[STATUS];"CANCELADO")/CONT.VALORES(fContratos[STATUS])
=MÉDIA(fContratos[PRAZO])
=ÍNDICE(Analises!B21:B27;CORRESP(MÁXIMO(Analises!C21:C27);Analises!C21:C27;0))
=MAIOR(Analises!C21:C27;1)
=TEXTO(MÁXIMO(Analises!F21:F28);"R$ #.##0,00")
```

## Principais aprendizados

Ao montar o dashboard, os pontos que mais chamaram atenção foram:

- a conversão ficou baixa em relação ao total de contratos digitados;
- o cancelamento ficou alto e precisa ser acompanhado de perto;
- parte importante do volume ficou concentrada em poucos bancos;
- alguns operadores concentraram boa parte da produção;
- contratos parados em assinatura digital podem impactar diretamente a averbação.

## Plano de ação sugerido

- acompanhar de perto os contratos cancelados;
- entender os motivos de cancelamento;
- atuar nos contratos parados em assinatura digital;
- comparar os bancos com maior e menor volume;
- observar boas práticas dos operadores com melhor desempenho;
- acompanhar a conversão por período para identificar gargalos.

## Ferramentas utilizadas

- Microsoft Excel
- Power Query
- Fórmulas Excel
- Tabelas Dinâmicas
- Gráficos Dinâmicos
- Segmentação de Dados

## Sobre o projeto

Este projeto foi criado para portfólio, com foco em tratamento de dados, construção de indicadores e apresentação de informações em formato de dashboard executivo no Excel.
