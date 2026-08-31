# Simulador de Investimentos em Fundos Imobiliários (FIIs)
Este projeto apresenta um Simulador de Investimentos em Fundos Imobiliários (FIIs) construído em Microsoft Excel, desenvolvido para transformar dados financeiros em tomadas de decisão estratégicas.

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-2016%2B-107C41?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Bootcamp DIO](https://img.shields.io/badge/DIO-Santander_Excel_com_IA-orange?style=for-the-badge)

## Descrição do Desafio
Este repositório contém a solução do desafio prático do bootcamp **Santander - Excel com IA e Claude** oferecido pela **DIO (Digital Innovation One)**. 

O objetivo do projeto é construir uma ferramenta prática e interativa em **Microsoft Excel** que auxilia investidores na simulação de aportes em Fundos Imobiliários (FIIs), projetando a evolução do patrimônio ao longo do tempo e a geração de renda passiva mensal (dividendos), além de sugerir a distribuição ideal de carteira de acordo com o Perfil do Investidor.

---

## Arquitetura e Funcionalidades da Planilha

### 1. Parâmetros de Entrada (Configurações do Investidor)
A planilha coleta dados de entrada essenciais para personalizar os cálculos:
* **Salário / Renda Mensal:** Renda base do investidor.
* **Aporte Mensal:** Valor a ser investido mensalmente (calculado ou definido pelo usuário).
* **Taxa de Rendimento Mensal (% a.m.):** Taxa média de retorno estimada dos FIIs.
* **Tempo de Investimento:** Período de acúmulo projetado em anos.

### 2. Cálculos Financeiros e Fórmulas Utilizadas

* **Patrimônio Acumulado (Juros Compostos):**
  Utilizou-se a função de **Valor Futuro (`VF` / `FV`)** para simular o valor final acumulado considerando os aportes periódicos e a taxa de juros:
  ```excel
  =VF(taxa_mensal; qtd_tempo * 12; -aporte_mensal)
