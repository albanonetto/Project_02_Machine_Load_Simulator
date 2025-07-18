# Simulador de Carga Máquina e Planejamento de Capacidade

## 📖 Descrição

Esta ferramenta é uma aplicação avançada em **Excel com Power Query e VBA**, projetada para simular a carga de trabalho em centros de usinagem e linhas de produção. O sistema permite que o Planejamento e Controle da Produção (PCP) analise a capacidade produtiva, identifique gargalos e tome decisões estratégicas com base em dados.

O simulador processa a carteira de pedidos ou planos operativos, permitindo ajustar dinamicamente parâmetros como turnos de trabalho, mix de máquinas, e metas de OEE (Overall Equipment Effectiveness) para prever a ocupação e otimizar a alocação de recursos.

***
![Imagem dashboard](Imagem2.jpg)

## ✨ Funcionalidades Principais

-   **Integração de Dados (ETL):** Utiliza **Power Query** para extrair, transformar e estruturar de forma robusta os dados de engenharia de produto a partir de fontes externas.
-   **Simulação de Cenários:** Permite a criação de múltiplos cenários para comparar o impacto de diferentes volumes de produção ou configurações de fábrica.
-   **Controle de Parâmetros:** Interface centralizada para ajustar variáveis críticas como OEE (Meta vs. Realizado), turnos e disponibilidade de máquinas.
-   **Roteirização Inteligente:** Define a rota de produção ideal para cada item com base em regras de negócio centralizadas (família, tipo, etc.).
-   **Reprogramação Automatizada:** Um motor de otimização em **VBA** analisa a ocupação e, ao detectar sobrecargas (>100%), busca e aloca automaticamente máquinas alternativas com capacidade disponível.
-   **Relatórios e Análises:** Gera visualizações claras da ocupação em horas por máquina, por semana e por mês, permitindo análises de capacidade e necessidade de equipes.

***

## 🏗️ Arquitetura Otimizada

A arquitetura do projeto foi desenhada para máxima clareza e manutenibilidade, operando com quatro motores principais que separam as responsabilidades do sistema:

| Componente | Tecnologia | Função |
| :--- | :--- | :--- |
| **Motor de Dados** | Power Query | Responsável por todo o ETL (Extração, Transformação e Carga), entregando dados limpos e estruturados para a planilha. |
| **Motor de Regras** | Tabelas Excel | Centraliza todas as regras de negócio (roteiros, famílias, setups, OEE) em tabelas de configuração, eliminando a lógica "escondida" em fórmulas complexas. |
| **Motor de Cálculo** | Fórmulas Excel | Realiza os cálculos de carga, horas e ocupação em tempo real, lendo os dados do Motor de Dados e as regras do Motor de Regras. |
| **Motor de Otimização**| VBA Generalizado | Executa a lógica de reprogramação de forma flexível e genérica, consultando o Motor de Regras para encontrar as melhores alternativas. |

***

## 🚀 Como Usar

1.  **Atualizar Base de Dados:** Garantir que o arquivo de origem .xls esteja atualizado com as últimas informações de engenharia.
2.  **Atualizar Consulta:** No Excel, vá em `Dados > Atualizar Tudo` para que o Power Query processe as novas informações.
3.  **Ajustar Parâmetros:** Na aba `Painel de Controle`, ajuste os parâmetros da simulação (cenário, OEE, turnos).
4.  **Executar Simulação:** Clique no botão para rodar a macro principal de reprogramação e otimização.
5.  **Analisar Resultados:** Verifique as abas de ocupação por máquina e os resumos gerados para tomar decisões.

***

## 💻 Tecnologias Utilizadas

-   Microsoft Excel
-   Power Query (Linguagem M)
-   Visual Basic for Applications (VBA)
