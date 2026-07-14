# Gestão e Controle de Obra 🏗️📊

Este é um dashboard interativo desenvolvido em **Angular** para acompanhamento físico-financeiro, planejamento de custos e controle de desvios em obras de engenharia e projetos. A ferramenta permite a importação de dados diretamente de planilhas do Excel para automatizar a geração de relatórios e gráficos.

---

## 🚀 Funcionalidades Principais

* **Importação Inteligente de Planilhas (Excel)**:
  * Suporte a arquivos `.xlsx` e `.xls` através da biblioteca `xlsx`.
  * Mapeamento inteligente de colunas, capaz de deduzir cabeçalhos mesmo com variações de escrita (ex: "orçado", "planejado", "previsto", "medido").
  * Algoritmo de busca robusto de metadados gerais (Valor do Contrato, Meses Corridos e Avanço Físico) nas primeiras linhas da planilha.
  * Normalização inteligente de porcentagens inseridas como números inteiros (ex: `8.33` convertido automaticamente para `8,33%`).
  * Preenchimento recíproco automático (recalcula valores se a porcentagem estiver presente, e vice-versa).

* **Painel de Indicadores (Cards)**:
  * **Valor do Contrato**: Exibição em formato monetário (BRL).
  * **Meses Corridos**: Prazo transcorrido da obra.
  * **Avanço da Obra**: Barra de progresso com porcentagem acumulada de execução física.

* **Tabela de Custos e Desvios**:
  * Detalhamento mês a mês contendo:
    * Valor Realizado e Valor Previsto.
    * **Desvio Valor**: Diferença absoluta em reais (Realizado - Previsto).
    * **Desvio / Previsto**: Variação percentual relativa ao planejado para o mês.
    * **Realizado (% Contrato)** e **Previsto (% Contrato)**: Peso de cada mês frente ao contrato global.
    * **Desvio (%)**: Desvio percentual do realizado frente ao previsto relativo ao valor total do contrato.
  * Destaques visuais para desvios positivos (verde) e negativos (vermelho).

* **Gráfico de Evolução (Curva S)**:
  * Gráfico vetorial SVG integrado e responsivo que exibe as curvas acumuladas de **Previsto Acumulado** (azul) e **Realizado Acumulado** (verde).
  * Linhas de grade monetárias automáticas e interatividade através de dicas de tela ao passar o mouse sobre os pontos do gráfico.

* **Diário de Acompanhamento**:
  * Ferramenta interativa para adicionar e salvar anotações ou justificativas de desvios de custos mês a mês.
  * Persistência em memória ao navegar entre diferentes meses selecionados.

---

## 🛠️ Tecnologias Utilizadas

* **Framework**: [Angular](https://angular.dev/) (Versão 22.0.5)
* **Estilização**: Vanilla CSS (com variáveis CSS e design responsivo moderno com efeito Glassmorphism)
* **Biblioteca de Excel**: [SheetJS (XLSX)](https://sheetjs.com/)
* **Compilador**: TypeScript 6
* **Ferramenta de Build**: Vite / Angular Builder

---

## 📦 Como Instalar e Rodar Localmente

### Pré-requisitos
* Node.js instalado (versão v20 ou superior recomendada)
* Gerenciador de pacotes `npm`

### Passos de Instalação

1. Clone o repositório ou navegue até o diretório do projeto:
   ```bash
   cd project/project
   ```

2. Instale as dependências:
   ```bash
   npm install
   ```

3. Inicie o servidor de desenvolvimento:
   ```bash
   npm start
   ```

4. Acesse o dashboard no seu navegador:
   * **URL local**: `http://localhost:4200/`

---

## 🏗️ Estrutura do Código

* [home.ts](file:///home/rafael/project/project/src/app/home/home.ts): Lógica em TypeScript para importação do Excel, processamento dos arrays, normalização de porcentagens, busca de metadados e cálculos acumulados da Curva S.
* [home.html](file:///home/rafael/project/project/src/app/home/home.html): Estrutura HTML do painel contendo os cards de resumo, tabela de dados mensais, o SVG dinâmico da Curva S e a interface do Diário.
* [home.css](file:///home/rafael/project/project/src/app/home/home.css): Estilos visuais personalizados, incluindo paleta de cores moderna, efeitos de hover, responsividade para diferentes tamanhos de tela e estilização dos gráficos.
