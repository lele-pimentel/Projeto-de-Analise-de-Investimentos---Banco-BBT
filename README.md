# **PROJETO DE ANÁLISE DE INVESTIMENTOS - BANCO BBT**

Este projeto demonstra o fluxo completo de dados, desde o tratamento, organização em ETL e automação em Excel até a criação de um dashboard estratégico em Power BI, simulando o cenário real de uma mesa de investimentos.

### **DESAFIO**
Como analista, recebi bases de dados fragmentadas, vazias, com erros e desogarnizadas, o que torna o acompanhamento de rentabilidade, lucro e perda propenso a erros manuais e dificulta a visualização de oportunidades ou riscos na carteira do cliente.

### **OBJETIVO**
Demonstrar o fluxo completo de dados de uma carteira de investimentos, desde o tratamento e automação em Excel até a criação de um dashboard analítico em Power BI, garantindo integridade de dados e suporte à tomada de decisão.

### **TECNOLOGIAS UTILIZADAS**
* **Excel** (Fórmulas PROCV, SOMA, SE, MEDIA, MAXIMO, MINIMO e Tabela Dinâmica)
* **VBA/Macro** (Automação de processos e dominio da linguagem)
* **Power BI** (Visualização de dados)
* **Power Query** (Processo de ETL)
* **DAX** (Cálculos do BI e função condicional)

---

### **ETAPA 1: EXCEL AVANÇADO & VBA**

Nesta etapa, transformei dados brutos em uma ferramenta operacional funcional:

* **Tratamento:** Padronização de dados, tranformar intervalos em tabelas estruturadas, identificar valores nulos e com erros para evitar falhas quando usar como referência.
* **Inteligência:** Aplicação de funções avançadas (PROCV, SE, MÉDIA e SOMA) para cálculo de performance.
* **Automação de Média de Alunos (VBA):** Desenvolvimento de um script para entrada de notas via `InputBox`, cálculo automático de médias usando formatação condicional (Verde para aprovação e Vermelho para reprovação) para treino e dominio da linguagem de automação de ponta a ponta
* **UX/UI:** Lista suspensa interativa por ativo exibindo informações, tabela condicional visualizando a porcentagem do valor atual, em R$ e seu total

---

**Visualização:**

<img width="1919" height="1009" alt="image" src="https://github.com/user-attachments/assets/8c303ba7-f273-4775-b2bd-9dd48d37800a" />

---

### **ETAPA 2: Dashboard (POWER BI)**

Aqui, os dados foram elevados a um nível analítico e visual:

* **ETL (Power Query):** Criação de colunas condicionais para segmentação de rentabilidade e classifcação de tipo de colunas em numero inteiros ou decimais, texto, em R$ e procentagem
* **Cálculos DAX:** Realizar medidas para Rentabilidade %, Lucro Total, Perda total e análises (Máx/Mín).
* **Visualização:** Dashboard interativo com navegação por botões, filtros dinâmicos e análise por categoria de ativo e tipo.
  
---

**Visualização:**

<img width="1437" height="804" alt="image" src="https://github.com/user-attachments/assets/9f41b4fc-9e2a-4946-b0a4-f78c58bb6690" />
<img width="1304" height="747" alt="image" src="https://github.com/user-attachments/assets/a6600a0f-3721-4647-a461-aadf0f82eb6d" />

---

### **RESULTADOS OBTIDOS**
* **Eficiência:** Reduz o tempo de tratamento de dados com o uso de Power Query e VBA.
* **O que esta acontecendo?:** Identifica instantaneamente quais ativos estão performando acima ou abaixo da meta.
* **Dashboard interativo e funcional:** Transforma dados brutos em informações visuais claras, permitindo o acompanhamento eficiente dos investimentos e apoiando a tomada de decisão.
* **Planilha funcional e profissional com busca interativa e visualização clara da performance da carteira, com integração com macro**
* **VBA com interação ao usuario para calculo de media de nota, para dominio da linguagem e adminstração do codigo para tarefas futuras**
---



