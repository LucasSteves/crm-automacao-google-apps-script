# 🎓 CRM Acadêmico & Automação de Retenção

Este repositório contém um ecossistema de automação desenvolvido em **JavaScript (Google Apps Script)** para otimizar a gestão de permanência, o tratamento de dados e o fluxo de comunicação entre setores de uma instituição de ensino. 

O sistema foi arquitetado para lidar com bases de dados volumosas (mais de 1.000 registros simultâneos), substituindo o cruzamento de dados manual por execuções em tempo real e mitigando 100% dos erros operacionais.

## 🚀 Impacto do Projeto
- **Escalabilidade:** Processamento automatizado e estruturação de grandes volumes de dados educacionais.
- **Eficiência Operacional:** Redução do tempo de execução de relatórios de dias/meses para **segundos**.
- **Precisão Sistêmica:** Implementação de lógicas de análise de risco e higienização de strings para evitar falhas humanas.

---

## 🏗️ Arquitetura do Sistema

O projeto foi dividido em três módulos independentes, respeitando o princípio de responsabilidade única e a estrutura das planilhas utilizadas pela operação:

### 1. Módulo de Gestão de Permanência (`gestao-permanencia.js`)
Focado na análise de risco e no primeiro contato com o aluno.
* **Algoritmo "Farol de Risco":** Calcula a criticidade do aluno com base em múltiplas variáveis (ausência >= 20%, requisições financeiras, mensalidades vencidas).
* **Normalização de Dados:** Padronização de nomenclaturas de cursos através de mapeamento em dicionários de dados.
* **Automação de CRM:** Disparo dinâmico de e-mails customizados de acordo com as necessidades específicas de cada aluno (financeiro, acadêmico, faltas).

### 2. Módulo de Pós-Vendas e Higienização (`automacao-pos-vendas.js`)
Focado em limpar "sujeiras" de extração de sistemas legados e gerar relatórios unificados.
* **Higienização de Banco de Dados:** Correção automática de células mescladas e preenchimento inteligente de dados faltantes (memória de variáveis).
* **Estruturação de Relatórios:** Organização dinâmica de arrays para categorizar os alunos por "Status" (Ingressantes, Evasão, Prouni).
* **Geração de Arquivos:** Conversão da base de dados em arquivos unificados via API.

### 3. Módulo de Envio em Lote (`envio-diario-gestores.js`)
Focado na comunicação entre a equipe de retenção e os gestores/coordenadores acadêmicos.
* **Motor de Expressões Regulares (Regex):** Direcionamento preciso de informações. Lê o curso do aluno, busca o gestor responsável em um dicionário usando Regex e agrupa os dados corretamente.
* **Criação Dinâmica de Excel (`.xlsx`):** Geração temporária de planilhas formatadas, conversão para `.xlsx` via requisição (UrlFetchApp / Drive API) e anexo no e-mail.
* **Envio Assíncrono:** Disparo massivo automatizado para dezenas de gestores em poucos cliques, deletando os arquivos temporários da nuvem logo em seguida.

---

## 🛠️ Tecnologias e Conceitos Aplicados
* **Linguagem:** JavaScript (ES6+), Google Apps Script (GAS)
* **APIs Integradas:** Google Sheets API, GmailApp, MailApp, Drive API, UrlFetchApp
* **Conceitos de Programação:**
  * Manipulação de Arrays e Objetos (Map, Filter, Sort, Reduce)
  * Expressões Regulares (Regex) e Higienização de Strings
  * Lógica de Negócios Complexa (IF/Else Aninhados)
  * Automação de Processos Robóticos (RPA)

---
*Projeto desenvolvido por [Lucas Steves] como demonstração de resolução de problemas reais de negócios utilizando tecnologia.*
