# Curso de IA: Soluções Práticas e Replicáveis

Este repositório contém materiais, scripts e "skills" (habilidades modulares) desenvolvidas para capacitar profissionais na utilização de Inteligência Artificial para automação de tarefas complexas, auditorias e análise de dados, com foco em soluções rápidas, eficientes e fáceis de replicar.

## 🚀 Objetivo do Projeto

O objetivo principal é democratizar o uso de agentes de IA, demonstrando como criar soluções que:
- **São Rápidas:** Implementação imediata para problemas reais.
- **São Replicáveis:** Estrutura modular que permite adaptar a lógica para diferentes contextos.
- **Focam em Resultados:** Automação de processos burocráticos (como extração de diários oficiais e auditoria orçamentária).

## 📂 Estrutura do Repositório

- **`/skills`**: Habilidades modulares prontas para uso (ex: downloaders de boletins, conversores de PDF para Markdown, geradores de apresentações).
- **`/material`**: Prompts estruturados, roteiros de instrutor e kits de apoio para alunos.
- **`/classificacao_orcamentaria`**: Estudos de caso reais e documentos de análise para prática de auditoria assistida por IA.
- **`/scratch`**: Scripts experimentais e ferramentas de suporte para manipulação de dados em Excel e HTML.

## 🛠️ Tecnologias Utilizadas

- **Python 3.12+**: Linguagem base para automações e scripts.
- **Selenium**: Extração de dados de portais dinâmicos.
- **Markdown**: Formato padrão para troca de informações entre agentes de IA e humanos, garantindo eficiência de tokens.
- **Git**: Controle de versão para garantir a rastreabilidade das evoluções do curso.

## ⚙️ Configuração Inicial

Para utilizar as ferramentas deste repositório, recomenda-se o uso do ambiente virtual:

```powershell
# Criação do ambiente virtual
python -m venv venv

# Ativação (Windows)
.\venv\Scripts\activate

# Instalação de dependências (exemplo)
pip install selenium webdriver-manager requests openpyxl
```

## 🧠 Filosofia de Desenvolvimento

Este projeto segue a regra de **"Agente Primeiro"**. Todo o código e documentação são estruturados para que um assistente de IA possa entender o contexto rapidamente (`AI_CONTEXT.md`), utilizar ferramentas predefinidas (`skills/`) e executar tarefas com o mínimo de intervenção humana, focando na tomada de decisão estratégica.

---
*Desenvolvido para transformar a produtividade no setor público e privado através da IA.*
