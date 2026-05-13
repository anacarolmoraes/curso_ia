---
name: tce-boletim-downloader
description: "Navega no Diário Oficial do Tribunal de Contas do Tocantins, faz o download do boletim mais recente e o salva na pasta padronizada."
category: tool
risk: safe
source: local
tags: "[web-scraping, automation, tce-to, pdf]"
date_added: "2026-05-11"
---

# tce-boletim-downloader

## Purpose

Esta skill automatiza o processo de acessar a página do Boletim Oficial Público do TCE-TO, identificar o boletim mais recente (PDF), fazer o download e salvá-lo na pasta `diarias/diario_oficial` com a nomenclatura padrão: `BO_[numero do boletim]_[data de publicacao].pdf`.

## When to Use This Skill

Use esta skill quando:
- Precisar analisar ou extrair dados do último Boletim Oficial do TCE-TO.
- O usuário solicitar o "download", "baixar", ou "buscar" o último boletim/diário oficial do Tribunal de Contas do Tocantins.

## Core Capabilities

1. **Acesso Dinâmico:** Utiliza Python (Selenium/Requests) para varrer a página do TCE-TO que possui conteúdo carregado via JS.
2. **Download Direto:** Extrai o link direto para o PDF do último boletim.
3. **Padronização:** Salva o arquivo na pasta correta seguindo uma convenção de nomenclatura rigorosa.

## Usage Guide

A skill requer o ambiente virtual do projeto ativado (`venv`), onde o script Python fará a extração.

### Comando Básico

```powershell
.\venv\Scripts\python.exe skills\tce-boletim-downloader\scripts\download.py
```

## Regras de Execução

1. **Acesso à Pasta:** O PDF deve ser sempre armazenado no diretório `diarias/diario_oficial/`. Se o diretório não existir, deve ser criado.
2. **Nomenclatura:** O arquivo baixado deve ter a sintaxe `BO_[numero do boletim]_[data de publicacao].pdf`. Exemplo: `BO_3948_11-05-2026.pdf`.
3. **Erros:** Em caso de indisponibilidade do site, o script deve retornar uma mensagem clara do erro.

## Limitations

- Como a página usa ExtJS, a extração depende de interceptação de rede ou renderização da página. Pode ser necessário manter bibliotecas como `selenium` ou similares instaladas no `venv`.
