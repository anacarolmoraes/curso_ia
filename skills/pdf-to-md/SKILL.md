---
name: pdf-to-md
description: "Transforma arquivos PDF em Markdown limpo e estruturado utilizando a ferramenta pdfmd local. Suporta extração de texto, tabelas, fórmulas matemáticas (LaTeX) e OCR para documentos digitalizados."
category: tool
risk: safe
source: local
tags: "[pdf, markdown, conversion, ocr, latex]"
date_added: "2026-05-08"
---

# pdf-to-md

## Purpose

Fornecer uma interface padronizada para converter documentos PDF em Markdown, preservando a estrutura original, tabelas e equações. Esta skill utiliza o projeto `pdfmd` instalado no ambiente virtual do projeto para garantir privacidade total e processamento offline.

## When to Use This Skill

Use esta skill quando:
- O usuário fornecer um arquivo PDF e precisar do conteúdo em texto para análise ou edição.
- For necessário extrair tabelas complexas de PDFs acadêmicos ou financeiros.
- O PDF contiver fórmulas matemáticas que devem ser preservadas em formato LaTeX.
- O PDF for uma imagem/digitalização e exigir OCR para ser lido.
- Você precisar preparar o conteúdo de um PDF para ser consumido por um LLM.

## Core Capabilities

1. **Conversão Nativa** - Extração rápida de texto de PDFs digitais.
2. **Reconstrução de Tabelas** - Detecção automática e formatação de tabelas em GFM (Pipe Tables).
3. **Preservação Matemática** - Normalização de Unicode para LaTeX em equações.
4. **Suporte a OCR** - Integração com Tesseract para documentos digitalizados (necessita Tesseract instalado).
5. **Extração de Imagens** - Exportação de imagens para uma pasta de assets.

## Usage Guide

### 1. Preparação
Certifique-se de que o ambiente virtual `venv` está ativo ou use o caminho direto para o executável.

### 2. Comandos Básicos

```powershell
# Conversão simples (gera arquivo .md na mesma pasta do PDF)
.\venv\Scripts\python.exe -m pdfmd.cli "caminho/do/arquivo.pdf"

# Conversão com detecção automática de OCR (se o PDF for imagem)
.\venv\Scripts\python.exe -m pdfmd.cli "caminho/do/arquivo.pdf" --ocr auto

# Exportação de imagens e estatísticas
.\venv\Scripts\python.exe -m pdfmd.cli "caminho/do/arquivo.pdf" --export-images --stats
```

### 3. Perfis Recomendados (Contextuais)

- **Artigos Acadêmicos**: Use `--stats` e verifique as fórmulas matemáticas.
- **Slides**: Use `--export-images` e `--page-breaks`.
- **Digitalizações**: Use `--ocr auto --lang por` (para português).

## Regras de Execução

1. **Privacidade**: Nunca envie o PDF para serviços externos. Todo o processamento deve ser via `pdfmd` local.
2. **Saída**: Por padrão, o arquivo Markdown deve ser gerado na mesma pasta ou em uma subpasta `output/` se houver muitos arquivos.
3. **Pós-processamento**: Após a conversão, leia o arquivo `.md` gerado para confirmar a qualidade da extração antes de responder ao usuário.

## Limitations

- A qualidade do OCR depende da instalação local do Tesseract e pacotes de linguagem.
- PDFs com layouts extremamente complexos (ex: revistas com muitas colunas sobrepostas) podem exigir revisão manual do Markdown gerado.
- Requer que a pasta `pdfmd/` na raiz do projeto seja mantida para funcionamento correto do modo editável.
