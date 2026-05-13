# buscar-diarias-tceto

Skill para extrair automaticamente informações de diárias de arquivos Markdown de boletins oficiais do TCE-TO e salvar em Excel.

## Funcionalidades
- Varredura de arquivos `.md` na pasta `diarias/diario_oficial/`.
- Extração de dados de portarias de viagem assinadas pelo Presidente.
- Criação ou atualização (append) do arquivo `diarias_tce_to.xlsx`.
- Movimentação de arquivos processados para `diarias/diario_oficial/processados/`.

## Requisitos
- Python 3.x
- pandas
- openpyxl

## Uso
```bash
python scripts/process_diarias.py
```

## Estrutura de Saída
A planilha Excel contém:
- Portaria
- Nome
- Cargo
- Matrícula
- Data Publicação
- Itinerário
- Período
- Qtd Diárias
- Total Diária
- Processo SEI
