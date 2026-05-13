# monitor-diarias-tceto

Skill orquestradora que executa o fluxo completo: Download -> Conversão -> Extração -> Excel.

## Fluxo
1. Executa `tce-boletim-downloader` para obter o boletim mais recente.
2. Executa `pdf-to-md` para converter o PDF em texto processável.
3. Executa `buscar-diarias-tceto` para alimentar a planilha Excel e organizar os arquivos.

## Uso
```bash
python scripts/run_monitor.py
```
