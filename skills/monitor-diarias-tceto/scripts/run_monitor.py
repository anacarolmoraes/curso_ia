import os
import subprocess
import sys
import re

# Configurações de caminhos
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
VENV_PYTHON = sys.executable  # Usa o interpretador atual (que já deve ser o do venv)

def run_script(path, args=None):
    """Executa um script python usando o venv do projeto."""
    full_path = os.path.join(BASE_DIR, path)
    # Garante que não estamos duplicando o BASE_DIR se o path já for absoluto
    if not os.path.isabs(full_path):
        full_path = os.path.abspath(full_path)
    cmd = [VENV_PYTHON, full_path]
    if args:
        cmd.extend(args)
    
    print(f"\n>>> Executando: {path}...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print(f"Sucesso: {result.stdout.strip()}")
        return True
    else:
        print(f"Erro ao executar {path}: {result.stderr.strip()}")
        return False

def get_processed_list():
    """Lê a lista de diários já processados do arquivo markdown."""
    log_path = os.path.join(BASE_DIR, "diarias", "diario_oficial", "processados.md")
    if not os.path.exists(log_path):
        return []
    with open(log_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return re.findall(r"- (\d+)", content)

def update_processed_list(bulletin_number):
    """Adiciona um novo número ao log de processados."""
    log_path = os.path.join(BASE_DIR, "diarias", "diario_oficial", "processados.md")
    if not os.path.exists(log_path):
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write("# Diários Processados - TCE-TO\n\n")
    
    processed = get_processed_list()
    if bulletin_number not in processed:
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"- {bulletin_number}\n")

def check_excel_for_bulletin(bulletin_number):
    """Verifica se o boletim já possui dados na planilha Excel."""
    excel_path = os.path.join(BASE_DIR, "diarias", "diario_oficial", "diarias_tce_to.xlsx")
    if not os.path.exists(excel_path):
        return False
    try:
        import pandas as pd
        df = pd.read_excel(excel_path)
        if "Boletim" in df.columns:
            return str(bulletin_number) in df["Boletim"].astype(str).values
    except:
        pass
    return False

def main():
    print("=== Iniciando Monitoramento de Diárias TCE-TO ===")
    
    import re
    # Captura o número do boletim se fornecido como argumento
    bulletin_number = sys.argv[1] if len(sys.argv) > 1 else None
    
    # 1. Verificações de Segurança e Economia de Tokens
    if bulletin_number:
        # Check 1: Já está no Log MD?
        if bulletin_number in get_processed_list():
            print(f"[!] O Boletim {bulletin_number} já consta como PROCESSADO no log (processados.md).")
            # Check 2: Já está no Excel?
            if check_excel_for_bulletin(bulletin_number):
                print(f"[!] AVISO: Já existem dados do Boletim {bulletin_number} na planilha Excel.")
                print("Interrompendo execução para evitar duplicidade e economizar tokens.")
                return

        # Check 3: Já existe o arquivo MD na pasta de processados?
        processed_dir = os.path.join(BASE_DIR, "diarias", "diario_oficial", "processados")
        existing_md = [f for f in os.listdir(processed_dir) if f"BO_{bulletin_number}" in f and f.endswith(".md")]
        if existing_md:
            print(f"[*] Boletim {bulletin_number} encontrado na pasta 'processados'. Re-processando extração...")
            # Move de volta temporariamente para re-processar
            src = os.path.join(processed_dir, existing_md[0])
            dst = os.path.join(BASE_DIR, "diarias", "diario_oficial", existing_md[0])
            import shutil
            shutil.copy2(src, dst)
            if run_script("skills/buscar-diarias-tceto/scripts/process_diarias.py"):
                update_processed_list(bulletin_number)
                print("=== Re-processamento concluído! ===")
            return

    # 2. Fluxo Normal de Download (se não for encontrado localmente)
    download_args = [bulletin_number] if bulletin_number else None
    if not run_script("skills/tce-boletim-downloader/scripts/download.py", download_args):
        print("Falha na etapa de Download. Abortando.")
        return
    
    # Identificar o número do boletim baixado se não foi passado via argumento
    if not bulletin_number:
        pdf_dir = os.path.join(BASE_DIR, "diarias", "diario_oficial")
        pdfs = [f for f in os.listdir(pdf_dir) if f.startswith("BO_") and f.endswith(".pdf")]
        if pdfs:
            match = re.search(r"BO_(\d+)", pdfs[0])
            if match: bulletin_number = match.group(1)

    # 3. Conversão
    pdf_dir = os.path.join(BASE_DIR, "diarias", "diario_oficial")
    pdfs = [f for f in os.listdir(pdf_dir) if f.endswith(".pdf")]
    
    if pdfs:
        for pdf in pdfs:
            pdf_path = os.path.join(pdf_dir, pdf)
            if run_script("skills/pdf-to-md/scripts/convert.py", [pdf_path]):
                # Remover o PDF após converter para evitar re-processamento
                os.remove(pdf_path)
                print(f"PDF {pdf} convertido e removido.")

    # 3. Extração e Excel
    if not run_script("skills/buscar-diarias-tceto/scripts/process_diarias.py"):
        print("Falha na etapa de Extração de Dados.")
        return

    print("\n=== Fluxo concluído com sucesso! ===")

if __name__ == "__main__":
    main()
