import sys
import subprocess
import os

def main():
    if len(sys.argv) < 2:
        print("Uso: python convert.py <input_pdf> [flags]")
        sys.exit(1)

    input_pdf = sys.argv[1]
    extra_flags = sys.argv[2:]

    # Caminho para o executável do pdfmd no venv
    # No Windows, geralmente está em venv/Scripts/python.exe -m pdfmd.cli
    venv_python = os.path.join(os.getcwd(), "venv", "Scripts", "python.exe")
    
    command = [venv_python, "-m", "pdfmd.cli", input_pdf] + extra_flags
    
    print(f"Executando: {' '.join(command)}")
    
    try:
        result = subprocess.run(command, check=True, capture_output=True, text=True, encoding='utf-8')
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Erro na conversão: {e}")
        print(e.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Erro inesperado: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
