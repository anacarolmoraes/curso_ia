import os
import sys
import time
import json
import base64
import re
from pathlib import Path
from datetime import datetime

# Importações do Selenium
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Erro: As bibliotecas necessárias não estão instaladas no seu ambiente virtual.")
    print("Por favor, execute: .\\venv\\Scripts\\pip install selenium webdriver-manager requests")
    sys.exit(1)

# Configurações de Diretório
BASE_DIR = Path(os.getcwd())
DEST_DIR = BASE_DIR / "diarias" / "diario_oficial"
DEST_DIR.mkdir(parents=True, exist_ok=True)

URL_ALVO = "https://app.tce.to.gov.br/boletim/publico/app/index.php"

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def save_as_pdf(driver, file_path):
    """Utiliza o Chrome DevTools Protocol para imprimir a página atual como PDF."""
    print(f"[*] Gerando PDF: {file_path.name}...")
    
    print_options = {
        'landscape': False,
        'displayHeaderFooter': False,
        'printBackground': True,
        'preferCSSPageSize': True,
    }
    
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    
    with open(file_path, "wb") as f:
        f.write(base64.b64decode(result['data']))
    
    print(f"[+] PDF salvo com sucesso em: {file_path}")

def extrair_e_baixar(numero_alvo=None):
    if numero_alvo:
        print(f"[*] Buscando Boletim Oficial Nº {numero_alvo}...")
    else:
        print(f"[*] Buscando o Boletim Oficial mais recente...")

    driver = setup_driver()
    
    try:
        driver.get(URL_ALVO)
        
        print("[*] Aguardando o carregamento da lista de boletins...")
        wait = WebDriverWait(driver, 20)
        
        # Aguarda a presença dos cartões de boletins
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "btnVerMais")))
        
        # Pega todos os containers de boletins para procurar o número
        boletins = driver.find_elements(By.XPATH, "//div[contains(@class, 'card') or contains(@class, 'row')]")
        
        target_boletim = None
        
        if numero_alvo:
            for b in boletins:
                if f"Nº {numero_alvo}" in b.text:
                    target_boletim = b
                    break
            if not target_boletim:
                print(f"[-] Boletim Nº {numero_alvo} não encontrado na lista atual.")
                return
        else:
            # Pega o primeiro da lista
            target_boletim = boletins[0]

        texto_completo = target_boletim.text
        botao_abrir = target_boletim.find_element(By.CLASS_NAME, "btnVerMais")

        # Regex para extrair Número e Data
        match_numero = re.search(r'N[º°]\s*(\d+)', texto_completo)
        match_data = re.search(r'(\d{2}/\d{2}/\d{4})', texto_completo)
        
        numero_boletim = match_numero.group(1) if match_numero else "0000"
        data_original = match_data.group(1) if match_data else datetime.now().strftime("%d/%m/%Y")
        data_formatada = data_original.replace("/", "-")
        
        print(f"[*] Boletim localizado: Nº {numero_boletim} (Data: {data_original})")
        
        # Clica no botão para abrir os detalhes e pegar o ID
        botao_abrir.click()
        time.sleep(3)
        
        # O subagent viu que uma nova aba aparece com o ID no atributo 'id' do link/tab
        tab_ativa = driver.find_element(By.CSS_SELECTOR, ".nav-link.active[id^='boletim']")
        tab_id_attr = tab_ativa.get_attribute("id") 
        
        boletim_id = re.search(r'boletim(\d+)', tab_id_attr).group(1)
        print(f"[*] ID Interno capturado: {boletim_id}")
        
        url_print = f"https://app.tce.to.gov.br/boletim/app/controllers/?&c=TCE_Boletim_PublicacoesCtrll&m=abrirHtml&id={boletim_id}&print=1"
        
        print(f"[*] Navegando para a versão de impressão: {url_print}")
        driver.get(url_print)
        
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(5)
        
        nome_arquivo = f"BO_{numero_boletim}_{data_formatada}.pdf"
        caminho_final = DEST_DIR / nome_arquivo
        
        save_as_pdf(driver, caminho_final)
        
    except Exception as e:
        print(f"[-] Erro durante o processo: {e}")
        driver.save_screenshot("scratch/error_screenshot.png")
    finally:
        driver.quit()

if __name__ == "__main__":
    print("=== TCE-TO Boletim Downloader (Print-to-PDF Mode) ===")
    num = sys.argv[1] if len(sys.argv) > 1 else None
    extrair_e_baixar(num)
