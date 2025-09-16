import os
import time
import shutil
import pandas as pd
import locale
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Config geral
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil')

hoje = datetime.now()
nome_mes_atual = hoje.strftime("%B").capitalize()
num_mes_atual = hoje.month + 100

DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'Downloads')
PASTA_BASE = r'\\fserver\cedoc_docs\Doc - EmbraerProdutivo\2025'
MES_ATUAL = f'{num_mes_atual} - {nome_mes_atual}'
PASTA_DESTINO = os.path.join(PASTA_BASE, MES_ATUAL)
LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')

os.makedirs(PASTA_DESTINO, exist_ok=True)

# Registra log
def registrar_log(mensagem):
    with open(LOG_PATH, 'a', encoding='utf-8') as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")

def esperar_download_concluir(pasta_download, timeout=60):
    segundos = 0
    while segundos < timeout:
        if not any(f.endswith('.crdownload') for f in os.listdir(pasta_download)):
            arquivos_pdf = [os.path.join(pasta_download, f) for f in os.listdir(pasta_download) if f.endswith('.pdf')]
            if arquivos_pdf:
                return max(arquivos_pdf, key=os.path.getmtime)
        time.sleep(1)
        segundos += 1
    return None

# Le os dados do excel ex: oc/oclinha
try:
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
    registrar_log(f"Arquivo Excel lido. {len(df)} itens para processar na pasta '{MES_ATUAL}'.")
except Exception as e:
    registrar_log(f"ERRO CRÍTICO ao ler o Excel: {e}")
    raise

# Configurações padronizada do Chrome
options = webdriver.ChromeOptions() 
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

# Aguarda o login manual
driver.get("https://web.embraer.com.br/irj/portal")
input("Faça TODA a navegação (Login, GFS, FSE, Busca FSe) manualmente. Quando a tela de busca estiver pronta, pressione ENTER...")

wait = WebDriverWait(driver, 30)

# O script começa a trabalhar diretamente no loop de processamento
registrar_log("Iniciando processamento do Excel...")

# Loop de buscar e realizar download
for index, row in df.iterrows():
    oc1 = row['OC_antes']
    oc2 = row['OC_depois']

    try:
        registrar_log(f"Processando OC: {oc1}/{oc2}")
        
        campo_oc1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']")))
        campo_oc1.clear()
        campo_oc1.send_keys(oc1)

        campo_oc2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']")))
        campo_oc2.clear()
        campo_oc2.send_keys(oc2)

        wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()

        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
        
        # --- CORREÇÃO 1: ESPERA MAIS ROBUSTA E PAUSA ESTRATÉGICA ---
        # Pausa de 1 segundo para dar tempo da nova página "se assentar"
        time.sleep(1) 
        lista_materiais_wait = WebDriverWait(driver, 30)
        lista_materiais_btn = lista_materiais_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Lista de Materiais')]")))
        lista_materiais_btn.click()
        # -----------------------------------------------------------

        caminho_arquivo_baixado = esperar_download_concluir(DOWNLOAD_DIR)

        if caminho_arquivo_baixado:
            nome_arquivo = os.path.basename(caminho_arquivo_baixado)
            destino = os.path.join(PASTA_DESTINO, nome_arquivo)

            if not os.path.exists(destino):
                shutil.move(caminho_arquivo_baixado, destino)
                registrar_log(f"Movido: {nome_arquivo} para {PASTA_DESTINO}")
            else:
                os.remove(caminho_arquivo_baixado)
                registrar_log(f"Arquivo já existe no destino: {nome_arquivo}. Download duplicado removido.")
        else:
            registrar_log(f"ERRO: Download não concluído a tempo para a OC {oc1}/{oc2}")

        # --- CORREÇÃO 2: VOLTAR PARA A PÁGINA DE BUSCA PARA CONTINUAR O LOOP ---
        registrar_log("Retornando à página de busca para a próxima OC...")
        driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
        # -------------------------------------------------------------------

    except Exception as e:
        timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_oc_{str(oc1).replace('/', '-')}_{timestamp_erro}.png"
        caminho_screenshot = os.path.join(os.getcwd(), nome_screenshot)
        try:
            driver.save_screenshot(caminho_screenshot)
            registrar_log(f"ERRO com OC {oc1}/{oc2}: {e} - Screenshot salvo em: {caminho_screenshot}")
        except Exception as screenshot_error:
            registrar_log(f"ERRO com OC {oc1}/{oc2}: {e} - FALHA AO SALVAR SCREENSHOT: {screenshot_error}")
        
        # A instrução de recuperação foi mantida
        input(f"Ocorreu um erro com a OC {oc1}/{oc2}. Por favor, coloque na tela de busca novamente e pressione ENTER para continuar...")

registrar_log("Automação finalizada.")
driver.quit()