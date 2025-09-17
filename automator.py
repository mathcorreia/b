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
from selenium.common.exceptions import TimeoutException

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
PASTA_MES = os.path.join(PASTA_BASE, MES_ATUAL)
LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')

PASTA_DESTINO_LM = os.path.join(PASTA_MES, 'LM')
PASTA_DESTINO_LP = os.path.join(PASTA_MES, 'LP')
PASTA_DESTINO_FS = os.path.join(PASTA_MES, 'FS')

os.makedirs(PASTA_DESTINO_LM, exist_ok=True)
os.makedirs(PASTA_DESTINO_LP, exist_ok=True)
os.makedirs(PASTA_DESTINO_FS, exist_ok=True)

# Registra log
def registrar_log(mensagem):
    with open(LOG_PATH, 'a', encoding='utf-8') as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")

def esperar_download_iniciar(pasta_download, arquivos_antes, timeout=10):
    """Verifica se um novo download começou em um curto período."""
    segundos = 0
    while segundos < timeout:
        arquivos_depois = set(os.listdir(pasta_download))
        novos_arquivos = arquivos_depois - arquivos_antes
        if any(f.endswith('.crdownload') or f.endswith('.tmp') for f in novos_arquivos):
            return True # Download iniciado
        time.sleep(1)
        segundos += 1
    return False # Nenhum download iniciado

def esperar_download_concluir(pasta_download, timeout=60):
    """Espera um download em progresso terminar."""
    segundos = 0
    while segundos < timeout:
        if not any(f.endswith('.crdownload') or f.endswith('.tmp') for f in os.listdir(pasta_download)):
            return True # Download concluído
        time.sleep(1)
        segundos += 1
    return False # Timeout

# Le os dados do excel ex: oc/oclinha
try:
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    df.rename(columns={df.columns[0]: 'OS'}, inplace=True)
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
    registrar_log(f"Arquivo Excel lido. {len(df)} itens para processar na pasta '{MES_ATUAL}'.")
except Exception as e:
    registrar_log(f"ERRO CRÍTICO ao ler o Excel: {e}")
    raise

# Configurações do Chrome
options = webdriver.ChromeOptions() 
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR, "download.prompt_for_download": False,
    "download.directory_upgrade": True, "safebrowsing.enabled": True
})
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

# Aguarda o login manual
driver.get("https://web.embraer.com.br/irj/portal")
input("Faça TODA a navegação (Login, GFS, FSE, Busca FSe) manualmente. Quando a tela de busca estiver pronta, pressione ENTER...")

wait = WebDriverWait(driver, 30)
registrar_log("Iniciando processamento do Excel...")

# Loop principal
for index, row in df.iterrows():
    os_num = str(row['OS'])
    oc1 = row['OC_antes']
    oc2 = row['OC_depois']

    try:
        registrar_log(f"--- Processando OS: {os_num} | OC: {oc1}/{oc2} ---")
        
        # Preenchimento e busca inicial
        campo_oc1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']")))
        campo_oc1.clear()
        campo_oc1.send_keys(oc1)

        campo_oc2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']")))
        campo_oc2.clear()
        campo_oc2.send_keys(oc2)

        wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
        
        try:
            registrar_log("Procurando botão: Lista de Materiais (LM)...")
            seletor_lm = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[1]")
            wait.until(EC.element_to_be_clickable(seletor_lm)).click()

            arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
            if esperar_download_iniciar(DOWNLOAD_DIR, arquivos_antes):
                if esperar_download_concluir(DOWNLOAD_DIR):
                    arquivos_depois = set(os.listdir(DOWNLOAD_DIR))
                    novo_arquivo_nome = (arquivos_depois - arquivos_antes).pop()
                    caminho_arquivo = os.path.join(DOWNLOAD_DIR, novo_arquivo_nome)
                    
                    novo_nome_lm = f"{os_num}_LM.pdf"
                    destino_lm = os.path.join(PASTA_DESTINO_LM, novo_nome_lm)
                    if not os.path.exists(destino_lm):
                        shutil.move(caminho_arquivo, destino_lm)
                        registrar_log(f"SUCESSO (LM): Movido e renomeado para {novo_nome_lm}")
                    else:
                        os.remove(caminho_arquivo)
                        registrar_log(f"AVISO (LM): Arquivo já existe. Download duplicado removido.")
                else:
                    registrar_log(f"ERRO (LM): Download iniciado mas não concluído para a OS {os_num}")
            else:
                registrar_log(f"AVISO (LM): Nenhum download iniciado. O documento pode não existir.")
        except TimeoutException:
            registrar_log(f"AVISO (LM): Botão 'Lista de Materiais' não encontrado para a OS {os_num}.")
        except Exception as e:
            registrar_log(f"ERRO inesperado no download de LM para a OS {os_num}: {e}")

        try:
            registrar_log("Procurando botão: Lista de Peças (LP)...")
            seletor_lp = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[2]")
            wait.until(EC.element_to_be_clickable(seletor_lp)).click()
            
            arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
            if esperar_download_iniciar(DOWNLOAD_DIR, arquivos_antes):
                if esperar_download_concluir(DOWNLOAD_DIR):
                    arquivos_depois = set(os.listdir(DOWNLOAD_DIR))
                    novo_arquivo_nome = (arquivos_depois - arquivos_antes).pop()
                    caminho_arquivo = os.path.join(DOWNLOAD_DIR, novo_arquivo_nome)

                    novo_nome_lp = f"{os_num}_LP.pdf"
                    destino_lp = os.path.join(PASTA_DESTINO_LP, novo_nome_lp)
                    if not os.path.exists(destino_lp):
                        shutil.move(caminho_arquivo, destino_lp)
                        registrar_log(f"SUCESSO (LP): Movido e renomeado para {novo_nome_lp}")
                    else:
                        os.remove(caminho_arquivo)
                        registrar_log(f"AVISO (LP): Arquivo já existe. Download duplicado removido.")
                else:
                    registrar_log(f"ERRO (LP): Download iniciado mas não concluído para a OS {os_num}")
            else:
                registrar_log(f"AVISO (LP): Nenhum download iniciado. O documento pode não existir.")
        except TimeoutException:
            registrar_log(f"AVISO (LP): Botão 'Lista de Peças' não encontrado para a OS {os_num}.")
        except Exception as e:
            registrar_log(f"ERRO inesperado no download de LP para a OS {os_num}: {e}")

        try:
            registrar_log("Procurando botão: Ficha de Serviço (FS)...")
            seletor_fs = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[3]/button[2]")
            wait.until(EC.element_to_be_clickable(seletor_fs)).click()

            arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
            if esperar_download_iniciar(DOWNLOAD_DIR, arquivos_antes):
                if esperar_download_concluir(DOWNLOAD_DIR):
                    arquivos_depois = set(os.listdir(DOWNLOAD_DIR))
                    novo_arquivo_nome = (arquivos_depois - arquivos_antes).pop()
                    caminho_arquivo = os.path.join(DOWNLOAD_DIR, novo_arquivo_nome)
                    
                    novo_nome_fs = f"{os_num}_FS.pdf"
                    destino_fs = os.path.join(PASTA_DESTINO_FS, novo_nome_fs)
                    if not os.path.exists(destino_fs):
                        shutil.move(caminho_arquivo, destino_fs)
                        registrar_log(f"SUCESSO (FS): Movido e renomeado para {novo_nome_fs}")
                    else:
                        os.remove(caminho_arquivo)
                        registrar_log(f"AVISO (FS): Arquivo já existe. Download duplicado removido.")
                else:
                    registrar_log(f"ERRO (FS): Download iniciado mas não concluído para a OS {os_num}")
            else:
                registrar_log(f"AVISO (FS): Nenhum download iniciado. O documento pode não existir.")
        except TimeoutException:
            registrar_log(f"AVISO (FS): Botão 'Ficha de Serviço' não encontrado para a OS {os_num}.")
        except Exception as e:
            registrar_log(f"ERRO inesperado no download de FS para a OS {os_num}: {e}")
            
        registrar_log(f"Processo da OS {os_num} concluído. Voltando para a página de busca.")
        driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")

    except Exception as e:
        timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_oc_{str(oc1).replace('/', '-')}_{timestamp_erro}.png"
        driver.save_screenshot(os.path.join(os.getcwd(), nome_screenshot))
        registrar_log(f"ERRO GERAL com OS {os_num}: {e} - Screenshot salvo.")
        input(f"Ocorreu um erro geral com a OS {os_num}. Por favor, coloque na tela de busca novamente e pressione ENTER para continuar...")

registrar_log("Automação finalizada.")
driver.quit()