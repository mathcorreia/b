import os
import time
import shutil
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Config geral
try:
    DOWNLOAD_DIR = os.path.join(os.environ['USERPROFILE'], 'Downloads')
except KeyError:
    DOWNLOAD_DIR = os.path.expanduser("~/Downloads")

PASTA_BASE = r'\\fserver\cedoc_docs\Doc - EmbraerProdutivo\2025'
# MANTIDO O SEU MÊS ORIGINAL PARA GARANTIR QUE FUNCIONE
MES_ATUAL = '110 - Setembro' 
PASTA_DESTINO = os.path.join(PASTA_BASE, MES_ATUAL)
LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')

# Registra log
def registrar_log(mensagem):
    with open(LOG_PATH, 'a', encoding='utf-8') as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")

# Le os dados do excel ex: oc/oclinha
try:
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
except Exception as e:
    registrar_log(f"Erro ao ler o Excel: {e}")
    raise

# Configurações padronizada do Edge
options = Options()
options.use_chromium = True
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

service = Service()  # EdgeDriver deve estar no PATH
driver = webdriver.Edge(service=service, options=options)

# Aguarda o login manual
driver.get("https://web.embraer.com.br/")
input("Faça login manualmente, entre na página para baixar notas e pressione ENTER para continuar...")

# Define um objeto de espera com tempo maior para mais segurança
wait = WebDriverWait(driver, 20)

# Loop de buscar e realizar download
for index, row in df.iterrows():
    oc1 = row['OC_antes']
    oc2 = row['OC_depois']

    try:
        # --- INÍCIO DA MUDANÇA MÍNIMA E SEGURA ---

        # Preencher campos OC
        campo_oc1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']")))
        campo_oc1.clear()
        campo_oc1.send_keys(oc1)

        campo_oc2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']")))
        campo_oc2.clear()
        campo_oc2.send_keys(oc2)

        # Clicar em Buscar
        wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()

        # Esperar e clicar na lupa (botão de detalhes da busca)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()

        # Clicar em Lista de Materiais
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Lista de Materiais')]"))).click()

        # --- FIM DA MUDANÇA ---
        
        # Tempo de espera do download/ajustar dependendo do tamanho do arquivo (SEU CÓDIGO ORIGINAL)
        time.sleep(5)  

        # Move o pdf para pasta definida em cedoc_docs (SEU CÓDIGO ORIGINAL)
        arquivos = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.pdf')]
        for arquivo in arquivos:
            origem = os.path.join(DOWNLOAD_DIR, arquivo)
            destino = os.path.join(PASTA_DESTINO, arquivo)

            if not os.path.exists(destino):
                shutil.move(origem, destino)
                registrar_log(f"Movido: {arquivo} para {PASTA_DESTINO}")
            else:
                registrar_log(f"Arquivo já existe: {arquivo}")

    except Exception as e:
        registrar_log(f"Erro com OC {oc1}/{oc2}: {e}")
        driver.refresh() # Tenta recarregar a página para recuperar
        time.sleep(3)

# Verifica o mês, se virou o mês, cria pasta do mês seguinte (SUA LÓGICA ORIGINAL)
try:
    # Adicionado try-except para evitar erro caso o nome do mês não seja encontrado
    mes_atual_data = datetime.strptime(MES_ATUAL.split(" - ")[1], "%B")
    mes_seguinte_data = mes_atual_data.replace(day=1) + pd.DateOffset(months=1)
    mes_seguinte_num = mes_seguinte_data.month + 100
    mes_seguinte_nome = mes_seguinte_data.strftime("%B")
    nova_pasta = os.path.join(PASTA_BASE, f"{mes_seguinte_num} - {mes_seguinte_nome}")

    if not os.path.exists(nova_pasta):
        os.makedirs(nova_pasta)
        registrar_log(f"Criada nova pasta: {nova_pasta}")

    # Copia os arquivos para a nova pasta (SUA LÓGICA ORIGINAL)
    for arquivo in os.listdir(PASTA_DESTINO):
        origem = os.path.join(PASTA_DESTINO, arquivo)
        destino = os.path.join(nova_pasta, arquivo)
        if not os.path.exists(destino):
            shutil.copy2(origem, destino)
            registrar_log(f"Copiado para nova pasta: {arquivo}")
except Exception as e:
    registrar_log(f"Aviso: Não foi possível processar a lógica de virada de mês. Erro: {e}")


registrar_log("Automação finalizada.")
driver.quit()