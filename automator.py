import os
import time
import shutil
import pandas as pd
import locale
import re
import traceback # <-- ADIÇÃO DA LINHA FALTANTE
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- INICIALIZAÇÃO DA INTERFACE GRÁFICA ---
# Cria uma janela raiz oculta para os pop-ups
root = tk.Tk()
root.withdraw()

# Função de log precisa ser definida antes do bloco principal
def registrar_log(mensagem):
    # O LOG_PATH será definido dentro do bloco try principal
    with open(LOG_PATH, 'a', encoding='utf-8') as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")

# --- BLOCO PRINCIPAL COM CAPTURA DE ERRO ---
try:
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
    LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt') # Definido aqui
    ERRO_LOG_PATH = os.path.join(os.getcwd(), 'log_erros.txt')

    PASTA_DESTINO_LM = os.path.join(PASTA_MES, 'LM')
    PASTA_DESTINO_LP = os.path.join(PASTA_MES, 'LP')
    PASTA_DESTINO_FS = os.path.join(PASTA_MES, 'FS')

    os.makedirs(PASTA_DESTINO_LM, exist_ok=True)
    os.makedirs(PASTA_DESTINO_LP, exist_ok=True)
    os.makedirs(PASTA_DESTINO_FS, exist_ok=True)

    def esperar_download_concluir(pasta_download, timeout=60):
        segundos = 0
        for item in os.listdir(pasta_download):
            if item.endswith(".pdf"):
                try:
                    os.remove(os.path.join(pasta_download, item))
                except OSError as e:
                    registrar_log(f"Aviso: Não foi possível limpar o arquivo antigo {item}. Erro: {e}")
        while segundos < timeout:
            if not any(f.endswith('.crdownload') for f in os.listdir(pasta_download)):
                arquivos_pdf = [os.path.join(pasta_download, f) for f in os.listdir(pasta_download) if f.endswith('.pdf')]
                if arquivos_pdf:
                    return max(arquivos_pdf, key=os.path.getmtime)
            time.sleep(1)
            segundos += 1
        return None

    def processar_uma_os(driver, wait, os_num, oc1, oc2):
        # Esta função não será usada nesta versão de teste simplificada
        pass 

    # Le os dados do excel
    registrar_log("Lendo arquivo Excel...")
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    df.rename(columns={df.columns[0]: 'OS'}, inplace=True) 
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
    registrar_log(f"Arquivo Excel lido com sucesso. {len(df)} itens para processar.")

    # Configurações do Chrome
    registrar_log("Configurando o Chrome...")
    options = webdriver.ChromeOptions() 
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR, "download.prompt_for_download": False,
        "download.directory_upgrade": True, "safebrowsing.enabled": True
    })

    # Inicialização do Navegador (PONTO PROVÁVEL DE FALHA)
    registrar_log("Iniciando o WebDriver do Chrome...")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    registrar_log("WebDriver iniciado com sucesso.")

    # Se chegou até aqui, o problema não é na inicialização.
    messagebox.showinfo("Sucesso Parcial", "O navegador foi iniciado com sucesso! O problema deve ocorrer durante a navegação ou o loop. Continue com o processo.")
    
    # O restante da sua lógica iria aqui...
    driver.get("https://web.embraer.com.br/irj/portal")
    # ... etc ...

    driver.quit()
    messagebox.showinfo("Finalizado", "O script de teste foi concluído sem erros de inicialização.")

except Exception as e:
    # --- CAPTURA E EXIBIÇÃO DO ERRO ---
    error_details = traceback.format_exc()
    mensagem_final = f"Ocorreu um erro e a automação não pôde continuar.\n\n" \
                     f"Este erro geralmente acontece se o Microsoft VC++ Redistributable não está instalado ou se o antivírus/firewall está bloqueando o chromedriver.\n\n" \
                     f"DETALHES DO ERRO:\n{error_details}"
    
    # Registra o erro detalhado no log
    try:
        # Tenta definir o LOG_PATH uma última vez se falhou antes
        if 'LOG_PATH' not in locals():
            LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')
        registrar_log("--- ERRO CRÍTICO ---")
        registrar_log(mensagem_final)
    except:
        pass # Se nem o log funcionar, apenas mostra o pop-up
    
    # Mostra o pop-up de erro que não fecha sozinho
    messagebox.showerror("Erro na Automação", mensagem_final)