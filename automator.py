import os
import time
import shutil
import pandas as pd
import locale
import re
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- INICIALIZAÇÃO DA INTERFACE GRÁFICA ---
# Cria uma janela raiz oculta para os pop-ups
root = tk.Tk()
root.withdraw()
# -----------------------------------------

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
ERRO_LOG_PATH = os.path.join(os.getcwd(), 'log_erros.txt') # Novo log de erros

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

# --- FUNÇÃO DEDICADA PARA PROCESSAR UMA OS ---
def processar_uma_os(driver, wait, os_num, oc1, oc2):
    try:
        registrar_log(f"--- Processando OS: {os_num} | OC: {oc1}/{oc2} ---")
        
        campo_oc1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']")))
        campo_oc1.clear()
        campo_oc1.send_keys(oc1)

        campo_oc2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']")))
        campo_oc2.clear()
        campo_oc2.send_keys(oc2)

        wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
        
        # Download 1: Lista de Materiais (LM)
        try:
            time.sleep(1)
            seletor_lm = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[1]")
            wait.until(EC.element_to_be_clickable(seletor_lm)).click()
            caminho_lm = esperar_download_concluir(DOWNLOAD_DIR)
            if caminho_lm:
                novo_nome_lm = f"{os_num}_LM.pdf"
                destino_lm = os.path.join(PASTA_DESTINO_LM, novo_nome_lm)
                shutil.move(caminho_lm, destino_lm)
                registrar_log(f"SUCESSO (LM): Arquivo salvo como {novo_nome_lm}")
            else:
                registrar_log(f"ERRO (LM): Download não concluído para a OS {os_num}")
        except TimeoutException:
            registrar_log(f"AVISO (LM): Botão 'Lista de Materiais' não encontrado para a OS {os_num}.")
        
        # Download 2: Lista de Peças (LP)
        try:
            time.sleep(2)
            seletor_lp = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[2]")
            wait.until(EC.element_to_be_clickable(seletor_lp)).click()
            caminho_lp = esperar_download_concluir(DOWNLOAD_DIR)
            if caminho_lp:
                novo_nome_lp = f"{os_num}_LP.pdf"
                destino_lp = os.path.join(PASTA_DESTINO_LP, novo_nome_lp)
                shutil.move(caminho_lp, destino_lp)
                registrar_log(f"SUCESSO (LP): Arquivo salvo como {novo_nome_lp}")
            else:
                registrar_log(f"ERRO (LP): Download não concluído para a OS {os_num}")
        except TimeoutException:
            registrar_log(f"AVISO (LP): Botão 'Lista de Peças' não encontrado para a OS {os_num}.")

        # Download 3: Ficha de Serviço (FS)
        try:
            time.sleep(2)
            seletor_fs = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[3]/button[2]")
            wait.until(EC.element_to_be_clickable(seletor_fs)).click()
            caminho_fs = esperar_download_concluir(DOWNLOAD_DIR)
            if caminho_fs:
                novo_nome_fs = f"{os_num}_FS.pdf"
                destino_fs = os.path.join(PASTA_DESTINO_FS, novo_nome_fs)
                shutil.move(caminho_fs, destino_fs)
                registrar_log(f"SUCESSO (FS): Arquivo salvo como {novo_nome_fs}")
            else:
                registrar_log(f"ERRO (FS): Download não concluído para a OS {os_num}")
        except TimeoutException:
            registrar_log(f"AVISO (FS): Botão 'Ficha de Serviço' não encontrado para a OS {os_num}.")
        
        registrar_log(f"Processo da OS {os_num} concluído. Voltando para a página de busca.")
        driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
        return True

    except Exception as e:
        timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_os_{os_num}_{timestamp_erro}.png"
        driver.save_screenshot(os.path.join(os.getcwd(), nome_screenshot))
        registrar_log(f"ERRO GERAL com OS {os_num}: {e} - Screenshot salvo.")
        driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
        return False

# --- SCRIPT PRINCIPAL ---
# Le os dados do excel ex: oc/oclinha
try:
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    df.rename(columns={df.columns[0]: 'OS'}, inplace=True) 
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
    registrar_log(f"Arquivo Excel lido. {len(df)} itens para processar na pasta '{MES_ATUAL}'.")
except Exception as e:
    registrar_log(f"ERRO CRÍTICO ao ler o Excel: {e}")
    messagebox.showerror("Erro Crítico", f"Não foi possível ler o arquivo 'lista.xlsx'.\nVerifique se o arquivo existe e se o nome da aba é 'baixar_lm'.\n\nErro: {e}")
    raise

# Configurações padronizada do Chrome
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR, "download.prompt_for_download": False,
    "download.directory_upgrade": True, "safebrowsing.enabled": True
})
driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)

# Aguarda o login manual
driver.get("https://web.embraer.com.br/irj/portal")
messagebox.showinfo("Ação Necessária", "Faça o login e, quando estiver na página principal do portal, clique em OK para continuar.")

wait = WebDriverWait(driver, 30)

try:
    # --- ETAPA DE NAVEGAÇÃO SEMI-AUTOMÁTICA ---
    registrar_log("Iniciando navegação para GFS...")
    original_window = driver.current_window_handle
    wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
    registrar_log("Clicou no link 'GFS'.")

    wait.until(EC.number_of_windows_to_be(2))
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            break
    registrar_log("Foco alterado para a nova aba da aplicação GFS.")

    messagebox.showinfo("Ação Necessária", "Robô na aba correta.\n\nAGORA, clique em 'FSE' > 'Busca FSe' e, quando a tela de busca carregar, clique em OK para o robô começar a trabalhar.")
    
    # Loop de processamento principal
    registrar_log("Iniciando processamento principal do Excel...")
    for index, row in df.iterrows():
        os_num = str(row['OS'])
        oc1 = row['OC_antes']
        oc2 = row['OC_depois']
        processar_uma_os(driver, wait, os_num, oc1, oc2)

    # --- BLOCO DE REPROCESSAMENTO DE ERROS ---
    registrar_log("--- Fim do processamento principal. Verificando erros para reprocessar. ---")
    erros_os = set()
    linhas_de_erro = []
    try:
        with open(LOG_PATH, 'r', encoding='utf-8') as log_file:
            for linha in log_file:
                if "ERRO" in linha or "AVISO" in linha:
                    linhas_de_erro.append(linha)
                    match = re.search(r'OS (\d+)', linha)
                    if match:
                        erros_os.add(match.group(1))
    except FileNotFoundError:
        registrar_log("Arquivo de log não encontrado. Nenhum item para reprocessar.")

    if linhas_de_erro:
        with open(ERRO_LOG_PATH, 'w', encoding='utf-8') as erro_log_file:
            erro_log_file.write(f"--- Resumo de Erros e Avisos da execução de {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n\n")
            erro_log_file.writelines(linhas_de_erro)
        registrar_log(f"Log de erros filtrado foi salvo em: {ERRO_LOG_PATH}")

    if not erros_os:
        messagebox.showinfo("Automação Concluída", "Nenhum erro encontrado na primeira passagem. Processo finalizado com sucesso!")
    else:
        registrar_log(f"Encontrados {len(erros_os)} itens com erro para reprocessar: {', '.join(sorted(erros_os))}")
        
        resposta = messagebox.askyesno("Reprocessamento de Erros", f"Foram encontrados {len(erros_os)} itens com erros ou avisos.\n\nO log de erros foi salvo em 'log_erros.txt'.\n\nDeseja tentar baixar novamente os itens com erro?")
        
        if resposta:
            df_erros = df[df['OS'].astype(str).isin(erros_os)]
            registrar_log("--- Iniciando reprocessamento dos erros. ---")
            for index, row in df_erros.iterrows():
                os_num = str(row['OS'])
                oc1 = row['OC_antes']
                oc2 = row['OC_depois']
                processar_uma_os(driver, wait, os_num, oc1, oc2)
            registrar_log("--- Fim do reprocessamento. ---")
            messagebox.showinfo("Automação Concluída", "Reprocessamento finalizado. Verifique os logs para mais detalhes.")
        else:
            registrar_log("Reprocessamento ignorado pelo usuário.")
            messagebox.showinfo("Automação Concluída", "Processo finalizado. Alguns itens apresentaram erros e não foram reprocessados.")

except Exception as e:
     registrar_log(f"ERRO CRÍTICO fora do loop principal: {e}")
     try:
        timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_critico_{timestamp_erro}.png"
        driver.save_screenshot(os.path.join(os.getcwd(), nome_screenshot))
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro grave e a automação será encerrada.\n\nVerifique o arquivo 'log_automacao.txt'.\n\nErro: {e}")
     except:
         pass

registrar_log("Automação finalizada.")
driver.quit()