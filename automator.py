import os
import time
import shutil
import pandas as pd
import locale
import re
import traceback
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- NOVA JANELA DE ALERTA VISUAL ---
class AlertaVisual:
    def __init__(self, title, message, color="green"):
        self.window = tk.Toplevel()
        self.window.title(title)
        self.window.geometry("450x170")
        self.window.resizable(False, False)
        
        # --- A MÁGICA PARA FICAR SEMPRE NA FRENTE ---
        self.window.attributes('-topmost', True)
        
        # Centraliza a janela
        self.window.eval('tk::PlaceWindow . center')
        
        label = tk.Label(self.window, text=message, font=("Arial", 14, "bold"), wraplength=420, fg=color)
        label.pack(expand=True, fill="both", padx=15, pady=15)
        
        ok_button = tk.Button(self.window, text="OK", command=self.window.destroy, width=15, font=("Arial", 10, "bold"))
        ok_button.pack(pady=(0, 15))
        
        # Garante que o foco vá para a janela
        self.window.focus_force()
        # Impede que o script continue até a janela ser fechada
        self.window.wait_window()

# Oculta a janela raiz principal do Tkinter que não será usada
root = tk.Tk()
root.withdraw()

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
    PASTA_RAIZ_VERIFICACAO = r'\\fserver\cedoc_docs\Doc - EmbraerProdutivo'
    PASTA_BASE_ANO_ATUAL = os.path.join(PASTA_RAIZ_VERIFICACAO, str(hoje.year))
    MES_ATUAL = f'{num_mes_atual} - {nome_mes_atual}'
    PASTA_MES = os.path.join(PASTA_BASE_ANO_ATUAL, MES_ATUAL)
    LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')
    ERRO_LOG_PATH = os.path.join(os.getcwd(), 'log_erros.txt')

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
            
            # Lista de Materiais (LM)
            try:
                time.sleep(1)
                seletor_lm = (By.XPATH, "//button[contains(., 'Lista de Materiais')]")
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
            
            # Lista de Peças (LP)
            try:
                time.sleep(2)
                seletor_lp = (By.XPATH, "//button[contains(., 'Lista de Peças')]")
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

            # Ficha de Serviço (FS)
            try:
                time.sleep(2)
                seletor_fs = (By.XPATH, "//button[contains(., 'Imprimir')]")
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
            registrar_log(f"ERRO GERAL com OS {os_num}: {e}")
            try:
                driver.save_screenshot(os.path.join(os.getcwd(), nome_screenshot))
                registrar_log(f"Screenshot salvo como {nome_screenshot}")
            except Exception as screenshot_error:
                registrar_log(f"FALHA AO SALVAR SCREENSHOT: {screenshot_error}")
            
            driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return False

    # Le os dados do excel
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    df.rename(columns={df.columns[0]: 'OS'}, inplace=True) 
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
    registrar_log(f"Arquivo Excel lido. Total de {len(df)} itens na lista.")

    # Verificação de Duplicidade Retroativa
    registrar_log("Iniciando verificação retroativa de duplicidade (até 2 anos). Isso pode levar alguns minutos...")
    arquivos_existentes = set()
    data_limite = hoje - pd.DateOffset(years=2)
    if os.path.exists(PASTA_RAIZ_VERIFICACAO):
        for root_dir, dirs, files in os.walk(PASTA_RAIZ_VERIFICACAO):
            nome_da_pasta_atual = os.path.basename(root_dir)
            try:
                if len(nome_da_pasta_atual) == 4 and nome_da_pasta_atual.isdigit():
                    if int(nome_da_pasta_atual) < data_limite.year:
                        dirs[:] = []
                        continue
            except ValueError:
                pass
            for nome_arquivo in files:
                if nome_arquivo.endswith(".pdf"):
                    os_num = nome_arquivo.split('_')[0]
                    if os_num.isdigit():
                        arquivos_existentes.add(os_num)
    
    if arquivos_existentes:
        df['OS'] = df['OS'].astype(str)
        df_original_len = len(df)
        df = df[~df['OS'].isin(arquivos_existentes)]
        removidos = df_original_len - len(df)
        registrar_log(f"Verificação concluída. {removidos} OSs foram removidas da lista por já terem sido baixadas.")
    else:
        registrar_log("Verificação concluída. Nenhuma duplicidade encontrada.")
    registrar_log(f"Total de {len(df)} itens restantes para processar.")
    
    if df.empty:
        messagebox.showinfo("Nenhum Item a Processar", "Todos os itens da lista já foram baixados anteriormente. Automação finalizada.")
    else:
        # Configurações do Navegador
        caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
        service = ChromeService(executable_path=caminho_chromedriver)
        options = webdriver.ChromeOptions() 
        options.add_argument("--start-maximized")
        options.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        driver = webdriver.Chrome(service=service, options=options)

        # Login e Navegação
        driver.get("https://web.embraer.com.br/irj/portal")
        AlertaVisual("Ação Necessária (1/2)", "Faça o login no portal e, quando a página principal carregar, clique em OK.")

        wait = WebDriverWait(driver, 30)
        
        # ETAPA DE NAVEGAÇÃO SEMI-AUTOMÁTICA
        registrar_log("Iniciando navegação para GFS e trocando de aba...")
        original_window = driver.current_window_handle
        wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                break
        registrar_log("Foco alterado para a nova aba da aplicação GFS.")
        
        AlertaVisual("Ação Necessária (2/2)", "Robô na aba correta.\n\nAGORA, clique em 'FSE' > 'Busca FSe' e, quando a tela de busca carregar, clique em OK.")

        # Loop de processamento principal
        registrar_log("Iniciando processamento principal do Excel...")
        for index, row in df.iterrows():
            os_num = str(row['OS'])
            oc1 = row['OC_antes']
            oc2 = row['OC_depois']
            processar_uma_os(driver, wait, os_num, oc1, oc2)

        # BLOCO DE REPROCESSAMENTO DE ERROS
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
            resposta = messagebox.askyesno("Reprocessamento de Erros", f"Foram encontrados {len(erros_os)} itens com erros ou avisos.\n\nO log de erros foi salvo em 'log_erros.txt'.\n\nDeseja tentar baixá-los novamente?")
            
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
     error_details = traceback.format_exc()
     registrar_log(f"ERRO CRÍTICO: {error_details}")
     messagebox.showerror("Erro Crítico", f"Ocorreu um erro grave e a automação será encerrada.\n\nVerifique o 'log_automacao.txt'.\n\nErro: {error_details}")

finally:
    if 'driver' in locals() and 'driver' in vars() and driver:
        registrar_log("Automação finalizada.")
        driver.quit()