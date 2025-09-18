import os
import time
import shutil
import pandas as pd
import locale
import traceback
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- JANELA DE STATUS VISUAL ---
class StatusWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Status da Automação")
        self.root.geometry("450x170")
        self.root.resizable(False, False)
        self.root.eval('tk::PlaceWindow . center')
        
        self.label = tk.Label(self.root, text="", font=("Arial", 16, "bold"), wraplength=420)
        self.label.pack(expand=True, fill="both", padx=15, pady=15)
        self.ok_button = tk.Button(self.root, text="OK", command=self.root.destroy, state="disabled", width=15, font=("Arial", 10, "bold"))
        self.ok_button.pack(pady=(0, 15))

    def show_wait(self, message):
        self.label.config(text=message, fg="#E69500") # Amarelo
        self.ok_button.config(state="disabled")
        self.root.deiconify()
        self.root.update()

    def show_ready(self, message):
        self.label.config(text=message, fg="#008A00") # Verde
        self.ok_button.config(state="normal")
        self.root.deiconify()
        self.root.mainloop()

    def hide(self):
        self.root.withdraw()

# Oculta a janela raiz principal do Tkinter
root = tk.Tk()
root.withdraw()

# Função de log
def registrar_log(mensagem):
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
    PASTA_RAIZ_VERIFICACAO = r'\\fserver\cedoc_docs\Doc - EmbraerProdutivo'
    PASTA_BASE_ANO_ATUAL = os.path.join(PASTA_RAIZ_VERIFICACAO, str(hoje.year))
    MES_ATUAL = f'{num_mes_atual} - {nome_mes_atual}'
    PASTA_DESTINO = os.path.join(PASTA_BASE_ANO_ATUAL, MES_ATUAL)
    LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')

    os.makedirs(PASTA_DESTINO, exist_ok=True)
    
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
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
    # Adiciona uma coluna 'OS' se não existir, baseada na primeira coluna
    if 'OS' not in df.columns:
        df.rename(columns={df.columns[0]: 'OS'}, inplace=True)
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
    registrar_log(f"Arquivo Excel lido. Total de {len(df)} itens na lista inicial.")
    
    # --- VERIFICAÇÃO DE DUPLICIDADE RETROATIVA ---
    registrar_log("Iniciando verificação retroativa de duplicidade (até 2 anos). Isso pode levar alguns minutos...")
    arquivos_existentes = set()
    data_limite = hoje - pd.DateOffset(years=2)
    
    if os.path.exists(PASTA_RAIZ_VERIFICACAO):
        for root_dir, dirs, files in os.walk(PASTA_RAIZ_VERIFICACAO):
            nome_da_pasta_atual = os.path.basename(root_dir)
            try:
                if len(nome_da_pasta_atual) == 4 and nome_da_pasta_atual.isdigit():
                    if int(nome_da_pasta_atual) < data_limite.year:
                        dirs[:] = [] # Poda a busca, não entra em pastas de anos antigos
                        continue
            except ValueError:
                pass
            
            for nome_arquivo in files:
                if nome_arquivo.endswith(".pdf"):
                    # Extrai o número da OS do nome do arquivo
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
    # ----------------------------------------------------------------
    
    if df.empty:
        messagebox.showinfo("Nenhum Item a Processar", "Todos os itens da lista já foram baixados anteriormente. Automação finalizada.")
    else:
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

        # Inicia a janela de status e a esconde
        janela_status = StatusWindow()
        janela_status.hide()

        # Aguarda o login manual
        driver.get("https://web.embraer.com.br/irj/portal")
        janela_status.show_ready("Faça o login e, quando estiver na página principal do portal, clique em OK para continuar...")
        
        wait = WebDriverWait(driver, 30)

        # ETAPA DE NAVEGAÇÃO SEMI-AUTOMÁTICA
        janela_status.show_wait("POR FAVOR, AGUARDE...\nNavegando para o sistema GFS e trocando de aba.")
        
        original_window = driver.current_window_handle
        wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
        registrar_log("Clicou no link 'GFS'.")

        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                break
        registrar_log("Foco alterado para a nova aba da aplicação GFS.")
        
        janela_status.show_ready("AGORA PODE CLICAR EM OK\n\nNo navegador, clique em 'FSE' > 'Busca FSe' e, quando a tela carregar, clique em OK aqui.")

        # Loop de buscar e realizar download
        for index, row in df.iterrows():
            oc1 = row['OC_antes']
            oc2 = row['OC_depois']
            os_num = str(row['OS']) # Pega a OS para o nome do arquivo

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
                
                time.sleep(1)
                lista_materiais_wait = WebDriverWait(driver, 30)
                
                seletor_final = (By.XPATH, "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[1]")
                lista_materiais_btn = lista_materiais_wait.until(EC.element_to_be_clickable(seletor_final))
                lista_materiais_btn.click()

                caminho_arquivo_baixado = esperar_download_concluir(DOWNLOAD_DIR)

                if caminho_arquivo_baixado:
                    nome_arquivo = os.path.basename(caminho_arquivo_baixado)
                    
                    novo_nome_arquivo = f"{os_num}_LM.pdf"
                    destino = os.path.join(PASTA_DESTINO, novo_nome_arquivo)

                    if not os.path.exists(destino):
                        shutil.move(caminho_arquivo_baixado, destino)
                        registrar_log(f"Movido e renomeado: {novo_nome_arquivo} para {PASTA_DESTINO}")
                    else:
                        os.remove(caminho_arquivo_baixado)
                        registrar_log(f"Arquivo já existe no destino: {novo_nome_arquivo}. Download duplicado removido.")
                else:
                    registrar_log(f"ERRO: Download não concluído a tempo para a OC {oc1}/{oc2}")
                
                registrar_log(f"Processo da OC {oc1}/{oc2} concluído. Voltando para a página de busca.")
                driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")

            except Exception as e:
                timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_screenshot = f"erro_oc_{str(oc1).replace('/', '-')}_{timestamp_erro}.png"
                caminho_screenshot = os.path.join(os.getcwd(), nome_screenshot)
                try:
                    driver.save_screenshot(caminho_screenshot)
                    registrar_log(f"ERRO com OC {oc1}/{oc2}: {e} - Screenshot salvo em: {caminho_screenshot}")
                except Exception as screenshot_error:
                    registrar_log(f"ERRO com OC {oc1}/{oc2}: {e} - FALHA AO SALVAR SCREENSHOT: {screenshot_error}")
                
                try:
                    messagebox.showerror("Erro em OC", f"Ocorreu um erro com a OC {oc1}/{oc2}.\n\nPor favor, coloque na tela de busca novamente e clique em OK para continuar com a próxima OC.")
                except Exception as refresh_error:
                    registrar_log(f"AVISO: Falha crítica ao tentar se recuperar. Erro: {refresh_error}")
                    break
        
        # Oculta a janela de status no final para não atrapalhar
        if 'janela_status' in locals():
            janela_status.hide()

        registrar_log("Automação finalizada.")
        driver.quit()

except Exception as e:
     error_details = traceback.format_exc()
     registrar_log(f"ERRO CRÍTICO: {error_details}")
     messagebox.showerror("Erro Crítico", f"Ocorreu um erro grave e a automação será encerrada.\n\nVerifique o 'log_automacao.txt'.\n\nErro: {error_details}")

finally:
    if 'driver' in locals() and 'driver' in vars() and driver:
        driver.quit()