import os
import time
import shutil
import re
import traceback
import tkinter as tk
from tkinter import scrolledtext, filedialog
import threading
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- CONSTANTES E CONFIGURAÇÕES ---
LOG_FILENAME = 'log_automacao_fiscal.log'
INPUT_FILENAME = 'lista_ocs.xlsx' # Nome do arquivo Excel de entrada
DOWNLOAD_DIR_NAME = "Notas_Fiscais_Baixadas" # Nome da pasta para salvar os PDFs

# --- Seletores do Portal SAP (centralizados para fácil manutenção) ---
PORTAL_URL = "http://web.embraer.com.br:55100/irj/portal"
IFRAME_CONTEUDO_PRINCIPAL = (By.ID, "contentAreaFrame")

# Navegação
MENU_SUPRIMENTOS = (By.ID, "tabIndex1")
MENU_ORDENS_COMPRA = (By.ID, "L2N0")
MENU_TODAS = (By.ID, "0L3N1")

# Busca de OC
# O ID pode ser dinâmico, então usamos um XPath mais genérico
CAMPO_ORDEM_COMPRA = (By.XPATH, "//*[contains(@id, 'WD_SELECT_OPTIONS_ID_') and @ct='I']")
BOTAO_APLICAR_BUSCA = (By.XPATH, "//button[.//span[text()='Aplicar']]")
BOTAO_EXIBE_PDF = (By.XPATH, "//button[contains(., 'Exibe PDF')]")


class DownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Download de Notas Fiscais")
        self.root.geometry("850x650")
        self.root.attributes('-topmost', True)
        
        self.user_action_event = threading.Event()
        self.driver = None

        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(expand=True, fill='both')

        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill='x', pady=(0, 5))

        self.label_status = tk.Label(top_frame, text="Pronto para iniciar.", font=("Helvetica", 12, "bold"), fg="#00529B", pady=10, wraplength=700, justify='center')
        self.label_status.pack()

        self.action_button = tk.Button(top_frame, text="Iniciar Automação", command=self.iniciar_automacao_thread, font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", padx=20, pady=10)
        self.action_button.pack(pady=(5, 10))

        log_label = tk.Label(main_frame, text="Log em Tempo Real:", font=("Helvetica", 10, "bold"))
        log_label.pack(fill='x', pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(main_frame, state='disabled', wrap=tk.WORD, font=("Courier New", 9))
        self.log_text.pack(expand=True, fill='both', pady=5)
        
        self.log_path = os.path.join(os.getcwd(), LOG_FILENAME)
        self.download_path = os.path.join(os.getcwd(), DOWNLOAD_DIR_NAME)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        if self.driver:
            self.driver.quit()
        self.root.destroy()

    def registrar_log(self, mensagem):
        log_entry = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n"
        with open(self.log_path, 'a', encoding='utf-8') as log_file:
            log_file.write(log_entry)
        
        def update_gui():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.root.after(0, update_gui)

    def update_status(self, text, color="#00529B"):
        self.root.after(0, lambda: self.label_status.config(text=text, fg=color))

    def iniciar_automacao_thread(self):
        self.action_button.config(state='disabled', text="Executando...")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        threading.Thread(target=self.run_automation, daemon=True).start()

    def prompt_user_action(self, message):
        self.user_action_event.clear()
        self.root.after(0, lambda: [
            self.update_status(message, color="#E69500"),
            self.action_button.config(text="Continuar", command=self.signal_user_action, state="normal")
        ])
        self.user_action_event.wait()
        self.root.after(0, lambda: self.action_button.config(state='disabled', text="Executando..."))

    def signal_user_action(self):
        self.user_action_event.set()

    def setup_driver(self):
        """Configura e retorna uma instância do WebDriver para download automático."""
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        
        caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
        service = ChromeService(executable_path=caminho_chromedriver)
        
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True, # Força o download em vez de abrir no navegador
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")

        self.driver = webdriver.Chrome(service=service, options=options)
        self.registrar_log(f"Navegador configurado. Downloads serão salvos em: {self.download_path}")
        return self.driver, WebDriverWait(self.driver, 20)

    def esperar_download_concluir(self, oc_num, timeout=60):
        """Espera um arquivo ser baixado e o renomeia."""
        self.registrar_log(f"Aguardando download do PDF para a OC {oc_num}...")
        segundos = 0
        while segundos < timeout:
            arquivos_na_pasta = os.listdir(self.download_path)
            if any(f.endswith('.crdownload') for f in arquivos_na_pasta):
                time.sleep(1)
                segundos += 1
            else:
                # Encontra o arquivo mais recente que não seja .crdownload
                time.sleep(2) # Garante que a escrita no disco terminou
                files = [os.path.join(self.download_path, f) for f in os.listdir(self.download_path) if f.endswith('.pdf')]
                if not files: continue

                latest_file = max(files, key=os.path.getctime)
                novo_nome = os.path.join(self.download_path, f"OC_{oc_num}.pdf")
                
                # Loop para tentar renomear, caso o arquivo ainda esteja bloqueado
                for _ in range(5):
                    try:
                        shutil.move(latest_file, novo_nome)
                        self.registrar_log(f"Download concluído e renomeado para: {os.path.basename(novo_nome)}")
                        return True
                    except PermissionError:
                        time.sleep(1)
                
                self.registrar_log(f"ERRO: Não foi possível renomear o arquivo baixado para a OC {oc_num} (permissão negada).")
                return False
        
        self.registrar_log(f"ERRO: Tempo limite excedido esperando o download da OC {oc_num}.")
        return False

    def run_automation(self):
        try:
            self.update_status("Iniciando automação...")
            
            df_input = pd.read_excel(INPUT_FILENAME, dtype={'OC': str})
            self.registrar_log(f"Arquivo '{INPUT_FILENAME}' lido com {len(df_input)} OCs.")

            # Filtra OCs que já foram baixadas
            ocs_ja_baixadas = {f.replace('OC_', '').replace('.pdf', '') for f in os.listdir(self.download_path) if f.endswith('.pdf')}
            df_a_processar = df_input[~df_input['OC'].isin(ocs_ja_baixadas)].copy()

            if df_a_processar.empty:
                self.update_status("Todas as OCs da lista já foram baixadas.", "#008A00")
                self.registrar_log("Nenhuma nova OC para processar.")
                return

            self.registrar_log(f"Encontradas {len(df_a_processar)} novas OCs para baixar.")
            driver, wait = self.setup_driver()
            
            driver.get(PORTAL_URL)
            self.prompt_user_action("Faça o login no portal SAP. Quando a página principal carregar, clique em 'Continuar'.")

            # Navegação inicial
            self.update_status("Navegando pelo menu...")
            wait.until(EC.element_to_be_clickable(MENU_SUPRIMENTOS)).click()
            self.registrar_log("Clicou em 'Suprimentos'.")
            wait.until(EC.element_to_be_clickable(MENU_ORDENS_COMPRA)).click()
            self.registrar_log("Clicou em 'Ordens de Compra'.")
            wait.until(EC.element_to_be_clickable(MENU_TODAS)).click()
            self.registrar_log("Clicou em 'Todas'.")

            # Entra no iframe principal onde a busca é realizada
            wait.until(EC.frame_to_be_available_and_switch_to_it(IFRAME_CONTEUDO_PRINCIPAL))
            self.registrar_log("Entrou no iframe de conteúdo.")

            # Loop para baixar cada OC
            for index, row in df_a_processar.iterrows():
                oc = str(row['OC'])
                self.update_status(f"Processando OC: {oc} ({index + 1}/{len(df_a_processar)})...")
                
                try:
                    campo_oc = wait.until(EC.visibility_of_element_located(CAMPO_ORDEM_COMPRA))
                    campo_oc.clear()
                    campo_oc.send_keys(oc)

                    wait.until(EC.element_to_be_clickable(BOTAO_APLICAR_BUSCA)).click()
                    self.registrar_log(f"Busca realizada para a OC {oc}.")
                    
                    # Aguarda um tempo para os resultados carregarem
                    time.sleep(3) 

                    # Rola até o botão de PDF e clica
                    botao_pdf = wait.until(EC.presence_of_element_located(BOTAO_EXIBE_PDF))
                    driver.execute_script("arguments[0].scrollIntoView(true);", botao_pdf)
                    time.sleep(0.5)
                    botao_pdf.click()
                    
                    self.esperar_download_concluir(oc)

                except Exception as e:
                    self.registrar_log(f"ERRO ao processar a OC {oc}: {e}")
                    self.tirar_print_de_erro(oc)

            self.update_status("Processo concluído com sucesso!", "#008A00")

        except FileNotFoundError:
            self.update_status(f"ERRO: Arquivo '{INPUT_FILENAME}' não encontrado!", "red")
            self.registrar_log(f"ERRO CRÍTICO: Arquivo de entrada '{INPUT_FILENAME}' não foi encontrado no mesmo diretório do programa.")
        except Exception as e:
            error_details = traceback.format_exc()
            self.registrar_log(f"ERRO CRÍTICO: {error_details}")
            self.update_status(f"Erro Crítico: {e}", "red")
        finally:
            if self.driver:
                self.registrar_log("Fechando navegador.")
                self.driver.quit()
                self.driver = None
            self.root.after(0, lambda: self.action_button.pack_forget())

    def tirar_print_de_erro(self, identificador):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_oc_{identificador}_{timestamp}.png"
        try:
            self.driver.save_screenshot(nome_screenshot)
            self.registrar_log(f"Screenshot de erro salvo: '{nome_screenshot}'")
        except Exception as e:
            self.registrar_log(f"FALHA AO SALVAR SCREENSHOT: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DownloaderGUI(root)
    root.mainloop()
