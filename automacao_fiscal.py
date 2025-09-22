import os
import time
import shutil
import traceback
import tkinter as tk
from tkinter import scrolledtext
import threading
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- CONSTANTES E CONFIGURAÇÕES ---
LOG_FILENAME = 'log_automacao_oc.log'
INPUT_FILENAME = 'lista.xlsx'

# --- Seletores do Portal SAP (para Ordens de Compra) ---
PORTAL_URL = "https://web.embraer.com.br"
IFRAME_CONTEUDO_PRINCIPAL = (By.ID, "contentAreaFrame")

# Navegação
MENU_SUPRIMENTOS = (By.ID, "tabIndex1")
MENU_TODAS = (By.ID, "0L3N1")

# Ações na página de busca
CAMPO_ORDEM_COMPRA = (By.ID, "GOCI.Wzsulmm100View.txtPO")
LINK_EXIBE_PDF = (By.ID, "GOCI.Wzsulmm100View.lnaPDF.0")


class DownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Download de Ordens de Compra")
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
        self.download_path = os.getcwd()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
        self.root.destroy()

    def registrar_log(self, mensagem):
        log_entry = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n"
        try:
            with open(self.log_path, 'a', encoding='utf-8') as log_file:
                log_file.write(log_entry)
        except Exception as e:
            print(f"Erro ao escrever no log: {e}")
        
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
        caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
        service = ChromeService(executable_path=caminho_chromedriver)
        
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")

        try:
            self.driver = webdriver.Chrome(service=service, options=options)
            self.registrar_log(f"Navegador configurado. Downloads serão salvos em: {self.download_path}")
            return self.driver, WebDriverWait(self.driver, 45)
        except Exception as e:
            self.registrar_log(f"ERRO ao iniciar o chromedriver: {e}")
            self.update_status("ERRO: Verifique se o chromedriver.exe está na pasta e é compatível com seu Chrome.", "red")
            return None, None

    def esperar_download_concluir(self, oc_num, timeout=90):
        self.registrar_log(f"Aguardando download do PDF para a OC {oc_num}...")
        self.update_status(f"Aguardando download para OC {oc_num}...")
        
        tempo_inicial = time.time()
        while time.time() - tempo_inicial < timeout:
            if any(f.endswith('.crdownload') for f in os.listdir(self.download_path)):
                time.sleep(1)
                continue

            pdfs_na_pasta = [os.path.join(self.download_path, f) for f in os.listdir(self.download_path) if f.lower().endswith('.pdf')]
            if pdfs_na_pasta:
                arquivo_recente = max(pdfs_na_pasta, key=os.path.getctime)
                if time.time() - os.path.getctime(arquivo_recente) < 60:
                    novo_nome = os.path.join(self.download_path, f"OC_{oc_num}.pdf")
                    time.sleep(2)

                    try:
                        shutil.move(arquivo_recente, novo_nome)
                        self.registrar_log(f"Download concluído e renomeado para: {os.path.basename(novo_nome)}")
                        return True
                    except Exception as e:
                        self.registrar_log(f"AVISO: Não foi possível renomear o arquivo para OC {oc_num} na primeira tentativa: {e}. Tentando novamente.")
                        time.sleep(3)
                        try:
                            shutil.move(arquivo_recente, novo_nome)
                            self.registrar_log(f"Sucesso na segunda tentativa de renomear para: {os.path.basename(novo_nome)}")
                            return True
                        except Exception as e2:
                            self.registrar_log(f"ERRO: Falha ao renomear o arquivo para OC {oc_num} na segunda tentativa: {e2}")
                            return False
            time.sleep(1)

        self.registrar_log(f"ERRO: Tempo limite de {timeout}s excedido esperando o download da OC {oc_num}.")
        return False

    def run_automation(self):
        try:
            self.update_status("Iniciando automação...")
            driver, wait = self.setup_driver()
            if not driver: return

            driver.get(PORTAL_URL)
            self.registrar_log(f"Navegador aberto em: {PORTAL_URL}")
            
            self.prompt_user_action("Por favor, faça o login e a autenticação no portal. Quando a página principal carregar, clique em 'Continuar'.")
            
            self.registrar_log("Usuário clicou em 'Continuar'. Retomando automação.")
            self.update_status("Login detectado. Iniciando busca pelas Ordens de Compra...")

            # <-- ALTERAÇÃO AQUI: Lendo a planilha 'lista' e a coluna 'Nº Os Cliente'
            df_input = pd.read_excel(INPUT_FILENAME, sheet_name='lista', dtype={'Nº Os Cliente': str})
            self.registrar_log(f"Arquivo '{INPUT_FILENAME}' lido com {len(df_input)} OCs.")

            # <-- ALTERAÇÃO AQUI: Processando a coluna 'Nº Os Cliente' para extrair o número principal
            df_input['OC_BASE'] = df_input['Nº Os Cliente'].str.split('/', expand=True)[0]

            # Filtrando OCs já baixadas usando a nova coluna 'OC_BASE'
            ocs_ja_baixadas = {f.replace('OC_', '').replace('.pdf', '') for f in os.listdir(self.download_path) if f.startswith('OC_') and f.endswith('.pdf')}
            df_a_processar = df_input[~df_input['OC_BASE'].isin(ocs_ja_baixadas)].copy()
            total_a_processar = len(df_a_processar)

            if total_a_processar == 0:
                self.update_status("Todas as OCs da lista já foram baixadas.", "#008A00")
                self.registrar_log("Nenhuma nova OC para processar. Finalizando.")
                return

            self.registrar_log(f"Encontradas {total_a_processar} novas OCs para baixar.")

            self.update_status("Navegando pelo menu do portal...")
            wait.until(EC.element_to_be_clickable(MENU_SUPRIMENTOS)).click()
            self.registrar_log("Clicou no menu 'Suprimentos'.")
            
            wait.until(EC.element_to_be_clickable(MENU_TODAS)).click()
            self.registrar_log("Clicou no submenu 'Todas'.")

            self.registrar_log("Aguardando o iframe de conteúdo...")
            wait.until(EC.frame_to_be_available_and_switch_to_it(IFRAME_CONTEUDO_PRINCIPAL))
            self.registrar_log("Entrou no iframe principal.")

            processadas_count = 0
            for index, row in df_a_processar.iterrows():
                # Usando a coluna 'OC_BASE' para a busca
                oc = str(row['OC_BASE']).strip()
                processadas_count += 1
                self.update_status(f"Processando OC: {oc} ({processadas_count}/{total_a_processar})...")
                
                try:
                    campo_oc = wait.until(EC.element_to_be_clickable(CAMPO_ORDEM_COMPRA))
                    campo_oc.clear()
                    campo_oc.send_keys(oc)
                    campo_oc.send_keys(Keys.RETURN)
                    self.registrar_log(f"Busca realizada para a OC {oc}.")
                    
                    self.registrar_log("Aguardando o link do PDF aparecer...")
                    link_pdf = wait.until(EC.element_to_be_clickable(LINK_EXIBE_PDF))
                    
                    link_pdf.click()
                    self.registrar_log(f"Clique no link para baixar o PDF da OC {oc}.")
                    
                    self.esperar_download_concluir(oc)

                except TimeoutException:
                    msg = f"ERRO: Não foi possível encontrar o resultado para a OC {oc} após a busca. Verifique se a OC é válida ou se a página demorou muito para carregar."
                    self.registrar_log(msg)
                    self.tirar_print_de_erro(oc)
                    driver.switch_to.default_content()
                    wait.until(EC.frame_to_be_available_and_switch_to_it(IFRAME_CONTEUDO_PRINCIPAL))
                except Exception as e:
                    self.registrar_log(f"ERRO inesperado ao processar a OC {oc}: {e}")
                    self.tirar_print_de_erro(oc)

            self.update_status("Processo concluído! Verifique a pasta de downloads.", "#008A00")
            self.registrar_log("Todas as OCs da lista foram processadas.")

        except FileNotFoundError:
            msg = f"ERRO CRÍTICO: Arquivo '{INPUT_FILENAME}' não encontrado. Verifique se ele está na mesma pasta do programa."
            self.update_status(msg, "red")
            self.registrar_log(msg)
        except KeyError:
            # <-- ALTERAÇÃO AQUI: Mensagem de erro específica se a coluna 'Nº Os Cliente' não for encontrada
            msg = "ERRO CRÍTICO: A coluna 'Nº Os Cliente' não foi encontrada na planilha 'lista' do arquivo 'lista.xlsx'. Por favor, verifique o nome da coluna."
            self.update_status(msg, "red")
            self.registrar_log(msg)
        except Exception as e:
            error_details = traceback.format_exc()
            self.registrar_log(f"ERRO CRÍTICO NA EXECUÇÃO: {error_details}")
            self.update_status(f"Erro Crítico: {e}", "red")
        finally:
            if self.driver:
                self.registrar_log("Fechando navegador.")
                self.driver.quit()
                self.driver = None
            self.root.after(0, lambda: self.action_button.pack(pady=(5, 10)))
            self.root.after(0, lambda: self.action_button.config(state='normal', text="Iniciar Automação"))

    def tirar_print_de_erro(self, identificador):
        screenshots_dir = "screenshots_de_erro"
        if not os.path.exists(screenshots_dir):
            os.makedirs(screenshots_dir)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = os.path.join(screenshots_dir, f"erro_oc_{identificador}_{timestamp}.png")
        try:
            self.driver.save_screenshot(nome_screenshot)
            self.registrar_log(f"Screenshot de erro salvo em: '{nome_screenshot}'")
        except Exception as e:
            self.registrar_log(f"FALHA AO SALVAR SCREENSHOT: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DownloaderGUI(root)
    root.mainloop()