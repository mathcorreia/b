import os
import time
import shutil
import pandas as pd
import openpyxl
import locale
import re
import traceback
import tkinter as tk
from tkinter import scrolledtext
import threading
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- CONSTANTES GLOBAIS ---
LOG_FILENAME = 'log_validador.txt'
EXCEL_FILENAME = 'Extracao_Dados_FSE.xlsx'

class ValidadorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Validador de Revisão de Engenharia")
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
        self.excel_path = os.path.join(os.getcwd(), EXCEL_FILENAME)
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

    def setup_excel(self):
        """Cria o ficheiro Excel com o cabeçalho se ele não existir."""
        if not os.path.exists(self.excel_path):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Dados FSE"
            self.headers = [
                "OS", "OC / Item", "CODEM / DT. REV. ROT.", "PN / REV. PN / LID", 
                "IND. RASTR.", "NÚMERO DE SERIAÇÃO", "PN extraído", "REV. FSE",
                "REV. Engenharia", "Status Comparação"
            ]
            sheet.append(self.headers)
            for cell in sheet[1]:
                cell.font = openpyxl.styles.Font(bold=True)
            workbook.save(self.excel_path)
            self.registrar_log(f"Arquivo Excel '{EXCEL_FILENAME}' criado com sucesso.")
        else:
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            self.headers = [cell.value for cell in sheet[1]]
            self.registrar_log(f"Arquivo Excel '{EXCEL_FILENAME}' já existe e será atualizado.")

    def run_automation(self):
        try:
            self.update_status("Iniciando configuração...")
            self.setup_excel()

            # --- LEITURA E FILTRAGEM DE DADOS ---
            df_input = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
            df_input.rename(columns={df_input.columns[0]: 'OS'}, inplace=True)
            df_input[['OC_antes', 'OC_depois']] = df_input.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
            df_input['OS'] = df_input['OS'].astype(str)
            
            df_input = df_input.head(10) # Modo de teste
            self.registrar_log(f"MODO DE TESTE: Execução limitada às primeiras 10 OCs.")

            os_ja_verificadas = set()
            try:
                df_existente = pd.read_excel(self.excel_path)
                if 'OS' in df_existente.columns:
                    os_ja_verificadas = set(df_existente['OS'].astype(str))
            except Exception: pass

            df_a_processar = df_input[~df_input['OS'].isin(os_ja_verificadas)].copy()
            if df_a_processar.empty:
                self.update_status("Nenhuma OS nova para processar na lista de teste.", "#008A00")
                return

            # --- INICIALIZAÇÃO DO NAVEGADOR ---
            self.update_status("Configurando o navegador...")
            caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
            service = ChromeService(executable_path=caminho_chromedriver)
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(self.driver, 15)

            self.driver.get("https://web.embraer.com.br/irj/portal")
            self.prompt_user_action("Faça o login no portal e, quando a página principal carregar, clique em 'Continuar'.")

            # --- LOOP PRINCIPAL "ITEM A ITEM" ---
            for index, row in df_a_processar.iterrows():
                os_num = str(row['OS'])
                self.update_status(f"Processando OS: {os_num} ({index + 1}/{len(df_a_processar)})...")

                # ETAPA 1: EXTRAIR DADOS DO GFS
                self.navegar_para_fse_busca(wait)
                dados_fse = self.extrair_dados_fse(wait, os_num, row['OC_antes'], row['OC_depois'])

                if not dados_fse:
                    self.registrar_log(f"Falha na extração da OS {os_num}. Registrando erro.")
                    dados_fse = {h: "" for h in self.headers}
                    dados_fse["OS"] = os_num
                    dados_fse["Status Comparação"] = "FALHA NA EXTRAÇÃO"
                else:
                    # ETAPA 2: BUSCAR REVISÃO NA ENGENHARIA
                    self.navegar_para_desenhos_engenharia(wait)
                    pn_extraido = dados_fse["PN extraído"]
                    rev_fse = dados_fse["REV. FSE"]
                    rev_engenharia = self.buscar_revisao_engenharia(wait, pn_extraido)

                    dados_fse["REV. Engenharia"] = rev_engenharia
                    status = "FALHA NA BUSCA"
                    if rev_engenharia != "Não encontrada" and rev_fse != "Não encontrada":
                        status = "OK" if rev_engenharia.strip().upper() == rev_fse.strip().upper() else "DIVERGENTE"
                    dados_fse["Status Comparação"] = status

                # SALVAR RESULTADO FINAL NO EXCEL
                workbook = openpyxl.load_workbook(self.excel_path)
                sheet = workbook.active
                sheet.append(list(dados_fse.values()))
                workbook.save(self.excel_path)
                self.registrar_log(f"OS {os_num} processada e salva com status: {dados_fse['Status Comparação']}")

            self.update_status("Processo de teste concluído com sucesso!", "#008A00")

        except Exception as e:
            error_details = traceback.format_exc()
            self.registrar_log(f"ERRO CRÍTICO: {error_details}")
            self.update_status(f"Erro Crítico: {e}", "red")
        finally:
            if self.driver:
                self.driver.quit()
            self.action_button.pack_forget()

    def navegar_para_fse_busca(self, wait):
        self.update_status("Navegando para busca FSE...")
        if len(self.driver.window_handles) > 1:
            self.driver.switch_to.window(self.driver.window_handles[1])
        else: # Se a janela GFS não estiver aberta, abre-a
            original_window = self.driver.current_window_handle
            wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
            wait.until(EC.number_of_windows_to_be(2))
            for handle in self.driver.window_handles:
                if handle != original_window: self.driver.switch_to.window(handle)
        
        self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
        # Pequena espera para garantir que a página de busca carregou
        wait.until(EC.visibility_of_element_located((By.ID, "searchBtn")))

    def extrair_dados_fse(self, wait, os_num, oc1, oc2):
        try:
            # ... (código de extração permanece o mesmo)
            self.registrar_log(f"Buscando OS: {os_num} | OC: {oc1}/{oc2}")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']"))).clear()
            self.driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderNumber']").send_keys(oc1)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']"))).clear()
            self.driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderLine']").send_keys(oc2)
            wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
            
            wait.until(EC.visibility_of_element_located((By.ID, "fseHeader")))
            
            dados = {"OS": os_num}
            dados["OC / Item"] = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[1]/div[5]").replace('\n', ' ')
            dados["CODEM / DT. REV. ROT."] = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[3]/div[1]").replace('CODEM / DT. REV. ROT.\n', '').replace('\n', ' | ')
            dados["PN / REV. PN / LID"] = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[3]/div[2]").replace('PN / REV. PN / LID\n', '').replace('\n', ' | ')
            dados["IND. RASTR."] = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[2]/div[3]").replace('IND. RASTR.\n', '').strip()
            
            seriacao_elements = self.driver.find_elements(By.XPATH, "//*[text()='NÚMERO DE SERIAÇÃO']/following-sibling::div//span")
            dados["NÚMERO DE SERIAÇÃO"] = ", ".join([el.text for el in seriacao_elements if el.text.strip()])

            pn_rev_raw = dados["PN / REV. PN / LID"]
            pn_match = re.search(r'(\d+-\d+-\d+)', pn_rev_raw)
            rev_match = re.search(r'\s+([A-Z])\s+', pn_rev_raw)
            
            dados["PN extraído"] = pn_match.group(1) if pn_match else "Não encontrado"
            dados["REV. FSE"] = rev_match.group(1) if rev_match else "Não encontrada"
            return dados
        except Exception as e:
            self.registrar_log(f"ERRO ao extrair dados da OS {os_num}: {e}")
            self.tirar_print_de_erro(os_num, "extracao_FSE")
            return None
    
    def navegar_para_desenhos_engenharia(self, wait):
        self.update_status("Navegando para Desenhos de Engenharia...")
        self.driver.switch_to.window(self.driver.window_handles[0])
        self.driver.get("https://web.embraer.com.br/irj/portal")
        wait.until(EC.element_to_be_clickable((By.ID, "L2N1"))).click()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
        self.driver.switch_to.default_content() # Volta para o contexto principal para o próximo comando

    def buscar_revisao_engenharia(self, wait, part_number):
        try:
            if not part_number or part_number == "Não encontrado": return "PN não fornecido"
            
            self.registrar_log(f"Buscando revisão para PN: {part_number}")
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ivuFrm_page0ivu0")))
            
            campo_pn = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id, 'PartNumber')]")))
            campo_pn.clear()
            campo_pn.send_keys(part_number)
            
            # Localizador mais específico e focado para o botão 'Desenho'
            button_locator = (By.XPATH, "//*[contains(@title, 'Desenho')] | //span[text()='Desenho'] | //a[text()='Desenho']")
            search_button = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(button_locator))
            self.driver.execute_script("arguments[0].click();", search_button)

            seletor_rev = f"//span[contains(text(), 'Rev ')]"
            rev_element = WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.XPATH, seletor_rev)))
            
            revisao = rev_element.text.split(" ")[-1]
            return revisao
        except Exception as e:
            self.registrar_log(f"ERRO ao buscar revisão para PN {part_number}: {e}")
            self.tirar_print_de_erro(part_number, "busca_revisao")
            return "Não encontrada"
        finally:
            self.driver.switch_to.default_content()

    def safe_find_text(self, by, value):
        try: return self.driver.find_element(by, value).text
        except: return ""

    def tirar_print_de_erro(self, identificador, etapa):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_{etapa}_{identificador}_{timestamp}.png"
        try:
            if self.driver:
                self.driver.save_screenshot(os.path.join(os.getcwd(), nome_screenshot))
                self.registrar_log(f"Screenshot de erro salvo: '{nome_screenshot}'")
        except Exception as e:
            self.registrar_log(f"FALHA AO SALVAR SCREENSHOT: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ValidadorGUI(root)
    root.mainloop()

