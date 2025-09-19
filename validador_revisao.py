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
        """Cria o arquivo Excel com o cabeçalho se ele não existir."""
        if not os.path.exists(self.excel_path):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Dados FSE"
            headers = [
                "OS", "OC / Item", "CODEM / DT. REV. ROT.", "PN / REV. PN / LID", 
                "IND. RASTR.", "NÚMERO DE SERIAÇÃO", "PN extraído", "REV. FSE",
                "REV. Engenharia", "Status Comparação"
            ]
            sheet.append(headers)
            # Formatação opcional
            for cell in sheet[1]:
                cell.font = openpyxl.styles.Font(bold=True)
            workbook.save(self.excel_path)
            self.registrar_log(f"Arquivo Excel '{EXCEL_FILENAME}' criado com sucesso.")

    def run_automation(self):
        try:
            # --- CONFIGURAÇÃO INICIAL ---
            self.update_status("Iniciando configuração...")
            self.setup_excel()

            # --- LEITURA DO EXCEL DE ENTRADA ---
            self.update_status("Lendo arquivo Excel 'lista.xlsx'...")
            df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
            df.rename(columns={df.columns[0]: 'OS'}, inplace=True)
            df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
            self.registrar_log(f"Arquivo Excel lido com {len(df)} itens para processar.")

            # --- CONFIGURAÇÃO DO NAVEGADOR ---
            self.update_status("Configurando o navegador...")
            caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
            service = ChromeService(executable_path=caminho_chromedriver)
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(self.driver, 15)

            # --- LOGIN E NAVEGAÇÃO INICIAL ---
            self.driver.get("https://web.embraer.com.br/irj/portal")
            self.prompt_user_action("Faça o login no portal e, quando a página principal carregar, clique em 'Continuar'.")

            # --- ETAPA 1: EXTRAÇÃO DE DADOS DA FSE ---
            self.update_status("Navegando para o GFS para extrair dados...")
            self.navegar_para_fse_busca(wait)
            
            resultados = []
            for index, row in df.iterrows():
                os_num = str(row['OS'])
                self.update_status(f"Extraindo dados da OS: {os_num}...")
                dados_fse = self.extrair_dados_fse(wait, os_num, row['OC_antes'], row['OC_depois'])
                if dados_fse:
                    resultados.append(dados_fse)
            
            # --- ETAPA 2: COMPARAÇÃO COM DADOS DA ENGENHARIA ---
            self.update_status("Navegando para Desenhos de Engenharia para comparação...")
            self.navegar_para_desenhos_engenharia(wait)

            for dados in resultados:
                os_num = dados["OS"]
                pn_extraido = dados["PN extraído"]
                self.update_status(f"Comparando revisão para OS: {os_num} (PN: {pn_extraido})...")
                rev_engenharia = self.buscar_revisao_engenharia(wait, pn_extraido)
                dados["REV. Engenharia"] = rev_engenharia
                
                # Comparação
                if rev_engenharia and dados["REV. FSE"]:
                    if rev_engenharia.strip().upper() == dados["REV. FSE"].strip().upper():
                        dados["Status Comparação"] = "OK"
                    else:
                        dados["Status Comparação"] = "DIVERGENTE"
                else:
                    dados["Status Comparação"] = "FALHA NA BUSCA"
                
                # Salva a linha completa no Excel
                workbook = openpyxl.load_workbook(self.excel_path)
                sheet = workbook.active
                sheet.append(list(dados.values()))
                workbook.save(self.excel_path)

            self.update_status("Processo concluído com sucesso!", "#008A00")

        except Exception as e:
            error_details = traceback.format_exc()
            self.registrar_log(f"ERRO CRÍTICO: {error_details}")
            self.update_status(f"Erro Crítico: {e}", "red")
        finally:
            if self.driver:
                self.registrar_log("Automação finalizada.")
                self.driver.quit()
                self.driver = None
            self.action_button.pack_forget()

    def navegar_para_fse_busca(self, wait):
        original_window = self.driver.current_window_handle
        wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click() # Link GFS
        wait.until(EC.number_of_windows_to_be(2))
        for handle in self.driver.window_handles:
            if handle != original_window:
                self.driver.switch_to.window(handle)
                break
        self.prompt_user_action("No navegador, navegue para 'FSE' > 'Busca FSe' e, quando a tela de busca carregar, clique em 'Continuar'.")

    def extrair_dados_fse(self, wait, os_num, oc1, oc2):
        try:
            self.registrar_log(f"Buscando OS: {os_num} | OC: {oc1}/{oc2}")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']"))).clear()
            self.driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderNumber']").send_keys(oc1)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']"))).clear()
            self.driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderLine']").send_keys(oc2)
            wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
            
            # Espera um elemento chave da página de detalhes carregar
            wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(text(), 'DATA EXPLOSÃO')]")))
            
            # Extração dos dados
            dados = {"OS": os_num}
            dados["OC / Item"] = self.safe_find_text(By.XPATH, "//span[contains(text(), 'OC')]/following-sibling::span")
            dados["CODEM / DT. REV. ROT."] = self.safe_find_text(By.XPATH, "//div[contains(text(), 'CODEM / DT. REV. ROT.')]/following-sibling::div")
            dados["PN / REV. PN / LID"] = self.safe_find_text(By.XPATH, "//div[contains(text(), 'PN / REV. PN / LID')]/following-sibling::div")
            dados["IND. RASTR."] = self.safe_find_text(By.XPATH, "//div[contains(text(), 'IND. RASTR.')]/following-sibling::div")
            
            seriacao_elements = self.driver.find_elements(By.XPATH, "//div[text()='NÚMERO DE SERIAÇÃO']/following-sibling::div//span")
            dados["NÚMERO DE SERIAÇÃO"] = ", ".join([el.text for el in seriacao_elements])

            # Lógica para extrair PN e Revisão do campo "PN / REV. PN / LID"
            pn_rev_raw = dados["PN / REV. PN / LID"]
            pn_match = re.search(r'(\d{4}-\d{4}-\d{3}|\d{4}-\d{4})', pn_rev_raw)
            rev_match = re.search(r'\n([A-Z])\n', pn_rev_raw) # Busca por uma letra maiúscula entre quebras de linha
            
            dados["PN extraído"] = pn_match.group(1) if pn_match else "Não encontrado"
            dados["REV. FSE"] = rev_match.group(1) if rev_match else "Não encontrada"

            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1") # Volta para busca
            return dados

        except Exception as e:
            self.registrar_log(f"ERRO ao extrair dados da OS {os_num}: {e}")
            self.tirar_print_de_erro(os_num, "extracao_FSE")
            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1") # Tenta voltar
            return None
    
    def navegar_para_desenhos_engenharia(self, wait):
        # Assumindo que a navegação parte do portal principal
        self.driver.switch_to.window(self.driver.window_handles[0]) # Volta para a aba principal do portal
        self.driver.get("https://web.embraer.com.br/irj/portal") # Recarrega o portal
        # O caminho de cliques conforme a imagem 2
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Minhas Aplicações')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Desenhos Engenharia')]"))).click()
        self.prompt_user_action("Valide se a tela 'Desenhos Engenharia' está aberta e clique em 'Continuar'.")
    
    def buscar_revisao_engenharia(self, wait, part_number):
        try:
            # Conforme imagem 2 e 3
            campo_pn = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@id, 'PartNumber')]")))
            campo_pn.clear()
            campo_pn.send_keys(part_number)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Consultar')]"))).click()

            # Espera o resultado da árvore de arquivos carregar
            # O seletor busca pela revisão, ex: 'Rev N'
            seletor_rev = f"//span[contains(text(), 'Rev ')]"
            rev_element = wait.until(EC.visibility_of_element_located((By.XPATH, seletor_rev)))
            
            revisao_raw = rev_element.text # Ex: "Rev N"
            revisao = revisao_raw.split(" ")[-1]
            self.registrar_log(f"Revisão encontrada para PN {part_number}: {revisao}")
            return revisao
        except Exception as e:
            self.registrar_log(f"ERRO ao buscar revisão de engenharia para PN {part_number}: {e}")
            self.tirar_print_de_erro(part_number, "busca_revisao")
            return "Não encontrada"

    def safe_find_text(self, by, value):
        try:
            return self.driver.find_element(by, value).text
        except NoSuchElementException:
            return ""

    def tirar_print_de_erro(self, identificador, etapa):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_{etapa}_{identificador}_{timestamp}.png"
        screenshot_path = os.path.join(os.getcwd(), nome_screenshot)
        try:
            if self.driver:
                self.driver.save_screenshot(screenshot_path)
                self.registrar_log(f"Screenshot de erro salvo em: '{screenshot_path}'")
        except Exception as e:
            self.registrar_log(f"FALHA AO SALVAR SCREENSHOT: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ValidadorGUI(root)
    root.mainloop()
