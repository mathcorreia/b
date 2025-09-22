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
            # Se o ficheiro existe, apenas carrega os cabeçalhos para uso posterior
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            self.headers = [cell.value for cell in sheet[1]]
            self.registrar_log(f"Arquivo Excel '{EXCEL_FILENAME}' já existe e será atualizado.")

    def run_automation(self):
        try:
            self.update_status("Iniciando configuração...")
            self.setup_excel()

            # --- LEITURA DO EXCEL DE ENTRADA ---
            self.update_status("Lendo arquivo Excel 'lista.xlsx'...")
            df_input = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
            df_input.rename(columns={df_input.columns[0]: 'OS'}, inplace=True)
            df_input[['OC_antes', 'OC_depois']] = df_input.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
            df_input['OS'] = df_input['OS'].astype(str) # Garante que OS é string para comparação
            self.registrar_log(f"Arquivo 'lista.xlsx' lido com {len(df_input)} itens.")

            # --- Limita o DataFrame para teste ---
            df_input = df_input.head(10)
            self.registrar_log(f"MODO DE TESTE: Execução limitada às primeiras 10 OCs da lista.")

            # --- LER OS JÁ VERIFICADAS E FILTRAR A LISTA ---
            self.update_status("Verificando OCs já processadas...")
            os_ja_verificadas = set()
            try:
                df_existente = pd.read_excel(self.excel_path)
                if 'OS' in df_existente.columns:
                    os_ja_verificadas = set(df_existente['OS'].astype(str))
                self.registrar_log(f"Encontradas {len(os_ja_verificadas)} OSs no arquivo de resultados.")
            except Exception as e:
                self.registrar_log(f"Aviso: Não foi possível ler o arquivo Excel existente. Verificando todas as OSs. Erro: {e}")

            df_a_processar = df_input[~df_input['OS'].isin(os_ja_verificadas)].copy()
            novas_os_count = len(df_a_processar)
            self.registrar_log(f"{len(df_input) - novas_os_count} OSs da lista de teste já foram processadas e serão ignoradas.")

            if novas_os_count == 0:
                self.update_status("Nenhuma nova OS para extrair na amostra de teste. Verificando comparações pendentes...", "#00529B")
            else:
                self.registrar_log(f"Iniciando extração de dados para {novas_os_count} novas OSs.")
                
                # --- CONFIGURAÇÃO DO NAVEGADOR (só se necessário) ---
                self.update_status("Configurando o navegador...")
                caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
                service = ChromeService(executable_path=caminho_chromedriver)
                options = webdriver.ChromeOptions()
                options.add_argument("--start-maximized")
                self.driver = webdriver.Chrome(service=service, options=options)
                wait = WebDriverWait(self.driver, 15)

                self.driver.get("https://web.embraer.com.br/irj/portal")
                self.prompt_user_action("Faça o login no portal e, quando a página principal carregar, clique em 'Continuar'.")

                # --- ETAPA 1: EXTRAÇÃO E SALVAMENTO INICIAL (APENAS PARA NOVAS OSs) ---
                self.update_status(f"ETAPA 1: Extraindo dados de {novas_os_count} novas OSs...")
                self.navegar_para_fse_busca(wait)
                
                for index, row in df_a_processar.iterrows():
                    os_num = str(row['OS'])
                    self.update_status(f"Extraindo dados da OS: {os_num} ({index + 1}/{len(df_a_processar)})...")
                    dados_fse = self.extrair_dados_fse(wait, os_num, row['OC_antes'], row['OC_depois'])
                    if dados_fse:
                        dados_fse["REV. Engenharia"] = ""
                        dados_fse["Status Comparação"] = ""
                        workbook = openpyxl.load_workbook(self.excel_path)
                        sheet = workbook.active
                        sheet.append(list(dados_fse.values()))
                        workbook.save(self.excel_path)
                self.registrar_log("Etapa 1 (Extração de novas OSs) concluída.")

            # --- ETAPA 2: ATUALIZAÇÃO COM COMPARAÇÃO (PARA ITENS PENDENTES) ---
            self.update_status("ETAPA 2: Verificando comparações pendentes...")
            
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            col_indices = {name: i+1 for i, name in enumerate(self.headers)}
            
            # Identifica as linhas que precisam de comparação
            linhas_a_comparar = []
            for i, row_cells in enumerate(sheet.iter_rows(min_row=2, values_only=False)):
                # Verifica se a OS da linha está na lista de teste antes de adicionar
                os_da_linha = str(row_cells[col_indices["OS"] - 1].value)
                if os_da_linha in df_input['OS'].values:
                    status_cell = row_cells[col_indices["Status Comparação"] - 1]
                    if not status_cell.value: # Se a célula de status está vazia
                        linhas_a_comparar.append(i + 2) # Guarda o número da linha

            if not linhas_a_comparar:
                self.update_status("Nenhuma comparação pendente na amostra de teste. Processo finalizado!", "#008A00")
            else:
                if not self.driver: # Inicia o driver se ainda não foi iniciado
                    self.update_status("Configurando o navegador para a Etapa 2...")
                    caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
                    service = ChromeService(executable_path=caminho_chromedriver)
                    options = webdriver.ChromeOptions()
                    options.add_argument("--start-maximized")
                    self.driver = webdriver.Chrome(service=service, options=options)
                    wait = WebDriverWait(self.driver, 15)
                    self.driver.get("https://web.embraer.com.br/irj/portal")
                    self.prompt_user_action("Faça o login para a etapa de comparação e clique em 'Continuar'.")

                self.update_status(f"ETAPA 2: Comparando {len(linhas_a_comparar)} itens...")
                self.navegar_para_desenhos_engenharia(wait)

                for row_num in linhas_a_comparar:
                    row_cells = sheet[row_num]
                    pn_extraido = row_cells[col_indices["PN extraído"] - 1].value
                    rev_fse = row_cells[col_indices["REV. FSE"] - 1].value
                    
                    self.update_status(f"Comparando PN: {pn_extraido} (linha {row_num})...")

                    if pn_extraido and pn_extraido != "Não encontrado":
                        rev_engenharia = self.buscar_revisao_engenharia(wait, pn_extraido)
                        
                        status = "FALHA NA BUSCA"
                        if rev_engenharia and rev_fse and rev_engenharia != "Não encontrada":
                            if rev_engenharia.strip().upper() == rev_fse.strip().upper():
                                status = "OK"
                            else:
                                status = "DIVERGENTE"
                        
                        sheet.cell(row=row_num, column=col_indices["REV. Engenharia"], value=rev_engenharia)
                        sheet.cell(row=row_num, column=col_indices["Status Comparação"], value=status)
                    else:
                        sheet.cell(row=row_num, column=col_indices["Status Comparação"], value="PN NÃO ENCONTRADO NA FSE")

                    workbook.save(self.excel_path)

            self.update_status("Processo de teste concluído com sucesso!", "#008A00")

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
        wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
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

            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return dados

        except Exception as e:
            self.registrar_log(f"ERRO ao extrair dados da OS {os_num}: {e}")
            self.tirar_print_de_erro(os_num, "extracao_FSE")
            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return None
    
    def navegar_para_desenhos_engenharia(self, wait):
        self.driver.switch_to.window(self.driver.window_handles[0])
        self.driver.get("https://web.embraer.com.br/irj/portal")
        wait.until(EC.element_to_be_clickable((By.ID, "L2N1"))).click()
        self.prompt_user_action("Valide se a tela 'Desenhos Engenharia' está aberta e clique em 'Continuar'.")
    
    def find_and_click(self, wait, selectors, description):
        """
        Tenta localizar um elemento usando uma lista de seletores XPath.
        Clica no primeiro que encontrar e retorna True.
        Se nenhum for encontrado, retorna False.
        """
        for i, selector in enumerate(selectors):
            try:
                self.registrar_log(f"Tentativa {i+1} para '{description}' com seletor: {selector}")
                element = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                self.registrar_log(f"SUCESSO: Elemento '{description}' encontrado.")
                self.driver.execute_script("arguments[0].click();", element)
                return True
            except TimeoutException:
                self.registrar_log(f"Tentativa {i+1} falhou.")
                continue # Tenta o próximo seletor
        
        self.registrar_log(f"ERRO: Não foi possível localizar o elemento '{description}' com nenhum dos seletores.")
        return False

    #
# COLOQUE ESTA VERSÃO DE DEBUG TEMPORARIAMENTE NO LUGAR DA SUA buscar_revisao_engenharia
#
def buscar_revisao_engenharia(self, wait, part_number):
    """
    VERSÃO DE DEBUG: Pausa o script para permitir a inspeção do botão 'Desenho'.
    """
    self.registrar_log(f"--- MODO DEBUG ATIVADO para o PN: {part_number} ---")
    self.driver.switch_to.default_content()

    try:
        if not part_number or part_number == "Não encontrado":
            return "PN não fornecido"

        # Navega até o formulário
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[starts-with(@id, 'ivuFrm_')]")))

        # Preenche o campo
        campo_pn = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@id, 'PartNumber')]")))
        campo_pn.clear()
        campo_pn.send_keys(part_number)
        self.registrar_log(f"Campo preenchido com: {part_number}")
        
        # PAUSA O SCRIPT AQUI!
        self.registrar_log("!!! SCRIPT PAUSADO !!!")
        self.registrar_log(">>> POR FAVOR, SIGA AS INSTRUÇÕES PARA INSPECIONAR O BOTÃO 'DESENHO' NO NAVEGADOR. <<<")
        self.registrar_log(">>> APÓS COPIAR O HTML, VOLTE AQUI NO TERMINAL E PRESSIONE 'ENTER' PARA CONTINUAR. <<<")
        
        # Esta linha vai pausar a execução até você apertar Enter no terminal
        input() 

        # Apenas para o script não quebrar depois da pausa
        self.registrar_log("Continuando execução após a pausa...")
        return "DEBUG_CONCLUIDO"

    except Exception as e:
        self.registrar_log(f"ERRO DURANTE O MODO DEBUG: {traceback.format_exc()}")
        input("Pressione Enter para fechar.")
        return "ERRO_DEBUG"
    finally:
        self.driver.switch_to.default_content()



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
