import os
import time
import pandas as pd
import openpyxl
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
from selenium.webdriver.common.action_chains import ActionChains

import pyodbc

LOG_FILENAME = 'log_validador.txt'
EXCEL_FILENAME = 'Extracao_Dados_FSE.xlsx'
ERROS_DIR = 'erros'

class ValidadorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Validador de Revisão de Engenharia")
        self.root.geometry("850x650")
        self.root.attributes('-topmost', True)
        
        self.user_action_event = threading.Event()
        self.driver = None
        
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.reprocess_choice = ""

        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(expand=True, fill='both')

        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill='x', pady=(0, 5))

        self.label_status = tk.Label(top_frame, text="Pronto para iniciar.", font=("Helvetica", 12, "bold"), fg="#00529B", pady=10, wraplength=700, justify='center')
        self.label_status.pack()

        self.action_frame = tk.Frame(top_frame)
        self.action_frame.pack(pady=(5,10))

        self.action_button = tk.Button(self.action_frame, text="Iniciar Automação", command=self.iniciar_automacao_thread, font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", padx=20, pady=10)
        self.action_button.pack(side='left', padx=5)

        self.pause_button = tk.Button(self.action_frame, text="Pausar", command=self.request_pause, font=("Helvetica", 12, "bold"), bg="#E69500", fg="white", padx=20, pady=10)

        self.reprocess_frame = tk.Frame(main_frame)
        self.reprocess_button = tk.Button(self.reprocess_frame, text="Reprocessar Itens com Erro", command=lambda: self.set_reprocess_choice("reprocess"), font=("Helvetica", 10, "bold"), bg="#FFA500", fg="white")
        self.finish_button = tk.Button(self.reprocess_frame, text="Finalizar", command=lambda: self.set_reprocess_choice("finish"), font=("Helvetica", 10))
        self.reprocess_button.pack(side='left', padx=5)
        self.finish_button.pack(side='left', padx=5)
        self.reprocess_frame.pack_forget()

        log_label = tk.Label(main_frame, text="Log de Atividades:", font=("Helvetica", 10, "bold"))
        log_label.pack(fill='x', pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(main_frame, state='disabled', wrap=tk.WORD, font=("Courier New", 9))
        self.log_text.pack(expand=True, fill='both', pady=5)
        
        self.log_path = os.path.join(os.getcwd(), LOG_FILENAME)
        self.excel_path = os.path.join(os.getcwd(), EXCEL_FILENAME)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.stop_event.set() 
        self.pause_event.set() 
        if self.driver:
            self.driver.quit()
        self.root.destroy()

    def registrar_log(self, mensagem):
        log_entry = f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {mensagem}\n"
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
        self.stop_event.clear()
        self.pause_event.set() 
        self.action_button.pack_forget()
        self.pause_button.config(text="Pausar", command=self.request_pause, bg="#E69500")
        self.pause_button.pack(side='left', padx=5)
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        threading.Thread(target=self.run_automation_loop, daemon=True).start()

    def request_pause(self):
        self.pause_event.clear()
        self.registrar_log("Automação pausada pelo usuário.")
        self.update_status("Automação Pausada. Clique em 'Retomar' para continuar.", color="#E69500")
        self.pause_button.config(text="Retomar", command=self.request_resume, bg="#4CAF50")

    def request_resume(self):
        self.pause_event.set()
        self.registrar_log("Automação retomada pelo usuário.")
        self.update_status("Retomando automação...", color="#00529B")
        self.pause_button.config(text="Pausar", command=self.request_pause, bg="#E69500")

    def request_stop(self): 
        self.stop_event.set()
        self.pause_event.set() 
        self.registrar_log("Solicitação de parada de emergência recebida.")
        
    def prompt_user_action(self, message):
        self.user_action_event.clear()
        self.root.after(0, lambda: [
            self.update_status(message, color="#E69500"),
            self.action_button.config(text="Continuar", command=self.signal_user_action, state="normal"),
            self.action_button.pack(side='left', padx=5),
            self.pause_button.pack_forget()
        ])
        self.user_action_event.wait()
        self.root.after(0, lambda: [
            self.action_button.pack_forget(),
            self.pause_button.pack(side='left', padx=5)
        ])

    def signal_user_action(self):
        self.user_action_event.set()
        
    def set_reprocess_choice(self, choice):
        self.reprocess_choice = choice

    def setup_excel(self):
        if not os.path.exists(self.excel_path):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Dados FSE"
            self.headers = [
                "OS", "OC", "Item", "CODEM", "DT. REV. ROT.", "PN", "REV. PN", "LID", "PLANTA",
                "IND. RASTR.", "NÚMERO DE SERIAÇÃO", "PN extraído", 
                "REV. FSE", "REV. Engenharia", "Revisão do Banco",
                "Status (Eng vs FSE)", "Detalhes (Eng vs FSE)",
                "Status (Banco vs FSE)", "Detalhes (Banco vs FSE)",
                "Status (Banco vs Eng)", "Detalhes (Banco vs Eng)",
                "COD_OS_COMPLETO", "COD_PECAS", "N_OS_CLIENTE", "N_DESENHO", "QTDE_PECAS", "DT_ENTRADA",
                "U_ZLT_REVISAO_2D", "U_ZLT_REVISAO_DES_3D", "U_ZLT_REVISAO_DI", "U_ZLT_REVISAO_DI_F2",
                "U_ZLT_REVISAO_DI_F3", "U_ZLT_REVISAO_FI", "U_ZLT_REVISAO_LP", "U_ZLT_REVISAO_LP_F2",
                "U_ZLT_REVISAO_LP_F3", "U_ZLT_REVISAO_PN", "U_ZLT_REVISAO_ROT", "U_REVISAO_3D_MF",
                "U_CLASSE_ELEB", "U_PN_DE_PROJETO", "U_OM", "CAMPO_ADICIONAL9"
            ]
            sheet.append(self.headers)
            for cell in sheet[1]: cell.font = openpyxl.styles.Font(bold=True)
            workbook.save(self.excel_path)
            self.registrar_log(f"Arquivo de resultados '{EXCEL_FILENAME}' criado com sucesso.")
        else:
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            self.headers = [cell.value for cell in sheet[1]]
            self.registrar_log(f"Arquivo de resultados '{EXCEL_FILENAME}' encontrado.")
    
    def consultar_dados_banco(self, part_number):
        self.registrar_log(f"Consultando banco de dados para o PN: {part_number}...")
        string_conexao = (r'DRIVER={SQL Server};SERVER=172.20.1.7;DATABASE=CPS;UID=sa;PWD=masterkey;')
        comando_sql = "SELECT COD_OS_COMPLETO, COD_PECAS, N_OS_CLIENTE, N_DESENHO, QTDE_PECAS, DT_ENTRADA, U_ZLT_REVISAO_2D, U_ZLT_REVISAO_DES_3D, U_ZLT_REVISAO_DI, U_ZLT_REVISAO_DI_F2, U_ZLT_REVISAO_DI_F3, U_ZLT_REVISAO_FI, U_ZLT_REVISAO_LP, U_ZLT_REVISAO_LP_F2, U_ZLT_REVISAO_LP_F3, U_ZLT_REVISAO_PN, U_ZLT_REVISAO_ROT, U_REVISAO_3D_MF, U_CLASSE_ELEB, U_PN_DE_PROJETO, U_OM, CAMPO_ADICIONAL9 FROM TOS_AUX WHERE N_DESENHO = ?"
        colunas_db = ["COD_OS_COMPLETO", "COD_PECAS", "N_OS_CLIENTE", "N_DESENHO", "QTDE_PECAS", "DT_ENTRADA", "U_ZLT_REVISAO_2D", "U_ZLT_REVISAO_DES_3D", "U_ZLT_REVISAO_DI", "U_ZLT_REVISAO_DI_F2", "U_ZLT_REVISAO_DI_F3", "U_ZLT_REVISAO_FI", "U_ZLT_REVISAO_LP", "U_ZLT_REVISAO_LP_F2", "U_ZLT_REVISAO_LP_F3", "U_ZLT_REVISAO_PN", "U_ZLT_REVISAO_ROT", "U_REVISAO_3D_MF", "U_CLASSE_ELEB", "U_PN_DE_PROJETO", "U_OM", "CAMPO_ADICIONAL9"]
        try:
            with pyodbc.connect(string_conexao) as conexao:
                resultado = pd.read_sql(comando_sql, conexao, params=[part_number])
                if not resultado.empty:
                    self.registrar_log(f"Dados encontrados no banco para o PN: {part_number}")
                    return resultado.iloc[0].to_dict()
                else:
                    self.registrar_log(f"PN {part_number} não encontrado no banco.")
                    return {col: "Não encontrado no BD" for col in colunas_db}
        except Exception as e:
            self.registrar_log(f"ERRO ao consultar o banco de dados: {e}")
            return {col: "Erro no BD" for col in colunas_db}

    def run_automation_loop(self):
        self.registrar_log("--- INICIANDO NOVO CICLO DE AUTOMAÇÃO ---")
        self.update_status("Iniciando automação...")
        reprocess_mode = False
        while not self.stop_event.is_set():
            try:
                if not reprocess_mode:
                    self.update_status("Lendo arquivo 'lista.xlsx'...")
                    df_input = pd.read_excel('lista.xlsx', sheet_name='lista', engine='openpyxl')
                    self.registrar_log(f"Arquivo 'lista.xlsx' lido com sucesso. Total de {len(df_input)} OCs na lista.")
                
                self.run_automation_cycle(df_input, reprocess_mode=reprocess_mode)

                if self.stop_event.is_set(): break

                self.update_status("Verificando se existem erros para reprocessar...", "#00529B")
                erros_df = self.check_for_errors()
                
                if not erros_df.empty:
                    self.reprocess_choice = ""
                    self.root.after(0, lambda: [
                        self.update_status(f"{len(erros_df)} OSs com falhas. Deseja reprocessá-las?", "#E69500"),
                        self.action_frame.pack_forget(),
                        self.reprocess_frame.pack()
                    ])
                    while not self.reprocess_choice and not self.stop_event.is_set(): time.sleep(0.1)
                    self.root.after(0, self.reprocess_frame.pack_forget)
                    
                    if self.reprocess_choice == "reprocess":
                        self.registrar_log(f"--- INICIANDO REPROCESSO PARA {len(erros_df)} ITENS COM FALHA ---")
                        df_input = erros_df.copy()
                        reprocess_mode = True
                        self.clear_error_status_in_excel(df_input)
                        continue
                    else:
                        self.update_status("Processo finalizado com itens pendentes.", "#00529B")
                        break
                else:
                    self.update_status("Processo concluído com sucesso e sem erros!", "#008A00")
                    break
            except Exception as e:
                self.update_status(f"Erro Crítico: {e}", "red")
                self.registrar_log(f"ERRO CRÍTICO NO LOOP: {traceback.format_exc()}")
                break
            finally:
                if self.driver:
                    self.driver.quit()
                    self.driver = None

        self.root.after(0, lambda: [
            self.pause_button.pack_forget(),
            self.action_button.config(state='normal', text="Iniciar Automação"),
            self.action_button.pack(side='left', padx=5)
        ])
        self.registrar_log("--- FIM DA EXECUÇÃO ---")
    
    def check_for_errors(self):
        df = pd.read_excel(self.excel_path)
        df_erros = df[(~df['Status (Eng vs FSE)'].astype(str).str.contains("OK", na=False))].copy()
        return df_erros

    def clear_error_status_in_excel(self, df_erros):
        workbook = openpyxl.load_workbook(self.excel_path)
        sheet = workbook.active
        col_indices = {name: i + 1 for i, name in enumerate(self.headers)}
        os_com_erro = set(df_erros['OS'].astype(str))
        for row_num in range(2, sheet.max_row + 1):
            if str(sheet.cell(row=row_num, column=col_indices["OS"]).value) in os_com_erro:
                for col in ["Status (Eng vs FSE)", "Detalhes (Eng vs FSE)", "Status (Banco vs FSE)", "Detalhes (Banco vs FSE)", "Status (Banco vs Eng)", "Detalhes (Banco vs Eng)"]:
                    sheet.cell(row=row_num, column=col_indices[col], value=None)
        workbook.save(self.excel_path)
        self.registrar_log(f"{len(os_com_erro)} status de itens com erro foram limpos para reprocessamento.")

    def run_automation_cycle(self, df_to_process, reprocess_mode=False):
        try:
            self.setup_excel()
            if 'OC_antes' not in df_to_process.columns:
                df_to_process.rename(columns={df_to_process.columns[0]: 'OS'}, inplace=True)
                df_to_process[['OC_antes', 'OC_depois']] = df_to_process.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
                df_to_process['OS'] = df_to_process['OS'].astype(str)

            if not reprocess_mode:
                self.update_status("Verificando OCs já extraídas...")
                os_ja_extraidas = set(pd.read_excel(self.excel_path)['OS'].astype(str)) if os.path.exists(self.excel_path) and 'OS' in pd.read_excel(self.excel_path).columns else set()
                self.registrar_log(f"Encontradas {len(os_ja_extraidas)} OSs no arquivo de resultados.")
                
                df_a_extrair = df_to_process[~df_to_process['OS'].isin(os_ja_extraidas)].copy()
                novas_os_count = len(df_a_extrair)
                self.registrar_log(f"{len(df_to_process) - novas_os_count} OSs da lista já foram extraídas e serão ignoradas.")
                
                if novas_os_count > 0:
                    self.registrar_log(f"Iniciando extração de dados para {novas_os_count} novas OCs.")
                    if not self.driver:
                        self._setup_driver()
                        self.driver.get("https://web.embraer.com.br/irj/portal")
                        self.prompt_user_action("Por favor, faça o login no portal.")
                    
                    if self.stop_event.is_set(): return
                    wait = WebDriverWait(self.driver, 20)
                    self.navegar_para_fse_busca(wait)

                    for i, (_, row) in enumerate(df_a_extrair.iterrows()):
                        self.pause_event.wait() 
                        if self.stop_event.is_set(): break
                        self.update_status(f"Extraindo {i+1} de {novas_os_count}: OC {row['OC_antes']}/{row['OC_depois']}")
                        dados_fse = self.extrair_dados_fse(wait, str(row['OS']), row['OC_antes'], row['OC_depois'])
                        if dados_fse: self.append_to_excel(dados_fse)
                    if self.stop_event.is_set(): return
                else:
                    self.registrar_log("Nenhuma OC nova para extrair.")

            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook.active
            col_indices = {name: i+1 for i, name in enumerate(self.headers)}
            
            self.update_status("Verificando itens que necessitam de comparação...")
            linhas_a_comparar = [i + 2 for i, row in enumerate(sheet.iter_rows(min_row=2)) if str(row[0].value) in df_to_process['OS'].values and not row[col_indices["Status (Eng vs FSE)"] - 1].value]
            self.registrar_log(f"Encontradas {len(linhas_a_comparar)} OSs com comparação pendente neste ciclo.")

            if linhas_a_comparar:
                if not self.driver:
                    self._setup_driver()
                    self.driver.get("https://web.embraer.com.br/irj/portal")
                    self.prompt_user_action("Faça o login para a etapa de comparação.")
                
                wait = WebDriverWait(self.driver, 20)
                if self.stop_event.is_set(): return
                self.navegar_para_desenhos_engenharia(wait)
                
                for i, row_num in enumerate(linhas_a_comparar):
                    self.pause_event.wait()
                    if self.stop_event.is_set(): break
                    self.update_status(f"Comparando item {i+1} de {len(linhas_a_comparar)}...")
                    self.processar_linha(sheet, row_num, col_indices, wait)
                
                workbook.save(self.excel_path)
                self.registrar_log("Dados de comparação foram salvos no Excel.")
        finally:
            if self.driver: self.driver.quit(); self.driver = None

    def _setup_driver(self):
        caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
        service = ChromeService(executable_path=caminho_chromedriver)
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument("--disable-blink-features=AutomationControlled")
        self.driver = webdriver.Chrome(service=service, options=options)

    def append_to_excel(self, dados):
        workbook = openpyxl.load_workbook(self.excel_path)
        sheet = workbook.active
        linha = [dados.get(h, "") for h in self.headers]
        sheet.append(linha)
        workbook.save(self.excel_path)
        self.registrar_log(f"Dados da OS {dados.get('OS')} salvos no Excel.")

    def processar_linha(self, sheet, row_num, col_indices, wait):
        pn_extraido = sheet.cell(row=row_num, column=col_indices["PN extraído"]).value
        rev_fse = sheet.cell(row=row_num, column=col_indices["REV. FSE"]).value
        self.registrar_log(f"Processando PN: {pn_extraido}...")
        
        if pn_extraido and pn_extraido != "Não encontrado":
            rev_engenharia = self.buscar_revisao_engenharia(wait, pn_extraido)
            dados_banco = self.consultar_dados_banco(pn_extraido)
            revisao_banco = dados_banco.get("U_ZLT_REVISAO_PN", "Chave não encontrada")
            status_eng_fse, d_eng_fse = self.comparar_revisoes(rev_engenharia, rev_fse, "ENG", "FSE")
            status_banco_fse, d_banco_fse = self.comparar_revisoes(revisao_banco, rev_fse, "BANCO", "FSE")
            status_banco_eng, d_banco_eng = self.comparar_revisoes(revisao_banco, rev_engenharia, "BANCO", "ENG")

            sheet.cell(row=row_num, column=col_indices["REV. Engenharia"], value=rev_engenharia)
            sheet.cell(row=row_num, column=col_indices["Revisão do Banco"], value=revisao_banco)
            sheet.cell(row=row_num, column=col_indices["Status (Eng vs FSE)"], value=status_eng_fse)
            sheet.cell(row=row_num, column=col_indices["Detalhes (Eng vs FSE)"], value=d_eng_fse)
            sheet.cell(row=row_num, column=col_indices["Status (Banco vs FSE)"], value=status_banco_fse)
            sheet.cell(row=row_num, column=col_indices["Detalhes (Banco vs FSE)"], value=d_banco_fse)
            sheet.cell(row=row_num, column=col_indices["Status (Banco vs Eng)"], value=status_banco_eng)
            sheet.cell(row=row_num, column=col_indices["Detalhes (Banco vs Eng)"], value=d_banco_eng)
            for col_nome, valor in dados_banco.items():
                if col_nome in col_indices: sheet.cell(row=row_num, column=col_indices[col_nome], value=valor)
        else:
            sheet.cell(row=row_num, column=col_indices["Status (Eng vs FSE)"], value="PN NÃO ENCONTRADO NA FSE")

    def comparar_revisoes(self, rev1, rev2, nome1, nome2):
        detalhes = f"{nome1}: {rev1} vs {nome2}: {rev2}"
        status = "FALHA"
        is_rev1_valida = rev1 and "Não encontrad" not in str(rev1) and "Erro" not in str(rev1)
        is_rev2_valida = rev2 and "Não encontrad" not in str(rev2) and "Erro" not in str(rev2)
        if is_rev1_valida and is_rev2_valida:
            if str(rev1).strip().upper() == str(rev2).strip().upper(): status = "OK"
            else: status = "DIVERGENTE"
        return status, detalhes

    def extrair_dados_fse(self, wait, os_num, oc1, oc2):
        try:
            oc_completa = f"{oc1}/{oc2}"
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']"))).clear()
            self.driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderNumber']").send_keys(oc1)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']"))).clear()
            self.driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderLine']").send_keys(oc2)
            wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
            wait.until(EC.visibility_of_element_located((By.ID, "fseHeader")))
            
            dados = {"OS": os_num}
            oc_item_raw = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[1]/div[5]").replace('\n', '/').strip()
            oc_item_split = [x.strip() for x in oc_item_raw.split('/')]; dados["OC"], dados["Item"] = (oc_item_split[0], oc_item_split[1]) if len(oc_item_split) > 1 else (oc_item_split[0] if oc_item_split else "", "")
            codem_raw = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[3]/div[1]").replace('CODEM / DT. REV. ROT.\n', '').strip()
            codem_split = [x.strip() for x in codem_raw.split('\n')]; dados["CODEM"], dados["DT. REV. ROT."] = (codem_split[0], codem_split[1]) if len(codem_split) > 1 else (codem_split[0] if codem_split else "", "")
            pn_raw = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[3]/div[2]").replace('PN / REV. PN / LID\n', '').strip()
            pn_parts = [p for p in pn_raw.replace('\n', ' ').split() if p]; dados["PN"], dados["REV. PN"], dados["LID"] = (pn_parts[0], pn_parts[1], pn_parts[2]) if len(pn_parts) > 2 else ((pn_parts[0], pn_parts[1], "") if len(pn_parts) > 1 else ((pn_parts[0], "", "") if len(pn_parts) > 0 else ("", "", "")))
            dados["PLANTA"] = self.safe_find_text(By.XPATH, "//*[normalize-space()='PLANTA']/parent::div/following-sibling::div").strip()
            dados["IND. RASTR."] = self.safe_find_text(By.XPATH, "//*[@id='fseHeader']/div[2]/div[3]").replace('IND. RASTR.\n', '').strip()
            seriacao_elements = self.driver.find_elements(By.XPATH, "//*[normalize-space()='NÚMERO DE SERIAÇÃO']/ancestor::div[@class='row']/following-sibling::div[@class='row']//div[contains(@class, 'ng-binding')]")
            dados["NÚMERO DE SERIAÇÃO"] = ", ".join([el.text.strip() for el in seriacao_elements if el.text.strip()])
            pn_match = re.search(r'(\d+-\d+-\d+)', dados.get("PN", "")); dados["PN extraído"] = pn_match.group(1) if pn_match else "Não encontrado"
            dados["REV. FSE"] = dados.get("REV. PN", "Não encontrada")

            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return dados
        except Exception:
            oc_str = f"{oc1}/{oc2}"
            self.registrar_log(f"ERRO: Falha ao extrair dados da FSE para a OC {oc_str}.")
            self.tirar_print_de_erro(oc_str.replace('/', '-'), "extracao_FSE")
            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return None
    
    def navegar_para_fse_busca(self, wait):
        original_window = self.driver.current_window_handle
        wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
        wait.until(EC.number_of_windows_to_be(2))
        for handle in self.driver.window_handles:
            if handle != original_window: self.driver.switch_to.window(handle); break
        self.prompt_user_action("No navegador, navegue para 'FSE' > 'Busca FSe' e clique em 'Continuar'.")

    def navegar_para_desenhos_engenharia(self, wait):
        self.driver.switch_to.window(self.driver.window_handles[0])
        self.driver.get("https://web.embraer.com.br/irj/portal")
        wait.until(EC.element_to_be_clickable((By.ID, "L2N1"))).click()
        self.prompt_user_action("Valide se a tela 'Desenhos Engenharia' está aberta e clique em 'Continuar'.")
    
    def find_and_click(self, wait, selectors, description):
        for selector in selectors:
            try:
                element = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                ActionChains(self.driver).move_to_element(element).click().perform()
                return True
            except TimeoutException: continue
        self.registrar_log(f"AVISO: Não foi possível clicar no elemento '{description}'.")
        return False

def buscar_revisao_engenharia(self, wait, part_number):
    try:
        if not part_number or part_number == "Não encontrado":
            return "PN não fornecido"

        self.driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[starts-with(@id, 'ivuFrm_')]")))

        campo_pn = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@id, 'PartNumber')]")))
        campo_pn.clear()
        campo_pn.send_keys(part_number)
        self.registrar_log(f"Buscando revisão de engenharia para o PN: {part_number}")
        
        self.find_and_click(wait, ['//*[@id="FOAH.Dplpl049View.cmdGBI"]'], "Botão Desenho")

        
        seletores_voltar_universal = [
            '//*[@id="FOAHJJEL.GbiMenu.cmdRetornarNaveg"]',  
            '//*[@id="FOAH.Dplpl049View.cmdVoltar"]',        
            "//div[contains(@ct, 'B') and .//span[normalize-space()='Voltar']]" 
        ]

        try:
            seletor_rev = '//*[@id="FOAHJJEL.GbiMenu.TreeNodeType1.0.childNode.0.childNode.0.childNode.0.childNode.0-cnt-start"]'
            rev_element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, seletor_rev)))
            revisao = rev_element.text.strip()
            self.registrar_log(f"Revisão de Engenharia encontrada: {revisao}")
            
            self.find_and_click(wait, seletores_voltar_universal, "Botão Voltar (Sucesso)")
            return revisao
            
        except TimeoutException:
            self.registrar_log(f"AVISO: PN {part_number} não foi encontrado no sistema de Engenharia.")
            self.tirar_print_de_erro(part_number, "busca_revisao_nao_encontrado")
            
            self.find_and_click(wait, seletores_voltar_universal, "Botão Voltar (Tela de Erro)")
            return "Não encontrado em ENG"

    except Exception:
        self.registrar_log(f"ERRO inesperado ao buscar revisão para PN {part_number}: {traceback.format_exc()}")
        self.tirar_print_de_erro(part_number, "busca_revisao_erro_inesperado")
        return "Falha na busca"
    finally:
        self.driver.switch_to.default_content()

    def safe_find_text(self, by, value):
        try: return self.driver.find_element(by, value).text
        except NoSuchElementException: return ""

    def tirar_print_de_erro(self, identificador, etapa):
        os.makedirs(ERROS_DIR, exist_ok=True)
        id_limpo = re.sub(r'[\\/*?:"<>|]', "", str(identificador))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(ERROS_DIR, f"erro_{etapa}_{id_limpo}_{timestamp}.png")
        try:
            if self.driver:
                self.driver.save_screenshot(screenshot_path)
                self.registrar_log(f"Screenshot do erro salvo em: '{screenshot_path}'")
        except Exception as e:
            self.registrar_log(f"FALHA AO SALVAR SCREENSHOT: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ValidadorGUI(root)
    root.mainloop()