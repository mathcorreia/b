import os
import time
import shutil
import pandas as pd
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
from selenium.common.exceptions import TimeoutException

DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'Downloads')
PASTA_RAIZ_VERIFICACAO = r'\\fserver\cedoc_docs\Doc - Embraer Produtivo'
LOG_FILENAME = 'log_automacao.txt'
ERRO_LOG_FILENAME = 'log_erros.txt'

class AutomatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automator Embraer Produtivo - Painel de Controle")
        self.root.geometry("850x650")
    
        
        # Evento para sincronizar a thread de automação com as ações do usuário na GUI
        self.user_action_event = threading.Event()
        self.reprocess_choice = "" # Armazena a escolha do usuário para reprocessamento

        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(expand=True, fill='both')

        # Frame para o status e botões de ação
        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill='x', pady=(0, 5))

        self.label_status = tk.Label(top_frame, text="Pronto para iniciar.", font=("Helvetica", 12, "bold"), fg="#00529B", pady=10, wraplength=700, justify='center')
        self.label_status.pack()

        self.action_button = tk.Button(top_frame, text="Iniciar Automação", command=self.iniciar_automacao_thread, font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", padx=20, pady=10)
        self.action_button.pack(pady=(5, 10))

        # Frame para os botões de reprocessamento (inicialmente oculto)
        self.reprocess_frame = tk.Frame(main_frame)
        self.reprocess_button = tk.Button(self.reprocess_frame, text="Reprocessar Erros", command=lambda: self.set_reprocess_choice("reprocess"), font=("Helvetica", 10, "bold"), bg="#FFA500", fg="white")
        self.finish_button = tk.Button(self.reprocess_frame, text="Finalizar", command=lambda: self.set_reprocess_choice("finish"), font=("Helvetica", 10))
        self.reprocess_button.pack(side='left', padx=5)
        self.finish_button.pack(side='left', padx=5)

        # Log em tempo real
        log_label = tk.Label(main_frame, text="Log em Tempo Real:", font=("Helvetica", 10, "bold"))
        log_label.pack(fill='x', pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(main_frame, state='disabled', wrap=tk.WORD, font=("Courier New", 9))
        self.log_text.pack(expand=True, fill='both', pady=5)
        
        self.log_path = os.path.join(os.getcwd(), LOG_FILENAME)
        self.erro_log_path = os.path.join(os.getcwd(), ERRO_LOG_FILENAME)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.driver = None

    def on_closing(self):
        """Garante que o driver do chrome seja fechado ao fechar a janela."""
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
        def do_update():
            self.label_status.config(text=text, fg=color)
        self.root.after(0, do_update)

    def iniciar_automacao_thread(self):
        self.action_button.config(state='disabled', text="Executando...")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        automation_thread = threading.Thread(target=self.run_automation)
        automation_thread.daemon = True
        automation_thread.start()

    def prompt_user_action(self, message):
        """Pausa a automação e pede uma ação do usuário na GUI principal."""
        self.user_action_event.clear()
        
        def setup_gui_for_action():
            self.update_status(message, color="#E69500") # Laranja para ação necessária
            self.action_button.config(text="Continuar (Após realizar a ação)", command=self.signal_user_action, state="normal")
        
        self.root.after(0, setup_gui_for_action)
        self.user_action_event.wait() 

        def reset_gui_after_action():
            self.action_button.config(state='disabled', text="Executando...")
        
        self.root.after(0, reset_gui_after_action)

    def signal_user_action(self):
        """Sinaliza para a thread de automação que o usuário completou a ação."""
        self.user_action_event.set()

    def run_automation(self):
        try:
            self.update_status("Iniciando configuração...")
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
            except locale.Error:
                locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil')

            hoje = datetime.now()
            
            meses_abrev_pt = {
                "janeiro": "Jan", "fevereiro": "Fev", "março": "Mar", "abril": "Abr",
                "maio": "Mai", "junho": "Jun", "julho": "Jul", "agosto": "Ago",
                "setembro": "Set", "outubro": "Out", "novembro": "Nov", "dezembro": "Dez"
            }
            nome_mes_atual = hoje.strftime("%B").lower()
            nome_mes_abrev = meses_abrev_pt.get(nome_mes_atual, "Err")

            PASTA_MES_NOME = f"{hoje.strftime('%Y_%m')}-{nome_mes_abrev}"
            PASTA_MES = os.path.join(PASTA_RAIZ_VERIFICACAO, PASTA_MES_NOME)

            pastas_destino = {
                'LM': os.path.join(PASTA_MES, 'LM'),
                'LP': os.path.join(PASTA_MES, 'LP'),
                'FS': os.path.join(PASTA_MES, 'FS')
            }
            for pasta in pastas_destino.values():
                os.makedirs(pasta, exist_ok=True)
            self.registrar_log(f"Pasta de destino do mês atual: {PASTA_MES}")

            self.update_status("Lendo arquivo Excel 'lista.xlsx'...")
            df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
            df.rename(columns={df.columns[0]: 'OS'}, inplace=True)
            df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
            self.registrar_log(f"Arquivo Excel lido. Total de {len(df)} itens na lista.")

            # ### INÍCIO DA ALTERAÇÃO: LÓGICA DE VERIFICAÇÃO COM LIMITE DE 2 ANOS ###
            self.update_status("Verificando arquivos já existentes (até 2 anos retroativos)...")
            arquivos_existentes = set()
            padrao_pasta_mes = re.compile(r'^\d{4}_\d{2}-\w{3}$') # Padrão YYYY_MM-Mon
            
            # Define os anos a serem verificados: o ano atual e os dois anteriores.
            anos_a_verificar = [str(ano) for ano in range(hoje.year, hoje.year - 3, -1)] # Ex: [ '2025', '2024', '2023' ]
            self.registrar_log(f"Anos a serem verificados para arquivos existentes: {', '.join(anos_a_verificar)}")

            if os.path.exists(PASTA_RAIZ_VERIFICACAO):
                self.registrar_log(f"Verificando pastas em: {PASTA_RAIZ_VERIFICACAO}...")
                
                for nome_pasta in os.listdir(PASTA_RAIZ_VERIFICACAO):
                    caminho_pasta = os.path.join(PASTA_RAIZ_VERIFICACAO, nome_pasta)
                    
                    # Extrai o ano do nome da pasta (primeiros 4 caracteres)
                    ano_da_pasta = nome_pasta[:4]

                    # CONDIÇÃO ADICIONADA: Verifica se o ano da pasta está na lista de anos permitidos
                    if os.path.isdir(caminho_pasta) and ano_da_pasta in anos_a_verificar and padrao_pasta_mes.match(nome_pasta):
                        # Percorre a estrutura de subpastas (LM, LP, FS)
                        for _, _, files in os.walk(caminho_pasta):
                            for nome_arquivo in files:
                                if nome_arquivo.endswith(".pdf"):
                                    os_num = nome_arquivo.split('_')[0]
                                    if os_num.isdigit():
                                        arquivos_existentes.add(f"{os_num}_LM.pdf")
                                        arquivos_existentes.add(f"{os_num}_LP.pdf")
                                        arquivos_existentes.add(f"{os_num}_FS.pdf")
            # ### FIM DA ALTERAÇÃO ###
            
            df['OS_str'] = df['OS'].astype(str)
            df['ja_existe_lm'] = df['OS_str'].apply(lambda x: f"{x}_LM.pdf" in arquivos_existentes)
            df['ja_existe_lp'] = df['OS_str'].apply(lambda x: f"{x}_LP.pdf" in arquivos_existentes)
            df['ja_existe_fs'] = df['OS_str'].apply(lambda x: f"{x}_FS.pdf" in arquivos_existentes)
            df_filtrado = df[~(df['ja_existe_lm'] & df['ja_existe_lp'] & df['ja_existe_fs'])]
            
            removidos = len(df) - len(df_filtrado)
            self.registrar_log(f"Verificação concluída. {removidos} OSs foram removidas por já estarem completas.")

            if df_filtrado.empty:
                self.update_status("Todos os itens já foram baixados. Automação finalizada.", "#008A00")
                self.action_button.pack_forget()
                return

            self.update_status("Configurando o navegador...")
            caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
            service = ChromeService(executable_path=caminho_chromedriver)
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_experimental_option("prefs", {"download.default_directory": DOWNLOAD_DIR})
            self.driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(self.driver, 30) # Wait longo para ações gerais

            self.driver.get("https://web.embraer.com.br/irj/portal")
            self.prompt_user_action("Faça o login no portal. Quando a página principal carregar, clique em 'Continuar'.")

            self.update_status("Navegando para a aplicação GFS...")
            original_window = self.driver.current_window_handle
            wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
            wait.until(EC.number_of_windows_to_be(2))
            for handle in self.driver.window_handles:
                if handle != original_window:
                    self.driver.switch_to.window(handle)
                    break
            self.registrar_log("Foco alterado para a nova aba da aplicação GFS.")

            self.prompt_user_action("No navegador, navegue para 'FSE' > 'Busca FSe'. Quando a tela de busca carregar, clique em 'Continuar'.")

            for index, row in df_filtrado.iterrows():
                self.update_status(f"Processando OS: {row['OS_str']}...")
                self.processar_uma_os(wait, row, pastas_destino)
            
            self.reprocessar_erros(df_filtrado, wait, pastas_destino)

        except Exception as e:
            error_details = traceback.format_exc()
            self.registrar_log(f"ERRO CRÍTICO: {error_details}")
            self.update_status(f"Erro Crítico: {e}", "red")
        finally:
            if self.driver:
                self.registrar_log("Automação finalizada.")
                self.driver.quit()
                self.driver = None
            if not self.reprocess_choice:
                self.update_status("Processo finalizado!", "#008A00")
                self.action_button.pack_forget()

    def esperar_download_concluir(self, pasta_download, timeout=45):
        arquivos_antes = set(f for f in os.listdir(pasta_download) if f.endswith('.pdf'))
        segundos = 0
        while segundos < timeout:
            if not any(f.endswith('.crdownload') for f in os.listdir(pasta_download)):
                arquivos_depois = set(f for f in os.listdir(pasta_download) if f.endswith('.pdf'))
                novos_arquivos = arquivos_depois - arquivos_antes
                if novos_arquivos:
                    nome_novo_arquivo = novos_arquivos.pop()
                    caminho_completo = os.path.join(pasta_download, nome_novo_arquivo)
                    self.registrar_log(f"Download detectado: {nome_novo_arquivo}")
                    time.sleep(0.5) 
                    return caminho_completo
            time.sleep(1)
            segundos += 1
        self.registrar_log(f"ERRO: Timeout ({timeout}s) esperando download.")
        return None

    def processar_uma_os(self, wait, row, pastas_destino):
        os_num = row['OS_str']
        oc1, oc2 = row['OC_antes'], row['OC_depois']
        try:
            self.registrar_log(f"--- Processando OS: {os_num} | OC: {oc1}/{oc2} ---")
            
            campo_oc1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']")))
            campo_oc1.clear()
            campo_oc1.send_keys(oc1)
            campo_oc2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']")))
            campo_oc2.clear()
            campo_oc2.send_keys(oc2)
            wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
            
            try:
                WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
            except TimeoutException:
                self.registrar_log(f"ERRO: OS {os_num} não encontrada ou a busca falhou. Pulando para a próxima.")
                self.tirar_print_de_erro(os_num)
                return 
            
            docs_a_processar = {
                'LM': {"seletor": "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[1]", "existe": row['ja_existe_lm']},
                'LP': {"seletor": "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[2]", "existe": row['ja_existe_lp']},
                'FS': {"seletor": "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[3]/button[2]", "existe": row['ja_existe_fs']}
            }

            for tipo, info in docs_a_processar.items():
                if info['existe']:
                    self.registrar_log(f"SKIP ({tipo}): Documento para OS {os_num} já existe.")
                    continue
                try:
                    time.sleep(1) 
                    button = wait.until(EC.element_to_be_clickable((By.XPATH, info['seletor'])))
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", button)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", button)
                    caminho_arquivo = self.esperar_download_concluir(DOWNLOAD_DIR)
                    if caminho_arquivo:
                        novo_nome = f"{os_num}_{tipo}.pdf"
                        destino = os.path.join(pastas_destino[tipo], novo_nome)
                        shutil.move(caminho_arquivo, destino)
                        self.registrar_log(f"SUCESSO ({tipo}): Arquivo salvo como {novo_nome}")
                    else:
                        self.registrar_log(f"ERRO ({tipo}): Download não concluído para a OS {os_num}")
                except TimeoutException:
                    self.registrar_log(f"AVISO ({tipo}): Botão para '{tipo}' não encontrado para a OS {os_num}.")
            
            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
        except Exception as e:
            self.registrar_log(f"ERRO GERAL com OS {os_num}: {e}")
            self.tirar_print_de_erro(os_num)
            self.driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")

    def tirar_print_de_erro(self, os_num):
        timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_screenshot = f"erro_os_{os_num}_{timestamp_erro}.png"
        screenshot_path = os.path.join(os.getcwd(), nome_screenshot)
        try:
            if self.driver:
                self.driver.save_screenshot(screenshot_path)
                self.registrar_log(f"Screenshot de erro salvo em: '{screenshot_path}'")
        except Exception as screenshot_error:
            self.registrar_log(f"FALHA AO SALVAR SCREENSHOT: {screenshot_error}")

    def reprocessar_erros(self, df_original, wait, pastas_destino):
        self.registrar_log("--- Verificando erros para reprocessar ---")
        erros_os = set()
        with open(self.log_path, 'r', encoding='utf-8') as log_file:
            linhas_de_erro = [linha for linha in log_file if "ERRO" in linha or "AVISO" in linha]
        if not linhas_de_erro:
            self.update_status("Nenhum erro encontrado. Processo finalizado com sucesso!", "#008A00")
            return

        with open(self.erro_log_path, 'w', encoding='utf-8') as erro_log_file:
            erro_log_file.writelines(linhas_de_erro)
        for linha in linhas_de_erro:
            match = re.search(r'OS (\d+)', linha)
            if match: erros_os.add(match.group(1))

        self.update_status(f"{len(erros_os)} OSs com erros ou avisos encontradas. Deseja tentar baixá-las novamente?", "#E69500")
        self.action_button.pack_forget()
        self.reprocess_frame.pack()

        while not self.reprocess_choice:
            time.sleep(0.1)
        self.reprocess_frame.pack_forget()

        if self.reprocess_choice == "reprocess":
            # Recalcula quais documentos precisam ser baixados
            df_erros = df_original[df_original['OS_str'].isin(erros_os)].copy()
            df_erros['ja_existe_lm'] = df_erros['OS_str'].apply(lambda x: os.path.exists(os.path.join(pastas_destino['LM'], f"{x}_LM.pdf")))
            df_erros['ja_existe_lp'] = df_erros['OS_str'].apply(lambda x: os.path.exists(os.path.join(pastas_destino['LP'], f"{x}_LP.pdf")))
            df_erros['ja_existe_fs'] = df_erros['OS_str'].apply(lambda x: os.path.exists(os.path.join(pastas_destino['FS'], f"{x}_FS.pdf")))
            
            self.registrar_log("--- Iniciando reprocessamento dos erros ---")
            for index, row in df_erros.iterrows():
                self.update_status(f"Reprocessando OS: {row['OS_str']}...")
                self.processar_uma_os(wait, row, pastas_destino)
            self.update_status("Reprocessamento finalizado!", "#008A00")
        else:
            self.registrar_log("Reprocessamento ignorado pelo usuário.")
            self.update_status("Finalizado. Itens com erro não foram reprocessados.", "#00529B")

    def set_reprocess_choice(self, choice):
        self.reprocess_choice = choice

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatorGUI(root)
    root.mainloop()