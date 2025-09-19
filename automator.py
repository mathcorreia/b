import os
import time
import shutil
import pandas as pd
import locale
import re
import traceback
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- CONSTANTES GLOBAIS DE CONFIGURAÇÃO ---
DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'Downloads')
# ATENÇÃO: Esta é a pasta raiz onde a verificação de 2 anos será feita.
PASTA_RAIZ_AUTOMATOR = r'\\fserver\cedoc_docs\Doc - EmbraerProdutivo'
LOG_FILENAME = 'log_automacao.txt'
ERRO_LOG_FILENAME = 'log_erros.txt'

class AutomatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automator Embraer Produtivo")
        self.root.geometry("800x600")

        self.label_status = tk.Label(root, text="Pronto para iniciar.", font=("Helvetica", 12), pady=10)
        self.label_status.pack()

        self.log_text = scrolledtext.ScrolledText(root, state='disabled', wrap=tk.WORD, font=("Courier New", 9))
        self.log_text.pack(expand=True, fill='both', padx=10, pady=5)

        self.start_button = tk.Button(root, text="Iniciar Automação", command=self.iniciar_automacao_thread, font=("Helvetica", 12, "bold"), bg="#4CAF50", fg="white", padx=20, pady=10)
        self.start_button.pack(pady=10)
        
        # Caminhos dos logs serão definidos no início da automação
        self.log_path = os.path.join(os.getcwd(), LOG_FILENAME)
        self.erro_log_path = os.path.join(os.getcwd(), ERRO_LOG_FILENAME)

    def registrar_log(self, mensagem):
        """Registra a mensagem no arquivo de log e na interface gráfica de forma thread-safe."""
        log_entry = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n"
        
        # Escreve no arquivo
        with open(self.log_path, 'a', encoding='utf-8') as log_file:
            log_file.write(log_entry)

        # Atualiza a GUI
        def update_gui():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END) # Auto-scroll
            self.log_text.config(state='disabled')
        
        self.root.after(0, update_gui)

    def update_status(self, text):
        """Atualiza o label de status na GUI de forma thread-safe."""
        def do_update():
            self.label_status.config(text=text)
        self.root.after(0, do_update)

    def iniciar_automacao_thread(self):
        """Inicia a lógica de automação em uma nova thread para não congelar a GUI."""
        self.start_button.config(state='disabled', text="Executando...")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END) # Limpa o log da execução anterior
        self.log_text.config(state='disabled')
        
        automation_thread = threading.Thread(target=self.run_automation)
        automation_thread.daemon = True
        automation_thread.start()

    def verificar_documento_existente(self, os_num, tipo_doc):
        """
        Verifica se um documento para uma OS já existe na PASTA_RAIZ_AUTOMATOR,
        com uma retroatividade de 2 anos.
        
        Args:
            os_num (str): O número da Ordem de Serviço.
            tipo_doc (str): O tipo de documento ('LM', 'LP' ou 'FS').

        Returns:
            bool: True se o arquivo for encontrado, False caso contrário.
        """
        nome_arquivo = f"{os_num}_{tipo_doc}.pdf"
        ano_atual = datetime.now().year
        anos_a_verificar = [str(ano) for ano in range(ano_atual, ano_atual - 3, -1)]

        self.registrar_log(f"VERIFICANDO: Buscando por '{nome_arquivo}' nos anos {anos_a_verificar}...")

        for ano in anos_a_verificar:
            caminho_ano = os.path.join(PASTA_RAIZ_AUTOMATOR, ano)
            if not os.path.isdir(caminho_ano):
                continue

            for dirpath, _, filenames in os.walk(caminho_ano):
                if nome_arquivo in filenames:
                    caminho_encontrado = os.path.join(dirpath, nome_arquivo)
                    self.registrar_log(f"EXISTENTE (SKIP): Documento '{nome_arquivo}' já encontrado em: {caminho_encontrado}")
                    return True
        
        self.registrar_log(f"NÃO ENCONTRADO: '{nome_arquivo}' não existe. Prosseguindo com o download.")
        return False

    def run_automation(self):
        """Contém toda a lógica principal da automação."""
        driver = None
        try:
            # --- CONFIGURAÇÃO INICIAL ---
            self.update_status("Iniciando configuração...")
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
            except locale.Error:
                locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil')

            hoje = datetime.now()
            nome_mes_atual = hoje.strftime("%B").capitalize()
            # O cálculo do mês estava somando 100, ajustei para um formato mais comum (ex: 09)
            num_mes_atual = f"{hoje.month:02d}"

            PASTA_ANO_ATUAL = os.path.join(PASTA_RAIZ_AUTOMATOR, str(hoje.year))
            MES_ATUAL = f'{num_mes_atual} - {nome_mes_atual}'
            PASTA_MES = os.path.join(PASTA_ANO_ATUAL, MES_ATUAL)

            PASTA_DESTINO_LM = os.path.join(PASTA_MES, 'LM')
            PASTA_DESTINO_LP = os.path.join(PASTA_MES, 'LP')
            PASTA_DESTINO_FS = os.path.join(PASTA_MES, 'FS')

            os.makedirs(PASTA_DESTINO_LM, exist_ok=True)
            os.makedirs(PASTA_DESTINO_LP, exist_ok=True)
            os.makedirs(PASTA_DESTINO_FS, exist_ok=True)
            
            self.registrar_log(f"Pasta de destino do mês atual: {PASTA_MES}")

            # --- LEITURA DO EXCEL ---
            self.update_status("Lendo arquivo Excel...")
            self.registrar_log("Lendo arquivo Excel 'lista.xlsx'...")
            try:
                df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl')
                df.rename(columns={df.columns[0]: 'OS'}, inplace=True)
                df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].astype(str).str.split('/', expand=True, n=1)
                self.registrar_log(f"Arquivo Excel lido com sucesso. {len(df)} itens para processar.")
            except FileNotFoundError:
                self.registrar_log("ERRO CRÍTICO: Arquivo 'lista.xlsx' não encontrado.")
                messagebox.showerror("Erro", "O arquivo 'lista.xlsx' não foi encontrado na pasta do programa.")
                return # Encerra a execução da automação

            # --- CONFIGURAÇÃO DO NAVEGADOR ---
            self.update_status("Configurando o navegador...")
            self.registrar_log("Configurando o Chrome...")
            caminho_chromedriver = os.path.join(os.getcwd(), "chromedriver.exe")
            if not os.path.exists(caminho_chromedriver):
                self.registrar_log("ERRO CRÍTICO: 'chromedriver.exe' não encontrado.")
                messagebox.showerror("Erro", "'chromedriver.exe' não foi encontrado na pasta do programa.")
                return

            service = ChromeService(executable_path=caminho_chromedriver)
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_experimental_option("prefs", {
                "download.default_directory": DOWNLOAD_DIR,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            })

            self.registrar_log("Iniciando o WebDriver do Chrome...")
            driver = webdriver.Chrome(service=service, options=options)
            wait = WebDriverWait(driver, 30)
            self.registrar_log("WebDriver iniciado com sucesso.")

            # --- LOGIN E NAVEGAÇÃO ---
            driver.get("https://web.embraer.com.br/irj/portal")
            messagebox.showinfo("Ação Necessária", "Faça o login e, quando estiver na página principal do portal, clique em OK para continuar.")

            self.update_status("Navegando para a aplicação GFS...")
            self.registrar_log("Iniciando navegação para GFS...")
            original_window = driver.current_window_handle
            wait.until(EC.element_to_be_clickable((By.ID, "L2N10"))).click()
            self.registrar_log("Clicou no link 'GFS'.")

            wait.until(EC.number_of_windows_to_be(2))
            for window_handle in driver.window_handles:
                if window_handle != original_window:
                    driver.switch_to.window(window_handle)
                    break
            self.registrar_log("Foco alterado para a nova aba da aplicação GFS.")

            messagebox.showinfo("Ação Necessária", "Robô na aba correta.\n\nAGORA, clique em 'FSE' > 'Busca FSe' e, quando a tela de busca carregar, clique em OK para o robô começar a trabalhar.")

            # --- LOOP DE PROCESSAMENTO PRINCIPAL ---
            self.registrar_log("Iniciando processamento principal do Excel...")
            for index, row in df.iterrows():
                os_num = str(row['OS'])
                oc1 = row['OC_antes']
                oc2 = row['OC_depois']
                self.update_status(f"Processando OS: {os_num}")
                self.processar_uma_os(driver, wait, os_num, oc1, oc2, {
                    'LM': PASTA_DESTINO_LM, 'LP': PASTA_DESTINO_LP, 'FS': PASTA_DESTINO_FS
                })
            
            self.update_status("Verificando erros para reprocessamento...")
            self.reprocessar_erros(df, driver, wait, {
                'LM': PASTA_DESTINO_LM, 'LP': PASTA_DESTINO_LP, 'FS': PASTA_DESTINO_FS
            })

            self.update_status("Automação concluída com sucesso!")
            messagebox.showinfo("Fim", "Processo de automação concluído. Verifique o log para detalhes.")

        except Exception as e:
            error_details = traceback.format_exc()
            self.registrar_log(f"ERRO CRÍTICO: {error_details}")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro grave e a automação será encerrada.\n\nVerifique o log.\n\nErro: {e}")
            self.update_status("Erro crítico. Automação encerrada.")

        finally:
            if driver:
                self.registrar_log("Automação finalizada. Fechando o navegador.")
                driver.quit()
            
            # Reabilita o botão ao final da execução
            self.start_button.config(state='normal', text="Iniciar Automação")
            self.update_status("Pronto.")
    
    def esperar_download_concluir(self, pasta_download, timeout=60):
        """Espera um arquivo PDF ser completamente baixado na pasta."""
        # Limpa PDFs antigos para evitar pegar o arquivo errado
        for item in os.listdir(pasta_download):
            if item.endswith(".pdf"):
                try:
                    os.remove(os.path.join(pasta_download, item))
                except OSError as e:
                    self.registrar_log(f"Aviso: Não foi possível limpar o arquivo antigo {item}. Erro: {e}")

        segundos = 0
        while segundos < timeout:
            if not any(f.endswith('.crdownload') for f in os.listdir(pasta_download)):
                arquivos_pdf = [os.path.join(pasta_download, f) for f in os.listdir(pasta_download) if f.endswith('.pdf')]
                if arquivos_pdf:
                    return max(arquivos_pdf, key=os.path.getmtime)
            time.sleep(1)
            segundos += 1
        return None

    def processar_uma_os(self, driver, wait, os_num, oc1, oc2, pastas_destino):
        """Processa uma única linha do Excel (uma OS)."""
        try:
            self.registrar_log(f"--- Processando OS: {os_num} | OC: {oc1}/{oc2} ---")
            
            docs_a_baixar = []
            if not self.verificar_documento_existente(os_num, 'LM'): docs_a_baixar.append('LM')
            if not self.verificar_documento_existente(os_num, 'LP'): docs_a_baixar.append('LP')
            if not self.verificar_documento_existente(os_num, 'FS'): docs_a_baixar.append('FS')
            
            if not docs_a_baixar:
                self.registrar_log(f"Todos os documentos para a OS {os_num} já existem. Pulando para a próxima.")
                return True

            campo_oc1 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderNumber']")))
            campo_oc1.clear()
            campo_oc1.send_keys(oc1)

            campo_oc2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@ng-model='vm.search.orderLine']")))
            campo_oc2.clear()
            campo_oc2.send_keys(oc2)

            wait.until(EC.element_to_be_clickable((By.ID, "searchBtn"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'vm.showFseDetails')]"))).click()
            
            # Dicionário de seletores para simplificar
            seletores = {
                'LM': "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[1]",
                'LP': "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[1]/button[2]",
                'FS': "/html/body/main/div/ui-view/div/div[3]/fse-operations-form/div[1]/div[2]/div/div[3]/button[2]"
            }

            for tipo_doc in ['LM', 'LP', 'FS']:
                if tipo_doc in docs_a_baixar:
                    try:
                        time.sleep(2) # Pequena pausa entre downloads
                        seletor = (By.XPATH, seletores[tipo_doc])
                        wait.until(EC.element_to_be_clickable(seletor)).click()
                        
                        caminho_arquivo = self.esperar_download_concluir(DOWNLOAD_DIR)
                        if caminho_arquivo:
                            novo_nome = f"{os_num}_{tipo_doc}.pdf"
                            destino = os.path.join(pastas_destino[tipo_doc], novo_nome)
                            shutil.move(caminho_arquivo, destino)
                            self.registrar_log(f"SUCESSO ({tipo_doc}): Arquivo salvo como {novo_nome}")
                        else:
                            self.registrar_log(f"ERRO ({tipo_doc}): Download não concluído para a OS {os_num}")
                    except TimeoutException:
                        self.registrar_log(f"AVISO ({tipo_doc}): Botão para '{tipo_doc}' não encontrado para a OS {os_num}.")
            
            self.registrar_log(f"Processo da OS {os_num} concluído. Voltando para a página de busca.")
            driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return True

        except Exception as e:
            timestamp_erro = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_screenshot = f"erro_os_{os_num}_{timestamp_erro}.png"
            screenshot_path = os.path.join(os.getcwd(), nome_screenshot)
            driver.save_screenshot(screenshot_path)
            self.registrar_log(f"ERRO GERAL com OS {os_num}: {e} - Screenshot salvo em '{screenshot_path}'.")
            driver.get("https://appscorp2.embraer.com.br/gfs/#/fse/search/1")
            return False

    def reprocessar_erros(self, df_original, driver, wait, pastas_destino):
        """Analisa o log, encontra OS com erros e pergunta ao usuário se deseja reprocessá-las."""
        self.registrar_log("--- Fim do processamento principal. Verificando erros para reprocessar. ---")
        erros_os = set()
        linhas_de_erro = []
        try:
            with open(self.log_path, 'r', encoding='utf-8') as log_file:
                for linha in log_file:
                    if "ERRO" in linha or "AVISO" in linha:
                        linhas_de_erro.append(linha)
                        match = re.search(r'OS (\d+)', linha)
                        if match:
                            erros_os.add(match.group(1))
        except FileNotFoundError:
            self.registrar_log("Arquivo de log não encontrado. Nenhum item para reprocessar.")
            return

        if linhas_de_erro:
            with open(self.erro_log_path, 'w', encoding='utf-8') as erro_log_file:
                erro_log_file.write(f"--- Resumo de Erros e Avisos da execução de {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n\n")
                erro_log_file.writelines(linhas_de_erro)
            self.registrar_log(f"Log de erros filtrado foi salvo em: {self.erro_log_path}")

        if not erros_os:
            # Não exibe o messagebox se a janela principal já foi fechada
            if self.root.winfo_exists():
                messagebox.showinfo("Automação Concluída", "Nenhum erro encontrado na primeira passagem. Processo finalizado com sucesso!")
        else:
            self.registrar_log(f"Encontrados {len(erros_os)} itens com erro para reprocessar: {', '.join(sorted(erros_os))}")
            
            if self.root.winfo_exists():
                resposta = messagebox.askyesno("Reprocessamento de Erros", f"Foram encontrados {len(erros_os)} itens com erros ou avisos.\n\nO log de erros foi salvo em '{ERRO_LOG_FILENAME}'.\n\nDeseja tentar baixá-los novamente?")
                
                if resposta:
                    df_erros = df_original[df_original['OS'].astype(str).isin(erros_os)]
                    self.registrar_log("--- Iniciando reprocessamento dos erros. ---")
                    for index, row in df_erros.iterrows():
                        os_num = str(row['OS'])
                        oc1 = row['OC_antes']
                        oc2 = row['OC_depois']
                        self.update_status(f"Reprocessando OS: {os_num}")
                        self.processar_uma_os(driver, wait, os_num, oc1, oc2, pastas_destino)
                    self.registrar_log("--- Fim do reprocessamento. ---")
                    if self.root.winfo_exists():
                        messagebox.showinfo("Automação Concluída", "Reprocessamento finalizado. Verifique os logs para mais detalhes.")
                else:
                    self.registrar_log("Reprocessamento ignorado pelo usuário.")
                    if self.root.winfo_exists():
                        messagebox.showinfo("Automação Concluída", "Processo finalizado. Alguns itens apresentaram erros e não foram reprocessados.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatorGUI(root)
    root.mainloop()