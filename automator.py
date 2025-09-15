import os
import time
import shutil
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- 1. CONFIGURAÇÕES DINÂMICAS E ROBUSTAS ---

# Dicionário para traduzir o nome do mês para Português
MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
    7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

# Obtém a data atual para criar pastas dinamicamente
hoje = datetime.now()
ano_atual_str = str(hoje.year)
# O formato do número do mês parece ser 100 + número do mês (ex: Setembro = 109)
# Ajustei o código original que usava '110' para '109' para Setembro, que é o mês 9.
# Se o padrão for outro, basta ajustar a linha abaixo.
mes_atual_str = f"{100 + hoje.month} - {MESES_PT[hoje.month]}"

# Define o diretório de downloads do usuário (funciona para Windows, Linux, Mac)
try:
    DOWNLOAD_DIR = os.path.join(os.environ['USERPROFILE'], 'Downloads')
except KeyError:
    DOWNLOAD_DIR = os.path.expanduser('~/Downloads')

# Define os caminhos de pasta de forma dinâmica
PASTA_BASE = r'\\fserver\cedoc_docs\Doc - EmbraerProdutivo'
PASTA_ANO = os.path.join(PASTA_BASE, ano_atual_str)
PASTA_DESTINO = os.path.join(PASTA_ANO, mes_atual_str)

# Cria as pastas de ano e mês se elas não existirem
os.makedirs(PASTA_DESTINO, exist_ok=True)

# Define o caminho do arquivo de log
LOG_PATH = os.path.join(os.getcwd(), 'log_automacao.txt')

# Função para registrar logs de execução
def registrar_log(mensagem):
    """Registra uma mensagem com data e hora no arquivo de log."""
    with open(LOG_PATH, 'a', encoding='utf-8') as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")

# --- 2. LEITURA DOS DADOS E INICIALIZAÇÃO DO NAVEGADOR ---

registrar_log("--- Início da Automação ---")

# Tenta ler os dados da planilha Excel
try:
    df = pd.read_excel('lista.xlsx', sheet_name='baixar_lm', engine='openpyxl', dtype=str)
    # Garante que a coluna seja lida como texto para evitar problemas com o split
    df[['OC_antes', 'OC_depois']] = df.iloc[:, 1].str.split('/', expand=True)
    registrar_log(f"Planilha 'lista.xlsx' lida com sucesso. {len(df)} itens para processar.")
except FileNotFoundError:
    registrar_log("Erro Crítico: Arquivo 'lista.xlsx' não encontrado. A automação não pode continuar.")
    raise
except Exception as e:
    registrar_log(f"Erro Crítico ao ler o Excel: {e}")
    raise

# Configurações do Microsoft Edge
options = Options()
options.use_chromium = True
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Inicializa o driver do Edge
# Garanta que o msedgedriver.exe está no PATH do sistema ou especifique o caminho.
try:
    service = Service()
    driver = webdriver.Edge(service=service, options=options)
except Exception as e:
    registrar_log(f"Erro Crítico ao iniciar o WebDriver: {e}")
    registrar_log("Verifique se o EdgeDriver está instalado e acessível no PATH do sistema.")
    raise

# --- 3. EXECUÇÃO DA AUTOMAÇÃO ---

# Abre a página e aguarda o login manual do usuário
driver.get("https://web.embraer.com.br/")
print(">>> AÇÃO NECESSÁRIA <<<")
input("Faça o login manualmente, navegue até a página de busca de notas e pressione ENTER para continuar...")
registrar_log("Login manual detectado. Iniciando o loop de downloads.")

# Loop principal para buscar e baixar os arquivos
for index, row in df.iterrows():
    # Converte para string para garantir que não haja erros de tipo
    oc1 = str(row['OC_antes']).strip()
    oc2 = str(row['OC_depois']).strip()

    try:
        registrar_log(f"Processando OC: {oc1}/{oc2}")

        # Aguarda até que o campo OC esteja presente e o preenche
        input_oc1 = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@ng-model='vm.search.orderNumber']"))
        )
        input_oc1.clear()
        input_oc1.send_keys(oc1)

        # Preenche o campo da linha da OC
        input_oc2 = driver.find_element(By.XPATH, "//input[@ng-model='vm.search.orderLine']")
        input_oc2.clear()
        input_oc2.send_keys(oc2)

        # Clica em "Buscar"
        driver.find_element(By.ID, "searchBtn").click()

        # Espera o resultado aparecer e clica na "lupa"
        lupa = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "glyphicon-search"))
        )
        lupa.click()

        # Clica em "Lista de Materiais"
        btn_lista_materiais = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Lista de Materiais')]"))
        )

        # Lógica para esperar o download de forma inteligente
        arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
        btn_lista_materiais.click()

        # Espera até que um novo arquivo apareça na pasta de downloads (timeout de 60s)
        tempo_espera = 0
        novo_arquivo = None
        while tempo_espera < 60:
            arquivos_depois = set(os.listdir(DOWNLOAD_DIR))
            novos_arquivos = arquivos_depois - arquivos_antes
            # Verifica se há um novo arquivo .pdf que não seja temporário
            arquivos_pdf_finalizados = [f for f in novos_arquivos if f.lower().endswith('.pdf') and not f.lower().endswith('.crdownload')]
            
            if arquivos_pdf_finalizados:
                novo_arquivo = arquivos_pdf_finalizados[0]
                break
            time.sleep(1)
            tempo_espera += 1
        
        if not novo_arquivo:
            raise Exception("O download do arquivo não foi detectado em 60 segundos.")

        registrar_log(f"Download detectado: {novo_arquivo}")
        
        # Espera um segundo extra para garantir que o arquivo foi completamente escrito no disco
        time.sleep(1)

        # Move e renomeia o arquivo
        nome_arquivo_novo = f"{oc1}-{oc2}.pdf"
        origem = os.path.join(DOWNLOAD_DIR, novo_arquivo)
        destino = os.path.join(PASTA_DESTINO, nome_arquivo_novo)

        if not os.path.exists(destino):
            shutil.move(origem, destino)
            registrar_log(f"Movido e renomeado: '{novo_arquivo}' para '{nome_arquivo_novo}' em '{PASTA_DESTINO}'")
        else:
            # Se o arquivo já existe, adiciona um carimbo de data/hora para não sobrescrever
            base, ext = os.path.splitext(nome_arquivo_novo)
            timestamp = datetime.now().strftime("_%Y%m%d%H%M%S")
            destino_alternativo = os.path.join(PASTA_DESTINO, f"{base}{timestamp}{ext}")
            shutil.move(origem, destino_alternativo)
            registrar_log(f"Arquivo '{nome_arquivo_novo}' já existia. Movido como '{os.path.basename(destino_alternativo)}'.")
            # Deleta o arquivo original que não foi movido
            if os.path.exists(origem):
                os.remove(origem)


    except Exception as e:
        registrar_log(f"ERRO com OC {oc1}/{oc2}: {e}")
        # Opcional: recarregar a página para tentar se recuperar de um estado de erro
        driver.get(driver.current_url) 
        time.sleep(2) # Pausa para a página recarregar

# --- 4. FINALIZAÇÃO ---

# A lógica de copiar para a pasta do mês seguinte foi removida por ser desnecessária
# e potencialmente incorreta. O script agora sempre salva na pasta do mês correto.

registrar_log("--- Automação Finalizada ---")
print("\nAutomação finalizada. Verifique o arquivo 'log_automacao.txt' para detalhes.")
driver.quit()