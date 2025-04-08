import logging
import os
import time
from dotenv import load_dotenv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options

# Configuração do logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# Carregar variáveis de ambiente
dotenv_path_env = r'Z:\Comissionamento\20 - AUTÔNOMOS\60 - LUCAS KANG\01 - KANG BOTS\.env'
load_dotenv(dotenv_path=dotenv_path_env)
time.sleep(1)

# Obter credenciais
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
time.sleep(1)

# Verificação de credenciais
if not EMAIL or not PASSWORD:
    logger.error("Erro: As variáveis de ambiente EMAIL e PASSWORD não foram carregadas corretamente.")
    exit(10)
time.sleep(1)

# Carregar lista de parceiros e cashback da planilha
name_path = r"Z:\Comissionamento\20 - AUTÔNOMOS\2024\09 - Setembro\Financeiro\Cashback - Nomes sistema.xlsx"
df = pd.read_excel(name_path)
time.sleep(1)

# Verificar se as colunas "Parceiros" e "Cashback" existem
if "Parceiros" not in df.columns or "Cashback" not in df.columns:
    logger.error("Erro: A planilha não contém as colunas 'Parceiros' e 'Cashback'.")
    exit(10)
time.sleep(1)

# Lista
lista = df["Parceiros"].tolist()
time.sleep(1)

# Diretório de download
download_directory = r'Z:\Comissionamento\20 - AUTÔNOMOS\60 - LUCAS KANG\01 - KANG BOTS\Relatorios-extraidos-cashback'
time.sleep(1)

# Criar pastas
folder_50 = os.path.join(download_directory, "Relatórios Cashback 50%")
folder_100 = os.path.join(download_directory, "Relatórios Cashback 100%")
os.makedirs(folder_50, exist_ok=True)
os.makedirs(folder_100, exist_ok=True)
time.sleep(1)

if os.path.exists(folder_50) and os.path.exists(folder_100):
    print("Pastas já existentes no diretório")
time.sleep(1)

# Configurações do Chrome
chrome_options = Options()
prefs = {
    'download.default_directory': download_directory,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
}
chrome_options.add_experimental_option('prefs', prefs)
time.sleep(1)

# Iniciar o navegador
driver = webdriver.Chrome(options=chrome_options)
time.sleep(2)

# Acessar a página de login
driver.get("https://backofficeweb.genialinvestimentos.com.br/comissao/competencia")
time.sleep(3)

# Fazer login
email = driver.find_element(By.ID, "userLogin")
password = driver.find_element(By.ID, "input_0")
email.send_keys(EMAIL)
password.send_keys(PASSWORD)
login_button = driver.find_element(By.XPATH, "/html/body/div/div/section/section/md-card-content/form/div[2]/md-input-container/button")
login_button.click()
time.sleep(3)

# Navegar para as seções necessárias
comissao_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/ul/li[6]/div")
comissao_button.click()
time.sleep(3)

# Selecionar a aba de Cashback
titulo_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/div[1]/div[2]/md-tabs/md-tabs-wrapper/md-tabs-canvas/md-pagination-wrapper/md-tab-item[3]")
titulo_button.click()
time.sleep(3)

# Selecionar o mês "Janeiro"
selected_click = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/div[2]/div/md-content/div[2]/div/div[1]/div/md-input-container[2]")
selected_click.click()
time.sleep(3)

selected_mes = driver.find_element(By.XPATH, "//md-option/div[contains(text(), 'Janeiro')]")
selected_mes.click()
time.sleep(3)

# Selecionar o campo "Cashback"
selected_campo = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/div[2]/div/md-content/div[2]/div/div[1]/div/md-input-container[3]/md-select")
selected_campo.click()
time.sleep(3)

cashback_button = driver.find_element(By.XPATH, "//div[@class='md-text' and text()='Cashback']")
cashback_button.click()
time.sleep(3)

# Selecionar o parceiro
parceiro = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/div[2]/div/md-content/div[2]/div/div[4]/div[1]/md-chips/md-chips-wrap/div/div/md-autocomplete/md-autocomplete-wrap/input")
parceiro.click()
time.sleep(3)


# Função principal para processar os parceiros
def processar_parceiro(lista):

    def mover_para_pasta_correcta(file_path, cashback, nome_arquivo):
        if cashback == "50%":
            novo_caminho = os.path.join(folder_50, nome_arquivo)
            os.rename(file_path, novo_caminho)
            logger.info(f"Arquivo movido para a pasta Relatórios Cashback 50%: {nome_arquivo}")
        elif cashback == "100%":
            novo_caminho = os.path.join(folder_100, nome_arquivo)
            os.rename(file_path, novo_caminho)
            logger.info(f"Arquivo movido para a pasta Relatórios Cashback 100%: {nome_arquivo}")
        else:
            logger.error(f"Cashback {cashback} não reconhecido para o parceiro. Ignorando o arquivo {nome_arquivo}.")

    for nome in lista:
        try:
            parceiro.clear()
            parceiro.send_keys(nome)
            time.sleep(3)

            submit_button = driver.find_element(By.XPATH, "/html/body/md-virtual-repeat-container[4]")
            submit_button.click()
            time.sleep(3)

            extrair_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/div[2]/div/md-content/div[2]/div/div[4]/div[2]/div[1]/button")
            extrair_button.click()
            time.sleep(10)

            files = os.listdir(download_directory)
            paths = [os.path.join(download_directory, basename) for basename in files]
            temp_download_path = max(paths, key=os.path.getctime)
            time.sleep(5)

            new_file_name = f"Relatório-cashback-{nome}.xlsx"
            new_file_path = os.path.join(download_directory, new_file_name)
            os.rename(temp_download_path, new_file_path)
            logger.info(f"Relatório extraído e renomeado para: {new_file_name}")
            time.sleep(3)

            parceiro_cashback = df[df["Parceiros"].str.strip() == nome.strip()]
            if parceiro_cashback.empty:
                logger.error(f"Parceiro {nome} não encontrado na planilha de Cashback.")
                continue

            cashback = str(parceiro_cashback["Cashback"].iloc[0]).strip()
            logger.info(f"Cashback para {nome}: {cashback}")

            if cashback == "0.5":
                cashback = "50%"
            elif cashback == "1.0":
                cashback = "100%"
            else:
                logger.error(f"Cashback {cashback} não reconhecido para o parceiro {nome}.")
                continue

            mover_para_pasta_correcta(new_file_path, cashback, new_file_name)
            time.sleep(3)
        except NoSuchElementException:
            logger.info(f"Nenhuma opção encontrada para o parceiro {nome}. Tentando o próximo parceiro.")
            continue
        except Exception as e:
            logger.error(f"Erro ao processar o parceiro {nome}: {e}")
        finally:
            time.sleep(3)
            try:
                remove_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/section/md-content/div[2]/div/md-content/div[2]/div/div[4]/div[1]/md-chips/md-chips-wrap/md-chip/div[2]/button")
                remove_button.click()
                logger.info(f"Parceiro {nome} removido da lista de seleção.")
            except NoSuchElementException:
                logger.error(f"Não foi possível encontrar o botão de remover para o parceiro {nome}.")
            time.sleep(3)

# Processar a lista de parceiros
processar_parceiro(lista)

# Finalizar o driver
driver.quit()
time.sleep(5)

# Função para compilar arquivos dentro de uma pasta
def compilar_arquivos_na_pasta(pasta):
    all_files = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith('.xlsx')]
    if not all_files:
        logger.error(f"Nenhum arquivo de relatório encontrado na pasta {pasta}.")
        return
    df_list = [pd.read_excel(file) for file in all_files]
    compiled_df = pd.concat(df_list)
    compiled_file_path = os.path.join(pasta, f"Compilado_{os.path.basename(pasta)}.xlsx")
    compiled_df.to_excel(compiled_file_path, index=False)
    logger.info(f"Todos os arquivos da pasta {pasta} foram compilados em {compiled_file_path}")

# Compilar os arquivos
compilar_arquivos_na_pasta(folder_50)
time.sleep(3)
compilar_arquivos_na_pasta(folder_100)
time.sleep(3)
