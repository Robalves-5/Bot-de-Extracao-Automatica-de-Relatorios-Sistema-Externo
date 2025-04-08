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
dotenv_path_env = r'[CAMINHO_DO_ARQUIVO_ENV]'
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
name_path = r"[CAMINHO_DA_PLANILHA]"
df = pd.read_excel(name_path)
time.sleep(1)

# Verificar se as colunas necessárias existem
if "Parceiros" not in df.columns or "Cashback" not in df.columns:
    logger.error("Erro: A planilha não contém as colunas 'Parceiros' e 'Cashback'.")
    exit(10)
time.sleep(1)

lista = df["Parceiros"].tolist()
time.sleep(1)

# Diretório de download
download_directory = r'[CAMINHO_DIRETORIO_DOWNLOAD]'
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
driver.get("[URL_PRIVADA_REMOVIDA]")
time.sleep(3)

# Fazer login
email = driver.find_element(By.ID, "userLogin")
password = driver.find_element(By.ID, "input_0")
email.send_keys(EMAIL)
password.send_keys(PASSWORD)
login_button = driver.find_element(By.XPATH, "[XPATH_LOGIN_BUTTON]")
login_button.click()
time.sleep(3)

# Navegação até a aba de Cashback
driver.find_element(By.XPATH, "[XPATH_COMISSAO_BUTTON]").click()
time.sleep(3)

driver.find_element(By.XPATH, "[XPATH_ABA_CASHBACK]").click()
time.sleep(3)

# Selecionar mês e tipo de campo
driver.find_element(By.XPATH, "[XPATH_MES]").click()
time.sleep(3)

driver.find_element(By.XPATH, "//md-option/div[contains(text(), 'Janeiro')]").click()
time.sleep(3)

driver.find_element(By.XPATH, "[XPATH_CAMPO]").click()
time.sleep(3)

driver.find_element(By.XPATH, "//div[@class='md-text' and text()='Cashback']").click()
time.sleep(3)

# Campo do parceiro
parceiro = driver.find_element(By.XPATH, "[XPATH_CAMPO_PARCEIRO]")
parceiro.click()
time.sleep(3)

# Função principal
def processar_parceiro(lista):
    def mover_para_pasta_correcta(file_path, cashback, nome_arquivo):
        destino = folder_50 if cashback == "50%" else folder_100 if cashback == "100%" else None
        if destino:
            novo_caminho = os.path.join(destino, nome_arquivo)
            os.rename(file_path, novo_caminho)
            logger.info(f"Arquivo movido para: {novo_caminho}")
        else:
            logger.error(f"Cashback '{cashback}' não reconhecido para: {nome_arquivo}")

    for nome in lista:
        try:
            parceiro.clear()
            parceiro.send_keys(nome)
            time.sleep(3)

            driver.find_element(By.XPATH, "[XPATH_SUBMIT]").click()
            time.sleep(3)

            driver.find_element(By.XPATH, "[XPATH_EXTRACAO]").click()
            time.sleep(10)

            files = os.listdir(download_directory)
            paths = [os.path.join(download_directory, basename) for basename in files]
            temp_download_path = max(paths, key=os.path.getctime)
            time.sleep(5)

            new_file_name = f"Relatório-cashback-{nome}.xlsx"
            new_file_path = os.path.join(download_directory, new_file_name)
            os.rename(temp_download_path, new_file_path)
            logger.info(f"Arquivo salvo como: {new_file_name}")
            time.sleep(3)

            parceiro_cashback = df[df["Parceiros"].str.strip() == nome.strip()]
            if parceiro_cashback.empty:
                logger.error(f"Parceiro {nome} não encontrado na planilha.")
                continue

            cashback = str(parceiro_cashback["Cashback"].iloc[0]).strip()
            if cashback == "0.5": cashback = "50%"
            elif cashback == "1.0": cashback = "100%"

            mover_para_pasta_correcta(new_file_path, cashback, new_file_name)
            time.sleep(3)
        except NoSuchElementException:
            logger.info(f"Nenhuma opção encontrada para o parceiro {nome}.")
            continue
        except Exception as e:
            logger.error(f"Erro com parceiro {nome}: {e}")
        finally:
            time.sleep(3)
            try:
                driver.find_element(By.XPATH, "[XPATH_REMOVER_PARCEIRO]").click()
                logger.info(f"Parceiro {nome} removido da seleção.")
            except NoSuchElementException:
                logger.warning(f"Não foi possível remover {nome}.")
            time.sleep(3)

# Processar parceiros
processar_parceiro(lista)

# Finalizar navegador
driver.quit()
time.sleep(5)

# Compilar arquivos
def compilar_arquivos_na_pasta(pasta):
    arquivos = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith('.xlsx')]
    if not arquivos:
        logger.error(f"Nenhum arquivo encontrado na pasta {pasta}.")
        return
    df_final = pd.concat([pd.read_excel(arq) for arq in arquivos])
    caminho_saida = os.path.join(pasta, f"Compilado_{os.path.basename(pasta)}.xlsx")
    df_final.to_excel(caminho_saida, index=False)
    logger.info(f"Arquivo compilado salvo em {caminho_saida}")

compilar_arquivos_na_pasta(folder_50)
time.sleep(3)
compilar_arquivos_na_pasta(folder_100)
time.sleep(3)
