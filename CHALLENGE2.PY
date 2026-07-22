import pandas as pd
from selenium import webdriver
#from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
#from webdriver_manager.chrome import ChromeDriverManager
import logging
from datetime import datetime
from pathlib import Path
import time


### Caminho do arquivo e URL do teste
base = Path("challenge.xlsx")
URL = "https://rpachallenge.com/"


### Lendo arquivo XLSX
df = pd.read_excel(base)


### Remover espaços nas colunas do arquivo XLSX
df.columns = df.columns.str.strip()
df = df.fillna("")


### Configurações necessárias para o arquivo de Logs
LOGS = Path("logs")
LOGS.mkdir(exist_ok=True)

data_execucao = datetime.now().strftime("%Y%m%d_%H%M%S")

logging.basicConfig(
    filename=LOGS / f"execucao_{data_execucao}.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8"
)

logging.info("Início da automação")
logging.info("Arquivo Excel carregado | Registros encontrados: %s", len(df))


### Executando Chrome
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

nav = webdriver.Chrome(options=options)

wait = WebDriverWait(nav, 10)

logging.info("Etapa iniciada: abrindo o site | URL: %s", URL)

nav.get(URL)

logging.info("Etapa concluída: site aberto")


### Clique no Start
wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[normalize-space()='Start']")
    )
).click()

logging.info("Etapa concluída: botão Start acionado | Status: OK")


### Loop para input dos dados
for numero_linha, (_, linha) in enumerate(df.iterrows(), start=1):

    first_name = str(linha["First Name"])
    last_name = str(linha["Last Name"])
    company_name = str(linha["Company Name"])
    role_company = str(linha["Role in Company"])
    address = str(linha["Address"])
    email = str(linha["Email"])
    phone = str(linha["Phone Number"])

    nome_registro = f"{first_name} {last_name}"
    timestamp_processamento = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        ### LOG: início do processamento da linha
        logging.info(
            "Etapa iniciada | Linha: %s | Registro: %s",
            numero_linha,
            nome_registro
        )

        wait.until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, "input[ng-reflect-name='labelFirstName']")
            )
        ).send_keys(first_name)

        nav.find_element(
            By.CSS_SELECTOR,
            "input[ng-reflect-name='labelLastName']"
        ).send_keys(last_name)

        nav.find_element(
            By.CSS_SELECTOR,
            "input[ng-reflect-name='labelCompanyName']"
        ).send_keys(company_name)

        nav.find_element(
            By.CSS_SELECTOR,
            "input[ng-reflect-name='labelRole']"
        ).send_keys(role_company)

        nav.find_element(
            By.CSS_SELECTOR,
            "input[ng-reflect-name='labelAddress']"
        ).send_keys(address)

        nav.find_element(
            By.CSS_SELECTOR,
            "input[ng-reflect-name='labelEmail']"
        ).send_keys(email)

        nav.find_element(
            By.CSS_SELECTOR,
            "input[ng-reflect-name='labelPhone']"
        ).send_keys(phone)

        logging.info(
            "Etapa concluída | Linha: %s | Formulário preenchido",
            numero_linha
        )

        nav.find_element(
            By.CSS_SELECTOR,
            "input[value='Submit']"
        ).click()

        ### LOG: resultado OK da linha
        logging.info(
            "RESULTADO | Status: OK | Linha: %s | Registro: %s | "
            "Observação: formulário preenchido e enviado com sucesso | "
            "Data/hora processamento: %s",
            numero_linha,
            nome_registro,
            timestamp_processamento
        )

    ### LOG: resultado NOK da linha, incluindo traceback
    except Exception as erro:
        
        logging.exception(
            "RESULTADO | Status: NOK | Linha: %s | Registro: %s | "
            "Observação: erro no preenchimento ou envio: %s | "
            "Data/hora processamento: %s",
            numero_linha,
            nome_registro,
            str(erro),
            timestamp_processamento
        )

        ### Screenshot no ponto exato da falha
        caminho_erro = LOGS / (
            f"erro_linha_{numero_linha}_{data_execucao}.png"
        )
        nav.save_screenshot(str(caminho_erro))

        logging.info(
            "Evidência de erro salva | Linha: %s | Arquivo: %s",
            numero_linha,
            caminho_erro
        )

        # Segue para a próxima linha do Excel.
        # Remova este continue se quiser que o script pare no primeiro erro.
        continue


time.sleep(0.3)

### Captura da tela final com score
try:
    logging.info("Etapa iniciada: captura da tela final do score")


    caminho_score = LOGS / f"score_final_{data_execucao}.png"
    nav.save_screenshot(str(caminho_score))

    logging.info(
        "Etapa concluída | Status: OK | "
        "Observação: screenshot do score final salvo | Arquivo: %s",
        caminho_score
    )

    print(f"Desafio finalizado. Screenshot do score salvo em:\n{caminho_score}")

except Exception as erro:
    logging.exception(
        "RESULTADO FINAL | Status: NOK | "
        "Observação: não foi possível capturar a tela do score: %s",
        str(erro)
    )

    print("Desafio finalizado, mas não foi possível capturar o score final.")


time.sleep(3)

### Fecha o navegador após 3 segundos
nav.quit()

logging.info("Automação encerrada")
