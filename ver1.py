from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import shutil
import time
from datetime import datetime

# Configurações
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DESTINO_DIR = r"C:\Users\bruno.martins\Desktop\RoboParis\extratos"
EXCEL_PATH = "empresas.xlsx"
URL_LOGIN = "https://portal.ssparisi.com.br/prime/login.php"
URL_EXTRATO = "https://portal.ssparisi.com.br/prime/app/ctrl/GestaoBankExtratoSS.php"

def inicializar_driver():
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    return driver

def fazer_login(driver, wait):
    driver.get(URL_LOGIN)
    
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "User")))
    campo_usuario.clear()
    campo_usuario.send_keys("mauro@conttrolare.com.br")

    campo_senha = wait.until(EC.presence_of_element_located((By.ID, "Pass")))
    campo_senha.clear()
    campo_senha.send_keys("Juni4724")

    botao_login = wait.until(EC.element_to_be_clickable((By.ID, "SubLogin")))
    botao_login.click()

def processar_empresa(driver, wait, empresa, data_inicial, data_final):
    driver.get(URL_EXTRATO)
    
    # Preencher empresa
    campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="autocompleter-empresa-autocomplete"]')))
    campo_empresa.clear()
    campo_empresa.send_keys(empresa)
    campo_empresa.send_keys(Keys.RETURN)
    time.sleep(3)

    # Selecionar banco
    botao_bancos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bankDiv"]')))
    botao_bancos.click()

    # Selecionar conta bancária
    botao_lancamentos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="account-7097"]')))
    botao_lancamentos.click()
    time.sleep(8)

    # Preencher datas
    campo_data_ini = wait.until(EC.presence_of_element_located((By.ID, 'initialDate')))
    campo_data_ini.clear()
    campo_data_ini.send_keys(data_inicial)

    campo_data_fim = wait.until(EC.presence_of_element_located((By.ID, 'finalDate')))
    campo_data_fim.clear()
    campo_data_fim.send_keys(data_final)

    check_exportados = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="modalAccount"]/div/div/div[2]/div[3]/div[2]/label')))
    check_exportados.click()

    # Processar extrato
    botao_processar = wait.until(EC.element_to_be_clickable((By.ID, 'seeTransactions')))
    botao_processar.click()
    time.sleep(7)

    # Obter o valor do elemento HTML
    valor_elemento = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="accordion_parent"]/tr[1]/td[1]'))).text

    # Definir o ID do primeiro_input com base no valor do elemento
    primeiro_input_id = f"field-{valor_elemento}-0"
    primeiro_input = wait.until(EC.presence_of_element_located((By.ID, primeiro_input_id)))
    primeiro_input.clear()
    primeiro_input.send_keys("20")
    print("primeiro_input")

    segundo_input_id = f"field-{valor_elemento}-1"
    segundo_input = wait.until(EC.presence_of_element_located((By.ID, segundo_input_id)))
    segundo_input.clear()
    segundo_input.send_keys("10")
    print("segundo")

    terceiro_input_id = f"field-{valor_elemento}-2"
    terceiro_input = wait.until(EC.presence_of_element_located((By.ID, terceiro_input_id)))
    terceiro_input.clear()
    terceiro_input.send_keys("8")
    print("terceiro_input")

    # Exportar dados
    botao_exportar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, 'export-data'))
    )
    botao_exportar.click()

    # Baixar arquivo
    botao_baixar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'btn-success'))
    )
    botao_baixar.click()

def mover_arquivo(empresa, data_inicial):
    # Gerar nome único para o arquivo
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"Extrato_{empresa}_{data_inicial.replace('/', '-')}_{timestamp}.txt"
    
    # Esperar o download completar
    arquivo_encontrado = None
    for _ in range(10):
        arquivos = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.txt')]
        if arquivos:
            arquivo_encontrado = max(
                [os.path.join(DOWNLOAD_DIR, f) for f in arquivos],
                key=os.path.getmtime
            )
            break
        time.sleep(1)

    if arquivo_encontrado:
        caminho_destino = os.path.join(DESTINO_DIR, nome_arquivo)
        shutil.move(arquivo_encontrado, caminho_destino)
        print(f"Arquivo movido: {caminho_destino}")
    else:
        print("Arquivo não encontrado para mover")

def main():
    driver = inicializar_driver()
    wait = WebDriverWait(driver, 20)
    
    try:
        # Fazer login
        fazer_login(driver, wait)
        
        # Ler planilha
        df = pd.read_excel(EXCEL_PATH, parse_dates=['dataInicial', 'dataFinal'])
        
        # Processar cada empresa
        for index, row in df.iterrows():
            empresa = row['Empresa']
            data_inicial = row['dataInicial'].strftime('%d/%m/%Y')
            data_final = row['dataFinal'].strftime('%d/%m/%Y')
            
            print(f"Processando: {empresa} - {data_inicial} a {data_final}")
            
            try:
                processar_empresa(driver, wait, empresa, data_inicial, data_final)
                mover_arquivo(empresa, data_inicial)
            except Exception as e:
                print(f"Erro ao processar {empresa}: {str(e)}")
                continue
                
    finally:
        driver.quit()

if __name__ == "__main__":
    main()


# a variavel valor_elemento = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="accordion_parent"]/tr[1]/td[1]'))).text
# que contem o valor do primeiro tr (//*[@id="accordion_parent"]/tr[1]), muda conforme a linha, ou seja,
# a primeira linha é tr[1], a segunda tr[2] e assim por diante.

# Este exemplo faz apenas na primeira linha, preciso que faça nas linhas 
# onde o histórico bata com o histórico do arquivo excel.

# O programa deve verificar se os valores do Histórico do excel tem o mesmo valor de algum Histórico do site
# O XPATH do elemento vem desse jeito: //*[@id="accordion_parent"]/tr[1]/td[1]
    # Significa que é o XPATH do elemento da primeira linha
    # O XPATH da quarta linha é: //*[@id="accordion_parent"]/tr[4]/td[1]

# Após encontrar, deve pegar o valor do HTML da linha atual, exemplo do elemento (<td class="w-3p">4</td>) pelo XPATH da onde o loop está (verificando o histórico pelo excel)
    # E deve fazer com que o valor do input seja: field-(valor do elemento)-0'
    # O segundo input deverá ficar assim:  field-(valor do elemento)-1'