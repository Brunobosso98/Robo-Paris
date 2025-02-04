from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os

EXCEL_PATH = "empresas.xlsx"
URL_LOGIN = "https://portal.ssparisi.com.br/prime/login.php"
URL_RELACIONAMENTO = "https://portal.ssparisi.com.br/prime/app/ls/Relacionamento.php"

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

def relacionamento_empresa(driver, wait, empresa):
    driver.get(URL_RELACIONAMENTO)

    # Preencher empresa
    campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="empresa-autocomplete-main"]')))
    campo_empresa.clear()
    campo_empresa.send_keys(empresa)
    campo_empresa.send_keys(Keys.RETURN)
    time.sleep(3)

    botao_visualizar = wait.until(EC.element_to_be_clickable((By.ID, 'processa')))
    botao_visualizar.click()
    time.sleep(5)

    # Processar tabela de relacionamento
    processar_tabela(driver, wait, empresa)

def processar_tabela(driver, wait, empresa):
    dados = []

    try:
        tabela = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="rel_dataTable"]/tbody')))
        linhas = tabela.find_elements(By.TAG_NAME, "tr")

        for index, linha in enumerate(linhas, start=1):
            try:
                # XPath do td que contém o <span>
                xpath_sistema = f'//*[@id="rel_dataTable"]/tbody/tr[{index}]/td[2]/span'
                sistema_importacao = linha.find_element(By.XPATH, xpath_sistema).get_attribute("textContent").strip()


                # XPath do td que contém a string Hist. Padrão / Débito / Crédito
                xpath_hist = f'//*[@id="rel_dataTable"]/tbody/tr[{index}]/td[4]'
                hist_padrao = linha.find_element(By.XPATH, xpath_hist).text

                # Separar valores da string Hist. Padrão / Débito / Crédito
                hist_valores = hist_padrao.split("/")  # Exemplo: "3/100/2" → ["3", "100", "2"]
                
                if len(hist_valores) == 3:
                    hist_padrao, debito, credito = hist_valores
                else:
                    hist_padrao, debito, credito = hist_padrao, "", ""

                print(f"Linha {index}: {sistema_importacao} | {hist_padrao} | {debito} | {credito}")

                # Adicionar ao array de dados
                dados.append([sistema_importacao, hist_padrao, debito, credito, empresa])

            except Exception as e:
                print(f"Erro ao processar linha {index}: {e}")
                continue

        # Salvar no Excel
        salvar_dados(dados)

    except Exception as e:
        print(f"Erro ao processar tabela: {e}")

def salvar_dados(dados):
    df = pd.DataFrame(dados, columns=["Historico", "HP", "Déb", "Cré", "Empresa"])
    
    output_file = "mapeamento.xlsx"
    
    if os.path.exists(output_file):
        df_existente = pd.read_excel(output_file)
        df = pd.concat([df_existente, df], ignore_index=True)
    
    df.to_excel(output_file, index=False)
    print(f"Dados salvos em {output_file}")

def main():
    driver = inicializar_driver()
    wait = WebDriverWait(driver, 20)
    
    try:
        # Fazer login
        fazer_login(driver, wait)
        
        # Ler planilha
        df = pd.read_excel(EXCEL_PATH)
        
        # Processar cada empresa
        for index, row in df.iterrows():
            empresa = row['Empresa']
            
            print(f"Processando: {empresa}")
            
            try:
                relacionamento_empresa(driver, wait, empresa)
            except Exception as e:
                print(f"Erro ao processar {empresa}: {str(e)}")
                continue
                
    finally:
        driver.quit()

if __name__ == "__main__":
    main()

