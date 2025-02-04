import logging
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
import json

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('execucao.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configurações
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DESTINO_DIR = r"C:\Users\bruno.martins\Desktop\RoboParis\extratos"
EXCEL_PATH = "empresas.xlsx"
EXCEL_MAP = "mapeamento.xlsx"
URL_LOGIN = "https://portal.ssparisi.com.br/prime/login.php"
URL_EXTRATO = "https://portal.ssparisi.com.br/prime/app/ctrl/GestaoBankExtratoSS.php"

def inicializar_driver():
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    return driver

def carregar_credenciais(caminho_arquivo):
    """Carrega as credenciais de login de um arquivo JSON."""
    with open(caminho_arquivo, 'r') as arquivo:
        return json.load(arquivo)

def fazer_login(driver, wait):
    driver.get(URL_LOGIN)
    
    # Carregar credenciais do arquivo
    credenciais = carregar_credenciais('cred.json')

    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "User")))
    campo_usuario.clear()
    campo_usuario.send_keys(credenciais['username'])

    campo_senha = wait.until(EC.presence_of_element_located((By.ID, "Pass")))
    campo_senha.clear()
    campo_senha.send_keys(credenciais['password'])

    botao_login = wait.until(EC.element_to_be_clickable((By.ID, "SubLogin")))
    botao_login.click()

def processar_empresa(driver, wait, empresa, data_inicial, data_final):
    driver.get(URL_EXTRATO)
    
    # Preencher empresa
    campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="autocompleter-empresa-autocomplete"]')))
    campo_empresa.clear()
    campo_empresa.send_keys(empresa)
    campo_empresa.send_keys(Keys.RETURN)
    time.sleep(4)

    # Selecionar banco
    botao_bancos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bankDiv"]')))
    botao_bancos.click()

    # Selecionar conta bancária
    botao_lancamentos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="account-7097"]')))
    botao_lancamentos.click()
    time.sleep(9)

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
    time.sleep(5)

    processar_historicos(driver, wait)

    # Forçar rolagem para o topo absoluto
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)  # Pequeno delay para garantir que a página carregue

    # Garantir que um elemento fixo no topo esteja visível
    elemento_topo = driver.find_element(By.XPATH, '/html/body/div[4]')
    driver.execute_script("arguments[0].scrollIntoView();", elemento_topo)
    logger.info("Página rolada para o topo.")

    # Pequena pausa para evitar erros de carregamento
    time.sleep(0.5)

    # Localizar e clicar no botão "Exportar"
    botao_exportar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, 'export-data'))
    )
    botao_exportar.click()
    logger.info("Botão 'Exportar' clicado com sucesso!")

    # Baixar arquivo
    botao_baixar = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'btn-success'))
    )
    botao_baixar.click()

def carregar_mapeamento():
    """Carrega o mapeamento de históricos do Excel com logging"""
    try:
        logger.info("Carregando mapeamento do Excel...")
        df = pd.read_excel(EXCEL_MAP)
        df = df.fillna('')  # Tratar valores nulos
        mapeamento = df.set_index('Historico').to_dict(orient='index')
        logger.info(f"Total de regras carregadas: {len(mapeamento)}")
        return mapeamento
    except Exception as e:
        logger.error(f"Erro ao carregar mapeamento: {str(e)}")
        raise

def processar_historicos(driver, wait):
    """Processa cada linha da tabela de históricos comparando com o mapeamento do Excel"""
    mapeamento = carregar_mapeamento()  # Carregar os históricos do Excel

    try:
        linhas = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="accordion_parent"]/tr')))
        logger.info(f"Encontradas {len(linhas)} linhas para processar")
    except Exception as e:
        logger.error(f"Erro ao localizar linhas: {str(e)}")
        return

    for i, linha in enumerate(linhas, start=1):
        try:
            # 1. Obter o número do identificador 
            xpath_valor_elemento = f'//*[@id="accordion_parent"]/tr[{i}]/td[1]'
            valor_elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath_valor_elemento))).text.strip()

            if not valor_elemento.isdigit():
                logger.error(f"Valor inválido no campo identificador na linha {i}: '{valor_elemento}'")
                continue

            # 2. Construir o ID do input do histórico e obter o texto real do histórico
            input_historico_id = f"field-{valor_elemento}-hist"
            input_historico = wait.until(EC.presence_of_element_located((By.ID, input_historico_id)))
            historico = input_historico.get_attribute("value").strip()

            if not historico:
                logger.warning(f"Histórico vazio na linha {i}")
                continue

            # 3. Verificar se o histórico real está no mapeamento do Excel
            if historico not in mapeamento:
                logger.warning(f"Histórico não mapeado: '{historico}' na linha {i}")
                continue

            dados = mapeamento[historico]
            logger.info(f"Processando linha {i}: Histórico '{historico}' encontrado no mapeamento.")

            # 4. Preencher os campos correspondentes
            preencher_campos(driver, wait, valor_elemento, dados)

        except Exception as e:
            logger.error(f"Erro na linha {i}: {str(e)}", exc_info=True)
            continue

def preencher_campos(driver, wait, valor_elemento, dados):
    """Preenche os campos baseados no valor do elemento e nos dados do mapeamento"""
    logger.info(f"Preenchendo campos para ID {valor_elemento}...")

    try:
        # IDs dinâmicos para os inputs
        primeiro_input_id = f"field-{valor_elemento}-0"
        segundo_input_id = f"field-{valor_elemento}-1"
        terceiro_input_id = f"field-{valor_elemento}-2"

        # Preencher o primeiro input
        primeiro_input = wait.until(EC.presence_of_element_located((By.ID, primeiro_input_id)))
        primeiro_input.clear()
        primeiro_input.send_keys(str(dados['HP']))
        logger.info(f"Campo {primeiro_input_id} preenchido com: {dados['HP']}")

        # Preencher o segundo input
        segundo_input = wait.until(EC.presence_of_element_located((By.ID, segundo_input_id)))
        segundo_input.clear()
        segundo_input.send_keys(str(dados['Déb']))
        logger.info(f"Campo {segundo_input_id} preenchido com: {dados['Déb']}")

        # Preencher o terceiro input
        terceiro_input = wait.until(EC.presence_of_element_located((By.ID, terceiro_input_id)))
        terceiro_input.clear()
        terceiro_input.send_keys(str(dados['Cré']))
        logger.info(f"Campo {terceiro_input_id} preenchido com: {dados['Cré']}")

    except Exception as e:
        logger.error(f"Erro ao preencher campos para ID {valor_elemento}: {str(e)}", exc_info=True)

def clicar_botao_acao(linha, wait, tipo_botao):
    """Clica no botão de ação correspondente com verificação"""
    logger.debug(f"Tentando clicar no botão {tipo_botao}")
    
    mapeamento_botoes = {
        'verde': ('btn-success', 'Botão verde'),
        'amarelo': ('btn-warning', 'Botão amarelo'),
        'vermelho': ('btn-danger', 'Botão vermelho')
    }

    try:
        classe, nome = mapeamento_botoes[tipo_botao]
        botao = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, f'button.{classe}')
        ), message=f"{nome} não encontrado")
        
        botao.click()
        logger.info(f"Clicado no {nome} com sucesso")
        
        # Verificar mudança de estado se aplicável
        time.sleep(0.5)  # Ajustar conforme necessidade
        
    except Exception as e:
        logger.error(f"Falha ao clicar no botão {tipo_botao}: {str(e)}")
        raise

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
    
    if not os.path.exists(DESTINO_DIR):
        os.makedirs(DESTINO_DIR)
        logger.info(f"Pasta criada: {DESTINO_DIR}")
    
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
