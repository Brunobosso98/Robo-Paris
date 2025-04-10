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
import logging
from datetime import datetime
from selenium.common.exceptions import TimeoutException

# Configurações
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DESTINO_DIR = r"C:\Users\Administrador\Desktop\Projetos\Automações\RoboParis\extratos"
EXCEL_PATH = "empresas.xlsx"
URL_LOGIN = "https://portal.ssparisi.com.br/prime/login.php"
URL_EXTRATO = "https://portal.ssparisi.com.br/prime/app/ctrl/GestaoBankExtratoSS.php"

# Mapeamento de bancos e suas classes de botão
BANCO_CLASSES = {
    "sicredi": "button-34",
    "itau": "button-23",
    "santander": "button-32",
    "bradesco": "button-15"
}

# Configuração de logging
def setup_logging():
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_file = os.path.join(log_dir, f"robo_paris_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

    return logging.getLogger("RoboParis")

def inicializar_driver(logger):
    logger.info("Inicializando o driver do Chrome")
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.maximize_window()
        logger.info("Driver inicializado com sucesso")
        return driver
    except Exception as e:
        logger.error(f"Erro ao inicializar o driver: {str(e)}")
        raise

def fazer_login(driver, wait, logger):
    logger.info("Iniciando processo de login")
    try:
        driver.get(URL_LOGIN)
        logger.info(f"Acessando URL de login: {URL_LOGIN}")

        campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "User")))
        campo_usuario.clear()
        campo_usuario.send_keys("bruno.martins@conttrolare.com.br")
        logger.info("Usuário preenchido")

        campo_senha = wait.until(EC.presence_of_element_located((By.ID, "Pass")))
        campo_senha.clear()
        campo_senha.send_keys("1234")
        logger.info("Senha preenchida")

        botao_login = wait.until(EC.element_to_be_clickable((By.ID, "SubLogin")))
        botao_login.click()
        logger.info("Login realizado com sucesso")
    except TimeoutException as e:
        logger.error(f"Timeout ao fazer login: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Erro ao fazer login: {str(e)}")
        raise

def processar_empresa(driver, wait, empresa, data_inicial, data_final, logger):
    logger.info(f"Processando empresa: {empresa} - Período: {data_inicial} a {data_final}")

    try:
        driver.get(URL_EXTRATO)
        logger.info(f"Acessando URL de extrato: {URL_EXTRATO}")

        # Preencher empresa
        campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="autocompleter-empresa-autocomplete"]')))
        campo_empresa.clear()
        campo_empresa.send_keys(empresa)
        campo_empresa.send_keys(Keys.RETURN)
        logger.info(f"Empresa '{empresa}' selecionada")
        time.sleep(3)

        # Selecionar banco
        botao_bancos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bankDiv"]')))
        botao_bancos.click()
        logger.info("Botão de bancos clicado")
        time.sleep(2)

        # Verificar quais contas bancárias estão disponíveis
        contas_processadas = []

        for banco_nome, classe_botao in BANCO_CLASSES.items():
            try:
                # Verificar se existe o botão para este banco
                todos_botoes = driver.find_elements(By.CLASS_NAME, classe_botao)

                # Filtrar apenas os botões de conta bancária (excluir os botões de exclusão)
                botoes_banco = []
                for botao in todos_botoes:
                    id_botao = botao.get_attribute('id')
                    texto_botao = botao.text.strip()

                    # Verificar se o ID NÃO começa com "delete-" e se o texto é "Ver Lançamentos"
                    if (id_botao and not id_botao.startswith('delete-') and
                        texto_botao.lower() == "ver lançamentos"):
                        botoes_banco.append(botao)
                        logger.debug(f"Botão válido encontrado: ID={id_botao}, Texto={texto_botao}")

                if not botoes_banco:
                    logger.info(f"Banco {banco_nome} (classe {classe_botao}) não encontrado para esta empresa")
                    continue

                logger.info(f"Encontrado(s) {len(botoes_banco)} conta(s) do banco {banco_nome}")

                # Processar cada conta do banco encontrado
                for i, botao in enumerate(botoes_banco):
                    try:
                        # Se não for o primeiro banco processado, precisamos clicar no botão de bancos novamente
                        if contas_processadas:
                            driver.get(URL_EXTRATO)
                            logger.info("Retornando à página de extrato para processar próxima conta")

                            # Preencher empresa novamente
                            campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="autocompleter-empresa-autocomplete"]')))
                            campo_empresa.clear()
                            campo_empresa.send_keys(empresa)
                            campo_empresa.send_keys(Keys.RETURN)
                            time.sleep(3)

                            # Clicar no botão de bancos novamente
                            botao_bancos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bankDiv"]')))
                            botao_bancos.click()
                            time.sleep(2)

                            # Atualizar a referência ao botão
                            todos_botoes = driver.find_elements(By.CLASS_NAME, classe_botao)

                            # Filtrar novamente para excluir botões de exclusão
                            botoes_banco = []
                            for b in todos_botoes:
                                id_b = b.get_attribute('id')
                                texto_b = b.text.strip()

                                # Verificar se o ID NÃO começa com "delete-" e se o texto é "Ver Lançamentos"
                                if (id_b and not id_b.startswith('delete-') and
                                    texto_b.lower() == "ver lançamentos"):
                                    botoes_banco.append(b)
                                    logger.debug(f"Botão válido encontrado após recarga: ID={id_b}, Texto={texto_b}")

                            # Verificar se ainda temos botões suficientes
                            if i < len(botoes_banco):
                                botao = botoes_banco[i]
                            else:
                                logger.warning(f"Botão {i+1} não encontrado após recarregar a página")
                                continue

                        id_botao = botao.get_attribute('id')
                        logger.info(f"Processando conta {i+1} do banco {banco_nome} (ID: {id_botao})")
                        botao.click()
                        time.sleep(9)

                        # Preencher datas
                        campo_data_ini = wait.until(EC.presence_of_element_located((By.ID, 'initialDate')))
                        campo_data_ini.clear()
                        campo_data_ini.send_keys(data_inicial)
                        logger.info(f"Data inicial preenchida: {data_inicial}")

                        campo_data_fim = wait.until(EC.presence_of_element_located((By.ID, 'finalDate')))
                        campo_data_fim.clear()
                        campo_data_fim.send_keys(data_final)
                        logger.info(f"Data final preenchida: {data_final}")

                        # Processar extrato
                        botao_processar = wait.until(EC.element_to_be_clickable((By.ID, 'seeTransactions')))
                        botao_processar.click()
                        logger.info("Botão processar clicado, aguardando processamento...")
                        time.sleep(7)

                        # Exportar dados
                        botao_exportar = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable((By.ID, 'export-data'))
                        )
                        botao_exportar.click()
                        logger.info("Botão exportar clicado")

                        # Baixar arquivo
                        try:
                            botao_baixar = WebDriverWait(driver, 30).until(
                                EC.element_to_be_clickable((By.CLASS_NAME, 'btn-success'))
                            )
                            # botao_baixar.click()
                            logger.info(f"Download iniciado para {empresa} - {banco_nome} - conta {i+1}")
                            time.sleep(5)  # Aguardar o download iniciar
                        except TimeoutException:
                            logger.warning(f"Botão de download não encontrado para {empresa} - {banco_nome} - conta {i+1}")
                            # Verificar se há mensagem de 'sem dados'
                            try:
                                msg_sem_dados = driver.find_element(By.XPATH, "//div[contains(text(), 'Nenhum registro encontrado')]")
                                if msg_sem_dados:
                                    logger.info("Nenhum registro encontrado para esta conta no período selecionado")
                            except:
                                pass
                            continue

                        # Mover o arquivo baixado
                        nome_arquivo = mover_arquivo(empresa, data_inicial, banco_nome, i+1, logger)

                        if nome_arquivo:
                            logger.info(f"Arquivo processado e movido: {nome_arquivo}")
                            contas_processadas.append(f"{banco_nome}_{i+1}")
                        else:
                            logger.warning(f"Não foi possível mover o arquivo para {empresa} - {banco_nome} - conta {i+1}")
                    except Exception as e:
                        logger.error(f"Erro ao processar conta {i+1} do banco {banco_nome}: {str(e)}")
                        continue
            except Exception as e:
                logger.error(f"Erro ao verificar banco {banco_nome}: {str(e)}")
                continue

        if not contas_processadas:
            logger.warning(f"Nenhuma conta bancária foi processada para a empresa {empresa}")
            return False

        logger.info(f"Processamento concluído para empresa {empresa}. Contas processadas: {', '.join(contas_processadas)}")
        return True
    except TimeoutException as e:
        logger.error(f"Timeout ao processar empresa {empresa}: {str(e)}")
        return False
    except Exception as e:
        logger.error(f"Erro ao processar empresa {empresa}: {str(e)}")
        return False

def mover_arquivo(empresa, data_inicial, banco, num_conta, logger):
    # Gerar nome único para o arquivo
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"Extrato_{empresa}_{banco}_conta{num_conta}_{data_inicial.replace('/', '-')}_{timestamp}.txt"

    # Criar pasta da empresa se não existir
    pasta_empresa = os.path.join(DESTINO_DIR, empresa)
    if not os.path.exists(pasta_empresa):
        os.makedirs(pasta_empresa)
        logger.info(f"Pasta criada para empresa: {pasta_empresa}")

    # Esperar o download completar
    arquivo_encontrado = None
    for tentativa in range(15):  # Aumentado para 15 tentativas
        logger.info(f"Procurando arquivo baixado (tentativa {tentativa+1}/15)")
        arquivos = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.txt')]
        if arquivos:
            arquivo_encontrado = max(
                [os.path.join(DOWNLOAD_DIR, f) for f in arquivos],
                key=os.path.getmtime
            )
            logger.info(f"Arquivo encontrado: {os.path.basename(arquivo_encontrado)}")
            break
        time.sleep(1)

    if arquivo_encontrado:
        caminho_destino = os.path.join(pasta_empresa, nome_arquivo)
        try:
            shutil.move(arquivo_encontrado, caminho_destino)
            logger.info(f"Arquivo movido para: {caminho_destino}")
            return nome_arquivo
        except Exception as e:
            logger.error(f"Erro ao mover arquivo: {str(e)}")
            return None
    else:
        logger.warning("Arquivo não encontrado para mover")
        return None

def main():
    # Configurar logging
    logger = setup_logging()
    logger.info("=== INICIANDO ROBÔ PARIS ===")

    # Verificar se o diretório de destino existe
    if not os.path.exists(DESTINO_DIR):
        os.makedirs(DESTINO_DIR)
        logger.info(f"Diretório de destino criado: {DESTINO_DIR}")
    else:
        logger.info(f"Diretório de destino encontrado: {DESTINO_DIR}")

    # Mostrar diretório de download
    logger.info(f"Diretório de download configurado: {DOWNLOAD_DIR}")

    driver = None
    try:
        # Inicializar driver
        driver = inicializar_driver(logger)
        wait = WebDriverWait(driver, 30)  # Aumentado para 30 segundos

        # Fazer login
        fazer_login(driver, wait, logger)

        # Verificar se o arquivo Excel existe
        if not os.path.exists(EXCEL_PATH):
            logger.error(f"Arquivo de empresas não encontrado: {EXCEL_PATH}")
            return

        # Ler planilha
        try:
            logger.info(f"Lendo planilha de empresas: {EXCEL_PATH}")
            df = pd.read_excel(EXCEL_PATH, parse_dates=['dataInicial', 'dataFinal'])
            logger.info(f"Total de empresas na planilha: {len(df)}")
        except Exception as e:
            logger.error(f"Erro ao ler planilha: {str(e)}")
            return

        # Processar cada empresa
        empresas_processadas = 0
        empresas_com_erro = 0

        for index, row in df.iterrows():
            empresa = row['Empresa']
            data_inicial = row['dataInicial'].strftime('%d/%m/%Y')
            data_final = row['dataFinal'].strftime('%d/%m/%Y')

            logger.info(f"\n{'='*50}")
            logger.info(f"EMPRESA {index+1}/{len(df)}: {empresa}")
            logger.info(f"{'='*50}")

            # Tentar processar a empresa até 3 vezes em caso de falha
            max_tentativas = 3
            tentativa = 1
            sucesso = False

            while tentativa <= max_tentativas and not sucesso:
                try:
                    logger.info(f"Tentativa {tentativa}/{max_tentativas} para empresa {empresa}")
                    resultado = processar_empresa(driver, wait, empresa, data_inicial, data_final, logger)

                    if resultado:
                        logger.info(f"Empresa {empresa} processada com sucesso na tentativa {tentativa}")
                        empresas_processadas += 1
                        sucesso = True
                    else:
                        logger.warning(f"Falha ao processar empresa {empresa} na tentativa {tentativa}")
                        tentativa += 1
                        if tentativa <= max_tentativas:
                            logger.info(f"Aguardando 5 segundos antes da próxima tentativa...")
                            time.sleep(5)
                except Exception as e:
                    logger.error(f"Erro não tratado ao processar {empresa} na tentativa {tentativa}: {str(e)}")
                    tentativa += 1
                    if tentativa <= max_tentativas:
                        logger.info(f"Aguardando 5 segundos antes da próxima tentativa...")
                        time.sleep(5)

            if not sucesso:
                logger.error(f"Todas as {max_tentativas} tentativas falharam para empresa {empresa}")
                empresas_com_erro += 1

        # Resumo final
        logger.info("\n\n=== RESUMO DA EXECUÇÃO ===")
        logger.info(f"Total de empresas: {len(df)}")
        logger.info(f"Empresas processadas com sucesso: {empresas_processadas}")
        logger.info(f"Empresas com erro: {empresas_com_erro}")
        logger.info("=== FIM DA EXECUÇÃO ===\n")

    except Exception as e:
        if logger:
            logger.error(f"Erro crítico na execução: {str(e)}")
        else:
            print(f"Erro crítico na execução: {str(e)}")
    finally:
        if driver:
            logger.info("Fechando o navegador")
            driver.quit()

if __name__ == "__main__":
    main()
