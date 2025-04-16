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
import re
from datetime import datetime, timedelta
from selenium.common.exceptions import TimeoutException
from fpdf import FPDF
from collections import defaultdict

# Configurações
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
BASE_DESTINO_DIR = r"I:\Contabilidade\Banco Online"
EXCEL_PATH = "empresas.xlsx"
URL_LOGIN = "https://portal.ssparisi.com.br/prime/login.php"
URL_EXTRATO = "https://portal.ssparisi.com.br/prime/app/ctrl/GestaoBankExtratoSS.php"

# Dicionário de meses em português
MESES = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro"
}

# Mapeamento de bancos e suas classes de botão
BANCO_CLASSES = {
    "sicredi": "button-34",
    "itau": "button-23",
    "santander": "button-32",
    "bradesco": "button-15"
}

# Função para calcular as datas do mês anterior
def calcular_datas_mes_anterior():
    hoje = datetime.now()

    # Calcular o primeiro dia do mês atual
    primeiro_dia_mes_atual = hoje.replace(day=1)

    # Calcular o último dia do mês anterior (um dia antes do primeiro dia do mês atual)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)

    # Calcular o primeiro dia do mês anterior
    primeiro_dia_mes_anterior = ultimo_dia_mes_anterior.replace(day=1)

    # Formatar as datas no formato dd/mm/aaaa
    data_inicial = primeiro_dia_mes_anterior.strftime('%d/%m/%Y')
    data_final = ultimo_dia_mes_anterior.strftime('%d/%m/%Y')

    # Obter o ano e o mês para o caminho do diretório
    ano = ultimo_dia_mes_anterior.year
    mes = ultimo_dia_mes_anterior.month
    nome_mes = MESES[mes]

    return data_inicial, data_final, ano, nome_mes

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

def processar_empresa(driver, wait, empresa, data_inicial, data_final, ano, nome_mes, logger, erros_registrados=None, resumo_bancos=None):
    logger.info(f"Processando empresa: {empresa} - Período: {data_inicial} a {data_final}")

    # Inicializar o dicionário de erros registrados se não for fornecido
    if erros_registrados is None:
        erros_registrados = {}
    if resumo_bancos is None:
        resumo_bancos = []

    try:
        # Primeiro, identificar todos os bancos disponíveis para esta empresa
        bancos_disponiveis = identificar_bancos_disponiveis(driver, wait, empresa, logger)

        # Verificar se encontramos algum banco
        if not bancos_disponiveis:
            erro_msg = f"Nenhum banco encontrado para a empresa {empresa}"
            logger.warning(erro_msg)
            adicionar_resumo_banco_unico(resumo_bancos, empresa, "-", "Erro", erro_msg)
            return False

        logger.info(f"Total de bancos encontrados: {len(bancos_disponiveis)}")

        # Processar cada banco e suas contas
        contas_processadas = []

        for banco_nome, botoes_info in bancos_disponiveis.items():
            logger.info(f"Processando banco {banco_nome} com {len(botoes_info)} conta(s)...")

            for i, botao_info in enumerate(botoes_info):
                status_banco = "Sucesso"
                mensagem_banco = "Processado com sucesso"
                try:
                    # Para cada conta, voltar à página inicial e selecionar a empresa novamente
                    driver.get(URL_EXTRATO)
                    logger.info(f"Acessando página de extrato para processar conta {i+1} do banco {banco_nome}")

                    # Preencher empresa
                    campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="autocompleter-empresa-autocomplete"]')))
                    campo_empresa.clear()
                    campo_empresa.send_keys(empresa)
                    campo_empresa.send_keys(Keys.RETURN)
                    time.sleep(3)

                    # Clicar no botão de bancos
                    botao_bancos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bankDiv"]')))
                    botao_bancos.click()
                    time.sleep(3)

                    # Encontrar o botão correto usando as informações armazenadas
                    classe_botao = botao_info['classe']
                    id_botao = botao_info['id']

                    # Encontrar todos os botões com a classe correta
                    todos_botoes = driver.find_elements(By.CLASS_NAME, classe_botao)
                    botao_encontrado = None

                    # Procurar o botão com o ID correto
                    for b in todos_botoes:
                        if b.get_attribute('id') == id_botao:
                            botao_encontrado = b
                            break

                    if not botao_encontrado:
                        mensagem_banco = f"Botão com ID {id_botao} não encontrado para o banco {banco_nome}"
                        logger.warning(mensagem_banco)
                        status_banco = "Erro"
                        adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
                        continue

                    logger.info(f"Processando conta {i+1} do banco {banco_nome} (ID: {id_botao})")
                    botao_encontrado.click()
                    time.sleep(12)

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
                    time.sleep(3)  # Reduzido para verificar o modal mais rapidamente

                    # Verificar se aparece o modal "Sem lançamentos!"
                    try:
                        # Tentar encontrar o modal com timeout reduzido para não atrasar muito o processo
                        modal_sem_lancamentos = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Sem lançamentos')]"))
                        )

                        if modal_sem_lancamentos:
                            mensagem_banco = "Sem lançamentos disponíveis para esta conta no período selecionado"
                            logger.warning(mensagem_banco)
                            chave_erro = f"{empresa}_{banco_nome}_sem_lancamentos"
                            if chave_erro not in erros_registrados:
                                registrar_erro_no_arquivo(empresa, banco_nome, mensagem_banco, ano, nome_mes, logger)
                                erros_registrados[chave_erro] = True
                            status_banco = "Erro"
                            adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
                            continue
                    except TimeoutException:
                        # Não encontrou o modal, o que é bom - significa que há lançamentos
                        logger.info("Processando lançamentos...")

                    # Aguardar o processamento completo
                    time.sleep(4)

                    # Verificar se todos os lançamentos foram feitos
                    try:
                        info_element = wait.until(EC.presence_of_element_located((By.ID, 'empInfo')))
                        info_text = info_element.text
                        logger.info(f"Informação de lançamentos: {info_text}")

                        # Extrair a informação de lançamentos (ex: "0/136")
                        lancamentos_match = re.search(r'(\d+)/(\d+)$', info_text)

                        if lancamentos_match:
                            lancamentos_feitos = int(lancamentos_match.group(1))
                            lancamentos_total = int(lancamentos_match.group(2))

                            logger.info(f"Lançamentos: {lancamentos_feitos} de {lancamentos_total}")

                            # Verificar se todos os lançamentos foram feitos
                            if lancamentos_feitos < lancamentos_total:
                                mensagem_banco = f"Nem todos os lançamentos foram feitos: {lancamentos_feitos}/{lancamentos_total}"
                                logger.warning(mensagem_banco)
                                chave_erro = f"{empresa}_{banco_nome}_lancamentos"
                                if chave_erro not in erros_registrados:
                                    registrar_erro_no_arquivo(empresa, banco_nome, mensagem_banco, ano, nome_mes, logger)
                                    erros_registrados[chave_erro] = True
                                status_banco = "Erro"
                                adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
                                continue
                            else:
                                logger.info(f"Todos os lançamentos foram feitos corretamente: {lancamentos_feitos}/{lancamentos_total}")
                    except Exception as e:
                        logger.warning(f"Não foi possível verificar os lançamentos: {str(e)}")

                    # Exportar dados
                    botao_exportar = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.ID, 'export-data'))
                    )
                    botao_exportar.click()
                    logger.info("Botão exportar clicado")

                    # Variável para controlar se o download foi iniciado
                    download_iniciado = False

                    # Baixar arquivo
                    try:
                        botao_baixar = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable((By.CLASS_NAME, 'btn-success'))
                        )
                        botao_baixar.click()
                        logger.info(f"Download iniciado para {empresa} - {banco_nome} - conta {i+1}")
                        time.sleep(5)  # Aguardar o download iniciar
                        download_iniciado = True
                    except TimeoutException:
                        mensagem_banco = f"Botão de download não encontrado para {empresa} - {banco_nome} - conta {i+1}"
                        logger.warning(mensagem_banco)
                        status_banco = "Erro"
                        try:
                            msg_sem_dados = driver.find_element(By.XPATH, "//div[contains(text(), 'Nenhum registro encontrado')]")
                            if msg_sem_dados:
                                mensagem_banco = "Nenhum registro encontrado para esta conta no período selecionado"
                                logger.info(mensagem_banco)
                                adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
                        except:
                            adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)

                    # Mover o arquivo baixado apenas se o download foi iniciado
                    nome_arquivo = None
                    if download_iniciado:
                        nome_arquivo = mover_arquivo(ano, nome_mes, logger)

                    if nome_arquivo:
                        logger.info(f"Arquivo processado e movido: {nome_arquivo}")
                        contas_processadas.append(f"{banco_nome}_{i+1}")
                        mensagem_banco = f"Arquivo processado: {nome_arquivo}"
                        status_banco = "Sucesso"
                        adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
                    else:
                        mensagem_banco = f"Não foi possível mover o arquivo para {empresa} - {banco_nome} - conta {i+1}"
                        logger.warning(mensagem_banco)
                        status_banco = "Erro"
                        adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
                except Exception as e:
                    mensagem_banco = f"Erro ao processar conta {i+1} do banco {banco_nome}: {str(e)}"
                    logger.error(mensagem_banco)
                    status_banco = "Erro"
                    adicionar_resumo_banco_unico(resumo_bancos, empresa, banco_nome, status_banco, mensagem_banco)
        if not contas_processadas:
            erro_msg = f"Nenhuma conta bancária foi processada para a empresa {empresa}"
            logger.warning(erro_msg)
            return False

        logger.info(f"Processamento concluído para empresa {empresa}. Contas processadas: {', '.join(contas_processadas)}")
        return True
    except TimeoutException as e:
        erro_msg = f"Timeout ao processar empresa {empresa}: {str(e)}"
        logger.error(erro_msg)
        return False
    except Exception as e:
        erro_msg = f"Erro ao processar empresa {empresa}: {str(e)}"
        logger.error(erro_msg)
        return False

def identificar_bancos_disponiveis(driver, wait, empresa, logger):
    """
    Identifica todos os bancos disponíveis para uma empresa.
    Retorna um dicionário com os bancos e suas contas.
    """
    logger.info(f"Identificando bancos disponíveis para a empresa: {empresa}")

    # Acessar a página de extrato
    driver.get(URL_EXTRATO)
    logger.info(f"Acessando URL de extrato: {URL_EXTRATO}")

    # Preencher empresa
    time.sleep(3)
    campo_empresa = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="autocompleter-empresa-autocomplete"]')))
    campo_empresa.clear()
    campo_empresa.send_keys(empresa)
    campo_empresa.send_keys(Keys.RETURN)
    campo_empresa.send_keys(Keys.RETURN)
    logger.info(f"Empresa '{empresa}' selecionada")
    time.sleep(3)

    # Selecionar banco
    botao_bancos = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bankDiv"]')))
    botao_bancos.click()
    logger.info("Botão de bancos clicado")
    time.sleep(2)

    # Estrutura para armazenar informações sobre os bancos disponíveis
    bancos_disponiveis = {}

    # Verificar cada tipo de banco
    for banco_nome, classe_botao in BANCO_CLASSES.items():
        try:
            # Verificar se existe o botão para este banco
            todos_botoes = driver.find_elements(By.CLASS_NAME, classe_botao)

            # Filtrar apenas os botões de conta bancária (excluir os botões de exclusão)
            botoes_validos = []
            for botao in todos_botoes:
                id_botao = botao.get_attribute('id')
                texto_botao = botao.text.strip()

                # Armazenar informações sobre o botão para uso posterior
                if id_botao and not id_botao.startswith('delete-') and texto_botao.lower() == "ver lançamentos":
                    botoes_validos.append({
                        'id': id_botao,
                        'texto': texto_botao,
                        'classe': classe_botao
                    })
                    logger.debug(f"Botão válido encontrado: ID={id_botao}, Texto={texto_botao}")

            if not botoes_validos:
                logger.info(f"Banco {banco_nome} (classe {classe_botao}) não encontrado para esta empresa")
            else:
                logger.info(f"Encontrado(s) {len(botoes_validos)} conta(s) do banco {banco_nome}")
                # Armazenar informações sobre este banco para processamento posterior
                bancos_disponiveis[banco_nome] = botoes_validos
        except Exception as e:
            erro_msg = f"Erro ao verificar banco {banco_nome}: {str(e)}"
            logger.error(erro_msg)

    return bancos_disponiveis

# Função para registrar erros no arquivo de log
def registrar_erro_no_arquivo(empresa, banco, motivo, ano, nome_mes, logger):
    # Criar estrutura de diretórios: I:\Contabilidade\Banco Online\{ano}\{mes}
    destino_dir = os.path.join(BASE_DESTINO_DIR, str(ano), nome_mes)

    # Criar pasta do ano/mês se não existir
    if not os.path.exists(destino_dir):
        os.makedirs(destino_dir)
        logger.info(f"Pasta ano/mês criada: {destino_dir}")

    # Nome do arquivo de log de erros
    log_file = os.path.join(destino_dir, f"erros_download_{datetime.now().strftime('%Y%m%d')}.txt")

    # Verificar se o arquivo já existe
    arquivo_existe = os.path.exists(log_file)

    # Registrar o erro no arquivo
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    erro_msg = f"[{timestamp}] Empresa: {empresa} | Banco: {banco} | Erro: {motivo}\n"

    try:
        # Abrir o arquivo em modo append (ou criar se não existir)
        with open(log_file, 'a', encoding='utf-8') as f:
            # Se o arquivo não existia, adicionar um cabeçalho
            if not arquivo_existe:
                data_atual = datetime.now().strftime("%d/%m/%Y")
                cabecalho = f"=== REGISTRO DE ERROS DE DOWNLOAD - {data_atual} ===\n"
                cabecalho += f"Diretório de destino: {destino_dir}\n"
                cabecalho += "=" * 50 + "\n\n"
                f.write(cabecalho)

            f.write(erro_msg)
        logger.info(f"Erro registrado no arquivo de log: {log_file}")
    except Exception as e:
        logger.error(f"Erro ao registrar no arquivo de log: {str(e)}")

def adicionar_resumo_banco_unico(resumo_bancos, empresa, banco, status, mensagem):
    entrada = {
        "empresa": empresa,
        "banco": banco,
        "status": status,
        "mensagem": mensagem
    }
    if entrada not in resumo_bancos:
        resumo_bancos.append(entrada)

def mover_arquivo(ano, nome_mes, logger):
    # Criar estrutura de diretórios: I:\Contabilidade\Banco Online\{ano}\{mes}
    destino_dir = os.path.join(BASE_DESTINO_DIR, str(ano), nome_mes)

    # Criar pasta do ano/mês se não existir
    if not os.path.exists(destino_dir):
        os.makedirs(destino_dir)
        logger.info(f"Pasta ano/mês criada: {destino_dir}")

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
        # Manter o nome original do arquivo
        nome_arquivo_original = os.path.basename(arquivo_encontrado)
        caminho_destino = os.path.join(destino_dir, nome_arquivo_original)

        try:
            shutil.move(arquivo_encontrado, caminho_destino)
            logger.info(f"Arquivo movido para: {caminho_destino}")
            return nome_arquivo_original
        except Exception as e:
            logger.error(f"Erro ao mover arquivo: {str(e)}")
            return None
    else:
        logger.warning("Arquivo não encontrado para mover")
        return None

def gerar_relatorio_pdf(resumo_bancos, caminho_pdf="relatorios/relatorio_execucao_detalhado.pdf"):
    if not os.path.exists("relatorios"):
        os.makedirs("relatorios")
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    try:
        pdf.image("conttrolare.png", x=10, y=8, w=33)
    except Exception:
        pass
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Relatório de Execução - RoboParis", ln=True, align="C")
    pdf.ln(10)
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True, align="C")
    pdf.ln(10)

    empresas = set([item['empresa'] for item in resumo_bancos])
    bancos = set([item['banco'] for item in resumo_bancos])
    total = len(resumo_bancos)
    sucesso = sum(1 for r in resumo_bancos if r["status"] == "Sucesso")
    erro = total - sucesso
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f"Total de empresas: {len(empresas)}", ln=True)
    pdf.cell(0, 10, f"Total de bancos: {len(bancos)}", ln=True)
    pdf.cell(0, 10, f"Total de processamentos: {total}", ln=True)
    pdf.cell(0, 10, f"Processados com sucesso: {sucesso}", ln=True)
    pdf.cell(0, 10, f"Com erro: {erro}", ln=True)
    pdf.ln(10)

    pdf.set_font("Arial", 'B', 11)
    pdf.cell(50, 10, "Empresa", border=1)
    pdf.cell(35, 10, "Banco", border=1)
    pdf.cell(25, 10, "Status", border=1)
    pdf.cell(80, 10, "Mensagem", border=1)
    pdf.ln()
    pdf.set_font("Arial", '', 10)
    for item in resumo_bancos:
        if item["status"] != "Sucesso":
            pdf.cell(50, 10, item["empresa"], border=1)
            pdf.cell(35, 10, item["banco"], border=1)
            pdf.cell(25, 10, item["status"], border=1)
            mensagem = item["mensagem"]
            y_before = pdf.get_y()
            pdf.multi_cell(80, 10, mensagem, border=1)
            y_after = pdf.get_y()
            if y_after - y_before > 10:
                pdf.set_y(y_after)
            else:
                pdf.ln()
    pdf.output(caminho_pdf)

def main():
    # Configurar logging
    logger = setup_logging()
    logger.info("=== INICIANDO ROBÔ PARIS ===")

    # Dicionário para armazenar erros já registrados (evitar duplicação)
    erros_registrados = {}
    resumo_bancos = []

    # Verificar se o diretório base de destino existe
    if not os.path.exists(BASE_DESTINO_DIR):
        os.makedirs(BASE_DESTINO_DIR)
        logger.info(f"Diretório base de destino criado: {BASE_DESTINO_DIR}")
    else:
        logger.info(f"Diretório base de destino encontrado: {BASE_DESTINO_DIR}")

    # Mostrar diretório de download
    logger.info(f"Diretório de download configurado: {DOWNLOAD_DIR}")

    # Calcular datas do mês anterior
    data_inicial, data_final, ano, nome_mes = calcular_datas_mes_anterior()
    logger.info(f"Período calculado: {data_inicial} a {data_final} (Ano: {ano}, Mês: {nome_mes})")

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

        df = pd.read_excel(EXCEL_PATH)
        logger.info(f"Total de empresas na planilha: {len(df)}")

        # Processar cada empresa
        resumo_empresas = []
        empresas_processadas = 0
        empresas_com_erro = 0

        # Usamos as datas calculadas automaticamente para todas as empresas
        for index, row in df.iterrows():
            empresa = row['Empresa']

            logger.info(f"\n{'='*50}")
            logger.info(f"EMPRESA {index+1}/{len(df)}: {empresa}")
            logger.info(f"{'='*50}")

            # Tentar processar a empresa até 3 vezes em caso de falha
            max_tentativas = 2
            tentativa = 1
            sucesso = False
            erro_msg = ""
            while tentativa <= max_tentativas and not sucesso:
                try:
                    logger.info(f"Tentativa {tentativa}/{max_tentativas} para empresa {empresa}")
                    # Passar o dicionário de erros registrados para a função processar_empresa
                    resultado = processar_empresa(driver, wait, empresa, data_inicial, data_final, ano, nome_mes, logger, erros_registrados, resumo_bancos)

                    if resultado:
                        logger.info(f"Empresa {empresa} processada com sucesso na tentativa {tentativa}")
                        empresas_processadas += 1
                        sucesso = True
                        erro_msg = "Processada com sucesso"
                    else:
                        logger.warning(f"Falha ao processar empresa {empresa} na tentativa {tentativa}")
                        # Incrementar o contador de tentativas
                        tentativa += 1

                        # Se foi a última tentativa, registrar os erros no arquivo de log
                        if tentativa > max_tentativas:
                            erro_msg = f"Falha ao processar empresa {empresa} após {max_tentativas} tentativas"
                            logger.error(erro_msg)

                            # Verificar se este erro já foi registrado para evitar duplicação
                            chave_erro = f"{empresa}_todas_tentativas"
                            if chave_erro not in erros_registrados:
                                registrar_erro_no_arquivo(empresa, "Todos", erro_msg, ano, nome_mes, logger)
                                erros_registrados[chave_erro] = True
                        else:
                            logger.info(f"Aguardando 5 segundos antes da próxima tentativa {tentativa}/{max_tentativas}...")
                            time.sleep(5)
                except Exception as e:
                    erro_msg = f"Erro não tratado ao processar {empresa} na tentativa {tentativa}: {str(e)}"
                    logger.error(erro_msg)
                    # Incrementar o contador de tentativas
                    tentativa += 1
            if not sucesso:
                erro_msg = f"Todas as {max_tentativas} tentativas falharam para empresa {empresa}"
                logger.error(erro_msg)
                # Não precisamos registrar novamente aqui, já foi registrado na última tentativa
                empresas_com_erro += 1
            resumo_empresas.append({
                "empresa": empresa,
                "status": "Sucesso" if sucesso else "Erro",
                "mensagem": erro_msg
            })
        # Resumo final
        logger.info("\n\n=== RESUMO DA EXECUÇÃO ===")
        logger.info(f"Total de empresas: {len(df)}")
        logger.info(f"Empresas processadas com sucesso: {empresas_processadas}")
        logger.info(f"Empresas com erro: {empresas_com_erro}")
        logger.info("=== FIM DA EXECUÇÃO ===\n")
        if resumo_bancos:
            gerar_relatorio_pdf(resumo_bancos)
            logger.info("Relatório PDF detalhado gerado com sucesso.")
    except Exception as e:
        erro_msg = f"Erro crítico na execução: {str(e)}"
        if logger:
            logger.error(erro_msg)
            try:
                # Tentar registrar o erro no arquivo de log
                registrar_erro_no_arquivo("Sistema", "Todos", erro_msg, ano, nome_mes, logger)
            except Exception as log_error:
                logger.error(f"Não foi possível registrar o erro no arquivo: {str(log_error)}")
        else:
            print(erro_msg)
    finally:
        if driver:
            logger.info("Fechando o navegador")
            driver.quit()

if __name__ == "__main__":
    main()
