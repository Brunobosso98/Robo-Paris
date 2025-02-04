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
import tkinter as tk
from tkinter import ttk, messagebox
import sys

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

# Variáveis globais para armazenar as datas
data_inicial_global = None
data_final_global = None

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

            # Clicar no botão 'cadRelac' da linha atual
            # botao_adicionar = wait.until(EC.element_to_be_clickable(
            #     ((By.XPATH, f'//*[@id="accordion_parent"]/tr[{i}]/td[last()]/div/button[contains(@class, "cadRelac")]'))))
            # # //*[@id="accordion_parent"]/tr[{i}]/td[7]/div/button[1]
            # botao_adicionar.click()
            # logger.info(f"Botão 'cadRelac' da linha {i} clicado com sucesso!")

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

def criar_interface():
    """Cria uma interface gráfica moderna com design aprimorado."""
    def iniciar_processamento():
        global data_inicial_global, data_final_global
        try:
            # Obter e validar valores
            campos = [
                entry_dia_inicial.get(),
                entry_mes_inicial.get(),
                entry_ano_inicial.get(),
                entry_dia_final.get(),
                entry_mes_final.get(),
                entry_ano_final.get()
            ]
            
            if not all(campos):
                raise ValueError("Preencha todos os campos")
                
            dia_inicial = campos[0].zfill(2)
            mes_inicial = campos[1].zfill(2)
            ano_inicial = campos[2]
            dia_final = campos[3].zfill(2)
            mes_final = campos[4].zfill(2)
            ano_final = campos[5]

            # Validação numérica
            valores = [
                (int(dia_inicial), "Dia inicial", 1, 31),
                (int(mes_inicial), "Mês inicial", 1, 12),
                (int(dia_final), "Dia final", 1, 31),
                (int(mes_final), "Mês final", 1, 12)
            ]
            
            for valor, nome, minimo, maximo in valores:
                if not (minimo <= valor <= maximo):
                    raise ValueError(f"{nome} deve estar entre {minimo:02d} e {maximo:02d}")

            data_inicial_global = f"{dia_inicial}/{mes_inicial}/{ano_inicial}"
            data_final_global = f"{dia_final}/{mes_final}/{ano_final}"
            # Aviso que as datam foram lidas e o programa irá iniciar
            # messagebox.showinfo("Sucesso", f"Processando de {data_inicial_global} a {data_final_global}")
            root.quit()
        except ValueError as e:
            messagebox.showerror("Erro", str(e))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {str(e)}")

    def on_entry_click(event):
        """Remove the placeholder text on click."""
        if event.widget.get() == event.widget.placeholder:
            event.widget.delete(0, "end")  # Remove placeholder
            event.widget.config(foreground='black')  # Change text color to black

    root = tk.Tk()
    root.title("Seletor de Datas Premium")
    root.minsize(400, 450)
    root.configure(bg='#2a2a2a')

    # Estilo personalizado
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configurações de estilo
    style.configure('TFrame', background='#363636')
    style.configure('TLabel', background='#363636', foreground='#ffffff', font=('Segoe UI', 10))
    style.configure('TButton', font=('Segoe UI', 10, 'bold'), borderwidth=0)
    style.map('TButton',
              background=[('active', '#45a7ff'), ('!active', '#0078d4')],
              foreground=[('active', '#ffffff'), ('!active', '#ffffff')])

    # Container principal
    main_frame = ttk.Frame(root, style='TFrame')
    main_frame.pack(padx=30, pady=30, fill='both', expand=True)

    # Título
    ttk.Label(main_frame, text="Seletor de Período", font=('Segoe UI', 14, 'bold'), 
             foreground='#45a7ff', background='#363636').pack(pady=(0, 20))

    # Container das datas
    date_container = tk.Canvas(main_frame, bg='#363636', highlightthickness=0)
    date_container.pack()

    def create_date_frame(parent, title):
        """Cria um frame estilizado para cada grupo de datas"""
        frame = tk.Frame(parent, bg='#454545', bd=0, highlightthickness=0)
        
        # Borda decorativa
        tk.Frame(frame, bg='#45a7ff', height=2).pack(fill='x')
        ttk.Label(frame, text=title, font=('Segoe UI', 9, 'bold'), 
                 background='#454545', foreground='#ffffff').pack(pady=8)
        
        return frame

    # Data Inicial
    frame_inicial = create_date_frame(date_container, "DATA INICIAL")
    frame_inicial.pack(pady=10, ipadx=15, ipady=5)

    campos_inicial = [
        ("Dia", "DD", frame_inicial),
        ("Mês", "MM", frame_inicial),
        ("Ano", "AAAA", frame_inicial)
    ]

    entries_inicial = []
    for idx, (label, placeholder, frame) in enumerate(campos_inicial):
        ttk.Label(frame, text=label).pack(anchor='w', padx=10)
        entry = ttk.Entry(frame, width=15, font=('Segoe UI', 10))
        entry.placeholder = placeholder  # Store placeholder
        entry.insert(0, placeholder)
        entry.config(foreground='grey')  # Set placeholder color
        entry.bind("<FocusIn>", on_entry_click)  # Bind click event
        entry.pack(padx=10, pady=(0, 10 if idx < 2 else 0))
        entries_inicial.append(entry)

    entry_dia_inicial, entry_mes_inicial, entry_ano_inicial = entries_inicial

    # Data Final
    frame_final = create_date_frame(date_container, "DATA FINAL")
    frame_final.pack(pady=10, ipadx=15, ipady=5)

    campos_final = [
        ("Dia", "DD", frame_final),
        ("Mês", "MM", frame_final),
        ("Ano", "AAAA", frame_final)
    ]

    entries_final = []
    for idx, (label, placeholder, frame) in enumerate(campos_final):
        ttk.Label(frame, text=label).pack(anchor='w', padx=10)
        entry = ttk.Entry(frame, width=15, font=('Segoe UI', 10))
        entry.placeholder = placeholder  # Store placeholder
        entry.insert(0, placeholder)
        entry.config(foreground='grey')  # Set placeholder color
        entry.bind("<FocusIn>", on_entry_click)  # Bind click event
        entry.pack(padx=10, pady=(0, 10 if idx < 2 else 0))
        entries_final.append(entry)

    entry_dia_final, entry_mes_final, entry_ano_final = entries_final

    # Botão estilizado
    btn_style = ttk.Style()
    btn_style.configure('Modern.TButton', 
                       background='#0078d4', 
                       foreground='white',
                       bordercolor='#0078d4',
                       focuscolor='none',
                       font=('Segoe UI', 10, 'bold'),
                       padding=10)
    
    btn_style.map('Modern.TButton',
                 background=[('active', '#005a9e'), ('!active', '#0078d4')],
                 foreground=[('active', 'white'), ('!active', 'white')])

    btn_processar = ttk.Button(main_frame, text="PROCESSAR", 
                              style='Modern.TButton', 
                              command=iniciar_processamento)
    btn_processar.pack(pady=20, ipadx=30)

    root.protocol("WM_DELETE_WINDOW", sys.exit)
    root.mainloop()

# Adiciona a função de retângulo arredondado ao Canvas
tk.Canvas.create_rounded_rectangle = lambda self, x1, y1, x2, y2, radius=25, **kwargs: self.create_polygon(
    x1+radius, y1, x2-radius, y1, x2, y1, x2, y1+radius, x2, y2-radius, x2, y2,
    x2-radius, y2, x1+radius, y2, x1, y2, x1, y2-radius, x1, y1+radius, x1, y1,
    smooth=True, **kwargs
)

def main():
    criar_interface()  # Chama a função para criar a interface

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
            data_inicial = data_inicial_global  # Usar a data inicial da interface
            data_final = data_final_global  # Usar a data final da interface
            
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
