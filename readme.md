# Configurações Iniciais: 
Define diretórios de download, caminhos para arquivos de empresas (empresas.xlsx), URLs de login e extrato, além de dicionários para meses em português e mapeamento de bancos.
## Datas do Mês Anterior: 
Calcula automaticamente o período do mês anterior para buscar os extratos.
## Logging: 
Cria e configura arquivos de log para registrar todas as ações e possíveis erros do robô.
## Inicialização do Navegador: 
Usa Selenium WebDriver para abrir o Chrome, maximizar a janela e preparar para automação.
## Login Automático: 
Acessa a página de login do portal, insere usuário e senha, e faz login automaticamente.
## Processamento das Empresas:
 - Lê uma lista de empresas de um arquivo Excel.
 - Para cada empresa, tenta acessar o extrato bancário referente ao período do mês anterior.
 - Identifica bancos disponíveis para cada empresa e faz o download dos extratos.
 - Move os arquivos baixados para pastas organizadas por ano, mês e empresa.
 - Registra erros específicos no log caso algum passo falhe.
## Repetição e Tolerância a Falhas: 
Para cada empresa, tenta até 2 vezes em caso de falha antes de registrar erro definitivo.
## Resumo Final: 
Ao terminar, registra no log quantas empresas foram processadas com sucesso e quantas tiveram erro.