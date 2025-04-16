# Robô Paris - Automação Bancária

## Visão Geral
O Robô Paris é uma solução de automação desenvolvida em Python para extrair, organizar e gerenciar extratos bancários de múltiplas empresas e bancos via portal SS Parisi. O sistema automatiza login, navegação, download e arquivamento dos extratos, otimizando o fluxo de trabalho contábil e garantindo eficiência na coleta de dados financeiros.

Além das funcionalidades principais, o projeto conta com um sistema de geração de relatórios PDF detalhados, que apresenta apenas as empresas e bancos que tiveram erros durante o processamento, facilitando auditorias e o acompanhamento de exceções.

## Principais Funcionalidades
- Extração automatizada de extratos bancários
- Processamento em lote de múltiplas empresas
- Organização automática dos arquivos por ano/mês
- Geração de relatórios PDF detalhados para auditoria
- Tratamento robusto de erros com logging detalhado
- Execução em modo headless (ideal para servidores)
- Registro e resumo dos processos realizados

## Detalhes Técnicos
- **Linguagem:** Python
- **Principais Bibliotecas:** Selenium, Pandas, WebDriver Manager, FPDF
- **Arquitetura:**
  - `roboParis.py`: Versão com interface gráfica
  - `roboParisHeadless.py`: Execução sem interface gráfica
  - `relacionamentos.py`: Processamento de relacionamentos entre empresas

## Fluxo de Execução
1. Inicialização e configuração do ambiente
2. Login automático no portal SS Parisi
3. Leitura da planilha Excel com empresas
4. Identificação dos bancos disponíveis por empresa
5. Extração e organização dos extratos
6. Registro detalhado de erros e geração de relatório PDF

## Resumo do Projeto
Solução de automação bancária completa, com foco em eficiência, clareza para auditoria e robustez no tratamento de erros.
