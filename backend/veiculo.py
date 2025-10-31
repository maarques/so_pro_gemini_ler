# backend/veiculo.py

import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ==================================================================
# ATENÇÃO: VOCÊ PRECISA CONFIGURAR ESTAS VARIÁVEIS
# ==================================================================

# 1. URL da página de LOGIN do sistema (imagem 1)
URL_LOGIN = "https://www.valeshop.com.br/portal/valeshop/!AP_LOGIN?p_tipo=1"

# 2. Localizadores dos campos do formulário (IDs, Nomes, etc.)
#    Use o "Inspecionar Elemento" do navegador para encontrar os
#    atributos 'id' ou 'name' dos campos do formulário.
#    Estes são apenas CHUTES baseados na imagem.
LOCATORS = {
    'cliente_codigo': (By.NAME, 'p_id_cliente'),      # Palpite
    'contrato_informe': (By.NAME, 'p_nr_contrato'),  # Palpite
    'motorista_nome': (By.NAME, 'p_nm_motorista'),      # Palpite
    'motorista_matricula': (By.NAME, 'p_nr_matricula'),# Palpite
    'veiculo_placa': (By.NAME, 'p_nr_placa_veiculo'),        # Palpite
    'veiculo_marca': (By.NAME, 'p_ds_marca_veiculo'),        # Palpite
    'veiculo_modelo': (By.NAME, 'p_nm_modelo_veiculo'),      # Palpite
}
# ==================================================================


def extrair_dados_planilha(caminho_arquivo, logger):
    """
    Lê os dados de células específicas da planilha com base na imagem 2.
    """
    logger(f"Lendo planilha: {caminho_arquivo}")
    try:
        # data_only=True lê o valor da célula (ex: 5644517110) e não a fórmula
        workbook = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        sheet = workbook.active
        
        dados = {}
        
        # Mapeamento com base na sua explicação e na 'image_e6dd59.png'
        # D2 -> Condutor (Nome)
        dados['nome'] = sheet['D2'].value
        # E4 -> CPF (Matrícula)
        dados['matricula'] = sheet['E4'].value
        # D7 -> Placa
        dados['placa'] = sheet['D7'].value
        
        # D6 -> Veículo (Marca e Modelo)
        veiculo_completo = sheet['D6'].value
        if veiculo_completo and ' ' in veiculo_completo:
            partes = veiculo_completo.split(' ', 1)
            dados['marca'] = partes[0]  # Ex: "Ford"
            dados['modelo'] = partes[1] # Ex: "Ranger"
        else:
            dados['marca'] = veiculo_completo
            dados['modelo'] = ""
            
        # Validação simples
        if not all([dados['nome'], dados['matricula'], dados['placa'], dados['marca']]):
            logger(f"AVISO: Dados faltando na planilha {caminho_arquivo}.")
            logger(f"Verifique as células: D2 (Nome), E4 (Matrícula), D7 (Placa), D6 (Veículo).")
            return None
        
        logger("Dados extraídos com sucesso.")
        return dados
        
    except Exception as e:
        logger(f"ERRO ao ler a planilha '{caminho_arquivo}': {e}")
        return None

def preencher_formulario_web(dados_veiculo, logger):
    """
    Inicia o Selenium, pausa para login manual e preenche o formulário.
    """
    logger("Iniciando automação com Selenium...")
    
    if URL_LOGIN == "COLOQUE_A_URL_DE_LOGIN_DO_SISTEMA_AQUI":
        logger("ERRO FATAL: A 'URL_LOGIN' não foi definida em 'backend/veiculo.py'.")
        logger("Edite o arquivo e adicione a URL correta.")
        return

    driver = None
    try:
        # Inicia o serviço do ChromeDriver automaticamente
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get(URL_LOGIN)
        
        logger("="*50)
        logger("--- AÇÃO MANUAL NECESSÁRIA ---")
        logger("O navegador foi aberto.")
        logger("1. Faça o login no sistema.")
        logger("2. Resolva o CAPTCHA.")
        logger("3. Navegue até a tela 'Cadastrar Vendas sem Cartão'.")
        logger("\nIMPORTANTE: Quando o formulário (imagem 1) estiver VISÍVEL,")
        logger("pressione a tecla 'Enter' no CONSOLE (terminal) onde")
        logger("você iniciou a aplicação.")
        logger("="*50)
        
        # Esta linha vai pausar o script (na thread) e esperar
        # que o usuário pressione 'Enter' no console/terminal.
        input("Pressione Enter no CONSOLE (terminal) para continuar...")
        
        logger("Continuando automação... Preenchendo formulário.")
        
        # Espera o primeiro campo (Código do Cliente) estar disponível (até 30s)
        wait = WebDriverWait(driver, 30)
        wait.until(EC.visibility_of_element_located(LOCATORS['cliente_codigo']))
        
        # Preenche os campos
        logger("Preenchendo campos fixos...")
        driver.find_element(*LOCATORS['cliente_codigo']).send_keys("3359")
        driver.find_element(*LOCATORS['contrato_informe']).send_keys("00101033590125")
        
        logger("Preenchendo dados do motorista...")
        driver.find_element(*LOCATORS['motorista_nome']).send_keys(dados_veiculo['nome'])
        driver.find_element(*LOCATORS['motorista_matricula']).send_keys(dados_veiculo['matricula'])
        
        logger("Preenchendo dados do veículo...")
        driver.find_element(*LOCATORS['veiculo_placa']).send_keys(dados_veiculo['placa'])
        driver.find_element(*LOCATORS['veiculo_marca']).send_keys(dados_veiculo['marca'])
        driver.find_element(*LOCATORS['veiculo_modelo']).send_keys(dados_veiculo['modelo'])
        
        logger("Formulário preenchido com sucesso!")
        logger("A automação irá pausar por 15 segundos antes de fechar o navegador.")
        time.sleep(15)
        
    except Exception as e:
        logger(f"ERRO durante a automação Selenium: {e}")
        logger("Verifique se os 'LOCATORS' no arquivo 'backend/veiculo.py' estão corretos.")
        
    finally:
        if driver:
            driver.quit()
        logger("Navegador fechado. Automação concluída.")
