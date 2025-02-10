from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

# Configura as opções do Chrome para abrir em modo anônimo
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

# Inicializa o WebDriver com as opções configuradas
driver = webdriver.Chrome(options=chrome_options)

try:
    # 1. Acessa a página de login do SKY
    driver.get("https://app2.skychart.com.br/skyline-blu-83025/#/login")
    
    # 2. Faz o login
    # Localiza os campos de usuário e senha
    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "dsEmail"))
    )
    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "dsSenha"))
    )

    # Insere as credenciais
    username_field.send_keys("samuel.baecker@blulogistics.com.br")
    password_field.send_keys("060804Sam!")

    # Clica no botão de login
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "botao-login"))
    )
    login_button.click()

    # Aguarda o carregamento da página após o login
    WebDriverWait(driver, 10).until(
        EC.url_contains("painel-usuario")
    )

    # 3. Abre o menu lateral clicando no botão "menu-button"
    menu_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "menu-button"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(menu_button).perform()
    menu_button.click()

    # 4. Aguarda o carregamento do menu lateral e clica em "Utilidades"
    utilidades_menu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "Utilidades"))
    )
    actions.move_to_element(utilidades_menu).perform()
    utilidades_menu.click()

    # 5. Aguarda o carregamento do submenu "Relatórios"
    relatorios_menu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Reporting')]"))
    )
    actions.move_to_element(relatorios_menu).perform()
    relatorios_menu.click()

    # 6. Aguarda o carregamento do submenu "Operacional"
    operacional_menu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Operacional')]"))
    )
    actions.move_to_element(operacional_menu).perform()
    operacional_menu.click()

    # 7. Aguarda o carregamento do item "Massa de Dados Operacional Excel"
    massa_dados_menu = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Massa de Dados Operacional Excel')]"))
    )
    actions.move_to_element(massa_dados_menu).perform()
    massa_dados_menu.click()

    # 8. Aguarda o carregamento da página de relatórios
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "dtAberturalnicial"))
    )

    # 9. Preenche os filtros
    # Preenche a Data de Abertura (últimos dois dias)
    data_abertura = driver.find_element(By.ID, "dtAberturalnicial")
    data_abertura.clear()
    data_abertura.send_keys("01/10/2023 - 02/10/2023")

    # Seleciona o Produto (Importação)
    produto = driver.find_element(By.ID, "produto")
    produto.send_keys("Importação")

    # Seleciona o Status Processo (Aberto)
    status_processo = driver.find_element(By.ID, "status-processo")
    status_processo.send_keys("Aberto")

    # 10. Clica em "Solicitar Excel"
    solicitar_excel_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Solicitar excel')]"))
    )
    actions.move_to_element(solicitar_excel_button).perform()
    solicitar_excel_button.click()

    # Aguarda o download ser concluído
    time.sleep(10)  # Ajuste o tempo conforme necessário

finally:
    # Fecha o navegador
    driver.quit()
