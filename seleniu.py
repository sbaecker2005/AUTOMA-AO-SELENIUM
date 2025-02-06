from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Configura o WebDriver (certifique-se de ter o ChromeDriver instalado)
driver = webdriver.Chrome()

try:
    # 1. Acessa a página de login do SKY
    driver.get("https://app2.skychart.com.br/skyline-blu-83025/#/login")
    # 2. Faz o login
    # Localiza os campos de usuário e senha (substitua pelos seletores reais)
    username_field = driver.find_element(By.ID, "E-mail:")  # Substitua pelo seletor real
    password_field = driver.find_element(By.ID, "Senha:")  # Substitua pelo seletor real

    # Insere as credenciais
    username_field.send_keys("samuel.baecker@blulogistics.com.br")  # Substitua pelo seu usuário
    password_field.send_keys("060804Sam!")  # Substitua pela sua senha

    # Clica no botão de login (substitua pelo seletor real)
    login_button = driver.find_element(By.ID, "login-button")  # Substitua pelo seletor real
    login_button.click()

    # Aguarda o carregamento da página após o login
    wait = WebDriverWait(driver, 10)
    wait.until(EC.url_contains("dashboard"))  # Substitua por uma URL ou elemento da página após o login

    # 3. Abre o menu e navega até "Utilidades > Relatórios > Operacional > Massa de Dados Operacional Excel"
    # Clica no botão do menu (substitua pelo seletor real)
    menu_button = driver.find_element(By.ID, "menu-button")  # Substitua pelo seletor real
    menu_button.click()

    # Navega até "Utilidades > Relatórios > Operacional > Massa de Dados Operacional Excel"
    # Substitua pelos seletores reais dos itens do menu
    utilidades_menu = driver.find_element(By.XPATH, "//span[text()='Utilidades']")
    utilidades_menu.click()

    relatorios_menu = driver.find_element(By.XPATH, "//span[text()='Relatórios']")
    relatorios_menu.click()

    operacional_menu = driver.find_element(By.XPATH, "//span[text()='Operacional']")
    operacional_menu.click()

    massa_dados_menu = driver.find_element(By.XPATH, "//span[text()='Massa de Dados Operacional Excel']")
    massa_dados_menu.click()

    # Aguarda o carregamento da página de relatórios
    wait.until(EC.presence_of_element_located((By.ID, "data-abertura")))  # Substitua pelo seletor real

    # 4. Preenche os filtros
    # Preenche a Data de Abertura (últimos dois dias)
    data_abertura = driver.find_element(By.ID, "data-abertura")  # Substitua pelo seletor real
    data_abertura.clear()
    data_abertura.send_keys("01/10/2023 - 02/10/2023")  # Datas chumbadas (substitua pelas datas desejadas)

    # Seleciona o Produto (Importação)
    produto = driver.find_element(By.ID, "produto")  # Substitua pelo seletor real
    produto.send_keys("Importação")

    # Seleciona o Status Processo (Aberto)
    status_processo = driver.find_element(By.ID, "status-processo")  # Substitua pelo seletor real
    status_processo.send_keys("Aberto")

    # 5. Clica em "Solicitar Excel"
    solicitar_excel_button = driver.find_element(By.ID, "solicitar-excel-button")  # Substitua pelo seletor real
    solicitar_excel_button.click()

    # Aguarda o download ser concluído
    time.sleep(10)  # Ajuste o tempo conforme necessário

finally:
    # Fecha o navegador
    driver.quit()
