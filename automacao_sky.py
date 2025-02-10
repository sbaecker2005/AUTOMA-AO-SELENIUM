from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

driver = webdriver.Chrome(options=chrome_options)

try:
    driver.get("https://app2.skychart.com.br/skyline-blu-83025/#/login")

    username_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "dsEmail"))
    )
    password_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "dsSenha"))
    )

    username_field.send_keys("samuel.baecker@blulogistics.com.br")
    password_field.send_keys("060804Sam!")

    login_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "botao-login"))
    )
    login_button.click()

    WebDriverWait(driver, 20).until(EC.url_contains("painel-usuario"))

    menu_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "menu-button"))
    )
    menu_button.click()

    utilidades_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "Utilidades"))
    )
    utilidades_menu.click()

    relatorios_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Reporting')]"))
    )
    relatorios_menu.click()

    operacional_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Operacional')]"))
    )
    operacional_menu.click()

    time.sleep(3)

    if "operacional" not in driver.current_url:
        print("Erro: Página errada aberta!")
        driver.quit()

    try:
        toggle_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ui-treetable-toggler"))
        )
        toggle_button.click()
        time.sleep(2)  
    except:
        print("Nenhuma expansão necessária.")

    massa_dados_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Massa de Dados Operacional - Excel')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", massa_dados_menu)
    time.sleep(1)  

    driver.execute_script("arguments[0].click();", massa_dados_menu)

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "dtAberturalnicial"))
    )

    data_abertura = driver.find_element(By.ID, "dtAberturalnicial")
    data_abertura.clear()
    data_abertura.send_keys("01/10/2023 - 02/10/2023")

    produto = driver.find_element(By.ID, "produto")
    produto.send_keys("Importação")

    status_processo = driver.find_element(By.ID, "status-processo")
    status_processo.send_keys("Aberto")

    solicitar_excel_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Solicitar excel')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", solicitar_excel_button)
    driver.execute_script("arguments[0].click();", solicitar_excel_button)

    time.sleep(20)

finally:
    driver.quit()
