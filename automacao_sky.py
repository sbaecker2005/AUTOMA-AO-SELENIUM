from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

driver = webdriver.Chrome(options=chrome_options)

def clicar_com_scroll(driver, elemento):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    driver.execute_script("arguments[0].click();", elemento)

def login(driver):
    driver.get("https://app2.skychart.com.br/skyline-blu-83025/#/login")
    driver.maximize_window()
    username_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "dsEmail")))
    password_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "dsSenha")))
    username_field.send_keys("samuel.baecker@blulogistics.com.br")
    password_field.send_keys("060804Sam!")
    login_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "botao-login")))
    login_button.click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "menu-button")))

def remover_overlay(driver):
    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-dialog-mask")))

def navegar_para_relatorio(driver):
    menu_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "menu-button")))
    clicar_com_scroll(driver, menu_button)
    utilidades_menu = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "Utilidades")))
    clicar_com_scroll(driver, utilidades_menu)
    relatorios_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Reporting')]"))
    )
    clicar_com_scroll(driver, relatorios_menu)

def acessar_menu_operacional(driver):
    operacional_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Operacional')]"))
    )
    clicar_com_scroll(driver, operacional_menu)
    try:
        toggle_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ui-treetable-toggler"))
        )
        clicar_com_scroll(driver, toggle_button)
        time.sleep(2)
    except TimeoutException:
        pass

def acessar_massa_dados(driver):
    tentativa = 0
    max_tentativas = 5
    while tentativa < max_tentativas:
        try:
            massa_dados_menu = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Massa de Dados Operacional - Excel')]"))
            )
            clicar_com_scroll(driver, massa_dados_menu)
            time.sleep(2)
            if not massa_dados_menu.is_displayed():
                raise TimeoutException
            print("Elemento 'Massa de Dados Operacional - Excel' acessado com sucesso.")
            break
        except TimeoutException:
            tentativa += 1
            if tentativa == max_tentativas:
                driver.save_screenshot(f"erro_massa_dados_operacional_{int(time.time())}.png")
                raise TimeoutException("Elemento 'Massa de Dados Operacional - Excel' não pôde ser acessado.")
            time.sleep(2)

def preencher_formulario(driver):
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "dtAberturalnicial")))
    data_abertura = driver.find_element(By.ID, "dtAberturalnicial")
    data_abertura.clear()
    data_abertura.send_keys("01/10/2023 - 02/10/2023")
    produto = driver.find_element(By.ID, "produto")
    produto.send_keys("Importação")
    status_processo = driver.find_element(By.ID, "status-processo")
    status_processo.send_keys("Aberto")

def solicitar_excel(driver):
    solicitar_excel_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Solicitar excel')]"))
    )
    clicar_com_scroll(driver, solicitar_excel_button)
    time.sleep(20)

try:
    login(driver)
    remover_overlay(driver)
    navegar_para_relatorio(driver)
    acessar_menu_operacional(driver)
    acessar_massa_dados(driver)
    preencher_formulario(driver)
    solicitar_excel(driver)
except Exception as e:
    print(f"Erro durante a execução: {e}")
finally:
    driver.quit()
