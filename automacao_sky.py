from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

driver = webdriver.Chrome(options=chrome_options)

def log(message):
    print(f"[LOG] {message}")

def clicar_com_scroll(driver, elemento):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    driver.execute_script("arguments[0].click();", elemento)

def login(driver):
    log("Acessando o site e preenchendo o login...")
    driver.get("https://app2.skychart.com.br/skyline-blu-83025/#/login")
    driver.maximize_window()
    
    username_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "dsEmail")))
    password_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "dsSenha")))
    
    username_field.send_keys("samuel.baecker@blulogistics.com.br")
    password_field.send_keys("060804Sam!")
    
    login_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "botao-login")))
    login_button.click()
    
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "menu-button")))
    log("Login efetuado com sucesso.")

def remover_overlay(driver):
    log("Verificando se há overlays na tela...")
    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-dialog-mask")))
    log("Overlay removido.")

def navegar_para_relatorio(driver):
    log("Navegando até o menu de relatórios...")
    menu_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "menu-button")))
    clicar_com_scroll(driver, menu_button)
    log("Menu principal acessado.")
    
    utilidades_menu = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "Utilidades")))
    clicar_com_scroll(driver, utilidades_menu)
    log("Menu 'Utilidades' acessado.")
    
    relatorios_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Reporting')]"))
    )
    clicar_com_scroll(driver, relatorios_menu)
    log("Menu 'Relatórios' acessado.")

def acessar_menu_operacional(driver):
    log("Tentando acessar o menu 'Operacional'...")
    operacional_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Operacional')]"))
    )
    clicar_com_scroll(driver, operacional_menu)
    log("Menu 'Operacional' acessado.")
    
    try:
        toggle_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ui-treetable-toggler"))
        )
        clicar_com_scroll(driver, toggle_button)
        time.sleep(2)
        log("Botão de expansão clicado.")
    except TimeoutException:
        log("Nenhuma expansão necessária.")

def acessar_massa_dados(driver):
    log("Tentando acessar 'Massa de Dados Operacional - Excel'...")
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
            log("Elemento 'Massa de Dados Operacional - Excel' acessado com sucesso.")
            break
        except TimeoutException:
            tentativa += 1
            log(f"Tentativa {tentativa} de {max_tentativas} falhou.")
            if tentativa == max_tentativas:
                driver.save_screenshot(f"erro_massa_dados_operacional_{int(time.time())}.png")
                raise TimeoutException("Elemento 'Massa de Dados Operacional - Excel' não pôde ser acessado.")
            time.sleep(2)

def preencher_formulario(driver):
    log("Preenchendo o formulário...")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "dtAberturalnicial")))
    
    data_abertura = driver.find_element(By.ID, "dtAberturalnicial")
    data_abertura.clear()
    data_abertura.send_keys("01/10/2023 - 02/10/2023")
    
    produto = driver.find_element(By.ID, "produto")
    produto.send_keys("Importação")
    
    status_processo = driver.find_element(By.ID, "status-processo")
    status_processo.send_keys("Aberto")
    log("Formulário preenchido.")

def solicitar_excel(driver):
    log("Solicitando o Excel...")
    solicitar_excel_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Solicitar excel')]"))
    )
    clicar_com_scroll(driver, solicitar_excel_button)
    log("Solicitação de Excel realizada.")
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
    log(f"Erro durante a execução: {e}")
finally:
    log("Fechando o navegador...")
    driver.quit()
