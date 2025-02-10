from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

# Configurações do Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

# Inicializa o WebDriver
driver = webdriver.Chrome(options=chrome_options)

try:
    # 1. Acessa a página de login
    driver.get("https://app2.skychart.com.br/skyline-blu-83025/#/login")

    # 2. Login
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

    # Aguarda o carregamento do painel principal
    WebDriverWait(driver, 20).until(EC.url_contains("painel-usuario"))

    # 3. Abre o menu lateral
    menu_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "menu-button"))
    )
    menu_button.click()

    # 4. Aguarda e clica no menu "Utilidades"
    utilidades_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "Utilidades"))
    )
    utilidades_menu.click()

    # 5. Aguarda e clica no submenu "Relatórios"
    relatorios_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Reporting')]"))
    )
    relatorios_menu.click()

    # 6. Aguarda e clica no submenu "Operacional"
    operacional_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Operacional')]"))
    )
    operacional_menu.click()

    # Aguarda um pouco para garantir que a página carregue corretamente
    time.sleep(3)

    # Verifica se a URL correta está carregada antes de continuar
    if "operacional" not in driver.current_url:
        print("Erro: Página errada aberta!")
        driver.quit()

    # 7. Expande a linha da árvore apenas se necessário
    try:
        toggle_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.ui-treetable-toggler"))
        )
        toggle_button.click()
        time.sleep(2)  # Tempo extra para carregar
    except:
        print("Nenhuma expansão necessária.")

    # 8. Clica no elemento "Massa de Dados Operacional - Excel"
    massa_dados_menu = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Massa de Dados Operacional - Excel')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", massa_dados_menu)
    driver.execute_script("arguments[0].click();", massa_dados_menu)  # Usa JavaScript para evitar problemas de clique

    # Aguarda o carregamento correto da página antes de continuar
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "dtAberturalnicial")))

    # 9. Preenche os filtros
    data_abertura = driver.find_element(By.ID, "dtAberturalnicial")
    data_abertura.clear()
    data_abertura.send_keys("01/10/2023 - 02/10/2023")

    produto = driver.find_element(By.ID, "produto")
    produto.send_keys("Importação")

    status_processo = driver.find_element(By.ID, "status-processo")
    status_processo.send_keys("Aberto")

    # 10. Clica em "Solicitar Excel"
    solicitar_excel_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Solicitar excel')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView();", solicitar_excel_button)
    driver.execute_script("arguments[0].click();", solicitar_excel_button)  # Evita erro de clique invisível

    # Aguarda o download ser concluído
    time.sleep(20)

finally:
    # Fecha o navegador
    driver.quit()
