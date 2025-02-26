from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.common.action_chains import ActionChains
# import os
# import pyautogui
import time

# Fazendo o download do WebDriver mais recente e capturando o caminho do arquivo
driver_path = ChromeDriverManager().install()

# Verifica se o arquivo foi baixado e mostra um print
print(f"WebDriver baixado em: {driver_path}")

# Inicia o Chrome com o WebDriver baixado
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# Abrir site GerencieAqui
driver.get("https://app.gerencieaqui.com.br/entrar")

wait = WebDriverWait(driver, 10)

def gerencieAqui():
    # Logar no site GerencieAqui
    email_login = wait.until(EC.visibility_of_element_located((By.ID, "email")))
    email_login.send_keys("as")
    senha_login = wait.until(EC.visibility_of_element_located((By.ID, "senhaLogin")))
    senha_login.send_keys("as")
    senha_login.send_keys(Keys.RETURN)

gerencieAqui()

time.sleep(1800)