from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
# excel
import pandas as pd
import os

# Fazendo o download do WebDriver mais recente e capturando o caminho do arquivo
driver_path = ChromeDriverManager().install()

# Verifica se o arquivo foi baixado e mostra um print
print(f"WebDriver baixado em: {driver_path}")

# Inicia o Chrome com o WebDriver baixado
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

wait = WebDriverWait(driver, 10)
# Abrir site GerencieAqui
driver.get("https://app.gerencieaqui.com.br/entrar")
# Logar no site GerencieAqui
email_login = wait.until(EC.visibility_of_element_located((By.ID, "email")))
email_login.send_keys("barbara.moreira@workongroup.com.br")
senha_login = wait.until(EC.visibility_of_element_located((By.ID, "senhaLogin")))
senha_login.send_keys("Barbara308")
senha_login.send_keys(Keys.RETURN)

# escolher a empresa

acessar_people = wait.until(EC.visibility_of_element_located((By.XPATH, "//tr[@data-ri='1']//td[@role='gridcell'][4]//a[contains(@onclick, 'mojarra.jsfcljs')]")))
acessar_people.click() 
    

#entrar em Pedido de Venda
pedido_venda = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@class='ga-col-md-6 ga-text-decoration-none ga-text-primary atalhos-lista-item' and @href='/venda/']")))
pedido_venda.click()

#novo pedido de venda
novo_pedido = wait.until(EC.visibility_of_element_located((By.ID, "j_idt81")))
novo_pedido.click()


caminho_planilha_nome = pd.read_excel("bases/PROMOTORES_PERIFERIA.xlsx")
caminho_planilha_pedido = pd.read_excel("bases/PITU_Logix.xlsx")
caminho_planilha_nome_periferia = pd.read_excel("bases/PROMOTORES_PERIFERIA.xlsx")
caminho_planilha_pedido_PITU_LOGIX = pd.read_excel("bases/Pasta1.xlsx")
caminho_planilha_pedido_PITU_LOGIX = pd.read_excel("bases/PITU_Logix.xlsx", sheet_name="RESULTADO")

# opcao_combobox_nome = combo_box_nome.currentText()
# if opcao_combobox_nome == "Periferia":
#     caminho_planilha_nome = caminho_planilha_nome_periferia

# opcao_combobox_pedido = combo_box_pedido.currentText()
# if opcao_combobox_pedido == "PITU LOGIX":
#     caminho_planilha_pedido = caminho_planilha_pedido_PITU_LOGIX


for _, linha_produto in caminho_planilha_pedido.iterrows():
    produto = linha_produto["PRODUTO"]
    quantidade = linha_produto["QTDREAL"]
    print(f"Adicionando produto: {produto} | Quantidade: {quantidade}")

    # preencher produto
    campo_produto = wait.until(EC.visibility_of_element_located((By.ID, "frmNovo:basicPojo_input")))
    campo_produto.clear()
    campo_produto.send_keys(produto)
    time.sleep(2)
    campo_produto.send_keys(Keys.RETURN)
    time.sleep(2)

    # preencher quantidade
    # campo_quantidade = wait.until(EC.visibility_of_element_located((By.ID, "frmNovo:quant")))
    campo_quantidade = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "ui-inputfield ui-inputtext ui-widget ui-state-default ui-corner-all")))
    campo_quantidade.clear()
    time.sleep(0.5)
    campo_quantidade.send_keys(quantidade)

    # inserir o produto
    inserir = wait.until(EC.visibility_of_element_located((By.ID, "frmNovo:j_idt152")))
    inserir.click()