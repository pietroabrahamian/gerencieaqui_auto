from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
# excel
import pandas as pd
import os
#interface grafica
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QGraphicsDropShadowEffect, QComboBox, QFileDialog, QMessageBox
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt




def gerencieAqui():
    driver = webdriver.Chrome()

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
    opcao_combobox = combo_box.currentText()
    if opcao_combobox == "WorkOn":
        # acessar_workon = wait.until(EC.visibility_of_element_located((By.XPATH, '//tr[@datari="0"//td[@role="gridcell][4]//a')))
        acessar_workon = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(@onclick, 'mojarra.jsfcljs')]")))
        acessar_workon.click()
    if opcao_combobox == "ON JOB":
        acessar_people = wait.until(EC.visibility_of_element_located((By.XPATH, "//tr[@data-ri='1']//td[@role='gridcell'][4]//a[contains(@onclick, 'mojarra.jsfcljs')]")))
        acessar_people.click()    

    #entrar em Pedido de Venda
    pedido_venda = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@class='ga-col-md-6 ga-text-decoration-none ga-text-primary atalhos-lista-item' and @href='/venda/']")))
    pedido_venda.click()

    #novo pedido de venda
    novo_pedido = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[text()='Nova Venda']")))
    novo_pedido.click()

    # vendedor
    vendedor_responsavel = wait.until(EC.visibility_of_element_located((By.ID, 'frmNovo:listaVend_input')))
    vendedor_responsavel.clear()
    vendedor_responsavel.send_keys("Alice")
    li_vendedor = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@class='ui-autocomplete-items ui-autocomplete-list ui-widget-content ui-widget ui-corner-all ui-helper-reset']/li[1]")))
    li_vendedor.click()

    # lendo o excel

    #caminho das planilhas

    caminho_planilha_nome = ""
    caminho_planilha_pedido = ""
    caminho_planilha_nome_periferia = pd.read_excel("bases/PROMOTORES_PERIFERIA.xlsx")
    caminho_planilha_nome_vip = pd.read_excel("bases/PROMOTORES_VIP.xlsx")
    caminho_planilha_pedido_PITU_LOGIX = pd.read_excel("bases/PITU_Logix.xlsx")
    # caminho_planilha_pedido_PITU_LOGIX = pd.read_excel("bases/PITU_Logix.xlsx", sheet_name="RESULTADO")

    opcao_combobox_nome = combo_box_nome.currentText()
    if opcao_combobox_nome == "Periferia":
        caminho_planilha_nome = caminho_planilha_nome_periferia

    if opcao_combobox_nome == "VIP":
        caminho_planilha_nome = caminho_planilha_nome_vip
    
    opcao_combobox_pedido = combo_box_pedido.currentText()
    if opcao_combobox_pedido == "PITU LOGIX":
        caminho_planilha_pedido = caminho_planilha_pedido_PITU_LOGIX

    for _, pessoa in caminho_planilha_nome.iterrows():
        nome_pessoa = pessoa["NOME"]

        try:
            campo_nome = driver.find_element(By.ID, "frmNovo:listaCli_input")
            campo_nome.clear()
            campo_nome.send_keys(nome_pessoa)
            li_nome = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@class='ui-autocomplete-items ui-autocomplete-list ui-widget-content ui-widget ui-corner-all ui-helper-reset']/li[1]")))
            li_nome.click()

            for _, linha_produto in caminho_planilha_pedido.iterrows():
                produto = linha_produto["PRODUTO"]
                quantidade = linha_produto["QTDREAL"]
                print(f"Adicionando produto: {produto} | Quantidade: {quantidade}")

                #preencher produto
                try:
                    campo_produto = wait.until(EC.visibility_of_element_located((By.ID, "frmNovo:basicPojo_input")))
                    campo_produto.clear()
                    campo_produto.send_keys(produto)
                    time.sleep(1)
                    campo_produto.send_keys(Keys.RETURN)
                    time.sleep(1)
                except Exception as e:
                    print(f"⚠️ Erro ao adicionar o produto {produto}: {e}")
                    continue 
                

                # preencher quantidade
                campo_quantidade = wait.until(EC.visibility_of_element_located((By.ID, "frmNovo:quant")))
                campo_quantidade.clear()
                time.sleep(1)
                campo_quantidade = wait.until(EC.visibility_of_element_located((By.ID, "frmNovo:quant")))
                campo_quantidade.click()
                campo_quantidade.send_keys(quantidade)
                time.sleep(1)

                # inserir o produto
                inserir = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[text()='Inserir']")))
                inserir.click()
                time.sleep(1)

            vendedor_responsavel = wait.until(EC.visibility_of_element_located((By.ID, 'frmNovo:listaVend_input')))
            vendedor_responsavel.clear()
            vendedor_responsavel.send_keys("Alice")
            li_vendedor = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@class='ui-autocomplete-items ui-autocomplete-list ui-widget-content ui-widget ui-corner-all ui-helper-reset']/li[1]")))
            li_vendedor.click()


            def esperar_confirmacao():
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Question)
                msg.setWindowTitle("Confirmação")
                msg.setText("Deseja continuar com o processamento?")
                msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                resposta = msg.exec_()  # Exibe a caixa de diálogo e espera pela resposta do usuário
                return resposta == QMessageBox.Yes
            if not esperar_confirmacao():
                print("Processamento interrompido pelo usuário.")
                return
            
            time.sleep(2)

            novo_pedido = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[text()='Nova Venda']")))
            novo_pedido.click()

        except Exception as e:
            print(f"⚠️ Erro ao processar pedido para {nome_pessoa}: {e}")





# Interface Gráfica

#criando a base da interface
app = QApplication(sys.argv)

#criando a jenala principal 
janela = QWidget()
janela.setWindowTitle("GerencieAqui | Automação")
janela.setGeometry( 600, 250, 800, 600)
janela.setStyleSheet("""            
    QWidget {
        background-color: white; 
        font-size: 18px;              
    }                  

    QPushButton{          
        font-size: 22px;       
        background-color: white;
        color: black;
        border: 2px solid blue;
        border-radius: 11px;
        padding: 5;
                                 
    }                         

    QPushButton:hover{
     background-color: blue;
     color: white;                               
    } 
                     
    #button{
        border: 2px solid green;               
    }            

    #button:hover{
     background-color: green;
     color: white;                               
    }                  

    QLabel {
        font-size: 20px;
        padding: 5;
    }
                     
    QLineEdit {         
        border: 0.5px solid black;
        border-radius: 5;
        padding: 2;                        
    }   

    QComboBox {         
        border: 0.5px solid black;
        border-radius: 5;
        padding: 5;                        
    }
    #erro {         
        color: red;                        
    }                                                                        

""")

container = QWidget(janela)
container.setFixedSize(600, 550)


# layout da janela

#layout vertical
layout = QVBoxLayout() # vertical
layout.setAlignment(Qt.AlignCenter)
#layout horiontal 1
h_layout = QHBoxLayout() # horizontal
# h_layout.setAlignment(Qt.AlignCenter)
#layout horiontal 2
h2_layout = QHBoxLayout()
# h2_layout.setAlignment(Qt.AlignCenter)
# layout horizontal 3
h3_layout = QHBoxLayout()
h3_layout.setAlignment(Qt.AlignCenter)
# layout horizontal 4
h4_layout = QHBoxLayout()
# h4_layout.setAlignment(Qt.AlignCenter)

# selecionar o arquivo que contém os colaboradores
label_nome = QLabel("Tipo de Promotor:")
combo_box_nome = QComboBox()
combo_box_nome.setFixedWidth(300)
combo_box_nome.addItem("Periferia")
combo_box_nome.addItem("VIP")

# insira o pedido
label_pedido = QLabel("Insira o pedido:")
combo_box_pedido = QComboBox()
combo_box_pedido.setFixedWidth(300)
combo_box_pedido.addItem("PITU LOGIX")

# escolha a empresa 
label_combo_box = QLabel("Escolha a Empresa:")
combo_box = QComboBox()
combo_box.setFixedWidth(300)
combo_box.addItem("ON JOB")
combo_box.addItem("WorkOn")

# label para aparecer os produtos inclusos
label_resultado = QLabel("")

# label para mostrar se houve erro 
label_erro = QLabel("")
label_erro.setObjectName("erro")


# button acessar o GerencieAqui
button_auto = QPushButton("GerencieAqui")
button_auto.clicked.connect(gerencieAqui)
button_auto.setFixedWidth(230)
button_auto.setFixedHeight(80)
button_auto.setCursor(Qt.PointingHandCursor)

shadow_effect = QGraphicsDropShadowEffect()
shadow_effect.setOffset(0, 4)  # Deslocamento da sombra (x, y)
shadow_effect.setBlurRadius(10)  # Raio de desfoque da sombra
shadow_effect.setColor(QColor(0, 0, 0, 160))  # Cor da sombra (preto com opacidade de 160)

# Aplicando o efeito de sombra ao botão
button_auto.setGraphicsEffect(shadow_effect)



h_layout.addWidget(label_nome)
h_layout.addWidget(combo_box_nome, alignment=Qt.AlignLeft)
h2_layout.addWidget(label_pedido)
h2_layout.addWidget(combo_box_pedido, alignment=Qt.AlignLeft)
h3_layout.addWidget(button_auto)
h4_layout.addWidget(label_combo_box)
h4_layout.addWidget(combo_box, alignment=Qt.AlignLeft)

# h_layout.addStretch()
# h2_layout.addStretch()
layout.addLayout(h_layout)
layout.addLayout(h2_layout)
layout.setSpacing(20)
layout.addLayout(h4_layout)
layout.addWidget(label_resultado)
layout.addWidget(label_erro)
layout.addLayout(h3_layout)
# layout.addStretch()



container.setLayout(layout)
container.move((janela.width() - container.width()) // 2, (janela.height() - container.height()) // 2)
janela.show()
sys.exit((app.exec_()))