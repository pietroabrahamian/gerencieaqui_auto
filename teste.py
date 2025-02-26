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
#interface grafica
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QGraphicsDropShadowEffect, QComboBox
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt




def gerencieAqui():
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
    opcao_combobox = combo_box.currentText()
    if opcao_combobox == "WorkOn":
        # acessar_workon = wait.until(EC.visibility_of_element_located((By.XPATH, '//tr[@datari="0"//td[@role="gridcell][4]//a')))
        acessar_workon = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(@onclick, 'mojarra.jsfcljs')]")))
        acessar_workon.click()
    if opcao_combobox == "People":
        acessar_people = wait.until(EC.visibility_of_element_located((By.XPATH, "//tr[@data-ri='1']//td[@role='gridcell'][4]//a[contains(@onclick, 'mojarra.jsfcljs')]")))
        acessar_people.click()    

    #entrar em Pedido de Venda
    pedido_venda = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@class='ga-col-md-6 ga-text-decoration-none ga-text-primary atalhos-lista-item' and @href='/venda/']")))
    pedido_venda.click()

    #novo pedido de venda
    novo_pedido = wait.until(EC.visibility_of_element_located((By.ID, "j_idt79")))
    novo_pedido.click()

    # vendedor
    vendedor_responsavel = wait.until(EC.visibility_of_element_located((By.ID, 'frmNovo:listaVend_input')))
    vendedor_responsavel.clear()
    vendedor_responsavel.send_keys("25431")
    li_vendedor = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@class='ui-autocomplete-items ui-autocomplete-list ui-widget-content ui-widget ui-corner-all ui-helper-reset']/li[1]")))
    li_vendedor.click()

# Interface Gráfica


# criando a função para salvar no excel
caminho_arquivo = "bases/base_produtos.xlsx"

def salvar_no_excel():
    produtos_texto = input_pedido.text().strip()

    if produtos_texto:
        produtos_lista = [produto.strip() for produto in produtos_texto.split(";") if produtos_texto.strip()]

        if produtos_lista:
            df = pd.DataFrame({"Produto":produtos_lista})

            try: 
    
                df.to_excel(caminho_arquivo, index=False)
                label_resultado.setText(f"{len(produtos_lista)} produtos salvos com sucesso")
            except Exception as e:
                label_resultado.setText(f"Erro ao salvar os produtos: {e}")
        else:
            label_resultado.setText("Nenhum produto válido inserido")
    else:
        label_resultado.setText("O campo está vazio. Insira produtos para salvar.")

    input_pedido.clear()


#cirando a base da interface
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

    container {
        border: 2px solid blue;
    }                 

    QPushButton{                 
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

""")

container = QWidget(janela)
container.setFixedSize(600, 450)


# layout da janela

#layout vertical
layout = QVBoxLayout() # vertical
layout.setAlignment(Qt.AlignCenter)
#layout horiontal 1
h_layout = QHBoxLayout() # horizontal
h_layout.setAlignment(Qt.AlignCenter)
#layout horiontal 2
h2_layout = QHBoxLayout()
h2_layout.setAlignment(Qt.AlignCenter)
# layout horizontal 3
h3_layout = QHBoxLayout()
h3_layout.setAlignment(Qt.AlignCenter)
# layout horizontal 4
h4_layout = QHBoxLayout()
# h4_layout.setAlignment(Qt.AlignCenter)

# inserir nome label + input
label_nome = QLabel("Nome do Promotor:")
input_nome = QLineEdit()
input_nome.setFixedWidth(300) # tamanho fixo de 300px

# insira o pedido
label_pedido = QLabel("Insira o pedido:")
input_pedido = QLineEdit()
input_pedido.setFixedWidth(300)

# escolha a empresa 
label_combo_box = QLabel("Escolha a Empresa:")
combo_box = QComboBox()
combo_box.setFixedWidth(300)
combo_box.addItem("WorkOn")
combo_box.addItem("People")

# label para aparecer os produtos inclusos
label_resultado = QLabel("")

# button Salvar no Excel
button = QPushButton("Salvar no Excel")
button.setObjectName("button")
button.clicked.connect(salvar_no_excel)
button.setStyleSheet("")
button.setFixedWidth(190) # tamanho fixo
button.setFixedHeight(50) # tamanho fixo
button.setCursor(Qt.PointingHandCursor) # muda o cursor para a mãozinha

# button acessar o GerencieAqui
button_auto = QPushButton("GerencieAqui")
button_auto.clicked.connect(gerencieAqui)
button_auto.setFixedWidth(190)
button_auto.setFixedHeight(50)
button_auto.setCursor(Qt.PointingHandCursor)

# Criando o efeito de sombra
shadow_effect = QGraphicsDropShadowEffect()
shadow_effect.setOffset(0, 4)  # Deslocamento da sombra (x, y)
shadow_effect.setBlurRadius(10)  # Raio de desfoque da sombra
shadow_effect.setColor(QColor(0, 0, 0, 160))  # Cor da sombra (preto com opacidade de 160)

# Aplicando o efeito de sombra ao botão
button.setGraphicsEffect(shadow_effect)

shadow_effect1 = QGraphicsDropShadowEffect()
shadow_effect1.setOffset(0, 4)  # Deslocamento da sombra (x, y)
shadow_effect1.setBlurRadius(10)  # Raio de desfoque da sombra
shadow_effect1.setColor(QColor(0, 0, 0, 160))  # Cor da sombra (preto com opacidade de 160)

# Aplicando o efeito de sombra ao botão
button_auto.setGraphicsEffect(shadow_effect1)



h_layout.addWidget(label_nome)
h_layout.addWidget(input_nome)
h2_layout.addWidget(label_pedido)
h2_layout.addWidget(input_pedido)
h3_layout.addWidget(button)
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
layout.addLayout(h3_layout)
# layout.addStretch()



container.setLayout(layout)
container.move((janela.width() - container.width()) // 2, (janela.height() - container.height()) // 2)
janela.show()
sys.exit((app.exec_()))



