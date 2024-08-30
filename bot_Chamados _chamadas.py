# MouseInfo pega a posição do mouse
# No terminal: pip install mouseinfo
# python
# from mouseinfo import mouseInfo
# mouseInfo()

#*****************#

import openpyxl
import pyautogui
import pyperclip
import time
import os

pyautogui.PAUSE = 0.5
# Link portal de serviços dos trip
link_formulario = ('https://bracell.service-now.com/sp?id=sc_cat_item&sys_id=0ab3eb5c1bf1c21017fd404be54bcbd4&referrer=popular_items')  
# Link para portal de serviços onde mostra os numeros dos chamados
link_chamados_portalS = ('https://bracell.service-now.com/sp?id=requests')
# Link para nossa pagina de chamados 
link_fecharRitm = ('https://bracell.service-now.com/nav_to.do?uri=%2F$pa_dashboard.do')

def abrir_navegador():
    # Abrir o navegador
    pyautogui.press('winleft')
    pyautogui.write('edge')
    pyautogui.press('enter')    
    time.sleep(2)

def acessar_formulario_chamado():
    # Navegar para o link do formulário para abrir o chamado
    pyperclip.copy(link_formulario)
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(5)                               # Aguardar o formulário carregar
    pyautogui.click(561,427, duration=1)        # move o mouse ate  NOME LOGIN e clica

# Entrar na planilha 
workbook = openpyxl.load_workbook('chamados_database.xlsx')
# seleciona a aba da planilha 
sheet_chamados = workbook['PrevRonda']

# chama as funções Abrir navegador e depois acessar o formulario
abrir_navegador()
acessar_formulario_chamado()

# Copiar informações de um capy e cola no formulario de chamado
for linha in sheet_chamados.iter_rows(min_row=2):
    
    nome_re = str(linha[0].value)               # move o mouse ate  NOME LOGIN e clica
    pyperclip.copy(nome_re)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('tab')  



    Solicitacao = linha[2].value
    pyautogui.write(Solicitacao)
    time.sleep(0.5)
    pyautogui.press('tab')

 

    descricao = linha[6].value
    descricao_2 = str(linha[7].value)
    texto_concatenado = descricao + descricao_2
    pyautogui.write(texto_concatenado)
    # Tab para botão de adicionar ao carrinho  
    
    time.sleep(2)  
    # Tab para botão  de salvar o chamado
    pyautogui.press('tab') 
    pyautogui.press('tab') 
    pyautogui.press('enter')
    time.sleep(3)

#***************** PEGANDO NUMERO DO CHAMADO *******************

    pyperclip.copy(link_chamados_portalS)                         # Salva o link do portal no ctrl + C para acessar a pagina 
    pyautogui.hotkey('ctrl', 'l')                               
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(5)        

    pyautogui.click(418,422, duration=0.5)                         # move o mouse para selecionar o numero do chamado e abrir ele
    time.sleep(3) 
    pyautogui.click(421,380, duration=0.5)                         # move o mouse para selecionar o numero do chamado e abrir ele
    time.sleep(3) 
    pyautogui.doubleClick(1421,281, duration=0.5)                   # move o mouse e da duplo click no numero do chamado 
    pyautogui.hotkey('ctrl', 'c')                                  # Salva o numero do chamado no ctrl + C
    numero_chamado = pyperclip.paste()                             # salvei o numero do chamado copiado numa variavel.
    print(numero_chamado)
    time.sleep(2) 

#************* Indo Fechar o chamado *****************

    # Acessa a pagina de chamados nosso 
    pyperclip.copy(link_fecharRitm)
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(4)                                                   # tempo para carregar a pagina de chamados

    pyautogui.click(1738,93, duration=1)                            # move o mouse para o campo de BUSCAR
    pyperclip.copy(numero_chamado)                                  # copia o numero do chamado 
    pyautogui.hotkey('ctrl', 'v')                                   # cola pra dar enter
    pyautogui.press('enter')
    time.sleep(3)   

    pyautogui.click(1453,349, duration=1)                         # move o mouse para selecionar o ESTADO do chamado
    estado = linha[8].value
    pyperclip.copy(estado)
    pyautogui.write(estado)
    time.sleep(1)   
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab') 
    
    
    analista = linha[9].value
    pyperclip.copy(analista)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)   
    pyautogui.press('tab')

    pyautogui.rightClick(1031,145, duration=1)  # move o mouse para selecionar Salvar 
    pyautogui.press('enter')
     
   
    pyautogui.click(1437,295, duration=1)                         # move o mouse para selecionar o ESTADO do chamado
    resolvido = linha[10].value
    pyperclip.copy(resolvido)
    pyautogui.write(resolvido)
    time.sleep(1)   
     

    
                                    
    pyautogui.click(1782,143, duration=1)                         # move o mouse para selecionar o botao ATUALIZAR
   # time.sleep(0.5) 

    # Abre uma nova aba e acessa o site novamente
    pyautogui.hotkey('ctrl', 't')
    acessar_formulario_chamado()

# Aparece um alerta na tela para o usuário final do código
pyautogui.alert("O código acabou de rodar, Click Ok para fechar a janela")
