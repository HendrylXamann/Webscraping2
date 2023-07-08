#importações:
import pandas as pd 
import time as time
import pyautogui
from selenium import webdriver #navegador
from selenium.webdriver.common.by import By #Localizar elementos
from selenium.webdriver.common.keys import Keys #Envio de teclas
from selenium.webdriver.support.wait import WebDriverWait #Espera
from selenium.webdriver.support import expected_conditions as EC #Atender certas condições 
from selenium.webdriver.support.expected_conditions import (presence_of_element_located) #Aguardar elemento 
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException #Exceção
import datetime #Data do dia
from PySimpleGUI import PySimpleGUI as sg #Interface simpples 

#Interface:
sg.theme('Reddit')
layout = [
     [sg.Text('Usuário'), sg.Input(key='usuario', size=(20, 1))],
    [sg.Text('Senha'), sg.Input(key='senha', password_char='*', size=(20, 1))],
     [sg.Text('Token'), sg.Input(key='token', size=(20, 1))],
     [sg.Button('Entrar')]

 ]
janela = sg.Window('Tela de Login', layout)

event, values = janela.read()
usuario = values['usuario']
senha = values['senha']
token = values['token']

#Variáveis para gerar o arquivo excel com data atual:
data_atual = datetime.datetime.now().strftime("%Y%m%d") #Parte 1: Gerar arquivo final com data do dia;
nome_da_planilha = f"{data_atual}.xlsx" #Parte 2: Gerar arquivo final com data do dia;

#Abrir site:
url_claro = "o link fica aqui" #link do site reservado (empresa privada)
chrome = webdriver.Chrome(url_claro)
chrome.get(url_claro) #Abre

#Inserir as infos e fazer login:
def login2():
        e_usuario = chrome.find_element(By.XPATH,'/html/body/div/div/div/form/div[1]/div/div/div[1]/input')
        e_usuario.send_keys(usuario)
        e_senha = chrome.find_element(By.XPATH,'/html/body/div/div/div/form/div[2]/div/div/div/input')
        e_senha.send_keys(senha)
        time.sleep(1)
        e_entrar = chrome.find_element(By.XPATH,'/html/body/div/div/div/form/div[3]/button')
        e_entrar.click()
        time.sleep(1.3)
        e_token = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[2]/form/div/div[1]/div/div/div[1]/input')
        e_token.send_keys(token)
        time.sleep(1)
        enter_token = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[2]/form/div/div[2]/div/button')
        enter_token.click()
        time.sleep(15)
        return e_usuario,e_senha,e_entrar,e_token, enter_token
e_usuario, e_senha, e_entrar, e_token, enter_token = login2() 

#Leitura do arquivo com os dados a serem analisados e criação de lista para armazenamento:    
nome_do_arquivo_1 = "FonteDeDados.xlsx"
df = pd.read_excel(nome_do_arquivo_1)
lista = []

#Realização da automação:
while not df.empty: #Todas as variáveis tem nomenclaturas correlatas a suas funções, fiz isso para auxiliar na manutenção;
      for index,row in df.iterrows():
                wait = WebDriverWait (chrome, 10)
                chrome.get(url_claro)
                time.sleep(4)               
                e_movlinha = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="pesquisar-cliente"]/div/div/div[1]/div/div/div/span/div/div'))).click()
                actions = ActionChains(chrome)
                actions.send_keys(Keys.DOWN)
                time.sleep(0.5)
                actions.send_keys(Keys.ENTER)
                actions.perform() #executa os dois comandos de cima
                time.sleep(2)
                elemento_cpf = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[2]/form/div/div/div[2]/div/div/div[1]/input') 
                elemento_cpf.send_keys(int(row["CPF"]))
                time.sleep(1)
                elemento_entrar_2 = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[2]/form/div/div/div[3]/div/div/button') 
                elemento_entrar_2.click()
                time.sleep(6) 
                try:  
                        linha_nao_encontrada = chrome.find_element(By.XPATH, '/html/body/div/div/div/div[2]/h1')
                        print("lINHA NÃO LOCALIZADA")
                        continue
                except NoSuchElementException:
                        pass
                        wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="Layer_1"]'))).click() #clica nos tracinhos, canto superior esquerdo
                        time.sleep(1)             
                        clickativacao = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[1]/div/div[2]/div/div/ul/li[2]')
                        clickativacao.click()
                        time.sleep(2)
                        clickativaosimplificada = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[1]/div/div[2]/div/div/div/div/ul/li[1]')
                        clickativaosimplificada.click()
                        time.sleep(4)
                        select_regional = wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="Passo1-body"]/div[1]/div/div'))).click() 
                        actions.send_keys(Keys.DOWN)
                        time.sleep(0.5)
                        actions.send_keys(Keys.ENTER)
                        actions.perform()
                        time.sleep(1.7)
                        plataformaclick = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div[1]/div/div/form/div/div[2]/div[2]/div/div[2]/div/div/label[3]/p')
                        plataformaclick.click()
                        time.sleep(1.7)
                        clickaparelho = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div[1]/div/div/form/div/div[2]/div[2]/div/div[3]/div/div/label[1]/p')
                        clickaparelho.click()
                        time.sleep(1.7)
                        clickportabilidade = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div[1]/div/div/form/div/div[2]/div[2]/div/div[4]/div/div/label[2]/p')
                        clickportabilidade.click()
                        time.sleep(1.7)
                        subtipoooclick = chrome.find_element(By.XPATH,'//*[@id="Passo1-body"]/div[3]/div/div')
                        subtipoooclick.click()
                        time.sleep(1.7)
                        actions.send_keys(Keys.ENTER)
                        actions.perform()
                        time.sleep(2)
                        nextbuton = chrome.find_element(By.XPATH,'//*[@id="btn-wizard-next"]')#OK
                        nextbuton.click()
                        time.sleep(8)
                        clear_nome = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div/div/div[2]/form/div/div[2]/div[1]/div/div/div[3]/div/div/div/div[1]/input')
                        clear_nome.clear()
                        time.sleep(0.2)
                        formnome = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div/div/div[2]/form/div/div[2]/div[1]/div/div/div[3]/div/div/div/div[1]/input')
                        formnome.send_keys(str('N C'))
                        time.sleep(1.7)
                        selectsexo = chrome.find_element(By.XPATH,'//*[@id="cadastroGeneroRadio"]/div[1]/div[1]')
                        selectsexo.click()
                        time.sleep(1.7)
                        nascimentodate = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div/div/div[2]/form/div/div[2]/div[1]/div/div/div[5]/div/div/div/div[1]/input')
                        nascimentodate.send_keys(int('20012000'))
                        time.sleep(2)
                        nomedamae = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div/div/div[2]/form/div/div[2]/div[1]/div/div/div[6]/div/div/div/div[1]/input')
                        nomedamae.send_keys(str('N C'))
                        time.sleep(1.7)
                        rg11 = chrome.find_element(By.XPATH,'//*[@id="cadastroDocumentoText"]')
                        rg11.send_keys(int('4002'))
                        time.sleep(1.7)
                        orgaoexp1 = chrome.find_element(By.XPATH,'//*[@id="rgEmissor"]') 
                        orgaoexp1.send_keys(str('S'))
                        time.sleep(1.7)
                        orgaoexp2 = chrome.find_element(By.XPATH,'//*[@id="cadastroOrgaoEmissorText"]/div[2]/div/div/span') #primeiro click pra abrir o leque de estados
                        orgaoexp2.click()
                        time.sleep(1.7)
                        pyautogui.hotkey('ENTER')
                        time.sleep(2)
                        telefone11 = chrome.find_element(By.XPATH,'//*[@id="cadastroTelefoneText"]')
                        telefone11.send_keys(int('61987654321'))
                        time.sleep(1.7)
                        clear_email = chrome.find_element(By.XPATH,'/html/body/div/div/div/div[3]/div/div/div[2]/form/div/div[2]/div[1]/div/div/div[10]/div/div/div/div[1]/input')
                        clear_email.clear()
                        time.sleep(0.2)
                        pyautogui.hotkey('F11')
                        time.sleep(0.7)
                        novoendereco = chrome.find_element(By.XPATH,'//*[@id="cadastroTrocaEnderecoButton"]/button')
                        novoendereco.click()
                        time.sleep(2)
                        cep = chrome.find_element(By.XPATH,'//*[@id="cadastroCepText"]')
                        cep.send_keys(int('72035506'))
                        time.sleep(3)
                        actions.send_keys(Keys.ENTER)
                        numerocasa = chrome.find_element(By.XPATH,'//*[@id="cadastroNumeroText"]')
                        numerocasa.click()
                        time.sleep(1)
                        clear_casa = chrome.find_element(By.XPATH,'//*[@id="cadastroNumeroText"]')
                        clear_casa.clear()
                        time.sleep(1)
                        numerocasa.send_keys(int('21'))
                        time.sleep(3)
                        botaoproximo1 = chrome.find_element(By.XPATH,'//*[@id="btn-wizard-next"]')
                        botaoproximo1.click()
                        time.sleep(8)
                        wait.until(EC.visibility_of_element_located((By.XPATH,'(/html/body/span/div/div/div/div[2]/div[1]) | (/html/body/span/div/div/div/h1/tag)')))
                        status = chrome.find_element(By.XPATH,'(/html/body/span/div/div/div/div[2]/div[1]) | (/html/body/span/div/div/div/h1/tag)').text
                        time.sleep(1)            
                        # Criação do arquivo, excel, com os resultados da pesquisa: 
                        resultadoss = {'Linha': row['CPF'], 'Resultado': status}
                        lista.append(resultadoss)
                        df_lista = pd.DataFrame(lista, columns=['Linha', 'Resultado'])
                        df_lista.to_excel(nome_da_planilha, index=False)
                
print("Portanto alegra-te, Ó Iniciado, pois quanto maior a tua prova, maior o teu Triunfo.")        
                
