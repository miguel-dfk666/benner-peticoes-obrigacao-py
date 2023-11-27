from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from datetime import datetime
from ..config import LOGIN, PASSWORD
import pandas as pd
import time
import pyautogui
import os
import numpy as np



# iniciar webdriver
class AutomacaoSantanderBenner():
  # Ler Exel
    def __init__(self):
      options = webdriver.ChromeOptions()
      options.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
      })
      # Service initialization parameters
      service = Service(ChromeDriverManager().install())
      self.driver = webdriver.Chrome(service=service, options=options)
      self.df = None
      self.df = pd.read_excel('Pasta1.xlsx')
      # self.new_df = pd.read_excel('analisado.xlsx')
      
      
  # acessar benner
    def conectar_internet(self):
      self.driver.get("https://www.santandernegocios.com.br/portaldenegocios/#/externo")
      
      
  # fazer login no benner
    def logar_santander(self):
      login_input = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "userLogin__input")))
      login_input.send_keys(LOGIN)
      time.sleep(2)
      password_input = self.driver.find_element(By.ID, "userPassword__input")
      password_input.send_keys(PASSWORD)
      time.sleep(2)
      login_button = self.driver.find_element(By.XPATH, "/html/body/app/ui-view/login/div/div/div/div/div[2]/div[3]/button[2]")
      login_button.click()
      time.sleep(6)
        
        
  # Mover para Segunda tela do benner
    def ir_para_segunda_tela(self):
      botao_login = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "icon")))
      botao_login.click()
      botao_entrada = self.driver.find_element(By.XPATH, '//*[@id="header"]/div[4]/user-menu/div/nav/div/div[2]/ul/li[2]/a')
      botao_entrada.click()
      time.sleep(6)
      self.aba_nova = self.driver.window_handles[1]
      self.aba_original = self.driver.window_handles[0]
      self.driver.switch_to.window(self.aba_nova)
      
      
  # procurar pelo dossiê
    def pesquisar_processo(self):
       # clicar na tarefa
      tarefas_button = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tarefas"]/a')))
      tarefas_button.click()
      time.sleep(5)
      if 'Número Integração' in self.df.columns.astype(str):
        dossie = self.df.loc[self.df['Número Integração'] == 'Número Integração'].iloc[1:]
        # Clicar no botão de tarefas 
        for index, row in self.df.iterrows():  
          try:   
            numero_dossie = row['Número Integração']
            numero_codigo = row['Número Localizador']
            print(f"Dossiê: {numero_dossie}")
            
            textentry = self.driver.find_element(By.XPATH, "//input[contains(@id,'ctl00_Main_WFL_TASKS_INBOX_FilterControl_GERAL_1__TITULO')]")
            textentry.send_keys(str(numero_dossie))
            time.sleep(1)
            
            self.driver.find_element(By.XPATH,'//*[@id="ctl00_Main_WFL_TASKS_INBOX_FilterControl_FilterButton"]').click()
            time.sleep(3)
            
            # of = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH,f"//td[@class='text-left'][contains(.,'Peticionar em Juízo - Obrigação de fazer - Dossiê: {numero_dossie} - Código: {numero_codigo}')]")))          
            elemnto_click = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Main_WFL_TASKS_INBOX_SimpleGrid"]/tbody/tr[1]/td[3]/a')))
            self.driver.execute_script("arguments[0].click()", elemnto_click)
            time.sleep(0.5)
            
            
            # clicar em anexar
            self.driver.execute_script("__doPostBack('ctl00$Main$K9_INFORMAES','CMD_ANEXARPETICAO')")
           
            
            # Alterar para iframe
            iframe_xpath = '//*[@id="ModalCommand_modal"]/div/div/div[2]/iframe'
            WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, iframe_xpath)))
            
            
            # inserir  nome
            nome = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='ctl00_Main_PR_PROCESSODOCUMENTOS_FORM_PageControl_GERAL_GERAL_NOME']")))
            nome.send_keys(str(row["Ação"]))
            time.sleep(2)
            
            # inserir pasta corpo principal
            pasta = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Main_PR_PROCESSODOCUMENTOS_FORM_PageControl_GERAL_GERAL"]/div[1]/div[4]/div/span/div/span[1]/span[1]/span')))
            pasta.click()
            self.driver.find_element(By.XPATH, '//*[@id="ctl00_Body"]/span/span/span[1]/input').send_keys('Corpo Principal')
            time.sleep(2)
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-ctl00_Main_PR_PROCESSODOCUMENTOS_FORM_PageControl_GERAL_GERAL_ctl19_ctl01_select-results"]/li'))).click()
            time.sleep(3)
            
            # inserir arquivo
            self.driver.find_element(By.XPATH, '//*[@id="ctl00_Main_PR_PROCESSODOCUMENTOS_FORM_PageControl_GERAL_GERAL_dropzone_ARQUIVO"]').click()
            time.sleep(2)
            
            
            self.df.loc[self.df['Documentos'] == 'Documentos']
            pyautogui.write(str(row['Documentos']))
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(10)
            
            # salvar
            self.driver.execute_script("javascript:__doPostBack('ctl00$Main$PR_PROCESSODOCUMENTOS_FORM','Save')")
            time.sleep(6)
            
            # Clicar no botão X
            self.driver.switch_to.parent_frame()
            self.driver.find_element(By.XPATH, '//*[@id="ModalCommand_modalCloseButton"]').click()
            
            # clicar em editar
            self.driver.execute_script("javascript:__doPostBack('ctl00$Main$K9_INFORMAES','Edit')")
            
            # clicar em concluir
            time.sleep(5)
            self.driver.execute_script("javascript:__doPostBack('ctl00$Main$K9_INFORMAES','Finish')")
          
            # Exclua a linha no DataFrame original
            self.df.drop(index, inplace=True)
            time.sleep(5)

            # Salva a planilha atualizada
            self.df.to_excel("Pasta1.xlsx", index=False)
            self.new_df.to_excel('analisado.xlsx', index=False)
            
            
            documento = str(row['Documentos'])
            if documento.lower() == 'nan':
                print("Encerrando o programa, pois a entrada é NaN.")
                self.driver.quit()
                break
          except Exception as e:
            print(f"Error: {e}")
          
  

    def reiniciar_programa(self):
      self.driver.quit()
      # Aguardar antes de reiniciar
      print("Reiniciando o programa em 3 segundos...")
      time.sleep(3)
      self.executar()
      
      
  # Executar programa
    def executar(self):
      while True:
          try:
              self.conectar_internet()
              self.logar_santander()
              self.ir_para_segunda_tela()
              self.pesquisar_processo()
          except Exception as e:
              print(f"Erro: {e}")

          # Após a tentativa de pesquisa ou em caso de erro, encerre o WebDriver
          self.reiniciar_programa()
          

if __name__ == '__main__':
    bot = AutomacaoSantanderBenner()
    bot.executar()
  
  