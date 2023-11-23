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
import pandas as pd
import time
import pyautogui
import win32com.client as win32
import os
from datetime import datetime
import numpy as np


# iniciar webdriver
class AutomacaoSantanderBenner():
  # Ler Exel
    def __init__(self):
      self.driver = webdriver.Chrome()
      self.df = None
      self.df = pd.read_excel('Pasta1.xlsx')
      # self.new_df = pd.read_excel('analisado.xlsx')
      
      
  # acessar benner
    def conectar_internet(self):
      self.driver.get("https://www.santandernegocios.com.br/portaldenegocios/#/externo")
      
      
  # fazer login no benner
    def logar_santander(self):
      login_input = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "userLogin__input")))
      login_input.send_keys("EX26078")
      time.sleep(2)
      password_input = self.driver.find_element(By.ID, "userPassword__input")
      password_input.send_keys("@fer2305")
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
      tarefas_button = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tarefas"]/a')))
      tarefas_button.click()
      time.sleep(5)
      if 'Número Integração' in self.df.columns.astype(str):
        dossie = self.df.loc[self.df['Número Integração'] == 'Número Integração'].iloc[1:]
        # Clicar no botão de tarefas 
        for index, row in self.df.iterrows():  
          try:   
            numero_dossie = row['Número Integração']
            print(f"Dossiê: {numero_dossie}")
            
            textentry = self.driver.find_element(By.XPATH, "//input[contains(@id,'ctl00_Main_WFL_TASKS_INBOX_FilterControl_GERAL_1__TITULO')]")
            textentry.send_keys(str(numero_dossie))
            time.sleep(1)
            
            self.driver.find_element(By.XPATH,'//*[@id="ctl00_Main_WFL_TASKS_INBOX_FilterControl_FilterButton"]').click()
            time.sleep(3)
          except Exception as e:
            print(f"Error: {e}")          
          
  # clicar na tarefa
  # clicar em anexar
  # inserir  nome
  # inserir pasta acordo principal
  # inserir arquivo
  # salvar
  # clicar em editar
  # clicar em concluir
  # seguir para o próximo dossiê
  # deletar o numero já concluido
  # mover o numero já concluido para uma próxima planilha 
  # se der error reiniciar o driver
    def reiniciar_programa(self):
      self.driver.quit()
      # Aguardar antes de reiniciar
      print("Reiniciando o programa em 3 segundos...")
      time.sleep(3)
      
      
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
  
  