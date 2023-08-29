# Importações
import pandas as pd
import time
import pyautogui
import requests
from openpyxl import Workbook
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.webdriver_chrome import webDriver_chrome
import keyboard

# Importação do webdriver
dr = webDriver_chrome()


def login():
    # Função para realizar o login
    dr.driver.get('https://kroton.elaw.com.br/login.elaw')
    dr.driver.find_element(By.ID, 'username').send_keys('AMiceli2')
    dr.driver.find_element(By.ID, 'password').send_keys('COGN@2023')
    dr.driver.find_element(By.XPATH, '//*[@id="j_id_a_1_5_f"]').click()


def search():
    # Função para buscar o elemento de pesquisa
    dr.driver.find_element(
      By.XPATH, '//*[@id="j_id_2d_1"]/ul/li[2]/a'
    ).click()
    dr.driver.find_element(
      By.XPATH, '//*[@id="menu-form-contencioso:j_id_2d_a_4"]/a'
    ).click()


def workbook_reading():
    # Função para leitura da tabela, percorrendo os números de processo
    df = pd.read_excel('Pasta1.xlsx')
    if 'Número do Processo' in df.columns:
      numero_processos = df.loc[df['Número do Processo'] == 'Número do Processo']
      url = 'https://kroton.elaw.com.br/processoView.elaw'
      for row in df.iter_rows(min_row=2, values_only=True):
          try:
              print(f"Processo: {numero_processo}")
              campo_busca = dr.driver.find_element(
                By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
              )
              campo_busca.clear()
              campo_busca.send_keys(numero_processo)

              dr.driver.find_element(
                By.XPATH, '//*[@id="btnPesquisar"]'
              ).click()

              dr.driver.find_element(
                By.ID, 'dtProcessoResults:0:btnProcesso'
              ).click()
              tipo_processo = df.loc[df['Evento'] == 'Evento']
              
              if tipo_processo == 'Inicial':
                # Procurar por elementos <td> com o texto "Cópia Inicial"
                inicial = dr.driver.find_elements(By.XPATH, "//td[contains(text(), 'Cópia Inicial')]")
                if inicial:
                    print("Cópia Inicial encontrada na página.")
                    dr.driver.find_element(
                          By.XPATH, '''
                          //button[contains(@id,'btnVoltar')]
                          '''
                        ).click()
                    dr.driver.find_element(
                        By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                      ).clear()
                    continue                
                else:
                    print("Cópia Inicial não encontrada na página.")
                    dr.driver.find_element(
                  By.XPATH, '''
                    //*[@id="tabViewProcesso:j_id_i3_4_1_3_ac:j_id_i3_4_1_3_ag_1"]
                    '''
                  ).click()

                    wait = WebDriverWait(dr.driver, 10)
                    element = wait.until(EC.presence_of_element_located(
                      (By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]')
                      ))
                    element.click()

                    dr.driver.find_element(
                      By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                    ).send_keys('Acompanhamento')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(3)
                    pyautogui.press('tab')
                    time.sleep(5)

                    wait = WebDriverWait(dr.driver, 10)
                    text = wait.until(EC.presence_of_element_located((
                      By.XPATH, '''
                        //textarea[contains(@id,'j_id_2n:txtOcorrenciaAndamento')]
                      '''
                    )))
                    text.send_keys('Cópia Inicial')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(5)

                    wait = WebDriverWait(dr.driver, 10)
                    elemento = wait.until(EC.presence_of_element_located((
                      By.XPATH, '''
                        //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                      '''
                    )))
                    elemento.click()
                    
                    dr.driver.find_element(
                      By.XPATH, '''
                        //input[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter']
                        '''
                    ).send_keys('Inicial', Keys.ARROW_DOWN, Keys.ARROW_DOWN)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(5)

                    dr.driver.find_element(
                      By.XPATH, '''
                        //span[contains(@id,'label')]
                      '''
                      ).click()
                    nome_arquivo = df.loc[df['Documentos'] == 'Documentos']

                    time.sleep(2)

                    pyautogui.write(nome_arquivo)
                    time.sleep(5)
                    pyautogui.press('enter')
                    time.sleep(7)

                    dr.driver.find_element(
                              By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                            ).click()

                    waitoclick = WebDriverWait(dr.driver, 20)
                    bttnclick = waitoclick.until(EC.presence_of_element_located((
                      By.XPATH, '''
                        //button[contains(@id,'btnVoltar')]
                      '''
                    )))
                    bttnclick.click()
                    pass   
              if tipo_processo == 'Petições Diversas':
                try:
                  print('Petições Diversas encontradas na página')
                  dr.driver.find_element(
                  By.XPATH, '''
                    //*[@id="tabViewProcesso:j_id_i3_4_1_3_ac:j_id_i3_4_1_3_ag_1"]
                    '''
                  ).click()

                  wait = WebDriverWait(dr.driver, 10)
                  element = wait.until(EC.presence_of_element_located(
                      (By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]')
                      ))
                  element.click()

                  dr.driver.find_element(
                      By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                    ).send_keys('Acompanhamento')
                  time.sleep(0.5)
                  pyautogui.press('enter')
                  time.sleep(3)
                  pyautogui.press('tab')
                  time.sleep(5)

                  tipo_entrada = df.loc[df['Ação'] == 'Ação']
                  wait = WebDriverWait(dr.driver, 10)
                  text = wait.until(EC.presence_of_element_located((
                      By.XPATH, '''
                        //textarea[contains(@id,'j_id_2n:txtOcorrenciaAndamento')]
                      '''
                    )))
                  text.send_keys(tipo_entrada)
                  time.sleep(0.5)
                  pyautogui.press('enter')
                  time.sleep(5)

                  wait = WebDriverWait(dr.driver, 10)
                  elemento = wait.until(EC.presence_of_element_located((
                      By.XPATH, '''
                        //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                      '''
                    )))
                  elemento.click()
                    
                  dr.driver.find_element(
                      By.XPATH, '''
                        //input[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter']
                        '''
                    ).send_keys('Petições Diversas', Keys.ARROW_DOWN)
                  time.sleep(0.5)
                  pyautogui.press('enter')
                  time.sleep(5)

                  dr.driver.find_element(
                      By.XPATH, '''
                        //span[contains(@id,'label')]
                      '''
                      ).click()
                  nome_arquivo = df.loc[df['Documentos'] == 'Documentos']

                  time.sleep(2)

                  pyautogui.write(nome_arquivo)
                  time.sleep(5)
                  pyautogui.press('enter')
                  time.sleep(7)

                  dr.driver.find_element(
                              By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                            ).click()

                  waitoclick = WebDriverWait(dr.driver, 20)
                  bttnclick = waitoclick.until(EC.presence_of_element_located((
                      By.XPATH, '''
                        //button[contains(@id,'btnVoltar')]
                      '''
                    )))
                  bttnclick.click()
                except Exception as e:
                  print(f'Mensagem error: {e}')
                  file_path = 'Pasta1.xlsx'
                  df = pd.read_excel(file_path)
                  
                  df = df.iloc[:-1]
                  
                  new_file_path = r'C:\Users\miguel.silva\Python\CD1' 
                  new_workbook = Workbook()
                  new_sheet = new_workbook.active

                    # Adicionar os dados do DataFrame ao novo arquivo Excel
                  for _, row in df.iterrows():
                          new_sheet.append(row)

                      # Salvar o novo arquivo Excel
                  new_workbook.save(new_file_path)

                  print('Última linha removida e novo arquivo Excel criado com sucesso.')
          except Exception as e:
                print(f"Erro ao processar o número de processo: {numero_processo}")
                print(f"Mensagem de erro: {str(e)}")
                break
      else: 
        print('Coluna não existente.')


def main():
    login()
    search()
    workbook_reading()


if __name__ == '__main__':
    main()

input('Press Enter to exit...')
dr.driver.quit()
print('Driver quited successfully.')
