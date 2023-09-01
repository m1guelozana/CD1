import pandas as pd
import time
import pyautogui
from openpyxl import Workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.webdriver_chrome import webDriver_chrome

  # Importação do webdriver
dr = webDriver_chrome()

class ProcessAutomation:
      def __init__(self):
          self.df = None
          # Função para leitura da tabela, percorrendo os números de processo
          self.df = pd.read_excel('Pasta1.xlsx')
      
      def login(self):
          # Função para realizar o login
          dr.driver.get('https://kroton.elaw.com.br/login.elaw')
          dr.driver.find_element(By.ID, 'username').send_keys('AMiceli2')
          dr.driver.find_element(By.ID, 'password').send_keys('COGN@2023')
          dr.driver.find_element(By.XPATH, '//*[@id="j_id_a_1_5_f"]').click()
      
      def search(self):
          # Função para buscar o elemento de pesquisa
          dr.driver.find_element(
            By.XPATH, '//*[@id="j_id_2d_1"]/ul/li[2]/a'
          ).click()
          dr.driver.find_element(
            By.XPATH, '//*[@id="menu-form-contencioso:j_id_2d_a_4"]/a'
          ).click()

        
      def process_inicial(self, numero_processo):
        numero_processos = self.df.loc[self.df['Número do Processo'] == 'Número do Processo'].iloc[1:]
        for index, row in self.df.iterrows():
                try:
                    numero_processo = row[1]
                    print(f"Processo: {numero_processo}")
                    campo_busca = dr.driver.find_element(
                        By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                    )
                    campo_busca.clear()
                    campo_busca.send_keys(str(row['Número do Processo']))

                    dr.driver.find_element(
                        By.XPATH, '//*[@id="btnPesquisar"]'
                    ).click()

                    dr.driver.find_element(
                        By.ID, 'dtProcessoResults:0:btnProcesso'
                    ).click()

                    tipo_processo = row['Evento']
                    if tipo_processo == 'Inicial':
                        # Procurar por elementos <td> com o texto "Cópia Inicial"
                        inicial = dr.driver.find_elements(By.XPATH, "//td[contains(text(), 'Cópia Inicial')]")
                        if inicial:
                            print("Cópia Inicial encontrada na página.")
                            dr.driver.find_element(By.XPATH, '''
                                //button[contains(@id,'btnVoltar')]
                            ''').click()
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
                            element = wait.until(EC.presence_of_element_located((
                                By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]')
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
                            self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                            time.sleep(2)
                            pyautogui.write(str(row['Documentos']))
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
                     print(f"Error message: {e}")
                     file_path = 'Pasta1.xlsx'
                     self.df = pd.read_excel(file_path)
                     self.df = self.df.iloc[:-1]
                    
                     new_file_path = r'C:\Users\miguel.silva\Python\CD'
                     new_workbook = Workbook()
                     new_sheet = new_workbook.active

                     for _, row in self.df.iterrows():
                        new_sheet.append(row)
                        # Salvar o novo arquivo Excel
                        new_workbook.save(new_file_path)
    
      def process_petition(self, numero_processo):
        numero_processos = self.df.loc[self.df['Número do Processo'] == 'Número do Processo'].iloc[1:]
        for index, row in self.df.iterrows():
            try: 
                numero_processo = row[1]
                print(f"Processo: {numero_processo}")
                campo_busca = dr.driver.find_element(
                    By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                )
                campo_busca.clear()
                campo_busca.send_keys(str(row['Número do Processo']))

                dr.driver.find_element(
                    By.XPATH, '//*[@id="btnPesquisar"]'
                ).click()

                dr.driver.find_element(
                    By.ID, 'dtProcessoResults:0:btnProcesso'
                ).click()

                dr.driver.find_element(
                    By.XPATH, ' //*[@id="tabViewProcesso:j_id_i3_4_1_3_ac:j_id_i3_4_1_3_ag_1"]'
                ).click()
                wait = WebDriverWait(dr.driver, 10)
                element = wait.until(EC.presence_of_element_located((
                    By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]'
                )))
                element.click()
                
                dr.driver.find_element(
                    By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                ).send_keys('Acompanhamento')
                pyautogui.press('enter')
                time.sleep(3)
                pyautogui.press('tab')
                time.sleep(5)

                tipo_entrada = self.df.loc[self.df['Observação Tramitação'] == 'Observação Tramitação'].iloc[1:]
                wait = WebDriverWait(dr.driver, 10)
                text = wait.until(EC.presence_of_element_located((
                    By.XPATH, "//textarea[contains(@id,'j_id_2n:txtOcorrenciaAndamento')]"
                )))
                text.send_keys(str(row['Observação Tramitação']))
                time.sleep(1)
                

                waitoclick = WebDriverWait(dr.driver, 10)
                elementor = wait.until(EC.presence_of_element_located((
                         By.XPATH, '''
                         //textarea[contains(@id,'j_id_2n:txtOcorrenciaAndamento')]
                         '''
                         )))
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(5)

                wait = WebDriverWait(dr.driver, 10)
                elemento2 = wait.until(EC.presence_of_element_located((
                    By.XPATH, '''
                    //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                    '''
                )))
                elemento2.click()

                dr.driver.find_element(
                    By.XPATH, '''
                    //input[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter']
                    '''
                ).send_keys(str(row['Evento']), Keys.ARROW_DOWN, Keys.ARROW_DOWN)
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(5)


                dr.driver.find_element(
                By.XPATH, '''
                //span[contains(@id,'label')]
                '''
                ).click()  

                waitertowrite = WebDriverWait(dr.driver, 20)
                elementwrite = waitertowrite.until(EC.presence_of_element_located((
                By.XPATH, "//input[contains(@id,'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter')]"
                )))
                elementwrite.send_keys(str(row['Evento']))
                time.sleep(0.5)
                pyautogui.press('enter')
                
                
                dr.driver.find_element(
                    By.XPATH, "//span[contains(@id,'label')]"
                ).click()
                
                self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                time.sleep(2)
                pyautogui.write(str(row['Documentos']))
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(3)

                waitonclick = WebDriverWait(dr.driver, 20)
                bttnclick = waitonclick.until(EC.presence_of_element_located((
                    By.XPATH, "//button[contains(@id,'btnVoltar')]"
                )))
                bttnclick.click()

            except Exception as e:
                print(f"Error message: {e}")
    
      def process_row(self, row):   
        numero_processo = row['Número do Processo']
        tipo_processo = row['Evento']

        if tipo_processo == 'Inicial':
            self.process_initial(numero_processo)
        elif tipo_processo == 'Petições Diversas':
            self.process_petition(numero_processo)
            

      def run(self):
        self.login()
        self.search()

        for index, row in self.df.iterrows():
            self.process_row(row)

if __name__ == '__main__':
      automation = ProcessAutomation()
      automation.run()

      input('Press Enter to exit...')
      dr.driver.quit()
      print('Driver quited successfully.')
