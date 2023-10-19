import pandas as pd
import time
import pyautogui
from openpyxl import Workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.webdriver_chrome import webDriver_chrome
from selenium.webdriver.common.action_chains import ActionChains

# Inicializar o driver do navegador
dr = webDriver_chrome()


class ProcessAutomation:
    def __init__(self):
        self.df = None

    def login(self):
        # Realizar login
        dr.driver.get('https://kroton.elaw.com.br/login.elaw')
        dr.driver.find_element(By.ID, 'username').send_keys('AMiceli2')
        dr.driver.find_element(By.ID, 'password').send_keys('COGN@2023')
        dr.driver.find_element(By.XPATH, '//*[@id="j_id_a_1_5_f"]').click()

    def search(self):        
        try:
            # Aguarde até que o elemento de menu esteja visível e interagível
            marretadothor= WebDriverWait(dr.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="j_id_2d_1"]/ul/li[2]/a'))
            )
            marretadothor.click()

            menu_contencioso = WebDriverWait(dr.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="menu-form-contencioso:j_id_2d_a_4"]/a'))
            )
            menu_contencioso.click()
        except Exception as e:
            print(f"Erro ao clicar no elemento de menu: {str(e)}")


    def process_row(self):
        # Leitura da tabela, percorrendo os números de processo
        self.df = pd.read_excel('Pasta1.xlsx')
        
        if 'Número do Processo' in self.df.columns.astype(str):
            numero_processos = self.df.loc[self.df['Número do Processo'] == 'Número do Processo'].iloc[1:]
            for index, row in self.df.iterrows():                          
                
                try:
                    numero_processo = row['Número do Processo']
                    print(f"Processo: {numero_processo}")
                    # Aguarde até que o campo de busca esteja disponível
                    campo_busca = WebDriverWait(dr.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'))
                    )
                    campo_busca.clear()
                    campo_busca.send_keys(str(numero_processo))

                    dr.driver.find_element(By.XPATH, '//*[@id="btnPesquisar"]').click()
                    dr.driver.find_element(By.ID, 'dtProcessoResults:0:btnProcesso').click()
                    
                    try:
                        WebDriverWait(dr.driver, 30).until(EC.presence_of_element_located((                            
                                By.XPATH, '''
                                //*[@id="tabViewProcesso:j_id_i3_4_1_3_ac:j_id_i3_4_1_3_ag_1"]
                                '''
                            ))).click()
                    except:
                        waitoclick = WebDriverWait(dr.driver, 20)
                        bttnclick = waitoclick.until(EC.presence_of_element_located((
                        By.XPATH, '''
                        //*[@id="btnVoltar"]/span[1]
                        '''
                            )))
                        bttnclick.click()                            
                        time.sleep(7)
                        pass 

                    tipo_processo = row['Evento']
                    if tipo_processo == 'Inicial':
                        try:
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

                            # Verifica se a coluna "Documentos" não está vazia
                            if not pd.isna(row['Documentos']) and row['Documentos'] != '':
                                elementar = wait.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                                '''
                                )))
                                elementar.click()
                                
                                
                                elementwrite = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//input[contains(@id,'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter')]"
                                )))
                                elementwrite.send_keys(str(row['Evento']))
                                time.sleep(0.5)
                                pyautogui.press('enter')
                                time.sleep(5)
                                
                                
                                dr.driver.find_element(
                                    By.XPATH, "//span[contains(@id,'label')]"
                                ).click()
                                
                                nome_arquivo = self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                                time.sleep(4)
                                pyautogui.write(str(row['Documentos']))
                                time.sleep(0.5)
                                pyautogui.hotkey('alt', 'a')
                                time.sleep(7)

                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                )))
                                bttnclick.click()                            
                                time.sleep(7)
                            else:
                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                 )))
                                bttnclick.click()                            
                                time.sleep(7)
                                pass 
                            
                        except Exception as e:
                            # Este bloco será executado se ocorrer algum erro durante a busca ou clique
                            print(f"Erro ao processar o número de processo: {numero_processo}")
                            print(f"Mensagem de erro: {str(e)}")
                            dr.driver.find_element(By.XPATH, '''
                                //button[contains(@id,'btnVoltar')]
                            ''').click()
                            dr.driver.find_element(
                                By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                            ).clear()
                            pass
                        
                        
                        
                    elif tipo_processo == 'Petições Diversas':
                        try:
                            wait = WebDriverWait(dr.driver, 10)
                            element = wait.until(EC.presence_of_element_located((
                            By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]'
                            )))
                            element.click()
                            
                            dr.driver.find_element(
                                By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                            ).send_keys('Acompanhamento')                        
                            time.sleep(3)
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
                            time.sleep(5)
                            pyautogui.press('tab')
                            time.sleep(5)                     

                            # Verifica se a coluna "Documentos" não está vazia
                            if not pd.isna(row['Documentos']) and row['Documentos'] != '':                                                        
                                elementar = wait.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                                '''
                                )))
                                elementar.click()
                                
                                
                                elementwrite = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//input[contains(@id,'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter')]"
                                )))
                                elementwrite.send_keys(str(row['Evento']))
                                time.sleep(0.5)
                                pyautogui.press('enter')
                                time.sleep(5)
                                
                                
                                dr.driver.find_element(
                                    By.XPATH, "//span[contains(@id,'label')]"
                                ).click()
                                
                                nome_arquivo = self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                                time.sleep(4)
                                pyautogui.write(str(row['Documentos']))
                                time.sleep(0.5)
                                pyautogui.hotkey('alt', 'a')
                                time.sleep(7)

                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                )))
                                bttnclick.click()                            
                                time.sleep(7)
                            else:
                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                 )))
                                bttnclick.click()                            
                                time.sleep(7)
                                pass  

                        except Exception as e:
                            # Este bloco será executado se ocorrer algum erro durante a busca ou clique
                            print(f"Erro ao processar o número de processo: {numero_processo}")
                            print(f"Mensagem de erro: {str(e)}")
                            dr.driver.find_element(By.XPATH, '''
                                //button[contains(@id,'btnVoltar')]
                            ''').click()
                            dr.driver.find_element(
                                By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                            ).clear()
                            pass
                    elif tipo_processo == 'Contestação':
                        try:
                            wait = WebDriverWait(dr.driver, 10)
                            element = wait.until(EC.presence_of_element_located((
                            By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]'
                            )))
                            element.click()
                            
                            dr.driver.find_element(
                                By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                            ).send_keys('Acompanhamento')                        
                            time.sleep(3)
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
                            time.sleep(5)
                            pyautogui.press('tab')
                            time.sleep(5)
                           

                            # Verifica se a coluna "Documentos" não está vazia
                            if not pd.isna(row['Documentos']) and row['Documentos'] != '':                                                        
                                elementar = wait.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                                '''
                                )))
                                elementar.click()
                                
                                
                                elementwrite = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//input[contains(@id,'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter')]"
                                )))
                                elementwrite.send_keys(str(row['Evento']))
                                time.sleep(0.5)
                                pyautogui.press('enter')
                                time.sleep(5)
                                
                                
                                dr.driver.find_element(
                                    By.XPATH, "//span[contains(@id,'label')]"
                                ).click()
                                
                                nome_arquivo = self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                                time.sleep(4)
                                pyautogui.write(str(row['Documentos']))
                                time.sleep(0.5)
                                pyautogui.hotkey('alt', 'a')
                                time.sleep(7)

                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                    )))
                                bttnclick.click()                            
                                time.sleep(7) 
                            else:
                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                 )))
                                bttnclick.click()                            
                                time.sleep(7)
                                pass      

                        except Exception as e:
                            # Este bloco será executado se ocorrer algum erro durante a busca ou clique
                            print(f"Erro ao processar o número de processo: {numero_processo}")
                            print(f"Mensagem de erro: {str(e)}")
                            dr.driver.find_element(By.XPATH, '''
                                //button[contains(@id,'btnVoltar')]
                            ''').click()
                            dr.driver.find_element(
                                By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                            ).clear()
                            pass               
                    
                    elif tipo_processo == 'Entregar ata da audiência e Prazos' or 'Audiência Realizada':
                        try:                            
                            wait = WebDriverWait(dr.driver, 10)
                            element = wait.until(EC.presence_of_element_located((
                            By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]'
                            )))
                            element.click()
                            
                            dr.driver.find_element(
                                By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                            ).send_keys('Audiência Realizada')                        
                            time.sleep(3)
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
                            time.sleep(5)
                            pyautogui.press('tab')
                            time.sleep(5)
                           



                            # Verifica se a coluna "Documentos" não está vazia
                            if not pd.isna(row['Documentos']) and row['Documentos'] != '':
                                                        
                                elementar = wait.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                                '''
                                )))
                                elementar.click()
                                
                                
                                elementwrite = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//input[contains(@id,'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter')]"
                                )))
                                elementwrite.send_keys('Ata de Audiência')
                                time.sleep(0.5)
                                pyautogui.press('enter')
                                time.sleep(5)
                                
                                
                                dr.driver.find_element(
                                    By.XPATH, "//span[contains(@id,'label')]"
                                ).click()
                                
                                nome_arquivo = self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                                time.sleep(4)
                                pyautogui.write(str(row['Documentos']))
                                time.sleep(0.5)
                                pyautogui.hotkey('alt', 'a')
                                time.sleep(7)

                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                    )))
                                bttnclick.click()                            
                                time.sleep(7)
                            else:
                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                 )))
                                bttnclick.click()                            
                                time.sleep(7)
                                pass       

                        except Exception as e:
                            # Este bloco será executado se ocorrer algum erro durante a busca ou clique
                            print(f"Erro ao processar o número de processo: {numero_processo}")
                            print(f"Mensagem de erro: {str(e)}")
                            dr.driver.find_element(By.XPATH, '''
                                //button[contains(@id,'btnVoltar')]
                            ''').click()
                            dr.driver.find_element(
                                By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                            ).clear()
                            pass              
                    
                    elif tipo_processo == 'Documentos':
                        try:                            
                            wait = WebDriverWait(dr.driver, 10)
                            element = wait.until(EC.presence_of_element_located((
                            By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]'
                            )))
                            element.click()
                            
                            dr.driver.find_element(
                                By.XPATH, "//input[@id='j_id_2n:j_id_2r_2_d_filter']"
                            ).send_keys('Acompanhamento')                        
                            time.sleep(3)
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
                            time.sleep(5)
                            pyautogui.press('tab')
                            time.sleep(5)
                           



                            # Verifica se a coluna "Documentos" não está vazia
                            if not pd.isna(row['Documentos']) and row['Documentos'] != '':
                                
                                elementar = wait.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //label[@id='j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label']
                                '''
                                )))
                                elementar.click()
                                
                                
                                elementwrite = wait.until(EC.presence_of_element_located((
                                By.XPATH, "//input[contains(@id,'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_filter')]"
                                )))
                                elementwrite.send_keys('Documentos')
                                time.sleep(2)
                                # Pressione a tecla para baixo cinco vezes
                                pyautogui.press('down', presses=4)
                                pyautogui.press('enter')
                                time.sleep(5)
                                
                                
                                dr.driver.find_element(
                                    By.XPATH, "//span[contains(@id,'label')]"
                                ).click()
                                
                                nome_arquivo = self.df.loc[self.df['Documentos'] == 'Documentos'].iloc[1:]
                                time.sleep(4)
                                pyautogui.write(str(row['Documentos']))
                                time.sleep(0.5)
                                pyautogui.hotkey('alt', 'o')
                                time.sleep(7)

                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                 ).click()

                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                 )))
                                bttnclick.click()   



                                # waitonclick = WebDriverWait(dr.driver, 20)
                                # bttnclick = waitonclick.until(EC.presence_of_element_located((
                                #     By.XPATH, "//button[contains(@id,'btnVoltar')]"
                                # )))
                                # bttnclick.click()
                                time.sleep(7)  

                            else:
                                dr.driver.find_element(
                                By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
                                ).click()
                                time.sleep(7)
                                waitoclick = WebDriverWait(dr.driver, 20)
                                bttnclick = waitoclick.until(EC.presence_of_element_located((
                                By.XPATH, '''
                                //*[@id="btnVoltar"]/span[1]
                                '''
                                 )))
                                bttnclick.click()                            
                                time.sleep(7)
                                pass

                        except Exception as e:
                        # Este bloco será executado se ocorrer algum erro durante a busca ou clique
                            print(f"Erro ao processar o número de processo: {numero_processo}")
                            print(f"Mensagem de erro: {str(e)}")
                            dr.driver.find_element(By.XPATH, '''
                                //button[contains(@id,'btnVoltar')]
                            ''').click()
                            dr.driver.find_element(
                                By.XPATH, '//*[@id="tabSearchTab:txtSearch"]'
                            ).clear()
                            pass      
                                  
                                   
                    
                    # Exclua a linha no DataFrame original
                    self.df.drop(index, inplace=True)
                    time.sleep(5)

                    # Salva a planilha atualizada
                    self.df.to_excel("pasta1.xlsx", index=False)
                    # Voltar para a página de pesquisa
                    dr.driver.find_element(By.XPATH, "//button[contains(@id,'btnVoltar')]").click()
                    campo_busca = dr.driver.find_element(By.XPATH, '//*[@id="tabSearchTab:txtSearch"]')
                    campo_busca.clear()
                

                except Exception as e:
                    print(f"Erro ao processar o número de processo: {numero_processo}")
                    print(f"Mensagem de erro: {str(e)}")
                    self.dr.driver.quit()
            return        
                

   
    def run(self):
        while True:
            try:
                self.login()
                self.search()
                self.process_row()
                print('Código concluído.')
                break  # Saia do loop quando o código for concluído com sucesso
            except Exception as e:
                print(f'Ocorreu um erro: {str(e)}')
                print('Reiniciando o site...')
                self.dr.driver.quit()
                

if __name__ == '__main__':
    while True:
        automation = ProcessAutomation()
        automation.run()

        input('Pressione Enter para sair...')
        automation.dr.driver.quit()
        print('Driver encerrado com sucesso.')