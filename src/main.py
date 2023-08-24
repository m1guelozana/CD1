# Importações
import openpyxl as pyxl
import time
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.webdriver_chrome import webDriver_chrome

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
    workbook = pyxl.load_workbook('Elaw.xlsx')
    worksheet = workbook.active
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        numero_processo = row[0]
        try:
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

            tag_html = dr.driver.find_element(
              By.XPATH, '''
              //*[@id="tabViewProcesso:j_id_i3_4_1_3_ac:dtProcessoAndamentoResults_data"]/tr[1]/td[5]
              '''
            )
            valor_tag = tag_html.get_attribute('value')

            if valor_tag != 'Cópia Inicial':
                print(f"Processo: {numero_processo}")
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
                time.sleep(2)
                pyautogui.press('enter')
                pyautogui.press('tab')
                time.sleep(0.5)

                wait = WebDriverWait(dr.driver, 10)
                text = wait.until(EC.presence_of_element_located((
                  By.XPATH, '''
                  //textarea[contains(@id,'j_id_2n:txtOcorrenciaAndamento')]
                  '''
                )))
                text.send_keys('Cópia Inicial')
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(0.5)

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
                time.sleep(2)
                pyautogui.press('enter')
                time.sleep(0.5)

                dr.driver.find_element(
                  By.XPATH, "//span[contains(@role,'button')]"
                ).click()

                wait = WebDriverWait(dr.driver, 10)
                upload_element = wait.until(EC.presence_of_element_located((
                  By.XPATH, '''
                  //*[@id='j_id_2n:j_id_2r_2_5ni:uploadGedEFile_input']
                  '''
                )))

                file_name = row[4]
                upload_element.send_keys(file_name)

                dr.driver.find_element(
                  By.XPATH, "//span[contains(@role,'button')]"
                ).click()
                pass
            else:
                continue

        except Exception as e:
            print(f"Erro ao processar o número de processo: {numero_processo}")
            print(f"Mensagem de erro: {str(e)}")
            break


def main():
    login()
    search()
    workbook_reading()


if __name__ == '__main__':
    main()

input('Press Enter to exit...')
dr.driver.quit()
print('Driver quited successfully.')
