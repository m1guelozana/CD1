# Importações
from utils.webdriver_chrome import webDriver_chrome
from selenium.webdriver.common.by import By
import openpyxl as pyxl
from selenium.webdriver.common.keys import Keys

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

                dr.driver.find_element(
                  By.ID, 'j_id_2n:j_id_2r_2_d_label'
                ).click().send_keys('Acompanhamento')

                dr.driver.find_element(
                  By.ID, 'j_id_2n:txtOcorrenciaAndamento'
                ).send_keys('Cópia Inicial')

                dr.driver.find_element(
                  By.ID, 'j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label'
                ).click().send_keys('Inicial', Keys.ENTER)

                dr.driver.find_element(
                  By.ID, 'j_id_2n:j_id_2r_2_5ni:uploadGedEFile_input'
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
print('Driver quited successfully. Press F5 to run again')
