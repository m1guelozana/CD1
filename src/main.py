from utils.webdriver_chrome import webDriver_chrome
from utils.xls_reader import xlsReader
from selenium.webdriver.common.by import By

dr = webDriver_chrome()
df = xlsReader()


def login():
    dr.driver.get('https://kroton.elaw.com.br/login.elaw')
    dr.driver.find_element(By.ID, 'username').send_keys('AMiceli2')
    dr.driver.find_element(By.ID, 'password').send_keys('COGN@2023')
    dr.driver.find_element(By.XPATH, '//*[@id="j_id_a_1_5_f"]').click()


def search():
    dr.driver.find_element(
      By.XPATH, '//*[@id="j_id_2d_1"]/ul/li[2]/a'
    ).click()
    dr.driver.find_element(
      By.XPATH, '//*[@id="menu-form-contencioso:j_id_2d_a_4"]/a'
    ).click()


def protocols():
    for row in df.sheet.iter_rows(min_rows=2, values_only=True):
        protocolos = row
        value = dr.driver.find_element(
          By.XPATH, "//td[contains(text(), 'Cópia de inicial')]"
        ).send_keys(protocolos)
        dr.driver.find_element(By.XPATH, '//*[@id="btnPesquisar"]').click()
        dr.driver.find_element(
          By.XPATH, '//*[@id="dtProcessoResults:0:btnProcesso"]'
        ).click()

        if value == 'Cópia Inicial':
            protocolo_enviado = True
            break
        if not protocolo_enviado:
            dr.driver.find_element(
              By.XPATH,
              '//*[@id="tabViewProcesso:j_id_i3_4_1_3_ac:j_id_i3_4_1_3_ag_1"]'
            ).click()
            dr.driver.find_element(
              By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_d_label"]'
            ).send_keys('acompanhamento')
            dr.driver.find_element(
              By.XPATH, '//*[@id="j_id_2n:txtOcorrenciaAndamento"]'
            ).send_keys('Cópia Inicial')
            dr.driver.find_element(
              By.XPATH, '//*[@id="j_id_2n:j_id_2r_2_5ni:eFileTipoCombo_label"]'
            ).send_keys('Inicial')
            dr.driver.find_element(
              By.XPATH, '//*[@id="btnSalvarAndamentoProcesso"]'
            ).click()


def main():
    login()
    search()
    protocols()


main()

input('Press Enter to exit...')
dr.driver.quit()
print('Driver quited successfully. Press F5 to run again')
