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
    last_row = df.worksheet.max_row
    ultimo_processo_verificado = None
    for row in range(1, last_row+1):
        numero_processo = str(df.worksheet.cell(row, 1).value)
        if numero_processo == ultimo_processo_verificado:
          continue

        tag_html = '<td role="gridcell">'
    
        if 'Cópia Inicial' in tag_html:
          continue
    
        print(f"Realizanod processo para o número de processo: {numero_processo}")
        ultimo_processo_verificado = numero_processo
    


def main():
    login()
    search()
    protocols()


main()

input('Press Enter to exit...')
dr.driver.quit()
print('Driver quited successfully. Press F5 to run again')
