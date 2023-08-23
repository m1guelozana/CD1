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


def verify_protocols(Elaw, td):
    table = df.load_workbook(Elaw)
    firs_table = table.worksheets[1]
    
    for line in first_table.iter_rows():
        table_value = line[1].value
        
        if table_value == td:
            continue
        else:
            # continue...
            continue

    print(f"Realizando procedimento para o protocolo {table_value}")


def main():
    login()
    search()


main()

input('Press Enter to exit...')
dr.driver.quit()
print('Driver quited successfully. Press F5 to run again')
