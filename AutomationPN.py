#IMPORTS

from re import search
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver 
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time

# ABRIR SITE

PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

driver.get("https://www.thule.com/pt-br")
driver.set_window_size(1600, 900)

# MANIPULAR JANELAS

janela = driver.window_handles
print(janela)

# FECHAR JANELA DE SEGURANÇA GOOGLE

#driver.switch_to.window(driver.window_handles[0])
#driver.close()

# VOLTAR PARA JANELA PRINCIPAL

#driver.switch_to.window(driver.window_handles[-1])

driver.find_element_by_id("onetrust-accept-btn-handler").click()

# PEGAR DADOS EXCEL

wb = load_workbook(filename='D:\VerificacaoDeProdutos.xlsx')
sh = wb['RESUMO']
n = 1
element = "teste"
erro = "alo"

for row in sh['A']:
    A1 = row.value

    driver.switch_to.window(driver.window_handles[-1])

    #VERIFICAR SE A PAGINA NÃO DEU ERRO
    try:
        erro = driver.find_element_by_xpath('/html/body/div/div/h1').text
    except NoSuchElementException:
        print("sem erro")

    if erro == "Site down for maintenance":
        print("ERRO")
        driver.refresh()

    procura = driver.find_element_by_name("search")
    procura.clear()

    try:
        procura.send_keys(A1)
    except:
        driver.close()
        wb.save("D:\VerificacaoDeProdutos.xlsx")
        break

    procura.send_keys(Keys.RETURN)

    time.sleep(3)
    #texto = driver.find_element_by_class_name("product-item").text

    #print(texto)

    try:
        element = driver.find_element_by_xpath('//*[@id="main_0_mainframed_0_noContentHitsDiv"]/div[1]/div/h2').text
    except NoSuchElementException:  
        print('Element not found') # or do something else here
        
    try:
        driver.find_element_by_xpath('//*[@id="main_0_mainframed_0_rptSearchresultGroups_productsColumn_0"]/div/div/div/div[1]/div/article')
    except NoSuchElementException:
        print(str(A1) + " " + "Duvida")
        sh['A' + str(n)] = 'Duvida'
        wb.save("D:\VerificacaoDeProdutos.xlsx")

    if element == "Não encontramos uma correspondência":
        print(str(A1) + " " + "Não encontrado")
        sh['A' + str(n)] = 'Não encontrado'
        wb.save("D:\VerificacaoDeProdutos.xlsx")

    n = n + 1

