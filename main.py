import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import numpy
import openpyxl

options = webdriver.ChromeOptions()
options.add_argument("--headless=new")

bf = pd.read_excel("Produtos.xlsx")
array = numpy.asarray(bf['Produtos'])
arrayFinal = []


def buscarNome (produtos):
    contagem = 0

    
    driver = webdriver.Chrome(options=options)
    driver.get("http://www.google.com.br")

    book = openpyxl.Workbook()
    book.create_sheet("Produtos com preço")
    page = book["Produtos com preço"]
    
    for i in produtos:
        arProd = []
        elem = driver.find_element(By.NAME, "q")
        elem.clear()
        elem.send_keys(produtos[contagem])
        elem.send_keys(Keys.RETURN)
        try:
            name=driver.find_element(By.CLASS_NAME,"RnJeZd.top.pla-unit-title")
            arProd.append(name.text)
        except:
            arProd.append(produtos[contagem] + "---NÃO ENCONTRADO---")
        try:
            name=driver.find_element(By.CLASS_NAME,"e10twf.ONTJqd")
            arProd.append(name.text)
        except :
            name=driver.find_element(By.CLASS_NAME,"T4OwTb")
            arProd.append(name.text)
        name=driver.find_element(By.ID,"vplap0")
        url = name.get_attribute('href')
        arProd.append(str(url))
        print(arProd)
        page.append(arProd)
        contagem = contagem + 1
    return book.save("ProdOrçado.xlsx")

buscarNome(array)
