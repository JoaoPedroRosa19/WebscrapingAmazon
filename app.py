from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
import time

#acessar o site
driver = webdriver.Chrome()
driver.get('https://www.amazon.com.br/gp/bestsellers/books/ref=sv_b_1?pf_rd_r=6F3HG2D4HC1BK5FZMJW5&pf_rd_p=07d93345-dcfb-4f12-845c-d1945b239d7d&pf_rd_m=A1ZZFT5FULY4LN&pf_rd_s=merchandised-search-2&pf_rd_t=&pf_rd_i=6740748011')
time.sleep(10)
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(5)

#Extrair Posição
posicoes = driver.find_elements(By.XPATH, "//span[@class='zg-bdg-text']")

#Extrair todos os Títulos
elementos = driver.find_elements(By.XPATH, "//div[@class='_cDEzb_p13n-sc-css-line-clamp-1_1Fn1y']")

# Pegar apenas os títulos usando índices pares
titulos = elementos[::2]

#Extrair todos os preços
precos = driver.find_elements(By.XPATH,"//span[@class='_cDEzb_p13n-sc-price_3mJ9Z']")


#Criação da planilha
workbook = openpyxl.Workbook()
workbook.create_sheet('Livros')
sheet_produtos = workbook['Livros']
sheet_produtos['A1'].value = 'Posicao'
sheet_produtos['B1'].value = 'Titulo'
sheet_produtos['C1'].value = 'Preco'


#Inserir no Excel
for posicao, titulo, preco in zip(posicoes, titulos, precos):
    sheet_produtos.append([posicao.text,titulo.text,preco.text])
workbook.save('Top50Livros.xlsx')