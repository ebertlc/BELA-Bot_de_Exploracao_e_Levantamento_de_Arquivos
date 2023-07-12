from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import time
#import pandas as pd

print('....................................................................................')
print('....................................................................................')
print('............ B.E.L.A. - Bot de Exploração e Levantamento de Arquivos ...............')
print('....................................................................................')
print('....................................................................................')

#....................................................................................
#.............B.E.L.A. - Bot de Exploração e Levantamento de Arquivos................
#....................................................................................
print('iniciando BELA')
# Configurar o driver do Firefox
driver = webdriver.Firefox()
#caminho_dataset = 'C:\\Users\\eber_\\Documents\\delveloper\\MIDR\\aquivos\\usuarios_desabilitar.xlsx'
#usuarios = base_usuarios(caminho_dataset)

# Acessar a página de pesquisa
driver.get('https://homologacaos2id.mdr.gov.br/')

print('sistema acessado')

botao_ok = driver.find_element(By.ID, "j_idt35")
botao_ok.click()

usuario_login = driver.find_element(By.ID, 'usuario')
usuario_login.send_keys('dip@mdr.gov.br')

senha_login = driver.find_element(By.ID, 'j_idt56')
senha_login.send_keys('123456')

botao_login = driver.find_element(By.ID, 'btnEnter')
botao_login.click()

print('login efetuado')

time.sleep(1)

rec_federal = driver.find_element(By.CSS_SELECTOR, '[title="Reconhecimento Federal"]')
rec_federal.click()

time.sleep(10)
print('entrando no modulo Reconhcimento')


campo_pesquisa = driver.find_element(By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento:j_idt193:filter')
campo_pesquisa.clear()
campo_pesquisa.send_keys('PI-F-2200053-14110-20230621')
print('pesquisa adicionada')
time.sleep(2)

campo_pesquisa.send_keys(Keys.ENTER)
print('clicado enter')

time.sleep(5)

tabela = driver.find_element(By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento_data')
tabela_linha = tabela.find_element(By.CSS_SELECTOR, '[data-ri="0"]')
print('tabela encontrada')

time.sleep(1)

ActionChains(driver) \
    .double_click(tabela_linha) \
    .perform()

print('duplo click')

time.sleep(5)

sair = driver.find_element(By.ID, 'sair')
sair.click()

print('Script Execultado')