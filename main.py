from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from pywinauto import Application
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
driver.get('https://s2id.mi.gov.br/paginas/index.xhtml')

print('sistema acessado')

botao_ok = driver.find_element(By.ID, "j_idt35")
botao_ok.click()

usuario_login = driver.find_element(By.ID, 'usuario')
usuario_login.send_keys('email@dominio.br')

senha_login = driver.find_element(By.ID, 'j_idt56')
senha_login.send_keys('123456')

botao_login = driver.find_element(By.ID, 'btnEnter')
botao_login.click()

print('login efetuado')

time.sleep(5)

rec_federal = driver.find_element(By.CSS_SELECTOR, '[title="Reconhecimento Federal"]')
rec_federal.click()

print('entrando no modulo Reconhcimento')

time.sleep(10)
print('timer exedido')

protocolo = 'AL-F-2708501-13214-20230707'


campo_pesquisa = driver.find_element(By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento:j_idt123:filter')
campo_pesquisa.clear()
campo_pesquisa.send_keys(protocolo)
print('pesquisa adicionada')
time.sleep(2)

campo_pesquisa.send_keys(Keys.ENTER)
print('clicado enter')

time.sleep(5)

tabela = driver.find_element(By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento_data')
tabela_linha = tabela.find_element(By.CSS_SELECTOR, '[data-ri="0"]')
print('tabela encontrada')
time.sleep(1)
ActionChains(driver).double_click(tabela_linha).perform()
print('duplo click')

time.sleep(5)

rolar = driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

time. sleep(2)

detalhes = driver.find_element(By.ID, 'accordion:j_idt152')

arquivos_disponiveis = ''

print("procurando DMATE")
for pasta in range(0, 50):
    id = f"accordion:arquivos_disponiveis:{pasta}"
    print(f'pasta {pasta}')

    try:
        arquivo = detalhes.find_element(By.ID, id)
    except:
        continue

    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    if tag.text == 'RECONHECIMENTO':
        print(f"RECONHECIMENTO encontrado na pasta {pasta}")
        arquivos_disponiveis = id
        reconhecimento = arquivo.find_element(By.TAG_NAME, 'span')
        time.sleep(2)
        espandir = reconhecimento.find_element(By.CLASS_NAME, 'ui-tree-toggler')
        espandir.click()
        break

time. sleep(2)

print("procurando DMATE")
for pasta in range(0, 50):
    id = f"{arquivos_disponiveis}_{pasta}"
    print(f'pasta {pasta}')

    try:
        arquivo = detalhes.find_element(By.ID, id)
    except:
        continue

    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    if tag.text == 'DMATE':
        print(f"DMATE encontrado na pasta {pasta}")
        checkbox = arquivo.find_element(By.TAG_NAME, 'span')
        checkbox_cls = checkbox.find_element(By.CLASS_NAME, 'ui-chkbox-icon')
        checkbox_cls.click()
        break

print("procurando FIDE")
for pasta in range(0, 50):
    id = f"{arquivos_disponiveis}_{pasta}"
    print(f'pasta {pasta}')

    try:
        arquivo = driver.find_element(By.ID, id)
    except:
        continue

    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    if tag.text == 'FIDE':
        print(f"DMATE encontrado na pasta {pasta}")
        checkbox = arquivo.find_element(By.TAG_NAME, 'span')
        checkbox_cls = checkbox.find_element(By.CLASS_NAME, 'ui-chkbox-icon')
        checkbox_cls.click()
        break


time.sleep(5)

gerar_pdf = driver.find_element(By.ID, 'accordion:j_idt161').click()
print("Gerado o PDF")

time.sleep(6)

canva = detalhes.find_element(By.TAG_NAME, 'object')

driver.switch_to.frame(canva)
container = driver.find_element(By.ID, 'mainContainer')
toolbar = container.find_element(By.ID, 'toolbarContainer')
toolbar_viewer = toolbar.find_element(By.ID, 'toolbarViewer')
toolbar_viewer_direita = toolbar_viewer.find_element(By.ID, 'toolbarViewerRight')
download = toolbar_viewer_direita.find_element(By.ID, 'download').click()

driver.switch_to.default_content()

print('Download concluído.')

time.sleep(5)

app = Application(backend="win32").connect(title='Salvar como')

dlg = app.window(title='Salvar como')

caminho = f'C:\\Users\\eber_\\Documents\\S2iD\\Relatorios\\FIDE-DEMATE_municipio_{protocolo}_status.pdf'

# Localizar o campo "Nome do arquivo" na janela "Salvar como" pelo índice
filename_input = dlg.child_window(class_name='Edit', found_index=0)
filename_input.set_edit_text(caminho)

time.sleep(2)

filename_input.type_keys('{ENTER}')  # Digita a tecla Enter para ativar o botão "Salvar"

print('Arquivo salvo')


sair = driver.find_element(By.ID, 'sair')
sair.click()

print('Script Execultado')