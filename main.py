from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
print('\n\n...................................Iniciando BELA...................................')
# Configurar o driver do Firefox
driver = webdriver.Firefox()
#caminho_dataset = 'C:\\Users\\eber_\\Documents\\delveloper\\MIDR\\aquivos\\usuarios_desabilitar.xlsx'
#usuarios = base_usuarios(caminho_dataset)

# Acessar a página de pesquisa
driver.get('https://s2id.mi.gov.br/paginas/index.xhtml')

botao_ok = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "j_idt35")))
botao_ok.click()

print('\n\n.....................................S2iD Ativo.....................................')



usuario_login = driver.find_element(By.ID, 'usuario').send_keys(Login)
senha_login = driver.find_element(By.ID, 'j_idt56').send_keys(Senha)
time.sleep(1)

botao_login = driver.find_element(By.ID, 'btnEnter').click()

print('\n...................................Usuario logado...................................')

time.sleep(5)

rec_federal = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[title="Reconhecimento Federal"]')))
rec_federal.click()

print('\n..........................Entrando no modulo Reconhcimento..........................')

# Iniciar o cronômetro
start_time = time.time()

protocolo = 'AL-F-2708501-13214-20230707'

campo_pesquisa = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento:j_idt123:filter')))
campo_pesquisa.clear()
campo_pesquisa.send_keys(protocolo)
time.sleep(1)
campo_pesquisa.send_keys(Keys.ENTER)

print('\n.................................Pesquisa Concluida.................................')
time.sleep(2)

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento_data')))

tabela = driver.find_element(By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento_data')
tabela_linha = tabela.find_element(By.CSS_SELECTOR, '[data-ri="0"]')
time.sleep(1)

print('\n................................Protocolo Localizado................................')

ActionChains(driver).double_click(tabela_linha).perform()

print('\n................................Acessando  Protocolo................................')

detalhes = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accordion:j_idt152')))

#rolar até o fim da pagina
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

arquivos_disponiveis = ''


print('\n.............................Procurando  RECONHECIMENTO.............................')
for pasta in range(0, 50):
    id = f"accordion:arquivos_disponiveis:{pasta}"
    print(f'......................................pasta  {pasta+1}......................................')

    try:
        arquivo = detalhes.find_element(By.ID, id)
    except:
        continue

    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    if tag.text == 'RECONHECIMENTO':
        print(f'\n........................RECONHECIMENTO Encontrado na Pasta {pasta + 1}........................')
        arquivos_disponiveis = id
        reconhecimento = arquivo.find_element(By.TAG_NAME, 'span')
        espandir = reconhecimento.find_element(By.CLASS_NAME, 'ui-tree-toggler')
        espandir.click()
        break
time. sleep(2)

print('\n.................................Procurando  DMATE..................................')
for pasta in range(0, 50):
    id = f"{arquivos_disponiveis}_{pasta}"
    print(f'......................................pasta  {pasta+1}......................................')

    try:
        arquivo = detalhes.find_element(By.ID, id)
    except:
        continue

    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    if tag.text == 'DMATE':
        print(f'............................DMATE Encontrado na Pasta {pasta+1}.............................')
        checkbox = arquivo.find_element(By.TAG_NAME, 'span')
        checkbox_cls = checkbox.find_element(By.CLASS_NAME, 'ui-chkbox-icon')
        checkbox_cls.click()
        break

print('\n..................................Procurando FIDE...................................')
for pasta in range(0, 50):
    id = f"{arquivos_disponiveis}_{pasta}"
    print(f'......................................pasta  {pasta+1}......................................')

    try:
        arquivo = driver.find_element(By.ID, id)
    except:
        continue

    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    if tag.text == 'FIDE':
        print(f'............................FIDE Encontrado na Pasta {pasta+1}..............................')
        checkbox = arquivo.find_element(By.TAG_NAME, 'span')
        checkbox_cls = checkbox.find_element(By.CLASS_NAME, 'ui-chkbox-icon')
        checkbox_cls.click()
        break

time.sleep(1)

driver.find_element(By.ID, 'accordion:j_idt161').click()
time.sleep(1)

canva = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'object')))

print('\n....................................PDF  GERADO.....................................')

time.sleep(1)

driver.switch_to.frame(canva)

container = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'mainContainer')))
toolbar = container.find_element(By.ID, 'toolbarContainer')
toolbar_viewer = toolbar.find_element(By.ID, 'toolbarViewer')
toolbar_viewer_direita = toolbar_viewer.find_element(By.ID, 'toolbarViewerRight')
download = toolbar_viewer_direita.find_element(By.ID, 'download').click()

driver.switch_to.default_content()

time.sleep(3)

caminho = f'C:\\Users\\eber_\\Documents\\S2iD\\teste\\FIDE-DEMATE_municipio_{protocolo}_status3.pdf'

app = Application(backend="win32").connect(title='Salvar como')
dlg = app.window(title='Salvar como')

# Localizar o campo "Nome do arquivo" na janela "Salvar como" pelo índice
filename_input = dlg.child_window(class_name='Edit', found_index=0)
filename_input.set_edit_text(caminho)
time.sleep(1)
filename_input.type_keys('{ENTER}')  # Digita a tecla Enter para ativar o botão "Salvar"

print('\n...................................Arquivo Salvo....................................')

botao_voltar = driver.find_element(By.ID, 'j_idt26').click()

print('\n............................Voltar a Pagina de Pesquisa.............................')

end_time = time.time()
execution_time = end_time - start_time

# Imprimir o tempo de execução
print(f"Tempo de execução: {execution_time} segundos")

time.sleep(10)
sair = driver.find_element(By.ID, 'sair').click()

print('\n.................................Script Execultado..................................')