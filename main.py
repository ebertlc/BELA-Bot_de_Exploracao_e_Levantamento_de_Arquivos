#....................................................................................
#.............B.E.L.A. - Bot de Exploração e Levantamento de Arquivos................
#....................................................................................

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from pywinauto import Application
import time
import pandas as pd

print('....................................................................................')
print('....................................................................................')
print('............ B.E.L.A. - Bot de Exploração e Levantamento de Arquivos ...............')
print('....................................................................................')
print('....................................................................................')

print('\n\n')
print('...................................Iniciando BELA...................................')
# Configurar o driver do Firefox

# Criar um objeto Options para configurar o perfil do Firefox
options = Options()

# Desativar o cache
options.set_preference("browser.cache.disk.enable", False)
options.set_preference("browser.cache.memory.enable", False)
options.set_preference("browser.cache.offline.enable", False)
options.set_preference("network.http.use-cache", False)

# Iniciar o driver do Firefox com o perfil personalizado
driver = webdriver.Firefox(options=options)

# Acessar a página de pesquisa
driver.get('https://s2id.mi.gov.br/paginas/index.xhtml')

# Fecha a Caixa e Dialogo Inicial
botao_ok = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "j_idt35")))
botao_ok.click()

print('\n\n')
print('.....................................S2iD Ativo.....................................')

# Fazer login
usuario_login = driver.find_element(By.ID, 'usuario').send_keys(Login)
senha_login = driver.find_element(By.ID, 'j_idt56').send_keys(Senha)
time.sleep(1)

botao_login = driver.find_element(By.ID, 'btnEnter').click()

print('...................................Usuario logado...................................')

# Acessar o módulo "Reconhecimento Federal"
rec_federal = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[title="Reconhecimento Federal"]')))
rec_federal.click()

print('..........................Entrando no modulo Reconhcimento..........................')

caminho_protocolos = 'C:\\Users\\eber_\\Documents\\delveloper\\MIDR\\aquivos\\protocolos_MA.xlsx'
caminho_relatorio = 'C:\\Users\\eber_\\Documents\\delveloper\\MIDR\\aquivos\\modelorelatorio_arquivosgerados.xlsx'
protocolos = pd.read_excel(caminho_protocolos)
relatorio = pd.read_excel(caminho_relatorio)

for index, row in protocolos.iterrows():
    protocolo = row['PROTOCOLO']
    municipio = row['MUNICIPIO']
    status = row['STATUS']
    data = row['DATA']

    relatorio.at[index, 'MUNICIPIO'] = municipio
    relatorio.at[index, 'DATA'] = data
    relatorio.at[index, 'PROTOCOLO'] = protocoloE

    # Iniciar o cronômetro
    start_time = time.time()

    # Pesquisar o protocolo
    campo_pesquisa = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento:j_idt123:filter')))
    campo_pesquisa.clear()
    campo_pesquisa.send_keys(protocolo)
    time.sleep(1)
    campo_pesquisa.send_keys(Keys.ENTER)

    print('\n\n')
    print('.................................Pesquisa Concluida.................................')
    time.sleep(4)

    # Localizar a tabela de processos
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento_data')))

    try:
        tabela = driver.find_element(By.ID, 'accordion:form-reconhecimento:tbl-processos-reconhecimento_data')
        tabela_linha = tabela.find_element(By.CSS_SELECTOR, '[data-ri="0"]')
        time.sleep(2)
    except NoSuchElementException:
        print("Elemento não encontrado. Reiniciando o loop.")

        relatorio.at[index, 'RESULTADO'] = 'Protocolo não localizado no S2id'
        continue

    print('................................Protocolo Localizado................................')

    # Acessar o detalhe do protocolo
    ActionChains(driver).double_click(tabela_linha).perform()

    print('................................Acessando  Protocolo................................')

    # Aguardar até o elemento onde os arquivos estão apareça
    detalhes = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'accordion:j_idt152')))

    #rolar até o fim da pagina
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    arquivos_disponiveis = ''

print('.............................Procurando  RECONHECIMENTO.............................')
# Iterar através de um range de 0 a 49 para percorrer as ids até Encontrar o pasta do Reconhecimento e expandi-la
for pasta in range(0, 50):
    id = f"accordion:arquivos_disponiveis:{pasta}"
    print(f'......................................pasta  {pasta+1}......................................')

    # Verifica se o id existe
    try:
        arquivo = detalhes.find_element(By.ID, id)
    except:
        print("RECONHECIMENTO não encontrado em nenhuma pasta.")
        for pasta in range(0, 50):
            print(f'......................................pasta  {pasta + 1}......................................')

            # Verifica se o id existe
            try:
                arquivo_if = detalhes.find_element(By.ID, 'accordion:selecionar_todos')
            except:
                continue

            # Encontrar o elemento de tag que contém o texto do nó
            subarquivo_if = detalhes.find_element(By.ID, 'accordion:detalhes_pdf')
            tag_if = subarquivo_if.find_element(By.TAG_NAME, 'label')
            # Verificar se o texto da tag é igual a 'RECONHECIMENTO'
            if tag_if.text == 'Selecionar todos':
                # Encontrar o elemento de checkbox para marcar os subelementos
                time.sleep(1)
                checkbox_if = subarquivo_if.find_element(By.TAG_NAME, 'span').click()
                #checkbox_if.find_element(By.CLASS_NAME, 'ui-chkbox-icon').click()
                # Sair do loop
                break
        else:
            print("RECONHECIMENTO não encontrado em nenhuma pasta.")

            relatorio.at[index, 'RESULTADO'] = 'Protocolo não possui arquivos'

            botao_voltar = driver.find_element(By.ID, 'j_idt26').click()
            time. sleep(2)
            continue

    # Encontrar o elemento de tag que contém o texto do nó
    subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
    tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
    # Verificar se o texto da tag é igual a 'RECONHECIMENTO'
    if tag.text == 'RECONHECIMENTO':
        print(f'........................RECONHECIMENTO Encontrado na Pasta {pasta + 1}........................')
        arquivos_disponiveis = id
        # Encontrar o elemento de expansão para mostrar os subelementos
        reconhecimento = arquivo.find_element(By.TAG_NAME, 'span')
        espandir = reconhecimento.find_element(By.CLASS_NAME, 'ui-tree-toggler').click()

        documentos = {'CAPA', 'FIDE', 'DEMATE', 'DEATE'}

        for docs in documentos:

            procurar_elemento(detalhes, arquivos_disponiveis, docs)

    time.sleep(1)

    # Gerar o PDF
    driver.find_element(By.ID, 'accordion:j_idt161').click()
    time.sleep(1)

    try:
        # Localizar o elemento "object" e mudar para o frame
        canva = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'object')))
    except:
        print('Protocolo não possui arquivos')
        relatorio.at[index, 'RESULTADO'] = 'Protocolo não possui arquivos'

        botao_voltar = driver.find_element(By.ID, 'j_idt26').click()

        continue

    print('....................................PDF  GERADO.....................................')

    time.sleep(1)

    # Mudar para o frame do objeto "object"
    driver.switch_to.frame(canva)
    # Isso permite interagir com elementos dentro do frame

    viewerContainer = driver.find_element(By.ID, 'viewerContainer')
    page = WebDriverWait(viewerContainer, 20).until(EC.presence_of_element_located((By.CLASS_NAME, 'page')))

    # Localizar o elemento de download dentro da parte direita do visualizador de ferramentas e clicar nele
    container = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'mainContainer')))
    toolbar = container.find_element(By.ID, 'toolbarContainer')
    toolbar_viewer = toolbar.find_element(By.ID, 'toolbarViewer')
    toolbar_viewer_direita = toolbar_viewer.find_element(By.ID, 'toolbarViewerRight')
    download = toolbar_viewer_direita.find_element(By.ID, 'download').click()

    # Mudar para o conteúdo padrão da página
    driver.switch_to.default_content()
    # Isso é necessário para interagir com elementos fora do frame

    time.sleep(5)
    #definir o caminho onde será salvo e o nome do aqruivo
    caminho = f'C:\\Users\\eber_\\Documents\\S2iD\\Arquivos_gerados\\FIDE-DEMATE_{municipio}_{protocolo}_{status}.pdf'

    app = Application(backend="win32").connect(title='Salvar como')
    dlg = app.window(title='Salvar como')

    # Localizar o campo "Nome do arquivo" na janela "Salvar como" pelo índice
    filename_input = dlg.child_window(class_name='Edit', found_index=0)
    filename_input.set_edit_text(caminho)
    time.sleep(2)
    filename_input.type_keys('{ENTER}')  # Digita a tecla Enter para ativar o botão "Salvar"
    time.sleep(2)

    loop_verify = 0
    while True:  # Loop infinito
        try:
            loop_verify += 1

            print(f'Verificação: {loop_verify}')
            app_verify = Application(backend='win32').connect(title='Salvar como')
            tela_verify = app_verify.window(title='Salvar como')

            filename_input_verify = tela_verify.child_window(class_name='Edit', found_index=0)
            time.sleep(2)
            filename_input_verify.type_keys('{ENTER}')
            time.sleep(2)

            continue
        except:
            break

    print('...................................Arquivo Salvo....................................')


    relatorio.at[index, 'RESULTADO'] = 'Arquivo gerado'

    # Voltar para a página de pesquisa
    botao_voltar = driver.find_element(By.ID, 'j_idt26').click()

    print('\n')
    print('............................Voltar a Pagina de Pesquisa.............................')

    end_time = time.time()
    execution_time = end_time - start_time

    # Imprimir o tempo de execução
    print(f"Tempo de execução: {execution_time} segundos")

relatorio.to_excel('C:\\Users\\eber_\\Documents\\S2iD\\Arquivos_gerados\\FIDE-DEMATE_Relatorio.xlsx', index=False)

# Deslogar do sistema
time.sleep(10)
sair = driver.find_element(By.ID, 'sair').click()

driver.close()

print('\n\n')
print('.................................Script Execultado..................................')

def procurar_elemento(driver, arquivos_disponiveis, documento_procurado):

    print(f'.................................Procurando  {documento_procurado}..................................')

    for pasta in range(0, 50):

        id = f"{arquivos_disponiveis}_{pasta}"

        print(f'......................................pasta  {pasta+1}......................................')

        # Verifica se o id existe
        try:
            arquivo = driver.find_element(By.ID, id)
        except:
            break

        # Encontrar o elemento de tag que contém o texto do nó
        subarquivo = arquivo.find_element(By.TAG_NAME, 'span')
        tag = subarquivo.find_element(By.CLASS_NAME, 'ui-treenode-label')
        # Verificar se o texto da tag é igual a 'FIDE'
        if tag.text == documento_procurado:
            print(f'............................{documento_procurado} Encontrado na Pasta {pasta+1}..............................')
            # Encontrar o elemento de checkbox para marcar os subelementos
            checkbox = arquivo.find_element(By.TAG_NAME, 'span')
            checkbox_cls = checkbox.find_element(By.CLASS_NAME, 'ui-chkbox-icon')
            checkbox_cls.click()
            # Sair do loop
            break
