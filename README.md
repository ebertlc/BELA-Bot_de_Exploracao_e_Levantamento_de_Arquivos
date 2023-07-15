# B.E.L.A. - Bot de Exploração e Levantamento de Arquivos

Este é um bot de automação escrito em Python que utiliza a biblioteca Selenium para explorar e coletar arquivos de um sistema web chamado S2iD. O objetivo deste bot é fazer login no sistema, pesquisar um protocolo específico e obter arquivos relacionados a esse protocolo.

## Pré-requisitos

Antes de executar este bot, certifique-se de ter as seguintes dependências instaladas:

- [Python](https://www.python.org) (versão 3.6 ou superior)
- [Selenium](https://selenium-python.readthedocs.io) (biblioteca Python para automação de navegador)
- [Pywinauto](https://pywinauto.readthedocs.io) (biblioteca Python para automação de aplicativos do Windows)

Além disso, o bot foi desenvolvido para funcionar com o navegador Firefox. Portanto, certifique-se de ter o navegador Firefox instalado em seu sistema.

## Configuração

Antes de executar o bot, é necessário configurar as seguintes informações:

- `Login`: substitua esta variável pelo seu endereço de e-mail de login no sistema S2iD.
- `Senha`: substitua esta variável pela sua senha de login no sistema S2iD.

```python
Login = 'seu_email@example.com'
Senha = 'sua_senha_secreta'
```

Além disso, se você deseja salvar o arquivo PDF gerado em um local específico, defina o caminho desejado na variável `caminho`:

```python
caminho = 'C:\\caminho\\para\\salvar\\arquivo.pdf'
```

## Executando o bot

Após a configuração, você pode executar o bot executando o arquivo Python. Certifique-se de que todas as dependências estejam instaladas e que o navegador Firefox esteja em execução.

O bot irá abrir o navegador Firefox, fazer login no sistema S2iD, pesquisar o protocolo especificado, obter os arquivos relacionados a esse protocolo, gerar um arquivo PDF e salvá-lo no caminho especificado. Em seguida, o bot voltará à página de pesquisa, deslogará do sistema e encerrará a execução.

## Licença

Este projeto está licenciado sob a Licença MIT. Consulte o arquivo LICENSE para obter mais informações.
