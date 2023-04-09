import dados
from time import sleep

login_usuario = None
senha_usuario = None
relacao_a_enviar = []
diretorio = None
arquivo = 'C:\\Users\\rodri\\OneDrive\\Área de Trabalho\\'
nome_do_arquivo = None
menu_opcoes = ['Enviar Email', 'Configurar Email',
               'Ver lista à enviar', 'Diretório do Excel', 'Sair']
assunto = ''
conteudo = f'Olá, {dados.saudacao()}!\nEste é um teste de envio de email.'

while True:

    resposta = dados.menu(menu_opcoes)

    # Opção para enviar Email
    if resposta == 1:
        if login_usuario is None and senha_usuario is None:
            print('\033[31mVerificar Configuração de Email.\033[m')
            continue
        if diretorio is None:
            print('\033[31mVerificar Diretório do Excel.\033[m')
            continue
        relacao_a_enviar = dados.buscarnoArquivo(nome_do_arquivo)
        for pos, valor in enumerate(relacao_a_enviar):
            if pos == 0:
                continue
            assunto = f'{valor[1]} - {valor[0]}'
            email = valor[3]
            anexo = valor[4]
            dados.enviarEmail(login_usuario, senha_usuario, assunto,
                              email, conteudo, anexo)
            assunto = ''
            sleep(0.3)

    # Opção para configurar Email
    elif resposta == 2:
        dados.titulo('Login')
        login_usuario = input('Usuário: ')
        senha_usuario = input('Senha: ')

    # Ver lista de itens a enviar
    elif resposta == 3:
        if diretorio is None:
            print('\033[31mVerificar Diretório do Excel.\033[m')
            continue
        relacao_a_enviar = dados.buscarnoArquivo(nome_do_arquivo)
        dados.imprimirLista(relacao_a_enviar)

    # Buscar caminho do arquivo
    elif resposta == 4:
        diretorio = input('Informe o diretório: ')
        nome_do_arquivo = f'{arquivo}{diretorio}'
        if not dados.inserirdiretorio(nome_do_arquivo):
            diretorio = None

    # Opção para sair
    elif resposta == len(menu_opcoes):
        dados.titulo('Até logo')
        break
