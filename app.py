import dados
from time import sleep

login_usuario = None
senha_usuario = None
assinatura = None
com_copia = None
diretorio = None
conteudo = None
nome_do_arquivo = None
assunto = None
relacao_a_enviar = []
email_enviados = []
arquivo = 'C:\\Users\\rodri\\OneDrive\\Área de Trabalho\\'
menu_opcoes = ['Enviar Email', 'Configurar Email',
               'Ver lista à enviar', 'Ver lista de enviados', 'Diretório do Excel', 'Sair']

while True:

    resposta = dados.menu(menu_opcoes)

    # Opção para enviar Email
    if resposta == 1:
        dados.titulo('Enviando Email')

        # validações de informações importantes que precisam ser preenchidas
        if login_usuario is None and senha_usuario is None:
            print('\033[31mVerificar Configuração de Email.\033[m')
            continue
        if diretorio is None:
            print('\033[31mVerificar Diretório do Excel.\033[m')
            continue
        if conteudo is None:
            print('\033[31mVerificar Configuração de Email.\033[m')

        # Cria um lista com dados importados da planilha
        email_enviados = []
        relacao_a_enviar = dados.buscarnoArquivo(nome_do_arquivo)
        for pos, valor in enumerate(relacao_a_enviar):
            if pos == 0:
                continue
            assunto = f'{valor[1]} - {valor[0]}'
            nome = valor[0]
            email = valor[3]
            anexo = valor[4]
            email_enviados.append(dados.enviarEmail(login_usuario, senha_usuario, assunto,
                                                    email, com_copia, conteudo, anexo, assinatura, nome))
            assunto = ''
            sleep(0.2)

    # Opção para configurar Email
    elif resposta == 2:
        # Informações Sobre o Login do email Remetente
        dados.titulo('Login')
        login_usuario = input('Usuário: ')
        senha_usuario = input('Senha: ')

        # Informações que serão repetidas em todos os emails
        dados.titulo('Conteúdo do email. Deixe vazio se preferir.')
        conteudo_unificado = ''
        for valor in range(0, 2):
            conteudo_personalizado = input(
                f'Digite o Conteúdo do email({valor + 1}° parágrafo): ')
            conteudo_unificado += f'{conteudo_personalizado}\n'
        dados.linha()
        conteudo = f'Olá, {dados.saudacao()}!\n{conteudo_unificado}'

        dados.titulo('Assinatura e CC(com cópia)')
        assinatura = input('Conteúdo da Assinatura: ')
        com_copia = input('Copiar Alguem? ')
        if not com_copia:
            com_copia = None

    # Ver lista de itens a enviar
    elif resposta == 3:
        if diretorio is None:
            print('\033[31mVerificar Diretório do Excel.\033[m')
            continue
        relacao_a_enviar = dados.buscarnoArquivo(nome_do_arquivo)
        dados.imprimirLista(relacao_a_enviar)

    # Ver lista de enviados
    elif resposta == 4:
        if email_enviados == []:
            print('\033[31mAinda não foram enviados email.\033[m')
            continue
        while True:
            resposta = dados.questionSimNao('Imprimir Arquivo?')
            if resposta == 'S':
                dados.salvanoArquivo(email_enviados)
                break
            print(f'{"NOME":>33}{"ASSUNTO":>20}{"EMAIL":>50}{"STATUS":>20}{"OBS":>30}')
            for pos, valor in enumerate(email_enviados):
                print(
                    f'{pos + 1:<3}{valor["nome"]:>30}{valor["assunto"]:>20}{valor["email"]:>50}{valor["status"]:>20}{valor["obs"]:>30}')
        dados.linha()

    # Buscar caminho do arquivo
    elif resposta == 5:
        diretorio = None
        dados.titulo('Local do arquivo')
        diretorio = input('Informe o diretório: ')
        nome_do_arquivo = f'{arquivo}{diretorio}'
        if not dados.inserirdiretorio(nome_do_arquivo):
            diretorio = None

    # Opção para sair
    elif resposta == len(menu_opcoes):
        dados.titulo('Até logo')
        break
