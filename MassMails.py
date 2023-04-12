import dados
from time import sleep
<<<<<<< HEAD
from tkinter import *


=======
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e

login_usuario = None
senha_usuario = None
assinatura = None
com_copia = None
diretorio = None
conteudo = None
assunto = None
diretorio_arquivo = None
diretorio_pastas = None
relacao_a_enviar = []
email_enviados = []
menu_opcoes = ['Enviar E-mails', 'Configurar Email',
               'Lista de envios à realizar', 'Impressão de lista de e-mails enviados', 'Informar Diretórios', 'Instruções', 'Sair']

while True:

    resposta = dados.menu(menu_opcoes)

    # Opção para enviar Email
    if resposta == 1:
        dados.titulo('Enviando Email')

        # validações de informações importantes que precisam ser preenchidas
        if login_usuario is None and senha_usuario is None:
            print('\033[31mVerificar Configuração de Email.\033[m')
            continue
        if diretorio_arquivo is None:
            print('\033[31mVerificar Diretório do Excel.\033[m')
            continue
        if conteudo is None:
            print('\033[31mVerificar Configuração de Email.\033[m')
            continue

        # Cria um lista com dados importados da planilha
        email_enviados = []
        relacao_a_enviar = dados.buscarnoArquivo(diretorio_arquivo)
        for pos, valor in enumerate(relacao_a_enviar):
            if pos == 0:
                continue
            assunto = f'{valor[1]} - {valor[0]}'
            nome = valor[0]
            email = valor[2]
            anexo = f'{diretorio_pastas}\\{nome}'
<<<<<<< HEAD
=======
<<<<<<< HEAD
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e
            retorno_envio = dados.enviarEmail(login_usuario, senha_usuario, assunto,
                                                    email, com_copia, conteudo, anexo, assinatura, nome)
            if retorno_envio == 'Erro':
                break
            email_enviados.append(retorno_envio)
<<<<<<< HEAD
=======
=======
            email_enviados.append(dados.enviarEmail(login_usuario, senha_usuario, assunto,
                                                    email, com_copia, conteudo, anexo, assinatura, nome))
>>>>>>> 4d6802f14ad209652456edfd0ce8bb7f8e4d694a
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e
            assunto = ''
            sleep(0.2)
        print('\n\n\n\033[32mProcesso de envio concluído\033[m')
        sleep(2)

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
<<<<<<< HEAD
        nome_assinatura = input('Nome da Assinatura: ')
        cargo_assinatura = input('Cargo/Empresa da Assinatura: ')
        contato_assinatura = input('Contato da Assinatura: ')
        assinatura = f'{nome_assinatura}\n{cargo_assinatura}\n{contato_assinatura}'
=======
        assinatura = input('Conteúdo da Assinatura: ')
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e
        com_copia = input('Copiar Alguem? ')
        if not com_copia:
            com_copia = None

    # Ver lista de itens a enviar
    elif resposta == 3:
        if diretorio_arquivo is None:
            print('\033[31mVerificar Diretório do Excel.\033[m')
            continue
        relacao_a_enviar = dados.buscarnoArquivo(diretorio_arquivo)
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
<<<<<<< HEAD
=======
<<<<<<< HEAD
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e
            print(
                f'{"NOME":<36}{"ASSUNTO":<40}{"EMAIL":<50}{"STATUS":<20}')
            for pos, valor in enumerate(email_enviados):
                print(
                    f'{pos + 1:<3} - {valor["nome"]:<30}{valor["assunto"]:<40}{valor["email"]:<50}{valor["status"]:<20}')
            break
<<<<<<< HEAD
=======
=======
            else:
                print(
                    f'{"NOME":<36}{"ASSUNTO":<40}{"EMAIL":<50}{"STATUS":<20}')
                for pos, valor in enumerate(email_enviados):
                    print(
                        f'{pos + 1:<3} - {valor["nome"]:<30}{valor["assunto"]:<40}{valor["email"]:<50}{valor["status"]:<20}')
                break
>>>>>>> 4d6802f14ad209652456edfd0ce8bb7f8e4d694a
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e
        dados.linha()

    # Buscar caminho do arquivo
    elif resposta == 5:
        diretorio_arquivo = None
        diretorio_pastas = None
        dados.titulo('Local do arquivo')
        diretorio_arquivo = input(
            'Diretório do Excel: ')
        # nome_do_arquivo = f'{diretorio_arquivo}\\{diretorio}'
        if not dados.inserirdiretorio(diretorio_arquivo):
            diretorio_arquivo = None
        diretorio_pastas = input(
            'Diretório das Pastas: ')
        print(
            f'\033[32mDiretório salvo.Certifique-se de que esteja correto: {diretorio_pastas} \033[m')

    # Instruções
    elif resposta == 6:
        dados.titulo('Instruções')
        print('\n\033[32m1) Criando o Arquivo Origem\033[m')
        print('     --> Crie uma planilha Excel com três colunas: (Nome, Assunto, E-mail).')
        print(
            '     \033[31m--> IMPORTANTE: NÃO DEIXE NENHUM CAMPO EM BRANCO NA PLANILHA.\033[m')
        print('     --> Na coluna nome, insira o nome exatamente igual ao nome da pasta onde será inserido os arquivos.')
        print(
            '     --> Na coluna assunto, coloque exatamente o título do assunto do email.')
        print('     --> Na colua E-mail, preencha o email corretamente. pode inserir mais de um email, porém    devem ser separados por ponto e vírgula.')

        print('\n\033[32m2) Informar Diretórios (opção 5)\033[m')
        print('     --> Vá na diretório onde você salvou o arquivo excel, clique com botão esquerdo do mouse e selecione "copiar como caminho".')
        print('     --> cole o caminho na opção "Diretório do Excel".')
        print(
            '     \033[31m--> IMPORTANTE: SE COLAR AS ASPAS DUPLAS " NO INICION E NO FINAL, APAGUE-AS.\033[m')
        print('     --> Da mesma forma, insira o local do diretório onde se encontram as pastas que contém os arquivos na opção "Diretório das Pastas"')
        print(
            '     \033[31m--> IMPORTANTE: Cada pasta deve ser nomeada com o mesmo nome na planilha.')
        print(
            '     --> IMPORTANTE: Só enviará arquivo com extensão ".pdf".\033[m')

        print('\n\033[32m3) Configurar Email (opção 2)\033[m')
        print('     --> Informe o login e senha do email que disparará os envios.')
        print('     --> SÓ ENVIA POR CONTA GOOGLE GMAIL.')
        print(
            '     \033[31m--> IMPORTANTE: A senha para utilização é a senha para aplicativos.\033[m')
        print('             --> ao acessar a conta de disparo vá na imagem da conta;')
        print('             --> clique em "Gerenciar sua conta do Google";')
        print('             --> selecione "segurança";')
        print('             --> vá na opção de senha de app;')
        print('             --> digite essa senha.')
        print(
            '     --> Insira o conteúdo do email. Você poderá digitar até dois parágrafos.')
        print('     --> Digite a assinatura.')
        print('     --> Informe os emails que deseja colocar em cópia. Separe por ";" caso seja vários.')

        print('\n\033[32m4) Enviar E-mails (opção 1)\033[m')
        print('     --> Execute somente após realizar os passos acima')
        print('     --> Aguarde receber a mensagem "Processo de envio concluído".')
        print('\n\n\n')
        dados.linha()

    # Opção para sair
    elif resposta == len(menu_opcoes):
        dados.titulo('Até logo')
        break
<<<<<<< HEAD

=======
>>>>>>> 273fbe32ae6a35a8bc65a5c749331c135438468e
