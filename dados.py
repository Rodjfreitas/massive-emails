def enviarEmail(usuario, senha_usuario, assunto, destinatario, conteudo, caminho_anexo):
    import os
    import smtplib
    # from email.message import EmailMessage
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.utils import COMMASPACE
    from email import encoders

    from email.mime.application import MIMEApplication

    # configurar email, senha
    EMAIL_ADRESS = usuario
    EMAIL_PASSWORD = senha_usuario

    try:
        # Startar o servidor
        host = 'smtp.gmail.com'
        port = '587'
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        server.login(EMAIL_ADRESS, EMAIL_PASSWORD)

        # Criar um email
        msg = MIMEMultipart()
        msg['Subject'] = assunto
        msg['From'] = usuario
        msg['To'] = destinatario
        msg.attach(MIMEText(conteudo, 'plain'))

        # Anexar arquivos
        pdf_folder = caminho_anexo
        pdf_files = [os.path.join(pdf_folder, f) for f in os.listdir(
            pdf_folder) if f.endswith('.pdf')]

        for file in pdf_files:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(open(file, 'rb').read())
            encoders.encode_base64(part)
            part.add_header('content-Disposition', 'attachment',
                            filename=os.path.basename(file))
            msg.attach(part)

        # Enviar email
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
        print(f'\033[32mEmail Enviado\033[m - {assunto}')
    except:
        print(f'\033[31mEmail Não Enviado\033[m - {assunto}')
        return 'Erro no envio'
    return 'Sucesso no envio'


def linha():
    print('=' * 60)


def titulo(msg):
    print('=' * 60)
    print(msg.center(60))
    print('=' * 60)


def menu(lista):
    titulo('Massive Emails')
    for pos, valor in enumerate(lista):
        print(f'{pos + 1:>3} -- {valor:30}')
    linha()
    while True:
        try:
            opc = int(input('Opção: '))
            if opc < 1 or opc > len(lista):
                print('\033[31mOpção inexistente\033[m')
                continue
        except (ValueError, TypeError):
            print('\033[31mOpção inexistente\033[m')
            continue
        except (KeyboardInterrupt):
            opc = len(lista)
        else:
            return opc


def saudacao():
    import datetime
    horas = datetime.datetime.now().hour
    if 0 <= horas <= 11:
        return 'Bom dia'
    if horas <= 17:
        return 'Boa tarde'
    return 'Boa noite'


def buscarnoArquivo(arquivo):
    import openpyxl
    try:
        # Abre o arquivo excel
        workbook = openpyxl.load_workbook(arquivo)

        # Seleciona a planilha desejada
        sheet = workbook['Planilha1']

        # Itera sobre todas as células da planilha
        valores = []
        linhas = {}
        i = 0
        for row in sheet.iter_rows():
            for cell in row:
                linhas[i] = cell.value
                i += 1
            valores.append(linhas.copy())
            linhas.clear()
            i = 0
    except:
        return 0
    return valores


def imprimirLista(lista):
    titulo('RELAÇÃO PARA ENVIO')
    for pos, valor in enumerate(lista):
        if pos == 0:
            print(
                f'{valor[0]:23}{valor[1]:20}{valor[3]:50}{valor[4]:50}')
            continue
        print(
            f'{pos:3}{valor[0]:20}{valor[1]:20}{valor[3]:50}{valor[4]:50}')
    linha()


def inserirdiretorio(nome):
    import os
    try:
        if os.path.exists(nome):
            ...
        else:
            print('\033[31mArquivo não existe\033[m')
            return False
    except:
        print('\033[31mArquivo não existe\033[m')
        return False
    print('\033[32mArquivo carregado com sucesso\033[m')
    return True
