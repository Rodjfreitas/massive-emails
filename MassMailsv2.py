from tkinter import *

janela = Tk()
janela.title('Mass&Mails - Disparador Massivo de E-mails')
janela.geometry('800x600')
janela.config(background='white')

texto_orientacao = Label(janela, text='Selecione as opções abaixo:', width=30, fg='blue')
texto_orientacao.grid(column=0, row=3)

botao = Button(janela, text='Enviar E-mail', command=...)
botao.grid(column=0, row=1)


janela.mainloop()