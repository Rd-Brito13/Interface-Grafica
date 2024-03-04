import tkinter as tk
import sqlite3
from tkinter import ttk
from tkinter import messagebox
from tkinter import *
from Funcoes import  Inserir_Dados,  Atualizar_Exibicao_Registros, Excluir_Registros, Consultar_Banco_Relatorio, Atualizar_Dados_Pedido


Tela = tk.Tk()
# VARIAVEL QEU ABRIGA A TELA PRINCIPAL 

Tela.title("Sistema de Cadastro de Clientes")
# Definindo o titulo da janela criada 
Tela.config(bg="#722B2B")
# Definindo cor de fundo da interface

Tela.geometry("650x450")
# Definindo o tamanho da janela

Tela.resizable(width=False, height=False)
# responsável tornar impossível alterar largura ou altura da interface
 

Logo = PhotoImage(file="Logo_Novo.png")
lb = Label(Tela, image=Logo, bg="#722B2B").place(x=150, y=10)
# Difinindo a imagem na interface e posicionando ela.

# Função acionada quando o botão dos clicado, para que seja inserido os dados no banco e atualizado automaticamente
def Janela_Coleta():
    Janela_Auxiliar1 = tk.Toplevel()
    Janela_Auxiliar1.title("Coleta de Dados")
    Janela_Auxiliar1.geometry("700x400")
    Janela_Auxiliar1.resizable(width=False,height=False)
    Janela_Auxiliar1.config(bg="#722B2B")
    tk.Label(Janela_Auxiliar1,text="Dados Cadastrais",fg="#000000", bg="#722B2B",font=("Italic", 17, "bold")).place(x=270,y=50)

    
    # CÓDIGO DAS CAIXAS DE ENTRADA E DOS TEXTOS
    var_Codigo = tk.StringVar()
    tk.Entry(Janela_Auxiliar1, width="10", textvariable=var_Codigo).place(x=168, y=120)
    tk.Label(Janela_Auxiliar1, text="Código:", bg="#722B2B",font="bold").place(x=100, y= 116)
    # Entrada_Codigo é a caixa de texto utilizada para capturar o Código, Texto_Codigo é a mensagem simbolizando o pedido do cliente.

    var_Nome = tk.StringVar()
    tk.Entry(Janela_Auxiliar1,width="57", textvariable=var_Nome).place(x=168, y=145)
    tk.Label(Janela_Auxiliar1, text="Nome:",bg="#722B2B",font="bold").place(x=110, y= 140)
    # Entrada_Nome é a caixa de texto utilizada para capturar o nome, Texto_Nome é a mensagem simbolizando o pedido do cliente.
    var_Endereco = tk.StringVar()
    tk.Entry(Janela_Auxiliar1, width="57", textvariable=var_Endereco).place(x=168, y=170)
    tk.Label(Janela_Auxiliar1, text="Endereço:", bg="#722B2B",font="bold").place(x=85, y= 165)
    # Entrada_Endereco é a caixa de texto utilizada para capturar endereco, Texto_Endereco é a mensagem simbolizando o pedido do cliente.
    var_Cep = tk.StringVar()
    tk.Entry(Janela_Auxiliar1, width="12", textvariable=var_Cep).place(x=168, y=195)
    tk.Label(Janela_Auxiliar1, text="Cep:", bg="#722B2B",font="bold").place(x=122, y= 190)
    # Entrada_Cep é a caixa de texto utilizada para capturar Cep, Texto_Endereco é a mensagem simbolizando o pedido do cliente.

    var_Telefone= tk.StringVar()
    tk.Entry(Janela_Auxiliar1, width="25", textvariable=var_Telefone).place(x=168, y=220)
    tk.Label(Janela_Auxiliar1, text="Telefone:", bg="#722B2B",font="bold").place(x=92, y= 215)
    # Entrada_Telefone é a caixa de texto utilizada para capturar o Telefone Texto_Telefone é a mensagem simbolizando o pedido do cliente.

    var_Pedido= tk.StringVar()
    tk.Entry(Janela_Auxiliar1, width="30", textvariable=var_Pedido).place(x=168, y=245)
    tk.Label(Janela_Auxiliar1, text="Pedido:", bg="#722B2B",font="bold").place(x=100, y= 240)
    # Entrada_Pedido é a caixa de texto utilizada para capturar o Pedido, Texto_Pedido é a mensagem simbolizando o pedido do cliente.
    
    def Botao_Inserir_Dados():
        
        Codigo = var_Codigo.get()
        Nome = var_Nome.get()
        Endereco = var_Endereco.get()
        Cep = var_Cep.get()
        Telefone = var_Telefone.get()
        Pedido = var_Pedido.get()
        if Inserir_Dados(Codigo, Nome, Endereco, Cep, Telefone,Pedido):
            Janela_Auxiliar1.destroy()
            messagebox.showinfo(title="Sucesso",message="Dados cadastrados com Sucesso.")
        # Chamada a função para inserir os dados coletados pelo input, Utiliza as variaveis de controle para coletar os dados de input.
            Atualizar_Exibicao_Registros(treeview=Exibir_Registros)
        # Função para manter atualizado o banco de dados a cada entrada de novos dados.
        else:
            messagebox.showerror("Falha","Dados Não cadastrados")
        # Com esta função sempre que o botão for clicado acionamos uma caixa de mensagem.
        
    tk.Button(Janela_Auxiliar1, background="#C66B1D",fg="#FFFFFF",text="Confirmar e Salvar", command=Botao_Inserir_Dados).place(x=300, y=320)
    



# Função Para criar uma janela auxiliar, onde, vai ter somente um campo de entrada que vai ser utilizado para buscar um código no banco de dados e apagar todo o seu registro, através do código do pedido
def Botao_Abrir_Janela_Exclusao():
    Janela_Auxiliar = tk.Toplevel()
    Janela_Auxiliar.title("Exclusão de Registro")
    Janela_Auxiliar.geometry("500x350")
    Janela_Auxiliar.resizable(width=False, height=False)
    Janela_Auxiliar.config(bg="#722B2B")
    tk.Label(Janela_Auxiliar, text="Codigo:",font="bold",bg="#722B2B").place(x = "150", y= "150")
    tk.Label(Janela_Auxiliar,text="Insira o Código que deseja o Excluir", font=" Helvetica  17 bold ", fg="black",bg="#722B2B").place(x= "50", y = "75")
# Cria a janela auxiliar, Define o tamanho e que ela não pode ser alterada, define o titulo, cria o titulo pro campo e define a posição dele na tela.
    var_Codigo_Exclusao = tk.StringVar()
    tk.Entry(Janela_Auxiliar, textvariable=var_Codigo_Exclusao).place(x="215", y="153")
# Cria uma variavel auxiliar para receber o valor do campo e cria o campo de entrada de dados, define a posição dele na tela.

    def Confirmar_Exclusao():
        Codigo = var_Codigo_Exclusao.get()
        Janela_Auxiliar.destroy()
        Excluir_Registros(Codigo, Atualizar_Exibicao_Registros)
# Função para buscar o codigo no banco de dados e realizar a exclusão de todos os dados buscado pelo código.
    tk.Button(Janela_Auxiliar, text="Confirmar",fg="#FFFFFF", bg="#C66B1D", command=Confirmar_Exclusao).place(x = "240", y = "200")
# Botão para concluir a operação de exclusão.




def Botao_Abrir_Janela_Relatorio():
    Janela_Auxiliar2 = tk.Toplevel()
    Janela_Auxiliar2.title("Buscar Dados")
    Janela_Auxiliar2.geometry("500x350")
    Janela_Auxiliar2.resizable(width=False, height=False)
    Janela_Auxiliar2.config(bg="#722B2B")
    tk.Label(Janela_Auxiliar2,text="Código:",bg="#722B2B",font="bold").place(x="150", y="150")
    tk.Label(Janela_Auxiliar2,text="Insira o Código que deseja o relátorio", bg="#722B2B",font=" Helvetica  17 bold ", fg="black").place(x= "50", y = "75")
# Cria a janela auxiliar, Define o tamanho e que ela não pode ser alterada, define o tipo do campo e define o titulo da janela.
    var_Codigo_Busca = tk.StringVar()
    tk.Entry(Janela_Auxiliar2,textvariable=var_Codigo_Busca).place(x="215",y="153")
# Cria a variavel de controle e cria o campo de entrada de dados.
    def Buscar_Dados(Codigo, Janela_Relatorio):    
        Codigo = var_Codigo_Busca.get()
        Dados = Consultar_Banco_Relatorio(Codigo)
        if Dados:
            Janela_Relatorio = tk.Toplevel()
            Janela_Relatorio.title("Janela de Relatório")
            Janela_Relatorio.geometry("500x300")
            Janela_Relatorio.resizable(width=False,height=False)
            Janela_Relatorio.config(bg="#722B2B")
            Janela_Auxiliar2.destroy()
            # Cria a janela de exibição dos dados coletaos
            for i, Rotulos in enumerate(["Codigo:", "Nome:", "Endereço:","CEP:","Telefone:","Pedido:", "Valor:"]):
                tk.Label(Janela_Relatorio,bg="#722B2B",fg="#FFFFFF",font="bold", text=Dados[i]).grid(row=i, column=1,padx=2,pady=5, sticky="w")
                tk.Label(Janela_Relatorio,bg="#722B2B",fg="#FFFFFF",font="bold", text=Rotulos).grid(row=i, column=0,padx=10,pady=5, sticky="w")
             # exibe os dados definidos pelo codigo e o formata na GUI
            
        else: 
            messagebox.showerror(title="Error",message="Código Não foi encontrado")  
# Função parar criar a janela de exibição do relatorio informado pelo código.
    tk.Button(Janela_Auxiliar2,background="#C66B1D",fg="#FFFFFF",text="Confirmar", command=lambda:Buscar_Dados(var_Codigo_Busca.get(),Janela_Auxiliar2)).place(x="240", y="200")
    
# Botão para confiar a busca no banco de dados. 
    
def Botao_Alterar_Dados():
    Janela_Alterar = tk.Toplevel()
    Janela_Alterar.title("Alterar Clientes")
    Janela_Alterar.geometry("500x350")
    Janela_Alterar.resizable(width=False, height=False)
    Janela_Alterar.config(bg="#722B2B")
    tk.Label(Janela_Alterar, text="Código:", font="bold",bg="#722B2B").place(x="150", y="150")
    tk.Label(Janela_Alterar,text="Insira o Código que deseja alterar", font=" Helvetica  17 bold ", fg="black",bg="#722B2B").place(x= "80", y = "75")
    # Cria a janela que pede o código que tera os dados alterados
    var_Codigo_Alterar = StringVar()
    tk.Entry(Janela_Alterar,textvariable=var_Codigo_Alterar).place(x="215",y="153")
    # Cria o Botao para disparar o evento de busca.
    def Inserir_Dados_Alteracao():
        Codigo = var_Codigo_Alterar.get()
        Dados = Consultar_Banco_Relatorio(Codigo)
        if Dados:
            Janela_Inserir_Alterar = tk.Toplevel()
            Janela_Inserir_Alterar.title("Inserir Dados a Serem Alterados")
            tk.Label(Janela_Inserir_Alterar,text="Insira os Dados", bg="#722B2B",font=" Helvetica  17 bold ", fg="black").place(x= "250", y = "40")
            Janela_Inserir_Alterar.geometry("600x400")
            Janela_Inserir_Alterar.config(bg="#722B2B")
            Janela_Inserir_Alterar.resizable(width=False,height=False)
            Janela_Alterar.destroy()
            # Cria a janela para a inserção de dados
            
            var_Alterar_Nome = tk.StringVar()
            tk.Entry(Janela_Inserir_Alterar,width="57", textvariable=var_Alterar_Nome).place(x=168, y=103)
            tk.Label(Janela_Inserir_Alterar, bg="#722B2B",text="Nome:", font="bold").place(x=110, y= 100)
            # Entrada_Nome é a caixa de texto utilizada para capturar o nome, Texto_Nome é a mensagem simbolizando o pedido do cliente.
            var_Alterar_Endereco = tk.StringVar()
            tk.Entry(Janela_Inserir_Alterar, width="57", textvariable=var_Alterar_Endereco).place(x=168, y=129)
            tk.Label(Janela_Inserir_Alterar, bg="#722B2B",text="Endereço:", font="bold").place(x=85, y= 127)
            # Entrada_Endereco é a caixa de texto utilizada para capturar endereco, Texto_Endereco é a mensagem simbolizando o pedido do cliente.
            var_Alterar_Cep = tk.StringVar()
            tk.Entry(Janela_Inserir_Alterar, width="12", textvariable=var_Alterar_Cep).place(x=168, y=155)
            tk.Label(Janela_Inserir_Alterar, bg="#722B2B",text="Cep:", font="bold").place(x=122, y= 152)
            # Entrada_Cep é a caixa de texto utilizada para capturar Cep, Texto_Endereco é a mensagem simbolizando o pedido do cliente.

            var_Alterar_Telefone= tk.StringVar()
            tk.Entry(Janela_Inserir_Alterar, width="25", textvariable=var_Alterar_Telefone).place(x=168, y=180)
            tk.Label(Janela_Inserir_Alterar, bg="#722B2B",text="Telefone:",  font="bold").place(x=92, y= 180)
            # Entrada_Telefone é a caixa de texto utilizada para capturar o Telefone Texto_Telefone é a mensagem simbolizando o pedido do cliente.

            var_Alterar_Pedido= tk.StringVar()
            tk.Entry(Janela_Inserir_Alterar, width="30", textvariable=var_Alterar_Pedido).place(x=168, y=205)
            tk.Label(Janela_Inserir_Alterar, bg="#722B2B",text="Pedido:", font="bold").place(x=100, y= 202)
            #Rotulos, Campos de entrada e Botão, utilizado na alteração de dados do código desejado.
            def Salvar_Alteracoes():
                Novo_Nome_Valor = var_Alterar_Nome.get() or None
                Novo_Endereco_Valor = var_Alterar_Endereco.get() or None
                Novo_Cep_Valor = var_Alterar_Cep.get() or None
                Novo_Telefone_Valor = var_Alterar_Telefone.get() or None
                Novo_Pedido_Valor =var_Alterar_Pedido.get() or None

                if Novo_Nome_Valor is not None or Novo_Endereco_Valor is not None or Novo_Cep_Valor is not None or Novo_Telefone_Valor is not None or Novo_Pedido_Valor is not None:
                    Atualizar_Dados_Pedido(Codigo, Novo_Nome_Valor, Novo_Endereco_Valor, Novo_Cep_Valor, Novo_Telefone_Valor, Novo_Pedido_Valor)
                    messagebox.showinfo(title="Sucesso", message="Dados Alterados com Sucesso!")
                    Janela_Inserir_Alterar.destroy()
                    Atualizar_Exibicao_Registros
            tk.Button(Janela_Inserir_Alterar, background="#C66B1D",fg="#FFFFFF",text="Confirmar e Alterar", command=Salvar_Alteracoes).place(x="250", y="275")

        else: 
            messagebox.showerror(title="Error",message="Código Não foi encontrado")
    tk.Button(Janela_Alterar,text="Confirmar", background="#C66B1D",fg="#FFFFFF",command=Inserir_Dados_Alteracao).place(x="240", y="200")

def Exibir_Registros():     
    Janela_Dados = tk.Toplevel()
    Janela_Dados.title("Banco de Dados")
    Janela_Dados.geometry("1400x700")
    Janela_Dados.resizable(width=False, height=False)


    VisualizarDados = tk.ttk.Treeview(Janela_Dados, columns=("Codigo", "Nome", "Endereço", "CEP", "Telefone", "Pedido", "Valor"), show="headings")
    VisualizarDados.heading("Codigo", text="Codigo")
    VisualizarDados.heading("Nome", text="Nome")
    VisualizarDados.heading("Endereço", text="Endereço")
    VisualizarDados.heading("CEP", text="CEP")
    VisualizarDados.heading("Telefone", text="Telefone")
    VisualizarDados.heading("Pedido", text="Pedido")
    VisualizarDados.heading("Valor", text="Valor")
    VisualizarDados.pack(fill=tk.BOTH,expand=True)
    Atualizar_Exibicao_Registros(treeview=VisualizarDados)
    





#CODIGO DOS BOTÕES
tk.Button(Tela,text = "Cadastrar Pedido", padx=12, background="#C66B1D",fg="#FFFFFF",command=Janela_Coleta).place(x=280, y = 196)
# Botão para confirmar um pedido e salvar as informações do cliente e do pedido no banco
tk.Button(Tela,text = "Alterar Pedido", padx=20, background="#C66B1D",fg="#FFFFFF",command=Botao_Alterar_Dados).place(x=280, y = 225)
# Botão para alterar o pedido.
tk.Button(Tela,text = "Relatório do Pedido", padx=6, background="#C66B1D",fg="#FFFFFF",command=Botao_Abrir_Janela_Relatorio).place(x=280, y = 255)
# Botão parar gerar um relatório pelo código do pedido.
tk.Button(Tela,text = "Cancelar Pedido", padx=15, bg="#C66B1D",fg="#FFFFFF",command=Botao_Abrir_Janela_Exclusao).place(x=280, y = 285)
# Botão para cancelar o pedido do cliente.
tk.Button(Tela,text="Exibir Dados", padx=26, background="#C66B1D",fg="#FFFFFF",command=Exibir_Registros).place(x=280, y=315)


    





tk.mainloop()