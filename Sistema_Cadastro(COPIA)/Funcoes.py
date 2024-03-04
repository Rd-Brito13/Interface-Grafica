import sqlite3
from tkinter import messagebox
import tkinter as tk


def Criar_Banco():
    conn = sqlite3.connect("Banco_Dados.db")
    cursor = conn.cursor()
    cursor.execute(
        '''
        CREATE TABLE IF NOT EXISTS Clientes(
        Codigo INTEGER PRIMARY KEY,
        Nome TEXT,
        Endereco TEXT,
        Cep TEXT,
        Telefone TEXT,
        Pedido TEXT,
        Valor REAL

        )
        '''
    )
    conn.commit()
    conn.close()

# Função para criar o banco e a tabela

def Carregar_Dados():
    conn = sqlite3.connect("Banco_Dados.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Clientes")

    Registros = cursor.fetchall()
    conn.close()

    return Registros
# Função para carregar os dados do banco 

def calcular_valor_pedido(Pedido):
    valor_total = 0
    Menu = {"familia:": 55.99,
            "grande:": 45.89,
            "media:": 35.79,
            "pequena:": 25.79,
            "brotinho:": 10.00
            }
    palavras = Pedido.split()
    print("Tudo certo")
    for palavra in palavras:
        if palavra.lower() in Menu:
            valor_total += Menu[palavra.lower()]
    valor_formatado = "{:.2f}".format(valor_total)
    return valor_formatado
# Função para capturar as palavras chaves no campo pedido e decidir o preço da nota no campo valor




# Função para coletar dados dos input e inserir no banco
def Inserir_Dados(Codigo, Nome, Endereco, Cep, Telefone, Pedido):
    conn = sqlite3.connect("Banco_Dados.db")
    cursor = conn.cursor()

    # Verifique se o código já existe no banco
    cursor.execute("SELECT MIN(Codigo) FROM Clientes")
    menor_codigo_pedido = cursor.fetchone()[0]
    if menor_codigo_pedido is None:
        menor_codigo_pedido = 1
        
    # O código não existe, então podemos prosseguir com a inserção
    else:
        cursor.execute("SELECT Codigo FROM Clientes ORDER BY Codigo")
        codigos_usados = set(row[0] for row in cursor.fetchall())
        while menor_codigo_pedido in codigos_usados:
            menor_codigo_pedido +=1
    
    valor_total = calcular_valor_pedido(Pedido)

    try: 
        cursor.execute("INSERT INTO Clientes (Codigo, Nome, Endereco, Cep, Telefone, Pedido,Valor) VALUES (?,?,?,?,?,?,?)",(menor_codigo_pedido, Nome,Endereco,Cep,Telefone,Pedido, valor_total))         
        print("sucesso")
        conn.commit()
        conn.close()
        return True
    except sqlite3.Error as e:
        print(f"Error ao inserir dados: {e}")
        conn.rollback()
        conn.close()
        return False
# função auxiliar que calcula o valor dos pedidos presentes no menu

def Excluir_Registros(Codigo, treeview):
    if not Codigo:
        return
    conn = sqlite3.connect("Banco_Dados.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Clientes WHERE Codigo = ?", (Codigo,))
    Registro = cursor.fetchone()

    if Registro:
        confirmar = messagebox.askyesno(
            "Confirmação",
            f"Você Deseja excluir o registro do código {Codigo}?\n Essa ação não podera ser desfeita."
        )
        if confirmar:
            cursor.execute("DELETE FROM Clientes WHERE Codigo = ?",(Codigo,))
            conn.commit()
            conn.close()
            messagebox.showinfo(title="Exclusão Concluida",message="Sucesso, o registro do codigo: {}, excluido com sucesso.".format(Codigo,))
            Atualizar_Exibicao_Registros(treeview)
        
            
        
    else: 
        conn.close()
        messagebox.showerror(title="Error na Busca pelo Código", message="O Código: {}, não existe no banco de dados.".format(Codigo))
# Função para fazer a exclusão dos registros

def Atualizar_Dados_Pedido(Codigo, Novo_Nome=None, Novo_Endereco=None, Novo_Cep=None, Novo_Telefone=None, Novo_Pedido=None):
    conn = sqlite3.connect("Banco_Dados.db")
    cursor = conn.cursor()

    # Consulta os dados atuais
    cursor.execute("SELECT Nome, Endereco, Cep, Telefone, Pedido FROM Clientes WHERE Codigo = ?", (Codigo,))
    dados_atuais = cursor.fetchone()

    if not dados_atuais:
        conn.close()
        return  # Não há registro com o código fornecido

    nome_atual, endereco_atual, cep_atual, telefone_atual, pedido_atual = dados_atuais

    # Verifica se cada campo está definido, caso contrário, mantém o valor atual
    if Novo_Nome is None:
        Novo_Nome = nome_atual
    if Novo_Endereco is None:
        Novo_Endereco = endereco_atual
    if Novo_Cep is None:
        Novo_Cep = cep_atual
    if Novo_Telefone is None:
        Novo_Telefone = telefone_atual
    if Novo_Pedido is None:
        Novo_Pedido = pedido_atual

    # Atualiza apenas os campos que não são None
    cursor.execute(
        "UPDATE Clientes SET Nome = ?, Endereco = ?, Cep = ?, Telefone = ?, Pedido = ? WHERE Codigo = ?",
        (Novo_Nome, Novo_Endereco, Novo_Cep, Novo_Telefone, Novo_Pedido, Codigo)
    )

    if Novo_Pedido is not None:
        novo_valor_pedido = calcular_valor_pedido(Novo_Pedido)
        cursor.execute("UPDATE Clientes SET Valor = ? WHERE Codigo = ?",(novo_valor_pedido, Codigo))


    conn.commit()
    conn.close()
   

def Consultar_Banco_Relatorio(Codigo):
    try:
        conn = sqlite3.connect("Banco_Dados.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Clientes WHERE Codigo = ?",(Codigo,))
        Dados = cursor.fetchone()
    
        conn.close()

        return Dados
    except sqlite3.Error as e:
        print("Erro ao Consultar o Banco de Dados", e)
        return None
    
# Função para consultar o banco de dados e selecionar o codigo que será informado pelo usuario


def Atualizar_Exibicao_Registros(treeview):
    conn = sqlite3.connect("Banco_Dados.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Clientes")
    Registros =  cursor.fetchall()
    conn.close()

    for item in treeview.get_children():
        treeview.delete(item)

    for Registro in Registros:
        treeview.insert("","end", values=Registro)

# Função para atualizar os dados inseridos em tempo real


    
 







