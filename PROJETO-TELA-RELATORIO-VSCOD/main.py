import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from tkinter import PhotoImage

# Definindo a senha padrão
senha_padrao = "BST123"
arquivos_selecionados = []
tela_auxiliar = None
tela_auxiliar_aberta = False
tela_principal_congelada = False

def verificar_senha(entry):
    global tela_auxiliar_aberta, tela_principal_congelada
    senha = entry.get()
    if senha == senha_padrao:
        messagebox.showinfo("Acesso", "Acesso Liberado!")
        descongelar_tela_principal()
        tela_principal_congelada = False
        tela_auxiliar.destroy()
    else:
        root.destroy()

def fechar_janela_auxiliar():
    global tela_auxiliar_aberta, tela_auxiliar
    tela_auxiliar_aberta = False
    if tela_auxiliar is not None and tela_auxiliar.winfo_exists():  # Verifica se a janela auxiliar ainda existe
        tela_auxiliar.destroy()

def congelar_tela(event):
    global tela_auxiliar_aberta, tela_principal_congelada, tela_auxiliar
    if not tela_auxiliar_aberta and not tela_principal_congelada:
        root.config(cursor="wait")  # Muda o cursor para "espera"
        tela_auxiliar_aberta = True
        tela_auxiliar = tk.Toplevel(root)
        tela_auxiliar.title("Desbloquear Aplicativo")
        tela_auxiliar.configure(bg="DarkGrey")
        tela_auxiliar.geometry("400x300")
        tk.Label(tela_auxiliar, text="Desbloquear Aplicativo", font=("League Spartan", 15, "bold", "italic"),bg="DarkGrey").place(
            x=100, y=50)
        tk.Label(tela_auxiliar, text="SENHA:", font=("League Spartan", 10, "bold", "italic"),bg="DarkGrey").place(x=60, y=148)
        entry_senha = tk.Entry(tela_auxiliar, width=30, show="*")
        entry_senha.place(x=130, y=150)
        tk.Button(tela_auxiliar, text="Confirmar", width=20, height=2,
                  command=lambda: verificar_senha(entry_senha)).place(x=140, y=200)
        root.after(3000, lambda: root.config(cursor=""))
        congelar_tela_principal()

def congelar_tela_principal():
    for child in tela_principal.winfo_children():
        child.configure(state="disabled")

def descongelar_tela_principal():
    for child in tela_principal.winfo_children():
        child.configure(state="normal")
        
def procurar_arquivos():
    global arquivos_selecionados
    arquivos_selecionados = []  # Limpa a lista de arquivos
    file_names = filedialog.askopenfilenames(filetypes=[("Arquivos Excel", "*.xlsx")])
    if 2 <= len(file_names) <= 5:  # Verifica se o número de arquivos selecionados está entre 2 e 5
        # Define as partes comuns nos nomes dos arquivos que serão usadas para ordenação
        parte_comum_apuracao = "APURACAO_LUCRATIVIDADE"
        parte_comum_romaneio = "ROMANEIO_EMITIDOS"

        # Verifica se os arquivos contêm as partes comuns nos nomes e os coloca por último na lista
        arquivos_selecionados = sorted(file_names, key=lambda x: (parte_comum_apuracao in os.path.basename(x), parte_comum_romaneio in os.path.basename(x)))

        messagebox.showinfo("Arquivos Selecionados", f"Arquivos selecionados:\n{', '.join(os.path.basename(arquivo) for arquivo in arquivos_selecionados)}")
        print(arquivos_selecionados)
        return arquivos_selecionados
    elif len(file_names) < 2:
        messagebox.showwarning("Número Insuficiente de Arquivos", "Selecione pelo menos dois arquivos.")
    else:
        messagebox.showwarning("Número Excessivo de Arquivos", "Selecione no máximo cinco arquivos.")
        return None


def encapsulated_function(caminho_arquivo1, caminho_arquivo2):
    coletas_solicitadas = pd.read_excel(caminho_arquivo1)
    apuracao_lucratividade = pd.read_excel(caminho_arquivo2)

    def format_coletas_solicitadas(dataframe):
        #Excluindo a primeira linha do dataframe
        dataframe.drop([0],axis=0, inplace=True)
        #criando o dataframe com o numero do ct e o numero de romaneio
        dataframe2 = dataframe[["ROMANEIO","NUMERO"]]
        #utilizando a função (replace, para trocar os dados pela função , pd.NA e assim transformar os dados em NaN
        #e utilizando a função dropna(para remover todas os dados NaN)
        dataframe2 = dataframe2.replace("0.0", pd.NA).dropna(subset=["NUMERO","ROMANEIO"])
        #retornando o dataframe2
        return dataframe2
    
        # Sua função format_coletas_solicitadas aqui

    def assimilar_valores_coletas(dataframe):
         #esta linha de código, cria um novo dataframe, e utiliza da função groupby para agrupar os numeros com base no romaneio e transformar em lista
        coleta_solicitadas_agrupado = dataframe.groupby("NUMERO")["ROMANEIO"].apply(list).reset_index()
        #utilizando a expressão lambda pora remover os [colchestes] dos valores da coluna romaneio
        coleta_solicitadas_agrupado["ROMANEIO"] = coleta_solicitadas_agrupado["ROMANEIO"].apply(lambda x: x[0])
        #formatando os valores na coluna NUMERO e ROMANEIO
        coleta_solicitadas_agrupado['ROMANEIO'] = coleta_solicitadas_agrupado['ROMANEIO'].astype(int)
        coleta_solicitadas_agrupado["NUMERO"] = coleta_solicitadas_agrupado["NUMERO"].astype(int)
        #Criando um novo dataframe e renomeando as colunas
        coleta_solicitadas_agrupado = coleta_solicitadas_agrupado[["ROMANEIO","NUMERO"]]
        coleta_solicitadas_agrupado.rename(columns={"NUMERO":"ORDEM COLETA"}, inplace=True)
        #definindo a coluna romaneio com "str" e adicionando sempre 6 numero ao valor das celulas
        coleta_solicitadas_agrupado['ROMANEIO'] = coleta_solicitadas_agrupado['ROMANEIO'].astype(str)
        coleta_solicitadas_agrupado['ROMANEIO'] = coleta_solicitadas_agrupado['ROMANEIO'].str.zfill(6)
        #salvando o arquivo
        #coleta_solicitadas_agrupado.to_excel("teste1.xlsx")
        return coleta_solicitadas_agrupado
            # Sua função assimilar_valores_coletas aqui

    def format_apuracao_lucratividade(dataframe):
        #criando outro dataframe com base nas, FILIAL, SERIE, CTE
        apuracao_lucratividade2 = apuracao_lucratividade[["Unnamed: 5","Unnamed: 6","Unnamed: 7"]]
        #renomeando as colunas extraidas
        apuracao_lucratividade2.rename(columns={"Unnamed: 5":"FILIAL","Unnamed: 6":"SERIE","Unnamed: 7":"CTE"},inplace=True)
        #concatenando as colunas "ordem coleta" e "CTE" e formando um novo dataframe
        apuracao_lucratividade3 = pd.concat([coleta_solicitadas_agrupado["ORDEM COLETA"], apuracao_lucratividade2["CTE"]], axis=1)
        #Assimilando os valores da coluna "ordem coleta" em "cte"
        apuracao_lucratividade4 = apuracao_lucratividade3.groupby("ORDEM COLETA")["CTE"].apply(list).reset_index()
        #removendo os "[colchetes]" do valores das celulas 
        apuracao_lucratividade4["CTE"] = apuracao_lucratividade4["CTE"].apply(lambda x: x[0])
        #adicionando o valor "0" a celula da linha 0 da coluna "CTE", para substituir o valor "NaN" adicionando pela função
        apuracao_lucratividade4.at[0,"CTE"] = 0
        #seprando a coluna "romaneio" e criando um novo dataframe
        apuracao_lucratividade5 = coleta_solicitadas_agrupado[["ROMANEIO"]]
        #concatenando dos dataframes de, "ordem coleta, cte" com o "romaneio"
        apuracao_lucratividade6 = pd.concat([apuracao_lucratividade5, apuracao_lucratividade4],axis=1)
        #transformando os dados de "ordem coleta, cte" em inteiros
        novos_tipos_de_dados = {'CTE': int,"ORDEM COLETA":int}
        apuracao_lucratividade6 = apuracao_lucratividade6.astype(novos_tipos_de_dados)
        #apuracao_lucratividade6.to_excel("PRE_FINALIZADO_MARCO_BFS1.xlsx")
        apuracao_lucratividade_copia = apuracao_lucratividade.copy()
        
        #retornando o dataframe fortamdo
        return apuracao_lucratividade6, apuracao_lucratividade_copia
            # Sua função format_apuracao_lucratividade aqui

    def finalizar_dataframe(caminho1, caminho2):
         #criar o dataframe com base nos dados oferecidos
        finalizado1 = caminho1
        #remover a 1 linha do 1 dataframe
        #finalizado1.drop(["Unnamed: 0"],axis=1,inplace=True)
        # ----------------------------------------------------
        #criando dataframe para ser concatenado
        apuracao = caminho2
        #concatenando eles
        finalizado2 = pd.concat([apuracao,finalizado1],axis=1)
        #excluindo colunas
        finalizado2.drop(["CONTROLE","CTE"],axis=1,inplace=True)
        #renomeando as colunas
        finalizado2.rename(columns={"Unnamed: 5":"FILIAL","Unnamed: 6":"SERIE","Unnamed: 10":"ORIGEM CIDADE","Unnamed: 12":"CIDADE DESTINO","Unnamed: 7":"CTE"},inplace=True)
        #função utilizada para transformar todas as linhas do dataframe que estão no formato "NaN" para "0"
        finalizado2 = finalizado2.fillna(0)
       
        #transformando as colunas especificadas no tipo inteiro e atribuindo 6 valores a coluna romaneio
        finalizado2["ROMANEIO"] = finalizado2["ROMANEIO"].astype(int)
        finalizado2["ORDEM COLETA"] = finalizado2["ORDEM COLETA"].astype(int)
        finalizado2["SERIE"] = finalizado2["SERIE"].astype(int)
        finalizado2["CTE"] = finalizado2["CTE"].astype(int)
        finalizado2['ROMANEIO'] = finalizado2['ROMANEIO'].astype(str)
        finalizado2['ROMANEIO'] = finalizado2['ROMANEIO'].str.zfill(6)
       
         #criando uma lista para ree-ordenar as colunas e executando isto
        troca_colunas = ["MANIFESTO","MOTORISTAS","DATA EMISSAO", "ROMANEIO","ORDEM COLETA","DOC EMBARQUE","FILIAL","SERIE","CTE","PAGADOR"
                         ,"ORIGEM","ORIGEM CIDADE","DESTNO","CIDADE DESTINO","KGS", "FRETE","KGS COLETA","CUSTO COLETA","CUSTO TRANSFERENCIA","CUSTO ENTREGA","RCF","RCTRC","FLUVIAL",
                         "IMPOSTO", "TOTAL CUSTO","SALDO"]
        finalizado2 = finalizado2[troca_colunas]
        
        finalizado2["DATA EMISSAO"] = pd.to_datetime(finalizado2["DATA EMISSAO"]) 
        finalizado2['DATA EMISSAO'] = finalizado2['DATA EMISSAO'].dt.strftime('%d/%m/%Y')
        finalizado2.at[0,"DATA EMISSAO"] = 0
       

        
        #salvando o novo dataframe em formato de excel
        def salvar_arquivo():
            local_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                  filetypes=[("Arquivo Excel", "*.xlsx")])
            if local_arquivo:  # Verifica se o usuário selecionou um local de arquivo  
                finalizado2.to_excel(local_arquivo, index=False)
                print(f'O arquivo {local_arquivo} foi salvo com sucesso.')

        # Chama a função para salvar o arquivo usando a caixa de diálogo do sistema operacional Windows
        salvar_arquivo()

    coletas_solicitadas2 = format_coletas_solicitadas(coletas_solicitadas)
    coleta_solicitadas_agrupado = assimilar_valores_coletas(coletas_solicitadas2)
    apuracao_lucratividade6, apuracao_lucratividade_copia = format_apuracao_lucratividade(apuracao_lucratividade)
    finalizado = finalizar_dataframe(apuracao_lucratividade6, apuracao_lucratividade_copia)

    return finalizado
    
def romaneios_executados(caminho1, caminho2):
    romaneio_executado = pd.read_excel(caminho1, dtype={'ROMANEIO': str})
    romaneio_emitido = pd.read_excel(caminho2, dtype={'ROMANEIO': str})
    
   
    # Remove espaços em branco e preenche com zeros à esquerda para ter 6 dígitos
    romaneio_executado['ROMANEIO'] = romaneio_executado['ROMANEIO'].str.strip().str.zfill(6)
    romaneio_emitido['ROMANEIO'] = romaneio_emitido['ROMANEIO'].str.strip().str.zfill(6)
    
    #remove a 1 linha dos emitods
    romaneio_emitido.drop(0,axis=0, inplace=True)
    romaneio_executado.drop(0, axis=0, inplace=True)
    
    #função para comparar os romaneios e criar o documento de romaneio sem valor
    romaneios_sem_valor = {"ROMANEIOS_SEM_VALOR": [], "FILIAL":[]}
    for index, row in romaneio_emitido.iterrows():
        romaneio = row["ROMANEIO"]
        filial = row["EMISSORA"]
    
        if romaneio not in romaneio_executado["ROMANEIO"].values:
            print("OK")
            romaneios_sem_valor["ROMANEIOS_SEM_VALOR"].append(romaneio)
            romaneios_sem_valor["FILIAL"].append(filial)
    
    # Cria um DataFrame com os valores de ROMANEIO sem valor
    romaneio_sem_valor = pd.DataFrame(romaneios_sem_valor)
    romaneio_sem_valor['ROMANEIOS_SEM_VALOR'] = romaneio_sem_valor['ROMANEIOS_SEM_VALOR'].astype(str)
    # Remove espaços em branco e preenche com zeros à esquerda para ter 6 dígitos
    romaneio_sem_valor['ROMANEIOS_SEM_VALOR'] = romaneio_sem_valor['ROMANEIOS_SEM_VALOR'].astype(str).str.strip().str.zfill(6)
    
    
    
    
    
    dataframe_auxiliar = romaneio_emitido[['ROMANEIO', 'FRT.AGREGADO APURADO']]
    dataframe_auxiliar.rename(columns={'FRT.AGREGADO APURADO': "APURADO"},inplace=True)
    
    
    romaneio_final = pd.merge(romaneio_executado, dataframe_auxiliar, on="ROMANEIO",how="inner")
    
        
    romaneio_final["EMISSAO"] = pd.to_datetime(romaneio_final["EMISSAO"]) 
    romaneio_final["BAIXA"] = pd.to_datetime(romaneio_final["BAIXA"])

    romaneio_final['EMISSAO'] = romaneio_final['EMISSAO'].dt.strftime('%d/%m/%Y')
    romaneio_final['BAIXA'] = romaneio_final['BAIXA'].dt.strftime('%d/%m/%Y')
    
    
    
    def salvar_arquivo():
            local_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                  filetypes=[("Arquivo Excel", "*.xlsx")])
            if local_arquivo:  # Verifica se o usuário selecionou um local de arquivo  
                romaneio_final.to_excel(local_arquivo, index=False)
                print(f'O arquivo {local_arquivo} foi salvo com sucesso.')
                
                local_romaneio_sem_valor = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Arquivo Excel", "*.xlsx")])

                if local_romaneio_sem_valor:
                    romaneio_sem_valor.to_excel(local_romaneio_sem_valor, index=False)
                    print(f'O arquivo {local_romaneio_sem_valor} foi salvo com sucesso.')
                
        # Chama a função para salvar o arquivo usando a caixa de diálogo do sistema operacional Windows
    salvar_arquivo()

    return romaneio_final


def gerar_romaneio_sem_valor(arquivos_selecionados):
        

        # Carregar os arquivos Excel em DataFrames separados
        spo, mto, mtz, bhz, fsa = (pd.read_excel(caminho) for caminho in arquivos_selecionados)
        # Carregar os DataFrames

        # Converter as colunas para strings e adicionar zeros à esquerda
        for df in [mto, bhz, spo, fsa,mtz]:
            df['ROMANEIOS_SEM_VALOR'] = df['ROMANEIOS_SEM_VALOR'].astype(str).str.zfill(6)

        # Concatenar os DataFrames
        romaneio_sem_valor_final = pd.concat([mto,bhz,spo,fsa,mtz], axis=1)
        # Lista de nomes das colunas que você deseja preencher
        colunas_para_preencher = ['ROMANEIOS_SEM_VALOR']

        # Dicionário de valores de preenchimento para as colunas selecionadas
        valores_para_preencher = {coluna: 0 for coluna in colunas_para_preencher}

        # Preenche os valores NaN nas colunas selecionadas com os valores especificados
        romaneio_sem_valor_final[colunas_para_preencher] = romaneio_sem_valor_final[colunas_para_preencher].fillna(valores_para_preencher)
        romaneio_sem_valor_final.rename(columns={"ROMANEIOS_SEM_VALOR":"ROMANEIOS"},inplace=True)
        def salvar_arquivo():
            local_romaneio_sem_valor = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Arquivo Excel", "*.xlsx")])

            if local_romaneio_sem_valor:
                romaneio_sem_valor_final.to_excel(local_romaneio_sem_valor, index=False)
                print(f'O arquivo {local_romaneio_sem_valor} foi salvo com sucesso.')
                
        salvar_arquivo()        
        return romaneio_sem_valor_final
        
    


def mostrar_tela_principal():
    global tela_principal, tela_secundaria
    tela_secundaria.pack_forget()
    tela_principal.pack(fill="both", expand=True)

def mostrar_tela_secundaria():
    global tela_principal, tela_secundaria
    tela_principal.pack_forget()
    tela_secundaria.pack(fill="both", expand=True)

def iniciar_apuracao_lucratividade():
    global arquivos_selecionados
    if len(arquivos_selecionados) == 2:
        resultado_final = encapsulated_function(arquivos_selecionados[0], arquivos_selecionados[1])
        # Faça algo com o resultado, por exemplo, exibição em uma nova janela ou salvamento em um arquivo
        messagebox.showinfo("Apuração Lucratividade", "Apuração Lucratividade concluída com sucesso!")
    else:
        messagebox.showwarning("Arquivos Insuficientes", "Selecione os arquivos necessários antes de iniciar a apuração de lucratividade.")

def iniciar_romaneios_executados():
    global arquivos_selecionados
    if len(arquivos_selecionados) == 2:
        # Chama a função para processar os romaneios executados com os arquivos selecionados
        resultado_final = romaneios_executados(arquivos_selecionados[0], arquivos_selecionados[1])
        # Mostra uma mensagem informando que o relatório foi gerado com sucesso
        messagebox.showinfo("Romaneios Executados", "Relatório de Romaneios Executados gerado com sucesso!")
    else:
        messagebox.showwarning("Arquivos Insuficientes", "Selecione os arquivos necessários antes de gerar o relatório de romaneios executados.")

def inicar_romaneios_sem_valor():
    global arquivos_selecionados
    if len(arquivos_selecionados) == 5:
        resultado_final = gerar_romaneio_sem_valor(arquivos_selecionados)
        messagebox.showinfo("Romaneios Sem Valor", "Relatório de Romaneio Sem Valor gerado com sucesso")
    else:
        messagebox.showwarning("Arquivos Insuficientes", "Selecione os arquivos necessários antes de gerar o relatório de romaneios sem valor.")


root = tk.Tk()
root.title("BST - Relatórios Operativos")
root.geometry("1200x768")
root.configure(bg="DarkGrey")
root.resizable(False, False)

# Carregar a imagem de logo
Logo = tk.PhotoImage(file=r"logo_bahiasul.PNG")

# Criando tela principal
tela_principal = tk.Frame(root, bg="DarkGrey")
tela_principal.pack(expand=True, fill="both")

# Criando a label da tela principal
label_imagem = tk.Label(tela_principal, image=Logo, bg="DarkGrey")
label_imagem.place(x=220,y=150)



# Botão iniciar a dela principal
botao_iniciar = tk.Button(tela_principal, text="Iniciar", width=30, height=2, command=mostrar_tela_secundaria).place(x=500,y=550)


# Criando tela secundária
tela_secundaria = tk.Frame(root, bg="DarkGray")

# Label da tela secundária
label_secundaria = tk.Label(tela_secundaria, image=Logo, bg="DarkGray", )
label_secundaria.place(x=220,y=150)

# Botão para iniciar a busca de arquivos
browse_button = tk.Button(tela_secundaria, text="Procurar Arquivos", command=procurar_arquivos, width=25, height=2)
browse_button.place(x=620, y=550)

# Botão para voltar a tela principal
back_button = tk.Button(tela_secundaria, text="Voltar ao Menu Inicial", command=mostrar_tela_principal, width=25, height=2)
back_button.place(x=820, y=550)

# Botão para iniciar a apuração de lucratividade
botao_apuracao_lucra = tk.Button(tela_secundaria, text="Apuracao Lucratividade", width=25, height=2, command=iniciar_apuracao_lucratividade)
botao_apuracao_lucra.place(x=220, y=550)

# Botão para iniciar os relatórios de romaneios executados
botao_romaneio_execu = tk.Button(tela_secundaria, text="Romaneios Executados", width=25, height=2, command=iniciar_romaneios_executados)
botao_romaneio_execu.place(x=420, y=550)

botao_gerar_romaneios_sem_valor = tk.Button(tela_secundaria, text="Romaneios Sem Valor", width=25, height=2, command=inicar_romaneios_sem_valor)
botao_gerar_romaneios_sem_valor.place(x=420,y=600)

# Mostrar tela principal inicialmente
mostrar_tela_principal()

# Mantendo a tela executada
root.bind("<Button-1>", congelar_tela)
root.mainloop()