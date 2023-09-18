import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog, ttk

def atualizar_excel():
    # Solicita ao usuário que escolha o arquivo Excel de origem
    origem_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    
    if origem_path:
        # Solicita ao usuário que escolha o arquivo Excel de destino (relatório)
        destino_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])

        if destino_path:
            # Carregue o arquivo Excel de origem em um dataframe
            sheet_name = 'Planilha1'  # Nome da planilha que você deseja atualizar

            df_origem = pd.read_excel(origem_path, sheet_name=sheet_name)

            # Carregue o arquivo Excel de destino (relatório) em um dataframe
            df_destino = pd.read_excel(destino_path, sheet_name=sheet_name)

            # Obtenha a opção selecionada pelo usuário (Substituir ou Adicionar)
            opcao = opcao_var.get()

            # Obtenha a coluna selecionada pelo usuário
            coluna = coluna_combobox.get()

            if opcao == "Substituir":
                # Substitua os dados na coluna selecionada com base na coluna "Chave"
                df_destino[coluna] = df_origem.set_index('Chave')[coluna]

            elif opcao == "Adicionar":
                # Adicione uma nova coluna ao dataframe de destino com base na coluna de origem
                df_destino[coluna] = df_destino['Chave'].map(df_origem.set_index('Chave')[coluna])

            # Salve o dataframe de destino atualizado de volta para o arquivo Excel de destino
            with pd.ExcelWriter(destino_path, engine='openpyxl', mode='a') as writer:
                writer.book = load_workbook(destino_path)
                writer.sheets = {ws.title: ws for ws in writer.book.worksheets}
                df_destino.to_excel(writer, sheet_name=sheet_name, index=False)

            resultado_label.config(text="Base de relatório atualizada e salva no arquivo Excel de destino.")

# Crie a janela principal da interface gráfica
root = tk.Tk()
root.title("Atualizar Base de Relatório")

# Crie um botão para escolher o arquivo Excel de origem
escolher_origem_button = tk.Button(root, text="Escolher Origem de Dados", command=atualizar_excel)
escolher_origem_button.pack(pady=20)

# Crie uma etiqueta para exibir o resultado
resultado_label = tk.Label(root, text="")
resultado_label.pack()

# Crie caixas de seleção (radiobuttons) para escolher entre Substituir e Adicionar
opcao_var = tk.StringVar()
opcao_var.set("Substituir")  # Padrão selecionado
opcao_substituir = tk.Radiobutton(root, text="Substituir", variable=opcao_var, value="Substituir")
opcao_adicionar = tk.Radiobutton(root, text="Adicionar", variable=opcao_var, value="Adicionar")
opcao_substituir.pack()
opcao_adicionar.pack()

# Crie uma caixa de combinação (combobox) para escolher a coluna
colunas_disponiveis = ['Coluna1', 'Coluna2', 'Coluna3']  # Substitua com suas próprias colunas
coluna_combobox = ttk.Combobox(root, values=colunas_disponiveis)
coluna_combobox.set(colunas_disponiveis[0])  # Padrão selecionado
coluna_combobox.pack()

root.mainloop()
                                
