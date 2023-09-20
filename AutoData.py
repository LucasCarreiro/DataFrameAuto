import tkinter as tk
from tkinter import filedialog
import pandas as pd

# DataFrame que irá armazenar todas as colunas e linhas
dataframe = None

def escolher_arquivo():
    global dataframe
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        # Carregar o arquivo Excel selecionado em um DataFrame
        dataframe = pd.read_excel(arquivo)
        abrir_arquivo_excel()

def abrir_arquivo_excel():
    if dataframe is None:
        return

    arquivo_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo_excel:
        # Ler o arquivo Excel selecionado em um DataFrame
        dataframe_excel = pd.read_excel(arquivo_excel)

        # Substituir colunas correspondentes no DataFrame Excel
        for coluna in dataframe.columns:
            if coluna in dataframe_excel.columns:
                dataframe_excel[coluna] = dataframe[coluna]

        # Salvar o DataFrame Excel resultante em um novo arquivo
        arquivo_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        if arquivo_salvar:
            dataframe_excel.to_excel(arquivo_salvar, index=False)

        voltar_ao_inicio()

def voltar_ao_inicio():
    escolher_arquivo()

if __name__ == "__main__":
    root = tk.Tk()
    escolher_button = tk.Button(root, text="Escolha a Base Primaria", command=escolher_arquivo)
    escolher_button.pack()
    root.mainloop()
