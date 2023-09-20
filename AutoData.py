import tkinter as tk
from tkinter import filedialog, ttk
import openpyxl
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
        workbook = openpyxl.load_workbook(arquivo_excel)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]

            for coluna in dataframe.columns:
                if coluna in sheet[1]:
                    coluna_excel = sheet[1][coluna]
                    for i, valor in enumerate(dataframe[coluna], start=2):
                        coluna_excel[i].value = valor

        arquivo_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        if arquivo_salvar:
            workbook.save(arquivo_salvar)

if __name__ == "__main__":
    root = tk.Tk()
    escolher_button = tk.Button(root, text="Escolha o arquivo com colunas e linhas", command=escolher_arquivo)
    escolher_button.pack()
    root.mainloop()
