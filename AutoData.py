import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import threading

# DataFrame que irá armazenar todas as colunas e linhas
dataframe = None

def escolher_arquivo():
    global dataframe
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        # Carregar o arquivo Excel selecionado em um DataFrame
        dataframe = pd.read_excel(arquivo)
        selecionar_segundo_arquivo()

def selecionar_segundo_arquivo():
    root.withdraw()  # Esconder a janela principal
    arquivo_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo_excel:
        # Ler o arquivo Excel selecionado em um DataFrame
        dataframe_excel = pd.read_excel(arquivo_excel)

        # Substituir colunas correspondentes no DataFrame Excel
        for coluna in dataframe.columns:
            if coluna in dataframe_excel.columns:
                dataframe_excel[coluna] = dataframe[coluna]

        # Criar uma barra de progresso
        progress_bar = ttk.Progressbar(root, length=300, mode='indeterminate')
        progress_bar.pack()
        progress_bar.start()

        # Função para realizar as operações em segundo plano
        def processamento_em_segundo_plano():
            # Simule algum processamento demorado (você pode remover isso)
            import time
            time.sleep(5)

            progress_bar.stop()  # Parar a barra de progresso
            progress_bar.pack_forget()  # Remover a barra de progresso
            messagebox.showinfo("Concluído", "Análise e Modificações concluídas com sucesso!👍")

            # Pedir ao usuário para escolher onde salvar o arquivo resultante
            arquivo_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
            if arquivo_salvar:
                dataframe_excel.to_excel(arquivo_salvar, index=False)

            voltar_ao_inicio()

        # Iniciar uma thread para o processamento em segundo plano
        processing_thread = threading.Thread(target=processamento_em_segundo_plano)
        processing_thread.start()

def voltar_ao_inicio():
    root.deiconify()  # Mostrar a janela principal
    escolher_arquivo()

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1280x720")  # Definir o tamanho da janela principal

    escolher_button = tk.Button(root, text="Escolha a Base Primaria", command=escolher_arquivo)
    escolher_button.pack()

    root.mainloop()
