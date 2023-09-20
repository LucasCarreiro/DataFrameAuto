import tkinter as tk
from tkinter import filedialog
import openpyxl
import smtplib
from tkinter.simpledialog import askstring

def substituir_colunas(arquivo_principal, arquivo_secundario, planilha_principal_nome, planilha_secundaria_nome):
    planilha_principal = openpyxl.load_workbook(arquivo_principal)
    planilha_secundaria = openpyxl.load_workbook(arquivo_secundario)
    
    planilha_principal_atual = planilha_principal[planilha_principal_nome]
    planilha_secundaria_atual = planilha_secundaria[planilha_secundaria_nome]
    
    # Obtém as colunas da planilha secundária
    colunas_secundarias = list(planilha_secundaria_atual.iter_cols(values_only=True))
    
    # Substitui as colunas correspondentes na planilha principal
    for col_num, coluna_secundaria in enumerate(colunas_secundarias, start=1):
        for row_num, valor in enumerate(coluna_secundaria, start=1):
            planilha_principal_atual.cell(row=row_num, column=col_num, value=valor)
    
    resultado = "Resultado.xlsx"
    planilha_principal.save(resultado)
    return resultado

def selecionar_arquivos():
    arquivo_principal = filedialog.askopenfilename(title="Selecione o arquivo Excel principal")
    arquivo_secundario = filedialog.askopenfilename(title="Selecione o arquivo Excel secundário")
    
    if arquivo_principal and arquivo_secundario:
        planilha_principal_nome = selecionar_planilha(arquivo_principal, "Selecione a planilha principal:")
        planilha_secundaria_nome = selecionar_planilha(arquivo_secundario, "Selecione a planilha secundária:")
        
        resultado = substituir_colunas(arquivo_principal, arquivo_secundario, planilha_principal_nome, planilha_secundaria_nome)
        
        opcao = askstring("Opções", "Escolha a opção:\n1. Enviar por e-mail\n2. Salvar no Computador")
        
        if opcao == "1":
            email_destino = askstring("E-mail", "Digite o e-mail de destino:")
            enviar_email(resultado, email_destino)
        elif opcao == "2":
            salvar_no_computador(resultado)

def selecionar_planilha(arquivo_excel, titulo):
    workbook = openpyxl.load_workbook(arquivo_excel, read_only=True)
    planilhas = workbook.sheetnames
    
    janela = tk.Toplevel(root)
    tk.Label(janela, text=titulo).pack()
    
    planilha_var = tk.StringVar(janela)
    planilha_var.set(planilhas[0])
    
    def confirmar():
        janela.destroy()
    
    tk.OptionMenu(janela, planilha_var, *planilhas).pack()
    tk.Button(janela, text="Confirmar", command=confirmar).pack()
    
    janela.wait_window()
    
    return planilha_var.get()

def enviar_email(anexo, destinatario):
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        remetente = "jamarri833@hungeral.com"  # Coloque seu e-mail aqui
        senha = "d07b9ec5800f67d48810b1e218d049c9"  # Coloque sua senha aqui
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, "Envio de arquivo Excel", anexo)
        server.quit()
        resultado_label.config(text=f"Arquivo de resultado enviado por e-mail para {destinatario}")
    except Exception as e:
        resultado_label.config(text=f"Erro ao enviar e-mail: {str(e)}")

def salvar_no_computador(arquivo):
    destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if destino:
        try:
            import shutil
            shutil.copyfile(arquivo, destino)
            resultado_label.config(text=f"Arquivo de resultado salvo em {destino}")
        except Exception as e:
            resultado_label.config(text=f"Erro ao salvar no computador: {str(e)}")

root = tk.Tk()
root.title("Substituir Colunas Excel")
root.geometry("1280x720")  # Defina a resolução da janela principal

resultado_label = tk.Label(root, text="")
resultado_label.pack()

tk.Button(root, text="Selecionar Arquivos Excel", command=selecionar_arquivos).pack()

root.mainloop()
