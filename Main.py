import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os


def remover_duplicatas():
    # Selecionar o arquivo Excel
    arquivo_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel",
                                               filetypes=[("Excel files", "*.xlsx *.xls")])

    if not arquivo_excel:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return

    # Nome do arquivo de saída com o prefixo "-semduplicatas"
    diretorio_entrada = os.path.dirname(arquivo_excel)
    nome_arquivo, extensao_arquivo = os.path.splitext(os.path.basename(arquivo_excel))
    arquivo_saida = os.path.join(diretorio_entrada, f"{nome_arquivo}-semduplicatas{extensao_arquivo}")

    try:
        # Ler o arquivo Excel
        xls = pd.ExcelFile(arquivo_excel, engine='openpyxl')

        # Iterar sobre cada aba no arquivo Excel
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            for aba in xls.sheet_names:
                # Ler a aba específica
                df = pd.read_excel(arquivo_excel, sheet_name=aba, engine='openpyxl')

                # Remover itens duplicados
                df_sem_duplicatas = df.drop_duplicates()

                # Remover colunas 'Unnamed'
                df_sem_duplicatas = df_sem_duplicatas.loc[:, ~df_sem_duplicatas.columns.str.contains('^Unnamed')]

                # Formatar a coluna "DATA DA VISITA"
                if 'DATA DA VISITA' in df_sem_duplicatas.columns:
                    df_sem_duplicatas['DATA DA VISITA'] = pd.to_datetime(
                        df_sem_duplicatas['DATA DA VISITA']).dt.strftime('%d/%m/%Y')

                # Salvar a aba sem duplicatas no novo arquivo Excel
                df_sem_duplicatas.to_excel(writer, sheet_name=aba, index=False)

        messagebox.showinfo("Sucesso", f"Arquivo sem duplicatas salvo em: {arquivo_saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo: {str(e)}")


# Criar a interface gráfica
root = tk.Tk()
root.title("Remover Duplicatas de Arquivo Excel")

# Tamanho da janela
root.geometry("400x200")

# Centralizar o botão
frame_central = tk.Frame(root)
frame_central.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

# Botão para iniciar o processo
btn_remover_duplicatas = tk.Button(frame_central, text="Selecionar arquivo e remover duplicatas",
                                   command=remover_duplicatas)
btn_remover_duplicatas.pack(pady=20)

# Iniciar a interface
root.mainloop()
