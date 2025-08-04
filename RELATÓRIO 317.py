import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np


class PlanilhaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tratamento de Planilhas")
        self.root.geometry("420x350")

        self.df = None

        # Botão para selecionar planilha
        self.botao_arquivo = tk.Button(
            root, text="Selecionar Planilha", command=self.selecionar_arquivo)
        self.botao_arquivo.pack(pady=10)

        # Label de status
        self.label_arquivo = tk.Label(root, text="Nenhum arquivo selecionado")
        self.label_arquivo.pack()

        # Botão de tratamento
        self.botao_tratar = tk.Button(
            root, text="Tratar Dados", command=self.tratar_dados, state=tk.DISABLED)
        self.botao_tratar.pack(pady=10)

        # Label para visualização rápida
        self.label_preview = tk.Label(
            root, text="", wraplength=350, justify="left")
        self.label_preview.pack(pady=10)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if caminho:
            try:
                self.df = pd.read_excel(caminho)
                self.label_arquivo.config(
                    text=f"Arquivo: {caminho.split('/')[-1]}")
                self.label_preview.config(
                    text=f"Colunas: {', '.join(self.df.columns[:5])}...")
                self.botao_tratar.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir o arquivo:\n{e}")

    def tratar_dados(self):
        try:
            DF = self.df.copy()

            # Aplicando o seu tratamento
            DF = DF.dropna(axis=1, how='all').dropna(axis=0, how='all')
            DF.reset_index(drop=True, inplace=True)
            DF = DF.iloc[9:].reset_index(drop=True)
            DF = DF.dropna(axis=1, how='all').dropna(axis=0, how='all')
            DF.columns = DF.iloc[0]
            DF = DF[1:].reset_index(drop=True)

            DF['Data'] = DF['Data'].astype(str)
            DF = DF[~DF['Data'].str.contains('TOTAL GERAL:', na=False)]

            DF['Tipo'] = DF['Tipo'].astype(str)
            DF = DF[~DF['Tipo'].str.contains('SubTotal', na=False)]

            DF['Nome'] = DF['Nome'].astype(str)
            DF = DF[~DF['Nome'].str.contains(
                'SubTotal', na=False)].reset_index(drop=True)

            colunas_para_remover = ['Centro de Custo',
                                    '%', 'Conta', 'Boletim', 'Nota', 'Permuta']
            DF = DF.drop(
                columns=[col for col in colunas_para_remover if col in DF.columns])

            DF['Tipo'] = DF['Tipo'].replace('nan', np.nan)
            DF['Nome'] = DF['Nome'].replace('nan', np.nan)
            DF['Tipo'] = DF['Tipo'].fillna(method='ffill')
            DF['Nome'] = DF['Nome'].fillna(method='ffill')

            DF['Data Comp'] = pd.to_datetime(DF['Data Comp'], errors='coerce')
            DF['Data'] = pd.to_datetime(DF['Data'], errors='coerce')
            DF['Data Comp'] = DF['Data Comp'].dt.strftime('%d/%m/%y')
            DF['Data'] = DF['Data'].dt.strftime('%d/%m/%y')

            DF['Classe'] = DF['Nome']
            colunas_finais = ['Tipo', 'Classe', 'Nome', 'Data Comp', 'Data', 'Filial', 'Total',
                              'Referência', 'Histórico', 'Parceiro', 'Cod Desp', 'Parc.', 'Doc.']
            DF = DF[[col for col in colunas_finais if col in DF.columns]]

            # Atualiza a visualização
            self.df = DF

            # Salvar o arquivo tratado
            salvar_em = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if salvar_em:
                self.df.to_excel(salvar_em, index=False)
                messagebox.showinfo(
                    "Arquivo Salvo", f"Arquivo tratado salvo em:\n{salvar_em}")
            else:
                messagebox.showinfo(
                    "Aviso", "Tratamento concluído, mas o arquivo não foi salvo.")

        except Exception as e:
            messagebox.showerror("Erro no tratamento",
                                 f"Ocorreu um erro:\n{e}")


# Executa o app
if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaApp(root)
    root.mainloop()
