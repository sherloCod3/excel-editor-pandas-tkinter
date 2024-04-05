import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox

import numpy as np
import pandas as pd

# Criando a janela principal
janela = Tk()
janela.title("Editor de Excel com Pandas")
janela.attributes("-fullscreen", True)


class ExcelEditor:
    def __init__(self, janela_principal):
        """
        Inicializa a classe com uma instância da janela principal, cria uma Treeview para exibir os dados do Excel
        e inicializa outras variáveis necessárias.
        """
        self.janela_principal = janela_principal
        self.tree = ttk.Treeview(self.janela_principal)  # Criando a Treeview para exibir os dados
        self.resultado_label = Label(self.janela_principal, text="Total: ", font="Arial 16", bg="#F5F5F5")
        self.df = pd.DataFrame()  # Dataframe vazio para armazenar os dados do Excel
        self.cria_widgets()  # Criando os widgets da interface

    def cria_widgets(self):
        """
        Cria os widgets da interface, como menus e configurações da Treeview.
        """
        menu_bar = tk.Menu(self.janela_principal)

        # Menu Arquivo
        menu_de_arquivos = tk.Menu(menu_bar, tearoff=0)
        menu_de_arquivos.add_command(label="Abrir", command=self.carregar_excel)
        menu_de_arquivos.add_separator()
        menu_de_arquivos.add_command(label="Salvar Como", command=janela.destroy)
        menu_de_arquivos.add_separator()
        menu_de_arquivos.add_command(label="Sair", command=janela.destroy)
        menu_bar.add_cascade(label="Arquivo", menu=menu_de_arquivos)

        # Menu Editar
        menu_edicao = tk.Menu(menu_bar, tearoff=0)
        menu_edicao.add_command(label="Renomear Coluna", command=janela.destroy)
        menu_edicao.add_command(label="Remover Coluna", command=janela.destroy)
        menu_edicao.add_command(label="Filtrar", command=janela.destroy)
        menu_edicao.add_command(label="Pivot", command=janela.destroy)
        menu_edicao.add_command(label="Group", command=janela.destroy)
        menu_edicao.add_command(label="Remover Linhas em Branco", command=janela.destroy)
        menu_edicao.add_command(label="Remover Linhas Alternadas", command=janela.destroy)
        menu_edicao.add_command(label="Remover Duplicados", command=janela.destroy)
        menu_bar.add_cascade(label="Editar", menu=menu_edicao)

        # Menu Merge
        merge_menu = tk.Menu(menu_bar, tearoff=0)
        merge_menu.add_command(label="Inner Join", command=janela.destroy)
        merge_menu.add_command(label="Join Full", command=janela.destroy)
        merge_menu.add_command(label="Left Join", command=janela.destroy)
        merge_menu.add_command(label="Merge Outer", command=janela.destroy)
        menu_bar.add_cascade(label="Merge", menu=merge_menu)

        # Menu Relatórios
        relatorio_menu = tk.Menu(menu_bar, tearoff=0)
        relatorio_menu.add_command(label="Consolidar", command=janela.destroy)
        relatorio_menu.add_command(label="Quebra", command=janela.destroy)
        menu_bar.add_cascade(label="Relatórios", menu=relatorio_menu)

        # Configurando a barra de menu
        self.janela_principal.config(menu=menu_bar)

        # Configurando a Treeview
        self.tree.pack(expand=False)

    def carregar_excel(self):
        """
        Abre uma janela para selecionar o arquivo Excel, lê o arquivo usando Pandas e atualiza a Treeview com os dados.
        """
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        self.nome_do_arquivo = filedialog.askopenfilename(title="Selecione o Arquivo", filetypes=tipo_de_arquivo)

        try:
            self.df = pd.read_excel(self.nome_do_arquivo)
            self.atualiza_treeview()
        except Exception as e:
            messagebox.showerror("Erro!", f"Não foi possível abrir o arquivo: {e}")

    def atualiza_treeview(self):
        """
        Atualiza a Treeview com os dados do DataFrame.
        """
        self.tree.delete(*self.tree.get_children())  # Limpa a Treeview

        # Define as colunas da Treeview com base nas colunas do DataFrame
        self.tree["columns"] = list(self.df.columns)
        for column in self.df.columns:
            self.tree.heading(column, text=column)  # Define os cabeçalhos das colunas

        # Insere os dados do DataFrame na Treeview
        for _, row in self.df.iterrows():
            values = [np.asscalar(value) if isinstance(value, np.generic) else value for value in row]
            self.tree.insert("", tk.END, values=values)


# Instancia a classe ExcelEditor passando a janela principal como parâmetro
editor = ExcelEditor(janela)

# Inicia o loop principal da interface gráfica
janela.mainloop()
