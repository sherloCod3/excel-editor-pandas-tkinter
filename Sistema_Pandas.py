"""
Sistema_Pandas.py

Este sistema fornece uma interface gráfica para manipulação de DataFrames utilizando Pandas e Tkinter.
Funcionalidades incluem agrupamento, filtragem, criação de tabelas dinâmicas e salvamento/carregamento de arquivos Excel.

"""

# Importações das bibliotecas necessárias
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
        self.tree = ttk.Treeview(
            self.janela_principal
        )  # Criando a Treeview para exibir os dados
        self.resultado_label = Label(
            self.janela_principal, text="Total: ", font="Arial 16", bg="#F5F5F5"
        )
        self.resultado_label.pack(
            side=TOP, padx=10, pady=10
        )  # Configuração do rótulo de resultado
        self.df = pd.DataFrame()  # Dataframe vazio para armazenar os dados do Excel
        self.cria_widgets()  # Criando os widgets da interface

    def cria_widgets(self):
        """
        Cria os widgets da interface, como menus e configurações da Treeview.
        """
        # Criação da barra de menu
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
        menu_edicao.add_command(label="Renomear Coluna", command=self.renomear_coluna)
        menu_edicao.add_command(label="Remover Coluna", command=self.remover_coluna)

        menu_edicao.add_command(label="Filtrar", command=self.filtrar)
        menu_edicao.add_command(label="Pivot", command=janela.destroy)
        menu_edicao.add_command(label="Group", command=self.group)
        menu_edicao.add_command(
            label="Remover Linhas em Branco", command=self.remover_linhas_em_branco
        )
        menu_edicao.add_command(
            label="Remover Linhas Alternadas", command=self.remover_linhas_selecionadas
        )
        menu_edicao.add_command(
            label="Remover Duplicados", command=self.remover_duplicados
        )
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

        # Menu Temas
        menu_temas = tk.Menu(menu_bar, tearoff=0)
        menu_temas.add_command(label="Material Dark", command=self.aplicar_tema_dark)
        menu_temas.add_command(label="Material Light", command=self.aplicar_tema_light)
        menu_bar.add_cascade(label="Temas", menu=menu_temas)

        # Configurando a barra de menu
        self.janela_principal.config(menu=menu_bar)

        # Criando a Treeview para exibir os dados
        self.tree = tk.ttk.Treeview(self.janela_principal)

        # Configurando a Treeview (Insere o widget de árvore na janela principal)
        self.tree.pack(expand=False)

    def soma_colunas_com_valor(self):
        """
        Calcula a soma das colunas numéricas do DataFrame e exibe o resultado no rótulo de resultado.
        """
        resultados = []

        # Itera sobre cada coluna no DataFrame
        for coluna in self.df.columns:
            # Verifica se a coluna contém dados numéricos
            if pd.api.types.is_numeric_dtype(self.df[coluna]):
                # Seleciona os valores da coluna, excluindo o primeiro elemento (o cabeçalho)
                valores_numericos = self.df[coluna]

                # Converte os valores para tipo numérico, tratando erros com 'coerce'
                valores_numericos = pd.to_numeric(valores_numericos, errors="coerce")

                # Remove os valores NaN (não é um número) da série
                valores_numericos = valores_numericos[~np.isnan(valores_numericos)]

                # Calcula a soma dos valores numéricos da coluna
                soma = valores_numericos.sum()

                # Cria uma string com a mensagem indicando a soma da coluna atual
                resultado = f"A soma da coluna {coluna} é {soma}"

                # Adiciona o resultado à lista de resultados
                resultados.append(resultado)

        # Atualiza o texto do Label com os resultados das somas das colunas
        self.resultado_label.config(text="\n".join(resultados))

    def aplicar_tema_dark(self):
        # Define as cores do tema escuro
        cor_fundo = "#121212"  # Cor de fundo do tema escuro
        cor_texto = "#ecf0f1"  # Cor do texto do tema escuro
        cor_selecao = "#2e2e2e"  # Cor de seleção do tema escuro

        # Aplica as cores aos elementos da interface
        self.janela_principal.config(bg=cor_fundo)  # Cor de fundo da janela principal
        self.resultado_label.config(
            bg=cor_fundo, fg=cor_texto
        )  # Cores do rótulo de resultado
        self.tree.config(
            bg=cor_fundo, fg=cor_texto, selectbackground=cor_selecao
        )  # Cores da Treeview

    def aplicar_tema_light(self):
        # Define as cores do tema claro
        cor_fundo = "#ecf0f1"  # Cor de fundo do tema claro
        cor_texto = "#121212"  # Cor do texto do tema claro
        cor_selecao = "#3498db"  # Cor de seleção do tema claro

        # Aplica as cores aos elementos da interface
        self.janela_principal.config(bg=cor_fundo)  # Cor de fundo da janela principal
        self.resultado_label.config(
            bg=cor_fundo, fg=cor_texto
        )  # Cores do rótulo de resultado
        self.tree.config(
            bg=cor_fundo, fg=cor_texto, selectbackground=cor_selecao
        )  # Cores da Treeview

    def carregar_excel(self):
        """
        Abre uma janela para selecionar o arquivo Excel, lê o arquivo usando Pandas e atualiza a Treeview com os dados.
        """
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        self.nome_do_arquivo = filedialog.askopenfilename(
            title="Selecione o Arquivo", filetypes=tipo_de_arquivo
        )

        try:
            # Tenta ler o arquivo Excel especificado pelo nome_do_arquivo usando o Pandas
            self.df = pd.read_excel(self.nome_do_arquivo)

            # Atualiza a Treeview com o conteúdo do arquivo Excel lido
            self.atualiza_treeview()

            # Calcula a soma das colunas com valores
            self.soma_colunas_com_valor()

        except Exception as e:
            # Se ocorrer algum erro ao abrir o arquivo Excel, exibe uma mensagem de erro
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
        # Itera sobre cada linha do DataFrame e insere os valores na Treeview
        for i, row in self.df.iterrows():
            # Cria uma lista para armazenar os valores da linha
            values = list(row)

            # Itera sobre cada valor na lista de valores
            for j, value in enumerate(values):
                # Verifica se o valor é do tipo np.generic
                if isinstance(value, np.generic):
                    # Converte o valor para o tipo nativo do Python
                    values[j] = np.asscalar(value)

            # Insere os valores processados na Treeview
            self.tree.insert("", tk.END, values=values)

    def renomear_coluna(self):
        """
        Abre uma nova janela para permitir que o usuário renomeie uma coluna do DataFrame.

        Esta função cria uma nova janela (usando Toplevel) que contém widgets para inserir o nome da coluna a ser renomeada
        e o novo nome desejado. Um botão "Renomear" é fornecido para confirmar a ação.
        """
        # Cria uma nova janela para renomear a coluna
        janela_renomear_coluna = tk.Toplevel(self.janela_principal)
        janela_renomear_coluna.title("Renomear Coluna")

        # Configuração da geometria da janela
        largura_janela = 400
        altura_janela = 250
        largura_tela = janela_renomear_coluna.winfo_screenwidth()
        altura_tela = janela_renomear_coluna.winfo_screenheight()
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)
        janela_renomear_coluna.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        # Configurações visuais da janela (bg color)
        janela_renomear_coluna.configure(bg="#FFFFFF")

        # Rótulo e campo de entrada para o nome da coluna atual
        label_coluna = tk.Label(
            janela_renomear_coluna,
            text="Digite o nome da coluna que deseja renomear:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = tk.Entry(
            janela_renomear_coluna, font=("Arial", 12), bg="#FFFFFF"
        )
        entry_coluna.pack()

        # Rótulo e campo de entrada para o novo nome da coluna
        label_novo_nome = tk.Label(
            janela_renomear_coluna,
            text="Digite o novo nome:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_novo_nome.pack(pady=10)
        entry_novo_nome = tk.Entry(
            janela_renomear_coluna, font=("Arial", 12), bg="#FFFFFF"
        )
        entry_novo_nome.pack()

        # Botão para executar a ação de renomear a coluna
        botao_renomear = tk.Button(
            janela_renomear_coluna,
            text="Renomear",
            font=("Arial", 12),
            command=lambda: self.funcao_renomear_coluna(
                entry_coluna.get(), entry_novo_nome.get(), janela_renomear_coluna
            ),
        )
        botao_renomear.pack(pady=20)

        # Cria / exibe na tela
        janela_renomear_coluna.mainloop()

    def funcao_renomear_coluna(self, column, novo_nome, janela_renomear_coluna):
        """
        Renomeia a coluna especificada com o novo nome fornecido pelo usuário.

        Esta função é chamada quando o usuário clica no botão "Renomear". Ela verifica se um novo nome foi fornecido
        e, se sim, renomeia a coluna correspondente no DataFrame. Em seguida, a Treeview é atualizada com os novos dados
        e a janela de renomear coluna é fechada.
        """
        # Verifica se foi fornecido um novo nome para a coluna
        if novo_nome:
            # Renomeia a coluna no DataFrame
            self.df = self.df.rename(columns={column: novo_nome})

            # Atualiza a Treeview com os novos dados do DataFrame
            self.atualiza_treeview()

            # Fecha a janela de renomear coluna após a conclusão da operação
            janela_renomear_coluna.destroy()

    def remover_linhas_em_branco(self):
        # Exibe uma caixa de diálogo com a pergunta "Tem certeza que deseja remover as linhas em branco?"
        resposta = messagebox.askyesno(
            "Remover linhas em branco",
            "Tem certeza que deseja remover as linhas em branco?",
        )

        # Verifica se a resposta é positiva (1)
        if resposta == 1:
            # Remove as linhas com valores ausentes (NaN) do DataFrame self.df
            self.df = self.df.dropna(axis=0)

            # Atualiza a visualização da Treeview com os novos dados do DataFrame
            self.atualiza_treeview()

            # Calcula a soma das colunas que possuem valores
            self.soma_colunas_com_valor()

    def remover_linhas_selecionadas(self):
        """
        Abre uma nova janela para permitir que o usuário remova linhas selecionadas do DataFrame.

        Esta função cria uma nova janela (usando Toplevel) que contém widgets para inserir o número da linha a ser removida.
        Um botão "Remover" é fornecido para confirmar a ação.
        """
        # Cria uma nova janela para remover linhas selecionadas
        janela_remover_linhas_selecionadas = tk.Toplevel(self.janela_principal)
        janela_remover_linhas_selecionadas.title("Remover linhas selecionadas")

        # Configuração da geometria da janela
        largura_janela = 400
        altura_janela = 250
        largura_tela = janela_remover_linhas_selecionadas.winfo_screenwidth()
        altura_tela = janela_remover_linhas_selecionadas.winfo_screenheight()
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)
        janela_remover_linhas_selecionadas.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        # Configurações visuais da janela (cor de fundo)
        janela_remover_linhas_selecionadas.configure(bg="#FFFFFF")

        # Rótulo e campo de entrada para o número da linha a ser removida
        label_linha_inicio = tk.Label(
            janela_remover_linhas_selecionadas,
            text="Digite o número da linha a ser removida:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_linha_inicio.pack(pady=10)
        entry_linha_inicio = tk.Entry(
            janela_remover_linhas_selecionadas, font=("Arial", 12), bg="#FFFFFF"
        )
        entry_linha_inicio.pack()

        # Rótulo e campo de entrada para o novo número da linha a ser removida
        label_linha_fim = tk.Label(
            janela_remover_linhas_selecionadas,
            text="Digite o número da última linha a ser removida:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_linha_fim.pack(pady=10)
        entry_linha_fim = tk.Entry(
            janela_remover_linhas_selecionadas, font=("Arial", 12), bg="#FFFFFF"
        )
        entry_linha_fim.pack()

        # Botão para executar a ação de remover as linhas
        botao_Remover = tk.Button(
            janela_remover_linhas_selecionadas,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_remover_linhas_selecionadas(
                entry_linha_inicio.get(),
                entry_linha_fim.get(),
                janela_remover_linhas_selecionadas,
            ),
        )
        botao_Remover.pack(pady=20)

        # Cria e exibe a janela
        janela_remover_linhas_selecionadas.mainloop()

    def funcao_remover_linhas_selecionadas(
        self, linha_inicio, linha_fim, janela_remover_linhas_selecionadas
    ):
        primeira_linha = int(linha_inicio)
        ultima_linha = int(linha_fim)

        # Remove as linhas do DataFrame
        self.df = self.df.drop(self.df.index[primeira_linha - 1 : ultima_linha])

        # Atualiza a visualização do DataFrame (treeview)
        self.atualiza_treeview()

        # Realiza alguma operação relacionada às colunas com valores
        self.soma_colunas_com_valor()

        # Fecha a janela de remoção de linhas
        janela_remover_linhas_selecionadas.destroy()

    def remover_duplicados(self):
        """
        Abre uma nova janela para permitir que o usuário remova duplicados de uma coluna específica do DataFrame.
        """

        # Cria uma nova janela para remover duplicados
        janela_remover_duplicados = tk.Toplevel(self.janela_principal)
        janela_remover_duplicados.title("Remover Duplicados")

        # Configuração da geometria da janela
        largura_janela = 600
        altura_janela = 250
        largura_tela = janela_remover_duplicados.winfo_screenwidth()
        altura_tela = janela_remover_duplicados.winfo_screenheight()
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)
        janela_remover_duplicados.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        # Configurações visuais da janela (cor de fundo)
        janela_remover_duplicados.configure(bg="#FFFFFF")

        # Rótulo e campo de entrada para o nome da coluna com itens duplicados
        label_coluna = tk.Label(
            janela_remover_duplicados,
            text="Digite o nome da coluna com itens duplicados:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = tk.Entry(
            janela_remover_duplicados, font=("Arial", 12), bg="#FFFFFF"
        )
        entry_coluna.pack()

        # Botão para executar a ação de remover duplicados
        botao_Remover = tk.Button(
            janela_remover_duplicados,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_remover_duplicados(
                entry_coluna.get(),
                janela_remover_duplicados,
            ),
        )
        botao_Remover.pack(pady=20)

        # Cria e exibe a janela
        janela_remover_duplicados.mainloop()

    def funcao_remover_duplicados(self, coluna, janela_remover_duplicados):
        """
        Remove itens duplicados na coluna especificada do DataFrame.

        Args:
            coluna (str): Nome da coluna em que duplicados serão removidos.
            janela_remover_duplicados (Toplevel): Janela de interface para remoção de duplicados.
        """
        # Verifica se o usuário digitou uma coluna
        if coluna:
            # Remove os itens duplicados na coluna, mantendo apenas a primeira ocorrência
            self.df = self.df.drop_duplicates(subset=coluna, keep="first")

        # Atualiza a visualização do DataFrame (treeview)
        self.atualiza_treeview()

        # Calcula e exibe a soma das colunas com valores
        self.soma_colunas_com_valor()

        # Fecha a janela de remoção de duplicados
        janela_remover_duplicados.destroy()

    def remover_coluna(self):
        """
        Abre uma nova janela para permitir que o usuário remova duplicados de uma coluna específica do DataFrame.
        """

        # Cria uma nova janela para remover duplicados
        janela_remover_coluna = tk.Toplevel(self.janela_principal)
        janela_remover_coluna.title("Remover Coluna")

        # Configuração da geometria da janela
        largura_janela = 600
        altura_janela = 250
        largura_tela = janela_remover_coluna.winfo_screenwidth()
        altura_tela = janela_remover_coluna.winfo_screenheight()
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)
        janela_remover_coluna.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        # Configurações visuais da janela (cor de fundo)
        janela_remover_coluna.configure(bg="#FFFFFF")

        # Rótulo e campo de entrada para o nome da coluna
        label_coluna = tk.Label(
            janela_remover_coluna,
            text="Digite o nome da coluna a ser removida:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = tk.Entry(janela_remover_coluna, font=("Arial", 12), bg="#FFFFFF")
        entry_coluna.pack()

        # Botão para executar a ação de remover coluna
        botao_Remover = tk.Button(
            janela_remover_coluna,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_remover_coluna(
                entry_coluna.get(),
                janela_remover_coluna,
            ),
        )
        botao_Remover.pack(pady=20)

        # Cria e exibe a janela
        janela_remover_coluna.mainloop()

    def funcao_remover_coluna(self, coluna, janela_remover_coluna):
        """
        Remove itens duplicados na coluna especificada do DataFrame.

        Args:
            coluna (str): Nome da coluna em que duplicados serão removidos.
            janela_remover_coluna (Toplevel): Janela de interface para remoção de duplicados.
        """
        # Verifica se o usuário digitou uma coluna
        if coluna:

            # Remove a coluna selecionada
            self.df = self.df.drop(columns=coluna)

        # Atualiza a visualização do DataFrame (treeview)
        self.atualiza_treeview()

        # Calcula e exibe a soma das colunas com valores
        self.soma_colunas_com_valor()

        # Fecha a janela de remoção de duplicados
        janela_remover_coluna.destroy()

    def filtrar(self):
        """
        Abre uma nova janela para permitir que o usuário remova duplicados de uma coluna específica do DataFrame.
        """

        # Cria uma nova janela para remover duplicados
        janela_filtrar = tk.Toplevel(self.janela_principal)
        janela_filtrar.title("Filtrar")

        # Configuração da geometria da janela
        largura_janela = 600
        altura_janela = 250
        largura_tela = janela_filtrar.winfo_screenwidth()
        altura_tela = janela_filtrar.winfo_screenheight()
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)
        janela_filtrar.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        # Configurações visuais da janela (cor de fundo)
        janela_filtrar.configure(bg="#FFFFFF")

        # Rótulo e campo de entrada para o nome da coluna
        label_coluna = tk.Label(
            janela_filtrar,
            text="Digite o nome da coluna a ser filtrada:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = tk.Entry(janela_filtrar, font=("Arial", 12), bg="#FFFFFF")
        entry_coluna.pack()

        label_valor = tk.Label(
            janela_filtrar,
            text="Digite o valor a ser filtrado:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_valor.pack(pady=10)
        entry_valor = tk.Entry(janela_filtrar, font=("Arial", 12), bg="#FFFFFF")
        entry_valor.pack()

        # Botão para executar a ação de remover coluna
        botao_Remover = tk.Button(
            janela_filtrar,
            text="Remover",
            font=("Arial", 12),
            command=lambda: self.funcao_filtrar(
                entry_coluna.get(),
                entry_valor.get(),
                janela_filtrar,
            ),
        )
        botao_Remover.pack(pady=20)

        # Cria e exibe a janela
        janela_filtrar.mainloop()

    def funcao_filtrar(self, coluna, valor, janela_filtrar):
        """
        Remove itens duplicados na coluna especificada do DataFrame.

        Args:
            coluna (str): Nome da coluna em que duplicados serão removidos.
            janela_filtrar (Toplevel): Janela de interface para remoção de duplicados.
        """
        # Verifica se o usuário digitou uma coluna
        if coluna:

            # Verifica se o usuário digitou um valor
            if valor:

                # Filtra o dataframe com base na coluna e valor selecionados
                self.df = self.df[self.df[coluna] == valor]

        # Atualiza a visualização do DataFrame (treeview)
        self.atualiza_treeview()

        # Calcula e exibe a soma das colunas com valores
        self.soma_colunas_com_valor()

        # Fecha a janela de remoção de duplicados
        janela_filtrar.destroy()

    def group(self):
        """
        Abre uma nova janela para permitir que o usuário remova duplicados de uma coluna específica do DataFrame.
        """

        # Cria uma nova janela para remover duplicados
        janela_group = tk.Toplevel(self.janela_principal)
        janela_group.title("Agrupar")

        # Configuração da geometria da janela
        largura_janela = 600
        altura_janela = 250
        largura_tela = janela_group.winfo_screenwidth()
        altura_tela = janela_group.winfo_screenheight()
        posicao_x = (largura_tela // 2) - (largura_janela // 2)
        posicao_y = (altura_tela // 2) - (altura_janela // 2)
        janela_group.geometry(
            f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}"
        )

        # Configurações visuais da janela (cor de fundo)
        janela_group.configure(bg="#FFFFFF")

        # Rótulo e campo de entrada para o nome da coluna
        label_coluna = tk.Label(
            janela_group,
            text="Digite o nome da coluna a ser agrupada:",
            font=("Arial", 12),
            bg="#FFFFFF",
        )
        label_coluna.pack(pady=10)
        entry_coluna = tk.Entry(janela_group, font=("Arial", 12), bg="#FFFFFF")
        entry_coluna.pack()

        # Botão para executar a ação de remover coluna
        botao_agrupar = tk.Button(
            janela_group,
            text="Agrupar",
            font=("Arial", 12),
            command=lambda: self.funcao_agrupar(
                entry_coluna.get(),
                janela_group,
            ),
        )
        botao_agrupar.pack(pady=20)

        # Cria e exibe a janela
        janela_group.mainloop()

    def funcao_agrupar(self, coluna, janela_group):

        # Limpa os dados da treeview
        self.tree.delete(*self.tree.get_children())

        if coluna:

            dadosAgrupados = self.df.groupby(coluna).sum()

            # for
            for i, linha in dadosAgrupados.iterrows():

                values = list(linha)

                for j, value in enumerate(values):

                    if isinstance(value, np.generic):

                        values[j] = np.asscalar(value)

                # Inserindo linhas na treeview
                self.tree.insert("", tk.END, values=[i] + values)

                # Fecha a janela secundária
                janela_group.destroy()

    def merge_inner_join(self):

        # Define os tipos de arquivos que serão selecionados para a função Inner Join
        tipo_de_arquivo = (("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))

        # Abre a janela de seleção de arquivos e armazena o primeiro arquivo em uma variável
        nome_do_arquivo_1 = filedialog.askopenfilename(
            title="Selecione o primeiro arquivo", filetypes=tipo_de_arquivo
        )

        # Abre a janela de seleção de arquivos e armazena o segundo arquivo em uma variável
        nome_do_arquivo_2 = filedialog.askopenfilename(
            title="Selecione o segundo arquivo", filetypes=tipo_de_arquivo
        )

        # Lê os arquivos com extensões de planilhas
        arquivo_1 = pd.read_excel(nome_do_arquivo_1)
        arquivo_2 = pd.read_excel(nome_do_arquivo_2)

        coluna_join = simpledialog.askstring(
            "Coluna do Inner Join",
            "Digite o nome da coluna que será utilizada para o Inner Join",
        )

        # On - qual coluna
        # How - tipo de Join
        # Procura e exibe os vendedores que estão em ambas as tabelas
        self.df = pd.merge(arquivo_1, arquivo_2, on=coluna_join, how="inner")

        # Atualiza a Treeview com o resultado do merge
        self.atualiza_treeview()


# Instancia a classe ExcelEditor passando a janela principal como parâmetro
editor = ExcelEditor(janela)

# Inicia o loop principal da interface gráfica
janela.mainloop()
