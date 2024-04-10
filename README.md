## Projeto em Construção 🚧
<br>

> [!CAUTION]
> ⚠️ Este projeto está em desenvolvimento ativo.
> Algumas funcionalidades podem estar incompletas ou sujeitas a mudanças. ⚠️

<br>

# 📊 Editor de Excel com Pandas

Este é um aplicativo simples de editor de Excel usando Python, Tkinter e Pandas. Ele permite abrir arquivos Excel, exibir os dados em uma interface gráfica e oferece algumas opções básicas de edição.

---

## Funcionalidades

### Arquivo:
- **Carregar Excel:** Abre uma janela para selecionar um arquivo Excel, lê o arquivo usando Pandas e atualiza a Treeview com os dados.
- **Salvar Excel:** Salva o DataFrame atual em um novo arquivo Excel.

### Editar:
- **Renomear Coluna:** Abre uma nova janela para permitir que o usuário renomeie uma coluna do DataFrame.
- **Remover Coluna:** Remove a coluna selecionada do DataFrame.
- **Filtrar:** Abre uma janela para permitir que o usuário filtre os dados com base em critérios específicos.
- **Pivot:** Abre uma janela para permitir que o usuário faça uma operação de pivotamento nos dados.
- **Group:** Abre uma janela para permitir que o usuário agrupe os dados com base em uma ou mais colunas.
- **Remover Linhas em Branco:** Remove as linhas em branco do DataFrame.
- **Remover Linhas Alternadas:** Remove linhas alternadas do DataFrame.
- **Remover Duplicados:** Remove linhas duplicadas do DataFrame.

### Merge:
- **Inner Join:** Realiza uma operação de inner join em dois DataFrames.
- **Join Full:** Realiza uma operação de full join em dois DataFrames.
- **Left Join:** Realiza uma operação de left join em dois DataFrames.
- **Merge Outer:** Realiza uma operação de outer merge em dois DataFrames.

### Relatórios:
- **Consolidar:** Abre uma janela para permitir que o usuário consolide os dados de várias planilhas ou arquivos em um único conjunto de dados.
- **Quebra:** Abre uma janela para permitir que o usuário divida os dados com base em critérios específicos.

### Temas:
- **Aplicar Tema Dark:** Aplica um tema escuro à interface.
- **Aplicar Tema Light:** Aplica um tema claro à interface.

---

## Requisitos de Instalação

Para executar este aplicativo, você precisa ter o Python instalado no seu sistema. Você pode baixá-lo em [python.org](https://www.python.org/downloads/).

Além disso, instale as seguintes bibliotecas Python:

```bash
pip install pandas
pip install numpy
pip install pandastable
```

## Como usar 🚀
1. **Clone o repositório:** 
```git clone https://github.com/sherloCod3/excel-editor-pandas-tkinter.git```

2. **Execute o programa:**
```python Sistema_Pandas.py```

3. **Abra um arquivo Excel:**
Clique em "Arquivo" no menu e selecione "Abrir".
Escolha o arquivo Excel que deseja editar e clique em "Abrir".

4. **Edite o Excel:**
Utilize as opções de edição disponíveis no menu.

---

## Recursos 🛠️
Abrir arquivos Excel: Importe arquivos Excel para editar.
Editar colunas: Renomeie, remova e filtre colunas.
Ferramentas de merge: Realize operações de merge como Inner Join, Join Full, Left Join e Merge Outer.
Geração de relatórios: Crie relatórios consolidados e quebras de dados.

---

## Screenshot

Aqui está uma captura de tela do aplicativo em ação:

![Screenshot do Aplicativo](https://github.com/sherloCod3/excel-editor-pandas-tkinter/blob/main/assets/Shot-2024-04-08-094712.png)

---

## Contribuindo 🤝
Sinta-se à vontade para contribuir com novos recursos ou correções de bugs.
Faça um fork do projeto, implemente suas mudanças e envie uma pull request.

---

## Licença 📝
Este projeto está licenciado sob a Licença MIT - veja o arquivo LICENSE para detalhes.
