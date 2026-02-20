# Excel Editor — Pandas & Tkinter

Interface desktop para manipulação de dados Excel sem necessidade de fórmulas complexas. Construído com Python, Pandas e Tkinter.

> **Status:** Em desenvolvimento. Funcionalidades core estáveis; novas features em progresso.

![Screenshot](https://github.com/sherloCod3/excel-editor-pandas-tkinter/blob/main/assets/Shot-2024-04-08-094712.png)

---

## Motivação

Usuários não-técnicos precisavam fazer operações como joins, pivots e consolidações em planilhas Excel — tarefas que o Excel nativo torna complexas e propensas a erro. Este projeto expõe o poder do Pandas em uma interface visual acessível.

---

## Funcionalidades

### Manipulação de dados
- Renomear, remover e filtrar colunas
- Remover linhas em branco, duplicadas ou alternadas
- Pivot e agrupamento por uma ou mais colunas

### Merge entre planilhas
- Inner Join, Left Join, Full Join e Outer Merge entre dois DataFrames
- Consolidação de múltiplas planilhas em um único conjunto de dados
- Quebra de dados por critério (split inverso da consolidação)

### Interface
- Visualização tabular via `pandastable`
- Temas Dark e Light
- Salvar resultado em novo arquivo `.xlsx`

---

## Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=flat-square&logo=pandas&logoColor=white)
![NumPy](https://img.shields.io/badge/NumPy-013243?style=flat-square&logo=numpy&logoColor=white)

`Python` · `Pandas` · `NumPy` · `Tkinter` · `pandastable`

---

## Instalação

**Pré-requisito:** Python 3.8+

```bash
git clone https://github.com/sherloCod3/excel-editor-pandas-tkinter.git
cd excel-editor-pandas-tkinter

pip install pandas numpy pandastable
```

---

## Uso

```bash
python Sistema_Pandas.py
```

1. **Arquivo → Carregar Excel** — selecione o `.xlsx` de origem
2. Use o menu **Editar** para transformações na tabela
3. Use o menu **Merge** para cruzar com um segundo arquivo
4. **Arquivo → Salvar Excel** — exporta o resultado

---

## Licença

MIT — veja [LICENSE](./LICENSE).
