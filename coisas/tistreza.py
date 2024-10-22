import pyodbc
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

# Função para conectar ao banco de dados
def conectar_banco(database=None):
    # String de conexão com o banco de dados
    conn_str = (
        'DRIVER={SQL Server};'
        'SERVER=PCR-13717\MSSQLSERVER1;'  # Substitua pelo seu servidor
        'Trusted_Connection=yes;'
       # 'UID=;' #Caso utilize um login o usuario é aqui
       # 'PWD=' #Caso utilize um login a senha é aqui
    )
    if database:
        conn_str += f'DATABASE={database};'
    return pyodbc.connect(conn_str)

# Função para consultar e carregar os bancos de dados do servidor
def carregar_bancos():
    try:
        conn = conectar_banco()
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.databases")  # Consulta os bancos de dados
        bancos = [row[0] for row in cursor.fetchall()]  # Armazena os nomes dos bancos em uma lista
        conn.close()
        return bancos
    except Exception as e:
        messagebox.showerror("Erro de Conexão", str(e))  # Exibe erro em caso de falha
        return []

# Função para consultar e carregar as tabelas do banco de dados
def carregar_tabelas(database):
    try:
        conn = conectar_banco(database)
        cursor = conn.cursor()
        cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'")  # Consulta as tabelas
        tabelas = [row[0] for row in cursor.fetchall()]  # Armazena os nomes das tabelas em uma lista
        conn.close()
        return tabelas
    except Exception as e:
        messagebox.showerror("Erro de Conexão", str(e))  # Exibe erro em caso de falha
        return []

# Função para consultar e carregar as colunas da tabela
def carregar_colunas(database, tabela):
    try:
        conn = conectar_banco(database)
        cursor = conn.cursor()
        cursor.execute(f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tabela}'")  # Consulta as colunas da tabela
        colunas = [row[0] for row in cursor.fetchall()]  # Armazena os nomes das colunas em uma lista
        conn.close()
        return colunas
    except Exception as e:
        messagebox.showerror("Erro de Conexão", str(e))  # Exibe erro em caso de falha
        return []

# Função para consultar dados da tabela
def consultar_dados(database, tabela):
    try:
        conn = conectar_banco(database)
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {tabela}")  # Executa a consulta para selecionar todos os dados da tabela
        rows = cursor.fetchall()  # Armazena as linhas retornadas
        
        # Limpar tabela que foi pesquisada 
        limpar_tabela()

        # Carregar colunas da tabela
        colunas = carregar_colunas(database, tabela)

        # Configurar as colunas do Treeview dinamicamente
        tree['columns'] = colunas
        tree['show'] = 'headings'
        
        for col in colunas:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        for row in rows:
            clean_row = [str(col).strip() for col in row]  # Limpa espaços dos dados
            tree.insert("", tk.END, values=clean_row)  # Insere os dados no Treeview

        conn.close()
        
    except Exception as e:
        messagebox.showerror("Erro de Conexão", str(e))  # Exibe erro em caso de falha

# Função para limpar os dados da tabela
def limpar_tabela():
    for item in tree.get_children():
        tree.delete(item)  # Remove todos os itens do Treeview

# Função para executar o atalho
def executar_atalho():
    db_escolhido = db_var.get()
    nome_cod = nome_cod_var.get()
    
    # Conexão com o banco de dados
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=SEU_SERVIDOR;DATABASE=' + db_escolhido + ';UID=SEU_USUARIO;PWD=SUA_SENHA')
    cursor = conn.cursor()
    
    # Comando SQL para obter o código do atalho
    cursor.execute("SELECT cods FROM atalhos WHERE nome_cod = ?", nome_cod)
    resultado = cursor.fetchone()

    if resultado:
        # Executar o código SQL
        codigo_sql = resultado[0]
        try:
            cursor.execute(codigo_sql)
            # Se for uma consulta, obtém os resultados
            if cursor.description:  # Verifica se há resultados
                colunas = [column[0] for column in cursor.description]
                resultados = cursor.fetchall()
                # Formata os resultados
                texto_resultados = '\n'.join([' | '.join(map(str, row)) for row in resultados])
                resultado_texto.delete(1.0, tk.END)  # Limpa a área de texto
                resultado_texto.insert(tk.END, texto_resultados)  # Insere os resultados
            else:
                conn.commit()  # Caso de um comando que não retorna resultados
                resultado_texto.delete(1.0, tk.END)
                resultado_texto.insert(tk.END, "Comando executado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar o atalho: {str(e)}")
    else:
        messagebox.showwarning("Atenção", "Atalho não encontrado.")

# Configuração da interface Tkinter
root = tk.Tk()
root.title("Executar Atalhos")

# Janela principal
root = tk.Tk()
root.title("Sistema de Consultas SQL")

# Treeview para exibir os dados das consultas
tree = ttk.Treeview(root)
tree.pack(pady=20)

# Variáveis
db_var = tk.StringVar()
nome_cod_var = tk.StringVar()

# Interface para seleção do banco de dados e nome do atalho
tk.Label(root, text="Banco de Dados:").pack()
db_entry = tk.Entry(root, textvariable=db_var)
db_entry.pack()

tk.Label(root, text="Nome do Atalho:").pack()
nome_cod_entry = tk.Entry(root, textvariable=nome_cod_var)
nome_cod_entry.pack()

# Botão para executar o atalho
executar_btn = tk.Button(root, text="Executar Atalho", command=executar_atalho)
executar_btn.pack()

# Área de texto para exibir resultados
resultado_texto = tk.Text(root, height=10, width=50)
resultado_texto.pack()

root.mainloop()