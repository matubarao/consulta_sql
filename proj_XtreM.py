#O projeto é pensado para a escalabilidade, e não abrange muitas funcionalidades 
#mas o que for necessário consegue-se fazer na parte de atalhos. leia 253.

import pyodbc
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import logging

# Configurando o logger
logging.basicConfig(filename='logs.txt',  # Define o arquivo de log
                    level=logging.INFO,    # Define o nível de log
                    format='%(asctime)s - %(levelname)s - %(message)s')  # Formato do log

# Função de exemplo para registrar logs
def registrar_log(mensagem):
    logging.info(mensagem)  # Registra uma mensagem de informação

# Uso do logger
registrar_log("Aqui ficam os logs")

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

# Função para abrir a nova janela de consulta geral
def abrir_janela_consulta_geral():
    consulta_geral_janewlas = tk.Toplevel(root)
    consulta_geral_janewlas.title("Consulta Geral")  # Título da nova janela
    consulta_geral_janewlas.attributes("-fullscreen", True)  # Tamanho da nova janela
    
    # Carregar bancos de dados
    bancos = carregar_bancos()
    
    tk.Label(consulta_geral_janewlas, text="Selecione o Banco de Dados:").pack(pady=5)
    banco_var = tk.StringVar(consulta_geral_janewlas)
    banco_dropdown = ttk.Combobox(consulta_geral_janewlas, textvariable=banco_var, values=bancos, state="readonly")
    banco_dropdown.pack(pady=5)

    # Label e combobox para selecionar tabela
    tk.Label(consulta_geral_janewlas, text="Selecione a Tabela:").pack(pady=5)
    tabela_var = tk.StringVar(consulta_geral_janewlas)
    tabela_dropdown = ttk.Combobox(consulta_geral_janewlas, textvariable=tabela_var, state="readonly")
    tabela_dropdown.pack(pady=5)

    # Função para atualizar as tabelas ao selecionar um banco de dados
    def atualizar_tabelas(event):
        banco_selecionado = banco_var.get()
        tabelas = carregar_tabelas(banco_selecionado)  # Carrega as tabelas do banco selecionado
        tabela_dropdown['values'] = tabelas  # Atualiza as opções da tabela
        tabela_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de banco à atualização de tabelas
    banco_dropdown.bind("<<ComboboxSelected>>", atualizar_tabelas)

    # Botão para executar a consulta geral
    def executar_consulta_geral():
        banco = banco_var.get()
        tabela = tabela_var.get()

        if not banco or not tabela:
            messagebox.showerror("Erro", "Selecione um banco de dados e uma tabela!")  # Verifica se os campos estão preenchidos
            return

        consultar_dados(banco, tabela)  # Executa a consulta
        consulta_geral_janewlas.destroy()  # Fechar a janela após a consulta

    btn_executar_consulta_geral = tk.Button(consulta_geral_janewlas, text="Executar Consulta Geral", command=executar_consulta_geral)
    btn_executar_consulta_geral.pack(pady=10)

# Função para abrir a nova janela de consulta detalhada
def abrir_janela_consulta_detalhada():
    consulta_detalhada_window = tk.Toplevel(root)
    consulta_detalhada_window.title("Consulta Detalhada")  # Título da nova janela
    consulta_detalhada_window.attributes("-fullscreen", True)  # Tamanho da nova janela
    
    # Carregar bancos de dados
    bancos = carregar_bancos()
    
    tk.Label(consulta_detalhada_window, text="Selecione o Banco de Dados:").pack(pady=5)
    banco_var = tk.StringVar(consulta_detalhada_window)
    banco_dropdown = ttk.Combobox(consulta_detalhada_window, textvariable=banco_var, values=bancos, state="readonly")
    banco_dropdown.pack(pady=5)

    # Label e combobox para selecionar tabela
    tk.Label(consulta_detalhada_window, text="Selecione a Tabela:").pack(pady=5)
    tabela_var = tk.StringVar(consulta_detalhada_window)
    tabela_dropdown = ttk.Combobox(consulta_detalhada_window, textvariable=tabela_var, state="readonly")
    tabela_dropdown.pack(pady=5)

    # Função para atualizar as tabelas ao selecionar um banco de dados
    def atualizar_tabelas_detalhadas(event):
        banco_selecionado = banco_var.get()
        tabelas = carregar_tabelas(banco_selecionado)  # Carrega as tabelas do banco selecionado
        tabela_dropdown['values'] = tabelas  # Atualiza as opções da tabela
        tabela_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de banco à atualização de tabelas
    banco_dropdown.bind("<<ComboboxSelected>>", atualizar_tabelas_detalhadas)

    # Label e combobox para selecionar a coluna do WHERE
    tk.Label(consulta_detalhada_window, text="Selecione a Coluna (WHERE):").pack(pady=5)
    coluna_var = tk.StringVar(consulta_detalhada_window)
    coluna_dropdown = ttk.Combobox(consulta_detalhada_window, textvariable=coluna_var, state="readonly")
    coluna_dropdown.pack(pady=5)

    # Label e caixa de texto para o valor do WHERE
    tk.Label(consulta_detalhada_window, text="Valor para o WHERE:").pack(pady=5)
    valor_where_entry = tk.Entry(consulta_detalhada_window)
    valor_where_entry.pack(pady=5)

    # Função para atualizar as colunas ao selecionar uma tabela
    def atualizar_colunas_detalhadas(event):
        banco_selecionado = banco_var.get()
        tabela_selecionada = tabela_var.get()
        colunas = carregar_colunas(banco_selecionado, tabela_selecionada)  # Carrega as colunas da tabela selecionada
        coluna_dropdown['values'] = colunas  # Atualiza as opções da coluna
        coluna_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de tabela à atualização de colunas
    tabela_dropdown.bind("<<ComboboxSelected>>", atualizar_colunas_detalhadas)

    # Botão para executar a consulta detalhada
    def executar_consulta_detalhada():
        banco = banco_var.get()
        tabela = tabela_var.get()
        coluna = coluna_var.get()
        valor_where = valor_where_entry.get()

        if not banco or not tabela or not coluna or not valor_where:
            messagebox.showerror("Erro", "Preencha todos os campos!")  # Verifica se todos os campos estão preenchidos
            return

        try:
            conn = conectar_banco(banco)
            cursor = conn.cursor()
            consulta_sql = f"SELECT * FROM {tabela} WHERE {coluna} = ?"  # Consulta com cláusula WHERE
            cursor.execute(consulta_sql, (valor_where,))  # Executa a consulta com o valor do WHERE
            rows = cursor.fetchall()  # Armazena as linhas retornadas
            
            # Limpar tabela antes de carregar novos dados
            limpar_tabela()

            # Configurar colunas dinamicamente no Treeview
            colunas = carregar_colunas(banco, tabela)
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

        consulta_detalhada_window.destroy()  # Fecha a janela de consulta detalhada após a execução

    btn_executar_consulta_detalhada = tk.Button(consulta_detalhada_window, text="Executar Consulta Detalhada", command=executar_consulta_detalhada)
    btn_executar_consulta_detalhada.pack(pady=10)


### Para o funcionamento desta parte, necessario criar a tabela dbo.atalhos com as colunas nome_cod, cods.###

# Função para abrir a nova janela de atalhos
def abrir_janela_atalhos():
    atalhos_window = tk.Toplevel(root)
    atalhos_window.title("Criar Atalhos")  # Título da nova janela
    atalhos_window.attributes("-fullscreen", True)  # Tamanho da nova janela

    # Label e caixa de texto para o nome do atalho
    tk.Label(atalhos_window, text="Nome do Atalho:").pack(pady=5)
    nome_atalho_entry = tk.Entry(atalhos_window)
    nome_atalho_entry.pack(pady=5)

    # Carregar bancos de dados
    bancos = carregar_bancos()
    
    tk.Label(atalhos_window, text="Selecione o Banco de Dados:").pack(pady=5)
    banco_var = tk.StringVar(atalhos_window)
    banco_dropdown = ttk.Combobox(atalhos_window, textvariable=banco_var, values=bancos, state="readonly")
    banco_dropdown.pack(pady=5)

    # Label e caixa de texto para o código SQL
    tk.Label(atalhos_window, text="Código SQL:").pack(pady=5)
    codigo_sql_entry = tk.Text(atalhos_window, height=10, width=40)
    codigo_sql_entry.pack(pady=5)

    # Função para salvar o atalho no banco de dados
    def salvar_atalho():
        nome_atalho = nome_atalho_entry.get()
        banco = banco_var.get()
        codigo_sql = codigo_sql_entry.get("1.0", tk.END).strip()

        if not nome_atalho or not banco or not codigo_sql:
            messagebox.showerror("Erro", "Preencha todos os campos!")  # Verifica se os campos estão preenchidos
            return

        try:
            conn = conectar_banco(banco)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO atalhos (nome_cod, cods) VALUES (?, ?)", (nome_atalho, codigo_sql))  # Insere o atalho no banco
            conn.commit()
            conn.close()
            messagebox.showinfo("Sucesso", "Atalho salvo com sucesso!")
            atalhos_window.destroy()  # Fecha a janela após salvar o atalho

        except Exception as e:
            messagebox.showerror("Erro", str(e))  # Exibe erro em caso de falha

    btn_salvar_atalho = tk.Button(atalhos_window, text="Salvar Atalho", command=salvar_atalho)
    btn_salvar_atalho.pack(pady=10)

# Função para abrir a janela de execução de atalhos
def abrir_janela_executar_atalho():
    executar_atalho_window = tk.Toplevel(root)
    executar_atalho_window.title("Executar Atalhos")
    executar_atalho_window.attributes("-fullscreen", True)  # Tamanho da nova janela

    # Label e combobox para selecionar o banco de dados
    tk.Label(executar_atalho_window, text="Selecione o Banco de Dados:").pack(pady=5)
    banco_var = tk.StringVar(executar_atalho_window)
    bancos = carregar_bancos()
    banco_dropdown = ttk.Combobox(executar_atalho_window, textvariable=banco_var, values=bancos, state="readonly")
    banco_dropdown.pack(pady=5)

    # Label e combobox para selecionar o atalho
    tk.Label(executar_atalho_window, text="Selecione o Atalho:").pack(pady=5)
    atalho_var = tk.StringVar(executar_atalho_window)
    atalho_dropdown = ttk.Combobox(executar_atalho_window, textvariable=atalho_var, state="readonly")
    atalho_dropdown.pack(pady=5)

    # Função para atualizar os atalhos ao selecionar um banco de dados
    def atualizar_atalhos(event):
        banco_selecionado = banco_var.get()
        atalhos = carregar_atalhos(banco_selecionado)  # Função que carrega os atalhos do banco selecionado
        atalho_dropdown['values'] = [a[0] for a in atalhos]  # Atualiza os atalhos com os nomes dos atalhos
        atalho_dropdown.set('')

    # Vincular a mudança de banco à atualização de atalhos
    banco_dropdown.bind("<<ComboboxSelected>>", atualizar_atalhos)

    # Label e caixa de texto para mostrar e permitir a edição do código SQL do atalho
    tk.Label(executar_atalho_window, text="Código SQL:").pack(pady=5)
    codigo_sql_text = tk.Text(executar_atalho_window, height=10, width=40)
    codigo_sql_text.pack(pady=5)

    # Função para carregar o código SQL ao selecionar um atalho
    def carregar_codigo_atalho(event):
        banco_selecionado = banco_var.get()
        atalho_selecionado = atalho_var.get()
        codigo_sql = buscar_codigo_atalho(banco_selecionado, atalho_selecionado)  # Função que carrega o SQL do atalho
        codigo_sql_text.delete("1.0", tk.END)
        codigo_sql_text.insert(tk.END, codigo_sql)  # Insere o código SQL na caixa de texto

    # Vincular a mudança de atalho à exibição do código SQL
    atalho_dropdown.bind("<<ComboboxSelected>>", carregar_codigo_atalho)

    # Função para executar o código SQL do atalho
    def executar_atalho():
        banco_selecionado = banco_var.get()
        codigo_sql = codigo_sql_text.get("1.0", tk.END).strip()

        if not banco_selecionado or not codigo_sql:
            messagebox.showerror("Erro", "Selecione um banco de dados e um atalho válido!")
            return

        try:
            conn = conectar_banco(banco_selecionado)
            cursor = conn.cursor()
            
            # Executa o código SQL do atalho
            cursor.execute(codigo_sql)  

            # Se o cursor tem descrição, é uma consulta
            if cursor.description:  
                rows = cursor.fetchall()

                # Limpar tabela antes de carregar novos dados
                limpar_tabela()

                # Configurar colunas dinamicamente no Treeview
                colunas = [desc[0] for desc in cursor.description]  # Pega os nomes das colunas
                tree['columns'] = colunas
                tree['show'] = 'headings'

                for col in colunas:
                    tree.heading(col, text=col)
                    tree.column(col, width=100)

                for row in rows:
                    clean_row = [str(col).strip() for col in row]  # Limpa espaços dos dados
                    tree.insert("", tk.END, values=clean_row)
            else:
                # Se não houver descrição, é uma execução de INSERT, UPDATE ou DELETE
                conn.commit()  # Confirma as mudanças no banco
                messagebox.showinfo("Sucesso", "Atalho executado com sucesso, mas não retornou resultados.")

            conn.close()

        except Exception as e:
                registrar_log(f"Erro ao executar atalho: {str(e)}")

    # Botão para executar o atalho
    btn_executar_atalho = tk.Button(executar_atalho_window, text="Executar Atalho", command=executar_atalho)
    btn_executar_atalho.pack(pady=10)

# Função para carregar os atalhos a partir do banco de dados
def carregar_atalhos(banco):
    try:
        conn = conectar_banco(banco)
        cursor = conn.cursor()
        cursor.execute("SELECT nome_cod FROM atalhos")
        atalhos = cursor.fetchall()
        conn.close()
        return atalhos
    except Exception as e:
        messagebox.showerror("Erro ao carregar atalhos", str(e))
        return []

# Função para buscar o código SQL de um atalho específico
def buscar_codigo_atalho(banco, nome_atalho):
    try:
        conn = conectar_banco(banco)
        cursor = conn.cursor()
        cursor.execute("SELECT cods FROM atalhos WHERE nome_cod = ?", (nome_atalho,))
        codigo_sql = cursor.fetchone()[0]
        conn.close()
        return codigo_sql
    except Exception as e:
        messagebox.showerror("Erro ao carregar código do atalho", str(e))
        return ""

def registrar_log(mensagem):
    with open("logs.txt", "a") as log_file:  # Abre ou cria logs.txt no modo de anexar
        log_file.write(mensagem + "\n")  # Escreve a mensagem seguida de uma nova linha

def abrir_janela_logs():
    logs_window = tk.Toplevel(root)
    logs_window.title("Logs")
    logs_window.attributes("-fullscreen", True)  # Tamanho da nova janela

    try:
        with open("logs.txt", "r") as log_file:
            logs = log_file.readlines()
            registrar_log("Logs exibidos com sucesso.")  # Registro ao abrir os logs
    except FileNotFoundError:
        logs = ["Nenhum log encontrado."]
        registrar_log("Tentativa de exibir logs falhou: arquivo não encontrado.")

    logs_text = tk.Text(logs_window, height=15, width=50)
    logs_text.pack(pady=10)
    logs_text.insert(tk.END, "".join(logs))
    logs_text.config(state=tk.DISABLED)


# Janela principal
root = tk.Tk()
root.title("Sistema de Consultas SQL")

# Adiciona o nome do projeto
nome_projeto = tk.Label(root, text="XtreM",
                        font=("Brush script",100),fg="blue", bg="lightgray")
nome_projeto.grid(row=0, column=0,columnspan=2, padx=20, pady=20)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Botões para abrir as diferentes janelas
btn_consulta_geral = tk.Button(root, text="Consulta Geral", command=abrir_janela_consulta_geral)
btn_consulta_geral.grid(row=1, column=0, padx=20, pady=20)

btn_consulta_detalhada = tk.Button(root, text="Consulta Detalhada", command=abrir_janela_consulta_detalhada)
btn_consulta_detalhada.grid(row=1, column=1, padx=20, pady=20)

btn_atalhos = tk.Button(root, text="Criar Atalhos", command=abrir_janela_atalhos)
btn_atalhos.grid(row=2, column=0, padx=20, pady=20)

# Adicionando o botão na janela principal para abrir a execução de atalhos
btn_executar_atalho = tk.Button(root, text="Executar Atalhos", command=abrir_janela_executar_atalho)
btn_executar_atalho.grid(row=2, column=1, padx=20, pady=20)

# Adicionar o botão para ver os logs na janela principal
btn_ver_logs = tk.Button(root, text="Ver Logs", command=abrir_janela_logs)
btn_ver_logs.grid(row=3, column=0,columnspan=2, padx=20, pady=20)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)


# Treeview para exibir os dados das consultas
tree = ttk.Treeview(root)
tree.grid(row=4, column=0,columnspan=2, padx=20, pady=50)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

root.mainloop()