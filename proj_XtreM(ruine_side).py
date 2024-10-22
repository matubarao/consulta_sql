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
    consulta_detalhada_window.geometry("400x300")  # Tamanho da nova janela
    
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
    
    # Label e combobox para selecionar coluna
    tk.Label(consulta_detalhada_window, text="Selecione a Coluna (Chave):").pack(pady=5)
    coluna_var = tk.StringVar(consulta_detalhada_window)
    coluna_dropdown = ttk.Combobox(consulta_detalhada_window, textvariable=coluna_var, state="readonly")
    coluna_dropdown.pack(pady=5)

    # Função para atualizar as tabelas ao selecionar um banco de dados
    def atualizar_tabelas(event):
        banco_selecionado = banco_var.get()
        tabelas = carregar_tabelas(banco_selecionado)
        tabela_dropdown['values'] = tabelas  # Atualiza as opções da tabela
        tabela_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de banco à atualização de tabelas
    banco_dropdown.bind("<<ComboboxSelected>>", atualizar_tabelas)
    
    # Função para atualizar as colunas ao selecionar uma tabela
    def atualizar_colunas(event):
        banco_selecionado = banco_var.get()
        tabela_selecionada = tabela_var.get()
        colunas = carregar_colunas(banco_selecionado, tabela_selecionada)
        coluna_dropdown['values'] = colunas  # Atualiza as opções da coluna
        coluna_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de tabela à atualização de colunas
    tabela_dropdown.bind("<<ComboboxSelected>>", atualizar_colunas)

    # Caixa de texto para inserir o valor da chave
    tk.Label(consulta_detalhada_window, text="Valor da Chave:").pack(pady=5)
    chave_entry = tk.Entry(consulta_detalhada_window)
    chave_entry.pack(pady=5)

    # Função para executar a consulta detalhada
    def executar_consulta_detalhada():
        banco = banco_var.get()
        tabela = tabela_var.get()
        coluna_chave = coluna_var.get()
        valor_chave = chave_entry.get()

        if not banco or not tabela or not coluna_chave or not valor_chave:
            messagebox.showerror("Erro", "Preencha todos os campos!")  # Verifica se os campos estão preenchidos
            return

        try:
            conn = conectar_banco(banco)
            cursor = conn.cursor()
            query = f"SELECT * FROM {tabela} WHERE {coluna_chave} = ?"
            cursor.execute(query, valor_chave)
            rows = cursor.fetchall()  # Armazena os dados retornados

            limpar_tabela()

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

    btn_executar_consulta_detalhada = tk.Button(consulta_detalhada_window, text="Executar Consulta Detalhada", command=executar_consulta_detalhada)
    btn_executar_consulta_detalhada.pack(pady=10)

# Função para abrir a nova janela de atualização de dados (UPDATE)
def abrir_janela_update():
    update_window = tk.Toplevel(root)
    update_window.title("Atualizar Dados")  # Título da nova janela
    update_window.geometry("400x420")  # Tamanho da nova janela

    # Carregar bancos de dados
    bancos = carregar_bancos()

    tk.Label(update_window, text="Selecione o Banco de Dados:").pack(pady=5)
    banco_var = tk.StringVar(update_window)
    banco_dropdown = ttk.Combobox(update_window, textvariable=banco_var, values=bancos, state="readonly")
    banco_dropdown.pack(pady=5)

    # Label e combobox para selecionar tabela
    tk.Label(update_window, text="Selecione a Tabela:").pack(pady=5)
    tabela_var = tk.StringVar(update_window)
    tabela_dropdown = ttk.Combobox(update_window, textvariable=tabela_var, state="readonly")
    tabela_dropdown.pack(pady=5)

    # Label e combobox para selecionar a coluna a ser atualizada
    tk.Label(update_window, text="Selecione a Coluna a Ser Atualizada:").pack(pady=5)
    coluna_update_var = tk.StringVar(update_window)
    coluna_update_dropdown = ttk.Combobox(update_window, textvariable=coluna_update_var, state="readonly")
    coluna_update_dropdown.pack(pady=5)

    # Caixa de texto para inserir o novo valor
    tk.Label(update_window, text="Novo Valor:").pack(pady=5)
    novo_valor_entry = tk.Entry(update_window)
    novo_valor_entry.pack(pady=5)

    # Label e combobox para selecionar a coluna do WHERE
    tk.Label(update_window, text="Selecione a Coluna do WHERE:").pack(pady=5)
    coluna_where_var = tk.StringVar(update_window)
    coluna_where_dropdown = ttk.Combobox(update_window, textvariable=coluna_where_var, state="readonly")
    coluna_where_dropdown.pack(pady=5)

    # Caixa de texto para inserir o valor do WHERE
    tk.Label(update_window, text="Valor do WHERE:").pack(pady=5)
    valor_where_entry = tk.Entry(update_window)
    valor_where_entry.pack(pady=5)

    # Função para atualizar as tabelas ao selecionar um banco de dados
    def atualizar_tabelas(event):
        banco_selecionado = banco_var.get()
        tabelas = carregar_tabelas(banco_selecionado)
        tabela_dropdown['values'] = tabelas
        tabela_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de banco à atualização de tabelas
    banco_dropdown.bind("<<ComboboxSelected>>", atualizar_tabelas)

    # Função para atualizar as colunas ao selecionar uma tabela
    def atualizar_colunas(event):
        banco_selecionado = banco_var.get()
        tabela_selecionada = tabela_var.get()
        colunas = carregar_colunas(banco_selecionado, tabela_selecionada)
        coluna_update_dropdown['values'] = colunas  # Atualiza as opções da coluna
        coluna_update_dropdown.set('')  # Limpa a seleção anterior
        coluna_where_dropdown['values'] = colunas  # Atualiza também as opções da coluna do WHERE
        coluna_where_dropdown.set('')  # Limpa a seleção anterior

    # Vincular a mudança de tabela à atualização de colunas
    tabela_dropdown.bind("<<ComboboxSelected>>", atualizar_colunas)

    # Função para executar a atualização (UPDATE)
    def executar_update():
        banco = banco_var.get()
        tabela = tabela_var.get()
        coluna_update = coluna_update_var.get()
        novo_valor = novo_valor_entry.get()
        coluna_where = coluna_where_var.get()
        valor_where = valor_where_entry.get()

        if not banco or not tabela or not coluna_update or not novo_valor or not coluna_where or not valor_where:
            messagebox.showerror("Erro", "Preencha todos os campos!")  # Verifica se os campos estão preenchidos
            return

        try: #se  algum dia alguem otimizar isso por favor entre em contato
            conn = conectar_banco(banco)
            cursor = conn.cursor()
            query = f"UPDATE {tabela} SET {coluna_update} = ? WHERE {coluna_where} = ?"
            cursor.execute(query, (novo_valor, valor_where))
            conn.commit()  # Confirma a transação
            conn.close()
            messagebox.showinfo("Sucesso", "Dados atualizados com sucesso!")  # Mensagem de sucesso

        except Exception as e:
            messagebox.showerror("Erro de Conexão", str(e))  # Exibe erro em caso de falha

    btn_executar_update = tk.Button(update_window, text="Executar Update", command=executar_update)
    btn_executar_update.pack(pady=10)

# Função para criar a tabela de atalhos
def criar_tabela_atalhos():
    try:
        conn = conectar_banco()
        cursor = conn.cursor()
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='atalhos' AND xtype='U')
            CREATE TABLE atalhos (
                nome_cod VARCHAR(255),
                cods VARCHAR(10000)
            )
        """) # ai eu ja quis dms n funca mas se possivel alguem faz funfar ai
        conn.commit()
        conn.close()
    except Exception as e:
        messagebox.showerror("Erro de Conexão", str(e))

# Função para abrir a janela de atalhos
def abrir_janela_atalhos():
    atalhos_window = tk.Toplevel(root)
    atalhos_window.title("Atalhos")
    atalhos_window.geometry("500x400")

    # Carregar bancos de dados
    bancos = carregar_bancos()

    # Seção de criar atalho
    tk.Label(atalhos_window, text="Criar Atalho").pack(pady=5)

    # Campo para o nome do atalho
    tk.Label(atalhos_window, text="Nome do Atalho:").pack(pady=5)
    nome_atalho_entry = tk.Entry(atalhos_window)
    nome_atalho_entry.pack(pady=5)

    # Dropdown para selecionar banco de dados
    tk.Label(atalhos_window, text="Selecione o Banco de Dados:").pack(pady=5)
    banco_var = tk.StringVar(atalhos_window)
    banco_dropdown = ttk.Combobox(atalhos_window, textvariable=banco_var, values=bancos, state="readonly")
    banco_dropdown.pack(pady=5)

    # Campo para inserir o código SQL
    tk.Label(atalhos_window, text="Código SQL:").pack(pady=5)
    cod_sql_entry = tk.Text(atalhos_window, height=10, width=50)
    cod_sql_entry.pack(pady=5)

    # Função para salvar o atalho no banco
    def salvar_atalho():
        nome_atalho = nome_atalho_entry.get()
        banco = banco_var.get()
        codigo_sql = cod_sql_entry.get("1.0", tk.END).strip()

        if not nome_atalho or not banco or not codigo_sql:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        try:
            conn = conectar_banco(banco)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO dbo.atalhos (nome_cod, cods) VALUES (?, ?)", (nome_atalho, codigo_sql))
            conn.commit()
            conn.close()
            messagebox.showinfo("Sucesso", "Atalho salvo com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao salvar atalho", str(e))

    # Botão para salvar o atalho
    btn_salvar_atalho = tk.Button(atalhos_window, text="Salvar Atalho", command=salvar_atalho)
    btn_salvar_atalho.pack(pady=10)

    # Função para executar o atalho
    def executar_atalho():
        nome_atalho = nome_atalho_entry.get()
        banco = banco_var.get()

        if not nome_atalho or not banco:
            messagebox.showerror("Erro", "Preencha o nome do atalho e selecione o banco de dados!")
            return

        try:
            conn = conectar_banco(banco)
            cursor = conn.cursor()
            cursor.execute("SELECT cods FROM atalhos WHERE nome_cod = ?", (nome_atalho,))
            resultado = cursor.fetchone()

            if resultado:
                codigo_sql = resultado[0]
                # Executar o código SQL
                cursor.execute(codigo_sql)
                conn.commit()
                messagebox.showinfo("Sucesso", "Atalho executado com sucesso!")
            else:
                messagebox.showerror("Erro", "Atalho não encontrado.")
            conn.close()
        except Exception as e:
            messagebox.showerror("Erro ao executar atalho", str(e))
        
        # Botão para executar o atalho
        btn_executar_atalho = tk.Button(atalhos_window, text="Executar Atalho", command=executar_atalho)
        btn_executar_atalho.pack(pady=10)

        # Executar a função para criar a tabela de atalhos, se ela ainda não existir
        criar_tabela_atalhos()
        


# Configurações principais da interface
root = tk.Tk()
root.title("Interface de Banco de Dados")
root.geometry("600x400")

tree = ttk.Treeview(root)
tree.pack(expand=True, fill=tk.BOTH)

# Botão para abrir a janela de consulta geral
btn_consulta_geral = tk.Button(root, text="Consulta Geral", command=abrir_janela_consulta_geral)
btn_consulta_geral.pack(pady=10)

# Botão para abrir a janela de consulta detalhada
btn_consulta_detalhada = tk.Button(root, text="Consulta Detalhada", command=abrir_janela_consulta_detalhada)
btn_consulta_detalhada.pack(pady=10)

# Botão para abrir a janela de atualização (UPDATE)
btn_update = tk.Button(root, text="Atualizar Dados", command=abrir_janela_update)
btn_update.pack(pady=10)

# Botão para abrir a janela de atualização (UPDATE)
btn_update = tk.Button(root, text="Criar atalhos", command=abrir_janela_atalhos)
btn_update.pack(pady=10)

root.mainloop()