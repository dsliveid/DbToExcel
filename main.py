import json
import os
import threading

import pandas as pd
import psycopg2
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk, filedialog

import pyodbc

# 定义配置文件路径
json_file_name = "config.json"
sql_file_name = "sql.sql"

# 假设这是一个全局变量来持有数据库连接信息
db_settings = {
    'dbType': 'sqlserver',
    'schema': '',
    'server': '',
    'database': '',
    'username': '',
    'password': '',
    'port': '',
    'table': '',
    'sql': '',
}


def run_query_sqlserver(server, database, username, password, port, table):
    sqlStr = text_area.get("1.0", "end-1c")  # 正确获取Text组件中的全部文本
    try:
        server = server + "," + port if port is not None and len(port) > 0 else server
        connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(connection_string)

        query = sqlStr
        if sqlStr == "":
            query = f" select * from {table} "

        # 直接读取到DataFrame
        df = pd.read_sql_query(query, conn)
        export_to_excel(df)

    except Exception as e:
        messagebox.showerror("Error", f"Could not connect to database: {e}")


def export_to_excel(df):
    try:
        # 弹出文件保存对话框让用户选择导出文件名称及位置
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if output_path:  # 检查用户是否选择了文件路径
            df.to_excel(output_path, index=False)
            messagebox.showinfo("Success", f"Data exported successfully to {output_path}")
        else:
            messagebox.showinfo("Cancelled", "Export cancelled.")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def run_query_pgsql(server, database, username, password, port, table):
    sqlStr = text_area.get("1.0", "end-1c")  # 正确获取Text组件中的全部文本
    schema = schema_entry.get()
    try:
        # PostgreSQL连接字符串
        conn_string = f"host={server} port={port} dbname={database} user={username} password={password}"
        conn = psycopg2.connect(conn_string)
        cursor = conn.cursor()

        if schema is None or len(schema) == 0:
            messagebox.showerror("Error", f"Schema Could not None! ")
            return None
        # 设置模式
        cursor.execute(f'SET search_path TO {schema}')

        query = sqlStr
        if sqlStr == "":
            query = f" select * from {table} "

        # 直接读取到DataFrame
        df = pd.read_sql_query(query, conn)
        export_to_excel(df)

    except Exception as e:
        messagebox.showerror("Error", f"Could not connect to database: {e}")


def update_config_json(new_config):
    # 写入内容到新文件中
    with open(json_file_name, 'w', encoding='utf-8') as file:
        json.dump(new_config, file, ensure_ascii=False, indent=4)


def update_config_sql(sql_content):
    # 写入到SQL文件
    with open(sql_file_name, 'w', encoding='utf-8') as file:
        file.write(sql_content)


def db_to_excel():
    # 从输入字段获取数据库连接信息
    dbType = db_type_combobox.get()
    schema = schema_entry.get()
    server = server_entry.get()
    database = database_entry.get()
    username = username_entry.get()
    password = password_entry.get()
    port = port_entry.get()
    table = table_entry.get()
    sql = text_area.get("1.0", "end-1c")
    if server is None or len(server) < 1:
        messagebox.showerror("Error", "数据库链接地址server不能为空")
        return
    if database is None or len(database) < 1:
        messagebox.showerror("Error", "数据库名称database不能为空")
        return
    if username is None or len(username) < 1:
        messagebox.showerror("Error", "数据库登录名username不能为空")
        return
    if password is None or len(password) < 1:
        messagebox.showerror("Error", "数据库链接密码password不能为空")
        return

    new_config = {
        'dbType': dbType,
        'schema': schema,
        'server': server,
        'database': database,
        'username': username,
        'password': password,
        'port': port,
        'table': table,
    }
    update_config_json(new_config)
    update_config_sql(sql)

    dbType = db_type_combobox.get()
    if dbType == "sqlserver":
        run_query_sqlserver(server, database, username, password, port, table)
    elif dbType == "pgsql":
        run_query_pgsql(server, database, username, password, port, table)


def load_db_settings():
    # 检查配置文件是否存在
    if not os.path.isfile(json_file_name):
        return db_settings

    # 加载配置文件
    with open(json_file_name, 'r') as config_file:
        jsonObj = json.load(config_file)
        # 加载sql文件
        if not os.path.isfile(sql_file_name):
            jsonObj["sql"] = ""
            return jsonObj
        with open(sql_file_name, 'r', encoding='utf-8') as file:
            sql_content_from_file = file.read()
            jsonObj["sql"] = sql_content_from_file
            return jsonObj


# 配置文件的数据库设置
db_settings = load_db_settings()

# 创建窗体
root = tk.Tk()
root.title("SQL导出Excel工具")

# 创建一个容器Frame，并设置宽度
container = tk.Frame(root, width=200, height=100)
container.pack_propagate(False)  # 阻止容器自动调整大小以适应其内部的小部件
container.grid(row=0, column=0, padx=10, pady=10)  # 放置容器

# 数据库类型选项
db_options = ["sqlserver", "pgsql"]
# 创建一个Label
label_db_type = tk.Label(container, text="DbType:")
label_db_type.grid(row=0, column=0, sticky="E")
# 创建一个Combobox，包含数据库类型选项
db_type_combobox = ttk.Combobox(container, values=db_options, width=17)
db_type_combobox.grid(row=0, column=1)
db_type_combobox.set(db_settings['dbType'])  # 设置初始显示文本


# 添加on_db_type_changed功能以处理数据库类型变化时的界面更新
def on_db_type_changed(*args):
    if db_type_combobox.get() == "pgsql":
        schema_label.grid(row=1, column=0, sticky="E")
        schema_entry.grid(row=1, column=1)
    else:
        schema_label.grid_remove()  # 用grid_remove代替pack_forget以隐藏控件，并允许之后再恢复到其格子中
        schema_entry.grid_remove()


# 绑定事件到下拉框以侦测选择变化
db_type_combobox.bind("<<ComboboxSelected>>", on_db_type_changed)

# 创建用于schema的Label和Entry
schema_label = tk.Label(container, text="Schema:")
schema_entry = tk.Entry(container)
schema_entry.insert(0, db_settings['schema'])

# 根据当前数据库类型设置调整schema设置框的显示状态
on_db_type_changed()  # 初始调用以根据当前设置调整界面

# 创建标签和输入字段
label_server = tk.Label(container, text="Server:")
label_server.grid(row=2, column=0, sticky="E")
server_entry = tk.Entry(container)
server_entry.grid(row=2, column=1)
server_entry.insert(0, db_settings['server'])

label_database = tk.Label(container, text="Database:")
label_database.grid(row=3, column=0, sticky="E")
database_entry = tk.Entry(container)
database_entry.grid(row=3, column=1)
database_entry.insert(0, db_settings['database'])

label_username = tk.Label(container, text="Username:")
label_username.grid(row=4, column=0, sticky="E")
username_entry = tk.Entry(container)
username_entry.grid(row=4, column=1)
username_entry.insert(0, db_settings['username'])

label_password = tk.Label(container, text="Password:")
label_password.grid(row=5, column=0, sticky="E")
password_entry = tk.Entry(container, show="*")
password_entry.grid(row=5, column=1)
password_entry.insert(0, db_settings['password'])

label_port = tk.Label(container, text="Port:")
label_port.grid(row=6, column=0, sticky="E")
port_entry = tk.Entry(container)
port_entry.grid(row=6, column=1)
port_entry.insert(0, db_settings['port'])

label_table = tk.Label(container, text="table:")
label_table.grid(row=7, column=0, sticky="E")
table_entry = tk.Entry(container)
table_entry.grid(row=7, column=1)
table_entry.insert(0, db_settings['table'])


def open_progressbar_window(fun):
    # 在单独的线程中开始任务
    threading.Thread(target=fun, args=()).start()


open_db_button = tk.Button(root, text="导出Excel",
                           command=lambda: open_progressbar_window(db_to_excel))
open_db_button.grid(row=11, column=0, pady=10, sticky="nwe")

text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD)
text_area.grid(row=12, column=0, pady=10, sticky="nwe")
text_area.insert(tk.INSERT, db_settings['sql'])

root.mainloop()
