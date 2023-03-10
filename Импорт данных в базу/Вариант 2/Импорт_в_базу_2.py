# Общий случай
# У нас несколько файлов с данными, которые нужно залить в базу. Важно, чтобы кол-во столбцов совпадало у всех.

# импортируем все нужные библиотеки для выполения кода
import pandas as pd
import pyodbc
import os

# Для конвертации расширения .ipynb в .py прописываем в командной строке:
# jupyter nbconvert Импорт_в_базу_2.ipynb --to script

# указываем путь к папке, куда были сохранены данные
path_to_dir = 'Import data'

# сохраняем в отдельный лист, пути к файлам
paths_list = []
for file in os.listdir(path_to_dir):
    if file.endswith(".xlsx"):
        paths_list.append(file)

# параметры соединения с базой
sql_server = '192.168.2.128'
database = 'analitic'
driver = '{SQL Server Native Client 10.0}'
conn_string = f'SERVER={sql_server};DATABASE={database};DRIVER={driver};Trusted_connection=Yes'

# устанавливаем соединение с базой
conn = pyodbc.connect(conn_string)

# создаем таблицу в базе
with conn.cursor() as cur:
        create_table_query = '''
                                CREATE TABLE tander_x5 (
                                                        id int IDENTITY(1,1) PRIMARY KEY,
                                                        network varchar (64),
                                                        ...
                                                         )
                                   '''
        cur.execute(create_table_query)
        #cur.commit() указываем коммит, если по умолчанию не стоит autocommit=True

# insert функция
def insert_data(connection, df):
    with connection.cursor() as cur:
        for row in df.itertuples(name=None, index=False):
            insert_query = '''INSERT INTO tander_x5 VALUES{}'''.format(row)
            cur.execute(insert_query)

# пробегаемся по всем файлам, чтобы поочередно залить в базу каждый из них  
for path in paths_list:
    df = pd.read_excel(path) # sheet_name=1, skiprows=6, usecols='F:W' - дополнительные условия
    insert_data(conn, df) # вносим данные в базу

# отключаемся от базы
conn.close()





