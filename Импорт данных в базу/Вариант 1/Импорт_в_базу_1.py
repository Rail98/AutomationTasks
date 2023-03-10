# Простой случай
# У нас только один файл с данными, который нужно залить в базу.

# импортируем все нужные библиотеки для выполения кода
import pandas as pd
import pyodbc

# Для конвертации расширения .ipynb в .py прописываем в командной строке:
# jupyter nbconvert Импорт_в_базу_1.ipynb --to script

# укажем, где хранится файл
path = 'excel.xlsx'

# загрузим данные в df
df = pd.read_excel(path) # sheet_name=1, skiprows=6, usecols='F:W' - дополнительные условия

# покажем первые пять строк
df.head()

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
                                                        trade_point varchar (64),
                                                        week int,
                                                        producer varchar (64),
                                                        brand varchar (64),
                                                        product varchar (255),
                                                        federal_district varchar (64),
                                                        region varchar (64),
                                                        city varchar (64),
                                                        address varchar (max),
                                                        category varchar (64),

                                                        total_cost_price float,
                                                        turnover_rub float,
                                                        turnover_pieces float,
                                                        weight float,
                                                        barcode varchar (64)
                                                         )
                                   '''
        cur.execute(create_table_query)
        #cur.commit() указываем коммит, если нужно


# In[ ]:


#%%time
# вносим данные в базу
with conn.cursor() as cur:
    for row in df.itertuples(name=None, index=False):
        insert_query = '''INSERT INTO tander_x5 VALUES{}'''.format(row)
        cur.execute(insert_query)

# отключаемся от базы
conn.close()




