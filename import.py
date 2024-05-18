import pyodbc
import pandas as pd
from datetime import datetime

# Parameter koneksi
driver   = 'ODBC Driver 17 for SQL Server'
server   = ''
database = ''
trusted  = 'yes'

# String koneksi
conn_str = (
    f"DRIVER={driver};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"Trusted_Connection={trusted};"
)

try:
    # Membuat koneksi
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Load data CSV
    file_path = './csv_data/googleplaystore.csv'
    df_data = pd.read_csv(file_path, sep=',', encoding='utf-8')

    # Inisialisasi kolom dan nama sheet
    table_name = 'data'
    columns = ['App', 'Category', 'Genres', 'Size', 'Type', 'Reviews']
    sheet_name = 'googleplaystore'

    # Generate hdrid
    def hdrid(index):
       now = datetime.now()
       array_str = "RCDS"
       hdrid_prefix = now.strftime("%Y%m")
       hdrid_suffix = str(index).zfill(3)  
       return f'{array_str}{hdrid_prefix}{hdrid_suffix}'

    # Filter kolom
    df_data = df_data[columns]

    # Ambil 15 baris pertama
    df_data = df_data.head(15)

    print(f'Loading Worksheet: {sheet_name}')

    # generate hdrid
    df_data['hdrid'] = [hdrid(i) for i in range(1, len(df_data) + 1)]

    # generate transaction date
    df_data['transaction_date'] = datetime.now().strftime("%Y-%m-%d")

    # Iterate through DataFrame and insert rows using cursor
    for index, row in df_data.iterrows():
        # SQL statement for insertion
        sql_insert = f"INSERT INTO {table_name} ({', '.join(df_data.columns)}) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
        # Execute SQL statement with row values
        cursor.execute(sql_insert, tuple(row))
        # Commit after each row insertion
        conn.commit()

    # Close cursor and connection
    cursor.close()
    conn.close()

    print(df_data)
    print("Data has been successfully inserted into the database!")

except Exception as e:
    print("Error:", e)
