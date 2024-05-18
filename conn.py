import pyodbc

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
    print("Koneksi berhasil!")
    
    # Memeriksa versi SQL Server
    cursor = conn.cursor()
    cursor.execute("SELECT @@version;")
    row = cursor.fetchone()
    print(f"Versi SQL Server: {row[0]}")
    
    # Menutup koneksi
    conn.close()
    
except Exception as e:
    print("Koneksi gagal:", e)
