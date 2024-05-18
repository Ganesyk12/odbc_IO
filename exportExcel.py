import pyodbc
import pandas as pd
import os
from datetime import datetime
from plyer import notification

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
   #create SQL conn
   conn = pyodbc.connect(conn_str)
   cursor = conn.cursor()
   # print(pyodbc.drivers())
   # ['SQL Server', 'Microsoft Access Driver (*.mdb, *.accdb)', 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)', 'Microsoft Access Text Driver (*.txt, *.csv)', 'ODBC Driver 17 for SQL Server', 'SQL Server Native Client RDA 11.0']
   # SQL command to read data
   sqlQuery = "SELECT * FROM dbo.data"
   cursor.execute(sqlQuery)
   # getting data from sql into pandas dataFrame
   columns = [columns[0] for columns in cursor.description]
   data = cursor.fetchall()
   # df = pd.read_sql(sqlQuery, conn)
   df = pd.DataFrame(data, columns=columns)
   # close SQL conn
   cursor.close()
   conn.close()

   # initial path
   directory = "D:\\##PROJECT\\Python\\SQL_backup_auto"

   # create directory if not exist
   if not os.path.exists(directory):
      os.makedirs(directory)

   # merge path with filename
   file_path = os.path.join(directory, "SQL_Data_" + datetime.now().strftime("%d-%m-%Y") + ".xlsx")
   # export data to path
   df.to_excel(file_path, index = False)

   # display notification status
   notification.notify(title = "Report Status!!!", 
                     message = f"Exported data successfully save into Excel.\
                        \nTotal Rows: {df.shape[0]} \nTotal Columns: {df.shape[1]}", 
                        timeout = 10)

except Exception as e:
   print(e)
