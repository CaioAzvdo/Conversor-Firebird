import pandas as pd
import fdb

# Conectar ao banco Firebird
con = fdb.connect(
    dsn=r'localhost:C:\Users\kaioc\Downloads\DADOS - Copia\DADOS\INTEGRADO.GDB',  # Exemplo: 'localhost:C:/meubanco/dados.fdb'
    user='SYSDBA',
    password='masterkey',
    charset='UTF8'
    # fb_library_name=r'C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll'

)

query = "SELECT * FROM TABPRO"
cur = con.cursor()
cur.execute(query)
columns = [desc[0] for desc in cur.description]
rows = cur.fetchall()
df = pd.DataFrame(rows, columns=columns)
df.to_excel("dados_exportados.xlsx", index=False)
con.close()