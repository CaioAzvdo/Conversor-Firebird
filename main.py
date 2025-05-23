import pandas as pd
import fdb


con = fdb.connect(
    dsn=r'localhost:C:\Users\kaioc\Downloads\DADOS\DADOS\INTEGRADO.GDB',
    user='SYSDBA',
    password='masterkey',
    charset='UTF8'
)

query = """
SELECT
    t.CODPRO AS "COD_PRODUTO",
    '' AS COD_GRADE,
    t.DESCPRO AS "DESC_PRODUTO",
    t.CST AS "COD_IMPOSTO",
    '' AS "DESC_IMPOSTO",
    '' AS "REFERENCIA",
    t.CODBARUN AS "COD_BARRA",
    '' AS "COD_BALANCA",
    t.PESOLIQUIDO AS "PESO_LIQUIDO",
    t.PESOBRUTO  AS "PESO_BRUTO",
    t.UNIDADE AS "UNIDADE",
    (SELECT FIRST 1 tf.PRVAPRO 
     FROM TABPROFIL tf 
     WHERE tf.CODPRO = t.CODPRO 
     ORDER BY tf.PRVAPRO DESC) AS "VR_PRAZO",
    '' AS "VR_VISTA",
    (SELECT FIRST 1 tf.PRCUPRO 
     FROM TABPROFIL tf 
     WHERE tf.CODPRO = t.CODPRO 
     ORDER BY tf.PRVAPRO DESC) AS "VR_CUSTO",
    '' AS "OBS",
    t.CLASFISCAL AS "NCM",
    '' AS "FLAG_ATIVO",
    t.GRUPRO AS "COD_GRUPO",
    tg.NOMGRU AS "DESC_GRUPO",
    t.CEST AS "CEST"
FROM TABPRO t
LEFT JOIN TABGRU tg ON tg.CODGRU = t.GRUPRO;

"""
cur = con.cursor()
cur.execute(query)
columns = [desc[0] for desc in cur.description]
rows = cur.fetchall()
df = pd.DataFrame(rows, columns=columns)
df.to_excel("dados_exportados.xlsx", index=False)
con.close()