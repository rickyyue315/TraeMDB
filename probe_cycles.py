import pyodbc
import pandas as pd

DB = r"c:\\Users\\kf_yue\\Documents\\trae_projects\\MDB\\Upload Old Article RP Parameter (ideal stock added).mdb"
conn = pyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={DB};")

rp = pd.read_sql("SELECT TOP 200 [Site], [Article             ] AS ArticleId FROM [RP_List]", conn)
am = pd.read_sql("SELECT [Article Number] AS ArticleId, [Major Vendor(SAP)] AS Vendor FROM [Article Master]", conn)
vs = pd.read_sql("SELECT [Shop] AS Site, [Vendor], [Delivery S] AS DCode, [Planning S] AS PCode FROM [Vendor Schedule]", conn)
wh = pd.read_sql("SELECT [Shop] AS Site, [P] AS PCodeShop, [D] AS DCodeShop FROM [Warehouse Calendar]", conn)

rp['ArticleId'] = rp['ArticleId'].map(lambda x: str(int(x)) if pd.notna(x) else None)
am['ArticleId'] = am['ArticleId'].astype(str)
merged = rp.merge(am, on='ArticleId', how='left')
merged = merged.merge(vs, on=['Site','Vendor'], how='left')
merged = merged.merge(wh, on='Site', how='left')

print(merged.head(20).to_string())

conn.close()
