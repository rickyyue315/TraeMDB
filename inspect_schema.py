import pyodbc

DB = r"c:\\Users\\kf_yue\\Documents\\trae_projects\\MDB\\Upload Old Article RP Parameter (ideal stock added).mdb"

conn = pyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={DB};")
cur = conn.cursor()

tables = [
    'Article Master','RP_List','Vendor Schedule','Warehouse Calendar','Shop_Class',
    'MSS List','D001 MOQ','Exemption Qty','Final Result','Paste Errors',
    'Problem Transactions','RF to ND','ND to RF'
]

for t in tables:
    try:
        cols = [row.column_name for row in cur.columns(table=t)]
        print('TABLE', t, 'COLUMNS:', cols)
    except Exception as e:
        print('ERR columns', t, e)

views = ['Final Result 2','Check MOQ']
for v in views:
    try:
        cur.execute(f"SELECT TOP 1 * FROM [{v}]")
        cols = [d[0] for d in cur.description]
        print('VIEW', v, 'COLUMNS:', cols)
    except Exception as e:
        print('ERR view', v, e)

cur.close(); conn.close()

