import streamlit as st
import pandas as pd
try:
    import pyodbc
except Exception:
    pyodbc = None
import io
import os
from datetime import datetime, timedelta
try:
    from zoneinfo import ZoneInfo
    _HK_TZ = ZoneInfo("Asia/Hong_Kong")
except Exception:
    _HK_TZ = None
from datetime import datetime, timedelta
try:
    from zoneinfo import ZoneInfo
    _HK_TZ = ZoneInfo("Asia/Hong_Kong")
except Exception:
    _HK_TZ = None

DB_DEFAULT = r"c:\\Users\\kf_yue\\Documents\\trae_projects\\MDB\\Upload Old Article RP Parameter (ideal stock added).mdb"

def get_conn(db_path: str):
    return pyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};")

def list_objects(conn):
    cur = conn.cursor()
    tables = [row.table_name for row in cur.tables(tableType='TABLE')]
    views = [row.table_name for row in cur.tables(tableType='VIEW')]
    cur.close()
    return tables, views

def get_columns(conn, name: str):
    cur = conn.cursor()
    cols = [row.column_name for row in cur.columns(table=name)]
    cur.close()
    return cols

def fetch_df(conn, name: str, top_n: int):
    q = f"SELECT TOP {top_n} * FROM [{name}]"
    df = pd.read_sql(q, conn)
    return df

def insert_df(conn, table: str, df: pd.DataFrame, mode: str):
    cols = get_columns(conn, table)
    use_cols = [c for c in df.columns if c in cols]
    if len(use_cols) == 0:
        raise ValueError("無可寫入的欄位，請確認上傳檔欄位名稱與目標表一致")
    cur = conn.cursor()
    if mode == 'replace':
        cur.execute(f"DELETE FROM [{table}]")
    placeholders = ",".join(["?"] * len(use_cols))
    col_list = ",".join([f"[{c}]" for c in use_cols])
    sql = f"INSERT INTO [{table}] ({col_list}) VALUES ({placeholders})"
    for _, row in df[use_cols].iterrows():
        cur.execute(sql, list(row.values))
    conn.commit()
    cur.close()

def hk_date_yyyymmdd():
    if _HK_TZ is not None:
        return datetime.now(_HK_TZ).strftime("%Y%m%d")
    return (datetime.utcnow() + timedelta(hours=8)).strftime("%Y%m%d")

def build_result_name(base: str):
    return f"{base}_Trae{hk_date_yyyymmdd()}" + ".xlsx"

st.set_page_config(page_title="RP 參數上傳與檢核", layout="wide")
st.title("Access 移植到 Streamlit: RP 參數作業")

def hk_date_yyyymmdd():
    if _HK_TZ is not None:
        return datetime.now(_HK_TZ).strftime("%Y%m%d")
    return (datetime.utcnow() + timedelta(hours=8)).strftime("%Y%m%d")

def build_result_name(base: str):
    return f"{base}_Trae{hk_date_yyyymmdd()}" + ".xlsx"

def svg_data_url(path: str):
    try:
        with open(path, "rb") as f:
            data = f.read()
        import base64
        b64 = base64.b64encode(data).decode()
        return f"data:image/svg+xml;base64,{b64}"
    except Exception:
        return ""

def inject_ui_style():
    st.markdown(
        """
        <style>
        :root { --primary:#0F766E; --primary-600:#115E59; --accent:#2563EB; --bg:#0B1220; --card:#0F172A; --muted:#94A3B8; --text:#E2E8F0; }
        html, body, [data-testid="stAppViewContainer"] { background: var(--bg); color: var(--text); }
        [data-testid="stSidebar"] { background: #0D1324; }
        .brand-bar { display:flex; align-items:center; gap:12px; padding:12px 16px; border-radius:12px; background: linear-gradient(90deg, var(--card), #0C1A34); border:1px solid #1E293B; }
        .brand-title { font-weight:600; font-size:18px; letter-spacing:0.5px; }
        .brand-sub { font-size:12px; color: var(--muted); display:flex; gap:16px; }
        .tag { display:inline-flex; align-items:center; gap:6px; padding:4px 10px; border-radius:999px; background: rgba(15,118,110,0.15); color: #A7F3D0; border:1px solid rgba(15,118,110,0.35); }
        .section-header { display:flex; align-items:center; gap:10px; margin:18px 0 10px; }
        .section-name { font-weight:600; font-size:16px; }
        img.icon { width:20px; height:20px; }
        .card { background: var(--card); border:1px solid #1E293B; border-radius:12px; padding:12px; }
        .muted { color: var(--muted); }
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_brand():
    logo = svg_data_url(os.path.join("assets","icons","logo.svg"))
    info = svg_data_url(os.path.join("assets","icons","info.svg"))
    st.markdown(
        f"""
        <div class="brand-bar">
          <img class="icon" src="{logo}">
          <div>
            <div class="brand-title">Access 移植到 Streamlit: RP 參數作業</div>
            <div class="brand-sub">
              <span class="tag">雲端模式</span>
              <span class="muted">本系統僅限SASA RP team測試使用</span>
              <span class="muted">開發者：Ricky Yue</span>
            </div>
          </div>
          <img class="icon" src="{info}">
        </div>
        """,
        unsafe_allow_html=True,
    )

inject_ui_style()
render_brand()

db_path = st.text_input("MDB 檔案路徑", DB_DEFAULT) if pyodbc else ""
if pyodbc and not os.path.isfile(db_path):
    st.warning("找不到 MDB 檔案，請改用『雲端計算(無MDB)』或提供正確路徑")

conn = None
tables, views = [], []
if pyodbc is not None and os.path.isfile(db_path):
    try:
        conn = get_conn(db_path)
        tables, views = list_objects(conn)
    except Exception:
        conn = None
        tables, views = [], []

section = st.sidebar.selectbox("功能", ["雲端計算(無MDB)", "資料瀏覽", "RP 參數上傳", "視圖與檢核", "計算與匯出", "匯出資料", "檔案導入+計算"])

if section == "雲端計算(無MDB)":
    cloud = svg_data_url(os.path.join("assets","icons","cloud.svg"))
    st.markdown(f"<div class='section-header'><img class='icon' src='{cloud}'><div class='section-name'>雲端計算</div></div>", unsafe_allow_html=True)
    st.write("使用 RP_List.txt、Article Master.xlsx、Planning Cycle.xls，無需 Access 連線")
    rp_file = st.file_uploader("上傳 RP List.txt", type=["txt","csv"])
    am_file = st.file_uploader("上傳 Article Master.xlsx", type=["xlsx"])
    pc_file = st.file_uploader("上傳 Planning Cycle.xls", type=["xls","xlsx"])

    RESULT_COLUMNS = [
        'Site','Article','Article Description','Brand','MC','MC Description','Article categor','Article Type','Status','First Sales Dat','Season category','Available to','Launch Date','Sales Qty 20000101 -  20061203','Sales Price','Avg Weekly Sales','Cal Stock Turnover','Stock On Hand  20070827','Safety Stock','Purchase Group','RP Type','Planning Cycle','Delivery Cycle','Stock Planner','Reorder Point','Delivery Days','Target Coverage','Supply Source (1=Vendor/2=DC)','ABC Indicator','Smooth Promotion','Forecast Model','Historical periods','Forecast periods','Periods per season','Current consumption qty','Week 1 forecast value','Week 2 forecast value','Week 3 forecast value','Week 4 forecast value','Week 5 forecast value','New Safety Qty','New Purchase Group','New RP Typ','New Planning Cycle','New Delivery Cycle','New Stock Planner','New Reorder Point','New Delivery Days','New Traget Coverage','New Supply Source','New ABC Indicator','New Smoothing (0/1)','New Forecast Model','New Historical perio','New Forecast periods','New Periods per season','New Current consumption qty','New Week 1 forecast value','New Week 2 forecast value','New Week 3 forecast value','New Week 4 forecast value','New Week 5 forecast value','A QTY','B QTY','C QTY'
    ]

    def read_rp_buffer(buf):
        try:
            return pd.read_csv(buf, sep='\t', dtype=str)
        except Exception:
            buf.seek(0)
            return pd.read_csv(buf, sep='[,;|]', engine='python', dtype=str)

    def normalize_rp(df: pd.DataFrame):
        m = {
            'Article             ': 'Article',
            'Article Description                     ': 'Article Description',
            'Brand             ': 'Brand',
            'MC       ': 'MC',
            'MC Description      ': 'MC Description',
            'Article categor': 'Article categor',
            'Article Type   ': 'Article Type',
            'Status         ': 'Status',
            'First Sales Dat': 'First Sales Dat',
            'Season category': 'Season category',
            'Available to   ': 'Available to',
            'Launch Date    ': 'Launch Date',
            'Sales Qty. 20000101 -  20240214    ': 'Sales Qty 20000101 -  20061203',
            'Sales Price                   ': 'Sales Price',
            'Avg. Weekly Sales                  ': 'Avg Weekly Sales',
            'Cal. Stock Turnover           ': 'Cal Stock Turnover',
            'Stock On Hand  20251113       ': 'Stock On Hand  20070827',
            'Safety Stock        ': 'Safety Stock',
            'Purchase Group ': 'Purchase Group',
            'RP Type   ': 'RP Type',
            'Planning Cycle      ': 'Planning Cycle',
            'Delivery Cycle      ': 'Delivery Cycle',
            'Stock Planner       ': 'Stock Planner',
            'Reorder Point       ': 'Reorder Point',
            'Delivery Days       ': 'Delivery Days',
            'Target Coverage     ': 'Target Coverage',
            'Supply Source (1=Vendor/2=DC) ': 'Supply Source (1=Vendor/2=DC)',
            'ABC Indicator       ': 'ABC Indicator',
            'Smooth Promotion    ': 'Smooth Promotion',
            'Forecast Model      ': 'Forecast Model',
            'Historical periods  ': 'Historical periods',
            'Forecast periods    ': 'Forecast periods',
            'Periods per season  ': 'Periods per season',
            'Current consumption qty       ': 'Current consumption qty',
            'Week 1 forecast value         ': 'Week 1 forecast value',
            'Week 2 forecast value         ': 'Week 2 forecast value',
            'Week 3 forecast value         ': 'Week 3 forecast value',
            'Week 4 forecast value         ': 'Week 4 forecast value',
            'Week 5 forecast value         ': 'Week 5 forecast value',
            'New Safety Qty      ': 'New Safety Qty',
            'New Purchase Group  ': 'New Purchase Group',
            'New RP Typ': 'New RP Typ',
            'New Planning Cycle  ': 'New Planning Cycle',
            'New Delivery Cycle  ': 'New Delivery Cycle',
            'New Stock Planner   ': 'New Stock Planner',
            'New Reorder Point   ': 'New Reorder Point',
            'New Delivery Days   ': 'New Delivery Days',
            'New Traget Coverage ': 'New Traget Coverage',
            'New Supply Source   ': 'New Supply Source',
            'New ABC Indicator   ': 'New ABC Indicator',
            'New Smoothing (0/1) ': 'New Smoothing (0/1)',
            'New Forecast Model  ': 'New Forecast Model',
            'New Historical perio': 'New Historical perio',
            'New Forecast periods          ': 'New Forecast periods',
            'New Periods per season        ': 'New Periods per season',
            'New consumption qty           ': 'New Current consumption qty',
            'New Week 1 forecast           ': 'New Week 1 forecast value',
            'New Week 2 forecast           ': 'New Week 2 forecast value',
            'New Week 3 forecast           ': 'New Week 3 forecast value',
            'New Week 4 forecast           ': 'New Week 4 forecast value',
            'New Week 5 forecast           ': 'New Week 5 forecast value',
        }
        df = df.rename(columns={c: m.get(c, c) for c in df.columns})
        return df

    def read_am_buffer(buf):
        df = pd.read_excel(buf, dtype=str)
        m = {
            'Article Number (SAP)': 'Article Number',
            'Major vendor\n(SAP)': 'Major Vendor(SAP)',
            'Purchase Group': 'Purchase Group',
            'Supply Site': 'Supply Site',
            'Brand': 'Brand',
            'Article Description': 'Article Description',
            'Merchandise Category': 'Merchandise Category',
            'Supply Source': 'Supply Source',
        }
        df = df.rename(columns={c: m.get(c, c) for c in df.columns})
        return df

    def read_pc_buffer(buf):
        x = pd.ExcelFile(buf)
        df = x.parse(x.sheet_names[0])
        df.columns = [str(c).strip() for c in df.columns]
        df = df[[c for c in df.columns if c in ['Calendar','Final Code']]]
        df = df.dropna(subset=['Calendar','Final Code']).astype(str)
        return dict(zip(df['Calendar'].str.strip(), df['Final Code'].str.strip()))

    def cycles_transform(code: str, mapping: dict):
        if pd.isna(code) or code is None:
            return None
        c = str(code).strip()
        return mapping.get(c, c)

    if rp_file and am_file and pc_file:
        rp_df_raw = read_rp_buffer(rp_file)
        rp_df = normalize_rp(rp_df_raw)
        am_df = read_am_buffer(am_file)
        mapping = read_pc_buffer(pc_file)

        rp_df['New Planning Cycle'] = rp_df['Planning Cycle'].map(lambda x: cycles_transform(x, mapping) if 'Planning Cycle' in rp_df.columns else None)
        rp_df['New Delivery Cycle'] = rp_df['Delivery Cycle'].map(lambda x: cycles_transform(x, mapping) if 'Delivery Cycle' in rp_df.columns else None)

        out = rp_df.copy()
        if 'Article' not in out.columns and 'Article             ' in rp_df_raw.columns:
            out['Article'] = rp_df_raw['Article             ']

        if 'Article Number' in am_df.columns:
            out = out.merge(am_df[['Article Number','Purchase Group']], left_on='Article', right_on='Article Number', how='left', suffixes=('','_am'))
            out['Purchase Group'] = out['Purchase Group'].fillna(out.get('Purchase Group', None))
            out = out.drop(columns=['Article Number'], errors='ignore')

        for col in ['A QTY','B QTY','C QTY']:
            if col not in out.columns:
                out[col] = 0

        missing = [c for c in RESULT_COLUMNS if c not in out.columns]
        for c in missing:
            out[c] = None
        out = out[RESULT_COLUMNS]
        sort_cols = [c for c in ['Site','Article'] if c in out.columns]
        if sort_cols:
            out = out.sort_values(sort_cols)

        st.success('計算完成（雲端模式）')
        st.dataframe(out.head(1000), use_container_width=True)
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine='xlsxwriter') as w:
            out.to_excel(w, index=False, sheet_name='Sheet1')
        st.download_button('下載 Result.xlsx', bio.getvalue(), file_name=build_result_name('Result'), mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

elif section == "資料瀏覽":
    eye = svg_data_url(os.path.join("assets","icons","eye.svg"))
    st.markdown(f"<div class='section-header'><img class='icon' src='{eye}'><div class='section-name'>資料瀏覽</div></div>", unsafe_allow_html=True)
    if conn is None:
        st.error("MDB 連線不可用，請改用『雲端計算(無MDB)』")
        st.stop()
    obj_type = st.radio("類型", ["TABLE", "VIEW"], horizontal=True)
    opts = tables if obj_type == "TABLE" else views
    name = st.selectbox("選擇物件", opts)
    top_n = st.slider("顯示筆數", 100, 10000, 1000, 100)
    if name:
        df = fetch_df(conn, name, top_n)
        st.dataframe(df, use_container_width=True)

elif section == "RP 參數上傳":
    upload_ic = svg_data_url(os.path.join("assets","icons","upload.svg"))
    st.markdown(f"<div class='section-header'><img class='icon' src='{upload_ic}'><div class='section-name'>RP 參數上傳</div></div>", unsafe_allow_html=True)
    if conn is None:
        st.error("MDB 連線不可用，請改用『雲端計算(無MDB)』")
        st.stop()
    upload_target_candidates = [n for n in tables if n in [
        'RP_List','Article Master','MSS List','Exemption Qty','D001 MOQ','Vendor Schedule','Warehouse Calendar','Shop_Class'
    ]] or tables
    target = st.selectbox("目標資料表", upload_target_candidates)
    mode = st.radio("寫入模式", ["append", "replace"], horizontal=True)
    f = st.file_uploader("上傳 CSV 或 Excel", type=["csv","xlsx","xls"])
    if f is not None:
        try:
            if f.name.lower().endswith(".csv"):
                df = pd.read_csv(f)
            else:
                df = pd.read_excel(f)
        except Exception as e:
            st.error(f"讀檔錯誤: {e}")
            st.stop()
        st.write("預覽")
        st.dataframe(df.head(100), use_container_width=True)
        acc_cols = get_columns(conn, target)
        st.write("目標欄位", acc_cols)
        use_cols = [c for c in df.columns if c in acc_cols]
        st.write("可寫入欄位", use_cols)
        if st.button("開始寫入"):
            try:
                insert_df(conn, target, df, mode)
                st.success("寫入完成")
            except Exception as e:
                st.error(f"寫入失敗: {e}")

elif section == "視圖與檢核":
    db_ic = svg_data_url(os.path.join("assets","icons","database.svg"))
    st.markdown(f"<div class='section-header'><img class='icon' src='{db_ic}'><div class='section-name'>視圖與檢核</div></div>", unsafe_allow_html=True)
    if conn is None:
        st.error("MDB 連線不可用，請改用『雲端計算(無MDB)』")
        st.stop()
    default_views = [v for v in views if v in ['Final Result 2','Check MOQ']]
    default_tables = [t for t in tables if t in ['Paste Errors','貼上錯誤','Problem Transactions','Final Result']]
    choices = default_views + default_tables
    if not choices:
        choices = views + tables
    name = st.selectbox("選擇視圖/檢核", choices)
    top_n = st.slider("顯示筆數", 100, 10000, 1000, 100)
    if name:
        df = fetch_df(conn, name, top_n)
        st.dataframe(df, use_container_width=True)

def load_cycles_mapping(path: str):
    x = pd.ExcelFile(path)
    df = x.parse(x.sheet_names[0])
    df = df[[c for c in df.columns if c.strip() in ["Calendar","Final Code"]]]
    df.columns = ["Calendar","FinalCode"]
    df = df.dropna(subset=["Calendar","FinalCode"]).astype(str)
    m = dict(zip(df["Calendar"].str.strip(), df["FinalCode"].str.strip()))
    return m

def cycles_transform(code: str, mapping: dict):
    if pd.isna(code) or code is None:
        return None
    c = str(code).strip()
    return mapping.get(c, c)

def apply_mdb_logic(conn, planning_cycle_xls_path: str):
    mapping = load_cycles_mapping(planning_cycle_xls_path)
    wh = pd.read_sql("SELECT [Shop] AS Site, [P] AS PCode, [D] AS DCode FROM [Warehouse Calendar]", conn)
    wh["FinalP"] = wh["PCode"].map(lambda x: cycles_transform(x, mapping))
    wh["FinalD"] = wh["DCode"].map(lambda x: cycles_transform(x, mapping))
    cur = conn.cursor()
    for _, r in wh.iterrows():
        cur.execute(
            "UPDATE [RP_List] SET [New Planning Cycle] = ?, [New Delivery Cycle] = ? WHERE [Site] = ?",
            (r["FinalP"], r["FinalD"], r["Site"]),
        )
    conn.commit()
    cur.close()
    try:
        df = pd.read_sql("SELECT * FROM [Final Result 2]", conn)
    except Exception:
        df = pd.read_sql("SELECT * FROM [Final Result]", conn)
    sort_cols = [c for c in ["Site","Article"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols)
    return df

if section == "計算與匯出":
    calc_ic = svg_data_url(os.path.join("assets","icons","calc.svg"))
    st.markdown(f"<div class='section-header'><img class='icon' src='{calc_ic}'><div class='section-name'>計算與匯出</div></div>", unsafe_allow_html=True)
    if conn is None:
        st.error("MDB 連線不可用，請改用『雲端計算(無MDB)』")
        st.stop()
    planning_path = os.path.join(os.path.dirname(db_path), "Planning Cycle.xls")
    st.write("套用 New Planning Cycle / New Delivery Cycle 以計算 RP/理想庫存，並輸出 Result.xlsx")
    st.write("來源規則檔:", planning_path)
    do = st.button("計算並產出 Result.xlsx")
    if do:
        try:
            df = apply_mdb_logic(conn, planning_path)
            out_xlsx = os.path.join(os.path.dirname(db_path), build_result_name("Result"))
            with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as w:
                df.to_excel(w, index=False, sheet_name="Sheet1")
            st.success(f"已輸出 {out_xlsx}")
            st.dataframe(df.head(1000), use_container_width=True)
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
                df.to_excel(w, index=False, sheet_name="Sheet1")
            st.download_button("下載 Result.xlsx", bio.getvalue(), file_name=os.path.basename(out_xlsx), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"計算失敗: {e}")

elif section == "匯出資料":
    export_ic = svg_data_url(os.path.join("assets","icons","export.svg"))
    st.markdown(f"<div class='section-header'><img class='icon' src='{export_ic}'><div class='section-name'>匯出資料</div></div>", unsafe_allow_html=True)
    if conn is None:
        st.error("MDB 連線不可用，請改用『雲端計算(無MDB)』")
        st.stop()
    obj_type = st.radio("類型", ["TABLE", "VIEW"], horizontal=True)
    opts = tables if obj_type == "TABLE" else views
    name = st.selectbox("選擇物件", opts)
    top_n = st.slider("匯出筆數", 100, 10000, 5000, 100)
    if name:
        df = fetch_df(conn, name, top_n)
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("下載 CSV", csv, file_name=f"{name}.csv", mime="text/csv")

def read_rp_list_txt(path: str):
    try:
        df = pd.read_csv(path, sep="\t", dtype=str)
    except Exception:
        df = pd.read_csv(path, sep="[,;|]", engine="python", dtype=str)
    return df

def normalize_am_df(df: pd.DataFrame):
    m = {
        "Article Number (SAP)": "Article Number",
        "Major vendor\n(SAP)": "Major Vendor(SAP)",
        "Purchase Group": "Purchase Group",
        "Supply Site": "Supply Site",
        "Brand": "Brand",
        "Article Description": "Article Description",
        "Merchandise Category": "Merchandise Category",
        "Supply Source": "Supply Source",
    }
    cols = {c: m.get(c, c) for c in df.columns}
    df = df.rename(columns=cols)
    return df

def import_files_to_mdb(conn, rp_path: str, am_path: str):
    rp_df = read_rp_list_txt(rp_path)
    am_df = pd.read_excel(am_path, dtype=str)
    am_df = normalize_am_df(am_df)
    insert_df(conn, "RP_List", rp_df, mode="replace")
    insert_df(conn, "Article Master", am_df, mode="replace")

def compare_result(df: pd.DataFrame, ref_path: str):
    try:
        ref = pd.read_excel(ref_path)
        same_cols = list(df.columns) == list(ref.columns)
        st.write("欄位一致:", same_cols)
        st.write("行數", len(df), "vs", len(ref))
        show = min(5, len(df))
        st.write("新檔前幾行")
        st.dataframe(df.head(show), use_container_width=True)
        st.write("既有檔前幾行")
        st.dataframe(ref.head(show), use_container_width=True)
    except Exception as e:
        st.warning(f"比對失敗: {e}")

if section == "檔案導入+計算":
    st.markdown(f"<div class='section-header'><img class='icon' src='{upload_ic}'><div class='section-name'>檔案導入+計算</div></div>", unsafe_allow_html=True)
    if conn is None:
        st.error("MDB 連線不可用，請改用『雲端計算(無MDB)』")
        st.stop()
    base = os.path.dirname(db_path)
    rp_path = os.path.join(base, "RP List.txt")
    am_path = os.path.join(base, "Article Master.xlsx")
    planning_path = os.path.join(base, "Planning Cycle.xls")
    st.write("資料來源:")
    st.write("RP_List:", rp_path)
    st.write("Article Master:", am_path)
    st.write("Planning Cycle:", planning_path)
    do_import = st.button("匯入 RP List 與 Article Master 至 MDB")
    if do_import:
        try:
            import_files_to_mdb(conn, rp_path, am_path)
            st.success("匯入完成")
        except Exception as e:
            st.error(f"匯入失敗: {e}")
    do_all = st.button("一鍵導入 + 計算 + 匯出 + 比對")
    if do_all:
        try:
            import_files_to_mdb(conn, rp_path, am_path)
            df = apply_mdb_logic(conn, planning_path)
            out_xlsx = os.path.join(base, build_result_name("Result"))
            with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as w:
                df.to_excel(w, index=False, sheet_name="Sheet1")
            st.success(f"已輸出 {out_xlsx}")
            compare_result(df, out_xlsx)
        except Exception as e:
            st.error(f"執行失敗: {e}")
