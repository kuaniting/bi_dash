import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime
import os

st.set_page_config(page_title="全院證照總覽", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# 隱藏 Streamlit 的預設選單與 footer 以呈現乾淨的嵌入畫面
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .block-container {
                padding-top: 2rem;
                padding-bottom: 2rem;
            }
            /* 根據使用者要求，全域字體再放大 */
            html, body, p, span, label, [data-testid="stMarkdownContainer"] {
                font-size: 1.25rem !important;
            }
            h1 { font-size: 2.5rem !important; }
            h2 { font-size: 2.0rem !important; }
            h3 { font-size: 1.7rem !important; }
            h4 { font-size: 1.5rem !important; }
            
            /* 針對 Segmented Control (藥丸切換器) 與所有按鈕特別放大 */
            [data-testid="stSegmentedControl"] span, 
            [data-testid="stSegmentedControl"] label,
            .stButton button {
                font-size: 1.25rem !important;
                padding: 0.6rem 1.2rem !important;
            }
            
            /* 強制側邊欄的 st.pills 呈現雙欄併排，且縮小體積以塞進一頁 */
            section[data-testid="stSidebar"] [data-testid="stPills"] [data-testid="stButtonGroup"] {
                display: flex !important;
                flex-wrap: wrap !important;
                gap: 4px !important;
            }
            section[data-testid="stSidebar"] [data-testid="stPills"] button {
                width: calc(50% - 4px) !important; /* 強制 50% 減去 gap，雙欄併排 */
                flex: none !important;
                padding: 4px !important; /* 縮小內距 */
                margin: 0 !important;
            }
            section[data-testid="stSidebar"] [data-testid="stPills"] button p {
                font-size: 0.85rem !important; /* 縮小字體 */
                line-height: 1.2 !important;
                white-space: normal !important; /* 允許文字換行 */
                text-align: center !important;
                margin: 0 !important;
            }
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# ====== 定義路徑 ======
DICT_DIR = "./lice_db"
USERS_FILE = "./users.xlsx"

# 【優化 1】將快取時間從 10 秒延長為 300 秒 (5 分鐘)，避免頻USERS_FILE = "./users.xlsx"繁讀取 Excel 癱瘓效能
@st.cache_data(ttl=300, show_spinner=False) 
def load_dictionaries():
    """
    讀取所有提供的證照字典檔，建立 ID 與 中文名稱 的映射表。
    同時整理出各分類下所有的證照完整清單，供下拉式選單使用。
    """
    dict_files = {
        'PRO': ('醫師專科證照', os.path.join(DICT_DIR, 'epro.xlsx'), 'NAME'),
        'BASICLICE': ('基本證書', os.path.join(DICT_DIR, 'prsn_basiclice.xlsx'), 'KIND'),
        'PROLICE': ('專業證照', os.path.join(DICT_DIR, 'prsn_prolice.xlsx'), 'KIND'),
        'DEPTLICE': ('單位證照', os.path.join(DICT_DIR, 'prsn_deptlice.xlsx'), 'NAME'),
        'LICE': ('健保專區', os.path.join(DICT_DIR, 'prsn_lice.xlsx'), 'LICENAME')
    }
    
    id_to_name = {}
    category_to_certs = {
        '醫師專科證照': [],
        '基本證書': [],
        '專業證照': [],
        '單位證照': [],
        '健保專區': []
    }
    
    for sheet_code, (sheet_name_zh, dict_file, name_col) in dict_files.items():
        if os.path.exists(dict_file):
            try:
                df_dict = pd.read_excel(dict_file)
                if 'ID' in df_dict.columns and name_col in df_dict.columns:
                    for _, row in df_dict.iterrows():
                        # 排除空值
                        if pd.isna(row['ID']) or pd.isna(row[name_col]):
                            continue
                            
                        # 【優化 2】統一轉為大寫並去除空白，解決大小寫不一致導致的「未知名稱」
                        lid = str(row['ID']).replace(".0", "").strip().upper()
                        cname = str(row[name_col]).strip()
                        
                        if lid == 'NAN' or cname.lower() == 'nan' or not cname: 
                            continue
                            
                        # 組合「中文及代號」格式 (移除前面的大類前綴)
                        full_name = f"{cname} ({lid})"
                        id_to_name[lid] = cname
                        category_to_certs[sheet_name_zh].append(full_name)
                        
                        # 同時映射 COMMON_ID / COMMONID 以防格式差異
                        for alt_col in ['COMMON_ID', 'COMMONID']:
                            if alt_col in df_dict.columns and pd.notna(row[alt_col]):
                                cid = str(row[alt_col]).replace(".0", "").strip().upper()
                                if cid != 'NAN':
                                    id_to_name[cid] = cname
            except Exception as e:
                # 【優化 3】將錯誤印出，方便除錯
                print(f"讀取字典檔 {dict_file} 發生錯誤: {e}")
        else:
            print(f"找不到字典檔路徑: {dict_file}")
                
    # 針對每個分類的清單進行去重與排序
    for k in category_to_certs:
        category_to_certs[k] = sorted(list(set(category_to_certs[k])))
        
    return id_to_name, category_to_certs


# 【優化 1】快取時間延長為 300 秒
@st.cache_data(ttl=300, show_spinner=False) 
def load_data():
    # 載入證照字典檔的映射資料
    id_to_name, _ = load_dictionaries()
    
    # 直接讀取指定的 users.xlsx
    if not os.path.exists(USERS_FILE):
        return pd.DataFrame() 

    try:
        xls = pd.ExcelFile(USERS_FILE)
        df_users = pd.read_excel(xls, sheet_name="users")
    except Exception as e:
        print(f"讀取 {USERS_FILE} 發生錯誤: {e}")
        return pd.DataFrame()
        
    # 建立映射 (CARDNO -> ADM_DEPT_NAME, name)
    user_mapping = df_users[['CARDNO', 'ADM_DEPT_NAME', 'name']].drop_duplicates(subset=['CARDNO'])
    
    cert_sheets = {
        'PRO': '醫師專科證照', 
        'BASICLICE': '基本證書', 
        'PROLICE': '專業證照', 
        'DEPTLICE': '單位證照', 
        'LICE': '健保專區'
    }
    dfs = []
    
    for sheet_code, sheet_name_zh in cert_sheets.items():
        if sheet_code in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_code)
                if not df.empty:
                    df['CATEGORY'] = sheet_name_zh
                    
                    # 決定要用哪一個欄位當作證照 ID
                    id_candidates = ['TYPE', 'PROID', 'CERTID', 'COMMON_ID', 'COMMONID']
                    best_col = None
                    max_notna = -1
                    for col in id_candidates:
                        if col in df.columns:
                            notna_count = df[col].notna().sum()
                            if notna_count > max_notna:
                                max_notna = notna_count
                                best_col = col
                                
                    if best_col and max_notna > 0:
                        raw_ids = df[best_col].fillna('未知').astype(str).str.replace(".0", "", regex=False)
                        
                        # 套用字典映射，產生含中文名稱的完整字串
                        def get_full_name(lice_id):
                            # 【優化 2】比對時同樣轉為大寫
                            lice_id_clean = str(lice_id).strip().upper()
                            if lice_id_clean == '未知' or lice_id_clean == 'NAN':
                                return "未知項目"
                            cname = id_to_name.get(lice_id_clean, "未知名稱")
                            return f"{cname} ({lice_id_clean})"
                            
                        df['LICE_NAME'] = raw_ids.apply(get_full_name)
                    else:
                        df['LICE_NAME'] = "未知項目"
                    dfs.append(df)
            except Exception as e:
                print(f"解析 {sheet_code} 發生錯誤: {e}")
            
    if not dfs:
        return pd.DataFrame()
        
    df_certs = pd.concat(dfs, ignore_index=True)
    
    # 關聯
    df_merged = pd.merge(df_certs, user_mapping, on='CARDNO', how='left')
    df_merged['ADM_DEPT_NAME'] = df_merged['ADM_DEPT_NAME'].fillna('未知科別')
    df_merged['name'] = df_merged['name'].fillna('未知')
    df_merged['name_display'] = df_merged['name'] + " (" + df_merged['CARDNO'].astype(str) + ")"
    
    # 日期解析
    def parse_date(d):
        if pd.isna(d): return pd.NaT
        s = str(d).strip().replace(".0", "")
        if s == '永久有效' or '9999' in s: return pd.to_datetime('2099-12-31')
        try:
            if len(s) == 8 and s.isdigit():
                return pd.to_datetime(s, format='%Y%m%d')
            # 嘗試通用解析
            return pd.to_datetime(s)
        except:
            return pd.NaT
            
    df_merged['DATEE_PARSED'] = df_merged['DATEE'].apply(parse_date)

    # 🌟 優化 1：強化日期解析與完美備用機制，產生 APP_DATE
    def parse_create_dtm(d):
        if pd.isna(d): return pd.NaT
        s = str(d).strip()
        # 完美相容您的 2020/12/23 AM 09:41:17 格式
        if " AM " in s or " PM " in s:
            ampm = " AM" if " AM " in s else " PM"
            s = s.replace(" AM ", " ").replace(" PM ", " ") + ampm
        return pd.to_datetime(s, errors='coerce')
        
    if 'CREATE_DTM' in df_merged.columns:
        df_merged['APP_DATE'] = df_merged['CREATE_DTM'].apply(parse_create_dtm)
    else:
        df_merged['APP_DATE'] = pd.NaT
        
    # 【關鍵防呆】如果 CREATE_DTM 是空的，強制用 DATES (發證日) 補上！保證圖表不斷線！
    if 'DATES' in df_merged.columns:
        df_merged['APP_DATE'] = df_merged['APP_DATE'].fillna(df_merged['DATES'].apply(parse_date))

    return df_merged


def main():
    # 顯示置中的載入中訊息
    loading_placeholder = st.empty()
    with loading_placeholder.container():
        st.markdown("<br><br><br><br><br><h3 style='text-align: center; color: #64748b;'>⏳ 讀取與處理最新資料中，請稍候...</h3>", unsafe_allow_html=True)
        
    # 獲取各字典檔內所記載的完整證照清單
    _, category_to_certs = load_dictionaries()
    df = load_data()
    
    # 資料載入完畢，清除置中訊息
    loading_placeholder.empty()
    
    if df.empty:
        st.warning(f"讀取失敗。請確認 {USERS_FILE} 與字典檔是否正確設定且內容格式正確。")
        return
        
    # ==========================
    # 讀取 Query Params 作為預設狀態 (解決點擊字卡後篩選器被重置的問題)
    # ==========================
    params = st.query_params
    default_dept = params.get("dept", "全部科別")
    default_category = params.get("category", "全部大類")
    default_cert = params.get("cert", "全部證照")
    
    # ====== 主畫面上方：標題與清除按鈕 (使用佔位符先卡位) ======
    header_placeholder = st.container()
    metrics_placeholder = st.container()
    st.markdown("<br>", unsafe_allow_html=True) # 與下方篩選器保留一點間距
            
    # 將三個篩選器並排為三欄
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # 1. 動態取得科別清單 
        all_depts = ["全部科別"] + sorted(df['ADM_DEPT_NAME'].dropna().unique().tolist())
        with st.container(height=280):
            selected_dept = st.pills("🏢 選擇科別", all_depts, default=default_dept if default_dept in all_depts else "全部科別", key="dept_pills")
            if not selected_dept: selected_dept = "全部科別"  
            
    with col2:
        # 2. 證照大類清單
        all_categories = ["全部大類"] + sorted(df['CATEGORY'].dropna().unique().tolist())
        with st.container(height=280):
            selected_category = st.pills("📂 選擇證照大類", all_categories, default=default_category if default_category in all_categories else "全部大類", key="category_pills")
            if not selected_category: selected_category = "全部大類"
            
    with col3:
        # 3. 依照所選的大類，篩選出對應的個別證照
        if selected_category == "全部大類":
            known_certs = []
            for certs in category_to_certs.values():
                known_certs.extend(certs)
            df_certs = df['LICE_NAME'].dropna().unique().tolist()
        else:
            known_certs = category_to_certs.get(selected_category, [])
            df_certs = df[df['CATEGORY'] == selected_category]['LICE_NAME'].dropna().unique().tolist()
            
        # 將兩邊的清單整合去重並排序
        all_certs_list = sorted(list(set(known_certs + df_certs)))
        
        # 新增搜尋功能 (因為搜尋框會佔據高度，所以下方的 container 高度要減少)
        search_query = st.text_input("🔍 搜尋個別證照", placeholder="輸入證照名稱關鍵字...", key="cert_search")
        if search_query:
            all_certs_list = [c for c in all_certs_list if search_query.lower() in c.lower()]
            
        all_certs = ["全部證照"] + all_certs_list
        
        with st.container(height=188): 
            selected_cert = st.pills("📜 選擇個別證照", all_certs, default=default_cert if default_cert in all_certs else "全部證照", key="cert_pills")
            if not selected_cert: selected_cert = "全部證照"

    # 狀態同步：把使用者的選擇寫回 Query Params，確保換頁或點擊字卡時狀態一致
    if selected_dept != "全部科別": st.query_params["dept"] = selected_dept
    elif "dept" in st.query_params: del st.query_params["dept"]
    
    if selected_category != "全部大類": st.query_params["category"] = selected_category
    elif "category" in st.query_params: del st.query_params["category"]
    
    if selected_cert != "全部證照": st.query_params["cert"] = selected_cert
    elif "cert" in st.query_params: del st.query_params["cert"]

    # 執行最終過濾
    df_final = df
    title_parts = []
    
    if selected_dept != "全部科別":
        df_final = df_final[df_final['ADM_DEPT_NAME'] == selected_dept]
        title_parts.append(selected_dept)
        
    if selected_category != "全部大類":
        df_final = df_final[df_final['CATEGORY'] == selected_category]
        title_parts.append(selected_category)
        
    if selected_cert != "全部證照":
        df_final = df_final[df_final['LICE_NAME'] == selected_cert]
        title_parts.append(selected_cert)
        
    # 動態顯示標題
    base_prefix = "全院" if not title_parts else " ❯ ".join(title_parts)
    
    status_map = {
        "valid": "有效",
        "expiring": "快過期",
        "expired": "已逾期"
    }
    selected_status = st.query_params.get("status", "all")
    status_text = status_map.get(selected_status)
    
    if status_text:
        prefix_text = f"{base_prefix} ❯ {status_text}"
        page_title = f"{prefix_text} 證照總覽"
    else:
        prefix_text = base_prefix
        page_title = f"{base_prefix} 證照總覽"
        
    # 回到最上方的佔位符，將標題與清除按鈕並排顯示
    with header_placeholder:
        col_title, col_clear = st.columns([0.85, 0.15], vertical_alignment="bottom")
        with col_title:
            st.markdown(f"<h2 style='text-align: left; color: #1E293B; margin: 0;'>{page_title}</h2>", unsafe_allow_html=True)
            st.markdown(f"<div style='font-size: 15px; color: #94a3b8; margin-top: -5px;'>🔄 數據更新至：{pd.Timestamp.now().strftime('%Y/%m/%d %H:%M')}</div>", unsafe_allow_html=True)
        with col_clear:
            def clear_all_filters():
                st.query_params.clear()
                st.session_state["dept_pills"] = "全部科別"
                st.session_state["category_pills"] = "全部大類"
                st.session_state["cert_pills"] = "全部證照"
                st.session_state["cert_search"] = ""
                
            st.button("🔄 清除所有選項", use_container_width=True, on_click=clear_all_filters)
        
    render_dashboard_content(df_final, selected_dept, selected_category, selected_cert, metrics_placeholder, prefix_text)

# 【優化 1】移除 @st.fragment 裝飾器，讓畫面只在使用者操作時才重繪，消除無意義的背景運算卡頓
def render_dashboard_content(df_final, selected_dept, selected_category, selected_cert, metrics_placeholder, prefix_text):
    # ====== 狀態計算 ======
    now = pd.Timestamp.now().normalize()
    
    def get_status(row):
        if pd.isna(row['DATEE_PARSED']): return "未知"
        diff_days = (row['DATEE_PARSED'] - now).days
        if diff_days < 0:
            return "已逾期"
        elif diff_days <= 30:
            return "快過期 (30天內)"
        else:
            return "有效"
            
    df_final['STATUS_LABEL'] = df_final.apply(get_status, axis=1)
    
    total_count = len(df_final)
    valid_count = len(df_final[df_final['STATUS_LABEL'] == "有效"])
    expiring_count = len(df_final[df_final['STATUS_LABEL'] == "快過期 (30天內)"])
    expired_count = len(df_final[df_final['STATUS_LABEL'] == "已逾期"])
    
    # ====== UI 渲染 ======
    
    # 計算佔比
    valid_pct = round((valid_count / total_count * 100), 1) if total_count > 0 else 0
    expiring_pct = round((expiring_count / total_count * 100), 1) if total_count > 0 else 0
    expired_pct = round((expired_count / total_count * 100), 1) if total_count > 0 else 0

    # ==========================
    # 透過 Query Params 進行字卡點擊過濾 (100% 完美 CSS 支援)
    # ==========================
    params = st.query_params
    selected_status = params.get("status", "all")
    
    # 根據選擇設定邊框高亮樣式
    border_all = "border: 2px solid #3B82F6;" if selected_status == "all" else "border: 1px solid #f0f0f0;"
    border_valid = "border: 2px solid #22C55E;" if selected_status == "valid" else "border: 1px solid #f0f0f0;"
    border_expiring = "border: 2px solid #F97316;" if selected_status == "expiring" else "border: 1px solid #f0f0f0;"
    border_expired = "border: 2px solid #EF4444;" if selected_status == "expired" else "border: 1px solid #f0f0f0;"

    st.markdown("""
        <style>
        .metric-card-container {
            display: flex;
            flex-wrap: wrap; /* 🌟 關鍵修復：允許手機版自動把卡片往下折行 */
            gap: 20px;
            margin-bottom: 24px;
        }
        .metric-card-link {
            flex: 1 1 140px; /* 🌟 關鍵修復：設定最小寬度，手機上會自動完美變成 2x2 的雙欄方塊！ */
            text-decoration: none !important;
            color: inherit !important;
        }
        .metric-card {
            background-color: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
            display: flex;
            align-items: center;
            gap: 16px;
            transition: all 0.2s;
            height: 100%;
        }
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        }
        .metric-icon {
            width: 56px;
            height: 56px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
        }
        .icon-blue { background-color: #EBF5FF; color: #3B82F6; }
        .icon-green { background-color: #DCFCE7; color: #22C55E; }
        .icon-orange { background-color: #FFEDD5; color: #F97316; }
        .icon-red { background-color: #FEE2E2; color: #EF4444; }
        
        .metric-content {
            display: flex;
            flex-direction: column;
            text-align: left;
        }
        .metric-title {
            color: #64748B;
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 4px;
        }
        .metric-value {
            color: #0F172A;
            font-size: 28px;
            font-weight: 700;
            line-height: 1.2;
            margin-bottom: 4px;
        }
        .metric-subtext {
            color: #475569;
            font-size: 22px;
            font-weight: 600;
        }
        .trend-up {
            color: #22C55E;
            font-weight: 500;
        }
        </style>
    """, unsafe_allow_html=True)

    # 為了保留當前的篩選器狀態 (科別、證照)，我們要在 href 中帶上它們
    import urllib.parse
    base_query = {}
    if selected_dept != "全部科別": base_query["dept"] = selected_dept
    if selected_category != "全部大類": base_query["category"] = selected_category
    if selected_cert != "全部證照": base_query["cert"] = selected_cert
    
    def make_href(status):
        q = base_query.copy()
        q["status"] = status
        query_string = urllib.parse.urlencode(q)
        return f"?{query_string}"

    cards_html = f"""
    <div class="metric-card-container">
        <a href="{make_href('all')}" target="_self" class="metric-card-link">
            <div class="metric-card" style="{border_all}">
                <div class="metric-icon icon-blue">
                    <svg width="28" height="28" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"></path></svg>
                </div>
                <div class="metric-content">
                    <div class="metric-title">總證照數</div>
                    <div class="metric-value">{total_count:,}</div>
                </div>
            </div>
        </a>
        <a href="{make_href('valid')}" target="_self" class="metric-card-link">
            <div class="metric-card" style="{border_valid}">
                <div class="metric-icon icon-green">
                    <svg width="28" height="28" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                </div>
                <div class="metric-content">
                    <div class="metric-title">有效證照數</div>
                    <div class="metric-value">{valid_count:,}</div>
                    <div class="metric-subtext">佔比 {valid_pct}%</div>
                </div>
            </div>
        </a>
        <a href="{make_href('expiring')}" target="_self" class="metric-card-link">
            <div class="metric-card" style="{border_expiring}">
                <div class="metric-icon icon-orange">
                    <svg width="28" height="28" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                </div>
                <div class="metric-content">
                    <div class="metric-title">即將到期數</div>
                    <div class="metric-value">{expiring_count:,}</div>
                    <div class="metric-subtext">佔比 {expiring_pct}%</div>
                </div>
            </div>
        </a>
        <a href="{make_href('expired')}" target="_self" class="metric-card-link">
            <div class="metric-card" style="{border_expired}">
                <div class="metric-icon icon-red">
                    <svg width="28" height="28" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                </div>
                <div class="metric-content">
                    <div class="metric-title">已過期數</div>
                    <div class="metric-value">{expired_count:,}</div>
                    <div class="metric-subtext">佔比 {expired_pct}%</div>
                </div>
            </div>
        </a>
    </div>
    """
    with metrics_placeholder:
        st.markdown(cards_html, unsafe_allow_html=True)

    # 套用選擇的狀態篩選到後續的圖表與資料
    if selected_status == 'valid':
        df_final = df_final[df_final['STATUS_LABEL'] == "有效"]
    elif selected_status == 'expiring':
        df_final = df_final[df_final['STATUS_LABEL'] == "快過期 (30天內)"]
    elif selected_status == 'expired':
        df_final = df_final[df_final['STATUS_LABEL'] == "已逾期"]
    
    st.markdown("---")
    
    # 圖表區塊
    chart_col1, chart_col2, chart_col3 = st.columns(3)
    
    with chart_col1:
        st.markdown(f"#### 【{prefix_text}】證照狀態分佈")
        st.markdown(f"<div style='font-size: 15px; color: #94a3b8; margin-top: -5px;'>🔄 數據更新至：{pd.Timestamp.now().strftime('%Y/%m/%d %H:%M')}</div>", unsafe_allow_html=True)
        status_counts = df_final['STATUS_LABEL'].value_counts().reset_index()
        status_counts.columns = ['狀態', '數量']
        
        if not status_counts.empty:
            total_s = status_counts['數量'].sum()
            status_counts['百分比'] = (status_counts['數量'] / total_s * 100).round(1).astype(str) + "%"
            
            color_map = {"有效": "#10b981", "快過期 (30天內)": "#f59e0b", "已逾期": "#ef4444", "未知": "#94a3b8"}
            domain_list = status_counts['狀態'].tolist()
            range_list = [color_map.get(s, "#94a3b8") for s in status_counts['狀態']]
            
            # 🌟 解法二：隱藏 Altair 原生圖例 (legend=None)
            base_pie = alt.Chart(status_counts).encode(
                theta=alt.Theta(field="數量", type="quantitative"),
                color=alt.Color(field="狀態", type="nominal", 
                                scale=alt.Scale(domain=domain_list, range=range_list),
                                legend=None),
                tooltip=['狀態', '數量', '百分比']
            )
            pie_chart = base_pie.mark_arc(innerRadius=80, outerRadius=115)
            
            text_val = alt.Chart(pd.DataFrame({'t': [f"{total_s:,}"]})).mark_text(size=42, fontWeight='bolder', color='#0F172A', dy=-12).encode(text='t:N')
            text_lbl = alt.Chart(pd.DataFrame({'t': ['總證照數']})).mark_text(size=18, fontWeight='bold', color='#475569', dy=25).encode(text='t:N')
            
            final_pie = (pie_chart + text_val + text_lbl).properties(height=300)
            st.altair_chart(final_pie, use_container_width=True)
            
            # 🌟 解法二：改用 DataFrame 作為 RWD 圖例
            st.dataframe(
                status_counts,
                column_config={
                    "狀態": st.column_config.TextColumn("狀態", width="medium"),
                    "數量": st.column_config.NumberColumn("數量", format="%d", width="small"),
                    "百分比": st.column_config.TextColumn("佔比", width="small")
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("尚無數據可繪製")
            
    with chart_col2:
        single_cat = (selected_category != "全部大類")
        title_text = f"【{prefix_text}】{selected_category} 內部證照分佈" if single_cat else f"【{prefix_text}】各大類證照分佈"
        st.markdown(f"#### {title_text}")
        st.markdown(f"<div style='font-size: 15px; color: #94a3b8; margin-top: -5px;'>🔄 數據更新至：{pd.Timestamp.now().strftime('%Y/%m/%d %H:%M')}</div>", unsafe_allow_html=True)
        
        is_single_category = not df_final.empty and df_final['CATEGORY'].nunique() == 1
        group_col = 'LICE_NAME' if is_single_category else 'CATEGORY'
        category_counts = df_final[group_col].value_counts().reset_index()
        # 移除限制，顯示所有類別
        category_counts.columns = ['分類名稱', '數量']
        
        if not category_counts.empty:
            total_c = category_counts['數量'].sum()
            category_counts['百分比'] = (category_counts['數量'] / total_c * 100).round(1).astype(str) + "%"
            
            # 🌟 解法二：隱藏 Altair 原生圖例 (legend=None)
            base_pie2 = alt.Chart(category_counts).encode(
                theta=alt.Theta(field="數量", type="quantitative"),
                color=alt.Color(field="分類名稱", type="nominal", 
                                scale=alt.Scale(scheme="category10"),
                                legend=None),
                tooltip=['分類名稱', '數量', '百分比']
            )
            pie_chart2 = base_pie2.mark_arc(innerRadius=80, outerRadius=115)
            
            text_val2 = alt.Chart(pd.DataFrame({'t': [f"{total_c:,}"]})).mark_text(size=42, fontWeight='bolder', color='#0F172A', dy=-12).encode(text='t:N')
            text_lbl2 = alt.Chart(pd.DataFrame({'t': ['總數量']})).mark_text(size=18, fontWeight='bold', color='#475569', dy=25).encode(text='t:N')
            
            final_pie2 = (pie_chart2 + text_val2 + text_lbl2).properties(height=300)
            st.altair_chart(final_pie2, use_container_width=True)
            
            # 🌟 解法二：改用 DataFrame 作為 RWD 圖例 (支援完美換行)
            st.dataframe(
                category_counts,
                column_config={
                    "分類名稱": st.column_config.TextColumn("證照/分類名稱", width="large"),
                    "數量": st.column_config.NumberColumn("數量", format="%d", width="small"),
                    "百分比": st.column_config.TextColumn("佔比", width="small")
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("尚無數據可繪製")
            
    with chart_col3:
        st.markdown(f"#### 【{prefix_text}】證照狀態分佈")
        st.markdown(f"<div style='font-size: 15px; color: #94a3b8; margin-top: -5px;'>🔄 數據更新至：{pd.Timestamp.now().strftime('%Y/%m/%d %H:%M')}</div>", unsafe_allow_html=True)
        
        is_single_category = not df_final.empty and df_final['CATEGORY'].nunique() == 1
        if not df_final.empty:
            group_col = 'LICE_NAME' if is_single_category else 'CATEGORY'
            
            top_cats = df_final[group_col].value_counts().index.tolist()
            df_stacked = df_final[df_final[group_col].isin(top_cats)]
            
            stacked_counts = df_stacked.groupby([group_col, 'STATUS_LABEL']).size().reset_index(name='數量')
            sort_map = {"有效": 1, "快過期 (30天內)": 2, "已逾期": 3, "未知": 4}
            stacked_counts['sort_order'] = stacked_counts['STATUS_LABEL'].map(sort_map)
            
            color_scale = alt.Scale(
                domain=["有效", "快過期 (30天內)", "已逾期", "未知"],
                range=["#10b981", "#f59e0b", "#ef4444", "#94a3b8"]
            )
            
            totals = stacked_counts.groupby(group_col)['數量'].sum().reset_index()
            totals.rename(columns={'數量': '總數'}, inplace=True)
            
            # 計算佔比與手動中心點 (為了完美置中文字)
            stacked_counts = pd.merge(stacked_counts, totals, on=group_col)
            stacked_counts['佔比'] = (stacked_counts['數量'] / stacked_counts['總數'] * 100).round(1).astype(str) + '%'
            stacked_counts['顯示標籤'] = (stacked_counts['數量'] / stacked_counts['總數'] * 100).round(0).astype(int).astype(str) + '%'
            
            # 依照圖表排序順序計算累積值，以求出 X 軸的完美中心點
            stacked_counts = stacked_counts.sort_values([group_col, 'sort_order'], ascending=[True, False])
            stacked_counts['x_end'] = stacked_counts.groupby(group_col)['數量'].cumsum()
            stacked_counts['x_start'] = stacked_counts['x_end'] - stacked_counts['數量']
            stacked_counts['x_mid'] = stacked_counts['x_start'] + (stacked_counts['數量'] / 2)
            
            # 增加 headroom，避免文字被切掉
            max_x = int(totals['總數'].max() * 1.25) if not totals.empty else 10
            
            base_bar = alt.Chart(stacked_counts).encode(
                y=alt.Y(f'{group_col}:N', sort=top_cats, axis=alt.Axis(title=None)),
                x=alt.X('數量:Q', title=None, scale=alt.Scale(domain=[0, max_x])),
                color=alt.Color('STATUS_LABEL:N', scale=color_scale, legend=alt.Legend(title=None, orient="top", direction="horizontal", labelFontSize=16, labelFontWeight="bold", symbolSize=200)),
                order=alt.Order('sort_order:Q', sort='descending')
            )
            
            # 加粗柱子，並增加乾淨的 Tooltip (包含佔比)
            bars = base_bar.mark_bar(size=40).encode(
                tooltip=[
                    alt.Tooltip(f'{group_col}:N', title='類別/名稱'),
                    alt.Tooltip('STATUS_LABEL:N', title='狀態'), 
                    alt.Tooltip('數量:Q', title='數量'),
                    alt.Tooltip('佔比:N', title='佔比')
                ]
            )
            
            # 動態計算閥值 (約佔全圖長度的 4%)，若數值太小則隱藏內部文字，避免擠壓重疊
            threshold = max_x * 0.04
            text_inside = alt.Chart(stacked_counts).mark_text(fontWeight='bold', size=16, color='white').encode(
                y=alt.Y(f'{group_col}:N', sort=top_cats),
                x=alt.X('x_mid:Q'),
                text=alt.condition(alt.datum.數量 > threshold, alt.Text('顯示標籤:N'), alt.value('')),
                tooltip=alt.value(None)
            )
            
            # 加大右側總數標籤
            total_text = alt.Chart(totals).mark_text(dx=10, align='left', fontWeight='bolder', color='#1e293b', size=18).encode(
                y=alt.Y(f'{group_col}:N', sort=top_cats),
                x=alt.X('總數:Q'),
                text='總數:Q',
                tooltip=alt.value(None)
            )
            
            # 3. 關鍵修改：將固定的 height=400 替換為 height=alt.Step(50)
            # 這樣每個長條會分配 50px 的高度，資料越多圖表就自動長越高！
            final_stacked = (bars + text_inside + total_text).properties(height=alt.Step(50)).configure_axis(labelFontSize=15)
            with st.container(height=450, border=True):
                st.altair_chart(final_stacked, use_container_width=True)
        else:
            st.info("尚無數據可繪製")
            
    st.markdown("---")
    
    # 🌟 新增：各證照類別統計 & 近期證照申請統計
    new_col1, new_col2 = st.columns(2)
    
    with new_col1:
        st.markdown(f"#### 【{prefix_text}】部門證照有效率排行")
        st.markdown(f"<div style='font-size: 15px; color: #94a3b8; margin-top: -5px;'>🔄 數據更新至：{pd.Timestamp.now().strftime('%Y/%m/%d %H:%M')}</div>", unsafe_allow_html=True)
        if not df_final.empty:
            cat_stats = df_final.groupby('ADM_DEPT_NAME').agg(
                總證照數=('STATUS_LABEL', 'count'),
                有效證照數=('STATUS_LABEL', lambda x: (x == '有效').sum()),
                逾期證照數=('STATUS_LABEL', lambda x: (x == '已逾期').sum())
            ).reset_index()
            
            cat_stats['各科別有效證照率'] = (cat_stats['有效證照數'] / cat_stats['總證照數']) * 100
            cat_stats['有效/總數'] = cat_stats['有效證照數'].astype(str) + " / " + cat_stats['總證照數'].astype(str)
            cat_stats['逾期率 (對比)'] = (cat_stats['逾期證照數'] / cat_stats['總證照數']) * 100
            
            cat_stats.rename(columns={'ADM_DEPT_NAME': '部門'}, inplace=True)
            
            display_stats = cat_stats[['部門', '各科別有效證照率', '有效/總數', '逾期率 (對比)', '逾期證照數']].sort_values(by='各科別有效證照率', ascending=False)
            
            st.dataframe(
                display_stats,
                column_config={
                    "各科別有效證照率": st.column_config.ProgressColumn(
                        "各科別有效證照率",
                        help="有效證照佔比",
                        format="%.1f%%",
                        min_value=0,
                        max_value=100,
                    ),
                    "有效/總數": st.column_config.TextColumn(
                        "有效證照數 / 總數"
                    ),
                    "逾期率 (對比)": st.column_config.NumberColumn(
                        "逾期率 (對比)",
                        help="已逾期的佔比",
                        format="%.1f%%",
                    ),
                    "逾期證照數": st.column_config.NumberColumn(
                        "已逾期數"
                    )
                },
                hide_index=True,
                use_container_width=True,
                height=250
            )
        else:
            st.info("尚無數據可統計")

    with new_col2:
        col_title, col_filter = st.columns([0.45, 0.55], vertical_alignment="bottom")
        with col_title:
            st.markdown(f"#### 【{prefix_text}】近期證照申請統計")
        with col_filter:
            f_col1, f_col2 = st.columns([0.4, 0.6])
            with f_col1:
                st.markdown('<div id="freq-pills-marker"></div>', unsafe_allow_html=True)
                st.markdown("""
                    <style>
                    /* 透過隱藏的 marker 精準鎖定該 Column 內的 Pills，強制排成 2x2 正方形 */
                    div[data-testid="stVerticalBlock"]:has(#freq-pills-marker) [data-testid="stPills"] [data-testid="stButtonGroup"] {
                        display: flex !important;
                        flex-wrap: wrap !important;
                        gap: 6px !important;
                        width: 100% !important;
                    }
                    div[data-testid="stVerticalBlock"]:has(#freq-pills-marker) [data-testid="stPills"] button {
                        width: calc(50% - 6px) !important;
                        flex: none !important;
                        padding: 6px 0px !important;
                        margin: 0 !important;
                        justify-content: center !important;
                    }
                    div[data-testid="stVerticalBlock"]:has(#freq-pills-marker) [data-testid="stPills"] button p,
                    div[data-testid="stVerticalBlock"]:has(#freq-pills-marker) [data-testid="stPills"] button span {
                        font-size: 1rem !important;
                        margin: 0 !important;
                        padding: 0 !important;
                    }
                    </style>
                """, unsafe_allow_html=True)
                freq_choice = st.pills("維度", ["按日", "按周", "按月", "按年"], default="按月", label_visibility="collapsed", key="freq_choice")
            with f_col2:
                chart_date = st.date_input("區間", value=(), key="chart_date_filter", label_visibility="collapsed")
        
        # ====== 實作策略一：以申請日為基準的真實數據分析 ======
        # 1. 決定時間區間
        if len(chart_date) == 2:
            start_d, end_d = chart_date
            start_dt = pd.to_datetime(start_d)
            end_dt = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        else:
            end_dt = pd.Timestamp.now()
            if freq_choice == "按日": start_dt = end_dt - pd.DateOffset(days=6)
            elif freq_choice == "按周": start_dt = end_dt - pd.DateOffset(weeks=5)
            elif freq_choice == "按月": start_dt = end_dt - pd.DateOffset(months=12) # 🌟 放大預設為 12 個月
            else: start_dt = end_dt - pd.DateOffset(years=10) # 🌟 放大預設為 10 年，包容歷史資料
            
        start_dt = start_dt.normalize()
        
        # 2. 決定 pandas frequency，並建立完整的時間基準，避免某些時段無資料而消失
        freq_map = {"按日": "D", "按周": "W-MON", "按月": "MS", "按年": "YS"}
        pandas_freq = freq_map.get(freq_choice, "MS")
        
        idx_start = start_dt
        if freq_choice == "按月": idx_start = start_dt.replace(day=1)
        elif freq_choice == "按年": idx_start = start_dt.replace(month=1, day=1)
        
        full_idx = pd.date_range(start=idx_start, end=end_dt, freq=pandas_freq)
        if len(full_idx) == 0: full_idx = pd.DatetimeIndex([idx_start])
        
        trend_df = pd.DataFrame({'APP_DATE': full_idx, '申請數': 0.0, '通過數': 0.0})
        trend_df = trend_df.set_index('APP_DATE')
        
        # 3. 過濾真實資料並透過 APP_DATE 進行 Resample 統計
        if 'APP_DATE' in df_final.columns:
            df_actual = df_final.dropna(subset=['APP_DATE']).copy()
            df_actual = df_actual[(df_actual['APP_DATE'] >= start_dt) & (df_actual['APP_DATE'] <= end_dt)]
            
            if not df_actual.empty:
                # 定義：該申請案後續狀態若為「有效」，即視為已通過 (世代分析)
                df_actual['is_passed'] = (df_actual['STATUS_LABEL'] == '有效').astype(int)
                
                actual_agg = df_actual.set_index('APP_DATE').resample(pandas_freq).agg(
                    申請數=('is_passed', 'count'),
                    通過數=('is_passed', 'sum')
                )
                
                # 將真實統計資料更新到完整時間軸上
                trend_df.update(actual_agg)
                
        trend_df = trend_df.reset_index()
        
        # 4. 產生 X 軸時間標籤與最終陣列
        def format_time_label(d, f):
            if f == "按日": 
                return d.strftime('%m/%d')
            elif f == "按周": 
                # 計算當月第幾周 (以該月1號為基準)
                first_day = d.replace(day=1)
                adjusted_dom = d.day - 1 + first_day.weekday()
                week_num = int(adjusted_dom / 7) + 1
                return f"{d.year}年{d.month}月第{week_num}周"
            elif f == "按月": 
                return d.strftime('%Y/%m')
            else: 
                return d.strftime('%Y年')
            
        time_labels = trend_df['APP_DATE'].apply(lambda x: format_time_label(x, freq_choice)).tolist()
        app_counts = trend_df['申請數'].fillna(0).astype(int).tolist()
        pass_counts = trend_df['通過數'].fillna(0).astype(int).tolist()
            
        mock_trend = pd.DataFrame({
            '時間': time_labels,
            '申請數': app_counts,
            '通過數': pass_counts
        })
        trend_melted = mock_trend.melt('時間', var_name='統計項目', value_name='數量')
        
        # 🌟 優化 2：動態 Y 軸與平滑曲線視覺升級
        max_val3 = int(trend_melted['數量'].max())
        y_max = int(max_val3 * 1.35) if max_val3 > 0 else 5  # 避免全 0 時 Y 軸太扁
        
        base_line = alt.Chart(trend_melted).encode(
            x=alt.X('時間', axis=alt.Axis(labelAngle=0), title=None),
            y=alt.Y('數量:Q', title=None, scale=alt.Scale(domain=[0, y_max])),
            color=alt.Color('統計項目', scale=alt.Scale(domain=['申請數', '通過數'], range=['#3B82F6', '#10B981']), legend=alt.Legend(title=None, orient="top-right")),
            tooltip=['時間', '統計項目', '數量']
        )
        
        # 加入 interpolate='monotone' 讓折線變成企業級的平滑曲線
        line_chart = base_line.mark_line(interpolate='monotone', point=alt.OverlayMarkDef(size=80, color="white", strokeWidth=2))
        
        # 優化數值標籤，若數值為 0 則稍微淡化，凸顯有數據的月份
        text_labels = base_line.mark_text(
            clip=False, align='center', baseline='bottom', dy=-12, fontWeight='bold', size=15
        ).encode(
            text='數量:Q',
            color=alt.condition(alt.datum.數量 == 0, alt.value('#CBD5E1'), alt.value('#1E293B'))
        )
        
        final_line = (line_chart + text_labels).properties(height=250).configure_axis(
            labelFontSize=13, titleFontSize=15, gridColor="#F1F5F9"
        ).configure_legend(labelFontSize=13, titleFontSize=14)
        
        st.altair_chart(final_line, use_container_width=True)
            
    st.markdown("---")
    list_title_col, list_search_col = st.columns([0.7, 0.3], vertical_alignment="bottom")
    with list_title_col:
        st.markdown(f"#### 【{prefix_text}】證照詳細清單")
        st.markdown(f"<div style='font-size: 15px; color: #94a3b8; margin-top: -5px;'>🔄 數據更新至：{pd.Timestamp.now().strftime('%Y/%m/%d %H:%M')}</div>", unsafe_allow_html=True)
    with list_search_col:
        person_query = st.text_input("🔍 搜尋姓名或卡號", placeholder="輸入姓名或卡號...", label_visibility="collapsed")
        
    df_list = df_final.copy()
    if not df_list.empty and person_query:
        # 支援姓名與卡號的模糊搜尋
        df_list = df_list[df_list['name_display'].str.contains(person_query, case=False, na=False) | df_list['CARDNO'].astype(str).str.contains(person_query, case=False, na=False)]
        
    if not df_list.empty:
        # 計算剩餘天數
        df_list['剩餘天數'] = (df_list['DATEE_PARSED'] - now).dt.days
        df_list['到期日'] = df_list['DATEE_PARSED'].dt.strftime('%Y-%m-%d')
        df_list_display = df_list[['LICE_NAME', 'name_display', 'ADM_DEPT_NAME', '到期日', '剩餘天數', 'STATUS_LABEL']].copy()
        df_list_display['剩餘天數'] = df_list_display['剩餘天數'].fillna(0).astype(int)
        df_list_display.columns = ['證照名稱', '持有人', '所屬科別', '到期日', '剩餘天數', '狀態']
        
        # 依照狀態與剩餘天數排序，讓緊急的排前面
        df_list_display = df_list_display.sort_values(by=['狀態', '剩餘天數'])
        
        # 凸顯各種狀態顏色
        def highlight_status(val):
            if val == '已逾期': return 'color: #ef4444; font-weight: bold'
            if val == '快過期 (30天內)': return 'color: #f59e0b; font-weight: bold'
            if val == '有效': return 'color: #10b981; font-weight: bold'
            return 'color: #64748b'
            
        st.dataframe(df_list_display.style.map(highlight_status, subset=['狀態']), use_container_width=True, hide_index=True)
    else:
        st.info("查無符合條件的證照資料。")

# 執行儀表板
if __name__ == "__main__":
    main()
