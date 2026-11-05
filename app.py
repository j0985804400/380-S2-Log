import json
import math
import io
from datetime import datetime, timedelta, timezone
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
import pandas as pd
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ================= 系統安全轉型防呆函數 =================
def safe_int(val, default=0):
    try:
        if val is None or str(val).strip() == "" or str(val) == "None":
            return default
        return int(float(val))
    except:
        return default

def safe_float(val, default=0.0):
    try:
        if val is None or str(val).strip() == "" or str(val) == "None":
            return default
        return float(val)
    except:
        return default

def parse_date_safe(date_str, default_date):
    if not date_str or str(date_str).strip() == "" or str(date_str) == "None":
        return default_date
    try:
        clean_str = str(date_str).strip().replace("/", "-")
        return datetime.strptime(clean_str, "%Y-%m-%d").date()
    except:
        return default_date

# ================= 系統固定參數與時區設定 =================
TW_TZ = timezone(timedelta(hours=8))
now_time = datetime.now(TW_TZ)
TODAY = now_time.date()
if now_time.hour < 4:
    SHIFT_DATE = (now_time - timedelta(days=1)).date()
else:
    SHIFT_DATE = TODAY

# ================= 車隊多車架構 (純車牌導向，不需手動維護容量與尿素) =================
DEFAULT_FLEET = {
    "380-S2": {
        "plate": "380-S2",
        "keeper": "劉嘉舜",
        "default_user": "劉嘉舜",
        "doc_code": "4-33-DC0007B",
        "form_name": "公務車輛行車紀錄表 Rev.1",
        "sheet_name": "380-S2_公務車行車紀錄系統"
    }
}

if "fleet_registry" not in st.session_state:
    st.session_state.fleet_registry = DEFAULT_FLEET.copy()

# ================= 前端介面基礎設定 =================
st.set_page_config(page_title="車隊管理總控台", layout="centered", initial_sidebar_state="collapsed")

st.sidebar.title("🚚 車隊管理中心")
vehicle_list = list(st.session_state.fleet_registry.keys())
sidebar_options = vehicle_list + ["➕ 新增車輛 (擴充專用)"]
selected_vehicle = st.sidebar.selectbox("選擇管理車輛", sidebar_options, index=0)

# 新增車輛邏輯 (純車牌導向，極簡填寫)
if selected_vehicle == "➕ 新增車輛 (擴充專用)":
    st.title("➕ 新增車輛 (支援 50 台車隊擴充)")
    st.markdown("僅需輸入車牌號碼與保管人，系統將自動配置專屬行車資料庫：")
    with st.form("new_vehicle_form"):
        new_plate = st.text_input("車牌號碼 (如：380-S2 / ABC-1234)").strip().upper()
        c1, c2 = st.columns(2)
        with c1:
            new_keeper = st.text_input("保管人姓名", value="劉嘉舜")
        with c2:
            new_user = st.text_input("預設駕駛人", value="劉嘉舜")
            
        custom_sheet = st.text_input("對應 Google Sheets 試算表名稱 (選填，預設為 車牌_公務車行車紀錄系統)", value="")
        
        btn_register = st.form_submit_button("🚀 完成註冊並切換至此車輛", use_container_width=True)
        if btn_register:
            if not new_plate:
                st.error("請填寫車牌號碼！")
            else:
                st.session_state.fleet_registry[new_plate] = {
                    "plate": new_plate,
                    "keeper": new_keeper if new_keeper else "公務保管人",
                    "default_user": new_user if new_user else "公務使用人",
                    "doc_code": "4-33-DC0007B",
                    "form_name": "公務車輛行車紀錄表 Rev.1",
                    "sheet_name": custom_sheet.strip() if custom_sheet.strip() else f"{new_plate}_公務車行車紀錄系統"
                }
                st.success(f"車輛 【{new_plate}】 註冊成功！已切換至該車。")
                st.rerun()
    st.stop()

# 當前選取車輛參數
CURRENT_VEHICLE = st.session_state.fleet_registry[selected_vehicle]
DEFAULT_PLATE = CURRENT_VEHICLE["plate"]
DEFAULT_KEEPER = CURRENT_VEHICLE["keeper"]
DEFAULT_USER = CURRENT_VEHICLE["default_user"]
DOC_CODE = CURRENT_VEHICLE.get("doc_code", "4-33-DC0007B")
FORM_NAME = CURRENT_VEHICLE.get("form_name", "公務車輛行車紀錄表 Rev.1")
ACTIVE_SHEET_NAME = CURRENT_VEHICLE.get("sheet_name", f"{DEFAULT_PLATE}_公務車行車紀錄系統")

st.sidebar.markdown(f"**目前監控車牌**：`{DEFAULT_PLATE}`")
st.sidebar.caption(f"👤 保管人：{DEFAULT_KEEPER}")

# ================= Google API 連線池快取化 =================
@st.cache_resource(ttl=86400)
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "google_key" in st.secrets:
        creds_dict = json.loads(st.secrets["google_key"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    return gspread.authorize(creds)

@st.cache_resource(ttl=3600)
def get_spreadsheet(sheet_name):
    client = get_gspread_client()
    return client.open(sheet_name)

def get_worksheet(sheet_name, ws_name):
    sh = get_spreadsheet(sheet_name)
    return sh.worksheet(ws_name)

# ================= 智慧欄位名稱對齊與雙向安全寫入 =================
def normalize_records(records, mapping):
    norm_list = []
    for r in records:
        nr = {}
        for std_k, pos_k in mapping.items():
            found = False
            for pk in pos_k:
                if pk in r and r[pk] is not None and str(r[pk]).strip() != "":
                    nr[std_k] = r[pk]
                    found = True
                    break
            if not found:
                nr[std_k] = ""
        for k, v in r.items():
            if k not in nr:
                nr[k] = v
        norm_list.append(nr)
    return norm_list

daily_map = {
    "日期": ["日期", "時間", "紀錄日期"],
    "使用人": ["使用人", "駕駛"],
    "目的地": ["目的地", "地點"],
    "出發時間": ["出發時間", "紀錄時間"],
    "出發里程": ["出發里程", "出發里程數"],
    "返廠時間": ["返廠時間"],
    "返廠里程": ["返廠里程", "返廠里程數"],
    "備註": ["備註"]
}

fuel_map = {
    "加油日期": ["加油日期", "日期", "紀錄日期"],
    "加油時間": ["加油時間", "時間", "紀錄時間"],
    "加油里程": ["加油里程", "里程", "當下里程", "紀錄里程"],
    "公升數": ["公升數", "公升"],
    "加油費用": ["加油費用", "費用", "金額", "實付金額", "柴油費用"],
    "加油站名": ["加油站名", "加油站", "站名"]
}

urea_map = {
    "添加日期": ["添加日期", "日期", "加油日期", "紀錄日期"],
    "添加時間": ["添加時間", "時間", "紀錄時間"],
    "公升數": ["公升數", "公升", "添加量"],
    "費用": ["費用", "金額", "實付金額", "尿素費用"],
    "備註": ["備註", "說明"],
    "添加當下里程": ["添加當下里程", "添加里程", "里程", "當下里程", "紀錄里程"]
}

def safe_append_row(sheet_name, ws_name, field_dict, default_order, alias_map=None):
    """依照試算表第 1 列實際標題動態對齊寫入，結合別名字典防止錯位"""
    try:
        ws = get_worksheet(sheet_name, ws_name)
        headers = ws.row_values(1)
        if headers and len(headers) >= len(default_order):
            row_to_write = []
            for h in headers:
                val = field_dict.get(h, None)
                if val is None and alias_map:
                    for std_k, aliases in alias_map.items():
                        if h in aliases and std_k in field_dict:
                            val = field_dict[std_k]
                            break
                row_to_write.append(val if val is not None else "")
            ws.append_row(row_to_write)
        else:
            row_to_write = [field_dict.get(k, "") for k in default_order]
            ws.append_row(row_to_write)
    except Exception as e:
        ws = get_worksheet(sheet_name, ws_name)
        row_to_write = [field_dict.get(k, "") for k in default_order]
        ws.append_row(row_to_write)

# ================= 快取資料讀取 =================
@st.cache_data(ttl=600, show_spinner="連線雲端同步中...")
def fetch_all_data(sheet_name):
    client = get_gspread_client()
    try:
        sh = client.open(sheet_name)
    except gspread.SpreadsheetNotFound:
        sh = client.create(sheet_name)
        
    sheet_titles = [w.title for w in sh.worksheets()]
    
    if "Settings" not in sheet_titles:
        ws_settings = sh.add_worksheet(title="Settings", rows=10, cols=2)
        default_vals = [
            ["SERVICE_TARGET", 622000],
            ["FUEL_TANK_CAPACITY", 160.0],
            ["NEXT_VEHICLE_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(months=6)).strftime("%Y-%m-%d")],
            ["NEXT_EMISSIONS_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(years=1)).strftime("%Y-%m-%d")]
        ]
        ws_settings.update(values=default_vals, range_name="A1:B4")
        
    if "Daily_Logs" not in sheet_titles:
        ws_d = sh.add_worksheet(title="Daily_Logs", rows=100, cols=8)
        ws_d.append_row(["日期", "使用人", "目的地", "出發時間", "出發里程", "返廠時間", "返廠里程", "備註"])
        
    if "Fuel_Logs" not in sheet_titles:
        ws_f = sh.add_worksheet(title="Fuel_Logs", rows=100, cols=6)
        ws_f.append_row(["加油日期", "加油時間", "加油里程", "公升數", "加油費用", "加油站名"])
        
    if "Urea_Logs" not in sheet_titles:
        ws_u = sh.add_worksheet(title="Urea_Logs", rows=100, cols=6)
        ws_u.append_row(["添加日期", "添加時間", "公升數", "費用", "備註", "添加當下里程"])

    raw_settings = sh.worksheet("Settings").get_all_values()
    raw_daily = sh.worksheet("Daily_Logs").get_all_records()
    raw_fuel = sh.worksheet("Fuel_Logs").get_all_records()
    
    raw_urea = []
    if "Urea_Logs" in sheet_titles:
        try:
            raw_urea = sh.worksheet("Urea_Logs").get_all_records()
        except:
            raw_urea = []
            
    return raw_settings, raw_daily, raw_fuel, raw_urea

try:
    raw_settings, raw_daily, raw_fuel, raw_urea = fetch_all_data(ACTIVE_SHEET_NAME)
    st.session_state.settings_dict = {row[0]: row[1] for row in raw_settings if len(row) >= 2}
    st.session_state.daily_logs = normalize_records(raw_daily, daily_map) if raw_daily else []
    st.session_state.fuel_logs = normalize_records(raw_fuel, fuel_map) if raw_fuel else []
    st.session_state.urea_logs = normalize_records(raw_urea, urea_map) if raw_urea else []
    
    if "destinations" not in st.session_state:
        st.session_state.destinations = ["美光", "順益", "廠內調度"]
except Exception as e:
    st.error(f"資料讀取失敗，請確認 Google 試算表【{ACTIVE_SHEET_NAME}】已分享給機器人金鑰或設定正確。錯誤細節: {e}")
    st.stop()

# ================= 參數與動態油箱容量學習演算法 =================
st.session_state.service_target = safe_int(st.session_state.settings_dict.get("SERVICE_TARGET"), 622000)

next_v_insp = st.session_state.settings_dict.get("NEXT_VEHICLE_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(months=6)).strftime("%Y-%m-%d"))
next_e_insp = st.session_state.settings_dict.get("NEXT_EMISSIONS_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(years=1)).strftime("%Y-%m-%d"))
next_v_date = parse_date_safe(next_v_insp, TODAY)
next_e_date = parse_date_safe(next_e_insp, TODAY)
days_to_v = (next_v_date - TODAY).days
days_to_e = (next_e_date - TODAY).days

# 【Bug 修正：最新里程與對應紀錄日期精準提取】
def get_verified_latest_mileage_and_date(daily_logs, fuel_logs, urea_logs):
    valid_entries = []
    for d in daily_logs:
        mil = safe_int(d.get("返廠里程")) or safe_int(d.get("出發里程"))
        date_str = str(d.get("日期", "")).replace("/", "-")
        time_str = str(d.get("返廠時間", "23:59"))
        if 1000 <= mil <= 1500000 and date_str:
            valid_entries.append((date_str, time_str, mil))
            
    for f in fuel_logs:
        mil = safe_int(f.get("加油里程"))
        date_str = str(f.get("加油日期", "")).replace("/", "-")
        time_str = str(f.get("加油時間", "23:59"))
        if 1000 <= mil <= 1500000 and date_str:
            valid_entries.append((date_str, time_str, mil))
            
    for u in urea_logs:
        mil = safe_int(u.get("添加當下里程"))
        date_str = str(u.get("添加日期", "")).replace("/", "-")
        time_str = str(u.get("添加時間", "23:59"))
        if 1000 <= mil <= 1500000 and date_str:
            valid_entries.append((date_str, time_str, mil))
            
    if valid_entries:
        valid_entries.sort(key=lambda x: (x[0], x[1]))
        latest_date = parse_date_safe(valid_entries[-1][0], TODAY)
        return valid_entries[-1][2], latest_date
    return (615215 if DEFAULT_PLATE == "380-S2" else 100000), TODAY

latest_return_mileage, latest_mil_date = get_verified_latest_mileage_and_date(
    st.session_state.daily_logs,
    st.session_state.fuel_logs,
    st.session_state.urea_logs
)
rem_service = st.session_state.service_target - latest_return_mileage

# ================= 柴油消耗運算 (基準錨點零漂移演算法) =================
sorted_fuel_logs = sorted(st.session_state.fuel_logs, key=lambda x: safe_int(x.get("加油里程", 0)))
latest_fuel_mileage = safe_int(sorted_fuel_logs[-1].get("加油里程")) if sorted_fuel_logs else latest_return_mileage

global_fuel_km_l = 3.8
if len(sorted_fuel_logs) >= 2:
    last_f = sorted_fuel_logs[-1]
    prev_f = sorted_fuel_logs[-2]
    km_driven = safe_int(last_f.get("加油里程")) - safe_int(prev_f.get("加油里程"))
    liters = safe_float(last_f.get("公升數"))
    if liters > 0 and km_driven > 0:
        global_fuel_km_l = round(km_driven / liters, 2)
global_fuel_km_l = max(0.5, global_fuel_km_l)

historical_km_l = global_fuel_km_l
if len(sorted_fuel_logs) >= 3:
    first_mil = safe_int(sorted_fuel_logs[0].get("加油里程"))
    last_mil = safe_int(sorted_fuel_logs[-1].get("加油里程"))
    total_km_hist = last_mil - first_mil
    total_liters_hist = sum(safe_float(x.get("公升數")) for x in sorted_fuel_logs[1:])
    if total_liters_hist > 0 and total_km_hist > 0:
        historical_km_l = round(total_km_hist / total_liters_hist, 2)

def get_avg_daily_km(logs):
    sorted_logs = sorted(logs, key=lambda x: str(x.get("日期", "")))
    recent_logs = sorted_logs[-30:] if len(sorted_logs) >= 30 else sorted_logs
    if len(recent_logs) >= 2:
        try:
            first_date = parse_date_safe(recent_logs[0]["日期"], TODAY - timedelta(days=1))
            last_date = parse_date_safe(recent_logs[-1]["日期"], TODAY)
            days_diff = (last_date - first_date).days + 1
            if days_diff > 0:
                mileage_diff = safe_int(recent_logs[-1]["返廠里程"]) - safe_int(recent_logs[0]["出發里程"])
                if mileage_diff > 0:
                    avg = mileage_diff / days_diff
                    return max(1.0, avg)
        except:
            pass
    return 170.0

avg_daily_km = get_avg_daily_km(st.session_state.daily_logs)

# 油箱容量動態自動學習並鎖定
max_liters_added = max([safe_float(f.get("公升數")) for f in sorted_fuel_logs]) if sorted_fuel_logs else 130.0
learned_tank_cap = float(math.ceil(max_liters_added / 10.0) * 10)
settings_tank_cap = safe_float(st.session_state.settings_dict.get("FUEL_TANK_CAPACITY"), 130.0)
tank_cap = max(130.0, settings_tank_cap, learned_tank_cap)

# 即時剩餘油量推估
used_km_since_fuel = max(0, latest_return_mileage - latest_fuel_mileage)
used_vol_since_fuel = used_km_since_fuel / global_fuel_km_l
vol_after_today = max(0.0, min(tank_cap, tank_cap - used_vol_since_fuel))

# ================= 尿素消耗運算 =================
UREA_TANK_CAPACITY = 20.0
urea_km_l = 170.0
rem_urea_vol = UREA_TANK_CAPACITY
driven_since_urea = 0
HAS_UREA = len(st.session_state.urea_logs) > 0 or DEFAULT_PLATE == "380-S2"

if HAS_UREA and st.session_state.urea_logs:
    sorted_urea = sorted(st.session_state.urea_logs, key=lambda x: str(x.get("添加日期", "")))
    urea_with_mil = [u for u in sorted_urea if str(u.get("添加當下里程", "")).isdigit()]
    urea_with_mil = sorted(urea_with_mil, key=lambda x: safe_int(x["添加當下里程"]))
    
    if len(urea_with_mil) >= 2:
        first_u_mil = safe_int(urea_with_mil[0]["添加當下里程"])
        last_u_mil = safe_int(urea_with_mil[-1]["添加當下里程"])
        total_u_km = last_u_mil - first_u_mil
        total_u_liters = sum(safe_float(x.get("公升數")) for x in urea_with_mil[1:])
        if total_u_liters > 0 and total_u_km > 0:
            urea_km_l = round(total_u_km / total_u_liters, 2)
    urea_km_l = max(10.0, urea_km_l)
    
    last_u_mil = safe_int(urea_with_mil[-1]["添加當下里程"]) if urea_with_mil else latest_return_mileage
    driven_since_urea = max(0, latest_return_mileage - last_u_mil)
    used_urea_since_last = driven_since_urea / urea_km_l
    rem_urea_vol = max(0.0, min(UREA_TANK_CAPACITY, UREA_TANK_CAPACITY - used_urea_since_last))

def calculate_service_prediction(target_mileage):
    km_remaining = target_mileage - latest_return_mileage
    if km_remaining <= 0:
        return "已達保養里程"
    trips_remaining = math.ceil(km_remaining / avg_daily_km)
    est_date = latest_mil_date + timedelta(days=trips_remaining)
    return est_date.strftime("%m/%d")

est_service_date = calculate_service_prediction(st.session_state.service_target)

# ================= 1:1 行高極限特化 A4 報表生成 =================
def generate_kertz_excel(year: int, month: int):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{month}月行車紀錄"
    ws.views.sheetView[0].showGridLines = False
    
    font_family = "Microsoft JhengHei"
    font_bold_md = Font(name=font_family, size=11, bold=True)
    font_regular = Font(name=font_family, size=11)
    font_header_sub = Font(name=font_family, size=10)
    
    thin = Side(border_style="thin", color="000000")
    border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=False)
    
    ws.page_margins.top = 0.8
    ws.page_margins.bottom = 0.3 
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.print_options.horizontalCentered = True  

    def apply_outer_border(ws, r_start, c_start, r_end, c_end):
        for r in range(r_start, r_end + 1):
            for c in range(c_start, c_end + 1):
                top_side = thin if r == r_start else None
                bot_side = thin if r == r_end else None
                left_side = thin if c == c_start else None
                right_side = thin if c == c_end else None
                ws.cell(row=r, column=c).border = Border(top=top_side, bottom=bot_side, left=left_side, right=right_side)

    merges = [(1, 1, 1, 7), (1, 8, 1, 9), (2, 1, 2, 2), (2, 3, 2, 7), (2, 8, 2, 9), (3, 1, 3, 9)]
    for r1, c1, r2, c2 in merges:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        apply_outer_border(ws, r1, c1, r2, c2)

    ws["A1"] = "KT.T. Pro Clean Tech  科治新技股份有限公司 標準作業表格"
    ws["A1"].font = font_bold_md
    ws["A1"].alignment = align_center

    ws["H1"] = "頁次："
    ws["H1"].font = font_header_sub
    ws["H1"].alignment = align_left
    
    ws["A2"] = "表格名稱"
    ws["A2"].font = Font(name=font_family, size=12)
    ws["A2"].alignment = align_center

    ws["C2"] = FORM_NAME
    ws["C2"].font = Font(name=font_family, size=12)
    ws["C2"].alignment = align_center

    ws["H2"] = f"編號：{DOC_CODE}"
    ws["H2"].font = Font(name=font_family, size=10)
    ws["H2"].alignment = align_left

    roc_year = year - 1911
    ws["A3"] = f"  行車日誌               月份： {roc_year} 年 {month} 月          車牌號碼： {DEFAULT_PLATE}          保管人： {DEFAULT_KEEPER}"
    ws["A3"].font = font_bold_md
    ws["A3"].alignment = align_left 

    headers = ["日期", "使用人", "目的地", "出發時間", "出發里程數", "返廠時間", "返廠里程數", "加油里程數", "加油費用"]
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=text)
        cell.font = font_bold_md
        cell.alignment = align_center 
        cell.border = border_all

    for r in range(1, 5):
        ws.row_dimensions[r].height = 25 

    prefix = f"{year:04d}-{month:02d}"
    m_logs = [l for l in st.session_state.daily_logs if str(l.get("日期", "")).startswith(prefix)]
    m_logs = sorted(m_logs, key=lambda x: str(x.get("日期", "")))
    
    f_map_grouped = {}
    for fl in st.session_state.fuel_logs:
        d_str = str(fl.get("加油日期", "")).strip()
        if d_str.startswith(prefix):
            f_map_grouped.setdefault(d_str, []).append(fl)

    combined_rows = []
    for log in m_logs:
        d_key = log.get("日期", "")
        f_info = f_map_grouped[d_key].pop(0) if (d_key in f_map_grouped and f_map_grouped[d_key]) else None
        combined_rows.append({
            "日期": d_key,
            "使用人": log.get("使用人", ""),
            "目的地": log.get("目的地", ""),
            "出發時間": log.get("出發時間", ""),
            "出發里程": log.get("出發里程", ""),
            "返廠時間": log.get("返廠時間", ""),
            "返廠里程": log.get("返廠里程", ""),
            "加油里程": f_info.get("加油里程", "") if f_info else "",
            "加油費用": f_info.get("加油費用", "") if f_info else ""
        })

    for d_key, leftover_fuels in f_map_grouped.items():
        for fl in leftover_fuels:
            combined_rows.append({
                "日期": d_key,
                "使用人": DEFAULT_KEEPER,
                "目的地": "加油補給",
                "出發時間": fl.get("加油時間", ""),
                "出發里程": "",
                "返廠時間": "",
                "返廠里程": "",
                "加油里程": fl.get("加油里程", ""),
                "加油費用": fl.get("加油費用", "")
            })

    combined_rows.sort(key=lambda x: str(x.get("日期", "")))
    
    start_row = 5
    total_print_rows = max(31, len(combined_rows))
    
    for i in range(total_print_rows):
        curr_row = start_row + i
        ws.row_dimensions[curr_row].height = 21.5 
        
        if i < len(combined_rows):
            row_data = combined_rows[i]
            d_val = row_data["日期"]
            try:
                day_num = int(d_val.split("-")[2])
            except:
                day_num = ""
                
            ws.cell(row=curr_row, column=1, value=str(day_num)).alignment = align_center
            ws.cell(row=curr_row, column=2, value=str(row_data.get("使用人", ""))).alignment = align_center
            ws.cell(row=curr_row, column=3, value=str(row_data.get("目的地", ""))).alignment = align_center
            ws.cell(row=curr_row, column=4, value=str(row_data.get("出發時間", ""))).alignment = align_center
            
            st_mil = safe_int(row_data.get("出發里程"))
            ws.cell(row=curr_row, column=5, value=st_mil if st_mil > 0 else "").alignment = align_right
            
            ws.cell(row=curr_row, column=6, value=str(row_data.get("返廠時間", ""))).alignment = align_center
            
            ret_mil = safe_int(row_data.get("返廠里程"))
            ws.cell(row=curr_row, column=7, value=ret_mil if ret_mil > 0 else "").alignment = align_right
            
            f_mil = safe_int(row_data.get("加油里程"))
            ws.cell(row=curr_row, column=8, value=f_mil if f_mil > 0 else "").alignment = align_right
            
            f_cost = safe_int(row_data.get("加油費用"))
            ws.cell(row=curr_row, column=9, value=f_cost if f_cost > 0 else "").alignment = align_right
        
        for c in range(1, 10):
            cell = ws.cell(row=curr_row, column=c)
            if not cell.font:
                cell.font = font_regular
            cell.border = border_all
            if c in [5, 7, 8, 9] and isinstance(cell.value, (int, float)) and cell.value > 0:
                cell.number_format = "#,##0"

    col_widths = {
        "A": 5.0,
        "B": 15.0,  
        "C": 9.0,   
        "D": 11.0,  
        "E": 12.5,  
        "F": 11.0,  
        "G": 12.5,  
        "H": 12.0,  
        "I": 11.0   
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0 
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ================= 前端儀表板 UI =================
st.markdown(f"""
<style>
    .main-title {{ font-size: 28px !important; font-weight: 800; margin-bottom: 5px; white-space: nowrap; }}
</style>
<div class='main-title'>{DEFAULT_PLATE} 車隊管理</div>
""", unsafe_allow_html=True)
st.caption(f"連線狀態：🟢 雲端快取極速模式 | 👤 保管人：{DEFAULT_KEEPER}")

svc_color = "#00c04b" if "月" not in est_service_date else "gray"
svc_text = f"預估 {est_service_date}"

v_alert_num = f"逾期 {abs(days_to_v)} 天" if days_to_v < 0 else f"剩 {days_to_v} 天"
v_color = "#ff4b4b" if days_to_v < 0 else ("#ffa421" if days_to_v <= 30 else "#00c04b")

e_alert_num = f"逾期 {abs(days_to_e)} 天" if days_to_e < 0 else f"剩 {days_to_e} 天"
e_color = "#ff4b4b" if days_to_e < 0 else ("#ffa421" if days_to_e <= 30 else "#00c04b")

st.markdown(f"""
<div style="display: flex; justify-content: space-between; align-items: center; background-color: rgba(128,128,128,0.1); padding: 12px 5px; border-radius: 8px; margin-bottom: 15px;">
    <div style="flex: 1; text-align: center; border-right: 1px solid rgba(128,128,128,0.3);">
        <div style="font-size: 12px; opacity: 0.8;">🔧 距保養剩餘</div>
        <div style="font-size: 18px; font-weight: bold; margin: 4px 0;">{rem_service:,} <span style="font-size: 10px; font-weight: normal;">km</span></div>
        <div style="font-size: 12px; color: {svc_color};">{svc_text}</div>
    </div>
    <div style="flex: 1; text-align: center; border-right: 1px solid rgba(128,128,128,0.3);">
        <div style="font-size: 12px; opacity: 0.8;">📋 驗車 {next_v_date.strftime('%m/%d')}</div>
        <div style="font-size: 16px; font-weight: bold; color: {v_color}; margin: 4px 0;">{v_alert_num}</div>
    </div>
    <div style="flex: 1; text-align: center;">
        <div style="font-size: 12px; opacity: 0.8;">💨 排氣 {next_e_date.strftime('%m/%d')}</div>
        <div style="font-size: 16px; font-weight: bold; color: {e_color}; margin: 4px 0;">{e_alert_num}</div>
    </div>
</div>
""", unsafe_allow_html=True)

if "sys_msg" in st.session_state:
    if st.session_state.sys_msg_type == "warning":
        st.warning(st.session_state.sys_msg)
    else:
        st.success(st.session_state.sys_msg)
    del st.session_state.sys_msg
    del st.session_state.sys_msg_type

tab_daily, tab_history, tab_stats, tab_export, tab_settings = st.tabs(["📝 登記", "📋 修改", "📊 預測與防呆", "🖨️ 匯出", "⚙️ 設定"])

with tab_daily:
    st.subheader("每日返廠單次結算")
    with st.form("quick_daily_form"):
        f_date = st.date_input("日期", SHIFT_DATE)
        f_user = st.text_input("使用人", value=DEFAULT_USER)
        f_dest_select = st.selectbox("目的地 (點選快捷)", st.session_state.destinations)
        f_dest_custom = st.text_input("或輸入新目的地 (選填)")

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            f_start_time = st.text_input("出發時間", value="07:20")
            f_start_mil = st.number_input("出發里程 (預設前次，可手動修改)", value=int(latest_return_mileage), step=1)
        with col_t2:
            f_ret_time = st.text_input("返廠時間", value=now_time.strftime("%H:%M"))
            expected_add = 170 if "美光" in f_dest_select else (23 if "順益" in f_dest_select else (5 if "廠內調度" in f_dest_select else 50))
            def_return = int(f_start_mil) + expected_add
            f_ret_mil = st.number_input("返廠里程 (請看儀表板輸入)", value=int(def_return), step=1)

        # 司機下班一鍵同時結算加油發票快捷勾選
        st.markdown("---")
        chk_today_fuel = st.checkbox("⛽ 今天有順便加油？（勾選後直接填寫發票，一鍵同步）", value=False)
        col_df1, col_df2, col_df3 = st.columns(3)
        with col_df1:
            df_liters = st.number_input("加油公升數 (L)", value=150.0, step=0.1) if chk_today_fuel else 0.0
        with col_df2:
            df_cost = st.number_input("實付加油費用 (元)", value=4300, step=10) if chk_today_fuel else 0
        with col_df3:
            df_station = st.selectbox("加油站", ["中油苗栗站", "中油新竹站", "其他"], key="df_station") if chk_today_fuel else ""

        submit_daily = st.form_submit_button("💾 儲存今日紀錄，同步至雲端", use_container_width=True)
        
        if submit_daily:
            dest_final = f_dest_custom.strip() if f_dest_custom.strip() else f_dest_select
            diff_km = int(f_ret_mil) - int(f_start_mil)
            
            if diff_km < 0:
                st.error("❌ 【輸入錯誤】返廠里程不可小於出發里程！請重新檢查儀表板數字。")
            elif diff_km > 1000:
                st.error(f"❌ 【異常防呆】本次增加里程達 {diff_km} km (超過單日 1,000 km 上限)！疑似多輸入一位數字，已阻擋寫入。")
            else:
                expected_km = 170 if "美光" in dest_final else (23 if "順益" in dest_final else (5 if "廠內調度" in dest_final else 50))
                
                daily_record_dict = {
                    "日期": f_date.strftime("%Y-%m-%d"),
                    "使用人": f_user,
                    "目的地": dest_final,
                    "出發時間": f_start_time,
                    "出發里程": int(f_start_mil),
                    "返廠時間": f_ret_time,
                    "返廠里程": int(f_ret_mil),
                    "備註": ""
                }
                default_d_order = ["日期", "使用人", "目的地", "出發時間", "出發里程", "返廠時間", "返廠里程", "備註"]
                safe_append_row(ACTIVE_SHEET_NAME, "Daily_Logs", daily_record_dict, default_d_order, daily_map)
                
                if chk_today_fuel and float(df_liters) > 0:
                    fuel_dict = {
                        "加油日期": f_date.strftime("%Y-%m-%d"),
                        "加油時間": f_ret_time,
                        "加油里程": int(f_ret_mil),
                        "公升數": float(df_liters),
                        "加油費用": int(df_cost),
                        "加油站名": df_station
                    }
                    default_f_order = ["加油日期", "加油時間", "加油里程", "公升數", "加油費用", "加油站名"]
                    safe_append_row(ACTIVE_SHEET_NAME, "Fuel_Logs", fuel_dict, default_f_order, fuel_map)
                
                if diff_km > (expected_km + 80):
                    st.session_state.sys_msg = f"⚠️ 【里程較長提醒】本次增加了 {diff_km} km。紀錄已成功存入，請確認是否有繞道或漏填！"
                    st.session_state.sys_msg_type = "warning"
                else:
                    fuel_tag = "（含加油發票）" if (chk_today_fuel and float(df_liters) > 0) else ""
                    st.session_state.sys_msg = f"✅ {f_date.strftime('%m/%d')} 的行車紀錄{fuel_tag}已寫入 Google Sheets！"
                    st.session_state.sys_msg_type = "success"
                    
                fetch_all_data.clear() 
                st.rerun()

with tab_history:
    st.subheader("📱 手機專用歷史修改")
    c_btn1, c_btn2 = st.columns([0.7, 0.3])
    mode_choices = ["日常行車", "柴油紀錄"]
    if HAS_UREA:
        mode_choices.append("尿素紀錄")
        
    with c_btn1:
        edit_mode = st.radio("請選擇要修改的紀錄類型：", mode_choices, horizontal=True, label_visibility="collapsed")
    with c_btn2: 
        if st.button("🔄 強制同步"): 
            fetch_all_data.clear()
            st.rerun()
            
    if edit_mode == "日常行車":
        if not st.session_state.daily_logs:
            st.info("目前沒有資料可修改")
        else:
            all_months = sorted(list(set([str(log.get('日期', ''))[:7] for log in st.session_state.daily_logs if log.get('日期')])), reverse=True)
            selected_month = st.selectbox("📅 第一步：選擇月份", all_months, key="hist_m1")
            log_options = []
            for idx, log in enumerate(st.session_state.daily_logs):
                if str(log.get('日期', '')).startswith(selected_month):
                    unique_label = f"#{idx+2} [{log.get('日期')}] {log.get('目的地')} (出發:{log.get('出發里程')} ➔ 返廠:{log.get('返廠里程')})"
                    log_options.append((idx + 2, unique_label, log))
            log_options = sorted(log_options, key=lambda x: str(x[2].get('日期', '')), reverse=True)
            
            if not log_options:
                st.warning("該月份無紀錄")
            else:
                selected_label = st.selectbox("✏️ 第二步：選擇紀錄", [item[1] for item in log_options])
                target_row_num, list_idx, target_log = next((item[0], item[0]-2, item[2]) for item in log_options if item[1] == selected_label)
                
                with st.form("mobile_edit_form"):
                    e_date = st.date_input("日期", parse_date_safe(target_log.get("日期"), TODAY))
                    e_user = st.text_input("使用人", value=str(target_log.get("使用人", "")))
                    e_dest = st.text_input("目的地", value=str(target_log.get("目的地", "")))
                    c_e1, c_e2 = st.columns(2)
                    with c_e1:
                        e_st_time = st.text_input("出發時間", value=str(target_log.get("出發時間", "")))
                        e_st_mil = st.number_input("出發里程", value=safe_int(target_log.get("出發里程")), step=1)
                    with c_e2:
                        e_ret_time = st.text_input("返廠時間", value=str(target_log.get("返廠時間", "")))
                        e_ret_mil = st.number_input("返廠里程", value=safe_int(target_log.get("返廠里程")), step=1)
                    
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        btn_save_edit = st.form_submit_button("💾 儲存修改", use_container_width=True)
                    with c_btn2:
                        btn_delete = st.form_submit_button("🗑️ 刪除紀錄", use_container_width=True)
                    
                    if btn_save_edit:
                        updated_row = [e_date.strftime("%Y-%m-%d"), e_user, e_dest, e_st_time, int(e_st_mil), e_ret_time, int(e_ret_mil), str(target_log.get("備註", ""))]
                        get_worksheet(ACTIVE_SHEET_NAME, "Daily_Logs").update(range_name=f"A{target_row_num}:H{target_row_num}", values=[updated_row])
                        st.session_state.sys_msg = "✅ 日常紀錄已更新！"
                        st.session_state.sys_msg_type = "success"
                        fetch_all_data.clear()
                        st.rerun()
                    if btn_delete:
                        get_worksheet(ACTIVE_SHEET_NAME, "Daily_Logs").delete_rows(target_row_num)
                        st.session_state.sys_msg = "🗑️ 該筆紀錄已刪除！"
                        st.session_state.sys_msg_type = "warning"
                        fetch_all_data.clear()
                        st.rerun()

    elif edit_mode == "柴油紀錄":
        if not st.session_state.fuel_logs:
            st.info("目前沒有柴油資料")
        else:
            all_months = sorted(list(set([str(log.get('加油日期', ''))[:7] for log in st.session_state.fuel_logs if log.get('加油日期')])), reverse=True)
            selected_month = st.selectbox("📅 第一步：選擇月份", all_months, key="hist_m2")
            
            log_options = []
            for idx, log in enumerate(st.session_state.fuel_logs):
                if str(log.get('加油日期', '')).startswith(selected_month):
                    unique_label = f"#{idx+2} [{log.get('加油日期')}] {log.get('加油站名')} - {log.get('公升數')}L (里程: {log.get('加油里程')})"
                    log_options.append((idx + 2, unique_label, log))
            log_options = sorted(log_options, key=lambda x: str(x[2].get('加油日期', '')), reverse=True)
            
            if not log_options:
                st.warning("該月份無紀錄")
            else:
                selected_label = st.selectbox("✏️ 第二步：選擇要修改的柴油紀錄", [item[1] for item in log_options])
                target_row_num, list_idx, target_log = next((item[0], item[0]-2, item[2]) for item in log_options if item[1] == selected_label)
                
                with st.form("edit_fuel_form"):
                    e_fd = st.date_input("加油日期", parse_date_safe(target_log.get("加油日期"), TODAY))
                    e_fm = st.number_input("加油里程", value=safe_int(target_log.get("加油里程")), step=1)
                    e_fl = st.number_input("公升數 (L)", value=safe_float(target_log.get("公升數")), step=0.1)
                    e_fc = st.number_input("費用 (元)", value=safe_int(target_log.get("加油費用")), step=1)
                    
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        btn_save = st.form_submit_button("💾 儲存修改", use_container_width=True)
                    with c_btn2:
                        btn_del = st.form_submit_button("🗑️ 刪除", use_container_width=True)
                    
                    if btn_save:
                        upd_row = [e_fd.strftime("%Y-%m-%d"), target_log.get("加油時間", ""), int(e_fm), float(e_fl), int(e_fc), target_log.get("加油站名", "")]
                        get_worksheet(ACTIVE_SHEET_NAME, "Fuel_Logs").update(range_name=f"A{target_row_num}:F{target_row_num}", values=[upd_row])
                        st.session_state.sys_msg = "✅ 柴油紀錄已更新！"
                        st.session_state.sys_msg_type = "success"
                        fetch_all_data.clear()
                        st.rerun()
                    if btn_del:
                        get_worksheet(ACTIVE_SHEET_NAME, "Fuel_Logs").delete_rows(target_row_num)
                        st.session_state.sys_msg = "🗑️ 柴油紀錄已刪除！"
                        st.session_state.sys_msg_type = "warning"
                        fetch_all_data.clear()
                        st.rerun()

    elif edit_mode == "尿素紀錄" and HAS_UREA:
        if not st.session_state.urea_logs:
            st.info("目前沒有尿素資料")
        else:
            def get_urea_date(log):
                return log.get('添加日期') or log.get('日期') or log.get('紀錄日期') or '未填'

            valid_dates = [get_urea_date(log) for log in st.session_state.urea_logs if get_urea_date(log) != '未填']
            all_months = sorted(list(set([d[:7] for d in valid_dates])), reverse=True)
            if not all_months:
                all_months = ["未填"]
            
            selected_month = st.selectbox("📅 第一步：選擇月份", all_months, key="hist_m3")
            
            log_options = []
            for idx, log in enumerate(st.session_state.urea_logs):
                u_date = get_urea_date(log)
                if u_date.startswith(selected_month) or (selected_month == "未填" and u_date == "未填"):
                    unique_label = f"#{idx+2} [{u_date}] {log.get('公升數', 0)}L (里程: {log.get('添加當下里程', '未填')})"
                    log_options.append((idx + 2, unique_label, log))
            log_options = sorted(log_options, key=lambda x: get_urea_date(x[2]), reverse=True)
            
            if not log_options:
                st.warning("該月份無紀錄")
            else:
                selected_label = st.selectbox("✏️ 第二步：選擇要修改的尿素紀錄", [item[1] for item in log_options])
                target_row_num, list_idx, target_log = next((item[0], item[0]-2, item[2]) for item in log_options if item[1] == selected_label)
                
                with st.form("edit_urea_form"):
                    e_ud = st.date_input("添加日期", parse_date_safe(get_urea_date(target_log), TODAY))
                    e_um = st.number_input("添加當下里程", value=safe_int(target_log.get("添加當下里程")), step=1)
                    e_ul = st.number_input("公升數 (L)", value=safe_float(target_log.get("公升數")), step=1.0)
                    e_uc = st.number_input("費用 (元)", value=safe_int(target_log.get("費用")), step=1)
                    
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        btn_save = st.form_submit_button("💾 儲存修改", use_container_width=True)
                    with c_btn2:
                        btn_del = st.form_submit_button("🗑️ 刪除", use_container_width=True)
                    
                    if btn_save:
                        upd_row = [e_ud.strftime("%Y-%m-%d"), target_log.get("添加時間", ""), float(e_ul), int(e_uc), target_log.get("備註", ""), int(e_um)]
                        get_worksheet(ACTIVE_SHEET_NAME, "Urea_Logs").update(range_name=f"A{target_row_num}:F{target_row_num}", values=[upd_row])
                        st.session_state.sys_msg = "✅ 尿素紀錄已更新！"
                        st.session_state.sys_msg_type = "success"
                        fetch_all_data.clear()
                        st.rerun()
                    if btn_del:
                        get_worksheet(ACTIVE_SHEET_NAME, "Urea_Logs").delete_rows(target_row_num)
                        st.session_state.sys_msg = "🗑️ 尿素紀錄已刪除！"
                        st.session_state.sys_msg_type = "warning"
                        fetch_all_data.clear()
                        st.rerun()

with tab_stats:
    default_station_mil = latest_return_mileage
    
    st.subheader("⚠️ 硬體異常防呆監控")
    if historical_km_l > 0.5 and len(sorted_fuel_logs) >= 3:
        drop_ratio = (historical_km_l - global_fuel_km_l) / historical_km_l
        if drop_ratio >= 0.15:
            st.error(f"🔴 **警告！油耗異常崩跌**\n\n歷史平均油耗為 **{historical_km_l} km/L**，但最新一期驟降至 **{global_fuel_km_l} km/L** (跌幅 {drop_ratio*100:.1f}%)。\n\n**請盡速檢查：** 胎壓過低、煞車咬死、DPF 濾清器阻塞，或管路漏油。")
        elif drop_ratio <= -0.15:
            st.success(f"🟢 **油耗表現優異**：最新油耗 {global_fuel_km_l} km/L 優於歷史平均 {historical_km_l} km/L！")
        else:
            st.info(f"✅ **車況健康**：最新油耗 {global_fuel_km_l} km/L 與歷史平均一致，無異常阻力。")
    else:
        st.caption("累積滿 3 次加油紀錄後，將自動啟動硬體異常防呆雷達。")

    st.divider()
    st.subheader("📊 智慧進站預測 (依車上里程表為準)")
    
    safe_fuel_vol = 60 / global_fuel_km_l
    usable_fuel_vol = max(0.0, vol_after_today - safe_fuel_vol)
    fuel_runnable_km = usable_fuel_vol * global_fuel_km_l
    fuel_days_left = fuel_runnable_km / avg_daily_km
    target_fuel_mil = latest_return_mileage + math.floor(fuel_runnable_km)
    
    # 【核心架構修復：以最新里程的紀錄日期為基準，計算未來出車到達進站點之精準日期】
    # 當今日 (11/05) 已經返廠結算：剩餘油量是提供給「未來出車」使用
    # 第 1 趟出車為 11/06，第 2 趟為 11/07 ...
    # trips_until_fuel = math.ceil(fuel_days_left)
    trips_until_fuel = max(1, math.ceil(fuel_days_left)) if usable_fuel_vol > 0 else 1
    target_fuel_date = latest_mil_date + timedelta(days=trips_until_fuel)
    
    if HAS_UREA and urea_km_l > 10.0 and global_fuel_km_l > 0.5:
        usable_urea_vol = max(0.0, rem_urea_vol - 8.0)
        urea_runnable_km = usable_urea_vol * urea_km_l
        urea_days_left = urea_runnable_km / avg_daily_km
        target_urea_mil = latest_return_mileage + math.floor(urea_runnable_km)
        
        trips_until_urea = max(1, math.ceil(urea_days_left)) if usable_urea_vol > 0 else 1
        target_urea_date = latest_mil_date + timedelta(days=trips_until_urea)
        
        if urea_runnable_km >= fuel_runnable_km:
            default_station_mil = target_fuel_mil
            future_urea_vol = max(0.0, rem_urea_vol - (fuel_runnable_km / urea_km_l))
            
            st.success(f"✅ **預估進站點 (柴油主導)：【 里程 {target_fuel_mil:,} km 】**\n\n*(最晚請於 **{target_fuel_date.strftime('%m月%d日')}** 進站加油，請以車上里程表為準)*")
            
            if future_urea_vol >= 8.0:
                st.info(f"💡 **決策：【這次只要加柴油就好】**\n\n開到 **{target_fuel_mil:,} km** 時，屆時尿素將從目前 {rem_urea_vol:.1f}L 消耗至約 **{future_urea_vol:.1f} 公升** (安全)，單純補滿柴油即可！")
            else:
                st.warning(f"🟡 **決策：【柴油 ＋ 尿素 順便一起加】**\n\n開到 **{target_fuel_mil:,} km** 時，屆時尿素將消耗至剩 **{future_urea_vol:.1f} 公升** (逼近防護線)，建議兩管齊下一次加滿！")
        else:
            default_station_mil = target_urea_mil
            future_fuel_vol = max(0.0, vol_after_today - (urea_runnable_km / global_fuel_km_l))
            
            st.error(f"🚨 **尿素提早告急：【 里程 {target_urea_mil:,} km 】**\n\n*(最晚請於 **{target_urea_date.strftime('%m月%d日')}** 進站補充尿素，請以車上里程表為準)*")
            st.warning(f"💡 **決策：【以補尿素為主，柴油順便加滿】**\n\n開到 **{target_urea_mil:,} km** 時尿素將跌破 8L 警戒！屆時柴油約剩 **{future_fuel_vol:.1f} 公升**，請一併補滿。")
    else:
        default_station_mil = target_fuel_mil
        st.success(f"✅ **預估進站點：【 里程 {target_fuel_mil:,} km 】**\n\n*(最晚請於 **{target_fuel_date.strftime('%m月%d日')}** 進站，剩餘可用油量約 {usable_fuel_vol:.1f} L)*")

    st.divider()
    st.markdown(f"**目前車輛剩餘狀況** (以最新里程 {latest_return_mileage:,} km 結算)")
    
    c_dash1, c_dash2 = st.columns(2)
    with c_dash1:
        st.metric(label="距上次加油已跑", value=f"{used_km_since_fuel} km", delta=f"柴油即時剩餘 {vol_after_today:.1f} L (油箱 {int(tank_cap)}L)", delta_color="off")
    with c_dash2:
        st.metric("最新一期油耗", f"{global_fuel_km_l} km/L")
        
    rem_safe_km = int(vol_after_today * global_fuel_km_l)
    if rem_safe_km < 60:
        st.error(f"🔴 柴油極危險！含油底續航僅剩 {rem_safe_km} km，請立刻補油！")
    elif 60 <= rem_safe_km < 100:
        st.warning(f"🟡 柴油警戒！續航剩 {rem_safe_km} km，上高速前請確認無塞車。")
    
    if HAS_UREA:
        st.divider()
        c_u1, c_u2 = st.columns(2)
        with c_u1:
            st.metric(label="距上次加尿已跑", value=f"{driven_since_urea} km", delta=f"尿素即時剩餘 {rem_urea_vol:.1f} L", delta_color="off")
        with c_u2:
            st.metric("歷史尿素平均耗能", f"{urea_km_l} km/L")
                
        if 4.0 <= rem_urea_vol < 8.0:
            st.warning(f"⚠️ 尿素警戒：即時剩餘 {rem_urea_vol:.1f} L (逼近 8L 安全防線)")
        elif rem_urea_vol < 4.0:
            st.error(f"🔴 尿素危險：即時剩餘 {rem_urea_vol:.1f} L！避免抽不到底鎖死，請盡速補加！")

    st.divider()
    st.markdown("### ⛽💧 補給站登記中心")
    if default_station_mil > latest_return_mileage:
        st.caption(f"🤖 **神預測同步**：已自動預載預估進站里程，您只需微調最後幾碼即可！")
    else:
        st.caption("💡 輸入當下儀表板里程，勾選登記柴油或尿素，一鍵上傳！")
    
    with st.form("station_master_form"):
        shared_station_mil = st.number_input("補給當下儀表板總里程 (km)", value=int(default_station_mil), step=1)
        station_date = st.date_input("補給日期", SHIFT_DATE)
        
        st.markdown("---")
        chk_fuel = st.checkbox("☑️ 登記柴油", value=True)
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fuel_l = st.number_input("柴油公升數 (L)", value=150.0, step=0.1)
        with col_f2:
            fuel_cost = st.number_input("柴油實付費用 (元)", value=4300, step=10)
        with col_f3:
            fuel_station = st.selectbox("加油站", ["中油苗栗站", "中油新竹站", "其他"])
            
        chk_urea = False
        if HAS_UREA:
            st.markdown("---")
            chk_urea = st.checkbox("☑️ 登記尿素", value=False)
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                urea_l = st.number_input("尿素公升數 (L)", value=20.0, step=1.0)
            with col_u2:
                urea_cost = st.number_input("尿素費用 (元)", value=500, step=10)
            
        submit_station = st.form_submit_button("🚀 一鍵同步儲存 (只會上傳有打勾的項目)", use_container_width=True)
        
        if submit_station:
            min_allowed_mil = max(1000, latest_return_mileage - 500)
            if not chk_fuel and not chk_urea:
                st.warning("⚠️ 請至少勾選「柴油」或「尿素」其中一項！")
            elif int(shared_station_mil) < min_allowed_mil:
                st.error(f"❌ 補給里程過低（不可小於 {min_allowed_mil} km）！請確認是否看錯儀表板數字。")
            else:
                msg_list = []
                cur_hhmm = now_time.strftime("%H:%M")
                
                if chk_fuel:
                    fuel_dict = {
                        "加油日期": station_date.strftime("%Y-%m-%d"),
                        "加油時間": cur_hhmm,
                        "加油里程": int(shared_station_mil),
                        "公升數": float(fuel_l),
                        "加油費用": int(fuel_cost),
                        "加油站名": fuel_station
                    }
                    default_f_order = ["加油日期", "加油時間", "加油里程", "公升數", "加油費用", "加油站名"]
                    safe_append_row(ACTIVE_SHEET_NAME, "Fuel_Logs", fuel_dict, default_f_order, fuel_map)
                    msg_list.append("柴油")
                    
                if chk_urea and HAS_UREA:
                    urea_dict = {
                        "添加日期": station_date.strftime("%Y-%m-%d"),
                        "添加時間": cur_hhmm,
                        "公升數": float(urea_l),
                        "費用": int(urea_cost),
                        "備註": "",
                        "添加當下里程": int(shared_station_mil)
                    }
                    default_u_order = ["添加日期", "添加時間", "公升數", "費用", "備註", "添加當下里程"]
                    safe_append_row(ACTIVE_SHEET_NAME, "Urea_Logs", urea_dict, default_u_order, urea_map)
                    msg_list.append("尿素")
                    
                st.session_state.sys_msg = f"✅ 【{' ＋ '.join(msg_list)}】補給紀錄已成功寫入 Google Sheets！"
                st.session_state.sys_msg_type = "success"
                fetch_all_data.clear()
                st.rerun()

with tab_export:
    st.subheader("標準紙本 1:1 還原列印")
    exp_y = st.selectbox("年份", [2026, 2027], index=0)
    exp_m = st.selectbox("月份", list(range(1, 13)), index=datetime.now(TW_TZ).month - 1)
    excel_file = generate_kertz_excel(exp_y, exp_m)
    st.download_button(
        label=f"📥 下載 {exp_y} 年 {exp_m} 月行車紀錄表 (.xlsx)",
        data=excel_file,
        file_name=f"行車紀錄表_{DEFAULT_PLATE}_{exp_y}年{exp_m}月.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with tab_settings:
    st.subheader("⚙️ 系統與法規參數設定")
    with st.form("settings_form"):
        new_target = st.number_input("下次保養里程目標 (km)", value=st.session_state.service_target, step=1000)
        st.markdown("---")
        st.markdown("**法規檢驗排程 (請依照行照指定日期填寫)**")
        new_v_date = st.date_input("下次驗車到期日", next_v_date)
        new_e_date = st.date_input("下次驗排氣到期日", next_e_date)
        st.markdown("---")
        st.markdown(f"**車輛總油箱容量**：目前由系統動態鎖定為 **{int(tank_cap)} L** (根據歷史加油紀錄自動校正)")
        
        submit_settings = st.form_submit_button("💾 儲存所有設定", use_container_width=True)
        if submit_settings:
            update_vals = [
                ["SERVICE_TARGET", new_target],
                ["FUEL_TANK_CAPACITY", float(tank_cap)],
                ["NEXT_VEHICLE_INSP", new_v_date.strftime("%Y-%m-%d")],
                ["NEXT_EMISSIONS_INSP", new_e_date.strftime("%Y-%m-%d")]
            ]
            get_worksheet(ACTIVE_SHEET_NAME, "Settings").update(range_name="A1:B4", values=update_vals)
            st.session_state.sys_msg = "✅ 所有設定已更新並存入雲端！"
            st.session_state.sys_msg_type = "success"
            fetch_all_data.clear()
            st.rerun()
