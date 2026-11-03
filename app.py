import json
import math
from datetime import datetime, timedelta, timezone
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
import pandas as pd
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ================= 系統安全轉型防呆函數 =================
def safe_int(val, default=0):
    try:
        if val is None or val == "":
            return default
        return int(float(val))
    except:
        return default

def safe_float(val, default=0.0):
    try:
        if val is None or val == "":
            return default
        return float(val)
    except:
        return default

def parse_date_safe(date_str, default_date):
    try:
        return datetime.strptime(str(date_str).strip(), "%Y-%m-%d").date()
    except:
        return default_date

# ================= 系統固定參數與時區設定 =================
DEFAULT_PLATE = "380-S2"
DEFAULT_KEEPER = "劉嘉舜"
DEFAULT_USER = "劉嘉舜"
DOC_CODE = "4-33-DC0007B"
FORM_NAME = "公務車輛行車紀錄表 Rev.1"
UREA_TANK_CAPACITY = 20.0  
TW_TZ = timezone(timedelta(hours=8))

now_time = datetime.now(TW_TZ)
TODAY = now_time.date()

# 🚚 跨夜防呆邏輯：凌晨 00:00 ~ 04:00 打開系統，預設班別日期為「昨天」
if now_time.hour < 4:
    SHIFT_DATE = (now_time - timedelta(days=1)).date()
else:
    SHIFT_DATE = TODAY

# ================= Google Sheets 雲端連線設定 =================
@st.cache_resource
def get_gsheets_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if "google_key" in st.secrets:
        creds_dict = json.loads(st.secrets["google_key"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    return gspread.authorize(creds)

client = get_gsheets_client()
sh = client.open("380-S2_公務車行車紀錄系統")
ws_daily = sh.worksheet("Daily_Logs")
ws_fuel = sh.worksheet("Fuel_Logs")
ws_urea = sh.worksheet("Urea_Logs")

# ================= 設定檔初始化與讀取 =================
sheet_titles = [w.title for w in sh.worksheets()]
if "Settings" not in sheet_titles:
    ws_settings = sh.add_worksheet(title="Settings", rows=10, cols=2)
    default_vals = [
        ["SERVICE_TARGET", 622000],
        ["FUEL_TANK_CAPACITY", 130.0],
        ["NEXT_VEHICLE_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(months=6)).strftime("%Y-%m-%d")],
        ["NEXT_EMISSIONS_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(years=1)).strftime("%Y-%m-%d")]
    ]
    ws_settings.update(values=default_vals, range_name="A1:B4")
else:
    ws_settings = sh.worksheet("Settings")

try:
    settings_vals = ws_settings.get_all_values()
    settings_dict = {row[0]: row[1] for row in settings_vals if len(row) >= 2}
except:
    settings_dict = {}

st.session_state.service_target = safe_int(settings_dict.get("SERVICE_TARGET"), 622000)
base_tank_capacity = safe_float(settings_dict.get("FUEL_TANK_CAPACITY"), 130.0)

next_v_insp = settings_dict.get("NEXT_VEHICLE_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(months=6)).strftime("%Y-%m-%d"))
next_e_insp = settings_dict.get("NEXT_EMISSIONS_INSP", (pd.to_datetime(TODAY) + pd.DateOffset(years=1)).strftime("%Y-%m-%d"))
next_v_date = parse_date_safe(next_v_insp, TODAY)
next_e_date = parse_date_safe(next_e_insp, TODAY)
days_to_v = (next_v_date - TODAY).days
days_to_e = (next_e_date - TODAY).days

# ================= 資料同步與快取 =================
def load_data():
    try:
        daily_data = ws_daily.get_all_records()
        fuel_data = ws_fuel.get_all_records()
        urea_data = ws_urea.get_all_records()
    except Exception as e:
        st.error(f"雲端資料讀取失敗: {e}")
        return

    if not daily_data:
        daily_data = [{"日期": "2026-10-31", "使用人": "劉嘉舜", "目的地": "美光", "出發時間": "07:20", "出發里程": 615049, "返廠時間": "18:30", "返廠里程": 615215, "備註": "系統基準"}]
    if not fuel_data:
        fuel_data = [{"加油日期": "2026-10-30", "加油時間": "18:40", "加油里程": 614997, "公升數": 130.0, "加油費用": 3816, "加油站名": "中油苗栗站"}]
    if not urea_data:
        urea_data = [{"添加日期": "2026-10-30", "添加時間": "18:45", "公升數": 20.0, "費用": 500, "備註": "初始添加", "添加當下里程": 614997}]
        
    st.session_state.daily_logs = daily_data
    st.session_state.fuel_logs = fuel_data
    st.session_state.urea_logs = urea_data
    if "destinations" not in st.session_state:
        st.session_state.destinations = ["美光", "順益", "廠內調度"]

load_data()

# 跨表全域最新真實里程
all_recorded_mil = [safe_int(x.get("返廠里程")) for x in st.session_state.daily_logs] + \
                     [safe_int(x.get("加油里程")) for x in st.session_state.fuel_logs] + \
                     [safe_int(x.get("添加當下里程")) for x in st.session_state.urea_logs]
latest_return_mileage = max(all_recorded_mil) if all_recorded_mil else 615215
rem_service = st.session_state.service_target - latest_return_mileage

# ================= 柴油消耗運算 =================
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
global_fuel_km_l = max(0.1, global_fuel_km_l)

historical_km_l = global_fuel_km_l
if len(sorted_fuel_logs) >= 3:
    first_mil = safe_int(sorted_fuel_logs[0].get("加油里程"))
    prev_mil = safe_int(sorted_fuel_logs[-2].get("加油里程"))
    total_km_hist = prev_mil - first_mil
    total_liters_hist = sum(safe_float(x.get("公升數")) for x in sorted_fuel_logs[1:-1])
    if total_liters_hist > 0 and total_km_hist > 0:
        historical_km_l = round(total_km_hist / total_liters_hist, 2)

# ================= 日均里程演算法 =================
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
                avg = mileage_diff / days_diff
                return max(1.0, avg)
        except:
            pass
    return 170.0 

avg_daily_km = get_avg_daily_km(st.session_state.daily_logs)

max_liters_added = max([safe_float(f.get("公升數")) for f in sorted_fuel_logs]) if sorted_fuel_logs else 130.0
learned_tank_cap = float(math.ceil(max_liters_added / 10.0) * 10)
tank_cap = max(base_tank_capacity, learned_tank_cap)

curr_vol = tank_cap  
last_f_mil = safe_int(sorted_fuel_logs[0].get("加油里程")) if sorted_fuel_logs else latest_return_mileage
for f_log in sorted_fuel_logs:
    f_mil = safe_int(f_log.get("加油里程"))
    add_l = safe_float(f_log.get("公升數"))
    dist = f_mil - last_f_mil
    if dist > 0:
        curr_vol -= (dist / global_fuel_km_l)
    if curr_vol < 0:
        curr_vol = 0
    curr_vol += add_l
    if curr_vol > tank_cap:
        curr_vol = tank_cap
    last_f_mil = f_mil

used_km_since_fuel = max(0, latest_return_mileage - latest_fuel_mileage)
used_vol_since_fuel = used_km_since_fuel / global_fuel_km_l
vol_after_today = max(0.0, curr_vol - used_vol_since_fuel)

# ================= 尿素消耗運算 =================
sorted_urea = sorted(st.session_state.urea_logs, key=lambda x: str(x.get("添加日期", "")))
urea_with_mil = [u for u in sorted_urea if "添加當下里程" in u and str(u["添加當下里程"]).isdigit()]
urea_with_mil = sorted(urea_with_mil, key=lambda x: safe_int(x["添加當下里程"]))

urea_km_l = 170.0 
if len(urea_with_mil) >= 2:
    first_u_mil = safe_int(urea_with_mil[0]["添加當下里程"])
    last_u_mil = safe_int(urea_with_mil[-1]["添加當下里程"])
    total_u_km = last_u_mil - first_u_mil
    total_u_liters = sum(safe_float(x.get("公升數")) for x in urea_with_mil[1:])
    if total_u_liters > 0 and total_u_km > 0:
        urea_km_l = round(total_u_km / total_u_liters, 2)
urea_km_l = max(1.0, urea_km_l)

curr_urea_vol = UREA_TANK_CAPACITY  
last_u_mil = safe_int(urea_with_mil[0]["添加當下里程"]) if urea_with_mil else latest_return_mileage

for u_log in urea_with_mil:
    u_mil = safe_int(u_log["添加當下里程"])
    add_l = safe_float(u_log.get("公升數"))
    dist = u_mil - last_u_mil
    if dist > 0:
        curr_urea_vol -= (dist / urea_km_l)
    if curr_urea_vol < 0:
        curr_urea_vol = 0
    curr_urea_vol += add_l
    if curr_urea_vol > UREA_TANK_CAPACITY:
        curr_urea_vol = UREA_TANK_CAPACITY
    last_u_mil = u_mil

driven_since_urea = max(0, latest_return_mileage - last_u_mil)
used_urea_since_last = driven_since_urea / urea_km_l
rem_urea_vol = max(0.0, curr_urea_vol - used_urea_since_last)

def calculate_service_prediction(target_mileage):
    km_remaining = target_mileage - latest_return_mileage
    if km_remaining <= 0:
        return "已達保養里程"
    days_remaining = int(km_remaining / avg_daily_km)
    est_date = datetime.now(TW_TZ) + timedelta(days=days_remaining)
    return est_date.strftime("%m/%d")

est_service_date = calculate_service_prediction(st.session_state.service_target)

# ================= 1:1 原樣 A4 報表生成 =================
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
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    ws.merge_cells("A1:G1")
    rich_text = CellRichText(
        TextBlock(InlineFont(b=False, sz=12, rFont=font_family), 'KT.T. '), 
        TextBlock(InlineFont(i=True, sz=9, rFont=font_family), 'Pro Clean Tech   '), 
        TextBlock(InlineFont(b=False, sz=12, rFont=font_family), '科治新技股份有限公司 標準作業表格')
    )
    ws["A1"] = rich_text
    ws["A1"].alignment = align_center
    ws.merge_cells("H1:I1")
    ws["H1"] = "頁次："
    ws["H1"].font = font_header_sub
    ws["H1"].alignment = align_left
    
    ws.merge_cells("A2:B2")
    ws["A2"] = "表格名稱"
    ws["A2"].font = Font(name=font_family, size=12)
    ws["A2"].alignment = align_center
    ws.merge_cells("C2:G2")
    ws["C2"] = "公務車輛行車紀錄表 Rev.1"
    ws["C2"].font = Font(name=font_family, size=12)
    ws["C2"].alignment = align_center
    ws.merge_cells("H2:I2")
    ws["H2"] = f"編號：{DOC_CODE}"
    ws["H2"].font = Font(name=font_family, size=10)
    ws["H2"].alignment = align_left

    roc_year = year - 1911
    ws.merge_cells("A3:B3")
    ws["A3"] = "行車日誌"
    ws["A3"].font = font_bold_md
    ws["A3"].alignment = align_center
    
    ws.merge_cells("C3:I3")
    ws["C3"] = f"月份： {roc_year} 年 {month} 月          車牌號碼： {DEFAULT_PLATE}          保管人： {DEFAULT_KEEPER}"
    ws["C3"].font = font_bold_md
    ws["C3"].alignment = align_center

    headers = ["日期", "使用人", "目的地", "出發時間", "出發里程數", "返廠時間", "返廠里程數", "加油里程數", "加油費用"]
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=text)
        cell.font = font_bold_md
        cell.alignment = align_center

    for r in range(1, 5):
        ws.row_dimensions[r].height = 28
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = border_all

    prefix = f"{year:04d}-{month:02d}"
    m_logs = [l for l in st.session_state.daily_logs if str(l.get("日期", "")).startswith(prefix)]
    m_logs = sorted(m_logs, key=lambda x: str(x.get("日期", "")))
    f_map = {str(fl["加油日期"]): fl for fl in st.session_state.fuel_logs if str(fl.get("加油日期", "")).startswith(prefix)}

    start_row = 5
    for i in range(31):
        curr_row = start_row + i
        ws.row_dimensions[curr_row].height = 24.5
        
        if i < len(m_logs):
            log = m_logs[i]
            try:
                day_num = int(log["日期"].split("-")[2])
            except:
                day_num = ""
                
            ws.cell(row=curr_row, column=1, value=str(day_num)).alignment = align_center
            ws.cell(row=curr_row, column=2, value=str(log.get("使用人", ""))).alignment = align_center
            ws.cell(row=curr_row, column=3, value=str(log.get("目的地", ""))).alignment = align_center
            ws.cell(row=curr_row, column=4, value=str(log.get("出發時間", ""))).alignment = align_center
            ws.cell(row=curr_row, column=5, value=safe_int(log.get("出發里程"))).alignment = align_right
            ws.cell(row=curr_row, column=6, value=str(log.get("返廠時間", ""))).alignment = align_center
            ws.cell(row=curr_row, column=7, value=safe_int(log.get("返廠里程"))).alignment = align_right
            
            if log["日期"] in f_map:
                fl = f_map[log["日期"]]
                ws.cell(row=curr_row, column=8, value=safe_int(fl.get("加油里程"))).alignment = align_right
                ws.cell(row=curr_row, column=9, value=safe_int(fl.get("加油費用"))).alignment = align_right
        
        for c in range(1, 10):
            cell = ws.cell(row=curr_row, column=c)
            if not cell.font:
                cell.font = font_regular
            cell.border = border_all
            if c in [5, 7, 8, 9] and isinstance(cell.value, (int, float, str)) and str(cell.value).isdigit():
                cell.number_format = "#,##0"

    col_widths = {"A": 4.5, "B": 7.5, "C": 14.0, "D": 9.5, "E": 12.5, "F": 9.5, "G": 12.5, "H": 12.5, "I": 10.5}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ================= 前端介面 =================
st.set_page_config(page_title="380-S2 車隊管理", layout="centered")

st.markdown("""
<style>
    .main-title { font-size: 28px !important; font-weight: 800; margin-bottom: 5px; white-space: nowrap; }
</style>
<div class='main-title'>380-S2 車隊管理</div>
""", unsafe_allow_html=True)
st.caption("連線狀態：🟢 Google Sheets 雲端連線成功")

# ================= 綜合提醒儀表板 =================
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
            f_start_mil = st.number_input("出發里程 (鎖定前次)", value=latest_return_mileage, disabled=True)
        with col_t2:
            f_ret_time = st.text_input("返廠時間", value=now_time.strftime("%H:%M"))
            def_return = latest_return_mileage + 23 if "順益" in f_dest_select else latest_return_mileage + 170
            f_ret_mil = st.number_input("返廠里程 (請看儀表板輸入)", value=def_return, step=1)

        submit_daily = st.form_submit_button("💾 儲存今日紀錄，同步至雲端", use_container_width=True)
        
        if submit_daily:
            dest_final = f_dest_custom.strip() if f_dest_custom.strip() else f_dest_select
            diff_km = int(f_ret_mil) - int(f_start_mil)
            expected_km = 170 if "美光" in dest_final else (23 if "順益" in dest_final else 50)
            
            new_row = [
                f_date.strftime("%Y-%m-%d"), f_user, dest_final, 
                f_start_time, int(f_start_mil), f_ret_time, int(f_ret_mil), ""
            ]
            ws_daily.append_row(new_row)
            
            if diff_km > (expected_km + 80):
                st.session_state.sys_msg = f"⚠️ 【漏登提醒】本次增加了 {diff_km} km。紀錄已存入，請至「修改」確認是否漏填！"
                st.session_state.sys_msg_type = "warning"
            else:
                st.session_state.sys_msg = f"✅ {f_date.strftime('%m/%d')} 的紀錄已寫入 Google Sheets！"
                st.session_state.sys_msg_type = "success"
            st.rerun()

with tab_history:
    st.subheader("📱 手機專用歷史修改")
    edit_mode = st.radio("請選擇要修改的紀錄類型：", ["日常行車", "柴油紀錄", "尿素紀錄"], horizontal=True)
    
    if edit_mode == "日常行車":
        if not st.session_state.daily_logs:
            st.info("目前沒有資料可修改")
        else:
            all_months = sorted(list(set([str(log.get('日期', ''))[:7] for log in st.session_state.daily_logs if log.get('日期')])), reverse=True)
            selected_month = st.selectbox("📅 第一步：選擇月份", all_months, key="hist_m1")
            log_options = []
            for idx, log in enumerate(st.session_state.daily_logs):
                if str(log.get('日期', '')).startswith(selected_month):
                    log_options.append((idx + 2, f"[{log.get('日期')}] {log.get('目的地')} (返廠: {log.get('返廠里程')})", log))
            log_options = sorted(log_options, key=lambda x: str(x[2].get('日期', '')), reverse=True)
            
            if not log_options:
                st.warning("該月份無紀錄")
            else:
                selected_label = st.selectbox("✏️ 第二步：選擇紀錄", [item[1] for item in log_options])
                target_row_num, _, target_log = next(item for item in log_options if item[1] == selected_label)
                
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
                        ws_daily.update(values=[updated_row], range_name=f"A{target_row_num}:H{target_row_num}")
                        st.session_state.sys_msg = "✅ 日常紀錄已更新！"
                        st.session_state.sys_msg_type = "success"
                        st.rerun()
                    if btn_delete:
                        ws_daily.delete_rows(target_row_num)
                        st.session_state.sys_msg = "🗑️ 該筆紀錄已刪除！"
                        st.session_state.sys_msg_type = "warning"
                        st.rerun()

    elif edit_mode == "柴油紀錄":
        if not st.session_state.fuel_logs:
            st.info("目前沒有柴油資料")
        else:
            log_options = [(idx + 2, f"[{log.get('加油日期')}] {log.get('加油站名')} - {log.get('公升數')}L (里程: {log.get('加油里程')})", log) for idx, log in enumerate(st.session_state.fuel_logs)]
            log_options = sorted(log_options, key=lambda x: str(x[2].get('加油日期', '')), reverse=True)
            
            selected_label = st.selectbox("✏️ 選擇要修改的柴油紀錄", [item[1] for item in log_options])
            target_row_num, _, target_log = next(item for item in log_options if item[1] == selected_label)
            
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
                    ws_fuel.update(values=[upd_row], range_name=f"A{target_row_num}:F{target_row_num}")
                    st.session_state.sys_msg = "✅ 柴油紀錄已更新！"
                    st.session_state.sys_msg_type = "success"
                    st.rerun()
                if btn_del:
                    ws_fuel.delete_rows(target_row_num)
                    st.session_state.sys_msg = "🗑️ 柴油紀錄已刪除！"
                    st.session_state.sys_msg_type = "warning"
                    st.rerun()

    elif edit_mode == "尿素紀錄":
        if not st.session_state.urea_logs:
            st.info("目前沒有尿素資料")
        else:
            log_options = [(idx + 2, f"[{log.get('添加日期')}] {log.get('公升數')}L (里程: {log.get('添加當下里程', '未填')})", log) for idx, log in enumerate(st.session_state.urea_logs)]
            log_options = sorted(log_options, key=lambda x: str(x[2].get('添加日期', '')), reverse=True)
            
            selected_label = st.selectbox("✏️ 選擇要修改的尿素紀錄", [item[1] for item in log_options])
            target_row_num, _, target_log = next(item for item in log_options if item[1] == selected_label)
            
            with st.form("edit_urea_form"):
                e_ud = st.date_input("添加日期", parse_date_safe(target_log.get("添加日期"), TODAY))
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
                    ws_urea.update(values=[upd_row], range_name=f"A{target_row_num}:F{target_row_num}")
                    st.session_state.sys_msg = "✅ 尿素紀錄已更新！"
                    st.session_state.sys_msg_type = "success"
                    st.rerun()
                if btn_del:
                    ws_urea.delete_rows(target_row_num)
                    st.session_state.sys_msg = "🗑️ 尿素紀錄已刪除！"
                    st.session_state.sys_msg_type = "warning"
                    st.rerun()

with tab_stats:
    default_station_mil = latest_return_mileage
    
    st.subheader("⚠️ 硬體異常防呆監控")
    if historical_km_l > 0.1 and len(sorted_fuel_logs) >= 3:
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
    st.subheader("📊 未來進站預測引擎 (絕對里程防呆)")
    if urea_km_l > 1.0 and global_fuel_km_l > 0.1 and avg_daily_km > 1.0:
        safe_fuel_vol = 60 / global_fuel_km_l
        usable_fuel_vol = max(0, vol_after_today - safe_fuel_vol)
        fuel_runnable_km = usable_fuel_vol * global_fuel_km_l
        fuel_days_left = fuel_runnable_km / avg_daily_km
        
        usable_urea_vol = max(0, rem_urea_vol - 8.0)
        urea_runnable_km = usable_urea_vol * urea_km_l
        urea_days_left = urea_runnable_km / avg_daily_km

        target_fuel_mil = latest_return_mileage + math.floor(fuel_runnable_km)
        target_urea_mil = latest_return_mileage + math.floor(urea_runnable_km)

        st.markdown("### 🔄 智慧進站策略 (看里程表最準！)")
        
        if urea_runnable_km >= fuel_runnable_km:
            default_station_mil = target_fuel_mil
            target_date = datetime.now(TW_TZ) + timedelta(days=math.floor(fuel_days_left))
            future_urea_vol = rem_urea_vol - (fuel_runnable_km / urea_km_l)
            
            st.success(f"✅ **預估進站點 (柴油主導)：【 里程 {target_fuel_mil:,} km 】**\n\n*(約在 {target_date.strftime('%m月%d日')} 抵達，請以車上里程表為準)*")
            
            if future_urea_vol >= 8.0:
                st.info(f"💡 **決策：【這次只要加柴油就好】**\n\n開到 **{target_fuel_mil:,} km** 時，尿素約還剩 **{future_urea_vol:.1f} 公升** (安全)，單純補滿柴油即可！")
            else:
                st.warning(f"🟡 **決策：【柴油 ＋ 尿素 順便一起加】**\n\n開到 **{target_fuel_mil:,} km** 時，尿素約剩 **{future_urea_vol:.1f} 公升** (逼近防護線)，建議兩管齊下一次加滿！")
                
        else:
            default_station_mil = target_urea_mil
            target_date = datetime.now(TW_TZ) + timedelta(days=math.floor(urea_days_left))
            future_fuel_vol = vol_after_today - (urea_runnable_km / global_fuel_km_l)
            
            st.error(f"🚨 **尿素提早告急：【 里程 {target_urea_mil:,} km 】**\n\n*(約在 {target_date.strftime('%m月%d日')} 抵達，請以車上里程表為準)*")
            st.warning(f"💡 **決策：【以補尿素為主，柴油順便加滿】**\n\n開到 **{target_urea_mil:,} km** 時尿素將跌破 8L 警戒！屆時柴油約剩 **{future_fuel_vol:.1f} 公升**，請一併補滿。")
    else:
        st.info("🔄 智慧預測引擎需要至少 2 筆以上的柴/尿紀錄來運算。")

    st.divider()
    st.markdown(f"**目前車輛剩餘狀況** (以最新里程 {latest_return_mileage:,} km 結算)")
    
    c_dash1, c_dash2 = st.columns(2)
    with c_dash1:
        st.metric(label="距上次加油已跑", value=f"{used_km_since_fuel} km", delta=f"柴油約剩 {vol_after_today:.1f} L", delta_color="off")
    with c_dash2:
        st.metric("最新一期油耗", f"{global_fuel_km_l} km/L")
        
    rem_safe_km = int(vol_after_today * global_fuel_km_l)
    if rem_safe_km < 60:
        st.error(f"🔴 柴油極危險！含油底續航僅剩 {rem_safe_km} km，請立刻補油！")
    elif 60 <= rem_safe_km < 100:
        st.warning(f"🟡 柴油警戒！續航剩 {rem_safe_km} km，上高速前請確認無塞車。")
    
    st.divider()
    c_u1, c_u2 = st.columns(2)
    with c_u1:
        if urea_with_mil:
            st.metric(label="距上次加尿已跑", value=f"{driven_since_urea} km", delta=f"尿素約剩 {rem_urea_vol:.1f} L", delta_color="off")
        else:
            st.metric(label="距上次加尿已跑", value="資料不足")
    with c_u2:
        if urea_km_l > 1.0:
            st.metric("歷史尿素平均耗能", f"{urea_km_l} km/L")
        else:
            st.metric("近期尿素耗能", "資料不足")
            
    if urea_km_l > 1.0:
        if 4.0 <= rem_urea_vol < 8.0:
            st.warning(f"⚠️ 尿素警戒：僅剩 {rem_urea_vol:.1f} L (逼近 8L 安全防線)")
        elif rem_urea_vol < 4.0:
            st.error(f"🔴 尿素危險：僅剩 {rem_urea_vol:.1f} L！避免抽不到底鎖死，請盡速補加！")

    st.divider()
    
    # ================= 🚀 補給站整合主控中心 (支援雙向一鍵同步) =================
    st.markdown("### ⛽💧 補給站登記中心")
    if default_station_mil > latest_return_mileage:
        st.caption(f"🤖 **神預測同步**：AI 已自動預載預估進站里程，您只需微調最後幾碼即可！")
    else:
        st.caption("💡 輸入當下儀表板里程，可勾選同時登記柴油或尿素，一鍵上傳！")
    
    with st.form("station_master_form"):
        shared_station_mil = st.number_input("補給當下儀表板總里程 (km)", value=int(default_station_mil), step=1)
        station_date = st.date_input("補給日期", SHIFT_DATE)
        
        st.markdown("---")
        # 彈性開關：預設勾選柴油，尿素依據當天是否添加勾選
        chk_fuel = st.checkbox("⛽ 登記柴油", value=True)
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fuel_l = st.number_input("柴油公升數 (L)", value=150.0, step=0.1, disabled=not chk_fuel)
        with col_f2:
            fuel_cost = st.number_input("柴油實付費用 (元)", value=4500, step=10, disabled=not chk_fuel)
        with col_f3:
            fuel_station = st.selectbox("加油站", ["中油苗栗站", "中油新竹站", "其他"], disabled=not chk_fuel)
            
        st.markdown("---")
        chk_urea = st.checkbox("💧 登記尿素", value=False)
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            urea_l = st.number_input("尿素公升數 (L)", value=20.0, step=1.0, disabled=not chk_urea)
        with col_u2:
            urea_cost = st.number_input("尿素費用 (元)", value=500, step=10, disabled=not chk_urea)
            
        submit_station = st.form_submit_button("🚀 一鍵同步儲存所有勾選紀錄", use_container_width=True)
        
        if submit_station:
            if not chk_fuel and not chk_urea:
                st.warning("⚠️ 請至少勾選「柴油」或「尿素」其中一項！")
            else:
                msg_list = []
                cur_hhmm = now_time.strftime("%H:%M")
                
                if chk_fuel:
                    new_fuel_row = [station_date.strftime("%Y-%m-%d"), cur_hhmm, int(shared_station_mil), float(fuel_l), int(fuel_cost), fuel_station]
                    ws_fuel.append_row(new_fuel_row)
                    msg_list.append("柴油")
                    
                if chk_urea:
                    new_urea_row = [station_date.strftime("%Y-%m-%d"), cur_hhmm, float(urea_l), int(urea_cost), "", int(shared_station_mil)]
                    ws_urea.append_row(new_urea_row)
                    msg_list.append("尿素")
                    
                st.session_state.sys_msg = f"✅ 【{' ＋ '.join(msg_list)}】補給紀錄已成功寫入 Google Sheets！"
                st.session_state.sys_msg_type = "success"
                st.rerun()

with tab_export:
    st.subheader("標準紙本 1:1 還原列印")
    exp_y = st.selectbox("年份", [2026, 2027], index=0)
    exp_m = st.selectbox("月份", list(range(1, 13)), index=datetime.now(TW_TZ).month - 1)
    excel_file = generate_kertz_excel(exp_y, exp_m)
    st.download_button(label=f"📥 下載 {exp_y} 年 {exp_m} 月行車紀錄表 (.xlsx)", data=excel_file, file_name=f"行車紀錄表_{DEFAULT_PLATE}_{exp_y}年{exp_m}月.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with tab_settings:
    st.subheader("⚙️ 系統與法規參數設定")
    with st.form("settings_form"):
        new_target = st.number_input("下次保養里程目標 (km)", value=st.session_state.service_target, step=1000)
        
        st.markdown("---")
        st.markdown("**法規檢驗排程 (請依照行照指定日期填寫)**")
        new_v_date = st.date_input("下次驗車到期日", next_v_date)
        new_e_date = st.date_input("下次驗排氣到期日", next_e_date)
        
        st.markdown("---")
        st.markdown(f"**車輛總油箱容量 (公升 L)**\n\n🤖 *目前由 AI 自動微調鎖定在：{tank_cap} L*")
        
        submit_settings = st.form_submit_button("💾 儲存所有設定", use_container_width=True)
        if submit_settings:
            update_vals = [
                ["SERVICE_TARGET", new_target],
                ["FUEL_TANK_CAPACITY", base_tank_capacity],
                ["NEXT_VEHICLE_INSP", new_v_date.strftime("%Y-%m-%d")],
                ["NEXT_EMISSIONS_INSP", new_e_date.strftime("%Y-%m-%d")]
            ]
            ws_settings.update(values=update_vals, range_name="A1:B4")
            
            st.session_state.service_target = new_target
            st.session_state.sys_msg = "✅ 所有保養與檢驗目標已更新並存入雲端！"
            st.session_state.sys_msg_type = "success"
            st.rerun()
