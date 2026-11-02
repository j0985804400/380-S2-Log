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

# ================= 系統固定參數與時區設定 =================
DEFAULT_PLATE = "380-S2"
DEFAULT_KEEPER = "劉嘉舜"
DEFAULT_USER = "劉嘉舜"
DOC_CODE = "4-33-DC0007B"
FORM_NAME = "公務車輛行車紀錄表 Rev.1"
UREA_TANK_CAPACITY = 20.0  
TW_TZ = timezone(timedelta(hours=8))

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
    ws_settings.update(values=[["SERVICE_TARGET", 622000], ["FUEL_TANK_CAPACITY", 130.0]], range_name="A1:B2")
else:
    ws_settings = sh.worksheet("Settings")

try:
    settings_vals = ws_settings.get_all_values()
    settings_dict = {row[0]: row[1] for row in settings_vals if len(row) >= 2}
except:
    settings_dict = {}

st.session_state.service_target = int(settings_dict.get("SERVICE_TARGET", 622000))
base_tank_capacity = float(settings_dict.get("FUEL_TANK_CAPACITY", 130.0))

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

sorted_all_logs = sorted(st.session_state.daily_logs, key=lambda x: str(x.get("日期", "")))
latest_log = sorted_all_logs[-1]
latest_return_mileage = int(latest_log["返廠里程"])
rem_service = st.session_state.service_target - latest_return_mileage

sorted_fuel_logs = sorted(st.session_state.fuel_logs, key=lambda x: int(x.get("加油里程", 0)))
latest_fuel_log = sorted_fuel_logs[-1]
latest_fuel_mileage = int(latest_fuel_log["加油里程"])

if len(sorted_fuel_logs) >= 2:
    try:
        last_f = sorted_fuel_logs[-1]
        prev_f = sorted_fuel_logs[-2]
        km_driven = int(last_f.get("加油里程", 0)) - int(prev_f.get("加油里程", 0))
        liters = float(last_f.get("公升數", 1))
        global_fuel_km_l = round(km_driven / liters, 2) if liters > 0 else 3.8
    except:
        global_fuel_km_l = 3.8
else:
    global_fuel_km_l = 3.8

# ================= 提取日均行駛里程 =================
def get_avg_daily_km(logs):
    recent_logs = logs[-30:] if len(logs) >= 30 else logs
    if len(recent_logs) >= 2:
        try:
            first_date = datetime.strptime(str(recent_logs[0]["日期"]), "%Y-%m-%d")
            last_date = datetime.strptime(str(recent_logs[-1]["日期"]), "%Y-%m-%d")
            days_diff = (last_date - first_date).days
            if days_diff > 0:
                mileage_diff = int(recent_logs[-1]["返廠里程"]) - int(recent_logs[0]["出發里程"])
                return mileage_diff / days_diff
        except:
            pass
    return 170.0 

avg_daily_km = get_avg_daily_km(sorted_all_logs)

# ================= AI 油箱容量自動微調 =================
max_liters_added = max([float(f.get("公升數", 0)) for f in sorted_fuel_logs]) if sorted_fuel_logs else 130.0
learned_tank_cap = float(math.ceil(max_liters_added / 10.0) * 10)
tank_cap = max(base_tank_capacity, learned_tank_cap)

# ================= 連續虛擬油箱 =================
curr_vol = tank_cap  
last_f_mil = int(sorted_fuel_logs[0].get("加油里程", 0)) if sorted_fuel_logs else latest_return_mileage
for f_log in sorted_fuel_logs:
    f_mil = int(f_log.get("加油里程", 0))
    add_l = float(f_log.get("公升數", 0))
    dist = f_mil - last_f_mil
    if dist > 0: curr_vol -= (dist / global_fuel_km_l)
    if curr_vol < 0: curr_vol = 0
    curr_vol += add_l
    if curr_vol > tank_cap: curr_vol = tank_cap
    last_f_mil = f_mil

used_km_since_fuel = latest_return_mileage - latest_fuel_mileage
used_vol_since_fuel = used_km_since_fuel / global_fuel_km_l
vol_after_today = max(0, curr_vol - used_vol_since_fuel)

# ================= 尿素消耗推算 =================
sorted_urea = sorted(st.session_state.urea_logs, key=lambda x: str(x.get("添加日期", "")))
urea_with_mil = [u for u in sorted_urea if "添加當下里程" in u and str(u["添加當下里程"]).isdigit()]

urea_km_l = 0.0
driven_since_urea = 0
rem_urea_vol = 0.0

if urea_with_mil:
    last_urea_mil = int(urea_with_mil[-1]["添加當下里程"])
    driven_since_urea = latest_return_mileage - last_urea_mil
    
if len(urea_with_mil) >= 2:
    try:
        last_u = urea_with_mil[-1]
        prev_u = urea_with_mil[-2]
        km_u_driven = int(last_u["添加當下里程"]) - int(prev_u["添加當下里程"])
        u_liters = float(last_u.get("公升數", 1))
        if u_liters > 0:
            urea_km_l = round(km_u_driven / u_liters, 2)
            used_urea_vol = driven_since_urea / urea_km_l
            rem_urea_vol = max(0.0, UREA_TANK_CAPACITY - used_urea_vol)
    except:
        pass

# ================= 動態保養預測 =================
def calculate_service_prediction(logs, target_mileage):
    if len(logs) < 2: return "資料不足"
    current_mileage = int(logs[-1].get("返廠里程", 0))
    km_remaining = target_mileage - current_mileage
    if km_remaining <= 0: return "已達保養里程"
    if avg_daily_km <= 0: return "計算中"
    days_remaining = int(km_remaining / avg_daily_km)
    est_date = datetime.now(TW_TZ) + timedelta(days=days_remaining)
    return est_date.strftime("%m月%d日")

est_service_date = calculate_service_prediction(sorted_all_logs, st.session_state.service_target)

# ================= 1:1 原樣 A4 報表生成 =================
def generate_kertz_excel(year: int, month: int):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{month}月行車紀錄"
    ws.views.sheetView[0].showGridLines = True
    font_family = "Microsoft JhengHei"
    font_bold_md = Font(name=font_family, size=10, bold=True)
    font_regular = Font(name=font_family, size=9.5)
    font_header_sub = Font(name=font_family, size=8.5)
    thin = Side(border_style="thin", color="000000")
    border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_center_no_wrap = Alignment(horizontal="center", vertical="center", wrap_text=False)

    ws.merge_cells("A1:G1")
    rich_text = CellRichText(TextBlock(InlineFont(b=False, sz=12, rFont=font_family), 'KT.T. '), TextBlock(InlineFont(i=True, sz=8, rFont=font_family), 'Pro Clean Tech   '), TextBlock(InlineFont(b=False, sz=12, rFont=font_family), '科治新技股份有限公司 標準作業表格'))
    ws["A1"] = rich_text
    ws["A1"].alignment = align_center_no_wrap
    ws.merge_cells("H1:I1")
    ws["H1"] = "頁次："
    ws["H1"].font = font_header_sub
    ws["H1"].alignment = align_left
    ws.merge_cells("A2:B2")
    ws["A2"] = "表格名稱"
    ws["A2"].font = Font(name=font_family, size=11)
    ws["A2"].alignment = align_center
    ws.merge_cells("C2:G2")
    ws["C2"] = "公務車輛行車紀錄表 Rev.1"
    ws["C2"].font = Font(name=font_family, size=11)
    ws["C2"].alignment = align_center_no_wrap
    ws.merge_cells("H2:I2")
    ws["H2"] = f"編號：{DOC_CODE}"
    ws["H2"].font = Font(name=font_family, size=9)
    ws["H2"].alignment = align_left

    for r in range(1, 3):
        for c in range(1, 10): ws.cell(row=r, column=c).border = border_all

    roc_year = year - 1911
    ws.row_dimensions[3].height = 24
    ws["A3"] = "行車\n日誌"
    ws["A3"].font = font_bold_md
    ws["A3"].alignment = align_center
    ws.merge_cells("B3:C3")
    ws["B3"] = f"月份： {roc_year} 年 {month} 月"
    ws["B3"].font = font_bold_md
    ws["B3"].alignment = align_center
    ws.merge_cells("D3:E3")
    ws["D3"] = f"車牌號碼： {DEFAULT_PLATE}"
    ws["D3"].font = font_bold_md
    ws["D3"].alignment = align_center
    ws.merge_cells("F3:I3")
    ws["F3"] = f"保管人： {DEFAULT_KEEPER}"
    ws["F3"].font = font_bold_md
    ws["F3"].alignment = align_center

    for c in range(1, 10): ws.cell(row=3, column=c).border = border_all

    headers = ["日期", "使用人", "目的地", "出發時間", "出發里程數", "返廠時間", "返廠里程數", "加油里程數", "加油費用"]
    ws.row_dimensions[4].height = 26
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=text)
        cell.font = font_bold_md
        cell.alignment = align_center
        cell.border = border_all
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    prefix = f"{year:04d}-{month:02d}"
    m_logs = [l for l in st.session_state.daily_logs if str(l.get("日期", "")).startswith(prefix)]
    m_logs = sorted(m_logs, key=lambda x: str(x.get("日期", "")))
    f_map = {str(fl["加油日期"]): fl for fl in st.session_state.fuel_logs if str(fl.get("加油日期", "")).startswith(prefix)}

    start_row = 5
    for i in range(31):
        curr_row = start_row + i
        ws.row_dimensions[curr_row].height = 20
        if i < len(m_logs):
            log = m_logs[i]
            day_num = int(log["日期"].split("-")[2])
            ws.cell(row=curr_row, column=1, value=str(day_num)).alignment = align_center
            ws.cell(row=curr_row, column=2, value=log["使用人"]).alignment = align_center
            ws.cell(row=curr_row, column=3, value=log["目的地"]).alignment = align_center
            ws.cell(row=curr_row, column=4, value=log["出發時間"]).alignment = align_center
            ws.cell(row=curr_row, column=5, value=log["出發里程"]).alignment = align_right
            ws.cell(row=curr_row, column=6, value=log["返廠時間"]).alignment = align_center
            ws.cell(row=curr_row, column=7, value=log["返廠里程"]).alignment = align_right
            if log["日期"] in f_map:
                fl = f_map[log["日期"]]
                ws.cell(row=curr_row, column=8, value=fl["加油里程"]).alignment = align_right
                ws.cell(row=curr_row, column=9, value=fl["加油費用"]).alignment = align_right
        else:
            for c in range(1, 10): ws.cell(row=curr_row, column=c, value="")

        for c in range(1, 10):
            cell = ws.cell(row=curr_row, column=c)
            cell.font = font_regular
            cell.border = border_all
            if c in [5, 7, 8, 9] and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"

    col_widths = {"A": 7.5, "B": 9, "C": 10, "D": 10, "E": 12, "F": 10, "G": 12, "H": 12, "I": 11}
    for col, width in col_widths.items(): ws.column_dimensions[col].width = width

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
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

st.metric(label="目前最新里程", value=f"{latest_return_mileage:,} km", delta=f"距保養剩 {rem_service:,} km ｜ 預估：{est_service_date}", delta_color="normal")

if "sys_msg" in st.session_state:
    if st.session_state.sys_msg_type == "warning": st.warning(st.session_state.sys_msg)
    else: st.success(st.session_state.sys_msg)
    del st.session_state.sys_msg
    del st.session_state.sys_msg_type

tab_daily, tab_history, tab_stats, tab_export, tab_settings = st.tabs(["📝 登記", "📋 修改", "📊 油耗預測", "🖨️ 匯出", "⚙️ 設定"])

with tab_daily:
    st.subheader("每日返廠單次結算")
    with st.form("quick_daily_form"):
        f_date = st.date_input("日期", datetime.now(TW_TZ))
        f_user = st.text_input("使用人", value=DEFAULT_USER)
        f_dest_select = st.selectbox("目的地 (點選快捷)", st.session_state.destinations)
        f_dest_custom = st.text_input("或輸入新目的地 (選填)")

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            f_start_time = st.text_input("出發時間", value="07:20")
            f_start_mil = st.number_input("出發里程 (鎖定前次)", value=latest_return_mileage, disabled=True)
        with col_t2:
            f_ret_time = st.text_input("返廠時間", value=datetime.now(TW_TZ).strftime("%H:%M"))
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
    if not st.session_state.daily_logs:
        st.info("目前沒有歷史資料可修改")
    else:
        all_months = sorted(list(set([str(log.get('日期', ''))[:7] for log in st.session_state.daily_logs if log.get('日期')])), reverse=True)
        selected_month = st.selectbox("📅 第一步：選擇月份", all_months)
        
        log_options = []
        for idx, log in enumerate(st.session_state.daily_logs):
            sheet_row_num = idx + 2
            if str(log.get('日期', '')).startswith(selected_month):
                label = f"[{log.get('日期')}] {log.get('目的地')} (返廠: {log.get('返廠里程')})"
                log_options.append((sheet_row_num, label, log))
                
        log_options = sorted(log_options, key=lambda x: str(x[2].get('日期', '')), reverse=True)
        
        if not log_options:
            st.warning("該月份無紀錄")
        else:
            option_labels = [item[1] for item in log_options]
            selected_label = st.selectbox("✏️ 第二步：選擇要修改的紀錄", option_labels)
            selected_item = next(item for item in log_options if item[1] == selected_label)
            target_row_num, _, target_log = selected_item
            
            with st.form("mobile_edit_form"):
                e_date = st.text_input("日期 (YYYY-MM-DD)", value=str(target_log.get("日期")))
                e_user = st.text_input("使用人", value=str(target_log.get("使用人")))
                e_dest = st.text_input("目的地", value=str(target_log.get("目的地")))
                
                c_e1, c_e2 = st.columns(2)
                with c_e1:
                    e_st_time = st.text_input("出發時間", value=str(target_log.get("出發時間")))
                    e_st_mil = st.number_input("出發里程", value=int(target_log.get("出發里程", 0)), step=1)
                with c_e2:
                    e_ret_time = st.text_input("返廠時間", value=str(target_log.get("返廠時間")))
                    e_ret_mil = st.number_input("返廠里程", value=int(target_log.get("返廠里程", 0)), step=1)
                
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    btn_save_edit = st.form_submit_button("💾 儲存修改", use_container_width=True)
                with c_btn2:
                    btn_delete = st.form_submit_button("🗑️ 刪除紀錄", use_container_width=True)
                
                if btn_save_edit:
                    updated_row = [e_date, e_user, e_dest, e_st_time, int(e_st_mil), e_ret_time, int(e_ret_mil), str(target_log.get("備註", ""))]
                    ws_daily.update(values=[updated_row], range_name=f"A{target_row_num}:H{target_row_num}")
                    st.session_state.sys_msg = "✅ 資料已更新至雲端！"
                    st.session_state.sys_msg_type = "success"
                    st.rerun()
                if btn_delete:
                    ws_daily.delete_rows(target_row_num)
                    st.session_state.sys_msg = "🗑️ 該筆紀錄已徹底刪除！"
                    st.session_state.sys_msg_type = "warning"
                    st.rerun()

with tab_stats:
    st.subheader("📊 未來進站日預測引擎")
    
    # ---------------- 智慧決策分離運算 ----------------
    if urea_km_l > 0 and global_fuel_km_l > 0 and avg_daily_km > 0:
        safe_fuel_vol = 60 / global_fuel_km_l
        usable_fuel_vol = max(0, vol_after_today - safe_fuel_vol)
        fuel_runnable_km = usable_fuel_vol * global_fuel_km_l
        fuel_days_left = fuel_runnable_km / avg_daily_km
        
        # 尿素容錯機制：以 8L 為絕對安全進站基準 (防護儀表非線性與抽不到底)
        usable_urea_vol = max(0, rem_urea_vol - 8.0)
        urea_runnable_km = usable_urea_vol * urea_km_l
        urea_days_left = urea_runnable_km / avg_daily_km

        st.markdown("### 🔄 智慧進站策略 (下次該加什麼？)")
        
        if urea_runnable_km >= fuel_runnable_km:
            target_date = datetime.now(TW_TZ) + timedelta(days=math.floor(fuel_days_left))
            future_urea_vol = rem_urea_vol - (fuel_runnable_km / urea_km_l)
            
            st.success(f"✅ **預估進站日 (柴油主導)：【 {target_date.strftime('%m月%d日')} 】 (約再跑 {math.floor(fuel_days_left)} 天)**")
            
            # 以 8L 安全線為基準，避免柴油剛加滿隔天尿素就見底的尷尬狀況
            if future_urea_vol >= 8.0:
                st.info(f"💡 **決策：【這次只要加柴油就好】**\n\n預計到了那天進站時，尿素大約還剩 **{future_urea_vol:.1f} 公升 (大於 8L 安全底線，非常充足)**，單純補滿柴油即可！")
            else:
                st.warning(f"🟡 **決策：【柴油 ＋ 尿素 順便一起加】**\n\n預計到了那天進站時，尿素大約只剩 **{future_urea_vol:.1f} 公升** (準備逼近 8L 安全防護線)，建議直接兩管齊下一次加滿！")
                
        else:
            target_date = datetime.now(TW_TZ) + timedelta(days=math.floor(urea_days_left))
            future_fuel_vol = vol_after_today - (urea_runnable_km / global_fuel_km_l)
            
            st.error(f"🚨 **尿素告急！需提早進站：【 {target_date.strftime('%m月%d日')} 】 (約再跑 {math.floor(urea_days_left)} 天)**")
            st.warning(f"💡 **決策：【以補尿素為主，柴油順便加滿】**\n\n因為尿素即將跌破 8L 危險警戒，必須提早進站！雖然屆時柴油大約還剩 **{future_fuel_vol:.1f} 公升**，但請直接連同柴油一起補滿。")
    else:
        st.info("🔄 智慧預測引擎需要至少 2 筆以上的柴/尿紀錄，才能為您精準預測下次進站時機。")

    # ---------------- 基礎耗能與當下狀態 ----------------
    st.divider()
    st.markdown(f"**目前車輛剩餘狀況** (以今日返廠 {latest_return_mileage} km 結算)")
    
    c_dash1, c_dash2 = st.columns(2)
    with c_dash1:
        st.metric(label="距上次加油已跑", value=f"{used_km_since_fuel} km", delta=f"柴油約剩 {vol_after_today:.1f} L", delta_color="off")
    with c_dash2:
        st.metric("近期平均油耗", f"{global_fuel_km_l} km/L")
        
    rem_safe_km = int(vol_after_today * global_fuel_km_l)
    if rem_safe_km < 60: st.error(f"🔴 柴油極危險！含油底續航僅剩 {rem_safe_km} km，請立刻就近補油！")
    elif 60 <= rem_safe_km < 100: st.warning(f"🟡 柴油警戒！續航剩 {rem_safe_km} km，上高速前請確認無塞車。")
    
    st.divider()
    c_u1, c_u2 = st.columns(2)
    with c_u1:
        if urea_with_mil:
            st.metric(label="距上次加尿已跑", value=f"{driven_since_urea} km", delta=f"尿素約剩 {rem_urea_vol:.1f} L", delta_color="off")
        else:
            st.metric(label="距上次加尿已跑", value="資料不足")
    with c_u2:
        if urea_km_l > 0:
            st.metric("近期尿素耗能", f"{urea_km_l} km/L")
        else:
            st.metric("近期尿素耗能", "資料不足")
            
    if urea_km_l > 0:
        if 4.0 <= rem_urea_vol < 8.0: st.warning(f"⚠️ 尿素警戒：僅剩 {rem_urea_vol:.1f} L (逼近 8L 安全防線)")
        elif rem_urea_vol < 4.0: st.error(f"🔴 尿素危險：僅剩 {rem_urea_vol:.1f} L！避免抽不到底鎖死，請盡速補加！")

    # ---------------- 實體表單 ----------------
    st.divider()
    c_f1, c_f2 = st.columns(2)
    with c_f1:
        st.markdown("### ⛽ 柴油加油")
        with st.form("fuel_form"):
            fuel_d = st.date_input("加油日期", datetime.now(TW_TZ), key="fd")
            fuel_m = st.number_input("加油當下里程", value=latest_return_mileage, step=1)
            fuel_l = st.number_input("公升數 (L)", value=150.0, step=0.1)
            fuel_cost = st.number_input("實付費用 (元)", value=4500, step=10)
            fuel_station = st.selectbox("加油站", ["中油苗栗站", "中油頭份站", "其他"])
            sub_fuel = st.form_submit_button("紀錄柴油", use_container_width=True)
            if sub_fuel:
                new_fuel_row = [fuel_d.strftime("%Y-%m-%d"), datetime.now(TW_TZ).strftime("%H:%M"), int(fuel_m), float(fuel_l), int(fuel_cost), fuel_station]
                ws_fuel.append_row(new_fuel_row)
                st.session_state.sys_msg = "✅ 柴油紀錄已上傳！"
                st.session_state.sys_msg_type = "success"
                st.rerun()

    with c_f2:
        st.markdown("### 💧 尿素添加")
        with st.form("urea_form"):
            urea_d = st.date_input("添加日期", datetime.now(TW_TZ), key="ud")
            urea_m = st.number_input("添加當下里程", value=latest_return_mileage, step=1)
            urea_l = st.number_input("尿素公升數 (L)", value=20.0, step=1.0)
            urea_cost = st.number_input("實付費用 (元)", value=500, step=10)
            sub_urea = st.form_submit_button("紀錄尿素", use_container_width=True)
            if sub_urea:
                new_urea_row = [urea_d.strftime("%Y-%m-%d"), datetime.now(TW_TZ).strftime("%H:%M"), float(urea_l), int(urea_cost), "", int(urea_m)]
                ws_urea.append_row(new_urea_row)
                st.session_state.sys_msg = "✅ 尿素紀錄已上傳！"
                st.session_state.sys_msg_type = "success"
                st.rerun()

with tab_export:
    st.subheader("標準紙本 1:1 還原列印")
    exp_y = st.selectbox("年份", [2026, 2027], index=0)
    exp_m = st.selectbox("月份", list(range(1, 13)), index=datetime.now(TW_TZ).month - 1)
    excel_file = generate_kertz_excel(exp_y, exp_m)
    st.download_button(label=f"📥 下載 {exp_y} 年 {exp_m} 月行車紀錄表 (.xlsx)", data=excel_file, file_name=f"行車紀錄表_{DEFAULT_PLATE}_{exp_y}年{exp_m}月.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with tab_settings:
    st.subheader("⚙️ 系統參數設定")
    with st.form("settings_form"):
        new_target = st.number_input("下次保養里程目標 (km)", value=st.session_state.service_target, step=1000)
        st.markdown(f"**車輛總油箱容量 (公升 L)**\n\n🤖 *目前由 AI 自動微調鎖定在：{tank_cap} L*")
        submit_settings = st.form_submit_button("💾 儲存保養目標", use_container_width=True)
        if submit_settings:
            ws_settings.update(values=[["SERVICE_TARGET", new_target], ["FUEL_TANK_CAPACITY", base_tank_capacity]], range_name="A1:B2")
            st.session_state.service_target = new_target
            st.session_state.sys_msg = f"✅ 保養目標已永久更新為 {new_target:,} km！"
            st.session_state.sys_msg_type = "success"
            st.rerun()
