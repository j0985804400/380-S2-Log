import json
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
SERVICE_TARGET = 622000  
SERVICE_INTERVAL = 10000 
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

# ================= 資料同步與快取 =================
def load_data():
    daily_data = ws_daily.get_all_records()
    fuel_data = ws_fuel.get_all_records()
    
    if not daily_data:
        daily_data = [{
            "日期": "2026-10-31", "使用人": "劉嘉舜", "目的地": "美光",
            "出發時間": "07:20", "出發里程": 615049, "返廠時間": "18:30", "返廠里程": 615215, "備註": "系統基準"
        }]
    if not fuel_data:
        fuel_data = [{
            "加油日期": "2026-10-30", "加油時間": "18:40", "加油里程": 614997,
            "公升數": 130.0, "加油費用": 3816, "加油站名": "中油苗栗站"
        }]
        
    st.session_state.daily_logs = daily_data
    st.session_state.fuel_logs = fuel_data
    if "destinations" not in st.session_state:
        st.session_state.destinations = ["美光", "順益", "廠內調度"]

load_data()

latest_log = st.session_state.daily_logs[-1]
latest_return_mileage = int(latest_log["返廠里程"])
latest_fuel_log = st.session_state.fuel_logs[-1]
latest_fuel_mileage = int(latest_fuel_log["加油里程"])

# ================= 動態保養預測邏輯 =================
def calculate_service_prediction(logs, current_mileage, target_mileage):
    if len(logs) < 2: return "資料不足，無法預測"
    recent_logs = logs[-7:] if len(logs) >= 7 else logs
    try:
        first_date = datetime.strptime(str(recent_logs[0]["日期"]), "%Y-%m-%d")
        last_date = datetime.strptime(str(recent_logs[-1]["日期"]), "%Y-%m-%d")
        days_diff = (last_date - first_date).days
        if days_diff == 0: return "需更多天數資料"
        mileage_diff = int(recent_logs[-1]["返廠里程"]) - int(recent_logs[0]["出發里程"])
        avg_daily_km = mileage_diff / max(days_diff, 1) 
        if avg_daily_km <= 0: return "里程未增加"
        km_remaining = target_mileage - current_mileage
        if km_remaining <= 0: return "已達保養里程！"
        days_remaining = int(km_remaining / avg_daily_km)
        est_date = datetime.now(TW_TZ) + timedelta(days=days_remaining)
        return est_date.strftime("%m月%d日")
    except:
        return "計算中..."

est_service_date = calculate_service_prediction(st.session_state.daily_logs, latest_return_mileage, SERVICE_TARGET)

# ================= 智慧加油均算邏輯 =================
driven_since_fuel = latest_return_mileage - latest_fuel_mileage
def get_fuel_recommendation(delta_km):
    if delta_km < 340: return ("🟢 油量充裕（今日免加油）", f"距上次加油已跑 {delta_km} km，油箱尚有餘裕，跑完今日行程完全沒問題。", "success")
    elif 340 <= delta_km < 420: return ("🟡 建議今日【南下】在中油苗栗站加油", f"已跑 {delta_km} km。去程順下苗栗站加滿最穩當，避免回程塞車逼近警戒線。", "warning")
    elif 420 <= delta_km < 480: return ("🔵 建議今日【北上】回程在中油苗栗站加油", f"已跑 {delta_km} km。足夠跑到后里美光，回程北上剛好湊滿約 500 km 進站加滿，節奏最佳。", "info")
    else: return ("🔴 油量警戒：今日【南下】必須立刻加油", f"已跑 {delta_km} km，已達進站極限，去程請務必在中油苗栗站補油！", "error")

fuel_status, fuel_tip, fuel_level = get_fuel_recommendation(driven_since_fuel)
rem_service = SERVICE_TARGET - latest_return_mileage

# ================= 1:1 原樣 A4 報表生成引擎 =================
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
    rich_text = CellRichText(
        TextBlock(InlineFont(b=False, sz=12, rFont=font_family), 'KT.T. '),
        TextBlock(InlineFont(i=True, sz=8, rFont=font_family), 'Pro Clean Tech   '),
        TextBlock(InlineFont(b=False, sz=12, rFont=font_family), '科治新技股份有限公司 標準作業表格')
    )
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
    
    # 修正：加上換行符號 \n，讓「行車」與「日誌」完美上下排列
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
st.set_page_config(page_title="380-S2 智慧行車日誌", page_icon="🚛", layout="centered")
st.title("🚛 380-S2 行車日誌管理系統")
st.caption("連線狀態：🟢 Google Sheets 雲端連線成功")

c_dash1, c_dash2 = st.columns(2)
with c_dash1:
    st.metric(label="目前最新里程", value=f"{latest_return_mileage:,} km", delta=f"距保養剩 {rem_service:,} km ｜ 預估：{est_service_date}", delta_color="normal")
with c_dash2:
    st.metric(label="距上次加油已跑", value=f"{driven_since_fuel} km", delta="目標線 500 km", delta_color="off")

if fuel_level == "success": st.success(f"**{fuel_status}**\n\n{fuel_tip}")
elif fuel_level == "warning": st.warning(f"**{fuel_status}**\n\n{fuel_tip}")
elif fuel_level == "info": st.info(f"**{fuel_status}**\n\n{fuel_tip}")
else: st.error(f"**{fuel_status}**\n\n{fuel_tip}")

tab_daily, tab_history, tab_fuel, tab_export = st.tabs(["📝 返廠登記", "📋 歷史與修改", "⛽ 油/尿素", "🖨️ A4 匯出"])

with tab_daily:
    st.subheader("每日返廠單次結算")
    with st.form("quick_daily_form"):
        f_date = st.date_input("日期", datetime.now(TW_TZ))
        f_user = st.text_input("使用人", value=DEFAULT_USER)
        f_dest_select = st.selectbox("目的地 (點選快捷)", st.session_state.destinations)
        f_dest_custom = st.text_input("或輸入新目的地")

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
            new_row = [
                f_date.strftime("%Y-%m-%d"), f_user, dest_final, 
                f_start_time, int(f_start_mil), f_ret_time, int(f_ret_mil), ""
            ]
            ws_daily.append_row(new_row)
            st.success("✅ 雲端同步成功！資料已寫入 Google Sheets。")
            st.rerun()

with tab_history:
    st.subheader("📋 歷史紀錄與雲端修改")
    st.caption("直接在下方表格「點擊兩下」即可修改打錯的數字或文字，也可選取整列刪除。修改完畢後請務必按下方的【儲存】同步至雲端。")
    
    df_daily = pd.DataFrame(st.session_state.daily_logs)
    edited_df = st.data_editor(df_daily, num_rows="dynamic", use_container_width=True)
    
    if st.button("💾 將修改同步覆蓋至 Google 試算表", use_container_width=True):
        new_data = edited_df.fillna("").values.tolist()
        header = edited_df.columns.tolist()
        ws_daily.clear()
        ws_daily.append_rows([header] + new_data)
        st.success("✅ 修改已成功同步至雲端！")
        st.rerun()

with tab_fuel:
    c_f1, c_f2 = st.columns(2)
    with c_f1:
        st.markdown("### ⛽ 柴油加油")
        with st.form("fuel_form"):
            fuel_d = st.date_input("加油日期", datetime.now(TW_TZ), key="fd")
            fuel_m = st.number_input("加油當下里程", value=latest_return_mileage, step=1)
            fuel_l = st.number_input("公升數 (L)", value=150.0, step=0.1)
            fuel_cost = st.number_input("實付費用 (元)", value=4500, step=10)
            fuel_station = st.selectbox("加油站", ["中油苗栗站", "中油新竹站", "其他"])
            sub_fuel = st.form_submit_button("紀錄柴油並同步", use_container_width=True)

            if sub_fuel:
                new_fuel_row = [
                    fuel_d.strftime("%Y-%m-%d"), datetime.now(TW_TZ).strftime("%H:%M"),
                    int(fuel_m), float(fuel_l), int(fuel_cost), fuel_station
                ]
                ws_fuel.append_row(new_fuel_row)
                st.success("✅ 柴油紀錄已上傳！")
                st.rerun()

    with c_f2:
        st.markdown("### 💧 尿素添加")
        with st.form("urea_form"):
            urea_l = st.number_input("尿素公升數 (L)", value=20.0, step=1.0)
            urea_cost = st.number_input("實付費用 (元)", value=500, step=10)
            sub_urea = st.form_submit_button("紀錄尿素並同步", use_container_width=True)

            if sub_urea:
                new_urea_row = [
                    datetime.now(TW_TZ).strftime("%Y-%m-%d"), datetime.now(TW_TZ).strftime("%H:%M"),
                    float(urea_l), int(urea_cost), ""
                ]
                ws_urea.append_row(new_urea_row)
                st.success("✅ 尿素紀錄已上傳！")
                st.rerun()

with tab_export:
    st.subheader("標準紙本 1:1 還原列印")
    exp_y = st.selectbox("年份", [2026, 2027], index=0)
    exp_m = st.selectbox("月份", list(range(1, 13)), index=datetime.now(TW_TZ).month - 1)
    excel_file = generate_kertz_excel(exp_y, exp_m)

    st.download_button(
        label=f"📥 下載民國 {exp_y-1911} 年 {exp_m} 月行車紀錄表 (.xlsx)",
        data=excel_file,
        file_name=f"公務車輛行車紀錄表_{DEFAULT_PLATE}_{exp_y-1911}年{exp_m}月.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
