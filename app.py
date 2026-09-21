import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

st.set_page_config(
    page_title="DO SUNG - Delivery Schedule 2026", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Tên file Excel dữ liệu
EXCEL_FILE = "2026 DO SUNG会社の納品スケジュール.xlsx"
if not os.path.exists(EXCEL_FILE) and os.path.exists("dosung_schedule.xlsx"):
    EXCEL_FILE = "dosung_schedule.xlsx"

# ----------------- BẢO MẬT ĐĂNG NHẬP ĐƠN GIẢN CHO 2 CÔNG TY -----------------
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.subheader("🔒 ĐĂNG NHẬP HỆ THỐNG GIAO NHẬN DO SUNG 2026")
        pwd = st.text_input("Vui lòng nhập mã bảo mật truy cập:", type="password")
        if st.button("Truy Cập"):
            if pwd == "dosung2026":  # <-- Bạn có thể đổi mật khẩu này theo ý muốn
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Mã bảo mật không đúng, vui lòng thử lại!")
        return False
    return True

if not check_password():
    st.stop()
# -----------------------------------------------------------------------------

st.title("🚚 HỆ THỐNG VẬN HÀNH & THEO DÕI GIAO NHẬN DO SUNG (2026)")
st.caption(f"📁 Tệp dữ liệu kết nối: **{EXCEL_FILE}** | Trạng thái: Đang hoạt động")

if not os.path.exists(EXCEL_FILE):
    st.error(f"❌ Chưa tìm thấy tệp `{EXCEL_FILE}` trong thư mục `dosung-app`!")
    st.stop()

# Đọc danh sách sheet
wb_inspect = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
available_sheets = wb_inspect.sheetnames

tab1, tab2, tab3 = st.tabs(["📅 Quản Lý & Cập Nhật Ngày", "📊 Báo Cáo Tổng Nhập & Hải Quan", "📥 Tải File Excel"])

# ==================== TAB 1: DASHBOARD & CẬP NHẬT TIẾN ĐỘ ====================
with tab1:
    col_sel1, col_sel2 = st.columns([1, 1])
    with col_sel1:
        part_sheets = [s for s in available_sheets if s != "Tổng nhập"]
        selected_part = st.selectbox("📌 Chọn mã linh kiện / Sheet:", part_sheets, index=0)
    with col_sel2:
        selected_month = st.selectbox("🗓️ Chọn tháng theo dõi (2026):", [f"Tháng {m}" for m in range(1, 11)], index=0)
        month_idx = int(selected_month.replace("Tháng ", ""))

    # Đọc dữ liệu sheet và tháng
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    sheet = wb[selected_part]
    
    # Cấu trúc: Tháng 1 bắt đầu ở row 5, mỗi tháng cách nhau 9 dòng
    month_start_row = 5 + (month_idx - 1) * 9
    row_date = month_start_row + 1
    row_po = month_start_row + 3
    row_plan = month_start_row + 4
    row_actual = month_start_row + 5
    row_ng = month_start_row + 6

    days_data = []
    for day in range(1, 32):
        col = 6 + day
        d_val = sheet.cell(row=row_date, column=col).value
        po_val = sheet.cell(row=row_po, column=col).value or 0
        plan_val = sheet.cell(row=row_plan, column=col).value or 0
        act_val = sheet.cell(row=row_actual, column=col).value or 0
        ng_val = sheet.cell(row=row_ng, column=col).value or 0
        
        if d_val is not None:
            days_data.append({
                "Ngày": f"Ngày {day:02d}",
                "Day_Num": day,
                "PO": int(po_val),
                "Kế hoạch (PCS)": int(plan_val),
                "Thực nhập (PCS)": int(act_val),
                "Số lượng NG": int(ng_val),
                "Chênh lệch (Actual - Plan)": int(act_val - plan_val)
            })

    df_month = pd.DataFrame(days_data)

    tot_plan = df_month["Kế hoạch (PCS)"].sum()
    tot_actual = df_month["Thực nhập (PCS)"].sum()
    tot_ng = df_month["Số lượng NG"].sum()
    rate = (tot_actual / tot_plan * 100) if tot_plan > 0 else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🎯 Kế Hoạch Cả Tháng", f"{tot_plan:,.0f} PCS")
    m2.metric("📦 Thực Nhập Đến Nay", f"{tot_actual:,.0f} PCS", delta=f"{tot_actual - tot_plan:,.0f} PCS")
    m3.metric("⚠️ Tổng Lỗi (NG)", f"{tot_ng:,.0f} PCS")
    m4.metric("📈 Tỷ Lệ Hoàn Thành", f"{rate:.1f}%")

    st.divider()

    col_form, col_view = st.columns([1, 2])

    with col_form:
        st.subheader("📝 Nhập Số Lượng Thực Tế")
        with st.form("entry_form"):
            input_day = st.number_input("Chọn ngày giao hàng:", min_value=1, max_value=len(df_month), value=1)
            
            # Lấy số lượng cũ để hiển thị
            cur_act = int(df_month.loc[df_month['Day_Num'] == input_day, 'Thực nhập (PCS)'].values[0])
            cur_ng = int(df_month.loc[df_month['Day_Num'] == input_day, 'Số lượng NG'].values[0])
            
            input_actual = st.number_input("Số lượng THỰC NHẬP (PCS):", min_value=0, step=50, value=cur_act)
            input_ng = st.number_input("Số lượng NG phát sinh:", min_value=0, step=1, value=cur_ng)
            
            btn_save = st.form_submit_button("💾 XÁC NHẬN LƯU VÀO FILE")

            if btn_save:
                wb_w = openpyxl.load_workbook(EXCEL_FILE)
                ws_w = wb_w[selected_part]
                target_col = 6 + int(input_day)
                
                ws_w.cell(row=row_actual, column=target_col, value=int(input_actual))
                ws_w.cell(row=row_ng, column=target_col, value=int(input_ng))
                wb_w.save(EXCEL_FILE)
                
                st.success(f"✅ Đã ghi nhận Ngày {input_day:02d}: Nhập {input_actual:,} PCS | NG {input_ng:,} PCS")
                st.rerun()

    with col_view:
        st.subheader(f"📋 Bảng Chi Tiết {selected_month} - {selected_part}")
        def highlight_diff(val):
            if val < 0:
                return 'color: #ff4b4b; font-weight: bold;'
            elif val > 0:
                return 'color: #09ab3b; font-weight: bold;'
            return ''

        st.dataframe(
            df_month[["Ngày", "PO", "Kế hoạch (PCS)", "Thực nhập (PCS)", "Số lượng NG", "Chênh lệch (Actual - Plan)"]].style.map(
                highlight_diff, subset=['Chênh lệch (Actual - Plan)']
            ),
            use_container_width=True,
            height=380
        )

# ==================== TAB 2: SHEET TỔNG NHẬP & HẢI QUAN ====================
with tab2:
    st.subheader("📑 Báo Cáo Tổng Nhập & Đối Soát Hải Quan")
    wb_tn = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    if "Tổng nhập" in wb_tn.sheetnames:
        ws_tn = wb_tn["Tổng nhập"]
        
        tn_rows = []
        for r in range(4, 12):
            month_label = ws_tn.cell(row=r, column=2).value
            if month_label is not None:
                tn_rows.append({
                    "Tháng": str(month_label),
                    "PAR-011: Đã Giao": ws_tn.cell(row=r, column=3).value or 0,
                    "PAR-011: Đã Khai HQ": ws_tn.cell(row=r, column=4).value or 0,
                    "PAR-011: Tồn Chưa Khai": ws_tn.cell(row=r, column=5).value or 0,
                    "Woodone: Đã Giao": ws_tn.cell(row=r, column=6).value or 0,
                    "Woodone: Đã Khai HQ": ws_tn.cell(row=r, column=7).value or 0,
                    "mda0016: Đã Giao": ws_tn.cell(row=r, column=9).value or 0,
                })
        st.dataframe(pd.DataFrame(tn_rows), use_container_width=True)

        st.subheader("📦 Giám Sát Tồn PO (Đơn Hàng)")
        po_rows = []
        for r in range(4, 11):
            po_code = ws_tn.cell(row=r, column=13).value
            po_qty = ws_tn.cell(row=r, column=14).value
            if po_code:
                po_rows.append({"Mã Đơn Hàng (PO)": str(po_code), "Số Lượng (PCS)": po_qty})
        
        c_po1, c_po2 = st.columns([1, 1])
        with c_po1:
            st.dataframe(pd.DataFrame(po_rows), use_container_width=True)
        with c_po2:
            st.info(f"""
            - **Tổng PO Đã Ký:** {ws_tn.cell(row=13, column=14).value:,.0f} PCS
            - **Tổng Số Lượng Đã Giao:** {ws_tn.cell(row=14, column=14).value:,.0f} PCS
            - **Tồn PO Chưa Giao:** **{ws_tn.cell(row=15, column=14).value:,.0f} PCS**
            """)

# ==================== TAB 3: TẢI FILE ====================
with tab3:
    st.subheader("📥 Tải Lại Tệp Excel Cập Nhật Mới Nhất")
    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            label="Tải tệp Excel về máy (.xlsx)",
            data=f,
            file_name=f"DO_SUNG_Schedule_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )