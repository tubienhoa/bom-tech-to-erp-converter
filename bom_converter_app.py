import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: EBOM to MBOM (New Template)",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Chuyển Đổi BOM Kỹ Thuật Mới Sang BOM ERP")
st.caption("Tương thích Template EBOM 5 Cấp bậc, Routing QTSX_01 - QTSX_10 & Công thức hao hụt chuẩn xưởng thép")

# Sidebar cấu hình
with st.sidebar:
    st.header("🔧 Cấu hình")
    cust_code = st.text_input("Mã Khách hàng (KKK):", value="STV", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm (FG):", ["Bộ", "Cái"], index=0)
    st.markdown("---")
    st.markdown("""
    **Quy chuẩn Routing tự động:**
    - **QTSX_01**: Cắt (`CUT`), Chấn/Uốn (`BD`)
    - **QTSX_02**: Hàn mài (`WD`), Phay tiện (`MC`)
    - **QTSX_03**: Sơn (`PNT`)
    - **QTSX_04**: Đóng gói (`PK`)
    - **QTSX_09**: Lắp ráp / Đóng PEM (`AS`)
    - **QTSX_10**: Xi mạ (`PL`)
    """)

def resolve_routing(row):
    """Bóc tách Routing từ cột 31 đến 40 của Template BOM kỹ thuật mới"""
    has_buy_out = pd.notna(row[31]) and str(row[31]).strip() != ''
    has_cut     = pd.notna(row[32]) and str(row[32]).strip() != ''
    has_bend    = pd.notna(row[33]) and str(row[33]).strip() != ''
    has_weld    = pd.notna(row[34]) and str(row[34]).strip() != ''
    has_mach    = pd.notna(row[35]) and str(row[35]).strip() != ''
    has_form    = pd.notna(row[36]) and str(row[36]).strip() != ''
    has_paint   = pd.notna(row[37]) and str(row[37]).strip() != ''
    has_plate   = pd.notna(row[38]) and str(row[38]).strip() != ''
    has_assy    = pd.notna(row[39]) and str(row[39]).strip() != ''
    has_pack    = pd.notna(row[40]) and str(row[40]).strip() != ''

    stages = []
    qtsx = ""

    if has_cut and has_bend:
        qtsx = "QTSX_01"
        stages.append("CUT, BD")
    elif has_cut:
        qtsx = "QTSX_05"
        stages.append("CUT")
    elif has_bend:
        stages.append("BD")

    if has_weld and has_mach:
        qtsx = "QTSX_02"
        stages.append("WD, MC")
    elif has_weld:
        qtsx = "QTSX_06" if not qtsx else qtsx
        stages.append("WD")
    elif has_mach:
        qtsx = "QTSX_07" if not qtsx else qtsx
        stages.append("MC")

    if has_paint:
        stages.append("PNT")
    if has_plate:
        stages.append("PL")
    if has_assy:
        stages.append("AS")
    if has_pack:
        stages.append("PK")

    if not qtsx:
        if has_paint:
            qtsx = "QTSX_03"
        elif has_pack:
            qtsx = "QTSX_04"
        elif has_assy:
            qtsx = "QTSX_09"
        elif has_plate:
            qtsx = "QTSX_10"

    return qtsx, ", ".join(stages)

def parse_tech_bom_new(file_bytes, customer_code):
    df_raw = pd.read_excel(file_bytes, sheet_name='Form', header=None)

    # Lấy thông tin Header Dự án
    proj_raw = str(df_raw.iloc[1, 6]).strip() if pd.notna(df_raw.iloc[1, 6]) else "PROJECT"
    pcode_match = re.search(r'([A-Za-z0-9\-]+)', proj_raw)
    pcode = pcode_match.group(1) if pcode_match else proj_raw
    fg_name = str(df_raw.iloc[6, 6]).strip() if pd.notna(df_raw.iloc[6, 6]) else proj_raw

    erp_rows = []

    # 1. Dòng Thành phẩm Level 0 (FG)
    erp_rows.append({
        'STT': 'A',
        'QTSX': '',
        'Mã sản phẩm': f"FG - {customer_code} - {pcode}",
        'Loại': 'THÀNH PHẨM',
        'Tên sản phẩm': fg_name,
        'SL sản phẩm': 1,
        'ĐVT sp': default_uom_fg,
        'Mã cũ': '',
        'Mã NVL': '',
        'Tên NVL': '',
        'ĐVT (NVL)': '',
        'Số lượng': '',
        'ĐVT Kho': '',
        'Số lượng kho': '',
        'Số lượng setup': '',
        '% Tỉ lệ hao hụt': '',
        'Công đoạn': ''
    })

    current_idx = [0, 0, 0, 0, 0]

    # Duyệt từ row 7 trở đi
    r = 7
    while r < len(df_raw):
        row = df_raw.iloc[r]

        # Xác định Level dựa trên số thực/nguyên ở Cột 0 -> 4
        level = None
        for c in range(4, -1, -1):
            val = row[c]
            if pd.notna(val) and isinstance(val, (int, float)):
                level = c
                break

        part_no = str(row[5]).strip() if pd.notna(row[5]) else ""
        desc = str(row[6]).strip() if pd.notna(row[6]) else ""

        if level is None or (not part_no and not desc):
            r += 1
            continue

        current_idx[level] += 1
        for lower in range(level + 1, 5):
            current_idx[lower] = 0

        # Tạo chuỗi STT phân cấp ERP (1, 1.1, 1.1.1...)
        active_indices = [str(current_idx[i]) for i in range(1, level + 1) if current_idx[i] > 0]
        stt_display = ".".join(active_indices) if active_indices else str(current_idx[0])

        qty = row[18] if pd.notna(row[18]) else 1
        uom = str(row[26]).strip() if pd.notna(row[26]) else "Cái"
        mat_type = str(row[29]).strip().lower() if pd.notna(row[29]) else ""
        thickness = row[13]

        output_wt = float(row[23]) if pd.notna(row[23]) and str(row[23]).strip() != '' else 0.0
        tot_raw_wt = float(row[20]) if pd.notna(row[20]) and str(row[20]).strip() != '' else 0.0

        loss_pct = ""
        if output_wt > 0 and tot_raw_wt > 0:
            loss_pct = round(((tot_raw_wt - output_wt) / output_wt) * 100, 2)

        qtsx_code, stage_text = resolve_routing(row)

        # Cấp 1, 2, 3: Cụm lắp ráp / Bán thành phẩm (SG)
        if level in [1, 2]:
            erp_rows.append({
                'STT': stt_display,
                'QTSX': '',
                'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                'Loại': 'BTP',
                'Tên sản phẩm': desc if desc else part_no,
                'SL sản phẩm': qty,
                'ĐVT sp': uom if uom else 'Bộ',
                'Mã cũ': '', 'Mã NVL': '', 'Tên NVL': '', 'ĐVT (NVL)': '', 'Số lượng': '',
                'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                'Công đoạn': stage_text if stage_text else 'WD'
            })

        # Cấp 3: Cụm chi tiết con hoặc Chi tiết đơn
        elif level == 3:
            has_child = (r + 1 < len(df_raw) and pd.notna(df_raw.iloc[r + 1, 4]) and isinstance(df_raw.iloc[r + 1, 4], (int, float)))

            if has_child:
                # Cụm có lắp phụ kiện/PEM
                erp_rows.append({
                    'STT': stt_display,
                    'QTSX': '',
                    'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': desc if desc else f"CỤM {part_no}",
                    'SL sản phẩm': qty,
                    'ĐVT sp': uom,
                    'Mã cũ': '', 'Mã NVL': '', 'Tên NVL': '', 'ĐVT (NVL)': '', 'Số lượng': '',
                    'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                    'Công đoạn': stage_text
                })
            else:
                # Chi tiết đơn tấm không có phụ kiện đi kèm -> QTSX_01
                erp_rows.append({
                    'STT': stt_display,
                    'QTSX': qtsx_code if qtsx_code else 'QTSX_01',
                    'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': desc if desc else part_no,
                    'SL sản phẩm': qty,
                    'ĐVT sp': uom,
                    'Mã cũ': '',
                    'Mã NVL': '',
                    'Tên NVL': f"Thép tấm {thickness}mm" if pd.notna(thickness) else "Thép tấm",
                    'ĐVT (NVL)': 'Kg',
                    'Số lượng': round(tot_raw_wt, 3) if tot_raw_wt > 0 else '',
                    'ĐVT Kho': 'Kg',
                    'Số lượng kho': round(tot_raw_wt, 3) if tot_raw_wt > 0 else '',
                    'Số lượng setup': '',
                    '% Tỉ lệ hao hụt': loss_pct,
                    'Công đoạn': stage_text if stage_text else 'CUT, BD'
                })

        # Cấp 4: Phôi thép tấm hoặc Vật tư phụ (PEM, Pin, Bulong)
        elif level == 4:
            if "phụ" in mat_type or "fa" in part_no.lower():
                # Vật tư phụ tiêu hao (Đóng PEM/Lắp ráp)
                erp_rows.append({
                    'STT': '',
                    'QTSX': '',
                    'Mã sản phẩm': f"PT - {customer_code} - {part_no}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': desc,
                    'SL sản phẩm': qty,
                    'ĐVT sp': uom,
                    'Mã cũ': '',
                    'Mã NVL': part_no,
                    'Tên NVL': desc,
                    'ĐVT (NVL)': uom,
                    'Số lượng': qty,
                    'ĐVT Kho': uom,
                    'Số lượng kho': qty,
                    'Số lượng setup': '',
                    '% Tỉ lệ hao hụt': '',
                    'Công đoạn': 'ĐÓNG PEM'
                })
            else:
                # Phôi thép tấm gia công -> QTSX_01
                erp_rows.append({
                    'STT': '',
                    'QTSX': qtsx_code if qtsx_code else 'QTSX_01',
                    'Mã sản phẩm': f"PT - {customer_code} - {part_no}",
                    'Loại': '',
                    'Tên sản phẩm': desc if desc else '',
                    'SL sản phẩm': qty,
                    'ĐVT sp': uom,
                    'Mã cũ': '',
                    'Mã NVL': '',
                    'Tên NVL': f"Thép tấm {thickness}mm" if pd.notna(thickness) else "Thép tấm",
                    'ĐVT (NVL)': 'Kg',
                    'Số lượng': round(tot_raw_wt, 3) if tot_raw_wt > 0 else '',
                    'ĐVT Kho': 'Kg',
                    'Số lượng kho': round(tot_raw_wt, 3) if tot_raw_wt > 0 else '',
                    'Số lượng setup': '',
                    '% Tỉ lệ hao hụt': loss_pct,
                    'Công đoạn': stage_text if stage_text else 'CUT, BD'
                })
        r += 1

    return pd.DataFrame(erp_rows)

uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật mới (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_result = parse_tech_bom_new(uploaded_file, cust_code)
        st.success(f"✅ Đã xử lý thành công {len(df_result)} dòng BOM chuẩn ERP!")
        st.dataframe(df_result, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, sheet_name='BOM_ERP', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['BOM_ERP']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            for col_num, val in enumerate(df_result.columns.values):
                worksheet.write(0, col_num, val, header_format)
                max_len = max(df_result[val].astype(str).map(len).max(), len(val)) + 3
                worksheet.set_column(col_num, col_num, min(max_len, 35))

        output.seek(0)
        st.download_button(
            label="📥 Tải xuống file BOM ERP (.xlsx)",
            data=output,
            file_name=f"BOM_ERP_{cust_code}_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Lỗi: {str(e)}")
