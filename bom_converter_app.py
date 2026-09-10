import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: EBOM to MBOM",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Chuyển Đổi BOM Kỹ Thuật (EBOM) Sang BOM ERP (MBOM)")
st.caption("Chuẩn hóa theo Quy trình sản xuất QTSX_01 - QTSX_10 & Công thức hao hụt chuẩn")

# Tham số Sidebar
with st.sidebar:
    st.header("🔧 Cấu hình")
    cust_code = st.text_input("Mã Khách hàng (KKK):", value="STV", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm (FG):", ["Bộ", "Cái"], index=0)
    st.markdown("---")
    st.markdown("""
    **Ma trận Routing tích hợp:**
    - **QTSX_01**: Cắt (`CUT`), Chấn/Uốn (`BD`)
    - **QTSX_02**: Hàn mài (`WD`), Phay tiện (`MC`)
    - **QTSX_03**: Sơn (`PNT`)
    - **QTSX_04**: Đóng gói (`PK`)
    - **QTSX_09**: Lắp ráp / Đóng PEM (`AS`)
    """)

def resolve_routing(row):
    """Xác định QTSX và Mã công đoạn dựa trên checkbox từ Cột 30 đến 38 của BOM Kỹ thuật"""
    has_cut = pd.notna(row[30])
    has_bend = pd.notna(row[31])
    has_weld = pd.notna(row[32])
    has_paint = pd.notna(row[33])
    has_plating = pd.notna(row[34])
    has_pack = pd.notna(row[35])
    has_mill = pd.notna(row[36])
    has_lathe = pd.notna(row[37])

    stages = []
    qtsx = ""
    
    if has_cut and has_bend:
        qtsx = "QTSX_01"
        stages.append("CUT, BD")
    elif has_cut:
        qtsx = "QTSX_05" if not has_bend else "QTSX_01"
        stages.append("CUT")
    elif has_bend:
        stages.append("BD")

    if has_weld and (has_mill or has_lathe):
        qtsx = "QTSX_02"
        stages.append("WD, MC")
    elif has_weld:
        stages.append("WD")
    elif has_mill or has_lathe:
        stages.append("MC")

    if has_paint:
        stages.append("PNT")
    if has_pack:
        stages.append("PK")

    return qtsx if qtsx else "QTSX_01", ", ".join(stages)

def parse_tech_bom(file_bytes, customer_code):
    df_raw = pd.read_excel(file_bytes, sheet_name='Form', header=None)
    
    proj_raw = str(df_raw.iloc[1, 5]).strip() if pd.notna(df_raw.iloc[1, 5]) else "PROJECT"
    pcode_match = re.search(r'([A-Za-z0-9\-]+)', proj_raw)
    pcode = pcode_match.group(1) if pcode_match else proj_raw
    fg_name = str(df_raw.iloc[6, 5]).strip() if pd.notna(df_raw.iloc[6, 5]) else proj_raw

    erp_rows = []
    
    # 1. Dòng Thành phẩm Cấp A (FG)
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

    current_sg_stt = 1
    current_sub_stt = 1
    
    r = 7
    while r < len(df_raw):
        row = df_raw.iloc[r]
        
        # Xác định Level thụt dòng (0: FG, 1: Cụm SG, 2: BTP/Part, 3: Phôi/PEM)
        level = None
        for l in range(4):
            if pd.notna(row[l]):
                level = l
                break
                
        if level is None and pd.isna(row[4]):
            r += 1
            continue

        part_no = str(row[4]).strip() if pd.notna(row[4]) else ""
        desc = str(row[5]).strip() if pd.notna(row[5]) else ""
        qty = row[17] if pd.notna(row[17]) else 1
        uom = str(row[25]).strip() if pd.notna(row[25]) else "Cái"
        mat_type = str(row[28]).strip().lower() if pd.notna(row[28]) else ""
        thickness = row[12]
        
        # Áp dụng công thức tính khối lượng & % hao hụt
        output_wt = float(row[21]) if pd.notna(row[21]) and str(row[21]).strip() != '' else 0.0
        tot_raw_wt = float(row[19]) if pd.notna(row[19]) and str(row[19]).strip() != '' else 0.0
        
        loss_pct = ""
        if output_wt > 0 and tot_raw_wt > 0:
            loss_pct = round(((tot_raw_wt - output_wt) / output_wt) * 100, 2)

        qtsx_code, stage_text = resolve_routing(row)

        # Cấp 1: Cụm lắp ráp chính (SG)
        if level == 1:
            erp_rows.append({
                'STT': str(current_sg_stt),
                'QTSX': '',
                'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                'Loại': 'BTP',
                'Tên sản phẩm': desc,
                'SL sản phẩm': qty,
                'ĐVT sp': 'Bộ',
                'Mã cũ': '', 'Mã NVL': '', 'Tên NVL': '', 'ĐVT (NVL)': '', 'Số lượng': '',
                'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                'Công đoạn': 'WD'
            })
            current_sub_stt = 1

        # Cấp 2: BTP trung gian hoặc Chi tiết đơn tấm
        elif level == 2:
            stt_str = f"{current_sg_stt}.{current_sub_stt}"
            has_child = (r + 1 < len(df_raw) and pd.notna(df_raw.iloc[r + 1, 3]))

            if has_child:
                # Cụm có lắp PEM / Chốt pin
                erp_rows.append({
                    'STT': stt_str,
                    'QTSX': '',
                    'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': desc if desc else f"CỤM {part_no}",
                    'SL sản phẩm': qty,
                    'ĐVT sp': 'Cái',
                    'Mã cũ': '', 'Mã NVL': '', 'Tên NVL': '', 'ĐVT (NVL)': '', 'Số lượng': '',
                    'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                    'Công đoạn': ''
                })
            else:
                # Chi tiết tấm đơn cắt/chấn -> QTSX_01
                erp_rows.append({
                    'STT': stt_str,
                    'QTSX': 'QTSX_01',
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
            current_sub_stt += 1

        # Cấp 3: Chi tiết phôi hoặc Fasteners (PEM, Bulong, Tán)
        elif level == 3:
            if "phụ" in mat_type or "fa" in part_no.lower():
                # Phụ kiện lắp ráp / ép PEM -> QTSX_09
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
                # Phôi tấm kim loại -> QTSX_01
                erp_rows.append({
                    'STT': '',
                    'QTSX': 'QTSX_01',
                    'Mã sản phẩm': f"PT - {customer_code} - {part_no}.01",
                    'Loại': '',
                    'Tên sản phẩm': '',
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

uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_result = parse_tech_bom(uploaded_file, cust_code)
        st.success(f"✅ Đã xử lý {len(df_result)} dòng BOM theo đúng quy chuẩn QTSX và Công thức hao hụt.")
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