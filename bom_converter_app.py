import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: Tech to ERP",
    page_icon="⚙️",
    layout="wide"
)

st.title("⚙️ Chuyển Đổi BOM Kỹ Thuật (EBOM) Sang BOM ERP (MBOM)")
st.caption("Công cụ chuẩn hóa định mức sản xuất cho ngành gia công & sản xuất cơ khí thép xuất khẩu")

# 1. Sidebar - Thiết lập cấu hình hệ thống
with st.sidebar:
    st.header("🔧 Tham số cấu hình")
    cust_code = st.text_input("Mã viết tắt Khách hàng (KKK):", value="STV", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm (FG):", ["Bộ", "Cái"], index=0)
    st.markdown("---")
    st.markdown("""
    **Quy tắc sinh mã tự động:**
    - Thành phẩm: `FG - [KKK] - [PCODE]`
    - Bán thành phẩm: `SG - [KKK] - [PCODE]`
    - Chi tiết gia công: `PT - [KKK] - [PCODE]`
    - Phôi NVL Thép: `Thép tấm [T]mm` (ĐVT: Kg)
    """)

# 2. Logic xử lý ETL
def parse_tech_bom(file_bytes, customer_code):
    df_raw = pd.read_excel(file_bytes, sheet_name='Form', header=None)
    
    # Bóc tách Dự án / Tên thành phẩm
    proj_raw = str(df_raw.iloc[1, 5]).strip() if pd.notna(df_raw.iloc[1, 5]) else "PROJECT"
    # Tìm mã PCODE từ chuỗi dự án (ví dụ: 750-108 _ WELLINGTON 5 -> 750-108)
    pcode_match = re.search(r'([A-Za-z0-9\-]+)', proj_raw)
    pcode = pcode_match.group(1) if pcode_match else proj_raw
    fg_name = str(df_raw.iloc[6, 5]).strip() if pd.notna(df_raw.iloc[6, 5]) else proj_raw

    erp_rows = []
    
    # 2.1 Dòng cấp A: Thành phẩm FG
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
    
    # 2.2 Duyệt các dòng chi tiết kỹ thuật
    r = 7
    while r < len(df_raw):
        row = df_raw.iloc[r]
        
        # Xác định Level dựa trên cột 0, 1, 2, 3
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
        raw_weight = row[19] if pd.notna(row[19]) else row[18]

        # Cấp 1: Cụm BTP chính (SG)
        if level == 1:
            erp_rows.append({
                'STT': str(current_sg_stt),
                'QTSX': '',
                'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                'Loại': 'BTP',
                'Tên sản phẩm': desc,
                'SL sản phẩm': qty if pd.notna(qty) else 1,
                'ĐVT sp': 'Bộ',
                'Mã cũ': '', 'Mã NVL': '', 'Tên NVL': '', 'ĐVT (NVL)': '', 'Số lượng': '',
                'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                'Công đoạn': 'ASSY'
            })
            current_sub_stt = 1
            
        # Cấp 2: BTP trung gian hoặc Chi tiết đơn
        elif level == 2:
            stt_str = f"{current_sg_stt}.{current_sub_stt}"
            # Kiểm tra xem dòng kế tiếp có phải là component con (Level 3) không
            has_child = False
            if r + 1 < len(df_raw) and pd.notna(df_raw.iloc[r + 1, 3]):
                has_child = True

            if has_child:
                # Cụm hàn hoặc cụm ép PEM
                erp_rows.append({
                    'STT': stt_str,
                    'QTSX': '',
                    'Mã sản phẩm': f"SG - {customer_code} - {part_no}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': desc if desc else f"CỤM CHI TIẾT {part_no}",
                    'SL sản phẩm': 1,
                    'ĐVT sp': 'Cái',
                    'Mã cũ': '', 'Mã NVL': '', 'Tên NVL': '', 'ĐVT (NVL)': '', 'Số lượng': '',
                    'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                    'Công đoạn': ''
                })
            else:
                # Chi tiết tấm đơn lẻ, xuất dòng QTSX_01
                erp_rows.append({
                    'STT': stt_str,
                    'QTSX': 'QTSX_01',
                    'Mã sản phẩm': f"PT - {customer_code} - {part_no}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': desc if desc else part_no,
                    'SL sản phẩm': qty,
                    'ĐVT sp': uom,
                    'Mã cũ': '',
                    'Mã NVL': '',
                    'Tên NVL': f"Thép tấm {thickness}mm" if pd.notna(thickness) else "Thép tấm",
                    'ĐVT (NVL)': 'Kg',
                    'Số lượng': round(float(raw_weight), 3) if pd.notna(raw_weight) else '',
                    'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                    'Công đoạn': 'CUT, BD'
                })
            current_sub_stt += 1

        # Cấp 3: Chi tiết phôi hoặc Vật tư phụ (PEM, Pin tán, Fasteners)
        elif level == 3:
            if "phụ" in mat_type or "fa" in part_no.lower():
                # Vật tư phụ / Fastener tiêu hao
                erp_rows.append({
                    'STT': '',
                    'QTSX': '',
                    'Mã sản phẩm': '',
                    'Loại': '',
                    'Tên sản phẩm': '',
                    'SL sản phẩm': qty,
                    'ĐVT sp': uom,
                    'Mã cũ': '',
                    'Mã NVL': part_no,
                    'Tên NVL': desc,
                    'ĐVT (NVL)': uom,
                    'Số lượng': qty,
                    'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                    'Công đoạn': 'ĐÓNG PEM'
                })
            else:
                # Phôi tấm kim loại chính
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
                    'Số lượng': round(float(raw_weight), 3) if pd.notna(raw_weight) else '',
                    'ĐVT Kho': '', 'Số lượng kho': '', 'Số lượng setup': '', '% Tỉ lệ hao hụt': '',
                    'Công đoạn': 'CUT, BD'
                })
        r += 1

    return pd.DataFrame(erp_rows)

# 3. Giao diện người dùng (UI)
uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_result = parse_tech_bom(uploaded_file, cust_code)
        
        st.success(f"✅ Đã chuyển đổi thành công! Tổng số dòng định mức: {len(df_result)}")
        
        # Bảng hiển thị preview
        st.subheader("📋 Dữ liệu BOM ERP Xem Trước (Preview)")
        st.dataframe(df_result, use_container_width=True)
        
        # Tạo file Excel để tải về
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, sheet_name='BOM_ERP', index=False)
            
            # Auto-fit column widths
            workbook  = writer.book
            worksheet = writer.sheets['BOM_ERP']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            for col_num, value in enumerate(df_result.columns.values):
                worksheet.write(0, col_num, value, header_format)
                max_len = max(df_result[value].astype(str).map(len).max(), len(value)) + 3
                worksheet.set_column(col_num, col_num, min(max_len, 35))
                
        output.seek(0)
        
        st.download_button(
            label="📥 Tải xuống file BOM ERP (.xlsx)",
            data=output,
            file_name=f"BOM_ERP_{cust_code}_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Lỗi xử lý file: {str(e)}")
        st.info("Vui lòng đảm bảo file kỹ thuật đúng mẫu sheet 'Form' với các cột Level A, B, C, D.")