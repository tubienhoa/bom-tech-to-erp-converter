import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: EBOM to ERP MBOM",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Chuyển Đổi BOM Kỹ Thuật Sang BOM ERP (Multi-Level Routing)")
st.caption("Tự động bóc tách phân tầng Cha - Con, sinh mã BTP từng công đoạn, Master Data NVL và Routing gia công ngoài")

# Sidebar cấu hình
with st.sidebar:
    st.header("🔧 Cấu hình hệ thống")
    cust_code = st.text_input("Mã Khách hàng (KKK):", value="SVC", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm (FG):", ["Cái", "Bộ"], index=0)
    mat_spec_default = st.text_input("Mã phân loại thép (NVL Code Prefix):", value="S01")
    st.markdown("---")
    st.markdown("""
    **Logic nghiệp vụ tích hợp:**
    - Cụm cha tự động sinh chuỗi BTP theo thứ tự công đoạn tick chọn (`_WELDED`, `_PLATED`, `_PAINTED`...)
    - Các chi tiết con liền kề được tính toán quy cách phôi, mã Master Data NVL (`R-P-...`, `R-H-...`, `R-B-...`)
    - Phân định rõ ràng **Công đoạn nội bộ** (`WD`, `PK`, `CUT, BD`) và **Gia công ngoài** (`PL`, `PNT`)
    """)

def get_raw_mat_info(item, mat_spec_code="S01"):
    """Tạo Tên NVL, Quy cách Phôi và Mã Master Data NVL"""
    desc = item['desc'].lower() if item['desc'] else ""
    T = item['T']
    W = item['W']
    W_B = item['W_B']
    L = item['L']
    
    if "hộp" in desc:
        t_int = int(round(float(T))) if pd.notna(T) else 3
        mat_name = f"Thép hộp {W:.0f}x{W_B:.0f}x{T:.0f}mm" if pd.notna(W) and pd.notna(W_B) and pd.notna(T) else f"Thép hộp {T}mm"
        phoi = f"{W:.0f} x {W_B:.0f} x {L:.0f}" if pd.notna(W) and pd.notna(W_B) and pd.notna(L) else f"{L}"
        mat_code = f"R-H-{mat_spec_code}-{t_int:02d}-00-000"
    elif "tròn đặc" in desc:
        w_int = int(round(float(W))) if pd.notna(W) else 32
        mat_name = f"Thép tròn đặc phi {W:.0f}mm" if pd.notna(W) else "Thép tròn đặc"
        phoi = f"dài {L:.0f}" if pd.notna(L) else ""
        mat_code = f"R-B-{mat_spec_code}-{w_int:02d}-00-000"
    elif "ống" in desc:
        t_int = int(round(float(T))) if pd.notna(T) else 3
        mat_name = f"Thép ống phi {W:.0f}x{T:.0f}mm" if pd.notna(W) and pd.notna(T) else "Thép ống"
        phoi = f"phi {W:.0f} x dài {L:.0f}" if pd.notna(W) and pd.notna(L) else ""
        mat_code = f"R-T-{mat_spec_code}-{t_int:02d}-00-000"
    else:  # Thép tấm mặc định
        t_val = float(T) if pd.notna(T) else 4.0
        t_str = f"{t_val:.0f}" if t_val.is_integer() else f"{t_val}"
        t_int = int(round(t_val))
        mat_name = f"Thép tấm {t_str}mm"
        phoi = f"{L} x {W}" if pd.notna(L) and pd.notna(W) else ""
        mat_code = f"R-P-{mat_spec_code}-{t_int:02d}-00-000"
        
    return mat_name, phoi, mat_code

def convert_tech_to_erp(file_bytes, customer_code="SVC", mat_spec_code="S01", default_uom="Cái"):
    df_raw = pd.read_excel(file_bytes, sheet_name='Form', header=None)
    
    # Bóc tách danh sách chi tiết từ BOM Kỹ thuật
    items = []
    op_cols = [
        (32, 'CUT', 'Cắt', 'QTSX_05', False),
        (33, 'BD', 'Chấn/uốn', 'QTSX_01', False),
        (34, 'WD', 'Hàn', 'QTSX_06', False),
        (35, 'MC', 'Phay tiện', 'QTSX_07', False),
        (36, 'FM', 'Dập tạo hình', 'QTSX_08', False),
        (37, 'PNT', 'Sơn', 'QTSX_03', True),    # Gia công ngoài
        (38, 'PL', 'Xi mạ', 'QTSX_10', True),   # Gia công ngoài
        (39, 'AS', 'Lắp ráp', 'QTSX_09', False),
        (40, 'PK', 'Đóng gói', 'QTSX_04', False)
    ]
    
    for r in range(6, len(df_raw)):
        level = None
        for c in range(5):
            val = df_raw.iloc[r, c]
            if pd.notna(val) and isinstance(val, (int, float)):
                level = c
                break
                
        part_no = str(df_raw.iloc[r, 5]).strip() if pd.notna(df_raw.iloc[r, 5]) else ""
        desc = str(df_raw.iloc[r, 6]).strip() if pd.notna(df_raw.iloc[r, 6]) else ""
        
        if level is None and not part_no:
            continue
            
        qty = df_raw.iloc[r, 18] if pd.notna(df_raw.iloc[r, 18]) else 1
        uom = str(df_raw.iloc[r, 26]).strip() if pd.notna(df_raw.iloc[r, 26]) else default_uom
        mat_type = str(df_raw.iloc[r, 29]).strip() if pd.notna(df_raw.iloc[r, 29]) else ""
        
        L = df_raw.iloc[r, 7]
        W = df_raw.iloc[r, 9]
        W_B = df_raw.iloc[r, 11]
        T = df_raw.iloc[r, 13]
        
        raw_wt = df_raw.iloc[r, 20] if pd.notna(df_raw.iloc[r, 20]) else df_raw.iloc[r, 19]
        out_wt = df_raw.iloc[r, 23] if pd.notna(df_raw.iloc[r, 23]) else df_raw.iloc[r, 22]
        
        active_ops = []
        for col_idx, code, name, qtsx, is_outside in op_cols:
            val = df_raw.iloc[r, col_idx]
            if pd.notna(val) and str(val).strip().lower() in ['x', 'sx', 'có', 'yes', '1']:
                active_ops.append({
                    'col': col_idx, 'code': code, 'name': name, 'qtsx': qtsx, 'outside': is_outside
                })
                
        items.append({
            'row': r, 'level': level, 'part_no': part_no, 'desc': desc,
            'qty': qty, 'uom': uom, 'mat_type': mat_type,
            'L': L, 'W': W, 'W_B': W_B, 'T': T,
            'raw_wt': raw_wt, 'out_wt': out_wt,
            'ops': active_ops
        })
        
    # Gom nhóm theo Cụm Cha (Level 0) và Chi tiết con (Level > 0)
    parents = []
    current_parent = None
    for it in items:
        if it['level'] == 0:
            current_parent = {'item': it, 'children': []}
            parents.append(current_parent)
        else:
            if current_parent is not None:
                current_parent['children'].append(it)
                
    rows = []
    for p_idx, p_data in enumerate(parents, 1):
        p_item = p_data['item']
        children = p_data['children']
        p_code = p_item['part_no']
        p_desc = p_item['desc']
        
        fg_code = f"FG - {customer_code} - {p_code}"
        
        # 1. Dòng Khởi tạo Thành phẩm (FG)
        rows.append({
            'STT': p_idx,
            'QTSX': None,
            'Mã sản phẩm': fg_code,
            'Loại': 'THÀNH PHẨM',
            'Tên sản phẩm': p_desc,
            'SL SP': 1,
            'ĐVT sp': default_uom,
            'Mã NVL': None,
            'Phôi (mm)': None,
            'Tên NVL': None,
            'ĐVT (NVL)': None,
            'Số lượng': None,
            '% Hao hụt': None,
            'Công đoạn': None,
            'Gia công ngoài': None,
            'BTP': None,
            'TP': 'x',
            'Số Kg/NVL': None
        })
        
        # 2. Dòng BTP Cụm hàn đầu tiên (WELDED)
        welded_part_code = f"SG - {customer_code} - {p_code}.01"
        welded_part_name = f"{p_desc}_WELDED"
        
        rows.append({
            'STT': f"{p_idx}.1",
            'QTSX': None,
            'Mã sản phẩm': welded_part_code,
            'Loại': 'BTP',
            'Tên sản phẩm': welded_part_name,
            'SL SP': 1,
            'ĐVT sp': default_uom,
            'Mã NVL': None,
            'Phôi (mm)': None,
            'Tên NVL': None,
            'ĐVT (NVL)': None,
            'Số lượng': None,
            '% Hao hụt': None,
            'Công đoạn': None,
            'Gia công ngoài': None,
            'BTP': 'x',
            'TP': None,
            'Số Kg/NVL': None
        })
        
        # 3. Liệt kê các chi tiết con liền kề (Tiêu hao NVL phôi)
        for child in children:
            c_code = child['part_no']
            c_desc = child['desc']
            c_raw_wt = child['raw_wt']
            
            mat_name, phoi, mat_code = get_raw_mat_info(child, mat_spec_code)
            
            has_cut = any(o['code'] == 'CUT' for o in child['ops'])
            has_bd = any(o['code'] == 'BD' for o in child['ops'])
            has_mc = any(o['code'] == 'MC' for o in child['ops'])
            
            if has_cut and has_bd:
                qtsx = "QTSX_01"
                cong_doan = "CUT, BD"
            elif has_cut:
                qtsx = "QTSX_05"
                cong_doan = "CUT" if not has_mc else "CUT "
            else:
                qtsx = "QTSX_01"
                cong_doan = "CUT"
                
            rows.append({
                'STT': None,
                'QTSX': qtsx,
                'Mã sản phẩm': f"SG - {customer_code} - {c_code}",
                'Loại': 'BTP',
                'Tên sản phẩm': c_code,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': mat_code,
                'Phôi (mm)': phoi,
                'Tên NVL': mat_name,
                'ĐVT (NVL)': 'Kg',
                'Số lượng': round(float(c_raw_wt), 2) if pd.notna(c_raw_wt) else None,
                '% Hao hụt': None,
                'Công đoạn': cong_doan,
                'Gia công ngoài': None,
                'BTP': 'x',
                'TP': None,
                'Số Kg/NVL': None
            })
            
        # 4. Định mức lắp ráp/hàn Cụm hàn WELDED (QTSX_06: WD)
        first_c = children[0] if children else None
        if first_c:
            rows.append({
                'STT': None,
                'QTSX': 'QTSX_06',
                'Mã sản phẩm': welded_part_code,
                'Loại': 'BTP',
                'Tên sản phẩm': welded_part_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': first_c['part_no'] if p_idx > 1 else None,
                'Phôi (mm)': None,
                'Tên NVL': first_c['part_no'],
                'ĐVT (NVL)': default_uom,
                'Số lượng': first_c['qty'],
                '% Hao hụt': None,
                'Công đoạn': 'WD',
                'Gia công ngoài': None,
                'BTP': None,
                'TP': None,
                'Số Kg/NVL': None
            })
            for c in children[1:]:
                rows.append({
                    'STT': None,
                    'QTSX': None,
                    'Mã sản phẩm': None,
                    'Loại': None,
                    'Tên sản phẩm': None,
                    'SL SP': None,
                    'ĐVT sp': None,
                    'Mã NVL': c['part_no'] if p_idx > 1 else None,
                    'Phôi (mm)': None,
                    'Tên NVL': c['part_no'],
                    'ĐVT (NVL)': default_uom,
                    'Số lượng': c['qty'],
                    '% Hao hụt': None,
                    'Công đoạn': None,
                    'Gia công ngoài': None,
                    'BTP': None,
                    'TP': None,
                    'Số Kg/NVL': None
                })
                
        # 5. Công đoạn Xi mạ (QTSX_10 / PL) nếu cha có tick mạ
        has_pl = any(o['code'] == 'PL' for o in p_item['ops'])
        plated_part_code = f"SG - {customer_code} - {p_code}.01.01"
        plated_part_name = f"{p_desc}_PLATED"
        
        if has_pl:
            rows.append({
                'STT': f"{p_idx}.2",
                'QTSX': 'QTSX_10',
                'Mã sản phẩm': plated_part_code,
                'Loại': 'BTP',
                'Tên sản phẩm': plated_part_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': f"{p_code}.01",
                'Phôi (mm)': None,
                'Tên NVL': welded_part_name,
                'ĐVT (NVL)': default_uom,
                'Số lượng': 1,
                '% Hao hụt': None,
                'Công đoạn': None,
                'Gia công ngoài': 'PL',
                'BTP': 'x' if p_idx > 1 else None,
                'TP': None,
                'Số Kg/NVL': None
            })
            
        # 6. Công đoạn Sơn nếu cha có tick sơn
        has_pnt = any(o['code'] == 'PNT' for o in p_item['ops'])
        if has_pnt:
            painted_part_code = f"SG - {customer_code} - {p_code}.01.02" if has_pl else f"SG - {customer_code} - {p_code}.01.01"
            painted_part_name = f"{p_desc}_PAINTED"
            prev_code = f"{p_code}.01.01" if has_pl else f"{p_code}.01"
            prev_name = plated_part_name if has_pl else welded_part_name
            rows.append({
                'STT': f"{p_idx}.3",
                'QTSX': 'QTSX_03',
                'Mã sản phẩm': painted_part_code,
                'Loại': 'BTP',
                'Tên sản phẩm': painted_part_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': prev_code,
                'Phôi (mm)': None,
                'Tên NVL': prev_name,
                'ĐVT (NVL)': default_uom,
                'Số lượng': 1,
                '% Hao hụt': None,
                'Công đoạn': None,
                'Gia công ngoài': 'PNT',
                'BTP': 'x',
                'TP': None,
                'Số Kg/NVL': None
            })
            
        # 7. Công đoạn Đóng gói / Hoàn thiện Thành phẩm (QTSX_04 / PK)
        has_pk = any(o['code'] == 'PK' for o in p_item['ops'])
        if has_pk:
            if has_pl:
                last_code = f"{p_code}.01.01"
                last_name = plated_part_name
            else:
                last_code = f"{p_code}.01"
                last_name = welded_part_name
                
            stt_pk = f"{p_idx}.3" if (p_idx > 1 and not has_pnt) else None
            rows.append({
                'STT': stt_pk,
                'QTSX': 'QTSX_04',
                'Mã sản phẩm': fg_code,
                'Loại': 'THÀNH PHẨM',
                'Tên sản phẩm': p_desc,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': last_code,
                'Phôi (mm)': None,
                'Tên NVL': last_name,
                'ĐVT (NVL)': default_uom,
                'Số lượng': 1,
                '% Hao hụt': None,
                'Công đoạn': 'PK',
                'Gia công ngoài': None,
                'BTP': None,
                'TP': None,
                'Số Kg/NVL': None
            })
            
    return pd.DataFrame(rows)

uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_result = convert_tech_to_erp(uploaded_file, cust_code, mat_spec_default, default_uom_fg)
        st.success(f"✅ Đã chuyển đổi thành công {len(df_result)} dòng BOM chuẩn ERP!")
        st.dataframe(df_result, use_container_width=True)
        
        # Xuất file Excel an toàn 100% không bị lỗi float length
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, sheet_name='B.O.M', index=False)
            ws = writer.sheets['B.O.M']
            # Tự động căn chỉnh độ rộng cột
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        if len(val_str) > max_len:
                            max_len = len(val_str)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

        output.seek(0)
        st.download_button(
            label="📥 Tải xuống file BOM ERP (.xlsx)",
            data=output,
            file_name=f"BOM_ERP_{cust_code}_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Lỗi: {str(e)}")
