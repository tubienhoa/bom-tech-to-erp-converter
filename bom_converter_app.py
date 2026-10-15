import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: EBOM to MBOM",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Chuyển Đổi BOM Kỹ Thuật Sang BOM ERP (Kiểm Tra Chặt Chẽ)")
st.caption("Tự động bóc tách phân cấp MỤC động, kiểm soát tính hợp lệ của Template và xuất dữ liệu MBOM chuẩn ERP")

# Sidebar cấu hình
with st.sidebar:
    st.header("🔧 Cấu hình hệ thống")
    cust_code = st.text_input("Mã Khách hàng (KKK):", value="STV", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm (FG):", ["Bộ", "Cái"], index=0)
    mat_spec_default = st.text_input("Mã phân loại thép (NVL Code Prefix):", value="S01")
    st.markdown("---")
    st.markdown("""
    **Cơ chế kiểm soát dữ liệu:**
    - Tự động nhận diện số cấp của cột `MỤC` (từ cột 0 đến trước `MÃ SỐ CHI TIẾT`).
    - Bắt buộc kiểm tra đúng tên và thứ tự các cột cố định.
    - Cảnh báo chi tiết từng dòng nếu thiếu dữ liệu, sai số lượng hoặc thiếu kích thước phôi.
    """)

def validate_and_find_columns(df_raw):
    """
    Kiểm tra tính hợp lệ của Template BOM:
    - Tìm vị trí cột MÃ SỐ CHI TIẾT
    - Vùng trước đó là MỤC (cấp thụt lề động)
    - Kiểm tra danh sách các cột cố định từ MÃ SỐ CHI TIẾT đến cuối
    """
    errors = []
    
    # 1. Tìm hàng tiêu đề (thông thường là hàng 4, tức dòng thứ 5 trong Excel)
    header_row_idx = None
    for r in range(min(10, len(df_raw))):
        row_str = [str(x).strip().upper() for x in df_raw.iloc[r] if pd.notna(x)]
        if any("MÃ SỐ CHI TIẾT" in x or "MÃ CHI TIẾT" in x for x in row_str):
            header_row_idx = r
            break
            
    if header_row_idx is None:
        errors.append("Không tìm thấy hàng tiêu đề chứa cột 'MÃ SỐ CHI TIẾT' trong 10 dòng đầu tiên.")
        return None, errors

    header_row = df_raw.iloc[header_row_idx]
    
    # 2. Định vị cột MÃ SỐ CHI TIẾT
    idx_part_no = None
    for c in range(len(header_row)):
        val = str(header_row[c]).strip().upper() if pd.notna(header_row[c]) else ""
        if "MÃ SỐ CHI TIẾT" in val or "MÃ CHI TIẾT" in val:
            idx_part_no = c
            break
            
    if idx_part_no is None:
        errors.append("Không tìm thấy cột 'MÃ SỐ CHI TIẾT'.")
        return None, errors

    # Vùng cột MỤC là từ 0 đến idx_part_no - 1
    num_level_cols = idx_part_no
    if num_level_cols < 1:
        errors.append(f"Cột 'MÃ SỐ CHI TIẾT' nằm ở vị trí đầu tiên (cột {idx_part_no}), không có vùng cột 'MỤC' phía trước.")

    # 3. Định vị và kiểm tra các cột cố định
    col_mapping = {'level_zone_end': idx_part_no, 'part_no': idx_part_no}
    
    required_keywords = [
        ('desc', ['MÔ TẢ'], "MÔ TẢ"),
        ('length', ['CHIỀU DÀI'], "CHIỀU DÀI"),
        ('width', ['CHIỀU RỘNG'], "CHIỀU RỘNG"),
        ('thickness', ['ĐỘ DÀY'], "ĐỘ DÀY"),
        ('qty', ['SỐ LƯỢNG'], "SỐ LƯỢNG"),
        ('tot_raw_wt', ['TỔNG KL NGUYÊN LIỆU', 'KL NGUYÊN LIỆU'], "TỔNG KL NGUYÊN LIỆU (kg)"),
        ('uom', ['ĐƠN VỊ TÍNH'], "ĐƠN VỊ TÍNH"),
        ('mat_type', ['LOẠI VẬT TƯ', 'LOẠI VẬT LIỆU'], "LOẠI VẬT TƯ")
    ]
    
    for key, kw_list, label in required_keywords:
        found_c = None
        for c in range(idx_part_no + 1, len(header_row)):
            val = str(header_row[c]).strip().upper() if pd.notna(header_row[c]) else ""
            if any(kw in val for kw in kw_list):
                found_c = c
                break
        if found_c is not None:
            col_mapping[key] = found_c
        else:
            errors.append(f"Thiếu cột bắt buộc: '{label}' phía sau cột MÃ SỐ CHI TIẾT.")

    # Cột chiều rộng phụ (W_B) cho thép hộp (nếu có)
    for c in range(idx_part_no + 1, len(header_row)):
        v5 = str(df_raw.iloc[header_row_idx + 1, c]).strip().lower() if len(df_raw) > header_row_idx + 1 and pd.notna(df_raw.iloc[header_row_idx + 1, c]) else ""
        if "mm(b)" in v5:
            col_mapping['width_b'] = c
            break

    # 4. Định vị các cột Quy trình sản xuất (Routing)
    routing_keys = [
        ('CUT', ['CẮT', '(CUT)'], 'QTSX_05', False),
        ('BD', ['CHẤN', 'UỐN', '(BD)'], 'QTSX_01', False),
        ('WD', ['HÀN', '(WD)'], 'QTSX_06', False),
        ('MC', ['PHAY', 'TIỆN', '(MC)'], 'QTSX_07', False),
        ('FM', ['DẬP', '(FM)'], 'QTSX_08', False),
        ('PNT', ['SƠN', '(PNT)'], 'QTSX_03', True),   # Gia công ngoài
        ('PL', ['XI MẠ', 'MẠ', '(PL)'], 'QTSX_10', True), # Gia công ngoài
        ('AS', ['LẮP RÁP', '(AS)'], 'QTSX_09', False),
        ('PK', ['ĐÓNG GÓI', '(PK)'], 'QTSX_04', False)
    ]
    
    col_mapping['routings'] = []
    for code, kws, qtsx, is_outside in routing_keys:
        for c in range(idx_part_no + 1, len(header_row)):
            val = str(header_row[c]).strip().upper() if pd.notna(header_row[c]) else ""
            if any(kw in val for kw in kws):
                col_mapping['routings'].append({
                    'col': c, 'code': code, 'qtsx': qtsx, 'outside': is_outside, 'name': val
                })
                break

    return {'header_row': header_row_idx, 'cols': col_mapping}, errors

def get_raw_mat_info(item, mat_spec_code="S01"):
    """Tạo Tên NVL, Quy cách Phôi và Mã Master Data NVL chuẩn"""
    desc = item['desc'].lower() if item['desc'] else ""
    T = item['T']
    W = item['W']
    W_B = item.get('W_B')
    L = item['L']
    
    if "hộp" in desc:
        t_int = int(round(float(T))) if (T is not None and str(T).strip() != '') else 3
        wb_val = W_B if W_B is not None else W
        mat_name = f"Thép hộp {W:.0f}x{wb_val:.0f}x{T:.0f}mm" if (W is not None and T is not None) else f"Thép hộp {T}mm"
        phoi = f"{W:.0f} x {wb_val:.0f} x {L:.0f}" if (W is not None and L is not None) else f"{L}"
        mat_code = f"R-H-{mat_spec_code}-{t_int:02d}-00-000"
    elif "tròn đặc" in desc:
        w_int = int(round(float(W))) if (W is not None and str(W).strip() != '') else 32
        mat_name = f"Thép tròn đặc phi {W:.0f}mm" if W is not None else "Thép tròn đặc"
        phoi = f"dài {L:.0f}" if L is not None else ""
        mat_code = f"R-B-{mat_spec_code}-{w_int:02d}-00-000"
    elif "ống" in desc:
        t_int = int(round(float(T))) if (T is not None and str(T).strip() != '') else 3
        mat_name = f"Thép ống phi {W:.0f}x{T:.0f}mm" if (W is not None and T is not None) else "Thép ống"
        phoi = f"phi {W:.0f} x dài {L:.0f}" if (W is not None and L is not None) else ""
        mat_code = f"R-T-{mat_spec_code}-{t_int:02d}-00-000"
    else:  # Thép tấm mặc định
        t_val = float(T) if (T is not None and str(T).strip() != '') else 4.0
        t_str = f"{t_val:.0f}" if t_val.is_integer() else f"{t_val}"
        t_int = int(round(t_val))
        mat_name = f"Thép tấm {t_str}mm"
        phoi = f"{L} x {W}" if (L is not None and W is not None) else ""
        mat_code = f"R-P-{mat_spec_code}-{t_int:02d}-00-000"
        
    return mat_name, phoi, mat_code

def parse_and_validate_bom(df_raw, cfg, customer_code="STV", mat_spec_code="S01", default_uom="Cái"):
    """Duyệt từng dòng dữ liệu và kiểm tra chặt chẽ các lỗi logic dữ liệu"""
    row_errors = []
    items = []
    
    header_idx = cfg['header_row']
    cols = cfg['cols']
    num_level_cols = cols['level_zone_end']
    
    start_row = header_idx + 2  # Dòng bắt đầu dữ liệu sau hàng sub-header đơn vị mm/inch
    
    for r in range(start_row, len(df_raw)):
        row = df_raw.iloc[r]
        
        # 1. Xác định cấp độ Level từ vùng cột MỤC (từ 0 đến num_level_cols - 1)
        found_levels = []
        for c in range(num_level_cols):
            val = row[c]
            if pd.notna(val) and isinstance(val, (int, float)):
                found_levels.append(c)
                
        part_no = str(row[cols['part_no']]).strip() if pd.notna(row[cols['part_no']]) else ""
        desc = str(row[cols['desc']]).strip() if pd.notna(row[cols['desc']]) else ""
        
        # Nếu cả hàng trống hoàn toàn thì bỏ qua
        if not found_levels and not part_no and not desc:
            continue
            
        # Kiểm tra tính hợp lệ của Level
        if len(found_levels) == 0:
            row_errors.append(f"Dòng Excel {r + 1}: Không xác định được cấp độ MỤC (các cột MỤC từ A đến trước MÃ SỐ CHI TIẾT đều trống).")
            continue
        elif len(found_levels) > 1:
            row_errors.append(f"Dòng Excel {r + 1}: Có {len(found_levels)} cấp độ MỤC được điền cùng lúc trên một dòng.")
            continue
            
        level = found_levels[0]
        
        if not part_no and not desc:
            row_errors.append(f"Dòng Excel {r + 1}: Thiếu cả 'MÃ SỐ CHI TIẾT' và 'MÔ TẢ'.")
            continue

        # Kiểm tra Số lượng
        qty_val = row[cols['qty']]
        if pd.isna(qty_val):
            qty = 1
        else:
            try:
                qty = float(qty_val)
                if qty <= 0:
                    row_errors.append(f"Dòng Excel {r + 1} [{part_no or desc}]: Số lượng '{qty_val}' phải lớn hơn 0.")
            except:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no or desc}]: Số lượng '{qty_val}' không phải là định dạng số.")
                qty = 1

        # Lấy thông tin kích thước và khối lượng
        L = row[cols['length']] if cols.get('length') and pd.notna(row[cols['length']]) else None
        W = row[cols['width']] if cols.get('width') and pd.notna(row[cols['width']]) else None
        T = row[cols['thickness']] if cols.get('thickness') and pd.notna(row[cols['thickness']]) else None
        W_B = row[cols['width_b']] if cols.get('width_b') and pd.notna(row[cols['width_b']]) else None
        
        tot_raw_wt = row[cols['tot_raw_wt']] if cols.get('tot_raw_wt') and pd.notna(row[cols['tot_raw_wt']]) else None
        uom = str(row[cols['uom']]).strip() if cols.get('uom') and pd.notna(row[cols['uom']]) else default_uom
        mat_type = str(row[cols['mat_type']]).strip() if cols.get('mat_type') and pd.notna(row[cols['mat_type']]) else ""

        # Kiểm tra kích thước phôi với chi tiết tấm
        if "thép tấm" in desc.lower() and level > 0:
            if T is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}]: Là Thép tấm nhưng thiếu giá trị tại cột 'ĐỘ DÀY'.")
            if tot_raw_wt is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}]: Thiếu 'TỔNG KL NGUYÊN LIỆU (kg)'.")

        # Lấy các công đoạn và thứ tự thực hiện (1, 2, 3...)
        active_ops = []
        for r_info in cols['routings']:
            c = r_info['col']
            if c < len(row):
                val = row[c]
                if pd.notna(val) and str(val).strip() != '':
                    order = 99
                    try:
                        order = float(val)
                    except:
                        order = 1
                    active_ops.append({
                        'code': r_info['code'],
                        'qtsx': r_info['qtsx'],
                        'outside': r_info['outside'],
                        'order': order
                    })
        # Sắp xếp theo thứ tự công đoạn 1, 2, 3...
        active_ops.sort(key=lambda x: x['order'])

        items.append({
            'excel_row': r + 1,
            'level': level,
            'part_no': part_no,
            'desc': desc,
            'qty': qty,
            'uom': uom,
            'L': L, 'W': W, 'W_B': W_B, 'T': T,
            'tot_raw_wt': tot_raw_wt,
            'mat_type': mat_type,
            'ops': active_ops
        })
        
    return items, row_errors

def build_erp_mbom_multi_level(items, customer_code="STV", mat_spec_code="S01", default_uom="Cái"):
    """Chuyển đổi cây đa tầng (bất kể số cấp) sang MBOM ERP hoàn chỉnh"""
    if not items:
        return pd.DataFrame()

    rows = []
    
    # 1. Tìm các cụm cha cấp cao nhất (Level nhỏ nhất trong danh sách, thường là 0)
    min_level = min(it['level'] for it in items)
    
    # Phân nhóm cây Cha - Con
    # Một Cụm Cha bao gồm chính nó và tất cả các chi tiết con liền kề có Level > min_level
    parents = []
    curr_parent = None
    for it in items:
        if it['level'] == min_level:
            curr_parent = {'item': it, 'children': []}
            parents.append(curr_parent)
        else:
            if curr_parent is not None:
                curr_parent['children'].append(it)

    for p_idx, p_data in enumerate(parents, 1):
        p_item = p_data['item']
        children = p_data['children']
        p_code = p_item['part_no'] if p_item['part_no'] else f"PART_{p_idx}"
        p_desc = p_item['desc'] if p_item['desc'] else p_code
        
        fg_code = f"FG - {customer_code} - {p_code}"
        
        # Dòng 1: Thành phẩm hoàn chỉnh (FG)
        rows.append({
            'STT': p_idx,
            'QTSX': None,
            'Mã sản phẩm': fg_code,
            'Loại': 'THÀNH PHẨM',
            'Tên sản phẩm': p_desc,
            'SL SP': 1,
            'ĐVT sp': default_uom,
            'Mã NVL': None, 'Phôi (mm)': None, 'Tên NVL': None, 'ĐVT (NVL)': None, 'Số lượng': None,
            '% Hao hụt': None, 'Công đoạn': None, 'Gia công ngoài': None, 'BTP': None, 'TP': 'x', 'Số Kg/NVL': None
        })
        
        # Tách các công đoạn của Cụm Cha (Hàn WD, Xi mạ PL, Sơn PNT, Đóng gói PK)
        p_ops = p_item['ops']
        has_weld = any(o['code'] == 'WD' for o in p_ops)
        has_pl   = any(o['code'] == 'PL' for o in p_ops)
        has_pnt  = any(o['code'] == 'PNT' for o in p_ops)
        has_pk   = any(o['code'] == 'PK' for o in p_ops)
        
        welded_part_code = f"SG - {customer_code} - {p_code}.01"
        welded_part_name = f"{p_desc}_WELDED"
        
        # Dòng BTP Cụm hàn đầu tiên nếu có hàn
        if has_weld:
            rows.append({
                'STT': f"{p_idx}.1",
                'QTSX': None,
                'Mã sản phẩm': welded_part_code,
                'Loại': 'BTP',
                'Tên sản phẩm': welded_part_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': None, 'Phôi (mm)': None, 'Tên NVL': None, 'ĐVT (NVL)': None, 'Số lượng': None,
                '% Hao hụt': None, 'Công đoạn': None, 'Gia công ngoài': None, 'BTP': 'x', 'TP': None, 'Số Kg/NVL': None
            })
            
        # Liệt kê các chi tiết con cấu thành
        for child in children:
            c_code = child['part_no']
            c_desc = child['desc']
            c_wt   = child['tot_raw_wt']
            
            # Nếu là chi tiết con cấp dưới cùng dùng phôi NVL
            if child['tot_raw_wt'] is not None or "thép" in c_desc.lower():
                mat_name, phoi, mat_code = get_raw_mat_info(child, mat_spec_code)
                
                # Routing công đoạn của chi tiết con
                c_ops = [o['code'] for o in child['ops']]
                if 'CUT' in c_ops and 'BD' in c_ops:
                    qtsx = "QTSX_01"
                    c_stage = "CUT, BD"
                elif 'CUT' in c_ops:
                    qtsx = "QTSX_05"
                    c_stage = "CUT"
                else:
                    qtsx = "QTSX_01"
                    c_stage = "CUT"
                    
                rows.append({
                    'STT': None,
                    'QTSX': qtsx,
                    'Mã sản phẩm': f"SG - {customer_code} - {c_code}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': c_code if c_code else c_desc,
                    'SL SP': 1,
                    'ĐVT sp': default_uom,
                    'Mã NVL': mat_code,
                    'Phôi (mm)': phoi,
                    'Tên NVL': mat_name,
                    'ĐVT (NVL)': 'Kg',
                    'Số lượng': round(float(c_wt), 2) if (c_wt is not None and str(c_wt).strip() != '') else None,
                    '% Hao hụt': None,
                    'Công đoạn': c_stage,
                    'Gia công ngoài': None,
                    'BTP': 'x',
                    'TP': None,
                    'Số Kg/NVL': None
                })
            else:
                # Phụ kiện (PEM, Pin, Bulong)
                rows.append({
                    'STT': None,
                    'QTSX': None,
                    'Mã sản phẩm': f"PT - {customer_code} - {c_code}",
                    'Loại': 'BTP',
                    'Tên sản phẩm': c_desc,
                    'SL SP': child['qty'],
                    'ĐVT sp': child['uom'],
                    'Mã NVL': c_code,
                    'Phôi (mm)': None,
                    'Tên NVL': c_desc,
                    'ĐVT (NVL)': child['uom'],
                    'Số lượng': child['qty'],
                    '% Hao hụt': None,
                    'Công đoạn': 'ĐÓNG PEM',
                    'Gia công ngoài': None,
                    'BTP': None,
                    'TP': None,
                    'Số Kg/NVL': None
                })

        # Định mức gá tổ hợp hàn (WD) nếu có
        if has_weld and children:
            first_c = children[0]
            rows.append({
                'STT': None,
                'QTSX': 'QTSX_06',
                'Mã sản phẩm': welded_part_code,
                'Loại': 'BTP',
                'Tên sản phẩm': welded_part_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': first_c['part_no'],
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
                    'STT': None, 'QTSX': None, 'Mã sản phẩm': None, 'Loại': None, 'Tên sản phẩm': None,
                    'SL SP': None, 'ĐVT sp': None, 'Mã NVL': c['part_no'], 'Phôi (mm)': None,
                    'Tên NVL': c['part_no'], 'ĐVT (NVL)': default_uom, 'Số lượng': c['qty'],
                    '% Hao hụt': None, 'Công đoạn': None, 'Gia công ngoài': None, 'BTP': None, 'TP': None, 'Số Kg/NVL': None
                })

        # Công đoạn Xi mạ (QTSX_10 / PL)
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
                'BTP': 'x',
                'TP': None,
                'Số Kg/NVL': None
            })

        # Công đoạn Sơn (QTSX_03 / PNT)
        if has_pnt:
            pnt_code = f"SG - {customer_code} - {p_code}.01.02" if has_pl else f"SG - {customer_code} - {p_code}.01.01"
            pnt_name = f"{p_desc}_PAINTED"
            prev_code = f"{p_code}.01.01" if has_pl else f"{p_code}.01"
            prev_name = plated_part_name if has_pl else welded_part_name
            rows.append({
                'STT': f"{p_idx}.3",
                'QTSX': 'QTSX_03',
                'Mã sản phẩm': pnt_code,
                'Loại': 'BTP',
                'Tên sản phẩm': pnt_name,
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

        # Hoàn thiện Đóng gói Thành phẩm (QTSX_04 / PK)
        if has_pk:
            last_code = plated_part_code if has_pl else welded_part_code
            last_name = plated_part_name if has_pl else welded_part_name
            rows.append({
                'STT': f"{p_idx}.3" if (p_idx > 1 and not has_pnt) else None,
                'QTSX': 'QTSX_04',
                'Mã sản phẩm': fg_code,
                'Loại': 'THÀNH PHẨM',
                'Tên sản phẩm': p_desc,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': last_code.replace(f"SG - {customer_code} - ", ""),
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

# Giao diện người dùng
uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name='Form', header=None)
        
        # 1. Kiểm tra cấu trúc cột Template
        cfg, tpl_errors = validate_and_find_columns(df_raw)
        
        if tpl_errors:
            st.error("🚨 PHÁT HIỆN LỖI CẤU TRÚC TEMPLATE:")
            for err in tpl_errors:
                st.markdown(f"- ❌ {err}")
        else:
            # 2. Kiểm tra chi tiết dữ liệu từng dòng
            items, data_errors = parse_and_validate_bom(df_raw, cfg, cust_code, mat_spec_default, default_uom_fg)
            
            if data_errors:
                st.warning(f"⚠️ Phát hiện {len(data_errors)} cảnh báo dữ liệu cần lưu ý:")
                with st.expander("Bấm để xem danh sách lỗi chi tiết từng dòng", expanded=True):
                    for derr in data_errors[:20]:
                        st.markdown(f"- {derr}")
                    if len(data_errors) > 20:
                        st.markdown(f"*... và còn {len(data_errors) - 20} cảnh báo khác.*")
                        
            # 3. Tạo dữ liệu BOM ERP
            df_result = build_erp_mbom_multi_level(items, cust_code, mat_spec_default, default_uom_fg)
            
            if len(df_result) > 0:
                st.success(f"✅ Đã phân tích và chuyển đổi thành công {len(df_result)} dòng BOM ERP!")
                st.dataframe(df_result, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_result.to_excel(writer, sheet_name='B.O.M', index=False)
                    ws = writer.sheets['B.O.M']
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
            else:
                st.error("Không có dòng dữ liệu hợp lệ nào được tìm thấy để chuyển đổi.")
                
    except Exception as e:
        st.error(f"❌ Lỗi xử lý hệ thống: {str(e)}")
