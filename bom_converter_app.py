import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: EBOM to MBOM",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Chuyển Đổi BOM Kỹ Thuật Sang BOM ERP (Kiểm Tra Chặt Chẽ 100%)")
st.caption("Tự động nhận diện cấp MỤC linh hoạt & Kiểm tra đối soát tuần tự 35 cột cố định bằng thuật toán LCS")

# Sidebar cấu hình
with st.sidebar:
    st.header("🔧 Cấu hình hệ thống")
    cust_code = st.text_input("Mã Khách hàng (KKK):", value="STV", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm (FG):", ["Bộ", "Cái"], index=0)
    mat_spec_default = st.text_input("Mã phân loại thép (NVL Code Prefix):", value="S01")
    st.markdown("---")
    st.markdown("""
    **Cơ chế kiểm soát Template & Dữ liệu:**
    - **Vùng cột `MỤC`**: Tự động co giãn linh hoạt theo số cấp thụt lề thực tế (từ cột A đến trước `MÃ SỐ CHI TIẾT`).
    - **Khóa cứng 35 cột cố định**: Đối soát 1-1 từng cột từ `MÃ SỐ CHI TIẾT` đến `Đóng gói (PK)`. Bất kỳ cột nào bị xóa, chèn thêm, đảo vị trí hoặc đổi tên đều bị chỉ đích danh.
    - **Kiểm tra dữ liệu dòng**: Cảnh báo chính xác tọa độ dòng/cột bị thiếu thông số hoặc sai định dạng.
    """)

# Danh mục chuẩn 35 cột cố định tính từ MÃ SỐ CHI TIẾT (Offset +0 đến +34)
EXPECTED_FIXED_COLS = [
    (0, "MÃ SỐ CHI TIẾT", "", ["MÃ SỐ CHI TIẾT", "MÃ CHI TIẾT"]),
    (1, "MÔ TẢ", "", ["MÔ TẢ"]),
    (2, "CHIỀU DÀI (mm)", "mm", ["CHIỀU DÀI"]),
    (3, "CHIỀU DÀI (inch)", "inch", ["INCH"]),
    (4, "CHIỀU RỘNG mm(A)", "mm(A)", ["CHIỀU RỘNG"]),
    (5, "CHIỀU RỘNG inch(A)", "inch(A)", ["INCH(A)"]),
    (6, "CHIỀU RỘNG mm(B)", "mm(B)", ["MM(B)"]),
    (7, "CHIỀU RỘNG inch(B)", "inch(B)", ["INCH(B)"]),
    (8, "ĐỘ DÀY (mm)", "mm", ["ĐỘ DÀY"]),
    (9, "ĐỘ DÀY (inch)", "inch", ["INCH"]),
    (10, "LOẠI VẬT LIỆU (Nhóm)", "", ["LOẠI VẬT LIỆU"]),
    (11, "KHỐI LƯỢNG RIÊNG (g/cm3)", "", ["KHỐI LƯỢNG RIÊNG"]),
    (12, "TỈ LỆ HAO HỤT", "", ["TỈ LỆ HAO HỤT", "TỶ LỆ HAO HỤT"]),
    (13, "SỐ LƯỢNG", "", ["SỐ LƯỢNG"]),
    (14, "KHỐI LƯỢNG NGUYÊN LIỆU (kg)", "", ["KHỐI LƯỢNG NGUYÊN LIỆU"]),
    (15, "TỔNG KL NGUYÊN LIỆU (kg)", "", ["TỔNG KL NGUYÊN LIỆU"]),
    (16, "KHỐI LƯỢNG THÀNH PHẨM (lbs)", "", ["KHỐI LƯỢNG THÀNH PHẨM (LBS)", "(LBS)"]),
    (17, "KHỐI LƯỢNG THÀNH PHẨM (kg)", "", ["KHỐI LƯỢNG THÀNH PHẨM (KG)"]),
    (18, "TỔNG KL THÀNH PHẨM (kg)", "", ["TỔNG KL THÀNH PHẨM"]),
    (19, "DIỆN TÍCH BỀ MẶT SƠN (m2)", "", ["DIỆN TÍCH BỀ MẶT SƠN"]),
    (20, "TỔNG DIỆN TÍCH BỀ MẶT SƠN (m2)", "", ["TỔNG DIỆN TÍCH BỀ MẶT SƠN"]),
    (21, "ĐƠN VỊ TÍNH", "", ["ĐƠN VỊ TÍNH"]),
    (22, "GHI CHÚ", "", ["GHI CHÚ"]),
    (23, "HOÀN THIỆN BỀ MẶT", "", ["HOÀN THIỆN BỀ MẶT"]),
    (24, "LOẠI VẬT TƯ / VẬT LIỆU", "", ["LOẠI VẬT TƯ", "LOẠI VẬT LIỆU"]),
    (25, "MÁC VẬT TƯ", "", ["MÁC VẬT TƯ"]),
    (26, "Cắt (CUT)", "", ["CẮT", "(CUT)"]),
    (27, "Chấn/uốn (BD)", "", ["CHẤN", "UỐN", "(BD)"]),
    (28, "Hàn (WD)", "", ["HÀN", "(WD)"]),
    (29, "Phay tiện (MC)", "", ["PHAY", "TIỆN", "(MC)"]),
    (30, "Dập tạo hình (FM)", "", ["DẬP", "(FM)"]),
    (31, "Sơn (PNT)", "", ["SƠN", "(PNT)"]),
    (32, "Xi mạ (PL)", "", ["XI MẠ", "(PL)"]),
    (33, "Lắp ráp (AS)", "", ["LẮP RÁP", "(AS)"]),
    (34, "Đóng gói (PK)", "", ["ĐÓNG GÓI", "(PK)"])
]

def clean_header_str(val):
    """Chuẩn hóa chuỗi tiêu đề, loại bỏ ký tự xuống dòng \n, \r và khoảng trắng thừa"""
    if pd.isna(val):
        return ""
    return re.sub(r'[\r\n\s]+', ' ', str(val)).strip()

def excel_col_name(n):
    """Đổi chỉ số cột (0-indexed) sang ký hiệu cột Excel (A, B, ..., Z, AA, AB...)"""
    name = ""
    while n >= 0:
        name = chr(n % 26 + 65) + name
        n = n // 26 - 1
    return name

def match_col_spec(h_main, h_sub, spec):
    """Kiểm tra tiêu đề chính và tiêu đề phụ (đơn vị mm/inch) có khớp với đặc tả cột thứ offset hay không"""
    offset, label, req_sub, kw_list = spec
    h_main_up = h_main.upper()
    h_sub_up = h_sub.upper()
    
    if req_sub:
        if req_sub.upper() != h_sub_up:
            return False
        if h_main_up and not any(k in h_main_up for k in kw_list):
            return False
        return True
    else:
        # Phân biệt chặt các cột có tên gần giống nhau
        if offset == 15 and "TỔNG" not in h_main_up:
            return False
        if offset == 14 and "TỔNG" in h_main_up:
            return False
        if offset == 18 and "TỔNG" not in h_main_up:
            return False
        if offset == 17 and ("TỔNG" in h_main_up or "LBS" in h_main_up):
            return False
        if offset == 16 and "LBS" not in h_main_up:
            return False
        if offset == 20 and "TỔNG" not in h_main_up:
            return False
        if offset == 19 and "TỔNG" in h_main_up:
            return False
        if offset == 31 and "DIỆN TÍCH" in h_main_up:
            return False
        return any(k in h_main_up for k in kw_list)

def validate_and_find_columns(df_raw):
    """
    Thuật toán LCS (Longest Common Subsequence) đối soát chặt chẽ toàn bộ 35 cột cố định:
    - Tự động phát hiện chính xác bất kỳ cột nào bị xóa, bị đổi tên, đảo vị trí hoặc chèn thêm.
    """
    errors = []
    
    # 1. Tìm dòng tiêu đề chứa MÃ SỐ CHI TIẾT
    header_idx = None
    for r in range(min(20, len(df_raw))):
        row_vals = [clean_header_str(x).upper() for x in df_raw.iloc[r] if pd.notna(x)]
        if any("MÃ SỐ CHI TIẾT" in x or "MÃ CHI TIẾT" in x for x in row_vals):
            header_idx = r
            break
            
    if header_idx is None:
        return None, ["Không tìm thấy dòng tiêu đề chứa cột **'MÃ SỐ CHI TIẾT'** trong 20 dòng đầu tiên của sheet 'Form'."]
        
    # 2. Định vị cột MÃ SỐ CHI TIẾT
    idx_part = None
    for c in range(df_raw.shape[1]):
        if any(k in clean_header_str(df_raw.iloc[header_idx, c]).upper() for k in ["MÃ SỐ CHI TIẾT", "MÃ CHI TIẾT"]):
            idx_part = c
            break
            
    if idx_part is None or idx_part < 1:
        return None, ["Không tìm thấy cột **'MÃ SỐ CHI TIẾT'** hoặc thiếu vùng cột phân cấp **'MỤC'** phía trước."]
        
    muc_header = clean_header_str(df_raw.iloc[header_idx, 0]).upper()
    if "MỤC" not in muc_header:
        errors.append(f"Cột A (Dòng {header_idx + 1}) phải là tiêu đề **'MỤC'**, hiện tại đang là **'{clean_header_str(df_raw.iloc[header_idx, 0])}'**.")

    # 3. Thu thập danh sách các cột thực tế từ MÃ SỐ CHI TIẾT đến hết
    actual_cols = []
    for c in range(idx_part, df_raw.shape[1]):
        hm = clean_header_str(df_raw.iloc[header_idx, c])
        hs = clean_header_str(df_raw.iloc[header_idx + 1, c]) if header_idx + 1 < len(df_raw) else ""
        actual_cols.append((c, hm, hs))
        
    # 4. Đối soát chuỗi 35 cột cố định bằng Quy hoạch động (LCS)
    M = len(EXPECTED_FIXED_COLS)
    N = len(actual_cols)
    dp = [[0] * (N + 1) for _ in range(M + 1)]
    
    for i in range(M - 1, -1, -1):
        for j in range(N - 1, -1, -1):
            if match_col_spec(actual_cols[j][1], actual_cols[j][2], EXPECTED_FIXED_COLS[i]):
                dp[i][j] = 1 + dp[i + 1][j + 1]
            else:
                dp[i][j] = max(dp[i + 1][j], dp[i][j + 1])
                
    i, j = 0, 0
    missing_specs = []
    extra_actuals = []
    
    while i < M and j < N:
        if match_col_spec(actual_cols[j][1], actual_cols[j][2], EXPECTED_FIXED_COLS[i]):
            i += 1
            j += 1
        elif dp[i + 1][j] >= dp[i][j + 1]:
            exp_col_letter = excel_col_name(idx_part + EXPECTED_FIXED_COLS[i][0])
            missing_specs.append((EXPECTED_FIXED_COLS[i][0], EXPECTED_FIXED_COLS[i][1], exp_col_letter))
            i += 1
        else:
            col_let = excel_col_name(actual_cols[j][0])
            extra_actuals.append((col_let, actual_cols[j][1] or actual_cols[j][2]))
            j += 1
            
    while i < M:
        exp_col_letter = excel_col_name(idx_part + EXPECTED_FIXED_COLS[i][0])
        missing_specs.append((EXPECTED_FIXED_COLS[i][0], EXPECTED_FIXED_COLS[i][1], exp_col_letter))
        i += 1
    while j < N:
        col_let = excel_col_name(actual_cols[j][0])
        extra_actuals.append((col_let, actual_cols[j][1] or actual_cols[j][2]))
        j += 1
        
    for offset, label, exp_col_letter in missing_specs:
        errors.append(f"Bị xóa hoặc sai tên cột bắt buộc: **'{label}'** (Vị trí chuẩn: Cột **{exp_col_letter}** - đứng thứ {offset + 1} tính từ MÃ SỐ CHI TIẾT).")
        
    for col_let, val_name in extra_actuals:
        errors.append(f"Phát hiện cột lạ / sai vị trí chuẩn tại Cột **{col_let}**: **'{val_name}'**.")
        
    if errors:
        return None, errors

    # Khi toàn bộ 35 cột cố định khớp 100%, ánh xạ chính xác tuyệt đối theo offset chuẩn
    col_mapping = {
        'level_zone_end': idx_part,
        'part_no': idx_part + 0,
        'desc': idx_part + 1,
        'length': idx_part + 2,
        'width': idx_part + 4,
        'width_b': idx_part + 6,
        'thickness': idx_part + 8,
        'mat_group': idx_part + 10,
        'density': idx_part + 11,
        'loss_rate': idx_part + 12,
        'qty': idx_part + 13,
        'unit_raw_wt': idx_part + 14,
        'tot_raw_wt': idx_part + 15,
        'uom': idx_part + 21,
        'note': idx_part + 22,
        'surface': idx_part + 23,
        'mat_type': idx_part + 24,
        'mat_spec': idx_part + 25,
        'routings': [
            {'col': idx_part + 26, 'code': 'CUT', 'qtsx': 'QTSX_05', 'outside': False, 'name': 'Cắt (CUT)'},
            {'col': idx_part + 27, 'code': 'BD',  'qtsx': 'QTSX_01', 'outside': False, 'name': 'Chấn/uốn (BD)'},
            {'col': idx_part + 28, 'code': 'WD',  'qtsx': 'QTSX_06', 'outside': False, 'name': 'Hàn (WD)'},
            {'col': idx_part + 29, 'code': 'MC',  'qtsx': 'QTSX_07', 'outside': False, 'name': 'Phay tiện (MC)'},
            {'col': idx_part + 30, 'code': 'FM',  'qtsx': 'QTSX_08', 'outside': False, 'name': 'Dập tạo hình (FM)'},
            {'col': idx_part + 31, 'code': 'PNT', 'qtsx': 'QTSX_03', 'outside': True,  'name': 'Sơn (PNT)'},
            {'col': idx_part + 32, 'code': 'PL',  'qtsx': 'QTSX_10', 'outside': True,  'name': 'Xi mạ (PL)'},
            {'col': idx_part + 33, 'code': 'AS',  'qtsx': 'QTSX_09', 'outside': False, 'name': 'Lắp ráp (AS)'},
            {'col': idx_part + 34, 'code': 'PK',  'qtsx': 'QTSX_04', 'outside': False, 'name': 'Đóng gói (PK)'}
        ]
    }
    return {'header_row': header_idx, 'cols': col_mapping}, []

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
    
    start_row = header_idx + 2  # Dòng bắt đầu dữ liệu sau hàng sub-header mm/inch
    
    for r in range(start_row, len(df_raw)):
        row = df_raw.iloc[r]
        
        # Dừng nếu gặp dòng TOTAL ở cuối bảng BOM
        first_cell_str = str(row[0]).strip().upper() if pd.notna(row[0]) else ""
        if "TOTAL" in first_cell_str or "TÍNH KHỐI LƯỢNG" in str(row[cols['unit_raw_wt']]).upper():
            break

        # 1. Xác định cấp độ Level từ vùng cột MỤC (từ 0 đến num_level_cols - 1)
        found_levels = []
        for c in range(num_level_cols):
            val = row[c]
            if pd.notna(val) and str(val).strip() != '':
                if isinstance(val, (int, float)):
                    found_levels.append(c)
                else:
                    col_let = excel_col_name(c)
                    row_errors.append(f"Dòng Excel {r + 1} (Cột {col_let}): Giá trị cấp MỤC '{val}' không hợp lệ (phải là số).")
                
        part_no = str(row[cols['part_no']]).strip() if pd.notna(row[cols['part_no']]) else ""
        desc = str(row[cols['desc']]).strip() if pd.notna(row[cols['desc']]) else ""
        
        # Nếu cả hàng trống hoàn toàn thì bỏ qua
        if not found_levels and not part_no and not desc:
            continue
            
        # Kiểm tra tính hợp lệ của Level
        if len(found_levels) == 0:
            row_errors.append(f"Dòng Excel {r + 1} [{part_no or desc}]: Chưa điền số thứ tự cấp MỤC (từ Cột A đến Cột {excel_col_name(num_level_cols - 1)} đều trống).")
            continue
        elif len(found_levels) > 1:
            cols_str = ", ".join(excel_col_name(c) for c in found_levels)
            row_errors.append(f"Dòng Excel {r + 1} [{part_no or desc}]: Điền trùng lặp nhiều cột MỤC cùng lúc tại các cột ({cols_str}).")
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
                    col_let = excel_col_name(cols['qty'])
                    row_errors.append(f"Dòng Excel {r + 1} [{part_no or desc}] - Cột {col_let}: Số lượng '{qty_val}' phải lớn hơn 0.")
            except:
                col_let = excel_col_name(cols['qty'])
                row_errors.append(f"Dòng Excel {r + 1} [{part_no or desc}] - Cột {col_let}: Số lượng '{qty_val}' không phải là số hợp lệ.")
                qty = 1

        # Lấy thông tin kích thước và khối lượng
        L = row[cols['length']] if pd.notna(row[cols['length']]) else None
        W = row[cols['width']] if pd.notna(row[cols['width']]) else None
        T = row[cols['thickness']] if pd.notna(row[cols['thickness']]) else None
        W_B = row[cols['width_b']] if pd.notna(row[cols['width_b']]) else None
        
        tot_raw_wt = row[cols['tot_raw_wt']] if pd.notna(row[cols['tot_raw_wt']]) else None
        uom = str(row[cols['uom']]).strip() if pd.notna(row[cols['uom']]) else default_uom
        mat_type = str(row[cols['mat_type']]).strip() if pd.notna(row[cols['mat_type']]) else ""

        # Kiểm tra kích thước phôi với chi tiết thép gia công
        if "thép tấm" in desc.lower() and level > 0:
            if T is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}] - Cột {excel_col_name(cols['thickness'])}: Là Thép tấm nhưng thiếu 'ĐỘ DÀY (mm)'.")
            if L is None or W is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}]: Là Thép tấm nhưng thiếu 'CHIỀU DÀI' hoặc 'CHIỀU RỘNG'.")
            if tot_raw_wt is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}] - Cột {excel_col_name(cols['tot_raw_wt'])}: Thiếu 'TỔNG KL NGUYÊN LIỆU (kg)'.")

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
    min_level = min(it['level'] for it in items)
    
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
        
        p_ops = p_item['ops']
        has_weld = any(o['code'] == 'WD' for o in p_ops)
        has_pl   = any(o['code'] == 'PL' for o in p_ops)
        has_pnt  = any(o['code'] == 'PNT' for o in p_ops)
        has_pk   = any(o['code'] == 'PK' for o in p_ops)
        
        welded_part_code = f"SG - {customer_code} - {p_code}.01"
        welded_part_name = f"{p_desc}_WELDED"
        
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
            
        for child in children:
            c_code = child['part_no']
            c_desc = child['desc']
            c_wt   = child['tot_raw_wt']
            
            mat_type_lower = child.get('mat_type', '').lower()
            is_fastener = ("phụ" in mat_type_lower) or any(c_code.upper().startswith(pfx) for pfx in ['FA', 'D2150', 'M0', 'PIN']) or (child['tot_raw_wt'] is None and not any(k in c_desc.lower() for k in ['thép tấm', 'thép hộp', 'thép tròn', 'thép ống']))
            
            if not is_fastener:
                mat_name, phoi, mat_code = get_raw_mat_info(child, mat_spec_code)
                
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

# Giao diện chính
uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name='Form', header=None)
        
        # 1. Kiểm tra chặt chẽ cấu trúc 35 cột cố định của Template
        cfg, tpl_errors = validate_and_find_columns(df_raw)
        
        if tpl_errors:
            st.error(f"🚨 PHÁT HIỆN {len(tpl_errors)} LỖI CẤU TRÚC TEMPLATE (ĐÃ CHẶN CHUYỂN ĐỔI):")
            for err in tpl_errors:
                st.markdown(f"- ❌ {err}")
        else:
            # 2. Kiểm tra chi tiết dữ liệu từng dòng
            items, data_errors = parse_and_validate_bom(df_raw, cfg, cust_code, mat_spec_default, default_uom_fg)
            
            if data_errors:
                st.warning(f"⚠️ Phát hiện {len(data_errors)} lỗi/cảnh báo dữ liệu trên dòng:")
                with st.expander("Bấm để xem danh sách lỗi chi tiết từng dòng", expanded=True):
                    for derr in data_errors:
                        st.markdown(f"- {derr}")
                        
            # 3. Tạo dữ liệu BOM ERP
            df_result = build_erp_mbom_multi_level(items, cust_code, mat_spec_default, default_uom_fg)
            
            if len(df_result) > 0:
                st.success(f"✅ Template chuẩn 100%! Đã chuyển đổi thành công {len(df_result)} dòng BOM ERP.")
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
