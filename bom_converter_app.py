import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(
    page_title="BOM Converter: EBOM to MBOM",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Chuyển Đổi BOM Kỹ Thuật Sang BOM ERP")
st.caption("Tự động nhận diện cấp MỤC linh hoạt, kiểm soát cấu trúc 35 cột cố định và bóc tách đệ quy đa tầng theo thứ tự công đoạn")

# Sidebar cấu hình
with st.sidebar:
    st.header("🔧 Cấu hình hệ thống")
    cust_code = st.text_input("Mã Khách hàng (KKK):", value="STV", max_chars=5)
    default_uom_fg = st.selectbox("ĐVT Thành phẩm / BTP:", ["Cái", "Bộ"], index=0)
    mat_spec_default = st.text_input("Mã phân loại thép (NVL Code Prefix):", value="S01")
    st.markdown("---")
    st.markdown("""
    **Logic nghiệp vụ tích hợp:**
    - **Vùng cột `MỤC`**: Nhận diện động mọi cấp độ thụt lề Cha - Con (từ cột A đến trước `MÃ SỐ CHI TIẾT`).
    - **Khóa cứng 35 cột cố định**: Đối soát tuần tự 1-1 từ `MÃ SỐ CHI TIẾT` đến `Đóng gói (PK)`.
    - **Tách dòng BTP đúng 2 công đoạn**: Chi tiết có đúng 2 công đoạn được tách dòng tiêu đề BTP riêng phía trên.
    - **Đánh số STT đồng cấp**: Tự động đánh số STT cho toàn bộ BTP đồng cấp, ngoại trừ các BTP ở cấp thấp nhất trên toàn bộ file BOM.
    - **Chuẩn hóa Tên NVL cho BTP**: Khi NVL đầu vào là BTP, cột `Tên NVL` lấy đúng giá trị tại cột `Mã sản phẩm` của BTP đó.
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
    if pd.isna(val):
        return ""
    return re.sub(r'[\r\n\s]+', ' ', str(val)).strip()

def excel_col_name(n):
    name = ""
    while n >= 0:
        name = chr(n % 26 + 65) + name
        n = n // 26 - 1
    return name

def match_col_spec(h_main, h_sub, spec):
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
    errors = []
    header_idx = None
    for r in range(min(20, len(df_raw))):
        row_vals = [clean_header_str(x).upper() for x in df_raw.iloc[r] if pd.notna(x)]
        if any("MÃ SỐ CHI TIẾT" in x or "MÃ CHI TIẾT" in x for x in row_vals):
            header_idx = r
            break
            
    if header_idx is None:
        return None, ["Không tìm thấy dòng tiêu đề chứa cột **'MÃ SỐ CHI TIẾT'** trong 20 dòng đầu tiên của sheet 'Form'."]
        
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

    actual_cols = []
    for c in range(idx_part, df_raw.shape[1]):
        hm = clean_header_str(df_raw.iloc[header_idx, c])
        hs = clean_header_str(df_raw.iloc[header_idx + 1, c]) if header_idx + 1 < len(df_raw) else ""
        actual_cols.append((c, hm, hs))
        
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
            {'col': idx_part + 26, 'code': 'CUT', 'qtsx': 'QTSX_05', 'name': 'Cắt (CUT)'},
            {'col': idx_part + 27, 'code': 'BD',  'qtsx': 'QTSX_01', 'name': 'Chấn/uốn (BD)'},
            {'col': idx_part + 28, 'code': 'WD',  'qtsx': 'QTSX_06', 'name': 'Hàn (WD)'},
            {'col': idx_part + 29, 'code': 'MC',  'qtsx': 'QTSX_07', 'name': 'Phay tiện (MC)'},
            {'col': idx_part + 30, 'code': 'FM',  'qtsx': 'QTSX_08', 'name': 'Dập tạo hình (FM)'},
            {'col': idx_part + 31, 'code': 'PNT', 'qtsx': 'QTSX_03', 'name': 'Sơn (PNT)'},
            {'col': idx_part + 32, 'code': 'PL',  'qtsx': 'QTSX_10', 'name': 'Xi mạ (PL)'},
            {'col': idx_part + 33, 'code': 'AS',  'qtsx': 'QTSX_09', 'name': 'Lắp ráp (AS)'},
            {'col': idx_part + 34, 'code': 'PK',  'qtsx': 'QTSX_04', 'name': 'Đóng gói (PK)'}
        ]
    }
    return {'header_row': header_idx, 'cols': col_mapping}, []

def get_raw_mat_info(item, mat_spec_code="S01"):
    desc = item['desc'].strip() if item['desc'] else ""
    desc_low = desc.lower()
    T = item['T']
    W = item['W']
    W_B = item.get('W_B')
    L = item['L']
    
    if "thép hộp" in desc_low or ("hộp" in desc_low and T is not None):
        t_int = int(round(float(T))) if (T is not None and str(T).strip() != '') else 0
        wb_val = W_B if W_B is not None else W
        mat_name = f"{desc} {W:.0f}x{wb_val:.0f}x{T:.0f}mm" if (W is not None and T is not None) else desc
        phoi = f"{W:.0f} x {wb_val:.0f} x {L:.0f}" if (W is not None and L is not None) else (f"{L}" if L is not None else None)
        mat_code = f"R-H-{mat_spec_code}-{t_int:02d}-00-000"
    elif "tròn đặc" in desc_low:
        w_int = int(round(float(W))) if (W is not None and str(W).strip() != '') else 0
        mat_name = f"{desc} phi {W:.0f}mm" if W is not None else desc
        phoi = f"dài {L:.0f}" if L is not None else None
        mat_code = f"R-B-{mat_spec_code}-{w_int:02d}-00-000"
    elif "thép ống" in desc_low or ("ống" in desc_low and "khói" not in desc_low and W is not None):
        t_int = int(round(float(T))) if (T is not None and str(T).strip() != '') else 0
        mat_name = f"{desc} phi {W:.0f}x{L:.0f}mm" if (W is not None and L is not None) else desc
        phoi = f"phi {W:.0f} x {L:.0f}" if (W is not None and L is not None) else None
        mat_code = f"R-T-{mat_spec_code}-{t_int:02d}-00-000"
    elif T is not None and str(T).strip() != '':
        t_val = float(T)
        t_str = f"{t_val:.0f}" if t_val.is_integer() else f"{t_val}"
        t_int = int(round(t_val))
        mat_name = f"{desc} {t_str}mm" if desc else f"Thép tấm {t_str}mm"
        phoi = f"{L} x {W}" if (L is not None and W is not None) else None
        mat_code = f"R-P-{mat_spec_code}-{t_int:02d}-00-000"
    else:
        mat_name = desc
        phoi = f"{L} x {W}" if (L is not None and W is not None) else None
        mat_code = item['part_no'] if item['part_no'] else None
        
    return mat_name, phoi, mat_code

def parse_and_validate_bom(df_raw, cfg, customer_code="STV", mat_spec_code="S01", default_uom="Cái"):
    row_errors = []
    items = []
    
    header_idx = cfg['header_row']
    cols = cfg['cols']
    num_level_cols = cols['level_zone_end']
    
    start_row = header_idx + 2
    
    for r in range(start_row, len(df_raw)):
        row = df_raw.iloc[r]
        
        first_cell_str = str(row[0]).strip().upper() if pd.notna(row[0]) else ""
        if "TOTAL" in first_cell_str or "TÍNH KHỐI LƯỢNG" in str(row[cols['unit_raw_wt']]).upper():
            break

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
        
        if not found_levels and not part_no and not desc:
            continue
            
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

        L = row[cols['length']] if pd.notna(row[cols['length']]) else None
        W = row[cols['width']] if pd.notna(row[cols['width']]) else None
        T = row[cols['thickness']] if pd.notna(row[cols['thickness']]) else None
        W_B = row[cols['width_b']] if pd.notna(row[cols['width_b']]) else None
        
        tot_raw_wt = row[cols['tot_raw_wt']] if pd.notna(row[cols['tot_raw_wt']]) else None
        uom = str(row[cols['uom']]).strip() if pd.notna(row[cols['uom']]) else default_uom
        mat_type = str(row[cols['mat_type']]).strip() if pd.notna(row[cols['mat_type']]) else ""

        if "thép tấm" in desc.lower() and level > 0:
            if T is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}] - Cột {excel_col_name(cols['thickness'])}: Là Thép tấm nhưng thiếu 'ĐỘ DÀY (mm)'.")
            if L is None or W is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}] - Cột {excel_col_name(cols['length'])} / {excel_col_name(cols['width'])}: Là Thép tấm nhưng thiếu 'CHIỀU DÀI' hoặc 'CHIỀU RỘNG'.")
            if tot_raw_wt is None:
                row_errors.append(f"Dòng Excel {r + 1} [{part_no}] - Cột {excel_col_name(cols['tot_raw_wt'])}: Thiếu 'TỔNG KL NGUYÊN LIỆU (kg)'.")

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

def group_operations(ops):
    if not ops:
        return []
        
    stages = []
    i = 0
    while i < len(ops):
        curr = ops[i]
        code = curr['code']
        if i + 1 < len(ops) and {code, ops[i + 1]['code']} == {'CUT', 'BD'}:
            stages.append({'qtsx': 'QTSX_01', 'stage': 'CUT, BD', 'suffix': '_CUT_BD'})
            i += 2
        elif i + 1 < len(ops) and {code, ops[i + 1]['code']} == {'WD', 'MC'}:
            stages.append({'qtsx': 'QTSX_02', 'stage': 'WD, MC', 'suffix': '_WELDED'})
            i += 2
        else:
            suffix_map = {
                'CUT': '_CUT', 'BD': '_BENT', 'WD': '_WELDED', 'MC': '_MACHINED',
                'FM': '_FORMED', 'PNT': '_PAINTED', 'PL': '_PLATED', 'AS': '_ASSY', 'PK': ''
            }
            qtsx_map = {
                'CUT': 'QTSX_05', 'BD': 'QTSX_01', 'WD': 'QTSX_06', 'MC': 'QTSX_07',
                'FM': 'QTSX_08', 'PNT': 'QTSX_03', 'PL': 'QTSX_10', 'AS': 'QTSX_09', 'PK': 'QTSX_04'
            }
            stages.append({
                'qtsx': qtsx_map.get(code, curr['qtsx']),
                'stage': code,
                'suffix': suffix_map.get(code, f"_{code}")
            })
            i += 1
    return stages

def build_tree_from_flat_items(items):
    roots = []
    stack = []
    for it in items:
        node = {'item': it, 'children': []}
        lvl = it['level']
        while stack and stack[-1][0] >= lvl:
            stack.pop()
        if not stack:
            roots.append(node)
        else:
            stack[-1][1]['children'].append(node)
        stack.append((lvl, node))
    return roots

def build_erp_mbom_multi_level(items, customer_code="STV", mat_spec_code="S01", default_uom="Cái"):
    if not items:
        return pd.DataFrame()
        
    roots = build_tree_from_flat_items(items)
    
    def is_btp_node(node):
        return bool(node['children']) or bool(group_operations(node['item']['ops']))
        
    btp_levels = []
    def collect_btp_levels(node):
        if is_btp_node(node):
            btp_levels.append(node['item']['level'])
        for ch in node['children']:
            collect_btp_levels(ch)
            
    for r_node in roots:
        collect_btp_levels(r_node)
        
    max_btp_level = max(btp_levels) if btp_levels else 0
    rows = []
    
    def process_node(node, stt_prefix, parent_code="", child_seq=1, is_root=False):
        it = node['item']
        children = node['children']
        lvl = it['level']
        
        node_stt = None if (lvl == max_btp_level and not is_root) else stt_prefix
        
        raw_part = it['part_no']
        if is_root and " _ " in raw_part:
            raw_code = raw_part.split(" _ ")[0].strip()
        elif raw_part:
            raw_code = raw_part
        else:
            raw_code = f"{parent_code}.0{child_seq}" if parent_code else it['desc']
            
        if not children and raw_code == parent_code:
            raw_code = f"{parent_code}.01"
            
        raw_desc = it['desc'] if it['desc'] else raw_code
        stages = group_operations(it['ops'])
        
        # 1. TRƯỜNG HỢP NODE LÁ
        if not children:
            has_raw_mat = (it['tot_raw_wt'] is not None) or any(k in raw_desc.lower() for k in ['thép tấm', 'thép hộp', 'thép tròn', 'thép ống', 'inox']) or bool(stages)
            if has_raw_mat and stages:
                mat_name, phoi, mat_code = get_raw_mat_info(it, mat_spec_code)
                c_wt = round(float(it['tot_raw_wt']), 2) if (it['tot_raw_wt'] is not None and str(it['tot_raw_wt']).strip() != '') else it['qty']
                raw_uom = 'Kg' if (it['tot_raw_wt'] is not None and str(it['tot_raw_wt']).strip() != '') else (it['uom'] if it['uom'] else default_uom)
                
                final_btp_full = f"SG - {customer_code} - {raw_code}"
                
                # CHỈ TÁCH DÒNG BTP KHI CÓ ĐÚNG 2 CÔNG ĐOẠN (len(stages) == 2)
                if len(stages) == 2:
                    rows.append({
                        'STT': node_stt,
                        'QTSX': None,
                        'Mã sản phẩm': final_btp_full,
                        'Loại': 'BTP',
                        'Tên sản phẩm': raw_code,
                        'SL SP': 1,
                        'ĐVT sp': default_uom,
                        'Mã NVL': None, 'Phôi (mm)': None, 'Tên NVL': None, 'ĐVT (NVL)': None, 'Số lượng': None,
                        '% Hao hụt': None, 'Công đoạn': None, 'BTP': 'x', 'TP': None, 'Số Kg/NVL': None
                    })
                    first_row_stt = None
                else:
                    first_row_stt = node_stt
                    
                st0 = stages[0]
                curr_btp_short = f"{raw_code}.01" if len(stages) >= 2 else raw_code
                curr_btp_full = f"SG - {customer_code} - {curr_btp_short}" if len(stages) >= 2 else final_btp_full
                curr_btp_name = f"{raw_code}{st0['suffix']}" if len(stages) >= 2 else raw_code
                
                rows.append({
                    'STT': first_row_stt,
                    'QTSX': st0['qtsx'],
                    'Mã sản phẩm': curr_btp_full,
                    'Loại': 'BTP',
                    'Tên sản phẩm': curr_btp_name,
                    'SL SP': 1,
                    'ĐVT sp': default_uom,
                    'Mã NVL': mat_code,
                    'Phôi (mm)': phoi,
                    'Tên NVL': mat_name,
                    'ĐVT (NVL)': raw_uom,
                    'Số lượng': c_wt,
                    '% Hao hụt': None,
                    'Công đoạn': st0['stage'],
                    'BTP': 'x',
                    'TP': None,
                    'Số Kg/NVL': None
                })
                
                for s_idx, st_next in enumerate(stages[1:], 2):
                    is_last_stage = (s_idx == len(stages))
                    next_btp_short = raw_code if is_last_stage else f"{raw_code}.0{s_idx}"
                    next_btp_full = f"SG - {customer_code} - {next_btp_short}" if not is_last_stage else final_btp_full
                    next_btp_name = raw_code if is_last_stage else f"{raw_code}{st_next['suffix']}"
                    
                    rows.append({
                        'STT': None,
                        'QTSX': st_next['qtsx'],
                        'Mã sản phẩm': next_btp_full,
                        'Loại': 'BTP',
                        'Tên sản phẩm': next_btp_name,
                        'SL SP': 1,
                        'ĐVT sp': default_uom,
                        'Mã NVL': curr_btp_short if not is_last_stage else curr_btp_short,
                        'Phôi (mm)': None,
                        'Tên NVL': curr_btp_full,
                        'ĐVT (NVL)': default_uom,
                        'Số lượng': 1,
                        '% Hao hụt': None,
                        'Công đoạn': st_next['stage'],
                        'BTP': 'x',
                        'TP': None,
                        'Số Kg/NVL': None
                    })
                    curr_btp_short = next_btp_short
                    curr_btp_full = next_btp_full
                    curr_btp_name = next_btp_name
                    
                return {'code': raw_code if len(stages) < 2 else curr_btp_short, 'name': final_btp_full, 'qty': it['qty'], 'uom': default_uom}
            else:
                return {'code': raw_code, 'name': raw_code, 'qty': it['qty'], 'uom': it['uom'] if it['uom'] else default_uom}

        # 2. TRƯỜNG HỢP MÃ CHA
        item_type = 'THÀNH PHẨM' if is_root else 'BTP'
        prefix_type = 'FG' if is_root else 'SG'
        main_full_code = f"{prefix_type} - {customer_code} - {raw_code}"
        
        rows.append({
            'STT': node_stt,
            'QTSX': None,
            'Mã sản phẩm': main_full_code,
            'Loại': item_type,
            'Tên sản phẩm': raw_desc,
            'SL SP': 1,
            'ĐVT sp': default_uom,
            'Mã NVL': None, 'Phôi (mm)': None, 'Tên NVL': None, 'ĐVT (NVL)': None, 'Số lượng': None,
            '% Hao hụt': None, 'Công đoạn': None,
            'BTP': 'x' if not is_root else None,
            'TP': 'x' if is_root else None,
            'Số Kg/NVL': None
        })
        
        # CHỈ TÁCH DÒNG BTP KHI CÓ ĐÚNG 2 CÔNG ĐOẠN (len(stages) == 2)
        if len(stages) == 2:
            st0 = stages[0]
            first_btp_short = f"{raw_code}.01"
            first_btp_full = f"SG - {customer_code} - {first_btp_short}"
            first_btp_name = f"{raw_desc}{st0['suffix']}"
            rows.append({
                'STT': f"{stt_prefix}.1" if node_stt is not None else None,
                'QTSX': None,
                'Mã sản phẩm': first_btp_full,
                'Loại': 'BTP',
                'Tên sản phẩm': first_btp_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': None, 'Phôi (mm)': None, 'Tên NVL': None, 'ĐVT (NVL)': None, 'Số lượng': None,
                '% Hao hụt': None, 'Công đoạn': None, 'BTP': 'x', 'TP': None, 'Số Kg/NVL': None
            })
        else:
            first_btp_short = raw_code
            first_btp_full = main_full_code
            first_btp_name = raw_desc

        child_outputs = []
        btp_child_counter = 1
        for c_idx, ch in enumerate(children, 1):
            if is_btp_node(ch):
                ch_stt = f"{stt_prefix}.{btp_child_counter}"
                btp_child_counter += 1
            else:
                ch_stt = None
            ch_out = process_node(ch, ch_stt, parent_code=raw_code, child_seq=c_idx, is_root=False)
            if ch_out:
                child_outputs.append(ch_out)
                
        if stages and child_outputs:
            st0 = stages[0]
            first_c = child_outputs[0]
            rows.append({
                'STT': None,
                'QTSX': st0['qtsx'],
                'Mã sản phẩm': first_btp_full,
                'Loại': 'BTP' if len(stages) >= 2 or not is_root else 'THÀNH PHẨM',
                'Tên sản phẩm': first_btp_name,
                'SL SP': 1,
                'ĐVT sp': default_uom,
                'Mã NVL': first_c['code'],
                'Phôi (mm)': None,
                'Tên NVL': first_c['name'],
                'ĐVT (NVL)': first_c['uom'],
                'Số lượng': first_c['qty'],
                '% Hao hụt': None,
                'Công đoạn': st0['stage'],
                'BTP': None,
                'TP': None,
                'Số Kg/NVL': None
            })
            for c_out in child_outputs[1:]:
                rows.append({
                    'STT': None, 'QTSX': None, 'Mã sản phẩm': None, 'Loại': None, 'Tên sản phẩm': None,
                    'SL SP': None, 'ĐVT sp': None,
                    'Mã NVL': c_out['code'],
                    'Phôi (mm)': None,
                    'Tên NVL': c_out['name'],
                    'ĐVT (NVL)': c_out['uom'],
                    'Số lượng': c_out['qty'],
                    '% Hao hụt': None, 'Công đoạn': None, 'BTP': None, 'TP': None, 'Số Kg/NVL': None
                })
                
            curr_short = first_btp_short
            curr_full = first_btp_full
            for s_idx, st_next in enumerate(stages[1:], 2):
                is_last = (s_idx == len(stages))
                if is_last and is_root:
                    next_short = raw_code
                    next_full = main_full_code
                    next_name = raw_desc
                    next_type = 'THÀNH PHẨM'
                else:
                    next_short = f"{curr_short}.01"
                    next_full = f"SG - {customer_code} - {next_short}" if not is_last else main_full_code
                    next_name = f"{raw_desc}{st_next['suffix']}"
                    next_type = 'BTP'
                    
                rows.append({
                    'STT': f"{stt_prefix}.{s_idx}" if node_stt is not None else None,
                    'QTSX': st_next['qtsx'],
                    'Mã sản phẩm': next_full,
                    'Loại': next_type,
                    'Tên sản phẩm': next_name,
                    'SL SP': 1,
                    'ĐVT sp': default_uom,
                    'Mã NVL': curr_short,
                    'Phôi (mm)': None,
                    'Tên NVL': curr_full,
                    'ĐVT (NVL)': default_uom,
                    'Số lượng': 1,
                    '% Hao hụt': None,
                    'Công đoạn': st_next['stage'],
                    'BTP': 'x' if next_type == 'BTP' else None,
                    'TP': None,
                    'Số Kg/NVL': None
                })
                curr_short = next_short
                curr_full = next_full
                
        return {'code': raw_code if len(stages) < 2 else curr_short, 'name': main_full_code, 'qty': it['qty'], 'uom': default_uom}

    for r_idx, root in enumerate(roots, 1):
        process_node(root, str(r_idx), parent_code="", child_seq=r_idx, is_root=True)
        
    return pd.DataFrame(rows)

# Giao diện chính
uploaded_file = st.file_uploader("Tải lên file BOM Kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name='Form', header=None)
        
        cfg, tpl_errors = validate_and_find_columns(df_raw)
        
        if tpl_errors:
            st.error(f"🚨 PHÁT HIỆN {len(tpl_errors)} LỖI CẤU TRÚC TEMPLATE (ĐÃ CHẶN CHUYỂN ĐỔI):")
            for err in tpl_errors:
                st.markdown(f"- ❌ {err}")
        else:
            items, data_errors = parse_and_validate_bom(df_raw, cfg, cust_code, mat_spec_default, default_uom_fg)
            
            if data_errors:
                st.warning(f"⚠️ Phát hiện {len(data_errors)} lỗi/cảnh báo dữ liệu trên dòng:")
                with st.expander("Bấm để xem danh sách lỗi chi tiết từng dòng", expanded=True):
                    for derr in data_errors:
                        st.markdown(f"- {derr}")
                        
            df_result = build_erp_mbom_multi_level(items, cust_code, mat_spec_default, default_uom_fg)
            
            if len(df_result) > 0:
                st.success(f"✅ Đã chuyển đổi thành công {len(df_result)} dòng BOM ERP.")
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
