# EBOM to MBOM Converter (New Template)

Ứng dụng web tự động chuyển đổi định mức BOM Kỹ thuật mới (5 Cấp bậc thụt dòng) sang BOM ERP phục vụ sản xuất cơ khí & gia công thép tấm xuất khẩu.

### Cấu trúc gói cài đặt:
- `bom_converter_app.py`: Mã nguồn chính chạy trên Streamlit.
- `requirements.txt`: Danh sách thư viện phụ thuộc Python.
- `README.md`: Hướng dẫn vận hành & mô tả dự án.

### Tính năng chính:
- Tự động bóc tách cấu trúc 5 Level (Level 0 -> Level 4).
- Tích hợp Routing quy trình xưởng: QTSX_01 đến QTSX_10.
- Áp dụng công thức tính % hao hụt và khối lượng tiêu hao (Kg).
- Xuất file Excel định dạng chuẩn để Import trực tiếp vào hệ thống ERP.
