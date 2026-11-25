import sys, os
sys.path.insert(0, os.path.abspath('.'))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from scratch.run_full_link_test import en_cases, l_defs, u_defs, cases

def export_case_enriched_to_excel(case_records, linked_defs, unlinked_defs, out_filepath):
    wb = openpyxl.Workbook()

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    # 1. Sheet 1: Danh sách vụ án đã làm giàu
    ws1 = wb.active
    ws1.title = "So_Vu_An_Da_Lam_Giau"
    ws1.views.sheetView[0].showGridLines = True

    headers_case = [
        ("stt", "STT", 6),
        ("ma_vu_an", "Mã vụ án", 28),
        ("ten_vu_an", "Tên vụ án / Nội dung", 38),
        ("tinh_thanh", "Tỉnh / Thành", 16),
        ("so_qd_khoi_to", "Số QĐ khởi tố", 14),
        ("ngay_qd_khoi_to", "Ngày ra QĐ", 14),
        ("so_thu_ly", "Số thụ lý", 14),
        ("ngay_thu_ly", "Ngày thụ lý", 14),
        ("vks_thu_ly", "Đơn vị VKS", 30),
        ("thoi_han_giai_quyet", "Thời hạn giải quyết", 16),
        ("bi_can_dau_vu", "Bị can đầu vụ & CCCD", 28),
        ("tong_bi_can", "Tổng số bị can", 14),
        ("danh_sach_bi_can_chi_tiet", "Danh sách chi tiết bị can đã làm giàu", 45),
        ("dieu_luat", "Tội danh / Điều luật", 36),
        ("loai_toi", "Loại tội", 18),
        ("giai_doan", "Giai đoạn", 14),
        ("trang_thai_vu_an", "Trạng thái vụ án", 24),
    ]

    ws1.row_dimensions[1].height = 28
    for c_idx, (_, title, w) in enumerate(headers_case, 1):
        cell = ws1.cell(1, c_idx, title)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws1.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, c in enumerate(case_records, 2):
        ws1.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(headers_case, 1):
            val = c.get(k, "")
            cell = ws1.cell(r_idx, c_idx, val)
            cell.font = Font(name="Times New Roman", size=10)
            cell.border = thin_border
            if k in ("stt", "tong_bi_can"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_vu_an", "so_qd_khoi_to", "so_thu_ly", "ngay_qd_khoi_to", "ngay_thu_ly", "thoi_han_giai_quyet", "giai_doan"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = f"A1:{get_column_letter(len(headers_case))}{len(case_records)+1}"

    # 2. Sheet 2: Chi tiết bị can liên thông
    ws2 = wb.create_sheet("Chi_Tiet_Bi_Can_Lien_Thong")
    ws2.views.sheetView[0].showGridLines = True

    headers_defs = [
        ("stt", "STT", 6),
        ("ma_vu_an_linked", "Mã vụ án liên thông", 28),
        ("ten_vu_an_linked", "Tên vụ án", 35),
        ("ma_bi_can", "Mã bị can", 28),
        ("ten_bi_can", "Họ và tên bị can", 22),
        ("ngay_sinh", "Ngày sinh", 14),
        ("gioi_tinh", "Giới tính", 10),
        ("so_cccd", "Số CCCD / CMND", 18),
        ("dia_chi", "Địa chỉ cư trú", 35),
        ("dieu_luat", "Tội danh / Điều khoản", 35),
        ("loai_toi", "Loại tội", 18),
        ("trang_thai", "Trạng thái bị can", 20),
        ("giai_doan_linked", "Giai đoạn vụ án", 14),
        ("trang_thai_va_linked", "Trạng thái vụ án", 22),
    ]

    ws2.row_dimensions[1].height = 28
    for c_idx, (_, title, w) in enumerate(headers_defs, 1):
        cell = ws2.cell(1, c_idx, title)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws2.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, d in enumerate(linked_defs, 2):
        ws2.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(headers_defs, 1):
            val = d.get(k, "")
            cell = ws2.cell(r_idx, c_idx, val)
            cell.font = Font(name="Times New Roman", size=10)
            cell.border = thin_border
            if k in ("stt", "ngay_sinh", "gioi_tinh", "giai_doan_linked"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_vu_an_linked", "ma_bi_can", "so_cccd"):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws2.freeze_panes = "A2"
    if linked_defs:
        ws2.auto_filter.ref = f"A1:{get_column_letter(len(headers_defs))}{len(linked_defs)+1}"

    # 3. Sheet 3: Bị can chờ xác minh (PM cũ)
    ws3 = wb.create_sheet("Bi_Can_Cho_Xac_Minh_PM_Cu")
    ws3.views.sheetView[0].showGridLines = True

    headers_unlinked = [
        ("stt", "STT", 6),
        ("ma_bi_can", "Mã bị can", 28),
        ("ten_bi_can", "Họ và tên bị can", 22),
        ("ngay_sinh", "Ngày sinh", 14),
        ("gioi_tinh", "Giới tính", 10),
        ("so_cccd", "Số CCCD / CMND", 18),
        ("dia_chi", "Địa chỉ cư trú", 35),
        ("dieu_luat", "Tội danh / Điều khoản", 35),
        ("loai_toi", "Loại tội", 18),
        ("trang_thai", "Trạng thái giải quyết", 20),
        ("ma_vu_an_linked", "Tình trạng mã vụ án", 18),
        ("ten_vu_an_linked", "Ghi chú đối soát", 35),
    ]

    ws3.row_dimensions[1].height = 28
    for c_idx, (_, title, w) in enumerate(headers_unlinked, 1):
        cell = ws3.cell(1, c_idx, title)
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws3.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, d in enumerate(unlinked_defs, 2):
        ws3.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(headers_unlinked, 1):
            val = d.get(k, "")
            cell = ws3.cell(r_idx, c_idx, val)
            cell.font = Font(name="Times New Roman", size=10)
            cell.border = thin_border
            if k == "ma_vu_an_linked":
                cell.font = Font(name="Times New Roman", size=10, bold=True, color="991B1B")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("stt", "ngay_sinh", "gioi_tinh"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_bi_can", "so_cccd"):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws3.freeze_panes = "A2"
    if unlinked_defs:
        ws3.auto_filter.ref = f"A1:{get_column_letter(len(headers_unlinked))}{len(unlinked_defs)+1}"

    wb.save(out_filepath)
    print(f"Exported successfully to {out_filepath}")

out_test = 'scratch/test_enriched_export.xlsx'
export_case_enriched_to_excel(cases, l_defs, u_defs, out_test)
print(f"File size: {os.path.getsize(out_test):,} bytes")
