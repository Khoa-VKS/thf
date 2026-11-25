import sys, os, re
sys.path.insert(0, os.path.abspath('.'))
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from Phan_Mem_Chuan_Hoa_QLAHS_Desktop import (
    load_excel_and_parse,
    reconcile_and_merge_multiple_case_lists,
    create_executive_pivot_sheet,
    create_simple_stage_pivot_sheet,
    clean_person_name,
    split_officers,
    normalize_case_stage_and_status
)

def reconcile_and_link_cases_defendants(case_files, defendant_files, progress_cb=None):
    def norm_s(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()
    def norm_txt(s):
        return re.sub(r'\s+', ' ', str(s or '')).strip().lower()

    if progress_cb:
        progress_cb("Đang tổng hợp danh sách vụ án...", "Khởi tạo dữ liệu...")

    if len(case_files) == 1:
        cases = [dict(c) for c in case_files[0][2]]
        case_stats = [{"fname": case_files[0][0], "total": len(cases), "matched": 0, "new": len(cases), "type": "Hồ sơ vụ án"}]
    else:
        cases, _, case_stats = reconcile_and_merge_multiple_case_lists(case_files, progress_cb)
        cases = [dict(c) for c in cases]
        for s in case_stats:
            s["type"] = "Hồ sơ vụ án"

    for c in cases:
        normalize_case_stage_and_status(c)

    case_by_code = {}
    for c in cases:
        cd = c.get("ma_vu_an")
        if cd:
            case_by_code[norm_s(cd)] = c

    def_records_all = []
    for fname, fmt, recs in defendant_files:
        for r in recs:
            r_copy = dict(r)
            r_copy["_fname"] = fname
            def_records_all.append(r_copy)

    has_blocks = any(r.get("block_idx") is not None for r in def_records_all)
    if has_blocks:
        grouped = {}
        for r in def_records_all:
            key = (r.get("_fname"), r.get("block_idx", 0))
            grouped.setdefault(key, []).append(r)
        blocks = list(grouped.values())
    else:
        grouped = {}
        blocks = []
        for r in def_records_all:
            mva = r.get("ma_vu_an")
            if mva:
                grouped.setdefault(norm_s(mva), []).append(r)
            else:
                blocks.append([r])
        blocks.extend(grouped.values())

    enriched_cases_map = {}
    linked_defs_all = []
    unlinked_defs_all = []

    for b_idx, blk in enumerate(blocks):
        if progress_cb:
            progress_cb(
                f"Đang đối chiếu & liên thông khối bị can {b_idx + 1}/{len(blocks)}...",
                f"Khối gồm {len(blk)} bị can ({blk[0].get('ten_bi_can', '')})"
            )

        target_case = None

        # 3.1 Khớp theo mã vụ án
        for d in blk:
            m_va = d.get("ma_vu_an")
            if m_va and norm_s(m_va) in case_by_code:
                target_case = case_by_code[norm_s(m_va)]
                break
            m_bc = d.get("ma_bi_can", "")
            m_match = re.match(r"^(VA_[A-Za-z0-9_\.]+?)_\d{2,3}$", m_bc)
            if m_match and norm_s(m_match.group(1)) in case_by_code:
                target_case = case_by_code[norm_s(m_match.group(1))]
                break

        # 3.2 Khớp theo CCCD hoặc Tên bị can (đặc biệt bị can đầu vụ)
        if not target_case:
            for d in blk:
                name_bc = norm_txt(d.get("ten_bi_can"))
                cccd_bc = str(d.get("so_cccd") or "").strip()
                for c in cases:
                    c_bc = norm_txt(c.get("bi_can_dau_vu") or "")
                    c_va = norm_txt(c.get("ten_vu_an") or "")
                    if cccd_bc and len(cccd_bc) >= 9 and cccd_bc in (c_bc + " " + c_va):
                        target_case = c
                        break
                    if name_bc and len(name_bc) >= 4 and (name_bc in c_bc or name_bc in c_va):
                        target_case = c
                        break
                if target_case:
                    break

        if target_case:
            c_code = norm_s(target_case.get("ma_vu_an", ""))
            enriched_cases_map[c_code] = target_case

            c_defs = []
            for idx, d in enumerate(blk, 1):
                cccd_str = f" (CCCD: {d['so_cccd']})" if d.get("so_cccd") else ""
                tt_str = f" - {d['trang_thai']}" if d.get("trang_thai") else ""
                c_defs.append(f"{idx}. {d.get('ten_bi_can', '')}{cccd_str}{tt_str}")

            existing_details = target_case.get("danh_sach_bi_can_chi_tiet", "")
            if existing_details:
                target_case["danh_sach_bi_can_chi_tiet"] = existing_details + "\n" + "\n".join(c_defs)
            else:
                target_case["danh_sach_bi_can_chi_tiet"] = "\n".join(c_defs)

            curr_total = 0
            try:
                curr_total = int(str(target_case.get("tong_bi_can") or 0).strip() or 0)
            except Exception:
                pass
            if curr_total < len(blk):
                target_case["tong_bi_can"] = len(blk)

            if blk:
                d0 = blk[0]
                lead_curr = str(target_case.get("bi_can_dau_vu") or "").lower()
                if not lead_curr or "chưa" in lead_curr:
                    cc_lead = f" - {d0.get('so_cccd')}" if d0.get("so_cccd") else ""
                    target_case["bi_can_dau_vu"] = f"{d0.get('ten_bi_can')}{cc_lead}"

            for d in blk:
                d["ma_vu_an_linked"] = target_case.get("ma_vu_an", "")
                d["ten_vu_an_linked"] = target_case.get("ten_vu_an", "")
                d["giai_doan_linked"] = target_case.get("giai_doan", "")
                d["trang_thai_va_linked"] = target_case.get("trang_thai_vu_an", "")
                d["vks_thu_ly_linked"] = target_case.get("vks_thu_ly", "")
                linked_defs_all.append(d)
        else:
            for d in blk:
                d["ma_vu_an_linked"] = "Xác minh sau"
                d["ten_vu_an_linked"] = "Chờ đối soát từ PM cũ / Bổ sung hồ sơ vụ án"
                unlinked_defs_all.append(d)

    def_stats = []
    for fname, fmt, recs in defendant_files:
        linked_in_file = sum(1 for d in linked_defs_all if d.get("_fname") == fname)
        unlinked_in_file = sum(1 for d in unlinked_defs_all if d.get("_fname") == fname)
        def_stats.append({
            "fname": fname,
            "total": len(recs),
            "matched": linked_in_file,
            "new": unlinked_in_file,
            "type": "Danh sách bị can"
        })

    all_file_stats = case_stats + def_stats
    total_enriched = len(enriched_cases_map)
    return cases, total_enriched, all_file_stats, linked_defs_all, unlinked_defs_all

def export_case_enriched_to_excel(case_records, linked_defs, unlinked_defs, out_filepath):
    for rec in case_records:
        normalize_case_stage_and_status(rec)
        if not rec.get("lanh_dao_vien") and not rec.get("lanh_dao_phong"):
            ld_v, ld_p, ksv, cb_tl = split_officers(rec.get("can_bo_giai_quyet", ""), rec.get("can_bo_thu_ly", ""))
            rec["lanh_dao_vien"] = ld_v
            rec["lanh_dao_phong"] = ld_p
            rec["can_bo_giai_quyet"] = ksv
            rec["can_bo_thu_ly"] = cb_tl
        else:
            if rec.get("lanh_dao_vien"):
                rec["lanh_dao_vien"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_vien"]).split(";")]))
            if rec.get("lanh_dao_phong"):
                rec["lanh_dao_phong"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["lanh_dao_phong"]).split(";")]))
            if rec.get("can_bo_giai_quyet"):
                rec["can_bo_giai_quyet"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_giai_quyet"]).split(";")]))
            if rec.get("can_bo_thu_ly"):
                rec["can_bo_thu_ly"] = "; ".join(filter(None, [clean_person_name(x) for x in str(rec["can_bo_thu_ly"]).split(";")]))

    wb_out = openpyxl.Workbook()
    create_executive_pivot_sheet(wb_out, case_records, title_prefix="BÁO CÁO TỔNG HỢP MA TRẬN PIVOT VỤ ÁN ĐÃ LIÊN THÔNG BỊ CAN", title="Tong_Quan_Pivot_Lanh_Dao", index=0)
    create_simple_stage_pivot_sheet(wb_out, case_records, title="Pivot_Giai_Doan_VKS", index=1)

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    # Sheet 3: Sổ vụ án đã làm giàu
    ws3 = wb_out.create_sheet(title="So_Vu_An_Da_Lam_Giau", index=2)
    if "Sheet" in wb_out.sheetnames:
        wb_out.remove(wb_out["Sheet"])
    ws3.views.sheetView[0].showGridLines = True

    COLUMNS_CASE = [
        ("stt", "STT", 6, "center", "0F2942"),
        ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 24, "center", "0F2942"),
        ("ma_vu_an", "Mã vụ án", 28, "center", "0F2942"),
        ("ten_vu_an", "Tên vụ án / Nội dung", 38, "left", "0F2942"),
        ("tinh_thanh", "Tỉnh / Thành phố", 18, "center", "0F2942"),
        ("so_qd_khoi_to", "Số QĐ khởi tố", 14, "center", "1E3A8A"),
        ("ngay_qd_khoi_to", "Ngày ra QĐ khởi tố", 16, "center", "1E3A8A"),
        ("don_vi_khoi_to", "Đơn vị ra QĐ", 28, "left", "1E3A8A"),
        ("so_thu_ly", "Số thụ lý", 14, "center", "0E7490"),
        ("ngay_thu_ly", "Ngày thụ lý", 16, "center", "0E7490"),
        ("loai_thu_ly", "Loại thụ lý", 16, "center", "0E7490"),
        ("vks_thu_ly", "Đơn vị VKS thụ lý", 34, "left", "0E7490"),
        ("thoi_han_giai_quyet", "Thời hạn giải quyết", 16, "center", "991B1B"),
        ("bi_can_dau_vu", "Bị can đầu vụ & CCCD", 28, "left", "166534"),
        ("tong_bi_can", "Tổng số bị can", 14, "center", "166534"),
        ("danh_sach_bi_can_chi_tiet", "Danh sách chi tiết bị can đã làm giàu", 46, "left", "166534"),
        ("dieu_luat", "Tội danh / Điều luật", 36, "left", "166534"),
        ("loai_toi", "Phân loại tội", 18, "center", "166534"),
        ("lanh_dao_vien", "Lãnh đạo Viện", 22, "center", "581C87"),
        ("lanh_dao_phong", "Lãnh đạo Phòng", 24, "center", "581C87"),
        ("can_bo_giai_quyet", "Kiểm sát viên", 24, "left", "581C87"),
        ("can_bo_thu_ly", "Cán bộ thụ lý", 20, "left", "581C87"),
        ("giai_doan", "Giai đoạn tố tụng", 16, "center", "334155"),
        ("trang_thai_vu_an", "Trạng thái vụ án", 24, "center", "334155"),
    ]

    # Banner rows 1-3
    ws3.merge_cells("A1:X1")
    ws3["A1"] = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    ws3["A1"].font = Font(name="Segoe UI", size=12, bold=True, color="FBBF24")
    ws3["A1"].fill = PatternFill(start_color="07172C", end_color="07172C", fill_type="solid")
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    ws3.merge_cells("A2:X2")
    ws3["A2"] = "SỔ QUẢN LÝ VỤ ÁN HÌNH SỰ ĐÃ LIÊN THÔNG & LÀM GIÀU DANH SÁCH BỊ CAN"
    ws3["A2"].font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    ws3["A2"].fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
    ws3["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[2].height = 34

    ws3.merge_cells("A3:X3")
    enriched_cnt = sum(1 for c in case_records if c.get("danh_sach_bi_can_chi_tiet"))
    ws3["A3"] = f"★ Tổng số: {len(case_records):,} vụ án — Đã liên thông & làm giàu thành công: {enriched_cnt:,} vụ án ({len(linked_defs):,} bị can) ★"
    ws3["A3"].font = Font(name="Segoe UI", size=10, bold=True, color="047857")
    ws3["A3"].fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    ws3["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[3].height = 24

    ws3.row_dimensions[4].height = 28
    for col_idx, (fkey, label, width, align_h, col_hex) in enumerate(COLUMNS_CASE, 1):
        cell = ws3.cell(4, col_idx, label)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    for r_idx, c in enumerate(case_records, 5):
        ws3.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _, align_h, _) in enumerate(COLUMNS_CASE, 1):
            val = c.get(k, "")
            cell = ws3.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if k in ("stt", "tong_bi_can", "so_qd_khoi_to", "so_thu_ly", "ngay_qd_khoi_to", "ngay_thu_ly", "thoi_han_giai_quyet", "giai_doan", "trang_thai_vu_an"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_vu_an",):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c.get("danh_sach_bi_can_chi_tiet") and k == "danh_sach_bi_can_chi_tiet":
                cell.fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")

    ws3.freeze_panes = "A5"
    ws3.auto_filter.ref = f"A4:{get_column_letter(len(COLUMNS_CASE))}{len(case_records)+4}"

    # Sheet 4: Chi tiết bị can đã liên thông
    ws4 = wb_out.create_sheet(title="Chi_Tiet_Bi_Can_Lien_Thong", index=3)
    ws4.views.sheetView[0].showGridLines = True

    COLUMNS_DEFS = [
        ("stt", "STT", 6),
        ("ma_vu_an_linked", "Mã vụ án liên thông", 28),
        ("ten_vu_an_linked", "Tên vụ án / Nội dung", 36),
        ("ma_bi_can", "Mã bị can", 28),
        ("ten_bi_can", "Họ và tên bị can", 22),
        ("ngay_sinh", "Ngày sinh", 14),
        ("gioi_tinh", "Giới tính", 10),
        ("so_cccd", "Số CCCD / CMND", 18),
        ("dia_chi", "Địa chỉ cư trú", 35),
        ("dieu_luat", "Tội danh / Điều luật khởi tố", 35),
        ("loai_toi", "Loại tội", 18),
        ("trang_thai", "Trạng thái bị can", 20),
        ("giai_doan_linked", "Giai đoạn vụ án", 16),
        ("trang_thai_va_linked", "Trạng thái vụ án", 22),
        ("vks_thu_ly_linked", "Đơn vị VKS thụ lý", 32),
    ]

    ws4.merge_cells("A1:O1")
    ws4["A1"] = "DANH SÁCH CHI TIẾT BỊ CAN ĐÃ LIÊN THÔNG CHÍNH XÁC VỚI HỒ SƠ VỤ ÁN"
    ws4["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    ws4["A1"].fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 30

    ws4.row_dimensions[2].height = 26
    for c_idx, (_, title, w) in enumerate(COLUMNS_DEFS, 1):
        cell = ws4.cell(2, c_idx, title)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws4.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, d in enumerate(linked_defs, 3):
        ws4.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(COLUMNS_DEFS, 1):
            val = d.get(k, "")
            cell = ws4.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if k in ("stt", "ngay_sinh", "gioi_tinh", "giai_doan_linked"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_vu_an_linked", "ma_bi_can", "so_cccd"):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws4.freeze_panes = "A3"
    if linked_defs:
        ws4.auto_filter.ref = f"A2:{get_column_letter(len(COLUMNS_DEFS))}{len(linked_defs)+2}"

    # Sheet 5: Bị can chờ xác minh PM cũ
    ws5 = wb_out.create_sheet(title="Bi_Can_Cho_Xac_Minh_PM_Cu", index=4)
    ws5.views.sheetView[0].showGridLines = True

    COLUMNS_UNLINKED = [
        ("stt", "STT", 6),
        ("ma_bi_can", "Mã bị can", 28),
        ("ten_bi_can", "Họ và tên bị can", 22),
        ("ngay_sinh", "Ngày sinh", 14),
        ("gioi_tinh", "Giới tính", 10),
        ("so_cccd", "Số CCCD / CMND", 18),
        ("dia_chi", "Địa chỉ cư trú", 35),
        ("dieu_luat", "Tội danh / Điều luật khởi tố", 35),
        ("loai_toi", "Loại tội", 18),
        ("trang_thai", "Trạng thái giải quyết", 20),
        ("ma_vu_an_linked", "Tình trạng mã vụ án", 20),
        ("ten_vu_an_linked", "Ghi chú đối soát", 36),
    ]

    ws5.merge_cells("A1:L1")
    ws5["A1"] = "DANH SÁCH BỊ CAN CHỜ ĐỐI SOÁT TỪ PHẦN MỀM CŨ / BỔ SUNG SỔ VỤ ÁN"
    ws5["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    ws5["A1"].fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 30

    ws5.row_dimensions[2].height = 26
    for c_idx, (_, title, w) in enumerate(COLUMNS_UNLINKED, 1):
        cell = ws5.cell(2, c_idx, title)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="7F1D1D", end_color="7F1D1D", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws5.column_dimensions[get_column_letter(c_idx)].width = w

    for r_idx, d in enumerate(unlinked_defs, 3):
        ws5.row_dimensions[r_idx].height = 24
        for c_idx, (k, _, _) in enumerate(COLUMNS_UNLINKED, 1):
            val = d.get(k, "")
            cell = ws5.cell(r_idx, c_idx, val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if k == "ma_vu_an_linked":
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
            elif k in ("stt", "ngay_sinh", "gioi_tinh"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif k in ("ma_bi_can", "so_cccd"):
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws5.freeze_panes = "A3"
    if unlinked_defs:
        ws5.auto_filter.ref = f"A2:{get_column_letter(len(COLUMNS_UNLINKED))}{len(unlinked_defs)+2}"

    wb_out.active = 0
    wb_out.save(out_filepath)
    print("Exported successfully to", out_filepath)

# Test run
f1 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162916746_407e1c4e.xlsx'
f2 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162929725_af85271c.xlsx'

fmt1, recs1 = load_excel_and_parse(f1)
fmt2, recs2 = load_excel_and_parse(f2)

case_files = [('Danh_sach_vu_an.xlsx', fmt1, recs1)]
def_files = [('Danh_sach_bi_can.xlsx', fmt2, recs2)]

cases, enriched_n, stats, l_defs, u_defs = reconcile_and_link_cases_defendants(case_files, def_files)
print(f"Total Cases: {len(cases)}, Enriched Cases: {enriched_n}")
print(f"Linked Defs: {len(l_defs)}, Unlinked Defs: {len(u_defs)}")
print("Stats:", stats)

out_file = 'scratch/test_5sheets_enriched.xlsx'
export_case_enriched_to_excel(cases, l_defs, u_defs, out_file)
print(f"File size: {os.path.getsize(out_file):,} bytes")
wb_chk = openpyxl.load_workbook(out_file)
print("Sheet names:", wb_chk.sheetnames)
