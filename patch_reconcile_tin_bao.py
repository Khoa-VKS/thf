# -*- coding: utf-8 -*-
"""
Patch thêm chức năng đối chiếu và hợp nhất nhiều file Tin Báo vào Phan_Mem_Chuan_Hoa_QLAHS_Desktop.py
"""

with open('Phan_Mem_Chuan_Hoa_QLAHS_Desktop.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Thêm COLUMNS_TIN_BAO_RECONCILE_36
tb_recon_def = '''COLUMNS_TIN_BAO_RECONCILE_36 = [
    ("stt", "STT", 50),
    ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 240),
] + [col for col in COLUMNS_TIN_BAO_35 if col[0] != "stt"]
'''

if 'COLUMNS_TIN_BAO_RECONCILE_36' not in content:
    idx_col = content.find('COLUMNS_TIN_BAO_30 = COLUMNS_TIN_BAO_35')
    assert idx_col != -1
    content = content[:idx_col] + 'COLUMNS_TIN_BAO_30 = COLUMNS_TIN_BAO_35\n' + tb_recon_def + content[idx_col + len('COLUMNS_TIN_BAO_30 = COLUMNS_TIN_BAO_35'):]
    print("Added COLUMNS_TIN_BAO_RECONCILE_36")

# 2. Thêm hàm reconcile_and_merge_multiple_tin_bao_lists
fn_recon_tb = '''def reconcile_and_merge_multiple_tin_bao_lists(parsed_files, progress_cb=None):
    """
    Rà soát, đối chiếu và hợp nhất nhiều file Tin báo tội phạm QLAHS (2, 3, ... N files).
    Tự động so khớp theo Mã tin báo & Số thụ lý (Số TL) + Đơn vị VKS, bù đắp thông tin còn thiếu
    từ tất cả các file để tạo danh sách hợp nhất 36 cột đầy đủ nhất.
    """
    def norm_key(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

    master_records = []
    index_by_code = {}   # norm_key(ma_tin_bao) -> idx in master_records
    index_by_stl = {}    # norm_key(so_thu_ly + vks_thu_ly) -> idx in master_records
    file_stats = []      # {"fname": fname, "total": ..., "matched": ..., "new": ...}
    total_matches = 0

    for file_idx, (fname, fmt, records) in enumerate(parsed_files):
        if progress_cb:
            progress_cb(
                f"Đang so khớp & bù đắp dữ liệu Tin báo Tập tin {file_idx + 1}/{len(parsed_files)}...",
                f"{fname} ({len(records):,} tin báo)"
            )

        file_matched = 0
        file_new = 0

        for r in records:
            code_k = norm_key(r.get("ma_tin_bao"))
            so_tl = str(r.get("so_thu_ly", "") or "").strip()
            vks_tl = str(r.get("vks_thu_ly", "") or "").strip()
            stl_k = norm_key(so_tl + vks_tl) if so_tl else ""

            target_idx = None
            if code_k and code_k in index_by_code:
                target_idx = index_by_code[code_k]
            elif stl_k and stl_k in index_by_stl:
                target_idx = index_by_stl[stl_k]

            if target_idx is not None:
                # Đã có trong danh sách -> Tiến hành bù đắp & làm giàu dữ liệu
                target = master_records[target_idx]
                file_matched += 1
                total_matches += 1

                sources = target.get("_sources", [])
                if fname not in sources:
                    sources.append(fname)
                    target["_sources"] = sources
                if len(sources) > 1:
                    target["nguon_du_lieu"] = f"Khớp nối ({len(sources)} file): " + " + ".join(sources)
                else:
                    target["nguon_du_lieu"] = f"Tệp: {sources[0]}"

                # Bù đắp từng trường thông tin còn trống
                for k, v in r.items():
                    if k.startswith("_") or k in ("stt", "nguon_du_lieu"):
                        continue
                    if v and not target.get(k):
                        target[k] = v
                    elif k == "noi_dung_tin_bao" and len(str(v or "")) > len(str(target.get("noi_dung_tin_bao") or "")):
                        target["noi_dung_tin_bao"] = v
                    elif k in ("ksv_chinh", "dieu_tra_vien") and v and (not target.get(k) or str(target.get(k)).startswith("=")):
                        target[k] = v
                    elif k in ("thoi_han_den_ngay", "muc_do_khan_cap") and v and not target.get(k):
                        target[k] = v

                if code_k and code_k not in index_by_code:
                    index_by_code[code_k] = target_idx
                if stl_k and stl_k not in index_by_stl:
                    index_by_stl[stl_k] = target_idx
            else:
                # Tin báo mới -> Thêm vào master
                file_new += 1
                new_rec = dict(r)
                new_rec["_sources"] = [fname]
                new_rec["nguon_du_lieu"] = f"Tệp: {fname}"

                new_idx = len(master_records)
                master_records.append(new_rec)

                if code_k:
                    index_by_code[code_k] = new_idx
                if stl_k:
                    index_by_stl[stl_k] = new_idx

        file_stats.append({
            "fname": fname,
            "total": len(records),
            "matched": file_matched,
            "new": file_new,
        })

    # Cập nhật STT và dọn biến tạm
    for idx, r in enumerate(master_records, 1):
        r["stt"] = idx
        r.pop("_sources", None)

    return master_records, total_matches, file_stats


'''

if 'def reconcile_and_merge_multiple_tin_bao_lists' not in content:
    idx_recon_case = content.find('def reconcile_and_merge_case_lists(list_a, list_b):')
    assert idx_recon_case != -1
    # Tìm dòng return của reconcile_and_merge_case_lists
    idx_next_fn = content.find('\ndef clean_vks_unit_name', idx_recon_case)
    assert idx_next_fn != -1
    content = content[:idx_next_fn] + '\n\n' + fn_recon_tb + content[idx_next_fn:]
    print("Added reconcile_and_merge_multiple_tin_bao_lists")

# 3. Nâng cấp export_tin_bao_to_excel hỗ trợ has_sources
old_export_tin_bao_start = content.find('def export_tin_bao_to_excel(records, out_filepath):')
old_export_tin_bao_end = content.find('\ndef parse_qlahs_case_rows', old_export_tin_bao_start)
assert old_export_tin_bao_start != -1 and old_export_tin_bao_end != -1

new_export_tin_bao = '''def export_tin_bao_to_excel(records, out_filepath):
    """Xuất file Excel chuẩn hóa Danh sách Tin báo về tội phạm (hỗ trợ cả 35 cột đơn file và 36 cột hợp nhất đa nguồn)."""
    wb_out = openpyxl.Workbook()

    ws1 = wb_out.active
    ws1.title = "Danh_Sach_Tin_Bao_Chuan_Hoa"
    ws1.views.sheetView[0].showGridLines = True

    has_sources = bool(records and any(r.get("nguon_du_lieu") for r in records))

    cols_config_base = [
        ("stt", "STT", 7, "center", "1E3A8A"),
        ("ma_tin_bao", "Mã tin báo", 18, "center", "1E3A8A"),
        ("loai_tin", "Loại tin", 12, "center", "1E3A8A"),
        ("nguon_tin", "Phân loại nguồn tin", 16, "center", "1E3A8A"),
        ("tinh_tiet_phuc_tap", "Phức tạp", 10, "center", "1E3A8A"),

        ("nguoi_bao_tin", "Người / Cơ quan báo tin", 24, "left", "1E40AF"),
        ("nguoi_bi_to_giac", "Người / Đối tượng bị tố giác", 24, "left", "1E40AF"),
        ("cccd_nguoi_bao", "CCCD / Định danh", 16, "center", "1E40AF"),
        ("sdt_nguoi_bao", "Số điện thoại", 14, "center", "1E40AF"),
        ("dia_chi_nguoi_bao", "Địa chỉ người báo", 25, "left", "1E40AF"),

        ("thoi_gian_dia_diem", "Thời gian & Địa điểm", 22, "left", "0F766E"),
        ("noi_dung_tin_bao", "Nội dung tóm tắt nguồn tin", 38, "left", "0F766E"),
        ("dieu_luat", "Thông tin điều luật / Tội danh", 30, "left", "0F766E"),
        ("chuong_toi", "Chương BLHS", 24, "left", "0F766E"),
        ("muc_do_nghiem_trong", "Mức độ tội", 16, "center", "0F766E"),

        ("don_vi_tiep_nhan", "Đơn vị tiếp nhận ban đầu", 24, "left", "0369A1"),
        ("ngay_tiep_nhan_tin", "Ngày tiếp nhận tin", 14, "center", "0369A1"),
        ("ngay_cqdt_nhan", "Ngày CQĐT thụ lý", 14, "center", "0369A1"),
        ("dieu_tra_vien", "Điều tra viên (ĐTV)", 20, "left", "0369A1"),
        ("ngay_vks_nhan", "Ngày VKS tiếp nhận", 14, "center", "0369A1"),
        ("so_thu_ly", "Số thụ lý (Số TL)", 14, "center", "0369A1"),
        ("ngay_vks_thu_ly", "Ngày VKS thụ lý", 14, "center", "0369A1"),
        ("vks_thu_ly", "Đơn vị VKS thụ lý", 24, "left", "0369A1"),

        ("ksv_chinh", "Kiểm sát viên thụ lý chính", 22, "left", "5B21B6"),
        ("ksv_phoi_hop", "KSV phối hợp / giúp việc", 20, "left", "5B21B6"),
        ("don_vi_giai_quyet", "Đơn vị giải quyết", 20, "left", "5B21B6"),

        ("thoi_han_tu_ngay", "Thời hạn từ ngày", 14, "center", "991B1B"),
        ("thoi_han_den_ngay", "Thời hạn đến ngày", 14, "center", "991B1B"),
        ("so_ngay_con_lai", "Số ngày còn lại", 15, "center", "991B1B"),
        ("muc_do_khan_cap", "Cảnh báo tiến độ / Thời hạn", 25, "left", "991B1B"),
        ("ket_qua_giai_quyet", "Tình trạng / Kết quả giải quyết", 18, "center", "991B1B"),

        ("ngay_tam_dinh_chi", "Ngày tạm đình chỉ", 14, "center", "334155"),
        ("ly_do_tam_dinh_chi", "Lý do tạm đình chỉ", 24, "left", "334155"),
        ("ngay_het_thoi_hieu", "Ngày hết thời hiệu truy cứu", 16, "center", "334155"),
        ("ghi_chu", "Ghi chú", 20, "left", "334155"),
    ]

    if has_sources:
        cols_config = [
            ("stt", "STT", 7, "center", "1E3A8A"),
            ("nguon_du_lieu", "Nguồn dữ liệu đối chiếu", 30, "left", "1E3A8A"),
        ] + [c for c in cols_config_base if c[0] != "stt"]
        groups_config = [
            ("A4:F4", "1. NHẬN DIỆN, ĐỐI CHIẾU & NGUỒN TIN", "1E3A8A", "E0E7FF"),
            ("G4:K4", "2. NGƯỜI BÁO TIN & ĐỐI TƯỢNG BỊ TỐ GIÁC", "1E40AF", "DBEAFE"),
            ("L4:P4", "3. NỘI DUNG NGUỒN TIN & TỘI DANH", "0F766E", "CCFBF1"),
            ("Q4:X4", "4. TIẾP NHẬN & THỤ LÝ KIỂM SÁT", "0369A1", "E0F2FE"),
            ("Y4:AA4", "5. CÁN BỘ KIỂM SÁT THỤ LÝ", "5B21B6", "EDE9FE"),
            ("AB4:AF4", "6. THỜI HẠN & TIẾN ĐỘ GIẢI QUYẾT", "991B1B", "FEE2E2"),
            ("AG4:AJ4", "7. TẠM ĐÌNH CHỈ & THỜI HIỆU", "334155", "F1F5F9"),
        ]
        freeze_cell = "D6"
        banner_sub = f"BẢNG HỢP NHẤT & ĐỐI CHIẾU ĐA NGUỒN DANH SÁCH TIN BÁO TỐ GIÁC TỘI PHẠM QLAHS (Tổng hợp {len(records):,} tin báo)"
    else:
        cols_config = cols_config_base
        groups_config = [
            ("A4:E4", "1. NHẬN DIỆN & NGUỒN TIN", "1E3A8A", "E0E7FF"),
            ("F4:J4", "2. NGƯỜI BÁO TIN & ĐỐI TƯỢNG BỊ TỐ GIÁC", "1E40AF", "DBEAFE"),
            ("K4:O4", "3. NỘI DUNG NGUỒN TIN & TỘI DANH", "0F766E", "CCFBF1"),
            ("P4:W4", "4. TIẾP NHẬN & THỤ LÝ KIỂM SÁT", "0369A1", "E0F2FE"),
            ("X4:Z4", "5. CÁN BỘ KIỂM SÁT THỤ LÝ", "5B21B6", "EDE9FE"),
            ("AA4:AE4", "6. THỜI HẠN & TIẾN ĐỘ GIẢI QUYẾT", "991B1B", "FEE2E2"),
            ("AF4:AI4", "7. TẠM ĐÌNH CHỈ & THỜI HIỆU", "334155", "F1F5F9"),
        ]
        freeze_cell = "C6"
        banner_sub = f"BẢNG TỔNG HỢP & CHUẨN HÓA DANH SÁCH TIN BÁO TỐ GIÁC TỘI PHẠM QLAHS (Tổng hợp {len(records):,} tin báo)"

    max_col_letter = get_column_letter(len(cols_config))

    ws1.row_dimensions[1].height = 28
    ws1.row_dimensions[2].height = 20
    ws1.row_dimensions[3].height = 10
    ws1.row_dimensions[4].height = 26
    ws1.row_dimensions[5].height = 32

    # Banner Header
    ws1.merge_cells(f"A1:{max_col_letter}1")
    cell_b1 = ws1["A1"]
    cell_b1.value = "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"
    cell_b1.font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    cell_b1.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    cell_b1.alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells(f"A2:{max_col_letter}2")
    cell_b2 = ws1["A2"]
    time_stamp_full = datetime.now().strftime("%d/%m/%Y %H:%M")
    cell_b2.value = f"{banner_sub} - Xuất ngày: {time_stamp_full}"
    cell_b2.font = Font(name="Segoe UI", size=10, italic=True, color="E0E7FF")
    cell_b2.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    cell_b2.alignment = Alignment(horizontal="center", vertical="center")

    # Dòng nhóm 4
    for rng, gtitle, bg_c, fg_c in groups_config:
        ws1.merge_cells(rng)
        first_c = ws1[rng.split(":")[0]]
        first_c.value = gtitle
        first_c.font = Font(name="Segoe UI", size=10, bold=True, color=bg_c)
        first_c.fill = PatternFill(start_color=fg_c, end_color=fg_c, fill_type="solid")
        first_c.alignment = Alignment(horizontal="center", vertical="center")
        set_range_border(ws1, rng)

    # Dòng 5: Chi tiết các cột
    for c_idx, (_, col_title, width, _, bg_c) in enumerate(cols_config, start=1):
        c = ws1.cell(row=5, column=c_idx)
        c.value = col_title
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        col_letter = get_column_letter(c_idx)
        ws1.column_dimensions[col_letter].width = width

    # Ghi dữ liệu dòng 6 trở đi
    thin_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    font_data = Font(name="Segoe UI", size=9)
    font_data_bold = Font(name="Segoe UI", size=9, bold=True)

    fill_even = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_odd = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_overdue = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_urgent = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    fill_warning = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    fill_safe = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    for r_idx, rec in enumerate(records, start=6):
        ws1.row_dimensions[r_idx].height = 22
        base_fill = fill_even if (r_idx % 2 == 0) else fill_odd
        urg = rec.get("muc_do_khan_cap", "")

        for c_idx, (col_key, _, _, align, _) in enumerate(cols_config, start=1):
            c = ws1.cell(row=r_idx, column=c_idx)
            val = rec.get(col_key, "")
            c.value = val
            c.border = cell_border
            c.font = font_data
            c.alignment = Alignment(horizontal=align, vertical="center")

            if col_key == "muc_do_khan_cap":
                c.font = font_data_bold
                if "🔴" in urg:
                    c.fill = fill_overdue
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")
                elif "🟠" in urg:
                    c.fill = fill_urgent
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="C2410C")
                elif "🟡" in urg:
                    c.fill = fill_warning
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="854D0E")
                elif "🟢" in urg:
                    c.fill = fill_safe
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="166534")
                else:
                    c.fill = base_fill
            elif col_key == "so_ngay_con_lai":
                c.font = font_data_bold
                if "-" in str(val):
                    c.font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")
                c.fill = base_fill
            else:
                c.fill = base_fill

            if col_key in ("cccd_nguoi_bao", "sdt_nguoi_bao", "so_thu_ly"):
                c.number_format = "@"

    ws1.freeze_panes = freeze_cell
    ws1.auto_filter.ref = f"A5:{max_col_letter}{len(records) + 5}"

    # Sheet 2: Báo cáo Tiến độ giải quyết Tin báo & Thống kê KSV
    ws2 = wb_out.create_sheet(title="Bao_Cao_Tien_Do_Tin_Bao")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G1")
    t2 = ws2["A1"]
    t2.value = "BÁO CÁO TIẾN ĐỘ GIẢI QUYẾT & PHÂN CÔNG KIỂM SÁT VIÊN THỤ LÝ TIN BÁO"
    t2.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t2.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    total_tb = len(records)
    overdue = sum(1 for r in records if "🔴" in r.get("muc_do_khan_cap", ""))
    urgent = sum(1 for r in records if "🟠" in r.get("muc_do_khan_cap", ""))
    warning = sum(1 for r in records if "🟡" in r.get("muc_do_khan_cap", ""))
    safe = sum(1 for r in records if "🟢" in r.get("muc_do_khan_cap", ""))
    solved = sum(1 for r in records if "✔" in r.get("muc_do_khan_cap", ""))
    suspended = sum(1 for r in records if "⏸" in r.get("muc_do_khan_cap", ""))

    kpi_rows = [
        ("1", "Tổng số tin báo tội phạm thụ lý trên hệ thống", total_tb, "100.0%", "Toàn bộ nguồn tin đã số hóa"),
        ("2", "Đã giải quyết (Khởi tố / Không khởi tố / Chuyển nơi khác)", solved, f"{(solved/total_tb*100) if total_tb else 0:.1f}%", "Đã hoàn thành tố tụng"),
        ("3", "Tạm đình chỉ giải quyết nguồn tin", suspended, f"{(suspended/total_tb*100) if total_tb else 0:.1f}%", "Chờ kết quả giám định / xác minh"),
        ("4", "Đang trong thời hạn giải quyết an toàn (> 15 ngày)", safe, f"{(safe/total_tb*100) if total_tb else 0:.1f}%", "Tiến độ trong tầm kiểm soát"),
        ("5", "Cảnh báo thời hạn giải quyết (Còn 8 - 15 ngày)", warning, f"{(warning/total_tb*100) if total_tb else 0:.1f}%", "Cần đẩy nhanh tiến độ"),
        ("6", "Khẩn cấp sắp hết thời hạn giải quyết (Còn ≤ 7 ngày)", urgent, f"{(urgent/total_tb*100) if total_tb else 0:.1f}%", "Yêu cầu ra quyết định khẩn"),
        ("7", "QUÁ HẠN GIẢI QUYẾT LUẬT ĐỊNH (Cần chỉ đạo xử lý gấp)", overdue, f"{(overdue/total_tb*100) if total_tb else 0:.1f}%", "Cần báo cáo Lãnh đạo Viện ngay"),
    ]

    ws2.row_dimensions[3].height = 24
    h_kpi = ["STT", "Chỉ Tiêu Đánh Giá Tiến Độ Giải Quyết", "Số Lượng", "Tỷ Lệ %", "Ghi Chú Đôn Đốc Nghiệp Vụ"]
    for i, h in enumerate(h_kpi, start=1):
        c = ws2.cell(row=3, column=i)
        c.value = h
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 50
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 12
    ws2.column_dimensions["E"].width = 40

    for idx, (stt, title, val, pct, note) in enumerate(kpi_rows, start=4):
        ws2.row_dimensions[idx].height = 22
        ws2.cell(row=idx, column=1, value=stt).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=idx, column=2, value=title).alignment = Alignment(horizontal="left", vertical="center")
        c_val = ws2.cell(row=idx, column=3, value=val)
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.font = font_data_bold
        ws2.cell(row=idx, column=4, value=pct).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=idx, column=5, value=note).alignment = Alignment(horizontal="left", vertical="center")

        for c_idx in range(1, 6):
            cell_item = ws2.cell(row=idx, column=c_idx)
            cell_item.border = cell_border
            if stt == "7" and overdue > 0:
                cell_item.fill = fill_overdue
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    # Bảng phân công chi tiết theo KSV
    ksv_stats = {}
    for r in records:
        ksv = r.get("ksv_chinh", "").strip() or "Chưa phân công KSV"
        if ksv not in ksv_stats:
            ksv_stats[ksv] = {"total": 0, "overdue": 0, "urgent": 0, "solved": 0}
        ksv_stats[ksv]["total"] += 1
        if "🔴" in r.get("muc_do_khan_cap", ""):
            ksv_stats[ksv]["overdue"] += 1
        elif "🟠" in r.get("muc_do_khan_cap", ""):
            ksv_stats[ksv]["urgent"] += 1
        elif "✔" in r.get("muc_do_khan_cap", ""):
            ksv_stats[ksv]["solved"] += 1

    sorted_ksv = sorted(ksv_stats.keys(), key=lambda x: ksv_stats[x]["total"], reverse=True)

    r_ksv_start = 13
    ws2.cell(row=r_ksv_start, column=1, value="BẢNG THỐNG KÊ CHI TIẾT KHỐI LƯỢNG TIN BÁO THEO TỪNG KIỂM SÁT VIÊN").font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    ws2.row_dimensions[r_ksv_start + 1].height = 24

    h_ksv = ["STT", "Họ và Tên Kiểm Sát Viên", "Tổng Số Tin Thụ Lý", "Quá Hạn (🔴)", "Khẩn Cấp (🟠)", "Đã Giải Quyết (✔)", "Tỷ Lệ %"]
    for i, h in enumerate(h_ksv, start=1):
        c = ws2.cell(row=r_ksv_start + 1, column=i)
        c.value = h
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws2.column_dimensions["F"].width = 18
    ws2.column_dimensions["G"].width = 12

    for k_idx, ksv_name in enumerate(sorted_ksv, start=1):
        curr_row = r_ksv_start + 1 + k_idx
        ws2.row_dimensions[curr_row].height = 20
        st = ksv_stats[ksv_name]
        pct = (st["total"] / total_tb * 100) if total_tb else 0
        vals = [k_idx, ksv_name, st["total"], st["overdue"], st["urgent"], st["solved"], f"{pct:.1f}%"]
        for c_idx, val in enumerate(vals, start=1):
            cell_item = ws2.cell(row=curr_row, column=c_idx, value=val)
            cell_item.border = cell_border
            align = "left" if c_idx == 2 else "center"
            cell_item.alignment = Alignment(horizontal=align, vertical="center")
            if c_idx == 4 and st["overdue"] > 0:
                cell_item.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    # Sheet 3: Ma Trận Pivot Table VKS Thụ Lý × Kết Quả Giải Quyết
    ws3 = wb_out.create_sheet(title="Ma_Tran_Pivot_Tin_Bao")
    ws3.views.sheetView[0].showGridLines = True

    matrix, outcomes, sorted_units = build_pivot_tin_bao_summary(records)

    ws3.merge_cells("A1:J1")
    t3 = ws3["A1"]
    t3.value = "MA TRẬN PIVOT TABLE: ĐƠN VỊ VKS THỤ LÝ × KẾT QUẢ GIẢI QUYẾT TIN BÁO"
    t3.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    t3.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    pivot_headers = ["STT", "Đơn vị VKS thụ lý", "Tổng số tin"] + outcomes + ["Tỷ lệ %"]
    ws3.row_dimensions[3].height = 24
    for i, h in enumerate(pivot_headers, start=1):
        c = ws3.cell(row=3, column=i)
        c.value = h
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 32
    ws3.column_dimensions["C"].width = 14
    for i in range(4, 4 + len(outcomes)):
        ws3.column_dimensions[get_column_letter(i)].width = 16
    ws3.column_dimensions[get_column_letter(4 + len(outcomes))].width = 12

    for u_idx, u in enumerate(sorted_units, start=1):
        curr_row = 3 + u_idx
        ws3.row_dimensions[curr_row].height = 20
        tot = matrix[u]["_total"]
        pct = (tot / total_tb * 100) if total_tb else 0
        vals = [u_idx, u, tot] + [matrix[u].get(k, 0) for k in outcomes] + [f"{pct:.1f}%"]
        for c_idx, val in enumerate(vals, start=1):
            cell_item = ws3.cell(row=curr_row, column=c_idx, value=val)
            cell_item.border = cell_border
            align = "left" if c_idx == 2 else "center"
            cell_item.alignment = Alignment(horizontal=align, vertical="center")

    tot_row_idx = 4 + len(sorted_units)
    ws3.row_dimensions[tot_row_idx].height = 24
    tot_vals = ["∑", "TỔNG CỘNG CHUNG", total_tb] + [sum(matrix[u].get(k, 0) for u in sorted_units) for k in outcomes] + ["100.0%"]
    for c_idx, val in enumerate(tot_vals, start=1):
        cell_item = ws3.cell(row=tot_row_idx, column=c_idx, value=val)
        cell_item.border = cell_border
        cell_item.font = font_data_bold
        cell_item.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        align = "left" if c_idx == 2 else "center"
        cell_item.alignment = Alignment(horizontal=align, vertical="center")

    wb_out.save(out_filepath)
    return out_filepath
'''

content = content[:old_export_tin_bao_start] + new_export_tin_bao + content[old_export_tin_bao_end:]
print("Updated export_tin_bao_to_excel")

# 4. Nâng cấp on_reconcile_multiple_files và on_export_excel
old_on_recon_start = content.find('    def on_reconcile_multiple_files(self):')
old_on_recon_end = content.find('\nif __name__ == "__main__":', old_on_recon_start)
assert old_on_recon_start != -1 and old_on_recon_end != -1

new_on_reconcile = '''    def on_reconcile_multiple_files(self):
        """Mở nhiều file Excel QLAHS cùng lúc (không giới hạn 2 file),
        tự động phân tích cấu trúc, nhận diện xem là Vụ án hay Tin báo,
        rà soát đối chiếu chéo, bù đắp làm giàu dữ liệu và hợp nhất theo đúng định dạng tương ứng."""
        if openpyxl is None:
            messagebox.showerror("Thiếu thư viện", "Vui lòng cài đặt openpyxl: pip install openpyxl")
            return

        fpaths = filedialog.askopenfilenames(
            title="Chọn các file Excel QLAHS cần đối chiếu & hợp nhất (Giữ Ctrl hoặc Shift để chọn nhiều file)",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not fpaths:
            return

        fpaths = list(fpaths)
        n_files = len(fpaths)

        if n_files == 1:
            ans = messagebox.askyesno(
                "Đang chọn 1 file",
                f"Bạn vừa chọn 1 file:\\n{os.path.basename(fpaths[0])}\\n\\n"
                "Bạn có muốn chọn thêm các file khác để đối chiếu chéo không?\\n"
                "(Bấm 'Yes' để chọn thêm file, bấm 'No' để nạp luôn file này)"
            )
            if ans:
                fpaths2 = filedialog.askopenfilenames(
                    title="Chọn thêm các file Excel QLAHS khác để đối chiếu",
                    filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
                )
                if fpaths2:
                    for fp in fpaths2:
                        if fp not in fpaths:
                            fpaths.append(fp)
                    n_files = len(fpaths)

        self.lbl_status.config(
            text=f"⏳ Đang đọc và đối chiếu {n_files} file QLAHS...",
            fg="#D97706"
        )
        self.update_idletasks()

        def _worker(update_cb):
            parsed_files = []
            for i, fp in enumerate(fpaths):
                fname = os.path.basename(fp)
                update_cb(
                    f"Đang đọc tập tin ({i+1}/{n_files}): {fname}...",
                    f"Phân tích cấu trúc dữ liệu..."
                )
                fmt, recs = load_excel_and_parse(fp, update_cb)
                parsed_files.append((fname, fmt, recs))

            # Tự động nhận diện loại dữ liệu đối chiếu: Tin Báo hay Vụ Án
            tin_bao_count = sum(1 for _, fmt, _ in parsed_files if fmt == "TIN_BAO")
            is_tin_bao_mode = (tin_bao_count >= len(parsed_files) / 2)

            if is_tin_bao_mode:
                update_cb(
                    f"Đang so khớp chéo & bù đắp thông tin {n_files} tập tin Tin báo...",
                    f"Tổng cộng {sum(len(r) for _, _, r in parsed_files):,} tin báo đầu vào"
                )
                merged, match_n, file_stats = reconcile_and_merge_multiple_tin_bao_lists(parsed_files, update_cb)
                update_cb(
                    f"Đang chuẩn bị hiển thị danh sách 36 cột hợp nhất...",
                    f"Tổng hợp trọn vẹn {len(merged):,} tin báo"
                )
                return "TIN_BAO", merged, match_n, file_stats, parsed_files
            else:
                update_cb(
                    f"Đang so khớp chéo & bù đắp thông tin {n_files} tập tin Vụ án...",
                    f"Tổng cộng {sum(len(r) for _, _, r in parsed_files):,} vụ án đầu vào"
                )
                merged, match_n, file_stats = reconcile_and_merge_multiple_case_lists(parsed_files, update_cb)
                update_cb(
                    f"Đang chuẩn bị hiển thị danh sách 27 cột...",
                    f"Tổng hợp trọn vẹn {len(merged):,} vụ án"
                )
                return "CASE", merged, match_n, file_stats, parsed_files

        def _on_success(res):
            mode, merged, match_n, file_stats, parsed_files = res
            self.data_format = mode
            self.records = merged
            self.refresh_views()

            if mode == "TIN_BAO":
                unit_label = "tin báo về tội phạm"
                action_title = "RÀ SOÁT & HỢP NHẤT TIN BÁO TỐ GIÁC TỘI PHẠM THÀNH CÔNG"
                cols_msg = "36 cột chuẩn hóa chi tiết (kèm cột nguồn dữ liệu đối chiếu)"
            else:
                unit_label = "vụ án"
                action_title = "RÀ SOÁT & HỢP NHẤT VỤ ÁN THÀNH CÔNG"
                cols_msg = "27 cột chuẩn hóa chi tiết (kèm cột nguồn dữ liệu đối chiếu)"

            self.lbl_status.config(
                text=f"✓ Đã hợp nhất thành công {len(merged):,} {unit_label} từ {len(fpaths)} file QLAHS",
                fg="#166534"
            )

            lines = [
                f"⚡ {action_title} ({len(fpaths)} TẬP TIN QLAHS):\\n",
                f"• Tổng số {unit_label} sau hợp nhất: {len(merged):,} bản ghi",
                f"• Tổng số lượt khớp nối & bù đắp thông tin: {match_n:,} lượt\\n",
                "Chi tiết từng file nguồn đã nạp:"
            ]
            for s in file_stats:
                lines.append(f"  - {s['fname']}: {s['total']:,} ({s['new']:,} mới, {s['matched']:,} khớp nối)")

            lines.append(
                f"\\nToàn bộ thông tin còn thiếu giữa các file đã được tự động bù đắp đầy đủ nhất có thể!\\n"
                f"Bảng dữ liệu {cols_msg} đã được cập nhật hiển thị.\\n"
                f"Bạn có thể nhấn nút '📥 3. Xuất Excel' để tải tệp kết quả hoàn chỉnh."
            )
            messagebox.showinfo("Hoàn thành đối chiếu đa tập tin", "\\n".join(lines))

        self.run_async_task(
            title=f"Đang đối chiếu & hợp nhất {n_files} File QLAHS",
            loading_msg=f"Hệ thống đang mở và so khớp {n_files} tập tin Excel...",
            worker_fn=_worker,
            on_success=_on_success,
            on_error=lambda err: messagebox.showerror("Lỗi đối chiếu", f"Không thể đối chiếu các file:\\n{err}")
        )

    # Alias tương thích ngược
    on_reconcile_two_files = on_reconcile_multiple_files
'''

content = content[:old_on_recon_start] + new_on_reconcile + content[old_on_recon_end:]
print("Updated on_reconcile_multiple_files")

with open('Phan_Mem_Chuan_Hoa_QLAHS_Desktop.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Patch applied successfully!")
