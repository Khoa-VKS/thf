import re

patch_code = '''
# ==============================================================================
# HÀM VẼ BIỂU ĐỒ TIẾN ĐỘ TRỰC QUAN (VISUAL PROGRESS & DISTRIBUTION CANVAS)
# ==============================================================================
def draw_visual_distribution_chart(canvas, width, height, segments):
    """
    Vẽ thanh phân bổ tỷ lệ trực quan (Stacked Horizontal Distribution Bar)
    segments: danh sách các tuple (nhãn, số_lượng, mã_màu)
    """
    try:
        canvas.delete("all")
        total = sum(c for _, c, _ in segments)
        if total <= 0:
            canvas.create_text(width // 2, height // 2, text="Chưa có dữ liệu thống kê tiến độ", fill="#64748B", font=("Segoe UI", 9, "italic"))
            return

        pad_x = 8
        bar_y1 = 6
        bar_y2 = 22
        usable_w = max(10, width - (2 * pad_x))

        cur_x = pad_x
        for lbl, count, color in segments:
            if count <= 0:
                continue
            seg_w = max(3, int((count / total) * usable_w))
            canvas.create_rectangle(cur_x, bar_y1, cur_x + seg_w, bar_y2, fill=color, outline="#FFFFFF", width=1)
            cur_x += seg_w

        # Vẽ Chú thích (Legend)
        leg_x = pad_x
        leg_y = bar_y2 + 8
        for lbl, count, color in segments:
            pct = (count / total * 100) if total else 0.0
            txt = f"{lbl}: {count:,} ({pct:.1f}%)"
            canvas.create_oval(leg_x, leg_y + 2, leg_x + 9, leg_y + 11, fill=color, outline="")
            t_id = canvas.create_text(leg_x + 13, leg_y + 6, text=txt, anchor="w", font=("Segoe UI", 8), fill="#1E293B")
            bbox = canvas.bbox(t_id)
            leg_x = (bbox[2] + 16) if bbox else (leg_x + 110)
    except Exception:
        pass


# ==============================================================================
# BÁO CÁO TRÍCH YẾU NGHIỆP VỤ (EXECUTIVE BRIEF GENERATOR)
# ==============================================================================
def generate_executive_brief_text(records, data_format="CASE", config=None):
    """
    Tự động sinh đoạn văn bản trích yếu số liệu báo cáo nghiệp vụ chuẩn văn phong ngành Kiểm sát.
    Sẵn sàng sao chép nhanh (Ctrl+C) hoặc lưu tệp Word/Text báo cáo Lãnh đạo Viện / Phòng.
    """
    cfg = config or (CURRENT_CONFIG if "CURRENT_CONFIG" in globals() else DEFAULT_APP_CONFIG)
    agency = cfg.get("agency_name", "VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH")
    author = cfg.get("author_title", "Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM")
    thresholds = cfg.get("warning_thresholds", {"urgent_days": 30, "very_urgent_days": 60, "warning_days": 90})

    now_str = datetime.now().strftime("%d/%m/%Y lúc %H:%M")
    is_tin_bao = (data_format == "TIN_BAO")
    label_type = "Nguồn tin về tội phạm (Tin báo & Tố giác)" if is_tin_bao else "Hồ sơ vụ án hình sự"

    total = len(records)
    if total == 0:
        return f"{agency}\\nBÁO CÁO TRÍCH YẾU NGHIỆP VỤ\\n(Chưa có dữ liệu nào được nạp vào phần mềm)"

    # Phân loại tiến độ & trạng thái
    resolved = sum(1 for r in records if any(k in str(r.get("ket_qua_giai_quyet") or r.get("trang_thai_vu_an") or r.get("trang_thai") or "").lower() for k in ["đã giải quyết", "kết luận", "đình chỉ", "khởi tố"]))
    suspended = sum(1 for r in records if any(k in str(r.get("ket_qua_giai_quyet") or r.get("muc_do_khan_cap") or r.get("trang_thai_vu_an") or "").lower() for k in ["tạm đình chỉ", "đình nã", "truy nã"]))
    in_progress = max(0, total - resolved - suspended)

    # Thống kê cảnh báo thời hiệu & thời hạn
    overdue = sum(1 for r in records if any("quá hạn" in str(r.get(f) or "").lower() for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    red = sum(1 for r in records if any("🔴" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    orange = sum(1 for r in records if any("🟠" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    yellow = sum(1 for r in records if any("🟡" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))
    green = sum(1 for r in records if any("🟢" in str(r.get(f) or "") for f in ["muc_do_khan_cap", "canh_bao_thoi_hieu"]))

    # Danh sách các hồ sơ khẩn cấp (Overdue, Red, Orange)
    urgent_records = [
        r for r in records
        if any(sym in (str(r.get("muc_do_khan_cap") or "") + " " + str(r.get("canh_bao_thoi_hieu") or "")).lower() for sym in ["🔴", "quá hạn", "khẩn cấp"])
    ]

    # Phân bổ theo KSV
    ksv_map = {}
    for r in records:
        ksv = r.get("ksv_chinh") or r.get("can_bo_giai_quyet") or "Chưa phân công KSV"
        ksv_map[ksv] = ksv_map.get(ksv, 0) + 1
    top_ksv = sorted(ksv_map.items(), key=lambda x: x[1], reverse=True)[:8]

    # Phân bổ theo Đơn vị VKS thụ lý
    unit_map = {}
    for r in records:
        u = r.get("vks_thu_ly") or "Chưa rõ đơn vị"
        unit_map[u] = unit_map.get(u, 0) + 1
    top_units = sorted(unit_map.items(), key=lambda x: x[1], reverse=True)

    lines = []
    lines.append(agency.upper())
    lines.append("=" * 78)
    lines.append("BÁO CÁO TRÍCH YẾU TỔNG HỢP TIẾN ĐỘ & THỜI HIỆU TRUY CỨU TRÁCH NHIỆM HÌNH SỰ")
    lines.append(f"(Số liệu tổng hợp tự động từ phần mềm QLAHS tính đến {now_str})")
    lines.append("-" * 78)
    lines.append("")

    lines.append(f"I. TỔNG QUAN TÌNH HÌNH QUẢN LÝ {label_type.upper()}:")
    lines.append(f"  • Tổng số hồ sơ đang quản lý theo dõi: {total:,} hồ sơ (100.0%)")
    lines.append(f"  • Hồ sơ đang trong thời hạn giải quyết: {in_progress:,} hồ sơ ({in_progress / total * 100:.1f}%)")
    lines.append(f"  • Hồ sơ đã giải quyết / kết luận xử lý: {resolved:,} hồ sơ ({resolved / total * 100:.1f}%)")
    lines.append(f"  • Hồ sơ đang tạm đình chỉ giải quyết: {suspended:,} hồ sơ ({suspended / total * 100:.1f}%)")
    lines.append("")

    lines.append("II. THEO DÕI TIẾN ĐỘ & THỜI HIỆU TRUY CỨU TNHS (ĐẶC BIỆT CHÚ Ý):")
    lines.append(f"  • ⚠️ Đã quá thời hạn giải quyết / Quá thời hiệu TNHS: {overdue:,} hồ sơ")
    lines.append(f"  • 🔴 Khẩn cấp / Cấp bách (Thời hạn còn ≤ {thresholds.get('urgent_days', 30)} ngày): {red:,} hồ sơ")
    lines.append(f"  • 🟠 Rất gấp (Thời hạn còn 31 - {thresholds.get('very_urgent_days', 60)} ngày): {orange:,} hồ sơ")
    lines.append(f"  • 🟡 Sắp hết hạn (Thời hạn còn 61 - {thresholds.get('warning_days', 90)} ngày): {yellow:,} hồ sơ")
    lines.append(f"  • 🟢 Trong hạn bình thường (Thời hạn còn > {thresholds.get('warning_days', 90)} ngày): {green:,} hồ sơ")
    lines.append("")

    if urgent_records:
        lines.append(f"III. DANH SÁCH {min(len(urgent_records), 20)} HỒ SƠ KHẨN CẤP / QUÁ HẠN CẦN TẬP TRUNG CHỈ ĐẠO XỬ LÝ:")
        for idx, r in enumerate(urgent_records[:20], 1):
            code = r.get("ma_tin_bao") or r.get("ma_vu_an") or f"STT {idx}"
            title = r.get("noi_dung_tin_bao") or r.get("ten_vu_an") or "Chưa có tên/nội dung"
            title = (title[:75] + "...") if len(title) > 75 else title
            ksv = r.get("ksv_chinh") or r.get("can_bo_giai_quyet") or "Chưa phân công"
            unit = r.get("vks_thu_ly") or ""
            due = r.get("thoi_han_den_ngay") or r.get("ngay_het_thoi_hieu") or "Chưa rõ"
            urg = r.get("muc_do_khan_cap") or r.get("canh_bao_thoi_hieu") or "Khẩn cấp"
            lines.append(f"  {idx}. [{code}] {title}")
            lines.append(f"     -> Đơn vị: {unit} | KSV: {ksv}")
            lines.append(f"     -> Hạn / Thời hiệu: {due} | Tiến độ: {urg}")
        if len(urgent_records) > 20:
            lines.append(f"  ... và còn {len(urgent_records) - 20} hồ sơ khẩn cấp khác (xem chi tiết tại bảng xuất Excel).")
        lines.append("")

    if top_units:
        lines.append("IV. CƠ CẤU HỒ SƠ THEO ĐƠN VI VKS THỤ LÝ:")
        for u_name, u_cnt in top_units[:10]:
            lines.append(f"  • {u_name}: {u_cnt:,} hồ sơ ({u_cnt / total * 100:.1f}%)")
        lines.append("")

    if top_ksv:
        lines.append("V. TÌNH HÌNH PHÂN CÔNG THEO KIỂM SÁT VIÊN THỤ LÝ (TOP 8):")
        for ksv_name, ksv_cnt in top_ksv:
            lines.append(f"  • KSV {ksv_name}: {ksv_cnt:,} hồ sơ ({ksv_cnt / total * 100:.1f}%)")
        lines.append("")

    lines.append("=" * 78)
    lines.append(f"Người lập trích yếu báo cáo: {author}")
    lines.append("Phần Mềm Chuẩn Hóa & Đối Chiếu Dữ Liệu QLAHS (Version 1.0 Offline)")
    return "\\n".join(lines)


# ==============================================================================
# HỆ THỐNG SO SÁNH BIẾN ĐỘNG 2 KỲ BÁO CÁO (DIFF ANALYSIS ENGINE)
# ==============================================================================
def compare_two_qlahs_periods(records_old, records_new, data_format=None):
    """
    So sánh biến động dữ liệu giữa 2 kỳ báo cáo QLAHS:
    - added: Danh sách hồ sơ mới phát sinh trong kỳ
    - resolved: Danh sách hồ sơ đã được giải quyết / kết thúc
    - worsened: Danh sách hồ sơ chuyển mức khẩn cấp về thời hạn / thời hiệu TNHS
    - continued: Danh sách hồ sơ tiếp tục tồn đọng / đang giải quyết
    """
    def norm_k(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

    def get_urgency_level(r):
        u1 = str(r.get("muc_do_khan_cap") or "").lower()
        u2 = str(r.get("canh_bao_thoi_hieu") or "").lower()
        cb = u1 + " " + u2
        if "quá hạn" in cb:
            return 5
        if "🔴" in cb:
            return 4
        if "🟠" in cb:
            return 3
        if "🟡" in cb:
            return 2
        if "🟢" in cb:
            return 1
        return 0

    old_by_code = {}
    old_by_sub = {}
    for r in records_old:
        code = norm_k(r.get("ma_tin_bao") or r.get("ma_vu_an"))
        if code:
            old_by_code[code] = r
        sub = norm_k(str(r.get("so_qd_khoi_to") or r.get("so_thu_ly") or "") + str(r.get("vks_thu_ly") or ""))
        if sub:
            old_by_sub[sub] = r

    matched_old_keys = set()
    added = []
    worsened = []
    continued = []

    for r_new in records_new:
        c_k = norm_k(r_new.get("ma_tin_bao") or r_new.get("ma_vu_an"))
        s_k = norm_k(str(r_new.get("so_qd_khoi_to") or r_new.get("so_thu_ly") or "") + str(r_new.get("vks_thu_ly") or ""))

        r_old = None
        if c_k and c_k in old_by_code:
            r_old = old_by_code[c_k]
            matched_old_keys.add(("code", c_k))
        elif s_k and s_k in old_by_sub:
            r_old = old_by_sub[s_k]
            matched_old_keys.add(("sub", s_k))

        if r_old is None:
            added.append(r_new)
        else:
            lvl_old = get_urgency_level(r_old)
            lvl_new = get_urgency_level(r_new)
            if lvl_new > lvl_old and lvl_new >= 2:
                rec_w = dict(r_new)
                rec_w["_old_urgency"] = r_old.get("muc_do_khan_cap") or r_old.get("canh_bao_thoi_hieu") or "Bình thường"
                rec_w["_new_urgency"] = r_new.get("muc_do_khan_cap") or r_new.get("canh_bao_thoi_hieu") or "Khẩn cấp"
                worsened.append(rec_w)
            else:
                continued.append(r_new)

    resolved = []
    for r_old in records_old:
        c_k = norm_k(r_old.get("ma_tin_bao") or r_old.get("ma_vu_an"))
        s_k = norm_k(str(r_old.get("so_qd_khoi_to") or r_old.get("so_thu_ly") or "") + str(r_old.get("vks_thu_ly") or ""))
        is_m = False
        if c_k and ("code", c_k) in matched_old_keys:
            is_m = True
        elif s_k and ("sub", s_k) in matched_old_keys:
            is_m = True
        if not is_m:
            resolved.append(r_old)

    stats = {
        "total_old": len(records_old),
        "total_new": len(records_new),
        "n_added": len(added),
        "n_resolved": len(resolved),
        "n_worsened": len(worsened),
        "n_continued": len(continued),
        "net_change": len(records_new) - len(records_old)
    }

    return {
        "stats": stats,
        "added": added,
        "resolved": resolved,
        "worsened": worsened,
        "continued": continued,
        "data_format": data_format or "CASE"
    }


def export_diff_to_excel(diff_data, out_filepath, title_old="Kỳ Trước", title_new="Kỳ Này"):
    """
    Xuất báo cáo phân tích biến động giữa 2 kỳ thành tệp Excel chuyên nghiệp đa Sheet có Biểu đồ BarChart.
    """
    if openpyxl is None:
        raise RuntimeError("Thư viện openpyxl chưa được cài đặt.")

    wb = openpyxl.Workbook()
    stats = diff_data["stats"]
    is_tin_bao = (diff_data.get("data_format") == "TIN_BAO")
    cols_def = COLUMNS_TIN_BAO_35 if is_tin_bao else COLUMNS_CASE_24

    font_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
    font_regular = Font(name="Segoe UI", size=9)
    thin_border = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
                         top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))

    # Sheet 1: Tổng quan biến động
    ws1 = wb.active
    ws1.title = "Tong_Quan_Bien_Dong"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:G1")
    t1 = ws1["A1"]
    t1.value = "BÁO CÁO PHÂN TÍCH SO SÁNH BIẾN ĐỘNG GIỮA 2 KỲ BÁO CÁO QLAHS"
    t1.font = font_title
    t1.fill = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32

    ws1.merge_cells("A2:G2")
    t2 = ws1["A2"]
    t2.value = f"Kỳ trước: {title_old}  |  Kỳ này: {title_new}  •  Thời điểm đối chiếu: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    t2.font = Font(name="Segoe UI", size=9, italic=True, color="475569")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20

    # Bảng chỉ số
    headers_stat = ["STT", "Chỉ Tiêu So Sánh Biến Động Nghiệp Vụ", "Số Lượng", "Tỷ Lệ %", "Đánh Giá Chiều Hướng Biến Động"]
    ws1.row_dimensions[4].height = 24
    for i, h in enumerate(headers_stat, 1):
        c = ws1.cell(row=4, column=i, value=h)
        c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    tot_n = stats["total_new"] or 1
    stat_rows = [
        ("1", f"Tổng số hồ sơ Kỳ trước ({title_old})", stats["total_old"], "-", "Mốc đối chiếu ban đầu"),
        ("2", f"Tổng số hồ sơ Kỳ này ({title_new})", stats["total_new"], "100.0%", "Hiện trạng theo dõi hiện tại"),
        ("3", "Chênh lệch ròng (Tăng / Giảm)", f"{'+' if stats['net_change'] >= 0 else ''}{stats['net_change']}", f"{abs(stats['net_change']) / (stats['total_old'] or 1) * 100:.1f}%", "Xu hướng khối lượng công việc"),
        ("4", "✨ Số hồ sơ mới thụ lý / Mới phát sinh trong kỳ", stats["n_added"], f"{stats['n_added'] / tot_n * 100:.1f}%", "🟢 Thụ lý mới"),
        ("5", "✅ Số hồ sơ đã giải quyết xong / Kết thúc trong kỳ", stats["n_resolved"], f"{stats['n_resolved'] / (stats['total_old'] or 1) * 100:.1f}%", "🔵 Đã giải quyết / Đạt kết quả"),
        ("6", "⚠️ Số hồ sơ chuyển mức khẩn cấp thời hiệu / Quá hạn", stats["n_worsened"], f"{stats['n_worsened'] / tot_n * 100:.1f}%", "🔴 Nguy cơ quá hạn thời hiệu"),
        ("7", "⏸ Số hồ sơ tiếp tục tồn đọng / Tạm đình chỉ kéo dài", stats["n_continued"], f"{stats['n_continued'] / tot_n * 100:.1f}%", "⚪ Cần tiếp tục theo dõi"),
    ]

    for r_i, (stt, name, val, pct, note) in enumerate(stat_rows, 5):
        ws1.row_dimensions[r_i].height = 20
        c1 = ws1.cell(row=r_i, column=1, value=stt)
        c2 = ws1.cell(row=r_i, column=2, value=name)
        c3 = ws1.cell(row=r_i, column=3, value=val)
        c4 = ws1.cell(row=r_i, column=4, value=pct)
        c5 = ws1.cell(row=r_i, column=5, value=note)
        for c in (c1, c2, c3, c4, c5):
            c.border = thin_border
            c.font = font_regular
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c5.alignment = Alignment(horizontal="left", vertical="center")
        if r_i == 8:
            c2.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
        elif r_i == 9:
            c2.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
        elif r_i == 10 and stats["n_worsened"] > 0:
            c2.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            c3.font = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 46
    ws1.column_dimensions["C"].width = 16
    ws1.column_dimensions["D"].width = 14
    ws1.column_dimensions["E"].width = 30

    # Nhúng biểu đồ cột BarChart trực quan
    if BarChart and Reference:
        try:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = "CƠ CẤU BIẾN ĐỘNG HỒ SƠ GIỮA 2 KỲ BÁO CÁO"
            chart.y_axis.title = "Số lượng hồ sơ"
            chart.x_axis.title = "Nhóm biến động"
            data_ref = Reference(ws1, min_col=3, min_row=7, max_row=11)
            cats_ref = Reference(ws1, min_col=2, min_row=8, max_row=11)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.width = 17
            chart.height = 10
            ws1.add_chart(chart, "G4")
        except Exception:
            pass

    # Hàm phụ ghi danh sách chi tiết từng sheet
    def write_diff_sheet(title, rec_list, header_color):
        ws = wb.create_sheet(title=title)
        ws.views.sheetView[0].showGridLines = True
        ws.row_dimensions[1].height = 26
        # Headers
        ws.cell(row=1, column=1, value="STT").alignment = Alignment(horizontal="center", vertical="center")
        for c_idx, (_, col_name, col_w) in enumerate(cols_def, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(c_idx)].width = max(10, min(col_w // 8, 35))

        for r_idx, r in enumerate(rec_list, start=2):
            ws.row_dimensions[r_idx].height = 20
            for c_idx, (k, _, _) in enumerate(cols_def, start=1):
                val = r.get(k, "")
                if c_idx == 1:
                    val = r_idx - 1
                c = ws.cell(row=r_idx, column=c_idx, value=val)
                c.border = thin_border
                c.font = font_regular
                c.alignment = Alignment(horizontal="center" if c_idx in (1, 2) else "left", vertical="center")

    write_diff_sheet("1.Moi_Thu_Ly_Trong_Ky", diff_data["added"], "047857")
    write_diff_sheet("2.Da_Giai_Quyet_Trong_Ky", diff_data["resolved"], "1D4ED8")
    write_diff_sheet("3.Chuyen_Cap_Khan_Cap", diff_data["worsened"], "B91C1C")
    write_diff_sheet("4.Tiep_Tuc_Ton_Dong", diff_data["continued"][:500], "475569")

    wb.save(out_filepath)
    return out_filepath


# ==============================================================================
# HỆ THỐNG CÁC CỬA SỔ HỘI THOẠI BỔ SUNG (DIALOGS)
# ==============================================================================
class SettingsDialog(tk.Toplevel):
    """
    Cửa sổ Cài đặt cấu hình tùy biến:
    - Tên cơ quan chủ quản
    - Chức danh tác giả
    - Ngưỡng cảnh báo thời hiệu (30, 60, 90 ngày)
    - Danh mục từ điển viết tắt đơn vị (Custom Aliases)
    """
    def __init__(self, parent, config=None, on_save_cb=None):
        super().__init__(parent)
        self.parent = parent
        self.config = dict(config or CURRENT_CONFIG)
        self.on_save_cb = on_save_cb

        self.title("⚙️ CÀI ĐẶT QUY TẮC CHUẨN HÓA & NGƯỠNG CẢNH BÁO QLAHS")
        self.geometry("780x560")
        self.minsize(700, 500)
        self.transient(parent)
        self.configure(bg="#F8FAFC")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="⚙️ CÀI ĐẶT HỆ THỐNG & TÙY BIẾN QUY TẮC NGHIỆP VỤ", font=("Segoe UI", 12, "bold"), fg="#FBBF24", bg="#07172C").pack(anchor="w")
        tk.Label(header, text="Tùy chỉnh tên cơ quan, ngưỡng cảnh báo thời hiệu và danh mục viết tắt đơn vị linh hoạt theo từng địa phương.", font=("Segoe UI", 8, "italic"), fg="#93C5FD", bg="#07172C").pack(anchor="w")

        # Notebook Tabs
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # Tab 1: Cơ quan & Tác giả
        tab1 = tk.Frame(nb, bg="#FFFFFF", padx=16, pady=16)
        nb.add(tab1, text="🏢 Cơ Quan & Tác Giả")

        tk.Label(tab1, text="Tên Cơ quan chủ quản (Xuất hiện trên tiêu đề & báo cáo):", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 4))
        self.ent_agency = ttk.Entry(tab1, font=("Segoe UI", 10), width=65)
        self.ent_agency.insert(0, self.config.get("agency_name", DEFAULT_APP_CONFIG["agency_name"]))
        self.ent_agency.pack(anchor="w", pady=(0, 16))

        tk.Label(tab1, text="Thông tin tác giả / Chức danh người lập báo cáo:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 4))
        self.ent_author = ttk.Entry(tab1, font=("Segoe UI", 10), width=65)
        self.ent_author.insert(0, self.config.get("author_title", DEFAULT_APP_CONFIG["author_title"]))
        self.ent_author.pack(anchor="w", pady=(0, 16))

        info_box = tk.Frame(tab1, bg="#EFF6FF", padx=12, pady=10, relief=tk.SOLID, bd=1)
        info_box.pack(fill=tk.X, pady=(10, 0))
        tk.Label(info_box, text="💡 Ghi chú:", font=("Segoe UI", 9, "bold"), fg="#1D4ED8", bg="#EFF6FF").pack(anchor="w")
        tk.Label(info_box, text="Các thông tin này sẽ được tự động điền vào tiêu đề phần mềm, các bảng báo cáo Word/Text và tiêu đề các file Excel xuất ra.", font=("Segoe UI", 8), fg="#1E40AF", bg="#EFF6FF").pack(anchor="w")

        # Tab 2: Ngưỡng cảnh báo thời hiệu
        tab2 = tk.Frame(nb, bg="#FFFFFF", padx=16, pady=16)
        nb.add(tab2, text="⚠️ Ngưỡng Cảnh Báo Thời Hiệu")

        th = self.config.get("warning_thresholds", DEFAULT_APP_CONFIG["warning_thresholds"])
        tk.Label(tab2, text="THIẾT LẬP MỐC ĐẾM NGƯỢC THỜI HẠN & THỜI HIỆU TRUY CỨU TNHS:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 12))

        # Urgent (Red)
        f_u = tk.Frame(tab2, bg="#FFFFFF")
        f_u.pack(anchor="w", pady=6)
        tk.Label(f_u, text="🔴 Mức KHẨN CẤP / CẤP BÁCH (Đỏ): Thời hạn còn ≤", font=("Segoe UI", 9, "bold"), fg="#991B1B", bg="#FFFFFF").pack(side=tk.LEFT)
        self.sp_urgent = ttk.Spinbox(f_u, from_=1, to=365, width=6, font=("Segoe UI", 9, "bold"))
        self.sp_urgent.set(th.get("urgent_days", 30))
        self.sp_urgent.pack(side=tk.LEFT, padx=6)
        tk.Label(f_u, text="ngày", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)

        # Very Urgent (Orange)
        f_vu = tk.Frame(tab2, bg="#FFFFFF")
        f_vu.pack(anchor="w", pady=6)
        tk.Label(f_vu, text="🟠 Mức RẤT GẤP (Cam): Thời hạn còn ≤", font=("Segoe UI", 9, "bold"), fg="#C2410C", bg="#FFFFFF").pack(side=tk.LEFT)
        self.sp_vurgent = ttk.Spinbox(f_vu, from_=1, to=365, width=6, font=("Segoe UI", 9, "bold"))
        self.sp_vurgent.set(th.get("very_urgent_days", 60))
        self.sp_vurgent.pack(side=tk.LEFT, padx=6)
        tk.Label(f_vu, text="ngày", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)

        # Warning (Yellow)
        f_w = tk.Frame(tab2, bg="#FFFFFF")
        f_w.pack(anchor="w", pady=6)
        tk.Label(f_w, text="🟡 Mức SẮP ĐẾN HẠN (Vàng): Thời hạn còn ≤", font=("Segoe UI", 9, "bold"), fg="#B45309", bg="#FFFFFF").pack(side=tk.LEFT)
        self.sp_warning = ttk.Spinbox(f_w, from_=1, to=365, width=6, font=("Segoe UI", 9, "bold"))
        self.sp_warning.set(th.get("warning_days", 90))
        self.sp_warning.pack(side=tk.LEFT, padx=6)
        tk.Label(f_w, text="ngày (Trên mốc này là 🟢 Bình thường)", font=("Segoe UI", 9), bg="#FFFFFF").pack(side=tk.LEFT)

        # Tab 3: Danh mục viết tắt đơn vị
        tab3 = tk.Frame(nb, bg="#FFFFFF", padx=12, pady=12)
        nb.add(tab3, text="🏷️ Viết Tắt Đơn Vị (Aliases)")

        tk.Label(tab3, text="DANH MỤC QUY TẮC NHẬN DIỆN VIẾT TẮT ĐƠN VỊ:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#1E293B").pack(anchor="w", pady=(0, 6))

        tree_frame = tk.Frame(tab3, bg="#FFFFFF")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree_alias = ttk.Treeview(tree_frame, columns=("alias", "target"), show="headings", height=8)
        self.tree_alias.heading("alias", text="Từ khóa viết tắt (Ví dụ: P2, KV4)")
        self.tree_alias.heading("target", text="Tên đơn vị chuẩn hóa thay thế (Ví dụ: Phòng 2 - VKSND TP.HCM)")
        self.tree_alias.column("alias", width=220, anchor=tk.CENTER)
        self.tree_alias.column("target", width=460, anchor=tk.W)

        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_alias.yview)
        self.tree_alias.configure(yscrollcommand=sb.set)
        self.tree_alias.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        curr_aliases = self.config.get("unit_aliases", DEFAULT_APP_CONFIG["unit_aliases"])
        for a_k, a_v in curr_aliases.items():
            self.tree_alias.insert("", "end", values=(a_k, a_v))

        btn_box = tk.Frame(tab3, bg="#FFFFFF", pady=6)
        btn_box.pack(fill=tk.X)

        tk.Button(btn_box, text="➕ Thêm Từ Khóa", font=("Segoe UI", 8, "bold"), bg="#0284C7", fg="white", padx=10, pady=3, command=self._on_add_alias).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_box, text="🗑️ Xóa Đã Chọn", font=("Segoe UI", 8), bg="#EF4444", fg="white", padx=10, pady=3, command=self._on_del_alias).pack(side=tk.LEFT)

        # Footer Buttons
        footer = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        btn_save = tk.Button(footer, text="💾 Lưu Cài Đặt (Áp dụng ngay)", font=("Segoe UI", 9, "bold"), bg="#15803D", fg="white", padx=16, pady=6, cursor="hand2", command=self._on_save)
        btn_save.pack(side=tk.LEFT, padx=(0, 8))

        btn_reset = tk.Button(footer, text="🔄 Khôi Phục Mặc Định", font=("Segoe UI", 9), bg="#475569", fg="white", padx=12, pady=6, cursor="hand2", command=self._on_reset)
        btn_reset.pack(side=tk.LEFT, padx=(0, 8))

        btn_close = tk.Button(footer, text="Đóng", font=("Segoe UI", 9), bg="#334155", fg="white", padx=14, pady=6, cursor="hand2", command=self.destroy)
        btn_close.pack(side=tk.RIGHT)

    def _on_add_alias(self):
        dlg = tk.Toplevel(self)
        dlg.title("Thêm Quy Tắc Viết Tắt")
        dlg.geometry("450x180")
        dlg.resizable(False, False)
        dlg.configure(bg="#F8FAFC")
        dlg.transient(self)

        tk.Label(dlg, text="Từ khóa viết tắt (Ví dụ: P2, KV4, P10):", font=("Segoe UI", 9, "bold"), bg="#F8FAFC").pack(anchor="w", padx=14, pady=(12, 2))
        ent_k = ttk.Entry(dlg, width=45)
        ent_k.pack(anchor="w", padx=14, pady=(0, 8))

        tk.Label(dlg, text="Tên đơn vị chuẩn hóa đầy đủ:", font=("Segoe UI", 9, "bold"), bg="#F8FAFC").pack(anchor="w", padx=14, pady=(2, 2))
        ent_v = ttk.Entry(dlg, width=45)
        ent_v.pack(anchor="w", padx=14, pady=(0, 10))

        def _do_add():
            k = ent_k.get().strip()
            v = ent_v.get().strip()
            if not k or not v:
                messagebox.showwarning("Thiếu dữ liệu", "Vui lòng nhập đầy đủ từ khóa và tên đơn vị.", parent=dlg)
                return
            self.tree_alias.insert("", "end", values=(k, v))
            dlg.destroy()

        btn_f = tk.Frame(dlg, bg="#F8FAFC")
        btn_f.pack(fill=tk.X, padx=14)
        tk.Button(btn_f, text="Lưu Quy Tắc", font=("Segoe UI", 9, "bold"), bg="#15803D", fg="white", padx=10, pady=4, command=_do_add).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_f, text="Hủy", font=("Segoe UI", 9), bg="#64748B", fg="white", padx=10, pady=4, command=dlg.destroy).pack(side=tk.LEFT)

    def _on_del_alias(self):
        sel = self.tree_alias.selection()
        if not sel:
            messagebox.showinfo("Chọn dòng", "Vui lòng chọn dòng quy tắc cần xóa.", parent=self)
            return
        for s in sel:
            self.tree_alias.delete(s)

    def _on_reset(self):
        if messagebox.askyesno("Khôi phục mặc định", "Anh có chắc chắn muốn khôi phục lại toàn bộ cài đặt gốc ban đầu?", parent=self):
            self.ent_agency.delete(0, tk.END)
            self.ent_agency.insert(0, DEFAULT_APP_CONFIG["agency_name"])
            self.ent_author.delete(0, tk.END)
            self.ent_author.insert(0, DEFAULT_APP_CONFIG["author_title"])
            self.sp_urgent.set(30)
            self.sp_vurgent.set(60)
            self.sp_warning.set(90)
            for item in self.tree_alias.get_children():
                self.tree_alias.delete(item)
            for k, v in DEFAULT_APP_CONFIG["unit_aliases"].items():
                self.tree_alias.insert("", "end", values=(k, v))

    def _on_save(self):
        try:
            urg = int(self.sp_urgent.get().strip() or 30)
            vurg = int(self.sp_vurgent.get().strip() or 60)
            war = int(self.sp_warning.get().strip() or 90)
        except ValueError:
            messagebox.showerror("Lỗi dữ liệu", "Ngưỡng ngày cảnh báo phải là số nguyên dương.", parent=self)
            return

        aliases = {}
        for item in self.tree_alias.get_children():
            vals = self.tree_alias.item(item, "values")
            if vals and len(vals) >= 2:
                aliases[str(vals[0]).strip()] = str(vals[1]).strip()

        new_cfg = {
            "agency_name": self.ent_agency.get().strip() or DEFAULT_APP_CONFIG["agency_name"],
            "author_title": self.ent_author.get().strip() or DEFAULT_APP_CONFIG["author_title"],
            "warning_thresholds": {
                "urgent_days": urg,
                "very_urgent_days": vurg,
                "warning_days": war
            },
            "unit_aliases": aliases
        }

        if save_app_config(new_cfg):
            global CURRENT_CONFIG
            CURRENT_CONFIG = new_cfg
            if self.on_save_cb:
                self.on_save_cb(new_cfg)
            messagebox.showinfo("Thành công", "Đã lưu cài đặt cấu hình thành công!\\nCác quy tắc mới đã được áp dụng ngay lập tức.", parent=self)
            self.destroy()
        else:
            messagebox.showerror("Lỗi", "Không thể lưu tệp cấu hình config_qlahs.json.", parent=self)


class ReportSummaryDialog(tk.Toplevel):
    """
    Cửa sổ Báo cáo trích yếu số liệu nghiệp vụ:
    - Hiển thị văn bản trích yếu chuẩn phong cách ngành Kiểm sát
    - 1-Click Copy toàn bộ nội dung vào Clipboard
    - Lưu tệp văn bản Word (.doc) hoặc Text (.txt)
    """
    def __init__(self, parent, records, data_format="CASE", config=None):
        super().__init__(parent)
        self.parent = parent
        self.records = records or []
        self.data_format = data_format
        self.config = config or CURRENT_CONFIG

        self.title("📝 BÁO CÁO TRÍCH YẾU TỔNG HỢP NGHIỆP VỤ & THỜI HIỆU TRUY CỨU TNHS")
        self.geometry("920x680")
        self.minsize(800, 550)
        self.transient(parent)
        self.configure(bg="#F8FAFC")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="📝 BÁO CÁO TRÍCH YẾU NGHIỆP VỤ & TIẾN ĐỘ THỜI HIỆU TRUY CỨU TNHS", font=("Segoe UI", 12, "bold"), fg="#FBBF24", bg="#07172C").pack(anchor="w")
        tk.Label(header, text="Đoạn văn bản trích yếu số liệu chuẩn hóa, sẵn sàng sao chép nhanh (Ctrl+C) hoặc xuất tệp gửi Lãnh đạo Viện / Lãnh đạo Phòng.", font=("Segoe UI", 8, "italic"), fg="#93C5FD", bg="#07172C").pack(anchor="w")

        # Toolbar
        tb = tk.Frame(self, bg="#FFFFFF", padx=14, pady=8, relief=tk.RIDGE, bd=1)
        tb.pack(fill=tk.X, padx=12, pady=(8, 4))

        btn_copy = tk.Button(tb, text="📋 Sao Chép Toàn Bộ Trích Yếu (Copy)", font=("Segoe UI", 9, "bold"), bg="#0284C7", fg="white", padx=14, pady=5, cursor="hand2", command=self._on_copy)
        btn_copy.pack(side=tk.LEFT, padx=(0, 8))

        btn_txt = tk.Button(tb, text="💾 Lưu File Text (.txt)", font=("Segoe UI", 9), bg="#15803D", fg="white", padx=12, pady=5, cursor="hand2", command=self._on_save_txt)
        btn_txt.pack(side=tk.LEFT, padx=(0, 8))

        btn_doc = tk.Button(tb, text="📄 Lưu File Word (.doc)", font=("Segoe UI", 9), bg="#4338CA", fg="white", padx=12, pady=5, cursor="hand2", command=self._on_save_doc)
        btn_doc.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_copy_toast = tk.Label(tb, text="", font=("Segoe UI", 9, "bold"), fg="#16A34A", bg="#FFFFFF")
        self.lbl_copy_toast.pack(side=tk.LEFT, padx=10)

        # Main Text Widget
        text_frame = tk.Frame(self, bg="#F8FAFC", padx=12, pady=4)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_content = scrolledtext.ScrolledText(text_frame, font=("Segoe UI", 10), bg="#FFFFFF", fg="#0F172A", wrap=tk.WORD, relief=tk.SOLID, bd=1, padx=12, pady=12)
        self.txt_content.pack(fill=tk.BOTH, expand=True)

        report_text = generate_executive_brief_text(self.records, self.data_format, self.config)
        self.txt_content.insert(tk.END, report_text)

        # Footer
        footer = tk.Frame(self, bg="#07172C", padx=16, pady=8)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(footer, text="⚡ Phát triển bởi: Trần Văn Khoa - Kiểm tra viên / VKSND TP.HCM", font=("Segoe UI", 8, "italic"), fg="#94A3B8", bg="#07172C").pack(side=tk.LEFT)
        tk.Button(footer, text="Đóng", font=("Segoe UI", 9), bg="#475569", fg="white", padx=14, pady=4, cursor="hand2", command=self.destroy).pack(side=tk.RIGHT)

    def _on_copy(self):
        text = self.txt_content.get("1.0", tk.END).strip()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.lbl_copy_toast.config(text="✔ Đã sao chép vào bộ nhớ đệm (Clipboard)!")
        self.after(3500, lambda: self.lbl_copy_toast.config(text=""))

    def _on_save_txt(self):
        f = filedialog.asksaveasfilename(
            parent=self,
            title="Lưu Báo Cáo Trích Yếu",
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt"), ("All Files", "*.*")],
            initialfile=f"Trich_Yeu_Bao_Cao_QLAHS_{datetime.now().strftime('%d%m%Y_%H%h%M')}.txt"
        )
        if f:
            try:
                with open(f, "w", encoding="utf-8") as fp:
                    fp.write(self.txt_content.get("1.0", tk.END))
                messagebox.showinfo("Thành công", f"Đã lưu tệp báo cáo trích yếu tại:\\n{f}", parent=self)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu tệp: {e}", parent=self)

    def _on_save_doc(self):
        f = filedialog.asksaveasfilename(
            parent=self,
            title="Lưu Báo Cáo Trích Yếu Word",
            defaultextension=".doc",
            filetypes=[("Word Document", "*.doc"), ("All Files", "*.*")],
            initialfile=f"Trich_Yeu_Bao_Cao_QLAHS_{datetime.now().strftime('%d%m%Y_%H%h%M')}.doc"
        )
        if f:
            try:
                body = self.txt_content.get("1.0", tk.END).replace("\\n", "<br/>\\n")
                doc_html = f"""<html>
<head><meta charset="utf-8"><title>Trích yếu Báo cáo</title>
<style>body {{ font-family: 'Times New Roman', serif; font-size: 13pt; line-height: 1.4; margin: 2cm; }}</style>
</head>
<body>{body}</body></html>"""
                with open(f, "w", encoding="utf-8") as fp:
                    fp.write(doc_html)
                messagebox.showinfo("Thành công", f"Đã lưu tệp văn bản Word tại:\\n{f}", parent=self)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể lưu tệp: {e}", parent=self)


class DiffAnalysisDialog(tk.Toplevel):
    """
    Cửa sổ Phân tích so sánh biến động giữa 2 kỳ báo cáo QLAHS:
    - Cho phép chọn File Kỳ Trước và File Kỳ Này
    - Tự động so khớp theo Mã vụ án/tin báo và Số QĐ/Số thụ lý
    - Phân tích 4 nhóm: Mới thụ lý, Đã giải quyết, Chuyển cấp khẩn cấp, Tiếp tục tồn đọng
    - Bảng trực quan + Xuất Excel phân tích đa Sheet có biểu đồ BarChart
    """
    def __init__(self, parent, current_records=None, current_format="CASE"):
        super().__init__(parent)
        self.parent = parent
        self.current_records = current_records or []
        self.current_format = current_format

        self.records_old = []
        self.records_new = list(self.current_records)
        self.name_old = ""
        self.name_new = "Dữ liệu hiện tại trên màn hình" if self.current_records else ""
        self.diff_res = None

        self.title("🔄 SO SÁNH BIẾN ĐỘNG DỮ LIỆU GIỮA 2 KỲ BÁO CÁO QLAHS")
        self.geometry("1100x720")
        self.minsize(950, 600)
        self.transient(parent)
        self.configure(bg="#F8FAFC")

        icon_path = get_resource_path("logo_vks.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self._build_ui()
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        # 1. Header Banner
        header = tk.Frame(self, bg="#07172C", padx=16, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="🏛️  VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH", font=("Segoe UI", 9, "bold"), fg="#FBBF24", bg="#07172C").pack(anchor="w")
        tk.Label(header, text="🔄 PHÂN TÍCH SO SÁNH BIẾN ĐỘNG DỮ LIỆU GIỮA 2 KỲ BÁO CÁO QLAHS", font=("Segoe UI", 13, "bold"), fg="#FFFFFF", bg="#07172C").pack(anchor="w", pady=(2, 0))

        # 2. File Selection Bar
        sel_box = tk.Frame(self, bg="#FFFFFF", padx=16, pady=12, relief=tk.RIDGE, bd=1)
        sel_box.pack(fill=tk.X, padx=12, pady=(8, 4))

        # Row 1: File Kỳ 1 (Cũ)
        r1 = tk.Frame(sel_box, bg="#FFFFFF")
        r1.pack(fill=tk.X, pady=(0, 6))
        tk.Button(r1, text="📂 1. Chọn File Kỳ Trước (Cũ)", font=("Segoe UI", 9, "bold"), bg="#1D4ED8", fg="white", padx=12, pady=4, cursor="hand2", command=self._on_choose_old).pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_old = tk.Label(r1, text="Chưa chọn tệp dữ liệu kỳ trước...", font=("Segoe UI", 9, "italic"), fg="#64748B", bg="#FFFFFF")
        self.lbl_old.pack(side=tk.LEFT)

        # Row 2: File Kỳ 2 (Mới)
        r2 = tk.Frame(sel_box, bg="#FFFFFF")
        r2.pack(fill=tk.X, pady=(0, 8))
        tk.Button(r2, text="📂 2. Chọn File Kỳ Này (Mới)", font=("Segoe UI", 9, "bold"), bg="#0D9488", fg="white", padx=12, pady=4, cursor="hand2", command=self._on_choose_new).pack(side=tk.LEFT, padx=(0, 10))
        txt_new = f"{self.name_new} ({len(self.records_new):,} hồ sơ)" if self.records_new else "Chưa chọn tệp dữ liệu kỳ này..."
        self.lbl_new = tk.Label(r2, text=txt_new, font=("Segoe UI", 9, "bold" if self.records_new else "italic"), fg="#047857" if self.records_new else "#64748B", bg="#FFFFFF")
        self.lbl_new.pack(side=tk.LEFT)

        # Row 3: Nút Chạy đối chiếu
        r3 = tk.Frame(sel_box, bg="#FFFFFF")
        r3.pack(fill=tk.X)
        btn_run = tk.Button(r3, text="⚡ TIẾN HÀNH PHÂN TÍCH SO SÁNH BIẾN ĐỘNG", font=("Segoe UI", 10, "bold"), bg="#7C3AED", fg="white", activebackground="#6D28D9", activeforeground="white", padx=16, pady=6, cursor="hand2", command=self._on_run_diff)
        btn_run.pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_diff_status = tk.Label(r3, text="Sẵn sàng nạp 2 tệp để so sánh biến động.", font=("Segoe UI", 9, "italic"), fg="#64748B", bg="#FFFFFF")
        self.lbl_diff_status.pack(side=tk.LEFT)

        # 3. KPI Cards Frame
        self.kpi_frame = tk.Frame(self, bg="#F8FAFC", padx=12, pady=4)
        self.kpi_frame.pack(fill=tk.X)

        self.kpi_labels = {}
        cards_cfg = [
            ("added", "✨ MỚI THỤ LÝ TRONG KỲ", "0", "Phát sinh mới", "#047857", "#ECFDF5", "#A7F3D0"),
            ("resolved", "✅ ĐÃ GIẢI QUYẾT XONG", "0", "Đạt kết quả trong kỳ", "#1D4ED8", "#EFF6FF", "#BFDBFE"),
            ("worsened", "⚠️ NGUY CƠ THỜI HIỆU MỚI", "0", "Chuyển cấp khẩn cấp", "#B91C1C", "#FEF2F2", "#FECACA"),
            ("continued", "⏸ TỒN ĐỌNG TIẾP TỤC", "0", "Đang tiếp tục thụ lý", "#475569", "#F8FAFC", "#E2E8F0"),
        ]

        for idx, (cid, title, val, sub, fg_c, bg_c, bd_c) in enumerate(cards_cfg):
            card = tk.Frame(self.kpi_frame, bg=bg_c, highlightthickness=1, highlightbackground=bd_c, padx=10, pady=6)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
            tk.Label(card, text=title, font=("Segoe UI", 8, "bold"), fg=fg_c, bg=bg_c).pack(anchor="w")
            lbl_v = tk.Label(card, text=val, font=("Segoe UI", 16, "bold"), fg=fg_c, bg=bg_c)
            lbl_v.pack(anchor="w")
            tk.Label(card, text=sub, font=("Segoe UI", 7, "italic"), fg="#64748B", bg=bg_c).pack(anchor="w")
            self.kpi_labels[cid] = lbl_v

        # 4. Canvas Thanh Biểu đồ trực quan
        self.chart_canvas = tk.Canvas(self, bg="#FFFFFF", height=38, highlightthickness=1, highlightbackground="#CBD5E1")
        self.chart_canvas.pack(fill=tk.X, padx=16, pady=(4, 6))

        # 5. Notebook Tabs hiển thị chi tiết
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        self.tree_added = self._create_tab_tree("tab_added", "🟢 Mới Thụ Lý (0)")
        self.tree_resolved = self._create_tab_tree("tab_resolved", "🔵 Đã Giải Quyết Xong (0)")
        self.tree_worsened = self._create_tab_tree("tab_worsened", "🔴 Chuyển Cấp Khẩn Cấp (0)")
        self.tree_continued = self._create_tab_tree("tab_continued", "⚪ Tồn Đọng Tiếp Tục (0)")

        # 6. Action Bar (Bottom)
        act_bar = tk.Frame(self, bg="#07172C", padx=16, pady=8)
        act_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.btn_export_diff = tk.Button(act_bar, text="📥 Xuất Báo Cáo Biến Động Ra Excel (.xlsx)", font=("Segoe UI", 9, "bold"), bg="#15803D", fg="white", padx=14, pady=5, cursor="hand2", command=self._on_export_diff)
        self.btn_export_diff.pack(side=tk.LEFT, padx=(0, 8))

        btn_brief = tk.Button(act_bar, text="📝 Trích Yếu Số Liệu", font=("Segoe UI", 9), bg="#0284C7", fg="white", padx=12, pady=5, cursor="hand2", command=self._on_show_brief)
        btn_brief.pack(side=tk.LEFT, padx=(0, 8))

        btn_close = tk.Button(act_bar, text="Đóng", font=("Segoe UI", 9), bg="#475569", fg="white", padx=14, pady=5, cursor="hand2", command=self.destroy)
        btn_close.pack(side=tk.RIGHT)

    def _create_tab_tree(self, tab_id, tab_title):
        tab_frame = tk.Frame(self.nb, bg="#FFFFFF")
        self.nb.add(tab_frame, text=tab_title)

        cols = ("stt", "code", "title", "unit", "ksv", "due", "status")
        tree = ttk.Treeview(tab_frame, columns=cols, show="headings")
        tree.heading("stt", text="STT")
        tree.heading("code", text="Mã số hồ sơ")
        tree.heading("title", text="Tên vụ án / Nội dung tin báo")
        tree.heading("unit", text="Đơn vị VKS thụ lý")
        tree.heading("ksv", text="Kiểm sát viên")
        tree.heading("due", text="Hạn / Thời hiệu")
        tree.heading("status", text="Trạng thái / Đánh giá")

        tree.column("stt", width=45, anchor=tk.CENTER)
        tree.column("code", width=140, anchor=tk.CENTER)
        tree.column("title", width=340, anchor=tk.W)
        tree.column("unit", width=180, anchor=tk.W)
        tree.column("ksv", width=140, anchor=tk.W)
        tree.column("due", width=110, anchor=tk.CENTER)
        tree.column("status", width=160, anchor=tk.W)

        tree.tag_configure("odd", background="#F8FAFC")
        tree.tag_configure("even", background="#FFFFFF")

        sb = ttk.Scrollbar(tab_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        return tree

    def _on_choose_old(self):
        f = filedialog.askopenfilename(parent=self, title="Chọn File QLAHS Kỳ Trước (Cũ)", filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
        if f:
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                raw_rows = list(wb.active.iter_rows(values_only=True))
                fmt = detect_data_format(raw_rows)
                if fmt == "TIN_BAO":
                    recs = parse_qlahs_tin_bao_rows(raw_rows)
                elif fmt == "CASE_REPORT":
                    recs = parse_qlahs_case_report_rows(raw_rows, filename=os.path.basename(f))
                else:
                    recs = parse_qlahs_case_rows(raw_rows)
                self.records_old = recs
                self.name_old = os.path.basename(f)
                self.lbl_old.config(text=f"{self.name_old} ({len(recs):,} hồ sơ)", fg="#1D4ED8", font=("Segoe UI", 9, "bold"))
            except Exception as e:
                messagebox.showerror("Lỗi đọc file", f"Không thể đọc file kỳ trước: {e}", parent=self)

    def _on_choose_new(self):
        f = filedialog.askopenfilename(parent=self, title="Chọn File QLAHS Kỳ Này (Mới)", filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
        if f:
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                raw_rows = list(wb.active.iter_rows(values_only=True))
                fmt = detect_data_format(raw_rows)
                if fmt == "TIN_BAO":
                    recs = parse_qlahs_tin_bao_rows(raw_rows)
                elif fmt == "CASE_REPORT":
                    recs = parse_qlahs_case_report_rows(raw_rows, filename=os.path.basename(f))
                else:
                    recs = parse_qlahs_case_rows(raw_rows)
                self.records_new = recs
                self.name_new = os.path.basename(f)
                self.lbl_new.config(text=f"{self.name_new} ({len(recs):,} hồ sơ)", fg="#047857", font=("Segoe UI", 9, "bold"))
            except Exception as e:
                messagebox.showerror("Lỗi đọc file", f"Không thể đọc file kỳ này: {e}", parent=self)

    def _on_run_diff(self):
        if not self.records_old:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn File Kỳ Trước (Cũ) trước khi tiến hành phân tích.", parent=self)
            return
        if not self.records_new:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn File Kỳ Này (Mới) trước khi tiến hành phân tích.", parent=self)
            return

        self.diff_res = compare_two_qlahs_periods(self.records_old, self.records_new, self.current_format)
        st = self.diff_res["stats"]

        # Cập nhật KPI
        self.kpi_labels["added"].config(text=f"+{st['n_added']:,}")
        self.kpi_labels["resolved"].config(text=f"-{st['n_resolved']:,}")
        self.kpi_labels["worsened"].config(text=f"{st['n_worsened']:,}")
        self.kpi_labels["continued"].config(text=f"{st['n_continued']:,}")

        # Vẽ thanh biểu đồ trực quan
        w = self.chart_canvas.winfo_width() or 800
        draw_visual_distribution_chart(self.chart_canvas, w, 38, [
            ("Mới thụ lý", st["n_added"], "#10B981"),
            ("Đã giải quyết", st["n_resolved"], "#3B82F6"),
            ("Khẩn cấp mới", st["n_worsened"], "#EF4444"),
            ("Tồn đọng", st["n_continued"], "#94A3B8"),
        ])

        # Điền dữ liệu vào các Treeview
        def populate_tree(tree, items, is_worsened=False):
            for itm in tree.get_children():
                tree.delete(itm)
            for idx, r in enumerate(items, start=1):
                code = r.get("ma_tin_bao") or r.get("ma_vu_an") or ""
                title = r.get("noi_dung_tin_bao") or r.get("ten_vu_an") or ""
                unit = r.get("vks_thu_ly") or ""
                ksv = r.get("ksv_chinh") or r.get("can_bo_giai_quyet") or ""
                due = r.get("thoi_han_den_ngay") or r.get("ngay_het_thoi_hieu") or ""
                if is_worsened:
                    status = f"Từ {r.get('_old_urgency')} -> {r.get('_new_urgency')}"
                else:
                    status = r.get("ket_qua_giai_quyet") or r.get("muc_do_khan_cap") or r.get("canh_bao_thoi_hieu") or "Đang thụ lý"
                tag = "even" if idx % 2 == 0 else "odd"
                tree.insert("", "end", values=(idx, code, title, unit, ksv, due, status), tags=(tag,))

        populate_tree(self.tree_added, self.diff_res["added"])
        populate_tree(self.tree_resolved, self.diff_res["resolved"])
        populate_tree(self.tree_worsened, self.diff_res["worsened"], is_worsened=True)
        populate_tree(self.tree_continued, self.diff_res["continued"][:500])

        self.nb.tab(0, text=f"🟢 Mới Thụ Lý ({st['n_added']:,})")
        self.nb.tab(1, text=f"🔵 Đã Giải Quyết ({st['n_resolved']:,})")
        self.nb.tab(2, text=f"🔴 Chuyển Cấp Khẩn Cấp ({st['n_worsened']:,})")
        self.nb.tab(3, text=f"⚪ Tồn Đọng ({st['n_continued']:,})")

        self.lbl_diff_status.config(
            text=f"✔ Hoàn tất phân tích! Chênh lệch ròng: {'+' if st['net_change'] >= 0 else ''}{st['net_change']} hồ sơ.",
            fg="#16A34A",
            font=("Segoe UI", 9, "bold")
        )

    def _on_export_diff(self):
        if not self.diff_res:
            messagebox.showwarning("Chưa có kết quả", "Vui lòng bấm 'Tiến Hành Phân Tích' trước khi xuất file.", parent=self)
            return

        now_str = datetime.now().strftime("%d%m%Y_%H%h%M")
        default_fn = f"Bao_Cao_Bien_Dong_QLAHS_{now_str}.xlsx"
        out_f = filedialog.asksaveasfilename(
            parent=self,
            title="Lưu Báo Cáo Biến Động Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialfile=default_fn
        )
        if out_f:
            try:
                export_diff_to_excel(self.diff_res, out_f, self.name_old or "Kỳ Trước", self.name_new or "Kỳ Này")
                messagebox.showinfo("Thành công", f"Đã xuất file báo cáo so sánh biến động Excel thành công tại:\\n{out_f}", parent=self)
            except Exception as e:
                messagebox.showerror("Lỗi xuất file", f"Không thể lưu file Excel: {e}", parent=self)

    def _on_show_brief(self):
        if not self.diff_res:
            messagebox.showwarning("Chưa có kết quả", "Vui lòng bấm 'Tiến Hành Phân Tích' trước.", parent=self)
            return
        st = self.diff_res["stats"]
        txt = f"""VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH
TRÍCH YẾU SO SÁNH BIẾN ĐỘNG GIỮA 2 KỲ BÁO CÁO QLAHS
--------------------------------------------------------------------------------
• Kỳ trước: {self.name_old} ({st['total_old']:,} hồ sơ)
• Kỳ này:   {self.name_new} ({st['total_new']:,} hồ sơ)
• Biến động ròng: {'+' if st['net_change'] >= 0 else ''}{st['net_change']:,} hồ sơ

1. HỒ SƠ PHÁT SINH MỚI: +{st['n_added']:,} hồ sơ mới thụ lý trong kỳ.
2. HỒ SƠ ĐÃ GIẢI QUYẾT: -{st['n_resolved']:,} hồ sơ đã giải quyết / kết thúc.
3. CẢNH BÁO NGUY CƠ:    {st['n_worsened']:,} hồ sơ chuyển mức khẩn cấp thời hiệu TNHS.
4. TỒN ĐỌNG TIẾP TỤC:   {st['n_continued']:,} hồ sơ tiếp tục theo dõi giải quyết.
--------------------------------------------------------------------------------
Tác giả: {CURRENT_CONFIG.get('author_title')}"""
        dlg_b = tk.Toplevel(self)
        dlg_b.title("Trích Yếu Biến Động")
        dlg_b.geometry("600x380")
        dlg_b.configure(bg="#F8FAFC")
        dlg_b.transient(self)

        t_area = scrolledtext.ScrolledText(dlg_b, font=("Segoe UI", 10), wrap=tk.WORD, padx=10, pady=10)
        t_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t_area.insert(tk.END, txt)

        def _cp():
            self.clipboard_clear()
            self.clipboard_append(txt)
            messagebox.showinfo("Đã sao chép", "Đã chép trích yếu vào Clipboard!", parent=dlg_b)

        tk.Button(dlg_b, text="📋 Sao Chép (Copy)", font=("Segoe UI", 9, "bold"), bg="#0284C7", fg="white", padx=12, pady=4, command=_cp).pack(pady=(0, 10))
'''

with open("Phan_Mem_Chuan_Hoa_QLAHS_Desktop.py", "r", encoding="utf-8") as f:
    content = f.read()

target = "class QLAHSApp(tk.Tk):"
if target in content:
    new_content = content.replace(target, patch_code.strip() + "\n\n\n" + target)
    with open("Phan_Mem_Chuan_Hoa_QLAHS_Desktop.py", "w", encoding="utf-8") as f:
        f.write(new_content)
    print("Patch 1 applied successfully!")
else:
    print("Target class QLAHSApp not found!")
