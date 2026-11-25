import openpyxl, re, os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from datetime import date

from Phan_Mem_Chuan_Hoa_QLAHS_Desktop import (
    COLUMNS_CASE_24, is_procedural_status_text, is_form_status_text,
    clean_vks_unit_name, split_officers, detect_procedural_stage,
    detect_province, parse_date_safely, detect_file_unit,
    normalize_case_stage_and_status
)

def parse_case_header_to_col_map(row):
    r_strs = [str(c or "").strip().lower() for c in row]
    col_map = {"va": 1, "tl": None, "bc": None, "td": None, "lt": None, "cb": None, "tt": None, "bm": None}
    for idx, c_text in enumerate(r_strs):
        if not c_text:
            continue
        if "thông tin vụ án" in c_text or "mã vụ án" in c_text:
            col_map["va"] = idx
        elif "thông tin thụ lý" in c_text:
            col_map["tl"] = idx
        elif "loại thụ lý" in c_text:
            if col_map["tl"] is None:
                col_map["tl"] = idx
        elif "thông tin bị can" in c_text or "bị can" in c_text:
            col_map["bc"] = idx
        elif "loại tội" in c_text:
            col_map["lt"] = idx
        elif ("tội danh" in c_text or c_text == "tội" or "thông tin tội danh" in c_text) and "loại tội" not in c_text:
            col_map["td"] = idx
        elif "cán bộ" in c_text or "kiểm sát viên" in c_text:
            col_map["cb"] = idx
        elif "trạng thái vụ án" in c_text:
            col_map["tt"] = idx
        elif "trạng thái biểu mẫu" in c_text or "biểu mẫu" in c_text:
            col_map["bm"] = idx
    return col_map

def detect_initial_case_col_map(raw_rows):
    for r in raw_rows[:25]:
        if not r or not isinstance(r, (list, tuple)):
            continue
        line_str = " ".join(str(c or "").strip().lower() for c in r)
        if "thông tin vụ án" in line_str or "mã vụ án" in line_str:
            return parse_case_header_to_col_map(r)

    col_map = {"va": 1, "tl": None, "bc": None, "td": None, "lt": None, "cb": None, "tt": None, "bm": None}
    col_scores = {}
    for r in raw_rows[:35]:
        if not r or not isinstance(r, (list, tuple)):
            continue
        for idx, val in enumerate(r):
            if val is None:
                continue
            v_s = str(val).strip().replace("\xa0", " ")
            v_low = v_s.lower()
            if not v_s:
                continue
            if idx not in col_scores:
                col_scores[idx] = set()
            if v_s.startswith("VA_") or "số qđ khởi tố:" in v_low:
                col_scores[idx].add("va")
            if "số thụ lý:" in v_low or "ngày thụ lý:" in v_low or "thụ lý mới" in v_low:
                col_scores[idx].add("tl")
            if any(x in v_low for x in ("đặc biệt nghiêm trọng", "ít nghiêm trọng", "rất nghiêm trọng", "nghiêm trọng")):
                col_scores[idx].add("lt")
            if idx != 1 and ("điều " in v_low or "khoản " in v_low or "tội " in v_low):
                col_scores[idx].add("td")
            if "cán bộ thụ lý:" in v_low or "cán bộ giải quyết:" in v_low:
                col_scores[idx].add("cb")
            if is_procedural_status_text(v_s):
                col_scores[idx].add("tt")
            if is_form_status_text(v_s):
                col_scores[idx].add("bm")

    for idx, tags in col_scores.items():
        if "va" in tags and col_map["va"] is None:
            col_map["va"] = idx
        if "tl" in tags and col_map["tl"] is None:
            col_map["tl"] = idx
        if "lt" in tags and col_map["lt"] is None:
            col_map["lt"] = idx
        if "td" in tags and col_map["td"] is None:
            col_map["td"] = idx
        if "cb" in tags and col_map["cb"] is None:
            col_map["cb"] = idx
        if "tt" in tags and col_map["tt"] is None:
            col_map["tt"] = idx
        if "bm" in tags and col_map["bm"] is None:
            col_map["bm"] = idx

    return col_map

def parse_qlahs_case_rows_new(raw_rows, progress_cb=None, filename=""):
    file_context_unit = detect_file_unit(filename, raw_rows) if filename else None
    blocks = []
    current = []
    current_col_map = detect_initial_case_col_map(raw_rows)

    for row in raw_rows:
        if not row:
            continue
        line_str = " ".join(str(c or "").strip().lower() for c in row)
        c0 = row[0] if len(row) > 0 else None
        c1 = row[1] if len(row) > 1 else None

        c0_s = str(c0 or "").strip()
        c1_s = str(c1 or "").strip()

        if "thông tin vụ án" in line_str or "mã vụ án" in line_str:
            if current:
                blocks.append((current, current_col_map))
                current = []
            current_col_map = parse_case_header_to_col_map(row)
            continue

        if c0_s.lower() == "stt":
            continue

        if re.match(r'^\d+(\.0*)?$', c0_s):
            if current:
                blocks.append((current, current_col_map))
            current = []

        if any(v is not None and str(v).strip() for v in row):
            current.append(row)

    if current:
        blocks.append((current, current_col_map))

    total_blocks = len(blocks)
    records = []

    for idx, (blk, col_map) in enumerate(blocks, 1):
        rec = {k: "" for k, _, _ in COLUMNS_CASE_24}
        rec["stt"] = idx
        ten_lines = []

        for r_idx, row in enumerate(blk):
            def get_cell_val(c_idx):
                if c_idx is not None and len(row) > c_idx and row[c_idx] is not None:
                    return str(row[c_idx]).strip().replace("\xa0", " ")
                return ""

            v_va = get_cell_val(col_map.get("va"))
            v_tl = get_cell_val(col_map.get("tl"))
            v_bc = get_cell_val(col_map.get("bc"))
            v_td = get_cell_val(col_map.get("td"))
            v_lt = get_cell_val(col_map.get("lt"))
            v_cb = get_cell_val(col_map.get("cb"))
            v_tt = get_cell_val(col_map.get("tt"))
            v_bm = get_cell_val(col_map.get("bm"))

            # Lưới an toàn tự chữa lành
            if is_procedural_status_text(v_cb):
                if is_form_status_text(v_tt):
                    v_bm = v_tt
                v_tt = v_cb
                v_cb = ""
            if is_procedural_status_text(v_bm) and not is_procedural_status_text(v_tt) and v_tt:
                v_cb = v_tt
                v_tt = v_bm
                v_bm = ""
            if is_form_status_text(v_tt) and not v_bm:
                v_bm = v_tt
                v_tt = ""

            if v_va:
                va_lines = [ln.strip() for ln in v_va.replace("\r\n", "\n").split("\n") if ln.strip()]
                for l_i, ln in enumerate(va_lines):
                    if r_idx == 0 and l_i == 0 and ln.startswith("VA_"):
                        rec["ma_vu_an"] = ln
                    else:
                        ln_low = ln.lower()
                        if ln_low.startswith("số qđ khởi tố:") or ln_low.startswith("mã số qđ khởi tố"):
                            rec["so_qd_khoi_to"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("ngày ra qđ khởi tố:"):
                            rec["ngay_qd_khoi_to"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("đơn vị ra qđ khởi tố:"):
                            rec["don_vi_khoi_to"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("ngày tiếp nhận hồ sơ:"):
                            rec["ngay_tiep_nhan"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("thời hạn giải quyết:") or ln_low.startswith("thời hạn xử lý:"):
                            rec["thoi_han_giai_quyet"] = ln.split(":", 1)[1].strip()
                        elif ln_low.startswith("ngày hết thời hiệu"):
                            rec["ngay_het_thoi_hieu"] = ln.split(":", 1)[1].strip()
                        elif "thời hiệu" in ln_low:
                            rec["canh_bao_thoi_hieu"] = ln
                        elif ln.startswith("VA_") and not rec["ma_vu_an"]:
                            rec["ma_vu_an"] = ln
                        else:
                            ten_lines.append(ln)

            if v_tl:
                tl_lines = [ln.strip() for ln in v_tl.replace("\r\n", "\n").split("\n") if ln.strip()]
                for ln in tl_lines:
                    v_tl_low = ln.lower()
                    if v_tl_low.startswith("số thụ lý:"):
                        rec["so_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_tl_low.startswith("ngày thụ lý:"):
                        rec["ngay_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_tl_low.startswith("loại thụ lý:"):
                        rec["loai_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_tl_low.startswith("đơn vị vks thụ lý:"):
                        rec["vks_thu_ly"] = clean_vks_unit_name(ln.split(":", 1)[1].strip())
                    elif not rec["loai_thu_ly"] and any(x in v_tl_low for x in ("thụ lý mới", "phục hồi", "tiếp nhận")):
                        rec["loai_thu_ly"] = ln

            if v_cb:
                cb_lines = [ln.strip() for ln in v_cb.replace("\r\n", "\n").split("\n") if ln.strip()]
                for ln in cb_lines:
                    v_cb_low = ln.lower()
                    if v_cb_low.startswith("cán bộ thụ lý:"):
                        rec["can_bo_thu_ly"] = ln.split(":", 1)[1].strip()
                    elif v_cb_low.startswith("cán bộ giải quyết:"):
                        rec["can_bo_giai_quyet"] = ln.split(":", 1)[1].strip()
                    else:
                        if rec["can_bo_giai_quyet"]:
                            rec["can_bo_giai_quyet"] += "; " + ln
                        else:
                            rec["can_bo_giai_quyet"] = ln

            if v_bc:
                bc_lines = [ln.strip() for ln in v_bc.replace("\r\n", "\n").split("\n") if ln.strip()]
                for ln in bc_lines:
                    v_bc_low = ln.lower()
                    if v_bc_low.startswith("bị can đầu vụ:"):
                        rec["bi_can_dau_vu"] = ln.split(":", 1)[1].strip()
                    elif v_bc_low.startswith("tổng bị can khởi tố:") or v_bc_low.startswith("tổng số bị can:"):
                        rec["tong_bi_can"] = ln.split(":", 1)[1].strip()
                    elif not rec["bi_can_dau_vu"]:
                        rec["bi_can_dau_vu"] = ln

            if v_td:
                if rec["dieu_luat"]:
                    rec["dieu_luat"] += " | " + v_td
                else:
                    rec["dieu_luat"] = v_td

            if v_lt and not rec["loai_toi"]:
                rec["loai_toi"] = v_lt

            if v_tt:
                rec["trang_thai_vu_an"] = v_tt
                rec["giai_doan"] = detect_procedural_stage(v_tt, filename=filename, loai_thu_ly=rec.get("loai_thu_ly"))

            if v_bm:
                if rec["trang_thai_bieu_mau"]:
                    rec["trang_thai_bieu_mau"] += "; " + v_bm
                else:
                    rec["trang_thai_bieu_mau"] = v_bm

        ld_v, ld_p, ksv, cb_tl = split_officers(rec.get("can_bo_giai_quyet", ""), rec.get("can_bo_thu_ly", ""))
        rec["lanh_dao_vien"] = ld_v
        rec["lanh_dao_phong"] = ld_p
        rec["can_bo_giai_quyet"] = ksv
        rec["can_bo_thu_ly"] = cb_tl
        rec["ten_vu_an"] = " - ".join(ten_lines)

        prov = detect_province(rec["don_vi_khoi_to"], rec["vks_thu_ly"] + " " + rec["ten_vu_an"])
        if prov in ("Khác / Chưa rõ", "Chưa xác định") and (
            "hồ chí minh" in (rec["vks_thu_ly"] + " " + rec["ten_vu_an"]).lower()
            or "hải quan" in rec["don_vi_khoi_to"].lower()
        ):
            prov = "TP. Hồ Chí Minh"
        rec["tinh_thanh"] = prov

        if filename:
            rec["nguon_du_lieu"] = filename

        normalize_case_stage_and_status(rec, filename=filename)
        records.append(rec)

    return records

wb1 = openpyxl.load_workbook(r"D:\C BÍCH\TỔNG ÁN\COPY TAY.xlsx", data_only=True)
recs1 = parse_qlahs_case_rows_new(list(wb1.active.iter_rows(values_only=True)), filename="COPY TAY.xlsx")
print(f"Successfully parsed {len(recs1)} cases from COPY TAY.xlsx!")

# Verify field integrity
officer_in_status = 0
status_in_officer = 0
for idx, r in enumerate(recs1, 1):
    tt = r["trang_thai_vu_an"]
    ksv = r["can_bo_giai_quyet"]
    cb = r["can_bo_thu_ly"]
    ldv = r["lanh_dao_vien"]
    ldp = r["lanh_dao_phong"]
    if any(is_procedural_status_text(x) for x in (ksv, cb, ldv, ldp)):
        status_in_officer += 1
        print(f"Status in officer at case #{idx}: ksv={ksv}, cb={cb}")
    if "cán bộ" in tt.lower() or "kiểm sát viên" in tt.lower():
        officer_in_status += 1
        print(f"Officer in status at case #{idx}: tt={tt}")

print(f"Integrity test results: Status in officer={status_in_officer}, Officer in status={officer_in_status}")

# Inspect cases 1, 181, 308 (from the user screenshots)
for c_i in [1, 181, 308]:
    case = recs1[c_i - 1]
    print(f"=== CASE #{c_i} ===")
    print("  Mã vụ án:", case["ma_vu_an"])
    print("  Tên vụ án:", case["ten_vu_an"][:60])
    print("  LĐ Viện:", case["lanh_dao_vien"])
    print("  LĐ Phòng:", case["lanh_dao_phong"])
    print("  KSV GQ:", case["can_bo_giai_quyet"])
    print("  CB Thụ lý:", case["can_bo_thu_ly"])
    print("  Trạng thái VA:", case["trang_thai_vu_an"])
    print("  Trạng thái BM:", case["trang_thai_bieu_mau"])
