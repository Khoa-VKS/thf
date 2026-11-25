import sys, os
sys.path.insert(0, os.path.abspath('.'))
import openpyxl

from Phan_Mem_Chuan_Hoa_QLAHS_Desktop import (
    parse_qlahs_case_rows,
    parse_qlahs_case_report_rows,
    split_officers,
    clean_person_name,
    is_procedural_status_text,
    is_form_status_text,
    export_case_enriched_to_excel,
    reconcile_and_link_cases_defendants,
    load_excel_and_parse
)

print("--- TEST 1: 8-column layout without headers ---")
mock_8_col_rows = [
    [308, "VA_E01.59_2605_1288\nVụ án trộm cắp", "Rất nghiêm trọng", "Điều 173", "", "Thụ lý mới", "Điều tra: Đang THQCT và kiểm sát điều tra", "Lưu hành nội bộ: 2"],
    [309, "VA_E01.59_2606_0651\nVụ án ma túy", "Đặc biệt nghiêm trọng", "Điều 251", "", "Thụ lý mới", "Điều tra: Đã chuyển để nhập", "Lưu hành nội bộ: 1"],
    [310, "VA_E01.59_2606_1297\nVụ án cướp", "Nghiêm trọng", "Điều 168", "", "Thụ lý mới", "Điều tra: Đang THQCT và kiểm sát điều tra", "Lưu hành nội bộ: 1"]
]

res1 = parse_qlahs_case_rows(mock_8_col_rows, filename="COPY TAY.xlsx")
for r in res1:
    print(f"STT {r['stt']}:")
    print(f"  KSV: repr={repr(r.get('can_bo_giai_quyet'))}")
    print(f"  TT_VA: repr={repr(r.get('trang_thai_vu_an'))}")
    print(f"  TT_BM: repr={repr(r.get('trang_thai_bieu_mau'))}")
    assert r.get('can_bo_giai_quyet') == "", f"FAIL: KSV should be empty, got {r.get('can_bo_giai_quyet')}"
    assert "Đang THQCT" in r.get('trang_thai_vu_an') or "Đã chuyển để nhập" in r.get('trang_thai_vu_an')
    assert "Lưu hành nội bộ" in r.get('trang_thai_bieu_mau')

print("TEST 1 PASSED!\n")

print("--- TEST 2: Real case file media_1795162916746_407e1c4e.xlsx parsed by parse_qlahs_case_rows ---")
f1 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162916746_407e1c4e.xlsx'
wb1 = openpyxl.load_workbook(f1, data_only=True)
raw_rows_1 = list(wb1.active.iter_rows(values_only=True))

res2 = parse_qlahs_case_rows(raw_rows_1, filename="COPY TAY.xlsx")
print(f"Total parsed: {len(res2)} cases")
bad_ksv_count = 0
for r in res2:
    ksv = r.get("can_bo_giai_quyet", "")
    if is_procedural_status_text(ksv) or is_form_status_text(ksv):
        bad_ksv_count += 1
        print("Bad KSV found:", ksv)

assert bad_ksv_count == 0, f"FAIL: Found {bad_ksv_count} cases with bad KSV!"
print("TEST 2 PASSED! 0 bad KSV out of 307 cases.\n")

print("--- TEST 3: split_officers guard test ---")
bad_inputs = [
    "Điều tra: Đang THQCT và kiểm sát điều tra",
    "Điều tra: Đã chuyển để nhập",
    "Lưu hành nội bộ: 2",
    "Đã phát hành văn bản: 2",
    "Nguyễn Văn A; Điều tra: Đang THQCT và kiểm sát điều tra; Bùi Xuân Thái"
]

for inp in bad_inputs:
    ld_v, ld_p, ksv, cb_tl = split_officers(inp, "")
    print(f"Input: {inp}")
    print(f"  -> LD_V: {repr(ld_v)}, LD_P: {repr(ld_p)}, KSV: {repr(ksv)}, CB_TL: {repr(cb_tl)}")
    assert not is_procedural_status_text(ksv), f"FAIL: ksv contains status: {ksv}"
    assert not is_form_status_text(ksv), f"FAIL: ksv contains form: {ksv}"

print("TEST 3 PASSED!\n")

print("--- TEST 4: Export to Excel and verify columns in actual workbook ---")
out_file = 'scratch/test_clean_officers_export.xlsx'
cases, enriched_n, stats, l_defs, u_defs = reconcile_and_link_cases_defendants(
    [('Ho_So_Vu_An.xlsx', 'CASE', res2)],
    []
)
export_case_enriched_to_excel(cases, l_defs, u_defs, out_file)
wb_out = openpyxl.load_workbook(out_file)
ws_cases = wb_out['So_Vu_An_Da_Lam_Giau']

# Find column index for 'Kiểm sát viên'
ksv_col = None
tt_col = None
for c in range(1, 30):
    val = ws_cases.cell(4, c).value
    if val and 'Kiểm sát viên' in str(val):
        ksv_col = c
    if val and 'Trạng thái vụ án' in str(val):
        tt_col = c

print(f"Header check: KSV col={ksv_col}, TT col={tt_col}")
assert ksv_col is not None
assert tt_col is not None

bad_excel_count = 0
for r in range(5, ws_cases.max_row + 1):
    cell_val = str(ws_cases.cell(r, ksv_col).value or "").strip()
    if is_procedural_status_text(cell_val) or is_form_status_text(cell_val):
        bad_excel_count += 1
        print(f"Row {r} has bad KSV: {cell_val}")

assert bad_excel_count == 0, f"FAIL: Found {bad_excel_count} bad rows in Excel!"
print("TEST 4 PASSED! 100% clean columns in exported Excel.\n")

print("ALL TESTS PASSED SUCCESSFULLY!")
