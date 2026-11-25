import openpyxl
import re

f1 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162916746_407e1c4e.xlsx'
f2 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162929725_af85271c.xlsx'

wb1 = openpyxl.load_workbook(f1, data_only=True)
ws1 = wb1.active

cases = []
for r in range(4, ws1.max_row + 1):
    vals = [ws1.cell(r, c).value for c in range(1, 10)]
    stt = vals[0]
    tt_va = str(vals[1] or '').strip()
    if not tt_va:
        continue
    
    # parse lines in tt_va
    lines = [l.strip() for l in tt_va.split('\n') if l.strip()]
    ma_va = lines[0] if lines else ''
    ten_va = lines[1] if len(lines) > 1 else ''
    
    qd_kt = ''
    ngay_kt = ''
    for ln in lines:
        if 'mã số qđ khởi tố' in ln.lower() or 'số qđ khởi tố' in ln.lower():
            qd_kt = ln.split(':', 1)[1].strip() if ':' in ln else ln
        elif 'ngày ra qđ khởi tố' in ln.lower():
            ngay_kt = ln.split(':', 1)[1].strip() if ':' in ln else ln
            
    cases.append({
        'stt': stt,
        'ma_vu_an': ma_va,
        'ten_vu_an': ten_va,
        'qd_kt': qd_kt,
        'ngay_kt': ngay_kt,
        'loai_toi': vals[2],
        'toi_danh': vals[3],
        'tt_bi_can': str(vals[4] or '').strip(),
        'loai_thu_ly': vals[5],
        'trang_thai_va': vals[6],
        'trang_thai_bm': vals[7]
    })

print(f'Total cases loaded from File 1: {len(cases)}')

# Load defendants from File 2
wb2 = openpyxl.load_workbook(f2, data_only=True)
ws2 = wb2.active

defs = []
for r in range(1, ws2.max_row + 1):
    vals = [ws2.cell(r, c).value for c in range(1, ws2.max_column + 1)]
    c0 = str(vals[0] or '').strip().lower()
    c1 = str(vals[1] or '').strip()
    c2 = str(vals[2] or '').strip()

    if c0 == 'stt' and 'mã bị can' in str(vals[1] or '').lower():
        continue
    
    if c1 and (c1.startswith('VA_') or c1.startswith('E01') or c1.startswith('010')):
        defs.append({
            'row': r,
            'stt': vals[0],
            'ma_bi_can': c1,
            'ho_ten': c2,
            'ngay_sinh': str(vals[3] or '')[:10],
            'gioi_tinh': vals[4],
            'cccd': str(vals[5] or '') if vals[5] is not None else '',
            'dia_chi': vals[6],
            'dieu_luat': vals[7],
            'loai_toi': vals[8],
            'trang_thai': vals[9] if len(vals) > 9 else ''
        })

print(f'Total defendants loaded from File 2: {len(defs)}')

# Now match!
# Group defendants by their apparent group/case
def norm(s):
    return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

# Check how each defendant matches with cases
matched_cases = {} # case_idx -> list of defs
unmatched_defs = []

for d in defs:
    m_bc = d['ma_bi_can']
    name_bc = d['ho_ten']
    cccd_bc = d['cccd']
    
    matched = False
    for idx, c in enumerate(cases):
        m_va = c['ma_vu_an']
        # 1. Exact prefix match (new format VA_..._00x)
        if m_bc.startswith(m_va) or norm(m_bc).startswith(norm(m_va)):
            matched_cases.setdefault(idx, []).append(d)
            matched = True
            break
        # 2. Match by case name containing defendant name or CCCD
        norm_name = norm(name_bc)
        norm_va = norm(c['ten_vu_an'])
        norm_ttbc = norm(c['tt_bi_can'])
        if norm_name and (norm_name in norm_va or norm_name in norm_ttbc):
            matched_cases.setdefault(idx, []).append(d)
            matched = True
            break
        if cccd_bc and len(cccd_bc) >= 9 and (cccd_bc in c['tt_bi_can'] or cccd_bc in c['ten_vu_an']):
            matched_cases.setdefault(idx, []).append(d)
            matched = True
            break
            
    if not matched:
        unmatched_defs.append(d)

print(f'\n=== MATCHING RESULTS ===')
print(f'Number of cases matched: {len(matched_cases)}')
for c_idx, d_list in matched_cases.items():
    c = cases[c_idx]
    print(f"\n[Case STT {c['stt']}] Code: {c['ma_vu_an']}")
    print(f"  Tên VA: {c['ten_vu_an']}")
    print(f"  TT Bị can trong File 1: {repr(c['tt_bi_can'])}")
    print(f"  -> Linked được {len(d_list)} bị can từ File 2:")
    for d in d_list[:5]:
        print(f"     + BC: {d['ma_bi_can']} | {d['ho_ten']} | CCCD: {d['cccd']} | {d['trang_thai']}")
    if len(d_list) > 5:
        print(f"     ... và {len(d_list)-5} bị can khác.")

print(f'\nTotal unmatched defendants: {len(unmatched_defs)}')
if unmatched_defs:
    print('Sample unmatched:')
    for d in unmatched_defs[:10]:
        print(f"  BC: {d['ma_bi_can']} | {d['ho_ten']}")
