import sys, os
sys.path.insert(0, os.path.abspath('.'))
import openpyxl, re
from Phan_Mem_Chuan_Hoa_QLAHS_Desktop import parse_qlahs_case_report_rows

f1 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162916746_407e1c4e.xlsx'
f2 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162929725_af85271c.xlsx'

wb1 = openpyxl.load_workbook(f1, data_only=True)
raw_cases = list(wb1.active.iter_rows(values_only=True))
case_records = parse_qlahs_case_report_rows(raw_cases)

# Parse File 2 into BLOCKS of defendants
wb2 = openpyxl.load_workbook(f2, data_only=True)
ws2 = wb2.active

def parse_defendant_blocks(ws):
    blocks = []
    current_block = []
    
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        c0 = str(vals[0] or '').strip().lower()
        c1 = str(vals[1] or '').strip()
        c2 = str(vals[2] or '').strip()

        # If header line
        if c0 == 'stt' and ('mã bị can' in str(vals[1] or '').lower() or 'họ tên' in str(vals[2] or '').lower()):
            if current_block:
                blocks.append(current_block)
                current_block = []
            continue
            
        # If empty line
        if not any(vals):
            if current_block:
                blocks.append(current_block)
                current_block = []
            continue
            
        # Continuation of crime description
        c7 = str(vals[7] or '').strip() if len(vals) > 7 and vals[7] is not None else ''
        if not c1 and not c2 and c7 and current_block:
            current_block[-1]['dieu_luat'] += '\n' + c7
            continue
            
        if c1 and (c1.startswith('VA_') or c1.startswith('E01') or c1.startswith('010') or re.match(r'^[A-Z0-9_\.]+$', c1)):
            ns_val = vals[3] if len(vals) > 3 else None
            ns_str = ''
            if hasattr(ns_val, 'strftime'):
                ns_str = ns_val.strftime('%d/%m/%Y')
            elif ns_val:
                ns_raw = str(ns_val).strip()[:10]
                m_dt = re.match(r'^(\d{4})-(\d{2})-(\d{2})', ns_raw)
                ns_str = f'{m_dt.group(3)}/{m_dt.group(2)}/{m_dt.group(1)}' if m_dt else ns_raw

            cccd_val = str(vals[5] or '').strip() if len(vals) > 5 and vals[5] is not None else ''
            if cccd_val.isdigit() and len(cccd_val) == 11:
                cccd_val = '0' + cccd_val

            m_va = ''
            m_match = re.match(r'^(VA_[A-Za-z0-9_\.]+?)_\d{2,3}$', c1)
            if m_match:
                m_va = m_match.group(1)

            current_block.append({
                'row': r,
                'stt': vals[0],
                'ma_bi_can': c1,
                'ten_bi_can': c2,
                'ngay_sinh': ns_str,
                'so_cccd': cccd_val,
                'gioi_tinh': str(vals[4] or '').strip() if len(vals) > 4 and vals[4] is not None else '',
                'dia_chi': str(vals[6] or '').strip() if len(vals) > 6 and vals[6] is not None else '',
                'dieu_luat': c7,
                'loai_toi': str(vals[8] or '').strip() if len(vals) > 8 and vals[8] is not None else '',
                'trang_thai': str(vals[9] or '').strip() if len(vals) > 9 and vals[9] is not None else '',
                'ma_vu_an': m_va
            })

    if current_block:
        blocks.append(current_block)
    return blocks

blocks = parse_defendant_blocks(ws2)
print(f'Total parsed blocks: {len(blocks)}')
for i, b in enumerate(blocks, 1):
    print(f'Block {i}: {len(b)} defs (First: {b[0]["ten_bi_can"]}, Last: {b[-1]["ten_bi_can"]})')

# Now link blocks to cases!
def norm_s(s):
    return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()

def norm_txt(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip().lower()

case_by_code = {norm_s(c.get('ma_vu_an')): c for c in case_records if c.get('ma_vu_an')}

for i, blk in enumerate(blocks, 1):
    target_case = None
    
    # 1. Check if any def has explicit or derived ma_vu_an that matches a case
    for d in blk:
        m_va = d.get('ma_vu_an')
        if m_va and norm_s(m_va) in case_by_code:
            target_case = case_by_code[norm_s(m_va)]
            break

    # 2. Check if any defendant matches case's bi_can_dau_vu or ten_vu_an
    if not target_case:
        for d in blk:
            name_bc = norm_txt(d.get('ten_bi_can'))
            cccd_bc = d.get('so_cccd')
            for c in case_records:
                c_bc = norm_txt(c.get('bi_can_dau_vu') or '')
                c_va = norm_txt(c.get('ten_vu_an') or '')
                if cccd_bc and len(cccd_bc) >= 9 and cccd_bc in (c_bc + ' ' + c_va):
                    target_case = c
                    break
                if name_bc and len(name_bc) >= 4 and (name_bc in c_bc or name_bc in c_va):
                    target_case = c
                    break
            if target_case:
                break

    if target_case:
        print(f"\n-> Block {i} ({len(blk)} defs) LINKED to Case: {target_case.get('ma_vu_an')}")
        print(f"   Tên VA: {target_case.get('ten_vu_an')}")
        print(f"   Bị can đầu vụ: {target_case.get('bi_can_dau_vu')}")
        for d in blk:
            d['ma_vu_an_linked'] = target_case.get('ma_vu_an')
    else:
        print(f"\n-> Block {i} ({len(blk)} defs) UNLINKED (Xác minh sau)")
        for d in blk:
            d['ma_vu_an_linked'] = 'Xác minh sau'
