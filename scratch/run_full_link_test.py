import sys, os
sys.path.insert(0, os.path.abspath('.'))
import openpyxl, re, datetime
from Phan_Mem_Chuan_Hoa_QLAHS_Desktop import parse_qlahs_case_report_rows
from scratch.test_link_final import blocks

f1 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162916746_407e1c4e.xlsx'
wb1 = openpyxl.load_workbook(f1, data_only=True)
cases = parse_qlahs_case_report_rows(list(wb1.active.iter_rows(values_only=True)), filename='Ho_So_Vu_An.xlsx')

def link_and_enrich(cases, blocks):
    def norm_s(s):
        return re.sub(r'[^a-zA-Z0-9]', '', str(s or '')).upper()
    def norm_txt(s):
        return re.sub(r'\s+', ' ', str(s or '')).strip().lower()

    case_by_code = {norm_s(c.get('ma_vu_an')): c for c in cases if c.get('ma_vu_an')}
    
    total_linked_defs = 0
    total_unlinked_defs = 0
    enriched_cases = []
    
    linked_defs_all = []
    unlinked_defs_all = []

    for blk in blocks:
        target_case = None
        # 1. Match by ma_vu_an
        for d in blk:
            m_va = d.get('ma_vu_an')
            if m_va and norm_s(m_va) in case_by_code:
                target_case = case_by_code[norm_s(m_va)]
                break
        # 2. Match by defendant name or CCCD
        if not target_case:
            for d in blk:
                name_bc = norm_txt(d.get('ten_bi_can'))
                cccd_bc = d.get('so_cccd')
                for c in cases:
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
            enriched_cases.append(target_case)
            c_defs = []
            for idx, d in enumerate(blk, 1):
                cccd_str = f" (CCCD: {d['so_cccd']})" if d['so_cccd'] else ""
                tt_str = f" - {d['trang_thai']}" if d['trang_thai'] else ""
                c_defs.append(f"{idx}. {d['ten_bi_can']}{cccd_str}{tt_str}")

            target_case['danh_sach_bi_can_chi_tiet'] = '\n'.join(c_defs)
            if not target_case.get('tong_bi_can') or int(str(target_case.get('tong_bi_can') or 0).strip() or 0) < len(blk):
                target_case['tong_bi_can'] = len(blk)
            if blk:
                d0 = blk[0]
                if not target_case.get('bi_can_dau_vu') or 'chưa' in str(target_case.get('bi_can_dau_vu')).lower():
                    cc_lead = f" - {d0['so_cccd']}" if d0['so_cccd'] else ""
                    target_case['bi_can_dau_vu'] = f"{d0['ten_bi_can']}{cc_lead}"

            for d in blk:
                d['ma_vu_an_linked'] = target_case.get('ma_vu_an')
                d['ten_vu_an_linked'] = target_case.get('ten_vu_an')
                d['giai_doan_linked'] = target_case.get('giai_doan')
                d['trang_thai_va_linked'] = target_case.get('trang_thai_vu_an')
                d['vks_thu_ly_linked'] = target_case.get('vks_thu_ly')
                linked_defs_all.append(d)
                total_linked_defs += 1
        else:
            for d in blk:
                d['ma_vu_an_linked'] = 'Xác minh sau'
                d['ten_vu_an_linked'] = 'Chờ đối soát từ PM cũ / Bổ sung hồ sơ vụ án'
                unlinked_defs_all.append(d)
                total_unlinked_defs += 1

    return enriched_cases, linked_defs_all, unlinked_defs_all

en_cases, l_defs, u_defs = link_and_enrich(cases, blocks)
print(f'Enriched cases: {len(en_cases)}')
print(f'Linked defs: {len(l_defs)}')
print(f'Unlinked defs (Xác minh sau): {len(u_defs)}')
for c in en_cases:
    print(f"\nCase {c['ma_vu_an']}: Total {c['tong_bi_can']} defs, Lead: {c['bi_can_dau_vu']}")
    print(f"Detail preview:\n{c['danh_sach_bi_can_chi_tiet'][:200]}...")
