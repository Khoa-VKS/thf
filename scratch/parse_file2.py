import openpyxl

f2 = r'C:/Users/VKS/.gemini/antigravity/brain/18114c75-83f9-4938-9863-88b05dabe0aa/.user_uploaded/media_1795162929725_af85271c.xlsx'
wb = openpyxl.load_workbook(f2, data_only=True)
ws = wb.active

records = []
for r in range(1, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    c0 = str(vals[0] or '').strip().lower()
    c1 = str(vals[1] or '').strip()
    c2 = str(vals[2] or '').strip()

    if c0 == 'stt' and 'mã bị can' in str(vals[1] or '').lower():
        continue
    
    if c1 and (c1.startswith('VA_') or c1.startswith('E01') or c1.startswith('010')):
        records.append({
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

print(f'Total extracted defendants: {len(records)}')
for idx, rec in enumerate(records):
    print(f"{idx+1:>2}. BC: {rec['ma_bi_can']:<30} | Tên: {rec['ho_ten']:<22} | CCCD: {rec['cccd']:<15} | TT: {rec['trang_thai']}")
