import json
from pathlib import Path
import openpyxl

requested = '01,02,13,14,15,16,17,19,20,21,22,23,25,26,27,28,29,30,31,33,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55'.split(',')
source = Path(r'C:\Users\UserPc\Downloads\칩셋명 정보 정리.xlsx')
wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
ws = wb.worksheets[0]
headers = list(next(ws.values))
rows = {str(row[0]).zfill(2): {'sourceRow': i, 'values': list(row)} for i, row in enumerate(ws.values, 1) if row[0] is not None and str(row[0]).zfill(2) in requested}
assert len(rows) == len(requested) == 40
result = {'filename': source.name, 'sheet': ws.title, 'headers': headers, 'requested': requested, 'rows': rows}
Path(__file__).with_name('source.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'count': len(rows), 'sheet': ws.title, 'headers': headers}, ensure_ascii=True))
