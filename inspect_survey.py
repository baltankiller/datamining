from pathlib import Path
import csv, json, re
import pandas as pd
import openpyxl

ROOT = Path(__file__).resolve().parent
SRC = ROOT / '데이터분석 데이터'
OUT = ROOT / '전처리결과'
OUT.mkdir(exist_ok=True)
result = []
for p in sorted(SRC.rglob('*')):
    if not p.is_file():
        continue
    name = p.name
    if not any(s in name for s in ['고령', '장노년']):
        continue
    entry = {'file': str(p.relative_to(SRC))}
    if p.suffix.lower() == '.csv':
        for enc in ['utf-8-sig','cp949','euc-kr']:
            try:
                text = p.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        rows = list(csv.reader(text.splitlines()))
        entry.update(encoding=enc, rows=len(rows)-1, columns=len(rows[0]), header=rows[0], sample=rows[1:3])
    elif p.suffix.lower() == '.xlsx':
        wb = openpyxl.load_workbook(p,read_only=True,data_only=True)
        entry['sheets'] = []
        for ws in wb:
            info = {'name':ws.title,'rows':ws.max_row,'columns':ws.max_column}
            if '변수' in ws.title:
                info['contents'] = [[i,list(r)] for i,r in enumerate(ws.iter_rows(values_only=True),1) if any(v is not None for v in r)]
            else:
                info['head'] = list(ws.iter_rows(min_row=1,max_row=2,values_only=True))
            entry['sheets'].append(info)
        wb.close()
    result.append(entry)
(OUT/'structure_inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
for e in result:
    print(e['file'])
    if 'sheets' in e:
        print([(s['name'],s['rows'],s['columns']) for s in e['sheets']])
        for s in e['sheets']:
            if 'contents' in s:
                matches = [r for i,r in s['contents'] if len(r)>2 and isinstance(r[2],str) and re.search('연령|만 나이|지역|소득|효능|자신|교육|역량|점수|등급',r[2])]
                print(json.dumps(matches,ensure_ascii=False,default=str))
    else:
        print(e['encoding'],e['rows'],e['columns'],e['header'][:15],e['header'][-25:])
