from pathlib import Path
import json,re
root=Path(__file__).resolve().parent
data=json.loads((root/'전처리결과/structure_inspection.json').read_text(encoding='utf-8'))
for e in data:
    if 'sheets' not in e: continue
    year=re.search(r'20\d\d',e['file']).group()
    sheets=e['sheets']
    labels={r[0]:r[2] for s in sheets if '정보' in s['name'] or '정보' in s['name'].replace(' ','') for i,r in s.get('contents',[]) if len(r)>2 and r[0] is not None and isinstance(r[1],(int,float))}
    vals={}
    for s in sheets:
        if '값' not in s['name']:continue
        v=None
        for i,r in s.get('contents',[]):
            if r[0] is not None:v=r[0]
            if v and r[1] is not None and len(r)>2:vals.setdefault(v,{})[str(r[1])]=r[2]
    print('\nYEAR',year)
    for v,label in labels.items():
        if v in ['ADQ1','ADQ9','Q2A1R','Q2A11','Q2K1_1_1','ADQ4'] or (re.search('효능감|정보화교육',str(label)) and not re.search('수요|모두|기회',str(label))):
            print(v,label,json.dumps(vals.get(v,{}),ensure_ascii=False))
    print('PC/MOBILE:',[(v,label[:22]) for v,label in labels.items() if re.search('이용 능력',str(label))][:2])
