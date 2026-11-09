from pathlib import Path
import json,re
import pandas as pd
ROOT=Path(__file__).resolve().parent
SRC=ROOT/'데이터분석 데이터'
meta=json.loads((ROOT/'전처리결과/structure_inspection.json').read_text(encoding='utf-8'))
for y in range(2019,2026):
    entries=[e for e in meta if re.search(r'20\d\d',e['file']).group()==str(y)]
    raw=next(e for e in entries if 'header' in e or ('sheets' in e and any('head' in s for s in e['sheets'])))
    p=SRC/raw['file']
    df=pd.read_csv(p,encoding='utf-8-sig',dtype=str,keep_default_na=False) if p.suffix=='.csv' else pd.read_excel(p,sheet_name=0,dtype=str,keep_default_na=False)
    cb=next(e for e in entries if 'sheets' in e)
    labels={r[0]:r[2] for s in cb['sheets'] if '정보' in s['name'] for i,r in s.get('contents',[]) if r[0] and len(r)>2 and isinstance(r[1],(int,float))}
    age=pd.to_numeric(df.ADQ1,errors='coerce')
    print(y,'shape',df.shape,'age range',(int(age.min()),int(age.max())),'65plus',int((age>=65).sum()),'head',df.columns[:5].tolist())
    for v,label in labels.items():
        if any(t in str(label) for t in ['배우는데','활용하는데','빠르게 알아','교육 과정 인지','교육 경험','교육을 받은','교육 참여']):
            print('candidate',v,label)
    print('PC14',df[[c for c in df if re.fullmatch('Q[45]A[1-7]|Q[67]_[1-7]',c)]].apply(pd.to_numeric,errors='coerce').stack().value_counts().to_dict())
    print('unknown age label',labels.get('ADQ1'))
