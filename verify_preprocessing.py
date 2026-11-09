"""Independent source/output reconciliation for the prepared survey files."""
from pathlib import Path
import json,re,hashlib,csv,io
import numpy as np
import pandas as pd
import openpyxl

ROOT=Path(__file__).resolve().parent
OUT=ROOT/'전처리결과'
SRC=ROOT/'데이터분석 데이터'
meta=json.loads((OUT/'04_점검/run_metadata.json').read_text(encoding='utf-8'))
pool=pd.read_csv(OUT/'02_통합자료/고령자60세이상_2019_2025_전처리.csv',keep_default_na=False)
assert len(pool)==12173 and pool.record_id.is_unique
verified=[]
for r in meta['summary']:
    year=r['year']; source=SRC/r['source_file']
    raw=(pd.read_excel(source,sheet_name=0,dtype=str,keep_default_na=False) if source.suffix=='.xlsx' else pd.read_csv(source,dtype=str,keep_default_na=False,encoding='utf-8-sig'))
    age=pd.to_numeric(raw.ADQ1)
    mask=age.isin([2,3]) if year==2025 else age.ge(60)
    actual=pool[pool.survey_year.eq(year)]
    assert len(actual)==int(mask.sum())==r['age60_rows']
    id_col='응답ID' if year==2025 else 'ID'
    ids=(str(year)+'_older_'+raw.loc[mask,id_col]).tolist()
    assert ids==actual.record_id.tolist()
    cols=([f'Q6_{i}' for i in range(1,8)]+[f'Q7_{i}' for i in range(1,8)] if year>=2024 else [f'Q4A{i}' for i in range(1,8)]+[f'Q5A{i}' for i in range(1,8)])
    items=raw.loc[mask,cols].apply(pd.to_numeric,errors='coerce')
    complete=items.isin([1,2,3,4]).all(axis=1)
    # Compare integer sums to avoid floating point ambiguity at the boundary.
    expected=np.where(~complete,'',np.where(items.sum(axis=1)<14*meta['threshold'],'at_risk','general'))
    assert expected.tolist()==actual.risk_class.tolist()
    if year>=2024:
        eff=raw.loc[mask,[f'Q22_{i}' for i in range(1,4)]].apply(pd.to_numeric).mean(axis=1)
        np.testing.assert_allclose(eff,pd.to_numeric(actual.self_efficacy_mean3))
        expected_training=np.where(pd.to_numeric(raw.loc[mask,'Q30']).eq(1),'yes','no')
        assert expected_training.tolist()==actual.training_experience.tolist()
    verified.append({'year':year,'rows_reconciled':len(actual),'labels_recomputed':True})

models={}
schema_reference=None
for tag in ['train_2024','test_2025','combined_2024_2025']:
    path=OUT/'03_모델입력'
    csv_data=pd.read_csv(path/(tag+'.csv'))
    text=(path/(tag+'.arff')).read_text(encoding='utf-8')
    attributes=re.findall(r'^@attribute (\w+) (.+)$',text,re.MULTILINE)
    if schema_reference is None: schema_reference=attributes
    assert attributes==schema_reference
    assert csv_data.columns.tolist()==[a[0] for a in attributes]==meta['model_features']+['risk_class']
    rows=list(csv.reader(io.StringIO(text.split('@data\n',1)[1]),quotechar="'"))
    assert len(rows)==len(csv_data)
    for i,(col,spec) in enumerate(attributes):
        values=[r[i] for r in rows]
        if spec=='numeric': np.testing.assert_allclose(csv_data[col],[float(x) for x in values])
        else:
            domain=next(csv.reader([spec[1:-1]],quotechar="'"))
            assert set(values)<=set(domain)
            assert csv_data[col].tolist()==values,(tag,col)
    keys=pd.read_csv(path/(tag+'_row_keys.csv'))
    assert len(keys)==len(csv_data) and keys.record_id.is_unique
    lookup=pool.set_index('record_id').loc[keys.record_id,csv_data.columns]
    lookup['self_efficacy_mean3']=pd.to_numeric(lookup.self_efficacy_mean3)
    pd.testing.assert_frame_equal(csv_data.reset_index(drop=True),lookup.reset_index(drop=True),check_dtype=False)
    assert not csv_data.isna().any().any()
    assert not any('ability' in c or c in ['record_id','survey_year','survey_weight'] for c in csv_data)
    models[tag]={'rows':len(csv_data),'columns':len(csv_data.columns),'class_counts':csv_data.risk_class.value_counts().to_dict(),'arff_matches_csv':True}
train_keys=pd.read_csv(OUT/'03_모델입력/train_2024_row_keys.csv')
test_keys=pd.read_csv(OUT/'03_모델입력/test_2025_row_keys.csv')
assert set(train_keys.survey_year)=={2024} and set(test_keys.survey_year)=={2025}
assert set(train_keys.record_id).isdisjoint(test_keys.record_id)
for r in pd.read_csv(OUT/'자료목록.csv').to_dict('records'):
    assert hashlib.sha256((SRC/r['path']).read_bytes()).hexdigest()==r['sha256']
book=ROOT/'outputs/01a17310-8737-7eb2-91b8-fa1cf953edbb/전처리_점검요약.xlsx'
wb=openpyxl.load_workbook(book,data_only=True,read_only=True)
assert wb['전처리 요약']['D17'].value==len(pool)
assert wb['전처리 요약']['C20'].value==models['train_2024']['rows']
assert wb['전처리 요약']['C21'].value==models['test_2025']['rows']
wb.close()
result={'source_reconciliation':verified,'models':models,'source_hashes_unchanged':True,'xlsx_totals_reopened':True,'temporal_split_verified':True}
(OUT/'04_점검/verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False,indent=2))
