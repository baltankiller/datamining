"""Reproducible NIA survey preparation. Raw files are read only.

Run: python preprocess_survey.py
Requires pandas and openpyxl. Outputs CSV/ARFF plus an audit report.
"""
from pathlib import Path
import argparse
import hashlib
import json
import re
import csv
import numpy as np
import pandas as pd
import openpyxl

ROOT = Path(__file__).resolve().parent
SRC = ROOT / '데이터분석 데이터'
OUT = ROOT / '전처리결과'
REGIONS = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주']
DOMAINS = {
    'age_band': ['60s','70plus'],
    'region': [f'r{i:02}' for i in range(1,18)],
    'income_band6': ['lt100','100_199','200_299','300_399','400_599','600plus'],
    'pc_access': ['no','yes'],
    'smartphone_access': ['no','yes'],
    'self_efficacy_mean3': None,
    'training_experience': ['no','yes'],
    'sex': ['male','female'],
    'school_level': ['primary_or_less','middle','high','college_plus'],
    'residence_type': ['dong','eup_myeon'],
    'risk_class': ['at_risk','general'],
}

def year_of(p):
    # First year is the survey year. A later date can be the publication date.
    return int(re.search(r'20\d{2}', str(p.relative_to(SRC))).group())

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def norm_code(v):
    s = str(v).strip()
    try:
        x = float(s)
        if x.is_integer(): return str(int(x))
    except (TypeError, ValueError): pass
    return s

def read_codebook(p):
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    labels, values = {}, {}
    for ws in wb:
        if '정보' in ws.title:
            for r in ws.iter_rows(min_row=3, values_only=True):
                if r[0] and isinstance(r[1], (int,float)):
                    labels[str(r[0]).strip()] = str(r[2] or '')
        elif '값' in ws.title:
            current = None
            for r in ws.iter_rows(min_row=3, values_only=True):
                if r[0] is not None: current = str(r[0]).strip()
                if current and r[1] is not None:
                    values.setdefault(current,{})[norm_code(r[1])] = str(r[2] or '')
    wb.close()
    return labels, values

def save_csv(df, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path,index=False,encoding='utf-8-sig',lineterminator='\n')

def write_arff(df,path,threshold):
    lines = [f'% Research proxy: mean of 14 ability items < {threshold}. Not an official NIA grade.',
             '% Target is last. IDs, weights, year and all target components are excluded.',
             '@relation nia_older60_risk_proxy','']
    for c in df.columns:
        domain = DOMAINS[c]
        spec = 'numeric' if domain is None else '{'+','.join("'"+s+"'" for s in domain)+'}'
        lines.append(f'@attribute {c} {spec}')
    lines += ['','@data']
    for row in df.itertuples(index=False,name=None):
        cells=[]
        for c,v in zip(df.columns,row):
            if pd.isna(v): cells.append('?')
            elif DOMAINS[c] is None: cells.append(format(float(v),'.15g'))
            else:
                assert v in DOMAINS[c],(c,v)
                cells.append("'"+str(v)+"'")
        lines.append(','.join(cells))
    path.write_text('\n'.join(lines)+'\n',encoding='utf-8')

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--threshold',type=float,default=2.5)
    args=parser.parse_args()
    if not 1 < args.threshold < 4: raise ValueError('threshold must be between 1 and 4')
    OUT.mkdir(exist_ok=True)
    inventory=[]
    for p in sorted(SRC.rglob('*')):
        if p.is_file(): inventory.append({'path':str(p.relative_to(SRC)),'year':year_of(p),'bytes':p.stat().st_size,'sha256':sha(p)})
    save_csv(pd.DataFrame(inventory),OUT/'자료목록.csv')
    older=[p for p in sorted(SRC.rglob('*')) if p.is_file() and any(x in p.name for x in ['고령','장노년'])]
    summaries=[]; dictionaries=[]; profile=[]; issues=[]; maps=[]; frames=[]
    for year in range(2019,2026):
        candidates=[p for p in older if year_of(p)==year]
        raw=[p for p in candidates if p.suffix.lower()=='.csv']
        if not raw: raw=[p for p in candidates if p.suffix.lower()=='.xlsx' and 'data' in p.name.lower()]
        assert len(raw)==1,(year,raw)
        raw=raw[0]
        cb=raw if raw.suffix=='.xlsx' else next(p for p in candidates if p.suffix.lower()=='.xlsx')
        labels,value_labels=read_codebook(cb)
        if raw.suffix=='.csv':
            encoding=None
            for enc in ['utf-8-sig','cp949','euc-kr']:
                try:
                    df=pd.read_csv(raw,encoding=enc,dtype=str,keep_default_na=False)
                    encoding=enc
                    break
                except UnicodeDecodeError: continue
            if encoding is None: raise ValueError(f'Cannot decode {raw}')
            with raw.open(encoding=encoding,newline='') as stream:
                reader=csv.reader(stream); header=next(reader)
                assert len(header)==len(set(header)), f'Duplicate headers: {raw}'
                assert all(len(r)==len(header) for r in reader), f'Ragged CSV: {raw}'
            raw_sheet='CSV'
        else:
            df=pd.read_excel(raw,sheet_name=0,dtype=str,keep_default_na=False)
            encoding='xlsx'
            with pd.ExcelFile(raw) as book: raw_sheet=book.sheet_names[0]
        df.columns=[str(c).strip() for c in df.columns]
        assert df.columns.is_unique
        before=df.copy()
        # Standardize whitespace only. Literal NA/0/9/99 are never globally missing.
        for c in df:
            df[c]=df[c].astype('string').str.strip().replace('',pd.NA)
        blank_cells=int(df.isna().sum().sum())
        declared_missing=0
        for c in df:
            values=value_labels.get(c,{})
            missing=[k for k,v in values.items() if re.search(r'무응답|모름|모르겠|비해당|해당\s*없음',v)]
            if missing:
                mask=df[c].map(lambda v: norm_code(v) in missing if pd.notna(v) else False)
                declared_missing+=int(mask.sum())
                df.loc[mask,c]=pd.NA
            dictionaries.append({'year':year,'variable':c,'label':labels.get(c,''),'value_labels':json.dumps(values,ensure_ascii=False),'source':str(cb.relative_to(SRC))})
            if values:
                unexpected=df[c].dropna().map(norm_code)
                unexpected=unexpected[~unexpected.isin(values)]
                for value,n in unexpected.value_counts().items():
                    issues.append({'year':year,'variable':c,'kind':'code_not_in_codebook','value':value,'count':int(n),'action':'원문항 유지. 분석용 후보는 허용범위를 별도로 검증'})
            profile.append({'year':year,'variable':c,'rows':len(df),'missing':int(df[c].isna().sum()),'missing_rate':float(df[c].isna().mean()),'distinct_nonmissing':int(df[c].nunique())})
        id_col='응답ID' if '응답ID' in df else 'ID'
        assert df[id_col].notna().all() and df[id_col].is_unique,(year,'ID must be unique')
        if 'TYP' in df: assert set(df.TYP.dropna())=={'7'},(year,'Not an older-adult sample')
        df.insert(0,'survey_year',year)
        df.insert(1,'record_id',str(year)+'_older_'+df[id_col])
        save_csv(df,OUT/f'01_연도별정리/{year}_고령층_전체문항.csv')
        z=pd.DataFrame(index=df.index)
        z['record_id']=df.record_id
        z['survey_year']=year
        z['source_row']=np.arange(len(df))+2
        def numeric(c, allowed=None):
            if c not in df: return pd.Series(np.nan,index=df.index)
            x=pd.to_numeric(df[c],errors='coerce').astype(float)
            bad=df[c].notna() & x.isna()
            if allowed is not None: bad|=x.notna() & ~x.isin(allowed)
            if bad.any(): issues.append({'year':year,'variable':c,'kind':'candidate_invalid','value':json.dumps(df.loc[bad,c].value_counts().to_dict(),ensure_ascii=False),'count':int(bad.sum()),'action':'분석용 파생변수만 결측으로 처리, 원문항 유지'})
            return x.mask(bad)
        age=numeric('ADQ1',range(1,4) if year==2025 else range(55,121))
        z['age_years']=age if year<2025 else np.nan
        z['age_band']=age.map({1:'50s',2:'60s',3:'70plus'}) if year==2025 else pd.cut(age,[49,59,69,120],labels=['50s','60s','70plus']).astype('string')
        z['eligible_age60']=z.age_band.isin(['60s','70plus'])
        z['sex']=numeric('ADQ2',[1,2]).map({1:'male',2:'female'})
        z['region']=numeric('ADQ101',range(1,18)).map({i:f'r{i:02}' for i in range(1,18)})
        z['region_name']=z.region.map({f'r{i:02}':n for i,n in enumerate(REGIONS,1)})
        inc=numeric('ADQ9',range(1,12))
        early={1:1,2:1,3:2,4:2,5:3,6:3,7:4,8:4,9:5,10:5,11:6}
        late={1:1,2:2,3:3,4:4,5:5,6:5,7:6,8:6,9:6,10:6,11:6}
        z['income_source_code']=inc
        z['income_band6']=inc.map(early if year<=2020 else late).map(dict(enumerate(DOMAINS['income_band6'],1)))
        z['school_level']=numeric('ADQ4',[1,2,3,4]).map(dict(enumerate(DOMAINS['school_level'],1)))
        z['residence_type']=numeric('ADQ103',[1,2]).map({1:'dong',2:'eup_myeon'})
        pc_cols=['Q1_1','Q1_2'] if year>=2024 else ['Q1A1','Q1A2']
        pc=pd.concat([numeric(c,[1,2]) for c in pc_cols],axis=1)
        z['pc_access']=pd.Series(pd.NA,index=df.index,dtype='string')
        z.loc[pc.eq(1).any(axis=1),'pc_access']='yes'
        z.loc[pc.eq(2).all(axis=1),'pc_access']='no'
        phone_cols=([f'Q2K1_1_{i}' for i in range(1,4)] if year>=2024 else ['Q2A1R'] if year==2021 else [c for c in ['Q2A11','Q2A12','Q2A13'] if c in df])
        phone=pd.concat([numeric(c,[1,2,3]) for c in phone_cols],axis=1)
        z['smartphone_access']=pd.Series(pd.NA,index=df.index,dtype='string')
        z.loc[phone.eq(1).any(axis=1),'smartphone_access']='yes'
        # Multiple slots may be blank when no second/third phone exists.
        z.loc[phone.isin([2,3]).any(axis=1) & ~phone.eq(1).any(axis=1),'smartphone_access']='no'
        pcols=[f'Q6_{i}' if year>=2024 else f'Q4A{i}' for i in range(1,8)]
        mcols=[f'Q7_{i}' if year>=2024 else f'Q5A{i}' for i in range(1,8)]
        ability=pd.concat([numeric(c,[1,2,3,4]) for c in pcols+mcols],axis=1)
        ability.columns=[f'pc_ability_{i}' for i in range(1,8)]+[f'mobile_ability_{i}' for i in range(1,8)]
        z=pd.concat([z,ability],axis=1)
        z['ability_valid_items']=ability.notna().sum(axis=1)
        z['ability_mean14']=ability.mean(axis=1).where(z.ability_valid_items==14)
        z['risk_class']=z.ability_mean14.map(lambda v: pd.NA if pd.isna(v) else 'at_risk' if v<args.threshold else 'general').astype('string')
        z['label_reason']=np.where(z.risk_class.isna(),'incomplete_or_invalid_ability','complete_14_items')
        effcols=[f'Q22_{i}' if year>=2024 else f'Q18A{i}' for i in range(1,4)] if year>=2021 else []
        if effcols:
            eff=pd.concat([numeric(c,[1,2,3,4]) for c in effcols],axis=1)
            z['self_efficacy_mean3']=eff.mean(axis=1).where(eff.notna().sum(axis=1)==3)
        else:
            # Different constructs and damaged 2019 labels prevent a justified harmonization.
            z['self_efficacy_mean3']=np.nan
        training=numeric('Q30',[1,2,3,4]) if year>=2024 else pd.Series(np.nan,index=df.index)
        z['training_source_code']=training
        z['training_experience']=training.map({1:'yes',2:'no',3:'no',4:'no'})
        z['self_efficacy_available']=year>=2021
        z['training_question_available']=year>=2024
        weight_col='WT_T' if 'WT_T' in df else 'WT'
        z['survey_weight']=numeric(weight_col)
        z['weight_source_variable']=weight_col
        z['source_file']=str(raw.relative_to(SRC))
        z['source_sheet']=raw_sheet
        frames.append(z)
        n60=int(z.eligible_age60.sum())
        selected=z.loc[z.eligible_age60]
        summaries.append({'year':year,'raw_rows':len(df),'raw_columns':before.shape[1],'age60_rows':n60,'labelled_age60':int(selected.risk_class.notna().sum()),'at_risk':int(selected.risk_class.eq('at_risk').sum()),'general':int(selected.risk_class.eq('general').sum()),'unlabelled':int(selected.risk_class.isna().sum()),'blank_cells':blank_cells,'declared_missing_cells':declared_missing,'duplicate_ids':0,'duplicate_response_rows':int(before.drop(columns=[id_col]).duplicated().sum()),'age_format':'age_band' if year==2025 else 'years','encoding':encoding,'source_file':str(raw.relative_to(SRC)),'codebook':str(cb.relative_to(SRC))})
        mapping={'age_band':'ADQ1','region':'ADQ101','income_band6':'ADQ9','pc_access':','.join(pc_cols),'smartphone_access':','.join(phone_cols),'self_efficacy_mean3':','.join(effcols),'training_experience':'Q30' if year>=2024 else '', 'ability_mean14':','.join(pcols+mcols),'survey_weight':weight_col}
        maps.extend({'year':year,'prepared_variable':v,'source_variables':s} for v,s in mapping.items())
        print(year,'source',len(df),'age60',n60,'labelled',int(selected.risk_class.notna().sum()),flush=True)
    all_rows=pd.concat(frames,ignore_index=True)
    assert all_rows.record_id.is_unique
    selected=all_rows.loc[all_rows.eligible_age60].copy()
    save_csv(selected,OUT/'02_통합자료/고령자60세이상_2019_2025_전처리.csv')
    # No ID, year, weight, raw ability items or derived target scores in model inputs.
    main_df=selected.loc[selected.survey_year.isin([2024,2025]) & selected.risk_class.notna()]
    model_dir=OUT/'03_모델입력'; model_dir.mkdir(exist_ok=True)
    cols=list(DOMAINS)
    for tag,ys in [('train_2024',[2024]),('test_2025',[2025]),('combined_2024_2025',[2024,2025])]:
        subset=main_df.loc[main_df.survey_year.isin(ys)]
        model=subset[cols]
        save_csv(model,model_dir/(tag+'.csv'))
        write_arff(model,model_dir/(tag+'.arff'),args.threshold)
        save_csv(subset[['record_id','survey_year','source_row','source_file','source_sheet','survey_weight']],model_dir/(tag+'_row_keys.csv'))
        save_csv(pd.DataFrame({'variable':cols,'missing':model.isna().sum().values,'rows':len(model)}),model_dir/(tag+'_missingness.csv'))
    summary=pd.DataFrame(summaries)
    save_csv(summary,OUT/'04_점검/연도별_점검요약.csv')
    save_csv(pd.DataFrame(dictionaries),OUT/'04_점검/원문항_코드북.csv')
    save_csv(pd.DataFrame(profile),OUT/'04_점검/원문항_결측점검.csv')
    save_csv(pd.DataFrame(issues,columns=['year','variable','kind','value','count','action']),OUT/'04_점검/코드_확인사항.csv')
    save_csv(pd.DataFrame(maps),OUT/'04_점검/연도별_변수대응표.csv')
    descriptions=[
        ('age_band','연령대','60s=60대, 70plus=70대 이상','2025년과 맞추기 위해 실제 나이 대신 공통 연령대 사용'),
        ('region','거주 시도','r01~r17','지역명은 통합자료의 region_name 또는 코드북 확인'),
        ('income_band6','가구 월평균 소득','100만원 미만 / 100~199 / 200~299 / 300~399 / 400~599 / 600만원 이상','명목형 6구간. 2019~2020년의 다른 구간을 정확히 묶어 통일'),
        ('pc_access','필요시 PC 이용 가능 여부','no / yes','데스크톱·노트북 중 하나 이상 이용 가능'),
        ('smartphone_access','스마트폰 이용 가능 여부','no / yes','휴대전화 응답 슬롯 중 하나라도 스마트폰'),
        ('self_efficacy_mean3','디지털기기 이용 효능감','1~4점','자신감 및 사용법 습득 3문항의 완전응답 평균'),
        ('training_experience','정보화교육 참여 경험','no / yes','Q30=1은 yes, Q30=2~4는 no. 2024~2025년만 확인'),
        ('sex','성별','male / female','ADQ2'),
        ('school_level','최종 학력','초등졸 이하 / 중졸 / 고졸 / 대졸 이상','ADQ4. 네 가지 명목 범주'),
        ('residence_type','거주지 형태','dong / eup_myeon','ADQ103. 동 / 읍면'),
        ('risk_class','연구용 소외 등급','at_risk / general',f'14문항 평균 {args.threshold} 미만은 소외. 공식 NIA 등급 아님'),
    ]
    variable_dict=pd.DataFrame(descriptions,columns=['variable','meaning','values','definition'])
    variable_dict['role']=['target' if v=='risk_class' else 'predictor' for v in variable_dict.variable]
    save_csv(variable_dict,OUT/'04_점검/분석변수_설명.csv')
    threshold_rows=[]
    for year,g in selected.groupby('survey_year'):
        for t in [2.0,2.5,3.0]:
            valid=g.ability_mean14.dropna()
            threshold_rows.append({'year':int(year),'threshold':t,'labelled':len(valid),'at_risk':int(valid.lt(t).sum()),'at_risk_rate':float(valid.lt(t).mean())})
    save_csv(pd.DataFrame(threshold_rows),OUT/'04_점검/등급기준_민감도.csv')
    # Verify original sources have not changed during preparation.
    assert all(sha(SRC/r['path'])==r['sha256'] for r in inventory)
    metadata={'min_age':60,'threshold':args.threshold,'required_ability_items':14,'age60_total':len(selected),'main_total':len(main_df),'model_features':cols[:-1],'summary':summaries,'variable_dictionary':variable_dict.to_dict(orient='records'),'source_files':len(inventory),'originals_unchanged':True,'issues':issues,'main_missingness':main_df[cols].isna().sum().to_dict()}
    (OUT/'04_점검/run_metadata.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
    lines=['# 고령자 디지털 소외 위험 예측 전처리',
        '',f'원본 {len(inventory)}개 파일을 보존하고, 7개 연도의 고령층 파일을 처리했습니다. 60세 이상 {len(selected):,}행을 정리했습니다.',
        '', '## 먼저 사용할 파일',
        '- `03_모델입력/train_2024.arff`: Weka 학습용. 마지막 속성 risk_class를 class로 지정합니다.',
        '- `03_모델입력/test_2025.arff`: Weka Supplied test set. 학습용과 동일한 속성 순서·범주입니다.',
        '- 동일한 이름의 CSV: 파이썬에서 같은 행과 변수로 비교할 때 사용합니다.',
        '- `02_통합자료/고령자60세이상_2019_2025_전처리.csv`: 전체 연도 점검·추가 분석용입니다. 점수와 원문항이 포함되므로 그대로 모델에 넣지 마세요.',
        '- `01_연도별정리/`: 55~59세를 포함하는 원래 고령층 표본 전체 문항을 별도로 보존한 정리본입니다.',
        '- `04_점검/`: 행수, 결측, 코드북, 변수 대응표, 코드 확인사항 및 기준 민감도입니다.',
        '', '## 이번에 적용한 기준',
        '- 대상: 2019~2024년 ADQ1의 실제 나이 60세 이상. 2025년은 코드북의 ADQ1=2(60대), 3(70대 이상). 2025년에 실제 나이를 추정해 채우지 않았습니다.',
        f'- 연구용 소외 등급: PC 이용능력 7문항 + 모바일 이용능력 7문항의 1~4점 단순평균이 {args.threshold}점 미만이면 at_risk(소외), 이상이면 general(일반). 공식 NIA 등급 또는 공식 종합지수로 해석하지 않습니다.',
        '- 14문항 중 하나라도 결측·범위 밖이면 등급을 만들지 않습니다. 해당 응답은 통합자료에 남기고 모델용 파일에서만 제외합니다.',
        '- 등급 산출 문항과 능력 평균은 모델 입력에서 모두 제외했습니다. 입력변수로 이용능력을 꼭 사용해야 한다면 종속변수를 별도로 재정의해야 합니다.',
        '- 효능감: 2021~2025년 자신감·습득 능력 3문항의 완전응답 평균. 4번째 문항(더 많이 이용하고 싶다)은 이용 의향이라 제외했습니다. 2019~2020년 유사 태도 문항을 같은 척도라고 가정하지 않았습니다.',
        '- 교육 경험: 2024~2025년 Q30=1은 yes, 2~4는 no. 인지 여부·희망 강좌를 교육 경험으로 바꾸지 않았습니다. 이전 연도는 이 파생변수를 결측으로 유지했습니다.',
        '- 소득: 월 가구소득을 100만원 미만 / 100~199 / 200~299 / 300~399 / 400~599 / 600만원 이상으로 일치시켰습니다. 소득코드 원값을 연도 간 그대로 합치거나 구간 중간값을 실제 소득으로 사용하지 않았습니다.',
        '- PC 보유는 필요할 때 이용 가능한 데스크톱·노트북 중 하나라도 있으면 yes. 스마트폰도 이용 가능한 기기 응답 중 하나라도 스마트폰이면 yes입니다. 개인 소유와 동일하지 않습니다.',
        '- 빈칸·공백은 결측으로 정리했습니다. 코드북에 명시된 무응답·모름 등만 변수별로 결측 처리했습니다. 0·9·99를 전역적으로 제거하지 않았습니다.',
        '- 모델용 후보 변수는 허용범위도 검사했습니다. 2019년 이용능력의 9 등 범위 밖 값은 원문항에 유지하고 파생점수 계산에서만 제외했습니다.',
        '- ID는 연도와 조사집단을 포함해 새 키를 만들었습니다. 중복 ID를 검사했으며 응답값이 비슷하다는 이유로 응답자를 삭제하지 않았습니다.',
        '- 코드북 값 목록과 일치하지 않는 값은 확인사항 파일에 기록했습니다. 이를 모두 데이터 오류라고 단정해 삭제하지 않았습니다. 2019년 성별·지역의 별표(*) 각 123건은 분석용 파생변수에서만 결측 처리했습니다. 2025년 ADQ8_1T처럼 이번 입력변수에 포함되지 않는 항목의 값 불일치도 원문항 그대로 남겼습니다.',
        '', '## 모델 비교에 쓰는 입력변수',
        'age_band(60대/70대 이상), region(시도), income_band6(가구소득 구간), pc_access, smartphone_access, self_efficacy_mean3, training_experience, sex, school_level, residence_type(동/읍면)입니다.',
        '', '## 분석 시 지킬 분할과 해석',
        '- 학습: 2024년, 최종 평가: 2025년. combined 파일 전체를 대상으로 튜닝한 뒤 같은 2025년으로 평가하면 안 됩니다.',
        '- 결측치 대체·원핫인코딩·표준화·변수선택·오버샘플링을 한다면 학습 폴드 안에서만 학습합니다. 미리 전체 자료에 적용하지 않았습니다.',
        '- Weka에서는 CSV보다 ARFF를 사용하면 지역·연령대 등의 명목형이 유지됩니다. Logistic/J48/RandomForest에 동일한 train/test를 사용하세요. 결측 처리 필터가 필요하면 FilteredClassifier 안에 두고 학습자료에서 학습시킵니다.',
        '- 파이썬에서는 문자 범주를 OneHotEncoder(handle_unknown="ignore")로 처리하고, 필요한 SimpleImputer/StandardScaler는 Pipeline 안에 둡니다.',
        '- 조사연도와 응답 ID는 예측변수가 아닙니다. 가중치는 row_keys 파일과 통합자료에 보존했습니다. 제공 ARFF에는 가중치를 적용하지 않았으므로 단순 표본 성능입니다.',
        '- 반복 횡단면 조사로 취급합니다. 연도 간 같은 ID를 동일인으로 연결하지 않습니다. 지금 만든 등급은 같은 시점의 자기보고 이용능력을 분류하는 대리지표이며 향후 소외 발생을 추적한 결과가 아닙니다.',
        '- 권장 평가지표: 소외 집단 recall, precision, F1, PR-AUC, ROC-AUC, 혼동행렬. 지역별 표본 수·가중치·오차를 함께 검토해야 하며, 모형 확률만으로 개인별 교육 우선순위를 확정하지 않습니다.',
        '- 표본 수는 기본 분류모델 실험을 시작할 수 있는 규모입니다. 지역 세부 단위나 일부 범주에는 표본이 적을 수 있습니다. 다른 자료는 우선 추가하지 않았습니다.',
        '', '## 연도별 행수',
        '', '| 연도 | 원본 행수 | 60세 이상 | 등급 생성 | 소외 | 일반 | 등급 없음 |',
        '|---|---:|---:|---:|---:|---:|---:|']
    lines += [f"| {r['year']} | {r['raw_rows']:,} | {r['age60_rows']:,} | {r['labelled_age60']:,} | {r['at_risk']:,} | {r['general']:,} | {r['unlabelled']:,} |" for r in summaries]
    lines += ['', '## 재실행', '', '프로젝트 폴더에서 `python preprocess_survey.py`를 실행합니다. 임시 기준 변경은 `--threshold 2.5` 옵션을 사용합니다. 원본은 수정하지 않으며 전처리결과 파일들은 다시 생성됩니다.', '', '패키지: pandas, numpy, openpyxl. 원본 파일별 SHA-256을 기록하고 실행 종료 시 원본 불변 여부를 다시 확인했습니다.']
    (OUT/'전처리_안내.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print(json.dumps({'age60_total':len(selected),'main_total':len(main_df),'main_missingness':metadata['main_missingness'],'issues':len(issues),'originals_unchanged':True},ensure_ascii=False),flush=True)

if __name__=='__main__': main()
