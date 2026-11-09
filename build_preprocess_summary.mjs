import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const root=process.cwd();
const meta=JSON.parse(await fs.readFile(path.join(root,'전처리결과/04_점검/run_metadata.json'),'utf8'));
const output=path.join(root,'outputs/01a17310-8737-7eb2-91b8-fa1cf953edbb');
await fs.mkdir(output,{recursive:true});
const wb=Workbook.create();
const summary=wb.worksheets.add('전처리 요약');
const dict=wb.worksheets.add('분석변수');
for (const s of [summary,dict]) {
  s.showGridLines=false;
  s.getRange('A1:H30').format.font={name:'Arial',size:11,color:'#172033'};
  s.getRange('A1:H30').format.verticalAlignment='center';
  s.getRange('A1:H30').format.rowHeight=25;
}
summary.getRange('A2').values=[['고령자 디지털 소외 위험 예측 전처리']];
summary.getRange('A2').format.font={name:'Arial',size:16,bold:true};
summary.getRange('A4:B6').values=[['대상','60세 이상'],['등급 기준',`이용능력 14문항 평균 < ${meta.threshold}`],['등급 해석','연구용 대리지표. 공식 NIA 등급 아님']];
summary.getRange('A8:H8').values=[['연도','원본 행수','원본 열수','60세 이상','소외','일반','등급 없음','소외 비율']];
summary.getRange('A9:G15').values=meta.summary.map(r=>[r.year,r.raw_rows,r.raw_columns,r.age60_rows,r.at_risk,r.general,r.unlabelled]);
summary.getRange('H9').formulas=[['=E9/SUM(E9:F9)']];
summary.getRange('H9:H15').fillDown();
summary.getRange('A17:G17').values=[['합계',null,null,null,null,null,null]];
for (const c of ['B','D','E','F','G']) summary.getRange(`${c}17`).formulas=[[`=SUM(${c}9:${c}15)`]];
summary.getRange('H17').formulas=[['=E17/SUM(E17:F17)']];
summary.getRange('B9:G17').setNumberFormat('#,##0');
summary.getRange('A9:A15').setNumberFormat('0');
summary.getRange('H9:H17').setNumberFormat('0.0%');
summary.getRange('A19:D22').values=[
  ['모델용 자료','연도','행수','용도'],
  ['train_2024',2024,meta.summary.find(r=>r.year===2024).labelled_age60,'학습 및 교차검증'],
  ['test_2025',2025,meta.summary.find(r=>r.year===2025).labelled_age60,'최종 평가'],
  ['입력변수',null,meta.model_features.length,'등급 산출 문항·점수 제외']
];
summary.getRange('B20:B21').setNumberFormat('0');
summary.getRange('C20:C22').setNumberFormat('#,##0');
summary.getRange('D20:D22').format.horizontalAlignment='left';
summary.getRange('A24').values=[['원본 파일과 코드북: 제공된 디지털정보격차 실태조사 2019~2025년 고령층 자료']];
summary.getRange('A25').values=[['2025년은 연령대만 공개. 60대·70대 이상을 선택하고 실제 나이는 추정하지 않음.']];
summary.getRange('A26').values=[['교육 경험을 확인할 수 있는 2024~2025년 자료를 모델용으로 구성. 입력변수 결측 0건.']];
summary.getRange('A27').values=[['2019년 5행은 이용능력 응답이 불완전해 등급을 부여하지 않고 통합자료에 보존.']];
summary.getRange('A24:A27').format.font={name:'Arial',size:10,color:'#4B5563'};
summary.getRange('A8:H8').format={fill:'#24364B',font:{name:'Arial',size:11,color:'#FFFFFF',bold:true},horizontalAlignment:'center'};
summary.getRange('A19:D19').format={fill:'#24364B',font:{name:'Arial',size:11,color:'#FFFFFF',bold:true},horizontalAlignment:'center'};
summary.getRange('A17:H17').format.fill='#E9EFF5';
summary.getRange('A17:H17').format.font={name:'Arial',size:11,bold:true};
summary.getRange('A1:A27').format.columnWidth=23;
summary.getRange('B1:B27').format.columnWidth=20;
summary.getRange('C1:H27').format.columnWidth=15;
summary.getRange('D1:D27').format.columnWidth=28;
summary.tabColor='#24364B';

dict.getRange('A2').values=[['모델 입력변수와 등급 정의']];
dict.getRange('A2').format.font={name:'Arial',size:16,bold:true};
dict.getRange('A4:E4').values=[['변수','의미','값 / 범위','정의','역할']];
dict.getRange('A5:E15').values=meta.variable_dictionary.map(r=>[r.variable,r.meaning,r.values,r.definition,r.role==='target'?'종속변수':'입력변수']);
dict.getRange('A4:E4').format={fill:'#24364B',font:{name:'Arial',size:11,color:'#FFFFFF',bold:true},horizontalAlignment:'center'};
dict.getRange('A1:A20').format.columnWidth=28;
dict.getRange('B1:B20').format.columnWidth=30;
dict.getRange('C1:C20').format.columnWidth=53;
dict.getRange('D1:D20').format.columnWidth=68;
dict.getRange('E1:E20').format.columnWidth=14;
dict.getRange('A5:E15').format.wrapText=true;
dict.getRange('A5:E15').format.rowHeight=48;
dict.getRange('A17').values=[['ARFF에서는 지역·연령대 등 범주형을 명목형으로 지정. 종속변수는 마지막 열.']];
dict.getRange('A18').values=[['개인 소유 대신 필요할 때 이용 가능한 기기를 조사한 응답임. 소득은 가구 월평균 기준.']];
dict.freezePanes.freezeRows(4);

wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'전처리 요약!A9:H17',include:'values,formulas',tableMaxRows:10,tableMaxCols:8,maxChars:2000})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},maxChars:1200})).ndjson);
for (const [s,range,name] of [[summary,'A1:H27','summary'],[dict,'A1:E18','variables']]) {
  const img=await wb.render({sheetName:s.name,range,scale:1.3,format:'png'});
  await fs.writeFile(path.join(output,`${name}.png`),new Uint8Array(await img.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(wb)).save(path.join(output,'전처리_점검요약.xlsx'));
console.log('Saved workbook and two rendered previews.');
