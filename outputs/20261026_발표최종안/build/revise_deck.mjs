import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {FileBlob, PresentationFile} from '@oai/artifact-tool';

const task='C:/Users/UserPc/Documents/ChatGPT/빅데이터분석/outputs/20261026_발표최종안';
const skill='C:/Users/UserPc/.codex/plugins/cache/openai-primary-runtime/presentations/26.1025.11240/skills/presentations';
const source='C:/Users/UserPc/Downloads/고령자 디지털 소외 위험 예측 중간보고 발표자료.pptx';
const py='C:/Users/UserPc/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const content=JSON.parse(await fs.readFile(new URL('./content.json',import.meta.url),'utf8'));
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const raw=await p.inspect({kind:'slide,textbox,shape,image,notes',maxChars:1000000});
const records=raw.ndjson.split('\n').filter(Boolean).map(x=>JSON.parse(x));
const find=(slide,name)=>{const r=records.find(r=>r.slideIndex===slide-1&&r.name===name&&['textbox','shape'].includes(r.kind));if(!r)throw new Error(`Missing ${slide} ${name}`);return p.resolve(r.id);};
const edit=(slide,name,text,opts={})=>{
  const original=records.find(r=>r.slideIndex===slide-1&&r.name===name&&['textbox','shape'].includes(r.kind)).position;
  const shape=find(slide,name);shape.text=text;
  shape.text.style={typeface:'Malgun Gothic',color:'#12377A',fontSize:opts.fontSize??18,bold:opts.bold??false,autoFit:'none',insets:0,wrap:'square'};
  if(opts.color)shape.text.color=opts.color;
  shape.position={...original,...opts.position};
  return shape;
};
const remove=(slide,name,kind)=>{const r=records.find(r=>r.slideIndex===slide-1&&r.name===name&&(!kind||r.kind===kind));if(r)p.delete(r.id);};
const addText=(slide,text,position,fontSize=12,color='#667085')=>{const s=p.slides.getItem(slide-1).shapes.add({name:`final-note-${slide}`,geometry:'textbox',position,fill:'none',line:{fill:'none',width:0}});s.text=text;s.text.style={typeface:'Malgun Gothic',fontSize,color,insets:0,autoFit:'none'};return s;};

// Keep the source theme, backgrounds, panels and positions, while explicitly
// replacing unsupported source fonts in the Korean text with an installed font.
for(const r of records.filter(r=>r.kind==='textbox')){const s=p.resolve(r.id);s.text.style={typeface:'Malgun Gothic'};s.position={...r.position};}

edit(1,'Text 1','고령자 디지털 소외\n위험 예측',{fontSize:41.33,bold:true});
edit(1,'Text 2','중간보고',{fontSize:20,color:'#3D8BC1'});
edit(1,'Text 3','고령층의 디지털 소외 관련 특성 분석',{fontSize:18,position:{left:66,width:780,height:27}});
edit(1,'Text 4','정보화교육을 우선 안내할 대상자 제안',{fontSize:18,position:{left:66,width:780,height:27}});
edit(1,'Text 5','NIA 2025년 조사 고령층 응답자 2,300명',{fontSize:18,position:{left:66,width:780,height:27}});
edit(1,'Text 6','고령층 디지털 소외 분류 분석 중간보고',{fontSize:10.5,color:'#667085',position:{width:500,height:20}});

edit(2,'Text 0','소외군 정의와 분석 방법',{fontSize:26,bold:true});
edit(2,'Text 2','분석 대상과 분류 기준',{fontSize:20,bold:true,color:'#3D8BC1'});
edit(2,'Text 3','소외군: Q8·Q9 18문항 평균 2.5점 미만',{fontSize:19,position:{left:66,width:470,height:30}});
edit(2,'Text 4','소외군 1,583명(68.8%)\n일반군 717명(31.2%)',{fontSize:18,position:{left:66,top:173,width:465,height:49}});
remove(2,'Text 7');
edit(2,'Text 8','분석 방법',{fontSize:20,bold:true,color:'#3D8BC1'});
edit(2,'Text 9','연령·학력·소득·기기·자신감 등 10개 변수',{fontSize:18,position:{left:66,width:470,height:30}});
edit(2,'Text 10','Weka 10겹 교차검증으로 모델 비교',{fontSize:18,position:{left:66,width:470,height:30}});
edit(2,'Text 15','본 분석의 분류 기준에 따른 표본 내 비율이며, 가중치를 적용하지 않음',{fontSize:12,color:'#667085',position:{top:493,height:28}});
remove(2,'Image 1','image');remove(2,'Text 14');
const pie=p.slides.getItem(1).charts.add('pie',{
  title:'표본 내 집단 비율',titleTextStyle:{fontSize:19,bold:true,fill:'#12377A'},
  position:{left:600,top:118,width:312,height:335},categories:['소외군 1,583명','일반군 717명'],
  series:[{name:'응답자 수',values:[1583,717],points:[{idx:0,fill:'#12377A'},{idx:1,fill:'#A7D4E3'}]}],
  hasLegend:true,legend:{position:'bottom',textStyle:{fontSize:13,fill:'#475467'}},
  dataLabels:{showPercent:true,position:'center',numberFormatCode:'0.0%',textStyle:{fontSize:18,bold:true,fill:'#FFFFFF'}},
  chartFill:'#FFFFFF',chartLine:{fill:'none',width:0},plotAreaLine:{fill:'none',width:0}
});applyPresentationChartFont(pie,{fontFamily:'Malgun Gothic'});

edit(3,'Text 0','로지스틱 회귀를 우선 후보로 제안',{fontSize:26,bold:true,position:{height:38}});
edit(3,'Text 1','비교한 학습 모델 중 정확도와 소외군 재현율이 가장 높음',{fontSize:18,color:'#3D8BC1',position:{top:65,height:27}});
remove(3,'Image 1','image');remove(3,'Text 3');
const bar=p.slides.getItem(2).charts.add('bar',{
  title:'모델별 정확도',titleTextStyle:{fontSize:17,bold:true,fill:'#12377A'},
  position:{left:37,top:109,width:431,height:304},
  categories:['랜덤포레스트','J48','로지스틱 회귀','ZeroR 기준선'],
  series:[{name:'정확도',values:[.799,.814,.819,.688],valuesFormatCode:'0.0%',fill:'#12377A',points:[{idx:0,fill:'#12377A'},{idx:1,fill:'#12377A'},{idx:2,fill:'#137D82'},{idx:3,fill:'#A7D4E3'}]}],
  barOptions:{direction:'bar',grouping:'clustered',gapWidth:65},hasLegend:false,
  xAxis:{textStyle:{fontSize:14,fill:'#344054'},majorGridlines:null},
  yAxis:{min:0,max:1,majorUnit:.25,numberFormatCode:'0%',textStyle:{fontSize:12,fill:'#667085'},line:{fill:'none',width:0},majorGridlines:{fill:'#E1E8EF',width:1}},
  dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0.0%',textStyle:{fontSize:14,bold:true,fill:'#12377A'}},
  chartFill:'#FFFFFF',chartLine:{fill:'none',width:0},plotAreaLine:{fill:'none',width:0}
});applyPresentationChartFont(bar,{fontFamily:'Malgun Gothic'});
edit(3,'Text 6','+13.1%p',{fontSize:34,bold:true,position:{height:44}});
edit(3,'Text 7','ZeroR 기준선 대비 정확도 개선',{fontSize:16,position:{top:156,height:25}});
edit(3,'Text 10','88.9%',{fontSize:34,bold:true,position:{height:44}});
edit(3,'Text 11','실제 소외군 중 찾아낸 비율',{fontSize:16,position:{top:252,height:25}});
edit(3,'Text 13','정확도 81.9%, J48보다 0.5%p 높음',{fontSize:16.5,position:{height:24}});
edit(3,'Text 14','최종 선택 전 정밀도·일반군 성능 확인',{fontSize:16.5,position:{height:24}});
addText(3,'출처: 팀원 보고서의 Weka 10겹 교차검증 결과',{left:30,top:485,width:900,height:26},12);

edit(4,'Text 0','이용 자신감과 학력이 주요 관련 요인',{fontSize:26,bold:true});
edit(4,'Text 2','연령과 학력',{fontSize:20,bold:true,color:'#3D8BC1'});
edit(4,'Text 3','70대 이상 집단',{fontSize:18,color:'#667085'});
edit(4,'Text 4','소외 비율 90.6%',{fontSize:24,bold:true});
edit(4,'Text 5','초졸 이하 집단',{fontSize:18,color:'#667085'});
edit(4,'Text 6','소외 비율 98.9%',{fontSize:24,bold:true});
edit(4,'Text 8','스마트폰 미보유 134명 모두 소외군\n전체 소외군의 약 8.5%',{fontSize:18,position:{top:405,height:54}});
edit(4,'Text 10','이용 자신감과 기기 접근성',{fontSize:20,bold:true,color:'#3D8BC1'});
edit(4,'Text 11','디지털 기기 이용 자신감',{fontSize:18,color:'#667085'});
edit(4,'Text 12','보고서에서 주요 관련 요인으로 제시',{fontSize:22,bold:true});
edit(4,'Text 13','스마트폰 보유 집단',{fontSize:18,color:'#667085'});
edit(4,'Text 14','소외 비율 66.9%',{fontSize:24,bold:true});
edit(4,'Text 15','이용 자신감을 키우는\n실습 교육을 제안',{fontSize:20,position:{top:333,height:64}});
edit(4,'Text 16','집단별 비율은 원본 재집계 결과이며, 다른 변수의 영향을 통제하지 않은 단순 비교',{fontSize:11.5,color:'#667085',position:{height:27}});

edit(5,'Text 0','교육 활용 제안과 후속 과제',{fontSize:26,bold:true});
for(const n of [3,12,21]){const sh=find(5,`Text ${n}`);sh.text.style={typeface:'Malgun Gothic',fontSize:21,bold:true,color:'#12377A',insets:0};}
for(const [name,text] of [['Text 4','여러 특성 함께 고려'],['Text 5','일상 과제 반복 실습'],['Text 6','작은 성공 경험 제공'],['Text 13','교육 경험자 55명'],['Text 14','관련성 중심의 분석'],['Text 15','교육 효과는 별도 검증'],['Text 22','분류 기준 근거 보완'],['Text 23','상세 성능 확인'],['Text 24','파이썬 재현 비교']]){
  const s=edit(5,name,text,{fontSize:18});s.position={...s.position,left:s.position.left+18,width:217,height:32};
}

for(let i=0;i<5;i++){
  p.slides.getItem(i).speakerNotes.text=content.scripts[i]+'\n\n[발표하지 않는 참고 메모]\n'+content.references[i];
}
await fs.mkdir(path.join(task,'build','draft_render'),{recursive:true});
for(let i=0;i<5;i++){
  const png=await p.export({slide:p.slides.getItem(i),format:'png',scale:1.5});
  await fs.writeFile(path.join(task,'build','draft_render',`slide-${i+1}.png`),new Uint8Array(await png.arrayBuffer()));
}
const candidate=path.join(task,'build','candidate-r2.pptx');
await (await PresentationFile.exportPptx(p)).save(candidate);
await fs.mkdir(path.join(task,'final'),{recursive:true});
const finalPath=path.join(task,'final','고령자_디지털소외_발표최종본.pptx');
const result=await finalizePresentation({
  workspaceDir:task,candidatePath:candidate,finalPath,pythonExecutable:py,
  integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','9144000,5143500','--validate-bullet-geometry','--validate-heading-fit'],
  explicitTotalSlideCount:5,requiredNativeChartOwnerSlides:[2,3],requiredNativeTableOwnerSlides:[],
  materializeLiteralChartWorkbooks:true,
  fontPolicy:{basis:'design',families:['Malgun Gothic']},verifyArtifactToolImport:true,
  receiptPath:path.join(task,'build','validation-r2.json')
});
console.log(JSON.stringify(result));
