import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const dir = path.dirname(fileURLToPath(import.meta.url));
const source = JSON.parse(await fs.readFile(path.join(dir, 'source.json'), 'utf8'));
const out = path.resolve(dir, '../outputs/01a1aeee-21cb-7c42-86fa-892bb8704286');
await fs.mkdir(out, { recursive: true });
const rows = [];
function add(id, part, maker, major, detail, status, note, url, extra = '') {
  if (!source.rows[id]) throw new Error(`Unexpected ID: ${id}`);
  rows.push({ id, original: source.rows[id].values[1], part, maker, major, detail, status, note, url, extra, originalRow: source.rows[id].sourceRow });
}
const siN = 'N채널 전력 MOSFET';
const siP = 'P채널 전력 MOSFET';
const ti = p => `https://www.ti.com/lit/ds/symlink/${p.toLowerCase()}.pdf`;
const on = p => `https://www.onsemi.com/download/data-sheet/pdf/${p.toLowerCase()}-d.pdf`;
const toshiba = p => `https://toshiba.semicon-storage.com/us/semiconductor/product/mosfets/detail.${p}.html`;
const navitas = p => `https://navitassemi.com/wp-content/plugins/gb-navitas-stock-checker/product_files/${p}_Datasheet.pdf`;
add('01','ATSAMD21G18A','Microchip (구 Atmel)','IC','MCU (Cortex-M0+)','확인','32비트 마이크로컨트롤러.','https://www.microchip.com/content/dam/mchp/documents/OTH/ProductDocuments/DataSheets/40001882A.pdf');
add('02','ESP32-S3','Espressif Systems','IC','무선 MCU / SoC','확인','원문 ESP32S3 표기를 ESP32-S3로 정규화. Wi-Fi·Bluetooth LE 내장.','https://www.espressif.com/en/products/socs/esp32-s3');
add('13','CSD18511Q5A','Texas Instruments','MOSFET',siN,'확인','NexFET 실리콘 전력 MOSFET.',ti('CSD18511Q5A'));
add('14','CSD17308Q3','Texas Instruments','MOSFET',siN,'확인','NexFET 실리콘 전력 MOSFET.',ti('CSD17308Q3'));
add('15','CSD18543Q3A','Texas Instruments','MOSFET',siN,'확인','NexFET 실리콘 전력 MOSFET.',ti('CSD18543Q3A'));
add('16','DI100N10PQ-AQ','Diotec Semiconductor','MOSFET',siN,'확인','실리콘 전력 MOSFET.','https://diotec.com/en/product/DI100N10PQ-AQ.html');
add('17','DI070P04PQ-AQ','Diotec Semiconductor','MOSFET',siP,'확인','실리콘 전력 MOSFET.','https://diotec.com/files/diotec/files/pdf/datasheets/di070p04pq.pdf');
add('19','FDMS0308AS','onsemi (구 Fairchild)','MOSFET',siN,'기본품번 확인','유통사 CT-ND 코드를 제외. Schottky 다이오드가 통합된 SyncFET.',on('FDMS0308AS'));
add('20','FDMS4435BZ','onsemi (구 Fairchild)','MOSFET',siP,'확인','PowerTrench 실리콘 전력 MOSFET.',on('FDMS4435BZ'));
add('21','NV6015C-RA (추정)','Navitas Semiconductor','GaN FET','N채널 e-mode GaN FET','오탈자 추정','FNV6015C-RA의 F는 오탈자 추정. 원본 시트2에도 NV6015C-RA가 기재됨. 실물 품번 확인 필요.',navitas('NV6015C'));
add('22','TP65H150BG4JSG','Renesas (구 Transphorm)','GaN FET','N채널 캐스코드 GaN FET','기본품번 확인','유통사 접두사 1707-와 포장 접미사 -TR 제외. GaN HEMT + Si MOSFET 구조.','https://www.renesas.com/en/document/dst/tp65h150bg4jsg-datasheet');
add('23','GANE1R8-100QBA','Nexperia','GaN FET','N채널 e-mode GaN FET','기본품번 확인','말미 Z는 주문용 접미사. 제조사 기본품번으로 확인.','https://assets.nexperia.com/documents/data-sheet/GANE1R8-100QBA.pdf');
add('25','IRFH5015','Infineon (구 IR)','MOSFET',siN,'기본품번 확인','TRPBF 접미사를 제외한 제조사 기본품번으로 확인.','https://www.infineon.com/part/IRFH5015');
add('26','FCMT125N65S3','onsemi','MOSFET','N채널 슈퍼정션 MOSFET','유통코드 정규화','원문 0SCT-ND는 DigiKey OSCT-ND의 O/0 오기로 판단. 기본품번 FCMT125N65S3 확인.',on('FCMT125N65S3'),'https://www.digikey.no/en/products/detail/onsemi/FCMT125N65S3/9764705');
add('27','FCMT360N65S3','onsemi','MOSFET','N채널 슈퍼정션 MOSFET','유통코드 정규화','원문 0SCT-ND는 DigiKey OSCT-ND의 O/0 오기로 판단. 기본품번 FCMT360N65S3 확인.',on('FCMT360N65S3'),'https://www.digikey.in/en/products/detail/onsemi/FCMT360N65S3/10273402');
add('28','DI025N20PQ (추정)','Diotec Semiconductor','MOSFET',siN,'오탈자 추정','DIO25N20PQ의 O는 숫자 0의 오탈자 추정. DI025N20PQ 기준 분류. 실물 품번 확인 필요.','https://diotec.com/tl_files/diotec/files/pdf/datasheets/di025n20pq.pdf');
add('29','NTMFS5C604NL','onsemi','MOSFET',siN,'기본품번 확인','T1G 포장·주문 접미사 제외.',on('NTMFS5C604NL'));
add('30','IRFH5300','Infineon (구 IR)','MOSFET',siN,'기본품번 확인','TRPBF 접미사를 제외한 제조사 기본품번으로 확인.','https://www.infineon.com/part/IRFH5300');
add('31','FDMS86263P','onsemi (구 Fairchild)','MOSFET',siP,'확인','PowerTrench 실리콘 전력 MOSFET.',on('FDMS86263P'));
add('33','NV6015C-RA','Navitas Semiconductor','GaN FET','N채널 e-mode GaN FET','확인','제조사는 Navitas. 원본 시트2 제조사 칸의 Renesas 표기는 맞지 않음.',navitas('NV6015C'));
add('36','TPH5200FNH','Toshiba','MOSFET',siN,'확인','실리콘 전력 MOSFET.','https://toshiba.semicon-storage.com/eu/semiconductor/product/mosfets/12v-300v-mosfets/detail.TPH5200FNH.html');
add('37','NV6029','Navitas Semiconductor','GaN FET','N채널 e-mode GaN FET','확인','GaNFast 전력 FET.',navitas('NV6029'));
add('38','TSM056NH04LCR','Taiwan Semiconductor','MOSFET',siN,'기본품번 확인','RLG 포장코드와 CT-ND 유통사 코드를 제외.','https://www.taiwansemi.com/en/perfet/','https://services.taiwansemi.com/storage/resources/mtbf-mttf-report-9/AMF20230111_MTTF%20%26%20FIT%20Report_TSM056NH04LCR.pdf');
add('39','CDF56G6511N','Central Semiconductor','GaN FET','N채널 GaN 전력 FET','기본품번 확인','TR13 PBFREE 주문·포장 표기를 제외.','https://my.centralsemi.com/datasheets/CDF56G6511N.PDF');
add('40','AON6266E','Alpha & Omega (AOS)','MOSFET',siN,'확인','AlphaSGT 실리콘 전력 MOSFET.','https://www.aosmd.com/products/mosfets/medium-voltage-mosfets-40v-400v/aon6266e');
add('41','NP20P06YLG','Renesas','MOSFET',siP,'기본품번 확인','-E1-AY 주문·포장 접미사 제외.','https://www.renesas.com/en/document/dst/np20p06ylg-data-sheet');
add('42','NVMTS4D3N15MC','onsemi','MOSFET',siN,'확인','실리콘 전력 MOSFET.',on('NVMTS4D3N15MC'));
add('43','FDMS5672','onsemi (구 Fairchild)','MOSFET',siN,'확인','UltraFET Trench 실리콘 전력 MOSFET.',on('FDMS5672'));
add('44','TPHR9003NL1','Toshiba','MOSFET',siN,'기본품번 확인',',LQ 주문 접미사 제외.','https://toshiba.semicon-storage.com/eu/semiconductor/product/mosfets/12v-300v-mosfets/detail.TPHR9003NL1.html');
add('45','TSM280NB06LCR','Taiwan Semiconductor','MOSFET',siN,'기본품번 확인','RLG 포장코드 제외.','https://services.ts.com.tw/storage/resources/datasheet/TSM280NB06LCR_B1804.pdf');
add('46','TPH1R712MD','Toshiba','MOSFET',siP,'기본품번 확인','원문 ,L1Q 접미사는 그대로 보존. 분류는 확인된 기본품번 기준.',toshiba('TPH1R712MD'));
add('47','TPH3R704PL','Toshiba','MOSFET',siN,'기본품번 확인','원문 ,L1Q 접미사는 그대로 보존. 분류는 확인된 기본품번 기준.','https://toshiba.semicon-storage.com/eu/semiconductor/product/mosfets/12v-300v-mosfets/detail.TPH3R704PL.html');
add('48','TP65H480G4JSG','Renesas (구 Transphorm)','GaN FET','N채널 캐스코드 GaN FET','기본품번 확인','-TR 포장 접미사 제외. GaN HEMT + Si MOSFET 구조.','https://www.renesas.com/en/document/dst/tp65h480g4jsg-datasheet');
add('49','TP65H300G4LSGB','Renesas (구 Transphorm)','GaN FET','N채널 캐스코드 GaN FET','기본품번 확인','-TR 포장 접미사 제외. GaN HEMT + Si MOSFET 구조.','https://www.renesas.com/en/products/tp65h300g4lsgb');
add('50','TP65H480G4JSGB','Renesas (구 Transphorm)','GaN FET','N채널 캐스코드 GaN FET','기본품번 확인','-TR 포장 접미사 제외. 48번과는 B 유무가 다른 별도 품번.','https://www.renesas.com/en/products/tp65h480g4jsgb');
add('51','TSM170N06PQ56','Taiwan Semiconductor','MOSFET',siN,'기본품번 확인','RLG 포장코드 제외.','https://services.taiwansemi.com/storage/resources/datasheet/TSM170N06PQ56_A1608.pdf');
add('52','NTMFS0D5N04XM','onsemi','MOSFET',siN,'기본품번 확인','T1G 포장·주문 접미사 제외.',on('NTMFS0D5N04XM'));
add('53','XPH4R714MC','Toshiba','MOSFET',siP,'기본품번 확인',',L1XHQ 주문 접미사를 제외한 기본품번으로 확인.',toshiba('XPH4R714MC'));
add('54','FDMS4D4N08C','onsemi','MOSFET',siN,'확인','Shielded Gate PowerTrench 실리콘 전력 MOSFET.',on('FDMS4D4N08C'));
add('55','NVMFS3D0P04M8L','onsemi','MOSFET',siP,'기본품번 확인','T1G 포장·주문 접미사 제외.',on('NVMFS3D0P04M8L'));
if (rows.length !== 40 || rows.map(r => r.id).join() !== source.requested.join()) throw new Error('ID coverage mismatch');
await fs.writeFile(path.join(dir, 'classification.json'), JSON.stringify(rows, null, 2));

const wb = Workbook.create();
const main = wb.worksheets.add('칩셋 분류');
const evidence = wb.worksheets.add('원문 및 근거');
const font = 'Arial';
const dark = '#24364B';
function base(s, range) {
  s.showGridLines = false;
  s.getRange(range).format.font = { name: font, size: 11, color: '#172033' };
  s.getRange(range).format.rowHeight = 27;
  s.getRange(range).format.verticalAlignment = 'center';
}
function header(s, range) {
  s.getRange(range).format = { fill: dark, font: { name: font, size: 11, bold: true, color: '#FFFFFF' }, horizontalAlignment: 'center', verticalAlignment: 'center', rowHeight: 30 };
}
base(main, 'A1:G51');
base(evidence, 'A1:L45');
main.tabColor = dark;
main.getRange('A2').values = [['사용 칩셋 분류']];
main.getRange('A2').format.font = { name: font, size: 16, bold: true, color: dark };
main.getRange('A3').values = [['검색일 2026-11-06. 수치는 요청한 40개 번호 기준이며, 21·28번은 오탈자 추정에 따른 분류입니다.']];
main.getRange('A3:G3').format.borders = { bottom: { style: 'thin', color: '#D5DDE5' } };
main.getRange('A3').format.font = { name: font, size: 11, color: '#4B5563' };
main.getRange('A5:B8').values = [['대분류','번호 수'],['IC',null],['MOSFET',null],['GaN FET',null]];
main.getRange('B6').formulas = [['=COUNTIFS($C$12:$C$51,A6)']];
main.getRange('B6:B8').fillDown();
main.getRange('D5:E8').values = [['세부 구분','번호 수'],['N채널 MOSFET',null],['P채널 MOSFET',null],['전체',null]];
main.getRange('E6').formulas = [['=COUNTIFS($D$12:$D$51,"N채널 전력 MOSFET")+COUNTIFS($D$12:$D$51,"N채널 슈퍼정션 MOSFET")']];
main.getRange('E7').formulas = [['=COUNTIFS($D$12:$D$51,"P채널 전력 MOSFET")']];
main.getRange('E8').formulas = [['=COUNTA($A$12:$A$51)']];
header(main, 'A5:B5');
header(main, 'D5:E5');
main.getRange('B6:B8').setNumberFormat('0');
main.getRange('E6:E8').setNumberFormat('0');
main.getRange('G5').values = [['분류 기준']];
main.getRange('G5').format.font = { name: font, size: 11, bold: true };
main.getRange('G6').values = [['MOSFET: 실리콘 소자. GaN FET: 질화갈륨 전력 FET.']];
main.getRange('G7').values = [['QFN·DFN·QFP는 패키지 이름입니다.']];
main.getRange('G8').values = [['원문·품번 보정·검색 링크는 다음 시트에서 확인합니다.']];
main.getRange('A10').values = [['01·02번은 IC입니다. 나머지는 전력용 FET이며, 개별 다이오드·BJT로 분류된 번호는 없습니다.']];
main.getRange('A11:G11').values = [['번호','확인한 기본품번','대분류','세부 분류','제조사','확인 상태','비고']];
header(main, 'A11:G11');

evidence.getRange('A2').values = [['원문 매칭과 검색 근거']];
evidence.getRange('A2').format.font = { name: font, size: 16, bold: true, color: dark };
evidence.getRange('A3').values = [[`원본: ${source.filename} / ${source.sheet}의 번호·칩셋명. 21번은 원본 시트2도 대조했습니다.`]];
evidence.getRange('A3').format.font = { name: font, size: 11, color: '#4B5563' };
evidence.getRange('A5:H5').values = [['번호','엑셀 원문 칩셋명','확인한 기본품번','제조사','대분류','세부 분류','확인 상태','매칭 및 분류 메모']];
evidence.getRange('J5:L5').values = [['근거 자료','제조사 자료 URL','매칭 보조 자료 URL']];
header(evidence, 'A5:H5');
header(evidence, 'J5:L5');
evidence.getRange('A6:H45').values = rows.map(r => [r.id,r.original,r.part,r.maker,r.major,r.detail,r.status,r.note]);
evidence.getRange('J6:L45').values = rows.map(r => [r.id === '38' ? 'Taiwan Semiconductor 제품군·신뢰성 자료' : `${r.maker} ${r.part.replace(' (추정)','')} 제조사 자료`,r.url,r.extra || null]);
evidence.getRange('A6:A45').setNumberFormat('00');
evidence.getRange('B6:C45').setNumberFormat('@');
evidence.getRange('A6:L45').format.rowHeight = 66;
evidence.getRange('B6:C45').format.wrapText = true;
evidence.getRange('H6:H45').format.wrapText = true;
evidence.getRange('J6:L45').format.wrapText = true;
evidence.getRange('K6:L45').format.font = { name: font, size: 11, color: '#2457A6' };
const mainNotes = new Map([
  ['19','Schottky 통합형 MOSFET'],
  ['21','FNV → NV로 추정. 실물 품번 확인 필요.'],
  ['26','DigiKey 주문코드의 O/0 표기 보정'],
  ['27','DigiKey 주문코드의 O/0 표기 보정'],
  ['28','DIO25 → DI025로 추정. 실물 품번 확인 필요.'],
  ['33','21번도 이 품번일 가능성. 제조사는 Navitas.'],
  ['46','P채널 MOSFET'],
  ['48','GaN HEMT + 저전압 Si MOSFET'],
  ['49','GaN HEMT + 저전압 Si MOSFET'],
  ['50','48번과 B 접미사가 다른 별도 품번'],
  ['53','P채널 MOSFET']
]);
for (let i=0; i<rows.length; i++) {
  const r = 12+i, er = 6+i;
  main.getRange(`A${r}:F${r}`).formulas = [[`='원문 및 근거'!A${er}`,`='원문 및 근거'!C${er}`,`='원문 및 근거'!E${er}`,`='원문 및 근거'!F${er}`,`='원문 및 근거'!D${er}`,`='원문 및 근거'!G${er}`]];
  main.getRange(`G${r}`).values = [[mainNotes.get(rows[i].id) || '']];
  if (i%2===1) {
    main.getRange(`A${r}:G${r}`).format.fill = '#F1F5F9';
    evidence.getRange(`A${er}:H${er}`).format.fill = '#F1F5F9';
    evidence.getRange(`J${er}:L${er}`).format.fill = '#F1F5F9';
  }
}
main.getRange('A12:A51').setNumberFormat('00');
main.getRange('A12:A51').format.horizontalAlignment = 'center';
main.getRange('G12:G51').format.wrapText = true;
main.getRange('A12:G51').format.rowHeight = 39;
main.getRange('F12:F51').conditionalFormats.add('containsText',{text:'오탈자 추정',format:{fill:'#FFF1C2',font:{bold:true,color:'#8A4B05'}}});
evidence.getRange('G6:G45').conditionalFormats.add('containsText',{text:'오탈자 추정',format:{fill:'#FFF1C2',font:{bold:true,color:'#8A4B05'}}});
const widthsMain = { A:12, B:30, C:13, D:28, E:29, F:20, G:56 };
const widthsEvidence = { A:7, B:42, C:30, D:29, E:13, F:28, G:20, H:65, I:3, J:47, K:100, L:100 };
for (const [col,width] of Object.entries(widthsMain)) main.getRange(`${col}1:${col}51`).format.columnWidth = width;
for (const [col,width] of Object.entries(widthsEvidence)) evidence.getRange(`${col}1:${col}45`).format.columnWidth = width;
main.freezePanes.freezeRows(11);
main.freezePanes.freezeColumns(2);
evidence.freezePanes.freezeRows(5);
evidence.freezePanes.freezeColumns(2);
const table = main.tables.add('A11:G51',true,'ChipClassification');
table.showFilterButton = true;
const evidenceTable = evidence.tables.add('A5:H45',true,'ChipMatchingEvidence');
evidenceTable.showFilterButton = true;
header(main, 'A11:G11');
header(evidence, 'A5:H5');
wb.recalculate();
const counts = Object.fromEntries(['IC','MOSFET','GaN FET'].map(t=>[t,rows.filter(r=>r.major===t).length]));
if (counts.IC!==2 || counts.MOSFET!==29 || counts['GaN FET']!==9) throw new Error('Classification totals mismatch');
const summaryValues = main.getRange('D6:E8').values;
if (summaryValues[0][1] !== 22 || summaryValues[1][1] !== 7 || summaryValues[2][1] !== 40) throw new Error(`Summary mismatch: ${JSON.stringify(summaryValues)}`);
console.log(JSON.stringify({rows:rows.length,counts,uncertain:rows.filter(r=>r.status==='오탈자 추정').map(r=>r.id)}));
console.log((await wb.inspect({kind:'table',range:'칩셋 분류!A5:F8',include:'values,formulas',tableMaxRows:4,tableMaxCols:6,maxChars:2500})).ndjson);
console.log((await wb.inspect({kind:'table',range:'칩셋 분류!A19:F22',include:'values',tableMaxRows:4,tableMaxCols:6,maxChars:1500})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},maxChars:1500})).ndjson);
for (const [sheetName,range,name] of [
  ['칩셋 분류','A1:G22','classification_top'],
  ['칩셋 분류','A37:G51','classification_bottom'],
  ['원문 및 근거','A1:H13','evidence_left'],
  ['원문 및 근거','J5:L12','evidence_sources']
]) {
  const image = await wb.render({sheetName,range,scale:1.2,format:'png'});
  await fs.writeFile(path.join(dir,`${name}.png`),new Uint8Array(await image.arrayBuffer()));
}
const result = await SpreadsheetFile.exportXlsx(wb);
const outputFile = path.join(out,'사용_칩셋_40개_분류결과.xlsx');
await result.save(outputFile);
console.log(JSON.stringify({outputFile}));
