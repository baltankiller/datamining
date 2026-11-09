import fs from 'node:fs/promises';
import {FileBlob, PresentationFile} from '@oai/artifact-tool';
const p = await PresentationFile.importPptx(await FileBlob.load('C:/Users/UserPc/Downloads/고령자 디지털 소외 위험 예측 중간보고 발표자료.pptx'));
const inspect = await p.inspect({kind:'slide,textbox,shape,image,notes',maxChars:1000000});
await fs.writeFile(new URL('./source_inspect.ndjson',import.meta.url),inspect.ndjson);
const r=JSON.parse(inspect.ndjson.split('\n').find(x=>JSON.parse(x).kind==='textbox'));
const shape=p.resolve(r.id);
console.log('TEXT_TYPE',typeof shape.text,'TEXT',String(shape.text));
console.log(JSON.stringify(await p.help('shape.text.style',{include:['examples','notes'],maxChars:7000})));
