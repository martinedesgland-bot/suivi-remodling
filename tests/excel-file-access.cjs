const fs=require('fs'),vm=require('vm'),assert=require('assert'),path=require('path'),XLSX=require('../docs/xlsx.mini.min.js');
const nodes=new Map();const node=id=>{if(!nodes.has(id))nodes.set(id,{value:'',innerHTML:'',textContent:'',classList:{add(){},remove(){},toggle(){}},addEventListener(){},scrollIntoView(){}});return nodes.get(id)};
const c={XLSX,TextEncoder,TextDecoder,ArrayBuffer,Uint8Array,document:{getElementById:node,body:{classList:{add(){},remove(){}}}},localStorage:{getItem(){},setItem(){}},setTimeout(){},alert:m=>{c.lastAlert=m}};
vm.createContext(c);vm.runInContext(fs.readFileSync(path.join(__dirname,'../docs/excel.js'),'utf8'),c);
vm.runInContext(fs.readFileSync(path.join(__dirname,'../docs/index.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1],c);
const project={id:'tablet',store:'Test',city:'Orval',planned:{Devis:7,Imagine:0},people:['Martine'],families:[{name:'Visserie',types:['Devis']}],entries:[],archived:false};
const bytes=c.styledWorkbookBytes(c.buildSummaryWorkbook(project));
const reset=()=>vm.runInContext('db.projects=[];active=null;category="ongoing";',c);
const projects=()=>vm.runInContext('db.projects',c);
c.confirmExcelImport=async()=>true;
(async()=>{
 reset();let release;const reading=new Promise(resolve=>release=resolve),input={value:'selected',files:[]};
 input.files=[{size:bytes.length,arrayBuffer:async()=>{await reading;if(!input.value)throw Object.assign(Error('Lost provider access'),{name:'NotFoundError'});return bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength)}}];
 const importing=c.importExcel({target:input});assert.equal(input.value,'selected');assert.equal(projects().length,0);release();await importing;
 assert.equal(projects().length,1,c.lastAlert);assert.equal(projects()[0].store,'Test');assert.equal(input.value,'');
 reset();c.confirmExcelImport=async()=>false;input.value='selected';await c.importExcel({target:input});assert.equal(projects().length,0);assert.equal(input.value,'');
 c.confirmExcelImport=async()=>true;
 for(const name of ['NotFoundError','NotReadableError','SecurityError','NotAllowedError']){input.value='selected';input.files=[{size:1,arrayBuffer:async()=>{throw Object.assign(Error('Provider failure'),{name})}}];await c.importExcel({target:input});assert.equal(projects().length,0);assert.equal(input.value,'');assert(c.lastAlert.includes('introuvable ou inaccessible'));}
 input.value='selected';input.files=[{size:11*1024*1024,arrayBuffer:async()=>{throw Error('Must not read oversized file')}}];await c.importExcel({target:input});assert(c.lastAlert.includes('10 Mo'));assert.equal(input.value,'');assert.equal(projects().length,0);
 console.log('PASS: selection retained during asynchronous provider read, reset after success/cancellation/failure, inaccessible files leave data unchanged, size validation preserved');
})().catch(e=>{console.error(e);process.exitCode=1});
