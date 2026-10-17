const fs=require('fs'),vm=require('vm'),assert=require('assert'),XLSX=require('../docs/xlsx.mini.min.js');
const c={XLSX,TextEncoder,TextDecoder,ArrayBuffer,Uint8Array,minutes:s=>{const [h,m]=s.split(':').map(Number);return h*60+m},hours:e=>{const t=s=>{const [h,m]=s.split(':').map(Number);return h*60+m};return (t(e.end)-t(e.start)-t(e.breakEnd)+t(e.breakStart))/60},esc:s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))};
vm.createContext(c);vm.runInContext(fs.readFileSync(require('path').join(__dirname,'../docs/excel.js'),'utf8'),c);
const project={store:'Test',city:'Orval',prior:{allocs:[{family:'Ancienne famille',type:'Devis',h:70}]},entries:[
{date:'2026-08-26',day:'J2',person:'Martine',start:'08:00',breakStart:'12:15',breakEnd:'13:15',end:'18:30',allocs:[{family:'Outils plombier',type:'Devis',h:9.5}],comment:'Rangement et nettoyage.'},
{date:'2026-08-25',day:'J1',person:'Ali',start:'09:00',breakStart:'12:45',breakEnd:'13:30',end:'18:30',allocs:[{family:'Quincaillerie',type:'Imagine',h:8.75}]},
{date:'2026-08-25',day:'J1',person:'Olivier',start:'09:00',breakStart:'09:00',breakEnd:'09:00',end:'11:30',allocs:[{family:'Quincaillerie',type:'Imagine',h:1},{family:'Gants',type:'Devis',h:1.5}],comment:'Deux familles, une plage.'}
]};
const built=c.buildPersonalReportWorkbook(project),bytes=c.styledWorkbookBytes(built),wb=XLSX.read(bytes,{type:'array',cellStyles:true}),ws=wb.Sheets.Horaires;
assert.deepEqual(wb.SheetNames,['Horaires']);assert.deepEqual(XLSX.utils.sheet_to_json(ws,{header:1})[0],['Date','Jour','Intervenant','Type intervention','Famille','Heure début','Début pause','Fin pause','Heure fin','Temps total (h)','Commentaire']);
assert.equal(ws.C2.v,'Ali');assert.equal(ws.J2.w,'8:45');assert.equal(ws.J3.w,'2:30');assert.equal(ws.J4.w,'9:30');assert.equal(ws.K4.v,'Rangement et nettoyage.');assert.equal(ws.G3.v,'');assert.equal(ws.H3.v,'');assert.equal(ws.D3.v,'Imagine (1:00) + Devis (1:30)');assert(ws.E3.v.includes('Gants — Devis 1:30'));
assert.equal(ws.J6.w,'11:00');assert.equal(ws.J7.w,'9:45');assert.equal(ws.J8.w,'20:45');assert.equal(ws.C2.s.fgColor.rgb,'41B1DB');assert.equal(ws.C4.s.fgColor.rgb,'E49EDD');assert.equal(ws.A2.s.fgColor.rgb,'EAD1E8');assert.equal(ws.A1.s.fgColor.rgb,'DAEAF5');assert(!XLSX.utils.sheet_to_csv(ws).includes('Ancienne famille'));
const preview=c.summaryPreviewHTML(built,'Horaires');assert(preview.includes('Temps total (h)'));assert(preview.includes('8:45'));assert(preview.includes('reportImagine'));assert(preview.includes('Rangement et nettoyage.'));
const empty=c.buildPersonalReportWorkbook({...project,entries:[]});assert.equal(empty.Sheets.Horaires.J5.v,0);
console.log('PASS: requested eleven-column reporting, chronological slots, time formats, Devis/Imagine colors, comments, no imported schedules or duplicate mixed-slot hours, preview matches export, empty report');
