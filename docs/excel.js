/* Excel import/export for Suivi Remodeling. SheetJS CE 0.20.3 (Apache-2.0). */
function chooseExcel(){document.getElementById('excelFile').click()}
const excelKey=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
function excelNumber(v,label){if(v==null||String(v).trim()==='')return 0;const n=Number(String(v).replace(/\s/g,'').replace(',','.'));if(!Number.isFinite(n)||n<0)throw Error('Nombre invalide : '+label);return n}
function parseExcel(wb){
 const sheets=wb.SheetNames.map(name=>({name,rows:XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''})}));
 const data={people:[],families:[],entries:[],planned:{Devis:0,Imagine:0}};let found=false;
 for(const {rows} of sheets){const hi=rows.findIndex(r=>r.some(v=>['famille','familles'].includes(excelKey(v)))&&r.some(v=>['ujtdevis','ujtimagine'].includes(excelKey(v))));if(hi<0)continue;found=true;
 const keys=rows[hi].map(excelKey),fi=keys.findIndex(k=>['famille','familles'].includes(k)),pi=keys.findIndex(k=>['intervenant','intervenants'].includes(k)),di=keys.indexOf('ujtdevis'),ii=keys.indexOf('ujtimagine'),ti=keys.indexOf('types');
 for(const [offset,r] of rows.slice(hi+1).entries()){const person=String(r[pi]??'').trim();if(person&&!data.people.includes(person))data.people.push(person);const name=String(r[fi]??'').trim();if(!name){if((di>=0&&r[di]!=='')||(ii>=0&&r[ii]!==''))throw Error('Famille manquante ligne '+(hi+offset+2));continue}const d=excelNumber(r[di],'Devis / '+name),i=excelNumber(r[ii],'Imagine / '+name);if(data.families.some(f=>excelKey(f.name)===excelKey(name)))throw Error('Famille en double : '+name);data.families.push({name,types:ti>=0&&String(r[ti]).split(',').every(t=>['Devis','Imagine'].includes(t))?String(r[ti]).split(','):[...(d?['Devis']:[]),...(i?['Imagine']:[])].length?[...(d?['Devis']:[]),...(i?['Imagine']:[])]:['Devis','Imagine'],planned:String(r[di]??'').trim()!==''||String(r[ii]??'').trim()!==''?{Devis:d*7,Imagine:i*7}:null});data.planned.Devis+=d*7;data.planned.Imagine+=i*7;}
 break;}
 if(!found||!data.families.length)throw Error('Colonnes attendues : Intervenant, Familles, UJT devis, UJT Imagine.');
 const meta=sheets.find(s=>s.name==='Chantier');if(meta){const m=Object.fromEntries(meta.rows.map(r=>[excelKey(r[0]),r[1]]));data.store=String(m.magasin||'');data.city=String(m.ville||'');data.sourceId=String(m.identifiant||'');if(m.devisheures!==undefined)data.planned.Devis=excelNumber(m.devisheures,'budget Devis');if(m.imagineheures!==undefined)data.planned.Imagine=excelNumber(m.imagineheures,'budget Imagine');}
 const history=sheets.find(s=>s.name==='Heures');if(history){const k=history.rows[0].map(excelKey),get=(r,key)=>r[k.indexOf(key)];for(const r of history.rows.slice(1)){if(r.every(v=>v===''))continue;const person=String(get(r,'intervenant')||'').trim(),family=String(get(r,'famille')||'').trim(),type=get(r,'type'),date=String(get(r,'date')||'');const h=excelNumber(get(r,'heures'),'heures réalisées');if(!person||!family||!['Devis','Imagine'].includes(type)||!/^\d{4}-\d{2}-\d{2}$/.test(date)||!h)throw Error('Ligne d’historique invalide.');const e={id:String(get(r,'identifiant')||uid()),person,date,day:String(get(r,'journee')||''),start:String(get(r,'debut')),breakStart:String(get(r,'debutpause')),breakEnd:String(get(r,'finpause')),end:String(get(r,'fin')),allocs:[{family,type,h}]};const prior=data.entries.find(v=>v.id===e.id);if(prior){if(prior.date!==e.date||prior.person!==e.person||prior.start!==e.start||prior.end!==e.end)throw Error('Identifiant de saisie incohérent.');prior.allocs.push(e.allocs[0]);}else data.entries.push(e);if(!data.people.includes(person))data.people.push(person);if(!data.families.some(f=>f.name===family))data.families.push({name:family,types:[type]});}}
 for(const e of data.entries)if(!Number.isFinite(hours(e))||Math.abs(hours(e)-e.allocs.reduce((n,a)=>n+a.h,0))>.001)throw Error('Les horaires et heures réalisées ne correspondent pas.');
 return data;
}
async function importExcel(ev){const file=ev.target.files[0],target=active,cat=category;ev.target.value='';if(!file)return;try{if(file.size>10*1024*1024)throw Error('Fichier trop volumineux (10 Mo maximum).');const data=parseExcel(XLSX.read(await file.arrayBuffer(),{type:'array'}));let x=target?db.projects.find(p=>p.id===target):null;if(x?.archived)throw Error('Ce chantier est archivé.');if(!x&&data.sourceId)x=db.projects.find(p=>p.id===data.sourceId&&!p.archived);if(x&&data.store&&(excelKey(data.store)!==excelKey(x.store)||excelKey(data.city)!==excelKey(x.city)))throw Error('Ce fichier correspond à un autre magasin ou une autre ville.');let store=x?.store||data.store||$('store').value.trim(),city=x?.city||data.city||$('city').value.trim();if(!store||!city)throw Error('Renseigne d’abord le magasin et la ville dans le formulaire de cette rubrique, puis importe le fichier.');
 const msg='Importer dans '+store+' — '+city+' ?\n'+data.people.length+' intervenant(s), '+data.families.length+' famille(s)\nDevis : '+fmt(data.planned.Devis)+' h · Imagine : '+fmt(data.planned.Imagine)+' h\n'+data.entries.length+' plage(s) réalisée(s).\nLes budgets des familles présentes seront mis à jour. Les saisies existantes seront conservées.';if(!confirm(msg))return;
 const next=x?JSON.parse(JSON.stringify(x)):{id:data.sourceId||uid(),store,city,people:[],families:[],entries:[],archived:false,started:cat==='ongoing'};
 next.people=[...new Set([...next.people,...data.people])];for(const f of data.families){const idx=next.families.findIndex(v=>excelKey(familyName(v))===excelKey(f.name));if(idx<0)next.families.push(f);else next.families[idx]=f;}
 next.planned=data.planned;const signature=e=>JSON.stringify([e.date,e.person,e.start,e.breakStart,e.breakEnd,e.end,e.allocs]);for(const e of data.entries)if(!next.entries.some(v=>signature(v)===signature(e)))next.entries.push(e);next.started=next.started||!!next.entries.length;
 const previous=db.projects;if(x)db.projects=db.projects.map(v=>v.id===x.id?next:v);else db.projects=[...db.projects,next];try{save()}catch(e){db.projects=previous;throw Error('Enregistrement impossible : espace de stockage insuffisant.');}openProject(next.id);note('Fichier Excel importé.');}catch(e){alert('Import impossible : '+e.message)}}
// One row per recorded time slot: a slot shared by families is never counted twice.
function buildReportWorkbook(x){
 const wb=XLSX.utils.book_new(),duration='[h]" h "mm',clock='hh:mm';
 const amount=h=>Math.round(Number(h)*60)/1440;
 const typed=(v,z)=>({t:'n',v,z});
 const dur=h=>typed(amount(h),duration);
 const time=s=>{const [h,m]=String(s).split(':').map(Number);return typed((h*60+m)/1440,clock)};
 const date=s=>typed((Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000,'dd/mm/yyyy');
 const add=(name,rows,widths,headerRow)=>{const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=widths.map(wch=>({wch}));ws['!rows']=rows.map((_,r)=>({hpt:r===0?30:23}));ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:widths.length-1}},{s:{r:1,c:0},e:{r:1,c:widths.length-1}}];if(headerRow!==undefined)ws['!autofilter']={ref:XLSX.utils.encode_range({r:headerRow,c:0},{r:Math.max(headerRow,rows.length-1),c:widths.length-1})};XLSX.utils.book_append_sheet(wb,ws,name);return ws};
 const entries=[...x.entries].sort((a,b)=>a.person.localeCompare(b.person,'fr')||a.date.localeCompare(b.date)||a.start.localeCompare(b.start));
 const total={Devis:0,Imagine:0},people=new Map(),days=new Map(),families=new Map();
 for(const name of x.people)people.set(name,{Devis:0,Imagine:0});
 for(const f of x.families)families.set(familyName(f),{Devis:0,Imagine:0,budget:typeof f==='object'?f.planned:null});
 const slots=[];
 for(const e of entries){const split={Devis:0,Imagine:0};let person=people.get(e.person);if(!person){person={Devis:0,Imagine:0};people.set(e.person,person)}const key=JSON.stringify([e.person,e.date]);let day=days.get(key);if(!day){day={person:e.person,date:e.date,Devis:0,Imagine:0};days.set(key,day)}
 for(const a of e.allocs){if(!['Devis','Imagine'].includes(a.type)||!Number.isFinite(a.h)||a.h<0)throw Error('Saisie d’heures invalide.');split[a.type]+=a.h;total[a.type]+=a.h;person[a.type]+=a.h;day[a.type]+=a.h;let f=families.get(a.family);if(!f){f={Devis:0,Imagine:0,budget:null};families.set(a.family,f)}f[a.type]+=a.h;}
 const h=split.Devis+split.Imagine;if(!Number.isFinite(hours(e))||Math.abs(hours(e)-h)>0.001)throw Error('Les heures affectées ne correspondent pas aux horaires de '+e.person+' le '+e.date+'.');
 slots.push([e.person,date(e.date),time(e.start),time(e.breakStart),time(e.breakEnd),time(e.end),[...new Set(e.allocs.map(a=>a.family))].join(', '),dur(split.Devis),dur(split.Imagine),dur(h)]);
 }
 const title=x.store+' — '+x.city;
 const summary=[['SUIVI DU CHANTIER'],[title],[],['HEURES DU CHANTIER','Prévues','Réalisées','Restant à faire','Dépassement'],...['Devis','Imagine'].map(t=>[t.toUpperCase(),dur(x.planned[t]),dur(total[t]),dur(Math.max(0,x.planned[t]-total[t])),dur(Math.max(0,total[t]-x.planned[t]))]),[],['Durées en heures et minutes. 1 UJT = 7 heures.'],['Le restant à faire est le budget prévu moins les heures réalisées.'],['Un dépassement est indiqué séparément lorsque le budget est consommé.'],[],['TOTAL PAR INTERVENANT','Devis','Imagine','Total réalisé'],...[...people].map(([name,v])=>[name,dur(v.Devis),dur(v.Imagine),dur(v.Devis+v.Imagine)]),['TOTAL CHANTIER',dur(total.Devis),dur(total.Imagine),dur(total.Devis+total.Imagine)]];
 const ws=add('Synthèse',summary,[32,20,20,22,20]);ws['!merges'].push({s:{r:7,c:0},e:{r:7,c:4}},{s:{r:8,c:0},e:{r:8,c:4}},{s:{r:9,c:0},e:{r:9,c:4}});
 add('Horaires',[['HORAIRES PAR INTERVENANT'],[title],['Une ligne par plage enregistrée. Les colonnes Devis et Imagine répartissent cette même plage.'],['Intervenant','Date','Début','Début pause','Fin pause','Fin','Familles réalisées','Devis','Imagine','Total plage'],...slots],[23,14,12,16,16,12,38,16,16,16],3);
 add('Totaux journaliers',[['TOTAUX JOURNALIERS PAR INTERVENANT'],[title],['Total des plages enregistrées de chaque personne, pour chaque date.'],['Intervenant','Date','Devis','Imagine','Total jour'],...[...days.values()].map(v=>[v.person,date(v.date),dur(v.Devis),dur(v.Imagine),dur(v.Devis+v.Imagine)])],[25,15,19,19,19],3);
 const fr=[['SUIVI DES FAMILLES'],[title],['Les budgets non renseignés restent vides : ils ne sont pas considérés comme zéro.'],['Famille','Type','Prévues','Réalisées','Restant à faire','Dépassement']];
 for(const [name,v] of families)for(const t of ['Devis','Imagine']){const known=v.budget&&Number.isFinite(v.budget[t]);fr.push([name,t,known?dur(v.budget[t]):'Non renseigné',dur(v[t]),known?dur(Math.max(0,v.budget[t]-v[t])):'Non renseigné',known?dur(Math.max(0,v[t]-v.budget[t])):'Non renseigné'])}
 add('Familles',fr,[32,16,22,20,22,22],3);
 // Preserve the source fields used by the colleague's import, without mixing them into the report.
 const raw=(name,rows)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),name);
 raw('Chantier',[['Magasin',x.store],['Ville',x.city],['Identifiant',x.id],['Devis heures',x.planned.Devis],['Imagine heures',x.planned.Imagine]]);
 raw('Preparation',[['Intervenant','Familles','UJT devis','UJT Imagine','Types'],...Array.from({length:Math.max(x.people.length,x.families.length)},(_,i)=>[x.people[i]||'',familyName(x.families[i]||''),x.families[i]?.planned?x.families[i].planned.Devis/7:'',x.families[i]?.planned?x.families[i].planned.Imagine/7:'',x.families[i]?familyTypes(x.families[i]).join(','):''])]);
 raw('Heures',[['Identifiant','Date','Journée','Intervenant','Début','Début pause','Fin pause','Fin','Famille','Type','Heures'],...x.entries.flatMap(e=>e.allocs.map(a=>[e.id,e.date,e.day,e.person,e.start,e.breakStart,e.breakEnd,e.end,a.family,a.type,a.h]))]);
 wb.Workbook={Sheets:wb.SheetNames.map(name=>({name,Hidden:['Chantier','Preparation','Heures'].includes(name)?1:0}))};
 return wb;
}
function buildSummaryWorkbook(x){
 const wb=buildReportWorkbook(x),ws={},styles={},merges=[],duration='[h]:mm';
 const put=(r,c,v,tone='plain',format)=>{const ref=XLSX.utils.encode_cell({r,c});ws[ref]=typeof v==='number'?{t:'n',v}:{t:'s',v:String(v??'')};if(format)ws[ref].z=format;styles[ref]=tone};
 const span=(r,c,end,text,tone)=>{put(r,c,text,tone);merges.push({s:{r,c},e:{r,c:end}});for(let i=c+1;i<=end;i++)put(r,i,'',tone)};
 const hoursCell=(r,c,h,tone)=>put(r,c,Math.round(h*60)/1440,tone,duration);
 const people=new Map(x.people.map(n=>[n,{Devis:0,Imagine:0}])),families=new Map(x.families.map(f=>[familyName(f),{Devis:0,Imagine:0}])),days=new Map(),total={Devis:0,Imagine:0};
 for(const e of x.entries){let person=people.get(e.person);if(!person){person={Devis:0,Imagine:0};people.set(e.person,person)}let day=days.get(e.date);if(!day){day={h:0,labels:new Set()};days.set(e.date,day)}if(e.day)day.labels.add(e.day);
 for(const a of e.allocs){person[a.type]+=a.h;total[a.type]+=a.h;day.h+=a.h;let f=families.get(a.family);if(!f){f={Devis:0,Imagine:0};families.set(a.family,f)}f[a.type]+=a.h;}}
 span(0,0,11,'SYNTHÈSE DU CHANTIER','title');span(1,0,11,x.store+' — '+x.city,'subtitle');
 ['Type','Heures réalisées','UJT prévues','Heures prévues','Restant à faire','Dépassement'].forEach((v,c)=>put(3,c,v,'header'));
 ['Devis','Imagine'].forEach((t,i)=>{const r=4+i,tone=t==='Devis'?'devis':'imagine';put(r,0,'Total '+t,tone);hoursCell(r,1,total[t],tone);put(r,2,x.planned[t]/7,tone,'0.##');hoursCell(r,3,x.planned[t],tone);hoursCell(r,4,Math.max(0,x.planned[t]-total[t]),'remaining');hoursCell(r,5,Math.max(0,total[t]-x.planned[t]),total[t]>x.planned[t]?'overrun':'total')});
 span(7,0,11,'Durées au format heures:minutes. 1 UJT = 7 heures.','note');
 ['Intervenant','Devis (h)','Imagine (h)','Total (h)'].forEach((v,c)=>put(9,c,v,c===1?'devisHeader':c===2?'imagineHeader':'header'));
 ['Famille','Devis (h)','Imagine (h)','Total (h)'].forEach((v,c)=>put(9,c+5,v,c===1?'devisHeader':c===2?'imagineHeader':'header'));
 put(9,10,'Jour','header');put(9,11,'Total (h)','header');
 let r=10;for(const [name,v] of people){put(r,0,name);hoursCell(r,1,v.Devis,'devis');hoursCell(r,2,v.Imagine,'imagine');hoursCell(r,3,v.Devis+v.Imagine,'total');r++}put(r,0,'TOTAL','total');hoursCell(r,1,total.Devis,'devisHeader');hoursCell(r,2,total.Imagine,'imagineHeader');hoursCell(r,3,total.Devis+total.Imagine,'total');
 r=10;for(const [name,v] of families){put(r,5,name);hoursCell(r,6,v.Devis,'devis');hoursCell(r,7,v.Imagine,'imagine');hoursCell(r,8,v.Devis+v.Imagine,'total');r++}put(r,5,'TOTAL','total');hoursCell(r,6,total.Devis,'devisHeader');hoursCell(r,7,total.Imagine,'imagineHeader');hoursCell(r,8,total.Devis+total.Imagine,'total');
 r=10;for(const [date,v] of [...days].sort(([a],[b])=>a.localeCompare(b))){const label=v.labels.size===1?[...v.labels][0]:'Jour '+(r-9);put(r,10,label+' · '+date.split('-').reverse().join('/'));hoursCell(r,11,v.h,'total');r++}put(r,10,'TOTAL','total');hoursCell(r,11,total.Devis+total.Imagine,'total');
 const bottom=Math.max(people.size,families.size,days.size)+10;
 ws['!ref']='A1:L'+(bottom+1);ws['!merges']=merges;ws['!cols']=[25,18,17,19,20,32,16,16,16,3,28,16].map(wch=>({wch}));ws['!rows']=Array.from({length:bottom+1},(_,i)=>({hpt:i===0?30:i===3||i===9?34:23}));
 wb.Sheets['Synthèse']=ws;wb.SheetNames=wb.SheetNames.filter(n=>!['Horaires','Totaux journaliers','Familles'].includes(n));for(const n of ['Horaires','Totaux journaliers','Familles'])delete wb.Sheets[n];wb.Workbook.Sheets=wb.SheetNames.map(name=>({name,Hidden:name==='Synthèse'?0:1}));wb.__reportStyles={'Synthèse':styles};
 return wb;
}
// Add report fonts, borders and fills to the generated XLSX while preserving number formats.
function styledWorkbookBytes(wb){
 const cfb=XLSX.CFB,zip=cfb.read(new Uint8Array(XLSX.write(wb,{type:'array',bookType:'xlsx'})),{type:'buffer'}),decoder=new TextDecoder(),encoder=new TextEncoder();
 const read=path=>decoder.decode(cfb.find(zip,'/'+path).content),write=(path,text)=>cfb.utils.cfb_add(zip,'/'+path,encoder.encode(text));
 let xml=read('xl/styles.xml');
 const append=(tag,items)=>{let base=0;xml=xml.replace(new RegExp('<'+tag+' count="(\\d+)">([\\s\\S]*?)</'+tag+'>'),(_,count,body)=>{base=+count;return '<'+tag+' count="'+(base+items.length)+'">'+body+items.join('')+'</'+tag+'>'});return base};
 const fontBase=append('fonts',[['20352A',false,11],['FFFFFF',true,11],['20352A',true,11],['FFFFFF',true,18]].map(([color,bold,size])=>'<font>'+(bold?'<b/>':'')+'<sz val="'+size+'"/><color rgb="FF'+color+'"/><name val="Calibri"/></font>'));
 const tones={plain:['FFFFFF',0],title:['006341',3],subtitle:['E2EFDA',2],header:['006341',1],devis:['E8B0DF',0],imagine:['A6DCEF',0],total:['E2EFDA',2],remaining:['FFF2CC',2],overrun:['F8B4B4',2],note:['FFFFFF',0],devisHeader:['CF85C6',2],imagineHeader:['73C5E5',2]};
 const names=Object.keys(tones),fillBase=append('fills',names.map(n=>'<fill><patternFill patternType="solid"><fgColor rgb="FF'+tones[n][0]+'"/><bgColor indexed="64"/></patternFill></fill>'));
 const borderBase=append('borders',['<border><left style="thin"><color rgb="FFAAB8AE"/></left><right style="thin"><color rgb="FFAAB8AE"/></right><top style="thin"><color rgb="FFAAB8AE"/></top><bottom style="thin"><color rgb="FFAAB8AE"/></bottom><diagonal/></border>']);
 const match=xml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/),original=match[2].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g),variants=[],cache=new Map();
 const variant=(id,tone)=>{const key=id+':'+tone;if(cache.has(key))return cache.get(key);const source=original[id]||original[0],numFmt=(source.match(/numFmtId="(\d+)"/)||[])[1]||'0',index=original.length+variants.length;variants.push('<xf numFmtId="'+numFmt+'" fontId="'+(fontBase+tones[tone][1])+'" fillId="'+(fillBase+names.indexOf(tone))+'" borderId="'+borderBase+'" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>');cache.set(key,index);return index};
 for(const [sheet,map] of Object.entries(wb.__reportStyles||{})){const path='xl/worksheets/sheet'+(wb.SheetNames.indexOf(sheet)+1)+'.xml';let doc=read(path);doc=doc.replace(/<c\b([^>]*?)>/g,(all,attrs)=>{const ref=(attrs.match(/\br="([^"]+)"/)||[])[1],tone=map[ref];if(!tone)return all;const old=+(attrs.match(/\bs="(\d+)"/)||[])[1]||0;return '<c'+attrs.replace(/\s+s="\d+"/,'')+' s="'+variant(old,tone)+'">'});doc=doc.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/,'<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>');write(path,doc);}
 xml=xml.replace(/<cellXfs count="\d+">[\s\S]*?<\/cellXfs>/,'<cellXfs count="'+(original.length+variants.length)+'">'+original.join('')+variants.join('')+'</cellXfs>');write('xl/styles.xml',xml);
 return new Uint8Array(cfb.write(zip,{type:'buffer',fileType:'zip',compression:true}));
}
function exportReportingExcel(){try{const x=p(),wb=buildReportWorkbook(x);download('Reporting_'+x.store.replace(/[^a-z0-9]/gi,'_')+'.xlsx',styledWorkbookBytes(wb),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');note('Reporting généré : consulte tes téléchargements.')}catch(e){alert('Export du reporting impossible : '+e.message)}}
function exportExcel(){try{const x=p(),wb=buildSummaryWorkbook(x);download('Synthese_'+x.store.replace(/[^a-z0-9]/gi,'_')+'.xlsx',styledWorkbookBytes(wb),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');note('Synthèse du chantier générée : consulte tes téléchargements.')}catch(e){alert('Export Excel impossible : '+e.message)}}
