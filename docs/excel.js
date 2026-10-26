function requestModelIdentity(store,city){return new Promise(resolve=>{const dialog=document.createElement('dialog');dialog.style.cssText='border:0;border-radius:16px;padding:22px;max-width:90vw;width:400px;color:#18304b';dialog.setAttribute('aria-label','Identifier le chantier');const form=document.createElement('form');const title=document.createElement('h3');title.textContent='Quel chantier ?';const text=document.createElement('p');text.textContent='Ce tableau ne contient pas le magasin et la ville. Complète-les pour reprendre le chantier.';form.append(title,text);const inputs=[];for(const [caption,value] of [['Magasin',store],['Ville',city]]){const label=document.createElement('label');label.textContent=caption;const input=document.createElement('input');input.required=true;input.value=value||'';label.append(input);form.append(label);inputs.push(input)}const cancel=document.createElement('button');cancel.type='button';cancel.className='secondary';cancel.textContent='Annuler';const next=document.createElement('button');next.type='submit';next.textContent='Continuer';let done=false;const finish=value=>{if(done)return;done=true;dialog.close();dialog.remove();resolve(value)};cancel.onclick=()=>finish(null);dialog.oncancel=e=>{e.preventDefault();finish(null)};form.onsubmit=e=>{e.preventDefault();const values=inputs.map(n=>n.value.trim());if(values.every(Boolean))finish({store:values[0],city:values[1]})};form.append(cancel,next);dialog.append(form);document.body.append(dialog);dialog.showModal();(inputs.find(i=>!i.value)||inputs[0]).focus()})}
function parseSummaryModel(wb){
 const hoursFromCell=cell=>{const v=cell?.v;if(v===undefined||v===null||v==='')return 0;if(typeof v==='number'&&Number.isFinite(v)&&v>=0)return v*24;if(v instanceof Date)return (v.getTime()-Date.UTC(1899,11,30))/3600000;const m=String(v).trim().match(/^(\d+)[h:]\s*(\d{2})$/i);if(m&&+m[2]<60)return +m[1]+ +m[2]/60;throw Error('Durée illisible dans le tableau de synthèse.');};
 for(const name of wb.SheetNames){const ws=wb.Sheets[name];if(!ws['!ref'])continue;const range=XLSX.utils.decode_range(ws['!ref']);if(range.e.r>5000||range.e.c>100)continue;const cell=(r,c)=>ws[XLSX.utils.encode_cell({r,c})],key=(r,c)=>excelKey(cell(r,c)?.v);let ph,fh,dh,budgetRows={};
 for(let r=0;r<=Math.min(range.e.r,100);r++)for(let c=0;c<=range.e.c;c++){const k=key(r,c);if(k==='intervenant'&&key(r,c+1).includes('devis')&&key(r,c+2).includes('imagine'))ph={r,c};if(['famille','familles'].includes(k)&&key(r,c+1).includes('devis')&&key(r,c+2).includes('imagine'))fh={r,c};if(k==='jour'&&key(r,c+1).startsWith('total'))dh={r,c};if(k.startsWith('totaldevis'))budgetRows.Devis={r,c};if(k.startsWith('totalimagine'))budgetRows.Imagine={r,c};}
 if(!ph||!fh||!budgetRows.Devis||!budgetRows.Imagine)continue;
 const people=[],families=[],personTotals=[],allocs=[],days=[],planned={};
 const scan=(head,isPeople)=>{const list=[];for(let r=head.r+1;r<=range.e.r;r++){const label=String(cell(r,head.c)?.v??'').trim();if(!label||/^total/i.test(label))break;const d=hoursFromCell(cell(r,head.c+1)),i=hoursFromCell(cell(r,head.c+2));if(!d&&!i&&new RegExp(isPeople?'^intervenant\\s*\\d+$':'^famille\\s*\\d+$','i').test(label))continue;if(list.some(v=>excelKey(v.name)===excelKey(label)))throw Error('Nom en double dans le tableau : '+label);list.push({name:label,Devis:d,Imagine:i});}return list};
 for(const v of scan(ph,true)){people.push(v.name);personTotals.push({person:v.name,Devis:v.Devis,Imagine:v.Imagine})}
 for(const v of scan(fh,false)){families.push({name:v.name,types:['Devis','Imagine'],planned:null});for(const type of ['Devis','Imagine'])if(v[type])allocs.push({person:'',family:v.name,type,h:v[type]})}
 if(!people.length||!families.length)throw Error('Le tableau ne contient aucun intervenant ou famille.');
 for(const type of ['Devis','Imagine']){const pos=budgetRows[type],real=hoursFromCell(cell(pos.r,pos.c+1));planned[type]=hoursFromCell(cell(pos.r,pos.c+3));const perPerson=personTotals.reduce((n,v)=>n+v[type],0),perFamily=allocs.filter(a=>a.type===type).reduce((n,a)=>n+a.h,0);if(Math.abs(real-perPerson)>.001||Math.abs(real-perFamily)>.001)throw Error('Les totaux '+type+' du tableau ne correspondent pas aux intervenants et aux familles.');}
 if(dh)for(let r=dh.r+1;r<=range.e.r;r++){const label=String(cell(r,dh.c)?.v??'').trim();if(!label||/^total$/i.test(label))break;const h=hoursFromCell(cell(r,dh.c+1));if(h)days.push({date:'',day:label.replace(/^total\s*/i,''),h})}
 return {people,families,planned,entries:[],prior:{allocs,peopleTotals:personTotals,days,includedIds:[]},fromVisibleModel:true};
 }
 return null;
}
function confirmExcelImport(message,label){return new Promise(resolve=>{const dialog=document.createElement('dialog');dialog.setAttribute('aria-label','Import du chantier');dialog.style.cssText='border:0;border-radius:16px;padding:22px;max-width:90vw;width:400px;color:#18304b';const text=document.createElement('p');text.textContent=message;text.style.lineHeight='1.5';const cancel=document.createElement('button');cancel.type='button';cancel.className='secondary';cancel.textContent='Annuler';const accept=document.createElement('button');accept.type='button';accept.textContent=label;let finished=false;const finish=value=>{if(finished)return;finished=true;dialog.close();dialog.remove();resolve(value)};cancel.onclick=()=>finish(false);accept.onclick=()=>finish(true);dialog.oncancel=e=>{e.preventDefault();finish(false)};dialog.append(text,cancel,accept);document.body.append(dialog);dialog.showModal();cancel.focus()})}
function progressAllocations(x){return [...(x.prior?.allocs||[]),...x.entries.flatMap(e=>e.allocs.map(a=>({...a,person:e.person})))];}
function makePrior(x){
 const allocs=new Map(),days=new Map();for(const a of progressAllocations(x)){const key=JSON.stringify([a.person,a.family,a.type]);const old=allocs.get(key);if(old)old.h+=a.h;else allocs.set(key,{person:a.person,family:a.family,type:a.type,h:a.h});}
 for(const d of x.prior?.days||[])days.set(d.date||'label:'+d.day,{...d});for(const e of x.entries){const old=days.get(e.date)||{date:e.date,day:e.day||'',h:0};old.h+=e.allocs.reduce((n,a)=>n+a.h,0);days.set(e.date,old)}
 return {allocs:[...allocs.values()],peopleTotals:JSON.parse(JSON.stringify(x.prior?.peopleTotals||[])),days:[...days.values()],includedIds:[...new Set([...(x.prior?.includedIds||[]),...x.entries.map(e=>e.id)])]};
}
function priorWithoutOwn(snapshot,entries){const p=JSON.parse(JSON.stringify(snapshot));for(const e of entries){if(!p.includedIds.includes(e.id))continue;for(const a of e.allocs){const v=p.allocs.find(v=>v.person===e.person&&v.family===a.family&&v.type===a.type);if(!v||v.h+0.001<a.h)throw Error('Le fichier ne correspond pas aux saisies locales.');v.h=Math.max(0,v.h-a.h)}const day=p.days.find(d=>d.date===e.date);if(day)day.h=Math.max(0,day.h-e.allocs.reduce((n,a)=>n+a.h,0));}return p;}
/* Excel import/export for Suivi Remodeling. SheetJS CE 0.20.3 (Apache-2.0). */
function chooseExcel(){document.getElementById('excelFile').click()}
const excelKey=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
function excelNumber(v,label){if(v==null||String(v).trim()==='')return 0;const n=Number(String(v).replace(/\s/g,'').replace(',','.'));if(!Number.isFinite(n)||n<0)throw Error('Nombre invalide : '+label);return n}
function parseExcel(wb,mode="auto"){
 const sheets=wb.SheetNames.map(name=>({name,rows:XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''})}));
 const data={people:[],families:[],entries:[],planned:{Devis:0,Imagine:0}};let found=false;
 for(const {rows} of sheets){const hi=rows.findIndex(r=>r.some(v=>['famille','familles'].includes(excelKey(v)))&&r.some(v=>['ujtdevis','ujtimagine'].includes(excelKey(v))));if(hi<0)continue;found=true;
 const keys=rows[hi].map(excelKey),fi=keys.findIndex(k=>['famille','familles'].includes(k)),pi=keys.findIndex(k=>['intervenant','intervenants'].includes(k)),di=keys.indexOf('ujtdevis'),ii=keys.indexOf('ujtimagine'),ti=keys.indexOf('types'),ci=keys.indexOf('terminee');
 for(const [offset,r] of rows.slice(hi+1).entries()){const person=String(r[pi]??'').trim();if(person&&!data.people.includes(person))data.people.push(person);const name=String(r[fi]??'').trim();if(!name){if((di>=0&&r[di]!=='')||(ii>=0&&r[ii]!==''))throw Error('Famille manquante ligne '+(hi+offset+2));continue}const d=excelNumber(r[di],'Devis / '+name),i=excelNumber(r[ii],'Imagine / '+name);if(data.families.some(f=>excelKey(f.name)===excelKey(name)))throw Error('Famille en double : '+name);data.families.push({name,types:ti>=0&&String(r[ti]).split(',').every(t=>['Devis','Imagine'].includes(t))?String(r[ti]).split(','):[...(d?['Devis']:[]),...(i?['Imagine']:[])].length?[...(d?['Devis']:[]),...(i?['Imagine']:[])]:['Devis','Imagine'],planned:String(r[di]??'').trim()!==''||String(r[ii]??'').trim()!==''?{Devis:d*7,Imagine:i*7}:null});if(ci>=0)data.families[data.families.length-1].completed=['oui','true','1','terminee'].includes(excelKey(r[ci]));data.planned.Devis+=d*7;data.planned.Imagine+=i*7;}
 break;}
 if((!found||!data.families.length)&&!sheets.some(s=>s.name==='Chantier')){const model=parseSummaryModel(wb);if(model)return model;throw Error(mode==='resume'?'Choisis une synthèse de chantier déjà enregistrée. Le fichier à quatre colonnes sert uniquement à préparer un nouveau chantier.':'Colonnes attendues : Intervenant, Familles, UJT devis, UJT Imagine.');}
 const meta=sheets.find(s=>s.name==='Chantier');if(meta){const m=Object.fromEntries(meta.rows.map(r=>[excelKey(r[0]),r[1]]));data.store=String(m.magasin||'');data.city=String(m.ville||'');data.sourceId=String(m.identifiant||'');if(m.noteaucollegue!==undefined)data.noteToColleague=String(m.noteaucollegue);if(m.dateexport)data.receivedExportAt=String(m.dateexport);if(m.devisheures!==undefined)data.planned.Devis=excelNumber(m.devisheures,'budget Devis');if(m.imagineheures!==undefined)data.planned.Imagine=excelNumber(m.imagineheures,'budget Imagine');}
 const history=sheets.find(s=>s.name==='Heures');if(history){const k=history.rows[0].map(excelKey),get=(r,key)=>r[k.indexOf(key)];for(const r of history.rows.slice(1)){if(r.every(v=>v===''))continue;const person=String(get(r,'intervenant')||'').trim(),family=String(get(r,'famille')||'').trim(),type=get(r,'type'),date=String(get(r,'date')||'');const h=excelNumber(get(r,'heures'),'heures réalisées');if(!person||!family||!['Devis','Imagine'].includes(type)||!/^\d{4}-\d{2}-\d{2}$/.test(date)||!h)throw Error('Ligne d’historique invalide.');const e={id:String(get(r,'identifiant')||uid()),person,date,day:String(get(r,'journee')||''),start:String(get(r,'debut')),breakStart:String(get(r,'debutpause')),breakEnd:String(get(r,'finpause')),end:String(get(r,'fin')),allocs:[{family,type,h}]};const prior=data.entries.find(v=>v.id===e.id);if(prior){if(prior.date!==e.date||prior.person!==e.person||prior.start!==e.start||prior.end!==e.end)throw Error('Identifiant de saisie incohérent.');prior.allocs.push(e.allocs[0]);}else data.entries.push(e);if(!data.people.includes(person))data.people.push(person);if(!data.families.some(f=>f.name===family))data.families.push({name:family,types:[type]});}}
 for(const e of data.entries)if(!Number.isFinite(hours(e))||Math.abs(hours(e)-e.allocs.reduce((n,a)=>n+a.h,0))>.001)throw Error('Les horaires et heures réalisées ne correspondent pas.');
 const progress=sheets.find(s=>s.name==='Avancement');if(progress){const rows=progress.rows.slice(1);data.prior={allocs:rows.filter(r=>r.some(v=>v!=='')).map(r=>{const h=excelNumber(r[3],'avancement');if(!r[0]||!r[1]||!['Devis','Imagine'].includes(r[2]))throw Error('Avancement invalide.');return {person:String(r[0]),family:String(r[1]),type:r[2],h}}),days:[],includedIds:[]};const ds=sheets.find(s=>s.name==='Jours réalisés');if(ds)data.prior.days=ds.rows.slice(1).filter(r=>r[0]||r[1]).map(r=>{if(r[0]&&!/^\d{4}-\d{2}-\d{2}$/.test(String(r[0])))throw Error('Date invalide.');return {date:String(r[0]),day:String(r[1]||''),h:excelNumber(r[2],'total jour')}});const ids=sheets.find(s=>s.name==='Saisies incluses');if(ids)data.prior.includedIds=ids.rows.slice(1).filter(r=>r[0]||r[1]).map(r=>String(r[0]));}
 const familyProgress=sheets.find(s=>s.name==='Avancement familles'),personProgress=sheets.find(s=>s.name==='Totaux intervenants');if(familyProgress){data.prior??={allocs:[],days:[],includedIds:[]};for(const r of familyProgress.rows.slice(1)){if(!r[0])continue;if(!['Devis','Imagine'].includes(r[1]))throw Error('Type d’avancement invalide.');data.prior.allocs.push({person:'',family:String(r[0]),type:r[1],h:excelNumber(r[2],'avancement famille')});}data.prior.peopleTotals=personProgress?personProgress.rows.slice(1).filter(r=>r[0]).map(r=>({person:String(r[0]),Devis:excelNumber(r[1],'Devis intervenant'),Imagine:excelNumber(r[2],'Imagine intervenant')})):[];}
 if(mode==='resume'&&!data.sourceId&&!data.prior)throw Error('Ce fichier prépare un nouveau chantier. Importe-le dans « Chantiers à réaliser ». Pour continuer, choisis une synthèse enregistrée.');
 return data;
}
async function importExcel(ev){const input=ev.target,file=input.files[0],target=active,cat=category;if(!file)return;try{if(file.size>10*1024*1024)throw Error('Fichier trop volumineux (10 Mo maximum).');const current=target?db.projects.find(v=>v.id===target):null;const mode=(current?projectState(current):cat)==='ongoing'?'resume':'auto';const data=parseExcel(XLSX.read(await file.arrayBuffer(),{type:'array'}),mode);let x=target?db.projects.find(p=>p.id===target):null;if(x?.archived)throw Error('Ce chantier est archivé.');if(!x&&data.sourceId)x=db.projects.find(p=>p.id===data.sourceId&&!p.archived);if(x&&data.store&&(excelKey(data.store)!==excelKey(x.store)||excelKey(data.city)!==excelKey(x.city)))throw Error('Ce fichier correspond à un autre magasin ou une autre ville.');let store=x?.store||data.store||$('store').value.trim(),city=x?.city||data.city||$('city').value.trim();if(!store||!city){const identity=await requestModelIdentity(store,city);if(!identity)return;store=identity.store;city=identity.city;}
 const incoming=data.prior||data.entries.length?makePrior(data):null;const msg=incoming?'Reprendre le chantier '+store+' — '+city+' avec l’avancement du fichier Excel ?':'Importer les intervenants, les familles et les budgets dans le chantier '+store+' — '+city+' ?';if(!await confirmExcelImport(msg,incoming?'Reprendre le chantier':'Importer le fichier'))return;
 const next=x?JSON.parse(JSON.stringify(x)):{id:data.sourceId||uid(),store,city,people:[],families:[],entries:[],archived:false,started:cat==='ongoing'};
 next.people=[...new Set([...next.people,...data.people])];for(const f of data.families){const idx=next.families.findIndex(v=>excelKey(familyName(v))===excelKey(f.name));if(idx<0)next.families.push(f);else next.families[idx]={...(typeof next.families[idx]==='object'?next.families[idx]:{}),...f};}
 if(data.noteToColleague!==undefined)next.noteToColleague=data.noteToColleague;if(data.receivedExportAt)next.receivedExportAt=data.receivedExportAt;next.planned=data.planned;if(incoming)next.prior=priorWithoutOwn(incoming,next.entries);next.started=next.started||!!incoming||!!next.entries.length;
 const previous=db.projects;if(x)db.projects=db.projects.map(v=>v.id===x.id?next:v);else db.projects=[...db.projects,next];try{save()}catch(e){db.projects=previous;throw Error('Enregistrement impossible : espace de stockage insuffisant.');}openProject(next.id);if(incoming)showTab('summary');note(incoming?'Avancement repris : tu peux continuer le chantier.':'Fichier Excel importé.');}catch(e){const message=['NotFoundError','NotReadableError','SecurityError','NotAllowedError'].includes(e.name)?'Le fichier sélectionné est introuvable ou inaccessible sur cet appareil. Vérifie son accès dans Drive ou dans Fichiers, puis sélectionne-le à nouveau.':e.message;alert('Import impossible : '+message)}finally{input.value=''}}
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
 for(const a of x.prior?.allocs||[]){total[a.type]+=a.h;if(a.person){let person=people.get(a.person);if(!person){person={Devis:0,Imagine:0};people.set(a.person,person)}person[a.type]+=a.h;}let f=families.get(a.family);if(!f){f={Devis:0,Imagine:0,budget:null};families.set(a.family,f)}f[a.type]+=a.h;}
 for(const v of x.prior?.peopleTotals||[]){let person=people.get(v.person);if(!person){person={Devis:0,Imagine:0};people.set(v.person,person)}person.Devis+=v.Devis;person.Imagine+=v.Imagine;}
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
 raw('Chantier',[['Magasin',x.store],['Ville',x.city],['Identifiant',x.id],['Devis heures',x.planned.Devis],['Imagine heures',x.planned.Imagine],['Note au collègue',x.noteToColleague||''],['Date export',x.lastExportAt||'']]);
 raw('Preparation',[['Intervenant','Familles','UJT devis','UJT Imagine','Types','Terminée'],...Array.from({length:Math.max(x.people.length,x.families.length)},(_,i)=>[x.people[i]||'',familyName(x.families[i]||''),x.families[i]?.planned?x.families[i].planned.Devis/7:'',x.families[i]?.planned?x.families[i].planned.Imagine/7:'',x.families[i]?familyTypes(x.families[i]).join(','):'',x.families[i]?.completed?'Oui':'Non'])]);
 raw('Heures',[['Identifiant','Date','Journée','Intervenant','Début','Début pause','Fin pause','Fin','Famille','Type','Heures'],...x.entries.flatMap(e=>e.allocs.map(a=>[e.id,e.date,e.day,e.person,e.start,e.breakStart,e.breakEnd,e.end,a.family,a.type,a.h]))]);
 wb.Workbook={Sheets:wb.SheetNames.map(name=>({name,Hidden:['Chantier','Preparation','Heures'].includes(name)?1:0}))};
 return wb;
}
function fixedSummarySheet(x){
 const ws={},styles={},merges=[],dur='[h]:mm',people=new Map(x.people.map(n=>[n,{Devis:0,Imagine:0}])),families=new Map(x.families.map(f=>[familyName(f),{Devis:0,Imagine:0}])),total={Devis:0,Imagine:0},prior=makePrior(x);
 const put=(r,c,v,tone='modelPlain',z,f)=>{const ref=XLSX.utils.encode_cell({r,c});ws[ref]={t:typeof v==='number'?'n':'s',v};if(z)ws[ref].z=z;if(f)ws[ref].f=f;styles[ref]=tone};
 const hh=(r,c,h,tone='modelPlain',f)=>put(r,c,Math.round(h*60)/1440,tone,dur,f);
 for(const a of progressAllocations(x)){total[a.type]+=a.h;const family=families.get(a.family)||{Devis:0,Imagine:0};family[a.type]+=a.h;families.set(a.family,family);if(a.person){const person=people.get(a.person)||{Devis:0,Imagine:0};person[a.type]+=a.h;people.set(a.person,person)}}
 for(const p of prior.peopleTotals||[]){const person=people.get(p.person)||{Devis:0,Imagine:0};person.Devis+=p.Devis;person.Imagine+=p.Imagine;people.set(p.person,person)}
 put(0,2,'▼','modelArrow');['Heures\nréalisées','UJT\nPrévues','Heures\nPrévues','Delta'].forEach((v,i)=>put(1,i+1,v));
 ['Devis','Imagine'].forEach((type,i)=>{const r=i+2,tone=i?'modelImagine':'modelDevis';put(r,0,'Total '+type+' (h)',tone);hh(r,1,total[type],tone);put(r,2,x.planned[type]/7,tone,Number.isInteger(x.planned[type]/7)?'0':'0.##');hh(r,3,x.planned[type],tone,'C'+(r+1)+'*7/24');const delta=Math.round((total[type]-x.planned[type])*60),abs=Math.abs(delta);put(r,4,(delta<=0?'-':'')+Math.floor(abs/60)+'h'+String(abs%60).padStart(2,'0'),'modelDelta',undefined,'IF(B'+(r+1)+'<=D'+(r+1)+',"-","")&TEXT(ABS(B'+(r+1)+'-D'+(r+1)+'),"[h]""h""mm")')});
 ['Intervenant','Devis (h)','Imagine (h)','Total (h)'].forEach((v,i)=>put(7,i,v,i===1?'modelDevis':i===2?'modelImagine':'modelPlain'));
 const personRows=Math.max(10,people.size),familyRows=Math.max(20,families.size),dayRows=Math.max(31,prior.days.length),personList=[...people],familyList=[...families];
 for(let i=0;i<personRows;i++){const [name,v]=personList[i]||['Intervenant '+(i+1),{Devis:0,Imagine:0}],r=i+8;put(r,0,name,'modelPerson');hh(r,1,v.Devis,'modelDevis');hh(r,2,v.Imagine,'modelImagine');hh(r,3,v.Devis+v.Imagine,'modelPlain','SUM(B'+(r+1)+':C'+(r+1)+')')}
 ['Famille','Devis (h)','Imagine (h)','Total (h)'].forEach((v,i)=>put(7,i+4,v,i===1?'modelDevis':i===2?'modelImagine':'modelPlain'));
 for(let i=0;i<familyRows;i++){const [name,v]=familyList[i]||['Famille '+(i+1),{Devis:0,Imagine:0}],r=i+8;put(r,4,name);hh(r,5,v.Devis,'modelDevis');hh(r,6,v.Imagine,'modelImagine');hh(r,7,v.Devis+v.Imagine,'modelPlain','SUM(F'+(r+1)+':G'+(r+1)+')')}
 put(0,9,'Total par jour','modelDay');put(0,10,'','modelDay');put(1,9,'Jour','modelDay');put(1,10,'Total (h)','modelDay');
 for(let i=0;i<dayRows;i++){const d=prior.days[i];const label=d?.date?d.date.split('-').reverse().join('/'):d?.day||'Jour '+(i+1);put(i+2,9,d?'Total '+label:'Total Jour '+(i+1));hh(i+2,10,d?.h||0,i===0||i===2?'modelDayDevis':i===1?'modelDayImagine':'modelPlain')}
 ws.B3.f='SUM(B9:B'+(personRows+8)+')';ws.B4.f='SUM(C9:C'+(personRows+8)+')';
 let bottom=Math.max(8+personRows,8+familyRows,2+dayRows);
 const line=text=>{const r=++bottom;put(r,0,text,'modelNote');merges.push({s:{r,c:0},e:{r,c:10}})};
 const done=x.families.filter(f=>f.completed).map(familyName);if(done.length){bottom++;line('Familles terminées : '+done.join(', '));}
 if(x.noteToColleague?.trim()){bottom++;line('Note pour le collègue');for(const text of x.noteToColleague.split('\n').flatMap(v=>v.match(/.{1,140}/g)||['']))line(text)}
 ws['!ref']='A1:L'+(bottom+2);ws['!cols']=[143,56,56,56,130,56,56,58,42,83,80,24].map(wpx=>({wpx}));ws['!rows']=Array.from({length:bottom+2},(_,i)=>({hpx:i===0?20:i===1?23:i===6?10:i>=33?20:12}));ws['!merges']=merges;
 return {sheet:ws,styles};
}
function buildSummaryWorkbook(x){
 const wb=buildReportWorkbook(x),fixed=fixedSummarySheet(x),prior=makePrior(x),raw=(name,rows)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),name);
 raw('Avancement',[['Intervenant','Famille','Type','Heures réalisées'],...prior.allocs.filter(a=>a.person).map(a=>[a.person,a.family,a.type,a.h])]);
 raw('Avancement familles',[['Famille','Type','Heures réalisées'],...prior.allocs.filter(a=>!a.person).map(a=>[a.family,a.type,a.h])]);
 raw('Totaux intervenants',[['Intervenant','Devis heures','Imagine heures'],...(prior.peopleTotals||[]).map(v=>[v.person,v.Devis,v.Imagine])]);
 raw('Jours réalisés',[['Date','Journée','Total heures'],...prior.days.map(d=>[d.date,d.day,d.h])]);raw('Saisies incluses',[['Identifiant'],...prior.includedIds.map(id=>[id])]);
 wb.Sheets['Synthèse']=fixed.sheet;const remove=['Horaires','Totaux journaliers','Familles','Heures'];wb.SheetNames=wb.SheetNames.filter(n=>!remove.includes(n));for(const n of remove)delete wb.Sheets[n];
 wb.Workbook.Sheets=wb.SheetNames.map(name=>({name,Hidden:name==='Synthèse'?0:1}));wb.__reportStyles={'Synthèse':fixed.styles};wb.__protectSummary=true;return wb;
}
// Add report fonts, borders and fills to the generated XLSX while preserving number formats.
function styledWorkbookBytes(wb){
 const cfb=XLSX.CFB,zip=cfb.read(new Uint8Array(XLSX.write(wb,{type:'array',bookType:'xlsx'})),{type:'buffer'}),decoder=new TextDecoder(),encoder=new TextEncoder();
 const read=path=>decoder.decode(cfb.find(zip,'/'+path).content),write=(path,text)=>cfb.utils.cfb_add(zip,'/'+path,encoder.encode(text));
 let xml=read('xl/styles.xml');
 const append=(tag,items)=>{let base=0;xml=xml.replace(new RegExp('<'+tag+' count="(\\d+)">([\\s\\S]*?)</'+tag+'>'),(_,count,body)=>{base=+count;return '<'+tag+' count="'+(base+items.length)+'">'+body+items.join('')+'</'+tag+'>'});return base};
 const fontBase=append('fonts',[['20352A',false,11],['FFFFFF',true,11],['20352A',true,11],['FFFFFF',true,18],['000000',false,7],['FFFF00',false,7],['FF4000',false,22]].map(([color,bold,size])=>'<font>'+(bold?'<b/>':'')+'<sz val="'+size+'"/><color rgb="FF'+color+'"/><name val="Calibri"/></font>'));
 const tones={reportHeader:['DAEAF5',2],reportDevis:['E49EDD',0],reportImagine:['41B1DB',0],reportDate:['EAD1E8',0],reportMixed:['E8EFF5',0],reportTotal:['DAEAF5',2],plain:['FFFFFF',0],title:['006341',3],subtitle:['E2EFDA',2],header:['006341',1],devis:['E8B0DF',0],imagine:['A6DCEF',0],total:['E2EFDA',2],remaining:['FFF2CC',2],overrun:['F8B4B4',2],note:['FFFFFF',0],devisHeader:['CF85C6',2],imagineHeader:['73C5E5',2],modelPlain:['FFFFFF',4],modelPerson:['FFFFFF',4],modelDevis:['E49EDD',4],modelImagine:['83CCEB',4],modelDay:['DAF2D0',4],modelDelta:['FF0000',5],modelArrow:['FFFFFF',6],modelNote:['FFFFFF',4],modelDayDevis:['F2CEEF',4],modelDayImagine:['D4EAF4',4]};
 const names=Object.keys(tones),fillBase=append('fills',names.map(n=>'<fill><patternFill patternType="solid"><fgColor rgb="FF'+tones[n][0]+'"/><bgColor indexed="64"/></patternFill></fill>'));
 const borderBase=append('borders',['<border><left style="thin"><color rgb="FFAAB8AE"/></left><right style="thin"><color rgb="FFAAB8AE"/></right><top style="thin"><color rgb="FFAAB8AE"/></top><bottom style="thin"><color rgb="FFAAB8AE"/></bottom><diagonal/></border>']);
 const match=xml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/),original=match[2].match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g),variants=[],cache=new Map();
 const variant=(id,tone)=>{const key=id+':'+tone;if(cache.has(key))return cache.get(key);const source=original[id]||original[0],numFmt=(source.match(/numFmtId="(\d+)"/)||[])[1]||'0',index=original.length+variants.length;variants.push('<xf numFmtId="'+numFmt+'" fontId="'+(fontBase+tones[tone][1])+'" fillId="'+(fillBase+names.indexOf(tone))+'" borderId="'+(['modelNote','modelArrow'].includes(tone)?0:borderBase)+'" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="'+(tone==='modelPerson'||tone==='modelNote'?'left':tone.startsWith('model')?'center':'left')+'" vertical="center" wrapText="1"/></xf>');cache.set(key,index);return index};
 for(const [sheet,map] of Object.entries(wb.__reportStyles||{})){const path='xl/worksheets/sheet'+(wb.SheetNames.indexOf(sheet)+1)+'.xml';let doc=read(path);doc=doc.replace(/<c\b([^>]*?)>/g,(all,attrs)=>{const ref=(attrs.match(/\br="([^"]+)"/)||[])[1],tone=map[ref];if(!tone)return all;const old=+(attrs.match(/\bs="(\d+)"/)||[])[1]||0;return '<c'+attrs.replace(/\s+s="\d+"/,'')+' s="'+variant(old,tone)+'">'});doc=doc.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/,'<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>');if(wb.__protectSummary)doc=doc.replace('</sheetData>','</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="0" selectUnlockedCells="0"/>');write(path,doc);}
 xml=xml.replace(/<cellXfs count="\d+">[\s\S]*?<\/cellXfs>/,'<cellXfs count="'+(original.length+variants.length)+'">'+original.join('')+variants.join('')+'</cellXfs>');write('xl/styles.xml',xml);
 return new Uint8Array(cfb.write(zip,{type:'buffer',fileType:'zip',compression:true}));
}
function exportFileName(prefix,store,stamp){const d=new Date(stamp),pad=n=>String(n).padStart(2,'0');return prefix+'_'+store.replace(/[^a-z0-9]/gi,'_')+'_'+d.getFullYear()+pad(d.getMonth()+1)+pad(d.getDate())+'_'+pad(d.getHours())+pad(d.getMinutes())+pad(d.getSeconds())+'.xlsx';}
function recordGeneratedExport(x,kind,stamp){x[kind==='summary'?'lastExportAt':'lastReportingExportAt']=stamp;if(typeof save==='function')save();if(typeof renderExportDates==='function')renderExportDates();}
// Personal reporting lists only this device's recorded slots; imported progress has no schedules.
function buildPersonalReportWorkbook(x){
 const wb=XLSX.utils.book_new(),styles={},rows=[['Date','Jour','Intervenant','Type intervention','Famille','Heure début','Début pause','Fin pause','Heure fin','Temps total (h)','Commentaire']],rowTones=['reportHeader'],totals={Devis:0,Imagine:0};
 const time=v=>{if(!v)return '';const n=minutes(v);return {t:'n',v:n/1440,z:'h:mm'}};
 const entries=[...x.entries].sort((a,b)=>a.date.localeCompare(b.date)||a.start.localeCompare(b.start)||a.person.localeCompare(b.person,'fr'));
 for(const e of entries){
  let h=0;for(const a of e.allocs){if(!['Devis','Imagine'].includes(a.type)||!Number.isFinite(a.h)||a.h<0)throw Error('Saisie d’heures invalide.');h+=a.h;totals[a.type]+=a.h;}
  if(!Number.isFinite(hours(e))||Math.abs(hours(e)-h)>.001)throw Error('Les heures affectées ne correspondent pas aux horaires de '+e.person+' le '+e.date+'.');
  const types=[...new Set(e.allocs.map(a=>a.type))],mixed=e.allocs.length>1;
  const duration=v=>Math.floor(Math.round(v*60)/60)+':'+String(Math.round(v*60)%60).padStart(2,'0');
  const typeLabel=types.length===1?types[0]:types.map(t=>t+' ('+duration(e.allocs.filter(a=>a.type===t).reduce((n,a)=>n+a.h,0))+')').join(' + ');
  const family=mixed?e.allocs.map(a=>a.family+' — '+a.type+' '+duration(a.h)).join('\n'):e.allocs[0]?.family||'';
  rows.push([{t:'n',v:(Date.parse(e.date+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000,z:'dd/mm/yyyy'},e.day||'',e.person,typeLabel,family,time(e.start),e.breakStart===e.breakEnd?'':time(e.breakStart),e.breakStart===e.breakEnd?'':time(e.breakEnd),time(e.end),{t:'n',v:Math.round(h*60)/1440,z:'[h]:mm'},e.comment||'']);
  rowTones.push(types.length===1?'report'+types[0]:'reportMixed');
 }
 rows.push(Array(11).fill(''));rowTones.push('plain');
 for(const [label,h,tone] of [['TOTAL DEVIS',totals.Devis,'reportDevis'],['TOTAL IMAGINE',totals.Imagine,'reportImagine'],['TOTAL',totals.Devis+totals.Imagine,'reportTotal']]){const row=Array(11).fill('');row[8]=label;row[9]={t:'n',v:Math.round(h*60)/1440,z:'[h]:mm'};rows.push(row);rowTones.push(tone)}
 const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[100,60,145,165,200,95,95,95,95,115,280].map(wpx=>({wpx}));
 ws['!rows']=rows.map((r,i)=>({hpt:i===0?30:Math.max(24,Math.min(180,16*Math.max(String(r[4]||'').split('\n').length,Math.ceil(String(r[10]||'').length/35))))}));
 ws['!autofilter']={ref:'A1:K'+Math.max(1,entries.length+1)};
 for(let r=0;r<rows.length;r++)for(let col=0;col<11;col++)styles[XLSX.utils.encode_cell({r,c:col})]=r>0&&r<=entries.length&&col===0?'reportDate':rowTones[r];
 XLSX.utils.book_append_sheet(wb,ws,'Horaires');wb.__reportStyles={Horaires:styles};wb.Props={Title:'Mes horaires — '+x.store+' — '+x.city};return wb;
}
function exportReportingExcel(){openExcelPreview('reporting')}
let pendingSummaryExport=null;
// Display the same visible sheet whose bytes are saved after the user's review.
function summaryPreviewHTML(wb,sheetName='Synthèse'){
 const ws=wb.Sheets[sheetName],styles=wb.__reportStyles[sheetName],range=XLSX.utils.decode_range(ws['!ref']),merges=ws['!merges']||[];
 let html='<table aria-label="'+esc(sheetName)+'"><colgroup>';
 for(let c=0;c<=range.e.c;c++)html+='<col style="width:'+Math.max(48,ws['!cols']?.[c]?.wpx||80)+'px">';
 html+='</colgroup><tbody>';
 for(let r=0;r<=range.e.r;r++){html+='<tr>';for(let c=0;c<=range.e.c;c++){
  const merge=merges.find(m=>r>=m.s.r&&r<=m.e.r&&c>=m.s.c&&c<=m.e.c);if(merge&&(r!==merge.s.r||c!==merge.s.c))continue;
  const ref=XLSX.utils.encode_cell({r,c}),cell=ws[ref],tone=styles[ref]||'modelPlain';
  const span=merge?' colspan="'+(merge.e.c-merge.s.c+1)+'" rowspan="'+(merge.e.r-merge.s.r+1)+'"':'';
  html+='<td class="'+esc(tone)+'"'+span+'>'+esc(cell?XLSX.utils.format_cell(cell):'')+'</td>';
 }html+='</tr>';}return html+'</tbody></table>';
}
function exportExcel(){openExcelPreview('summary')}
function openExcelPreview(kind){try{
 const x=p(),stamp=new Date().toISOString(),report=kind==='reporting',wb=report?buildPersonalReportWorkbook(x):buildSummaryWorkbook({...x,lastExportAt:stamp});
 const prepared={projectId:x.id,kind,stamp,name:exportFileName(report?'Reporting':'Synthese',x.store,stamp),bytes:styledWorkbookBytes(wb)};
 const html=summaryPreviewHTML(wb,report?'Horaires':'Synthèse');pendingSummaryExport=prepared;
 $('handoverPanel').classList.add('hide');$('handoverFallback').classList.add('hide');$('openHandoverButton').classList.toggle('hide',report);
 $('excelPreviewTitle').textContent=report?'Mes horaires pour le reporting':'Aperçu de la synthèse Excel';
 $('excelPreviewIdentity').textContent=x.store+' — '+x.city;
 $('excelPreviewTable').innerHTML=html;$('excelPreviewStatus').textContent=(report&&!x.entries.length?'Aucun horaire saisi sur cet appareil. ':'')+'Aucun fichier enregistré. Choisis « Enregistrer le fichier Excel » après vérification.';
 showTab('excelPreview');$('excelPreview').scrollIntoView({block:'start'});
}catch(e){pendingSummaryExport=null;alert('Aperçu Excel impossible : '+e.message)}}
function savePreviewExcel(){try{
 const prepared=pendingSummaryExport,x=p();if(!prepared||!x||prepared.projectId!==x.id)throw Error('Rouvre l’aperçu du fichier avant d’enregistrer.');
 download(prepared.name,prepared.bytes,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
 recordGeneratedExport(x,prepared.kind,prepared.stamp);
 $('excelPreviewStatus').textContent='Enregistrement demandé pour « '+prepared.name+' ». Vérifie l’emplacement proposé par ton appareil.';
}catch(e){alert('Enregistrement Excel impossible : '+e.message)}}
