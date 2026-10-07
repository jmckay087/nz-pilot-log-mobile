function fpNum(v){return v===''||v===null||v===undefined?'':Math.round(num(v)*1000000)/1000000}function rowFingerprint(row){row=row||[];const caa=[...row.slice(6,21).map(fpNum),fpNum(row[23]),fpNum(row[24])],ap=[fpNum(row[21]),fpNum(row[22])],duty=[fpNum(row[32]),fpNum(row[33])];return JSON.stringify([String(row[0]??''),String(row[1]??'').trim().toUpperCase(),normaliseReg(row[2]),String(row[3]??'').trim(),String(row[4]??'').trim(),String(row[5]??'').trim(),...caa,...ap,fpNum(row[28]),fpNum(row[29]),fpNum(row[30]),fpNum(row[31]),...duty])}function flightFingerprint(f){const c=caaColumnsFor(f),ap=approachSplit(f),on=excelTimeFraction(f.dutyStart),off=excelTimeFraction(f.dutyEnd);return JSON.stringify([String(excelSerial(f.date)),String(f.aircraftType||'').trim().toUpperCase(),normaliseReg(f.registration),String(f.pilotInCommand||'').trim(),String(f.coPilotStudent||'').trim(),detailsFor(f),...c.slice(0,15).map(fpNum),fpNum(c[15]),fpNum(c[16]),fpNum(ap.non),fpNum(ap.prec),fpNum(f.takeoffsDay),fpNum(f.takeoffsNight),fpNum(f.landingsDay),fpNum(f.landingsNight),on===null?'':fpNum(on),off===null?'':fpNum(off)])}
async function updateWorkbook(ab,flights){const rowsBefore=await readLogbookRows(ab);if(!rowsBefore[3]||String(rowsBefore[3][0]||'').trim().toLowerCase()!=='date'||String(rowsBefore[3][5]||'').trim().toLowerCase()!=='details of flight')throw Error('This does not match your CAANZ Excel Pilot Logbook layout.');const z=await unzipXlsx(ab),sheetPath=await logbookSheetPath(z),sheetBytes=await z.get(sheetPath),doc=parseXml(sheetBytes),ss=await loadSharedStringStore(z),written=[];for(const f of flights){let row=+(f.excelRow||f.proposedRow||0);if(row>=5&&row<=2032&&rowFingerprint(rowsBefore[row-1])===flightFingerprint(f)){written.push({id:f.id,row,already:true});continue}if(!(row>=5&&row<=2032)||logbookRowHasUserData([...doc.getElementsByTagNameNS(XLSX_NS,'row')].find(r=>+r.getAttribute('r')===row)))row=nextAppendLogbookRow(doc);writeFlightToDoc(doc,row,f,ss);written.push({id:f.id,row,already:false});f.proposedRow=row}const enc=new TextEncoder(),repl=new Map([[sheetPath,enc.encode(xlsxXmlString(doc))],['xl/sharedStrings.xml',ss.finish()]]);const wbBytes=await z.get('xl/workbook.xml');if(wbBytes){const wbDoc=parseXml(wbBytes);let calc=[...wbDoc.getElementsByTagNameNS(XLSX_NS,'calcPr')][0];if(!calc){calc=wbDoc.createElementNS(XLSX_NS,'calcPr');wbDoc.documentElement.appendChild(calc)}calc.setAttribute('calcMode','auto');calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');repl.set('xl/workbook.xml',enc.encode(xlsxXmlString(wbDoc)))}const out=buildZipArchive(ab,repl),rowsAfter=await readLogbookRows(out.buffer);for(const w of written){const f=flights.find(x=>x.id===w.id);if(rowFingerprint(rowsAfter[w.row-1])!==flightFingerprint(f))throw Error(`Excel verification failed on row ${w.row}. Nothing has been marked synced.`)}return{bytes:out,written}}

let prepared=null;
const excelInput=$('excelFile');
const originalChoose=$('chooseExcel');
let chooseExcelTrigger=originalChoose;
if(excelInput&&originalChoose){
  excelInput.hidden=false;
  excelInput.style.position='fixed';
  excelInput.style.left='-10000px';
  excelInput.style.top='0';
  excelInput.style.width='1px';
  excelInput.style.height='1px';
  excelInput.style.opacity='0';
  excelInput.style.pointerEvents='none';
  excelInput.setAttribute('tabindex','-1');
  const label=document.createElement('label');
  label.id='chooseExcel';
  label.className=originalChoose.className;
  label.setAttribute('for','excelFile');
  label.setAttribute('role','button');
  label.textContent=originalChoose.textContent;
  label.style.display='flex';
  label.style.alignItems='center';
  label.style.justifyContent='center';
  label.style.cursor='pointer';
  label.style.userSelect='none';
  originalChoose.replaceWith(label);
  chooseExcelTrigger=label;
  label.addEventListener('click',e=>{if(label.dataset.busy==='1')e.preventDefault()});
}

excelInput.onchange=async e=>{const file=e.target.files[0];e.target.value='';if(!file)return;const flights=pendingFlights();if(!flights.length){alert('There are no phone flights waiting for Excel.');return}const btn=chooseExcelTrigger;try{btn.dataset.busy='1';btn.setAttribute('aria-disabled','true');btn.style.opacity='.65';btn.textContent='Preparing…';const ab=await file.arrayBuffer(),result=await updateWorkbook(ab,flights),t=Date.now(),already=result.written.filter(w=>w.already),needsSave=result.written.filter(w=>!w.already);for(const w of result.written){const f=db.flights.find(x=>x.id===w.id);if(f)f.proposedRow=w.row}for(const w of already){const f=db.flights.find(x=>x.id===w.id);if(f){f.syncStatus='synced';f.excelRow=w.row;f.proposedRow=0;f.syncedAt=t;f.workbook=file.name}}db.settings.lastExcelName=file.name;if(already.length)db.settings.lastSyncAt=t;localStorage.setItem(STORE,JSON.stringify(db));if(!needsSave.length){prepared=null;alert(`The selected workbook already contains all ${already.length} queued flight${already.length===1?'':'s'}. They have been marked synced on this phone.`);renderAll();switchTab('home');return}prepared={name:file.name,bytes:result.bytes,written:needsSave,flightIds:needsSave.map(w=>w.id)};$('preparedBox').classList.remove('hidden');$('syncStatus').className='notice good';$('syncStatus').innerHTML=`Prepared <strong>${needsSave.length}</strong> flight${needsSave.length===1?'':'s'} for ${esc(file.name)}${already.length?` · ${already.length} was already present`:''}. Now save the updated workbook back to iCloud Drive.`;await sharePrepared()}catch(err){console.error(err);prepared=null;$('preparedBox').classList.add('hidden');$('syncStatus').className='notice bad';$('syncStatus').textContent='Could not prepare the Excel workbook: '+(err.message||err)}finally{btn.dataset.busy='0';btn.removeAttribute('aria-disabled');btn.style.opacity='';btn.textContent='Choose Excel & prepare update';renderAll();switchTab('sync')}};
async function sharePrepared(){if(!prepared)return;const file=new File([prepared.bytes],prepared.name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});try{if(navigator.share&&(!navigator.canShare||navigator.canShare({files:[file]}))){await navigator.share({title:'Updated NZ Pilot Log Excel workbook',files:[file]});return true}}catch(err){if(err?.name==='AbortError')return false;console.warn('Share failed',err)}const a=document.createElement('a');a.href=URL.createObjectURL(file);a.download=prepared.name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000);return true}
