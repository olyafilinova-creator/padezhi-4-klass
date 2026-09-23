const SPREADSHEET_ID = '1hJ_1cPte4hH5lzT14JpLfmJ36XXWtpxw3bvwOCj3Q-M';
const API_TOKEN = '_bHISMz1zIrscq7n4XSu1aYs';

const RESULT_HEADERS = [
  'session_id','Дата отправки','Имя ученика','Блок','block_id','Начало','Завершение','Время, сек',
  'Всего заданий','Верно в итоге','Верно с первого раза','Итоговый %','% с первого раза','Ошибки',
  'Подсказки','Лиса использована','Вызовов лисы','Попытка','Устройство'
];
const ANSWER_HEADERS = [
  'session_id','Дата отправки','Имя ученика','Блок','block_id','№ задания','task_id','Тип задания',
  'Текст задания','Проверяемое слово','Правильный ответ','Ответ ученика','Правильно в итоге','С первого раза',
  'Попыток','Подсказка','Лиса','Время на задание, сек'
];

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Тренажёр')
    .addItem('Настроить листы','setupSheets')
    .addItem('Обновить аналитику','refreshAnalytics')
    .addToUi();
}

function doGet(){
  return json_({ok:true, service:'Падежи 4 класс', message:'Web App работает'});
}

function doPost(e){
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(15000)) return json_({ok:false,error:'busy'});
  try{
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const p = JSON.parse(raw);
    if(p.token !== API_TOKEN) return json_({ok:false,error:'bad token'});
    validatePayload_(p);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureSheets_(ss);
    const results = ss.getSheetByName('Результаты');
    const answers = ss.getSheetByName('Ответы');

    if(sessionExists_(results,p.session_id)) return json_({ok:true,duplicate:true});
    const attempt = getAttemptNumber_(results,p.student_name,p.block_id);
    const sentAt = new Date();

    results.appendRow([
      p.session_id, sentAt, p.student_name, p.block_name, p.block_id,
      new Date(p.started_at), new Date(p.finished_at), num_(p.duration_sec), num_(p.total_tasks),
      num_(p.final_correct), num_(p.first_try_correct), num_(p.final_pct), num_(p.first_try_pct),
      num_(p.errors), num_(p.hints_used), p.fox_used ? 'Да':'Нет', num_(p.fox_count), attempt, p.device || ''
    ]);

    const rows = (p.answers || []).map((a,i)=>[
      p.session_id, sentAt, p.student_name, p.block_name, p.block_id, i+1, a.task_id || '', a.task_type || '',
      a.question || '', a.target_word || '', a.correct_answer || '', a.student_answer || '',
      a.final_correct ? 'Да':'Нет', a.first_try_correct ? 'Да':'Нет', num_(a.attempts),
      a.hint_used ? 'Да':'Нет', a.fox_used ? 'Да':'Нет', num_(a.time_spent)
    ]);
    if(rows.length) answers.getRange(answers.getLastRow()+1,1,rows.length,ANSWER_HEADERS.length).setValues(rows);

    formatDataRows_(results,answers);
    refreshAnalytics_();
    return json_({ok:true,attempt:attempt});
  }catch(err){
    console.error(err);
    return json_({ok:false,error:String(err && err.message ? err.message : err)});
  }finally{
    lock.releaseLock();
  }
}

function setupSheets(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets_(ss);
  styleSheet_(ss.getSheetByName('Результаты'), RESULT_HEADERS.length);
  styleSheet_(ss.getSheetByName('Ответы'), ANSWER_HEADERS.length);
  refreshAnalytics_();
  SpreadsheetApp.getUi().alert('Готово: листы настроены, аналитика обновлена.');
}

function ensureSheets_(ss){
  let r=ss.getSheetByName('Результаты'); if(!r) r=ss.insertSheet('Результаты');
  let a=ss.getSheetByName('Ответы'); if(!a) a=ss.insertSheet('Ответы');
  let an=ss.getSheetByName('Аналитика'); if(!an) an=ss.insertSheet('Аналитика');
  ensureHeader_(r,RESULT_HEADERS); ensureHeader_(a,ANSWER_HEADERS);
}

function ensureHeader_(sheet,headers){
  const cur=sheet.getRange(1,1,1,headers.length).getValues()[0];
  if(cur.join('|')!==headers.join('|')) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleSheet_(sheet,headers.length);
}

function styleSheet_(sheet,cols){
  sheet.setFrozenRows(1);
  const head=sheet.getRange(1,1,1,cols);
  head.setBackground('#f4518c').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  head.setVerticalAlignment('middle');
  sheet.setRowHeight(1,34);
  for(let c=1;c<=cols;c++) sheet.setColumnWidth(c,120);
  if(sheet.getName()==='Результаты'){
    sheet.setColumnWidth(3,170); sheet.setColumnWidth(4,220); sheet.setColumnWidth(19,120);
    sheet.getRange('B:B').setNumberFormat('dd.MM.yyyy HH:mm:ss');
    sheet.getRange('F:G').setNumberFormat('dd.MM.yyyy HH:mm:ss');
  }else if(sheet.getName()==='Ответы'){
    sheet.setColumnWidth(3,170); sheet.setColumnWidth(4,220); sheet.setColumnWidth(9,330);
    sheet.setColumnWidth(10,160); sheet.setColumnWidth(11,180); sheet.setColumnWidth(12,180);
    sheet.getRange('B:B').setNumberFormat('dd.MM.yyyy HH:mm:ss');
  }
}

function formatDataRows_(results,answers){
  if(results.getLastRow()>1) results.getRange(2,1,results.getLastRow()-1,RESULT_HEADERS.length).setVerticalAlignment('middle');
  if(answers.getLastRow()>1) answers.getRange(2,1,answers.getLastRow()-1,ANSWER_HEADERS.length).setVerticalAlignment('top').setWrap(true);
}

function sessionExists_(sheet,sessionId){
  if(sheet.getLastRow()<2) return false;
  return !!sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(String(sessionId)).matchEntireCell(true).findNext();
}

function getAttemptNumber_(sheet,name,blockId){
  if(sheet.getLastRow()<2) return 1;
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,19).getValues();
  let n=0;
  data.forEach(r=>{if(String(r[2]).trim().toLowerCase()===String(name).trim().toLowerCase() && String(r[4])===String(blockId)) n++;});
  return n+1;
}

function refreshAnalytics(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); ensureSheets_(ss); refreshAnalytics_();
  SpreadsheetApp.getUi().alert('Аналитика обновлена.');
}

function refreshAnalytics_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const rs=ss.getSheetByName('Результаты'), as=ss.getSheetByName('Ответы'), out=ss.getSheetByName('Аналитика');
  const rdata=rs.getLastRow()>1 ? rs.getRange(2,1,rs.getLastRow()-1,19).getValues() : [];
  const adata=as.getLastRow()>1 ? as.getRange(2,1,as.getLastRow()-1,18).getValues() : [];
  out.clear();
  out.setFrozenRows(2);
  out.getRange('A1:I1').merge().setValue('Аналитика тренажёра «Падежи имени существительного»')
    .setBackground('#f4518c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14).setVerticalAlignment('middle');
  out.setRowHeight(1,36);

  const byStudent={};
  rdata.forEach(r=>{
    const name=String(r[2]||'').trim(); if(!name)return;
    const o=byStudent[name]||(byStudent[name]={runs:0,final:0,first:0,errors:0,hints:0,fox:0,time:0,last:null});
    o.runs++; o.final+=num_(r[11]); o.first+=num_(r[12]); o.errors+=num_(r[13]); o.hints+=num_(r[14]); o.fox+=num_(r[16]); o.time+=num_(r[7]);
    const dt=r[1] instanceof Date?r[1]:new Date(r[1]); if(!o.last || dt>o.last)o.last=dt;
  });
  const studentRows=Object.keys(byStudent).sort((a,b)=>a.localeCompare(b,'ru')).map(name=>{
    const o=byStudent[name]; return [name,o.runs,round1_(o.final/o.runs),round1_(o.first/o.runs),o.errors,o.hints,o.fox,Math.round(o.time/o.runs),o.last||''];
  });
  let row=3;
  writeSection_(out,row,'По ученикам',['Ученик','Прохождений','Средний итог %','С первого раза %','Ошибки','Подсказки','Вызовы лисы','Среднее время, сек','Последняя активность'],studentRows); row+=studentRows.length+4;

  const byBlock={};
  rdata.forEach(r=>{
    const name=String(r[3]||'').trim(); if(!name)return;
    const o=byBlock[name]||(byBlock[name]={runs:0,final:0,first:0,errors:0,hints:0,fox:0});
    o.runs++; o.final+=num_(r[11]); o.first+=num_(r[12]); o.errors+=num_(r[13]); o.hints+=num_(r[14]); o.fox+=num_(r[16]);
  });
  const blockRows=Object.keys(byBlock).map(name=>{const o=byBlock[name];return [name,o.runs,round1_(o.final/o.runs),round1_(o.first/o.runs),o.errors,o.hints,o.fox];});
  writeSection_(out,row,'По блокам',['Блок','Прохождений','Средний итог %','С первого раза %','Ошибки','Подсказки','Вызовы лисы'],blockRows); row+=blockRows.length+4;

  const diffs={};
  adata.forEach(r=>{
    const firstOk=String(r[13])==='Да'; if(firstOk)return;
    const key=String(r[3])+'||'+String(r[8]);
    const o=diffs[key]||(diffs[key]={block:String(r[3]),task:String(r[8]),errors:0,attempts:0,hints:0,fox:0});
    o.errors++; o.attempts+=num_(r[14]); if(String(r[15])==='Да')o.hints++; if(String(r[16])==='Да')o.fox++;
  });
  const diffRows=Object.values(diffs).sort((a,b)=>b.errors-a.errors).slice(0,20).map(o=>[o.block,o.task,o.errors,o.attempts,o.hints,o.fox]);
  writeSection_(out,row,'Частые затруднения',['Блок','Задание','Ошибок с первой попытки','Всего попыток','Подсказок','Лиса'],diffRows);

  out.getRange('A:K').setVerticalAlignment('top');
  out.setColumnWidth(1,210); out.setColumnWidth(2,330);
  for(let c=3;c<=11;c++)out.setColumnWidth(c,130);
  out.getRange('A:K').setWrap(true);
  if(studentRows.length) out.getRange(5,9,studentRows.length,1).setNumberFormat('dd.MM.yyyy HH:mm');
}

function writeSection_(sheet,start,title,headers,rows){
  sheet.getRange(start,1,1,headers.length).merge().setValue(title).setBackground('#ffe5ef').setFontWeight('bold').setFontSize(12);
  sheet.getRange(start+1,1,1,headers.length).setValues([headers]).setBackground('#fff0f6').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  if(rows.length) sheet.getRange(start+2,1,rows.length,headers.length).setValues(rows);
}

function validatePayload_(p){
  ['session_id','student_name','block_id','block_name','started_at','finished_at'].forEach(k=>{if(!p[k])throw new Error('missing '+k)});
  if(!Array.isArray(p.answers))throw new Error('answers must be array');
}
function num_(v){const n=Number(v);return Number.isFinite(n)?n:0}
function round1_(n){return Math.round(n*10)/10}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
