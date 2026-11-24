const TABS = ['Усі відповіді','Реальний досвід','Ідеї партнерства','Мрії'];
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID',ss.getId());
  TABS.forEach((name,i)=>{
    let s=ss.getSheetByName(name)||ss.insertSheet(name);
    let headers=i===0?['ID','Дата','Учасник','Реальний досвід','Ідея партнерства','Мрія','Розподіл карток']:['Учасник',name];
    s.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#173e60').setFontColor('#ffffff').setFontWeight('bold');s.setFrozenRows(1);
    s.setColumnWidth(i===0?3:1,200);
    if(i===0){s.setColumnWidths(4,3,420);s.setColumnWidth(7,300);}else{s.setColumnWidth(2,700);const col=['','D','E','F'][i];const sep=ss.getSpreadsheetLocale().startsWith('en')?',':';'; for(const [dst,src] of [['A2','C'],['B2',col]])s.getRange(dst).setFormula(`=IFERROR(FILTER('Усі відповіді'!${src}2:${src}${sep}'Усі відповіді'!A2:A<>"")${sep}"")`);}
    s.getDataRange().setWrap(true);
  });
}
function reply(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){let lock=LockService.getScriptLock();try{let r=JSON.parse(e.postData.contents);if(typeof r.id!=='string'||r.id.length>100||!['real','idea','dream'].every(k=>typeof r[k]==='string'&&r[k].trim()&&r[k].length<=10000))throw Error('Недійсні відповіді');lock.waitLock(20000);let s=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName(TABS[0]);let existing=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,1).createTextFinder(r.id).matchEntireCell(true).findNext():null;let row=existing?existing.getRow():s.getLastRow()+1;let safe=x=>"'"+String(x||'').slice(0,10000);s.getRange(row,1,1,7).setValues([[safe(r.id),new Date(),safe(r.author||'Анонімно'),safe(r.real),safe(r.idea),safe(r.dream),safe(JSON.stringify(r.cards||[]))]]).setWrap(true);SpreadsheetApp.flush();return reply({ok:true});}catch(err){return reply({ok:false,error:String(err)});}finally{if(lock.hasLock())lock.releaseLock();}}
