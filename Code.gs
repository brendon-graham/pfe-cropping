// PFE Cropping App v2 — Apps Script Backend
// Sheet ID: 1nSzOt6nBKqnYYmz8Dw4r7UsY05l699jMmamhQMYIoNk

const CROP_COLS = ['id','paddock','crop','ha','drillDate','yieldKgHA','seedHA','chemHA','fertHA','opsHA','notes'];
const STOCK_COLS = ['id','species','cls','headPrev','headCurr','kgDMday','period','days','feedSource','notes'];
const SUPP_COLS = ['id','name','type','kgDM','costPerKgDM','notes'];
const HIST_COLS = ['id','year','paddock','crop','ha','yieldBudget','yieldActual','notes'];

function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1B3A1B').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheetToRows(sh, cols) {
  if (sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2,1,sh.getLastRow()-1,cols.length).getValues();
  return vals.filter(r => r[0] !== '').map(r => {
    const o = {};
    cols.forEach((c,i) => { o[c] = r[i] === '' ? null : r[i]; });
    return o;
  });
}

function rowsToSheet(sh, cols, rows) {
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last-1);
  if (!rows.length) return;
  sh.getRange(2,1,rows.length,cols.length).setValues(rows.map(r => cols.map(c => r[c] != null ? r[c] : '')));
}

function getTs(ss) {
  const sh = ss.getSheetByName('Settings');
  if (!sh) return 0;
  const d = sh.getDataRange().getValues();
  const row = d.find(r => r[0] === 'lastModified');
  return row ? Number(row[1]) || 0 : 0;
}

function setTs(ss, ts) {
  let sh = ss.getSheetByName('Settings');
  if (!sh) { sh = ss.insertSheet('Settings'); sh.appendRow(['key','value']); }
  const d = sh.getDataRange().getValues();
  const idx = d.findIndex(r => r[0] === 'lastModified');
  if (idx >= 1) sh.getRange(idx+1,2).setValue(ts);
  else sh.appendRow(['lastModified', ts]);
}

function getAINotes(ss) {
  const sh = ss.getSheetByName('AILog');
  if (!sh || sh.getLastRow() < 2) return '';
  return sh.getRange(sh.getLastRow(),3).getValue() || '';
}

function appendAILog(ss, notes, ts) {
  let sh = ss.getSheetByName('AILog');
  if (!sh) { sh = ss.insertSheet('AILog'); sh.appendRow(['date','ts','notes']); sh.getRange(1,1,1,3).setFontWeight('bold'); }
  sh.appendRow([new Date(ts).toISOString().split('T')[0], ts, notes]);
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'pull';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'push') {
      const payload = JSON.parse(e.parameter.data);
      const curTs = getTs(ss);
      if (payload.clientTs && payload.clientTs < curTs) {
        return respond(JSON.stringify({status:'stale', lastModified:curTs}), e);
      }
      const newTs = Date.now();
      if (Array.isArray(payload.cropPlan))    rowsToSheet(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, payload.cropPlan);
      if (Array.isArray(payload.stockReq))    rowsToSheet(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, payload.stockReq);
      if (Array.isArray(payload.supplements)) rowsToSheet(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS, payload.supplements);
      if (Array.isArray(payload.history))     rowsToSheet(ensureSheet(ss,'History',HIST_COLS), HIST_COLS, payload.history);
      if (typeof payload.aiNotes === 'string') appendAILog(ss, payload.aiNotes, newTs);
      setTs(ss, newTs);
      return respond(JSON.stringify({status:'ok', lastModified:newTs}), e);
    }

    // pull
    ensureSheet(ss,'Settings',['key','value']);
    const result = {
      cropPlan:    sheetToRows(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS),
      stockReq:    sheetToRows(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS),
      supplements: sheetToRows(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS),
      history:     sheetToRows(ensureSheet(ss,'History',HIST_COLS), HIST_COLS),
      aiNotes:     getAINotes(ss),
      lastModified: getTs(ss)
    };
    return respond(JSON.stringify(result), e);

  } catch(err) {
    return respond(JSON.stringify({error: err.message}), e);
  }
}

function respond(json, e) {
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
