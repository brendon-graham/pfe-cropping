// PFE Cropping App v2.2 — Apps Script Backend
// Sheet ID: 1nSzOt6nBKqnYYmz8Dw4r7UsY05l699jMmamhQMYIoNk
// 'year' added at end of CROP/STOCK/SUPP cols — old rows without it default to 2026/27 in the app.
// v2.2: 'wastage' added after yieldKgHA — old rows without it get crop-type default in the app on load.

const CROP_COLS  = ['id','paddock','crop','ha','drillDate','yieldKgHA','wastage','seedHA','chemHA','fertHA','opsHA','notes','year'];
const STOCK_COLS = ['id','species','cls','headPrev','headCurr','kgDMday','period','days','feedSource','notes','year'];
const SUPP_COLS  = ['id','name','type','kgDM','costPerKgDM','notes','year'];
const HIST_COLS  = ['id','year','paddock','crop','ha','yieldBudget','yieldActual','notes'];

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

function appendRows(sh, cols, rows) {
  const last = sh.getLastRow();
  sh.getRange(last+1, 1, rows.length, cols.length)
    .setValues(rows.map(r => cols.map(c => r[c] != null ? r[c] : '')));
  SpreadsheetApp.flush();
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
      aiNotes:     getAINotes(ss),
      lastModified: getTs(ss)
    };
    return respond(JSON.stringify(result), e);

  } catch(err) {
    return respond(JSON.stringify({error: err.message}), e);
  }
}

// ============================================================
// SEEDING — run each function separately from the editor.
// ORDER FOR FRESH SETUP:
//   1. seedCropPlan        — clears + writes 26/27 crop plan
//   2. seedStockReq        — clears + writes 26/27 stock req
//   3. seedSupplements     — clears + writes 26/27 supplements
//   Then APPEND historical seasons (each adds to existing rows):
//   4. seedCropPlan2526    5. seedStockReq2526    6. seedSupps2526
//   7. seedCropPlan2425    8. seedStockReq2425    9. seedSupps2425
//  10. seedCropPlan2324   11. seedStockReq2324
//  12. seedCropPlan2223
//  13. seedCropPlan2122
// ============================================================

// ---------- 2026/27 (base season) ----------

function seedCropPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2026/27';
  const rows = [
    {id:'k1',year:y,paddock:'North Harbour',crop:'Kale',ha:5.2,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k2',year:y,paddock:'Wanganui',crop:'Kale',ha:5.47,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k3',year:y,paddock:'Bull',crop:'Kale',ha:6.95,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k4',year:y,paddock:'Bens',crop:'Kale',ha:5.17,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k5',year:y,paddock:'Waikato',crop:'Kale',ha:11,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'s1',year:y,paddock:'River 1',crop:'Swedes',ha:7.9,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s2',year:y,paddock:'River 2',crop:'Swedes',ha:5.17,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s3',year:y,paddock:'River 3',crop:'Swedes',ha:5.14,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s4',year:y,paddock:'River 4',crop:'Swedes',ha:3.3,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s5',year:y,paddock:'Horsfall 3',crop:'Swedes',ha:5.6,drillDate:'11/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s6',year:y,paddock:'Ryans',crop:'Swedes',ha:4.7,drillDate:'12/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s7',year:y,paddock:'Horsfall 2',crop:'Swedes',ha:4.2,drillDate:'13/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s8',year:y,paddock:'Yards',crop:'Swedes',ha:5.03,drillDate:'14/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s9',year:y,paddock:'Cattle Yards 3',crop:'Swedes',ha:4.53,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s10',year:y,paddock:'Little Horsfall Swamp',crop:'Swedes',ha:4.0,drillDate:'11/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'fb1',year:y,paddock:'Deershed Drive',crop:'Fodder Beet',ha:3.7,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'Tested'},
    {id:'fb2',year:y,paddock:'Two Tanks 2',crop:'Fodder Beet',ha:4.54,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'fb3',year:y,paddock:'Mckerchers',crop:'Fodder Beet',ha:7.05,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'lb1',year:y,paddock:'Sowbie 2',crop:'Lifting Beet',ha:3.84,drillDate:'21/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'Tested'},
    {id:'lb2',year:y,paddock:'Sowbie 1',crop:'Lifting Beet',ha:4.68,drillDate:'22/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'lb3',year:y,paddock:'Sowbie 3',crop:'Lifting Beet',ha:4.93,drillDate:'23/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'m1',year:y,paddock:'River 5',crop:'Maize',ha:4.5,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:''},
    {id:'m2',year:y,paddock:'Golden Willow 3',crop:'Maize',ha:5.0,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:'No test'},
    {id:'m3',year:y,paddock:'Golden Willow 4',crop:'Maize',ha:3.0,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:'No test'},
    {id:'t1',year:y,paddock:'One Tree',crop:'Triticale',ha:7.8,drillDate:'10/04/2026',yieldKgHA:6000,seedHA:250,chemHA:195,fertHA:163,opsHA:160,notes:''},
    {id:'r1',year:y,paddock:'Horsfall 1',crop:'Rape',ha:6.8,drillDate:'11/04/2026',yieldKgHA:8000,seedHA:195,chemHA:195,fertHA:163,opsHA:160,notes:''},
    {id:'r2',year:y,paddock:'Horsfall 4',crop:'Rape',ha:6.1,drillDate:'12/04/2026',yieldKgHA:8000,seedHA:195,chemHA:195,fertHA:163,opsHA:160,notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, rows);
  SpreadsheetApp.flush();
  Logger.log('CropPlan 26/27 done: ' + rows.length + ' rows');
}

function seedStockReq() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2026/27';
  const rows = [
    {id:'d1',year:y,species:'Deer',cls:'Barn Stags',headPrev:1250,headCurr:1250,kgDMday:4.75,period:'22 May - 8 Aug',days:75,feedSource:'Barn: sugar beet, maize, silage',notes:''},
    {id:'d2',year:y,species:'Deer',cls:'Ma / Old Stags',headPrev:650,headCurr:650,kgDMday:4.7,period:'June - July',days:75,feedSource:'Swedes',notes:''},
    {id:'d3',year:y,species:'Deer',cls:'Ma Old Stags (May)',headPrev:1900,headCurr:1900,kgDMday:5,period:'May',days:15,feedSource:'Saved pasture Italians',notes:'Aim for 80ha'},
    {id:'d4',year:y,species:'Deer',cls:'Ma Old Stags (Aug-Oct)',headPrev:1900,headCurr:1900,kgDMday:5,period:'Aug - 1 Oct',days:60,feedSource:'Saved pasture',notes:''},
    {id:'d5',year:y,species:'Deer',cls:'R3 Stags',headPrev:340,headCurr:340,kgDMday:4.2,period:'15 May - 1 Aug',days:75,feedSource:'Swedes',notes:''},
    {id:'d6',year:y,species:'Deer',cls:'R3 Stags (Aug)',headPrev:340,headCurr:340,kgDMday:3.96,period:'1 Aug - 15 Sep',days:45,feedSource:'Barn diet',notes:''},
    {id:'d7',year:y,species:'Deer',cls:'R3 Stags Elite',headPrev:315,headCurr:315,kgDMday:4.2,period:'20 May - 10 Aug',days:83,feedSource:'Kale, lifting beet, PK',notes:''},
    {id:'d8',year:y,species:'Deer',cls:'R3 Stags Elite (Aug)',headPrev:315,headCurr:315,kgDMday:4.2,period:'10 Aug - 15 Sep',days:35,feedSource:'Autumn saved pasture',notes:''},
    {id:'d9',year:y,species:'Deer',cls:'R3 B11 Stags + Ma',headPrev:65,headCurr:65,kgDMday:5.5,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'d10',year:y,species:'Deer',cls:'R2 Stags',headPrev:790,headCurr:790,kgDMday:3.6,period:'15 May - 1 Aug',days:78,feedSource:'Swedes',notes:''},
    {id:'d11',year:y,species:'Deer',cls:'R2 Stags Barn',headPrev:540,headCurr:540,kgDMday:3.4,period:'25 Jul - 25 Sep',days:62,feedSource:'Barn: sugar beet, maize, silage',notes:''},
    {id:'d12',year:y,species:'Deer',cls:'R2 B11 Stags',headPrev:73,headCurr:73,kgDMday:4,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'d13',year:y,species:'Deer',cls:'Forresters',headPrev:309,headCurr:309,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture Homeblock 107ha',notes:''},
    {id:'d14',year:y,species:'Deer',cls:'Velvet and Trophy',headPrev:457,headCurr:457,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Swedes',notes:''},
    {id:'d15',year:y,species:'Deer',cls:'Recips',headPrev:355,headCurr:355,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture 320ha',notes:''},
    {id:'d16',year:y,species:'Deer',cls:'Commercial Weaners',headPrev:900,headCurr:900,kgDMday:2.9,period:'10 May - 7 Aug',days:89,feedSource:'Fodder Beet',notes:''},
    {id:'d17',year:y,species:'Deer',cls:'Com Females (Aug)',headPrev:440,headCurr:440,kgDMday:2.9,period:'1 Aug - 15 Sep',days:45,feedSource:'Kale with some beet',notes:''},
    {id:'d18',year:y,species:'Deer',cls:'Com Males (Aug)',headPrev:450,headCurr:450,kgDMday:3.2,period:'1 Aug - 15 Sep',days:45,feedSource:'Pasture at gorge',notes:''},
    {id:'d19',year:y,species:'Deer',cls:'Ma Hinds (recips & studs)',headPrev:260,headCurr:260,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d20',year:y,species:'Deer',cls:'Com Hinds on Flats',headPrev:200,headCurr:200,kgDMday:2.6,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d21',year:y,species:'Deer',cls:'Com Hinds on Flats (after hill)',headPrev:800,headCurr:800,kgDMday:2.6,period:'15 Aug - 15 Sep',days:30,feedSource:'Kale',notes:''},
    {id:'d22',year:y,species:'Deer',cls:'B11 Hinds',headPrev:220,headCurr:220,kgDMday:4.7,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d23',year:y,species:'Deer',cls:'B11 R2 Hinds',headPrev:81,headCurr:81,kgDMday:4.4,period:'15 May - 15 Sep',days:120,feedSource:'River block pasture and rape',notes:''},
    {id:'d24',year:y,species:'Deer',cls:'B11 Weaners',headPrev:225,headCurr:225,kgDMday:3.5,period:'16 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'c1',year:y,species:'Cattle',cls:'R1 Heifers',headPrev:90,headCurr:90,kgDMday:7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture',notes:''},
    {id:'c2',year:y,species:'Cattle',cls:'Ma Cows on Flats',headPrev:80,headCurr:80,kgDMday:5,period:'15 May - 15 Sep',days:120,feedSource:'',notes:''},
    {id:'c3',year:y,species:'Cattle',cls:'R1 Finishing Cattle',headPrev:170,headCurr:170,kgDMday:8,period:'15 May - 15 Sep',days:120,feedSource:'',notes:''},
    {id:'c4',year:y,species:'Cattle',cls:'Calving Cows',headPrev:355,headCurr:355,kgDMday:7,period:'20 Aug - 15 Sep',days:27,feedSource:'Saved pasture',notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, rows);
  SpreadsheetApp.flush();
  Logger.log('StockReq 26/27 done: ' + rows.length + ' rows');
}

function seedSupplements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2026/27';
  const rows = [
    {id:'x1',year:y,name:'Lucerne / Clover Silage',type:'Lucerne Silage',kgDM:260000,costPerKgDM:0.18,notes:''},
    {id:'x2',year:y,name:'Bulk Grass Silage',type:'Silage',kgDM:750000,costPerKgDM:0.18,notes:''},
    {id:'x3',year:y,name:'Balage',type:'Balage',kgDM:324000,costPerKgDM:0.28,notes:''},
    {id:'x4',year:y,name:'Hay',type:'Hay',kgDM:50000,costPerKgDM:0.18,notes:''},
    {id:'x5',year:y,name:'Purchased Feed',type:'Supplement',kgDM:100000,costPerKgDM:0.33,notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS, rows);
  SpreadsheetApp.flush();
  Logger.log('Supplements 26/27 done: ' + rows.length + ' rows');
}

// ---------- 2025/26 ----------
// Crop paddocks from Cropping plan 26 Master.xlsx; costs estimated using 24/25 rates
// (no per-ha cost sheet found for 25/26). Update once confirmed.

function seedCropPlan2526() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2025/26';
  const rows = [
    {id:'26k1',year:y,paddock:'Thames Valley',crop:'Kale',ha:7.14,drillDate:'',yieldKgHA:12000,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:'Costs estimated from 24/25 rates — update when 25/26 invoices available'},
    {id:'26k2',year:y,paddock:'East Coast',crop:'Kale',ha:6.61,drillDate:'',yieldKgHA:12000,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:''},
    {id:'26k3',year:y,paddock:'Poverty Bay',crop:'Kale',ha:5.5,drillDate:'',yieldKgHA:12000,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:''},
    {id:'26k4',year:y,paddock:'Kingcountry',crop:'Kale',ha:5.0,drillDate:'',yieldKgHA:12000,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:''},
    {id:'26k5',year:y,paddock:'Bottom 60',crop:'Kale',ha:9.9,drillDate:'',yieldKgHA:12000,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:''},
    {id:'26k6',year:y,paddock:'Boots',crop:'Kale',ha:5.5,drillDate:'',yieldKgHA:12000,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:''},
    {id:'26s1',year:y,paddock:'North Harbour',crop:'Swedes',ha:5.2,drillDate:'',yieldKgHA:14000,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:'Costs estimated from 24/25 rates'},
    {id:'26s2',year:y,paddock:'Wanganui',crop:'Swedes',ha:5.47,drillDate:'',yieldKgHA:14000,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'26s3',year:y,paddock:'Bull',crop:'Swedes',ha:6.95,drillDate:'',yieldKgHA:14000,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'26s4',year:y,paddock:'Bens',crop:'Swedes',ha:5.17,drillDate:'',yieldKgHA:14000,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'26s5',year:y,paddock:'Waikato',crop:'Swedes',ha:11.0,drillDate:'',yieldKgHA:14000,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'26fb1',year:y,paddock:'Cattle Yards 2',crop:'Fodder Beet',ha:5.3,drillDate:'',yieldKgHA:20000,seedHA:400,chemHA:1200,fertHA:650,opsHA:490,notes:'Costs from 24/25 rates'},
    {id:'26fb2',year:y,paddock:'Cattle Yards 4',crop:'Fodder Beet',ha:6.3,drillDate:'',yieldKgHA:20000,seedHA:400,chemHA:1200,fertHA:650,opsHA:490,notes:''},
    {id:'26fb3',year:y,paddock:'Woodings',crop:'Fodder Beet',ha:6.88,drillDate:'',yieldKgHA:20000,seedHA:400,chemHA:1200,fertHA:650,opsHA:490,notes:''},
    {id:'26lb1',year:y,paddock:'One Tree',crop:'Lifting Beet',ha:7.0,drillDate:'',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'Using 26/27 rates — no 25/26 cost data found'},
    {id:'26m1',year:y,paddock:'Horsfall 1',crop:'Maize',ha:6.8,drillDate:'',yieldKgHA:18000,seedHA:500,chemHA:800,fertHA:800,opsHA:0,notes:'Costs from 24/25 rates (ops not itemised)'},
    {id:'26m2',year:y,paddock:'Horsfall 4',crop:'Maize',ha:6.0,drillDate:'',yieldKgHA:18000,seedHA:500,chemHA:800,fertHA:800,opsHA:0,notes:''},
  ];
  appendRows(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, rows);
  Logger.log('CropPlan 25/26 done: ' + rows.length + ' rows');
}

function seedStockReq2526() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2025/26';
  // Stock numbers estimated from 26/27 plan — no separate 25/26 stock file in source.
  // Update headCurr and feedSource fields to match actual 25/26 records.
  const rows = [
    {id:'26sd1',year:y,species:'Deer',cls:'Barn Stags',headPrev:1250,headCurr:1250,kgDMday:4.75,period:'22 May - 8 Aug',days:75,feedSource:'Barn: sugar beet, maize, silage',notes:'Est from 26/27 — update with actuals'},
    {id:'26sd2',year:y,species:'Deer',cls:'Ma / Old Stags',headPrev:650,headCurr:650,kgDMday:4.7,period:'June - July',days:75,feedSource:'Swedes',notes:'Est from 26/27'},
    {id:'26sd3',year:y,species:'Deer',cls:'Ma Old Stags (May)',headPrev:1900,headCurr:1900,kgDMday:5,period:'May',days:15,feedSource:'Saved pasture',notes:'Est from 26/27'},
    {id:'26sd4',year:y,species:'Deer',cls:'Ma Old Stags (Aug-Oct)',headPrev:1900,headCurr:1900,kgDMday:5,period:'Aug - 1 Oct',days:60,feedSource:'Saved pasture',notes:'Est from 26/27'},
    {id:'26sd5',year:y,species:'Deer',cls:'R3 Stags',headPrev:340,headCurr:340,kgDMday:4.2,period:'15 May - 1 Aug',days:75,feedSource:'Swedes',notes:'Est from 26/27'},
    {id:'26sd6',year:y,species:'Deer',cls:'R3 Stags (Aug)',headPrev:340,headCurr:340,kgDMday:3.96,period:'1 Aug - 15 Sep',days:45,feedSource:'Barn diet',notes:'Est from 26/27'},
    {id:'26sd7',year:y,species:'Deer',cls:'R3 Stags Elite',headPrev:315,headCurr:315,kgDMday:4.2,period:'20 May - 10 Aug',days:83,feedSource:'Kale, lifting beet',notes:'Est from 26/27'},
    {id:'26sd8',year:y,species:'Deer',cls:'R3 Stags Elite (Aug)',headPrev:315,headCurr:315,kgDMday:4.2,period:'10 Aug - 15 Sep',days:35,feedSource:'Autumn pasture',notes:'Est from 26/27'},
    {id:'26sd9',year:y,species:'Deer',cls:'R3 B11 Stags + Ma',headPrev:65,headCurr:65,kgDMday:5.5,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:'Est from 26/27'},
    {id:'26sd10',year:y,species:'Deer',cls:'R2 Stags',headPrev:790,headCurr:790,kgDMday:3.6,period:'15 May - 1 Aug',days:78,feedSource:'Swedes',notes:'Est from 26/27'},
    {id:'26sd11',year:y,species:'Deer',cls:'R2 Stags Barn',headPrev:540,headCurr:540,kgDMday:3.4,period:'25 Jul - 25 Sep',days:62,feedSource:'Barn: sugar beet, maize, silage',notes:'Est from 26/27'},
    {id:'26sd12',year:y,species:'Deer',cls:'R2 B11 Stags',headPrev:73,headCurr:73,kgDMday:4,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:'Est from 26/27'},
    {id:'26sd13',year:y,species:'Deer',cls:'Forresters',headPrev:309,headCurr:309,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture Homeblock',notes:'Est from 26/27'},
    {id:'26sd14',year:y,species:'Deer',cls:'Velvet and Trophy',headPrev:457,headCurr:457,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Swedes',notes:'Est from 26/27'},
    {id:'26sd15',year:y,species:'Deer',cls:'Recips',headPrev:355,headCurr:355,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture 320ha',notes:'Est from 26/27'},
    {id:'26sd16',year:y,species:'Deer',cls:'Commercial Weaners',headPrev:900,headCurr:900,kgDMday:2.9,period:'10 May - 7 Aug',days:89,feedSource:'Fodder Beet',notes:'Est from 26/27'},
    {id:'26sd17',year:y,species:'Deer',cls:'Com Females (Aug)',headPrev:440,headCurr:440,kgDMday:2.9,period:'1 Aug - 15 Sep',days:45,feedSource:'Kale with some beet',notes:'Est from 26/27'},
    {id:'26sd18',year:y,species:'Deer',cls:'Com Males (Aug)',headPrev:450,headCurr:450,kgDMday:3.2,period:'1 Aug - 15 Sep',days:45,feedSource:'Pasture at gorge',notes:'Est from 26/27'},
    {id:'26sd19',year:y,species:'Deer',cls:'Ma Hinds (recips & studs)',headPrev:260,headCurr:260,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:'Est from 26/27'},
    {id:'26sd20',year:y,species:'Deer',cls:'Com Hinds on Flats',headPrev:200,headCurr:200,kgDMday:2.6,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:'Est from 26/27'},
    {id:'26sd21',year:y,species:'Deer',cls:'Com Hinds on Flats (after hill)',headPrev:800,headCurr:800,kgDMday:2.6,period:'15 Aug - 15 Sep',days:30,feedSource:'Kale',notes:'Est from 26/27'},
    {id:'26sd22',year:y,species:'Deer',cls:'B11 Hinds',headPrev:220,headCurr:220,kgDMday:4.7,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:'Est from 26/27'},
    {id:'26sd23',year:y,species:'Deer',cls:'B11 R2 Hinds',headPrev:81,headCurr:81,kgDMday:4.4,period:'15 May - 15 Sep',days:120,feedSource:'Rape',notes:'Est from 26/27'},
    {id:'26sd24',year:y,species:'Deer',cls:'B11 Weaners',headPrev:225,headCurr:225,kgDMday:3.5,period:'16 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:'Est from 26/27'},
    {id:'26sc1',year:y,species:'Cattle',cls:'R1 Heifers',headPrev:90,headCurr:90,kgDMday:7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture',notes:'Est from 26/27'},
    {id:'26sc2',year:y,species:'Cattle',cls:'Ma Cows on Flats',headPrev:80,headCurr:80,kgDMday:5,period:'15 May - 15 Sep',days:120,feedSource:'',notes:'Est from 26/27'},
    {id:'26sc3',year:y,species:'Cattle',cls:'R1 Finishing Cattle',headPrev:170,headCurr:170,kgDMday:8,period:'15 May - 15 Sep',days:120,feedSource:'',notes:'Est from 26/27'},
    {id:'26sc4',year:y,species:'Cattle',cls:'Calving Cows',headPrev:355,headCurr:355,kgDMday:7,period:'20 Aug - 15 Sep',days:27,feedSource:'Saved pasture',notes:'Est from 26/27'},
  ];
  appendRows(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, rows);
  Logger.log('StockReq 25/26 done: ' + rows.length + ' rows');
}

function seedSupps2526() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2025/26';
  const rows = [
    {id:'26x1',year:y,name:'Lucerne / Clover Silage',type:'Lucerne Silage',kgDM:260000,costPerKgDM:0.18,notes:'Estimated — update with actuals'},
    {id:'26x2',year:y,name:'Bulk Grass Silage',type:'Silage',kgDM:700000,costPerKgDM:0.18,notes:'Estimated'},
    {id:'26x3',year:y,name:'Balage',type:'Balage',kgDM:300000,costPerKgDM:0.28,notes:'Estimated'},
    {id:'26x4',year:y,name:'Hay',type:'Hay',kgDM:50000,costPerKgDM:0.18,notes:'Estimated'},
  ];
  appendRows(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS, rows);
  Logger.log('Supplements 25/26 done: ' + rows.length + ' rows');
}

// ---------- 2024/25 ----------
// Source: Peel Forest Estate Crop Costs 2024-25.xlsx + Winter Budget master.xlsx
// Costs confirmed from source. Yields from Winter Budget master.

function seedCropPlan2425() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2024/25';
  const rows = [
    {id:'25s1',year:y,paddock:'Thames Valley',crop:'Swedes',ha:7.14,drillDate:'',yieldKgHA:14600,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:'Source: Crop Costs 2024-25. Yield units unconfirmed DM vs fresh.'},
    {id:'25s2',year:y,paddock:'East Coast',crop:'Swedes',ha:6.61,drillDate:'',yieldKgHA:14600,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'25s3',year:y,paddock:'Poverty Bay',crop:'Swedes',ha:5.5,drillDate:'',yieldKgHA:14600,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'25s4',year:y,paddock:'Bottom 60',crop:'Swedes',ha:10.42,drillDate:'',yieldKgHA:14600,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:''},
    {id:'25s5',year:y,paddock:'Kingcountry',crop:'Swedes',ha:5.0,drillDate:'',yieldKgHA:14600,seedHA:195,chemHA:500,fertHA:450,opsHA:280,notes:'Season total 455,563.8 kgDM usable'},
    {id:'25k1',year:y,paddock:'North Otago',crop:'Kale',ha:6.68,drillDate:'',yieldKgHA:12600,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:'Source: Crop Costs 2024-25. Total yield 174,636 kgDM'},
    {id:'25k2',year:y,paddock:'Nelson Bays',crop:'Kale',ha:8.72,drillDate:'',yieldKgHA:12600,seedHA:150,chemHA:500,fertHA:450,opsHA:354,notes:''},
    {id:'25fb1',year:y,paddock:'Bottom Stump',crop:'Fodder Beet',ha:4.86,drillDate:'',yieldKgHA:19000,seedHA:400,chemHA:1200,fertHA:650,opsHA:490,notes:'Total 161,766 kgDM. Seed $400, Chem $1200, Fert $650, Ops $490 per ha.'},
    {id:'25fb2',year:y,paddock:'Top Johns',crop:'Fodder Beet',ha:4.6,drillDate:'',yieldKgHA:19000,seedHA:400,chemHA:1200,fertHA:650,opsHA:490,notes:''},
    {id:'25m1',year:y,paddock:'Bottom Oaks',crop:'Maize',ha:3.84,drillDate:'',yieldKgHA:19000,seedHA:500,chemHA:800,fertHA:800,opsHA:0,notes:'Total 295,830 kgDM. Ops not itemised in source.'},
    {id:'25m2',year:y,paddock:'Oaks 3',crop:'Maize',ha:6.32,drillDate:'',yieldKgHA:19000,seedHA:500,chemHA:800,fertHA:800,opsHA:0,notes:''},
    {id:'25m3',year:y,paddock:'Oaks 4',crop:'Maize',ha:5.41,drillDate:'',yieldKgHA:19000,seedHA:500,chemHA:800,fertHA:800,opsHA:0,notes:''},
  ];
  appendRows(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, rows);
  Logger.log('CropPlan 24/25 done: ' + rows.length + ' rows');
}

function seedStockReq2425() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2024/25';
  // From Winter Budget master.xlsx stock allocation. Period 1 Jun - 1 Oct, 120 days where not stated.
  const rows = [
    {id:'25d1',year:y,species:'Deer',cls:'Barn Stags',headPrev:1180,headCurr:1180,kgDMday:4.5,period:'Jun - Jul',days:60,feedSource:'Barn: sugar beet, maize, silage',notes:'Source: Winter Budget master 2025'},
    {id:'25d2',year:y,species:'Deer',cls:'Ma / Old Stags (Jun-Oct)',headPrev:1115,headCurr:1115,kgDMday:4.7,period:'Jun - Oct',days:120,feedSource:'Swedes',notes:''},
    {id:'25d3',year:y,species:'Deer',cls:'Ma Old Stags (May)',headPrev:2295,headCurr:2295,kgDMday:5,period:'May',days:15,feedSource:'Saved pasture',notes:''},
    {id:'25d4',year:y,species:'Deer',cls:'Ma Old Stags (Aug-Oct)',headPrev:2295,headCurr:2295,kgDMday:5,period:'Aug - Oct',days:60,feedSource:'Saved pasture',notes:''},
    {id:'25d5',year:y,species:'Deer',cls:'Sires',headPrev:45,headCurr:45,kgDMday:4.7,period:'Jun - Oct',days:120,feedSource:'',notes:''},
    {id:'25d6',year:y,species:'Deer',cls:'R3 Stags',headPrev:458,headCurr:458,kgDMday:4.2,period:'15 May - 1 Aug',days:75,feedSource:'Swedes',notes:''},
    {id:'25d7',year:y,species:'Deer',cls:'R3 Stags Elite',headPrev:400,headCurr:400,kgDMday:4.2,period:'20 May - 10 Aug',days:83,feedSource:'Kale, lifting beet',notes:''},
    {id:'25d8',year:y,species:'Deer',cls:'R3 B11 Stags',headPrev:59,headCurr:59,kgDMday:5.5,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'25d9',year:y,species:'Deer',cls:'R2 Stags',headPrev:918,headCurr:918,kgDMday:3.5,period:'15 May - 1 Aug',days:78,feedSource:'Swedes',notes:''},
    {id:'25d10',year:y,species:'Deer',cls:'R2 B11 Stags',headPrev:73,headCurr:73,kgDMday:4,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'25d11',year:y,species:'Deer',cls:'R1 Weaners',headPrev:1080,headCurr:1080,kgDMday:2.3,period:'Jun - Oct',days:120,feedSource:'Winterstar + balage',notes:''},
    {id:'25d12',year:y,species:'Deer',cls:'Commercial Weaners',headPrev:776,headCurr:776,kgDMday:2.5,period:'Jun - Oct',days:120,feedSource:'Fodder Beet + lucerne silage',notes:''},
    {id:'25d13',year:y,species:'Deer',cls:'MA Hinds (silage then kale)',headPrev:516,headCurr:516,kgDMday:2.65,period:'Jun - Oct',days:120,feedSource:'Silage then kale',notes:'2.5 kgDM then 2.8 kgDM'},
    {id:'25d14',year:y,species:'Deer',cls:'Milking Hinds',headPrev:200,headCurr:200,kgDMday:2.6,period:'Jun - Oct',days:120,feedSource:'',notes:''},
    {id:'25d15',year:y,species:'Deer',cls:'Com Hinds (spring off hill)',headPrev:750,headCurr:500,kgDMday:2.6,period:'Various',days:90,feedSource:'',notes:'Head count varies across periods — use 500 as working figure'},
    {id:'25d16',year:y,species:'Deer',cls:'B11 Hinds',headPrev:177,headCurr:177,kgDMday:4.7,period:'Jun - Oct',days:120,feedSource:'',notes:''},
    {id:'25d17',year:y,species:'Deer',cls:'B11 R2 Hinds',headPrev:81,headCurr:81,kgDMday:4.4,period:'Jun - Oct',days:120,feedSource:'',notes:''},
    {id:'25c1',year:y,species:'Cattle',cls:'R1 Heifers',headPrev:83,headCurr:83,kgDMday:7,period:'Jun - Oct',days:120,feedSource:'Pasture',notes:''},
    {id:'25c2',year:y,species:'Cattle',cls:'Ma Cows on Flats',headPrev:100,headCurr:100,kgDMday:5,period:'Jun - Oct',days:120,feedSource:'',notes:'Head count conflict in source (679 also appears) — 100 used'},
    {id:'25c3',year:y,species:'Cattle',cls:'R1 Finishing Cattle',headPrev:270,headCurr:270,kgDMday:8,period:'Jun - Oct',days:120,feedSource:'',notes:''},
    {id:'25c4',year:y,species:'Cattle',cls:'Calving Cows',headPrev:400,headCurr:400,kgDMday:7,period:'Aug - Sep',days:30,feedSource:'Saved pasture',notes:'No rate given in source — 7 kgDM estimated'},
  ];
  appendRows(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, rows);
  Logger.log('StockReq 24/25 done: ' + rows.length + ' rows');
}

function seedSupps2425() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2024/25';
  // From Winter Budget master 2025 supply figures
  const rows = [
    {id:'25x1',year:y,name:'Lucerne Silage',type:'Lucerne Silage',kgDM:300000,costPerKgDM:0.18,notes:'Source: Winter Budget master 2025'},
    {id:'25x2',year:y,name:'Bulk Grass Silage',type:'Silage',kgDM:700000,costPerKgDM:0.18,notes:'Old silage included in this figure'},
    {id:'25x3',year:y,name:'Balage',type:'Balage',kgDM:280000,costPerKgDM:0.28,notes:''},
    {id:'25x4',year:y,name:'Hay',type:'Hay',kgDM:50000,costPerKgDM:0.18,notes:''},
    {id:'25x5',year:y,name:'Purchased Feed',type:'Supplement',kgDM:100000,costPerKgDM:0.33,notes:'Season total supply 4,239,641 kgDM; surplus +322,017 kgDM'},
  ];
  appendRows(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS, rows);
  Logger.log('Supplements 24/25 done: ' + rows.length + ' rows');
}

// ---------- 2023/24 ----------
// Source: Winter feed budget 2024 master.xlsx
// NOTE: yieldKgHA values here are FRESH WEIGHT from source, not DM.
// DM content: Beet ~14%, Swedes ~10%, Kale ~14%, Rape ~12% typical.
// Update yields to kgDM/ha when confirmed. Costs not found in source files for this season.

function seedCropPlan2324() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2023/24';
  const rows = [
    {id:'24fb1',year:y,paddock:'Pump 1',crop:'Fodder Beet',ha:3.63,drillDate:'',yieldKgHA:20600,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT yield (20,600 kgFW/ha). Convert to kgDM: multiply by ~0.14. No cost data in source.'},
    {id:'24fb2',year:y,paddock:'Pump 2',crop:'Fodder Beet',ha:4.06,drillDate:'',yieldKgHA:24400,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT (24,400 kgFW/ha)'},
    {id:'24fb3',year:y,paddock:'Andys',crop:'Fodder Beet',ha:3.74,drillDate:'',yieldKgHA:20600,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT (20,600 kgFW/ha)'},
    {id:'24s1',year:y,paddock:'Two Tanks 1',crop:'Swedes',ha:4.67,drillDate:'',yieldKgHA:17500,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT (17,500 kgFW/ha). No cost data.'},
    {id:'24s2',year:y,paddock:'Top Stump',crop:'Swedes',ha:4.8,drillDate:'',yieldKgHA:17000,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24s3',year:y,paddock:'Gum 1',crop:'Swedes',ha:4.4,drillDate:'',yieldKgHA:17500,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24s4',year:y,paddock:'Mothering Up',crop:'Swedes',ha:5.68,drillDate:'',yieldKgHA:15700,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24s5',year:y,paddock:'Bondis',crop:'Swedes',ha:3.5,drillDate:'',yieldKgHA:15700,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24s6',year:y,paddock:'North Otago',crop:'Swedes',ha:6.6,drillDate:'',yieldKgHA:14600,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24s7',year:y,paddock:'Nelson Bays',crop:'Swedes',ha:8.6,drillDate:'',yieldKgHA:14600,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24k1',year:y,paddock:'Big Bills',crop:'Kale',ha:7.16,drillDate:'',yieldKgHA:6350,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT (6,350 kgFW/ha — low, check). No cost data.'},
    {id:'24k2',year:y,paddock:'Robs',crop:'Kale',ha:3.59,drillDate:'',yieldKgHA:6350,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'FRESH WEIGHT'},
    {id:'24r1',year:y,paddock:'Westcoast/Marlborough/Johnsons/Pigsty',crop:'Rape',ha:22.54,drillDate:'',yieldKgHA:5170,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'Winterstar. 4 paddocks combined — per-paddock ha not in source. FRESH WEIGHT (5,170 kgFW/ha avg).'},
    {id:'24r2',year:y,paddock:'Oregons',crop:'Rape',ha:7.85,drillDate:'',yieldKgHA:3270,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'Turnips. FRESH WEIGHT.'},
    {id:'24r3',year:y,paddock:'Golden Willows 5',crop:'Rape',ha:4.9,drillDate:'',yieldKgHA:3270,seedHA:0,chemHA:0,fertHA:0,opsHA:0,notes:'Turnips. FRESH WEIGHT.'},
  ];
  appendRows(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, rows);
  Logger.log('CropPlan 23/24 done: ' + rows.length + ' rows');
}

function seedStockReq2324() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2023/24';
  // From summary in Winter feed budget 2024 master.xlsx. Period 1 Jun - 1 Oct, 120 days.
  // kgDMday rates extracted from source; head counts from stock requirements summary.
  const rows = [
    {id:'24d1',year:y,species:'Deer',cls:'MA / Old Stags',headPrev:2194,headCurr:2194,kgDMday:4.7,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:'Source: Winter feed budget 2024 master'},
    {id:'24d2',year:y,species:'Deer',cls:'R3 Stags',headPrev:828,headCurr:828,kgDMday:4.2,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
    {id:'24d3',year:y,species:'Deer',cls:'R3 B11 Stags',headPrev:59,headCurr:59,kgDMday:5.5,period:'1 Jun - 1 Oct',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'24d4',year:y,species:'Deer',cls:'R2 Stags',headPrev:984,headCurr:984,kgDMday:3.5,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
    {id:'24d5',year:y,species:'Deer',cls:'R2 B11 Stags',headPrev:73,headCurr:73,kgDMday:4,period:'1 Jun - 1 Oct',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'24d6',year:y,species:'Deer',cls:'R1 Weaners',headPrev:2157,headCurr:2157,kgDMday:2.3,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
    {id:'24d7',year:y,species:'Deer',cls:'MA Hinds',headPrev:516,headCurr:516,kgDMday:3.5,period:'1 Jun - 1 Oct',days:120,feedSource:'Kale',notes:''},
    {id:'24d8',year:y,species:'Deer',cls:'R2 Com Hinds',headPrev:219,headCurr:219,kgDMday:2.6,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:'Rate not in source — 2.6 estimated'},
    {id:'24d9',year:y,species:'Deer',cls:'MA Com Hinds',headPrev:260,headCurr:260,kgDMday:2.6,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:'Rate not in source — 2.6 estimated'},
    {id:'24d10',year:y,species:'Deer',cls:'Com Hinds',headPrev:450,headCurr:450,kgDMday:2.6,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:'Rate not in source — 2.6 estimated'},
    {id:'24d11',year:y,species:'Deer',cls:'B11 Hinds',headPrev:177,headCurr:177,kgDMday:4.7,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
    {id:'24d12',year:y,species:'Deer',cls:'B11 R2 Hinds',headPrev:81,headCurr:81,kgDMday:4.4,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
    {id:'24d13',year:y,species:'Deer',cls:'MA Ewes',headPrev:93,headCurr:93,kgDMday:1.5,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
    {id:'24c1',year:y,species:'Cattle',cls:'R1 Heifers',headPrev:83,headCurr:83,kgDMday:7,period:'1 Jun - 1 Oct',days:120,feedSource:'Pasture',notes:''},
    {id:'24c2',year:y,species:'Cattle',cls:'MA Cows',headPrev:140,headCurr:140,kgDMday:5,period:'1 Jun - 1 Oct',days:120,feedSource:'',notes:''},
  ];
  appendRows(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, rows);
  Logger.log('StockReq 23/24 done: ' + rows.length + ' rows');
}

// ---------- 2022/23 ----------
// Source: Copy of cropping plan 2023 (version 1).xlsx
// No yield data or stock requirements found for this season — crop area and costs only.

function seedCropPlan2223() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2022/23';
  const rows = [
    {id:'23fb1',year:y,paddock:'Pump 1',crop:'Fodder Beet',ha:3.63,drillDate:'',yieldKgHA:0,seedHA:400,chemHA:1200,fertHA:650,opsHA:660,notes:'Source: Cropping plan 2023 v1. Group total 11.37ha $22,523.40. No yield data in source.'},
    {id:'23fb2',year:y,paddock:'Andys',crop:'Fodder Beet',ha:3.74,drillDate:'',yieldKgHA:0,seedHA:400,chemHA:1200,fertHA:650,opsHA:660,notes:''},
    {id:'23fb3',year:y,paddock:'Pump 2',crop:'Fodder Beet',ha:4.0,drillDate:'',yieldKgHA:0,seedHA:400,chemHA:1200,fertHA:650,opsHA:660,notes:''},
    {id:'23k1',year:y,paddock:'Plantation',crop:'Kale',ha:10.67,drillDate:'',yieldKgHA:0,seedHA:120,chemHA:500,fertHA:500,opsHA:0,notes:'Part of Kale+Swedes group 48.92ha $83,223.75. Ops not itemised.'},
    {id:'23s1',year:y,paddock:'Two Tanks 1 / Top Stump / Gum 1 / Mothering Up / Bondis / North Otago / Nelson Bays',crop:'Swedes',ha:38.25,drillDate:'',yieldKgHA:0,seedHA:235,chemHA:500,fertHA:500,opsHA:0,notes:'7 paddocks combined — per-paddock ha not in source. Ops not itemised. No yield or stock data in source.'},
    {id:'23m1',year:y,paddock:'Oaks 2 / Top Walnut / Charlies / Pump 3',crop:'Maize',ha:18.19,drillDate:'',yieldKgHA:0,seedHA:500,chemHA:800,fertHA:800,opsHA:0,notes:'4 paddocks combined — per-paddock ha not in source. Total $49,879.80. No yield data.'},
  ];
  appendRows(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, rows);
  Logger.log('CropPlan 22/23 done: ' + rows.length + ' rows');
}

// ---------- 2021/22 ----------
// Source: Cropping plan 22.xlsx, CROP YIELDS 22.xlsx, feed requirements 22.xlsx
// Per-paddock ha available for beet only. Swedes/Kale aggregated. No stock-req breakdown.

function seedCropPlan2122() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const y = '2021/22';
  const rows = [
    {id:'22fb1',year:y,paddock:'Oaks 2',crop:'Fodder Beet',ha:5.41,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:'Yield = season avg across 12 beet paddocks (source did not break out per-paddock). Group total 57.61ha $155,899.68. Ops not itemised.'},
    {id:'22fb2',year:y,paddock:'Horsfall 2',crop:'Fodder Beet',ha:4.24,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb3',year:y,paddock:'Pump 1',crop:'Fodder Beet',ha:3.63,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb4',year:y,paddock:'Pump 3',crop:'Fodder Beet',ha:3.86,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb5',year:y,paddock:'Top Walnut',crop:'Fodder Beet',ha:4.95,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb6',year:y,paddock:'Two Tanks 1',crop:'Fodder Beet',ha:4.67,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb7',year:y,paddock:'Gum 2',crop:'Fodder Beet',ha:4.5,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb8',year:y,paddock:'Top Stump',crop:'Fodder Beet',ha:4.75,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb9',year:y,paddock:'Georges 1',crop:'Fodder Beet',ha:4.32,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb10',year:y,paddock:'Sowbie 1',crop:'Fodder Beet',ha:4.68,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb11',year:y,paddock:'Left Deans',crop:'Fodder Beet',ha:4.0,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:''},
    {id:'22fb12',year:y,paddock:'Nelson Bays',crop:'Fodder Beet',ha:8.6,drillDate:'',yieldKgHA:22570,seedHA:380,chemHA:1130,fertHA:600,opsHA:0,notes:'12th beet paddock — largest.'},
    {id:'22s1',year:y,paddock:'Bondis/Mothering Up/CY1/Oaks1/Oscars/Caldwell Yds/Boots/North Otago',crop:'Swedes',ha:36.17,drillDate:'',yieldKgHA:17906,seedHA:225,chemHA:425,fertHA:480,opsHA:0,notes:'8 paddocks combined — per-paddock ha not in source. Yield avg of 5 matched paddocks. Total $98,187.48. Oaks 1 cost anomaly in source (flagged).'},
    {id:'22k1',year:y,paddock:'Tonys / Big Bill / Robs',crop:'Kale',ha:10.62,drillDate:'',yieldKgHA:0,seedHA:150,chemHA:621,fertHA:480,opsHA:0,notes:'3 paddocks combined — per-paddock ha not in source. Total $16,269.84. No yield figure in source.'},
    {id:'22m1',year:y,paddock:'Bottom 60',crop:'Maize',ha:10.42,drillDate:'',yieldKgHA:0,seedHA:1050,chemHA:345,fertHA:700,opsHA:0,notes:'Total $26,623.10. No yield figure in source.'},
  ];
  appendRows(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, rows);
  Logger.log('CropPlan 21/22 done: ' + rows.length + ' rows');
}

// ---------- Legacy / Deprecated ----------

function seedHistory() {
  Logger.log('seedHistory is deprecated. Historical data is now seeded into CropPlan/StockReq with year tags. Run seedCropPlan2526, seedStockReq2526, etc. instead.');
}

function seedData() {
  Logger.log('seedData is deprecated — timed out at 6 min. Run seedCropPlan, seedStockReq, seedSupplements individually, then the historical append functions.');
}

function respond(json, e) {
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
