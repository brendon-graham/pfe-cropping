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

// Run each seed function separately from the Apps Script editor (one at a time)
// 1. seedCropPlan  2. seedStockReq  3. seedSupplements  4. seedHistory
// Then check your Sheet — all data should be there.

function seedCropPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cropPlan = [
    {id:'k1',paddock:'North Harbour',crop:'Kale',ha:5.2,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k2',paddock:'Wanganui',crop:'Kale',ha:5.47,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k3',paddock:'Bull',crop:'Kale',ha:6.95,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k4',paddock:'Bens',crop:'Kale',ha:5.17,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k5',paddock:'Waikato',crop:'Kale',ha:11,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'s1',paddock:'River 1',crop:'Swedes',ha:7.9,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s2',paddock:'River 2',crop:'Swedes',ha:5.17,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s3',paddock:'River 3',crop:'Swedes',ha:5.14,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s4',paddock:'River 4',crop:'Swedes',ha:3.3,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s5',paddock:'Horsfall 3',crop:'Swedes',ha:5.6,drillDate:'11/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s6',paddock:'Ryans',crop:'Swedes',ha:4.7,drillDate:'12/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s7',paddock:'Horsfall 2',crop:'Swedes',ha:4.2,drillDate:'13/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s8',paddock:'Yards',crop:'Swedes',ha:5.03,drillDate:'14/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s9',paddock:'Cattle Yards 3',crop:'Swedes',ha:4.53,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s10',paddock:'Little Horsfall Swamp',crop:'Swedes',ha:4.0,drillDate:'11/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'fb1',paddock:'Deershed Drive',crop:'Fodder Beet',ha:3.7,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'Tested'},
    {id:'fb2',paddock:'Two Tanks 2',crop:'Fodder Beet',ha:4.54,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'fb3',paddock:'Mckerchers',crop:'Fodder Beet',ha:7.05,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'lb1',paddock:'Sowbie 2',crop:'Lifting Beet',ha:3.84,drillDate:'21/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'Tested'},
    {id:'lb2',paddock:'Sowbie 1',crop:'Lifting Beet',ha:4.68,drillDate:'22/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'lb3',paddock:'Sowbie 3',crop:'Lifting Beet',ha:4.93,drillDate:'23/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'m1',paddock:'River 5',crop:'Maize',ha:4.5,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:''},
    {id:'m2',paddock:'Golden Willow 3',crop:'Maize',ha:5.0,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:'No test'},
    {id:'m3',paddock:'Golden Willow 4',crop:'Maize',ha:3.0,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:'No test'},
    {id:'t1',paddock:'One Tree',crop:'Triticale',ha:7.8,drillDate:'10/04/2026',yieldKgHA:6000,seedHA:250,chemHA:195,fertHA:163,opsHA:160,notes:''},
    {id:'r1',paddock:'Horsfall 1',crop:'Rape',ha:6.8,drillDate:'11/04/2026',yieldKgHA:8000,seedHA:195,chemHA:195,fertHA:163,opsHA:160,notes:''},
    {id:'r2',paddock:'Horsfall 4',crop:'Rape',ha:6.1,drillDate:'12/04/2026',yieldKgHA:8000,seedHA:195,chemHA:195,fertHA:163,opsHA:160,notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, cropPlan);
  SpreadsheetApp.flush();
  Logger.log('CropPlan done: ' + cropPlan.length + ' rows');
}

function seedStockReq() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockReq = [
    {id:'d1',species:'Deer',cls:'Barn Stags',headPrev:1250,headCurr:1250,kgDMday:4.75,period:'22 May - 8 Aug',days:75,feedSource:'Barn: sugar beet, maize, silage',notes:''},
    {id:'d2',species:'Deer',cls:'Ma / Old Stags',headPrev:650,headCurr:650,kgDMday:4.7,period:'June - July',days:75,feedSource:'Swedes',notes:''},
    {id:'d3',species:'Deer',cls:'Ma Old Stags (May)',headPrev:1900,headCurr:1900,kgDMday:5,period:'May',days:15,feedSource:'Saved pasture Italians',notes:'Aim for 80ha'},
    {id:'d4',species:'Deer',cls:'Ma Old Stags (Aug-Oct)',headPrev:1900,headCurr:1900,kgDMday:5,period:'Aug - 1 Oct',days:60,feedSource:'Saved pasture',notes:''},
    {id:'d5',species:'Deer',cls:'R3 Stags',headPrev:340,headCurr:340,kgDMday:4.2,period:'15 May - 1 Aug',days:75,feedSource:'Swedes',notes:''},
    {id:'d6',species:'Deer',cls:'R3 Stags (Aug)',headPrev:340,headCurr:340,kgDMday:3.96,period:'1 Aug - 15 Sep',days:45,feedSource:'Barn diet',notes:''},
    {id:'d7',species:'Deer',cls:'R3 Stags Elite',headPrev:315,headCurr:315,kgDMday:4.2,period:'20 May - 10 Aug',days:83,feedSource:'Kale, lifting beet, PK',notes:''},
    {id:'d8',species:'Deer',cls:'R3 Stags Elite (Aug)',headPrev:315,headCurr:315,kgDMday:4.2,period:'10 Aug - 15 Sep',days:35,feedSource:'Autumn saved pasture',notes:''},
    {id:'d9',species:'Deer',cls:'R3 B11 Stags + Ma',headPrev:65,headCurr:65,kgDMday:5.5,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'d10',species:'Deer',cls:'R2 Stags',headPrev:790,headCurr:790,kgDMday:3.6,period:'15 May - 1 Aug',days:78,feedSource:'Swedes',notes:''},
    {id:'d11',species:'Deer',cls:'R2 Stags Barn',headPrev:540,headCurr:540,kgDMday:3.4,period:'25 Jul - 25 Sep',days:62,feedSource:'Barn: sugar beet, maize, silage',notes:''},
    {id:'d12',species:'Deer',cls:'R2 B11 Stags',headPrev:73,headCurr:73,kgDMday:4,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'d13',species:'Deer',cls:'Forresters',headPrev:309,headCurr:309,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture Homeblock 107ha',notes:''},
    {id:'d14',species:'Deer',cls:'Velvet and Trophy',headPrev:457,headCurr:457,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Swedes',notes:''},
    {id:'d15',species:'Deer',cls:'Recips',headPrev:355,headCurr:355,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture 320ha',notes:''},
    {id:'d16',species:'Deer',cls:'Commercial Weaners',headPrev:900,headCurr:900,kgDMday:2.9,period:'10 May - 7 Aug',days:89,feedSource:'Fodder Beet',notes:''},
    {id:'d17',species:'Deer',cls:'Com Females (Aug)',headPrev:440,headCurr:440,kgDMday:2.9,period:'1 Aug - 15 Sep',days:45,feedSource:'Kale with some beet',notes:''},
    {id:'d18',species:'Deer',cls:'Com Males (Aug)',headPrev:450,headCurr:450,kgDMday:3.2,period:'1 Aug - 15 Sep',days:45,feedSource:'Pasture at gorge',notes:''},
    {id:'d19',species:'Deer',cls:'Ma Hinds (recips & studs)',headPrev:260,headCurr:260,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d20',species:'Deer',cls:'Com Hinds on Flats',headPrev:200,headCurr:200,kgDMday:2.6,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d21',species:'Deer',cls:'Com Hinds on Flats (after hill)',headPrev:800,headCurr:800,kgDMday:2.6,period:'15 Aug - 15 Sep',days:30,feedSource:'Kale',notes:''},
    {id:'d22',species:'Deer',cls:'B11 Hinds',headPrev:220,headCurr:220,kgDMday:4.7,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d23',species:'Deer',cls:'B11 R2 Hinds',headPrev:81,headCurr:81,kgDMday:4.4,period:'15 May - 15 Sep',days:120,feedSource:'River block pasture and rape',notes:''},
    {id:'d24',species:'Deer',cls:'B11 Weaners',headPrev:225,headCurr:225,kgDMday:3.5,period:'16 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'c1',species:'Cattle',cls:'R1 Heifers',headPrev:90,headCurr:90,kgDMday:7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture',notes:''},
    {id:'c2',species:'Cattle',cls:'Ma Cows on Flats',headPrev:80,headCurr:80,kgDMday:5,period:'15 May - 15 Sep',days:120,feedSource:'',notes:''},
    {id:'c3',species:'Cattle',cls:'R1 Finishing Cattle',headPrev:170,headCurr:170,kgDMday:8,period:'15 May - 15 Sep',days:120,feedSource:'',notes:''},
    {id:'c4',species:'Cattle',cls:'Calving Cows',headPrev:355,headCurr:355,kgDMday:7,period:'20 Aug - 15 Sep',days:27,feedSource:'Saved pasture',notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, stockReq);
  SpreadsheetApp.flush();
  Logger.log('StockReq done: ' + stockReq.length + ' rows');
}

function seedSupplements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supplements = [
    {id:'x1',name:'Lucerne / Clover Silage',type:'Lucerne Silage',kgDM:260000,costPerKgDM:0.18,notes:''},
    {id:'x2',name:'Bulk Grass Silage',type:'Silage',kgDM:750000,costPerKgDM:0.18,notes:''},
    {id:'x3',name:'Balage',type:'Balage',kgDM:324000,costPerKgDM:0.28,notes:''},
    {id:'x4',name:'Hay',type:'Hay',kgDM:50000,costPerKgDM:0.18,notes:''},
    {id:'x5',name:'Purchased Feed',type:'Supplement',kgDM:100000,costPerKgDM:0.33,notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS, supplements);
  SpreadsheetApp.flush();
  Logger.log('Supplements done: ' + supplements.length + ' rows');
}

function seedHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const history = [
    {id:'h1',year:'2025/26',paddock:'Thames Valley',crop:'Kale',ha:7.14,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h2',year:'2025/26',paddock:'East Coast',crop:'Kale',ha:6.61,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h3',year:'2025/26',paddock:'Poverty Bay',crop:'Kale',ha:5.5,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h4',year:'2025/26',paddock:'Kingcountry',crop:'Kale',ha:5,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h5',year:'2025/26',paddock:'Bottom 60',crop:'Kale',ha:9.9,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h6',year:'2025/26',paddock:'Boots',crop:'Kale',ha:5.5,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h7',year:'2025/26',paddock:'North Harbour',crop:'Swedes',ha:5.2,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h8',year:'2025/26',paddock:'Wanganui',crop:'Swedes',ha:5.47,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h9',year:'2025/26',paddock:'Bull',crop:'Swedes',ha:6.95,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h10',year:'2025/26',paddock:'Bens',crop:'Swedes',ha:5.17,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h11',year:'2025/26',paddock:'Waikato',crop:'Swedes',ha:11,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h12',year:'2025/26',paddock:'Cattle Yards 2',crop:'Fodder Beet',ha:5.3,yieldBudget:20000,yieldActual:null,notes:''},
    {id:'h13',year:'2025/26',paddock:'Cattle Yards 4',crop:'Fodder Beet',ha:6.3,yieldBudget:20000,yieldActual:null,notes:''},
    {id:'h14',year:'2025/26',paddock:'Woodings',crop:'Fodder Beet',ha:6.88,yieldBudget:20000,yieldActual:null,notes:''},
    {id:'h15',year:'2025/26',paddock:'One Tree',crop:'Lifting Beet',ha:7,yieldBudget:16000,yieldActual:null,notes:''},
    {id:'h16',year:'2025/26',paddock:'Horsfall 1',crop:'Maize',ha:6.8,yieldBudget:18000,yieldActual:null,notes:''},
    {id:'h17',year:'2025/26',paddock:'Horsfall 4',crop:'Maize',ha:6,yieldBudget:18000,yieldActual:null,notes:''},
  ];
  rowsToSheet(ensureSheet(ss,'History',HIST_COLS), HIST_COLS, history);
  setTs(ss, Date.now());
  SpreadsheetApp.flush();
  Logger.log('History done: ' + history.length + ' rows');
}

// DEPRECATED — replaced by the four functions above
function seedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- CROP PLAN 26/27 ---
  const cropPlan = [
    // Kale
    {id:'k1',paddock:'North Harbour',crop:'Kale',ha:5.2,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k2',paddock:'Wanganui',crop:'Kale',ha:5.47,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k3',paddock:'Bull',crop:'Kale',ha:6.95,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k4',paddock:'Bens',crop:'Kale',ha:5.17,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    {id:'k5',paddock:'Waikato',crop:'Kale',ha:11,drillDate:'20/11/2026',yieldKgHA:12000,seedHA:147,chemHA:435,fertHA:566,opsHA:295,notes:'Ex Swedes'},
    // Swedes
    {id:'s1',paddock:'River 1',crop:'Swedes',ha:7.9,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s2',paddock:'River 2',crop:'Swedes',ha:5.17,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s3',paddock:'River 3',crop:'Swedes',ha:5.14,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s4',paddock:'River 4',crop:'Swedes',ha:3.3,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s5',paddock:'Horsfall 3',crop:'Swedes',ha:5.6,drillDate:'11/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s6',paddock:'Ryans',crop:'Swedes',ha:4.7,drillDate:'12/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s7',paddock:'Horsfall 2',crop:'Swedes',ha:4.2,drillDate:'13/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s8',paddock:'Yards',crop:'Swedes',ha:5.03,drillDate:'14/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    {id:'s9',paddock:'Cattle Yards 3',crop:'Swedes',ha:4.53,drillDate:'10/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'Tested'},
    {id:'s10',paddock:'Little Horsfall Swamp',crop:'Swedes',ha:4.0,drillDate:'11/11/2026',yieldKgHA:14000,seedHA:240,chemHA:543,fertHA:535,opsHA:395,notes:'No test'},
    // Fodder Beet
    {id:'fb1',paddock:'Deershed Drive',crop:'Fodder Beet',ha:3.7,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'Tested'},
    {id:'fb2',paddock:'Two Tanks 2',crop:'Fodder Beet',ha:4.54,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'fb3',paddock:'Mckerchers',crop:'Fodder Beet',ha:7.05,drillDate:'20/10/2026',yieldKgHA:20000,seedHA:430,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    // Lifting Beet
    {id:'lb1',paddock:'Sowbie 2',crop:'Lifting Beet',ha:3.84,drillDate:'21/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'Tested'},
    {id:'lb2',paddock:'Sowbie 1',crop:'Lifting Beet',ha:4.68,drillDate:'22/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    {id:'lb3',paddock:'Sowbie 3',crop:'Lifting Beet',ha:4.93,drillDate:'23/10/2026',yieldKgHA:16000,seedHA:521,chemHA:1078,fertHA:335,opsHA:378,notes:'No test'},
    // Maize
    {id:'m1',paddock:'River 5',crop:'Maize',ha:4.5,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:''},
    {id:'m2',paddock:'Golden Willow 3',crop:'Maize',ha:5.0,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:'No test'},
    {id:'m3',paddock:'Golden Willow 4',crop:'Maize',ha:3.0,drillDate:'20/10/2026',yieldKgHA:18000,seedHA:999,chemHA:326,fertHA:892,opsHA:525,notes:'No test'},
    // Triticale
    {id:'t1',paddock:'One Tree',crop:'Triticale',ha:7.8,drillDate:'10/04/2026',yieldKgHA:6000,seedHA:250,chemHA:195,fertHA:163,opsHA:160,notes:''},
    // Rape
    {id:'r1',paddock:'Horsfall 1',crop:'Rape',ha:6.8,drillDate:'11/04/2026',yieldKgHA:8000,seedHA:195,chemHA:195,fertHA:163,opsHA:160,notes:''},
    {id:'r2',paddock:'Horsfall 4',crop:'Rape',ha:6.1,drillDate:'12/04/2026',yieldKgHA:8000,seedHA:195,chemHA:195,fertHA:163,opsHA:160,notes:''},
  ];

  // --- STOCK REQUIREMENTS 26/27 ---
  const stockReq = [
    // Deer
    {id:'d1',species:'Deer',cls:'Barn Stags',headPrev:1250,headCurr:1250,kgDMday:4.75,period:'22 May - 8 Aug',days:75,feedSource:'Barn: sugar beet, maize, silage',notes:''},
    {id:'d2',species:'Deer',cls:'Ma / Old Stags',headPrev:650,headCurr:650,kgDMday:4.7,period:'June - July',days:75,feedSource:'Swedes',notes:''},
    {id:'d3',species:'Deer',cls:'Ma Old Stags (May)',headPrev:1900,headCurr:1900,kgDMday:5,period:'May',days:15,feedSource:'Saved pasture Italians',notes:'Aim for 80ha'},
    {id:'d4',species:'Deer',cls:'Ma Old Stags (Aug-Oct)',headPrev:1900,headCurr:1900,kgDMday:5,period:'Aug - 1 Oct',days:60,feedSource:'Saved pasture',notes:''},
    {id:'d5',species:'Deer',cls:'R3 Stags',headPrev:340,headCurr:340,kgDMday:4.2,period:'15 May - 1 Aug',days:75,feedSource:'Swedes',notes:''},
    {id:'d6',species:'Deer',cls:'R3 Stags (Aug)',headPrev:340,headCurr:340,kgDMday:3.96,period:'1 Aug - 15 Sep',days:45,feedSource:'Barn diet',notes:''},
    {id:'d7',species:'Deer',cls:'R3 Stags Elite',headPrev:315,headCurr:315,kgDMday:4.2,period:'20 May - 10 Aug',days:83,feedSource:'Kale, lifting beet, PK',notes:''},
    {id:'d8',species:'Deer',cls:'R3 Stags Elite (Aug)',headPrev:315,headCurr:315,kgDMday:4.2,period:'10 Aug - 15 Sep',days:35,feedSource:'Autumn saved pasture',notes:''},
    {id:'d9',species:'Deer',cls:'R3 B11 Stags + Ma',headPrev:65,headCurr:65,kgDMday:5.5,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'d10',species:'Deer',cls:'R2 Stags',headPrev:790,headCurr:790,kgDMday:3.6,period:'15 May - 1 Aug',days:78,feedSource:'Swedes',notes:''},
    {id:'d11',species:'Deer',cls:'R2 Stags Barn',headPrev:540,headCurr:540,kgDMday:3.4,period:'25 Jul - 25 Sep',days:62,feedSource:'Barn: sugar beet, maize, silage',notes:''},
    {id:'d12',species:'Deer',cls:'R2 B11 Stags',headPrev:73,headCurr:73,kgDMday:4,period:'15 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    {id:'d13',species:'Deer',cls:'Forresters',headPrev:309,headCurr:309,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture Homeblock 107ha',notes:''},
    {id:'d14',species:'Deer',cls:'Velvet and Trophy',headPrev:457,headCurr:457,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Swedes',notes:''},
    {id:'d15',species:'Deer',cls:'Recips',headPrev:355,headCurr:355,kgDMday:2.7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture 320ha',notes:''},
    {id:'d16',species:'Deer',cls:'Commercial Weaners',headPrev:900,headCurr:900,kgDMday:2.9,period:'10 May - 7 Aug',days:89,feedSource:'Fodder Beet',notes:''},
    {id:'d17',species:'Deer',cls:'Com Females (Aug)',headPrev:440,headCurr:440,kgDMday:2.9,period:'1 Aug - 15 Sep',days:45,feedSource:'Kale with some beet',notes:''},
    {id:'d18',species:'Deer',cls:'Com Males (Aug)',headPrev:450,headCurr:450,kgDMday:3.2,period:'1 Aug - 15 Sep',days:45,feedSource:'Pasture at gorge',notes:''},
    {id:'d19',species:'Deer',cls:'Ma Hinds (recips & studs)',headPrev:260,headCurr:260,kgDMday:2.5,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d20',species:'Deer',cls:'Com Hinds on Flats',headPrev:200,headCurr:200,kgDMday:2.6,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d21',species:'Deer',cls:'Com Hinds on Flats (after hill)',headPrev:800,headCurr:800,kgDMday:2.6,period:'15 Aug - 15 Sep',days:30,feedSource:'Kale',notes:''},
    {id:'d22',species:'Deer',cls:'B11 Hinds',headPrev:220,headCurr:220,kgDMday:4.7,period:'15 May - 15 Sep',days:120,feedSource:'Kale',notes:''},
    {id:'d23',species:'Deer',cls:'B11 R2 Hinds',headPrev:81,headCurr:81,kgDMday:4.4,period:'15 May - 15 Sep',days:120,feedSource:'River block pasture and rape',notes:''},
    {id:'d24',species:'Deer',cls:'B11 Weaners',headPrev:225,headCurr:225,kgDMday:3.5,period:'16 May - 15 Sep',days:120,feedSource:'Pasture with balage',notes:''},
    // Cattle
    {id:'c1',species:'Cattle',cls:'R1 Heifers',headPrev:90,headCurr:90,kgDMday:7,period:'15 May - 15 Sep',days:120,feedSource:'Pasture',notes:''},
    {id:'c2',species:'Cattle',cls:'Ma Cows on Flats',headPrev:80,headCurr:80,kgDMday:5,period:'15 May - 15 Sep',days:120,feedSource:'',notes:''},
    {id:'c3',species:'Cattle',cls:'R1 Finishing Cattle',headPrev:170,headCurr:170,kgDMday:8,period:'15 May - 15 Sep',days:120,feedSource:'',notes:''},
    {id:'c4',species:'Cattle',cls:'Calving Cows',headPrev:355,headCurr:355,kgDMday:7,period:'20 Aug - 15 Sep',days:27,feedSource:'Saved pasture',notes:''},
  ];

  // --- SUPPLEMENTS ---
  const supplements = [
    {id:'x1',name:'Lucerne / Clover Silage',type:'Lucerne Silage',kgDM:260000,costPerKgDM:0.18,notes:''},
    {id:'x2',name:'Bulk Grass Silage',type:'Silage',kgDM:750000,costPerKgDM:0.18,notes:''},
    {id:'x3',name:'Balage',type:'Balage',kgDM:324000,costPerKgDM:0.28,notes:''},
    {id:'x4',name:'Hay',type:'Hay',kgDM:50000,costPerKgDM:0.18,notes:''},
    {id:'x5',name:'Purchased Feed',type:'Supplement',kgDM:100000,costPerKgDM:0.33,notes:''},
  ];

  // --- HISTORY (25/26 actuals and 24/25) ---
  const history = [
    // 25/26
    {id:'h1',year:'2025/26',paddock:'Thames Valley',crop:'Kale',ha:7.14,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h2',year:'2025/26',paddock:'East Coast',crop:'Kale',ha:6.61,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h3',year:'2025/26',paddock:'Poverty Bay',crop:'Kale',ha:5.5,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h4',year:'2025/26',paddock:'Kingcountry',crop:'Kale',ha:5,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h5',year:'2025/26',paddock:'Bottom 60',crop:'Kale',ha:9.9,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h6',year:'2025/26',paddock:'Boots',crop:'Kale',ha:5.5,yieldBudget:12000,yieldActual:null,notes:''},
    {id:'h7',year:'2025/26',paddock:'North Harbour',crop:'Swedes',ha:5.2,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h8',year:'2025/26',paddock:'Wanganui',crop:'Swedes',ha:5.47,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h9',year:'2025/26',paddock:'Bull',crop:'Swedes',ha:6.95,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h10',year:'2025/26',paddock:'Bens',crop:'Swedes',ha:5.17,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h11',year:'2025/26',paddock:'Waikato',crop:'Swedes',ha:11,yieldBudget:14000,yieldActual:null,notes:''},
    {id:'h12',year:'2025/26',paddock:'Cattle Yards 2',crop:'Fodder Beet',ha:5.3,yieldBudget:20000,yieldActual:null,notes:''},
    {id:'h13',year:'2025/26',paddock:'Cattle Yards 4',crop:'Fodder Beet',ha:6.3,yieldBudget:20000,yieldActual:null,notes:''},
    {id:'h14',year:'2025/26',paddock:'Woodings',crop:'Fodder Beet',ha:6.88,yieldBudget:20000,yieldActual:null,notes:''},
    {id:'h15',year:'2025/26',paddock:'One Tree',crop:'Lifting Beet',ha:7,yieldBudget:16000,yieldActual:null,notes:''},
    {id:'h16',year:'2025/26',paddock:'Horsfall 1',crop:'Maize',ha:6.8,yieldBudget:18000,yieldActual:null,notes:''},
    {id:'h17',year:'2025/26',paddock:'Horsfall 4',crop:'Maize',ha:6,yieldBudget:18000,yieldActual:null,notes:''},
  ];

  // Write to sheets
  rowsToSheet(ensureSheet(ss,'CropPlan',CROP_COLS), CROP_COLS, cropPlan);
  rowsToSheet(ensureSheet(ss,'StockReq',STOCK_COLS), STOCK_COLS, stockReq);
  rowsToSheet(ensureSheet(ss,'Supplements',SUPP_COLS), SUPP_COLS, supplements);
  rowsToSheet(ensureSheet(ss,'History',HIST_COLS), HIST_COLS, history);
  setTs(ss, Date.now());

  SpreadsheetApp.getUi().alert('Seed complete. CropPlan: ' + cropPlan.length + ' rows, StockReq: ' + stockReq.length + ' rows, Supplements: ' + supplements.length + ' rows, History: ' + history.length + ' rows.');
}

function respond(json, e) {
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
