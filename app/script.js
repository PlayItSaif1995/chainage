// ═══════════════════════════════════════════════════════════════════
//  STATE
// ═══════════════════════════════════════════════════════════════════
let activities       = [];
let groups           = [];   // {id, name, parentId|null, color, useColor, visible, collapsed}
let risks            = [];   // {id, activityId, title, likelihood, impact, timeDays, costImpact, owner, mitigation, status}
let selectedActivity = null;
let selectedIds      = new Set();
let highlightedGroupId = null;
let timelineStart    = null;
let timelineEnd      = null;
let minChainage      = 0;
let maxChainage      = 1000;
let chainageSpacing  = 50;

let dateMarkers    = [];
let chainageZones  = [];
let activityLinks  = [];
let colourPalette  = JSON.parse(localStorage.getItem('tcplanner_palette')||'[]');
let searchFilter   = '';
let viewFilter     = 'all';

let zoneEditId     = null;
let markerEditId   = null;
let linkEditId     = null;
let groupEditId    = null;

let titleBlock = {
  projectName:'',drawingTitle:'',drawingNumber:'',revision:'',
  drawnBy:'',checkedBy:'',date:'',dataDate:'',scale:'NTS',logoSrc:null
};
let tbLogoImage = null;

let revisions  = [];
const BASELINE_TINTS=['#e67e22','#8e44ad','#16a085','#c0392b','#2980b9','#27ae60','#f39c12','#7f8c8d'];
let baselines  = [];

let chainageImage    = null;
let chainageImageSrc = null;
let chainageImageH   = 120;
let imgResizeDragging   = false;
let imgResizeDragStartY = 0;
let imgResizeDragStartH = 0;

let isDragging     = false;
let dragActivityId = -1;
let dragStartMX    = 0;
let dragStartMY    = 0;
let dragOrigStart  = null;
let dragOrigEnd    = null;
let dragOrigSCh    = 0;
let dragOrigECh    = 0;

let isResizing    = false;
let resizeActIdx  = -1;
let resizeEdge    = '';
let resizeStartMX = 0;
let resizeStartMY = 0;
let resizeOrigVal = null;

let isRubberBanding = false;
let rubberStart     = {x:0,y:0};
let rubberEnd       = {x:0,y:0};

let undoStack = [];
let redoStack = [];
const MAX_UNDO = 50;

// Critical path state — supports MULTIPLE paths simultaneously
// Each path: {id, targetIdx, color, label, highlightSet, chainInOrder, totalDays}
let cpPaths     = [];
let cpHighlightSet = new Set();  // legacy — union of all active paths for toggle checkbox
let cpTargetIdx = -1;

// Current open file name (for PC save)
let currentFileName = 'untitled';

// Spreadsheet editor state
let ssEditingCell  = null;
let ssDragFill     = null;

// Chart view mode: 'programme' | 'resource' | 'cost'
let chartViewMode  = 'programme';

// Currency
let projectCurrency = '£';

// ─── Currency data & custom dropdown ───────────────────────────────
const WORLD_CURRENCIES = [
  ['£','British Pound','GBP'],['€','Euro','EUR'],['$','US Dollar','USD'],
  ['AED','UAE Dirham','AED'],['AFN','Afghan Afghani','AFN'],['ALL','Albanian Lek','ALL'],
  ['AMD','Armenian Dram','AMD'],['ANG','Netherlands Antillean Guilder','ANG'],
  ['AOA','Angolan Kwanza','AOA'],['ARS','Argentine Peso','ARS'],
  ['AUD','Australian Dollar','AUD'],['AWG','Aruban Florin','AWG'],
  ['AZN','Azerbaijani Manat','AZN'],['BAM','Bosnia-Herzegovina Mark','BAM'],
  ['BBD','Barbadian Dollar','BBD'],['BDT','Bangladeshi Taka','BDT'],
  ['BGN','Bulgarian Lev','BGN'],['BHD','Bahraini Dinar','BHD'],
  ['BIF','Burundian Franc','BIF'],['BMD','Bermudian Dollar','BMD'],
  ['BND','Brunei Dollar','BND'],['BOB','Bolivian Boliviano','BOB'],
  ['BRL','Brazilian Real','BRL'],['BSD','Bahamian Dollar','BSD'],
  ['BTN','Bhutanese Ngultrum','BTN'],['BWP','Botswana Pula','BWP'],
  ['BYN','Belarusian Ruble','BYN'],['BZD','Belize Dollar','BZD'],
  ['CAD','Canadian Dollar','CAD'],['CDF','Congolese Franc','CDF'],
  ['CHF','Swiss Franc','CHF'],['CLP','Chilean Peso','CLP'],
  ['CNY','Chinese Yuan','CNY'],['COP','Colombian Peso','COP'],
  ['CRC','Costa Rican Colon','CRC'],['CUP','Cuban Peso','CUP'],
  ['CVE','Cape Verdean Escudo','CVE'],['CZK','Czech Koruna','CZK'],
  ['DJF','Djiboutian Franc','DJF'],['DKK','Danish Krone','DKK'],
  ['DOP','Dominican Peso','DOP'],['DZD','Algerian Dinar','DZD'],
  ['EGP','Egyptian Pound','EGP'],['ERN','Eritrean Nakfa','ERN'],
  ['ETB','Ethiopian Birr','ETB'],['FJD','Fijian Dollar','FJD'],
  ['FKP','Falkland Islands Pound','FKP'],['GEL','Georgian Lari','GEL'],
  ['GHS','Ghanaian Cedi','GHS'],['GIP','Gibraltar Pound','GIP'],
  ['GMD','Gambian Dalasi','GMD'],['GNF','Guinean Franc','GNF'],
  ['GTQ','Guatemalan Quetzal','GTQ'],['GYD','Guyanese Dollar','GYD'],
  ['HKD','Hong Kong Dollar','HKD'],['HNL','Honduran Lempira','HNL'],
  ['HTG','Haitian Gourde','HTG'],['HUF','Hungarian Forint','HUF'],
  ['IDR','Indonesian Rupiah','IDR'],['ILS','Israeli New Shekel','ILS'],
  ['INR','Indian Rupee','INR'],['IQD','Iraqi Dinar','IQD'],
  ['IRR','Iranian Rial','IRR'],['ISK','Icelandic Krona','ISK'],
  ['JMD','Jamaican Dollar','JMD'],['JOD','Jordanian Dinar','JOD'],
  ['JPY','Japanese Yen','JPY'],['KES','Kenyan Shilling','KES'],
  ['KGS','Kyrgyzstani Som','KGS'],['KHR','Cambodian Riel','KHR'],
  ['KMF','Comorian Franc','KMF'],['KPW','North Korean Won','KPW'],
  ['KRW','South Korean Won','KRW'],['KWD','Kuwaiti Dinar','KWD'],
  ['KYD','Cayman Islands Dollar','KYD'],['KZT','Kazakhstani Tenge','KZT'],
  ['LAK','Laotian Kip','LAK'],['LBP','Lebanese Pound','LBP'],
  ['LKR','Sri Lankan Rupee','LKR'],['LRD','Liberian Dollar','LRD'],
  ['LSL','Lesotho Loti','LSL'],['LYD','Libyan Dinar','LYD'],
  ['MAD','Moroccan Dirham','MAD'],['MDL','Moldovan Leu','MDL'],
  ['MGA','Malagasy Ariary','MGA'],['MKD','Macedonian Denar','MKD'],
  ['MMK','Myanmar Kyat','MMK'],['MNT','Mongolian Togrog','MNT'],
  ['MOP','Macanese Pataca','MOP'],['MRU','Mauritanian Ouguiya','MRU'],
  ['MUR','Mauritian Rupee','MUR'],['MVR','Maldivian Rufiyaa','MVR'],
  ['MWK','Malawian Kwacha','MWK'],['MXN','Mexican Peso','MXN'],
  ['MYR','Malaysian Ringgit','MYR'],['MZN','Mozambican Metical','MZN'],
  ['NAD','Namibian Dollar','NAD'],['NGN','Nigerian Naira','NGN'],
  ['NIO','Nicaraguan Cordoba','NIO'],['NOK','Norwegian Krone','NOK'],
  ['NPR','Nepalese Rupee','NPR'],['NZD','New Zealand Dollar','NZD'],
  ['OMR','Omani Rial','OMR'],['PAB','Panamanian Balboa','PAB'],
  ['PEN','Peruvian Sol','PEN'],['PGK','Papua New Guinean Kina','PGK'],
  ['PHP','Philippine Peso','PHP'],['PKR','Pakistani Rupee','PKR'],
  ['PLN','Polish Zloty','PLN'],['PYG','Paraguayan Guarani','PYG'],
  ['QAR','Qatari Riyal','QAR'],['RON','Romanian Leu','RON'],
  ['RSD','Serbian Dinar','RSD'],['RUB','Russian Ruble','RUB'],
  ['RWF','Rwandan Franc','RWF'],['SAR','Saudi Riyal','SAR'],
  ['SBD','Solomon Islands Dollar','SBD'],['SCR','Seychellois Rupee','SCR'],
  ['SDG','Sudanese Pound','SDG'],['SEK','Swedish Krona','SEK'],
  ['SGD','Singapore Dollar','SGD'],['SHP','Saint Helenian Pound','SHP'],
  ['SLL','Sierra Leonean Leone','SLL'],['SOS','Somali Shilling','SOS'],
  ['SRD','Surinamese Dollar','SRD'],['STN','Sao Tome Dobra','STN'],
  ['SVC','Salvadoran Colon','SVC'],['SYP','Syrian Pound','SYP'],
  ['SZL','Swazi Lilangeni','SZL'],['THB','Thai Baht','THB'],
  ['TJS','Tajikistani Somoni','TJS'],['TMT','Turkmenistani Manat','TMT'],
  ['TND','Tunisian Dinar','TND'],['TOP','Tongan Paanga','TOP'],
  ['TRY','Turkish Lira','TRY'],['TTD','Trinidad and Tobago Dollar','TTD'],
  ['TWD','New Taiwan Dollar','TWD'],['TZS','Tanzanian Shilling','TZS'],
  ['UAH','Ukrainian Hryvnia','UAH'],['UGX','Ugandan Shilling','UGX'],
  ['UYU','Uruguayan Peso','UYU'],['UZS','Uzbekistani Som','UZS'],
  ['VES','Venezuelan Bolivar','VES'],['VND','Vietnamese Dong','VND'],
  ['VUV','Vanuatu Vatu','VUV'],['WST','Samoan Tala','WST'],
  ['XAF','Central African CFA Franc','XAF'],['XCD','East Caribbean Dollar','XCD'],
  ['XOF','West African CFA Franc','XOF'],['XPF','CFP Franc','XPF'],
  ['YER','Yemeni Rial','YER'],['ZAR','South African Rand','ZAR'],
  ['ZMW','Zambian Kwacha','ZMW'],['ZWL','Zimbabwean Dollar','ZWL'],
];

function filterCurrencyList(query){
  const dd=document.getElementById('currencyDropdown');
  if(!dd)return;
  const q=(query||'').toLowerCase().trim();
  const matches=q
    ? WORLD_CURRENCIES.filter(([sym,name,code])=>
        name.toLowerCase().includes(q)||code.toLowerCase().includes(q)||sym.toLowerCase().includes(q))
    : WORLD_CURRENCIES;
  dd.style.display='block';
  dd.innerHTML=matches.slice(0,40).map(([sym,name,code])=>{
    const label=`${sym} — ${name} (${code})`;
    return `<div class="currency-opt" onclick="selectCurrency('${sym}','${label}')"
      style="padding:7px 10px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--border-light)"
      onmouseover="this.style.background='var(--hover-row)'"
      onmouseout="this.style.background=''">
      <strong>${sym}</strong> — ${name} <span style="color:var(--text-muted);font-size:11px">${code}</span>
    </div>`;
  }).join('')+(matches.length>40?`<div style="padding:6px 10px;font-size:12px;color:var(--text-muted)">${matches.length-40} more — keep typing to narrow down</div>`:'');
}

function selectCurrency(sym, label){
  projectCurrency=sym;
  const inp=document.getElementById('settingsCurrencySearch');
  if(inp) inp.value=label;
  document.getElementById('currencyDropdown').style.display='none';
  syncCurrencySymbol();
}

// Close dropdown when clicking outside
document.addEventListener('click', e=>{
  const dd=document.getElementById('currencyDropdown');
  const inp=document.getElementById('settingsCurrencySearch');
  if(dd&&inp&&!dd.contains(e.target)&&e.target!==inp){
    dd.style.display='none';
  }
});

function syncCurrencySymbol(){
  // projectCurrency is set by selectCurrency() when user picks from dropdown
  // Also handle direct typed input
  const raw = document.getElementById('settingsCurrencySearch')?.value||'';
  if(raw && !projectCurrency){
    projectCurrency = raw.includes(' — ') ? raw.split(' — ')[0].trim() : raw.trim() || '£';
  }

  const lbl=document.getElementById('actUnitCostLabel');
  if(lbl) lbl.textContent=`💼 Labour rate/day`;

  if(hotInstance){
    const ci=HOT_COLS.findIndex(c=>c.data==='unitCost');
    if(ci>=0){
      HOT_COLS[ci].title=`Labour rate/day (${projectCurrency})`;
      hotInstance.updateSettings({colHeaders:(i)=>_hotColHeader(i)});
    }
  }

  // Immediately redraw whichever chart is active — fixes Bug 4
  if(chartViewMode==='cost')          drawCostChart();
  else if(chartViewMode==='combined') drawCombinedChart();
  else if(chartViewMode==='resource') drawResourceChart();

  renderResourceTab();
}

// ═══════════════════════════════════════════════════════════════════
//  LAYOUT CONSTANTS
// ═══════════════════════════════════════════════════════════════════
const BASE_TOP_MARGIN   = 40;
const BOTTOM_MARGIN     = 40;
const LEFT_MARGIN       = 130;   // slightly wider to fit year labels
const RESIZE_HANDLE_H   = 8;
const MIN_CANVAS_HEIGHT = 650;
const MIN_PX_PER_MONTH  = 28;   // guaranteed pixels per month — enough for a label + breathing room
const TITLE_BLOCK_H     = 90;
const LEGEND_H          = 26;   // height of the subgroup legend strip below the header

// ── Legend helpers ──────────────────────────────────────────────────
function getLegendSubgroups(){
  if(!groups||!groups.length) return [];
  return groups.filter(g=>{
    if(g.parentId===null) return false;
    return activities.some(a=>a.groups&&a.groups.includes(g.id));
  });
}

function drawLegendStrip(c, ct, canvasW, y){
  // y = top of the legend strip
  const stripX = LEFT_MARGIN;
  const stripW = canvasW - LEFT_MARGIN - 40;
  const stripH = LEGEND_H;

  // Background
  c.fillStyle = ct.tbBg || ct.yearBandBg;
  c.fillRect(stripX, y, stripW, stripH);
  c.strokeStyle = ct.tbBorder || ct.yearBandBorder;
  c.lineWidth = 0.7;
  c.strokeRect(stripX, y, stripW, stripH);

  const subs = getLegendSubgroups();
  if(!subs.length) return;

  // Measure and lay out pills
  c.font = 'bold 10px Arial';
  const pillPadX = 8, pillPadY = 4, gap = 10, shapeW = 26;
  let x = stripX + 8;
  const cy2 = y + stripH / 2;

  subs.forEach(sg => {
    const cat = groups.find(g=>g.id===sg.parentId);
    const color = sg.color || (cat&&cat.color) || '#3498db';
    const labelW = c.measureText(sg.name).width;
    const pillW = shapeW + pillPadX + labelW + pillPadX;

    if(x + pillW > stripX + stripW - 8) return; // don't overflow

    // Shape preview — find first activity in this subgroup to get shape
    const sample = activities.find(a=>a.groups&&a.groups.includes(sg.id));
    const shape = sample ? (sample.shape||'rect') : 'rect';

    // Draw the mini shape swatch
    const sw = 16, sh2 = 12;
    const sx = x + 4, sy = cy2 - sh2/2;
    c.save();
    c.fillStyle = color;
    c.strokeStyle = color;
    c.globalAlpha = 0.9;
    switch(shape){
      case 'rect':
        c.fillRect(sx, sy, sw, sh2); break;
      case 'line':
        c.lineWidth=2.5; c.beginPath(); c.moveTo(sx,cy2); c.lineTo(sx+sw,cy2); c.stroke(); break;
      case 'diamond':
        c.beginPath();c.moveTo(sx+sw/2,sy);c.lineTo(sx+sw,cy2);c.lineTo(sx+sw/2,sy+sh2);c.lineTo(sx,cy2);c.closePath();c.fill(); break;
      case 'circle':
        c.beginPath();c.arc(sx+sw/2,cy2,sh2/2,0,Math.PI*2);c.fill(); break;
      case 'triangle':
        c.beginPath();c.moveTo(sx+sw/2,sy);c.lineTo(sx+sw,sy+sh2);c.lineTo(sx,sy+sh2);c.closePath();c.fill(); break;
      case 'star':
        _drawStarOn(c,sx+sw/2,cy2,sh2/2,5); c.fill(); break;
      case 'flag':
        c.lineWidth=1.5;c.beginPath();c.moveTo(sx+3,sy);c.lineTo(sx+3,sy+sh2);c.stroke();
        c.beginPath();c.moveTo(sx+3,sy);c.lineTo(sx+sw,sy+3);c.lineTo(sx+3,sy+6);c.closePath();c.fill(); break;
      default:
        c.fillRect(sx,sy,sw,sh2);
    }
    c.globalAlpha=1; c.restore();

    // Label
    c.fillStyle = ct.axisText;
    c.font = '10px Arial';
    c.textAlign = 'left';
    c.textBaseline = 'middle';
    c.fillText(sg.name, x + shapeW, cy2);

    x += pillW + gap;
  });
}

// Tiny star helper (reused from main draw)
function _drawStarOn(c,cx,cy,r,pts){
  const ri=r*0.42;
  c.beginPath();
  for(let i=0;i<pts*2;i++){
    const a=((i*Math.PI)/pts)-Math.PI/2;
    const rr=i%2===0?r:ri;
    i===0?c.moveTo(cx+rr*Math.cos(a),cy+rr*Math.sin(a)):c.lineTo(cx+rr*Math.cos(a),cy+rr*Math.sin(a));
  }
  c.closePath();
}
const MILESTONE_RADIUS  = 10;
const MILESTONE_SIZE    = 13;
const EDGE_HIT          = 6;
const YEAR_BAND_H       = 18;    // height of year band at top of grid

function getTopMargin(){
  // Extra space: year band + optional image
  const imgExtra = chainageImage ? chainageImageH + RESIZE_HANDLE_H : 0;
  return BASE_TOP_MARGIN + YEAR_BAND_H + imgExtra;
}

/** Calculate the canvas height needed to show every month without skipping.
 *  = topMargin + (months * MIN_PX_PER_MONTH) + BOTTOM_MARGIN + optional title block */
function getRequiredCanvasHeight(includeTitleBlock){
  if(!timelineStart||!timelineEnd) return MIN_CANVAS_HEIGHT;
  const months = (timelineEnd.getFullYear() - timelineStart.getFullYear()) * 12
               + (timelineEnd.getMonth()    - timelineStart.getMonth()) + 1;
  const gridH  = Math.max(months * MIN_PX_PER_MONTH, MIN_CANVAS_HEIGHT - getTopMargin() - BOTTOM_MARGIN);
  return getTopMargin() + gridH + BOTTOM_MARGIN + (includeTitleBlock ? TITLE_BLOCK_H : 0);
}

// ═══════════════════════════════════════════════════════════════════
//  THEME
// ═══════════════════════════════════════════════════════════════════
function getChartTheme(){
  const t=document.documentElement.getAttribute('data-theme')||'white';
  switch(t){
    case'dark':  return{axisText:'#c8dff0',gridH:'#1e3a5a',gridV:'#1a3050',canvasBg:'#16213e',imgBorder:'#2e5080',imgLabel:'#4da3ff',resizeBar:'rgba(77,163,255,0.55)',labelText:'#d0e8ff',tbBg:'#0d1520',tbText:'#c8dff0',tbBorder:'#2e4060',markerText:'#d0e8ff',zoneLabelText:'#c8dff0',yearBandBg:'#1a2e50',yearBandText:'#4da3ff',yearBandBorder:'#2e4a70'};
    case'black': return{axisText:'#e8e8e8',gridH:'#2a2a2a',gridV:'#252525',canvasBg:'#000000',imgBorder:'#444',imgLabel:'#aaa',resizeBar:'rgba(200,200,200,0.25)',labelText:'#f0f0f0',tbBg:'#111',tbText:'#e8e8e8',tbBorder:'#333',markerText:'#f0f0f0',zoneLabelText:'#e8e8e8',yearBandBg:'#1a1a1a',yearBandText:'#cccccc',yearBandBorder:'#333'};
    case'grey':  return{axisText:'#1a1a1a',gridH:'#a8a8a8',gridV:'#a0a0a0',canvasBg:'#dcdcdc',imgBorder:'#999',imgLabel:'#444',resizeBar:'rgba(0,0,0,0.18)',labelText:'#111',tbBg:'#c8c8c8',tbText:'#1a1a1a',tbBorder:'#999',markerText:'#111',zoneLabelText:'#1a1a1a',yearBandBg:'#c8d8ee',yearBandText:'#0056b3',yearBandBorder:'#aabbcc'};
    default:     return{axisText:'#111',gridH:'#e0e0e0',gridV:'#ddd',canvasBg:'#fff',imgBorder:'#ccc',imgLabel:'#555',resizeBar:'rgba(0,0,0,0.13)',labelText:'#111',tbBg:'#f0f4f8',tbText:'#111',tbBorder:'#aaa',markerText:'#111',zoneLabelText:'#333',yearBandBg:'#e8f0fe',yearBandText:'#1a56db',yearBandBorder:'#b8ccee'};
  }
}

// ═══════════════════════════════════════════════════════════════════
//  UNDO / REDO
// ═══════════════════════════════════════════════════════════════════
function pushUndo(){
  undoStack.push(serialiseDiagram());
  if(undoStack.length>MAX_UNDO)undoStack.shift();
  redoStack=[];
  // Auto-save silently to localStorage as a safety net
  try{localStorage.setItem('tcplanner_autosave',JSON.stringify({version:2,currentDiagram:serialiseDiagram(),_savedAt:Date.now()}));}catch(e){}
}
function undo(){if(!undoStack.length)return;redoStack.push(serialiseDiagram());deserialiseDiagram(undoStack.pop());if(hotInstance)hotRefresh();}
function redo(){if(!redoStack.length)return;undoStack.push(serialiseDiagram());deserialiseDiagram(redoStack.pop());if(hotInstance)hotRefresh();}

// ═══════════════════════════════════════════════════════════════════
//  VIEW FILTER
// ═══════════════════════════════════════════════════════════════════
function setViewFilter(f){
  viewFilter=f;
  document.querySelectorAll('.view-btn[id^="viewBtn-all"],[id^="viewBtn-1"],[id^="viewBtn-3"],[id^="viewBtn-6"]').forEach(b=>b.classList.remove('active'));
  document.getElementById('viewBtn-'+f)?.classList.add('active');
  if(chartViewMode==='programme') drawChart();
  else if(chartViewMode==='resource') drawResourceChart();
  else if(chartViewMode==='cost') drawCostChart();
  else if(chartViewMode==='combined') drawCombinedChart();
  // Also refresh Resources tab if it's currently open
  if(document.getElementById('resources')?.classList.contains('active')) renderResourceTab();
}
function getViewDateRange(){
  if(viewFilter==='all')return null;
  const dataDate=titleBlock.dataDate?parseDateLocal(titleBlock.dataDate):new Date();
  const from=new Date(dataDate),to=new Date(dataDate);
  if(viewFilter==='1m') to.setMonth(to.getMonth()+1);
  else if(viewFilter==='3m') to.setMonth(to.getMonth()+3);
  else if(viewFilter==='6m') to.setMonth(to.getMonth()+6);
  else if(viewFilter==='1y') to.setFullYear(to.getFullYear()+1);
  return{from,to};
}
function activityInViewFilter(a){
  const range=getViewDateRange();
  if(!range)return true;
  if(!a.start||!a.end)return false;
  return a.start<=range.to&&a.end>=range.from;
}

// ═══════════════════════════════════════════════════════════════════
//  DATA DATE SYNC
// ═══════════════════════════════════════════════════════════════════
function syncDataDateFromSettings(){
  const val=document.getElementById('settingsDataDate').value;
  titleBlock.dataDate=val;
  document.getElementById('tbDataDate').value=val;
  drawChart();
}

// ═══════════════════════════════════════════════════════════════════
//  PC FILE OPERATIONS (Feature 3 — replaces browser file explorer)
// ═══════════════════════════════════════════════════════════════════

/** Open the Save As dialog so user can name the file before downloading */
function openSaveAsDialog(){
  const input = document.getElementById('saveAsInput');
  // Pre-fill with current name or project name
  input.value = currentFileName !== 'untitled'
    ? currentFileName
    : (titleBlock.projectName || '');
  document.getElementById('saveAsDialog').style.display = 'flex';
  requestAnimationFrame(() => { input.select(); input.focus(); });
}

function closeSaveAsDialog(){
  document.getElementById('saveAsDialog').style.display = 'none';
}

function confirmSaveAs(){
  const raw   = document.getElementById('saveAsInput').value.trim();
  const name  = raw || titleBlock.projectName || 'untitled';
  const safe  = name.replace(/[/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
  currentFileName = safe;
  closeSaveAsDialog();
  const data = { version: 2, currentDiagram: serialiseDiagram() };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = safe + '.tcplan';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 1000);
  updateCurrentFileLabel();
  flashSaved('✅ Saved as ' + safe + '.tcplan');
}

/** quickSaveToPC — now routes through the Save As dialog every time */
function quickSaveToPC(){ openSaveAsDialog(); }

/** Open a .tcplan file from the user's PC */
function openTcplanFromPC(evt){
  const file = evt.target.files[0];
  if(!file) return;
  currentFileName = file.name.replace(/\.tcplan$/i,'');
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const p = JSON.parse(e.target.result);
      // Support both wrapper format {version,currentDiagram} and raw diagram
      const diag = p.currentDiagram || p;
      deserialiseDiagram(diag);
      updateCurrentFileLabel();
      flashSaved('✅ Opened!');
    }catch(err){alert('Failed to open file: '+err.message);}
  };
  reader.readAsText(file);
  evt.target.value='';
}

/** Download current diagram — also called from export modal for .tcplan */
function downloadCurrentDiagram(){
  quickSaveToPC();
}

function updateCurrentFileLabel(){
  const lbl=document.getElementById('currentFileLabel');
  lbl.textContent = currentFileName ? '📄 '+currentFileName+'.tcplan' : 'No file open';
}
function flashSaved(msg){
  const lbl=document.getElementById('currentFileLabel');
  const prev=lbl.textContent;
  lbl.textContent=msg||'✅ Saved!';
  lbl.style.color='#27ae60';
  setTimeout(()=>{lbl.textContent=prev;lbl.style.color='';},1800);
}

// Legacy quickSave — now routes to PC save
function quickSave(){ quickSaveToPC(); }

// ═══════════════════════════════════════════════════════════════════
//  SCHEDULE
// ═══════════════════════════════════════════════════════════════════
function scheduleDiagram(){
  if(!activityLinks.length){alert('No dependencies to schedule.');return;}
  const n=activities.length;
  const inDeg=new Array(n).fill(0);
  const adj=Array.from({length:n},()=>[]);
  activityLinks.forEach(l=>{
    if(l.fromIdx<n&&l.toIdx<n){
      adj[l.fromIdx].push({to:l.toIdx,type:l.type,lag:l.lag||0});
      inDeg[l.toIdx]++;
    }
  });
  const queue=[];for(let i=0;i<n;i++)if(inDeg[i]===0)queue.push(i);
  const order=[];while(queue.length){const v=queue.shift();order.push(v);adj[v].forEach(e=>{inDeg[e.to]--;if(inDeg[e.to]===0)queue.push(e.to);});}
  // Detect circular dependencies — if not all nodes were processed there's a cycle
  if(order.length<n){
    const cycleActs=activities
      .map((a,i)=>({a,i}))
      .filter(({i})=>!order.includes(i))
      .map(({a,i})=>`• ${a.p6Id?a.p6Id+' — ':''}${a.name||'Activity '+(i+1)}`)
      .join('\n');
    alert(`⚠️ Circular dependency detected!\n\nThe following activities form a loop and cannot be scheduled:\n${cycleActs}\n\nPlease check and remove the circular link in the Relationships or Dep. List tab.`);
    return;
  }
  pushUndo();
  order.forEach(fromIdx=>{
    adj[fromIdx].forEach(({to,type,lag})=>{
      const A=activities[fromIdx],B=activities[to];
      if(!A||!B||!A.start||!A.end||!B.start||!B.end)return;
      const dur=B.end-B.start;
      const lagMs=(lag||0)*86400000;
      let newStart=null;
      if(type==='FS') newStart=new Date(A.end.getTime()+lagMs);
      else if(type==='SS') newStart=new Date(A.start.getTime()+lagMs);
      else if(type==='FF'){const newEnd=new Date(A.end.getTime()+lagMs);newStart=new Date(newEnd.getTime()-dur);}
      else if(type==='SF'){const newEnd=new Date(A.start.getTime()+lagMs);newStart=new Date(newEnd.getTime()-dur);}
      if(newStart){B.start=newStart;B.end=new Date(newStart.getTime()+dur);}
    });
  });
  drawChart();renderActivityTable();
  const msg=`✅ Scheduled ${order.length} activities successfully.`;
  flashSaved(msg);
}

// ═══════════════════════════════════════════════════════════════════
//  FULL-SCREEN VIEW (Feature 5 — hover tooltip, read-only)
// ═══════════════════════════════════════════════════════════════════
let fsActivityCoords = [];  // [{activity, x1,y1,x2,y2}] for hover hit-testing

let fsShowCritical = false;

function openFullscreen(){
  document.getElementById('fullscreenOverlay').style.display='flex';
  document.getElementById('fullscreenTitle').textContent=titleBlock.projectName||'Time-Chainage Diagram';
  // Sync the CP toggle button state
  const btn=document.getElementById('fsCpToggle');
  if(btn) btn.textContent=fsShowCritical?'🔴 Hide Critical Path':'🔴 Show Critical Path';
  requestAnimationFrame(()=>renderFullscreenCanvas());
}
function closeFullscreen(){document.getElementById('fullscreenOverlay').style.display='none';}
function toggleFsCriticalPath(){
  fsShowCritical=!fsShowCritical;
  const btn=document.getElementById('fsCpToggle');
  if(btn) btn.textContent=fsShowCritical?'🔴 Hide Critical Path':'🔴 Show Critical Path';
  renderFullscreenCanvas();
}

function renderFullscreenCanvas(){
  const fsCanvas=document.getElementById('fullscreenCanvas');
  const body=document.querySelector('.fullscreen-body');
  fsCanvas.width=body.clientWidth;
  fsCanvas.height=body.clientHeight;
  const opts={showTitleBlock:true,showProgress:true,showBaseline:true,showImage:true,
    showLinks:     document.getElementById('showLinksToggle')?.checked ?? false,
    showDateMarkers:  document.getElementById('showDateMarkersToggle')?.checked ?? true,
    showRiskOverlays: document.getElementById('showRiskOverlaysToggle')?.checked ?? true,
    showPageBreaks:false,resolution:1,
    fsCritical:fsShowCritical};
  const oc=renderOffscreen(opts,fsCanvas.width,fsCanvas.height);
  if(oc){
    const fsCtx=fsCanvas.getContext('2d');
    fsCtx.drawImage(oc,0,0,fsCanvas.width,fsCanvas.height);
  }
  buildFsCoords(fsCanvas.width,fsCanvas.height);
}

function buildFsCoords(W,H){
  fsActivityCoords=[];
  if(!timelineStart||!timelineEnd)return;
  const topM=getTopMargin();
  const hasTB=hasTitleBlock();
  const gridW=W-LEFT_MARGIN-40;
  const gridH=H-topM-BOTTOM_MARGIN-(hasTB?TITLE_BLOCK_H:0);
  activities.forEach((a,i)=>{
    if(!a.start||!a.end)return;
    const y1=_dateToY(a.start,topM,gridH);
    const y2=_dateToY(a.end,topM,gridH);
    const x1=_chainageToX(a.startCh,W,gridW);
    const x2=_chainageToX(a.endCh,W,gridW);
    fsActivityCoords.push({a,i,x1,y1,x2,y2});
  });
}

// Raw coordinate helpers (used by both main canvas and fullscreen)
function _dateToY(date,topM,gridH){
  const totalMs=timelineEnd-timelineStart;
  return topM+((date-timelineStart)/totalMs)*gridH;
}
function _chainageToX(ch,canvasW,gridW){
  return LEFT_MARGIN+(ch-minChainage)/(maxChainage-minChainage)*gridW;
}

// Fullscreen hover
document.addEventListener('DOMContentLoaded',()=>{
  const fsCanvas=document.getElementById('fullscreenCanvas');
  const fsTip=document.getElementById('fsTooltip');

  fsCanvas.addEventListener('mousemove',e=>{
    const r=fsCanvas.getBoundingClientRect();
    const mx=e.clientX-r.left, my=e.clientY-r.top;

    // In non-programme modes, use chart hit regions for tooltip
    if(fsViewMode!=='programme'){
      // Hit test against stored chart regions (from last drawCostChart/drawResourceChart/drawCombinedChart)
      let found=null;
      for(const reg of chartHitRegions){
        if(reg.type==='bar'&&mx>=reg.x&&mx<=reg.x+reg.w&&my>=reg.y&&my<=reg.y+reg.h){found=reg;break;}
        if(reg.type==='point'&&Math.hypot(mx-reg.x,my-reg.y)<=reg.r){found=reg;break;}
      }
      if(found){
        const d=found.data;
        let html=`<strong>${d.label||''}</strong>`;
        if(d.type==='bar'){
          html+=`<br>💼 Labour: ${_fmtC(d.labour||0)}`;
          html+=`<br>🧱 Material: ${_fmtC(d.material||0)}`;
          html+=`<br>🚜 Equipment: ${_fmtC(d.equip||0)}`;
          html+=`<hr style="border:none;border-top:1px solid rgba(255,255,255,0.3);margin:3px 0">`;
          html+=`<strong>Monthly total: ${_fmtC(d.total||0)}</strong>`;
        } else if(d.type==='point'){
          html+=`<br>Cumulative: <strong>${_fmtC(d.cumulative||0)}</strong>`;
          if(d.monthly) html+=`<br>This month: ${_fmtC(d.monthly)}`;
        } else if(d.type==='resource'){
          html+=`<br>Worker-days: <strong>${Math.round(d.resources||0).toLocaleString()}</strong>`;
        }
        fsTip.style.display='block';
        fsTip.style.left=(e.clientX+14)+'px';
        fsTip.style.top=(e.clientY+10)+'px';
        fsTip.innerHTML=html;
      } else {
        fsTip.style.display='none';
      }
      return;
    }

    // Programme mode — hit test activities
    let found=null;
    for(let i=fsActivityCoords.length-1;i>=0;i--){
      const c=fsActivityCoords[i];
      const minX=Math.min(c.x1,c.x2)-5,maxX=Math.max(c.x1,c.x2)+5;
      const minY=Math.min(c.y1,c.y2)-5,maxY=Math.max(c.y1,c.y2)+5;
      if(mx>=minX&&mx<=maxX&&my>=minY&&my<=maxY){found=c;break;}
    }
    if(found){
      const a=found.a;
      fsTip.style.display='block';
      fsTip.style.left=(e.clientX+14)+'px';
      fsTip.style.top =(e.clientY+10)+'px';
      fsTip.innerHTML=buildActivityTooltipHTML(a);
    } else {
      fsTip.style.display='none';
    }
  });
  fsCanvas.addEventListener('mouseleave',()=>{fsTip.style.display='none';});
  // Fullscreen is read-only — no click-to-edit
  fsCanvas.addEventListener('click',e=>e.stopPropagation());
});

/** Format date as "06-Mar-26" */
function formatLongDate(d){
  if(!d||isNaN(d))return'—';
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return String(d.getDate()).padStart(2,'0')+'-'+months[d.getMonth()]+'-'+String(d.getFullYear()).slice(-2);
}

// ═══════════════════════════════════════════════════════════════════
//  EXPORT PREVIEW (Feature 6)
// ═══════════════════════════════════════════════════════════════════
function openExportPreview(){
  if(!timelineStart||!timelineEnd){alert('Set the timeline first.');return;}
  const opts=gatherExportOpts();
  const oc=renderOffscreen(opts);
  if(!oc){alert('Nothing to preview.');return;}
  const pc=document.getElementById('previewCanvas');
  // Fit within preview area without distorting — use full modal width
  const maxW=Math.min(900, window.innerWidth-80);
  const maxH=Math.min(600, window.innerHeight-200);
  const ratio=Math.min(maxW/oc.width, maxH/oc.height, 1);
  pc.width =Math.round(oc.width *ratio);
  pc.height=Math.round(oc.height*ratio);
  pc.style.width =pc.width +'px';
  pc.style.height=pc.height+'px';
  pc.getContext('2d').drawImage(oc,0,0,pc.width,pc.height);
  document.getElementById('exportPreviewModal').style.display='flex';
}
function closeExportPreview(){document.getElementById('exportPreviewModal').style.display='none';}

function onExpFormatChange(){
  // Show/hide PDF-specific and PNG-specific controls
  const fmt=document.querySelector('input[name="expFormat"]:checked')?.value;
  // (could grey out paper size for non-PDF — left as cosmetic for now)
}

function gatherExportOpts(){
  return{
    format:   document.querySelector('input[name="expFormat"]:checked')?.value||'png',
    size:     document.querySelector('input[name="expSize"]:checked')?.value||'a3',
    orientation: document.querySelector('input[name="expOrient"]:checked')?.value||'landscape',
    resolution:  parseInt(document.querySelector('input[name="expRes"]:checked')?.value||'2'),
    showTitleBlock: document.getElementById('expShowTitleBlock').checked,
    showProgress:   document.getElementById('expShowProgress').checked,
    showBaseline:   document.getElementById('expShowBaseline').checked,
    showImage:      document.getElementById('expShowImage').checked,
    showLinks:      document.getElementById('expShowLinks').checked,
    showPageBreaks: document.getElementById('expShowPageBreaks').checked,
    showDateMarkers:  document.getElementById('showDateMarkersToggle')?.checked ?? true,
    showRiskOverlays: document.getElementById('showRiskOverlaysToggle')?.checked ?? true,
    legendGroups:     document.getElementById('expLegendGroups').checked,
    legendActivities: document.getElementById('expLegendActivities').checked,
    legendStandalone: document.getElementById('expLegendStandalone').checked,
  };
}

function updateLegendExportOpts(){
  // Both can be checked independently — no mutual exclusion needed
}

// ═══════════════════════════════════════════════════════════════════
//  EXPORT TO EXCEL
// ═══════════════════════════════════════════════════════════════════
function exportToExcel(){
  if(typeof XLSX==='undefined'){alert('XLSX library not loaded.');return;}
  const rows=activities.map(a=>({
    'P6 ID':a.p6Id||'','Name':a.name||'',
    'Start':a.start?formatLongDate(a.start):'','End':a.end?formatLongDate(a.end):'',
    'Ch Start':a.startCh,'Ch End':a.endCh,
    'Progress %':a.progress||0,'Colour':a.color||'','Shape':a.shape||'',
    'Group(s)':getActivityGroupNames(a),
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Activities');
  XLSX.writeFile(wb,(titleBlock.projectName||'activities')+'.xlsx');
}

function getActivityGroupNames(a){
  if(!a.groups||!a.groups.length)return'';
  return a.groups.map(gid=>{const g=groups.find(x=>x.id===gid);return g?g.name:''}).filter(Boolean).join(', ');
}

// ═══════════════════════════════════════════════════════════════════
//  DATE MARKERS
// ═══════════════════════════════════════════════════════════════════
function addMarker(){
  const name=document.getElementById('markerName').value.trim();
  const date=document.getElementById('markerDate').value;
  const color=document.getElementById('markerColor').value;
  if(!name||!date){alert('Please enter a name and date.');return;}
  pushUndo();
  dateMarkers.push({id:genId(),name,date,color,hideLabel:false});
  document.getElementById('markerName').value='';
  document.getElementById('markerDate').value='';
  renderMarkerList();drawChart();
}
function openMarkerEditDialog(id){
  markerEditId=id;
  const m=dateMarkers.find(x=>x.id===id);if(!m)return;
  document.getElementById('markerEditName').value=m.name;
  document.getElementById('markerEditDate').value=m.date;
  document.getElementById('markerEditColor').value=m.color;
  document.getElementById('markerEditHideLabel').checked=!!m.hideLabel;
  document.getElementById('markerEditDialog').style.display='flex';
}
function confirmEditMarker(){
  const m=dateMarkers.find(x=>x.id===markerEditId);if(!m){closeMarkerEditDialog();return;}
  pushUndo();
  m.name=document.getElementById('markerEditName').value.trim()||m.name;
  m.date=document.getElementById('markerEditDate').value;
  m.color=document.getElementById('markerEditColor').value;
  m.hideLabel=document.getElementById('markerEditHideLabel').checked;
  closeMarkerEditDialog();renderMarkerList();drawChart();
}
function closeMarkerEditDialog(){document.getElementById('markerEditDialog').style.display='none';markerEditId=null;}
function deleteMarker(id){pushUndo();dateMarkers=dateMarkers.filter(m=>m.id!==id);renderMarkerList();drawChart();}
function renderMarkerList(){
  const el=document.getElementById('markerList');if(!el)return;
  if(!dateMarkers.length){el.innerHTML='<p style="color:var(--text-muted);font-size:13px">No markers added.</p>';return;}
  let html='<table class="activityTable"><thead><tr><th>Colour</th><th>Name</th><th>Date</th><th>Label</th><th>Actions</th></tr></thead><tbody>';
  dateMarkers.forEach(m=>{
    html+=`<tr>
      <td><span class="colorSwatch" style="background:${m.color}"></span></td>
      <td>${m.name}</td><td>${formatDisplayDate(m.date)}</td>
      <td>${m.hideLabel?'Hidden':'Visible'}</td>
      <td><button onclick="openMarkerEditDialog('${m.id}')">✏️ Edit</button>
          <button onclick="deleteMarker('${m.id}')" class="deleteBtn" style="margin-left:4px">Delete</button></td>
    </tr>`;
  });
  html+='</tbody></table>';el.innerHTML=html;
}

// ═══════════════════════════════════════════════════════════════════
//  CHAINAGE ZONES
// ═══════════════════════════════════════════════════════════════════
function addZone(){
  const name=document.getElementById('zoneName').value.trim();
  const sCh=parseFloat(document.getElementById('zoneStartCh').value);
  const eCh=parseFloat(document.getElementById('zoneEndCh').value);
  const color=document.getElementById('zoneColor').value;
  const alpha=parseFloat(document.getElementById('zoneAlpha').value);
  if(!name||isNaN(sCh)||isNaN(eCh)){alert('Fill in all zone fields.');return;}
  pushUndo();
  chainageZones.push({id:genId(),name,startCh:sCh,endCh:eCh,color,alpha,hideLabel:false});
  document.getElementById('zoneName').value='';
  renderZoneList();drawChart();
}
function openZoneEditDialog(id){
  zoneEditId=id;const z=chainageZones.find(x=>x.id===id);if(!z)return;
  document.getElementById('zoneEditName').value=z.name;
  document.getElementById('zoneEditStartCh').value=z.startCh;
  document.getElementById('zoneEditEndCh').value=z.endCh;
  document.getElementById('zoneEditColor').value=z.color;
  document.getElementById('zoneEditAlpha').value=z.alpha;
  document.getElementById('zoneEditHideLabel').checked=!!z.hideLabel;
  document.getElementById('zoneEditDialog').style.display='flex';
}
function confirmEditZone(){
  const z=chainageZones.find(x=>x.id===zoneEditId);if(!z){closeZoneEditDialog();return;}
  pushUndo();
  z.name=document.getElementById('zoneEditName').value.trim()||z.name;
  z.startCh=parseFloat(document.getElementById('zoneEditStartCh').value);
  z.endCh=parseFloat(document.getElementById('zoneEditEndCh').value);
  z.color=document.getElementById('zoneEditColor').value;
  z.alpha=parseFloat(document.getElementById('zoneEditAlpha').value);
  z.hideLabel=document.getElementById('zoneEditHideLabel').checked;
  closeZoneEditDialog();renderZoneList();drawChart();
}
function closeZoneEditDialog(){document.getElementById('zoneEditDialog').style.display='none';zoneEditId=null;}
function deleteZone(id){pushUndo();chainageZones=chainageZones.filter(z=>z.id!==id);renderZoneList();drawChart();}
function renderZoneList(){
  const el=document.getElementById('zoneList');if(!el)return;
  if(!chainageZones.length){el.innerHTML='<p style="color:var(--text-muted);font-size:13px">No zones added.</p>';return;}
  let html='<table class="activityTable"><thead><tr><th>Colour</th><th>Name</th><th>Start Ch</th><th>End Ch</th><th>Label</th><th>Actions</th></tr></thead><tbody>';
  chainageZones.forEach(z=>{
    html+=`<tr><td><span class="colorSwatch" style="background:${z.color}"></span></td>
    <td>${z.name}</td><td>${z.startCh}</td><td>${z.endCh}</td>
    <td>${z.hideLabel?'Hidden':'Visible'}</td>
    <td><button onclick="openZoneEditDialog('${z.id}')">✏️ Edit</button>
        <button onclick="deleteZone('${z.id}')" class="deleteBtn" style="margin-left:4px">Delete</button></td></tr>`;
  });
  html+='</tbody></table>';el.innerHTML=html;
}

// ═══════════════════════════════════════════════════════════════════
//  ACTIVITY LINKS
// ═══════════════════════════════════════════════════════════════════
function openLinkDialog(editId){
  if(activities.length<2){alert('Need at least 2 activities.');return;}
  linkEditId=editId||null;
  const from=document.getElementById('linkFrom'),to=document.getElementById('linkTo');
  from.innerHTML='';to.innerHTML='';
  activities.forEach((a,i)=>{const opt=`<option value="${i}">${a.p6Id||i} — ${a.name}</option>`;from.innerHTML+=opt;to.innerHTML+=opt;});
  if(linkEditId){
    const l=activityLinks.find(x=>x.id===linkEditId);
    if(l){from.value=l.fromIdx;to.value=l.toIdx;document.getElementById('linkType').value=l.type;document.getElementById('linkLag').value=l.lag||0;}
    document.getElementById('linkDialogTitle').textContent='✏️ Edit Dependency Link';
  } else {
    if(activities.length>1)to.selectedIndex=1;
    document.getElementById('linkLag').value=0;
    document.getElementById('linkDialogTitle').textContent='🔗 Add Dependency Link';
  }
  document.getElementById('linkDialog').style.display='flex';
}
function closeLinkDialog(){document.getElementById('linkDialog').style.display='none';linkEditId=null;}
function confirmSaveLink(){
  const fromIdx=parseInt(document.getElementById('linkFrom').value);
  const toIdx  =parseInt(document.getElementById('linkTo').value);
  const type   =document.getElementById('linkType').value;
  const lag    =parseFloat(document.getElementById('linkLag').value)||0;
  if(fromIdx===toIdx){alert('From and To must be different.');return;}
  pushUndo();
  if(linkEditId){const l=activityLinks.find(x=>x.id===linkEditId);if(l){l.fromIdx=fromIdx;l.toIdx=toIdx;l.type=type;l.lag=lag;}}
  else activityLinks.push({id:genId(),fromIdx,toIdx,type,lag});
  closeLinkDialog();renderLinkList();drawChart();
  populateCpTargetSelect();
}
function deleteLink(id){pushUndo();activityLinks=activityLinks.filter(l=>l.id!==id);renderLinkList();renderRelationshipsPanel();drawChart();}
function getActivityZoneName(a){
  if(!a||!chainageZones.length) return '—';
  const midCh=(a.startCh+a.endCh)/2;
  const z=chainageZones.find(z=>midCh>=z.startCh&&midCh<=z.endCh);
  return z?z.name:'—';
}

function renderLinkList(){
  const el=document.getElementById('linkList');if(!el)return;
  if(!activityLinks.length){el.innerHTML='<p style="color:var(--text-muted);font-size:13px">No dependencies added yet. Use the button above or Ctrl+click two activities on the chart.</p>';return;}
  let html=`<table class="activityTable"><thead><tr>
    <th>From Activity</th><th>Ch Range</th><th>Zone</th>
    <th>Type</th><th>Lag (days)</th>
    <th>To Activity</th><th>Ch Range</th><th>Zone</th>
    <th>Actions</th>
  </tr></thead><tbody>`;
  activityLinks.forEach(l=>{
    const from=activities[l.fromIdx],to=activities[l.toIdx];
    if(!from||!to)return;
    const fZone=getActivityZoneName(from);
    const tZone=getActivityZoneName(to);
    html+=`<tr>
      <td><strong>${from.name}</strong>${from.p6Id?'<br><span style="font-size:11px;color:var(--text-muted)">'+from.p6Id+'</span>':''}</td>
      <td style="font-size:11px">${from.startCh}–${from.endCh}</td>
      <td style="font-size:11px">${fZone}</td>
      <td>
        <select onchange="updateLinkField('${l.id}','type',this.value)" style="width:80px;font-size:12px;padding:2px;margin:0">
          <option value="FS"${l.type==='FS'?' selected':''}>FS</option>
          <option value="SS"${l.type==='SS'?' selected':''}>SS</option>
          <option value="FF"${l.type==='FF'?' selected':''}>FF</option>
          <option value="SF"${l.type==='SF'?' selected':''}>SF</option>
        </select>
      </td>
      <td>
        <input type="number" value="${l.lag||0}" style="width:60px;font-size:12px;padding:2px;margin:0"
          onchange="updateLinkField('${l.id}','lag',parseFloat(this.value)||0)">
      </td>
      <td><strong>${to.name}</strong>${to.p6Id?'<br><span style="font-size:11px;color:var(--text-muted)">'+to.p6Id+'</span>':''}</td>
      <td style="font-size:11px">${to.startCh}–${to.endCh}</td>
      <td style="font-size:11px">${tZone}</td>
      <td>
        <button onclick="openLinkDialog('${l.id}')" style="margin:0;padding:3px 7px;font-size:11px">✏️ Edit</button>
        <button onclick="deleteLink('${l.id}')" class="deleteBtn" style="margin:2px 0 0 2px;padding:3px 7px;font-size:11px">✕</button>
      </td>
    </tr>`;
  });
  html+='</tbody></table>';el.innerHTML=html;
}

function updateLinkField(id,field,value){
  const l=activityLinks.find(x=>x.id===id);
  if(!l)return;
  pushUndo();
  l[field]=value;
  drawChart();
}

// ═══════════════════════════════════════════════════════════════════
//  P6-STYLE RELATIONSHIPS PANEL
//  Activity search → Predecessors | Successors side by side
// ═══════════════════════════════════════════════════════════════════
let relPanelActIdx = -1;   // currently selected activity in the panel

// State kept between renders so inputs aren't lost
let relPendingPred = {actIdx:'', type:'FS', lag:0};
let relPendingSucc = {actIdx:'', type:'FS', lag:0};
let relPredSearch  = '';
let relSuccSearch  = '';

function renderRelationshipsPanel(){
  const el=document.getElementById('relationshipsPanel');if(!el)return;

  // Activity search list
  const actListHtml = ()=>{
    let h=`<div class="rel-search-bar">
      <input id="relSearchInput" class="search-input" placeholder="🔍 Search activity by name or P6 ID…"
        oninput="filterRelActivityList(this.value)"
        style="width:100%;max-width:400px;margin:0">
    </div>
    <div id="relActivityList" class="rel-act-list">`;
    activities.forEach((a,i)=>{
      const isSel=i===relPanelActIdx;
      h+=`<div class="rel-act-item${isSel?' rel-act-selected':''}" onclick="selectRelActivity(${i})">
        <span class="rel-act-name">${a.name||'Unnamed'}</span>
        ${a.p6Id?'<span class="rel-act-id">'+a.p6Id+'</span>':''}
      </div>`;
    });
    h+=`</div>`;
    return h;
  };

  if(relPanelActIdx<0||!activities[relPanelActIdx]){
    el.innerHTML=actListHtml()+`<div class="rel-empty">👆 Select an activity above to view and edit its relationships.</div>`;
    return;
  }

  const a=activities[relPanelActIdx];
  const preds=activityLinks.filter(l=>l.toIdx===relPanelActIdx);
  const succs=activityLinks.filter(l=>l.fromIdx===relPanelActIdx);
  const typeOpts=(sel)=>['FS','SS','FF','SF'].map(t=>`<option value="${t}"${t===sel?' selected':''}>${t}</option>`).join('');

  const predRows=preds.map(l=>{
    const fa=activities[l.fromIdx];if(!fa)return'';
    return`<tr>
      <td><span class="rel-act-pill" title="${fa.name}">${fa.name}${fa.p6Id?' <em>('+fa.p6Id+')</em>':''}</span></td>
      <td><select onchange="updateLinkField('${l.id}','type',this.value)" class="rel-type-sel">${typeOpts(l.type)}</select></td>
      <td><input type="number" value="${l.lag||0}" class="rel-lag-inp" onchange="updateLinkField('${l.id}','lag',parseFloat(this.value)||0)"></td>
      <td><button onclick="deleteLink('${l.id}')" class="deleteBtn" style="margin:0;padding:2px 6px;font-size:11px">✕</button></td>
    </tr>`;
  }).join('');

  const succRows=succs.map(l=>{
    const ta=activities[l.toIdx];if(!ta)return'';
    return`<tr>
      <td><span class="rel-act-pill" title="${ta.name}">${ta.name}${ta.p6Id?' <em>('+ta.p6Id+')</em>':''}</span></td>
      <td><select onchange="updateLinkField('${l.id}','type',this.value)" class="rel-type-sel">${typeOpts(l.type)}</select></td>
      <td><input type="number" value="${l.lag||0}" class="rel-lag-inp" onchange="updateLinkField('${l.id}','lag',parseFloat(this.value)||0)"></td>
      <td><button onclick="deleteLink('${l.id}')" class="deleteBtn" style="margin:0;padding:2px 6px;font-size:11px">✕</button></td>
    </tr>`;
  }).join('');

  // Searchable activity options for add rows
  const filteredPredOpts = activities.map((ac,i)=>{
    if(i===relPanelActIdx)return'';
    const text=((ac.p6Id||'')+' '+ac.name).toLowerCase();
    if(relPredSearch&&!text.includes(relPredSearch.toLowerCase()))return'';
    return`<option value="${i}"${String(i)===String(relPendingPred.actIdx)?' selected':''}>${ac.p6Id?ac.p6Id+' — ':''}${ac.name}</option>`;
  }).join('');

  const filteredSuccOpts = activities.map((ac,i)=>{
    if(i===relPanelActIdx)return'';
    const text=((ac.p6Id||'')+' '+ac.name).toLowerCase();
    if(relSuccSearch&&!text.includes(relSuccSearch.toLowerCase()))return'';
    return`<option value="${i}"${String(i)===String(relPendingSucc.actIdx)?' selected':''}>${ac.p6Id?ac.p6Id+' — ':''}${ac.name}</option>`;
  }).join('');

  el.innerHTML=`
    ${actListHtml()}
    <div class="rel-panel-header">
      <strong>Relationships for: <span style="color:var(--tab-active-bg)">${a.name}</span></strong>
      ${a.p6Id?`<span style="font-size:12px;color:var(--text-muted);margin-left:8px">${a.p6Id}</span>`:''}
    </div>
    <div class="rel-columns">

      <!-- PREDECESSORS -->
      <div class="rel-col">
        <div class="rel-col-header rel-col-pred">◀ Predecessors <span class="rel-count">${preds.length}</span></div>
        <table class="rel-table">
          <thead><tr><th>Activity</th><th>Type</th><th>Lag (d)</th><th></th></tr></thead>
          <tbody>
            ${predRows||'<tr><td colspan="4" style="color:var(--text-muted);font-size:12px;padding:8px">No predecessors</td></tr>'}
            <tr class="rel-add-row">
              <td colspan="4">
                <div style="display:flex;flex-direction:column;gap:4px">
                  <input class="search-input" placeholder="🔍 Search predecessor…" value="${relPredSearch}"
                    oninput="relPredSearch=this.value;_rerenderRelCols()" style="width:100%;margin:0;font-size:11px">
                  <div style="display:flex;gap:4px;align-items:center">
                    <select id="relAddPredAct" style="flex:1;font-size:12px;padding:3px;margin:0"
                      onchange="relPendingPred.actIdx=this.value">
                      <option value="">— Select predecessor —</option>${filteredPredOpts}
                    </select>
                    <select id="relAddPredType" style="width:58px;font-size:12px;padding:3px;margin:0"
                      onchange="relPendingPred.type=this.value">${typeOpts(relPendingPred.type)}</select>
                    <input type="number" id="relAddPredLag" value="${relPendingPred.lag}"
                      style="width:48px;font-size:12px;padding:3px;margin:0"
                      oninput="relPendingPred.lag=parseFloat(this.value)||0" title="Lag (days)">
                    <button onclick="relAddPredecessor()" style="margin:0;padding:3px 8px;font-size:12px;background:#2980b9;color:#fff;white-space:nowrap">＋ Add</button>
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- SUCCESSORS -->
      <div class="rel-col">
        <div class="rel-col-header rel-col-succ">Successors ▶ <span class="rel-count">${succs.length}</span></div>
        <table class="rel-table">
          <thead><tr><th>Activity</th><th>Type</th><th>Lag (d)</th><th></th></tr></thead>
          <tbody>
            ${succRows||'<tr><td colspan="4" style="color:var(--text-muted);font-size:12px;padding:8px">No successors</td></tr>'}
            <tr class="rel-add-row">
              <td colspan="4">
                <div style="display:flex;flex-direction:column;gap:4px">
                  <input class="search-input" placeholder="🔍 Search successor…" value="${relSuccSearch}"
                    oninput="relSuccSearch=this.value;_rerenderRelCols()" style="width:100%;margin:0;font-size:11px">
                  <div style="display:flex;gap:4px;align-items:center">
                    <select id="relAddSuccAct" style="flex:1;font-size:12px;padding:3px;margin:0"
                      onchange="relPendingSucc.actIdx=this.value">
                      <option value="">— Select successor —</option>${filteredSuccOpts}
                    </select>
                    <select id="relAddSuccType" style="width:58px;font-size:12px;padding:3px;margin:0"
                      onchange="relPendingSucc.type=this.value">${typeOpts(relPendingSucc.type)}</select>
                    <input type="number" id="relAddSuccLag" value="${relPendingSucc.lag}"
                      style="width:48px;font-size:12px;padding:3px;margin:0"
                      oninput="relPendingSucc.lag=parseFloat(this.value)||0" title="Lag (days)">
                    <button onclick="relAddSuccessor()" style="margin:0;padding:3px 8px;font-size:12px;background:#27ae60;color:#fff;white-space:nowrap">＋ Add</button>
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

    </div>
  `;
}

/** Only update the filtered dropdown options — never rebuild the inputs (avoids focus loss) */
function _rerenderRelCols(){
  if(relPanelActIdx<0||!activities[relPanelActIdx]) return;

  const filteredOpts = (searchVal, excludeIdx, pendingIdx) =>
    activities.map((ac,i)=>{
      if(i===excludeIdx) return '';
      const text=((ac.p6Id||'')+' '+ac.name).toLowerCase();
      if(searchVal && !text.includes(searchVal.toLowerCase())) return '';
      return `<option value="${i}"${String(i)===String(pendingIdx)?' selected':''}>${ac.p6Id?ac.p6Id+' — ':''}${ac.name}</option>`;
    }).join('');

  const predSel = document.getElementById('relAddPredAct');
  const succSel = document.getElementById('relAddSuccAct');

  if(predSel){
    const cur = predSel.value;
    predSel.innerHTML = `<option value="">— Select predecessor —</option>`
      + filteredOpts(relPredSearch, relPanelActIdx, relPendingPred.actIdx);
    if(cur) predSel.value = cur;
  }
  if(succSel){
    const cur = succSel.value;
    succSel.innerHTML = `<option value="">— Select successor —</option>`
      + filteredOpts(relSuccSearch, relPanelActIdx, relPendingSucc.actIdx);
    if(cur) succSel.value = cur;
  }
}

function selectRelActivity(idx){
  relPanelActIdx=idx;
  relPendingPred={actIdx:'',type:'FS',lag:0};
  relPendingSucc={actIdx:'',type:'FS',lag:0};
  relPredSearch='';relSuccSearch='';
  renderRelationshipsPanel();
}

function filterRelActivityList(query){
  const q=query.toLowerCase().trim();
  const list=document.getElementById('relActivityList');if(!list)return;
  list.querySelectorAll('.rel-act-item').forEach((el,i)=>{
    const a=activities[i];if(!a)return;
    const match=!q||a.name.toLowerCase().includes(q)||(a.p6Id&&a.p6Id.toLowerCase().includes(q));
    el.style.display=match?'':'none';
  });
}

function relAddPredecessor(){
  const fromIdx=parseInt(document.getElementById('relAddPredAct').value);
  if(isNaN(fromIdx)||fromIdx<0||fromIdx===''){alert('Select a predecessor activity first.');return;}
  const type=relPendingPred.type||document.getElementById('relAddPredType').value;
  const lag =relPendingPred.lag||parseFloat(document.getElementById('relAddPredLag').value)||0;
  if(activityLinks.find(l=>l.fromIdx===fromIdx&&l.toIdx===relPanelActIdx)){alert('This relationship already exists.');return;}
  pushUndo();
  activityLinks.push({id:genId(),fromIdx,toIdx:relPanelActIdx,type,lag});
  relPendingPred={actIdx:'',type:'FS',lag:0};  // clear only predecessor input
  renderRelationshipsPanel();renderLinkList();drawChart();
}

function relAddSuccessor(){
  const toIdx=parseInt(document.getElementById('relAddSuccAct').value);
  if(isNaN(toIdx)||toIdx<0||toIdx===''){alert('Select a successor activity first.');return;}
  const type=relPendingSucc.type||document.getElementById('relAddSuccType').value;
  const lag =relPendingSucc.lag||parseFloat(document.getElementById('relAddSuccLag').value)||0;
  if(activityLinks.find(l=>l.fromIdx===relPanelActIdx&&l.toIdx===toIdx)){alert('This relationship already exists.');return;}
  pushUndo();
  activityLinks.push({id:genId(),fromIdx:relPanelActIdx,toIdx,type,lag});
  relPendingSucc={actIdx:'',type:'FS',lag:0};  // clear only successor input
  renderRelationshipsPanel();renderLinkList();drawChart();
}

// ═══════════════════════════════════════════════════════════════════
//  CRITICAL PATH — Proper backwards trace (Feature 4)
// ═══════════════════════════════════════════════════════════════════

function populateCpTargetSelect(){
  const sel=document.getElementById('cpTargetSelect');
  if(!sel)return;
  const prevVals=new Set(Array.from(sel.selectedOptions).map(o=>o.value));
  sel.innerHTML='';
  activities.forEach((a,i)=>{
    const opt=document.createElement('option');
    opt.value=i;
    opt.textContent=(a.p6Id?a.p6Id+' — ':'')+a.name;
    if(prevVals.has(String(i)))opt.selected=true;
    sel.appendChild(opt);
  });
}

function filterCpTargetSelect(query){
  const q=query.toLowerCase().trim();
  const sel=document.getElementById('cpTargetSelect');
  if(!sel)return;
  Array.from(sel.options).forEach(opt=>{
    const a=activities[parseInt(opt.value)];
    const text=a?(((a.p6Id||'')+' '+a.name).toLowerCase()):'';
    opt.style.display=(!q||text.includes(q))?'':'none';
  });
}

// ═══════════════════════════════════════════════════════════════════
//  CRITICAL PATH — Multi-path with colour picking
// ═══════════════════════════════════════════════════════════════════
const CP_DEFAULT_COLOURS = ['#e74c3c','#f39c12','#8e44ad','#16a085','#2980b9','#27ae60','#e67e22','#c0392b'];

function computeAndShowCriticalPath(){
  const sel=document.getElementById('cpTargetSelect');
  const selectedIndices=Array.from(sel.selectedOptions)
    .map(o=>parseInt(o.value))
    .filter(i=>!isNaN(i)&&i>=0&&activities[i]);

  if(!selectedIndices.length){
    alert('Please select at least one target activity.\nHold Ctrl (or Cmd on Mac) to select multiple at once.');
    return;
  }

  let added=0;
  selectedIndices.forEach(targetIdx=>{
    if(cpPaths.find(p=>p.targetIdx===targetIdx)) return; // skip duplicates

    const result=_computePathToTarget(targetIdx);
    if(!result) return;

    const usedColours=cpPaths.map(p=>p.color);
    const nextColour=CP_DEFAULT_COLOURS.find(c=>!usedColours.includes(c))||CP_DEFAULT_COLOURS[cpPaths.length%CP_DEFAULT_COLOURS.length];

    cpPaths.push({
      id:genId(), targetIdx, color:nextColour,
      label:activities[targetIdx].name,
      highlightSet:result.highlightSet,
      chainInOrder:result.chainInOrder,
      totalDays:result.totalDays,
    });
    added++;
  });

  if(!added){
    document.getElementById('cpResults').innerHTML=
      '<div class="cp-summary">No new paths could be computed — they may already exist, or those activities have no incoming dependencies.</div>';
    return;
  }

  cpTargetIdx=selectedIndices[selectedIndices.length-1];
  _rebuildCpHighlightSet();
  renderAllCriticalPaths();
  drawChart();
}

function _computePathToTarget(targetIdx){
  const predecessors = Array.from({length:activities.length},()=>[]);
  activityLinks.forEach(l=>{
    if(l.toIdx>=0&&l.toIdx<activities.length)
      predecessors[l.toIdx].push({fromIdx:l.fromIdx,type:l.type,lag:l.lag||0,linkId:l.id});
  });
  const visited=new Set();
  function collectAncestors(idx){if(visited.has(idx))return;visited.add(idx);predecessors[idx].forEach(p=>collectAncestors(p.fromIdx));}
  collectAncestors(targetIdx);
  if(visited.size<=1)return null;

  const subAdj=Array.from({length:activities.length},()=>[]);
  activityLinks.forEach(l=>{if(visited.has(l.fromIdx)&&visited.has(l.toIdx))subAdj[l.fromIdx].push(l);});
  const inDeg=new Array(activities.length).fill(0);
  activityLinks.forEach(l=>{if(visited.has(l.fromIdx)&&visited.has(l.toIdx))inDeg[l.toIdx]++;});
  const q=[];visited.forEach(idx=>{if(inDeg[idx]===0)q.push(idx);});
  const topoOrder=[];const qq=[...q];
  while(qq.length){const v=qq.shift();topoOrder.push(v);subAdj[v].forEach(l=>{inDeg[l.toIdx]--;if(inDeg[l.toIdx]===0)qq.push(l.toIdx);});}

  const chainDays=new Array(activities.length).fill(0);
  [...topoOrder].reverse().forEach(idx=>{
    const a=activities[idx];
    const dur=a.start&&a.end?(a.end-a.start)/86400000:0;
    let maxSucc=0;
    subAdj[idx].forEach(l=>maxSucc=Math.max(maxSucc,(l.lag||0)+chainDays[l.toIdx]));
    chainDays[idx]=dur+maxSucc;
  });

  const maxChain=Math.max(...[...visited].map(i=>chainDays[i]));
  const highlightSet=new Set();
  function traceCritical(idx,rem){
    if(Math.abs(chainDays[idx]-rem)>0.5)return false;
    highlightSet.add(idx);
    const dur=activities[idx].start&&activities[idx].end?(activities[idx].end-activities[idx].start)/86400000:0;
    subAdj[idx].forEach(l=>traceCritical(l.toIdx,rem-dur-(l.lag||0)));
    return true;
  }
  visited.forEach(idx=>{if(Math.abs(chainDays[idx]-maxChain)<0.5)traceCritical(idx,maxChain);});

  return{
    highlightSet,
    chainInOrder: topoOrder.filter(i=>highlightSet.has(i)),
    totalDays: maxChain,
    subAdj,
  };
}

function _rebuildCpHighlightSet(){
  cpHighlightSet=new Set();
  cpPaths.forEach(p=>p.highlightSet.forEach(i=>cpHighlightSet.add(i)));
}

function getActivityCpColor(idx){
  // Returns colour of the first (or highest-priority) path that highlights this activity
  for(const p of cpPaths){
    if(p.highlightSet.has(idx))return p.color;
  }
  return null;
}

function removeCpPath(id){
  cpPaths=cpPaths.filter(p=>p.id!==id);
  _rebuildCpHighlightSet();
  renderAllCriticalPaths();
  drawChart();
}

function updateCpPathColor(id,color){
  const p=cpPaths.find(x=>x.id===id);if(!p)return;
  p.color=color;
  _rebuildCpHighlightSet();
  renderAllCriticalPaths();
  drawChart();
}

function clearCriticalPathAnalysis(){
  cpPaths=[];cpHighlightSet=new Set();cpTargetIdx=-1;
  document.getElementById('cpResults').innerHTML='';
  document.getElementById('cpTargetSelect').value='';
  const si=document.getElementById('cpSearchInput');if(si)si.value='';
  populateCpTargetSelect();
  drawChart();
}

function renderAllCriticalPaths(){
  const el=document.getElementById('cpResults');if(!el)return;
  if(!cpPaths.length){el.innerHTML='';return;}

  let html='';

  cpPaths.forEach((path,pi)=>{
    const targetName=activities[path.targetIdx]?.name||'Unknown';
    html+=`
      <div class="cp-path-block" style="border-left:4px solid ${path.color};margin-bottom:18px;padding-left:12px">
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap">
          <span style="font-size:15px;font-weight:bold;color:${path.color}">🎯 Path ${pi+1}: ${targetName}</span>
          <input type="color" value="${path.color}" title="Change path colour"
            onchange="updateCpPathColor('${path.id}',this.value)"
            style="width:32px;height:26px;padding:1px;border:1px solid var(--border);border-radius:4px;cursor:pointer;margin:0">
          <span style="font-size:12px;color:var(--text-muted)">~${Math.round(path.totalDays)} days · ${path.chainInOrder.length} activities</span>
          <button onclick="removeCpPath('${path.id}')" class="deleteBtn" style="margin:0;padding:2px 8px;font-size:11px">✕ Remove</button>
        </div>
        <div class="cp-chain">
    `;

    path.chainInOrder.forEach((idx,pos)=>{
      const a=activities[idx];if(!a)return;
      const isTarget=idx===path.targetIdx;
      const dur=a.start&&a.end?Math.round((a.end-a.start)/86400000):0;
      let linkLabel='';
      if(pos<path.chainInOrder.length-1){
        const nextIdx=path.chainInOrder[pos+1];
        const lnk=activityLinks.find(l=>l.fromIdx===idx&&l.toIdx===nextIdx);
        if(lnk)linkLabel=`<span class="cp-node-link">${lnk.type}${lnk.lag?` +${lnk.lag}d`:''}</span>`;
      }
      html+=`
        <div class="cp-node${isTarget?' cp-target':''}" style="border-left-color:${path.color}">
          <div style="display:flex;flex-direction:column;flex:1;gap:2px">
            <div style="display:flex;align-items:center;gap:8px">
              <span style="font-size:14px">${isTarget?'🎯':'●'}</span>
              <span class="cp-node-name">${a.name||'Unnamed'}${a.p6Id?' <span style="font-weight:normal;color:var(--text-muted);font-size:11px">('+a.p6Id+')</span>':''}</span>
              ${linkLabel}
            </div>
            <div class="cp-node-dates">
              ${a.start?formatLongDate(a.start):'?'} → ${a.end?formatLongDate(a.end):'?'}
              &nbsp;|&nbsp; <strong>${dur}d</strong>
              &nbsp;|&nbsp; Ch: ${a.startCh}–${a.endCh}
            </div>
          </div>
        </div>
        ${pos<path.chainInOrder.length-1?`<div class="cp-arrow" style="color:${path.color}">↓</div>`:''}
      `;
    });

    html+=`</div></div>`;
  });

  // Summary across all paths
  html+=`<div class="cp-summary" style="margin-top:8px">
    <strong>${cpPaths.length} path${cpPaths.length>1?'s':''} active</strong> — activities on any path are highlighted on the chart in their path colour.
    <br><span style="font-size:12px;color:var(--text-muted)">Add more paths by selecting another target activity above and clicking Analyse.</span>
  </div>`;

  el.innerHTML=html;
}

// Legacy function — kept for the Show Critical Path toggle checkbox
function computeCriticalPath(){
  if(cpHighlightSet.size>0)return cpHighlightSet;
  const n=activities.length;if(!n||!activityLinks.length)return new Set();
  const dur=activities.map(a=>(!a.start||!a.end)?0:(a.end-a.start)/86400000);
  const adj=Array.from({length:n},()=>[]);
  activityLinks.forEach(l=>{if(l.fromIdx<n&&l.toIdx<n)adj[l.fromIdx].push(l.toIdx);});
  const dist=new Array(n).fill(0),vis=new Array(n).fill(false);
  function dfs(v){vis[v]=true;adj[v].forEach(u=>{if(!vis[u])dfs(u);dist[v]=Math.max(dist[v],dur[v]+(dist[u]||0));});}
  for(let i=0;i<n;i++)if(!vis[i])dfs(i);
  const maxD=Math.max(...dist);const critSet=new Set();
  function mark(v,rem){critSet.add(v);adj[v].forEach(u=>{if(Math.abs(dist[u]-(rem-dur[v]))<0.5)mark(u,rem-dur[v]);});}
  dist.forEach((d,i)=>{if(Math.abs(d-maxD)<0.5)mark(i,maxD);});
  return critSet;
}

// ═══════════════════════════════════════════════════════════════════
//  COLOUR PALETTE
// ═══════════════════════════════════════════════════════════════════
function openPaletteModal(){renderPaletteSwatches();document.getElementById('paletteModal').style.display='flex';}
function closePaletteModal(){document.getElementById('paletteModal').style.display='none';}
function renderPaletteSwatches(){
  const el=document.getElementById('paletteSwatches');if(!el)return;
  el.innerHTML='';
  colourPalette.forEach((c,i)=>{
    const sw=document.createElement('div');sw.className='palette-swatch';sw.style.background=c;sw.title=c;
    sw.onclick=()=>{document.getElementById('actColor').value=c;closePaletteModal();};
    sw.oncontextmenu=e=>{e.preventDefault();colourPalette.splice(i,1);savePalette();renderPaletteSwatches();};
    el.appendChild(sw);
  });
  if(!colourPalette.length)el.innerHTML='<p style="font-size:13px;color:var(--text-muted)">No colours saved yet.</p>';
}
function addCurrentColourToPalette(){const c=document.getElementById('actColor').value;if(!colourPalette.includes(c)){colourPalette.push(c);savePalette();renderPaletteSwatches();}}
function savePalette(){localStorage.setItem('tcplanner_palette',JSON.stringify(colourPalette));}

// ═══════════════════════════════════════════════════════════════════
//  MULTI-SELECT
// ═══════════════════════════════════════════════════════════════════
function updateSelectionUI(){
  const count=selectedIds.size;
  ['btnBulkDelete','btnBulkGroup','btnBulkColour','btnDeselectAll'].forEach(id=>document.getElementById(id).style.display=count>0?'inline-block':'none');
  const info=document.getElementById('selectionInfo');
  if(count>0){info.style.display='inline';info.textContent=`${count} selected`;}
  else info.style.display='none';
}
function deselectAll(){selectedIds.clear();updateSelectionUI();drawChart();}
function bulkDelete(){
  if(!confirm(`Delete ${selectedIds.size} activities?`))return;
  pushUndo();
  const keep=new Set([...selectedIds]);
  activities=activities.filter((_,i)=>!keep.has(i));
  activityLinks=activityLinks.filter(l=>!keep.has(l.fromIdx)&&!keep.has(l.toIdx));
  selectedIds.clear();updateSelectionUI();drawChart();renderActivityTable();renderLinkList();
}
// ── Complaint 5: Bulk group assignment modal ────────────────────────
function openBulkGroupModal(){
  if(!groups.length){alert('Create a group first in the Grouping tab.');return;}
  const count=selectedIds.size;
  document.getElementById('bulkGroupCount').textContent=`${count} activit${count===1?'y':'ies'} selected`;
  document.getElementById('bulkGroupApplyCount').textContent=count;
  // Build group picker
  const picker=document.getElementById('bulkGroupPicker');
  const categories=groups.filter(g=>g.parentId===null);
  picker.innerHTML=categories.map(cat=>{
    const subs=groups.filter(g=>g.parentId===cat.id);
    return `<div style="margin-bottom:8px">
      <label style="display:flex;align-items:center;gap:7px;font-weight:bold;cursor:pointer">
        <input type="checkbox" class="bulk-grp-chk" value="${cat.id}" style="width:auto;margin:0">
        <span class="group-color-dot" style="background:${cat.color||'#3498db'}"></span>
        📂 ${cat.name}
      </label>
      ${subs.map(sg=>`
        <label style="display:flex;align-items:center;gap:7px;padding-left:20px;margin-top:4px;cursor:pointer">
          <input type="checkbox" class="bulk-grp-chk" value="${sg.id}" style="width:auto;margin:0">
          <span class="group-color-dot" style="background:${sg.color||'#e74c3c'}"></span>
          └ ${sg.name}
        </label>`).join('')}
    </div>`;
  }).join('');
  document.getElementById('bulkGroupModal').style.display='flex';
}

function confirmBulkGroup(){
  const mode=document.querySelector('input[name="bulkGroupMode"]:checked')?.value||'add';
  const selectedGroupIds=[...document.querySelectorAll('.bulk-grp-chk:checked')].map(c=>c.value);
  if(!selectedGroupIds.length&&mode!=='replace'){alert('Select at least one group.');return;}
  pushUndo();
  selectedIds.forEach(i=>{
    if(!activities[i])return;
    if(!activities[i].groups)activities[i].groups=[];
    if(mode==='replace'){
      activities[i].groups=selectedGroupIds;
    } else if(mode==='add'){
      selectedGroupIds.forEach(gid=>{if(!activities[i].groups.includes(gid))activities[i].groups.push(gid);});
    } else if(mode==='remove'){
      activities[i].groups=activities[i].groups.filter(gid=>!selectedGroupIds.includes(gid));
    }
  });
  document.getElementById('bulkGroupModal').style.display='none';
  drawChart();renderActivityTable();deselectAll();
}

// ── Complaint 6: Fullscreen view mode ──────────────────────────────
let fsViewMode='programme';

function setFsViewMode(mode){
  fsViewMode=mode;
  ['fsBtnProgramme','fsBtnResource','fsBtnCost','fsBtnCombined'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    el.style.background='transparent';el.style.color='#ccc';
  });
  const active=document.getElementById('fsBtn'+mode.charAt(0).toUpperCase()+mode.slice(1));
  if(active){active.style.background='var(--tab-active-bg,#2980b9)';active.style.color='#fff';}
  renderFullscreenCanvas();
}

// Patch renderFullscreenCanvas to support all modes
function renderFullscreenCanvas(){
  const fsCanvas=document.getElementById('fullscreenCanvas');
  const body=document.querySelector('.fullscreen-body');
  if(!fsCanvas||!body)return;
  fsCanvas.width=body.clientWidth;
  fsCanvas.height=body.clientHeight;

  if(fsViewMode==='programme'){
    const opts={showTitleBlock:true,showProgress:true,showBaseline:true,showImage:true,
      showLinks:     document.getElementById('showLinksToggle')?.checked ?? false,
      showDateMarkers:  document.getElementById('showDateMarkersToggle')?.checked ?? true,
      showRiskOverlays: document.getElementById('showRiskOverlaysToggle')?.checked ?? true,
      showPageBreaks:false,resolution:1,fsCritical:fsShowCritical};
    const oc=renderOffscreen(opts,fsCanvas.width,fsCanvas.height);
    if(oc){fsCanvas.getContext('2d').drawImage(oc,0,0,fsCanvas.width,fsCanvas.height);}
    buildFsCoords(fsCanvas.width,fsCanvas.height);
  } else {
    // Resource/cost/combined — render at full fsCanvas dimensions
    const fsCtx=fsCanvas.getContext('2d');
    const origW=canvas.width,origH=canvas.height;
    canvas.width=fsCanvas.width;
    canvas.height=fsCanvas.height;
    if(fsViewMode==='resource')      drawResourceChart();
    else if(fsViewMode==='cost')     drawCostChart();
    else if(fsViewMode==='combined') drawCombinedChart();
    fsCtx.drawImage(canvas,0,0);
    canvas.width=origW;canvas.height=origH;
    // Restore main canvas
    if(chartViewMode==='programme')      drawChart();
    else if(chartViewMode==='resource')  drawResourceChart();
    else if(chartViewMode==='cost')      drawCostChart();
    else if(chartViewMode==='combined')  drawCombinedChart();
  }
}

// ── Complaint 9: Print multiple views ──────────────────────────────
function openPrintModal(){document.getElementById('printModal').style.display='flex';}

function printCurrentView(){
  // If fullscreen is open, grab its canvas directly and print immediately
  const fsOverlay=document.getElementById('fullscreenOverlay');
  if(fsOverlay&&fsOverlay.style.display!=='none'){
    const fsCanvas=document.getElementById('fullscreenCanvas');
    if(fsCanvas){
      const url=fsCanvas.toDataURL('image/png');
      const projectName=titleBlock.projectName||'Time-Chainage Programme';
      const printDate=new Date().toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
      const modeLabel={programme:'Programme',resource:'Resources',cost:'Cost S-Curve',combined:'Combined'}[fsViewMode]||'View';
      const win=window.open('','_blank');
      if(!win){alert('Pop-up blocked — please allow pop-ups.');return;}
      win.document.write(`<!DOCTYPE html><html><head><title>${projectName} — ${modeLabel}</title>
        <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;background:#fff}
        .ph{display:flex;justify-content:space-between;align-items:center;padding:8mm 10mm 4mm;border-bottom:2px solid #333}
        .pt{font-size:16px;font-weight:bold}.pm{font-size:11px;color:#666}
        img{width:100%;height:auto;display:block;padding:8mm}
        @media print{*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}</style>
      </head><body>
        <div class="ph"><div class="pt">${projectName} — ${modeLabel}</div><div class="pm">${printDate}</div></div>
        <img src="${url}">
      </body></html>`);
      win.document.close();win.onload=()=>{win.focus();win.print();};
    }
    return;
  }
  // Otherwise open the print selection modal
  openPrintModal();
}

function executePrint(){
  const printProgramme=document.getElementById('printProgramme')?.checked;
  const printResources=document.getElementById('printResources')?.checked;
  const printCost     =document.getElementById('printCost')?.checked;
  const printCombined =document.getElementById('printCombined')?.checked;
  if(!printProgramme&&!printResources&&!printCost&&!printCombined){
    alert('Please select at least one view to print.');return;
  }
  document.getElementById('printModal').style.display='none';

  // Build print pages — each a canvas image
  const pages=[];
  const W=Math.min(1400,canvas.parentElement.offsetWidth||900);
  const H=Math.round(W*0.65);

  const origMode=chartViewMode;
  const origW=canvas.width,origH=canvas.height;

  function renderToDataURL(mode){
    canvas.width=W;canvas.height=mode==='programme'?canvas.height:H;
    if(mode==='programme'){
      const oc=renderOffscreen({
        showTitleBlock:true,showProgress:true,showBaseline:true,showImage:true,
        showLinks:     document.getElementById('showLinksToggle')?.checked ?? true,
        showDateMarkers:  document.getElementById('showDateMarkersToggle')?.checked ?? true,
        showRiskOverlays: document.getElementById('showRiskOverlaysToggle')?.checked ?? true,
        showPageBreaks:false,resolution:1
      },W,canvas.height);
      if(oc){ctx.drawImage(oc,0,0);}
    } else if(mode==='resource') drawResourceChart();
    else if(mode==='cost')       drawCostChart();
    else if(mode==='combined')   drawCombinedChart();
    return canvas.toDataURL('image/png');
  }

  if(printProgramme) pages.push({title:'📅 Programme',url:renderToDataURL('programme')});
  if(printResources) pages.push({title:'👷 Resources',url:renderToDataURL('resource')});
  if(printCost)      pages.push({title:'💰 Cost (S-Curve)',url:renderToDataURL('cost')});
  if(printCombined)  pages.push({title:'📊 Combined',url:renderToDataURL('combined')});

  // Restore canvas
  canvas.width=origW;canvas.height=origH;
  setChartViewMode(origMode);

  // Open print window
  const win=window.open('','_blank');
  if(!win){alert('Pop-up blocked — please allow pop-ups for this page.');return;}
  const projectName=titleBlock.projectName||'Time-Chainage Programme';
  const printDate=new Date().toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  win.document.write(`<!DOCTYPE html><html><head>
    <title>${projectName} — Print</title>
    <style>
      *{margin:0;padding:0;box-sizing:border-box}
      body{font-family:Arial,sans-serif;background:#fff;color:#000}
      .page{page-break-after:always;padding:12mm}
      .page:last-child{page-break-after:avoid}
      .page-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6mm;border-bottom:1.5px solid #333;padding-bottom:4mm}
      .page-title{font-size:16px;font-weight:bold}
      .page-meta{font-size:11px;color:#666}
      img{width:100%;height:auto;display:block;border:1px solid #e0e0e0}
      @media print{.page{padding:8mm}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    </style>
  </head><body>`);
  pages.forEach(p=>{
    win.document.write(`<div class="page">
      <div class="page-header">
        <div class="page-title">${projectName} — ${p.title}</div>
        <div class="page-meta">Printed: ${printDate}</div>
      </div>
      <img src="${p.url}" alt="${p.title}">
    </div>`);
  });
  win.document.write(`</body></html>`);
  win.document.close();
  win.onload=()=>{ win.focus(); win.print(); };
}

// ── Tooltip with resources/cost ─────────────────────────────────────
function buildActivityTooltipHTML(a){
  const fmtCost=v=>{const sym=projectCurrency||'£';return v>=1e6?sym+(v/1e6).toFixed(2)+'M':v>=1e3?sym+(v/1e3).toFixed(0)+'k':sym+Math.round(v);};
  const dur=a.start&&a.end?Math.round((a.end-a.start)/86400000):0;
  const gNames=getActivityGroupNames(a);
  let html=`<b>${a.name}</b>`;
  if(a.p6Id) html+=`<br><span style="color:#aaa;font-size:11px">${a.p6Id}</span>`;
  html+=`<br>Start: ${formatLongDate(a.start)}<br>End: ${formatLongDate(a.end)}`;
  html+=`<br>Duration: ${dur} days`;
  html+=`<br>Chainage: ${a.startCh}–${a.endCh}`;
  html+=`<br>Progress: ${a.progress||0}%`;
  // Chainage validation indicator
  if(a.startCh>a.endCh) html+=` <span style="color:#e74c3c">⚠️ Ch reversed</span>`;
  // Resources & Cost
  const res      = parseFloat(a.resourceCount)||0;
  const labRate  = parseFloat(a.unitCost)||0;
  const equipDay = parseFloat(a.equipmentCost)||0;
  const matFixed = parseFloat(a.materialCost)||0;
  const labTotal = res * labRate * dur;
  const equipTotal = equipDay * dur;
  const totalCostAct = labTotal + equipTotal + matFixed;

  if(res || labRate || equipDay || matFixed){
    html+=`<hr style="border:none;border-top:1px solid rgba(255,255,255,0.2);margin:4px 0">`;
    if(res)       html+=`<br>👷 Workers: ${res}/day`;
    if(labRate)   html+=`<br>💼 Labour: ${fmtCost(labRate)}/worker/day → ${fmtCost(labTotal)} total`;
    if(equipDay)  html+=`<br>🚜 Equipment: ${fmtCost(equipDay)}/day → ${fmtCost(equipTotal)} total`;
    if(matFixed)  html+=`<br>🧱 Materials: ${fmtCost(matFixed)} (fixed)`;
    if(totalCostAct) html+=`<br>💰 Activity total: <strong>${fmtCost(totalCostAct)}</strong>`;
    if(res&&dur)  html+=`<br>Worker-days: ${Math.round(res*dur)}`;
  }
  if(a.notes) html+=`<hr style="border:none;border-top:1px solid rgba(255,255,255,0.2);margin:4px 0"><em style="font-size:11px">${a.notes}</em>`;
  if(gNames)  html+=`<br><span style="font-size:11px;color:#aaa">Groups: ${gNames}</span>`;
  return html;
}

function bulkColour(){
  // Use a hidden colour input for a proper OS colour picker
  let inp=document.getElementById('_bulkColourInput');
  if(!inp){
    inp=document.createElement('input');
    inp.type='color';inp.id='_bulkColourInput';
    inp.style.cssText='position:absolute;opacity:0;pointer-events:none;width:0;height:0;';
    document.body.appendChild(inp);
  }
  inp.value='#e74c3c';
  inp.onchange=()=>{
    pushUndo();
    selectedIds.forEach(i=>{if(activities[i])activities[i].color=inp.value;});
    drawChart();renderActivityTable();deselectAll();
  };
  inp.click();
}

// ═══════════════════════════════════════════════════════════════════
//  SEARCH / VIEW FILTER
// ═══════════════════════════════════════════════════════════════════
function onSearchChange(val){searchFilter=val.trim().toLowerCase();drawChart();}
function clearSearch(){document.getElementById('actSearchInput').value='';searchFilter='';drawChart();}
function activityMatchesSearch(a){
  if(!searchFilter)return true;
  return(a.name&&a.name.toLowerCase().includes(searchFilter))||(a.p6Id&&a.p6Id.toLowerCase().includes(searchFilter))||(a.notes&&a.notes.toLowerCase().includes(searchFilter));
}

// ═══════════════════════════════════════════════════════════════════
//  MERGE FROM FILE (Option A — adds on top, does not replace)
// ═══════════════════════════════════════════════════════════════════
function mergeFromFile(evt){
  const file = evt.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const p    = JSON.parse(e.target.result);
      const diag = p.currentDiagram || p;
      const incoming = diag.activities || [];
      if(!incoming.length){ alert('No activities found in the file to merge.'); return; }

      pushUndo();
      const offset = activities.length; // index offset for links

      // Merge activities
      incoming.forEach(a=>{
        activities.push({
          ...a,
          start:  a.start ? new Date(a.start) : null,
          end:    a.end   ? new Date(a.end)   : null,
          groups: [], // don't merge group assignments — groups may differ
        });
      });

      // Merge links (offset indices)
      const incomingLinks = diag.activityLinks || [];
      incomingLinks.forEach(l=>{
        activityLinks.push({
          ...l,
          id:      genId(),
          fromIdx: l.fromIdx + offset,
          toIdx:   l.toIdx   + offset,
        });
      });

      // Merge zones (by name — skip duplicates)
      const existingZoneNames = new Set(chainageZones.map(z=>z.name));
      (diag.chainageZones||[]).forEach(z=>{
        if(!existingZoneNames.has(z.name)) chainageZones.push({...z, id:genId()});
      });

      // Merge date markers (by name — skip duplicates)
      const existingMarkerNames = new Set(dateMarkers.map(m=>m.name));
      (diag.dateMarkers||[]).forEach(m=>{
        if(!existingMarkerNames.has(m.name)) dateMarkers.push({...m, id:genId()});
      });

      drawChart();
      renderActivityTable();
      renderLinkList();
      renderZoneList();
      renderMarkerList();
      populateCpTargetSelect();
      if(hotInstance) hotRefresh();

      alert(`✅ Merged successfully:\n• ${incoming.length} activities added\n• ${incomingLinks.length} dependency links added\nYour existing activities were not changed.`);
    } catch(err){ alert('Failed to merge file: ' + err.message); }
  };
  reader.readAsText(file);
  evt.target.value = '';
}

// Keep importFromFile as an alias for backwards compatibility
function importFromFile(evt){ mergeFromFile(evt); }

// ═══════════════════════════════════════════════════════════════════
//  EXCEL IMPORT
// ═══════════════════════════════════════════════════════════════════
function importExcel(){
  const file=document.getElementById('xlsxFile').files[0];if(!file){alert('Select a file.');return;}
  const debugDiv=document.getElementById('xlsxDebug');
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length){debugDiv.innerHTML='<p class="csv-error">No rows found.</p>';return;}
      const norm=s=>String(s).toLowerCase().replace(/[^a-z0-9]/g,'');
      const VALID_SHAPES=['rect','line','circle','diamond','triangle','star','flag'];
      const VALID_LABELS=['no','left','center','right'];
      let imported=0,skipped=0;
      rows.forEach(row=>{
        const keys=Object.keys(row);
        const get=(...cands)=>{for(const c of cands){const k=keys.find(k=>norm(k)===c);if(k&&row[k]!=='')return String(row[k]).trim();}return null;};
        const name=get('name','activityname','description');
        const sRaw=get('start','startdate','plannedstart','earlystart');
        const eRaw=get('end','finish','enddate','finishdate','plannedfinish','earlyfinish');
        const sCh=parseFloat(get('startch','startchainage','chstart')||'NaN');
        const eCh=parseFloat(get('endch','endchainage','chend')||'NaN');
        if(!name||!sRaw||!eRaw||isNaN(sCh)||isNaN(eCh)){skipped++;return;}
        const start=parseExcelDate(sRaw),end=parseExcelDate(eRaw);
        if(!start||!end){skipped++;return;}
        // Shape — accept rect, line, circle, diamond, triangle
        const rawShape=(get('shape')||'rect').toLowerCase();
        const shape=VALID_SHAPES.includes(rawShape)?rawShape:'rect';
        // Alpha/transparency — accept 0–1 or 0–100
        let alpha=parseFloat(get('alpha','transparency','opacity')||'1');
        if(alpha>1)alpha=alpha/100; // treat >1 as percentage
        alpha=Math.max(0.05,Math.min(1,alpha||1));
        // Label position
        const rawLabel=(get('label','labelposition','showlabel')||'no').toLowerCase();
        const label=VALID_LABELS.includes(rawLabel)?rawLabel:'no';
        pushUndo();
        activities.push({
          p6Id:get('p6id','activityid','id','taskid')||'',
          name,start,end,startCh:sCh,endCh:eCh,
          color:get('colour','color')||'#3498db',
          alpha,shape,label,
          progress:Math.min(100,Math.max(0,parseFloat(get('progress','%complete','percentcomplete')||'0')||0)),
          notes:get('notes','comments')||'',
          groups:[],
          // Resource & cost fields
          resourceCount: parseFloat(get('workers','resourcecount','resources','workersday')||'0')||0,
          unitCost:      parseFloat(get('labourrate','unitcost','labourday','labourperday','rateperworker')||'0')||0,
          equipmentCost: parseFloat(get('equipmentday','equipday','equipmenthire','equipmentperday')||'0')||0,
          materialCost:  parseFloat(get('materialtotal','materialcost','materials','material')||'0')||0,
          // Outline fields
          outlineStyle: (['none','solid','dashed','dotted','dash-dot'].includes((get('outlinestyle','outline')||'none').toLowerCase()))?(get('outlinestyle','outline')||'none'):'none',
          outlineColor: get('outlinecolour','outlinecolor')||'#000000',
          outlineWidth: parseFloat(get('outlinewidth','outlinethickness')||'1.5')||1.5,
        });
        imported++;
      });
      debugDiv.innerHTML=`<p class="csv-success">✅ ${imported} imported, ${skipped} skipped.</p>`;
      drawChart();renderActivityTable();populateCpTargetSelect();
    }catch(err){debugDiv.innerHTML=`<p class="csv-error">❌ ${err.message}</p>`;}
  };
  reader.readAsArrayBuffer(file);
}
function parseExcelDate(val){
  if(!val)return null;
  if(typeof val==='number'||/^\d{4,5}$/.test(String(val))){
    const d=XLSX.SSF.parse_date_code(Number(val));if(d)return new Date(d.y,d.m-1,d.d);
  }
  return parseDateLocal(String(val));
}

// ═══════════════════════════════════════════════════════════════════
//  TITLE BLOCK
// ═══════════════════════════════════════════════════════════════════
function hasTitleBlock(){return !!(titleBlock.projectName||titleBlock.drawingNumber||titleBlock.drawingTitle);}
function saveTitleBlock(){
  titleBlock.projectName  =document.getElementById('tbProjectName').value.trim();
  titleBlock.drawingTitle =document.getElementById('tbDrawingTitle').value.trim();
  titleBlock.drawingNumber=document.getElementById('tbDrawingNumber').value.trim();
  titleBlock.revision     =document.getElementById('tbRevision').value.trim();
  titleBlock.drawnBy      =document.getElementById('tbDrawnBy').value.trim();
  titleBlock.checkedBy    =document.getElementById('tbCheckedBy').value.trim();
  titleBlock.date         =document.getElementById('tbDate').value;
  titleBlock.scale        =document.getElementById('tbScale').value.trim();
  drawChart();
  const msg=document.getElementById('tbSaveMsg');msg.textContent='✅ Saved';setTimeout(()=>msg.textContent='',2000);
}
function loadTbLogo(evt){
  const file=evt.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=e=>{
    titleBlock.logoSrc=e.target.result;
    const img=new Image();
    img.onload=()=>{tbLogoImage=img;document.getElementById('tbLogoName').textContent=file.name;document.getElementById('tbLogoPreview').style.display='block';document.getElementById('tbLogoImg').src=titleBlock.logoSrc;document.getElementById('tbLogoRemove').style.display='inline-block';drawChart();};
    img.src=titleBlock.logoSrc;
  };
  reader.readAsDataURL(file);
}
function removeTbLogo(){
  titleBlock.logoSrc=null;tbLogoImage=null;
  document.getElementById('tbLogoInput').value='';document.getElementById('tbLogoName').textContent='No logo loaded';
  document.getElementById('tbLogoPreview').style.display='none';document.getElementById('tbLogoRemove').style.display='none';
  drawChart();
}
function syncTitleBlockUI(){
  document.getElementById('tbProjectName').value  =titleBlock.projectName  ||'';
  document.getElementById('tbDrawingTitle').value =titleBlock.drawingTitle ||'';
  document.getElementById('tbDrawingNumber').value=titleBlock.drawingNumber||'';
  document.getElementById('tbRevision').value     =titleBlock.revision     ||'';
  document.getElementById('tbDrawnBy').value      =titleBlock.drawnBy      ||'';
  document.getElementById('tbCheckedBy').value    =titleBlock.checkedBy    ||'';
  document.getElementById('tbDate').value         =titleBlock.date         ||'';
  document.getElementById('tbDataDate').value     =titleBlock.dataDate     ||'';
  document.getElementById('tbScale').value        =titleBlock.scale        ||'NTS';
  document.getElementById('settingsDataDate').value=titleBlock.dataDate    ||'';
  if(titleBlock.logoSrc){
    const img=new Image();
    img.onload=()=>{tbLogoImage=img;document.getElementById('tbLogoPreview').style.display='block';document.getElementById('tbLogoImg').src=titleBlock.logoSrc;document.getElementById('tbLogoRemove').style.display='inline-block';document.getElementById('tbLogoName').textContent='Saved logo';};
    img.src=titleBlock.logoSrc;
  } else {
    tbLogoImage=null;document.getElementById('tbLogoPreview').style.display='none';document.getElementById('tbLogoRemove').style.display='none';document.getElementById('tbLogoName').textContent='No logo loaded';
  }
}

function formatShortDate(date){
  if(!date||isNaN(date))return'';
  const d=String(date.getDate()).padStart(2,'0'),m=String(date.getMonth()+1).padStart(2,'0'),y=String(date.getFullYear()).slice(-2);
  return`${d}/${m}/${y}`;
}
function formatDisplayDate(isoStr){
  if(!isoStr)return'—';
  const d=new Date(isoStr+'T12:00:00');
  return d.toLocaleDateString(undefined,{day:'2-digit',month:'short',year:'numeric'});
}

function drawTitleBlock(ct,cv,c,logicalW,logicalH){
  cv=cv||canvas;c=c||ctx;
  // logicalW/logicalH are the drawing-space dimensions (before any scale transform)
  const W=logicalW||cv.width;
  const H=logicalH||cv.height;
  const tbY=H-TITLE_BLOCK_H,tbX=LEFT_MARGIN,tbW=W-LEFT_MARGIN-40,tbH=TITLE_BLOCK_H;
  c.fillStyle=ct.tbBg;c.fillRect(tbX,tbY,tbW,tbH);
  c.strokeStyle=ct.tbBorder;c.lineWidth=1;c.strokeRect(tbX,tbY,tbW,tbH);
  const logoW=tbLogoImage?140:0,detailW=340,midW=tbW-logoW-detailW;
  if(tbLogoImage){
    const lP=6,lH=tbH-lP*2,lW=logoW-lP*2;
    const ratio=Math.min(lW/tbLogoImage.width,lH/tbLogoImage.height);
    const dw=tbLogoImage.width*ratio,dh=tbLogoImage.height*ratio;
    c.drawImage(tbLogoImage,tbX+lP+(lW-dw)/2,tbY+lP+(lH-dh)/2,dw,dh);
    c.beginPath();c.moveTo(tbX+logoW,tbY);c.lineTo(tbX+logoW,tbY+tbH);c.strokeStyle=ct.tbBorder;c.stroke();
  }
  const midX=tbX+logoW+10;
  c.fillStyle=ct.tbText;c.textBaseline='middle';c.textAlign='left';
  c.font='bold 13px Arial';c.fillText(titleBlock.projectName||'Project Name',midX,tbY+tbH*0.22);
  c.font='11px Arial';c.fillText(titleBlock.drawingTitle||'Time-Chainage Programme',midX,tbY+tbH*0.50);
  c.font='10px Arial';c.fillStyle=ct.axisText;c.fillText('Scale: '+(titleBlock.scale||'NTS'),midX,tbY+tbH*0.78);
  const divX=tbX+logoW+midW;
  c.beginPath();c.moveTo(divX,tbY);c.lineTo(divX,tbY+tbH);c.strokeStyle=ct.tbBorder;c.lineWidth=1;c.stroke();
  const dX=divX+8,row=tbH/4,halfW=(tbW-(divX-tbX))/2;
  const details=[['Dwg No:',titleBlock.drawingNumber||'—'],['Revision:',titleBlock.revision||'—'],['Drawn by:',titleBlock.drawnBy||'—'],['Checked:',titleBlock.checkedBy||'—'],['Date:',formatDisplayDate(titleBlock.date)],['Data Date:',formatDisplayDate(titleBlock.dataDate)],[' ',' '],[' ',' ']];
  for(let i=0;i<4;i++){
    const rowY=tbY+row*i+row/2,c1=details[i],c2=details[i+4]||['',''];
    c.font='bold 9px Arial';c.fillStyle=ct.tbText;c.textAlign='left';c.fillText(c1[0],dX,rowY);
    c.font='9px Arial';c.fillText(c1[1],dX+58,rowY);
    c.font='bold 9px Arial';c.fillText(c2[0],dX+halfW,rowY);
    c.font='9px Arial';c.fillText(c2[1],dX+halfW+58,rowY);
  }
  for(let i=1;i<4;i++){c.beginPath();c.moveTo(divX,tbY+row*i);c.lineTo(tbX+tbW,tbY+row*i);c.strokeStyle=ct.tbBorder;c.lineWidth=0.5;c.stroke();}
  c.beginPath();c.moveTo(divX+halfW,tbY);c.lineTo(divX+halfW,tbY+tbH);c.strokeStyle=ct.tbBorder;c.lineWidth=0.5;c.stroke();
}

// ═══════════════════════════════════════════════════════════════════
//  EXPORT MODAL
// ═══════════════════════════════════════════════════════════════════
function openExportModal(){document.getElementById('exportModal').style.display='flex';document.getElementById('exportStatus').textContent='';}
function closeExportModal(){document.getElementById('exportModal').style.display='none';}

function runExport(){
  const opts=gatherExportOpts();
  const status=document.getElementById('exportStatus');
  status.textContent='Rendering…';

  if(opts.format==='excel'){exportToExcel();status.textContent='✅ Excel downloaded.';setTimeout(()=>closeExportModal(),1200);return;}
  if(opts.format==='tcplan'){downloadCurrentDiagram();status.textContent='✅ .tcplan downloaded.';setTimeout(()=>closeExportModal(),1200);return;}

  requestAnimationFrame(()=>{
    const oc=renderOffscreen(opts);
    if(!oc){status.textContent='❌ Set timeline first.';return;}

    // Build legend canvas(es) based on user selection
    const legendGroups    = opts.legendGroups    ? _buildLegendCanvas(oc.width,'groups')     : null;
    const legendActivities= opts.legendActivities ? _buildLegendCanvas(oc.width,'activities') : null;

    // Stitch all legend canvases into one combined legend if both selected
    const legendOc = _stitchLegendCanvases(oc.width, legendGroups, legendActivities);

    // Standalone legend download
    if(opts.legendStandalone && legendOc){
      const sl=document.createElement('a');
      sl.download=(titleBlock.projectName||'legend')+'-legend.png';
      sl.href=legendOc.toDataURL('image/png');
      sl.click();
    }

    if(opts.format==='png'){
      const combined = legendOc ? _stitchProgrammeAndLegend(oc, legendOc) : oc;
      const link=document.createElement('a');
      link.download=(titleBlock.projectName||'time-chainage')+'.png';
      link.href=combined.toDataURL('image/png');
      link.click();
      status.textContent='✅ PNG downloaded'+(legendOc?' (legend included)':'')+'.';

    } else if(opts.format==='pdf'){
      if(typeof window.jspdf==='undefined'){status.textContent='❌ PDF library not loaded.';return;}
      const{jsPDF}=window.jspdf;
      const sizeMap={a1:[594,841],a3:[297,420],a4:[210,297]};
      const[pw,ph]=sizeMap[opts.size]||[297,420];
      const isLand=opts.orientation==='landscape';
      const pdfW=isLand?Math.max(pw,ph):Math.min(pw,ph);
      const pdfH=isLand?Math.min(pw,ph):Math.max(pw,ph);
      const DPI=150;
      const renderW=Math.round(pdfW*DPI/25.4);
      const renderH=Math.round(pdfH*DPI/25.4);
      const pdfOc=renderOffscreen({...opts,resolution:1},renderW,renderH);
      if(!pdfOc){status.textContent='❌ Render failed.';return;}
      const margin=10,availW=pdfW-margin*2,availH=pdfH-margin*2;
      const cRatio=pdfOc.width/pdfOc.height,pRatio=availW/availH;
      let imgW,imgH;
      if(cRatio>pRatio){imgW=availW;imgH=availW/cRatio;}
      else{imgH=availH;imgW=availH*cRatio;}
      const pdf=new jsPDF({orientation:opts.orientation,unit:'mm',format:opts.size});
      pdf.addImage(pdfOc.toDataURL('image/png'),'PNG',margin+(availW-imgW)/2,margin+(availH-imgH)/2,imgW,imgH);
      // Legend page(s)
      if(legendOc){
        pdf.addPage(opts.size,opts.orientation);
        const lRatio=Math.min(availW/legendOc.width,(availH-16)/legendOc.height);
        const lW=legendOc.width*lRatio, lH=legendOc.height*lRatio;
        pdf.setFontSize(13);pdf.setFont(undefined,'bold');
        pdf.text('Legend — '+(titleBlock.projectName||'Programme'),margin,margin+6);
        pdf.addImage(legendOc.toDataURL('image/png'),'PNG',margin,margin+14,lW,lH);
      }
      pdf.save((titleBlock.projectName||'time-chainage')+'.pdf');
      status.textContent='✅ PDF downloaded'+(legendOc?' (legend on page 2)':'')+'.';
    }
    setTimeout(()=>closeExportModal(),1800);
  });
}

function _stitchLegendCanvases(w, groupsOc, activitiesOc){
  if(!groupsOc && !activitiesOc) return null;
  if(groupsOc && !activitiesOc) return groupsOc;
  if(!groupsOc && activitiesOc) return activitiesOc;
  // Both — stack vertically with a divider
  const DIVIDER=32;
  const combined=document.createElement('canvas');
  combined.width=w;
  combined.height=groupsOc.height+DIVIDER+activitiesOc.height;
  const c=combined.getContext('2d');
  c.drawImage(groupsOc,0,0);
  c.fillStyle='#e0e4e8';c.fillRect(0,groupsOc.height,w,DIVIDER);
  c.fillStyle='#2c3e50';c.font='bold 12px Arial';c.textAlign='left';c.textBaseline='middle';
  c.fillText('All Activities',12,groupsOc.height+DIVIDER/2);
  c.drawImage(activitiesOc,0,groupsOc.height+DIVIDER);
  return combined;
}

function _stitchProgrammeAndLegend(progOc, legendOc){
  const GAP=24, HEADER=44;
  const combined=document.createElement('canvas');
  combined.width=progOc.width;
  combined.height=progOc.height+GAP+HEADER+legendOc.height;
  const c=combined.getContext('2d');
  c.drawImage(progOc,0,0);
  // Legend section background
  c.fillStyle='#f0f4f8';
  c.fillRect(0,progOc.height+GAP,combined.width,HEADER+legendOc.height);
  // Divider line
  c.strokeStyle='#c0c8d4';c.lineWidth=2;
  c.beginPath();c.moveTo(0,progOc.height+GAP);c.lineTo(combined.width,progOc.height+GAP);c.stroke();
  // Legend header
  c.fillStyle='#2c3e50';c.font='bold 18px Arial';c.textAlign='left';c.textBaseline='middle';
  c.fillText('Legend — '+(titleBlock.projectName||'Programme'),16,progOc.height+GAP+HEADER/2);
  c.drawImage(legendOc,0,progOc.height+GAP+HEADER);
  return combined;
}

/** Build a legend canvas. mode = 'groups' | 'activities' */
function _buildLegendCanvas(targetW, mode='groups'){
  const cats=groups.filter(g=>g.parentId===null);
  if(!cats.length && mode==='groups') return null;
  if(!activities.length && mode==='activities') return null;

  const items=[];

  if(mode==='groups'){
    cats.forEach(cat=>{
      items.push({type:'cat',name:cat.name,color:cat.color||'#3498db'});
      groups.filter(g=>g.parentId===cat.id).forEach(sg=>{
        const sample=activities.find(a=>a.groups&&a.groups.includes(sg.id));
        const n=activities.filter(a=>a.groups&&a.groups.includes(sg.id)).length;
        items.push({type:'item',name:sg.name,shape:sample?.shape||'rect',
          color:sg.color||cat.color||'#3498db',count:n});
      });
    });
    if(!items.length) return null;
  } else {
    // Activities mode — grouped by category
    const shown=new Set();
    cats.forEach(cat=>{
      const catActs=activities.filter((a,i)=>!shown.has(i)&&a.groups&&a.groups.some(gid=>{
        const g=groups.find(x=>x.id===gid);return g&&(g.id===cat.id||g.parentId===cat.id);
      }));
      if(catActs.length){
        items.push({type:'cat',name:cat.name,color:cat.color||'#3498db'});
        catActs.forEach(a=>{shown.add(activities.indexOf(a));
          items.push({type:'item',name:a.name+(a.p6Id?` (${a.p6Id})`:''),
            shape:a.shape||'rect',color:getActivityDrawColor(a)});
        });
      }
    });
    const ungrouped=activities.filter((_,i)=>!shown.has(i));
    if(ungrouped.length){
      if(cats.length) items.push({type:'cat',name:'Ungrouped',color:'#95a5a6'});
      ungrouped.forEach(a=>items.push({type:'item',name:a.name+(a.p6Id?` (${a.p6Id})`:''),
        shape:a.shape||'rect',color:getActivityDrawColor(a)}));
    }
    if(!items.length) return null;
  }

  // Layout: 3 columns
  const COLS=3,ROW_H=28,SWATCH_W=42,PAD=12,ITEM_W=Math.floor((targetW-PAD*2)/COLS);
  const CAT_H=30;
  const rows=[];let curRow=[];
  items.forEach(item=>{
    if(item.type==='cat'){
      if(curRow.length){rows.push({type:'items',cells:curRow});curRow=[];}
      rows.push({type:'cat',item});
    } else {
      curRow.push(item);
      if(curRow.length===COLS){rows.push({type:'items',cells:curRow});curRow=[];}
    }
  });
  if(curRow.length) rows.push({type:'items',cells:curRow});

  const totalH=rows.reduce((h,r)=>h+(r.type==='cat'?CAT_H:ROW_H),0)+PAD*2;
  const lc=document.createElement('canvas');
  lc.width=targetW;lc.height=totalH;
  const c=lc.getContext('2d');
  c.fillStyle='#f8f9fa';c.fillRect(0,0,targetW,totalH);

  let y=PAD;
  rows.forEach(row=>{
    if(row.type==='cat'){
      c.fillStyle='#e8edf2';c.fillRect(PAD,y,targetW-PAD*2,CAT_H-4);
      c.fillStyle=row.item.color;c.beginPath();c.arc(PAD+10,y+CAT_H/2-2,6,0,Math.PI*2);c.fill();
      c.fillStyle='#1a1a2e';c.font='bold 13px Arial';c.textAlign='left';c.textBaseline='middle';
      c.fillText(row.item.name,PAD+22,y+CAT_H/2-2);
      y+=CAT_H;
    } else {
      row.cells.forEach((item,col)=>{
        const x=PAD+col*ITEM_W;
        c.fillStyle=col%2===0?'#fff':'#f4f6f8';c.fillRect(x,y,ITEM_W-4,ROW_H-3);
        // Shape swatch
        const sw=SWATCH_W-8,sh=16,sx2=x+4,sy2=y+(ROW_H-sh)/2;
        const cx2=sx2+sw/2,cy2=sy2+sh/2,r=6;
        c.fillStyle=item.color;c.strokeStyle=item.color;c.globalAlpha=0.9;
        switch(item.shape){
          case 'rect':    c.fillRect(sx2,sy2+2,sw,sh-4);break;
          case 'line':    c.lineWidth=2;c.beginPath();c.moveTo(sx2,cy2);c.lineTo(sx2+sw,cy2);c.stroke();break;
          case 'circle':  c.beginPath();c.arc(cx2,cy2,r,0,Math.PI*2);c.fill();break;
          case 'diamond': c.beginPath();c.moveTo(cx2,sy2);c.lineTo(sx2+sw,cy2);c.lineTo(cx2,sy2+sh);c.lineTo(sx2,cy2);c.closePath();c.fill();break;
          case 'triangle':c.beginPath();c.moveTo(cx2,sy2);c.lineTo(sx2+sw,sy2+sh);c.lineTo(sx2,sy2+sh);c.closePath();c.fill();break;
          case 'star':{const pts=5,ri=r*.42;c.beginPath();for(let i=0;i<pts*2;i++){const a2=((i*Math.PI)/pts)-Math.PI/2;const rr=i%2===0?r:ri;i===0?c.moveTo(cx2+rr*Math.cos(a2),cy2+rr*Math.sin(a2)):c.lineTo(cx2+rr*Math.cos(a2),cy2+rr*Math.sin(a2));}c.closePath();c.fill();break;}
          case 'flag':    c.lineWidth=1.5;c.beginPath();c.moveTo(cx2,sy2);c.lineTo(cx2,sy2+sh);c.stroke();c.beginPath();c.moveTo(cx2,sy2);c.lineTo(sx2+sw,cy2-2);c.lineTo(cx2,cy2+2);c.closePath();c.fill();break;
          default:        c.fillRect(sx2,sy2+2,sw,sh-4);
        }
        c.globalAlpha=1;
        c.fillStyle='#1a1a2e';c.font='11px Arial';c.textAlign='left';c.textBaseline='middle';
        const maxLblW=ITEM_W-SWATCH_W-16;let lbl=item.name;
        while(lbl.length>3&&c.measureText(lbl).width>maxLblW)lbl=lbl.slice(0,-1);
        if(lbl!==item.name)lbl=lbl.slice(0,-1)+'…';
        c.fillText(lbl,x+SWATCH_W,y+ROW_H/2);
        if(item.count!=null){c.fillStyle='#888';c.font='10px Arial';c.textAlign='right';c.fillText(item.count+'×',x+ITEM_W-6,y+ROW_H/2);}
      });
      y+=ROW_H;
    }
  });
  return lc;
}
function renderOffscreen(opts,forceW,forceH){
  if(!timelineStart||!timelineEnd)return null;
  const ct=getChartTheme();
  const hasTB=opts.showTitleBlock&&hasTitleBlock();
  const scale=opts.resolution||1;
  const srcW=forceW||(canvas.parentElement.offsetWidth||800);
  const baseH=forceH||getRequiredCanvasHeight(hasTB);
  const oc=document.createElement('canvas');
  oc.width=srcW*(forceW?1:scale);
  oc.height=baseH*(forceH?1:scale);
  const c=oc.getContext('2d');
  if(!forceW)c.scale(scale,scale);

  const topM=BASE_TOP_MARGIN+YEAR_BAND_H+(opts.showImage&&chainageImage?chainageImageH+RESIZE_HANDLE_H:0);
  const gridW=srcW-LEFT_MARGIN-40;
  const gridH=baseH-topM-BOTTOM_MARGIN-(hasTB?TITLE_BLOCK_H:0);

  function oY(d){const tot=timelineEnd-timelineStart,off=d-timelineStart;return topM+(off/tot)*gridH;}
  function oX(ch){return LEFT_MARGIN+(ch-minChainage)/(maxChainage-minChainage)*gridW;}

  c.fillStyle=ct.canvasBg;c.fillRect(0,0,srcW,baseH);

  // Year bands (Feature 8) — pass full baseH so it can compute gridH correctly
  drawYearBandsOn(c,ct,srcW,gridW,baseH);

  // Chainage axis labels
  c.fillStyle=ct.axisText;c.textAlign='center';c.textBaseline='top';c.font='11px Arial';
  for(let ch=minChainage;ch<=maxChainage;ch+=chainageSpacing)c.fillText(ch,oX(ch),6);

  // Reference image
  if(opts.showImage&&chainageImage){
    const imgY=BASE_TOP_MARGIN+YEAR_BAND_H;
    c.save();c.beginPath();c.rect(LEFT_MARGIN,imgY,gridW,chainageImageH);c.clip();
    c.drawImage(chainageImage,LEFT_MARGIN,imgY,gridW,chainageImageH);
    c.restore();c.strokeStyle=ct.imgBorder;c.lineWidth=1;c.strokeRect(LEFT_MARGIN,imgY,gridW,chainageImageH);
  }

  // Zones
  chainageZones.forEach(z=>{
    const x1=oX(Math.max(z.startCh,minChainage)),x2=oX(Math.min(z.endCh,maxChainage));
    c.save();c.globalAlpha=z.alpha;c.fillStyle=z.color;c.fillRect(x1,topM,x2-x1,gridH);c.restore();
    if(!z.hideLabel){
      const labelY=topM-12, labelX=(x1+x2)/2;
      c.save();c.globalAlpha=0.9;c.fillStyle=z.color;
      const tw=c.measureText(z.name).width;
      c.fillRect(labelX-tw/2-4,labelY-7,tw+8,14);
      c.fillStyle='#fff';c.font='bold 9px Arial';c.textAlign='center';c.textBaseline='middle';
      c.fillText(z.name,labelX,labelY);
      c.restore();
    }
  });

  // Date grid lines + month labels — every month, no skipping
  const ticks=buildDateTicks();
  ticks.forEach((tick) => {
    const y = oY(tick);
    const isJan = tick.getMonth() === 0;
    c.strokeStyle = isJan ? ct.yearBandBorder : ct.gridH;
    c.lineWidth   = isJan ? 1.5 : 1;
    c.beginPath(); c.moveTo(LEFT_MARGIN, y); c.lineTo(srcW - 40, y); c.stroke();
    c.lineWidth = 1;
    c.fillStyle    = isJan ? ct.yearBandText : ct.axisText;
    c.font         = isJan ? 'bold 10px Arial' : '10px Arial';
    c.textAlign    = 'right';
    c.textBaseline = 'middle';
    c.fillText(tick.toLocaleDateString(undefined, {month:'short'}), LEFT_MARGIN - 4, y);
  });

  // Vertical chainage grid
  for(let ch=minChainage;ch<=maxChainage;ch+=chainageSpacing){
    const x=oX(ch);c.strokeStyle=ct.gridV;c.lineWidth=1;c.beginPath();c.moveTo(x,topM);c.lineTo(x,baseH-BOTTOM_MARGIN-(hasTB?TITLE_BLOCK_H:0));c.stroke();
  }

  // Date markers — drawn AFTER activities (see after activity loop below)
  // Data date line — also drawn after activities

  // Baselines — draw ghost overlay for each visible baseline
  if(opts.showBaseline)baselines.filter(b=>b.visible).forEach(b=>{
    b.activities.forEach(ba=>{
      if(!ba.start||!ba.end)return;
      const y1=oY(new Date(ba.start)),y2=oY(new Date(ba.end));
      const x1=oX(ba.startCh),x2=oX(ba.endCh);
      const shape=ba.shape||'rect';
      c.save();
      c.globalAlpha=0.22;
      c.fillStyle=b.tint;
      c.strokeStyle=b.tint;
      drawShapeOn(c,shape,x1,y1,x2,y2);
      c.globalAlpha=0.55;
      c.lineWidth=1.5;
      c.setLineDash([5,3]);
      // Draw outline for all shapes
      drawOutlineOn(c,{shape,outlineStyle:'dashed',outlineColor:b.tint,outlineWidth:1.5},x1,y1,x2,y2);
      c.setLineDash([]);
      c.restore();
    });
  });

  // Activities
  const offCritSet = opts.fsCritical ? (cpHighlightSet.size>0?cpHighlightSet:computeCriticalPath()) : new Set();
  activities.forEach((a,i)=>{
    if(!a.start||!a.end||isNaN(a.start)||isNaN(a.end))return;
    if(!activityInViewFilter(a))return;
    const y1=oY(a.start),y2=oY(a.end),x1=oX(a.startCh),x2=oX(a.endCh);
    // Critical path ring in fullscreen
    if(offCritSet.has(i)){
      c.save();c.strokeStyle='#e74c3c';c.lineWidth=3;c.strokeRect(x1-2,y1-2,x2-x1+4,y2-y1+4);c.restore();
    }
    const drawColor=getActivityDrawColor(a);
    c.globalAlpha=a.alpha||1;c.fillStyle=drawColor;c.strokeStyle=drawColor;
    drawShapeOn(c,a.shape,x1,y1,x2,y2);
    drawOutlineOn(c,a,x1,y1,x2,y2);
    if(opts.showProgress&&(a.progress||0)>0)drawProgressOn(c,a,x1,y1,x2,y2);
    if(a.label&&a.label!=='no'){
      c.globalAlpha=1;c.fillStyle=ct.labelText;c.font='bold 12px Arial';c.textBaseline='middle';
      const cx=(x1+x2)/2,midY=(y1+y2)/2;let lx;
      if(a.label==='left'){c.textAlign='right';lx=isMilestone(a.shape)?cx-MILESTONE_SIZE-4:Math.min(x1,x2)-4;}
      else if(a.label==='right'){c.textAlign='left';lx=isMilestone(a.shape)?cx+MILESTONE_SIZE+4:Math.max(x1,x2)+4;}
      else{c.textAlign='center';lx=cx;}
      c.fillText(a.name,lx,midY);
    }
    c.globalAlpha=1;a.coords={x1,y1,x2,y2};
  });

  // Dependency links — elbow-routed, correct time-axis edges
  if(opts.showLinks){
    activityLinks.forEach(l=>{
      const from=activities[l.fromIdx],to=activities[l.toIdx];
      if(!from||!to||!from.coords||!to.coords)return;
      const fc=from.coords,tc=to.coords;
      const fcX=(fc.x1+fc.x2)/2, tcX=(tc.x1+tc.x2)/2;
      let sx,sy,ex,ey;
      if(l.type==='FS'){sx=fcX;sy=Math.max(fc.y1,fc.y2);ex=tcX;ey=Math.min(tc.y1,tc.y2);}
      else if(l.type==='SS'){sx=fcX;sy=Math.min(fc.y1,fc.y2);ex=tcX;ey=Math.min(tc.y1,tc.y2);}
      else if(l.type==='FF'){sx=fcX;sy=Math.max(fc.y1,fc.y2);ex=tcX;ey=Math.max(tc.y1,tc.y2);}
      else{sx=fcX;sy=Math.min(fc.y1,fc.y2);ex=tcX;ey=Math.max(tc.y1,tc.y2);}
      drawArrowOn(c,sx,sy,ex,ey,'rgba(100,100,100,0.7)');
      const lagStr=l.lag?`${l.lag>0?'+':''}${l.lag}d`:'';
      drawLinkLabel(c,sx,sy,ex,ey,l.type+(lagStr?` ${lagStr}`:''));
    });
  }

  // ── Risk indicators on top of activities ──
  if(opts.showRiskOverlays!==false) _drawRiskIndicators(c,gridW,gridH,ct,opts.showRiskOverlays!==false);
  // ── Risk P-lines ──
  if(opts.showRiskOverlays!==false) _drawRiskPLines(c,srcW,gridH,ct,opts.showRiskOverlays!==false);

  // ── Date markers drawn ON TOP of activities ──
  if(opts.showDateMarkers!==false){
    dateMarkers.forEach(m=>{
      const d=parseDateLocal(m.date);if(!d)return;const y=oY(d);
      c.save();c.strokeStyle=m.color;c.lineWidth=2;c.setLineDash([6,4]);
      c.beginPath();c.moveTo(LEFT_MARGIN,y);c.lineTo(srcW-40,y);c.stroke();c.setLineDash([]);
      if(!m.hideLabel){
        c.font='bold 10px Arial';
        const tw=c.measureText(m.name).width;
        const px=LEFT_MARGIN+6,py=y-18;
        c.globalAlpha=0.88;c.fillStyle=m.color;
        c.beginPath();c.roundRect?c.roundRect(px-3,py,tw+10,15,4):c.fillRect(px-3,py,tw+10,15);
        c.fill();
        c.globalAlpha=1;c.fillStyle='#fff';c.textAlign='left';c.textBaseline='top';
        c.fillText(m.name,px+2,py+2);
      }
      c.restore();
    });
  }

  // Data date line on top — respects showDateMarkers opt
  if(opts.showDateMarkers!==false && titleBlock.dataDate){
    const d=parseDateLocal(titleBlock.dataDate);
    if(d){
      const y=oY(d);
      c.save();c.strokeStyle='#e74c3c';c.lineWidth=2.5;c.setLineDash([5,4]);
      c.beginPath();c.moveTo(LEFT_MARGIN,y);c.lineTo(srcW-40,y);c.stroke();c.setLineDash([]);
      c.font='bold 11px Arial';
      const lbl='▶ Data Date';const tw=c.measureText(lbl).width;
      c.fillStyle='#e74c3c';c.globalAlpha=0.9;c.fillRect(srcW-44-tw-8,y-17,tw+10,15);
      c.globalAlpha=1;c.fillStyle='#fff';c.textAlign='right';c.textBaseline='bottom';
      c.fillText(lbl,srcW-48,y-4);
      c.restore();
    }
  }

  if(hasTB)drawTitleBlock(ct,oc,c,srcW,baseH);
  return oc;
}

// ═══════════════════════════════════════════════════════════════════
//  YEAR BANDS (Feature 8)
//  TIME IS ON THE Y AXIS. Chainage is on the X axis.
//  Year bands are horizontal stripes across the full chart width.
//  A narrow column on the LEFT of the grid shows the year number,
//  sitting just above the first month tick of that year.
// ═══════════════════════════════════════════════════════════════════
function drawYearBandsOn(c, ct, canvasW, gridW, canvasH){
  if(!timelineStart||!timelineEnd) return;

  // Grid geometry — must match drawChart / renderOffscreen exactly
  const imgExtra = chainageImage ? chainageImageH + RESIZE_HANDLE_H : 0;
  const topM     = BASE_TOP_MARGIN + YEAR_BAND_H + imgExtra;
  const hasTB    = hasTitleBlock();
  const gridH    = canvasH - topM - BOTTOM_MARGIN - (hasTB ? TITLE_BLOCK_H : 0);
  const totalMs  = timelineEnd - timelineStart;

  // Convert a Date → Y pixel
  function yOf(d){ return topM + ((d - timelineStart) / totalMs) * gridH; }

  const startYear = timelineStart.getFullYear();
  const endYear   = timelineEnd.getFullYear();

  // ── 1. Alternating horizontal band fills (full chart width) ──
  for(let yr = startYear; yr <= endYear; yr++){
    const bandTop    = new Date(Math.max(new Date(yr,   0, 1).getTime(), timelineStart.getTime()));
    const bandBottom = new Date(Math.min(new Date(yr+1, 0, 1).getTime(), timelineEnd.getTime()));
    const y1 = yOf(bandTop);
    const y2 = yOf(bandBottom);

    if(yr % 2 === 0){
      c.save();
      c.globalAlpha = 0.13;
      c.fillStyle   = ct.yearBandBg;
      c.fillRect(LEFT_MARGIN, y1, gridW, y2 - y1);
      c.restore();
    }

    // Dashed separator at the top of each year (except the very first)
    if(yr > startYear){
      c.save();
      c.strokeStyle = ct.yearBandBorder;
      c.lineWidth   = 1.5;
      c.setLineDash([8, 4]);
      c.beginPath();
      c.moveTo(LEFT_MARGIN, y1);
      c.lineTo(canvasW - 40, y1);
      c.stroke();
      c.setLineDash([]);
      c.restore();
    }
  }

  // ── 2. Year label — large, bold, vertically centred within each year's band ──
  for(let yr = startYear; yr <= endYear; yr++){
    const bandTop    = new Date(Math.max(new Date(yr,   0, 1).getTime(), timelineStart.getTime()));
    const bandBottom = new Date(Math.min(new Date(yr+1, 0, 1).getTime(), timelineEnd.getTime()));
    const y1 = yOf(bandTop);
    const y2 = yOf(bandBottom);
    const bandPx = y2 - y1;
    if(bandPx < 12) continue;

    // Clamp label to within the visible band
    const labelY = Math.min(Math.max(y1 + bandPx * 0.5, y1 + 10), y2 - 10);

    c.save();
    c.font         = 'bold 16px Arial';
    c.fillStyle    = ct.yearBandText;
    c.globalAlpha  = 0.9;
    c.textAlign    = 'center';
    c.textBaseline = 'middle';
    // Sit in the leftmost ~48px of the margin, leaving room for month labels
    c.fillText(String(yr), 24, labelY);
    c.restore();
  }

  // ── 3. Top strip — subtle shaded bar only, NO text (removed per user request) ──
  c.save();
  c.fillStyle   = ct.yearBandBg;
  c.globalAlpha = 0.35;
  c.fillRect(LEFT_MARGIN, BASE_TOP_MARGIN, gridW, YEAR_BAND_H);
  c.restore();
  c.save();
  c.strokeStyle = ct.yearBandBorder;
  c.lineWidth   = 1;
  c.globalAlpha = 0.5;
  c.strokeRect(LEFT_MARGIN, BASE_TOP_MARGIN, gridW, YEAR_BAND_H);
  c.restore();
}

// ═══════════════════════════════════════════════════════════════════
//  REVISIONS
// ═══════════════════════════════════════════════════════════════════
function saveRevision(){
  const name=document.getElementById('revisionName').value.trim();if(!name){alert('Enter a name.');return;}
  revisions.push({name,timestamp:new Date().toISOString(),data:serialiseDiagram()});
  document.getElementById('revisionName').value='';
  renderRevisionList();
}
function renderRevisionList(){
  const el=document.getElementById('revisionListInline');if(!el)return;
  if(!revisions.length){el.innerHTML='<p style="color:var(--text-muted);font-size:13px">No revisions saved.</p>';return;}
  let html='<table class="activityTable"><thead><tr><th>Name</th><th>Saved</th><th>Actions</th></tr></thead><tbody>';
  revisions.forEach((r,i)=>{
    const d=new Date(r.timestamp).toLocaleString(undefined,{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
    html+=`<tr><td><strong>${r.name}</strong></td><td style="font-size:12px;color:var(--text-muted)">${d}</td>
    <td><button onclick="restoreRevision(${i})">⏪ Restore</button>
        <button onclick="deleteRevision(${i})" class="deleteBtn" style="margin-left:4px">Delete</button></td></tr>`;
  });
  html+='</tbody></table>';el.innerHTML=html;
}
function restoreRevision(i){if(!confirm(`Restore "${revisions[i].name}"?`))return;deserialiseDiagram(revisions[i].data);}
function deleteRevision(i){if(!confirm(`Delete "${revisions[i].name}"?`))return;revisions.splice(i,1);renderRevisionList();}

// ═══════════════════════════════════════════════════════════════════
//  BASELINES
// ═══════════════════════════════════════════════════════════════════
function addBaseline(){
  const name=document.getElementById('baselineName').value.trim();if(!name){alert('Enter a name.');return;}
  if(!activities.length){alert('No activities.');return;}
  const tint=BASELINE_TINTS[baselines.length%BASELINE_TINTS.length];
  baselines.push({id:genId(),name,timestamp:new Date().toISOString(),visible:true,tint,
    activities:activities.map(a=>({p6Id:a.p6Id,name:a.name,start:a.start?a.start.toISOString():null,end:a.end?a.end.toISOString():null,startCh:a.startCh,endCh:a.endCh,shape:a.shape,color:a.color}))});
  document.getElementById('baselineName').value='';
  renderBaselineList();drawChart();
}
function toggleBaselineVisibility(id){const b=baselines.find(x=>x.id===id);if(b){b.visible=!b.visible;renderBaselineList();drawChart();}}
function renameBaseline(id){const b=baselines.find(x=>x.id===id);if(!b)return;const n=prompt('Rename:',b.name);if(n&&n.trim()){b.name=n.trim();renderBaselineList();}}
function deleteBaseline(id){const b=baselines.find(x=>x.id===id);if(!b)return;if(!confirm(`Delete "${b.name}"?`))return;baselines=baselines.filter(x=>x.id!==id);renderBaselineList();drawChart();}
function renderBaselineList(){
  const el=document.getElementById('baselineListContainer');if(!el)return;
  if(!baselines.length){el.innerHTML='<p style="color:var(--text-muted);font-size:13px">No baselines saved.</p>';return;}
  let html='<table class="activityTable"><thead><tr><th>Colour</th><th>Name</th><th>Saved</th><th>Acts</th><th>Visible</th><th>Actions</th></tr></thead><tbody>';
  baselines.forEach(b=>{
    const d=new Date(b.timestamp).toLocaleString(undefined,{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
    html+=`<tr>
      <td>
        <input type="color" value="${b.tint}" title="Click to change baseline colour"
          onchange="updateBaselineTint('${b.id}',this.value)"
          style="width:32px;height:24px;padding:1px;border:1px solid var(--border);border-radius:4px;cursor:pointer;margin:0">
      </td>
      <td><strong>${b.name}</strong></td>
      <td style="font-size:12px;color:var(--text-muted)">${d}</td>
      <td>${b.activities.length}</td>
      <td><label class="toggle-switch"><input type="checkbox" ${b.visible?'checked':''} onchange="toggleBaselineVisibility('${b.id}')"><span class="toggle-slider"></span></label></td>
      <td>
        <button onclick="renameBaseline('${b.id}')" style="margin:0;padding:3px 8px;font-size:11px">✏️ Rename</button>
        <button onclick="deleteBaseline('${b.id}')" class="deleteBtn" style="margin:2px 0 0 2px;padding:3px 8px;font-size:11px">✕ Delete</button>
      </td>
    </tr>`;
  });
  html+='</tbody></table>';el.innerHTML=html;
}

function updateBaselineTint(id,color){
  const b=baselines.find(x=>x.id===id);if(!b)return;
  b.tint=color;drawChart();
}

// ═══════════════════════════════════════════════════════════════════
//  GROUPS — Feature 7: collapse, reorder, edit name, colour override,
//           chart filter checkboxes, group colour applied to activities
// ═══════════════════════════════════════════════════════════════════

function createCategory(){
  const name=document.getElementById('categoryName').value.trim();if(!name)return;
  const color=document.getElementById('categoryColor')?.value||'#3498db';
  groups.push({id:genId(),name,parentId:null,color,useColor:false,visible:true,collapsed:false});
  document.getElementById('categoryName').value='';
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();
}

function createSubgroup(){
  const name=document.getElementById('subgroupName').value.trim();
  const parentId=document.getElementById('subgroupParent').value;
  if(!name){alert('Enter a sub-group name.');return;}
  if(!parentId){alert('Select a category.');return;}
  const color=document.getElementById('subgroupColor')?.value||'#e74c3c';
  groups.push({id:genId(),name,parentId,color,useColor:false,visible:true,collapsed:false});
  document.getElementById('subgroupName').value='';
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();
}

function deleteGroup(id){
  const childIds=groups.filter(g=>g.parentId===id).map(g=>g.id);
  groups=groups.filter(g=>g.id!==id&&g.parentId!==id);
  activities.forEach(a=>{if(a.groups)a.groups=a.groups.filter(gid=>gid!==id&&!childIds.includes(gid));});
  if(highlightedGroupId===id||childIds.includes(highlightedGroupId))highlightedGroupId=null;
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();drawChart();
}

// Move group up/down in the array (among siblings)
function moveGroup(id,dir){
  const g=groups.find(x=>x.id===id);if(!g)return;
  const siblings=groups.filter(x=>x.parentId===g.parentId);
  const sibIdx=siblings.findIndex(x=>x.id===id);
  const newSibIdx=sibIdx+dir;
  if(newSibIdx<0||newSibIdx>=siblings.length)return;
  // Find actual indices in groups array
  const gIdx=groups.findIndex(x=>x.id===id);
  const swapId=siblings[newSibIdx].id;
  const swapIdx=groups.findIndex(x=>x.id===swapId);
  // Swap
  [groups[gIdx],groups[swapIdx]]=[groups[swapIdx],groups[gIdx]];
  renderGroupTree();renderGroupPickerInEditor();drawChart();
}

// Toggle collapse of a category
function toggleGroupCollapse(id){
  const g=groups.find(x=>x.id===id);if(!g)return;
  g.collapsed=!g.collapsed;
  renderGroupTree();
}

// Open edit dialog for group/category
function openGroupEditDialog(id){
  groupEditId=id;
  const g=groups.find(x=>x.id===id);if(!g)return;
  document.getElementById('groupEditName').value=g.name;
  document.getElementById('groupEditColor').value=g.color||'#3498db';
  document.getElementById('groupEditUseColor').checked=!!g.useColor;
  document.getElementById('groupEditDialog').style.display='flex';
}
function confirmEditGroup(){
  const g=groups.find(x=>x.id===groupEditId);if(!g){closeGroupEditDialog();return;}
  pushUndo();
  g.name=document.getElementById('groupEditName').value.trim()||g.name;
  g.color=document.getElementById('groupEditColor').value;
  g.useColor=document.getElementById('groupEditUseColor').checked;
  closeGroupEditDialog();
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();drawChart();renderActivityTable();
}
function closeGroupEditDialog(){document.getElementById('groupEditDialog').style.display='none';groupEditId=null;}

// Toggle chart visibility
function toggleGroupVisible(id){
  const g=groups.find(x=>x.id===id);if(!g)return;
  g.visible=!g.visible;
  renderGroupTree();drawChart();
}

function checkAllGroups(){groups.forEach(g=>g.visible=true);renderGroupTree();drawChart();}
function uncheckAllGroups(){groups.forEach(g=>g.visible=false);renderGroupTree();drawChart();}

function highlightGroup(id){
  highlightedGroupId=(highlightedGroupId===id)?null:id;
  renderGroupTree();drawChart();
}
function getHighlightedGroupIds(){
  if(!highlightedGroupId)return null;
  const g=groups.find(x=>x.id===highlightedGroupId);if(!g)return null;
  if(g.parentId===null){const children=groups.filter(x=>x.parentId===highlightedGroupId).map(x=>x.id);return new Set([highlightedGroupId,...children]);}
  return new Set([highlightedGroupId]);
}
function activityIsHighlighted(a){
  const hids=getHighlightedGroupIds();if(!hids)return true;
  if(!a.groups||!a.groups.length)return false;
  return a.groups.some(gid=>hids.has(gid));
}

/** Get the draw colour for an activity, respecting group colour overrides */
function getActivityDrawColor(a){
  if(!a.groups||!a.groups.length)return a.color;
  // Check if any assigned group has useColor on — prefer subgroup over category
  for(const gid of a.groups){
    const g=groups.find(x=>x.id===gid);
    if(g&&g.useColor&&g.color)return g.color;
  }
  // Check parent categories
  for(const gid of a.groups){
    const sg=groups.find(x=>x.id===gid);
    if(sg&&sg.parentId){
      const cat=groups.find(x=>x.id===sg.parentId);
      if(cat&&cat.useColor&&cat.color)return cat.color;
    }
  }
  return a.color;
}

/** Is this activity visible based on group chart filter? */
function activityVisibleByGroup(a){
  // No groups defined at all → show everything
  if(!groups.length) return true;

  // How many groups are currently unchecked (hidden)?
  const hiddenGroups = groups.filter(g=>!g.visible);

  // Nothing unchecked → no filter active → show everything
  if(!hiddenGroups.length) return true;

  // Activity has NO group assigned → always show (ungrouped items not affected by group filter)
  if(!a.groups||!a.groups.length) return true;

  // Activity is visible if ANY of its assigned groups is checked (visible)
  return a.groups.some(gid=>{
    const g=groups.find(x=>x.id===gid);
    if(!g) return true; // unknown group — show by default
    if(g.parentId===null){
      // Category — visible if category itself is checked
      // AND not all its subgroups are unchecked
      const subs=groups.filter(sg=>sg.parentId===g.id);
      if(!subs.length) return g.visible;
      // If the category is checked AND at least one of its subgroups is checked, show
      return g.visible && subs.some(sg=>sg.visible);
    }
    // Subgroup — visible only if the subgroup itself is checked
    return g.visible;
  });
}

// ── Group drag-and-drop state ──
let groupDragId   = null;   // id being dragged
let groupDragOver = null;   // id of item being hovered over
let groupDragMode = null;   // 'before' | 'after' | 'into' (reparent subgroup into another cat)

function renderGroupTree(){
  const div=document.getElementById('groupTree');if(!div)return;
  div.innerHTML='';
  const categories=groups.filter(g=>g.parentId===null);
  if(!categories.length){
    div.innerHTML='<p style="color:var(--text-muted);font-size:13px">No categories yet. Add one above.</p>';
    return;
  }

  categories.forEach((cat)=>{
    const subgroups=groups.filter(g=>g.parentId===cat.id);
    const isHighlighted=highlightedGroupId===cat.id;

    // ── Category row ──
    const catEl=document.createElement('div');
    catEl.className='group-category'+(isHighlighted?' group-highlighted':'');
    catEl.dataset.id=cat.id;
    catEl.dataset.type='category';
    catEl.draggable=true;

    catEl.innerHTML=`
      <span class="group-drag-handle" title="Drag to reorder">⠿</span>
      <input type="checkbox" class="group-visibility-check" ${cat.visible?'checked':''}
        onchange="toggleGroupVisible('${cat.id}')" title="Show/hide on chart">
      <button class="group-collapse-btn" onclick="toggleGroupCollapse('${cat.id}')" title="${cat.collapsed?'Expand':'Collapse'}">
        ${cat.collapsed?'▶':'▼'}
      </button>
      <span class="group-color-dot" style="background:${cat.color||'#3498db'}" title="${cat.useColor?'Colour override ON':'Colour override off'}"></span>
      ${cat.useColor?'<span style="font-size:11px;color:#f39c12;flex-shrink:0" title="Colour override active">★</span>':'<span style="font-size:11px;color:transparent;flex-shrink:0">★</span>'}
      <span class="group-cat-name" onclick="highlightGroup('${cat.id}')" title="Click to highlight on chart">📂 ${cat.name}</span>
      <div style="display:flex;gap:3px;flex-shrink:0">
        <button onclick="openGroupEditDialog('${cat.id}')" style="margin:0;padding:2px 7px;font-size:11px" title="Edit">✏️</button>
        <button onclick="deleteGroup('${cat.id}')" class="deleteBtn" style="margin:0;padding:2px 7px;font-size:11px" title="Delete">✕</button>
      </div>
    `;
    _attachGroupDragListeners(catEl, cat.id, 'category');
    div.appendChild(catEl);

    // ── Subgroup rows ──
    const sgContainer=document.createElement('div');
    sgContainer.className='group-subgroups-container'+(cat.collapsed?' collapsed':'');
    sgContainer.style.maxHeight=cat.collapsed?'0':(subgroups.length*48+8)+'px';

    subgroups.forEach((sg)=>{
      const isSgHighlighted=highlightedGroupId===sg.id;
      const sgEl=document.createElement('div');
      sgEl.className='group-subgroup'+(isSgHighlighted?' group-highlighted':'');
      sgEl.dataset.id=sg.id;
      sgEl.dataset.type='subgroup';
      sgEl.dataset.parentId=cat.id;
      sgEl.draggable=true;
      sgEl.innerHTML=`
        <span class="group-drag-handle" title="Drag to reorder or move to another category">⠿</span>
        <input type="checkbox" class="group-visibility-check" ${sg.visible?'checked':''}
          onchange="toggleGroupVisible('${sg.id}')" title="Show/hide on chart">
        <span class="group-color-dot" style="background:${sg.color||'#e74c3c'}" title="${sg.useColor?'Colour override ON':'Colour override off'}"></span>
        ${sg.useColor?'<span style="font-size:11px;color:#f39c12;flex-shrink:0" title="Colour override active">★</span>':'<span style="font-size:11px;color:transparent;flex-shrink:0">★</span>'}
        <span class="group-sg-name" onclick="highlightGroup('${sg.id}')" title="Click to highlight on chart">└ ${sg.name}</span>
        <div style="display:flex;gap:3px;flex-shrink:0">
          <button onclick="openGroupEditDialog('${sg.id}')" style="margin:0;padding:2px 7px;font-size:11px" title="Edit">✏️</button>
          <button onclick="deleteGroup('${sg.id}')" class="deleteBtn" style="margin:0;padding:2px 7px;font-size:11px" title="Delete">✕</button>
        </div>
      `;
      _attachGroupDragListeners(sgEl, sg.id, 'subgroup');
      sgContainer.appendChild(sgEl);
    });
    div.appendChild(sgContainer);
  });
}

function _attachGroupDragListeners(el, id, type){
  el.addEventListener('dragstart', e=>{
    groupDragId=id;
    e.dataTransfer.effectAllowed='move';
    setTimeout(()=>el.classList.add('group-dragging'),0);
  });
  el.addEventListener('dragend', ()=>{
    el.classList.remove('group-dragging');
    // Clear all drop indicators
    document.querySelectorAll('.group-drop-above,.group-drop-below,.group-drop-into')
      .forEach(x=>{x.classList.remove('group-drop-above','group-drop-below','group-drop-into');});
    groupDragId=null;groupDragOver=null;groupDragMode=null;
  });
  el.addEventListener('dragover', e=>{
    e.preventDefault();
    if(!groupDragId||groupDragId===id)return;
    const rect=el.getBoundingClientRect();
    const midY=rect.top+rect.height/2;
    const dragGrp=groups.find(g=>g.id===groupDragId);
    const overGrp=groups.find(g=>g.id===id);
    // Clear others
    document.querySelectorAll('.group-drop-above,.group-drop-below,.group-drop-into')
      .forEach(x=>{x.classList.remove('group-drop-above','group-drop-below','group-drop-into');});
    groupDragOver=id;
    // Dragging subgroup over a category → reparent
    if(dragGrp&&dragGrp.parentId!==null&&overGrp&&overGrp.parentId===null){
      el.classList.add('group-drop-into');
      groupDragMode='into';
    } else if(e.clientY<midY){
      el.classList.add('group-drop-above');
      groupDragMode='before';
    } else {
      el.classList.add('group-drop-below');
      groupDragMode='after';
    }
  });
  el.addEventListener('dragleave', ()=>{
    el.classList.remove('group-drop-above','group-drop-below','group-drop-into');
  });
  el.addEventListener('drop', e=>{
    e.preventDefault();
    if(!groupDragId||groupDragId===id){return;}
    _applyGroupDrop(groupDragId, id, groupDragMode);
    el.classList.remove('group-drop-above','group-drop-below','group-drop-into');
  });
}

function _applyGroupDrop(dragId, overId, mode){
  const dragGrp=groups.find(g=>g.id===dragId);
  const overGrp=groups.find(g=>g.id===overId);
  if(!dragGrp||!overGrp)return;

  // Reparent: move subgroup into a different category
  if(mode==='into'&&overGrp.parentId===null&&dragGrp.parentId!==null){
    dragGrp.parentId=overId;
    renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();drawChart();
    return;
  }

  // Reorder within same scope (both categories, or both subgroups of same parent)
  const sameScopeCategories = dragGrp.parentId===null && overGrp.parentId===null;
  const sameScopeSubs = dragGrp.parentId!==null && overGrp.parentId!==null && dragGrp.parentId===overGrp.parentId;

  if(sameScopeCategories||sameScopeSubs){
    const dragIdx=groups.findIndex(g=>g.id===dragId);
    const overIdx=groups.findIndex(g=>g.id===overId);
    groups.splice(dragIdx,1);
    const newOverIdx=groups.findIndex(g=>g.id===overId);
    const insertAt=mode==='before'?newOverIdx:newOverIdx+1;
    groups.splice(insertAt,0,dragGrp);
    renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();drawChart();
    return;
  }

  // Move subgroup to a different parent — position before/after a sibling subgroup
  if(dragGrp.parentId!==null && overGrp.parentId!==null && dragGrp.parentId!==overGrp.parentId){
    dragGrp.parentId=overGrp.parentId;
    const dragIdx=groups.findIndex(g=>g.id===dragId);
    groups.splice(dragIdx,1);
    const newOverIdx=groups.findIndex(g=>g.id===overId);
    groups.splice(mode==='before'?newOverIdx:newOverIdx+1,0,dragGrp);
    renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();drawChart();
  }
}

function renderGroupDropdown(){
  const sel=document.getElementById('actGroup');if(!sel)return;
  sel.innerHTML='';
  const categories=groups.filter(g=>g.parentId===null);
  if(!categories.length){sel.innerHTML='<option value="" disabled>No groups created yet</option>';return;}
  categories.forEach(cat=>{
    const optCat=document.createElement('option');optCat.value=cat.id;optCat.textContent='📂 '+cat.name;optCat.style.fontWeight='bold';sel.appendChild(optCat);
    groups.filter(g=>g.parentId===cat.id).forEach(sg=>{const opt=document.createElement('option');opt.value=sg.id;opt.textContent='   └ '+sg.name;sel.appendChild(opt);});
  });
  const spSel=document.getElementById('subgroupParent');
  if(spSel){spSel.innerHTML='<option value="">— Select category —</option>';categories.forEach(cat=>{spSel.innerHTML+=`<option value="${cat.id}">${cat.name}</option>`;});}
}

// ═══════════════════════════════════════════════════════════════════
//  GROUP PICKER IN ACTIVITY EDITOR (Feature 7 — collapsible accordion)
// ═══════════════════════════════════════════════════════════════════
let pickerCollapsedCats=new Set();  // category ids that are collapsed in picker

function renderGroupPickerInEditor(selectedGroupIds){
  const container=document.getElementById('groupPickerContainer');
  if(!container)return;
  if(!selectedGroupIds){
    // Read current selection from hidden select
    const sel=document.getElementById('actGroup');
    selectedGroupIds=sel?new Set(Array.from(sel.selectedOptions).map(o=>o.value)):new Set();
  }
  container.innerHTML='';
  const categories=groups.filter(g=>g.parentId===null);
  if(!categories.length){
    container.innerHTML='<p style="font-size:12px;color:var(--text-muted);padding:6px 8px">No groups yet — create them in the Grouping tab.</p>';
    return;
  }
  categories.forEach(cat=>{
    const isCollapsed=pickerCollapsedCats.has(cat.id);
    const subgroups=groups.filter(g=>g.parentId===cat.id);
    const catSelected=selectedGroupIds.has(cat.id);
    const anyChildSelected=subgroups.some(sg=>selectedGroupIds.has(sg.id));

    // Category header
    const catDiv=document.createElement('div');
    catDiv.className='group-picker-category'+(catSelected||anyChildSelected?' cat-selected':'');
    catDiv.innerHTML=`
      <span class="group-picker-collapse" onclick="togglePickerCat('${cat.id}')">${isCollapsed?'▶':'▼'}</span>
      <span class="group-color-dot" style="background:${cat.color||'#3498db'};width:9px;height:9px;border-radius:50%;flex-shrink:0"></span>
      <span style="flex:1" onclick="togglePickerGroupSelect('${cat.id}')">${cat.name}</span>
      ${catSelected?'<span style="color:#27ae60;font-size:10px">✓</span>':''}
    `;
    container.appendChild(catDiv);

    if(!isCollapsed){
      subgroups.forEach(sg=>{
        const sgSelected=selectedGroupIds.has(sg.id);
        const sgDiv=document.createElement('div');
        sgDiv.className='group-picker-subgroup'+(sgSelected?' selected':'');
        sgDiv.innerHTML=`
          <span class="group-color-dot" style="background:${sg.color||'#e74c3c'};width:8px;height:8px;border-radius:50%;flex-shrink:0"></span>
          <span style="flex:1">${sg.name}</span>
          ${sgSelected?'<span style="font-size:10px">✓</span>':''}
        `;
        sgDiv.onclick=()=>togglePickerGroupSelect(sg.id);
        container.appendChild(sgDiv);
      });
    }
  });
}

function togglePickerCat(catId){
  if(pickerCollapsedCats.has(catId))pickerCollapsedCats.delete(catId);
  else pickerCollapsedCats.add(catId);
  // Re-render with current selection
  const sel=document.getElementById('actGroup');
  const cur=sel?new Set(Array.from(sel.selectedOptions).map(o=>o.value)):new Set();
  renderGroupPickerInEditor(cur);
}

function togglePickerGroupSelect(groupId){
  const sel=document.getElementById('actGroup');if(!sel)return;
  const cur=new Set(Array.from(sel.selectedOptions).map(o=>o.value));
  if(cur.has(groupId))cur.delete(groupId);
  else cur.add(groupId);
  // Sync to hidden select
  Array.from(sel.options).forEach(o=>o.selected=cur.has(o.value));
  renderGroupPickerInEditor(cur);
}

// ═══════════════════════════════════════════════════════════════════
//  CHAINAGE IMAGE
// ═══════════════════════════════════════════════════════════════════
function loadChainageImage(evt){
  const file=evt.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=e=>{chainageImageSrc=e.target.result;const img=new Image();img.onload=()=>{chainageImage=img;syncImgUI();requestAnimationFrame(()=>drawChart());};img.src=chainageImageSrc;};
  reader.readAsDataURL(file);
  document.getElementById('chainageImageName').textContent=file.name;
}
function removeChainageImage(){chainageImage=null;chainageImageSrc=null;document.getElementById('chainageImageInput').value='';document.getElementById('chainageImageName').textContent='No image loaded';syncImgUI();requestAnimationFrame(()=>drawChart());}
function onImgHeightSlider(val){chainageImageH=parseInt(val,10);document.getElementById('imgHeightVal').textContent=chainageImageH;requestAnimationFrame(()=>drawChart());}
function syncImgUI(){const row=document.getElementById('imgHeightRow');row.style.display=chainageImage?'block':'none';if(chainageImage){document.getElementById('imgHeightSlider').value=chainageImageH;document.getElementById('imgHeightVal').textContent=chainageImageH;}}
function getResizeHandleY(){return BASE_TOP_MARGIN+YEAR_BAND_H+chainageImageH;}
function isOverResizeHandle(cy){if(!chainageImage)return false;const hy=getResizeHandleY();return cy>=hy-4&&cy<=hy+RESIZE_HANDLE_H+4;}

// ═══════════════════════════════════════════════════════════════════
//  SERIALISE / DESERIALISE
// ═══════════════════════════════════════════════════════════════════
function genId(){return'n_'+Date.now()+'_'+Math.random().toString(36).substr(2,6);}

function serialiseDiagram(){
  return{
    version:2,
    activities:activities.map(a=>({...a,start:a.start?a.start.toISOString():null,end:a.end?a.end.toISOString():null,coords:undefined})),
    groups,
    settings:{timelineStart:timelineStart?timelineStart.toISOString():null,timelineEnd:timelineEnd?timelineEnd.toISOString():null,minChainage,maxChainage,chainageSpacing,uiStart:document.getElementById('timelineStart').value,uiEnd:document.getElementById('timelineEnd').value},
    image:{src:chainageImageSrc,height:chainageImageH},
    titleBlock,revisions,baselines,dateMarkers,chainageZones,activityLinks,risks,
    cpPaths: cpPaths.map(p=>({...p, highlightSet:[...p.highlightSet]})),
    projectCurrency,
  };
}

function deserialiseDiagram(data){
  const rawActs=data.activities||[];
  activities=rawActs.map(a=>{
    const act={...a,start:a.start?new Date(a.start):null,end:a.end?new Date(a.end):null};
    if(!act.groups){act.groups=act.group!=null?[act.group]:[];}
    delete act.group;
    return act;
  });
  groups        =data.groups       ||[];
  revisions     =data.revisions    ||[];
  baselines     =data.baselines    ||[];
  dateMarkers   =data.dateMarkers  ||[];
  chainageZones =data.chainageZones||[];
  activityLinks =data.activityLinks||[];
  risks         =data.risks        ||[];
  // Ensure group fields exist (migrate older saves)
  groups.forEach(g=>{
    if(g.visible===undefined)g.visible=true;
    if(g.collapsed===undefined)g.collapsed=false;
    if(!g.color)g.color=g.parentId?'#e74c3c':'#3498db';
    if(g.useColor===undefined)g.useColor=false;
  });
  const s=data.settings||{};
  timelineStart=s.timelineStart?new Date(s.timelineStart):null;
  timelineEnd  =s.timelineEnd  ?new Date(s.timelineEnd)  :null;
  minChainage  =s.minChainage  ??0;maxChainage=s.maxChainage??1000;chainageSpacing=s.chainageSpacing??50;
  document.getElementById('timelineStart').value=s.uiStart||'';
  document.getElementById('timelineEnd').value  =s.uiEnd  ||'';
  document.getElementById('minChainage').value  =minChainage;
  document.getElementById('maxChainage').value  =maxChainage;
  document.getElementById('chainSpacing').value =chainageSpacing;
  if(data.titleBlock){titleBlock={...titleBlock,...data.titleBlock};syncTitleBlockUI();}
  if(data.projectCurrency){
    projectCurrency=data.projectCurrency;
    const inp=document.getElementById('settingsCurrencySearch');
    if(inp){
      const match=WORLD_CURRENCIES.find(([sym])=>sym===projectCurrency);
      if(match) inp.value=`${match[0]} — ${match[1]} (${match[2]})`;
    }
    syncCurrencySymbol();
  }
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();renderActivityTable();
  renderRevisionList();renderBaselineList();renderMarkerList();renderZoneList();renderLinkList();
  populateCpTargetSelect();
  cpPaths=(data.cpPaths||[]).map(p=>({...p,highlightSet:new Set(p.highlightSet||[])}));
  cpHighlightSet=new Set();cpTargetIdx=-1;
  _rebuildCpHighlightSet();
  renderAllCriticalPaths();
  resetEditorFields();selectedIds.clear();updateSelectionUI();highlightedGroupId=null;
  const imgData=data.image||{};chainageImageH=imgData.height||120;chainageImageSrc=imgData.src||null;chainageImage=null;
  document.getElementById('chainageImageInput').value='';document.getElementById('chainageImageName').textContent='No image loaded';
  if(chainageImageSrc){
    const img=new Image();
    img.onload=()=>{chainageImage=img;document.getElementById('chainageImageName').textContent='Saved image loaded';syncImgUI();requestAnimationFrame(()=>drawChart());};
    img.src=chainageImageSrc;syncImgUI();requestAnimationFrame(()=>drawChart());
  } else {syncImgUI();requestAnimationFrame(()=>drawChart());}
}

// ═══════════════════════════════════════════════════════════════════
//  NEW DIAGRAM
// ═══════════════════════════════════════════════════════════════════
function newDiagram(){
  if(!confirm('New diagram? Unsaved changes will be lost.'))return;
  activities=[];groups=[];selectedActivity=null;selectedIds.clear();highlightedGroupId=null;
  timelineStart=null;timelineEnd=null;minChainage=0;maxChainage=1000;chainageSpacing=50;
  currentFileName='untitled';
  chainageImage=null;chainageImageSrc=null;chainageImageH=120;
  revisions=[];baselines=[];dateMarkers=[];chainageZones=[];activityLinks=[];risks=[];
  cpPaths=[];cpHighlightSet=new Set();cpTargetIdx=-1;
  titleBlock={projectName:'',drawingTitle:'',drawingNumber:'',revision:'',drawnBy:'',checkedBy:'',date:'',dataDate:'',scale:'NTS',logoSrc:null};
  tbLogoImage=null;undoStack=[];redoStack=[];viewFilter='all';
  document.getElementById('timelineStart').value='';document.getElementById('timelineEnd').value='';
  document.getElementById('minChainage').value=0;document.getElementById('maxChainage').value=1000;document.getElementById('chainSpacing').value=50;
  document.getElementById('chainageImageInput').value='';document.getElementById('chainageImageName').textContent='No image loaded';
  document.getElementById('settingsDataDate').value='';
  syncTitleBlockUI();syncImgUI();updateSelectionUI();setViewFilter('all');
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();renderActivityTable();
  renderRevisionList();renderBaselineList();renderMarkerList();renderZoneList();renderLinkList();populateCpTargetSelect();
  resetEditorFields();updateCurrentFileLabel();
  ctx.clearRect(0,0,canvas.width,canvas.height);
}

// ═══════════════════════════════════════════════════════════════════
//  CANVAS
// ═══════════════════════════════════════════════════════════════════
const canvas=document.getElementById('chartCanvas');
const ctx   =canvas.getContext('2d');
const tooltip=document.createElement('div');tooltip.className='chart-tooltip';document.body.appendChild(tooltip);

function openTab(name){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById(name).classList.add('active');
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  event.target.classList.add('active');

  // Hide the canvas area + toolbar completely when on the dashboard — it has its own content
  const isDash = name === 'dashboard';
  const canvasWrap = document.querySelector('.canvas-wrap');
  const canvasToolbar = document.querySelector('.canvas-toolbar');
  if(canvasWrap)    canvasWrap.style.display    = isDash ? 'none' : '';
  if(canvasToolbar) canvasToolbar.style.display = isDash ? 'none' : '';

  if(name==='list')renderActivityTable();
  if(name==='revisions')renderRevisionList();
  if(name==='baselines')renderBaselineList();
  if(name==='markers')renderMarkerList();
  if(name==='zones')renderZoneList();
  if(name==='links')renderLinkList();
  if(name==='relationships'){renderRelationshipsPanel();}
  if(name==='groups'){renderGroupTree();renderGroupDropdown();}
  if(name==='criticalpath')populateCpTargetSelect();
  if(name==='resources')renderResourceTab();
  if(name==='risks')openTab_risks();
  if(name==='dashboard')renderDashboard();
}

function applySettings(){
  const tsRaw = document.getElementById('timelineStart').value;
  const teRaw = document.getElementById('timelineEnd').value;
  const ts = parseDateLocal(tsRaw);
  const te = parseDateLocal(teRaw);
  if(ts&&te&&te<=ts){
    alert('⚠️ Timeline End must be after Timeline Start. Please correct the dates.');
    return;
  }
  const minCh = parseFloat(document.getElementById('minChainage').value);
  const maxCh = parseFloat(document.getElementById('maxChainage').value);
  if(!isNaN(minCh)&&!isNaN(maxCh)&&minCh>=maxCh){
    alert(`⚠️ Minimum Chainage (${minCh}) must be less than Maximum Chainage (${maxCh}). Please correct these values.`);
    return;
  }
  timelineStart  = ts;
  timelineEnd    = te;
  minChainage    = minCh;
  maxChainage    = maxCh;
  chainageSpacing= parseFloat(document.getElementById('chainSpacing').value);
  drawChart();
}

function syncDataDateFromSettings(){
  const val = document.getElementById('settingsDataDate').value;
  if(val && timelineStart){
    const dd = parseDateLocal(val);
    if(dd && dd < timelineStart){
      alert('⚠️ Data Date cannot be before the Timeline Start. Please correct it.');
      // Reset to current value
      document.getElementById('settingsDataDate').value = titleBlock.dataDate||'';
      return;
    }
  }
  titleBlock.dataDate=val;
  document.getElementById('tbDataDate').value=val;
  drawChart();
}

function parseDateLocal(str){
  if(!str)return null;
  const s=str.trim().replace(/^["']|["']$/g,'');
  // YYYY-MM-DD
  const iso=s.match(/^(\d{4})-(\d{2})-(\d{2})/);if(iso)return new Date(+iso[1],+iso[2]-1,+iso[3]);
  // DD/MM/YYYY
  const dmy4=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(dmy4)return new Date(+dmy4[3],+dmy4[2]-1,+dmy4[1]);
  // DD/MM/YY — 2-digit year, assume 2000s
  const dmy2=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);if(dmy2)return new Date(2000+ +dmy2[3],+dmy2[2]-1,+dmy2[1]);
  // DD-Mon-YYYY  e.g. 25-Mar-2026
  const dmy3=s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if(dmy3){const months={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};const m=months[dmy3[2].toLowerCase()];if(m!==undefined)return new Date(+dmy3[3],m,+dmy3[1]);}
  // DD-Mon-YY  e.g. 25-Mar-26
  const dmy3s=s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
  if(dmy3s){const months={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};const m=months[dmy3s[2].toLowerCase()];if(m!==undefined)return new Date(2000+ +dmy3s[3],m,+dmy3s[1]);}
  return null;
}

function dateToY(date,gridHeight){
  const topM=getTopMargin(),totalMs=timelineEnd-timelineStart,offMs=date-timelineStart;
  return topM+(offMs/totalMs)*gridHeight;
}
function chainageToX(ch,gridWidth){return LEFT_MARGIN+(ch-minChainage)/(maxChainage-minChainage)*gridWidth;}

function buildDateTicks(){
  const ticks=[];let y=timelineStart.getFullYear(),m=timelineStart.getMonth();
  const eY=timelineEnd.getFullYear(),eM=timelineEnd.getMonth();
  while(y<eY||(y===eY&&m<=eM)){ticks.push(new Date(y,m,1));m++;if(m>11){m=0;y++;}}
  return ticks;
}
function isMilestone(s){return s==='circle'||s==='diamond'||s==='triangle'||s==='star'||s==='flag';}

/** Draw a shape fill. Call drawOutlineOn separately after to add an outline. */
function drawShapeOn(c,shape,x1,y1,x2,y2){
  const cx=(x1+x2)/2,cy=(y1+y2)/2,R=MILESTONE_RADIUS,S=MILESTONE_SIZE;
  switch(shape){
    case'rect':c.fillRect(x1,y1,x2-x1,y2-y1);break;
    case'line':c.beginPath();c.moveTo(x1,y1);c.lineTo(x2,y2);c.lineWidth=3;c.stroke();c.lineWidth=1;break;
    case'circle':c.beginPath();c.arc(cx,cy,R,0,Math.PI*2);c.fill();break;
    case'diamond':c.beginPath();c.moveTo(cx,cy-S);c.lineTo(cx+S,cy);c.lineTo(cx,cy+S);c.lineTo(cx-S,cy);c.closePath();c.fill();break;
    case'triangle':c.beginPath();c.moveTo(cx,cy-S);c.lineTo(cx+S,cy+S*0.75);c.lineTo(cx-S,cy+S*0.75);c.closePath();c.fill();break;
    case'star':   _drawStar(c,cx,cy,S,true);break;
    case'flag':   _drawFlag(c,cx,cy,S,true);break;
  }
}

/** Draw a 5-pointed star centred at cx,cy with outer radius r */
function _drawStar(c,cx,cy,r,fill){
  const inner=r*0.42,pts=5;
  c.beginPath();
  for(let i=0;i<pts*2;i++){
    const rad=(i*Math.PI/pts)-Math.PI/2;
    const rr=i%2===0?r:inner;
    i===0?c.moveTo(cx+Math.cos(rad)*rr,cy+Math.sin(rad)*rr):c.lineTo(cx+Math.cos(rad)*rr,cy+Math.sin(rad)*rr);
  }
  c.closePath();
  if(fill)c.fill();else c.stroke();
}

/** Draw a flag: vertical pole + filled rectangular flag */
function _drawFlag(c,cx,cy,S,fill){
  const poleX=cx-S*0.3;
  const poleTop=cy-S;
  const poleBot=cy+S;
  const flagW=S*1.1, flagH=S*0.65;
  // Pole
  c.beginPath();c.moveTo(poleX,poleTop);c.lineTo(poleX,poleBot);
  c.lineWidth=1.5;c.stroke();c.lineWidth=1;
  // Flag body
  c.beginPath();
  c.moveTo(poleX,poleTop);
  c.lineTo(poleX+flagW,poleTop+flagH*0.35);
  c.lineTo(poleX,poleTop+flagH);
  c.closePath();
  if(fill)c.fill();else c.stroke();
}

/** Draw an outline on top of a shape if the activity has outline settings */
function drawOutlineOn(c, a, x1, y1, x2, y2){
  const style  = a.outlineStyle  || 'none';
  if(style === 'none' || style === '') return;
  const color  = a.outlineColor  || '#000000';
  const width  = parseFloat(a.outlineWidth) || 1.5;
  const cx=(x1+x2)/2, cy=(y1+y2)/2, R=MILESTONE_RADIUS, S=MILESTONE_SIZE;

  c.save();
  c.strokeStyle = color;
  c.lineWidth   = width;
  c.globalAlpha = 1;

  if     (style==='dashed')   c.setLineDash([width*4, width*3]);
  else if(style==='dotted')   c.setLineDash([width, width*2]);
  else if(style==='dash-dot') c.setLineDash([width*5, width*2, width, width*2]);
  else                        c.setLineDash([]);

  switch(a.shape||'rect'){
    case'rect':     c.strokeRect(x1,y1,x2-x1,y2-y1);break;
    case'line':     c.beginPath();c.moveTo(x1,y1);c.lineTo(x2,y2);c.stroke();break;
    case'circle':   c.beginPath();c.arc(cx,cy,R,0,Math.PI*2);c.stroke();break;
    case'diamond':  c.beginPath();c.moveTo(cx,cy-S);c.lineTo(cx+S,cy);c.lineTo(cx,cy+S);c.lineTo(cx-S,cy);c.closePath();c.stroke();break;
    case'triangle': c.beginPath();c.moveTo(cx,cy-S);c.lineTo(cx+S,cy+S*0.75);c.lineTo(cx-S,cy+S*0.75);c.closePath();c.stroke();break;
    case'star':     _drawStar(c,cx,cy,S,false);break;
    case'flag':     _drawFlag(c,cx,cy,S,false);break;
  }
  c.setLineDash([]);
  c.restore();
}
function drawProgressOn(c,a,x1,y1,x2,y2){
  const prog=(a.progress||0)/100;if(prog<=0)return;
  if(a.shape==='rect'){
    const h=y2-y1,w=x2-x1;
    c.save();c.globalAlpha=0.55;c.fillStyle='#000';c.fillRect(x1,y1+h*(1-prog),w,h*prog);c.restore();
    if(h*prog>12&&Math.abs(w)>20){c.save();c.globalAlpha=1;c.fillStyle='#fff';c.font='bold 10px Arial';c.textAlign='center';c.textBaseline='middle';c.fillText((a.progress||0)+'%',(x1+x2)/2,y1+h*(1-prog/2));c.restore();}
  } else {
    const cx=(x1+x2)/2,cy=(y1+y2)/2,r=MILESTONE_RADIUS+4;
    c.save();c.globalAlpha=0.9;c.strokeStyle='#fff';c.lineWidth=3;c.beginPath();c.arc(cx,cy,r,-Math.PI/2,-Math.PI/2+(a.progress/100)*Math.PI*2);c.stroke();c.restore();
  }
}
function drawArrowOn(c, sx, sy, ex, ey, color){
  // Elbow-routed line: horizontal jog then vertical, like P6
  // On a time-chainage diagram: Y = time (vertical), X = chainage (horizontal)
  // So a FS link goes: from bottom of predecessor → down a bit → across → to top of successor
  c.save();
  c.strokeStyle = color;
  c.fillStyle   = color;
  c.lineWidth   = 1.5;

  const JOG = 12; // pixels of vertical jog before going horizontal

  // Midpoint horizontally
  const midX = (sx + ex) / 2;
  const midY = sy + JOG;

  c.beginPath();
  c.moveTo(sx, sy);
  // Jog down from start
  c.lineTo(sx, midY);
  // Cross horizontally to midpoint x
  c.lineTo(ex, midY);
  // Drop to target y
  c.lineTo(ex, ey);
  c.stroke();

  // Arrowhead pointing down toward the successor (along Y axis)
  const aL = 7, aW = 4;
  const dir = ey > midY ? 1 : -1;
  c.beginPath();
  c.moveTo(ex, ey);
  c.lineTo(ex - aW, ey - aL * dir);
  c.lineTo(ex + aW, ey - aL * dir);
  c.closePath();
  c.fill();

  c.restore();
}

function drawLinkLabel(c, sx, sy, ex, ey, label){
  if(!label) return;
  const lx = (sx + ex) / 2;
  const ly = sy + 8;
  c.save();
  c.font = '9px Arial';
  c.fillStyle = 'rgba(80,80,80,0.85)';
  c.textAlign = 'center';
  c.textBaseline = 'middle';
  c.fillText(label, lx, ly);
  c.restore();
}
function getEdgeHit(a,mx,my){
  if(!a.coords)return null;
  const{x1,y1,x2,y2}=a.coords;
  const minX=Math.min(x1,x2),maxX=Math.max(x1,x2),minY=Math.min(y1,y2),maxY=Math.max(y1,y2);
  if(mx<minX-EDGE_HIT||mx>maxX+EDGE_HIT||my<minY-EDGE_HIT||my>maxY+EDGE_HIT)return null;
  if(Math.abs(my-minY)<EDGE_HIT)return'top';if(Math.abs(my-maxY)<EDGE_HIT)return'bottom';
  if(Math.abs(mx-minX)<EDGE_HIT)return'left';if(Math.abs(mx-maxX)<EDGE_HIT)return'right';
  return null;
}
function snapChainage(ch){if(!document.getElementById('snapToggle')?.checked)return ch;return Math.round(ch/chainageSpacing)*chainageSpacing;}
function snapDate(date){if(!document.getElementById('snapToggle')?.checked)return date;const d=new Date(date);const mid=new Date(d.getFullYear(),d.getMonth(),15);return d<mid?new Date(d.getFullYear(),d.getMonth(),1):new Date(d.getFullYear(),d.getMonth()+1,1);}

// ═══════════════════════════════════════════════════════════════════
//  ACTIVITIES
// ═══════════════════════════════════════════════════════════════════
function saveActivity(){
  const start=parseDateLocal(document.getElementById('actStart').value);
  const end  =parseDateLocal(document.getElementById('actEnd').value);
  if(start&&end&&end<start){
    alert('⚠️ End date cannot be before the start date. Please correct the dates before saving.');
    return;
  }
  const startCh=parseFloat(document.getElementById('actStartCh').value);
  const endCh  =parseFloat(document.getElementById('actEndCh').value);
  if(!isNaN(startCh)&&!isNaN(endCh)&&startCh>endCh){
    if(!confirm(`⚠️ Start Chainage (${startCh}) is greater than End Chainage (${endCh}).\n\nThis will draw the activity in reverse on the chart. Save anyway?`)) return;
  }
  pushUndo();
  const sel=document.getElementById('actGroup');
  const selectedGroups=sel?Array.from(sel.selectedOptions).map(o=>o.value).filter(v=>v!==''&&v!=='none'):[];
  const obj={
    p6Id:document.getElementById('actP6').value.trim(),
    name:document.getElementById('actName').value.trim(),
    start,end,
    startCh:parseFloat(document.getElementById('actStartCh').value),
    endCh:parseFloat(document.getElementById('actEndCh').value),
    color:document.getElementById('actColor').value,
    alpha:parseFloat(document.getElementById('actAlpha').value),
    shape:document.getElementById('actShape').value,
    label:document.getElementById('actLabel').value,
    progress:parseFloat(document.getElementById('actProgress').value)||0,
    notes:document.getElementById('actNotes').value,
    groups:selectedGroups,
    outlineColor: document.getElementById('actOutlineColor').value||'#000000',
    outlineStyle: document.getElementById('actOutlineStyle').value||'none',
    outlineWidth: parseFloat(document.getElementById('actOutlineWidth').value)||1.5,
    resourceCount: parseFloat(document.getElementById('actResourceCount').value)||0,
    unitCost:      parseFloat(document.getElementById('actUnitCost').value)||0,
    materialCost:  parseFloat(document.getElementById('actMaterialCost').value)||0,
    equipmentCost: parseFloat(document.getElementById('actEquipmentCost').value)||0,
  };
  if(selectedActivity!==null){activities[selectedActivity]=obj;selectedActivity=null;}
  else activities.push(obj);
  resetEditorFields();drawChart();renderActivityTable();populateCpTargetSelect();
}
function deleteActivity(){
  if(selectedActivity===null)return;pushUndo();activities.splice(selectedActivity,1);selectedActivity=null;
  resetEditorFields();drawChart();renderActivityTable();populateCpTargetSelect();
}
function resetEditorFields(){
  selectedActivity=null;
  ['actP6','actName','actStart','actEnd','actStartCh','actEndCh','actNotes'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('actColor').value='#0000ff';
  document.getElementById('actAlpha').value=1;
  document.getElementById('actAlphaVal').value=100;
  document.getElementById('actShape').value='rect';document.getElementById('actLabel').value='no';
  document.getElementById('actProgress').value=0;
  document.getElementById('progressVal').value=0;
  document.getElementById('actOutlineColor').value='#000000';
  document.getElementById('actOutlineStyle').value='none';
  document.getElementById('actOutlineWidth').value=1.5;
  document.getElementById('outlineWidthVal').value=1.5;
  document.getElementById('actResourceCount').value=0;
  document.getElementById('actUnitCost').value=0;
  document.getElementById('actMaterialCost').value=0;
  document.getElementById('actEquipmentCost').value=0;
  const sel=document.getElementById('actGroup');if(sel)Array.from(sel.options).forEach(o=>o.selected=false);
  renderGroupPickerInEditor(new Set());
}
function formatDateInput(d){
  if(!d||isNaN(d))return'';
  return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function editActivity(i){
  const a=activities[i];selectedActivity=i;
  document.getElementById('actP6').value=a.p6Id||'';document.getElementById('actName').value=a.name;
  document.getElementById('actStart').value=formatDateInput(a.start);document.getElementById('actEnd').value=formatDateInput(a.end);
  document.getElementById('actStartCh').value=a.startCh;document.getElementById('actEndCh').value=a.endCh;
  document.getElementById('actColor').value=a.color;
  document.getElementById('actAlpha').value=a.alpha;
  document.getElementById('actAlphaVal').value=Math.round((a.alpha||1)*100);
  document.getElementById('actShape').value=a.shape;document.getElementById('actLabel').value=a.label;
  document.getElementById('actProgress').value=a.progress||0;
  document.getElementById('progressVal').value=a.progress||0;
  document.getElementById('actNotes').value=a.notes||'';
  document.getElementById('actOutlineColor').value=a.outlineColor||'#000000';
  document.getElementById('actOutlineStyle').value=a.outlineStyle||'none';
  document.getElementById('actOutlineWidth').value=a.outlineWidth||1.5;
  document.getElementById('outlineWidthVal').value=a.outlineWidth||1.5;
  document.getElementById('actResourceCount').value=a.resourceCount||0;
  document.getElementById('actUnitCost').value=a.unitCost||0;
  document.getElementById('actMaterialCost').value=a.materialCost||0;
  document.getElementById('actEquipmentCost').value=a.equipmentCost||0;
  const sel=document.getElementById('actGroup');
  if(sel){Array.from(sel.options).forEach(o=>{o.selected=!!(a.groups&&a.groups.includes(o.value));});}
  renderGroupPickerInEditor(new Set(a.groups||[]));
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));document.getElementById('editor').classList.add('active');
  document.querySelectorAll('.tab').forEach(t=>{if(t.textContent.trim()==='Activity Editor')t.classList.add('active');});
}

function renderActivityTable(){
  const container=document.getElementById('activityTableContainer');
  if(!activities.length){container.innerHTML='<p>No activities added yet.</p>';return;}
  let html=`<table class="activityTable"><thead><tr><th></th><th>P6 ID</th><th>Name</th><th>Start</th><th>End</th><th>Ch Start</th><th>Ch End</th><th>Progress</th><th>Colour</th><th>Shape</th><th>Groups</th><th>Actions</th></tr></thead><tbody>`;
  activities.forEach((a,i)=>{
    const gNames=getActivityGroupNames(a);
    const startStr=a.start instanceof Date&&!isNaN(a.start)?formatLongDate(a.start):'';
    const endStr  =a.end   instanceof Date&&!isNaN(a.end)  ?formatLongDate(a.end)  :'';
    const prog=a.progress||0;const isSel=selectedIds.has(i);
    const drawColor=getActivityDrawColor(a);
    html+=`<tr style="${isSel?'background:var(--hover-row)':''}">
      <td><input type="checkbox" ${isSel?'checked':''} onchange="toggleSelect(${i},this.checked)" style="width:auto;margin:0"></td>
      <td>${a.p6Id||''}</td><td>${a.name}</td><td>${startStr}</td><td>${endStr}</td>
      <td>${a.startCh}</td><td>${a.endCh}</td>
      <td><div class="progress-cell"><div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${prog}%;background:${drawColor}"></div></div><span>${prog}%</span></div></td>
      <td><span class="colorSwatch" style="background:${drawColor}"></span>${a.color}</td>
      <td>${a.shape}</td><td style="font-size:12px">${gNames||'—'}</td>
      <td><button onclick="editActivity(${i})">Edit</button><button onclick="deleteActivityByIndex(${i})" class="deleteBtn">Delete</button></td></tr>`;
  });
  html+='</tbody></table>';container.innerHTML=html;
}
function toggleSelect(i,checked){if(checked)selectedIds.add(i);else selectedIds.delete(i);updateSelectionUI();drawChart();}
function deleteActivityByIndex(i){
  pushUndo();activities.splice(i,1);
  if(selectedActivity===i){selectedActivity=null;resetEditorFields();}
  else if(selectedActivity!==null&&selectedActivity>i)selectedActivity--;
  selectedIds.clear();drawChart();renderActivityTable();populateCpTargetSelect();
}

// ═══════════════════════════════════════════════════════════════════
//  MAIN DRAW CHART — Feature 8 integrated
// ═══════════════════════════════════════════════════════════════════
function drawChart(){
  if(!timelineStart||!timelineEnd)return;
  const ct=getChartTheme(),topM=getTopMargin(),hasTB=hasTitleBlock();
  const showDataDate  =document.getElementById('showDataDateToggle')?.checked;
  const showPageBreaks=document.getElementById('showPageBreaksToggle')?.checked;
  const showCritical  =document.getElementById('showCriticalToggle')?.checked;
  const showLinks     =document.getElementById('showLinksToggle')?.checked;

  const parentW=canvas.parentElement.offsetWidth;
  canvas.width=parentW>0?parentW:800;
  canvas.height=getRequiredCanvasHeight(hasTB);
  ctx.fillStyle=ct.canvasBg;ctx.fillRect(0,0,canvas.width,canvas.height);

  const gridW=canvas.width-LEFT_MARGIN-40;
  const gridH=canvas.height-topM-BOTTOM_MARGIN-(hasTB?TITLE_BLOCK_H:0);

  // ── FEATURE 8: Year bands ──
  drawYearBandsOn(ctx,ct,canvas.width,gridW,canvas.height);

  // Chainage axis labels
  ctx.fillStyle=ct.axisText;ctx.textAlign='center';ctx.textBaseline='top';ctx.font='11px Arial';
  for(let ch=minChainage;ch<=maxChainage;ch+=chainageSpacing)ctx.fillText(ch,chainageToX(ch,gridW),6);

  // Reference image
  if(chainageImage){
    const imgY=BASE_TOP_MARGIN+YEAR_BAND_H;
    ctx.save();ctx.beginPath();ctx.rect(LEFT_MARGIN,imgY,gridW,chainageImageH);ctx.clip();
    ctx.drawImage(chainageImage,LEFT_MARGIN,imgY,gridW,chainageImageH);
    ctx.restore();ctx.strokeStyle=ct.imgBorder;ctx.lineWidth=1;ctx.strokeRect(LEFT_MARGIN,imgY,gridW,chainageImageH);
    const hY=imgY+chainageImageH;
    ctx.fillStyle=ct.resizeBar;ctx.fillRect(LEFT_MARGIN,hY,gridW,RESIZE_HANDLE_H);
    ctx.fillStyle=ct.imgLabel;
    const dC=7,dS=14,dX=LEFT_MARGIN+gridW/2-((dC-1)*dS)/2;
    for(let d=0;d<dC;d++){ctx.beginPath();ctx.arc(dX+d*dS,hY+RESIZE_HANDLE_H/2,2,0,Math.PI*2);ctx.fill();}
  }

  // Zones
  chainageZones.forEach(z=>{
    const x1=chainageToX(Math.max(z.startCh,minChainage),gridW),x2=chainageToX(Math.min(z.endCh,maxChainage),gridW);
    ctx.save();ctx.globalAlpha=z.alpha;ctx.fillStyle=z.color;ctx.fillRect(x1,topM,x2-x1,gridH);ctx.restore();
    if(!z.hideLabel){
      // Draw zone label in the header band above the grid (between chainage numbers and grid)
      const labelY = topM - 12;
      const labelX = (x1+x2)/2;
      ctx.save();
      ctx.globalAlpha=0.9;
      ctx.fillStyle=z.color;
      const tw=ctx.measureText(z.name).width;
      ctx.fillRect(labelX-tw/2-4,labelY-7,tw+8,14);
      ctx.fillStyle='#fff';ctx.font='bold 9px Arial';ctx.textAlign='center';ctx.textBaseline='middle';
      ctx.fillText(z.name,labelX,labelY);
      ctx.restore();
    }
  });

  // Month grid lines — every single month shown, canvas is tall enough to fit them all
  const ticks=buildDateTicks();
  ticks.forEach((tick) => {
    const y = dateToY(tick, gridH);
    const isJan = tick.getMonth() === 0;
    ctx.strokeStyle = isJan ? ct.yearBandBorder : ct.gridH;
    ctx.lineWidth   = isJan ? 1.5 : 1;
    ctx.beginPath(); ctx.moveTo(LEFT_MARGIN, y); ctx.lineTo(canvas.width - 40, y); ctx.stroke();
    ctx.lineWidth = 1;
    ctx.fillStyle    = isJan ? ct.yearBandText : ct.axisText;
    ctx.font         = isJan ? 'bold 10px Arial' : '10px Arial';
    ctx.textAlign    = 'right';
    ctx.textBaseline = 'middle';
    ctx.fillText(tick.toLocaleDateString(undefined, {month:'short'}), LEFT_MARGIN - 4, y);
  });

  // Vertical chainage grid
  for(let ch=minChainage;ch<=maxChainage;ch+=chainageSpacing){
    const x=chainageToX(ch,gridW);
    if(chainageImage){
      const imgY=BASE_TOP_MARGIN+YEAR_BAND_H;
      ctx.save();ctx.globalAlpha=0.45;ctx.strokeStyle=ct.axisText;ctx.setLineDash([3,4]);ctx.lineWidth=1;
      ctx.beginPath();ctx.moveTo(x,imgY);ctx.lineTo(x,imgY+chainageImageH);ctx.stroke();ctx.setLineDash([]);ctx.restore();
    }
    ctx.strokeStyle=ct.gridV;ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(x,topM);ctx.lineTo(x,canvas.height-BOTTOM_MARGIN-(hasTB?TITLE_BLOCK_H:0));ctx.stroke();
  }

  // Date markers — drawn AFTER activities (see below, after activity loop)

  // Data date — also drawn after activities
  // Baselines
  baselines.filter(b=>b.visible).forEach((b,bi)=>{
    b.activities.forEach(ba=>{
      if(!ba.start||!ba.end)return;
      const y1=dateToY(new Date(ba.start),gridH),y2=dateToY(new Date(ba.end),gridH),x1=chainageToX(ba.startCh,gridW),x2=chainageToX(ba.endCh,gridW);
      ctx.save();ctx.globalAlpha=0.28;ctx.fillStyle=b.tint;ctx.strokeStyle=b.tint;ctx.setLineDash([4,3]);
      drawShapeOn(ctx,ba.shape||'rect',x1,y1,x2,y2);ctx.globalAlpha=0.5;ctx.lineWidth=1.5;
      if((ba.shape||'rect')==='rect')ctx.strokeRect(x1,y1,x2-x1,y2-y1);ctx.setLineDash([]);ctx.restore();
    });
    ctx.save();ctx.globalAlpha=0.8;ctx.fillStyle=b.tint;ctx.font='bold 10px Arial';ctx.textAlign='left';ctx.textBaseline='top';
    ctx.fillText('▬ '+b.name,LEFT_MARGIN+4,topM+4+bi*14);ctx.restore();
  });

  // Determine which activities to highlight via critical path
  const critSet=showCritical?computeCriticalPath():new Set();
  const hGroupIds=getHighlightedGroupIds();

  // Activities
  activities.forEach((a,i)=>{
    if(!a.start||!a.end||isNaN(a.start)||isNaN(a.end))return;
    if(!activityInViewFilter(a))return;
    if(!activityVisibleByGroup(a))return;

    const y1=dateToY(a.start,gridH),y2=dateToY(a.end,gridH);
    const x1=chainageToX(a.startCh,gridW),x2=chainageToX(a.endCh,gridW);
    const matchSearch=activityMatchesSearch(a);
    const isGroupHighlighted=!hGroupIds||activityIsHighlighted(a);
    const isSel=selectedIds.has(i);
    const isOnCritPath=cpHighlightSet.size>0?cpHighlightSet.has(i):(critSet.has(i));
    const cpColor = getActivityCpColor(i) || '#e74c3c';

    let alpha=a.alpha;
    if(!isGroupHighlighted)alpha=0.12;
    if(searchFilter&&!matchSearch)alpha=0.1;

    // Critical path highlight ring — uses the path's own colour
    if(isOnCritPath){
      ctx.save();ctx.strokeStyle=cpColor;ctx.lineWidth=3;ctx.strokeRect(x1-2,y1-2,x2-x1+4,y2-y1+4);ctx.restore();
    }
    // Selection ring
    if(isSel){
      ctx.save();ctx.strokeStyle='#f39c12';ctx.lineWidth=2.5;ctx.setLineDash([4,3]);
      ctx.strokeRect(Math.min(x1,x2)-3,Math.min(y1,y2)-3,Math.abs(x2-x1)+6,Math.abs(y2-y1)+6);
      ctx.setLineDash([]);ctx.restore();
    }

    const drawColor=getActivityDrawColor(a);
    ctx.globalAlpha=alpha;ctx.fillStyle=drawColor;ctx.strokeStyle=drawColor;
    drawShapeOn(ctx,a.shape,x1,y1,x2,y2);
    drawOutlineOn(ctx,a,x1,y1,x2,y2);
    if((a.progress||0)>0)drawProgressOn(ctx,a,x1,y1,x2,y2);

    if(a.label&&a.label!=='no'){
      const lAlpha=(!isGroupHighlighted)?0.12:(searchFilter&&!matchSearch?0.1:1);
      ctx.globalAlpha=lAlpha;ctx.fillStyle=ct.labelText;ctx.font='bold 12px Arial';ctx.textBaseline='middle';
      const cx=(x1+x2)/2,midY=(y1+y2)/2;let lx;
      if(a.label==='left'){ctx.textAlign='right';lx=isMilestone(a.shape)?cx-MILESTONE_SIZE-4:Math.min(x1,x2)-4;}
      else if(a.label==='right'){ctx.textAlign='left';lx=isMilestone(a.shape)?cx+MILESTONE_SIZE+4:Math.max(x1,x2)+4;}
      else{ctx.textAlign='center';lx=cx;}
      ctx.fillText(a.name,lx,midY);
    }
    ctx.globalAlpha=1;
    a.coords={x1,y1,x2,y2};
  });

  // Dependency links — elbow-routed, correct edge connections
  if(showLinks){
    activityLinks.forEach(l=>{
      const from=activities[l.fromIdx],to=activities[l.toIdx];
      if(!from||!to||!from.coords||!to.coords)return;
      const fc=from.coords,tc=to.coords;
      // Y axis = time: y1=start, y2=end. X axis = chainage: x1=startCh, x2=endCh
      // Centre X of each bar
      const fcX=(fc.x1+fc.x2)/2, tcX=(tc.x1+tc.x2)/2;
      let sx,sy,ex,ey;
      if(l.type==='FS'){
        // From end of predecessor (bottom) to start of successor (top)
        sx=fcX; sy=Math.max(fc.y1,fc.y2);
        ex=tcX; ey=Math.min(tc.y1,tc.y2);
      } else if(l.type==='SS'){
        // From start of predecessor to start of successor
        sx=fcX; sy=Math.min(fc.y1,fc.y2);
        ex=tcX; ey=Math.min(tc.y1,tc.y2);
      } else if(l.type==='FF'){
        // From end of predecessor to end of successor
        sx=fcX; sy=Math.max(fc.y1,fc.y2);
        ex=tcX; ey=Math.max(tc.y1,tc.y2);
      } else { // SF
        // From start of predecessor to end of successor
        sx=fcX; sy=Math.min(fc.y1,fc.y2);
        ex=tcX; ey=Math.max(tc.y1,tc.y2);
      }
      drawArrowOn(ctx,sx,sy,ex,ey,'rgba(100,100,100,0.7)');
      const lagStr=l.lag?`${l.lag>0?'+':''}${l.lag}d`:'';
      drawLinkLabel(ctx,sx,sy,ex,ey,l.type+(lagStr?` ${lagStr}`:''));
    });
  }

  // Page break guides
  if(showPageBreaks){
    const sizeMap={a1:[594,841],a3:[297,420],a4:[210,297]};
    const sel=document.querySelector('input[name="expSize"]:checked');
    const[pw,ph]=sizeMap[sel?.value||'a3']||[297,420];
    const scX=canvas.width/Math.max(pw,ph),scY=canvas.height/Math.min(pw,ph);
    ctx.save();ctx.strokeStyle='rgba(255,0,0,0.25)';ctx.lineWidth=1;ctx.setLineDash([8,4]);
    for(let x=Math.max(pw,ph)*scX;x<canvas.width;x+=Math.max(pw,ph)*scX){ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x,canvas.height);ctx.stroke();}
    for(let y=Math.min(pw,ph)*scY;y<canvas.height;y+=Math.min(pw,ph)*scY){ctx.beginPath();ctx.moveTo(0,y);ctx.lineTo(canvas.width,y);ctx.stroke();}
    ctx.setLineDash([]);ctx.restore();
  }

  // Rubber band selection
  if(isRubberBanding){
    ctx.save();ctx.strokeStyle='#007bff';ctx.lineWidth=1.5;ctx.setLineDash([4,3]);ctx.fillStyle='rgba(0,123,255,0.07)';
    const rx=Math.min(rubberStart.x,rubberEnd.x),ry=Math.min(rubberStart.y,rubberEnd.y),rw=Math.abs(rubberEnd.x-rubberStart.x),rh=Math.abs(rubberEnd.y-rubberStart.y);
    ctx.fillRect(rx,ry,rw,rh);ctx.strokeRect(rx,ry,rw,rh);ctx.setLineDash([]);ctx.restore();
  }

  // ── Risk indicators — coloured dots on activities with open risks ──
  _drawRiskIndicators(ctx,gridW,gridH,ct);
  // ── Risk P-lines (P50/P80/P90) ──
  _drawRiskPLines(ctx,canvas.width,gridH,ct);

  // ── Date markers — drawn ON TOP of activities ──
  if(document.getElementById('showDateMarkersToggle')?.checked ?? true)
    _drawDateMarkersOn(ctx, canvas.width, gridH, ct);

  // ── Data date line — on top of everything ──
  if(showDataDate&&titleBlock.dataDate&&(document.getElementById('showDateMarkersToggle')?.checked??true)){
    const dd=parseDateLocal(titleBlock.dataDate);
    if(dd){
      const y=dateToY(dd,gridH);
      ctx.save();ctx.strokeStyle='#e74c3c';ctx.lineWidth=2.5;ctx.setLineDash([5,4]);
      ctx.beginPath();ctx.moveTo(LEFT_MARGIN,y);ctx.lineTo(canvas.width-40,y);ctx.stroke();ctx.setLineDash([]);
      const lbl='▶ Data Date';
      ctx.font='bold 11px Arial';
      const tw=ctx.measureText(lbl).width;
      ctx.fillStyle='#e74c3c';ctx.globalAlpha=0.9;ctx.fillRect(canvas.width-44-tw-8,y-17,tw+10,15);
      ctx.globalAlpha=1;ctx.fillStyle='#fff';ctx.textAlign='right';ctx.textBaseline='bottom';
      ctx.fillText(lbl,canvas.width-48,y-4);
      ctx.restore();
    }
  }

  if(hasTB)drawTitleBlock(ct);
}

/** Draw date marker lines + pill labels on top of activities */
function _drawDateMarkersOn(c, canvasW, gridH, ct){
  dateMarkers.forEach(m=>{
    const d=parseDateLocal(m.date);if(!d)return;
    const y=dateToY(d,gridH);
    c.save();
    c.strokeStyle=m.color;c.lineWidth=2;c.setLineDash([6,4]);
    c.beginPath();c.moveTo(LEFT_MARGIN,y);c.lineTo(canvasW-40,y);c.stroke();c.setLineDash([]);
    if(!m.hideLabel){
      c.font='bold 10px Arial';
      const tw=c.measureText(m.name).width;
      const px=LEFT_MARGIN+6, py=y-18;
      // Pill background
      c.globalAlpha=0.88;c.fillStyle=m.color;
      c.beginPath();c.roundRect?c.roundRect(px-3,py,tw+10,15,4):c.fillRect(px-3,py,tw+10,15);
      c.fill();
      // Text
      c.globalAlpha=1;c.fillStyle='#fff';
      c.textAlign='left';c.textBaseline='top';
      c.fillText(m.name,px+2,py+2);
    }
    c.restore();
  });
}

// ═══════════════════════════════════════════════════════════════════
//  CANVAS INTERACTIONS
// ═══════════════════════════════════════════════════════════════════
function getCanvasPos(e){const r=canvas.getBoundingClientRect();return{mx:e.clientX-r.left,my:e.clientY-r.top};}
function findActivityAtPoint(mx,my){
  for(let i=activities.length-1;i>=0;i--){
    const a=activities[i];if(!a.coords)continue;
    const minX=Math.min(a.coords.x1,a.coords.x2),maxX=Math.max(a.coords.x1,a.coords.x2);
    const minY=Math.min(a.coords.y1,a.coords.y2),maxY=Math.max(a.coords.y1,a.coords.y2);
    if(mx>=minX-5&&mx<=maxX+5&&my>=minY-5&&my<=maxY+5)return i;
  }
  return -1;
}

// Ctrl+click link selection state
let ctrlLinkFirst = -1;  // index of first Ctrl-clicked activity

canvas.addEventListener('contextmenu', e => {
  // Right-click while Ctrl held — if we have a first activity selected, open link dialog
  if((e.ctrlKey||e.metaKey) && ctrlLinkFirst !== -1){
    e.preventDefault();
    const{mx,my}=getCanvasPos(e);
    const idx=findActivityAtPoint(mx,my);
    if(idx !== -1 && idx !== ctrlLinkFirst){
      // Open link dialog pre-filled with from=ctrlLinkFirst, to=idx
      openLinkDialogForPair(ctrlLinkFirst, idx);
    }
    ctrlLinkFirst=-1;
    selectedIds.clear();updateSelectionUI();drawChart();
    return;
  }
});

function openLinkDialogForPair(fromIdx, toIdx){
  if(activities.length<2) return;
  linkEditId=null;
  const from=document.getElementById('linkFrom'),to=document.getElementById('linkTo');
  from.innerHTML='';to.innerHTML='';
  activities.forEach((a,i)=>{const opt=`<option value="${i}">${a.p6Id||i} — ${a.name}</option>`;from.innerHTML+=opt;to.innerHTML+=opt;});
  from.value=fromIdx; to.value=toIdx;
  document.getElementById('linkType').value='FS';
  document.getElementById('linkLag').value=0;
  document.getElementById('linkDialogTitle').textContent='🔗 Add Dependency Link';
  document.getElementById('linkDialog').style.display='flex';
}

canvas.addEventListener('mousedown',e=>{
  if(e.button!==0)return;
  const{mx,my}=getCanvasPos(e);
  if(chainageImage&&isOverResizeHandle(my)){imgResizeDragging=true;imgResizeDragStartY=e.clientY;imgResizeDragStartH=chainageImageH;e.preventDefault();return;}
  if(chainageImage&&my<getTopMargin())return;
  const idx=findActivityAtPoint(mx,my);

  // Ctrl+click: select first or second activity for link creation
  if((e.ctrlKey||e.metaKey)&&idx!==-1){
    e.preventDefault();
    if(ctrlLinkFirst===-1){
      // First click — store it, highlight it
      ctrlLinkFirst=idx;
      selectedIds.clear();selectedIds.add(idx);updateSelectionUI();drawChart();
      // Show hint
      const info=document.getElementById('selectionInfo');
      info.style.display='inline';
      info.textContent='Ctrl+right-click a second activity to add a link';
    } else {
      // Second Ctrl+click — also open dialog (alternative to right-click)
      if(idx!==ctrlLinkFirst){
        openLinkDialogForPair(ctrlLinkFirst,idx);
      }
      ctrlLinkFirst=-1;
      selectedIds.clear();updateSelectionUI();drawChart();
    }
    return;
  }

  // If Ctrl was released, cancel any pending link selection
  if(!e.ctrlKey&&!e.metaKey&&ctrlLinkFirst!==-1){
    ctrlLinkFirst=-1;
  }

  if(idx!==-1&&!isMilestone(activities[idx].shape)){
    const edge=getEdgeHit(activities[idx],mx,my);
    if(edge){isResizing=true;resizeActIdx=idx;resizeEdge=edge;resizeStartMX=mx;resizeStartMY=my;const a=activities[idx];resizeOrigVal={start:new Date(a.start),end:new Date(a.end),startCh:a.startCh,endCh:a.endCh};pushUndo();e.preventDefault();return;}
  }
  if(idx!==-1){
    pushUndo();
    if(!selectedIds.has(idx)){selectedIds.clear();selectedIds.add(idx);updateSelectionUI();}
    isDragging=true;dragActivityId=idx;dragStartMX=mx;dragStartMY=my;
    const a=activities[idx];dragOrigStart=new Date(a.start);dragOrigEnd=new Date(a.end);dragOrigSCh=a.startCh;dragOrigECh=a.endCh;
    canvas.style.cursor='grabbing';e.preventDefault();
  } else {
    isRubberBanding=true;rubberStart={x:mx,y:my};rubberEnd={x:mx,y:my};
    if(!e.ctrlKey&&!e.shiftKey&&!e.metaKey){selectedIds.clear();updateSelectionUI();}
    resetEditorFields();
  }
});

window.addEventListener('mousemove',e=>{
  // Never show canvas tooltips when dashboard is active
  if(document.getElementById('dashboard')?.classList.contains('active')){
    tooltip.style.display='none';
    return;
  }
  if(imgResizeDragging){
    const delta=e.clientY-imgResizeDragStartY;
    chainageImageH=Math.max(40,Math.min(400,imgResizeDragStartH+delta));
    document.getElementById('imgHeightSlider').value=chainageImageH;document.getElementById('imgHeightVal').textContent=chainageImageH;
    requestAnimationFrame(()=>drawChart());return;
  }
  const{mx,my}=getCanvasPos(e);
  if(isRubberBanding){rubberEnd={x:mx,y:my};drawChart();return;}
  if(isResizing){
    const a=activities[resizeActIdx];if(!a)return;
    const gridW=canvas.width-LEFT_MARGIN-40,gridH=canvas.height-getTopMargin()-BOTTOM_MARGIN-(hasTitleBlock()?TITLE_BLOCK_H:0);
    const dMX=mx-resizeStartMX,dMY=my-resizeStartMY;
    const totalMs=timelineEnd-timelineStart,msPxY=totalMs/gridH,chPxX=(maxChainage-minChainage)/gridW;
    if(resizeEdge==='top')a.start=snapDate(new Date(resizeOrigVal.start.getTime()+dMY*msPxY));
    if(resizeEdge==='bottom')a.end=snapDate(new Date(resizeOrigVal.end.getTime()+dMY*msPxY));
    if(resizeEdge==='left')a.startCh=snapChainage(resizeOrigVal.startCh+dMX*chPxX);
    if(resizeEdge==='right')a.endCh=snapChainage(resizeOrigVal.endCh+dMX*chPxX);
    requestAnimationFrame(()=>drawChart());return;
  }
  if(isDragging&&dragActivityId!==-1){
    const a=activities[dragActivityId];if(!a)return;
    const gridW=canvas.width-LEFT_MARGIN-40,gridH=canvas.height-getTopMargin()-BOTTOM_MARGIN-(hasTitleBlock()?TITLE_BLOCK_H:0);
    const dMX=mx-dragStartMX,dMY=my-dragStartMY;
    const totalMs=timelineEnd-timelineStart,msPxY=totalMs/gridH,chPxX=(maxChainage-minChainage)/gridW;
    const dur=dragOrigEnd-dragOrigStart;
    const ns=snapDate(new Date(dragOrigStart.getTime()+dMY*msPxY));
    const ne=new Date(ns.getTime()+dur);
    const dCh=snapChainage(dMX*chPxX);
    a.start=ns;a.end=ne;a.startCh=snapChainage(dragOrigSCh+dCh);a.endCh=snapChainage(dragOrigECh+dCh);
    if(selectedIds.size>1){
      selectedIds.forEach(si=>{
        if(si===dragActivityId)return;
        const sa=activities[si];if(!sa)return;
        const sd=sa.end-sa.start;const sns=snapDate(new Date(sa.start.getTime()+dMY*msPxY));
        sa.start=sns;sa.end=new Date(sns.getTime()+sd);sa.startCh=snapChainage(sa.startCh+dCh);sa.endCh=snapChainage(sa.endCh+dCh);
      });
    }
    requestAnimationFrame(()=>drawChart());return;
  }
  if(chainageImage&&isOverResizeHandle(my)){canvas.style.cursor='ns-resize';tooltip.style.display='none';return;}
  const idx=findActivityAtPoint(mx,my);
  if(idx!==-1){
    const edge=getEdgeHit(activities[idx],mx,my);
    if(edge==='top'||edge==='bottom')canvas.style.cursor='ns-resize';
    else if(edge==='left'||edge==='right')canvas.style.cursor='ew-resize';
    else canvas.style.cursor='grab';
    const a=activities[idx];
    tooltip.style.display='block';tooltip.style.left=e.pageX+10+'px';tooltip.style.top=e.pageY+10+'px';
    tooltip.innerHTML=buildActivityTooltipHTML(a);
  } else if(chartViewMode!=='programme'){
    // Show cost/resource hover tooltips
    _checkChartHover(e);
    canvas.style.cursor='default';
  } else {canvas.style.cursor='default';tooltip.style.display='none';}
});

window.addEventListener('mouseup',e=>{
  if(imgResizeDragging){imgResizeDragging=false;canvas.style.cursor='default';}
  if(isRubberBanding){
    const rx1=Math.min(rubberStart.x,rubberEnd.x),rx2=Math.max(rubberStart.x,rubberEnd.x);
    const ry1=Math.min(rubberStart.y,rubberEnd.y),ry2=Math.max(rubberStart.y,rubberEnd.y);
    if(rx2-rx1>4&&ry2-ry1>4){
      activities.forEach((a,i)=>{
        if(!a.coords)return;
        const ax=Math.min(a.coords.x1,a.coords.x2),bx=Math.max(a.coords.x1,a.coords.x2);
        const ay=Math.min(a.coords.y1,a.coords.y2),by=Math.max(a.coords.y1,a.coords.y2);
        if(ax>=rx1&&bx<=rx2&&ay>=ry1&&by<=ry2)selectedIds.add(i);
      });
      updateSelectionUI();
    }
    isRubberBanding=false;drawChart();
  }
  if(isDragging){isDragging=false;dragActivityId=-1;canvas.style.cursor='default';renderActivityTable();}
  if(isResizing){isResizing=false;resizeActIdx=-1;canvas.style.cursor='default';renderActivityTable();}
});

canvas.addEventListener('click',e=>{
  if(isDragging||isResizing||isRubberBanding)return;
  const{mx,my}=getCanvasPos(e);
  if(chainageImage&&my<getTopMargin())return;
  const idx=findActivityAtPoint(mx,my);
  if(idx!==-1){
    if(!e.ctrlKey&&!e.shiftKey&&!e.metaKey){selectedIds.clear();selectedIds.add(idx);updateSelectionUI();}
    editActivity(idx);
  } else {
    if(!e.ctrlKey&&!e.shiftKey&&!e.metaKey){selectedIds.clear();updateSelectionUI();}
    resetEditorFields();
  }
});
canvas.addEventListener('mouseleave',()=>{tooltip.style.display='none';});

// ═══════════════════════════════════════════════════════════════════
//  KEYBOARD SHORTCUTS
// ═══════════════════════════════════════════════════════════════════
document.addEventListener('keydown',e=>{
  const tag=document.activeElement.tagName.toLowerCase();
  const inInput=tag==='input'||tag==='textarea'||tag==='select';
  if((e.ctrlKey||e.metaKey)&&e.key==='z'&&!e.shiftKey){e.preventDefault();undo();return;}
  if((e.ctrlKey||e.metaKey)&&(e.key==='y'||(e.key==='z'&&e.shiftKey))){e.preventDefault();redo();return;}
  if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();quickSaveToPC();return;}
  if((e.ctrlKey||e.metaKey)&&e.key==='a'&&!inInput){e.preventDefault();activities.forEach((_,i)=>selectedIds.add(i));updateSelectionUI();drawChart();return;}
  if(inInput)return;
  if(e.key==='Delete'||e.key==='Backspace'){if(selectedIds.size>0)bulkDelete();else if(selectedActivity!==null)deleteActivity();return;}
  if(e.key==='Escape'){deselectAll();resetEditorFields();return;}
  if(e.key==='F11'||e.key==='f'){openFullscreen();return;}
});

// ═══════════════════════════════════════════════════════════════════
//  P6 CSV IMPORT
// ═══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════
//  SYNC DATES FROM P6 — uses same XLSX engine as bulk import
//  Matches by P6 Activity ID and updates start/end dates only
// ═══════════════════════════════════════════════════════════════════
function importP6(){
  const file=document.getElementById('p6File').files[0];
  if(!file){alert('Please select a file.');return;}
  const debugDiv=document.getElementById('csvDebug');
  debugDiv.innerHTML='<p style="color:var(--text-muted);font-size:13px">Reading file…</p>';

  const ext=file.name.split('.').pop().toLowerCase();
  const reader=new FileReader();

  reader.onload=e=>{
    try{
      let rows=[];

      if(ext==='xlsx'||ext==='xls'){
        // Excel path — same as bulk import
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
        const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      } else {
        // CSV path — parse manually into same row-object format
        const text=new TextDecoder().decode(e.target.result);
        const lines=text.split(/\r?\n/).filter(l=>l.trim()!=='');
        if(lines.length<2){debugDiv.innerHTML='<p class="csv-error">❌ File has fewer than 2 rows.</p>';return;}
        const headers=parseCSVRow(lines[0]);
        lines.slice(1).forEach(line=>{
          const vals=parseCSVRow(line);
          const obj={};
          headers.forEach((h,i)=>obj[h]=vals[i]||'');
          rows.push(obj);
        });
      }

      if(!rows.length){debugDiv.innerHTML='<p class="csv-error">❌ No data rows found.</p>';return;}

      const norm=s=>String(s).toLowerCase().replace(/[^a-z0-9]/g,'');
      let updated=0,notFound=0,skipped=0;

      rows.forEach(row=>{
        const keys=Object.keys(row);
        const get=(...cands)=>{
          for(const c of cands){
            const k=keys.find(k=>norm(k)===c);
            if(k&&row[k]!=='')return String(row[k]).trim();
          }
          return null;
        };

        const id   =get('p6id','activityid','activity_id','id','taskid','task_id');
        const sRaw =get('start','startdate','start_date','plannedstart','planned_start','earlystart','early_start','actualstart','actual_start');
        const eRaw =get('end','finish','finishdate','finish_date','enddate','end_date','plannedfinish','planned_finish','earlyfinish','early_finish','actualfinish','actual_finish');

        if(!id||!sRaw||!eRaw){skipped++;return;}

        const sD=parseExcelDate(sRaw);
        const eD=parseExcelDate(eRaw);
        if(!sD||!eD||isNaN(sD)||isNaN(eD)){skipped++;return;}

        let matched=false;
        activities.forEach(a=>{
          if(a.p6Id&&a.p6Id.trim()===id){
            a.start=sD;a.end=eD;updated++;matched=true;
          }
        });
        if(!matched)notFound++;
      });

      const total=rows.length;
      debugDiv.innerHTML=`
        <div class="csv-debug-panel">
          <strong>Sync complete</strong><br>
          Rows read: <b>${total}</b><br>
          ✅ Activities updated: <b>${updated}</b><br>
          ${notFound?`⚠️ P6 IDs not found on chart: <b>${notFound}</b> — make sure activities have their P6 ID set in the Activity Editor<br>`:''}
          ${skipped?`⏭ Rows skipped (missing ID or dates): <b>${skipped}</b>`:''}
        </div>
        <p class="${updated>0?'csv-success':'csv-error'}">${updated>0?`✅ ${updated} activit${updated===1?'y':'ies'} updated.`:'❌ No activities were updated — check that P6 IDs match.'}</p>
      `;
      if(updated>0){drawChart();renderActivityTable();}
    }catch(err){
      debugDiv.innerHTML=`<p class="csv-error">❌ Error reading file: ${err.message}</p>`;
    }
  };

  // Read as ArrayBuffer for Excel, as ArrayBuffer for CSV too (we'll decode manually)
  reader.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════════════════════════════
//  EXCEL TEMPLATE DOWNLOAD
// ═══════════════════════════════════════════════════════════════════
function downloadExcelTemplate(){
  if(typeof XLSX==='undefined'){alert('XLSX library not loaded.');return;}
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: Instructions ──
  const instructions = [
    ['Time-Chainage Planner — Excel Import Template'],[''],
    ['HOW TO USE THIS TEMPLATE'],
    ['1. Fill in the "Activities" sheet — one row per activity. Delete the two example rows before importing.'],
    ['2. Do not rename column headers in row 1.'],
    ['3. Save as .xlsx and use "Bulk Create Activities from Excel" in the Sync tab to import.'],[''],
    ['REQUIRED COLUMNS'],
    ['Column','What it means','Example'],
    ['Name','Activity name shown on the chart','Drainage Run A'],
    ['Start','Start date (DD/MM/YYYY)','01/03/2026'],
    ['End','End / finish date (DD/MM/YYYY)','30/06/2026'],
    ['StartCh','Start chainage (distance from datum)','100'],
    ['EndCh','End chainage','450'],[''],
    ['OPTIONAL COLUMNS — leave blank to use defaults'],
    ['Column','What it means','Example / Options'],
    ['P6ID','Primavera P6 Activity ID — needed for the Sync Dates tool','A1050'],
    ['Colour','Fill colour as a hex code','#e74c3c'],
    ['Shape','Bar shape on chart','rect · line · circle · diamond · triangle · star · flag'],
    ['Alpha','Opacity: 1.0 = solid, 0.5 = 50% transparent (or enter 50 for 50%)','0.8'],
    ['Label','Show name label on chart and where','no · left · center · right'],
    ['Notes','Free-text notes shown in the hover tooltip','Hold for design approval'],[''],
    ['RESOURCE & COST COLUMNS'],
    ['Column','What it means','Example'],
    ['Workers','Number of workers per day on this activity','12'],
    ['LabourRate','Cost per worker per day. Total labour = LabourRate × Workers × Duration','350'],
    ['EquipmentDay','Daily plant or equipment hire rate. Total equip = EquipmentDay × Duration','1200'],
    ['MaterialTotal','Fixed one-off material cost — does NOT multiply by duration','50000'],[''],
    ['APPEARANCE COLUMNS'],
    ['Column','What it means','Example / Options'],
    ['OutlineStyle','Outline drawn around the activity bar','none · solid · dashed · dotted · dash-dot'],
    ['OutlineColour','Outline colour as hex','#000000'],
    ['OutlineWidth','Outline thickness in pixels (0.5–6)','1.5'],[''],
    ['DATE FORMATS ACCEPTED'],
    ['DD/MM/YYYY  e.g. 01/03/2026'],
    ['YYYY-MM-DD  e.g. 2026-03-01'],
    ['DD-Mon-YYYY e.g. 01-Mar-2026'],
    ['DD/MM/YY    e.g. 01/03/26 (2-digit year, 00–49 = 2000s)'],
  ];
  const wsInstr = XLSX.utils.aoa_to_sheet(instructions);
  wsInstr['!cols'] = [{wch:18},{wch:60},{wch:40}];
  XLSX.utils.book_append_sheet(wb, wsInstr, 'Instructions');

  // ── Sheet 2: Activities template ──
  const headers = [['Name','Start','End','StartCh','EndCh','P6ID','Colour','Shape','Alpha','Label',
    'Workers','LabourRate','EquipmentDay','MaterialTotal','OutlineStyle','OutlineColour','OutlineWidth','Notes']];
  const examples = [
    ['Example: Drainage Run A','01/03/2026','30/06/2026',100,450,'A1050','#3498db','rect',1,'no',
      8,350,800,25000,'dashed','#1a1a1a',1.5,'Delete this row before importing'],
    ['Example: Piling Milestone','15/04/2026','15/04/2026',200,200,'A1060','#e74c3c','flag',1,'right',
      0,0,0,0,'none','#000000',1.5,'Delete this row before importing'],
  ];
  const wsActs = XLSX.utils.aoa_to_sheet([...headers,...examples]);
  wsActs['!cols']=[
    {wch:28},{wch:13},{wch:13},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:7},{wch:8},
    {wch:10},{wch:12},{wch:14},{wch:15},{wch:14},{wch:14},{wch:14},{wch:30}
  ];
  XLSX.utils.book_append_sheet(wb, wsActs, 'Activities');

  XLSX.writeFile(wb, 'TC-Planner-Import-Template.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
//  SPREADSHEET ACTIVITY EDITOR — powered by Handsontable
//  Full Excel-like experience: drag-fill, paste from Excel,
//  right-click menus, keyboard nav, column resize, undo/redo
// ═══════════════════════════════════════════════════════════════════

let hotInstance = null;   // the Handsontable instance
/** Flexible date parser for HOT — accepts 2-digit years, validates day/month ranges */
function parseHotDate(val){
  if(!val) return null;
  const s = String(val).trim();
  // DD/MM/YYYY or DD/MM/YY
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if(dmy){
    const day=parseInt(dmy[1]), mon=parseInt(dmy[2]);
    let yr = parseInt(dmy[3]);
    if(yr < 100) yr += yr < 50 ? 2000 : 1900;
    if(mon<1||mon>12||day<1||day>31) return null;  // reject impossible dates
    const d=new Date(yr,mon-1,day);
    if(d.getMonth()!==mon-1) return null; // catches 31/02 etc
    return d;
  }
  // YYYY-MM-DD
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(iso){
    const d=new Date(+iso[1],+iso[2]-1,+iso[3]);
    if(d.getMonth()!==+iso[2]-1) return null;
    return d;
  }
  // DD-Mon-YYYY or DD-Mon-YY
  const dmy2 = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if(dmy2){
    const months={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const m = months[dmy2[2].toLowerCase()];
    if(m !== undefined){
      let yr = parseInt(dmy2[3]);
      if(yr < 100) yr += yr < 50 ? 2000 : 1900;
      const day=parseInt(dmy2[1]);
      if(day<1||day>31) return null;
      const d=new Date(yr,m,day);
      if(d.getMonth()!==m) return null;
      return d;
    }
  }
  return null;
}

const HOT_COLS = [
  { data:'p6Id',    title:'P6 ID',              width:90,  type:'text' },
  { data:'name',    title:'Name',               width:220, type:'text' },
  { data:'start',   title:'Start (DD/MM/YYYY)', width:130, type:'text',
    validator:(val,cb)=>cb(!val||val===''||!!parseHotDate(val)), allowInvalid:true },
  { data:'end',     title:'End (DD/MM/YYYY)',   width:130, type:'text',
    validator:(val,cb)=>cb(!val||val===''||!!parseHotDate(val)), allowInvalid:true },
  { data:'startCh',      title:'Start Ch',       width:85,  type:'numeric', numericFormat:{pattern:'0'} },
  { data:'endCh',        title:'End Ch',         width:85,  type:'numeric', numericFormat:{pattern:'0'} },
  { data:'resourceCount',title:'Workers/day',         width:105, type:'numeric', numericFormat:{pattern:'0'} },
  { data:'unitCost',     title:'Labour rate/day',      width:130, type:'numeric', numericFormat:{pattern:'0.00'} },
  { data:'equipmentCost',title:'Equipment hire/day',   width:140, type:'numeric', numericFormat:{pattern:'0.00'} },
  { data:'materialCost', title:'Material cost (total)',width:150, type:'numeric', numericFormat:{pattern:'0.00'} },
  { data:'category',     title:'Category',       width:130, type:'text', readOnly:true },
  { data:'subgroup',     title:'Sub-group',      width:130, type:'text', readOnly:true },
  { data:'color',        title:'Colour',         width:80,  type:'text' },
  { data:'alpha',        title:'Opacity %',      width:85,  type:'numeric', numericFormat:{pattern:'0'} },
  { data:'shape',        title:'Shape',          width:90,  type:'dropdown',
    source:['rect','line','circle','diamond','triangle','star','flag'] },
  { data:'label',        title:'Label',          width:80,  type:'dropdown',
    source:['no','left','center','right'] },
  { data:'progress',     title:'Progress %',     width:90,  type:'numeric', numericFormat:{pattern:'0'} },
  { data:'notes',        title:'Notes',          width:180, type:'text' },
  { data:'outlineStyle', title:'Outline',        width:90,  type:'dropdown',
    source:['none','solid','dashed','dotted','dash-dot'] },
  { data:'outlineColor', title:'Outline Colour', width:110, type:'text' },
  { data:'outlineWidth', title:'Outline Width',  width:100, type:'numeric', numericFormat:{pattern:'0.0'} },
];

/** Convert one activity to a flat HOT row object */
function actToHotRow(a){
  return {
    p6Id:    a.p6Id    || '',
    name:    a.name    || '',
    start:   a.start   ? formatDateInput(a.start)  : '',
    end:     a.end     ? formatDateInput(a.end)    : '',
    startCh: a.startCh ?? '',
    endCh:   a.endCh   ?? '',
    color:   a.color   || '#3498db',
    alpha:   a.alpha != null ? Math.round(a.alpha * 100) : 100,
    shape:   a.shape   || 'rect',
    label:   a.label   || 'no',
    progress:a.progress|| 0,
    notes:   a.notes   || '',
    outlineStyle: a.outlineStyle || 'none',
    outlineColor: a.outlineColor || '#000000',
    outlineWidth: a.outlineWidth || 1.5,
    resourceCount:a.resourceCount||0,
    unitCost:     a.unitCost||0,
    materialCost: a.materialCost||0,
    equipmentCost:a.equipmentCost||0,
    category:     '',  // filled dynamically in buildHotData
    subgroup:     '',
  };
}

/** Convert a HOT row back to an activity-compatible object */
function hotRowToAct(row, existing){
  const base = existing ? {...existing} : {groups:[], id:genId()};
  base.p6Id    = row.p6Id    || '';
  base.name    = row.name    || '';
  base.start   = row.start   ? parseHotDate(String(row.start))   : null;
  base.end     = row.end     ? parseHotDate(String(row.end))     : null;
  base.startCh = parseFloat(row.startCh) || 0;
  base.endCh   = parseFloat(row.endCh)   || 0;
  base.color   = row.color   || '#3498db';
  // Alpha: stored as 0-1 internally, shown as 0-100 in grid
  const rawAlpha = parseFloat(row.alpha);
  base.alpha = isNaN(rawAlpha) ? 1 : (rawAlpha > 1 ? rawAlpha / 100 : rawAlpha);
  base.alpha = Math.max(0.05, Math.min(1, base.alpha));
  base.shape   = row.shape   || 'rect';
  base.label   = row.label   || 'no';
  base.progress= parseFloat(row.progress)|| 0;
  base.notes        = row.notes        || '';
  base.outlineStyle = row.outlineStyle || 'none';
  base.outlineColor = row.outlineColor || '#000000';
  base.outlineWidth = parseFloat(row.outlineWidth) || 1.5;
  base.resourceCount= parseFloat(row.resourceCount)||0;
  base.unitCost     = parseFloat(row.unitCost)||0;
  base.materialCost = parseFloat(row.materialCost)||0;
  base.equipmentCost= parseFloat(row.equipmentCost)||0;
  return base;
}

// Tracks which HOT column indices are currently hidden
let hotHiddenCols = new Set();

// Column tooltips — shown on header hover
const HOT_COL_TOOLTIPS = {
  p6Id:         'Primavera P6 Activity ID. Used by the Sync Dates tool to match activities when importing updated dates from P6.',
  name:         'Activity name shown on the chart.',
  start:        'Start date. Format: DD/MM/YYYY. Must be before the End date.',
  end:          'End / Finish date. Format: DD/MM/YYYY. Must be after the Start date.',
  startCh:      'Start chainage — the distance (in your project units) where this activity begins on the route.',
  endCh:        'End chainage — where the activity ends. Can equal Start Ch for a point event.',
  color:        'Fill colour as a hex code e.g. #e74c3c. Leave blank to use the default blue.',
  alpha:        'Opacity as a percentage (0–100). 100 = fully solid, 50 = semi-transparent.',
  shape:        'Bar shape on the chart. rect = standard bar, diamond/star/flag = milestone point.',
  label:        'Whether to show the activity name as a label on the chart, and where.',
  progress:     'Percentage complete (0–100). Shown as a dark overlay on the activity bar.',
  notes:        'Free-text notes. Shown in the tooltip when you hover over the activity on the chart.',
  outlineStyle: 'Outline style around the activity bar. Useful for distinguishing overlapping activities.',
  outlineColor: 'Outline colour as a hex code e.g. #000000.',
  outlineWidth: 'Outline thickness in pixels (0.5 – 6).',
  resourceCount:'Number of workers assigned to this activity per day.',
  unitCost:     'Labour rate per worker per day. Total labour cost = this × Workers/day × Duration.',
  equipmentCost:'Daily plant or equipment hire rate. Total equipment cost = this × Duration.',
  materialCost: 'Fixed one-off material purchase cost. This does NOT multiply by duration — enter the total material spend for the activity.',
  category:     'Read-only. Shows the first Category (group) this activity belongs to.',
  subgroup:     'Read-only. Shows the first Sub-group this activity belongs to.',
};

// Sort state
let hotSortCol   = null;   // data key being sorted
let hotSortDir   = 'asc';  // 'asc' | 'desc'
let hotFilterVals = {};    // {dataKey: filterString}

function initHot(){
  if(hotInstance) return;
  const container = document.getElementById('hotContainer');
  if(!container || typeof Handsontable === 'undefined') return;

  hotInstance = new Handsontable(container, {
    data:           buildHotData(),
    columns:        _visibleHotCols(),
    colHeaders:     (ci)=>_hotColHeader(ci),
    rowHeaders:     true,
    width:          '100%',
    height:         520,
    stretchH:       'none',       // Don't stretch last col — allows horizontal scroll
    autoWrapRow:    false,
    autoWrapCol:    false,
    manualColumnResize: true,
    fillHandle:     true,
    allowInsertRow: true,
    allowRemoveRow: true,
    copyPaste:      { pasteMode:'overwrite', rowsLimit:5000 },
    undo:           true,
    outsideClickDeselects: false,
    licenseKey:     'non-commercial-and-evaluation',

    // Right-click context menu — rows + column visibility
    contextMenu: {
      items: {
        'row_above':  {name:'Insert row above'},
        'row_below':  {name:'Insert row below'},
        'remove_row': {name:'Delete row'},
        'sep1': Handsontable.plugins.ContextMenu.SEPARATOR,
        'copy': {name:'Copy'},
        'cut':  {name:'Cut'},
        'sep2': Handsontable.plugins.ContextMenu.SEPARATOR,
        'hide_col': {
          name(){ return 'Hide this column'; },
          callback(key,[{start}]){
            const colIdx=start.col;
            const visKeys=_visibleColKeys();
            const dataKey=visKeys[colIdx];
            if(dataKey&&dataKey!=='name'){  // never hide Name
              hotHiddenCols.add(dataKey);
              hotDestroyAndReinit();
            }
          },
          disabled(){ return false; }
        },
        'sep3': Handsontable.plugins.ContextMenu.SEPARATOR,
        'show_cols': {
          name(){ return hotHiddenCols.size>0?`Show hidden columns (${hotHiddenCols.size})`:'No hidden columns'; },
          callback(){
            hotHiddenCols.clear();
            hotDestroyAndReinit();
          },
          disabled(){ return hotHiddenCols.size===0; }
        },
        'sep4': Handsontable.plugins.ContextMenu.SEPARATOR,
        'sort_asc': {
          name(){ return 'Sort A → Z'; },
          callback(key,[{start}]){
            const visKeys=_visibleColKeys();
            hotSortCol=visKeys[start.col];hotSortDir='asc';
            hotRefresh();
          }
        },
        'sort_desc': {
          name(){ return 'Sort Z → A'; },
          callback(key,[{start}]){
            const visKeys=_visibleColKeys();
            hotSortCol=visKeys[start.col];hotSortDir='desc';
            hotRefresh();
          }
        },
        'clear_sort': {
          name(){ return 'Reset sort & filters'; },
          callback(){ hotSortCol=null;hotFilterVals={};hotRefresh(); }
        },
      }
    },

    afterChange(changes, source){
      if(!changes || source==='loadData'||source==='revert') return;
      let blocked=false;
      changes.forEach(([row,prop,oldVal,newVal])=>{
        if((prop==='start'||prop==='end')&&newVal){
          const d=hotInstance.getSourceDataAtRow(row);
          const s=parseHotDate(String(d.start||''));
          const e=parseHotDate(String(d.end||''));
          if(s&&e&&e<s){
            setTimeout(()=>{
              alert(`⚠️ Row ${row+1}: End date cannot be before Start date. Reverting.`);
              hotInstance.setDataAtRowProp(row,prop,oldVal,'revert');
            },50);
            blocked=true;
          }
        }
      });
      if(!blocked) hotSyncToActivities();
    },
    afterRemoveRow(){ hotSyncToActivities(); },

    afterRender(){
      _attachHotHeaderTooltips();
      if(!document.getElementById('hotFilterBar')) _attachHotFilterRow();
    },

    afterColumnResize(newSize, column){
      // Update the HOT_COLS width so filter bar rebuilds with correct widths
      const visKeys = _visibleColKeys();
      const key = visKeys[column];
      if(key){
        const col = HOT_COLS.find(c=>c.data===key);
        if(col) col.width = newSize;
      }
      _rebuildHotFilterBar();
    },
  });
}

function hotDestroyAndReinit(){
  const bar=document.getElementById('hotFilterBar');if(bar)bar.remove();
  if(hotInstance){hotInstance.destroy();hotInstance=null;}
  initHot();
  hotRefresh();
  _rebuildHotFilterBar();
}

function _visibleColKeys(){
  return HOT_COLS.filter(c=>!hotHiddenCols.has(c.data)).map(c=>c.data);
}

function _visibleHotCols(){
  return HOT_COLS.filter(c=>!hotHiddenCols.has(c.data));
}

function _hotColHeader(ci){
  const visKeys=_visibleColKeys();
  const key=visKeys[ci];
  if(!key)return'';
  const col=HOT_COLS.find(c=>c.data===key);
  const label=col?col.title:key;
  // Sort indicator
  const sortIndicator = hotSortCol===key ? (hotSortDir==='asc'?' ↑':' ↓') : '';
  // Filter indicator
  const filterIndicator = hotFilterVals[key] ? ' 🔍' : '';
  return label+sortIndicator+filterIndicator;
}

function _attachHotHeaderTooltips(){
  const container=document.getElementById('hotContainer');
  if(!container)return;
  container.querySelectorAll('.ht_clone_top th').forEach((th,ci)=>{
    if(ci===0)return; // row header
    const visKeys=_visibleColKeys();
    const key=visKeys[ci-1];
    if(!key)return;
    const tip=HOT_COL_TOOLTIPS[key];
    if(tip){th.title=tip;th.style.cursor='help';}
  });
}

/** Render the filter bar above the HOT container — completely separate from HOT DOM */
function _attachHotFilterRow(){
  const hotContainer = document.getElementById('hotContainer');
  const ssMode = document.getElementById('editorSSMode');
  if(!hotContainer || !ssMode) return;
  if(document.getElementById('hotFilterBar')) return;

  const bar = document.createElement('div');
  bar.id = 'hotFilterBar';
  bar.style.cssText = [
    'display:flex','align-items:stretch',
    'background:var(--bg-secondary)','border:1px solid var(--border)',
    'border-bottom:2px solid var(--tab-active-bg)',
    'overflow-x:auto','flex-shrink:0','font-size:11px',
    'scrollbar-width:none',  // hide scrollbar on filter bar itself
  ].join(';');

  // Reset cell
  const resetCell = document.createElement('div');
  resetCell.style.cssText = 'display:flex;align-items:center;justify-content:center;min-width:50px;width:50px;flex-shrink:0;border-right:1px solid var(--border);padding:3px;';
  resetCell.innerHTML = `<button onclick="hotClearFilters()" title="Clear all filters and sort" style="margin:0;padding:2px 5px;font-size:10px;background:var(--border);color:var(--text);border:none;border-radius:3px;cursor:pointer">✕ Reset</button>`;
  bar.appendChild(resetCell);

  _visibleColKeys().forEach(key=>{
    const col = HOT_COLS.find(c=>c.data===key);
    const w = col ? col.width : 80;
    const cell = document.createElement('div');
    cell.dataset.filterKey = key;
    cell.style.cssText = `min-width:${w}px;width:${w}px;flex-shrink:0;border-right:1px solid var(--border);padding:2px 3px;display:flex;align-items:center;box-sizing:border-box;`;
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = 'Filter…';
    inp.value = hotFilterVals[key] || '';
    inp.style.cssText = 'width:100%;font-size:10px;padding:2px 3px;border:1px solid var(--input-border);border-radius:2px;background:var(--input-bg);color:var(--text);box-sizing:border-box;';
    inp.oninput = ()=>{ hotFilterVals[key]=inp.value.trim().toLowerCase(); hotRefresh(); };
    cell.appendChild(inp);
    bar.appendChild(cell);
  });

  // Sync horizontal scroll with HOT master table
  const masterHolder = hotContainer.querySelector('.ht_master .wtHolder');
  if(masterHolder){
    masterHolder.addEventListener('scroll', ()=>{ bar.scrollLeft = masterHolder.scrollLeft; });
    bar.addEventListener('scroll', ()=>{ masterHolder.scrollLeft = bar.scrollLeft; });
  }

  ssMode.insertBefore(bar, hotContainer);
}

function _rebuildHotFilterBar(){
  const existing = document.getElementById('hotFilterBar');
  if(existing) existing.remove();
  _attachHotFilterRow();
}

function hotClearFilters(){
  hotSortCol=null;hotSortDir='asc';hotFilterVals={};hotRefresh();
}


// Maps displayed HOT row index → original activities[] index (for sort/filter)
let hotRowMap = [];

function buildHotData(){
  hotRowMap = [];
  let rows = activities.map((a, origIdx) => {
    const r = actToHotRow(a);
    r._origIdx = origIdx;
    return r;
  });

  // Sort
  if(hotSortCol){
    rows.sort((a,b)=>{
      let av=a[hotSortCol]??'', bv=b[hotSortCol]??'';
      if(typeof av==='number'&&typeof bv==='number') return hotSortDir==='asc'?av-bv:bv-av;
      av=String(av).toLowerCase(); bv=String(bv).toLowerCase();
      return hotSortDir==='asc'?av.localeCompare(bv):bv.localeCompare(av);
    });
  }

  // Filter
  Object.entries(hotFilterVals).forEach(([key,fv])=>{
    if(!fv)return;
    rows=rows.filter(r=>String(r[key]??'').toLowerCase().includes(fv));
  });

  // Resolve category & subgroup
  rows = rows.map(r=>{
    const a = activities[r._origIdx];
    let cat='', sg='';
    if(a&&a.groups&&a.groups.length){
      for(const gid of a.groups){
        const g=groups.find(x=>x.id===gid);
        if(!g)continue;
        if(g.parentId===null&&!cat) cat=g.name;
        if(g.parentId!==null&&!sg)  sg=g.name;
      }
    }
    hotRowMap.push(r._origIdx);
    // Remove _origIdx from the actual data object so HOT doesn't see it
    const {_origIdx, ...clean} = {...r, category:cat, subgroup:sg};
    return clean;
  });

  // Blank rows for new entry — no _origIdx
  for(let i=0;i<10;i++){
    hotRowMap.push(-1);
    rows.push({
      p6Id:'',name:'',start:'',end:'',startCh:'',endCh:'',
      resourceCount:'',unitCost:'',category:'',subgroup:'',
      color:'',alpha:'',shape:'',label:'',progress:'',notes:'',
      outlineStyle:'',outlineColor:'',outlineWidth:'',
    });
  }
  return rows;
}

/** Push HOT grid data back into the activities array */
function hotSyncToActivities(){
  if(!hotInstance) return;
  const data = hotInstance.getData();
  const visKeys = _visibleColKeys();
  const newActs = [];

  data.forEach((row, ri)=>{
    const obj={};
    visKeys.forEach((k,ci)=>{ obj[k]=row[ci]; });
    // Skip entirely blank rows
    if(!obj.name&&!obj.p6Id&&!obj.start&&!obj.end&&(obj.startCh===''||obj.startCh==null)&&(obj.endCh===''||obj.endCh==null)) return;
    const origIdx = hotRowMap[ri] != null ? hotRowMap[ri] : -1;
    const existing = origIdx>=0 ? activities[origIdx] : null;
    newActs.push(hotRowToAct(obj, existing));
  });

  pushUndo();
  activities = newActs;
  drawChart(); renderActivityTable(); populateCpTargetSelect();
}

/** Reload the HOT grid from the activities array (e.g. after undo/load) */
function hotRefresh(){
  if(!hotInstance) return;
  hotInstance.updateSettings({
    data: buildHotData(),
    columns: _visibleHotCols(),
    colHeaders: (ci)=>_hotColHeader(ci),
  });
  hotInstance.render();
}

function hotAddRows(n){
  if(!hotInstance) return;
  const lastRow = hotInstance.countRows();
  hotInstance.alter('insert_row_below', lastRow-1, n);
}

function hotClearEmptyRows(){
  if(!hotInstance) return;
  const data    = hotInstance.getData();
  // Find rows where every meaningful cell is blank
  const toRemove = [];
  data.forEach((row, ri)=>{
    const p6Id  = String(row[0]||'').trim();
    const name  = String(row[1]||'').trim();
    const start = String(row[2]||'').trim();
    const end   = String(row[3]||'').trim();
    const sCh   = String(row[4]||'').trim();
    const eCh   = String(row[5]||'').trim();
    if(!p6Id && !name && !start && !end && !sCh && !eCh) toRemove.push(ri);
  });
  if(!toRemove.length){ alert('No empty rows to remove.'); return; }
  // Remove from bottom up so indices stay valid
  toRemove.reverse().forEach(ri=>{
    hotInstance.alter('remove_row', ri, 1);
  });
  hotSyncToActivities();
}

function toggleEditorMode(mode){
  const formMode = document.getElementById('editorFormMode');
  const ssMode   = document.getElementById('editorSSMode');
  const btnForm  = document.getElementById('editorToggleForm');
  const btnSS    = document.getElementById('editorToggleSS');
  if(mode==='form'){
    formMode.style.display='block';
    ssMode.style.display='none';
    btnForm.classList.add('active');btnSS.classList.remove('active');
    const bar=document.getElementById('hotFilterBar');if(bar)bar.remove();
  } else {
    formMode.style.display='none';
    ssMode.style.display='block';
    btnForm.classList.remove('active');btnSS.classList.add('active');
    // Guard: if no activities yet, show a helpful placeholder instead of letting HOT error
    const hotContainer=document.getElementById('hotContainer');
    if(!activities.length && hotContainer && !hotInstance){
      hotContainer.innerHTML=`<div style="padding:32px;text-align:center;color:var(--text-muted)">
        <div style="font-size:32px;margin-bottom:12px">📋</div>
        <div style="font-size:14px;font-weight:600;margin-bottom:6px">No activities yet</div>
        <div style="font-size:13px">Add your first activity using the <strong>Form Editor</strong> above,<br>then switch back to Grid view to edit in bulk.</div>
      </div>`;
      return;
    }
    if(hotContainer) hotContainer.innerHTML=''; // clear placeholder if activities now exist
    initHot();
    hotRefresh();
    requestAnimationFrame(()=>hotInstance?.render());
  }
}

// ═══════════════════════════════════════════════════════════════════
//  RESOURCE & COST VIEWS
// ═══════════════════════════════════════════════════════════════════

function renderResourceTab(){
  const cards=document.getElementById('resourceSummaryCards');
  const table=document.getElementById('resourceTableContainer');
  if(!cards||!table)return;

  // Respect the current view filter
  const range=getViewDateRange();
  const filterLabel=viewFilter==='all'?'Full Timeline':viewFilter.toUpperCase()+' window';

  const withRes=activities.filter(a=>{
    if(!a.resourceCount&&!a.unitCost&&!a.materialCost&&!a.equipmentCost)return false;
    if(range&&a.start&&a.end&&(a.end<range.from||a.start>range.to))return false;
    return true;
  });

  if(!withRes.length){
    cards.innerHTML='';
    table.innerHTML=`<p style="color:var(--text-muted);font-size:13px">No resource or cost data ${range?'in the <strong>'+filterLabel+'</strong>':''} — add values to activities in the Activity Editor or Grid view.</p>`;
    return;
  }

  const fmtCost=v=>{const sym=projectCurrency||'£';return v>=1e6?sym+(v/1e6).toFixed(2)+'M':v>=1e3?sym+(v/1e3).toFixed(0)+'k':sym+Math.round(v);};
  let totalWorkerDays=0,totalLabour=0,totalMaterial=0,totalEquip=0;

  withRes.forEach(a=>{
    const dur=a.start&&a.end?Math.max(0,(a.end-a.start)/86400000):0;
    const res      =parseFloat(a.resourceCount)||0;
    const labRate  =parseFloat(a.unitCost)||0;
    const equipDay =parseFloat(a.equipmentCost)||0;
    const matTotal =parseFloat(a.materialCost)||0;
    totalWorkerDays += res * dur;
    totalLabour     += res * labRate * dur;
    totalEquip      += equipDay * dur;
    totalMaterial   += matTotal;          // fixed — does not multiply by duration
  });
  const totalCost=totalLabour+totalMaterial+totalEquip;

  cards.innerHTML=`
    <div class="resource-card">
      <div class="resource-card-value">${Math.round(totalWorkerDays).toLocaleString()}</div>
      <div class="resource-card-label">👷 Total Worker-Days</div>
    </div>
    <div class="resource-card" style="border-left-color:#3498db">
      <div class="resource-card-value">${fmtCost(totalLabour)}</div>
      <div class="resource-card-label">💼 Total Labour Cost</div>
    </div>
    <div class="resource-card" style="border-left-color:#e67e22">
      <div class="resource-card-value">${fmtCost(totalMaterial)}</div>
      <div class="resource-card-label">🧱 Total Material Cost (fixed)</div>
    </div>
    <div class="resource-card" style="border-left-color:#8e44ad">
      <div class="resource-card-value">${fmtCost(totalEquip)}</div>
      <div class="resource-card-label">🚜 Total Equipment Cost</div>
    </div>
    <div class="resource-card resource-card--cost">
      <div class="resource-card-value">${fmtCost(totalCost)}</div>
      <div class="resource-card-label">💰 Total Project Cost</div>
    </div>
  `;

  let html=`<table class="activityTable" style="font-size:12px"><thead><tr>
    <th>Activity</th><th>P6 ID</th><th>Duration</th>
    <th title="Number of workers on this activity">👷 Workers</th>
    <th title="Workers × Duration">Worker-days</th>
    <th title="Cost per worker per day × workers × duration">💼 Labour total</th>
    <th title="Fixed one-off material purchase — does NOT scale with duration">🧱 Material (fixed)</th>
    <th title="Daily hire rate × duration">🚜 Equipment total</th>
    <th>💰 Activity total</th>
  </tr></thead><tbody>`;

  withRes.forEach(a=>{
    const dur     = a.start&&a.end ? Math.round(Math.max(0,(a.end-a.start)/86400000)) : 0;
    const res     = parseFloat(a.resourceCount)||0;
    const labRate = parseFloat(a.unitCost)||0;
    const equipDay= parseFloat(a.equipmentCost)||0;
    const matTotal= parseFloat(a.materialCost)||0;
    const wd      = Math.round(res*dur);
    const labTot  = res*labRate*dur;
    const equipTot= equipDay*dur;
    const actTotal= labTot+matTotal+equipTot;
    html+=`<tr>
      <td>${a.name}</td>
      <td>${a.p6Id||'—'}</td>
      <td>${dur}d</td>
      <td>${res||'—'}</td>
      <td>${wd?wd.toLocaleString():'—'}</td>
      <td>${labTot>0?fmtCost(labTot):'—'}</td>
      <td>${matTotal>0?fmtCost(matTotal):'—'}</td>
      <td>${equipTot>0?fmtCost(equipTot):'—'}</td>
      <td><strong>${actTotal>0?fmtCost(actTotal):'—'}</strong></td>
    </tr>`;
  });
  html+=`<tr style="font-weight:bold;background:var(--bg-secondary)">
    <td colspan="4">TOTAL</td>
    <td>${Math.round(totalWorkerDays).toLocaleString()}</td>
    <td>${fmtCost(totalLabour)}</td>
    <td>${fmtCost(totalMaterial)}</td>
    <td>${fmtCost(totalEquip)}</td>
    <td>${fmtCost(totalCost)}</td>
  </tr>
  </tbody></table>`;
  table.innerHTML=html;
}


function setChartViewMode(mode){
  chartViewMode=mode;
  ['viewBtn-programme','viewBtn-resource','viewBtn-cost','viewBtn-combined'].forEach(id=>{
    const el=document.getElementById(id);if(el)el.classList.remove('active');
  });
  const btn=document.getElementById('viewBtn-'+mode);if(btn)btn.classList.add('active');
  if(mode==='programme')      drawChart();
  else if(mode==='resource')  drawResourceChart();
  else if(mode==='cost')      drawCostChart();
  else if(mode==='combined')  drawCombinedChart();
}

// Stored hit regions for cost/resource chart hover tooltips
let chartHitRegions = [];  // [{type:'bar'|'point', x,y,w,h, data:{}}]

function _storeBarHit(x,y,w,h,data){
  chartHitRegions.push({type:'bar',x,y,w,h,data});
}
function _storePointHit(x,y,data){
  chartHitRegions.push({type:'point',x,y,r:10,data});
}

function _clearChartHits(){chartHitRegions=[];}

function _fmtC(v){const sym=projectCurrency||'£';return v>=1e6?sym+(v/1e6).toFixed(2)+'M':v>=1e3?sym+(v/1e3).toFixed(0)+'k':sym+Math.round(v);}

function _checkChartHover(e){
  if(chartViewMode==='programme')return; // programme has its own tooltip
  const rect=canvas.getBoundingClientRect();
  const mx=e.clientX-rect.left, my=e.clientY-rect.top;
  let found=null;
  for(const r of chartHitRegions){
    if(r.type==='bar'&&mx>=r.x&&mx<=r.x+r.w&&my>=r.y&&my<=r.y+r.h){found=r;break;}
    if(r.type==='point'&&Math.hypot(mx-r.x,my-r.y)<=r.r){found=r;break;}
  }
  if(found){
    const d=found.data;
    let html=`<strong>${d.label||''}</strong>`;
    if(d.type==='bar'){
      html+=`<br>🧱 Labour: ${_fmtC(d.labour||0)}`;
      html+=`<br>🔧 Material: ${_fmtC(d.material||0)}`;
      html+=`<br>🚜 Equipment: ${_fmtC(d.equip||0)}`;
      html+=`<hr style="border:none;border-top:1px solid rgba(255,255,255,0.3);margin:3px 0">`;
      html+=`<strong>Monthly total: ${_fmtC(d.total||0)}</strong>`;
    } else if(d.type==='point'){
      html+=`<br>Cumulative: <strong>${_fmtC(d.cumulative||0)}</strong>`;
      if(d.monthly) html+=`<br>This month: ${_fmtC(d.monthly)}`;
    } else if(d.type==='resource'){
      html+=`<br>Worker-days: <strong>${Math.round(d.resources||0).toLocaleString()}</strong>`;
    }
    tooltip.innerHTML=html;
    tooltip.style.display='block';
    tooltip.style.left=(e.pageX+12)+'px';
    tooltip.style.top =(e.pageY+10)+'px';
  } else {
    tooltip.style.display='none';
  }
}

function buildMonthBuckets(){
  if(!timelineStart||!timelineEnd)return[];
  const range=getViewDateRange();
  const from=range?range.from:timelineStart;
  const to  =range?range.to  :timelineEnd;
  const buckets=[];
  let y=from.getFullYear(),m=from.getMonth();
  const ey=to.getFullYear(),em=to.getMonth();
  while(y<ey||(y===ey&&m<=em)){
    buckets.push({year:y,month:m,
      label:new Date(y,m,1).toLocaleDateString(undefined,{month:'short',year:'2-digit'}),
      resources:0,cost:0,labourCost:0,materialCost:0,equipCost:0});
    m++;if(m>11){m=0;y++;}
  }
  return buckets;
}

function buildResourceData(){
  const buckets=buildMonthBuckets();if(!buckets.length)return buckets;
  const range=getViewDateRange();
  const totalDurCache={};

  activities.forEach(a=>{
    if(!a.start||!a.end)return;
    if(range&&(a.end<range.from||a.start>range.to))return;

    const res          = parseFloat(a.resourceCount)||0;  // workers/day
    const labourRate   = parseFloat(a.unitCost)||0;        // cost per worker per day
    const equipPerDay  = parseFloat(a.equipmentCost)||0;   // equipment hire per day
    const materialTotal= parseFloat(a.materialCost)||0;    // fixed one-off material cost

    if(!res && !labourRate && !equipPerDay && !materialTotal) return;

    // Total duration of this activity in days (for spreading fixed material cost)
    const totalDur = Math.max(1, (a.end - a.start) / 86400000);

    buckets.forEach(b=>{
      const bStart = new Date(b.year, b.month, 1);
      const bEnd   = new Date(b.year, b.month+1, 0, 23, 59, 59);
      const effStart = range ? new Date(Math.max(a.start.getTime(), range.from.getTime())) : a.start;
      const effEnd   = range ? new Date(Math.min(a.end.getTime(),   range.to.getTime()))   : a.end;
      const oStart = new Date(Math.max(effStart.getTime(), bStart.getTime()));
      const oEnd   = new Date(Math.min(effEnd.getTime(),   bEnd.getTime()));
      if(oEnd < oStart) return;

      const days = Math.max(0, (oEnd - oStart) / 86400000);
      const frac = days / totalDur; // proportion of activity in this bucket

      // Labour cost = workers × rate × days
      const labourCostBucket  = res * labourRate * days;
      // Equipment = daily hire × days
      const equipCostBucket   = equipPerDay * days;
      // Material = spread proportionally across activity duration (fixed total ÷ total days × bucket days)
      const materialCostBucket = materialTotal * frac;

      b.resources    += res * days;
      b.labourCost   += labourCostBucket;
      b.equipCost    += equipCostBucket;
      b.materialCost += materialCostBucket;
      b.cost         += labourCostBucket + equipCostBucket + materialCostBucket;
    });
  });
  return buckets;
}

function _drawViewAxes(ctx,ct,W,H,pad,maxY,yTicks,yFmt,title,yLabel){
  ctx.fillStyle=ct.canvasBg;ctx.fillRect(0,0,W,H);
  ctx.fillStyle=ct.axisText;ctx.font='bold 14px Arial';ctx.textAlign='center';ctx.textBaseline='top';
  ctx.fillText(title,W/2,8);
  const gW=W-pad.left-pad.right,gH=H-pad.top-pad.bottom;
  for(let i=0;i<=yTicks;i++){
    const val=maxY/yTicks*i;
    const y=pad.top+gH-(i/yTicks)*gH;
    ctx.strokeStyle=ct.gridH;ctx.lineWidth=1;
    ctx.beginPath();ctx.moveTo(pad.left,y);ctx.lineTo(W-pad.right,y);ctx.stroke();
    ctx.fillStyle=ct.axisText;ctx.font='11px Arial';ctx.textAlign='right';ctx.textBaseline='middle';
    ctx.fillText(yFmt(val),pad.left-6,y);
  }
  ctx.fillStyle=ct.axisText;ctx.font='12px Arial';
  ctx.save();ctx.translate(14,pad.top+gH/2);ctx.rotate(-Math.PI/2);
  ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(yLabel,0,0);ctx.restore();
  return{gW,gH};
}

function drawResourceChart(){
  if(!timelineStart||!timelineEnd)return;
  const buckets=buildResourceData();if(!buckets.length)return;
  _clearChartHits();
  const ct=getChartTheme();
  const W=canvas.parentElement?.offsetWidth||canvas.width||800;
  const H=canvas.height||canvas.parentElement?.offsetHeight||500;
  canvas.width=W;canvas.height=H;
  const pad={top:52,right:60,bottom:70,left:80};
  const maxRes=Math.max(1,...buckets.map(b=>b.resources));
  const {gW,gH}=_drawViewAxes(ctx,ct,W,H,pad,maxRes,5,v=>Math.round(v),
    `Resource Histogram — Workers × Days per Month (${viewFilter==='all'?'Full Timeline':viewFilter})`,'Workers × Days');
  const bSlot=gW/buckets.length;
  const barW=Math.max(4,bSlot*0.72);
  buckets.forEach((b,i)=>{
    const x=pad.left+i*bSlot+bSlot*0.14;
    const bH=(b.resources/maxRes)*gH;
    const y=pad.top+gH-bH;
    ctx.globalAlpha=0.82;ctx.fillStyle='#3498db';ctx.fillRect(x,y,barW,bH);
    ctx.globalAlpha=1;ctx.strokeStyle='#1a6fa0';ctx.lineWidth=0.7;ctx.strokeRect(x,y,barW,bH);
    if(bH>14){ctx.fillStyle='#fff';ctx.font='bold 9px Arial';ctx.textAlign='center';ctx.textBaseline='middle';
      ctx.fillText(Math.round(b.resources),x+barW/2,y+bH/2);}
    ctx.fillStyle=ct.axisText;ctx.font='10px Arial';ctx.textAlign='center';ctx.textBaseline='top';
    ctx.fillText(b.label,x+barW/2,pad.top+gH+4);
    _storeBarHit(x,y,barW,bH,{type:'resource',label:b.label,resources:b.resources});
  });
  const totalRes=buckets.reduce((s,b)=>s+b.resources,0);
  ctx.fillStyle=ct.axisText;ctx.font='bold 12px Arial';ctx.textAlign='right';ctx.textBaseline='top';
  ctx.fillText('Total: '+Math.round(totalRes).toLocaleString()+' worker-days',W-pad.right,pad.top-4);
}

function drawCostChart(){
  if(!timelineStart||!timelineEnd)return;
  const buckets=buildResourceData();if(!buckets.length)return;
  _clearChartHits();
  const ct=getChartTheme();
  const W=canvas.parentElement?.offsetWidth||canvas.width||800;
  const H=canvas.height||Math.max(420,canvas.parentElement?.offsetHeight||420);
  canvas.width=W;canvas.height=H;
  const pad={top:54,right:120,bottom:70,left:90};
  let cum=0;
  const pts=buckets.map((b,i)=>{cum+=b.cost;return{x:i,y:cum,label:b.label,labour:b.labourCost,material:b.materialCost,equip:b.equipCost,monthly:b.cost};});
  const maxCost=Math.max(1,pts[pts.length-1]?.y||1);
  const {gW,gH}=_drawViewAxes(ctx,ct,W,H,pad,maxCost,5,_fmtC,
    `S-Curve — Cumulative Cost (${projectCurrency||'£'}) — ${viewFilter==='all'?'Full Timeline':viewFilter}`,`Cumulative Cost`);
  if(pts.length<2)return;
  const xOf=i=>pad.left+(i/(pts.length-1))*gW;
  const yOf=y=>pad.top+gH-(y/maxCost)*gH;
  const maxMonthly=Math.max(1,...buckets.map(b=>b.cost));
  const bSlot=gW/pts.length;
  const barW=Math.max(2,bSlot*0.7);
  // Stacked bars
  buckets.forEach((b,i)=>{
    const x=pad.left+i*bSlot+bSlot*0.15;
    const totalH=(b.cost/maxMonthly)*gH*0.45;
    if(totalH<1)return;
    let yBase=pad.top+gH;
    [{v:b.labourCost,c:'#3498db'},{v:b.materialCost,c:'#e67e22'},{v:b.equipCost,c:'#8e44ad'}].forEach(s=>{
      if(!s.v)return;
      const h=(s.v/b.cost)*totalH;yBase-=h;
      ctx.globalAlpha=0.45;ctx.fillStyle=s.c;ctx.fillRect(x,yBase,barW,h);ctx.globalAlpha=1;
    });
    _storeBarHit(x,pad.top+gH-totalH,barW,totalH,{type:'bar',label:b.label,labour:b.labourCost,material:b.materialCost,equip:b.equipCost,total:b.cost});
  });
  // S-curve area
  ctx.beginPath();ctx.moveTo(xOf(0),pad.top+gH);
  pts.forEach(p=>ctx.lineTo(xOf(p.x),yOf(p.y)));
  ctx.lineTo(xOf(pts.length-1),pad.top+gH);ctx.closePath();
  ctx.fillStyle='#27ae60';ctx.globalAlpha=0.12;ctx.fill();ctx.globalAlpha=1;
  ctx.beginPath();pts.forEach((p,i)=>i===0?ctx.moveTo(xOf(p.x),yOf(p.y)):ctx.lineTo(xOf(p.x),yOf(p.y)));
  ctx.strokeStyle='#27ae60';ctx.lineWidth=2.5;ctx.stroke();
  pts.forEach(p=>{
    const px=xOf(p.x),py=yOf(p.y);
    ctx.beginPath();ctx.arc(px,py,3.5,0,Math.PI*2);ctx.fillStyle='#27ae60';ctx.fill();
    _storePointHit(px,py,{type:'point',label:p.label,cumulative:p.y,monthly:p.monthly});
  });
  const step=Math.max(1,Math.floor(pts.length/12));
  pts.forEach((p,i)=>{if(i%step===0||i===pts.length-1){ctx.fillStyle=ct.axisText;ctx.font='10px Arial';ctx.textAlign='center';ctx.textBaseline='top';ctx.fillText(p.label,xOf(p.x),pad.top+gH+4);}});
  const last=pts[pts.length-1];
  ctx.fillStyle='#27ae60';ctx.font='bold 12px Arial';ctx.textAlign='left';ctx.textBaseline='middle';
  ctx.fillText('Total: '+_fmtC(last.y),Math.min(xOf(last.x)+8,W-pad.right-10),yOf(last.y));
  const legend=[['💼 Labour','#3498db'],['🧱 Material','#e67e22'],['🚜 Equipment','#8e44ad'],['S-Curve','#27ae60']];
  let lx=pad.left;
  legend.forEach(([name,c])=>{ctx.fillStyle=c;ctx.globalAlpha=0.75;ctx.fillRect(lx,10,14,12);ctx.globalAlpha=1;ctx.fillStyle=ct.axisText;ctx.font='11px Arial';ctx.textAlign='left';ctx.textBaseline='top';ctx.fillText(name,lx+17,10);lx+=ctx.measureText(name).width+42;});
}


// ── Combined Resources + S-Curve ──────────────────────
function drawCombinedChart(){
  if(!timelineStart||!timelineEnd)return;
  const buckets=buildResourceData();if(!buckets.length)return;
  _clearChartHits();
  const ct=getChartTheme();
  const W=canvas.parentElement?.offsetWidth||canvas.width||800;
  const H=canvas.height||Math.max(500,canvas.parentElement?.offsetHeight||500);
  canvas.width=W;canvas.height=H;
  const splitY=Math.floor(H*0.52);
  const padR={top:40,right:60,bottom:12,left:80};
  const padC={top:14,right:60,bottom:56,left:80};
  ctx.fillStyle=ct.canvasBg;ctx.fillRect(0,0,W,H);
  ctx.fillStyle=ct.axisText;ctx.font='bold 13px Arial';ctx.textAlign='center';ctx.textBaseline='top';
  ctx.fillText(`Resources & Cost — Combined View (${viewFilter==='all'?'Full Timeline':viewFilter})`,W/2,6);

  // ─ Resource bars (top) ─
  const maxRes=Math.max(1,...buckets.map(b=>b.resources));
  const gW=W-padR.left-padR.right;
  const gHR=splitY-padR.top-padR.bottom;
  for(let i=0;i<=4;i++){
    const val=Math.round(maxRes/4*i),y=padR.top+gHR-(i/4)*gHR;
    ctx.strokeStyle=ct.gridH;ctx.lineWidth=0.7;ctx.beginPath();ctx.moveTo(padR.left,y);ctx.lineTo(W-padR.right,y);ctx.stroke();
    ctx.fillStyle=ct.axisText;ctx.font='10px Arial';ctx.textAlign='right';ctx.textBaseline='middle';ctx.fillText(val,padR.left-5,y);
  }
  ctx.save();ctx.translate(14,padR.top+gHR/2);ctx.rotate(-Math.PI/2);ctx.fillStyle=ct.axisText;ctx.font='11px Arial';ctx.textAlign='center';ctx.fillText('Workers×Days',0,0);ctx.restore();
  const bSlot=gW/buckets.length,barW=Math.max(3,bSlot*0.72);
  buckets.forEach((b,i)=>{
    const x=padR.left+i*bSlot+bSlot*0.14,bH=(b.resources/maxRes)*gHR,y=padR.top+gHR-bH;
    ctx.globalAlpha=0.8;ctx.fillStyle='#3498db';ctx.fillRect(x,y,barW,bH);
    ctx.globalAlpha=1;ctx.strokeStyle='#1a6fa0';ctx.lineWidth=0.7;ctx.strokeRect(x,y,barW,bH);
    ctx.fillStyle=ct.axisText;ctx.font='9px Arial';ctx.textAlign='center';ctx.textBaseline='top';
    ctx.fillText(b.label,x+barW/2,splitY-10);
    _storeBarHit(x,y,barW,bH,{type:'resource',label:b.label,resources:b.resources});
  });
  ctx.fillStyle=ct.axisText;ctx.font='bold 11px Arial';ctx.textAlign='right';ctx.textBaseline='top';
  ctx.fillText('Total: '+Math.round(buckets.reduce((s,b)=>s+b.resources,0)).toLocaleString()+' worker-days',W-padR.right,padR.top);

  // ─ Divider ─
  ctx.strokeStyle=ct.border||'#ccc';ctx.lineWidth=1;ctx.setLineDash([4,4]);
  ctx.beginPath();ctx.moveTo(padR.left,splitY);ctx.lineTo(W-padR.right,splitY);ctx.stroke();ctx.setLineDash([]);

  // ─ S-curve (bottom) ─
  let cum=0;
  const pts=buckets.map((b,i)=>{cum+=b.cost;return{x:i,y:cum,label:b.label,monthly:b.cost};});
  const maxCost=Math.max(1,pts[pts.length-1]?.y||1);
  const gHC=H-splitY-padC.top-padC.bottom,yBase2=splitY+padC.top;
  for(let i=0;i<=4;i++){
    const val=maxCost/4*i,y=yBase2+gHC-(i/4)*gHC;
    ctx.strokeStyle=ct.gridH;ctx.lineWidth=0.7;ctx.beginPath();ctx.moveTo(padC.left,y);ctx.lineTo(W-padC.right,y);ctx.stroke();
    ctx.fillStyle=ct.axisText;ctx.font='10px Arial';ctx.textAlign='right';ctx.textBaseline='middle';ctx.fillText(_fmtC(val),padC.left-5,y);
  }
  ctx.save();ctx.translate(14,yBase2+gHC/2);ctx.rotate(-Math.PI/2);ctx.fillStyle=ct.axisText;ctx.font='11px Arial';ctx.textAlign='center';ctx.fillText('Cumulative Cost',0,0);ctx.restore();
  const xOf=i=>padC.left+(i/(pts.length-1||1))*gW;
  const yOf=y=>yBase2+gHC-(y/maxCost)*gHC;
  ctx.beginPath();ctx.moveTo(xOf(0),yBase2+gHC);
  pts.forEach(p=>ctx.lineTo(xOf(p.x),yOf(p.y)));
  ctx.lineTo(xOf(pts.length-1),yBase2+gHC);ctx.closePath();
  ctx.fillStyle='#27ae60';ctx.globalAlpha=0.15;ctx.fill();ctx.globalAlpha=1;
  ctx.beginPath();pts.forEach((p,i)=>i===0?ctx.moveTo(xOf(p.x),yOf(p.y)):ctx.lineTo(xOf(p.x),yOf(p.y)));
  ctx.strokeStyle='#27ae60';ctx.lineWidth=2.5;ctx.stroke();
  pts.forEach(p=>{
    const px=xOf(p.x),py=yOf(p.y);
    ctx.beginPath();ctx.arc(px,py,3,0,Math.PI*2);ctx.fillStyle='#27ae60';ctx.fill();
    _storePointHit(px,py,{type:'point',label:p.label,cumulative:p.y,monthly:p.monthly});
  });
  const last=pts[pts.length-1];
  if(last){ctx.fillStyle='#27ae60';ctx.font='bold 11px Arial';ctx.textAlign='right';ctx.textBaseline='bottom';ctx.fillText('Total: '+_fmtC(last.y),W-padC.right,H-padC.bottom+2);}
}
  ctx.save();ctx.translate(12,padR.top+gHR/2);ctx.rotate(-Math.PI/2);ctx.fillStyle=ct.axisText;ctx.font='11px Arial';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText('Workers×Days',0,0);ctx.restore();
  const bSlot=gW/buckets.length;

// ═══════════════════════════════════════════════════════════════════
//  DASHBOARD
// ═══════════════════════════════════════════════════════════════════
function renderDashboard(){
  if(!document.getElementById('dashboard')?.classList.contains('active')) return;

  const now = new Date();
  const dataDate = titleBlock.dataDate ? parseDateLocal(titleBlock.dataDate) : now;
  const fmtD = d => d ? formatLongDate(d) : '—';
  const fmtC = v => { const s=projectCurrency||'£'; return v>=1e6?s+(v/1e6).toFixed(2)+'M':v>=1e3?s+(v/1e3).toFixed(0)+'k':s+Math.round(v); };

  // Header
  document.getElementById('dashProjectName').textContent = titleBlock.projectName || 'Programme Dashboard';
  document.getElementById('dashSubtitle').textContent =
    `${titleBlock.drawingTitle||''} · Data Date: ${fmtD(dataDate)} · Generated: ${fmtD(now)}`;

  // ── Scorecard data ────────────────────────────────────────────────
  const totalActs = activities.length;
  const completedActs = activities.filter(a=>a.end&&a.end<dataDate).length;
  const inProgressActs = activities.filter(a=>a.start&&a.start<=dataDate&&a.end&&a.end>=dataDate).length;
  const notStartedActs = activities.filter(a=>a.start&&a.start>dataDate).length;

  // Days remaining to last activity end
  const lastEnd = activities.reduce((mx,a)=>a.end&&a.end>mx?a.end:mx, new Date(0));
  const daysRemaining = lastEnd>now ? Math.round((lastEnd-dataDate)/86400000) : 0;

  // Total cost
  let totalLabour=0, totalMaterial=0, totalEquip=0;
  activities.forEach(a=>{
    const dur=a.start&&a.end?Math.max(0,(a.end-a.start)/86400000):0;
    totalLabour  += (parseFloat(a.resourceCount)||0)*(parseFloat(a.unitCost)||0)*dur;
    totalEquip   += (parseFloat(a.equipmentCost)||0)*dur;
    totalMaterial+= parseFloat(a.materialCost)||0;
  });
  const totalCost = totalLabour+totalMaterial+totalEquip;
  const totalWorkerDays = activities.reduce((s,a)=>{
    const dur=a.start&&a.end?Math.max(0,(a.end-a.start)/86400000):0;
    return s+(parseFloat(a.resourceCount)||0)*dur;
  },0);

  // ── Render scorecards ─────────────────────────────────────────────
  const cards=[
    {icon:'📋',label:'Total Activities',value:totalActs,sub:'in programme',color:'#2c3e50'},
    {icon:'✅',label:'Completed',value:completedActs,sub:`${totalActs?Math.round(completedActs/totalActs*100):0}% of activities`,color:'#27ae60'},
    {icon:'⚙️',label:'In Progress',value:inProgressActs,sub:'at data date',color:'#e67e22'},
    {icon:'📅',label:'Days Remaining',value:daysRemaining,sub:`to ${fmtD(lastEnd)}`,color:'#2980b9'},
    {icon:'💰',label:'Total Project Cost',value:fmtC(totalCost),sub:`${fmtC(totalLabour)} labour + ${fmtC(totalMaterial)} materials + ${fmtC(totalEquip)} equipment`,color:'#8e44ad'},
    {icon:'👷',label:'Total Worker-Days',value:Math.round(totalWorkerDays).toLocaleString(),sub:'across all activities',color:'#16a085'},
  ];
  document.getElementById('dashCards').innerHTML=cards.map(c=>`
    <div class="dash-card" style="border-top:4px solid ${c.color}">
      <div class="dash-card-icon">${c.icon}</div>
      <div class="dash-card-value" style="color:${c.color}">${c.value}</div>
      <div class="dash-card-label">${c.label}</div>
      <div class="dash-card-sub">${c.sub}</div>
    </div>`).join('');

  // ── Health indicators ─────────────────────────────────────────────
  const ddAge = dataDate ? Math.round((now-dataDate)/86400000) : 999;
  const hasLinks = activityLinks.length>0;
  const linkedPct = totalActs>0 ? [...new Set(activityLinks.flatMap(l=>[l.fromIdx,l.toIdx]))].length/totalActs*100 : 0;
  const hasCostData = activities.some(a=>a.unitCost||a.materialCost||a.equipmentCost);
  const hasP6Ids = activities.filter(a=>a.p6Id).length;
  const overrunning = activities.filter(a=>a.end&&timelineEnd&&a.end>timelineEnd).length;
  const hasCPaths = cpPaths.length>0;

  function healthDot(ok,warn,label,detail){
    const col=ok?'#27ae60':warn?'#f39c12':'#e74c3c';
    const icon=ok?'●':warn?'◐':'○';
    return `<div class="dash-health-item">
      <span class="dash-health-dot" style="color:${col}">${icon}</span>
      <div><div class="dash-health-label">${label}</div><div class="dash-health-detail">${detail}</div></div>
    </div>`;
  }

  document.getElementById('dashHealth').innerHTML=`
    <div class="dash-health-title">Programme Health</div>
    ${healthDot(ddAge<=14,ddAge<=30,'Data Date',ddAge<=14?'Up to date':'Last updated '+ddAge+' days ago')}
    ${healthDot(hasLinks&&linkedPct>50,hasLinks,'Dependencies',hasLinks?`${activityLinks.length} links (${Math.round(linkedPct)}% of activities)`:'No dependency links defined')}
    ${healthDot(hasP6Ids===totalActs&&totalActs>0,hasP6Ids>0,'P6 IDs',hasP6Ids+' of '+totalActs+' activities have P6 IDs')}
    ${healthDot(hasCostData,'','Cost Data',hasCostData?'Cost data present':'No cost data entered')}
    ${healthDot(overrunning===0,false,'Timeline',overrunning===0?'All activities within timeline':overrunning+' activities overrun timeline end')}
    ${healthDot(hasCPaths,false,'Critical Path',hasCPaths?cpPaths.length+' path'+(cpPaths.length>1?'s':'')+' defined':'No critical paths defined')}
  `;

  // ── Analytics charts ─────────────────────────────────────────────
  requestAnimationFrame(()=>_drawDashAnalytics());

  // ── Upcoming milestones ────────────────────────────────────────────
  const milestoneShapes=new Set(['diamond','star','flag','circle','triangle']);
  const upcoming=activities
    .filter(a=>milestoneShapes.has(a.shape)&&a.start&&a.start>=dataDate)
    .sort((a,b)=>a.start-b.start)
    .slice(0,8);

  const msEl=document.getElementById('dashMilestones');
  if(!upcoming.length){
    msEl.innerHTML='<p class="dash-empty">No upcoming milestones. Add diamond, star, or flag activities to see them here.</p>';
  } else {
    msEl.innerHTML=`<table class="dash-milestone-table">
      <thead><tr><th>Milestone</th><th>Date</th><th>Ch</th><th>In</th></tr></thead>
      <tbody>${upcoming.map(a=>{
        const days=Math.round((a.start-dataDate)/86400000);
        const urgency=days<7?'#e74c3c':days<30?'#f39c12':'#27ae60';
        const shapes={diamond:'♦',star:'★',flag:'⚑',circle:'●',triangle:'▲'};
        return`<tr>
          <td><span style="color:${a.color||'#3498db'};margin-right:5px">${shapes[a.shape]||'●'}</span>${a.name}</td>
          <td style="white-space:nowrap">${fmtD(a.start)}</td>
          <td>${a.startCh}</td>
          <td style="color:${urgency};font-weight:bold">${days}d</td>
        </tr>`;
      }).join('')}</tbody>
    </table>`;
  }

  // ── Zone breakdown ────────────────────────────────────────────────
  const zoneEl=document.getElementById('dashZones');
  if(!chainageZones.length){
    zoneEl.innerHTML='<p class="dash-empty">No chainage zones defined.</p>';
  } else {
    const zoneData=chainageZones.map(z=>{
      const zActs=activities.filter(a=>{
        const midCh=(a.startCh+a.endCh)/2;
        return midCh>=z.startCh&&midCh<=z.endCh;
      });
      const done=zActs.filter(a=>a.end&&a.end<dataDate).length;
      const pct=zActs.length?Math.round(done/zActs.length*100):0;
      return{name:z.name,total:zActs.length,done,pct,color:z.color||'#3498db'};
    }).filter(z=>z.total>0);
    if(!zoneData.length){
      zoneEl.innerHTML='<p class="dash-empty">No activities assigned to zones.</p>';
    } else {
      zoneEl.innerHTML=zoneData.map(z=>`
        <div class="dash-zone-row">
          <div style="display:flex;justify-content:space-between;margin-bottom:3px">
            <span style="font-size:13px;font-weight:500">${z.name}</span>
            <span style="font-size:12px;color:var(--text-muted)">${z.done}/${z.total} · ${z.pct}%</span>
          </div>
          <div class="dash-zone-bar-bg">
            <div class="dash-zone-bar-fill" style="width:${z.pct}%;background:${z.color}"></div>
          </div>
        </div>`).join('');
    }
  }

  // ── Critical path summary ─────────────────────────────────────────
  const cpEl=document.getElementById('dashCritical');
  if(!cpPaths.length){
    cpEl.innerHTML='<p class="dash-empty">No critical paths defined. Use the Critical Path tab to trace paths to your key dates.</p>';
  } else {
    cpEl.innerHTML=cpPaths.map((p,i)=>{
      const target=activities[p.targetIdx];
      return`<div class="dash-cp-row">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
          <div style="width:12px;height:12px;border-radius:50%;background:${p.color};flex-shrink:0"></div>
          <strong>Path ${i+1}:</strong> ${target?target.name:'Unknown target'}
        </div>
        <div style="font-size:12px;color:var(--text-muted);padding-left:20px">
          ${p.chainInOrder.length} activities · ~${Math.round(p.totalDays)} days driving duration
          ${target?.end?'<br>Target date: '+fmtD(target.end):''}
        </div>
      </div>`;
    }).join('');
  }

  // ── Risk panels ───────────────────────────────────────────────────
  _renderDashRisks();
}

function _drawDashAnalytics(){
  const ct = getChartTheme();
  const dataDate = titleBlock.dataDate ? parseDateLocal(titleBlock.dataDate) : new Date();

  // ─────────────────────────────────────────────────────────────────
  // Helper: draw a large, readable donut with external labels
  // ─────────────────────────────────────────────────────────────────
  function drawDonut(canvasEl, legendEl, slices, centreLines){
    if(!canvasEl) return;
    // Read parent width BEFORE setting canvas dimensions to avoid feedback loop
    const W = canvasEl.parentElement.offsetWidth || 360;
    const H = Math.round(W * 0.72);
    canvasEl.width  = W;
    canvasEl.height = H;
    const c = canvasEl.getContext('2d');
    c.fillStyle = ct.canvasBg; c.fillRect(0,0,W,H);

    const total = slices.reduce((s,x)=>s+x.value,0);
    if(!total || !slices.length){
      c.fillStyle = ct.axisText; c.font = '15px Arial';
      c.textAlign = 'center'; c.textBaseline = 'middle';
      c.fillText('No data', W/2, H/2); return;
    }

    const cx = W/2, cy = H/2;
    const R  = Math.min(W, H) * 0.38;
    const ri = R * 0.54;

    // Draw slices
    let angle = -Math.PI / 2;
    slices.forEach(s => {
      const sweep = (s.value / total) * Math.PI * 2;
      // Outer slice
      c.beginPath(); c.moveTo(cx,cy);
      c.arc(cx,cy,R,angle,angle+sweep); c.closePath();
      c.fillStyle = s.color; c.globalAlpha = 0.92; c.fill(); c.globalAlpha = 1;
      // Subtle border
      c.strokeStyle = ct.canvasBg; c.lineWidth = 2;
      c.beginPath(); c.moveTo(cx,cy);
      c.arc(cx,cy,R,angle,angle+sweep); c.closePath(); c.stroke();
      angle += sweep;
    });

    // Donut hole
    c.beginPath(); c.arc(cx,cy,ri,0,Math.PI*2);
    c.fillStyle = ct.canvasBg; c.fill();

    // Centre text
    c.textAlign = 'center'; c.textBaseline = 'middle';
    centreLines.forEach((line, i) => {
      const offset = (i - (centreLines.length-1)/2) * (line.size + 3);
      c.fillStyle = line.color || ct.axisText;
      c.font = `${line.bold?'bold ':''}${line.size}px Arial`;
      c.fillText(line.text, cx, cy + offset);
    });

    // Percentage labels on slices (only if slice > 6%)
    angle = -Math.PI / 2;
    slices.forEach(s => {
      const sweep = (s.value / total) * Math.PI * 2;
      const pct = Math.round(s.value / total * 100);
      if(pct >= 6){
        const mid = angle + sweep/2;
        const lr  = R * 0.74;
        c.fillStyle = '#fff';
        c.font = `bold ${Math.max(11, Math.round(R*0.13))}px Arial`;
        c.textAlign = 'center'; c.textBaseline = 'middle';
        c.fillText(pct + '%', cx + Math.cos(mid)*lr, cy + Math.sin(mid)*lr);
      }
      angle += sweep;
    });

    // Legend below (HTML)
    if(legendEl){
      legendEl.innerHTML = slices.map(s => {
        const pct = Math.round(s.value / total * 100);
        return `<div class="dash-legend-row">
          <div class="dash-legend-dot" style="background:${s.color}"></div>
          <span class="dash-legend-label">${s.label}</span>
          <strong class="dash-legend-val" style="color:${s.color}">${s.formatted}</strong>
          <span class="dash-legend-pct">${pct}%</span>
        </div>`;
      }).join('');
    }
  }

  // ─────────────────────────────────────────────────────────────────
  // 1. COST BREAKDOWN DONUT
  // ─────────────────────────────────────────────────────────────────
  {
    let totalLabour=0, totalMaterial=0, totalEquip=0;
    activities.forEach(a => {
      const dur = a.start&&a.end ? Math.max(0,(a.end-a.start)/86400000) : 0;
      totalLabour   += (parseFloat(a.resourceCount)||0) * (parseFloat(a.unitCost)||0) * dur;
      totalEquip    += (parseFloat(a.equipmentCost)||0) * dur;
      totalMaterial += parseFloat(a.materialCost)||0;
    });
    const totalCost = totalLabour + totalMaterial + totalEquip;
    const slices = [
      {label:'💼 Labour',   value:totalLabour,   color:'#3498db', formatted:_fmtC(totalLabour)},
      {label:'🧱 Material', value:totalMaterial,  color:'#e67e22', formatted:_fmtC(totalMaterial)},
      {label:'🚜 Equipment',value:totalEquip,     color:'#8e44ad', formatted:_fmtC(totalEquip)},
    ].filter(s => s.value > 0);
    drawDonut(
      document.getElementById('dashCostPie'),
      document.getElementById('dashCostLegend'),
      slices,
      [
        {text:'Total', size:13, bold:false},
        {text:_fmtC(totalCost), size:17, bold:true, color:ct.axisText},
      ]
    );
  }

  // ─────────────────────────────────────────────────────────────────
  // 2. ZONE / TYPE DONUT
  // ─────────────────────────────────────────────────────────────────
  {
    const COLOURS = ['#3498db','#27ae60','#e74c3c','#f39c12','#8e44ad','#16a085','#e67e22','#2c3e50','#1abc9c','#d35400'];
    let slices = [];
    const subtitle = chainageZones.length ? 'by zone' : 'by shape';

    if(chainageZones.length){
      slices = chainageZones.map((z,i) => {
        const cnt = activities.filter(a => {
          const mid = (a.startCh + a.endCh) / 2;
          return mid >= z.startCh && mid <= z.endCh;
        }).length;
        return {label: z.name, value: cnt, color: z.color || COLOURS[i % COLOURS.length], formatted: cnt + ' acts'};
      }).filter(s => s.value > 0);
    } else {
      const SHAPE_C = {rect:'#3498db',line:'#95a5a6',circle:'#f39c12',diamond:'#e74c3c',triangle:'#27ae60',star:'#8e44ad',flag:'#e67e22'};
      const counts = {};
      activities.forEach(a => { const sh = a.shape||'rect'; counts[sh] = (counts[sh]||0)+1; });
      slices = Object.entries(counts)
        .sort((a,b) => b[1]-a[1])
        .map(([k,v]) => ({label: k.charAt(0).toUpperCase()+k.slice(1), value: v, color: SHAPE_C[k]||'#aaa', formatted: v + ' acts'}));
    }

    drawDonut(
      document.getElementById('dashShapePie'),
      document.getElementById('dashShapeLegend'),
      slices,
      [
        {text: activities.length + '', size:18, bold:true, color:ct.axisText},
        {text:'activities', size:12},
        {text:subtitle, size:11, color: '#888'},
      ]
    );
  }

  // ─────────────────────────────────────────────────────────────────
  // 3. ACTIVITY LOAD — 6 months before + 6 months after data date
  // ─────────────────────────────────────────────────────────────────
  {
    const tlCanvas = document.getElementById('dashTimelineBars');
    if(!tlCanvas || !timelineStart || !timelineEnd) return;
    const W = tlCanvas.parentElement.offsetWidth || 500;
    const H = 260;
    tlCanvas.width = W; tlCanvas.height = H;
    const tc = tlCanvas.getContext('2d');
    tc.fillStyle = ct.canvasBg; tc.fillRect(0,0,W,H);

    // Show 6 months before data date and 6 months after — centred on "now"
    const MONTHS_BEFORE = 6, MONTHS_AFTER = 6;
    const TOTAL = MONTHS_BEFORE + MONTHS_AFTER;
    const months = [];
    for(let i = -MONTHS_BEFORE; i < MONTHS_AFTER; i++){
      const ms = new Date(dataDate.getFullYear(), dataDate.getMonth()+i, 1);
      const me = new Date(dataDate.getFullYear(), dataDate.getMonth()+i+1, 0, 23,59,59);
      const active = activities.filter(a => a.start&&a.end&&a.start<=me&&a.end>=ms).length;
      const lbl = ms.toLocaleDateString(undefined, {month:'short', year:'2-digit'});
      months.push({label:lbl, active, isPast: me < dataDate, isCurrentMonth: i === 0});
    }

    const maxCnt = Math.max(1, ...months.map(m => m.active));
    const pad = {top:36, right:16, bottom:44, left:44};
    const gW = W - pad.left - pad.right;
    const gH = H - pad.top - pad.bottom;
    const bSlot = gW / months.length;
    const bW = Math.max(8, bSlot * 0.68);

    // Gridlines + Y axis
    const ySteps = 5;
    for(let i = 0; i <= ySteps; i++){
      const val = Math.ceil(maxCnt / ySteps * i);
      const y = pad.top + gH - (i/ySteps)*gH;
      tc.strokeStyle = ct.gridH; tc.lineWidth = 0.7;
      tc.beginPath(); tc.moveTo(pad.left, y); tc.lineTo(W - pad.right, y); tc.stroke();
      tc.fillStyle = ct.axisText; tc.font = '11px Arial';
      tc.textAlign = 'right'; tc.textBaseline = 'middle';
      tc.fillText(val, pad.left - 5, y);
    }

    // Y axis label
    tc.save(); tc.translate(14, pad.top + gH/2); tc.rotate(-Math.PI/2);
    tc.fillStyle = ct.axisText; tc.font = 'bold 11px Arial';
    tc.textAlign = 'center'; tc.textBaseline = 'middle';
    tc.fillText('Active activities', 0, 0); tc.restore();

    // Title
    tc.fillStyle = ct.axisText; tc.font = 'bold 12px Arial';
    tc.textAlign = 'center'; tc.textBaseline = 'top';
    tc.fillText(`Activity Load — ${MONTHS_BEFORE} months before & after data date`, W/2, 8);

    // Bars
    months.forEach((m, i) => {
      const x = pad.left + i * bSlot + (bSlot - bW) / 2;
      const bH = (m.active / maxCnt) * gH;
      const y = pad.top + gH - bH;
      const col = m.isPast ? '#95a5a6' : '#3498db';
      tc.globalAlpha = m.isPast ? 0.5 : 0.88;
      tc.fillStyle = col; tc.fillRect(x, y, bW, bH);
      tc.globalAlpha = 1;
      // "Now" vertical marker
      if(m.isCurrentMonth){
        tc.save();
        tc.strokeStyle = '#e74c3c'; tc.lineWidth = 2.5; tc.setLineDash([5,3]);
        tc.beginPath(); tc.moveTo(x - 2, pad.top); tc.lineTo(x - 2, pad.top + gH); tc.stroke();
        tc.setLineDash([]); tc.restore();
        tc.fillStyle = '#e74c3c'; tc.font = 'bold 10px Arial';
        tc.textAlign = 'left'; tc.textBaseline = 'top';
        tc.fillText('▶ Now', x, pad.top + 2);
      }
      if(bH > 22){
        tc.fillStyle = '#fff'; tc.font = `bold ${Math.max(11, Math.round(bW*0.32))}px Arial`;
        tc.textAlign = 'center'; tc.textBaseline = 'middle';
        tc.fillText(m.active, x + bW/2, y + bH/2);
      } else if(m.active > 0){
        tc.fillStyle = ct.axisText; tc.font = '10px Arial';
        tc.textAlign = 'center'; tc.textBaseline = 'bottom';
        tc.fillText(m.active, x + bW/2, y - 2);
      }
      tc.fillStyle = ct.axisText; tc.font = '10px Arial';
      tc.textAlign = 'center'; tc.textBaseline = 'top';
      tc.fillText(m.label, x + bW/2, pad.top + gH + 5);
    });

    // Legend
    const legendItems = [['Past','#95a5a6'],['Upcoming','#3498db'],['▶ Data date','#e74c3c']];
    let lx = pad.left;
    legendItems.forEach(([label, col]) => {
      if(col === '#e74c3c'){
        tc.strokeStyle = col; tc.lineWidth = 2; tc.setLineDash([4,2]);
        tc.beginPath(); tc.moveTo(lx, H-10); tc.lineTo(lx+14, H-10); tc.stroke();
        tc.setLineDash([]);
      } else {
        tc.fillStyle = col; tc.globalAlpha = col==='#95a5a6'?0.5:0.88;
        tc.fillRect(lx, H-16, 12, 10); tc.globalAlpha = 1;
      }
      tc.fillStyle = ct.axisText; tc.font = '10px Arial';
      tc.textAlign = 'left'; tc.textBaseline = 'middle';
      tc.fillText(label, lx + 17, H - 11);
      lx += tc.measureText(label).width + 36;
    });
  }
}

function printDashboard(){
  const projectName = titleBlock.projectName||'Programme Dashboard';
  const dataDate    = titleBlock.dataDate ? parseDateLocal(titleBlock.dataDate) : new Date();
  const printDate   = new Date().toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  const win = window.open('','_blank');
  if(!win){alert('Pop-up blocked — please allow pop-ups for this page.');return;}

  // Capture all analytics canvases as data URLs
  const costPie   = document.getElementById('dashCostPie');
  const shapePie  = document.getElementById('dashShapePie');
  const tlBars    = document.getElementById('dashTimelineBars');
  const costUrl   = costPie  ? costPie.toDataURL('image/png')  : '';
  const shapeUrl  = shapePie ? shapePie.toDataURL('image/png') : '';
  const tlUrl     = tlBars   ? tlBars.toDataURL('image/png')   : '';

  // Capture HTML legend contents (text rows below each donut)
  const costLegendHTML  = document.getElementById('dashCostLegend')?.innerHTML  || '';
  const shapeLegendHTML = document.getElementById('dashShapeLegend')?.innerHTML || '';

  // Capture health row
  const healthHTML = document.getElementById('dashHealth')?.innerHTML || '';

  // Capture bottom panels
  const milestonesHTML = document.getElementById('dashMilestones')?.innerHTML || '';
  const zonesHTML      = document.getElementById('dashZones')?.innerHTML      || '';
  const cpHTML         = document.getElementById('dashCritical')?.innerHTML   || '';

  // Capture risk panels
  const riskMatrixCvs  = document.getElementById('dashRiskMatrix');
  const riskMatrixUrl  = riskMatrixCvs ? riskMatrixCvs.toDataURL('image/png') : '';
  const riskListHTML   = document.getElementById('dashRisks')?.innerHTML      || '';
  const riskExpHTML    = document.getElementById('dashRiskExposure')?.innerHTML || '';

  // Build scorecard HTML from DOM
  let cardsHTML = '';
  document.querySelectorAll('.dash-card').forEach(c=>{
    const col = c.querySelector('.dash-card-value')?.style.color || '#2980b9';
    cardsHTML += `<div class="card" style="border-top:4px solid ${col}">
      <div class="card-icon">${c.querySelector('.dash-card-icon')?.textContent||''}</div>
      <div class="card-val" style="color:${col}">${c.querySelector('.dash-card-value')?.textContent||''}</div>
      <div class="card-lbl">${c.querySelector('.dash-card-label')?.textContent||''}</div>
      <div class="card-sub">${c.querySelector('.dash-card-sub')?.textContent||''}</div>
    </div>`;
  });

  win.document.write(`<!DOCTYPE html><html><head>
    <title>${projectName} — Dashboard</title>
    <style>
      *{margin:0;padding:0;box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      body{font-family:'Segoe UI',Arial,sans-serif;background:#fff;color:#1a1a2e;padding:10mm;font-size:12px}
      h1{font-size:20px;font-weight:700;margin-bottom:3px}
      .sub{font-size:11px;color:#666;padding-bottom:8px;margin-bottom:12px;border-bottom:2px solid #2c3e50}
      /* Scorecards */
      .cards{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;margin-bottom:12px}
      .card{background:#f8f9fa;border-radius:6px;padding:10px 8px;text-align:center}
      .card-icon{font-size:18px;margin-bottom:4px}
      .card-val{font-size:18px;font-weight:700;line-height:1.1}
      .card-lbl{font-size:10px;color:#555;margin-top:3px;font-weight:600;text-transform:uppercase;letter-spacing:.03em}
      .card-sub{font-size:9px;color:#888;margin-top:2px}
      /* Health */
      .health{border:1px solid #ddd;border-radius:6px;padding:10px;display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px;background:#f8f9fa}
      .health .dash-health-title{width:100%;font-size:10px;font-weight:700;text-transform:uppercase;color:#888;margin-bottom:4px}
      .health .dash-health-item{display:flex;gap:6px;align-items:flex-start;flex:1;min-width:140px}
      .health .dash-health-dot{font-size:16px;line-height:1}
      .health .dash-health-label{font-size:11px;font-weight:600}
      .health .dash-health-detail{font-size:10px;color:#666;margin-top:1px}
      /* Analytics row */
      .analytics{display:grid;grid-template-columns:1fr 1fr 1.4fr;gap:10px;margin-bottom:12px}
      .panel{border:1px solid #ddd;border-radius:6px;padding:10px;background:#fff}
      .panel-title{font-size:10px;font-weight:700;text-transform:uppercase;color:#888;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid #eee;letter-spacing:.04em}
      .panel img{width:100%;height:auto;display:block;margin-bottom:6px}
      /* Legend rows */
      .dash-legend-row{display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:11px}
      .dash-legend-row:last-child{border-bottom:none}
      .dash-legend-dot{width:11px;height:11px;border-radius:50%;flex-shrink:0}
      .dash-legend-label{flex:1;color:#333}
      .dash-legend-val{font-weight:700;font-size:11px}
      .dash-legend-pct{color:#888;font-size:10px;min-width:30px;text-align:right}
      /* Bottom row */
      .bottom{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
      table{width:100%;border-collapse:collapse;font-size:10px}
      th{background:#f0f4f8;padding:3px 5px;text-align:left;font-size:9px;color:#555;font-weight:600}
      td{padding:4px 5px;border-bottom:1px solid #f0f0f0;vertical-align:middle}
      /* Zone bars */
      .dash-zone-row{margin-bottom:8px}
      .dash-zone-bar-bg{height:7px;background:#e8e8e8;border-radius:4px;overflow:hidden;margin-top:3px}
      .dash-zone-bar-fill{height:100%;border-radius:4px}
      /* Risk section */
      .risk-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:10px}
      .risk-exp-row{display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:11px}
      .risk-exp-badge{border:1.5px solid;border-radius:5px;padding:3px 7px;font-size:10px;font-weight:600;text-align:center}
      .dash-legend-row{display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:11px}
      @media print{body{padding:8mm}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    </style>
  </head><body>
    <h1>${projectName}</h1>
    <div class="sub">${titleBlock.drawingTitle||''} · Data Date: ${formatLongDate(dataDate)} · Printed: ${printDate}</div>
    <div class="cards">${cardsHTML}</div>
    <div class="health">${healthHTML}</div>
    <div class="analytics">
      <div class="panel">
        <div class="panel-title">💰 Cost Breakdown</div>
        ${costUrl?`<img src="${costUrl}">`:''} ${costLegendHTML}
      </div>
      <div class="panel">
        <div class="panel-title">🏗 Activities by Zone / Type</div>
        ${shapeUrl?`<img src="${shapeUrl}">`:''} ${shapeLegendHTML}
      </div>
      <div class="panel">
        <div class="panel-title">📅 Activity Load</div>
        ${tlUrl?`<img src="${tlUrl}"`+' style="margin-bottom:0">':''}
      </div>
    </div>
    <div class="bottom">
      <div class="panel"><div class="panel-title">🚩 Upcoming Milestones</div>${milestonesHTML}</div>
      <div class="panel"><div class="panel-title">🏗 Zone Breakdown</div>${zonesHTML}</div>
      <div class="panel"><div class="panel-title">🔴 Critical Path Summary</div>${cpHTML}</div>
    </div>
    ${risks.length?`
    <div class="risk-row" style="margin-top:12px">
      <div class="panel" style="grid-column:1/2">
        <div class="panel-title">⚠️ Top Open Risks</div>
        ${riskListHTML}
      </div>
      <div class="panel" style="grid-column:2/3;text-align:center">
        <div class="panel-title">📊 Risk Heat Map</div>
        ${riskMatrixUrl?`<img src="${riskMatrixUrl}" style="max-width:200px;height:auto;display:block;margin:0 auto">`:  ''}
      </div>
      <div class="panel" style="grid-column:3/4">
        <div class="panel-title">💡 Risk Exposure</div>
        ${riskExpHTML}
      </div>
    </div>`:''}
  </body></html>`);
  win.document.close();
  win.onload = ()=>{ win.focus(); win.print(); };
}
// ═══════════════════════════════════════════════════════════════════
//  LEGEND MODAL
// ═══════════════════════════════════════════════════════════════════
function openLegendModal(){
  renderLegendModal();
  document.getElementById('legendModal').style.display='flex';
}

function renderLegendModal(){
  const el=document.getElementById('legendModalContent');
  if(!el)return;
  const mode=document.querySelector('input[name="legendMode"]:checked')?.value||'groups';
  if(mode==='groups') _renderLegendGroups(el);
  else _renderLegendActivities(el);
}

function _legendShapeCanvas(shape,color){
  const w=36,h=22;
  const cvs=document.createElement('canvas');
  cvs.width=w;cvs.height=h;
  cvs.style.cssText=`width:${w}px;height:${h}px;vertical-align:middle;flex-shrink:0`;
  const c=cvs.getContext('2d');
  const cx=w/2,cy=h/2,r=7;
  c.fillStyle=color;c.strokeStyle=color;c.globalAlpha=0.9;
  switch(shape){
    case 'rect':   c.fillRect(2,cy-4,w-4,8);break;
    case 'line':   c.lineWidth=2.5;c.beginPath();c.moveTo(2,cy);c.lineTo(w-2,cy);c.stroke();break;
    case 'circle': c.beginPath();c.arc(cx,cy,r,0,Math.PI*2);c.fill();break;
    case 'diamond':c.beginPath();c.moveTo(cx,2);c.lineTo(w-2,cy);c.lineTo(cx,h-2);c.lineTo(2,cy);c.closePath();c.fill();break;
    case 'triangle':c.beginPath();c.moveTo(cx,2);c.lineTo(w-2,h-2);c.lineTo(2,h-2);c.closePath();c.fill();break;
    case 'star':{
      const pts=5,ri=r*.42;c.beginPath();
      for(let i=0;i<pts*2;i++){const a=((i*Math.PI)/pts)-Math.PI/2;const rr=i%2===0?r:ri;i===0?c.moveTo(cx+rr*Math.cos(a),cy+rr*Math.sin(a)):c.lineTo(cx+rr*Math.cos(a),cy+rr*Math.sin(a));}
      c.closePath();c.fill();break;
    }
    case 'flag':
      c.lineWidth=1.5;c.beginPath();c.moveTo(cx-2,2);c.lineTo(cx-2,h-2);c.stroke();
      c.beginPath();c.moveTo(cx-2,2);c.lineTo(w-2,cy-3);c.lineTo(cx-2,cy+2);c.closePath();c.fill();break;
    default: c.fillRect(2,cy-4,w-4,8);
  }
  c.globalAlpha=1;
  return cvs;
}

function _renderLegendGroups(el){
  el.innerHTML='';
  const categories=groups.filter(g=>g.parentId===null);
  if(!categories.length){
    el.innerHTML='<p style="color:var(--text-muted);font-size:13px;padding:16px">No groups defined yet. Create categories and sub-groups in the 📂 Grouping tab.</p>';
    return;
  }
  categories.forEach(cat=>{
    const subs=groups.filter(g=>g.parentId===cat.id);
    const catRow=document.createElement('div');
    catRow.style.cssText='display:flex;align-items:center;gap:10px;padding:8px 6px;margin-bottom:4px;border-bottom:2px solid var(--border);';
    const dot=document.createElement('div');
    dot.style.cssText=`width:14px;height:14px;border-radius:50%;background:${cat.color||'#3498db'};flex-shrink:0`;
    const lbl=document.createElement('strong');
    lbl.style.cssText='font-size:14px;color:var(--text)';
    lbl.textContent=cat.name;
    catRow.appendChild(dot);catRow.appendChild(lbl);
    el.appendChild(catRow);
    if(!subs.length){
      const ns=document.createElement('div');
      ns.style.cssText='padding:4px 16px;font-size:12px;color:var(--text-muted);margin-bottom:8px';
      ns.textContent='No sub-groups';el.appendChild(ns);return;
    }
    subs.forEach(sg=>{
      const sample=activities.find(a=>a.groups&&a.groups.includes(sg.id));
      const shape=sample?sample.shape||'rect':'rect';
      const color=sg.color||cat.color||'#3498db';
      const row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;gap:10px;padding:6px 16px;margin-bottom:3px;border-radius:5px;background:var(--bg-section)';
      const cvs=_legendShapeCanvas(shape,color);
      const name=document.createElement('span');
      name.style.cssText='font-size:13px;color:var(--text);flex:1';
      name.textContent=sg.name;
      const cnt=document.createElement('span');
      cnt.style.cssText='font-size:11px;color:var(--text-muted)';
      const n=activities.filter(a=>a.groups&&a.groups.includes(sg.id)).length;
      cnt.textContent=n+' activit'+(n===1?'y':'ies');
      row.appendChild(cvs);row.appendChild(name);row.appendChild(cnt);
      el.appendChild(row);
    });
    const sp=document.createElement('div');sp.style.height='10px';el.appendChild(sp);
  });
}

function _renderLegendActivities(el){
  el.innerHTML='';
  if(!activities.length){
    el.innerHTML='<p style="color:var(--text-muted);font-size:13px;padding:16px">No activities added yet.</p>';
    return;
  }
  const cats=groups.filter(g=>g.parentId===null);
  const shown=new Set();
  cats.forEach(cat=>{
    const catActs=activities.filter((a,i)=>{
      if(shown.has(i))return false;
      return a.groups&&a.groups.some(gid=>{
        const g=groups.find(x=>x.id===gid);
        return g&&(g.id===cat.id||g.parentId===cat.id);
      });
    });
    if(!catActs.length)return;
    const hdr=document.createElement('div');
    hdr.style.cssText='font-weight:700;font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em;padding:8px 0 4px;border-bottom:1px solid var(--border);margin-bottom:4px';
    hdr.textContent=cat.name;el.appendChild(hdr);
    catActs.forEach(a=>{shown.add(activities.indexOf(a));el.appendChild(_buildActivityLegendRow(a));});
    const sp=document.createElement('div');sp.style.height='6px';el.appendChild(sp);
  });
  const ungrouped=activities.filter((_,i)=>!shown.has(i));
  if(ungrouped.length){
    if(cats.length){
      const hdr=document.createElement('div');
      hdr.style.cssText='font-weight:700;font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em;padding:8px 0 4px;border-bottom:1px solid var(--border);margin-bottom:4px';
      hdr.textContent='Ungrouped';el.appendChild(hdr);
    }
    ungrouped.forEach(a=>el.appendChild(_buildActivityLegendRow(a)));
  }
}

function _buildActivityLegendRow(a){
  const row=document.createElement('div');
  row.style.cssText='display:flex;align-items:center;gap:10px;padding:5px 8px;margin-bottom:2px;border-radius:4px;background:var(--bg-section)';
  const color=getActivityDrawColor(a);
  row.appendChild(_legendShapeCanvas(a.shape||'rect',color));
  const lbl=document.createElement('span');
  lbl.style.cssText='font-size:12px;color:var(--text)';
  lbl.textContent=a.name+(a.p6Id?` (${a.p6Id})`:'');
  row.appendChild(lbl);
  return row;
}

function printLegend(){
  const mode=document.querySelector('input[name="legendMode"]:checked')?.value||'groups';
  const projectName=titleBlock.projectName||'Programme';
  const printDate=new Date().toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  const items=[];
  if(mode==='groups'){
    groups.filter(g=>g.parentId===null).forEach(cat=>{
      items.push({type:'cat',name:cat.name,color:cat.color||'#3498db'});
      groups.filter(g=>g.parentId===cat.id).forEach(sg=>{
        const sample=activities.find(a=>a.groups&&a.groups.includes(sg.id));
        const n=activities.filter(a=>a.groups&&a.groups.includes(sg.id)).length;
        items.push({type:'item',name:sg.name,shape:sample?.shape||'rect',color:sg.color||cat.color||'#3498db',count:n});
      });
    });
  } else {
    activities.forEach(a=>items.push({type:'item',name:a.name+(a.p6Id?` (${a.p6Id})`:''),shape:a.shape||'rect',color:getActivityDrawColor(a)}));
  }
  const win=window.open('','_blank');
  if(!win){alert('Pop-up blocked.');return;}
  win.document.write(`<!DOCTYPE html><html><head><title>${projectName} — Legend</title><style>
    *{margin:0;padding:0;box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
    body{font-family:'Segoe UI',Arial,sans-serif;padding:10mm;color:#111}
    h1{font-size:18px;margin-bottom:3px}.sub{font-size:11px;color:#666;margin-bottom:14px;padding-bottom:6px;border-bottom:2px solid #333}
    .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:6px 12px}
    .cat{grid-column:1/-1;font-weight:700;font-size:12px;text-transform:uppercase;color:#555;padding:8px 0 3px;border-bottom:1px solid #ddd;margin-top:8px;display:flex;align-items:center;gap:8px}
    .cat:first-child{margin-top:0}
    .row{display:flex;align-items:center;gap:8px;padding:5px 6px;border:1px solid #eee;border-radius:4px;background:#f9f9f9}
    .lbl{font-size:11px;color:#222;flex:1}.cnt{font-size:10px;color:#888}
    canvas{display:block;flex-shrink:0}
    @media print{body{padding:8mm}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  </style></head><body>
  <h1>${projectName} — Legend</h1>
  <div class="sub">${mode==='groups'?'Categories & Sub-groups':'All Activities'} · ${printDate}</div>
  <div class="grid" id="g"></div>
  <script>
  (function(){
    const items=${JSON.stringify(items)};
    const g=document.getElementById('g');
    function drawShape(cvs,shape,color){
      const c=cvs.getContext('2d'),w=cvs.width,h=cvs.height,cx=w/2,cy=h/2,r=7;
      c.fillStyle=color;c.strokeStyle=color;c.globalAlpha=.9;
      if(shape==='rect')c.fillRect(2,cy-4,w-4,8);
      else if(shape==='line'){c.lineWidth=2.5;c.beginPath();c.moveTo(2,cy);c.lineTo(w-2,cy);c.stroke();}
      else if(shape==='circle'){c.beginPath();c.arc(cx,cy,r,0,Math.PI*2);c.fill();}
      else if(shape==='diamond'){c.beginPath();c.moveTo(cx,2);c.lineTo(w-2,cy);c.lineTo(cx,h-2);c.lineTo(2,cy);c.closePath();c.fill();}
      else if(shape==='triangle'){c.beginPath();c.moveTo(cx,2);c.lineTo(w-2,h-2);c.lineTo(2,h-2);c.closePath();c.fill();}
      else if(shape==='star'){const pts=5,ri=r*.42;c.beginPath();for(let i=0;i<pts*2;i++){const a=((i*Math.PI)/pts)-Math.PI/2;const rr=i%2===0?r:ri;i===0?c.moveTo(cx+rr*Math.cos(a),cy+rr*Math.sin(a)):c.lineTo(cx+rr*Math.cos(a),cy+rr*Math.sin(a));}c.closePath();c.fill();}
      else if(shape==='flag'){c.lineWidth=1.5;c.beginPath();c.moveTo(cx-2,2);c.lineTo(cx-2,h-2);c.stroke();c.beginPath();c.moveTo(cx-2,2);c.lineTo(w-2,cy-3);c.lineTo(cx-2,cy+2);c.closePath();c.fill();}
      else c.fillRect(2,cy-4,w-4,8);
    }
    items.forEach(item=>{
      if(item.type==='cat'){
        const d=document.createElement('div');d.className='cat';
        const dot=document.createElement('span');dot.style.cssText='width:10px;height:10px;border-radius:50%;background:'+item.color+';display:inline-block;flex-shrink:0';
        d.appendChild(dot);d.appendChild(document.createTextNode(item.name));
        g.appendChild(d);return;
      }
      const row=document.createElement('div');row.className='row';
      const cvs=document.createElement('canvas');cvs.width=36;cvs.height=22;
      drawShape(cvs,item.shape,item.color);
      row.appendChild(cvs);
      const lbl=document.createElement('span');lbl.className='lbl';lbl.textContent=item.name;row.appendChild(lbl);
      if(item.count!=null){const c2=document.createElement('span');c2.className='cnt';c2.textContent=item.count+'×';row.appendChild(c2);}
      g.appendChild(row);
    });
    window.onload=()=>{window.focus();window.print();};
  })();
  <\/script></body></html>`);
  win.document.close();
}

// ═══════════════════════════════════════════════════════════════════
//  ONBOARDING
// ═══════════════════════════════════════════════════════════════════
function openTab2(name){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById(name)?.classList.add('active');
  document.querySelectorAll('.tab').forEach(t=>{
    t.classList.toggle('active',t.getAttribute('onclick')===`openTab('${name}')`);
  });
  const cw=document.querySelector('.canvas-wrap');
  const ct=document.querySelector('.canvas-toolbar');
  if(cw) cw.style.display=name==='dashboard'?'none':'';
  if(ct) ct.style.display=name==='dashboard'?'none':'';
}

function openSamplePicker(){
  document.getElementById('samplePickerModal').style.display='flex';
}

function loadSample(n){
  document.getElementById('samplePickerModal').style.display='none';
  if(activities.length>0){
    if(!confirm(`⚠️ Loading this sample will replace your current programme (${activities.length} activit${activities.length===1?'y':'ies'}).\n\nYour work is auto-saved — Ctrl+Z will recover it.\n\nContinue?`)) return;
  }
  if(n===1) _loadSample1();
  else if(n===2) _loadSample2();
  else if(n===3) _loadSample3();
  else if(n===4) _loadSample4();
  else if(n===5) _loadSample5();
}

// ── Shared helpers ────────────────────────────────────────────────
function _sAct(p6Id,name,start,end,sCh,eCh,color,alpha,shape,label,notes,grps,outSt,outC,outW,res,labRate,equip,mat){
  return{p6Id,name,start,end,startCh:sCh,endCh:eCh,color,alpha,shape,label,progress:0,notes,groups:grps,
    outlineStyle:outSt,outlineColor:outC,outlineWidth:outW,resourceCount:res,unitCost:labRate,equipmentCost:equip,materialCost:mat};
}
function _sLnk(f,t,type='FS',lag=0){return{id:genId(),fromIdx:f,toIdx:t,type,lag};}
// ═══════════════════════════════════════════════════════════════════
//  RISK REGISTER — CORE FUNCTIONS
// ═══════════════════════════════════════════════════════════════════

function riskScore(r){ return (parseInt(r.likelihood)||1)*(parseInt(r.impact)||1); }
function riskLevel(score){ return score>=16?'critical':score>=9?'high':score>=4?'medium':'low'; }
function riskLevelLabel(score){ return score>=16?'Critical':score>=9?'High':score>=4?'Medium':'Low'; }
function riskColor(score){ return {critical:'#8e44ad',high:'#e74c3c',medium:'#f39c12',low:'#27ae60'}[riskLevel(score)]; }

// Called when the ⚠️ Risk Register tab is opened
function openTab_risks(){
  populateRiskActivityDropdown();
  renderRiskTable();
  renderRiskMatrix('riskMatrix');
  renderRiskSummaryCards();
  renderRiskAdjustedCompletion();
  renderFloatRiskAlerts();
  const sym = document.getElementById('riskCurrSymbol');
  if(sym) sym.textContent = projectCurrency||'£';
}

function populateRiskActivityDropdown(){
  const sel = document.getElementById('riskActivity'); if(!sel) return;
  const cur = sel.value;
  sel.innerHTML = '<option value="">— None (programme-level risk) —</option>';
  activities.forEach((a,i)=>{
    const o = document.createElement('option');
    o.value = a.p6Id || String(i);
    o.textContent = (a.p6Id ? a.p6Id+' — ' : '') + (a.name||'Unnamed');
    sel.appendChild(o);
  });
  if(cur) sel.value = cur;
}

function saveRisk(){
  const title = (document.getElementById('riskTitle')?.value||'').trim();
  if(!title){ alert('Please enter a risk title.'); return; }
  const editId = document.getElementById('riskEditId')?.value || '';
  const r = {
    id:                 editId || genId(),
    activityId:         document.getElementById('riskActivity')?.value || '',
    title,
    likelihood:         parseInt(document.getElementById('riskLikelihood')?.value)||3,
    impact:             parseInt(document.getElementById('riskImpact')?.value)||3,
    timeDays:           parseFloat(document.getElementById('riskTimeDays')?.value)||0,
    costImpact:         parseFloat(document.getElementById('riskCost')?.value)||0,
    owner:              document.getElementById('riskOwner')?.value?.trim()||'',
    mitigation:         document.getElementById('riskMitigation')?.value?.trim()||'',
    mitigationProgress: parseInt(document.getElementById('riskMitigationProgress')?.value)||0,
    dueDate:            document.getElementById('riskDueDate')?.value||'',
    status:             document.getElementById('riskStatus')?.value||'Open',
    created:            editId ? (risks.find(x=>x.id===editId)||{}).created||new Date().toISOString() : new Date().toISOString(),
  };
  if(editId){ const i=risks.findIndex(x=>x.id===editId); if(i>=0)risks[i]=r; else risks.push(r); }
  else risks.push(r);
  pushUndo();
  clearRiskForm();
  renderRiskTable();
  renderRiskMatrix('riskMatrix');
  renderRiskSummaryCards();
  renderRiskAdjustedCompletion();
  renderFloatRiskAlerts();
  drawChart();
  flashSaved('✅ Risk saved');
}

function clearRiskForm(){
  ['riskEditId','riskTitle','riskOwner','riskDueDate'].forEach(id=>{ const el=document.getElementById(id); if(el)el.value=''; });
  const ra=document.getElementById('riskActivity'); if(ra)ra.value='';
  const rl=document.getElementById('riskLikelihood'); if(rl)rl.value='3';
  const ri=document.getElementById('riskImpact'); if(ri)ri.value='3';
  const rt=document.getElementById('riskTimeDays'); if(rt)rt.value='0';
  const rc=document.getElementById('riskCost'); if(rc)rc.value='0';
  const rs=document.getElementById('riskStatus'); if(rs)rs.value='Open';
  const rm=document.getElementById('riskMitigation'); if(rm)rm.value='';
  const rp=document.getElementById('riskMitigationProgress'); if(rp)rp.value='0';
  const rv=document.getElementById('riskProgressVal'); if(rv)rv.textContent='0%';
  const btn=document.querySelector('#riskFormBox button');
  if(btn) btn.textContent='＋ Add Risk';
}

function editRisk(id){
  const r=risks.find(x=>x.id===id); if(!r)return;
  document.getElementById('riskEditId').value=r.id;
  document.getElementById('riskTitle').value=r.title;
  document.getElementById('riskActivity').value=r.activityId||'';
  document.getElementById('riskLikelihood').value=r.likelihood;
  document.getElementById('riskImpact').value=r.impact;
  document.getElementById('riskTimeDays').value=r.timeDays||0;
  document.getElementById('riskCost').value=r.costImpact||0;
  document.getElementById('riskOwner').value=r.owner||'';
  document.getElementById('riskMitigation').value=r.mitigation||'';
  document.getElementById('riskStatus').value=r.status||'Open';
  const dp=document.getElementById('riskDueDate'); if(dp)dp.value=r.dueDate||'';
  const rp=document.getElementById('riskMitigationProgress'); if(rp)rp.value=r.mitigationProgress||0;
  const rv=document.getElementById('riskProgressVal'); if(rv)rv.textContent=(r.mitigationProgress||0)+'%';
  const btn=document.querySelector('#riskFormBox button');
  if(btn) btn.textContent='💾 Update Risk';
  document.getElementById('riskFormBox')?.scrollIntoView({behavior:'smooth'});
}

function deleteRisk(id){
  if(!confirm('Delete this risk?')) return;
  pushUndo();
  risks=risks.filter(r=>r.id!==id);
  renderRiskTable(); renderRiskMatrix('riskMatrix'); renderRiskSummaryCards();
  renderRiskAdjustedCompletion(); renderFloatRiskAlerts();
  drawChart();
}

function renderRiskTable(){
  const el=document.getElementById('riskTableContainer'); if(!el)return;
  const statusFilter=document.getElementById('riskFilterStatus')?.value||'';
  const likFilter=parseInt(document.getElementById('riskFilterLikelihood')?.value||'0');
  const sym=projectCurrency||'£';
  const fmtC=v=>v>=1e6?sym+(v/1e6).toFixed(1)+'M':v>=1000?sym+(v/1000).toFixed(0)+'k':sym+Math.round(v);

  let filtered=risks.filter(r=>{
    if(statusFilter&&r.status!==statusFilter)return false;
    if(likFilter&&(r.likelihood||1)<likFilter)return false;
    return true;
  }).sort((a,b)=>riskScore(b)-riskScore(a));

  const countEl=document.getElementById('riskCount');
  if(countEl) countEl.textContent=`${filtered.length} of ${risks.length} risk${risks.length===1?'':'s'}`;

  if(!filtered.length){
    el.innerHTML=`<div style="padding:28px;text-align:center;color:var(--text-muted);font-size:13px">
      ${risks.length?'No risks match the current filter.':'No risks logged yet — fill in the form above and click ＋ Add Risk.'}
    </div>`;
    return;
  }

  const statusStyle={
    'Open':       'background:#e74c3c;color:#fff',
    'Mitigated':  'background:#f39c12;color:#fff',
    'Closed':     'background:#27ae60;color:#fff',
    'Transferred':'background:#3498db;color:#fff',
  };

  const today = new Date(); today.setHours(0,0,0,0);

  const rows=filtered.map(r=>{
    const score=riskScore(r), col=riskColor(score), lvl=riskLevelLabel(score);
    const act=activities.find(a=>a.p6Id===r.activityId)||(r.activityId?null:null);
    const actLabel=act?`${act.p6Id||''} — ${act.name}`:r.activityId||'—';
    const ss=statusStyle[r.status]||'background:#95a5a6;color:#fff';
    const prog=r.mitigationProgress||0;
    const progColor=prog>=80?'#27ae60':prog>=40?'#f39c12':'#e74c3c';
    const due=r.dueDate?new Date(r.dueDate):null;
    const isOverdue=due&&r.status==='Open'&&due<today;
    const dueStr=due?due.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'2-digit'}):'—';
    return `<tr${isOverdue?' style="background:rgba(231,76,60,0.07)"':''}>
      <td><span style="display:inline-block;background:${col};color:#fff;font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;white-space:nowrap">${lvl} (${score})</span></td>
      <td style="font-weight:600;max-width:200px">${r.title}</td>
      <td style="text-align:center;font-weight:700">${r.likelihood}</td>
      <td style="text-align:center;font-weight:700">${r.impact}</td>
      <td style="font-size:11px;color:var(--text-muted);max-width:130px">${actLabel}</td>
      <td style="color:#e74c3c;font-weight:600;white-space:nowrap">${r.timeDays?'+'+r.timeDays+'d':'—'}</td>
      <td style="color:#e74c3c;font-weight:600;white-space:nowrap">${r.costImpact?fmtC(r.costImpact):'—'}</td>
      <td style="font-size:11px;color:var(--text-muted)">${r.owner||'—'}</td>
      <td><span style="${ss};font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;display:inline-block">${r.status}</span></td>
      <td style="min-width:110px">
        <div style="display:flex;align-items:center;gap:5px">
          <div style="flex:1;height:6px;background:var(--border);border-radius:3px;overflow:hidden">
            <div style="width:${prog}%;height:100%;background:${progColor};border-radius:3px"></div>
          </div>
          <span style="font-size:10px;color:var(--text-muted);flex-shrink:0">${prog}%</span>
        </div>
        <div style="font-size:10px;color:${isOverdue?'#e74c3c':'var(--text-muted)'};margin-top:2px">
          ${isOverdue?'⚠️ Overdue: ':'Due: '}${dueStr}
        </div>
      </td>
      <td style="font-size:11px;max-width:160px;color:var(--text-muted)">${r.mitigation||'—'}</td>
      <td style="white-space:nowrap">
        <button onclick="editRisk('${r.id}')" style="margin:0;padding:2px 8px;font-size:11px;background:var(--tab-active-bg);color:#fff;border:none;border-radius:4px;cursor:pointer">✏️</button>
        <button onclick="deleteRisk('${r.id}')" style="margin:0 0 0 4px;padding:2px 8px;font-size:11px;background:#e74c3c;color:#fff;border:none;border-radius:4px;cursor:pointer">✕</button>
      </td>
    </tr>`;
  }).join('');

  el.innerHTML=`<table class="risk-table">
    <thead><tr>
      <th>Score</th><th>Risk</th><th>L</th><th>I</th><th>Activity</th>
      <th>Time</th><th>Cost</th><th>Owner</th><th>Status</th><th>Mitigation Progress</th><th>Action</th><th></th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function renderRiskSummaryCards(){
  const el=document.getElementById('riskSummaryCards'); if(!el)return;
  const sym=projectCurrency||'£';
  const fmtC=v=>v>=1e6?sym+(v/1e6).toFixed(1)+'M':v>=1000?sym+(v/1000).toFixed(0)+'k':sym+Math.round(v);
  const open=risks.filter(r=>r.status==='Open');
  const critical=open.filter(r=>riskScore(r)>=16).length;
  const high=open.filter(r=>riskScore(r)>=9&&riskScore(r)<16).length;
  const weightedTime=open.reduce((s,r)=>s+(r.timeDays||0)*(r.likelihood/5),0);
  const weightedCost=open.reduce((s,r)=>s+(r.costImpact||0)*(r.likelihood/5),0);
  const resolved=risks.filter(r=>r.status==='Mitigated'||r.status==='Closed').length;
  el.innerHTML=`
    <div class="risk-summary-card" style="border-left:4px solid #e74c3c">
      <div class="risk-sc-val">${open.length}</div>
      <div class="risk-sc-lbl">Open Risks</div>
    </div>
    <div class="risk-summary-card" style="border-left:4px solid #8e44ad">
      <div class="risk-sc-val">${critical}</div>
      <div class="risk-sc-lbl">Critical</div>
    </div>
    <div class="risk-summary-card" style="border-left:4px solid #e74c3c">
      <div class="risk-sc-val">${high}</div>
      <div class="risk-sc-lbl">High</div>
    </div>
    <div class="risk-summary-card" style="border-left:4px solid #e67e22">
      <div class="risk-sc-val">${Math.round(weightedTime)}d</div>
      <div class="risk-sc-lbl">Weighted Time</div>
    </div>
    <div class="risk-summary-card" style="border-left:4px solid #e74c3c">
      <div class="risk-sc-val">${fmtC(weightedCost)}</div>
      <div class="risk-sc-lbl">Weighted Cost</div>
    </div>
    <div class="risk-summary-card" style="border-left:4px solid #27ae60">
      <div class="risk-sc-val">${resolved}</div>
      <div class="risk-sc-lbl">Resolved</div>
    </div>`;
}

function renderRiskMatrix(canvasId){
  const cvs=document.getElementById(canvasId); if(!cvs)return;
  const W=cvs.width||200, H=cvs.height||180;
  const c=cvs.getContext('2d');
  c.clearRect(0,0,W,H);

  const GRID=5, padL=22, padB=22, padT=8, padR=8;
  const gW=(W-padL-padR)/GRID, gH=(H-padT-padB)/GRID;

  // Colour-coded background cells
  for(let l=0;l<GRID;l++){
    for(let i=0;i<GRID;i++){
      const score=(l+1)*(i+1);
      const col=score>=16?'rgba(142,68,173,0.22)':score>=9?'rgba(231,76,60,0.22)':score>=4?'rgba(243,156,18,0.20)':'rgba(39,174,96,0.20)';
      c.fillStyle=col;
      c.fillRect(padL+l*gW, padT+(GRID-1-i)*gH, gW, gH);
      c.strokeStyle='rgba(180,180,180,0.3)'; c.lineWidth=0.5;
      c.strokeRect(padL+l*gW, padT+(GRID-1-i)*gH, gW, gH);
    }
  }

  // Axis labels
  c.fillStyle='#888'; c.font='9px Arial'; c.textAlign='center'; c.textBaseline='middle';
  for(let n=1;n<=5;n++){
    c.fillText(String(n), padL+(n-1)*gW+gW/2, H-padB+10);
    c.fillText(String(n), padL-9, padT+(GRID-n)*gH+gH/2);
  }
  c.font='bold 8px Arial'; c.fillStyle='#555';
  c.fillText('Likelihood →', padL+gW*2.5, H-3);
  c.save(); c.translate(8, padT+gH*2.5); c.rotate(-Math.PI/2);
  c.fillText('Impact ↑', 0, 0); c.restore();

  // Plot open risk dots
  const open=risks.filter(r=>r.status==='Open');
  const placed={};
  open.forEach(r=>{
    const l=Math.min(5,Math.max(1,parseInt(r.likelihood)||1));
    const imp=Math.min(5,Math.max(1,parseInt(r.impact)||1));
    const key=`${l},${imp}`;
    placed[key]=(placed[key]||0);
    const idx=placed[key]; placed[key]++;
    const cx=padL+(l-1)*gW+gW/2+(idx%3-1)*6;
    const cy=padT+(GRID-imp)*gH+gH/2+Math.floor(idx/3)*6;
    const col=riskColor(riskScore(r));
    c.beginPath(); c.arc(cx,cy,5.5,0,Math.PI*2);
    c.fillStyle=col; c.globalAlpha=0.88; c.fill();
    c.globalAlpha=1; c.strokeStyle='#fff'; c.lineWidth=1.5; c.stroke();
  });
}

// ── FEATURE 1: Risk-Adjusted Completion Dates ─────────────────────
// Uses open risks' weighted time impacts to calculate P50/P80/P90 dates
// and draws them as coloured marker lines on the chart.
// Returns the latest end date of physical work activities (rect or line shapes only — not milestones)
function getPhysicalCompletionDate(){
  const physical = activities.filter(a=>{
    const s = (a.shape||'rect').toLowerCase();
    return s==='rect'||s==='line';
  });
  if(!physical.length) return null;
  return physical.reduce((max,a)=>a.end&&a.end>max?a.end:max, new Date(0));
}

function computeRiskAdjustedDates(){
  if(!activities.length||!timelineEnd) return null;

  // Base completion = latest PHYSICAL activity end (rect or line only — not milestones)
  const physEnd = getPhysicalCompletionDate();
  if(!physEnd||physEnd.getTime()===0) return null;
  const baseEnd = physEnd;

  // Open risks with a time impact
  const openRisks = risks.filter(r=>r.status==='Open'&&(r.timeDays||0)>0);
  if(!openRisks.length) return {baseEnd, expectedDelay:0, p50:null, p80:null, p90:null};

  // Weighted expected delay: timeDays × (likelihood/5) as probability
  const expectedDelay = openRisks.reduce((s,r)=>s+(r.timeDays||0)*((r.likelihood||1)/5), 0);

  const MS = 86400000;
  const p50 = new Date(baseEnd.getTime() + Math.round(expectedDelay)        * MS);
  const p80 = new Date(baseEnd.getTime() + Math.round(expectedDelay * 1.6)  * MS);
  const p90 = new Date(baseEnd.getTime() + Math.round(expectedDelay * 2.2)  * MS);

  return {baseEnd, expectedDelay: Math.round(expectedDelay), p50, p80, p90};
}

function renderRiskAdjustedCompletion(){
  const el = document.getElementById('riskAdjustedCompletion'); if(!el) return;
  const open = risks.filter(r=>r.status==='Open'&&(r.timeDays||0)>0);

  if(!open.length){
    el.innerHTML=`<p style="font-size:12px;color:var(--text-muted)">No open risks with time impacts logged. Add time impact (days) to risks to see adjusted completion dates.</p>`;
    drawChart(); return;
  }

  const res = computeRiskAdjustedDates();
  if(!res){el.innerHTML='<p style="font-size:12px;color:var(--text-muted)">Set your timeline and add physical activities to see adjusted dates.</p>';return;}

  const fmt = d => d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  const fmtDiff = d => {
    const diff = Math.round((d-res.baseEnd)/86400000);
    return diff>0?`<span style="color:#e74c3c">+${diff}d</span>`:diff===0?'<span style="color:#27ae60">on time</span>':`<span style="color:#27ae60">${diff}d</span>`;
  };

  el.innerHTML=`
    <div style="font-size:11px;color:var(--text-muted);margin-bottom:8px">
      Based on ${open.length} open risk${open.length===1?'':'s'} with <strong>${res.expectedDelay} weighted days</strong> of exposure.
      Base = latest rect/line activity end (milestones excluded). Lines shown on chart when Risk Overlays is checked.
    </div>
    <div style="display:grid;grid-template-columns:auto 1fr auto;gap:5px 12px;align-items:center;font-size:12px">
      <span style="font-size:11px;font-weight:700;color:#2c3e50">📅 Base</span>
      <span>${fmt(res.baseEnd)}</span>
      <span style="font-size:10px;background:#2c3e50;color:#fff;padding:1px 8px;border-radius:8px">Physical Completion</span>

      ${res.p50?`
      <span style="font-size:11px;font-weight:700;color:#f39c12">P50</span>
      <span>${fmt(res.p50)}</span>
      <span style="font-size:11px;font-weight:600">${fmtDiff(res.p50)}</span>

      <span style="font-size:11px;font-weight:700;color:#e67e22">P80</span>
      <span>${fmt(res.p80)}</span>
      <span style="font-size:11px;font-weight:600">${fmtDiff(res.p80)}</span>

      <span style="font-size:11px;font-weight:700;color:#e74c3c">P90</span>
      <span>${fmt(res.p90)}</span>
      <span style="font-size:11px;font-weight:600">${fmtDiff(res.p90)}</span>`:''}
    </div>
    <div style="font-size:10px;color:var(--text-muted);margin-top:8px;border-top:1px solid var(--border);padding-top:6px">
      P50 = 50% confidence · P80 = 80% confidence (typical client commitment) · P90 = 90% confidence (contractual buffer)
    </div>`;

  drawChart(); // redraw to show P-lines
}

// Draw P50/P80/P90 lines on the canvas — called from drawChart and renderOffscreen
function _drawRiskPLines(c, canvasW, gridH, ct, forceShow){
  // Controlled by the Risk Overlays toggle (or forceShow override for exports)
  const show = forceShow!==undefined ? forceShow : document.getElementById('showRiskOverlaysToggle')?.checked;
  if(!show) return;
  if(!timelineStart||!timelineEnd) return;

  const totalMs = timelineEnd - timelineStart;
  const topM = getTopMargin();

  function yFor(date){
    const off = date - timelineStart;
    return topM + (off / totalMs) * gridH;
  }

  function drawPill(yPos, color, label){
    if(yPos < topM - 20 || yPos > topM + gridH + 20) return; // out of view
    c.save();
    c.strokeStyle=color; c.lineWidth=1.5; c.setLineDash([5,3]);
    c.beginPath(); c.moveTo(LEFT_MARGIN, yPos); c.lineTo(canvasW-40, yPos); c.stroke();
    c.setLineDash([]);
    c.font='bold 10px Arial';
    const lbl=`◀ ${label}`;
    const tw=c.measureText(lbl).width;
    c.fillStyle=color; c.globalAlpha=0.92;
    const px=canvasW-44-tw-6;
    c.beginPath();
    if(c.roundRect) c.roundRect(px,yPos-9,tw+10,17,4);
    else c.rect(px,yPos-9,tw+10,17);
    c.fill();
    c.globalAlpha=1; c.fillStyle='#fff';
    c.textAlign='right'; c.textBaseline='middle';
    c.fillText(lbl, canvasW-46, yPos);
    c.restore();
  }

  // ── Project Completion line — latest physical activity end ──
  const physEnd = getPhysicalCompletionDate();
  if(physEnd && physEnd.getTime()>0){
    drawPill(yFor(physEnd), '#2c3e50', 'Project Completion');
  }

  // ── P50 / P80 / P90 lines — only if risks with time impacts exist ──
  const res = computeRiskAdjustedDates();
  if(!res||!res.p50) return;

  [
    {date:res.p50, color:'#f39c12', label:'P50'},
    {date:res.p80, color:'#e67e22', label:'P80'},
    {date:res.p90, color:'#e74c3c', label:'P90'},
  ].forEach(pl=>drawPill(yFor(pl.date), pl.color, pl.label));
}

// ── FEATURE 3: Float vs Risk alerts ───────────────────────────────
// Compares each activity's total float (from CPM) against the worst
// open risk time impact attached to it. Flags where risk > float.
function computeActivityFloats(){
  if(!activities.length||!activityLinks.length) return {};

  // Forward pass — Earliest Start/Finish
  const ES={}, EF={};
  const topoOrder=[];
  const inDeg=activities.map(()=>0);
  const adj=activities.map(()=>[]);

  activityLinks.forEach(l=>{
    if(l.fromIdx<activities.length&&l.toIdx<activities.length){
      adj[l.fromIdx].push(l.toIdx);
      inDeg[l.toIdx]++;
    }
  });

  const queue=[];
  activities.forEach((_,i)=>{ if(inDeg[i]===0) queue.push(i); });
  while(queue.length){
    const n=queue.shift(); topoOrder.push(n);
    adj[n].forEach(s=>{ if(--inDeg[s]===0) queue.push(s); });
  }

  const dur=i=>{
    const a=activities[i];
    if(!a||!a.start||!a.end) return 0;
    return Math.max(0,Math.round((a.end-a.start)/86400000));
  };

  topoOrder.forEach(i=>{ ES[i]=0; EF[i]=dur(i); });
  topoOrder.forEach(i=>{
    adj[i].forEach(j=>{
      if((EF[i]||0)>(ES[j]||0)){
        ES[j]=EF[i]||0;
        EF[j]=(ES[j]||0)+dur(j);
      }
    });
  });

  // Project end = max EF
  const projEnd=Math.max(0,...Object.values(EF));

  // Backward pass — Latest Start/Finish
  const LS={}, LF={};
  topoOrder.forEach(i=>{ LF[i]=projEnd; LS[i]=projEnd-dur(i); });
  [...topoOrder].reverse().forEach(i=>{
    adj[i].forEach(j=>{
      if((LS[j]||0)<(LF[i]||projEnd)){
        LF[i]=LS[j]||0;
        LS[i]=(LF[i]||0)-dur(i);
      }
    });
  });

  const floats={};
  activities.forEach((_,i)=>{
    floats[i]=Math.max(0,(LS[i]||0)-(ES[i]||0));
  });
  return floats;
}

function renderFloatRiskAlerts(){
  const el=document.getElementById('floatRiskAlerts'); if(!el)return;
  if(!risks.length||!activityLinks.length){el.innerHTML='';return;}

  const floats=computeActivityFloats();
  const alerts=[];
  const today=new Date();

  activities.forEach((a,i)=>{
    const key=a.p6Id||String(i);
    const aRisks=risks.filter(r=>r.status==='Open'&&r.activityId===key&&(r.timeDays||0)>0);
    if(!aRisks.length) return;
    const worstRisk=aRisks.reduce((w,r)=>r.timeDays>w.timeDays?r:w,aRisks[0]);
    const float=floats[i]??null;
    if(float===null) return;

    const riskDays=worstRisk.timeDays||0;
    const exceeded=riskDays>float;
    const tight=!exceeded&&riskDays>=float*0.6&&float>0;

    if(exceeded||tight){
      alerts.push({a,i,float,riskDays,worstRisk,exceeded,tight});
    }
  });

  // Also flag overdue mitigations
  const overdue=risks.filter(r=>{
    if(r.status!=='Open'||!r.dueDate) return false;
    const due=new Date(r.dueDate); due.setHours(0,0,0,0);
    return due<today;
  });

  if(!alerts.length&&!overdue.length){el.innerHTML='';return;}

  const alertRows=alerts.map(({a,float,riskDays,worstRisk,exceeded})=>`
    <div class="risk-alert-row ${exceeded?'risk-alert-exceeded':'risk-alert-tight'}">
      <span class="risk-alert-icon">${exceeded?'🔴':'🟡'}</span>
      <div style="flex:1">
        <strong>${a.name}</strong>${a.p6Id?` <span style="font-size:11px;color:var(--text-muted)">(${a.p6Id})</span>`:''}
        <div style="font-size:12px;margin-top:2px">
          ${exceeded
            ?`Risk <strong style="color:#e74c3c">exceeds float</strong>: worst risk is <strong>+${riskDays}d</strong> but only <strong>${float}d float</strong> available — this activity is on a <strong>hidden critical path</strong>`
            :` Risk is <strong style="color:#f39c12">${Math.round(riskDays/float*100)}% of available float</strong> (${riskDays}d risk vs ${float}d float)`
          }
        </div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">Highest risk: "${worstRisk.title}"</div>
      </div>
    </div>`).join('');

  const overdueRows=overdue.map(r=>`
    <div class="risk-alert-row risk-alert-exceeded">
      <span class="risk-alert-icon">⚠️</span>
      <div style="flex:1">
        <strong>Overdue mitigation</strong>
        <div style="font-size:12px;margin-top:2px">"${r.title}"</div>
        <div style="font-size:11px;color:#e74c3c">Due: ${new Date(r.dueDate).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})} · Owner: ${r.owner||'—'}</div>
      </div>
    </div>`).join('');

  el.innerHTML=`
    <div style="font-size:11px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">
      ⚡ Risk Intelligence — Float &amp; Mitigation Alerts
    </div>
    ${alertRows}${overdueRows}`;
}

function _drawRiskIndicators(c, gridW, gridH, ct, forceShow){
  const show = forceShow!==undefined ? forceShow : document.getElementById('showRiskOverlaysToggle')?.checked;
  if(!show) return;
  if(!risks||!risks.length) return;
  const openByAct={};
  risks.filter(r=>r.status==='Open'&&r.activityId).forEach(r=>{
    if(!openByAct[r.activityId]) openByAct[r.activityId]=[];
    openByAct[r.activityId].push(r);
  });
  if(!Object.keys(openByAct).length) return;
  activities.forEach((a,i)=>{
    const key=a.p6Id||String(i);
    const aRisks=openByAct[key]; if(!aRisks||!a.coords) return;
    const maxScore=Math.max(...aRisks.map(r=>riskScore(r)));
    const col=riskColor(maxScore);
    const cx=Math.min(a.coords.x1,a.coords.x2)+9;
    const cy=Math.min(a.coords.y1,a.coords.y2)+9;
    c.save();
    c.beginPath(); c.arc(cx,cy,6.5,0,Math.PI*2);
    c.fillStyle=col; c.globalAlpha=0.93; c.fill();
    c.strokeStyle='#fff'; c.lineWidth=1.5; c.stroke();
    c.globalAlpha=1; c.fillStyle='#fff';
    c.font='bold 8px Arial'; c.textAlign='center'; c.textBaseline='middle';
    c.fillText(aRisks.length>9?'!':String(aRisks.length), cx, cy);
    c.restore();
  });
}

// Dashboard risk panels
function _renderDashRisks(){
  const el=document.getElementById('dashRisks');
  const expEl=document.getElementById('dashRiskExposure');
  if(!el) return;

  const sym=projectCurrency||'£';
  const fmtC=v=>v>=1e6?sym+(v/1e6).toFixed(1)+'M':v>=1000?sym+(v/1000).toFixed(0)+'k':sym+Math.round(v);
  const open=risks.filter(r=>r.status==='Open').sort((a,b)=>riskScore(b)-riskScore(a));

  if(!open.length){
    el.innerHTML='<p style="color:var(--text-muted);font-size:13px;padding:8px 0">No open risks. Go to the ⚠️ Risk Register tab to log risks.</p>';
  } else {
    el.innerHTML=open.slice(0,6).map(r=>{
      const score=riskScore(r), col=riskColor(score);
      return `<div style="display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid var(--border)">
        <span style="background:${col};color:#fff;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;flex-shrink:0;white-space:nowrap">${riskLevelLabel(score)}</span>
        <span style="font-size:12px;flex:1;color:var(--text)">${r.title}</span>
        ${r.timeDays?`<span style="font-size:11px;color:#e74c3c;font-weight:600;flex-shrink:0">+${r.timeDays}d</span>`:''}
      </div>`;
    }).join('');
    if(open.length>6) el.innerHTML+=`<div style="font-size:11px;color:var(--text-muted);padding:5px 0">+ ${open.length-6} more in Risk Register</div>`;
  }

  renderRiskMatrix('dashRiskMatrix');

  if(expEl){
    const open2=risks.filter(r=>r.status==='Open');
    const wTime=open2.reduce((s,r)=>s+(r.timeDays||0)*(r.likelihood/5),0);
    const wCost=open2.reduce((s,r)=>s+(r.costImpact||0)*(r.likelihood/5),0);
    const lv=(min,max)=>open2.filter(r=>{const s=riskScore(r);return s>=min&&s<max;}).length;
    const today=new Date(); today.setHours(0,0,0,0);
    const overdueMit=open2.filter(r=>r.dueDate&&new Date(r.dueDate)<today).length;
    const res=computeRiskAdjustedDates();
    const fmtD=d=>d?d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}):'—';
    expEl.innerHTML=`
      <div style="font-size:11px;color:var(--text-muted);margin-bottom:8px">Weighted exposure (likelihood × probability)</div>
      <div class="risk-exp-row"><span>⏱ Schedule exposure</span><strong style="color:#e74c3c">${Math.round(wTime)} days</strong></div>
      <div class="risk-exp-row"><span>💰 Cost exposure</span><strong style="color:#e74c3c">${fmtC(wCost)}</strong></div>
      ${res&&res.p50?`
      <div class="risk-exp-row" style="margin-top:6px"><span>📅 P50 completion</span><strong style="color:#f39c12">${fmtD(res.p50)}</strong></div>
      <div class="risk-exp-row"><span>📅 P80 completion</span><strong style="color:#e67e22">${fmtD(res.p80)}</strong></div>
      <div class="risk-exp-row"><span>📅 P90 completion</span><strong style="color:#e74c3c">${fmtD(res.p90)}</strong></div>`:''}
      ${overdueMit?`<div class="risk-exp-row" style="margin-top:6px"><span>⚠️ Overdue mitigations</span><strong style="color:#e74c3c">${overdueMit}</strong></div>`:''}
      <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:6px">
        <div class="risk-exp-badge" style="border-color:#8e44ad;color:#8e44ad">🔴 Critical: ${lv(16,26)}</div>
        <div class="risk-exp-badge" style="border-color:#e74c3c;color:#e74c3c">🟠 High: ${lv(9,16)}</div>
        <div class="risk-exp-badge" style="border-color:#f39c12;color:#f39c12">🟡 Medium: ${lv(4,9)}</div>
        <div class="risk-exp-badge" style="border-color:#27ae60;color:#27ae60">🟢 Low: ${lv(1,4)}</div>
      </div>`;
  }
}

function _sRisk(actId,title,likelihood,impact,timeDays,costImpact,owner,mitigation,status='Open'){
  return{id:genId(),activityId:actId,title,likelihood,impact,timeDays,costImpact,owner,mitigation,status};
}
function _sSetup(name,title,y,maxCh,spacing){
  pushUndo();
  timelineStart=new Date(y,0,1);timelineEnd=new Date(y+1,11,31);
  minChainage=0;maxChainage=maxCh;chainageSpacing=spacing;
  document.getElementById('timelineStart').value=`${y}-01-01`;
  document.getElementById('timelineEnd').value=`${y+1}-12-31`;
  document.getElementById('minChainage').value=0;
  document.getElementById('maxChainage').value=maxCh;
  document.getElementById('chainSpacing').value=spacing;
  titleBlock.projectName=name;titleBlock.drawingTitle=title;
  titleBlock.drawnBy='Sample';titleBlock.checkedBy='Planner';
  const td=new Date();
  titleBlock.dataDate=`${y}-${String(td.getMonth()+1).padStart(2,'0')}-${String(td.getDate()).padStart(2,'0')}`;
  syncTitleBlockUI();
}

// ═══════════════════════════════════════════════════════════════════
//  SAMPLE 1 — Highway Improvement Scheme
// ═══════════════════════════════════════════════════════════════════
function _loadSample1(){
  const y=new Date().getFullYear();
  _sSetup('Highway Improvement Scheme','Time-Chainage Construction Programme',y,1000,50);
  const d=(m,dy=1)=>new Date(y,m,dy);

  chainageZones=[
    {id:genId(),name:'Zone 1 — North',   startCh:0,  endCh:333, color:'#3498db',alpha:0.25,hideLabel:true},
    {id:genId(),name:'Zone 2 — Central', startCh:334,endCh:666, color:'#27ae60',alpha:0.25,hideLabel:true},
    {id:genId(),name:'Zone 3 — South',   startCh:667,endCh:1000,color:'#e67e22',alpha:0.25,hideLabel:true},
  ];
  dateMarkers=[
    {id:genId(),name:'Notice to Proceed',    date:`${y}-01-15`,   color:'#2c3e50'},
    {id:genId(),name:'IFC Drawings Issued',  date:`${y}-02-28`,   color:'#8e44ad'},
    {id:genId(),name:'Mid-Programme Review', date:`${y}-06-30`,   color:'#e67e22'},
    {id:genId(),name:'Contract Completion', date:`${y+1}-10-31`, color:'#27ae60'},
  ];

  const gP=genId(),gE=genId(),gD=genId(),gS=genId(),gV=genId();
  const sZ1=genId(),sZ2=genId(),sZ3=genId();
  const sDZ1=genId(),sDZ2=genId(),sDZ3=genId();
  const sC=genId(),sRW=genId();
  const sPZ1=genId(),sPZ2=genId(),sPZ3=genId();
  groups=[
    {id:gP,name:'Preparatory',   parentId:null,color:'#8e44ad',visible:true},
    {id:gE,name:'Excavation',    parentId:null,color:'#c0392b',visible:true},
    {id:gD,name:'Drainage',      parentId:null,color:'#2980b9',visible:true},
    {id:gS,name:'Structures',    parentId:null,color:'#16a085',visible:true},
    {id:gV,name:'Pavement',      parentId:null,color:'#d35400',visible:true},
    {id:sZ1,name:'Zone 1',parentId:gE,color:'#e74c3c',visible:true},
    {id:sZ2,name:'Zone 2',parentId:gE,color:'#c0392b',visible:true},
    {id:sZ3,name:'Zone 3',parentId:gE,color:'#a93226',visible:true},
    {id:sDZ1,name:'Zone 1',parentId:gD,color:'#5dade2',visible:true},
    {id:sDZ2,name:'Zone 2',parentId:gD,color:'#2980b9',visible:true},
    {id:sDZ3,name:'Zone 3',parentId:gD,color:'#1a5276',visible:true},
    {id:sC, name:'Culverts',       parentId:gS,color:'#1abc9c',visible:true},
    {id:sRW,name:'Retaining Walls',parentId:gS,color:'#148f77',visible:true},
    {id:sPZ1,name:'Zone 1',parentId:gV,color:'#f39c12',visible:true},
    {id:sPZ2,name:'Zone 2',parentId:gV,color:'#d68910',visible:true},
    {id:sPZ3,name:'Zone 3',parentId:gV,color:'#b7770d',visible:true},
  ];

  activities=[
    _sAct('P1000','Site Clearance',d(0),d(1,28),0,1000,'#8e44ad',0.85,'rect','no','Mobilisation',[gP],'none','#000',1.5,8,280,500,0),
    _sAct('P1010','Survey & Setting Out',d(0),d(1,15),0,1000,'#9b59b6',0.75,'line','no','Survey control',[gP],'none','#000',1.5,4,400,0,0),
    _sAct('E1000','Excavation — Zone 1',d(1),d(4,30),0,333,'#e74c3c',0.9,'rect','no','Cut to formation',[gE,sZ1],'none','#000',1.5,20,380,4200,0),
    _sAct('E1010','Excavation — Zone 2',d(3),d(6,30),334,666,'#c0392b',0.9,'rect','no','Cut to formation',[gE,sZ2],'none','#000',1.5,20,380,4200,0),
    _sAct('E1020','Excavation — Zone 3',d(5),d(8,30),667,1000,'#a93226',0.9,'rect','no','Cut to formation',[gE,sZ3],'none','#000',1.5,20,380,4200,0),
    _sAct('D1000','Drainage — Zone 1',d(3),d(6,30),0,333,'#2980b9',0.9,'rect','no','Surface water',[gD,sDZ1],'none','#000',1.5,12,350,900,48000),
    _sAct('D1010','Drainage — Zone 2',d(5),d(8,30),334,666,'#2980b9',0.9,'rect','no','Surface water',[gD,sDZ2],'none','#000',1.5,12,350,900,52000),
    _sAct('D1020','Drainage — Zone 3',d(7),d(10,30),667,1000,'#2980b9',0.9,'rect','no','Surface water',[gD,sDZ3],'none','#000',1.5,12,350,900,44000),
    _sAct('S1000','Culvert — Ch 180',d(3,15),d(5,30),160,200,'#1abc9c',0.9,'rect','no','Box culvert',[gS,sC],'dashed','#0e8a7a',1.5,8,600,2000,120000),
    _sAct('S1010','Culvert — Ch 550',d(6),d(8,30),530,570,'#1abc9c',0.9,'rect','no','Box culvert',[gS,sC],'dashed','#0e8a7a',1.5,8,600,2000,95000),
    _sAct('S1020','Retaining Wall',d(7),d(10,30),400,500,'#148f77',0.9,'rect','no','RC wall',[gS,sRW],'dashed','#0e6655',1.5,10,550,1500,200000),
    _sAct('V1000','Sub-base — Zone 1',d(5),d(7,31),0,333,'#f39c12',0.9,'rect','no','Type 1 & blinding',[gV,sPZ1],'none','#000',1.5,15,320,1800,85000),
    _sAct('V1010','Sub-base — Zone 2',d(7),d(9,30),334,666,'#d68910',0.9,'rect','no','Type 1 & blinding',[gV,sPZ2],'none','#000',1.5,15,320,1800,90000),
    _sAct('V1020','Sub-base — Zone 3',d(9),d(11,30),667,1000,'#b7770d',0.9,'rect','no','Type 1 & blinding',[gV,sPZ3],'none','#000',1.5,15,320,1800,78000),
    _sAct('V2000','FRC Slab — Zone 1',d(6,15),d(8,31),0,333,'#e67e22',0.9,'rect','no','FRC pour',[gV,sPZ1],'solid','#ca6f1e',2,18,420,2500,220000),
    _sAct('V2010','FRC Slab — Zone 2',d(8,15),d(10,31),334,666,'#e67e22',0.9,'rect','no','FRC pour',[gV,sPZ2],'solid','#ca6f1e',2,18,420,2500,240000),
    _sAct('V2020','FRC Slab — Zone 3',d(10,15),new Date(y+1,0,31),667,1000,'#e67e22',0.9,'rect','no','FRC pour',[gV,sPZ3],'solid','#ca6f1e',2,18,420,2500,195000),
    _sAct('V3000','Surfacing — Full Corridor',new Date(y+1,1,1),new Date(y+1,5,30),0,1000,'#d35400',0.9,'rect','no','Wearing course',[gV],'none','#000',1.5,22,390,3000,310000),
    _sAct('MS100','Start on Site',d(0,15),d(0,15),0,0,'#2c3e50',1,'circle','right','NTP',[gP],'none','#000',1.5,0,0,0,0),
    _sAct('MS110','Zone 1 Exc Complete',d(4,30),d(4,30),167,167,'#e74c3c',1,'diamond','right','',[gE,sZ1],'none','#000',1.5,0,0,0,0),
    _sAct('MS120','Zone 2 Exc Complete',d(6,30),d(6,30),500,500,'#c0392b',1,'diamond','right','',[gE,sZ2],'none','#000',1.5,0,0,0,0),
    _sAct('MS130','Zone 3 Exc Complete',d(8,30),d(8,30),834,834,'#a93226',1,'diamond','right','',[gE,sZ3],'none','#000',1.5,0,0,0,0),
    _sAct('MS200','Drainage Complete',d(10,30),d(10,30),500,500,'#2980b9',1,'star','right','All drainage signed off',[gD],'none','#000',1.5,0,0,0,0),
    _sAct('MS300','Structures Complete',new Date(y+1,0,31),new Date(y+1,0,31),420,420,'#16a085',1,'star','right','',[gS],'none','#000',1.5,0,0,0,0),
    _sAct('MS400','FRC Complete',new Date(y+1,0,31),new Date(y+1,0,31),580,580,'#e67e22',1,'star','right','',[gV],'none','#000',1.5,0,0,0,0),
    _sAct('MS500','Road Open',new Date(y+1,6,31),new Date(y+1,6,31),500,500,'#f39c12',1,'flag','right','',[gV],'none','#000',1.5,0,0,0,0),
  ];

  activityLinks=[
    _sLnk(0,2),_sLnk(2,3),_sLnk(3,4),
    _sLnk(2,5),_sLnk(3,6),_sLnk(4,7),
    _sLnk(2,8),_sLnk(3,9),_sLnk(3,10),
    _sLnk(5,11),_sLnk(6,12),_sLnk(7,13),
    _sLnk(11,14),_sLnk(12,15),_sLnk(13,16),
    _sLnk(14,17),_sLnk(15,17),_sLnk(16,17),
    _sLnk(0,18),_sLnk(2,19),_sLnk(3,20),_sLnk(4,21),
    _sLnk(7,22),_sLnk(10,23),_sLnk(16,24),
    _sLnk(17,25),_sLnk(25,26),
  ];

  risks=[
    _sRisk('E1000','Unexpected hard rock in Zone 1 — extends excavation duration',4,4,21,85000,'Site Engineer','Pre-excavation trial pits and GPR survey completed. Rock breaker on standby.','Open'),
    _sRisk('E1010','Contaminated ground discovered in Zone 2',3,5,35,220000,'Environmental Manager','Phase II ground investigation commissioned. Contingency disposal route agreed.','Open'),
    _sRisk('D1000','Utility clash — existing sewer not on records',4,3,14,45000,'Design Manager','All utility records re-requested. CAT & Genny sweep prior to any excavation.','Open'),
    _sRisk('V2000','FRC mix design approval delayed by client',3,4,10,0,'Commercial Manager','Early submission of mix design. Agreed fast-track 5-day review with client.','Mitigated'),
    _sRisk('V3000','Surfacing plant availability — only one approved supplier',3,3,7,30000,'Procurement','Alternative supplier being assessed. Early order placed.','Open'),
    _sRisk('S1020','Retaining wall design change mid-construction',2,5,28,180000,'Design Manager','Design freeze confirmed. Change control process implemented.','Open'),
    _sRisk('MS900','Practical completion delayed due to outstanding snags',3,3,14,0,'Project Manager','Weekly snag review meetings started 3 months before PC.','Open'),
    _sRisk('P1000','Key personnel unavailability at mobilisation',2,3,7,25000,'HR','Two senior engineers identified as backup. Retention bonuses agreed.','Closed'),
  ];

  _sFinish('✅ Sample 1 loaded — Highway Improvement Scheme');
}

// ═══════════════════════════════════════════════════════════════════
//  SAMPLE 2 — Multi-Zone Critical Path Demo
// ═══════════════════════════════════════════════════════════════════
function _loadSample2(){
  const y=new Date().getFullYear();
  _sSetup('Multi-Zone Critical Path Demo','Drainage & Foundations Programme',y,1000,50);
  const d=(m,dy=1)=>new Date(y,m,dy);

  const ZONES=[
    {name:'Zone A',ch0:0,  ch1:199,color:'#e74c3c'},
    {name:'Zone B',ch0:200,ch1:399,color:'#e67e22'},
    {name:'Zone C',ch0:400,ch1:599,color:'#2ecc71'},
    {name:'Zone D',ch0:600,ch1:799,color:'#3498db'},
    {name:'Zone E',ch0:800,ch1:1000,color:'#9b59b6'},
  ];
  chainageZones=ZONES.map(z=>({id:genId(),name:z.name,startCh:z.ch0,endCh:z.ch1,color:z.color,alpha:0.25,hideLabel:true}));
  dateMarkers=[
    {id:genId(),name:'Start on Site', date:`${y}-01-15`,color:'#2c3e50'},
    {id:genId(),name:'Programme Review',date:`${y}-07-01`,color:'#8e44ad'},
    {id:genId(),name:'Contract Completion',date:`${y+1}-06-30`,color:'#27ae60'},
  ];

  const cats=ZONES.map(()=>genId());
  const ddG=ZONES.map(()=>genId());
  const pcG=ZONES.map(()=>genId());
  const sdG=ZONES.map(()=>genId());
  groups=[
    ...ZONES.map((z,i)=>({id:cats[i],name:z.name,parentId:null,color:z.color,visible:true})),
    ...ZONES.map((z,i)=>({id:ddG[i], name:'Deep Drainage',    parentId:cats[i],color:z.color,visible:true})),
    ...ZONES.map((z,i)=>({id:pcG[i], name:'Pile Caps',         parentId:cats[i],color:z.color,visible:true})),
    ...ZONES.map((z,i)=>({id:sdG[i], name:'Shallow Drainage',  parentId:cats[i],color:z.color,visible:true})),
  ];

  activities=[];activityLinks=[];
  const W=7;
  const ZC=[
    {off:0,dd:[6,4,5],pc:[8,5],sd:[6,4],cpPc:0},
    {off:1,dd:[5,7,3],pc:[6,9],sd:[5,6],cpPc:1},
    {off:2,dd:[4,3,8],pc:[10,6],sd:[7,5],cpPc:0},
    {off:3,dd:[7,5,4],pc:[7,4],sd:[4,8],cpPc:0},
    {off:4,dd:[3,6,7],pc:[9,5],sd:[6,3],cpPc:0},
  ];

  ZONES.forEach((z,zi)=>{
    const cfg=ZC[zi];
    const off=cfg.off;
    const zMid=Math.round((z.ch0+z.ch1)/2);
    const nDD=cfg.dd.length,nPC=cfg.pc.length,nSD=cfg.sd.length;
    const catchW=Math.floor((z.ch1-z.ch0)/nDD);
    const pcW=Math.floor((z.ch1-z.ch0)/nPC);
    const sdW=Math.floor((z.ch1-z.ch0)/nSD);
    const baseIdx=activities.length;
    let ddStart=d(off,15);
    for(let c=0;c<nDD;c++){
      const dur=cfg.dd[c]*W;
      const cEnd=new Date(ddStart.getTime()+dur*86400000);
      activities.push(_sAct(`DD-Z${zi+1}-C${c+1}`,`${z.name} Deep Drain C${zi*10+21+c}`,new Date(ddStart),cEnd,z.ch0+c*catchW,z.ch0+(c+1)*catchW-1,z.color,0.9,'line','no',`${cfg.dd[c]}w`,[cats[zi],ddG[zi]],'none','#000',1.5,8,380,1200,35000));
      ddStart=new Date(ddStart);
    }
    const lastDDEnd=activities[baseIdx+nDD-1].end;
    for(let g=0;g<nPC;g++){
      const dur=cfg.pc[g]*W;
      const pStart=new Date(lastDDEnd);
      activities.push(_sAct(`PC-Z${zi+1}-G${g+1}`,`${z.name} Pile Cap G${g+1}`,pStart,new Date(pStart.getTime()+dur*86400000),z.ch0+g*pcW,z.ch0+(g+1)*pcW-1,z.color,0.9,'line','no',`${cfg.pc[g]}w`,[cats[zi],pcG[zi]],'dashed',z.color,1.5,10,450,1800,60000));
    }
    const critPcIdx=baseIdx+nDD+cfg.cpPc;
    let sdStart=new Date(activities[critPcIdx].end);
    for(let g=0;g<nSD;g++){
      const dur=cfg.sd[g]*W;
      const sEnd=new Date(sdStart.getTime()+dur*86400000);
      activities.push(_sAct(`SD-Z${zi+1}-G${g+1}`,`${z.name} Shallow Drain G${g+1}`,new Date(sdStart),sEnd,z.ch0+g*sdW,z.ch0+(g+1)*sdW-1,z.color,0.9,'line','no',`${cfg.sd[g]}w`,[cats[zi],sdG[zi]],'none','#000',1.5,8,320,800,28000));
      sdStart=new Date(sEnd);
    }
    const msDate=new Date(sdStart.getTime()+7*86400000);
    activities.push(_sAct(`MS-Z${zi+1}`,`${z.name} Complete`,msDate,msDate,zMid,zMid,z.color,1,'star','right',`${z.name} complete`,[cats[zi]],'none','#000',1.5,0,0,0,0));
    const ddEnd=baseIdx+nDD,pcEnd=ddEnd+nPC,sdEnd=pcEnd+nSD,msIdx=sdEnd;
    for(let c=0;c<nDD-1;c++) activityLinks.push(_sLnk(baseIdx+c,baseIdx+c+1));
    for(let g=0;g<nPC;g++) activityLinks.push(_sLnk(baseIdx+nDD-1,ddEnd+g));
    activityLinks.push(_sLnk(critPcIdx,pcEnd));
    for(let g=0;g<nPC;g++) if(g!==cfg.cpPc) activityLinks.push(_sLnk(ddEnd+g,pcEnd));
    for(let g=0;g<nSD-1;g++) activityLinks.push(_sLnk(pcEnd+g,pcEnd+g+1));
    activityLinks.push(_sLnk(sdEnd-1,msIdx));
  });

  risks=[
    _sRisk('DD-Z1-C1','Zone A — high groundwater table affecting deep drain excavation',4,4,14,55000,'Geotechnical Engineer','Dewatering pumps pre-ordered. Standby pump available on site.','Open'),
    _sRisk('PC-Z1-G1','Zone A — pile cap reinforcement delivery delay (supply chain)',3,4,10,0,'Procurement','Secondary supplier identified. Early procurement order placed.','Open'),
    _sRisk('DD-Z2-C1','Zone B — utilities conflict with deep drainage alignment',4,3,7,30000,'Design Manager','Utility diversion design completed. Awaiting statutory authority approval.','Open'),
    _sRisk('PC-Z3-G1','Zone C — contaminated spoil from pile cap excavation',3,5,21,120000,'Environmental Manager','Soil testing underway. Licensed disposal route secured.','Open'),
    _sRisk('DD-Z4-C1','Zone D — restricted working hours due to noise complaints',3,3,10,25000,'Site Manager','Acoustic barrier installed. Night shift application submitted.','Mitigated'),
    _sRisk('MS-Z5','Zone E — milestone delay due to third-party approval backlog',3,4,21,0,'Project Manager','Pre-application meetings held with approving authority.','Open'),
  ];

  _sFinish('✅ Sample 2 loaded — Multi-Zone Critical Path Demo');
}

// ═══════════════════════════════════════════════════════════════════
//  SAMPLE 3 — Subcontractor Programme
// ═══════════════════════════════════════════════════════════════════
function _loadSample3(){
  const y=new Date().getFullYear();
  _sSetup('Bridge & Highway Works','Subcontractor Programme — Sequential Handover',y,1000,50);
  const d=(m,dy=1)=>new Date(y,m,dy);

  chainageZones=[
    {id:genId(),name:"Blackwells Zone",startCh:0,  endCh:249, color:'#e67e22',alpha:0.25,hideLabel:true},
    {id:genId(),name:"EKFB Zone",      startCh:250,endCh:499, color:'#e74c3c',alpha:0.25,hideLabel:true},
    {id:genId(),name:"William Hare",   startCh:500,endCh:749, color:'#8e44ad',alpha:0.25,hideLabel:true},
    {id:genId(),name:"Clancy Zone",    startCh:750,endCh:1000,color:'#2980b9',alpha:0.25,hideLabel:true},
  ];
  dateMarkers=[
    {id:genId(),name:'Start on Site',       date:`${y}-01-15`,   color:'#2c3e50'},
    {id:genId(),name:'EKFB Possession',     date:`${y}-04-01`,   color:'#e74c3c'},
    {id:genId(),name:'Steel Delivery',      date:`${y}-09-01`,   color:'#8e44ad'},
    {id:genId(),name:'Clancy Mobilise',     date:`${y}-11-01`,   color:'#2980b9'},
    {id:genId(),name:'Contract Completion',date:`${y+1}-09-30`, color:'#27ae60'},
  ];

  const gBW=genId(),gEK=genId(),gWH=genId(),gCL=genId();
  const sBW_EW=genId(),sBW_CF=genId(),sEK_BL=genId(),sEK_FD=genId(),sEK_WW=genId(),sEK_AB=genId();
  const sWH_B1=genId(),sWH_B2=genId(),sCL_DR=genId(),sCL_EJ=genId();
  groups=[
    {id:gBW, name:'Blackwells',   parentId:null,color:'#e67e22',visible:true},
    {id:gEK, name:'EKFB',         parentId:null,color:'#e74c3c',visible:true},
    {id:gWH, name:'William Hare', parentId:null,color:'#8e44ad',visible:true},
    {id:gCL, name:'Clancy',       parentId:null,color:'#2980b9',visible:true},
    {id:sBW_EW,name:'Earthworks', parentId:gBW,color:'#d35400',visible:true},
    {id:sBW_CF,name:'Cut & Fill', parentId:gBW,color:'#e67e22',visible:true},
    {id:sEK_BL,name:'Blinding',   parentId:gEK,color:'#c0392b',visible:true},
    {id:sEK_FD,name:'Foundations',parentId:gEK,color:'#e74c3c',visible:true},
    {id:sEK_WW,name:'Wingwalls',  parentId:gEK,color:'#f1948a',visible:true},
    {id:sEK_AB,name:'Abutments',  parentId:gEK,color:'#922b21',visible:true},
    {id:sWH_B1,name:'Bridge 1',   parentId:gWH,color:'#8e44ad',visible:true},
    {id:sWH_B2,name:'Bridge 2',   parentId:gWH,color:'#6c3483',visible:true},
    {id:sCL_DR,name:'Drainage',   parentId:gCL,color:'#2980b9',visible:true},
    {id:sCL_EJ,name:'Exp Joints', parentId:gCL,color:'#1a5276',visible:true},
  ];

  activities=[
    _sAct('BW-1000','Earthworks — Cut (North)',d(0),d(3,31),0,120,'#d35400',0.9,'rect','no','Bulk cut',[gBW,sBW_EW],'none','#000',1.5,20,350,5000,0),
    _sAct('BW-1010','Earthworks — Cut (South)',d(2),d(5,31),130,249,'#e67e22',0.9,'rect','no','Bulk cut',[gBW,sBW_EW],'none','#000',1.5,20,350,5000,0),
    _sAct('BW-1020','Cut & Fill — North',d(3),d(6,30),0,120,'#f0a500',0.9,'rect','no','Fill to formation',[gBW,sBW_CF],'none','#000',1.5,15,300,3000,0),
    _sAct('BW-1030','Cut & Fill — South',d(5),d(7,31),130,249,'#d35400',0.9,'rect','no','Fill to formation',[gBW,sBW_CF],'none','#000',1.5,15,300,3000,0),
    _sAct('BW-MS','Earthworks Complete',d(7,31),d(7,31),125,125,'#e67e22',1,'diamond','right','BW handover',[gBW],'none','#000',1.5,0,0,0,0),
    _sAct('EK-1000','Blinding — Bridge 1',d(3),d(4,30),260,374,'#c0392b',0.9,'rect','no','50mm blinding',[gEK,sEK_BL],'none','#000',1.5,6,550,800,42000),
    _sAct('EK-1010','Blinding — Bridge 2',d(4),d(5,31),376,490,'#c0392b',0.9,'rect','no','50mm blinding',[gEK,sEK_BL],'none','#000',1.5,6,550,800,40000),
    _sAct('EK-2000','Foundations — Bridge 1',d(4,15),d(7,14),260,374,'#e74c3c',0.9,'rect','no','Pile caps & ground beams',[gEK,sEK_FD],'solid','#c0392b',2,12,580,1500,260000),
    _sAct('EK-2010','Foundations — Bridge 2',d(5,15),d(8,14),376,490,'#e74c3c',0.9,'rect','no','Pile caps & ground beams',[gEK,sEK_FD],'solid','#c0392b',2,12,580,1500,245000),
    _sAct('EK-3000','Wingwalls — Bridge 1',d(7),d(8,31),260,310,'#f1948a',0.9,'rect','no','RC wingwalls',[gEK,sEK_WW],'none','#000',1.5,8,520,900,88000),
    _sAct('EK-3010','Wingwalls — Bridge 2',d(8),d(9,30),376,426,'#f1948a',0.9,'rect','no','RC wingwalls',[gEK,sEK_WW],'none','#000',1.5,8,520,900,82000),
    _sAct('EK-4000','Abutments — Bridge 1',d(7),d(9,14),315,374,'#922b21',0.9,'rect','no','RC abutments',[gEK,sEK_AB],'none','#000',1.5,8,520,900,105000),
    _sAct('EK-4010','Abutments — Bridge 2',d(8),d(10,14),431,490,'#922b21',0.9,'rect','no','RC abutments',[gEK,sEK_AB],'none','#000',1.5,8,520,900,98000),
    _sAct('EK-MS','EKFB Works Complete',d(10,14),d(10,14),375,375,'#e74c3c',1,'diamond','right','Handover to WH',[gEK],'none','#000',1.5,0,0,0,0),
    _sAct('WH-1000','Steel Erection — Bridge 1',d(9),d(11,30),505,624,'#8e44ad',0.9,'rect','no','Truss lift & deck',[gWH,sWH_B1],'solid','#6c3483',2.5,14,650,4500,560000),
    _sAct('WH-2000','Steel Erection — Bridge 2',new Date(y+1,0,1),new Date(y+1,2,28),630,745,'#6c3483',0.9,'rect','no','Truss lift & deck',[gWH,sWH_B2],'solid','#4a235a',2.5,14,650,4500,520000),
    _sAct('WH-MS','Steel Erection Complete',new Date(y+1,2,28),new Date(y+1,2,28),625,625,'#8e44ad',1,'star','right','Both bridges erected',[gWH],'none','#000',1.5,0,0,0,0),
    _sAct('CL-1000','Drainage — North Section',d(10),new Date(y+1,1,28),750,874,'#2980b9',0.9,'rect','no','Highway drainage',[gCL,sCL_DR],'none','#000',1.5,12,340,900,115000),
    _sAct('CL-1010','Drainage — South Section',new Date(y+1,0,1),new Date(y+1,3,30),876,1000,'#5dade2',0.9,'rect','no','Highway drainage',[gCL,sCL_DR],'none','#000',1.5,12,340,900,108000),
    _sAct('CL-2000','Expansion Joints — Bridge 1',new Date(y+1,2,1),new Date(y+1,4,30),755,800,'#1a5276',0.9,'rect','no','Movement joints',[gCL,sCL_EJ],'dashed','#154360',1.5,6,480,600,46000),
    _sAct('CL-2010','Expansion Joints — Bridge 2',new Date(y+1,3,1),new Date(y+1,5,30),830,874,'#1a5276',0.9,'rect','no','Movement joints',[gCL,sCL_EJ],'dashed','#154360',1.5,6,480,600,43000),
    // 20 — no separate completion flag, date marker handles it
  ];

  activityLinks=[
    _sLnk(0,2),_sLnk(1,3),_sLnk(2,4),_sLnk(3,4),
    _sLnk(4,5),_sLnk(4,6),
    _sLnk(5,7),_sLnk(6,8),
    _sLnk(7,9),_sLnk(7,11),_sLnk(8,10),_sLnk(8,12),
    _sLnk(9,13),_sLnk(10,13),_sLnk(11,13),_sLnk(12,13),
    _sLnk(13,14),_sLnk(14,15),_sLnk(15,16),
    _sLnk(14,19),_sLnk(15,19),
    _sLnk(16,17),_sLnk(17,18),_sLnk(18,19),
  ];

  risks=[
    _sRisk('BW-1000','Unstable embankment slope — risk of collapse during bulk cut',4,5,28,200000,'Geotechnical Engineer','Slope stability analysis commissioned. Temporary propping design approved.','Open'),
    _sRisk('EK-2000','Concrete pour delayed by cold weather — FRC minimum temperature requirements',3,4,14,60000,'Site Manager','Heated enclosures on standby. Weather window monitoring system in place.','Open'),
    _sRisk('EK-2010','Piling rig breakdown — single rig on site, no immediate replacement',3,4,21,95000,'Plant Manager','Maintenance schedule brought forward. Backup rig sourced from hire company.','Open'),
    _sRisk('WH-1000','Steel fabrication error discovered on delivery — wrong camber on truss',2,5,35,320000,'Structural Engineer','Enhanced factory inspection regime agreed. QA hold points at fabrication stage.','Open'),
    _sRisk('WH-2000','Crane lift window lost due to high winds — bridge over live road',4,3,7,40000,'Lifting Supervisor','Wind monitoring system installed. Contingency lift windows identified.','Open'),
    _sRisk('CL-1000','Highway drainage alignment clashes with existing telecoms duct',3,3,10,35000,'Design Manager','BT/Openreach design liaison meetings ongoing. Diversion route agreed in principle.','Mitigated'),
    _sRisk('PC-MS','Third party inspections delaying PC sign-off',3,3,14,0,'Project Manager','Inspection schedule agreed with client 8 weeks in advance.','Open'),
  ];

  _sFinish('✅ Sample 3 loaded — Subcontractor Programme');
}

// ═══════════════════════════════════════════════════════════════════
//  SAMPLE 4 — Baseline Comparison
// ═══════════════════════════════════════════════════════════════════
function _loadSample4(){
  const y=2026; // Fixed year so dates match the spec exactly
  _sSetup('Urban Road Widening — Rev B','Baseline vs Current Programme (Delayed)',y,1000,50);

  // Timeline: Sept 2026 to end of 2026, plus a bit of 2027 for the baseline ghost
  timelineStart=new Date(2025,11,1); timelineEnd=new Date(2027,2,31);
  document.getElementById('timelineStart').value='2025-12-01';
  document.getElementById('timelineEnd').value='2027-03-31';

  chainageZones=[
    {id:genId(),name:'Zone A',startCh:1,  endCh:500, color:'#3498db',alpha:0.25,hideLabel:true},
    {id:genId(),name:'Zone B',startCh:501,endCh:1000,color:'#27ae60',alpha:0.25,hideLabel:true},
  ];
  dateMarkers=[
    {id:genId(),name:'Contract Completion',date:'2026-12-31',color:'#27ae60'},
  ];

  const gA=genId(),gB=genId();
  groups=[
    {id:gA,name:'Zone A',parentId:null,color:'#3498db',visible:true},
    {id:gB,name:'Zone B',parentId:null,color:'#27ae60',visible:true},
  ];

  // ── CURRENT (delayed) programme ──
  // Zone A: rect 01/09/26 → 01/10/26, line 02/10/26 → 01/11/26
  // Zone B: line 01/11/26 → 01/12/26, rect 02/12/26 → 20/12/26
  activities=[
    // 0 — Zone A rectangle (current)
    _sAct('ZA-R','Zone A — Construction Works',new Date(2026,8,1),new Date(2026,9,1),1,500,'#3498db',0.9,'rect','no','Current programme',[gA],'none','#000',1.5,15,380,2000,180000),
    // 1 — Zone A line (current)
    _sAct('ZA-L','Zone A — Finishing Works',new Date(2026,9,2),new Date(2026,10,1),1,500,'#2980b9',0.9,'line','no','Current programme',[gA],'none','#000',1.5,8,320,0,0),
    // 2 — Zone B line (current)
    _sAct('ZB-L','Zone B — Preparatory Works',new Date(2026,10,1),new Date(2026,11,1),501,1000,'#27ae60',0.9,'line','no','Current programme',[gB],'none','#000',1.5,8,320,0,0),
    // 3 — Zone B rectangle (current)
    _sAct('ZB-R','Zone B — Construction Works',new Date(2026,11,2),new Date(2026,11,20),501,1000,'#219a52',0.9,'rect','no','Current programme',[gB],'none','#000',1.5,15,380,2000,165000),
  ];

  activityLinks=[
    _sLnk(0,1), // ZA rect → ZA line
    _sLnk(1,2), // ZA line → ZB line
    _sLnk(2,3), // ZB line → ZB rect
  ];

  // ── BASELINE — original planned dates (earlier than current) ──
  // Zone A baseline: rect 01/01/26 → 01/02/26, line 02/02/26 → 01/03/26
  // Zone B baseline: line 01/03/26 → 01/04/26, rect 02/04/26 → 20/04/26
  baselines=[{
    id:genId(),
    name:'Rev A — Original Baseline',
    date:new Date().toISOString(),
    visible:true,
    activities:[
      {p6Id:'ZA-R',start:new Date(2026,0,1), end:new Date(2026,1,1), startCh:1,   endCh:500, shape:'rect',color:'#95a5a6',alpha:0.6},
      {p6Id:'ZA-L',start:new Date(2026,1,2), end:new Date(2026,2,1), startCh:1,   endCh:500, shape:'line',color:'#95a5a6',alpha:0.6},
      {p6Id:'ZB-L',start:new Date(2026,2,1), end:new Date(2026,3,1), startCh:501, endCh:1000,shape:'line',color:'#95a5a6',alpha:0.6},
      {p6Id:'ZB-R',start:new Date(2026,3,2), end:new Date(2026,3,20),startCh:501, endCh:1000,shape:'rect',color:'#95a5a6',alpha:0.6},
    ],
  }];

  risks=[
    _sRisk('ZA-R','Zone A delayed 8 months — unexpected utilities clash during site investigation',5,5,244,350000,'Design Manager','Utility diversions agreed and programme updated. EOT claim submitted to client.','Open'),
    _sRisk('ZB-R','Zone B at risk of further delay — ground conditions at Ch 700 require additional investigation',3,4,21,60000,'Geotechnical Engineer','Intrusive ground investigation commissioned. Results expected in 3 weeks.','Open'),
    _sRisk('ZA-L','Finishing works productivity lower than planned — smaller gang than tendered',3,3,10,15000,'Site Manager','Additional resource request submitted. Sub-contractor on standby.','Open'),
    _sRisk('ZB-R','Material supply — precast units lead time increased to 12 weeks',3,4,14,0,'Procurement','Early order placed. Confirmed delivery date from supplier.','Mitigated'),
  ];

  _sFinish('✅ Sample 4 loaded — Baseline Comparison. Enable "Show Baselines" in the chart toolbar to see the ghost overlay.');
}

function _loadSample5(){
  const y=new Date().getFullYear();
  _sSetup('East Coast Mainline — Track Renewal','Rail Possession Programme',y,120,5);
  const d=(m,dy=1)=>new Date(y,m,dy);

  // Track sections as chainage zones
  chainageZones=[
    {id:genId(),name:'Line 1 (Up Fast)',  startCh:0,  endCh:29,  color:'#e74c3c',alpha:0.25,hideLabel:true},
    {id:genId(),name:'Line 2 (Up Slow)',  startCh:30, endCh:59,  color:'#e67e22',alpha:0.25,hideLabel:true},
    {id:genId(),name:'Line 3 (Down Slow)',startCh:60, endCh:89,  color:'#3498db',alpha:0.25,hideLabel:true},
    {id:genId(),name:'Line 4 (Down Fast)',startCh:90, endCh:120, color:'#27ae60',alpha:0.25,hideLabel:true},
  ];

  // Date markers for key operational constraints
  dateMarkers=[
    {id:genId(),name:'Summer Embargo (no poss.)',date:`${y}-08-01`,color:'#e74c3c'},
    {id:genId(),name:'Embargo Lifts',            date:`${y}-10-01`,color:'#27ae60'},
    {id:genId(),name:'Timetable Change',         date:`${y}-12-14`,color:'#8e44ad'},
    {id:genId(),name:'Contract Completion',      date:`${y+1}-03-31`,color:'#2c3e50'},
  ];

  const gPoss=genId(),gRail=genId(),gCivil=genId(),gSignal=genId();
  const sPL1=genId(),sPL2=genId(),sPL3=genId(),sPL4=genId();
  const sRT=genId(),sRC=genId();
  const sDC=genId(),sDR=genId();
  groups=[
    {id:gPoss,  name:'Possessions',     parentId:null,color:'#e74c3c',visible:true},
    {id:gRail,  name:'Track Renewal',   parentId:null,color:'#2980b9',visible:true},
    {id:gCivil, name:'Civil Works',     parentId:null,color:'#e67e22',visible:true},
    {id:gSignal,name:'Signalling',      parentId:null,color:'#8e44ad',visible:true},
    {id:sPL1,name:'Line 1 Poss',parentId:gPoss,color:'#e74c3c',visible:true},
    {id:sPL2,name:'Line 2 Poss',parentId:gPoss,color:'#c0392b',visible:true},
    {id:sPL3,name:'Line 3 Poss',parentId:gPoss,color:'#e67e22',visible:true},
    {id:sPL4,name:'Line 4 Poss',parentId:gPoss,color:'#d35400',visible:true},
    {id:sRT, name:'Plain Line Track',parentId:gRail,color:'#2980b9',visible:true},
    {id:sRC, name:'Crossovers & Pts', parentId:gRail,color:'#1a5276',visible:true},
    {id:sDC, name:'Drainage & Culverts',parentId:gCivil,color:'#e67e22',visible:true},
    {id:sDR, name:'Drainage Renewal',  parentId:gCivil,color:'#d35400',visible:true},
  ];

  // Possession windows — wide bars spanning the track section chainage
  // Each possession is a wide rect at low opacity showing the blocked window
  // Works gangs appear as narrower bars within each possession
  activities=[
    // ── POSSESSION WINDOWS (wide, semi-transparent rects) ──
    // Line 1 possessions
    _sAct('PW-L1-01','POSSESSION — Line 1 (Up Fast) Block A',d(0,13),d(0,20),0,29,'#e74c3c',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL1],'solid','#c0392b',2,0,0,0,0),
    _sAct('PW-L1-02','POSSESSION — Line 1 (Up Fast) Block B',d(2,24),d(2,31),0,29,'#e74c3c',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL1],'solid','#c0392b',2,0,0,0,0),
    _sAct('PW-L1-03','POSSESSION — Line 1 (Up Fast) Block C',d(9,12),d(9,19),0,29,'#e74c3c',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL1],'solid','#c0392b',2,0,0,0,0),
    // Line 2 possessions
    _sAct('PW-L2-01','POSSESSION — Line 2 (Up Slow) Block A',d(1,10),d(1,17),30,59,'#c0392b',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL2],'solid','#922b21',2,0,0,0,0),
    _sAct('PW-L2-02','POSSESSION — Line 2 (Up Slow) Block B',d(3,21),d(3,28),30,59,'#c0392b',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL2],'solid','#922b21',2,0,0,0,0),
    _sAct('PW-L2-03','POSSESSION — Line 2 (Up Slow) Block C',d(10,5),d(10,12),30,59,'#c0392b',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL2],'solid','#922b21',2,0,0,0,0),
    // Line 3 possessions
    _sAct('PW-L3-01','POSSESSION — Line 3 (Down Slow) Block A',d(4,7),d(4,14),60,89,'#e67e22',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL3],'solid','#d35400',2,0,0,0,0),
    _sAct('PW-L3-02','POSSESSION — Line 3 (Down Slow) Block B',d(5,16),d(5,23),60,89,'#e67e22',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL3],'solid','#d35400',2,0,0,0,0),
    _sAct('PW-L3-03','POSSESSION — Line 3 (Down Slow) Block C',new Date(y+1,0,11),new Date(y+1,0,18),60,89,'#e67e22',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL3],'solid','#d35400',2,0,0,0,0),
    // Line 4 possessions
    _sAct('PW-L4-01','POSSESSION — Line 4 (Down Fast) Block A',d(5,28),d(6,4),90,120,'#d35400',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL4],'solid','#b7770d',2,0,0,0,0),
    _sAct('PW-L4-02','POSSESSION — Line 4 (Down Fast) Block B',d(10,19),d(10,26),90,120,'#d35400',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL4],'solid','#b7770d',2,0,0,0,0),
    _sAct('PW-L4-03','POSSESSION — Line 4 (Down Fast) Block C',new Date(y+1,1,8),new Date(y+1,1,15),90,120,'#d35400',0.45,'rect','no','STP Possession — 7 nights',[gPoss,sPL4],'solid','#b7770d',2,0,0,0,0),

    // ── TRACK RENEWAL GANGS (within each possession) ──
    _sAct('TR-L1-A','Track Renewal — L1 Ch 0–14',d(0,13),d(0,20),0,14,'#2980b9',0.95,'line','no','Plain line renewal — tamper & stone blow',[gRail,sRT],'none','#000',1.5,18,480,8500,145000),
    _sAct('TR-L1-B','Track Renewal — L1 Ch 15–29',d(2,24),d(2,31),15,29,'#2980b9',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,138000),
    _sAct('TR-L1-C','Track Renewal — L1 Ch 0–29 (complete)',d(9,12),d(9,19),0,29,'#1a5276',0.95,'line','no','Crossover renewal at Ch 12 & 24',[gRail,sRC],'dashed','#154360',2,12,580,6000,220000),
    _sAct('TR-L2-A','Track Renewal — L2 Ch 30–44',d(1,10),d(1,17),30,44,'#5dade2',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,132000),
    _sAct('TR-L2-B','Track Renewal — L2 Ch 45–59',d(3,21),d(3,28),45,59,'#5dade2',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,128000),
    _sAct('TR-L2-C','Points Renewal — L2 Ch 30–59',d(10,5),d(10,12),30,59,'#1a5276',0.95,'line','no','Diamond crossings & switches',[gRail,sRC],'dashed','#154360',2,12,580,6000,195000),
    _sAct('TR-L3-A','Track Renewal — L3 Ch 60–74',d(4,7),d(4,14),60,74,'#e67e22',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,140000),
    _sAct('TR-L3-B','Track Renewal — L3 Ch 75–89',d(5,16),d(5,23),75,89,'#e67e22',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,135000),
    _sAct('TR-L3-C','Points Renewal — L3 Ch 60–89',new Date(y+1,0,11),new Date(y+1,0,18),60,89,'#d35400',0.95,'line','no','Trailing crossovers',[gRail,sRC],'dashed','#b03a00',2,12,580,6000,180000),
    _sAct('TR-L4-A','Track Renewal — L4 Ch 90–104',d(5,28),d(6,4),90,104,'#27ae60',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,142000),
    _sAct('TR-L4-B','Track Renewal — L4 Ch 105–120',d(10,19),d(10,26),105,120,'#27ae60',0.95,'line','no','Plain line renewal',[gRail,sRT],'none','#000',1.5,18,480,8500,138000),
    _sAct('TR-L4-C','Points Renewal — L4 Ch 90–120',new Date(y+1,1,8),new Date(y+1,1,15),90,120,'#1e8449',0.95,'line','no','Buffer stop & crossovers',[gRail,sRC],'dashed','#145a32',2,12,580,6000,175000),

    // ── CIVIL & DRAINAGE (between possessions) ──
    _sAct('CW-01','Drainage Repair — Ch 5–20',d(1),d(3,31),5,20,'#e67e22',0.9,'rect','no','Culvert clearance & CCTV survey',[gCivil,sDC],'none','#000',1.5,8,280,400,28000),
    _sAct('CW-02','Drainage Renewal — Ch 45–60',d(4),d(6,30),45,60,'#d35400',0.9,'rect','no','New drainage channel',[gCivil,sDR],'none','#000',1.5,10,300,600,45000),
    _sAct('CW-03','Drainage Renewal — Ch 80–95',d(9),new Date(y+1,0,31),80,95,'#f0a500',0.9,'rect','no','New drainage channel',[gCivil,sDR],'none','#000',1.5,10,300,600,42000),

    // ── SIGNALLING ──
    _sAct('SIG-01','Signal Renewals — Ch 0–60',d(9),d(11,30),0,60,'#8e44ad',0.9,'line','no','AWS, TPWS & lineside signals',[gSignal],'dashed','#6c3483',1.5,6,620,1200,380000),
    _sAct('SIG-02','Signal Renewals — Ch 60–120',new Date(y+1,0,1),new Date(y+1,2,28),60,120,'#6c3483',0.9,'line','no','AWS, TPWS & lineside signals',[gSignal],'dashed','#4a235a',1.5,6,620,1200,360000),

    // ── MILESTONES ──
    _sAct('MS-L12','Lines 1&2 Renewal Complete',d(10,12),d(10,12),60,60,'#2980b9',1,'star','right','L1 & L2 handed back',[gRail],'none','#000',1.5,0,0,0,0),
    _sAct('MS-L34','Lines 3&4 Renewal Complete',new Date(y+1,1,15),new Date(y+1,1,15),60,60,'#27ae60',1,'star','right','L3 & L4 handed back',[gRail],'none','#000',1.5,0,0,0,0),
  ];

  // Key dependency links
  activityLinks=[
    // Track renewal follows possession windows
    _sLnk(0,12),_sLnk(1,13),_sLnk(2,14),
    _sLnk(3,15),_sLnk(4,16),_sLnk(5,17),
    _sLnk(6,18),_sLnk(7,19),_sLnk(8,20),
    _sLnk(9,21),_sLnk(10,22),_sLnk(11,23),
    // Line milestones
    _sLnk(14,30),_sLnk(17,30), // L1+L2 renewals → MS-L12
    _sLnk(20,31),_sLnk(23,31), // L3+L4 renewals → MS-L34
    // Signalling follows track
    _sLnk(14,28),_sLnk(17,28), // L1+L2 done → signal renewals L1
    _sLnk(20,29),_sLnk(23,29), // L3+L4 done → signal renewals L2
  ];

  risks=[
    _sRisk('PW-L1-01','Possession overrun — Line 1 Block A at risk if tamper unavailable',4,5,7,280000,'Possession Manager','Spare tamper allocated from Scotland depot. Overrun management plan in place.','Open'),
    _sRisk('TR-L1-A','Track geometry failure after renewal — failed acceptance test',3,5,7,95000,'Track Engineer','Enhanced QA regime. Two-pass tamping required. Acceptance criteria agreed upfront.','Open'),
    _sRisk('PW-L2-02','Possession conflict with emergency engineering works from another TOC',3,4,14,120000,'Network Rail Planner','Joint planning meetings in place. Conflict resolution protocol agreed.','Open'),
    _sRisk('CW-02','Drainage channel collapse during excavation — unstable formation',3,4,10,55000,'Civil Engineer','Ground investigation done. Sheet pile support specified.','Mitigated'),
    _sRisk('SIG-01','Signal cable route conflict with new drainage — not identified on records',4,3,7,40000,'Signalling Engineer','Full cable route check being carried out. Signalling design team engaged.','Open'),
    _sRisk('TR-L4-C','Points renewal scope increased — additional switch discovered on survey',3,3,7,85000,'Track Engineer','Survey completed. Additional materials ordered. Programme buffer sufficient.','Open'),
    _sRisk('MS-PC','Programme completion date at risk due to summer embargo restricting possessions',4,4,42,0,'Programme Manager','Possession applications submitted to avoid embargo. Contingency blocks in autumn.','Open'),
  ];

  _sFinish('✅ Sample 5 loaded — Rail Possession Programme');
}

function _sFinish(msg){
  renderGroupTree();renderGroupDropdown();renderGroupPickerInEditor();
  renderMarkerList();renderZoneList();renderBaselineList();
  drawChart();renderActivityTable();populateCpTargetSelect();
  flashSaved(msg);
  openTab2('settings');
}

function confirmLoadSample(){ openSamplePicker(); }
function loadSampleProject(){ _loadSample1(); }

function closeExplorer(){}
function openFromPC(){}
function closeMoveDialog(){document.getElementById('moveDialog').style.display='none';}
function confirmMove(){}
function openRenameDialog(id){document.getElementById('renameInput').value='';document.getElementById('renameDialog').style.display='flex';}
function confirmRename(){closeRenameDialog();}
function closeRenameDialog(){document.getElementById('renameDialog').style.display='none';}

// ═══════════════════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded',()=>{
  document.getElementById('renameInput')?.addEventListener('keydown',e=>{
    if(e.key==='Enter')confirmRename();if(e.key==='Escape')closeRenameDialog();
  });

  // Show onboarding on first load if no activities and no autosave
  try{
    const saved=localStorage.getItem('tcplanner_autosave');
    const hasSave=saved&&JSON.parse(saved)?.currentDiagram?.activities?.length>0;
    if(!hasSave&&!activities.length){
      setTimeout(()=>{document.getElementById('onboardingModal').style.display='flex';},400);
    }
  }catch(e){
    if(!activities.length) setTimeout(()=>{document.getElementById('onboardingModal').style.display='flex';},400);
  }
  window.addEventListener('beforeunload', e=>{
    if(activities.length){
      // Auto-save to localStorage as a safety net
      try{
        localStorage.setItem('tcplanner_autosave', JSON.stringify({
          version:2, currentDiagram:serialiseDiagram(), _savedAt: Date.now()
        }));
      }catch(err){}
      const msg='You have unsaved changes. Are you sure you want to leave? (Your work has been auto-saved in the browser.)';
      e.preventDefault();
      e.returnValue=msg;
      return msg;
    }
  });

  // Restore autosave if available and page looks empty
  try{
    const saved=localStorage.getItem('tcplanner_autosave');
    if(saved&&!activities.length){
      const data=JSON.parse(saved);
      const savedAt=data._savedAt?new Date(data._savedAt).toLocaleTimeString():'unknown time';
      const diag=data.currentDiagram||data;
      if(diag.activities&&diag.activities.length){
        const restore=confirm(`🔄 Auto-save detected from ${savedAt} with ${diag.activities.length} activities.\n\nRestore it?`);
        if(restore){ deserialiseDiagram(diag); flashSaved('✅ Auto-save restored'); }
      }
    }
  }catch(err){}
});

renderGroupTree();
renderGroupDropdown();
renderGroupPickerInEditor();
updateCurrentFileLabel();
drawChart();
