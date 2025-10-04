function el(id){ return document.getElementById(id) }
function num(id){ return parseFloat(el(id).value) }
let chartRef=null

function readParams(){
  let avgPriceDemon = num('avgPriceDemon')
  const addLong = (num('timeDemon') > 2.5) && !el('avgIncludesLongFee').checked
  if(addLong) avgPriceDemon += num('longFeeDemon')
  return {
    limitTime: num('limitTime'), limitClean: num('limitClean'),
    limitBigShows: Math.round(num('limitBigShows')), limitSmallShows: Math.round(num('limitSmallShows')),
    minPerTitle: Math.round(num('minPerTitle')),
    capBig: num('capBig'), capSmall: num('capSmall'),
    occDemonBig: num('occDemonBig'), occDemonSmall: num('occDemonSmall'),
    occF1Big: num('occF1Big'), occF1Small: num('occF1Small'),
    demandDemon: num('demandDemon'), demandF1: num('demandF1'),
    priceDemon: num('priceDemon'), timeDemon: num('timeDemon'),
    longFeeDemon: num('longFeeDemon'), avgPriceDemon,
    priceF1: num('priceF1'), timeF1: num('timeF1'), avgPriceF1: num('avgPriceF1'),
    cleanDemonBig: num('cleanDemonBig'), cleanDemonSmall: num('cleanDemonSmall'),
    cleanF1Big: num('cleanF1Big'), cleanF1Small: num('cleanF1Small'),
    varCostPerHead: num('varCostPerHead'),
    usedAvgAddLong: addLong
  }
}

function feasible(x1,x2,y1,y2,p){
  const time = p.timeDemon*(x1+x2) + p.timeF1*(y1+y2)
  if(time > p.limitTime + 1e-9) return {ok:false}
  const clean = p.cleanDemonBig*x1 + p.cleanDemonSmall*x2 + p.cleanF1Big*y1 + p.cleanF1Small*y2
  if(clean > p.limitClean + 1e-9) return {ok:false}
  if(x1 + y1 > p.limitBigShows) return {ok:false}
  if(x2 + y2 > p.limitSmallShows) return {ok:false}
  if(x1 + x2 < p.minPerTitle || y1 + y2 < p.minPerTitle) return {ok:false}

  const headDpos = (p.capBig*p.occDemonBig*0.01)*x1 + (p.capSmall*p.occDemonSmall*0.01)*x2
  const headFpos = (p.capBig*p.occF1Big*0.01)*y1 + (p.capSmall*p.occF1Small*0.01)*y2
  const headD = Math.min(headDpos, p.demandDemon)
  const headF = Math.min(headFpos, p.demandF1)
  const profit = headD*(p.avgPriceDemon - p.varCostPerHead) + headF*(p.avgPriceF1 - p.varCostPerHead)
  return {ok:true, profit, time, clean, headDpos, headFpos, headD, headF}
}

function solve(){
  const p = readParams()
  let best = {profit: -Infinity}
  for(let x1=0;x1<=p.limitBigShows;x1++){
    for(let y1=0;y1<=p.limitBigShows-x1;y1++){
      for(let x2=0;x2<=p.limitSmallShows;x2++){
        for(let y2=0;y2<=p.limitSmallShows-x2;y2++){
          const r = feasible(x1,x2,y1,y2,p)
          if(!r.ok) continue
          if(r.profit > best.profit){ best = {...r, x1,x2,y1,y2} }
        }
      }
    }
  }
  render(best, p); validate(best, p); window.__lastBest = {best, p}
}

function render(best, p){
  const k = el('kpi')
  if(best.profit === -Infinity){ k.innerHTML = '<div class="box"><div class="lbl">結果</div><div class="val">找不到可行解</div></div>'; return }
  k.innerHTML = `
    <div class="box"><div class="lbl">最佳利潤（近似）</div><div class="val">$${Math.round(best.profit).toLocaleString()}</div></div>
    <div class="box"><div class="lbl">鬼滅場次（大/小）</div><div class="val">${best.x1} / ${best.x2}</div></div>
    <div class="box"><div class="lbl">F1 場次（大/小）</div><div class="val">${best.y1} / ${best.y2}</div></div>
    <div class="box"><div class="lbl">觀影人次（鬼滅 / F1）</div><div class="val">${Math.round(best.headD).toLocaleString()} / ${Math.round(best.headF).toLocaleString()}</div></div>
    <div class="box"><div class="lbl">資源使用（時數/清潔）</div><div class="val">${best.time.toFixed(1)}h / ${best.clean.toFixed(1)}</div></div>
    <div class="box"><div class="lbl">平均票價（鬼滅/F1）</div><div class="val">$${p.avgPriceDemon} / $${p.avgPriceF1}</div></div>
  `
  const ctx = document.getElementById('barChart').getContext('2d')
  if(chartRef) chartRef.destroy()
  chartRef = new Chart(ctx, {
    type:'bar', data:{ labels:['鬼滅(大)','鬼滅(小)','F1(大)','F1(小)'], datasets:[{label:'場次',data:[best.x1,best.x2,best.y1,best.y2]}] },
    options:{ responsive:true, plugins:{legend:{display:true}}, scales:{y:{beginAtZero:true,ticks:{precision:0}}} }
  })
}

function addCheck(level, text){
  const li = document.createElement('li')
  li.innerHTML = `<span class="badge ${level}"></span><span>${text}</span>`
  el('checks').appendChild(li)
}

function validate(best, p){
  const ul = el('checks'); ul.innerHTML=''
  const occs = [p.occDemonBig,p.occDemonSmall,p.occF1Big,p.occF1Small]
  if(occs.some(v=> v>0 && v<=1)) addCheck('warn','上座率看起來像小數（≤1），請改用百分比 0~100。')
  if(p.timeDemon<=0 || p.timeF1<=0) addCheck('err','片長需 > 0。') else addCheck('ok','片長：格式正常。')
  if(p.avgPriceDemon<0 || p.avgPriceF1<0) addCheck('err','平均票價不可為負。') else addCheck('ok','平均票價：格式正常。')
  if(p.cleanDemonBig<0 || p.cleanDemonSmall<0 || p.cleanF1Big<0 || p.cleanF1Small<0) addCheck('err','清潔/場不可為負。') else addCheck('ok','清潔/場：格式正常。')
  if(p.capBig<=0 || p.capSmall<=0) addCheck('err','座位數需 > 0。') else addCheck('ok','座位數：格式正常。')
  if(best.profit === -Infinity){ addCheck('err','目前無可行解：請放寬約束或調整參數。'); return }
  addCheck(best.time<=p.limitTime+1e-9?'ok':'err', `放映時數：${best.time.toFixed(1)} / ${p.limitTime}`)
  addCheck(best.clean<=p.limitClean+1e-9?'ok':'err', `清潔工時：${best.clean.toFixed(1)} / ${p.limitClean}`)
  addCheck((best.x1+best.y1)<=p.limitBigShows?'ok':'err', `大廳場次：${best.x1+best.y1} / ${p.limitBigShows}`)
  addCheck((best.x2+best.y2)<=p.limitSmallShows?'ok':'err', `小廳場次：${best.x2+best.y2} / ${p.limitSmallShows}`)
  addCheck((best.x1+best.x2)>=p.minPerTitle?'ok':'err', `鬼滅最少場次：${best.x1+best.x2} ≥ ${p.minPerTitle}`)
  addCheck((best.y1+best.y2)>=p.minPerTitle?'ok':'err', `F1 最少場次：${best.y1+best.y2} ≥ ${p.minPerTitle}`)
  addCheck(best.headD<=p.demandDemon+1e-6?'ok':'err', `鬼滅人次 ≤ 上限：${Math.round(best.headD)} / ${p.demandDemon}`)
  addCheck(best.headF<=p.demandF1+1e-6?'ok':'err', `F1 人次 ≤ 上限：${Math.round(best.headF)} / ${p.demandF1}`)
  if(p.timeDemon>2.5){
    if(p.usedAvgAddLong) addCheck('ok', `長片加價：已加入 $${p.longFeeDemon}，鬼滅平均票價採用 $${p.avgPriceDemon}`)
    else addCheck('warn', `長片加價：未加入（勾選「平均票價已含長片加價」），目前採用 $${p.avgPriceDemon}`)
  }else{
    addCheck('ok','長片加價：片長 ≤ 2.5h，不適用。')
  }
}

// Import/Export config
function exportCfg(){
  const cfg = readParams()
  const blob = new Blob([JSON.stringify(cfg,null,2)], {type:'application/json'})
  const url = URL.createObjectURL(blob); const a = document.createElement('a')
  a.href = url; a.download = 'screening_config.json'; a.click(); URL.revokeObjectURL(url)
}
function importCfg(){ el('importCfgFile').click() }
function handleCfgFile(e){
  const file = e.target.files[0]; if(!file) return
  const r = new FileReader()
  r.onload = () => { try{ const cfg = JSON.parse(r.result); applyConfig(cfg); solve() }catch(e){ alert('JSON 格式錯誤') } }
  r.readAsText(file)
}
function applyConfig(cfg){
  const map = {
    limitTime:'limitTime', limitClean:'limitClean', limitBigShows:'limitBigShows', limitSmallShows:'limitSmallShows', minPerTitle:'minPerTitle',
    capBig:'capBig', capSmall:'capSmall', occDemonBig:'occDemonBig', occDemonSmall:'occDemonSmall', occF1Big:'occF1Big', occF1Small:'occF1Small',
    demandDemon:'demandDemon', demandF1:'demandF1', priceDemon:'priceDemon', timeDemon:'timeDemon', longFeeDemon:'longFeeDemon', avgPriceDemon:'avgPriceDemon',
    priceF1:'priceF1', timeF1:'timeF1', avgPriceF1:'avgPriceF1', cleanDemonBig:'cleanDemonBig', cleanDemonSmall:'cleanDemonSmall', cleanF1Big:'cleanF1Big', cleanF1Small:'cleanF1Small',
    varCostPerHead:'varCostPerHead'
  }
  Object.entries(map).forEach(([k,id])=>{ if(cfg[k] !== undefined){ el(id).value = cfg[k] } })
}

// Excel import + scenarios
const flatKeyMap = {
  "放映時數上限":"limitTime","清潔工時上限":"limitClean","大廳場次上限":"limitBigShows","小廳場次上限":"limitSmallShows",
  "每片最少場次":"minPerTitle","大廳座位數":"capBig","小廳座位數":"capSmall",
  "鬼滅大廳上座率":"occDemonBig","鬼滅小廳上座率":"occDemonSmall","F1大廳上座率":"occF1Big","F1小廳上座率":"occF1Small",
  "鬼滅人次上限":"demandDemon","F1人次上限":"demandF1",
  "鬼滅全票":"priceDemon","鬼滅片長":"timeDemon","鬼滅長片加價":"longFeeDemon","鬼滅平均票價":"avgPriceDemon",
  "F1全票":"priceF1","F1片長":"timeF1","F1平均票價":"avgPriceF1",
  "鬼滅大廳清潔工時":"cleanDemonBig","鬼滅小廳清潔工時":"cleanDemonSmall","F1大廳清潔工時":"cleanF1Big","F1小廳清潔工時":"cleanF1Small",
  "每人變動成本":"varCostPerHead"
}
const matrixHalls=['鬼滅大廳','鬼滅小廳','F1大廳','F1小廳']
const matrixRows=['片長(h)','上座率','清潔/場','座位數','票價(全票)','平均票價(估)']

function importExcel(){ el('excelFile').click() }
function handleExcelFile(e){
  const file = e.target.files[0]; if(!file) return
  const r = new FileReader()
  r.onload = (evt)=>{
    let wb; try{ wb = XLSX.read(evt.target.result, {type:'binary'}) }catch(err){ alert('讀取 Excel 失敗：'+err.message); return }
    pullGlobalParamsFromWorkbook(wb)
    const names = wb.SheetNames.filter(n=> /Base|Scenario|情境/i.test(n))
    const scenarios = []
    for(const s of names){
      const cfg = parseSheetToConfig(wb.Sheets[s])
      if(cfg) scenarios.push({name:s, cfg})
    }
    buildScenarioButtons(scenarios)
    if(scenarios.length){ applyConfig(scenarios[0].cfg); alert(`已載入 ${scenarios.length} 個情境（預設 ${scenarios[0].name}）`) }
    else{ alert('已載入 Excel（未找到情境頁）') }
    solve()
  }
  r.readAsBinaryString(file)
}
function buildScenarioButtons(items){
  const wrap = el('scenarioButtons'); wrap.innerHTML = ''
  items.forEach(({name, cfg})=>{
    const b = document.createElement('button'); b.className='secondary'; b.textContent = name
    b.addEventListener('click', ()=>{ applyConfig(cfg); solve() })
    wrap.appendChild(b)
  })
}
function parseSheetToConfig(sheet){
  const rows = XLSX.utils.sheet_to_json(sheet, {defval:"", raw:true})
  if(rows.length){
    const first = rows[0]; const cfg={}; let any=false
    for(const k of Object.keys(first)){ const std=flatKeyMap[k]||k; cfg[std]=first[k]; any=true }
    if(any) return cfg
  }
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null
  if(!range) return null
  const getCell = (r,c)=> sheet[XLSX.utils.encode_cell({r,c})]?.v ?? ""
  const R=range.e.r, C=range.e.c
  let headerRow=-1
  for(let r=0;r<=R;r++){
    let hit=0; for(let c=0;c<=C;c++){ const v=String(getCell(r,c)).trim(); if(matrixHalls.includes(v)) hit++ }
    if(hit>=2){ headerRow=r; break }
  }
  if(headerRow<0) return null
  const colIdx={}; for(let c=0;c<=C;c++){ const v=String(getCell(headerRow,c)).trim(); if(matrixHalls.includes(v)) colIdx[v]=c }
  let labelCol=-1; for(let c=0;c<=C;c++){ const v=String(getCell(headerRow,c)).trim(); if(!matrixHalls.includes(v) && v!==""){ labelCol=c; break } }
  if(labelCol<0) labelCol=0
  const labelMap={}
  for(let r=headerRow+1;r<=R;r++){ const lab=String(getCell(r,labelCol)).replace(/\s/g,''); if(matrixRows.includes(lab)) labelMap[lab]=r }
  if(Object.keys(labelMap).length===0) return null
  function read(lab,hall){ const rr=labelMap[lab]; const cc=colIdx[hall]; if(rr==null||cc==null) return null; const v=parseFloat(getCell(rr,cc)); return isFinite(v)?v:null }
  const cfg={}
  function setIf(k,val){ if(val!=null && !Number.isNaN(val)) cfg[k]=val }
  setIf('timeDemon', read('片長(h)','鬼滅大廳') ?? read('片長(h)','鬼滅小廳'))
  setIf('timeF1',    read('片長(h)','F1大廳')   ?? read('片長(h)','F1小廳'))
  setIf('occDemonBig',  read('上座率','鬼滅大廳'))
  setIf('occDemonSmall',read('上座率','鬼滅小廳'))
  setIf('occF1Big',     read('上座率','F1大廳'))
  setIf('occF1Small',   read('上座率','F1小廳'))
  setIf('cleanDemonBig',  read('清潔/場','鬼滅大廳'))
  setIf('cleanDemonSmall',read('清潔/場','鬼滅小廳'))
  setIf('cleanF1Big',     read('清潔/場','F1大廳'))
  setIf('cleanF1Small',   read('清潔/場','F1小廳'))
  setIf('capBig',   read('座位數','鬼滅大廳') ?? read('座位數','F1大廳'))
  setIf('capSmall', read('座位數','鬼滅小廳') ?? read('座位數','F1小廳'))
  setIf('priceDemon',    read('票價(全票)','鬼滅大廳') ?? read('票價(全票)','鬼滅小廳'))
  setIf('priceF1',       read('票價(全票)','F1大廳')   ?? read('票價(全票)','F1小廳'))
  setIf('avgPriceDemon', read('平均票價(估)','鬼滅大廳') ?? read('平均票價(估)','鬼滅小廳'))
  setIf('avgPriceF1',    read('平均票價(估)','F1大廳')   ?? read('平均票價(估)','F1小廳'))
  return cfg
}
function pullGlobalParamsFromWorkbook(wb){
  const keys = Object.keys(flatKeyMap).concat(['limitTime','limitClean','limitBigShows','limitSmallShows','minPerTitle','demandDemon','demandF1'])
  for(const s of wb.SheetNames){
    const sheet = wb.Sheets[s]
    const rows = XLSX.utils.sheet_to_json(sheet, {defval:"", raw:true})
    if(rows.length){
      const r0 = rows[0]
      for(const k of Object.keys(r0)){ const std=flatKeyMap[k]||k; if(keys.includes(std) && r0[k] !== ""){ const dom=el(std); if(dom) dom.value = r0[k] } }
    }
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null
    if(range){
      for(let r=range.s.r;r<=range.e.r;r++){
        const label = (sheet[XLSX.utils.encode_cell({r, c: range.s.c})]?.v ?? '').toString().trim()
        const value = sheet[XLSX.utils.encode_cell({r, c: range.s.c+1})]?.v
        const std = flatKeyMap[label] || label
        if(keys.includes(std) && value !== undefined && value!==""){ const dom = el(std); if(dom) dom.value = value }
      }
    }
  }
}

// Export Best
function exportBestCSV(){
  const st = window.__lastBest || {}; const best=st.best
  if(!best){ alert('請先求解'); return }
  const rows = [
    ['指標','數值'],
    ['最佳利潤(近似)', best.profit],
    ['鬼滅(大)場次', best.x1],
    ['鬼滅(小)場次', best.x2],
    ['F1(大)場次', best.y1],
    ['F1(小)場次', best.y2],
    ['鬼滅人次', Math.round(best.headD)],
    ['F1人次', Math.round(best.headF)],
    ['放映時數', best.time],
    ['清潔工時', best.clean],
  ]
  const content = rows.map(r=>r.join(',')).join('\\n')
  const blob = new Blob([content], {type:'text/csv;charset=utf-8'})
  saveAs(blob, 'best_solution.csv')
}
function exportBestXLSX(){
  const st = window.__lastBest || {}; const best=st.best
  if(!best){ alert('請先求解'); return }
  const wsData = [
    ['指標','數值'],
    ['最佳利潤(近似)', best.profit],
    ['鬼滅(大)場次', best.x1],
    ['鬼滅(小)場次', best.x2],
    ['F1(大)場次', best.y1],
    ['F1(小)場次', best.y2],
    ['鬼滅人次', Math.round(best.headD)],
    ['F1人次', Math.round(best.headF)],
    ['放映時數', best.time],
    ['清潔工時', best.clean],
  ]
  const ws = XLSX.utils.aoa_to_sheet(wsData)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'BestSolution')
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'})
  const blob = new Blob([wbout], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
  saveAs(blob, 'best_solution.xlsx')
}

function loadBase(){
  el('limitTime').value = 98; el('limitClean').value = 60
  el('limitBigShows').value = 30; el('limitSmallShows').value = 50; el('minPerTitle').value = 2
  el('avgIncludesLongFee').checked = false
  el('capBig').value = 400; el('capSmall').value = 200
  el('occDemonBig').value = 85; el('occDemonSmall').value = 90
  el('occF1Big').value = 90; el('occF1Small').value = 95
  el('demandDemon').value = 7000; el('demandF1').value = 6000
  el('priceDemon').value = 450; el('timeDemon').value = 2.6; el('longFeeDemon').value = 30; el('avgPriceDemon').value = 300
  el('priceF1').value = 420; el('timeF1').value = 2.1; el('avgPriceF1').value = 260
  el('cleanDemonBig').value = 3; el('cleanDemonSmall').value = 1.5; el('cleanF1Big').value = 2; el('cleanF1Small').value = 1
  el('varCostPerHead').value = 30
}

function bind(){
  document.querySelectorAll('input').forEach(i=>{
    i.addEventListener('input', ()=>{ clearTimeout(window.__t); window.__t=setTimeout(solve,150) })
    i.addEventListener('change', ()=>{ solve() })
  })
  el('btnSolve').addEventListener('click', solve)
  el('btnBase').addEventListener('click', ()=>{ loadBase(); solve() })
  el('btnExportCfg').addEventListener('click', exportCfg)
  el('btnImportCfg').addEventListener('click', importCfg)
  el('importCfgFile').addEventListener('change', handleCfgFile)
  el('btnExcel').addEventListener('click', importExcel)
  el('excelFile').addEventListener('change', handleExcelFile)
  el('btnExportCSV').addEventListener('click', exportBestCSV)
  el('btnExportXLSX').addEventListener('click', exportBestXLSX)
}

window.addEventListener('DOMContentLoaded', ()=>{ loadBase(); bind(); solve() })
