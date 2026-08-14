/* 地基雷达监测报告生成器 Mobile
 * 核心计算逻辑由桌面版 Python 迁移：面/点数据读取、6h/24h报告、24h对比表。
 */
(() => {
  'use strict';

  const FACE_GROUPS = [
    ['上部及后缘(全区变形峰值区)', {faceNames:['M6','M7','M8'], faceRange:'M6-M8', pointNames:['S6','S7','S8'], pointRange:'S6-S8'}],
    ['中部公路边坡', {faceNames:['M1','M2','M3','M4','M5'], faceRange:'M1-M5', pointNames:['S1','S2','S3','S4','S5'], pointRange:'S1-S5'}],
    ['下部边坡', {faceNames:['M9','M10','M11','M12','M13','M14'], faceRange:'M9-M14', pointNames:['S9','S10','S11','S12','S13','S14'], pointRange:'S9-S14'}]
  ];

  const state = {
    faceData: new Map(), pointData: new Map(), faceCols: [], pointCols: [], allTimes: [],
    faceFile: null, pointFile: null, last6: null, last24: null
  };

  const $ = (id) => document.getElementById(id);
  const els = {
    faceFile:$('faceFile'), pointFile:$('pointFile'), faceMeta:$('faceMeta'), pointMeta:$('pointMeta'),
    loadBtn:$('loadBtn'), status:$('status'), s6:$('s6'), e6:$('e6'), s24:$('s24'), e24:$('e24'), ps24:$('ps24'), pe24:$('pe24'),
    report6:$('report6'), report24:$('report24'), gen6:$('gen6'), gen24:$('gen24'), auto6:$('auto6'), auto24:$('auto24'),
    copy6:$('copy6'), copy24:$('copy24'), txt6:$('txt6'), txt24:$('txt24'), xlsx24:$('xlsx24'), share6:$('share6'), share24:$('share24'), toast:$('toast')
  };

  function setStatus(msg, type='') {
    els.status.textContent = msg;
    els.status.className = `status ${type}`.trim();
  }
  function toast(msg) {
    els.toast.textContent = msg; els.toast.classList.add('show');
    clearTimeout(toast._t); toast._t = setTimeout(()=>els.toast.classList.remove('show'), 1800);
  }

  function pad2(n){ return String(n).padStart(2,'0'); }
  function formatDateLocal(d){
    return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  }
  function normalizeTime(v){
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date && !isNaN(v)) return formatDateLocal(v);
    if (typeof v === 'number' && globalThis.XLSX?.SSF) {
      const p = XLSX.SSF.parse_date_code(v);
      if (p) return `${p.y}-${pad2(p.m)}-${pad2(p.d)} ${pad2(p.H)}:${pad2(p.M)}:${pad2(Math.floor(p.S))}`;
    }
    let s = String(v).trim();
    if (!s || s==='X' || s==='Y' || s==='Z' || s.startsWith('=')) return null;
    s = s.replace('T',' ').replace(/\.\d+$/,'');
    if (/^\d{4}-\d{2}-\d{2} \d{2}$/.test(s)) s += ':00:00';
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s)) s += ':00';
    return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(s) ? s : null;
  }
  function parseTime(s){
    const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
    if (!m) throw new Error(`无法解析时间: ${s}`);
    return new Date(+m[1], +m[2]-1, +m[3], +m[4], +m[5], +m[6]);
  }
  function roundHour(d){
    const x = new Date(d);
    if (x.getMinutes() >= 30) x.setHours(x.getHours()+1);
    x.setMinutes(0,0,0); return x;
  }
  function fmtTitle(d){ return formatDateLocal(roundHour(d)); }
  function fmtVal(v){ const n=Number(v)||0; return `${n>=0?'+':''}${n.toFixed(1)}`; }
  function num(v){ const n=Number(v); return Number.isFinite(n)?n:0; }

  function computeItemDisp(data, s, e, name){
    const sr=data.get(s), er=data.get(e); if(!sr||!er) return [null,null];
    const sv=num(sr[name]), ev=num(er[name]); return [ev-sv, ev];
  }
  function computeGroupDisp(data, s, e, names){
    const sr=data.get(s), er=data.get(e); if(!sr||!er) return [null,null,null,null];
    const arr=names.map(name=>{const sv=num(sr[name]),ev=num(er[name]);return {name,d:ev-sv,total:ev};});
    if(!arr.length) return [null,null,null,null];
    const mn=arr.reduce((a,b)=>b.d<a.d?b:a), mx=arr.reduce((a,b)=>b.d>a.d?b:a);
    return [mn.d,mx.d,mx.name,mx.total];
  }
  function sortNames(names){
    return [...names].sort((a,b)=>{const na=+(String(a).match(/\d+/)?.[0]||999), nb=+(String(b).match(/\d+/)?.[0]||999); return na-nb || String(a).localeCompare(String(b));});
  }

  async function workbookFromFile(file){
    if (!globalThis.XLSX) throw new Error('Excel 组件未加载。首次打开请保持网络连接后刷新页面。');
    const buf = await file.arrayBuffer();
    // nodim:true 很关键：雷达平台导出的 XLSX 自报工作表范围可能只有 A1，实际数据远超该范围。
    return XLSX.read(buf, {type:'array', nodim:true, cellDates:false, cellFormula:false});
  }

  async function loadRadarFile(file, kind){
    const wb=await workbookFromFile(file);
    const ws=wb.Sheets['位移'];
    if(!ws) throw new Error(`${file.name} 中未找到“位移”工作表`);
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});
    if(!rows.length || !rows[0] || rows[0].length<2) throw new Error(`${file.name} 的“位移”工作表未读取到有效列`);
    const headers=rows[0].slice(1).map(v=>v==null?'':String(v).trim());
    const cols=headers.filter(Boolean);
    const startIndex=kind==='face'?1:4;
    const data=new Map();
    for(let r=startIndex;r<rows.length;r++){
      const row=rows[r]||[]; const ts=normalizeTime(row[0]); if(!ts) continue;
      const values={};
      for(let c=1;c<=headers.length;c++){
        const name=headers[c-1]; if(!name) continue; values[name]=num(row[c]);
      }
      data.set(ts,values);
    }
    if(!data.size) throw new Error(`${file.name} 中没有识别到位移时间序列`);
    return {data,cols};
  }

  function detectKindByHeader(file){
    return workbookFromFile(file).then(wb=>{
      const ws=wb.Sheets['位移']; if(!ws) return null;
      const a1=ws['A1']?.v == null ? '' : String(ws['A1'].v);
      if(a1.includes('监测面')) return 'face'; if(a1.includes('监测点')) return 'point'; return null;
    }).catch(()=>null);
  }

  function fillSelect(sel,times){
    sel.innerHTML='';
    times.forEach(t=>{const o=document.createElement('option');o.value=t;o.textContent=t;sel.appendChild(o);});
  }
  function setSel(sel,val){ if(val && state.allTimes.includes(val)) sel.value=val; }

  function nearestTo(target, maxMinutes=90){
    if(!state.allTimes.length) return null;
    let best=null,bestDiff=Infinity;
    for(const t of state.allTimes){const diff=Math.abs(parseTime(t)-target);if(diff<bestDiff){best=t;bestDiff=diff;}}
    return bestDiff<=maxMinutes*60000?best:null;
  }
  function nearestBefore(target, maxMinutes=90){
    let best=null,bestDiff=Infinity;
    for(const t of state.allTimes){const dt=parseTime(t),diff=target-dt;if(diff>=0&&diff<bestDiff){best=t;bestDiff=diff;}}
    return bestDiff<=maxMinutes*60000?best:null;
  }

  function auto6(){
    if(!state.allTimes.length) return;
    const latest=parseTime(state.allTimes[state.allTimes.length-1]);
    const boundaryHours=[0,6,12,18,24];
    let endBoundary=null;
    for(let back=0;back<3 && !endBoundary;back++){
      const base=new Date(latest);base.setDate(base.getDate()-back);base.setHours(0,0,0,0);
      for(const h of boundaryHours){const x=new Date(base);x.setHours(h,0,0,0);if(x<=latest && (!endBoundary||x>endBoundary)) endBoundary=x;}
    }
    if(!endBoundary) return;
    const startBoundary=new Date(endBoundary.getTime()-6*3600000);
    let es=nearestTo(endBoundary,90) || nearestBefore(endBoundary,90);
    let ss=nearestTo(startBoundary,90) || nearestBefore(startBoundary,90);
    if(!es || !ss || parseTime(ss)>=parseTime(es)){
      es=state.allTimes[state.allTimes.length-1];
      const target=new Date(parseTime(es).getTime()-6*3600000); ss=nearestTo(target,120) || state.allTimes[0];
    }
    setSel(els.s6,ss); setSel(els.e6,es);
  }

  function auto24(){
    if(!state.allTimes.length) return;
    const es=state.allTimes[state.allTimes.length-1], ed=parseTime(es);
    const targetStart=new Date(ed); targetStart.setDate(targetStart.getDate()-1); targetStart.setHours(15,0,0,0);
    let ss=nearestTo(targetStart,120);
    if(!ss){ const target=new Date(ed.getTime()-24*3600000); ss=nearestTo(target,180) || state.allTimes[0]; }
    setSel(els.s24,ss); setSel(els.e24,es);

    const curStart=parseTime(ss), duration=parseTime(es)-curStart;
    const prevEndTarget=new Date(curStart), prevStartTarget=new Date(curStart.getTime()-duration);
    const pe=nearestTo(prevEndTarget,120), ps=nearestTo(prevStartTarget,180);
    if(ps&&pe&&parseTime(ps)<parseTime(pe)){setSel(els.ps24,ps);setSel(els.pe24,pe);}
  }

  function genReportText(ss,es,label){
    const lines=[`地基雷达${label}小时监测结果（${fmtTitle(parseTime(ss))} 至 ${fmtTitle(parseTime(es))}）`];
    for(const [gn,gi] of FACE_GROUPS){
      const [fMin,fMax,fn,ft]=computeGroupDisp(state.faceData,ss,es,gi.faceNames);
      const [pMin,pMax,pn,pt]=computeGroupDisp(state.pointData,ss,es,gi.pointNames);
      if(fMin===null||pMin===null) continue;
      lines.push(`- ${gn}：监测面（${gi.faceRange}）${label}小时累积位移${fmtVal(fMin)}~${fmtVal(fMax)}mm，最大值位于${fn}，总累积位移${fmtVal(ft)}mm；监测点（${gi.pointRange}）${label}小时累积位移${fmtVal(pMin)}~${fmtVal(pMax)}mm，最大值位于${pn}，总累积位移${fmtVal(pt)}mm。`);
    }
    return lines.length>1?lines.join('\n'):null;
  }

  function validatePeriod(s,e,name='周期'){
    if(!s||!e){toast(`请选择${name}起止时间`);return false;}
    if(parseTime(s)>=parseTime(e)){toast(`${name}开始时间必须早于结束时间`);return false;} return true;
  }

  function generate6(){
    if(!state.faceData.size){toast('请先加载数据');return;}
    const ss=els.s6.value,es=els.e6.value;if(!validatePeriod(ss,es,'6小时'))return;
    const text=genReportText(ss,es,'6'); if(!text){toast('生成报告失败');return;}
    els.report6.value=text; state.last6={text,ss,es}; enableOutputs(); toast('6小时报告已生成');
  }
  function generate24(){
    if(!state.faceData.size){toast('请先加载数据');return;}
    const ss=els.s24.value,es=els.e24.value,ps=els.ps24.value,pe=els.pe24.value;
    if(!validatePeriod(ss,es,'当前周期'))return;
    const hasPrev=!!(ps&&pe&&parseTime(ps)<parseTime(pe));
    const text=genReportText(ss,es,'24'); if(!text){toast('生成报告失败');return;}
    els.report24.value=text; state.last24={text,ss,es,ps,pe,hasPrev}; enableOutputs(); toast('24小时报告已生成');
  }

  function downloadBlob(blob,name){
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
    setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},1000);
  }
  function downloadText(kind){
    const x=kind==='6h'?state.last6:state.last24;if(!x)return;
    downloadBlob(new Blob([x.text+'\n'],{type:'text/plain;charset=utf-8'}),`${kind}_report.txt`);
  }
  async function copyText(kind){
    const x=kind==='6h'?state.last6:state.last24;if(!x)return;
    try{await navigator.clipboard.writeText(x.text);toast('报告已复制');}
    catch{const ta=document.createElement('textarea');ta.value=x.text;document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();toast('报告已复制');}
  }
  async function shareText(kind){
    const x=kind==='6h'?state.last6:state.last24;if(!x)return;
    try{
      if(navigator.share){await navigator.share({title:`地基雷达${kind==='6h'?'6':'24'}小时监测报告`,text:x.text});}
      else await copyText(kind);
    }catch(e){ if(e?.name!=='AbortError') toast('分享失败，可使用“复制报告”'); }
  }

  function dateLabel(a,b){const ad=parseTime(a),bd=parseTime(b);return `${ad.getMonth()+1}/${ad.getDate()}-${bd.getMonth()+1}/${bd.getDate()}`;}
  function buildExportAOA(data,cols,ss,es,ps,pe,hasPrev,isFace){
    const dl=dateLabel(ss,es), headers=['监测点/面',`${dl}位移量(mm)`];
    if(hasPrev) headers.push(`相对于${dateLabel(ps,pe)}变化量(mm)`); headers.push('总累积位移量(mm)');
    const rows=[headers];
    for(const n of sortNames(cols)){
      if(isFace && n==='W1') continue;
      const [da,ta]=computeItemDisp(data,ss,es,n);if(da===null)continue;
      const row=[n,Math.round(da*10)/10];
      if(hasPrev){const [db]=computeItemDisp(data,ps,pe,n);row.push(db===null?'':Math.round((da-db)*10)/10);}
      row.push(Math.round(ta*10)/10);rows.push(row);
    }
    return rows;
  }
  function downloadExcel24(){
    const x=state.last24;if(!x){toast('请先生成24小时报告');return;}
    const wb=XLSX.utils.book_new();
    const ws1=XLSX.utils.aoa_to_sheet(buildExportAOA(state.faceData,state.faceCols,x.ss,x.es,x.ps,x.pe,x.hasPrev,true));
    const ws2=XLSX.utils.aoa_to_sheet(buildExportAOA(state.pointData,state.pointCols,x.ss,x.es,x.ps,x.pe,x.hasPrev,false));
    ws1['!cols']=[{wch:14},{wch:19},{wch:24},{wch:19}];ws2['!cols']=ws1['!cols'];
    XLSX.utils.book_append_sheet(wb,ws1,'面');XLSX.utils.book_append_sheet(wb,ws2,'点');
    XLSX.writeFileXLSX(wb,'24h对比表.xlsx',{compression:true});
  }

  function enableOutputs(){
    const has6=!!state.last6, has24=!!state.last24;
    [els.copy6,els.txt6,els.share6].forEach(b=>b.disabled=!has6);
    [els.copy24,els.txt24,els.share24,els.xlsx24].forEach(b=>b.disabled=!has24);
  }

  async function loadData(){
    if(!state.faceFile||!state.pointFile){setStatus('请先选择面数据和点数据两个 Excel 文件。','warn');return;}
    els.loadBtn.disabled=true;setStatus('正在读取两份 Excel，请稍候…');
    try{
      const [fr,pr]=await Promise.all([loadRadarFile(state.faceFile,'face'),loadRadarFile(state.pointFile,'point')]);
      state.faceData=fr.data;state.faceCols=fr.cols;state.pointData=pr.data;state.pointCols=pr.cols;
      const pt=new Set(state.pointData.keys());state.allTimes=[...state.faceData.keys()].filter(t=>pt.has(t)).sort((a,b)=>parseTime(a)-parseTime(b));
      if(!state.allTimes.length) throw new Error('面数据和点数据没有共同时间点');
      [els.s6,els.e6,els.s24,els.e24,els.ps24,els.pe24].forEach(s=>fillSelect(s,state.allTimes));
      state.last6=state.last24=null;els.report6.value='';els.report24.value='';enableOutputs();
      auto6();auto24();
      setStatus(`已加载 ${state.allTimes.length} 个共同时间点\n${state.allTimes[0]}  ～  ${state.allTimes[state.allTimes.length-1]}`,'ok');
    }catch(e){console.error(e);setStatus(`加载失败：${e.message||e}`,'err');}
    finally{els.loadBtn.disabled=false;}
  }

  function fileMeta(file,kind){return file?`${file.name}\n${(file.size/1024/1024).toFixed(2)} MB · ${kind}`:'未选择文件';}
  els.faceFile.addEventListener('change',async()=>{state.faceFile=els.faceFile.files[0]||null;els.faceMeta.textContent=fileMeta(state.faceFile,'面数据');if(state.faceFile){const k=await detectKindByHeader(state.faceFile);if(k&&k!=='face')setStatus('当前“面数据”文件看起来是监测点文件，请检查是否选反。','warn');}});
  els.pointFile.addEventListener('change',async()=>{state.pointFile=els.pointFile.files[0]||null;els.pointMeta.textContent=fileMeta(state.pointFile,'点数据');if(state.pointFile){const k=await detectKindByHeader(state.pointFile);if(k&&k!=='point')setStatus('当前“点数据”文件看起来是监测面文件，请检查是否选反。','warn');}});
  els.loadBtn.addEventListener('click',loadData);els.auto6.addEventListener('click',auto6);els.auto24.addEventListener('click',auto24);els.gen6.addEventListener('click',generate6);els.gen24.addEventListener('click',generate24);
  els.copy6.addEventListener('click',()=>copyText('6h'));els.copy24.addEventListener('click',()=>copyText('24h'));els.txt6.addEventListener('click',()=>downloadText('6h'));els.txt24.addEventListener('click',()=>downloadText('24h'));els.xlsx24.addEventListener('click',downloadExcel24);els.share6.addEventListener('click',()=>shareText('6h'));els.share24.addEventListener('click',()=>shareText('24h'));

  document.querySelectorAll('.tab-btn').forEach(btn=>btn.addEventListener('click',()=>{
    document.querySelectorAll('.tab-btn').forEach(x=>x.classList.toggle('active',x===btn));
    document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('active',p.id===btn.dataset.tab));
  }));

  if('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('./sw.js').catch(()=>{});
  enableOutputs();
})();
