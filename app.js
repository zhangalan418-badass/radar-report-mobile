/* 地基雷达监测报告生成器 Mobile V2
 * 核心计算逻辑由桌面版 Python 迁移：面/点数据读取、6h/24h/一周报告、24h/一周对比表。
 */
(() => {
  'use strict';

  const FACE_GROUPS = [
    ['上部及后缘(全区变形峰值区)', {faceNames:['M6','M7','M8'], faceRange:'M6-M8', pointNames:['S6','S7','S8'], pointRange:'S6-S8'}],
    ['中部公路边坡', {faceNames:['M1','M2','M3','M4','M5'], faceRange:'M1-M5', pointNames:['S1','S2','S3','S4','S5'], pointRange:'S1-S5'}],
    ['下部边坡', {faceNames:['M9','M10','M11','M12','M13','M14'], faceRange:'M9-M14', pointNames:['S9','S10','S11','S12','S13','S14'], pointRange:'S9-S14'}]
  ];

  const state = {
    faceData:new Map(), pointData:new Map(), faceCols:[], pointCols:[], allTimes:[],
    faceFile:null, pointFile:null, last6:null, last24:null, lastWeek:null,
    readyTabs:{p6:false,p24:false,pweek:false}
  };

  const $ = id => document.getElementById(id);
  const els = {
    faceFile:$('faceFile'),pointFile:$('pointFile'),faceMeta:$('faceMeta'),pointMeta:$('pointMeta'),loadBtn:$('loadBtn'),status:$('status'),toast:$('toast'),
    s6:$('s6'),e6:$('e6'),auto6:$('auto6'),gen6:$('gen6'),report6:$('report6'),copy6:$('copy6'),share6:$('share6'),txt6:$('txt6'),
    s24:$('s24'),e24:$('e24'),ps24:$('ps24'),pe24:$('pe24'),auto24:$('auto24'),gen24:$('gen24'),report24:$('report24'),copy24:$('copy24'),share24:$('share24'),txt24:$('txt24'),xlsx24:$('xlsx24'),
    sw:$('sw'),ew:$('ew'),psw:$('psw'),pew:$('pew'),autoWeek:$('autoWeek'),genWeek:$('genWeek'),reportWeek:$('reportWeek'),copyWeek:$('copyWeek'),shareWeek:$('shareWeek'),txtWeek:$('txtWeek'),xlsxWeek:$('xlsxWeek')
  };

  const tabSelects = {
    p6:[els.s6,els.e6],
    p24:[els.s24,els.e24,els.ps24,els.pe24],
    pweek:[els.sw,els.ew,els.psw,els.pew]
  };

  function setStatus(msg,type=''){els.status.textContent=msg;els.status.className=`status ${type}`.trim();}
  function toast(msg){els.toast.textContent=msg;els.toast.classList.add('show');clearTimeout(toast._t);toast._t=setTimeout(()=>els.toast.classList.remove('show'),1800);}
  function pad2(n){return String(n).padStart(2,'0');}
  function formatDateLocal(d){return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;}
  function normalizeTime(v){
    if(v===null||v===undefined||v==='') return null;
    if(v instanceof Date&&!isNaN(v)) return formatDateLocal(v);
    if(typeof v==='number'&&globalThis.XLSX?.SSF){const p=XLSX.SSF.parse_date_code(v);if(p)return `${p.y}-${pad2(p.m)}-${pad2(p.d)} ${pad2(p.H)}:${pad2(p.M)}:${pad2(Math.floor(p.S))}`;}
    let s=String(v).trim();if(!s||s==='X'||s==='Y'||s==='Z'||s.startsWith('='))return null;
    s=s.replace('T',' ').replace(/\.\d+$/,'');if(/^\d{4}-\d{2}-\d{2} \d{2}$/.test(s))s+=':00:00';if(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s))s+=':00';
    return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(s)?s:null;
  }
  function parseTime(s){const m=String(s).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);if(!m)throw new Error(`无法解析时间: ${s}`);return new Date(+m[1],+m[2]-1,+m[3],+m[4],+m[5],+m[6]);}
  function roundHour(d){const x=new Date(d);if(x.getMinutes()>=30)x.setHours(x.getHours()+1);x.setMinutes(0,0,0);return x;}
  function fmtTitle(d){return formatDateLocal(roundHour(d));}
  function fmtVal(v){const n=Number(v)||0;return `${n>=0?'+':''}${n.toFixed(1)}`;}
  function num(v){const n=Number(v);return Number.isFinite(n)?n:0;}
  function round1(v){return Math.round((v+Number.EPSILON)*10)/10;}

  function computeItemDisp(data,s,e,name){const sr=data.get(s),er=data.get(e);if(!sr||!er)return[null,null];const sv=num(sr[name]),ev=num(er[name]);return[ev-sv,ev];}
  function computeGroupDisp(data,s,e,names){const sr=data.get(s),er=data.get(e);if(!sr||!er)return[null,null,null,null];const arr=names.map(name=>{const sv=num(sr[name]),ev=num(er[name]);return{name,d:ev-sv,total:ev};});if(!arr.length)return[null,null,null,null];const mn=arr.reduce((a,b)=>b.d<a.d?b:a),mx=arr.reduce((a,b)=>b.d>a.d?b:a);return[mn.d,mx.d,mx.name,mx.total];}
  function groupChangeRange(data,ss,es,ps,pe,names){const vals=[];for(const n of names){const[cur]=computeItemDisp(data,ss,es,n),[prev]=computeItemDisp(data,ps,pe,n);if(cur===null||prev===null)continue;vals.push(cur-prev);}if(!vals.length)return[null,null];return[Math.min(...vals),Math.max(...vals)];}
  function sortNames(names){return[...names].sort((a,b)=>{const na=+(String(a).match(/\d+/)?.[0]||999),nb=+(String(b).match(/\d+/)?.[0]||999);return na-nb||String(a).localeCompare(String(b));});}

  async function workbookFromFile(file){if(!globalThis.XLSX)throw new Error('Excel 组件未加载。首次打开请保持网络连接后刷新页面。');const buf=await file.arrayBuffer();return XLSX.read(buf,{type:'array',nodim:true,cellDates:false,cellFormula:false});}
  async function loadRadarFile(file,kind){
    const wb=await workbookFromFile(file),ws=wb.Sheets['位移'];if(!ws)throw new Error(`${file.name} 中未找到“位移”工作表`);
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});if(!rows.length||!rows[0]||rows[0].length<2)throw new Error(`${file.name} 的“位移”工作表未读取到有效列`);
    const headers=rows[0].slice(1).map(v=>v==null?'':String(v).trim()),cols=headers.filter(Boolean),startIndex=kind==='face'?1:4,data=new Map();
    for(let r=startIndex;r<rows.length;r++){const row=rows[r]||[],ts=normalizeTime(row[0]);if(!ts)continue;const values={};for(let c=1;c<=headers.length;c++){const name=headers[c-1];if(!name)continue;values[name]=num(row[c]);}data.set(ts,values);}
    if(!data.size)throw new Error(`${file.name} 中没有识别到位移时间序列`);return{data,cols};
  }
  function detectKindByHeader(file){return workbookFromFile(file).then(wb=>{const ws=wb.Sheets['位移'];if(!ws)return null;const a1=ws['A1']?.v==null?'':String(ws['A1'].v);if(a1.includes('监测面'))return'face';if(a1.includes('监测点'))return'point';return null;}).catch(()=>null);}

  function fillSelect(sel,times){const frag=document.createDocumentFragment();for(const t of times){const o=document.createElement('option');o.value=t;o.textContent=t;frag.appendChild(o);}sel.replaceChildren(frag);}
  function ensureTabSelects(tabId){if(state.readyTabs[tabId]||!state.allTimes.length)return;for(const s of tabSelects[tabId]||[])fillSelect(s,state.allTimes);state.readyTabs[tabId]=true;}
  function setSel(sel,val){if(val&&state.allTimes.includes(val)){if(!sel.options.length)fillSelect(sel,state.allTimes);sel.value=val;}}
  function clearSel(sel){if(sel.options.length)sel.selectedIndex=-1;}

  function nearestTime(target,maxGapHours=null){if(!state.allTimes.length)return null;let best=null,bestDiff=Infinity;for(const t of state.allTimes){const diff=Math.abs(parseTime(t)-target);if(diff<bestDiff){best=t;bestDiff=diff;}}if(maxGapHours!==null&&bestDiff>maxGapHours*3600000)return null;return best;}
  function nearestBefore(target,maxMinutes=90){let best=null,bestDiff=Infinity;for(const t of state.allTimes){const dt=parseTime(t),diff=target-dt;if(diff>=0&&diff<bestDiff){best=t;bestDiff=diff;}}return bestDiff<=maxMinutes*60000?best:null;}

  function auto6(){
    if(!state.allTimes.length)return;ensureTabSelects('p6');const latest=parseTime(state.allTimes[state.allTimes.length-1]),boundaryHours=[0,6,12,18,24];let endBoundary=null;
    for(let back=0;back<3&&!endBoundary;back++){const base=new Date(latest);base.setDate(base.getDate()-back);base.setHours(0,0,0,0);for(const h of boundaryHours){const x=new Date(base);x.setHours(h,0,0,0);if(x<=latest&&(!endBoundary||x>endBoundary))endBoundary=x;}}
    if(!endBoundary)return;const startBoundary=new Date(endBoundary.getTime()-6*3600000);let es=nearestTime(endBoundary,1.5)||nearestBefore(endBoundary,90),ss=nearestTime(startBoundary,1.5)||nearestBefore(startBoundary,90);
    if(!es||!ss||parseTime(ss)>=parseTime(es)){es=state.allTimes[state.allTimes.length-1];ss=nearestTime(new Date(parseTime(es).getTime()-6*3600000),2)||state.allTimes[0];}
    setSel(els.s6,ss);setSel(els.e6,es);
  }

  function auto24(){
    if(!state.allTimes.length)return;ensureTabSelects('p24');const es=state.allTimes[state.allTimes.length-1],ed=parseTime(es),targetStart=new Date(ed);targetStart.setDate(targetStart.getDate()-1);targetStart.setHours(15,0,0,0);
    let ss=nearestTime(targetStart,2);if(!ss)ss=nearestTime(new Date(ed.getTime()-24*3600000),3)||state.allTimes[0];setSel(els.s24,ss);setSel(els.e24,es);
    const curStart=parseTime(ss),duration=parseTime(es)-curStart,pe=nearestTime(new Date(curStart),2),ps=nearestTime(new Date(curStart.getTime()-duration),3);
    if(ps&&pe&&parseTime(ps)<parseTime(pe)){setSel(els.ps24,ps);setSel(els.pe24,pe);}else{clearSel(els.ps24);clearSel(els.pe24);}
  }

  function autoWeek(){
    if(!state.allTimes.length)return;ensureTabSelects('pweek');const end=state.allTimes[state.allTimes.length-1],ed=parseTime(end),start=nearestTime(new Date(ed.getTime()-7*24*3600000),12);
    if(!start||parseTime(start)>=ed){clearSel(els.sw);setSel(els.ew,end);clearSel(els.psw);clearSel(els.pew);return;}
    setSel(els.sw,start);setSel(els.ew,end);const prevEnd=start,prevStart=nearestTime(new Date(parseTime(start).getTime()-7*24*3600000),12);
    if(prevStart&&parseTime(prevStart)<parseTime(prevEnd)){setSel(els.psw,prevStart);setSel(els.pew,prevEnd);}else{clearSel(els.psw);clearSel(els.pew);}
  }

  function genReportText(ss,es,label){const lines=[`地基雷达${label}小时监测结果（${fmtTitle(parseTime(ss))} 至 ${fmtTitle(parseTime(es))}）`];for(const[gn,gi]of FACE_GROUPS){const[fMin,fMax,fn,ft]=computeGroupDisp(state.faceData,ss,es,gi.faceNames),[pMin,pMax,pn,pt]=computeGroupDisp(state.pointData,ss,es,gi.pointNames);if(fMin===null||pMin===null)continue;lines.push(`- ${gn}：监测面（${gi.faceRange}）${label}小时累积位移${fmtVal(fMin)}~${fmtVal(fMax)}mm，最大值位于${fn}，总累积位移${fmtVal(ft)}mm；监测点（${gi.pointRange}）${label}小时累积位移${fmtVal(pMin)}~${fmtVal(pMax)}mm，最大值位于${pn}，总累积位移${fmtVal(pt)}mm。`);}return lines.length>1?lines.join('\n'):null;}

  function genWeekReportText(ss,es,ps=null,pe=null){
    const lines=[`地基雷达一周监测结果（${fmtTitle(parseTime(ss))} 至 ${fmtTitle(parseTime(es))}）`],hasPrev=!!(ps&&pe&&parseTime(ps)<parseTime(pe));
    for(const[gn,gi]of FACE_GROUPS){
      const[fMin,fMax,fn,ft]=computeGroupDisp(state.faceData,ss,es,gi.faceNames),[pMin,pMax,pn,pt]=computeGroupDisp(state.pointData,ss,es,gi.pointNames);if(fMin===null||pMin===null)continue;
      let para=`${gn}：监测面（${gi.faceRange}）本周（7天）时段位移${fmtVal(fMin)}~${fmtVal(fMax)}mm，最大值位于${fn}，截至期末总累积位移${fmtVal(ft)}mm；监测点（${gi.pointRange}）本周（7天）时段位移${fmtVal(pMin)}~${fmtVal(pMax)}mm，最大值位于${pn}，截至期末总累积位移${fmtVal(pt)}mm。`;
      if(hasPrev){const[fcMin,fcMax]=groupChangeRange(state.faceData,ss,es,ps,pe,gi.faceNames),[pcMin,pcMax]=groupChangeRange(state.pointData,ss,es,ps,pe,gi.pointNames);if(fcMin!==null&&pcMin!==null)para+=`与上周相比，监测面位移变化${fmtVal(fcMin)}~${fmtVal(fcMax)}mm，监测点位移变化${fmtVal(pcMin)}~${fmtVal(pcMax)}mm。`;}
      else para+=' 历史数据不足14天，暂无法计算与上周相比的位移变化。';
      lines.push(para);
    }
    return lines.length>1?lines.join('\n\n'):null;
  }

  function validatePeriod(s,e,name='周期'){if(!s||!e){toast(`请选择${name}起止时间`);return false;}if(parseTime(s)>=parseTime(e)){toast(`${name}开始时间必须早于结束时间`);return false;}return true;}
  function generate6(){if(!state.faceData.size){toast('请先加载数据');return;}const ss=els.s6.value,es=els.e6.value;if(!validatePeriod(ss,es,'6小时'))return;const text=genReportText(ss,es,'6');if(!text){toast('生成报告失败');return;}els.report6.value=text;state.last6={text,ss,es};enableOutputs();toast('6小时报告已生成');}
  function generate24(){if(!state.faceData.size){toast('请先加载数据');return;}const ss=els.s24.value,es=els.e24.value,ps=els.ps24.value,pe=els.pe24.value;if(!validatePeriod(ss,es,'当前周期'))return;const hasPrev=!!(ps&&pe&&parseTime(ps)<parseTime(pe)),text=genReportText(ss,es,'24');if(!text){toast('生成报告失败');return;}els.report24.value=text;state.last24={text,ss,es,ps,pe,hasPrev};enableOutputs();toast('24小时报告已生成');}
  function generateWeek(){
    if(!state.faceData.size){toast('请先加载数据');return;}const ss=els.sw.value,es=els.ew.value,ps=els.psw.value,pe=els.pew.value;if(!validatePeriod(ss,es,'本周'))return;const hasPrev=!!(ps&&pe);if(hasPrev&&parseTime(ps)>=parseTime(pe)){toast('上周开始时间必须早于结束时间');return;}
    const text=genWeekReportText(ss,es,ps,pe);if(!text){toast('生成一周报告失败');return;}els.reportWeek.value=text;state.lastWeek={text,ss,es,ps,pe,hasPrev};enableOutputs();toast('一周报告已生成');
  }

  function getLast(kind){return kind==='6h'?state.last6:kind==='24h'?state.last24:state.lastWeek;}
  function reportTitle(kind){return kind==='6h'?'地基雷达6小时监测报告':kind==='24h'?'地基雷达24小时监测报告':'地基雷达一周监测报告';}
  function downloadBlob(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},1000);}
  function downloadText(kind){const x=getLast(kind);if(!x)return;const name=kind==='week'?'一周监测报告.txt':`${kind}_report.txt`;downloadBlob(new Blob([x.text+'\n'],{type:'text/plain;charset=utf-8'}),name);}
  async function copyText(kind){const x=getLast(kind);if(!x)return;try{await navigator.clipboard.writeText(x.text);toast('报告已复制');}catch{const ta=document.createElement('textarea');ta.value=x.text;document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();toast('报告已复制');}}
  async function shareText(kind){const x=getLast(kind);if(!x)return;try{if(navigator.share)await navigator.share({title:reportTitle(kind),text:x.text});else await copyText(kind);}catch(e){if(e?.name!=='AbortError')toast('分享失败，可使用“复制报告”');}}

  function dateLabel(a,b){const ad=parseTime(a),bd=parseTime(b);return`${ad.getMonth()+1}/${ad.getDate()}-${bd.getMonth()+1}/${bd.getDate()}`;}
  function buildExportAOA(data,cols,ss,es,ps,pe,hasPrev,isFace){const headers=['监测点/面',`${dateLabel(ss,es)}位移量(mm)`];if(hasPrev)headers.push(`相对于${dateLabel(ps,pe)}变化量(mm)`);headers.push('总累积位移量(mm)');const rows=[headers];for(const n of sortNames(cols)){if(isFace&&n==='W1')continue;const[da,ta]=computeItemDisp(data,ss,es,n);if(da===null)continue;const row=[n,round1(da)];if(hasPrev){const[db]=computeItemDisp(data,ps,pe,n);row.push(db===null?'':round1(da-db));}row.push(round1(ta));rows.push(row);}return rows;}
  function buildWeekExportAOA(data,cols,ss,es,ps,pe,hasPrev,isFace){const headers=['监测点/面','本周（7天）时段位移(mm)'];if(hasPrev)headers.push('上周（7天）时段位移(mm)','与上周相比变化量(mm)');headers.push('截至期末总累积位移(mm)');const rows=[headers];for(const n of sortNames(cols)){if(isFace&&n==='W1')continue;const[cur,total]=computeItemDisp(data,ss,es,n);if(cur===null)continue;const row=[n,round1(cur)];if(hasPrev){const[prev]=computeItemDisp(data,ps,pe,n);row.push(prev===null?'':round1(prev));row.push(prev===null?'':round1(cur-prev));}row.push(round1(total));rows.push(row);}return rows;}
  function setCols(ws,count){const widths=[14,22,22,22,24];ws['!cols']=widths.slice(0,count).map(wch=>({wch}));}
  function downloadExcel24(){const x=state.last24;if(!x){toast('请先生成24小时报告');return;}const wb=XLSX.utils.book_new(),a1=buildExportAOA(state.faceData,state.faceCols,x.ss,x.es,x.ps,x.pe,x.hasPrev,true),a2=buildExportAOA(state.pointData,state.pointCols,x.ss,x.es,x.ps,x.pe,x.hasPrev,false),ws1=XLSX.utils.aoa_to_sheet(a1),ws2=XLSX.utils.aoa_to_sheet(a2);setCols(ws1,a1[0].length);setCols(ws2,a2[0].length);XLSX.utils.book_append_sheet(wb,ws1,'面');XLSX.utils.book_append_sheet(wb,ws2,'点');XLSX.writeFileXLSX(wb,'24h对比表.xlsx',{compression:true});}
  function downloadExcelWeek(){const x=state.lastWeek;if(!x){toast('请先生成一周报告');return;}const wb=XLSX.utils.book_new(),a1=buildWeekExportAOA(state.faceData,state.faceCols,x.ss,x.es,x.ps,x.pe,x.hasPrev,true),a2=buildWeekExportAOA(state.pointData,state.pointCols,x.ss,x.es,x.ps,x.pe,x.hasPrev,false),ws1=XLSX.utils.aoa_to_sheet(a1),ws2=XLSX.utils.aoa_to_sheet(a2);setCols(ws1,a1[0].length);setCols(ws2,a2[0].length);XLSX.utils.book_append_sheet(wb,ws1,'面');XLSX.utils.book_append_sheet(wb,ws2,'点');XLSX.writeFileXLSX(wb,'一周对比表.xlsx',{compression:true});}

  function enableOutputs(){const h6=!!state.last6,h24=!!state.last24,hw=!!state.lastWeek;[els.copy6,els.share6,els.txt6].forEach(b=>b.disabled=!h6);[els.copy24,els.share24,els.txt24,els.xlsx24].forEach(b=>b.disabled=!h24);[els.copyWeek,els.shareWeek,els.txtWeek,els.xlsxWeek].forEach(b=>b.disabled=!hw);}
  function resetTabs(){state.readyTabs={p6:false,p24:false,pweek:false};for(const arr of Object.values(tabSelects))for(const s of arr)s.replaceChildren();}

  async function loadData(){
    if(!state.faceFile||!state.pointFile){setStatus('请先选择面数据和点数据两个 Excel 文件。','warn');return;}els.loadBtn.disabled=true;setStatus('正在读取两份 Excel，请稍候…');
    try{
      const[fr,pr]=await Promise.all([loadRadarFile(state.faceFile,'face'),loadRadarFile(state.pointFile,'point')]);state.faceData=fr.data;state.faceCols=fr.cols;state.pointData=pr.data;state.pointCols=pr.cols;const pt=new Set(state.pointData.keys());state.allTimes=[...state.faceData.keys()].filter(t=>pt.has(t)).sort((a,b)=>parseTime(a)-parseTime(b));if(!state.allTimes.length)throw new Error('面数据和点数据没有共同时间点');
      state.last6=state.last24=state.lastWeek=null;els.report6.value='';els.report24.value='';els.reportWeek.value='';resetTabs();enableOutputs();ensureTabSelects('p6');auto6();
      setStatus(`已加载 ${state.allTimes.length} 个共同时间点\n${state.allTimes[0]}  ～  ${state.allTimes[state.allTimes.length-1]}\n可切换“6 小时 / 24 小时 / 一周”报告。`,'ok');
    }catch(e){console.error(e);setStatus(`加载失败：${e.message||e}`,'err');}finally{els.loadBtn.disabled=false;}
  }

  function fileMeta(file,kind){return file?`${file.name}\n${(file.size/1024/1024).toFixed(2)} MB · ${kind}`:'未选择文件';}
  els.faceFile.addEventListener('change',async()=>{state.faceFile=els.faceFile.files[0]||null;els.faceMeta.textContent=fileMeta(state.faceFile,'面数据');if(state.faceFile){const k=await detectKindByHeader(state.faceFile);if(k&&k!=='face')setStatus('当前“面数据”文件看起来是监测点文件，请检查是否选反。','warn');}});
  els.pointFile.addEventListener('change',async()=>{state.pointFile=els.pointFile.files[0]||null;els.pointMeta.textContent=fileMeta(state.pointFile,'点数据');if(state.pointFile){const k=await detectKindByHeader(state.pointFile);if(k&&k!=='point')setStatus('当前“点数据”文件看起来是监测面文件，请检查是否选反。','warn');}});

  els.loadBtn.addEventListener('click',loadData);els.auto6.addEventListener('click',auto6);els.gen6.addEventListener('click',generate6);els.auto24.addEventListener('click',auto24);els.gen24.addEventListener('click',generate24);els.autoWeek.addEventListener('click',autoWeek);els.genWeek.addEventListener('click',generateWeek);
  els.copy6.addEventListener('click',()=>copyText('6h'));els.share6.addEventListener('click',()=>shareText('6h'));els.txt6.addEventListener('click',()=>downloadText('6h'));
  els.copy24.addEventListener('click',()=>copyText('24h'));els.share24.addEventListener('click',()=>shareText('24h'));els.txt24.addEventListener('click',()=>downloadText('24h'));els.xlsx24.addEventListener('click',downloadExcel24);
  els.copyWeek.addEventListener('click',()=>copyText('week'));els.shareWeek.addEventListener('click',()=>shareText('week'));els.txtWeek.addEventListener('click',()=>downloadText('week'));els.xlsxWeek.addEventListener('click',downloadExcelWeek);

  document.querySelectorAll('.tab-btn').forEach(btn=>btn.addEventListener('click',()=>{document.querySelectorAll('.tab-btn').forEach(x=>x.classList.toggle('active',x===btn));document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('active',p.id===btn.dataset.tab));if(state.allTimes.length){const wasReady=state.readyTabs[btn.dataset.tab];ensureTabSelects(btn.dataset.tab);if(!wasReady&&btn.dataset.tab==='p24')auto24();if(!wasReady&&btn.dataset.tab==='pweek')autoWeek();}}));

  if('serviceWorker'in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('./sw.js?v=2.0.0').catch(()=>{});enableOutputs();
})();
