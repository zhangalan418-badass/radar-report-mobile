/* 地基雷达监测报告生成器 V4.2
   逻辑按桌面版 wjz shibao_zhong.py 修正版迁移：
   短时数据 -> 6h；长时数据 -> 24h/7d；总累积位移从长时数据取值。
   V4.2 继承 V4.1 修正：修复下部边坡 pointRange 缺失导致“监测点（undefined）”；增加分组配置兜底校验；保留 V4 的 24h、Excel 与时间校验修正。 */

const GROUPS = [
  {name:"上部及后缘(全区变形峰值区)", faceNames:["M6","M7","M8"], faceRange:"M6-M8", pointNames:["S6","S7","S8"], pointRange:"S6-S8"},
  {name:"中部公路边坡", faceNames:["M1","M2","M3","M4","M5"], faceRange:"M1-M5", pointNames:["S1","S2","S3","S4","S5"], pointRange:"S1-S5"},
  {name:"下部边坡", faceNames:["M9","M10","M11","M12","M13","M14"], faceRange:"M9-M14", pointNames:["S9","S10","S11","S12","S13","S14"], pointRange:"S9-S14"}
];

// V4.1：防止分组显示范围字段缺失时在报告中出现 undefined。
function rangeFromNames(names){
  if(!Array.isArray(names) || names.length===0) return "";
  return names.length===1 ? names[0] : `${names[0]}-${names[names.length-1]}`;
}
for(const g of GROUPS){
  if(!g.faceRange) g.faceRange=rangeFromNames(g.faceNames);
  if(!g.pointRange) g.pointRange=rangeFromNames(g.pointNames);
}

const REQUIRED_FACE_NAMES = [...new Set(GROUPS.flatMap(g=>g.faceNames))];
const REQUIRED_POINT_NAMES = [...new Set(GROUPS.flatMap(g=>g.pointNames))];

function validateRequiredColumns(colMap, required, label){
  const missing=required.filter(n=>!(n in (colMap||{})));
  if(missing.length) throw new Error(`${label}缺少必要列：${missing.join("、")}`);
}

const state = {
  faceShort:null, pointShort:null, faceLong:null, pointLong:null,
  faceShortCols:null, pointShortCols:null, faceLongCols:null, pointLongCols:null,
  shortTimes:[], longTimes:[],
  files:{},
  last6:null, last24:null, last7:null
};

const $ = id => document.getElementById(id);

function toast(msg){
  const el=$("toast"); el.textContent=msg; el.classList.add("show");
  clearTimeout(toast._t); toast._t=setTimeout(()=>el.classList.remove("show"),2200);
}

function parseTime(s){
  if(s instanceof Date) return s.getTime();
  if(typeof s === "number" && Number.isFinite(s)){
    // Excel serial date.
    const d = XLSX.SSF.parse_date_code(s);
    if(d) return new Date(d.y,d.m-1,d.d,d.H||0,d.M||0,Math.floor(d.S||0)).getTime();
  }
  const str=String(s??"").trim();
  let m=str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?$/);
  if(!m) return NaN;
  return new Date(+m[1],+m[2]-1,+m[3],+m[4],+(m[5]||0),+(m[6]||0)).getTime();
}
function fmtFull(ms){
  const d=new Date(ms);
  const p=n=>String(n).padStart(2,"0");
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
function fmtDate(ms){
  const d=new Date(ms);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function addDays(ms,n){ return ms+n*86400000; }
function sameDate(ms,dateMs){
  const a=new Date(ms), b=new Date(dateMs);
  return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate();
}
function roundToHour(ms){
  const d=new Date(ms);
  if(d.getMinutes()>=30) d.setHours(d.getHours()+1);
  d.setMinutes(0,0,0);
  return d.getTime();
}
function to15Hour(ms){
  const d=new Date(ms); d.setHours(15,0,0,0); return d.getTime();
}
function fmtVal(v){
  const n=Number(v)||0;
  return n>=0 ? `+${n.toFixed(1)}` : n.toFixed(1);
}
function parseNumeric(v){
  if(v===null || v===undefined || v==="") return 0;
  if(typeof v==="number") return Number.isFinite(v)?v:0;
  const n=Number(String(v).replace(/,/g,"").trim());
  return Number.isFinite(n)?n:0;
}

function timeDisplayFromCell(cell){
  if(!cell) return "";
  if(cell.t==="d" && cell.v instanceof Date) return fmtFull(cell.v.getTime());
  if(typeof cell.v==="number"){
    const code = cell.z || "";
    // If formatted like a date/time, use parsed date code; otherwise return number.
    if(code && /[ymdhHs]/i.test(code)){
      const d=XLSX.SSF.parse_date_code(cell.v);
      if(d){
        const ms=new Date(d.y,d.m-1,d.d,d.H||0,d.M||0,Math.floor(d.S||0)).getTime();
        return fmtFull(ms);
      }
    }
  }
  return String(cell.w ?? cell.v ?? "").trim();
}

async function readWorkbook(file,isPoint){
  const buf=await file.arrayBuffer();
  const wb=XLSX.read(buf,{type:"array",cellDates:true,nodim:true});
  if(!wb.SheetNames.includes("位移")) throw new Error(`${file.name} 中没有“位移”工作表`);
  const ws=wb.Sheets["位移"];
  const range=XLSX.utils.decode_range(ws["!ref"]);
  const headers=[];
  for(let c=range.s.c;c<=range.e.c;c++){
    const cell=ws[XLSX.utils.encode_cell({r:range.s.r,c})];
    const h=cell?.v===null||cell?.v===undefined?null:String(cell.v).trim();
    headers[c]=h;
  }
  const colMap={};
  for(let c=1;c<headers.length;c++) if(headers[c]) colMap[headers[c]]=c;

  const data={};
  const startRow=(isPoint?4:1); // zero-based: Python row 5/2
  for(let r=Math.max(startRow,range.s.r); r<=range.e.r; r++){
    const timeCell=ws[XLSX.utils.encode_cell({r,c:range.s.c})];
    const rawTime=timeDisplayFromCell(timeCell);
    if(!rawTime) continue;
    const ts=String(rawTime).trim();
    if(isPoint && (ts==="X"||ts==="Y"||ts==="Z")) continue;
    const ms=parseTime(ts);
    if(!Number.isFinite(ms)) continue;
    const values={};
    for(const [name,c] of Object.entries(colMap)){
      const cell=ws[XLSX.utils.encode_cell({r,c})];
      const v=cell ? cell.v : 0;
      values[name]=v===null||v===undefined?0:parseNumeric(v);
    }
    const canonical=fmtFull(ms);
    data[canonical]=values;
  }
  const times=Object.keys(data).sort((a,b)=>parseTime(a)-parseTime(b));
  return {data, colMap, times};
}

async function loadAllData(){
  const required=["faceShort","pointShort","faceLong","pointLong"];
  if(required.some(k=>!state.files[k])){
    setStatus("请先选择全部 4 个数据文件。","err"); return;
  }
  try{
    setStatus("正在读取 4 个 Excel，请稍候……");
    const fdS=await readWorkbook(state.files.faceShort,false);
    const pdS=await readWorkbook(state.files.pointShort,true);
    const fdL=await readWorkbook(state.files.faceLong,false);
    const pdL=await readWorkbook(state.files.pointLong,true);

    validateRequiredColumns(fdS.colMap, REQUIRED_FACE_NAMES, "短时面数据");
    validateRequiredColumns(pdS.colMap, REQUIRED_POINT_NAMES, "短时点数据");
    validateRequiredColumns(fdL.colMap, REQUIRED_FACE_NAMES, "长时面数据");
    validateRequiredColumns(pdL.colMap, REQUIRED_POINT_NAMES, "长时点数据");

    state.faceShort=fdS.data; state.faceShortCols=fdS.colMap; state.shortTimes=fdS.times.filter(t=>pdS.data[t]);
    state.pointShort=pdS.data; state.pointShortCols=pdS.colMap;

    state.faceLong=fdL.data; state.faceLongCols=fdL.colMap; state.longTimes=fdL.times.filter(t=>pdL.data[t]);
    state.pointLong=pdL.data; state.pointLongCols=pdL.colMap;

    if(!state.shortTimes.length || !state.longTimes.length) throw new Error("短时或长时数据没有共同时间点");

    populateAllSelects();
    auto6h(); auto24h(); auto7d();
    $("reportArea").classList.remove("hidden");
    const ss=state.shortTimes, ls=state.longTimes;
    setStatus(`加载成功：短时 ${ss.length} 点（${ss[0]} ～ ${ss[ss.length-1]}）；长时 ${ls.length} 点（${ls[0]} ～ ${ls[ls.length-1]}）。`,"ok");
    toast("全部数据加载成功");
  }catch(e){
    console.error(e); setStatus("加载失败：" + e.message,"err");
  }
}

function setStatus(msg,cls=""){
  const el=$("status"); el.textContent=msg; el.className="status "+cls;
}

function wireFile(id,nameId,key){
  $(id).addEventListener("change",e=>{
    const f=e.target.files?.[0];
    if(f){state.files[key]=f; $(nameId).textContent=f.name;}
    else {delete state.files[key]; $(nameId).textContent="未选择";}
  });
}

function sortedNames(d){
  return Object.keys(d||{}).sort((a,b)=>{
    const na=(String(a).match(/\d+/)||["999"])[0];
    const nb=(String(b).match(/\d+/)||["999"])[0];
    return (+na)-(+nb);
  });
}

function computeItem(data,s,e,name){
  const sr=data?.[s], er=data?.[e];
  if(!sr || !er) return [null,null];
  const sv=parseNumeric(sr[name]), ev=parseNumeric(er[name]);
  return [ev-sv,ev];
}
function computeGroup(data,s,e,names){
  const sr=data?.[s], er=data?.[e];
  if(!sr||!er) return [null,null,null,null];
  const d={};
  for(const n of names){
    const sv=parseNumeric(sr[n]), ev=parseNumeric(er[n]);
    d[n]=[ev-sv,ev];
  }
  const keys=Object.keys(d);
  if(!keys.length) return [null,null,null,null];
  let mx=keys[0], mn=keys[0];
  for(const k of keys){ if(d[k][0]>d[mx][0]) mx=k; if(d[k][0]<d[mn][0]) mn=k; }
  return [d[mn][0],d[mx][0],mx,d[mx][1]];
}
function nearestTime(times,targetMs){
  if(!times.length) return null;
  let best=times[0], gap=Math.abs(parseTime(best)-targetMs);
  for(let i=1;i<times.length;i++){
    const g=Math.abs(parseTime(times[i])-targetMs);
    if(g<gap){gap=g;best=times[i];}
  }
  return best;
}

function findNearest(times,targetHour,targetMin=0,before=true,targetDateMs=null){
  let best=null;
  for(const t of times){
    const dt=new Date(parseTime(t));
    if(targetDateMs!==null && !sameDate(parseTime(t),targetDateMs)) continue;
    if(before && (dt.getHours()>targetHour || (dt.getHours()===targetHour && dt.getMinutes()>targetMin))) continue;
    if(!before && (dt.getHours()<targetHour || (dt.getHours()===targetHour && dt.getMinutes()<targetMin))) continue;
    best=t;
  }
  return best;
}
function nearestTimeFromData(data,timeKey){
  return nearestTime(Object.keys(data||{}),parseTime(timeKey));
}

function setSelectOptions(id,times){
  const el=$(id); el.innerHTML="";
  for(const t of times){
    const o=document.createElement("option"); o.value=t; o.textContent=t; el.appendChild(o);
  }
}
function setSelect(id,val){
  if(val!==null&&val!==undefined&&val!=="") $(id).value=val;
}
function getSelect(id){return $(id).value;}

function populateAllSelects(){
  ["s6","e6"].forEach(id=>setSelectOptions(id,state.shortTimes));
  ["s24","e24","ps24","pe24","s7","e7","ps7","pe7"].forEach(id=>setSelectOptions(id,state.longTimes));
}

function auto6h(){
  if(!state.shortTimes.length) return;
  const latestMs=parseTime(state.shortTimes[state.shortTimes.length-1]);
  const latest=new Date(latestMs);
  const sh=Math.floor(latest.getHours()/6)*6;
  const currentStart=new Date(latestMs);
  currentStart.setHours(sh,0,0,0);

  // V4：按真实 6 小时时间桶向前搜索，避免 V3 因只比较“小时数”而跨日期误选开始节点。
  for(let back=0; back<8; back++){
    const startMs=currentStart.getTime()-back*6*3600000;
    const endMs=startMs+6*3600000;
    const bucket=state.shortTimes.filter(t=>{
      const ms=parseTime(t);
      return ms>=startMs && ms<endMs && (back>0 || ms<=latestMs);
    });
    if(bucket.length>=2){
      setSelect("s6",bucket[0]);
      setSelect("e6",bucket[bucket.length-1]);
      return;
    }
  }

  // 极端缺数时仍给出可用范围，由用户手动调整。
  setSelect("s6",state.shortTimes[0]);
  setSelect("e6",state.shortTimes[state.shortTimes.length-1]);
}
function auto24h(){
  if(!state.longTimes.length) return;
  const latest=parseTime(state.longTimes[state.longTimes.length-1]);
  const todayStart=new Date(latest); todayStart.setHours(0,0,0,0);
  let end=findNearest(state.longTimes,15,0,true,todayStart.getTime());
  if(!end) end=state.longTimes[state.longTimes.length-1];

  const yesterday=addDays(todayStart.getTime(),-1);
  let start=findNearest(state.longTimes,15,0,true,yesterday);
  if(!start) start=state.longTimes[0];

  setSelect("s24",start); setSelect("e24",end);

  const prevDay=addDays(yesterday,-1);
  let ps=findNearest(state.longTimes,15,0,true,prevDay);
  if(!ps) ps=state.longTimes[0];
  setSelect("ps24",ps); setSelect("pe24",start);
}

function weekday(ms){ return new Date(ms).getDay()===0?6:new Date(ms).getDay()-1; } // Monday=0
function auto7d(){
  if(!state.longTimes.length) return;
  const latest=parseTime(state.longTimes[state.longTimes.length-1]);
  let found=null;
  for(let i=0;i<14;i++){
    const day=addDays(latest,-i);
    if(weekday(day)===4){ found=day;break; } // Friday
  }
  if(found===null) return;
  const end=findNearest(state.longTimes,15,0,true,found);
  const startDate=addDays(found,-7);
  const start=findNearest(state.longTimes,15,0,true,startDate) || state.longTimes[0];
  const psDate=addDays(startDate,-7);
  const ps=findNearest(state.longTimes,15,0,true,psDate) || state.longTimes[0];
  const pe=start;
  setSelect("s7",start);setSelect("e7",end||state.longTimes[state.longTimes.length-1]);
  setSelect("ps7",ps);setSelect("pe7",pe);
}

function getTotalValue(data,timeKey,name){
  const exact=data?.[timeKey];
  if(exact) return parseNumeric(exact[name]);
  const nearest=nearestTimeFromData(data,timeKey);
  return nearest ? parseNumeric(data[nearest]?.[name]) : 0;
}

function generateReportText(ss,es,label,dispFace,dispPoint,totalFace,totalPoint,compare=null,displayStart=null,displayEnd=null,compareLabel="与上周相比"){
  const ds=displayStart??ss, de=displayEnd??es;
  const lines=[`地基雷达${label}监测结果（${ds} 至 ${de}）`];
  for(const gi of GROUPS){
    const [fmin,fmax,fn]=computeGroup(dispFace,ss,es,gi.faceNames);
    const [pmin,pmax,pn]=computeGroup(dispPoint,ss,es,gi.pointNames);
    if(fmin===null||pmin===null) continue;
    const ft=getTotalValue(totalFace,es,fn);
    const pt=getTotalValue(totalPoint,es,pn);
    let line=`- ${gi.name}：监测面（${gi.faceRange}）${label}时段位移${fmtVal(fmin)}~${fmtVal(fmax)}mm，最大值位于${fn}，截止期末总累积位移${fmtVal(ft)}mm；监测点（${gi.pointRange}）${label}时段位移${fmtVal(pmin)}~${fmtVal(pmax)}mm，最大值位于${pn}，截止期末总累积位移${fmtVal(pt)}mm。`;
    if(compare){
      const [ps,pe]=compare;
      const fdelta={}, pdelta={};
      for(const n of gi.faceNames){
        const [a]=computeItem(dispFace,ss,es,n), [b]=computeItem(dispFace,ps,pe,n);
        if(a!==null&&b!==null) fdelta[n]=a-b;
      }
      for(const n of gi.pointNames){
        const [a]=computeItem(dispPoint,ss,es,n), [b]=computeItem(dispPoint,ps,pe,n);
        if(a!==null&&b!==null) pdelta[n]=a-b;
      }
      const fv=Object.values(fdelta), pv=Object.values(pdelta);
      if(fv.length){
        const fminD=Math.min(...fv), fmaxD=Math.max(...fv);
        if(pv.length){
          const pminD=Math.min(...pv), pmaxD=Math.max(...pv);
          line += ` ${compareLabel}，监测面位移变化${fmtVal(fminD)}~${fmtVal(fmaxD)}mm，监测点位移变化${fmtVal(pminD)}~${fmtVal(pmaxD)}mm。`;
        }else{
          line += ` ${compareLabel}，监测面位移变化${fmtVal(fminD)}~${fmtVal(fmaxD)}mm。`;
        }
      }
    }
    lines.push(line);
  }
  return lines.length>1?lines.join("\n"):null;
}

function makeWorkbook(ss,es,ps,pe,hasPrev,faceData,pointData,faceCols,pointCols){
  const sd=new Date(parseTime(ss)), ed=new Date(parseTime(es));
  const dl=`${sd.getMonth()+1}/${sd.getDate()}-${ed.getMonth()+1}/${ed.getDate()}`;
  const headers=["监测点/面",`${dl}位移量(mm)`];
  if(hasPrev){
    const p1=new Date(parseTime(ps)), p2=new Date(parseTime(pe));
    headers.push(`相对于${p1.getMonth()+1}/${p1.getDate()}-${p2.getMonth()+1}/${p2.getDate()}变化量(mm)`);
  }
  headers.push("总累积位移量(mm)");

  const rowsFace=[headers];
  for(const n of sortedNames(faceCols)){
    if(n==="W1") continue;
    const [da]=computeItem(faceData,ss,es,n); if(da===null) continue;
    const [db]=hasPrev?computeItem(faceData,ps,pe,n):[null,null];
    const nearest=nearestTimeFromData(faceData,es);
    const total=nearest ? parseNumeric(faceData[nearest]?.[n]) : 0;
    const row=[n,round1(da)];
    if(hasPrev) row.push(db!==null?round1(da-db):"");
    row.push(round1(total));
    rowsFace.push(row);
  }
  const rowsPoint=[headers];
  for(const n of sortedNames(pointCols)){
    const [da]=computeItem(pointData,ss,es,n); if(da===null) continue;
    const [db]=hasPrev?computeItem(pointData,ps,pe,n):[null,null];
    const nearest=nearestTimeFromData(pointData,es);
    const total=nearest ? parseNumeric(pointData[nearest]?.[n]) : 0;
    const row=[n,round1(da)];
    if(hasPrev) row.push(db!==null?round1(da-db):"");
    row.push(round1(total));
    rowsPoint.push(row);
  }
  const wb=XLSX.utils.book_new();
  const ws1=XLSX.utils.aoa_to_sheet(rowsFace);
  const ws2=XLSX.utils.aoa_to_sheet(rowsPoint);
  XLSX.utils.book_append_sheet(wb,ws1,"面");
  XLSX.utils.book_append_sheet(wb,ws2,"点");
  return wb;
}
function round1(v){return Math.round((Number(v)||0)*10)/10;}
function downloadWorkbook(wb,filename){
  XLSX.writeFile(wb,filename,{bookType:"xlsx"});
  toast("Excel 已生成");
}
function downloadTxt(text,filename){
  const blob=new Blob(["\uFEFF"+text+"\n"],{type:"text/plain;charset=utf-8"});
  const url=URL.createObjectURL(blob); const a=document.createElement("a");
  a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000); toast("TXT 已生成");
}
async function copyReport(text){
  try{await navigator.clipboard.writeText(text);toast("报告已复制");}
  catch{const t=document.createElement("textarea");t.value=text;document.body.appendChild(t);t.select();document.execCommand("copy");t.remove();toast("报告已复制");}
}
async function shareReport(text){
  if(navigator.share){try{await navigator.share({title:"地基雷达监测报告",text});}catch{}}
  else copyReport(text);
}

function requireData(kind){
  if(kind==="6h" && !state.faceShort) {toast("请先加载全部数据"); return false;}
  if((kind==="24h"||kind==="7d") && !state.faceLong) {toast("请先加载全部数据"); return false;}
  return true;
}
function showResult(id,text,meta){
  $(id).value=text||"";
  $(meta).textContent=text?meta:"";
  if(text) $(id).scrollTop=0;
}
function gen6(){
  if(!requireData("6h")) return;
  const ss=getSelect("s6"), es=getSelect("e6");
  if(!ss||!es){toast("请选择起止时间");return;}
  if(parseTime(ss)>=parseTime(es)){toast("开始时间必须早于结束时间");return;}
  const text=generateReportText(ss,es,"6小时",state.faceShort,state.pointShort,state.faceLong,state.pointLong,null,fmtFull(roundToHour(parseTime(ss))),fmtFull(roundToHour(parseTime(es))));
  if(text){showResult("r6",text,"meta6",);$("meta6").textContent=`${fmtFull(roundToHour(parseTime(ss)))} 至 ${fmtFull(roundToHour(parseTime(es)))}`;state.last6={text,ss,es};}
  else toast("生成6小时报告失败");
}
function gen24(){
  if(!requireData("24h")) return;
  const ss=getSelect("s24"),es=getSelect("e24"),ps=getSelect("ps24"),pe=getSelect("pe24");
  if(!ss||!es){toast("请选择当前周期");return;}
  if(parseTime(ss)>=parseTime(es)){toast("开始时间必须早于结束时间");return;}
  const hasPrev=!!(ps&&pe&&parseTime(ps)<parseTime(pe));
  // V4：24 小时日报正文只报告当前 24h，不输出任何“与上周相比”文字。
  // 对比周期仍保留，仅供“24h Excel 对比表”导出使用。
  const text=generateReportText(ss,es,"24小时",state.faceLong,state.pointLong,state.faceLong,state.pointLong,null,fmtFull(to15Hour(parseTime(ss))),fmtFull(to15Hour(parseTime(es))));
  if(text){$("r24").value=text;$("meta24").textContent=`${fmtFull(to15Hour(parseTime(ss)))} 至 ${fmtFull(to15Hour(parseTime(es)))}`;state.last24={text,ss,es,ps,pe,hasPrev};}
  else toast("生成24小时报告失败");
}
function gen7(){
  if(!requireData("7d")) return;
  const ss=getSelect("s7"),es=getSelect("e7"),ps=getSelect("ps7"),pe=getSelect("pe7");
  if(!ss||!es){toast("请选择当前周期");return;}
  if(parseTime(ss)>=parseTime(es)){toast("开始时间必须早于结束时间");return;}
  const hasPrev=!!(ps&&pe&&parseTime(ps)<parseTime(pe));
  const text=generateReportText(ss,es,"7天",state.faceLong,state.pointLong,state.faceLong,state.pointLong,hasPrev?[ps,pe]:null,fmtFull(to15Hour(parseTime(ss))),fmtFull(to15Hour(parseTime(es))),"与上周相比");
  if(text){$("r7").value=text;$("meta7").textContent=`${fmtFull(to15Hour(parseTime(ss)))} 至 ${fmtFull(to15Hour(parseTime(es)))}`;state.last7={text,ss,es,ps,pe,hasPrev};}
  else toast("生成周报失败");
}

function resetApp(){
  Object.assign(state,{faceShort:null,pointShort:null,faceLong:null,pointLong:null,faceShortCols:null,pointShortCols:null,faceLongCols:null,pointLongCols:null,shortTimes:[],longTimes:[],files:{},last6:null,last24:null,last7:null});
  ["faceShort","pointShort","faceLong","pointLong"].forEach(id=>$(id).value="");
  [["faceShort","faceShortName"],["pointShort","pointShortName"],["faceLong","faceLongName"],["pointLong","pointLongName"]].forEach(([a,b])=>$(b).textContent="未选择");
  $("reportArea").classList.add("hidden");
  setStatus("请选择全部 4 个 Excel 文件。");
  toast("已清空");
}

document.addEventListener("DOMContentLoaded",()=>{
  wireFile("faceShort","faceShortName","faceShort");
  wireFile("pointShort","pointShortName","pointShort");
  wireFile("faceLong","faceLongName","faceLong");
  wireFile("pointLong","pointLongName","pointLong");
  $("loadBtn").onclick=loadAllData;
  $("resetBtn").onclick=resetApp;

  document.querySelectorAll(".tab").forEach(btn=>{
    btn.onclick=()=>{
      document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(x=>x.classList.remove("active"));
      btn.classList.add("active"); $(`tab${btn.dataset.tab}`).classList.add("active");
    }
  });
  $("auto6").onclick=auto6h; $("gen6").onclick=gen6;
  $("auto24").onclick=auto24h; $("gen24").onclick=gen24;
  $("auto7").onclick=auto7d; $("gen7").onclick=gen7;

  $("copy6").onclick=()=>state.last6&&copyReport(state.last6.text);
  $("share6").onclick=()=>state.last6&&shareReport(state.last6.text);
  $("txt6").onclick=()=>state.last6&&downloadTxt(state.last6.text,"6h_report.txt");

  $("copy24").onclick=()=>state.last24&&copyReport(state.last24.text);
  $("share24").onclick=()=>state.last24&&shareReport(state.last24.text);
  $("txt24").onclick=()=>state.last24&&downloadTxt(state.last24.text,"24h_report.txt");
  $("xlsx24").onclick=()=>{
    if(!state.last24){toast("请先生成24小时报告");return;}
    const x=state.last24;
    downloadWorkbook(makeWorkbook(x.ss,x.es,x.ps,x.pe,x.hasPrev,state.faceLong,state.pointLong,state.faceLongCols,state.pointLongCols),"24h对比表.xlsx");
  };

  $("copy7").onclick=()=>state.last7&&copyReport(state.last7.text);
  $("share7").onclick=()=>state.last7&&shareReport(state.last7.text);
  $("txt7").onclick=()=>state.last7&&downloadTxt(state.last7.text,"7d_report.txt");
  $("xlsx7").onclick=()=>{
    if(!state.last7){toast("请先生成周报");return;}
    const x=state.last7;
    downloadWorkbook(makeWorkbook(x.ss,x.es,x.ps,x.pe,x.hasPrev,state.faceLong,state.pointLong,state.faceLongCols,state.pointLongCols),"周报对比表.xlsx");
  };

  if("serviceWorker" in navigator){
    // V4.2：清理旧版 Service Worker，再注册唯一版本文件，避免 V2/V3/V4 缓存反向覆盖新版页面。
    navigator.serviceWorker.getRegistrations().then(async regs=>{
      for(const reg of regs){
        const u=(reg.active||reg.waiting||reg.installing)?.scriptURL||"";
        if(!u.endsWith("/sw-v4.2.js")) await reg.unregister();
      }
      return navigator.serviceWorker.register("./sw-v4.2.js");
    }).then(()=>{$("swState").textContent="V4.2 离线缓存已启用";})
      .catch(()=>{$("swState").textContent="";});
  }
});
