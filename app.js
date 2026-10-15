(function(){
"use strict";
const C={gray:"#C0C0C0",purple:"#550E78",red:"#C30C3E",green:"#61BB2B",dot:"#7D7D7D",bg:"#FBFBFB",black:"#000000",white:"#FFFFFF",note:"#595959"};
const FONT='Carlito, Calibri, "Liberation Sans", Arial, sans-serif';
const $=id=>document.getElementById(id);
let wb=null, grid=null, model=null, lastPrims=null, lastSize=null;

/* ---------- leitura ---------- */
function parseNum(v){
  if(v==null||v==="") return null;
  if(typeof v==="number") return isFinite(v)?v:null;
  let s=String(v).trim(); if(!s||s==="-") return null;
  let neg=false;
  if(/^\(.*\)$/.test(s)){neg=true;s=s.slice(1,-1);}
  s=s.replace(/[R$\s%]/g,"");
  if(/,\d+$/.test(s)) s=s.replace(/\./g,"").replace(",",".");
  else if(/^-?\d{1,3}(\.\d{3})+$/.test(s)) s=s.replace(/\./g,"");
  const n=Number(s); if(!isFinite(n)) return null;
  return neg?-n:n;
}
function loadFile(file){
  $("fileName").textContent=file.name;
  const r=new FileReader();
  r.onload=e=>{
    try{
      wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
      const sel=$("sheet"); sel.innerHTML="";
      wb.SheetNames.forEach((n,i)=>sel.add(new Option(n,i)));
      $("sheetWrap").hidden=wb.SheetNames.length<2;
      readSheet(0);
    }catch(err){ showError("Não foi possível ler o arquivo. Confira se é um .xlsx ou .csv válido."); }
  };
  r.readAsArrayBuffer(file);
}
function showError(t){ $("stage").innerHTML='<div class="empty"><p><b>Arquivo não reconhecido</b>'+esc(t)+'</p></div>'; ["cfg","cfgSteps","cfgText","cfgExport"].forEach(id=>$(id).hidden=true); }
function esc(s){return String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));}

function readSheet(i){
  const ws=wb.Sheets[wb.SheetNames[i]];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});
  // cabeçalho = primeira linha com 3+ textos
  let h=rows.findIndex(r=>r.filter(c=>typeof c==="string"&&c.trim()).length>=3);
  if(h<0){showError("Não encontrei uma linha de cabeçalho com os nomes das empresas.");return;}
  const header=rows[h].map(c=>c==null?"":String(c).trim());
  const cols=[]; header.forEach((name,j)=>{ if(j>0&&name) cols.push({j,name}); });
  const data=[];
  rows.slice(h+1).forEach(r=>{
    const label=r[0]==null?"":String(r[0]).trim(); if(!label) return;
    const vals={}; let cnt=0;
    cols.forEach(c=>{const n=parseNum(r[c.j]); vals[c.j]=n; if(n!=null) cnt++;});
    if(cnt) data.push({label,vals});
  });
  if(data.length<2){showError("A planilha precisa de pelo menos duas linhas com valores: a base e o resultado.");return;}
  grid={cols,data};
  // agregados: "Oniz" quando existe "ONIZ - RS"
  const low=cols.map(c=>c.name.toLowerCase());
  cols.forEach(c=>{const p=c.name.toLowerCase()+" - "; c.agg=low.some(n=>n.startsWith(p));});
  const totIdx=cols.findIndex(c=>/braveo|total|consolidado/i.test(c.name));
  // selects
  const b=$("baseRow"), e=$("endRow"), t=$("totalCol");
  b.innerHTML=""; e.innerHTML=""; t.innerHTML="";
  data.forEach((d,k)=>{b.add(new Option(d.label,k)); e.add(new Option(d.label,k));});
  t.add(new Option("Nenhuma (somar os degraus)",-1));
  cols.forEach((c,k)=>t.add(new Option(c.name,k)));
  let bi=data.findIndex(d=>/6\s*\+\s*6|3\s*\+\s*9|base|orç/i.test(d.label)&&!/varia/i.test(d.label)); if(bi<0) bi=0;
  let ei=data.findIndex((d,k)=>k!==bi&&/fcst|forecast|real|atual/i.test(d.label)&&!/varia/i.test(d.label)); if(ei<0) ei=bi===0?1:0;
  b.value=bi; e.value=ei; t.value=totIdx;
  $("baseLabel").value=data[bi].label.replace(/\s*\+\s*/g,"+");
  $("endLabel").value=data[ei].label;
  // degraus
  const box=$("steps"); box.innerHTML="";
  cols.forEach((c,k)=>{
    if(k===totIdx) return;
    const id="st"+k;
    const row=document.createElement("label"); row.className="step"+(c.agg?" agg":""); row.htmlFor=id;
    row.innerHTML='<input type="checkbox" id="'+id+'" data-k="'+k+'"'+(c.agg?"":" checked")+'><span class="n">'+esc(c.name)+'</span><span class="v"></span>';
    box.appendChild(row);
  });
  ["cfg","cfgSteps","cfgText","cfgExport"].forEach(id=>$(id).hidden=false);
  $("hint").hidden=false;
  update();
}

/* ---------- modelo ---------- */
function build(){
  const bi=+$("baseRow").value, ei=+$("endRow").value, ti=+$("totalCol").value;
  const base=grid.data[bi], end=grid.data[ei];
  const steps=[];
  document.querySelectorAll("#steps input").forEach(inp=>{
    const k=+inp.dataset.k, c=grid.cols[k];
    const d=(end.vals[c.j]||0)-(base.vals[c.j]||0);
    const v=inp.parentElement.querySelector(".v"); v.textContent=fmtSigned(d,currentUnit()); v.classList.toggle("neg",d<0);
    if(inp.checked) steps.push({name:c.name,d});
  });
  let start,finish;
  if(ti>=0){ const c=grid.cols[ti]; start=base.vals[c.j]||0; finish=end.vals[c.j]||0; }
  else{
    start=0; finish=0;
    document.querySelectorAll("#steps input:checked").forEach(inp=>{const c=grid.cols[+inp.dataset.k]; start+=base.vals[c.j]||0; finish+=end.vals[c.j]||0;});
  }
  return {start,finish,steps};
}
function currentUnit(){return $("unit").value;}
function nf(v,unit){
  const dec=unit==="mi"?1:0, x=unit==="mi"?v/1000:v;
  const r=Number(x.toFixed(dec));
  return {txt:new Intl.NumberFormat("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec}).format(Math.abs(r)),neg:r<0,zero:r===0};
}
function fmtPar(v,unit){const f=nf(v,unit);return f.neg?"("+f.txt+")":f.txt;}
function fmtSigned(v,unit){const f=nf(v,unit);return f.zero?f.txt:(f.neg?"-":"+")+f.txt;}

/* ---------- desenho (primitivas -> SVG e Canvas) ---------- */
function splitLabel(s){
  s=String(s).trim();
  if(s.includes(" - ")){const i=s.indexOf(" - ");return [s.slice(0,i),s.slice(i+3)];}
  const i=s.indexOf(" "); return i>0?[s.slice(0,i),s.slice(i+1)]:[s];
}
function layout(m,unit,baseLabel,endLabel,note){
  const n=m.steps.length, slot=52, x0=14, bw=42, tw=48, Y0=330, TOP=80, MINH=16;
  const W=x0*2+4+slot*(n+2), H=392;
  const cx=i=>x0+slot*i+slot/2;
  const levels=[m.start]; let cum=m.start; m.steps.forEach(s=>{cum+=s.d;levels.push(cum);}); levels.push(m.finish);
  const mx=Math.max(...levels), mn=Math.min(...levels);
  let R=mx-mn; if(R<=0) R=Math.max(Math.abs(mx)*0.05,1);
  let B=Math.max(0,mn-R);
  const truncated=B>0;
  const k=(Y0-TOP)/((mx-B)||1);
  const y=v=>Y0-(v-B)*k;
  const P=[];
  P.push({t:"rect",x:0,y:0,w:W,h:H,fill:C.bg});
  const ys=y(m.start);
  P.push({t:"rect",x:cx(0)-tw/2,y:ys,w:tw,h:Y0-ys,fill:C.gray});
  P.push({t:"text",x:cx(0),y:ys-7,s:15,fill:C.black,a:"middle",text:fmtPar(m.start,unit)});
  cum=m.start; let px=cx(0)+tw/2;
  m.steps.forEach((s,i0)=>{
    const i=i0+1, nw=cum+s.d;
    let top=y(Math.max(cum,nw)), bot=y(Math.min(cum,nw)), h=bot-top;
    if(h<MINH){const mid=(top+bot)/2; top=mid-MINH/2; h=MINH;}
    const xl=cx(i)-bw/2;
    P.push({t:"line",x1:px,y1:y(cum),x2:xl,y2:y(cum),stroke:C.dot,w:1.2,dash:[2,2]});
    P.push({t:"rect",x:xl,y:top,w:bw,h:h,fill:s.d<0?C.red:C.green});
    P.push({t:"text",x:cx(i),y:top+h/2+4.5,s:13,fill:C.white,a:"middle",text:fmtPar(s.d,unit)});
    cum=nw; px=xl+bw;
  });
  const ie=n+1, ye=y(m.finish), xe=cx(ie)-tw/2;
  P.push({t:"line",x1:px,y1:y(cum),x2:xe,y2:y(cum),stroke:C.dot,w:1.2,dash:[2,2]});
  P.push({t:"rect",x:xe,y:ye,w:tw,h:Y0-ye,fill:C.purple});
  if(Y0-ye>28) P.push({t:"text",x:cx(ie),y:ye+20,s:13,fill:C.white,a:"middle",text:fmtPar(m.finish,unit)});
  else P.push({t:"text",x:cx(ie),y:ye-20,s:13,fill:C.black,a:"middle",text:fmtPar(m.finish,unit)});
  // seta + elipse
  const ya=30, tip=ye-(Y0-ye>28?1:24);
  P.push({t:"path",pts:[[cx(0),ys-26],[cx(0),ya],[cx(ie),ya],[cx(ie),tip-14]],stroke:C.black,w:1.5});
  P.push({t:"poly",pts:[[cx(ie)-7,tip-15],[cx(ie)+7,tip-15],[cx(ie),tip-1]],fill:C.black});
  const delta=m.finish-m.start, pct=m.start?delta/m.start:0;
  const t1=fmtSigned(delta,unit);
  const t2="("+(pct>0?"+":pct<0?"-":"")+new Intl.NumberFormat("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1}).format(Math.abs(pct*100))+"%)";
  const midx=(cx(0)+cx(ie))/2, rx=Math.max(52,Math.max(t1.length,t2.length)*4.6+16);
  P.push({t:"ellipse",cx:midx,cy:ya,rx,ry:21,fill:C.white,stroke:C.black,w:1.5});
  P.push({t:"text",x:midx,y:ya-3,s:15,fill:C.black,a:"middle",text:t1});
  P.push({t:"text",x:midx,y:ya+14,s:14,fill:C.black,a:"middle",text:t2});
  // eixo e rótulos
  P.push({t:"line",x1:8,y1:Y0,x2:W-8,y2:Y0,stroke:C.black,w:1});
  const labels=[baseLabel].concat(m.steps.map(s=>s.name)).concat([endLabel]);
  labels.forEach((lb,i)=>{const parts=splitLabel(lb);parts.forEach((p,j)=>P.push({t:"text",x:cx(i),y:Y0+20+16*j,s:13,fill:C.black,a:"middle",text:p}));});
  const foot=note!=null?note:("R$ "+(unit==="mi"?"milhões":"mil")+(truncated?" · eixo truncado nas barras de total":""));
  if(foot) P.push({t:"text",x:10,y:H-6,s:11,fill:C.note,a:"start",text:foot});
  return {P,W,H,truncated};
}
function toSVG(L){
  const o=['<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 '+L.W+' '+L.H+'" width="'+L.W+'" height="'+L.H+'" role="img" aria-label="Gráfico ponte">'];
  L.P.forEach(p=>{
    if(p.t==="rect") o.push('<rect x="'+p.x.toFixed(1)+'" y="'+p.y.toFixed(1)+'" width="'+p.w.toFixed(1)+'" height="'+p.h.toFixed(1)+'" fill="'+p.fill+'"/>');
    else if(p.t==="line") o.push('<line x1="'+p.x1.toFixed(1)+'" y1="'+p.y1.toFixed(1)+'" x2="'+p.x2.toFixed(1)+'" y2="'+p.y2.toFixed(1)+'" stroke="'+p.stroke+'" stroke-width="'+p.w+'"'+(p.dash?' stroke-dasharray="'+p.dash.join(" ")+'"':'')+'/>');
    else if(p.t==="path") o.push('<path d="M'+p.pts.map(q=>q[0].toFixed(1)+","+q[1].toFixed(1)).join(" L")+'" fill="none" stroke="'+p.stroke+'" stroke-width="'+p.w+'"/>');
    else if(p.t==="poly") o.push('<polygon points="'+p.pts.map(q=>q[0].toFixed(1)+","+q[1].toFixed(1)).join(" ")+'" fill="'+p.fill+'"/>');
    else if(p.t==="ellipse") o.push('<ellipse cx="'+p.cx.toFixed(1)+'" cy="'+p.cy+'" rx="'+p.rx.toFixed(1)+'" ry="'+p.ry+'" fill="'+p.fill+'" stroke="'+p.stroke+'" stroke-width="'+p.w+'"/>');
    else if(p.t==="text") o.push('<text x="'+p.x.toFixed(1)+'" y="'+p.y.toFixed(1)+'" text-anchor="'+p.a+'" font-family=\''+FONT+'\' font-size="'+p.s+'" fill="'+p.fill+'">'+esc(p.text)+'</text>');
  });
  o.push("</svg>"); return o.join("");
}
function toCanvas(L,scale){
  const cv=document.createElement("canvas"); cv.width=L.W*scale; cv.height=L.H*scale;
  const g=cv.getContext("2d"); g.scale(scale,scale);
  L.P.forEach(p=>{
    g.save();
    if(p.t==="rect"){g.fillStyle=p.fill;g.fillRect(p.x,p.y,p.w,p.h);}
    else if(p.t==="line"){g.strokeStyle=p.stroke;g.lineWidth=p.w;g.setLineDash(p.dash||[]);g.beginPath();g.moveTo(p.x1,p.y1);g.lineTo(p.x2,p.y2);g.stroke();}
    else if(p.t==="path"){g.strokeStyle=p.stroke;g.lineWidth=p.w;g.beginPath();p.pts.forEach((q,i)=>i?g.lineTo(q[0],q[1]):g.moveTo(q[0],q[1]));g.stroke();}
    else if(p.t==="poly"){g.fillStyle=p.fill;g.beginPath();p.pts.forEach((q,i)=>i?g.lineTo(q[0],q[1]):g.moveTo(q[0],q[1]));g.closePath();g.fill();}
    else if(p.t==="ellipse"){g.beginPath();g.ellipse(p.cx,p.cy,p.rx,p.ry,0,0,Math.PI*2);g.fillStyle=p.fill;g.fill();g.strokeStyle=p.stroke;g.lineWidth=p.w;g.stroke();}
    else if(p.t==="text"){g.fillStyle=p.fill;g.font=p.s+"px "+FONT;g.textAlign=p.a==="middle"?"center":p.a;g.textBaseline="alphabetic";g.fillText(p.text,p.x,p.y);}
    g.restore();
  });
  return cv;
}

/* ---------- atualização ---------- */
let noteTouched=false;
function update(){
  if(!grid) return;
  const unit=currentUnit();
  model=build();
  if(!model.steps.length){ $("stage").innerHTML='<div class="empty"><p><b>Nenhum degrau marcado</b>Marque ao menos uma empresa na lista de degraus.</p></div>'; lastPrims=null; return; }
  // checagem de fechamento
  const sum=model.steps.reduce((a,s)=>a+s.d,0), gap=(model.finish-model.start)-sum, ck=$("check");
  ck.hidden=false;
  if(Math.abs(gap)<0.5){ck.className="check ok";ck.textContent="Os degraus fecham com a diferença entre as barras ("+fmtSigned(sum,unit)+").";}
  else{ck.className="check bad";ck.textContent="Os degraus somam "+fmtSigned(sum,unit)+", mas a diferença entre as barras é "+fmtSigned(model.finish-model.start,unit)+". Falta "+fmtSigned(gap,unit)+": revise as colunas marcadas ou a coluna do total.";}
  const note=noteTouched?$("note").value:null;
  const L=layout(model,unit,$("baseLabel").value,$("endLabel").value,note);
  if(!noteTouched) $("note").value="R$ "+(unit==="mi"?"milhões":"mil")+(L.truncated?" · eixo truncado nas barras de total":"");
  lastPrims=L;
  $("stage").innerHTML=toSVG(L);
}

/* ---------- exportação ---------- */
let dl=null;
if(window.claude&&typeof window.claude.use==="function"){ window.claude.use("downloads").then(x=>{dl=x;}).catch(()=>{}); }
async function save(filename,blob){
  const msg=$("exportMsg"); msg.textContent="";
  if(dl){
    try{await dl.save({filename,data:blob}); msg.textContent="Arquivo salvo.";}
    catch(e){ msg.textContent=e&&e.code==="declined"?"Download cancelado.":"Não foi possível salvar o arquivo aqui."; }
    return;
  }
  if(window.claude){ msg.textContent="Download indisponível nesta visualização."; return; }
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href),2000); msg.textContent="Arquivo salvo.";
}
function fname(ext){return "grafico_ponte_"+($("endLabel").value||"fcst").replace(/[^\w]+/g,"_").replace(/^_|_$/g,"").toLowerCase()+"."+ext;}
$("btnPng").addEventListener("click",async()=>{
  if(!lastPrims) return;
  try{ await document.fonts.load("15px Carlito"); }catch(e){}
  const cv=toCanvas(lastPrims,3);
  cv.toBlob(b=>{ if(b) save(fname("png"),b); },"image/png");
});
$("btnSvg").addEventListener("click",()=>{
  if(!lastPrims) return;
  save(fname("svg"),new Blob([toSVG(lastPrims)],{type:"image/svg+xml"}));
});

/* ---------- eventos ---------- */
$("file").addEventListener("change",e=>{const f=e.target.files[0]; if(f) loadFile(f);});
const drop=$("drop");
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("over");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("over");}));
drop.addEventListener("drop",e=>{const f=e.dataTransfer.files[0]; if(f) loadFile(f);});
$("sheet").addEventListener("change",e=>readSheet(+e.target.value));
$("baseRow").addEventListener("change",()=>{$("baseLabel").value=grid.data[+$("baseRow").value].label.replace(/\s*\+\s*/g,"+");update();});
$("endRow").addEventListener("change",()=>{$("endLabel").value=grid.data[+$("endRow").value].label;update();});
["totalCol","unit"].forEach(id=>$(id).addEventListener("change",update));
["baseLabel","endLabel"].forEach(id=>$(id).addEventListener("input",update));
$("note").addEventListener("input",()=>{noteTouched=true;update();});
$("steps").addEventListener("change",update);
if(document.fonts&&document.fonts.ready) document.fonts.ready.then(()=>{ if(grid) update(); });
})();
