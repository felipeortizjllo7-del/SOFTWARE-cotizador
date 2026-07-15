# -*- coding: utf-8 -*-
"""Genera un cotizador HTML autonomo (self-contained) con datos embebidos."""
import json, base64, os, re

PROJ = r"C:\Users\felip\OneDrive\Escritorio\CLAUDE\CotizadorInnoba"
# version (fuente unica: la constante VERSION de cotizador_innoba.py)
_src = open(os.path.join(PROJ, "cotizador_innoba.py"), encoding="utf-8").read()
VERSION = re.search(r'VERSION\s*=\s*"([^"]+)"', _src).group(1)
precios = json.load(open(os.path.join(PROJ, "precios_2026.json"), encoding="utf-8"))
desc = json.load(open(os.path.join(PROJ, "descripciones_tours.json"), encoding="utf-8"))
try:
    clientes = json.load(open(os.path.join(PROJ, "clientes.json"), encoding="utf-8"))
except Exception:
    clientes = []
with open(os.path.join(PROJ, "logo_innoba.png"), "rb") as f:
    logo_b64 = base64.b64encode(f.read()).decode()
with open(os.path.join(PROJ, "app_icono.png"), "rb") as f:
    icon_b64 = base64.b64encode(f.read()).decode()

HTML = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Cotizador INNOBA Colombia DMC</title>
<link rel="icon" href="data:image/png;base64,__ICON__">
<style>
:root{--navy:#013984;--navy2:#00285f;--blue:#1466c7;--cyan:#2e8be6;--bg:#eef3fa;
--card:#fff;--card2:#f4f8fd;--text:#16233d;--muted:#64748b;--green:#1e9e5a;--line:#d7e1ef;--red:#c0392b}
*{box-sizing:border-box}
body{margin:0;font-family:'Segoe UI',system-ui,Arial,sans-serif;background:var(--bg);color:var(--text)}
button{font-family:inherit;cursor:pointer}
.hidden{display:none!important}
header{background:var(--card);display:flex;align-items:center;gap:14px;padding:12px 20px;box-shadow:0 1px 4px #0001}
header img{height:46px}
header .t1{color:var(--navy);font-size:20px;font-weight:700}
header .t2{color:var(--muted);font-size:12px}
header .sp{flex:1}
.btn{border:none;border-radius:10px;padding:10px 16px;font-weight:700;font-size:13px}
.btn-nav{background:var(--card2);color:var(--navy);border:1px solid var(--line)}
.btn-green{background:var(--green);color:#fff}.btn-green:hover{background:#178049}
.btn-navy{background:var(--navy);color:#fff}.btn-navy:hover{background:var(--blue)}
.trmbar{background:var(--navy);color:#fff;padding:9px 20px;font-size:13px;display:flex;align-items:center;gap:10px}
.trmbar b{color:#fff}.trmbar .st{color:#bfd4f0}
.trmbar .sp{flex:1}
.wrap{max-width:1180px;margin:0 auto;padding:14px}
.card{background:var(--card);border-radius:14px;padding:16px;margin-bottom:12px;box-shadow:0 1px 4px #0001}
.card h3{color:var(--navy);margin:0 0 10px}
.grid{display:grid;gap:12px}
label.lb{display:block;color:var(--muted);font-size:11px;text-transform:uppercase;margin-bottom:4px}
input[type=text],input[type=email],select{width:100%;padding:9px;border:1px solid var(--line);border-radius:8px;
background:var(--card2);color:var(--text);font-size:14px}
select{cursor:pointer}
.row{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-end}
.col{flex:1;min-width:150px}
.step{display:inline-flex;align-items:center;background:var(--card2);border-radius:8px;overflow:hidden;border:1px solid var(--line)}
.step button{background:var(--navy);color:#fff;border:none;width:30px;height:30px;font-size:17px;font-weight:700}
.step button:hover{background:var(--blue)}
.step span{min-width:34px;text-align:center;font-weight:700}
.ages{display:flex;gap:10px;flex-wrap:wrap;background:var(--card2);border-radius:8px;padding:8px;min-height:20px}
.ages .a{font-size:12px}
.ages .a label{display:block;color:var(--muted);font-size:10px}
.destbar{background:#e4edfa;border-radius:14px;padding:12px 14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.destbar b{color:var(--navy)}
.chip{display:inline-flex;align-items:center;border-radius:16px;overflow:hidden;background:var(--card);border:1px solid var(--line)}
.chip.active{background:var(--navy);border-color:var(--navy)}
.chip button{background:transparent;border:none;padding:7px 10px;font-weight:700;color:var(--navy);font-size:13px}
.chip.active button{color:#fff}
.chip .x{padding:6px 9px;color:var(--muted)}.chip.active .x{color:#fff}.chip .x:hover{background:var(--red);color:#fff}
.actcfg{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin:6px 0 10px;padding:0 4px}
.actcfg b{color:var(--navy)}
.tabs{display:flex;gap:4px;margin-bottom:8px}
.tab{background:var(--card2);border:none;border-radius:10px 10px 0 0;padding:9px 18px;font-weight:700;color:var(--text)}
.tab.active{background:var(--navy);color:#fff}
.panel{background:var(--card);border-radius:0 12px 12px 12px;padding:12px;min-height:120px}
.srch{width:280px;max-width:100%;margin-bottom:8px}
.list{max-height:380px;overflow:auto;background:var(--card2);border-radius:10px;padding:6px}
.item{display:flex;align-items:center;gap:8px;background:var(--card);border-radius:8px;padding:5px 10px;margin:2px}
.item input{width:18px;height:18px;accent-color:var(--navy)}
.item .pv{margin-left:auto;color:var(--muted);font-size:12px;white-space:nowrap}
.hotinfo{background:#eaf2fd;border-radius:12px;padding:12px;margin-top:10px;color:var(--navy);font-weight:700}
.hotinfo small{display:block;color:var(--text);font-weight:400;margin-top:4px}
.footer{position:sticky;bottom:0;background:var(--navy);color:#fff;border-radius:14px;padding:14px 20px;
display:flex;align-items:center;gap:16px;margin-top:6px}
.footer .sp{flex:1}
.footer .tot{font-size:26px;font-weight:800}.footer .ds{color:#bfd4f0;font-size:11px}
.modal{position:fixed;inset:0;background:#0007;display:flex;align-items:center;justify-content:center;z-index:50;padding:16px}
.modal .box{background:var(--card2);border-radius:14px;max-width:560px;width:100%;max-height:90vh;overflow:auto;padding:20px}
.modal h3{color:var(--navy);margin-top:0}
.modal label.lb{margin-top:8px}
.modal textarea{width:100%;padding:9px;border:1px solid var(--line);border-radius:8px;min-height:70px}
.mut{color:var(--muted);font-size:12px}
/* impresion */
#print{display:none}
@media print{
 body>*{display:none!important}
 #print{display:block!important}
 @page{margin:14mm}
}
#print{color:#16233d;font-size:12px}
#print .ph{display:flex;align-items:center;gap:14px;border-bottom:2px solid var(--navy);padding-bottom:8px}
#print .ph img{height:52px}
#print .pe{color:var(--navy);font-size:20px;font-weight:700}
#print .band{background:var(--navy);color:#fff;padding:7px 10px;font-weight:700;margin-top:14px;border-radius:4px}
#print .dband{background:var(--blue);color:#fff;padding:6px 10px;font-weight:700;margin-top:12px;border-radius:4px}
#print table{width:100%;border-collapse:collapse;margin-top:6px}
#print th{background:var(--navy);color:#fff;font-size:11px;padding:5px;text-align:left}
#print td{padding:5px;border-bottom:1px solid #e3ebf6;vertical-align:top;font-size:11px}
#print .desc{color:#6a7889;font-style:italic;font-size:10px}
#print .tot{background:var(--navy);color:#fff;font-size:15px;font-weight:700;padding:8px 10px;margin-top:14px;border-radius:4px;display:flex;justify-content:space-between}
#print .sub{text-align:right;font-weight:700;color:var(--navy);padding:4px 6px}
#print .firma{margin-top:40px}#print .firma .l{border-top:1px solid var(--navy);width:200px;margin-bottom:2px}
</style>
</head>
<body>
<header>
 <img src="data:image/png;base64,__LOGO__" alt="logo">
 <div><div class="t1">Cotizador de Paquetes</div><div class="t2">INNOBA Colombia DMC &middot; v__VERSION__ &middot; Itinerario hasta 5 destinos &middot; Version HTML</div></div>
 <div class="sp"></div>
 <button class="btn btn-navy" onclick="abrirItinerario()">Itinerario</button>
 <button class="btn btn-nav" onclick="abrirConfig()">Datos de mi empresa</button>
</header>
<div class="trmbar"><span>Tasa del dia (USD):</span> <b id="trmStatus">consultando...</b>
 <span class="sp"></span><button class="btn btn-navy" style="padding:6px 12px" onclick="cargarTRM(true)">Actualizar</button></div>
<div id="updbar" class="hidden" style="background:var(--green);color:#fff;padding:9px 20px;font-weight:700;text-align:center"></div>

<div class="wrap">
 <div class="card">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
   <h3 style="margin:0">Datos del viaje (cliente)</h3>
   <div><label class="lb" style="display:inline-block;margin:0 6px 0 0">Cotizado por:</label>
    <select id="cotizador" style="width:auto;display:inline-block">
     <option>Felipe Ortiz Jaramillo  -  Gerente - Innoba DMC</option>
     <option>Carlos Ortiz Jaramillo  -  Gerente Comercial - Innoba DMC</option>
    </select></div>
  </div>
  <div class="row" style="margin-top:8px">
   <div class="col">
    <div style="display:flex;justify-content:space-between;align-items:center">
     <label class="lb">Cliente</label>
     <div style="display:flex;gap:4px">
      <button class="btn btn-navy" style="padding:3px 10px;font-size:11px" onclick="buscarCliente()">&#128269; Buscar</button>
      <button class="btn btn-nav" style="padding:3px 10px;font-size:11px" onclick="editarClienteActual()">&#9998; Editar</button>
     </div></div>
    <input id="cli" type="text" placeholder="Nombre del cliente"></div>
   <div class="col"><label class="lb">Email cliente</label><input id="email" type="email" placeholder="correo@cliente.com"></div>
   <div class="col"><label class="lb">Fechas del viaje</label>
    <div style="display:flex;gap:6px;align-items:center">
     <input id="fdesde" type="date" style="flex:1" title="Fecha de ida">
     <span class="mut">al</span>
     <input id="fhasta" type="date" style="flex:1" title="Fecha de regreso"></div></div>
  </div>
  <div class="row" style="margin-top:10px">
   <div class="col"><label class="lb">Asesor (contacto)</label>
    <div style="display:flex;gap:6px">
     <input id="asesor" type="text" placeholder="Nombre del asesor" style="flex:1">
     <select id="asesorSel" style="width:150px" onchange="usarAsesor(this.value)"><option value="">(vendedor)</option></select>
    </div></div>
   <div class="col"><label class="lb">Telefono asesor</label><input id="asesorTel" type="text" placeholder="Telefono del asesor"></div>
   <div class="col"></div>
  </div>
  <div class="row" style="margin-top:12px">
   <div><label class="lb">Adultos</label><div id="stAd"></div></div>
   <div><label class="lb">Ninos</label><div id="stNi"></div></div>
   <div>
    <label class="lb">Habitaciones (Sen / Dob / Tri) <span id="habMsg" style="font-weight:700"></span></label>
    <div style="display:flex;align-items:center;gap:4px">
     <div id="stHs"></div><div id="stHd"></div><div id="stHt"></div>
     <button class="btn btn-nav" style="padding:6px 10px" onclick="sugerirHab()">Sug.</button>
    </div>
   </div>
   <div class="col"><label class="lb">Edad de cada nino</label><div class="ages" id="ages"></div></div>
  </div>
 </div>

 <div class="destbar">
  <b>Destinos del itinerario (max 5):</b>
  <div id="chips" style="display:flex;gap:8px;flex-wrap:wrap"></div>
  <span class="sp"></span>
  <select id="addDest" onchange="addDestino(this.value)" style="width:200px;background:var(--green);color:#fff;font-weight:700;border:none"></select>
 </div>

 <div class="actcfg" id="actcfg"></div>

 <div id="tabsWrap">
  <div class="tabs">
   <button class="tab active" data-t="hotel" onclick="setTab('hotel')">Hotel</button>
   <button class="tab" data-t="trans" onclick="setTab('trans')">Transportes</button>
   <button class="tab" data-t="act" onclick="setTab('act')">Actividades</button>
  </div>
  <div class="panel" id="panel"></div>
 </div>

 <div class="footer">
  <button class="btn btn-navy" onclick="nueva()">Nueva cotizacion</button>
  <span class="sp"></span>
  <div style="text-align:right"><div class="tot" id="total">USD 0.00</div><div class="ds" id="desg">Total del itinerario</div></div>
  <button class="btn btn-green" style="font-size:15px;padding:12px 22px" onclick="generar()">Generar PDF &#128229;</button>
 </div>
 <p class="mut">Sugerencia: al "Generar PDF", en el dialogo de impresion elige <b>"Guardar como PDF"</b> como destino.</p>
</div>

<div id="print"></div>

<script>
const VERSION = "__VERSION__";
const UPDATE_URL = "https://raw.githubusercontent.com/felipeortizjllo7-del/SOFTWARE-cotizador/main/version.json";
const PRECIOS = __PRECIOS__;
const DESC = __DESC__;
const CLIENTES_BASE = __CLIENTES__;
const LOGO = "data:image/png;base64,__LOGO__";
/* lista de trabajo de clientes (editable, persistida en el navegador) */
let CLIENTES = (()=>{try{const s=localStorage.getItem("innoba_clientes");if(s)return JSON.parse(s);}catch(e){}return CLIENTES_BASE.slice();})();
function guardarClientes(){localStorage.setItem("innoba_clientes",JSON.stringify(CLIENTES));}
let _vendActuales=[];
const EDADES = ["0-11 meses","1 ano","2 anos","3 anos","4 anos","5 anos","6 anos","7 anos","8 anos","9 anos"];
const DEF_CFG = {empresa:"INNOBA Colombia DMC",nit:"",direccion:"",telefono:"",email:"",web:"",
 firma_nombre:"Felipe Ortiz",firma_cargo:"Gerente - INNOBA Colombia DMC",trm_hoy:"",
 notas:"Tarifas sujetas a disponibilidad al momento de la reserva. Precios en dolares americanos (USD) por el total indicado."};

let cfg = Object.assign({}, DEF_CFG, JSON.parse(localStorage.getItem("innoba_cfg")||"{}"));
let st = {adultos:2, ages:[], tramos:[], activo:null, tab:"hotel", hab:{sencilla:0,doble:1,triple:0}, itinerario:""};

/* ---------- utilidades ---------- */
function norm(s){return (s||"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().replace(/\s+/g," ").trim();}
function usd(v){try{return "USD "+Number(v).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});}catch(e){return "USD "+v;}}
function fmtISO(s){if(!s)return"";const p=s.split("-");return p[2]+"/"+p[1]+"/"+p[0];}
function esTransporte(n){n=n.toLowerCase();return ["traslado","asistencia","asitencia","transporte"].some(k=>n.includes(k));}
function esPrivado(n){n=norm(n);return n.includes("privado")||n.includes("privada");}
const COTIZADORES=[["Felipe Ortiz Jaramillo","Gerente - Innoba DMC"],["Carlos Ortiz Jaramillo","Gerente Comercial - Innoba DMC"]];
function periodoFecha(f){if(f){const p=f.split("-"),y=+p[0],m=+p[1];
 if(y>=2027)return{desc:300,mh:0.82,mt:0.69};
 if(y===2026&&m>=9&&m<=12)return{desc:200,mh:null,mt:null};}
 return{desc:100,mh:null,mt:null};}
function periodoActual(){const el=document.getElementById("fdesde");return periodoFecha(el?el.value:"");}
function tasa(){const t=parseFloat((cfg.trm_hoy||"").toString().replace(/,/g,""));if(!(t>=1000))return null;const r=t-periodoActual().desc;return r>0?r:null;}
function margenes(dd){const p=periodoActual();const mh=p.mh||(dd.hoteles?dd.hoteles.margen:0.88);const mt=p.mt||(dd.terrestres?dd.terrestres.margen:0.75);return[mh,mt];}

/* ---------- precios ---------- */
function precioTerrestreUSD(serv,pax,ta,margen){
 const precios=serv.precios;const disp=Object.keys(precios).map(Number).sort((a,b)=>a-b);
 if(!disp.length||!ta)return 0;
 let col=disp.includes(pax)?pax:disp.reduce((p,c)=>Math.abs(c-pax)<Math.abs(p-pax)?c:p,disp[0]);
 if(pax>Math.max(...disp))col=Math.max(...disp);
 const ppc=precios[String(col)];const totalCop=ppc*pax;const ventaCop=margen?totalCop/margen:totalCop;
 return ventaCop/ta;
}
function precioHotelNoche(cop,ta,margen){if(!cop||!ta)return 0;const v=margen?cop/margen:cop;return v/ta;}
function precioHotelNino(ta,margen){if(!ta)return 0;return margen?(70000/margen)/ta:70000/ta;}
function precioServicioGrupo(serv,adultos,ages,ta,margen,privado){
 const N=Math.max(adultos+ages.filter(a=>a>=1).length,1);
 const pp=precioTerrestreUSD(serv,N,ta,margen)/N;
 let total=adultos*pp;
 for(const a of ages){if(a===0)continue;else if(a<=2)total+=privado?10:pp;else total+=0.5*pp;}
 return total;
}

/* ---------- descripciones (match difuso) ---------- */
const STOP=new Set(["de","del","la","el","en","y","a","por","con","los","las","tour","tours","visita","dia","full","the","compartido","compartida","privado","privada","especial","sencilla","sencillo","doble","grupo","pax","round","trip","in","out","ok","o","u","para","desde","hacia"]);
function toks(s){return new Set(norm(s).split(" ").filter(t=>!STOP.has(t)&&t.length>2));}
function bigr(s){s=norm(s);const b=[];for(let i=0;i<s.length-1;i++)b.push(s.slice(i,i+2));return b;}
function dice(a,b){const A=bigr(a),B=bigr(b);if(!A.length||!B.length)return 0;const m={};B.forEach(x=>m[x]=(m[x]||0)+1);let i=0;A.forEach(x=>{if(m[x]>0){i++;m[x]--;}});return 2*i/(A.length+B.length);}
function matchScore(a,b){const ta=toks(a),tb=toks(b);if(!ta.size||!tb.size)return[0,0];let inter=0;ta.forEach(t=>{if(tb.has(t))inter++;});if(inter===0)return[0,0];const cont=inter/ta.size;const uni=new Set([...ta,...tb]).size;const jacc=inter/uni;return[Math.max(cont,jacc),inter];}
function buscarDesc(nombre,destino){const cands=DESC[destino]||[];let best=null,bs=0;for(const c of cands){const[sc,inter]=matchScore(nombre,c.nombre);if(sc>bs){bs=sc;best=c;}}return(best&&bs>=0.34)?best:null;}
function textoDesc(reg,max=1200){let d=(reg.descripcion||"").trim();const dur=(reg.duracion||"").trim(),inc=(reg.incluye||"").trim();const ex=[];if(dur&&dur.length<40)ex.push("Duracion: "+dur);if(inc)ex.push("Incluye: "+inc);let t=d;if(ex.length)t=(d?d+"  ":"")+ex.join(" | ");t=t.replace(/\s+/g," ").trim();if(t.length>max)t=t.slice(0,max-1).replace(/\s\S*$/,"")+"...";return t;}

/* ---------- helpers de datos ---------- */
function temporadasDe(dest){const dd=PRECIOS[dest]||{};const t=[];if(dd.hoteles)dd.hoteles.hoteles.forEach(h=>{const x=(h.temporada||"Baja").trim();if(!t.includes(x))t.push(x);});return t.length?t:["Baja"];}
function hotelesDe(dest,temp){const dd=PRECIOS[dest]||{};const hs=[],ns=[];if(dd.hoteles)dd.hoteles.hoteles.forEach(h=>{const ht=(h.temporada||"Baja").trim();if(ht===temp){hs.push(h);ns.push(h.nombre+(h.zona?" ["+h.zona+"]":""));}});return[hs,ns];}
function hotelesSel(tr){const[hs,ns]=hotelesDe(tr.destino,tr.temporada);const d={};ns.forEach((n,i)=>d[n]=hs[i]);return tr.hoteles.filter(n=>d[n]).map(n=>[n,d[n]]);}
function serviciosPorTipo(dest,tipo){const dd=PRECIOS[dest]||{};if(!dd.terrestres)return[];const s=dd.terrestres.servicios;return tipo==="trans"?s.filter(x=>esTransporte(x.nombre)):s.filter(x=>!esTransporte(x.nombre));}

/* ---------- steppers / edades ---------- */
function stepper(el,get,set,min,max){el.className="step";el.innerHTML=`<button>-</button><span>${get()}</span><button>+</button>`;
 const b=el.querySelectorAll("button"),sp=el.querySelector("span");
 b[0].onclick=()=>{set(Math.max(min,get()-1));sp.textContent=get();};
 b[1].onclick=()=>{set(Math.min(max,get()+1));sp.textContent=get();};}
function renderPax(){
 stepper(document.getElementById("stAd"),()=>st.adultos,v=>{st.adultos=v;valHab();recalc();},1,60);
 stepper(document.getElementById("stNi"),()=>st.ages.length,v=>{
   while(st.ages.length<v)st.ages.push(3);while(st.ages.length>v)st.ages.pop();renderAges();recalc();},0,10);
 stepper(document.getElementById("stHs"),()=>st.hab.sencilla,v=>{st.hab.sencilla=v;valHab();recalc();},0,40);
 stepper(document.getElementById("stHd"),()=>st.hab.doble,v=>{st.hab.doble=v;valHab();recalc();},0,40);
 stepper(document.getElementById("stHt"),()=>st.hab.triple,v=>{st.hab.triple=v;valHab();recalc();},0,40);
 renderAges();valHab();
}
function habOcup(){return st.hab.sencilla*1+st.hab.doble*2+st.hab.triple*3;}
function sugerirHab(){const a=Math.max(st.adultos,0);st.hab={sencilla:a%2,doble:Math.floor(a/2),triple:0};renderPax();recalc();}
function valHab(){const el=document.getElementById("habMsg");if(!el)return;const o=habOcup(),a=st.adultos;
 if(o===a){el.textContent=`OK: ${o} plazas = ${a} adultos`;el.style.color="var(--green)";}
 else{el.textContent=`Ojo: ${o} plazas vs ${a} adultos`;el.style.color="var(--red)";}}
function renderAges(){const c=document.getElementById("ages");c.innerHTML="";
 st.ages.forEach((a,i)=>{const d=document.createElement("div");d.className="a";
  d.innerHTML=`<label>Nino ${i+1}</label><select>${EDADES.map((e,j)=>`<option value="${j}" ${j===a?"selected":""}>${e}</option>`).join("")}</select>`;
  d.querySelector("select").onchange=e=>{st.ages[i]=parseInt(e.target.value);recalc();};c.appendChild(d);});
 if(!st.ages.length)c.innerHTML='<span class="mut">Sin ninos</span>';}

/* ---------- destinos ---------- */
function addDestino(nombre){document.getElementById("addDest").value="";
 if(!PRECIOS[nombre])return;
 const ex=st.tramos.findIndex(t=>t.destino===nombre);if(ex>=0){st.activo=ex;render();return;}
 if(st.tramos.length>=5){alert("Puedes cotizar hasta 5 destinos.");return;}
 st.tramos.push({destino:nombre,temporada:temporadasDe(nombre)[0],noches:3,hoteles:[],hab:{s:0,d:1,t:0},trans:new Set(),act:new Set()});
 st.activo=st.tramos.length-1;render();}
function removeDestino(i){st.tramos.splice(i,1);st.activo=st.tramos.length?Math.min(i,st.tramos.length-1):null;render();}
function setActivo(i){st.activo=i;render();}

function renderChips(){const c=document.getElementById("chips");c.innerHTML="";
 st.tramos.forEach((t,i)=>{const d=document.createElement("div");d.className="chip"+(i===st.activo?" active":"");
  d.innerHTML=`<button>${i+1}. ${t.destino}</button><button class="x">&#10005;</button>`;
  d.children[0].onclick=()=>setActivo(i);d.children[1].onclick=e=>{e.stopPropagation();removeDestino(i);};c.appendChild(d);});
 const sel=document.getElementById("addDest");const disp=Object.keys(PRECIOS).filter(d=>!st.tramos.some(t=>t.destino===d));
 sel.innerHTML=`<option value="">+ Agregar destino</option>`+disp.map(d=>`<option>${d}</option>`).join("");}

/* ---------- config destino activo + tabs ---------- */
function renderActivo(){const w=document.getElementById("actcfg");
 if(st.activo===null){w.innerHTML='<span class="mut">Agrega un destino para comenzar.</span>';document.getElementById("tabsWrap").style.opacity=.4;return;}
 document.getElementById("tabsWrap").style.opacity=1;
 const tr=st.tramos[st.activo];const temps=temporadasDe(tr.destino);
 w.innerHTML=`<b>Destino: ${tr.destino}</b> <span class="mut">Temporada:</span>
  <select id="tmp" style="width:150px">${temps.map(t=>`<option ${t===tr.temporada?"selected":""}>${t}</option>`).join("")}</select>
  <span class="mut">Noches:</span> <span id="stNo"></span>`;
 document.getElementById("tmp").onchange=e=>{tr.temporada=e.target.value;tr.hoteles=[];render();};
 stepper(document.getElementById("stNo"),()=>tr.noches,v=>{tr.noches=v;recalc();},1,60);
}
function setTab(t){st.tab=t;document.querySelectorAll(".tab").forEach(b=>b.classList.toggle("active",b.dataset.t===t));renderPanel();}
function renderPanel(){const p=document.getElementById("panel");
 if(st.activo===null){p.innerHTML='<span class="mut">Sin destino.</span>';return;}
 const tr=st.tramos[st.activo];const ta=tasa();
 if(st.tab==="hotel"){
  const[hs,ns]=hotelesDe(tr.destino,tr.temporada);
  tr.hoteles=tr.hoteles.filter(n=>ns.includes(n));
  const dd=PRECIOS[tr.destino];const mh=margenes(dd)[0];const noches=Math.max(tr.noches,1);
  p.innerHTML=`<div class="row" style="align-items:flex-end;margin-bottom:6px">
    <div style="flex:2"><label class="lb">Buscar hotel</label><input id="srch" placeholder="nombre..."></div>
    <div style="flex:2" class="mut">Precio POR PERSONA en Sencilla / Doble / Triple</div></div>
   <div style="color:var(--navy);font-weight:700;font-size:13px;margin:2px 0 6px" id="hcount"></div>
   <div class="list" id="list"></div>`;
  function pintarH(){const q=norm(document.getElementById("srch").value);const L=document.getElementById("list");L.innerHTML="";
   document.getElementById("hcount").textContent=`Marca hasta 5 hoteles como opciones (${tr.hoteles.length}/5 elegidos)`;
   hs.forEach((h,i)=>{const n=ns[i];if(q&&!norm(n).includes(q))return;const sel=tr.hoteles.includes(n);
    const it=document.createElement("div");it.className="item";
    const pn=k=>h[k]&&ta?usd(precioHotelNoche(h[k],ta,mh)*noches/({sencilla:1,doble:2,triple:3}[k])):"N/D";
    const cat=h.categoria?`<span style="background:var(--cyan);color:#fff;border-radius:8px;padding:1px 8px;font-size:11px;font-weight:700">${h.categoria}</span>`:"";
    it.innerHTML=`<input type="checkbox" ${sel?"checked":""}><div style="flex:1;min-width:0"><b>${n}</b> ${cat}</div><div class="mut" style="font-size:11px;white-space:nowrap;text-align:right">Sen ${pn("sencilla")} · Dob ${pn("doble")} · Tri ${pn("triple")}</div>`;
    it.querySelector("input").onchange=e=>{if(e.target.checked){if(tr.hoteles.length>=5){e.target.checked=false;alert("Hasta 5 hoteles.");return;}tr.hoteles.push(n);}else{tr.hoteles=tr.hoteles.filter(x=>x!==n);}recalc();pintarH();};
    L.appendChild(it);});
   if(!L.children.length)L.innerHTML='<div class="mut" style="padding:16px">No hay hoteles.</div>';}
  document.getElementById("srch").oninput=pintarH;pintarH();
 } else {
  const tipo=st.tab;const sel=tr[tipo];const servs=serviciosPorTipo(tr.destino,tipo);
  p.innerHTML=`<input class="srch" id="srch" placeholder="Buscar servicio..."><div class="list" id="list"></div>`;
  const pintar=()=>{const q=norm(document.getElementById("srch").value);const L=document.getElementById("list");L.innerHTML="";
   servs.filter(s=>!q||norm(s.nombre).includes(q)).forEach(s=>{const it=document.createElement("div");it.className="item";
    const priv=esPrivado(s.nombre)?' <span class="mut">(privado)</span>':'';
    it.innerHTML=`<input type="checkbox" ${sel.has(s.nombre)?"checked":""}><div>${s.nombre}${priv}</div>`;
    it.querySelector("input").onchange=e=>{e.target.checked?sel.add(s.nombre):sel.delete(s.nombre);recalc();};L.appendChild(it);});
   if(!L.children.length)L.innerHTML='<div class="mut" style="padding:16px">No hay servicios.</div>';};
  document.getElementById("srch").oninput=pintar;pintar();
 }
}

/* ---------- calculo ---------- */
function calcularTramo(tr,ta){
 const dd=PRECIOS[tr.destino];const noches=Math.max(tr.noches,1);
 const adultos=Math.max(st.adultos,1);const ages=st.ages;
 const detPax=`${adultos} ad`+(ages.length?` + ${ages.length} nino(s)`:"");
 const[mh,mt]=margenes(dd);
 const baseSec=[];let base=0,adultPp=0;const ninosServ=ages.map(()=>0);
 if(dd.terrestres){const smap={};dd.terrestres.servicios.forEach(s=>smap[s.nombre]=s);const N=Math.max(adultos+ages.filter(a=>a>=1).length,1);
  [["trans","TRANSPORTE TERRESTRE"],["act","ACTIVIDADES"]].forEach(([tipo,tit])=>{const filas=[];let ss=0;
   [...tr[tipo]].sort().forEach(n=>{const s=smap[n];if(!s)return;const priv=esPrivado(n);
    const imp=precioServicioGrupo(s,adultos,ages,ta,mt,priv);ss+=imp;const pp=precioTerrestreUSD(s,N,ta,mt)/N;adultPp+=pp;
    ages.forEach((a,i)=>{if(a===0)return;else if(a<=2)ninosServ[i]+=priv?10:pp;else ninosServ[i]+=0.5*pp;});
    let d="";if(tipo==="act"){const r=buscarDesc(n,tr.destino);if(r)d=textoDesc(r);}filas.push([n,detPax,imp,pp,d]);});
   if(filas.length){baseSec.push([tit,filas,ss]);base+=ss;}});}
 const surch=precioHotelNino(ta,mh)*noches;
 const ninosTot=ages.map((a,i)=>ninosServ[i]+(a>=3?surch:0));
 const gr={};ages.forEach((a,i)=>{if(!gr[a])gr[a]=[0,0];gr[a][0]++;gr[a][1]=ninosTot[i];});
 const ninos=Object.keys(gr).map(Number).sort((x,y)=>x-y).map(a=>[a,gr[a][0],gr[a][1]]);
 const OCC={sencilla:1,doble:2,triple:3};
 const opciones=hotelesSel(tr).map(([n,h])=>{const o={nombre:h.nombre,categoria:h.categoria||""};
   ["sencilla","doble","triple"].forEach(k=>{o[k]=h[k]?adultPp+precioHotelNoche(h[k],ta,mh)*noches/OCC[k]:null;});return o;});
 return{destino:tr.destino,subtitulo:`${tr.destino}  ·  Temporada ${tr.temporada}  ·  ${noches} noches`,
        baseSec,base,opciones,ninos,n_adultos:adultos};
}
function calcularTodo(){const ta=tasa();if(ta===null)return[[],0,false];const bloques=[];let ref=0,opc=false;
 st.tramos.forEach(tr=>{const b=calcularTramo(tr,ta);if(b.opciones.length>1)opc=true;
  if(b.opciones.length){const o=b.opciones[0];ref+=o.doble||o.sencilla||o.triple||0;}else ref+=b.base;bloques.push(b);});
 return[bloques,ref,opc];}
function totalReserva(bloques){const OCC={sencilla:1,doble:2,triple:3};const conOp=bloques.filter(b=>b.opciones.length);
 if(!conOp.length||habOcup()===0)return null;let t=0;
 conOp.forEach(b=>{const o=b.opciones[0];
  ["sencilla","doble","triple"].forEach(acc=>{if(o[acc]&&st.hab[acc])t+=st.hab[acc]*OCC[acc]*o[acc];});
  t+=b.ninos.reduce((x,[a,c,pr])=>x+c*pr,0);});return t;}
function recalc(){const[bloques,total,opc]=calcularTodo();
 if(tasa()===null){document.getElementById("total").textContent="USD --";document.getElementById("desg").textContent="Falta la TRM (ponla en Datos de mi empresa)";return;}
 const res=totalReserva(bloques);
 if(res!==null){document.getElementById("total").textContent=usd(res);
  const h=st.hab;const det=["sencilla","doble","triple"].filter(k=>h[k]).map(k=>h[k]+" "+k).join(", ");
  document.getElementById("desg").textContent="Total reserva ("+det+") - 1a opcion";}
 else{document.getElementById("total").textContent=usd(total);
  document.getElementById("desg").textContent="Indica las habitaciones para el total";}}

/* ---------- TRM ---------- */
async function cargarTRM(manualBtn){const s=document.getElementById("trmStatus");s.textContent="consultando...";
 let trm=null;
 // Fuente OFICIAL de la TRM (datos.gov.co) - estable, con CORS, se actualiza a diario
 try{const r=await fetch("https://www.datos.gov.co/resource/32sa-8pi3.json?$order=vigenciadesde%20DESC&$limit=1",{cache:"no-store"});
  const d=await r.json();const v=parseFloat(String(d[0].valor).replace(/,/g,""));
  if(v>=1000&&v<=20000)trm=v;}catch(e){}
 if(trm){cfg.trm_hoy=String(trm);localStorage.setItem("innoba_cfg",JSON.stringify(cfg));s.textContent="actualizada (automatica) ✓";}
 else if(parseFloat(cfg.trm_hoy)>=1000){s.textContent="aplicada (guardada) ✓";}
 else{s.textContent="falta: ponla en 'Datos de mi empresa'";if(manualBtn)abrirConfig();}
 recalc();}

/* ---------- generar / imprimir ---------- */
function generar(){
 if(tasa()===null){alert("Falta la TRM. Ponla en 'Datos de mi empresa'.");abrirConfig();return;}
 if(!st.tramos.length){alert("Agrega al menos un destino.");return;}
 const[bloques,total,opc]=calcularTodo();
 if(total<=0||!bloques.some(b=>b.baseSec.length||b.opciones.length)){alert("Selecciona al menos un hotel, transporte o actividad.");return;}
 const fd=document.getElementById("fdesde").value, fh=document.getElementById("fhasta").value;
 if(!fd||!fh){alert("Debes elegir la FECHA DEL VIAJE (ida y regreso) en el calendario.");return;}
 if(fh<fd){alert("La fecha de regreso no puede ser anterior a la de ida.");return;}
 const fechasViaje=fmtISO(fd)+" al "+fmtISO(fh);
 if(!document.getElementById("email").value.trim()){alert("Debes ingresar el EMAIL del cliente.");return;}
 const hoy=new Date();const dd=String(hoy.getDate()).padStart(2,"0"),mm=String(hoy.getMonth()+1).padStart(2,"0");
 const fecha=`${dd}/${mm}/${hoy.getFullYear()}`;const v=new Date(hoy);v.setMonth(v.getMonth()+1);
 const valida=`${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;
 const ad=st.adultos;let paxtxt=`${ad} adultos`;if(st.ages.length)paxtxt+=`, ${st.ages.length} ninos (${st.ages.map(a=>a===0?"bebe":a+" anos").join(", ")})`;
 const multi=bloques.length>1;
 let html=`<div class="ph"><img src="${LOGO}"><div><div class="pe">${cfg.empresa}</div>
  <div style="font-size:11px">${[cfg.nit?"NIT/RUC: "+cfg.nit:"",cfg.telefono?"Tel: "+cfg.telefono:"",cfg.email,cfg.web].filter(Boolean).join(" | ")}</div></div></div>`;
 html+=`<div class="band">COTIZACION ${multi?"- ITINERARIO":"- "+bloques[0].destino}</div>`;
 html+=`<table style="margin-top:8px"><tr>
  <td style="width:50%"><b>Cliente:</b> ${document.getElementById("cli").value||"-"}<br><b>Email:</b> ${document.getElementById("email").value}${document.getElementById("asesor").value?`<br><b>Asesor:</b> ${document.getElementById("asesor").value}`:""}${document.getElementById("asesorTel").value?`<br><b>Tel. asesor:</b> ${document.getElementById("asesorTel").value}`:""}</td>
  <td><b>No.:</b> COT-${hoy.getFullYear()}${mm}${dd}-001<br><b>Fecha:</b> ${fecha} &nbsp; <b>Valida hasta:</b> ${valida}<br>
  <b>Fechas viaje:</b> ${fechasViaje} &nbsp; <b>Pasajeros:</b> ${paxtxt}</td></tr></table>`;
 const edadTxt=a=>a===0?"Bebe 0-11 meses":(a+(a===1?" ano":" anos"));
 const cel=v=>v?usd(v):"-";
 bloques.forEach(b=>{html+=`<div class="dband">${b.subtitulo}</div>`;
  b.baseSec.forEach(([tit,filas,sub])=>{html+=`<div style="color:var(--blue);font-weight:700;margin-top:6px">${tit}</div>
   <table><tr><th>Concepto</th><th>Detalle</th><th style="text-align:right">Por pasajero</th><th style="text-align:right">Total (USD)</th></tr>`;
   filas.forEach(([c,det,val,pp,d])=>{html+=`<tr><td><b>${c}</b>${d?`<div class="desc">${d}</div>`:""}</td><td>${det}</td><td style="text-align:right">${pp?usd(pp):"-"}</td><td style="text-align:right"><b>${usd(val)}</b></td></tr>`;});
   html+=`</table>`;});
  if(b.opciones.length){html+=`<div style="color:var(--blue);font-weight:700;margin-top:8px">${b.opciones.length>1?"OPCIONES DE HOTEL - precio por persona (el cliente elige)":"ALOJAMIENTO - precio por persona"}</div>
   <table><tr><th>Hotel</th><th>Categoria</th><th style="text-align:right">Sencilla</th><th style="text-align:right">Doble</th><th style="text-align:right">Triple</th></tr>`;
   b.opciones.forEach(o=>{html+=`<tr><td><b>${o.nombre}</b></td><td>${o.categoria||"-"}</td><td style="text-align:right"><b>${cel(o.sencilla)}</b></td><td style="text-align:right"><b>${cel(o.doble)}</b></td><td style="text-align:right"><b>${cel(o.triple)}</b></td></tr>`;});
   html+=`</table><div class="desc">Valores POR PERSONA (adulto) segun acomodacion, por todo el viaje. Incluye traslados y actividades.</div>`;}
  if(b.ninos.length){html+=`<div class="desc" style="margin-top:2px">Precio por nino: ${b.ninos.map(([a,c,pr])=>`${edadTxt(a)} (x${c}): ${usd(pr)}`).join("  |  ")}</div>`;}
 });
 // Costo total de la reserva con la 1a opcion de hotel y las habitaciones indicadas
 const conOp=bloques.filter(b=>b.opciones.length);
 const OCC={sencilla:1,doble:2,triple:3};const ocup=habOcup();
 if(conOp.length&&ocup>0){const nAd=conOp[0].n_adultos;const nNi=conOp.reduce((s,b)=>s+b.ninos.reduce((x,[a,c])=>x+c,0),0);
  let totalRes=0;const detHab=[];
  ["sencilla","doble","triple"].forEach(acc=>{const n=st.hab[acc];if(!n)return;detHab.push(n+" "+acc);
   conOp.forEach(b=>{if(b.opciones[0][acc])totalRes+=n*OCC[acc]*b.opciones[0][acc];});});
  totalRes+=conOp.reduce((s,b)=>s+b.ninos.reduce((x,[a,c,pr])=>x+c*pr,0),0);
  const hop=conOp.map(b=>b.opciones[0].nombre).join(" + ");
  html+=`<div class="tot"><span>COSTO TOTAL DE LA RESERVA - 1a opcion (${nAd} adulto(s)${nNi?` + ${nNi} nino(s)`:""})</span><span></span></div>
   <table><tr><td>Habitaciones solicitadas</td><td style="text-align:right">${detHab.join(", ")}</td></tr>
   <tr><td style="font-size:14px"><b>TOTAL DE LA RESERVA (USD)</b></td><td style="text-align:right;font-size:14px"><b>${usd(totalRes)}</b></td></tr></table>
   <div class="desc">Calculado con la 1a opcion: ${hop}</div>`;}
 // Itinerario dia por dia
 const itin=(st.itinerario||itinerarioAuto()).trim();
 if(itin){html+=`<div class="band" style="margin-top:14px">ITINERARIO DIA POR DIA</div>`;
  itin.split("\n").forEach(par=>{par=par.trim();if(!par)return;
   if(/^d[ií]a\s*\d/i.test(par))html+=`<div class="dband">${par}</div>`;
   else html+=`<div style="font-size:11px;margin:3px 0;line-height:1.4">${par}</div>`;});}
 html+=`<div style="margin-top:10px;font-size:11px"><b>Notas:</b> Vigencia: esta cotizacion tiene una validez de un (1) mes, hasta el ${valida}. ${cfg.notas}</div>`;
 const ci=document.getElementById("cotizador").selectedIndex; const cz=COTIZADORES[ci]||COTIZADORES[0];
 html+=`<div class="firma"><div>Cordialmente,</div><br><br><div class="l"></div><b style="color:var(--navy)">${cz[0]}</b><br>${cz[1]}</div>`;
 document.getElementById("print").innerHTML=html;
 window.print();
}

/* ---------- clientes: usar / buscar / editar ---------- */
function usarCliente(emp){const c=CLIENTES.find(x=>x.empresa===emp);if(!c)return;
 document.getElementById("cli").value=c.empresa||"";
 let email=c.email||"";if(!email&&c.vendedores&&c.vendedores.length)email=c.vendedores[0].email||"";
 if(email)document.getElementById("email").value=email;
 _vendActuales=c.vendedores||[];
 const sel=document.getElementById("asesorSel");
 sel.innerHTML=`<option value="">(vendedor)</option>`+_vendActuales.filter(v=>v.nombre).map(v=>`<option>${v.nombre}</option>`).join("");
 if(_vendActuales.length){sel.value=_vendActuales[0].nombre;usarAsesor(_vendActuales[0].nombre);}
 else{document.getElementById("asesor").value="";document.getElementById("asesorTel").value="";}}
function usarAsesor(nombre){const v=_vendActuales.find(x=>x.nombre===nombre);
 document.getElementById("asesor").value=nombre||"";
 if(v){document.getElementById("asesorTel").value=v.telefono||"";if(v.email)document.getElementById("email").value=v.email;}}
function cerrarModal(){const m=document.getElementById("modal");if(m)m.remove();}
function buscarCliente(){const m=document.createElement("div");m.className="modal";m.id="modal";
 m.innerHTML=`<div class="box" style="max-width:520px">
  <div style="display:flex;justify-content:space-between;align-items:center"><h3 style="margin:0">Buscar cliente</h3>
   <button class="btn btn-nav" style="padding:4px 10px" onclick="cerrarModal()">Cerrar</button></div>
  <input id="qcli" type="text" placeholder="Escribe la empresa o el asesor..." style="margin:10px 0" autofocus>
  <button class="btn btn-green" style="width:100%;margin-bottom:10px" onclick="editarCliente('')">+ Nueva empresa</button>
  <div class="list" id="qlist" style="max-height:52vh"></div></div>`;
 document.body.appendChild(m);
 const pintar=()=>{const q=norm(document.getElementById("qcli").value);const L=document.getElementById("qlist");L.innerHTML="";let n=0;
  for(const c of CLIENTES){const emp=c.empresa||"";const vs=c.vendedores||[];
   if(q&&!(norm(emp).includes(q)||vs.some(v=>norm(v.nombre||"").includes(q))))continue;
   if(++n>250)break;const sub=vs.map(v=>v.nombre).filter(Boolean).slice(0,3).join(", ");
   const it=document.createElement("div");it.className="item";it.style.cursor="pointer";
   it.innerHTML=`<div style="flex:1" onclick="usarCliente(${JSON.stringify(emp)});cerrarModal()"><b>${emp}</b>${sub?`<div class="mut" style="font-size:11px">${sub}</div>`:""}</div>
    <button class="btn btn-nav" style="padding:4px 9px" onclick="event.stopPropagation();editarCliente(${JSON.stringify(emp)})">&#9998;</button>
    <button class="btn btn-nav" style="padding:4px 9px;color:var(--red)" onclick="event.stopPropagation();eliminarClienteRapido(${JSON.stringify(emp)})">&#128465;</button>`;
   L.appendChild(it);}
  if(!n)L.innerHTML='<div class="mut" style="padding:16px">Sin resultados.</div>';};
 document.getElementById("qcli").oninput=pintar;pintar();}
function editarClienteActual(){editarCliente(document.getElementById("cli").value.trim());}
function editarCliente(emp){cerrarModal();
 const c=CLIENTES.find(x=>x.empresa===emp)||{empresa:emp||"",nit:"",telefono:"",email:"",web:"",pais:"",vendedores:[]};
 const esNuevo=!CLIENTES.includes(c);
 const m=document.createElement("div");m.className="modal";m.id="modal";
 const f=(k,l)=>`<label class="lb">${l}</label><input id="e_${k}" type="text" value="${(c[k]||"").replace(/"/g,"&quot;")}">`;
 m.innerHTML=`<div class="box"><div style="display:flex;justify-content:space-between;align-items:center">
   <h3 style="margin:0">${esNuevo?"Nueva empresa":"Editar cliente"}</h3>
   <button class="btn btn-nav" style="padding:4px 10px" onclick="buscarCliente()">&larr; Volver</button></div>
  ${f("empresa","Nombre de la empresa *")}${f("nit","NIT / Documento")}${f("telefono","Telefono")}
  ${f("email","Email")}${f("web","Sitio web")}${f("pais","Pais")}
  <div style="display:flex;justify-content:space-between;align-items:center;margin-top:12px">
   <b style="color:var(--navy)">Vendedores / contactos</b>
   <button class="btn btn-nav" style="padding:4px 10px" onclick="addVendRow()">+ Agregar</button></div>
  <div id="vends"></div>
  <div style="display:flex;gap:8px;justify-content:space-between;margin-top:14px">
   <button class="btn btn-nav" style="color:var(--red)" onclick="eliminarCliente(${JSON.stringify(c.empresa||"")})">Eliminar</button>
   <div style="display:flex;gap:8px">
    <button class="btn btn-nav" onclick="buscarCliente()">Cancelar</button>
    <button class="btn btn-green" onclick="guardarClienteForm(${JSON.stringify(c.empresa||"")})">Guardar</button></div></div>
  </div>`;
 document.body.appendChild(m);
 (c.vendedores||[]).forEach(addVendRow);if(!(c.vendedores||[]).length)addVendRow();}
function addVendRow(v){v=v||{};const cont=document.getElementById("vends");const row=document.createElement("div");
 row.className="vrow";row.style.cssText="display:flex;gap:4px;margin:4px 0";
 row.innerHTML=`<input placeholder="Nombre" value="${(v.nombre||"").replace(/"/g,'&quot;')}" style="flex:2">
  <input placeholder="Telefono" value="${(v.telefono||"").replace(/"/g,'&quot;')}" style="flex:1">
  <input placeholder="Email" value="${(v.email||"").replace(/"/g,'&quot;')}" style="flex:2">
  <input placeholder="Cargo" value="${(v.cargo||"").replace(/"/g,'&quot;')}" style="flex:1">
  <button class="btn btn-nav" style="color:var(--red);padding:4px 9px" onclick="this.parentNode.remove()">&#10005;</button>`;
 cont.appendChild(row);}
function guardarClienteForm(orig){const g=k=>document.getElementById("e_"+k).value.trim();
 const empresa=g("empresa");if(!empresa){alert("Escribe el nombre de la empresa.");return;}
 const vends=[...document.querySelectorAll("#vends .vrow")].map(r=>{const i=r.querySelectorAll("input");
  return{nombre:i[0].value.trim(),telefono:i[1].value.trim(),email:i[2].value.trim(),cargo:i[3].value.trim()};}).filter(v=>v.nombre);
 const nuevo={empresa,nit:g("nit"),telefono:g("telefono"),email:g("email"),web:g("web"),pais:g("pais"),vendedores:vends};
 const idx=CLIENTES.findIndex(x=>x.empresa===orig);
 if(idx>=0)CLIENTES[idx]=nuevo;else CLIENTES.push(nuevo);
 guardarClientes();cerrarModal();usarCliente(empresa);}
function eliminarCliente(emp){if(!emp)return;if(!confirm("¿Eliminar esta empresa?"))return;
 CLIENTES=CLIENTES.filter(x=>x.empresa!==emp);guardarClientes();cerrarModal();}
function eliminarClienteRapido(emp){if(!emp)return;
 if(!confirm("¿Eliminar definitivamente a:\n\n"+emp+"\n\nEsta accion no se puede deshacer."))return;
 CLIENTES=CLIENTES.filter(x=>x.empresa!==emp);guardarClientes();
 if(document.getElementById("cli").value===emp){document.getElementById("cli").value="";document.getElementById("asesor").value="";document.getElementById("asesorTel").value="";}
 const q=document.getElementById("qcli");if(q)q.dispatchEvent(new Event("input"));}

/* ---------- itinerario ---------- */
function itinerarioAuto(){let L=[],dia=1;
 if(st.tramos.length){L.push(`DIA 01: LLEGADA A ${st.tramos[0].destino.toUpperCase()}`);
  L.push("Recepcion en el aeropuerto por nuestro equipo y traslado al hotel seleccionado. Registro y alojamiento.");dia=2;}
 st.tramos.forEach(tr=>{[...tr.act].sort().forEach(act=>{const r=buscarDesc(act,tr.destino);const d=r?textoDesc(r):"";
  L.push(`DIA ${String(dia).padStart(2,"0")}: ${tr.destino.toUpperCase()} - ${act.toUpperCase()}`);
  L.push(d||"Actividad programada. Alojamiento.");dia++;});});
 L.push(`DIA ${String(dia).padStart(2,"0")}: TRASLADO AL AEROPUERTO`);
 L.push("Desayuno. A la hora indicada realizamos el traslado al aeropuerto. Fin de nuestros servicios.");
 return L.join("\n");}
function abrirItinerario(){if(!st.tramos.length){alert("Agrega al menos un destino primero.");return;}
 const m=document.createElement("div");m.className="modal";m.id="modal";
 m.innerHTML=`<div class="box" style="max-width:660px">
  <div style="display:flex;justify-content:space-between;align-items:center"><h3 style="margin:0">Itinerario dia por dia</h3>
   <button class="btn btn-nav" style="padding:4px 10px" onclick="cerrarModal()">Cerrar</button></div>
  <div class="mut" style="margin:6px 0">Editable. Las lineas que empiezan por 'DIA' salen resaltadas en el PDF.</div>
  <textarea id="itinTxt" style="width:100%;min-height:340px;font-size:13px">${(st.itinerario||itinerarioAuto()).replace(/</g,"&lt;")}</textarea>
  <div style="display:flex;justify-content:space-between;gap:8px;margin-top:12px">
   <button class="btn btn-nav" onclick="document.getElementById('itinTxt').value=itinerarioAuto()">Regenerar automatico</button>
   <button class="btn btn-green" onclick="st.itinerario=document.getElementById('itinTxt').value.trim();cerrarModal()">Guardar</button></div></div>`;
 document.body.appendChild(m);}

/* ---------- config ---------- */
function abrirConfig(){const m=document.createElement("div");m.className="modal";m.id="modal";
 const f=(k,l,t)=>`<label class="lb">${l}</label><input id="c_${k}" type="${t||"text"}" value="${(cfg[k]||"").replace(/"/g,"&quot;")}">`;
 m.innerHTML=`<div class="box"><h3>Datos de mi empresa</h3>
  ${f("empresa","Nombre de la empresa")}${f("nit","NIT / RUC")}${f("direccion","Direccion")}${f("telefono","Telefono")}
  ${f("email","Email")}${f("web","Sitio web")}${f("firma_nombre","Firma - Nombre")}${f("firma_cargo","Firma - Cargo")}
  <label class="lb" style="margin-top:8px;color:var(--navy)">TRM de hoy (COP) &mdash; cópiala de dolar-colombia.com</label>
  <input id="c_trm_hoy" type="text" value="${cfg.trm_hoy||""}" placeholder="ej. 3248.87">
  <div class="mut">El sistema le resta 100 automaticamente. No se muestra en la cotizacion.</div>
  <label class="lb" style="margin-top:8px">Notas por defecto</label><textarea id="c_notas">${cfg.notas||""}</textarea>
  <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:14px">
   <button class="btn btn-nav" onclick="document.getElementById('modal').remove()">Cancelar</button>
   <button class="btn btn-green" onclick="guardarConfig()">Guardar</button></div></div>`;
 document.body.appendChild(m);}
function guardarConfig(){["empresa","nit","direccion","telefono","email","web","firma_nombre","firma_cargo","trm_hoy","notas"].forEach(k=>{
  const el=document.getElementById("c_"+k);if(el)cfg[k]=el.value.trim();});
 localStorage.setItem("innoba_cfg",JSON.stringify(cfg));document.getElementById("modal").remove();cargarTRM(false);}

/* ---------- render maestro ---------- */
function render(){renderChips();renderActivo();renderPanel();recalc();}
function nueva(){st={adultos:2,ages:[],tramos:[],activo:null,tab:"hotel",hab:{sencilla:0,doble:1,triple:0},itinerario:""};
 document.getElementById("cli").value="";document.getElementById("email").value="";
 document.getElementById("asesor").value="";document.getElementById("asesorTel").value="";
 _vendActuales=[];document.getElementById("asesorSel").innerHTML='<option value="">(vendedor)</option>';
 document.getElementById("fdesde").value="";document.getElementById("fhasta").value="";
 document.querySelectorAll(".tab").forEach(b=>b.classList.toggle("active",b.dataset.t==="hotel"));
 renderPax();addDestino(Object.keys(PRECIOS)[0]);sugerirHab();}

/* fecha minima = hoy; al elegir la ida, la vuelta no puede ser antes */
(function(){const t=new Date();const iso=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}-${String(t.getDate()).padStart(2,"0")}`;
 const fd=document.getElementById("fdesde"),fh=document.getElementById("fhasta");
 fd.min=iso;fh.min=iso;fd.onchange=()=>{if(fd.value){fh.min=fd.value;if(fh.value&&fh.value<fd.value)fh.value=fd.value;}if(st.activo!==null)renderPanel();recalc();};})();

/* ---------- actualizaciones ---------- */
function verTuple(s){return (String(s||"0").match(/\d+/g)||["0"]).slice(0,3).map(Number);}
function verMayor(a,b){a=verTuple(a);b=verTuple(b);for(let i=0;i<3;i++){if((a[i]||0)>(b[i]||0))return true;if((a[i]||0)<(b[i]||0))return false;}return false;}
async function chequearActualizacion(){try{
  const r=await fetch(UPDATE_URL+"?t="+Date.now(),{cache:"no-store"});if(!r.ok)return;
  const info=await r.json();
  if(verMayor(info.version,VERSION)){const b=document.getElementById("updbar");
   b.innerHTML=`&#128260; Hay una nueva version disponible (v${info.version}, tienes v${VERSION}). `+
    `<a href="https://github.com/felipeortizjllo7-del/SOFTWARE-cotizador/releases/latest" target="_blank" style="color:#fff;text-decoration:underline">Descargar la ultima version</a>`;
   b.classList.remove("hidden");}
 }catch(e){}}

/* init */
renderPax();addDestino(Object.keys(PRECIOS)[0]);sugerirHab();cargarTRM(false);chequearActualizacion();
</script>
</body>
</html>"""

HTML = (HTML.replace("__PRECIOS__", json.dumps(precios, ensure_ascii=False))
            .replace("__DESC__", json.dumps(desc, ensure_ascii=False))
            .replace("__CLIENTES__", json.dumps(clientes, ensure_ascii=False))
            .replace("__VERSION__", VERSION)
            .replace("__LOGO__", logo_b64)
            .replace("__ICON__", icon_b64))

out = os.path.join(PROJ, "CotizadorInnoba.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(HTML)
print("HTML generado:", out, "  tamano:", round(len(HTML)/1024), "KB")
