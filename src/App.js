import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";

const SK = "fabric-sales-v3-trading";
const EMPTY = { customers:[], suppliers:[], products:[], tradingSales:[], tradingPayments:[], debitNotes:[], creditNotes:[] };
async function loadData(){try{const r=localStorage.getItem(SK);if(r)return{...EMPTY,...JSON.parse(r)};}catch(e){}return{...EMPTY};}
async function saveData(d){try{localStorage.setItem(SK,JSON.stringify(d));}catch(e){}}

const fmt   = n => Number(n||0).toLocaleString("en-IN",{maximumFractionDigits:2});
const fmtD  = d => d?new Date(d).toLocaleDateString("en-IN"):"—";
const today = () => new Date().toISOString().split("T")[0];
const uid   = () => Date.now()+"-"+Math.random().toString(36).slice(2,6);
function excelDateToJS(v){if(v instanceof Date)return v;if(typeof v==="number")return new Date(Math.round((v-25569)*86400*1000));if(typeof v==="string"){const d=new Date(v);if(!isNaN(d))return d;}return null;}
function exportBackup(data){const b=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=`FabricSales_${today()}.json`;a.click();}
function exportCSV(rows,fn){const csv=rows.map(r=>r.map(c=>`"${String(c||"").replace(/"/g,'""')}"`).join(",")).join("\n");const b=new Blob(["\uFEFF"+csv],{type:"text/csv"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=fn;a.click();}

// ── Financial Year helpers ────────────────────────────────────────
function getFY(dateStr){const d=new Date(dateStr);const y=d.getFullYear();const m=d.getMonth();return m>=3?`${y}-${String(y+1).slice(2)}`:`${y-1}-${String(y).slice(2)}`;}
function getFYRange(fy){const[startY]=fy.split("-");const y=parseInt(startY);return{start:new Date(y,3,1),end:new Date(y+1,2,31,23,59,59)};}
const ALL_FY="All Years";
function getAvailableFYs(sales){const s=new Set(sales.map(s=>getFY(s.date)));return[ALL_FY,...[...s].sort().reverse()];}
function filterByFY(sales,fy){if(fy===ALL_FY)return sales;const{start,end}=getFYRange(fy);return sales.filter(s=>{const d=new Date(s.date);return d>=start&&d<=end;});}

const C={navy:"#0F1923",navyMid:"#1A3A5C",gold:"#E8C97E",blue:"#2980B9",green:"#27AE60",red:"#E74C3C",orange:"#E67E22",purple:"#8E44AD",teal:"#16A085",bg:"#F0F4F8",card:"#FFFFFF",border:"#E2EAF4",muted:"#7A8A9A"};

const Row=({children,style})=><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",...style}}>{children}</div>;
const B=({children,style})=><div style={{fontWeight:700,fontSize:14,...style}}>{children}</div>;
const Mute=({children,style})=><div style={{fontSize:12.5,color:C.muted,marginTop:2,...style}}>{children}</div>;
const SecTitle=({children})=><div style={{fontWeight:800,fontSize:15,color:C.navy,marginBottom:12}}>{children}</div>;
const Empty=({text})=><div style={{textAlign:"center",color:"#bbb",fontSize:13.5,padding:"44px 10px"}}>{text}</div>;
const Btn=({children,color,onClick,style})=><button onClick={onClick} style={{background:color,color:color===C.navy?"#E8C97E":"#fff",border:"none",borderRadius:11,padding:"12px 18px",fontSize:13.5,fontWeight:700,cursor:"pointer",minHeight:46,...style}}>{children}</button>;
const IS={width:"100%",padding:"12px 14px",borderRadius:11,border:`1.5px solid ${C.border}`,fontSize:15,boxSizing:"border-box",outline:"none",background:"#fff",minHeight:46};
const hdrBtn=()=>({background:"rgba(255,255,255,0.13)",border:"1px solid rgba(255,255,255,0.22)",borderRadius:9,padding:"8px 11px",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:3,minHeight:44,minWidth:54});
function KpiCard({icon,label,val,color,sub}){return(<div style={{background:C.card,borderRadius:14,padding:14,borderLeft:`4px solid ${color}`,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}><div style={{fontSize:20}}>{icon}</div><div style={{fontSize:10.5,color:C.muted,marginTop:4,fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{label}</div><div style={{fontSize:17,fontWeight:900,color,marginTop:2}}>₹{fmt(val)}</div>{sub&&<div style={{fontSize:10.5,color:C.muted,marginTop:2}}>{sub}</div>}</div>);}
function SegCtrl({options,val,onChange}){return(<div style={{display:"flex",background:"#fff",borderRadius:11,overflow:"hidden",border:`1px solid ${C.border}`}}>{options.map(o=>(<button key={o.v} onClick={()=>onChange(o.v)} style={{flex:1,padding:"12px 4px",fontSize:12.5,fontWeight:val===o.v?700:500,color:val===o.v?"#E8C97E":C.muted,background:val===o.v?C.navy:"transparent",border:"none",cursor:"pointer",whiteSpace:"nowrap",minHeight:46}}>{o.l}</button>))}</div>);}
function ModalBase({title,onClose,children}){return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:200,display:"flex",alignItems:"flex-end"}} onClick={onClose}><div style={{background:"#fff",width:"100%",maxWidth:600,margin:"0 auto",borderRadius:"20px 20px 0 0",padding:"18px 16px calc(env(safe-area-inset-bottom,0px) + 28px)",maxHeight:"92vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}><Row style={{marginBottom:16}}><B style={{fontSize:16.5,color:C.navy}}>{title}</B><button onClick={onClose} style={{background:"#F0F4F8",border:"none",borderRadius:20,width:38,height:38,fontSize:16,cursor:"pointer",color:"#555",flexShrink:0}}>✕</button></Row>{children}</div></div>);}
function F({label,children}){return <div style={{marginBottom:13}}><label style={{fontSize:12.5,color:"#666",fontWeight:600,display:"block",marginBottom:5}}>{label}</label>{children}</div>;}
function SaveBtn({color,onClick,children}){return <button onClick={onClick} style={{background:color,color:"#fff",border:"none",borderRadius:12,padding:15,fontSize:15.5,fontWeight:800,cursor:"pointer",width:"100%",marginTop:10,minHeight:50}}>{children}</button>;}
function SmartInput({value,onChange,placeholder,list,idPrefix}){const id=`${idPrefix}-dl`;return(<><input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} list={id} style={IS} autoComplete="off"/><datalist id={id}>{list.map(l=><option key={l} value={l}/>)}</datalist></>);}

function generatePDF(title,content){
  const win=window.open("","_blank");
  win.document.write(`<!DOCTYPE html><html><head><title>${title}</title><style>body{font-family:Arial,sans-serif;font-size:13px;color:#111;padding:30px;max-width:800px;margin:auto;}h1{font-size:20px;border-bottom:2px solid #111;padding-bottom:8px;}table{width:100%;border-collapse:collapse;margin-top:16px;}th{background:#0F1923;color:#E8C97E;padding:8px 10px;text-align:left;font-size:12px;}td{padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;}tr:nth-child(even){background:#f9f9f9;}.tot td{background:#0F1923;color:#E8C97E;font-weight:bold;border:none;}.hdr{display:flex;justify-content:space-between;margin-bottom:20px;}.co{font-size:22px;font-weight:bold;color:#0F1923;}</style></head><body><div class="hdr"><div><div class="co">Amihem Sales</div><div style="font-size:12px;color:#666">Trading Sales Report</div></div><div style="text-align:right;font-size:12px;color:#666">Date: ${new Date().toLocaleDateString("en-IN")}</div></div><h1>${title}</h1>${content}<script>window.onload=()=>{window.print();}<\/script></body></html>`);
  win.document.close();
}

function parseExcelData(rows){
  let hi=-1;
  for(let i=0;i<Math.min(5,rows.length);i++){const r=rows[i];if(r&&r.some(c=>String(c||"").toLowerCase().includes("customer")||String(c||"").toLowerCase().includes("product"))){hi=i;break;}}
  if(hi<0)return null;
  const hd=rows[hi].map(h=>String(h||"").trim().toLowerCase());
  const col=(ns)=>{for(const n of ns){const i=hd.findIndex(h=>h.includes(n));if(i>=0)return i;}return -1;};
  const iDate=col(["date"]);const iSup=col(["supplier"]);const iCust=col(["customer"]);const iCity=col(["city"]);const iMob=col(["mobile","phone"]);const iProd=col(["product","item"]);const iQty=col(["qty","quantity"]);const iRate=col(["rate","price"]);const iSale=col(["tradingsales","sales amount","sales"]);const iPay=col(["tradingpayment","payment"]);const iComm=col(["commission","comm"]);const iGST=col(["gst"]);const iRem=col(["remarks"]);
  const sales=[];const payments=[];const customers=new Map();const suppliers=new Set();const products=new Map();
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i];if(!r||r.every(c=>c===null||c===undefined||c===""))continue;
    const gV=(idx)=>idx>=0?r[idx]:null;
    const cn=String(gV(iCust)||"").trim();const sn=String(gV(iSup)||"").trim();const pn=String(gV(iProd)||"").trim();
    if(!cn||!pn||cn.toLowerCase().startsWith("total"))continue;
    const city=String(gV(iCity)||"").trim();const mob=String(gV(iMob)||"").replace(/\D/g,"").slice(-10);const gst=String(gV(iGST)||"").trim();const rem=String(gV(iRem)||"").trim();
    const qty=parseFloat(gV(iQty))||0;const rate=parseFloat(gV(iRate))||0;const sa=parseFloat(gV(iSale))||0;const pa=parseFloat(gV(iPay))||0;const comm=parseFloat(gV(iComm))||0;
    const jd=gV(iDate)?excelDateToJS(gV(iDate)):new Date();const ds=jd&&!isNaN(jd)?jd.toISOString().split("T")[0]:today();
    if(!customers.has(cn))customers.set(cn,{id:uid(),name:cn,type:"Trading",phone:mob,city,gstin:gst});
    if(sn)suppliers.add(sn);
    if(!products.has(pn))products.set(pn,{id:uid(),name:pn,supplierName:sn,unit:"Mtr"});
    if(sa>0){const cid=customers.get(cn)?.id||"";const bn=`IMP-${i}`;sales.push({id:uid(),date:ds,billNo:bn,customerId:cid,customerName:cn,supplierName:sn,productName:pn,meters:qty,rate,amount:sa,remarks:rem,commission:comm});if(pa>0)payments.push({id:uid(),date:ds,billNo:bn,customerId:cid,customerName:cn,amount:pa,mode:"Import",remarks:"Auto from Excel",commissionEarned:comm});}
  }
  return{customers:[...customers.values()],suppliers:[...suppliers].map(s=>({id:uid(),name:s,phone:"",city:""})),products:[...products.values()],tradingSales:sales,tradingPayments:payments};
}

const TABS=[
  {id:"Dashboard",icon:"📊",label:"Dashboard"},
  {id:"Trading",icon:"🏪",label:"Trading"},
  {id:"Returns",icon:"↩️",label:"Returns"},
  {id:"Outstanding",icon:"⏳",label:"Outstanding"},
  {id:"Aging",icon:"📅",label:"Ageing"},
  {id:"Analytics",icon:"📈",label:"Analytics"},
  {id:"Commission",icon:"💰",label:"Commission"},
  {id:"Masters",icon:"⚙️",label:"Masters"},
  {id:"Reports",icon:"📋",label:"Reports"},
];

export default function App(){
  const[tab,setTab]=useState("Dashboard");
  const[data,setData]=useState(null);
  const[modal,setModal]=useState(null);
  const[toast,setToast]=useState(null);
  const restoreRef=useRef(null);const importRef=useRef(null);

  useEffect(()=>{loadData().then(setData);},[]);
  useEffect(()=>{if(data)saveData(data);},[data]);

  const showToast=(msg,err)=>{setToast({msg,err});setTimeout(()=>setToast(null),3000);};
  const add=(section,rec)=>{setData(p=>({...p,[section]:[...p[section],{...rec,id:uid()}]}));showToast("Saved successfully.");setModal(null);};
  const del=(section,id)=>{setData(p=>({...p,[section]:p[section].filter(r=>r.id!==id)}));showToast("Deleted.");};
  const updateMaster=(section,updated)=>{setData(p=>({...p,[section]:p[section].map(r=>r.id===updated.id?updated:r)}));showToast("Updated.");setModal(null);};

  const handleExcelImport=(e)=>{
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
        const parsed=parseExcelData(rows);
        if(!parsed||parsed.tradingSales.length===0){showToast("Could not parse Excel.",true);return;}
        setData(p=>{
          const ec=new Set(p.customers.map(c=>c.name.toLowerCase()));
          const es=new Set(p.suppliers.map(s=>s.name.toLowerCase()));
          const ep=new Set(p.products.map(x=>x.name.toLowerCase()));
          return{...p,customers:[...p.customers,...parsed.customers.filter(c=>!ec.has(c.name.toLowerCase()))],suppliers:[...p.suppliers,...parsed.suppliers.filter(s=>!es.has(s.name.toLowerCase()))],products:[...p.products,...parsed.products.filter(x=>!ep.has(x.name.toLowerCase()))],tradingSales:[...p.tradingSales,...parsed.tradingSales],tradingPayments:[...p.tradingPayments,...parsed.tradingPayments]};
        });
        showToast(`Imported ${parsed.tradingSales.length} sales, ${parsed.customers.length} customers.`);
      }catch{showToast("Error reading file.",true);}
      e.target.value="";
    };
    reader.readAsArrayBuffer(file);
  };

  const importBackup=(e)=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{try{setData({...EMPTY,...JSON.parse(ev.target.result)});showToast("Backup restored.");}catch{showToast("Invalid backup.",true);}e.target.value="";};r.onerror=()=>{showToast("Could not read file.",true);e.target.value="";};r.readAsText(f);};

  if(!data)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:C.navy,flexDirection:"column",gap:12}}><div style={{fontSize:48}}>🧵</div><div style={{color:C.gold,fontWeight:800,fontSize:16}}>Loading…</div></div>);

  // ── Computed ─────────────────────────────────────────────────
  const tradingOut={};
  data.tradingSales.forEach(s=>{
    if(!tradingOut[s.customerName])tradingOut[s.customerName]={name:s.customerName,due:0,paid:0,commissionEarned:0,meters:0,phone:"",debit:0,credit:0};
    tradingOut[s.customerName].due+=+s.amount||0;
    tradingOut[s.customerName].meters+=+s.meters||0;
  });
  data.tradingPayments.forEach(p=>{
    if(!tradingOut[p.customerName])tradingOut[p.customerName]={name:p.customerName,due:0,paid:0,commissionEarned:0,meters:0,phone:"",debit:0,credit:0};
    tradingOut[p.customerName].paid+=+p.amount||0;
    tradingOut[p.customerName].commissionEarned+=+p.commissionEarned||0;
  });
  (data.debitNotes||[]).forEach(n=>{if(!tradingOut[n.customerName])tradingOut[n.customerName]={name:n.customerName,due:0,paid:0,commissionEarned:0,meters:0,phone:"",debit:0,credit:0};tradingOut[n.customerName].debit+=+n.amount||0;});
  (data.creditNotes||[]).forEach(n=>{if(!tradingOut[n.customerName])tradingOut[n.customerName]={name:n.customerName,due:0,paid:0,commissionEarned:0,meters:0,phone:"",debit:0,credit:0};tradingOut[n.customerName].credit+=+n.amount||0;});
  data.customers.forEach(c=>{if(tradingOut[c.name])tradingOut[c.name].phone=c.phone||"";});

  const totTradingSale=data.tradingSales.reduce((a,s)=>a+(+s.amount||0),0);
  const totTradingPaid=data.tradingPayments.reduce((a,p)=>a+(+p.amount||0),0);
  const totTradingOut=Object.values(tradingOut).reduce((a,v)=>a+Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0)),0);
  const totComm=calcFIFOCommission(data.tradingSales,data.tradingPayments).totalCommission;
  const overdueCount=Object.values(tradingOut).filter(v=>Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))>0).length;

  const Sidebar=()=>(
    <div style={{width:220,background:C.navy,minHeight:"100vh",position:"fixed",left:0,top:0,zIndex:150,display:"flex",flexDirection:"column",padding:"20px 0"}}>
      <div style={{padding:"0 20px 24px"}}><div style={{fontSize:10,letterSpacing:2.5,color:C.gold,textTransform:"uppercase",fontWeight:700}}>Amihem</div><div style={{fontSize:20,fontWeight:900,color:"#fff",marginTop:4}}>Sales</div></div>
      {TABS.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?"rgba(232,201,126,0.12)":"transparent",border:"none",borderLeft:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",padding:"12px 20px",textAlign:"left",color:tab===t.id?C.gold:"rgba(255,255,255,0.65)",fontSize:13.5,fontWeight:tab===t.id?700:400,cursor:"pointer",display:"flex",alignItems:"center",gap:10}}><span>{t.icon}</span>{t.label}{t.id==="Outstanding"&&overdueCount>0&&<span style={{background:C.red,color:"#fff",borderRadius:20,fontSize:10,fontWeight:800,padding:"2px 7px",marginLeft:"auto"}}>{overdueCount}</span>}</button>))}
      <div style={{marginTop:"auto",padding:"16px 20px",borderTop:"1px solid rgba(255,255,255,0.08)",display:"flex",flexDirection:"column",gap:8}}>
        <button onClick={()=>importRef.current&&importRef.current.click()} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📊 <span style={{fontSize:12}}>Import Excel</span></button>
        <button onClick={()=>exportBackup(data)} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>☁️ <span style={{fontSize:12}}>Backup</span></button>
        <button onClick={()=>restoreRef.current&&restoreRef.current.click()} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📂 <span style={{fontSize:12}}>Restore</span></button>
      </div>
    </div>
  );

  return(
    <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:C.bg,minHeight:"100vh"}}>
      <style>{`@media(min-width:768px){.mob{display:none!important}.desk{display:flex!important}.main{margin-left:220px!important}}@media(max-width:767px){.mob{display:flex!important}.desk{display:none!important}.main{margin-left:0!important;padding-bottom:80px!important}}.main{max-width:860px;padding:18px 16px 40px;}`}</style>
      <div className="desk" style={{display:"none"}}><Sidebar/></div>
      <div className="mob" style={{display:"none",background:C.navy,padding:"calc(env(safe-area-inset-top,0px)+12px) 16px 10px",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 20px rgba(0,0,0,0.4)",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <div><div style={{fontSize:9,letterSpacing:2,color:C.gold,textTransform:"uppercase",fontWeight:700}}>Amihem</div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>Sales</div></div>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>importRef.current&&importRef.current.click()} style={hdrBtn()}><span style={{fontSize:18}}>📊</span><span style={{fontSize:9}}>Excel</span></button>
          <button onClick={()=>exportBackup(data)} style={hdrBtn()}><span style={{fontSize:18}}>☁️</span><span style={{fontSize:9}}>Backup</span></button>
          <button onClick={()=>restoreRef.current&&restoreRef.current.click()} style={hdrBtn()}><span style={{fontSize:18}}>📂</span><span style={{fontSize:9}}>Restore</span></button>
        </div>
      </div>
      <div className="mob" style={{display:"none",background:C.navyMid,overflowX:"auto",WebkitOverflowScrolling:"touch",scrollbarWidth:"none",position:"sticky",top:72,zIndex:99}}>
        {TABS.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{flex:"0 0 auto",padding:"12px 14px",fontSize:12,fontWeight:tab===t.id?800:500,color:tab===t.id?C.gold:"rgba(255,255,255,0.6)",background:"none",border:"none",borderBottom:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",cursor:"pointer",whiteSpace:"nowrap",minHeight:44,position:"relative"}}>{t.icon} {t.label}{t.id==="Outstanding"&&overdueCount>0&&<span style={{position:"absolute",top:6,right:4,background:C.red,color:"#fff",borderRadius:10,fontSize:9,fontWeight:800,padding:"1px 5px"}}>{overdueCount}</span>}</button>))}
      </div>
      <div className="main" style={{marginLeft:0,padding:"18px 16px 40px",maxWidth:860}}>
        {tab==="Dashboard"  &&<DashboardTab data={data} tots={{totTradingSale,totTradingPaid,totTradingOut,totComm}} onNav={setTab} tradingOut={tradingOut}/>}
        {tab==="Trading"    &&<TradingTab data={data} onAdd={()=>setModal({type:"sale"})} onAddPay={()=>setModal({type:"payment"})} onDel={del} tradingOut={tradingOut}/>}
        {tab==="Returns"    &&<ReturnsTab data={data} onAddDebit={()=>setModal({type:"debit"})} onAddCredit={()=>setModal({type:"credit"})} onDel={del}/>}
        {tab==="Outstanding"&&<OutstandingTab tradingOut={tradingOut} data={data} onAddPay={(pre)=>setModal({type:"payment",preCustomer:pre})} generatePDF={generatePDF}/>}
        {tab==="Aging"      &&<AgingTab data={data} generatePDF={generatePDF}/>}
        {tab==="Analytics"  &&<AnalyticsTab data={data} tradingOut={tradingOut}/>}
        {tab==="Commission" &&<CommissionTab data={data} generatePDF={generatePDF}/>}
        {tab==="Masters"    &&<MastersTab data={data} onAdd={setModal} onDel={del} onEdit={(type,rec)=>setModal({type:`edit-${type}`,rec})} onImportExcel={()=>importRef.current&&importRef.current.click()}/>}
        {tab==="Reports"    &&<ReportsTab data={data} tradingOut={tradingOut} tots={{totTradingSale,totTradingPaid,totTradingOut,totComm}} generatePDF={generatePDF}/>}
      </div>
      {modal?.type==="sale"          &&<SaleModal    data={data} onSave={r=>add("tradingSales",r)}    onClose={()=>setModal(null)}/>}
      {modal?.type==="payment"       &&<PaymentModal data={data} onSave={r=>add("tradingPayments",r)} onClose={()=>setModal(null)} preCustomer={modal.preCustomer}/>}
      {modal?.type==="debit"         &&<DebitModal   data={data} onSave={r=>add("debitNotes",r)}      onClose={()=>setModal(null)}/>}
      {modal?.type==="credit"        &&<CreditModal  data={data} onSave={r=>add("creditNotes",r)}     onClose={()=>setModal(null)}/>}
      {modal?.type==="customer"      &&<CustomerModal onSave={r=>add("customers",r)}                  onClose={()=>setModal(null)}/>}
      {modal?.type==="supplier"      &&<SupplierModal onSave={r=>add("suppliers",r)}                  onClose={()=>setModal(null)}/>}
      {modal?.type==="product"       &&<ProductModal  data={data} onSave={r=>add("products",r)}       onClose={()=>setModal(null)}/>}
      {modal?.type==="edit-customer" &&<CustomerModal onSave={r=>updateMaster("customers",r)}         onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-supplier" &&<SupplierModal onSave={r=>updateMaster("suppliers",r)}         onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-product"  &&<ProductModal  data={data} onSave={r=>updateMaster("products",r)} onClose={()=>setModal(null)} initial={modal.rec}/>}
      <input ref={importRef} type="file" accept=".xlsx,.xls" onChange={handleExcelImport} style={{display:"none"}}/>
      <input ref={restoreRef} type="file" accept=".json" onChange={importBackup} style={{display:"none"}}/>
      {toast&&<div style={{position:"fixed",bottom:"calc(env(safe-area-inset-bottom,0px) + 24px)",left:"50%",transform:"translateX(-50%)",background:toast.err?"#B03A2E":C.navy,color:C.gold,padding:"11px 24px",borderRadius:24,fontSize:13.5,fontWeight:700,zIndex:999,boxShadow:"0 4px 20px rgba(0,0,0,0.4)",whiteSpace:"nowrap",border:"1px solid rgba(232,201,126,0.3)",maxWidth:"90%",textAlign:"center"}}>{toast.msg}</div>}
    </div>
  );
}

// ─── DASHBOARD ───────────────────────────────────────────────────
function DashboardTab({data,tots,onNav,tradingOut}){
  const{totTradingSale,totTradingPaid,totTradingOut,totComm}=tots;
  const recent=[...data.tradingSales].sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,8);
  const outList=Object.values(tradingOut).filter(v=>Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))>0).sort((a,b)=>b.due-b.paid-(a.due-a.paid)).slice(0,5);
  return(<div>
    <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:16,padding:"18px 20px",marginBottom:16,border:"1px solid rgba(232,201,126,0.25)"}}>
      <div style={{fontSize:10.5,color:C.gold,letterSpacing:2,textTransform:"uppercase",fontWeight:700}}>Commission Earned (on Payments)</div>
      <div style={{fontSize:34,fontWeight:900,color:C.gold,marginTop:6}}>₹{fmt(totComm)}</div>
      <div style={{fontSize:12,color:"rgba(255,255,255,0.45)",marginTop:4}}>Calculated on actual payments received</div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10,marginBottom:16}}>
      <KpiCard icon="🏪" label="Total Sales"  val={totTradingSale} color={C.blue}/>
      <KpiCard icon="✅" label="Total Paid"   val={totTradingPaid} color={C.green}/>
      <KpiCard icon="⏳" label="Outstanding"  val={totTradingOut}  color={C.red}/>
      <KpiCard icon="💰" label="Commission"   val={totComm}        color={C.purple}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:10,marginBottom:16}}>
      {[{icon:"📈",t:"Analytics",bg:C.blue},{icon:"⏳",t:"Outstanding",bg:C.red},{icon:"📅",t:"Aging",bg:C.orange},{icon:"📋",t:"Reports",bg:C.purple}].map(q=>(
        <button key={q.t} onClick={()=>onNav(q.t)} style={{background:q.bg,color:"#fff",border:"none",borderRadius:13,padding:"14px 10px",fontSize:13,fontWeight:700,cursor:"pointer",textAlign:"center"}}>
          <div style={{fontSize:22,marginBottom:4}}>{q.icon}</div>{q.t}
        </button>
      ))}
    </div>
    {outList.length>0&&<div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.red,marginBottom:12}}>Top Outstanding</div>
      {outList.map(v=>(<div key={v.name} style={{display:"flex",justifyContent:"space-between",padding:"9px 0",borderBottom:"1px solid #F5F7FA"}}><div style={{fontSize:13,fontWeight:600}}>{v.name}</div><div style={{fontWeight:900,color:C.red}}>₹{fmt(Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0)))}</div></div>))}
    </div>}
    <div style={{background:C.card,borderRadius:14,padding:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>Recent Sales</div>
      {recent.length===0&&<Empty text="No transactions yet. Import Excel or add a sale."/>}
      {recent.map(item=>(<div key={item.id} style={{display:"flex",justifyContent:"space-between",padding:"10px 0",borderBottom:"1px solid #F5F7FA"}}><div><div style={{fontSize:13.5,fontWeight:600}}>{item.customerName}</div><div style={{fontSize:11.5,color:C.muted}}>{item.billNo&&`${item.billNo} · `}{item.productName} · {fmtD(item.date)}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:800,color:C.blue}}>₹{fmt(item.amount)}</div><div style={{fontSize:10.5,color:C.muted}}>{fmt(item.meters)}m</div></div></div>))}
    </div>
  </div>);
}

// ─── TRADING TAB ─────────────────────────────────────────────────
function TradingTab({data,onAdd,onAddPay,onDel,tradingOut}){
  const[view,setView]=useState("sales");const[search,setSearch]=useState("");
  const sales=[...data.tradingSales].filter(s=>!search||[s.customerName,s.productName,s.billNo].some(x=>x?.toLowerCase().includes(search.toLowerCase()))).sort((a,b)=>new Date(b.date)-new Date(a.date));
  const payments=[...data.tradingPayments].sort((a,b)=>new Date(b.date)-new Date(a.date));
  const totSales=data.tradingSales.reduce((a,s)=>a+(+s.amount||0),0);
  const totPaid=data.tradingPayments.reduce((a,p)=>a+(+p.amount||0),0);
  const totOut=Object.values(tradingOut).reduce((a,v)=>a+Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0)),0);
  const totComm=calcFIFOCommission(data.tradingSales,data.tradingPayments).totalCommission;
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
      <KpiCard icon="🏪" label="Total Sales" val={totSales} color={C.blue}/>
      <KpiCard icon="✅" label="Total Paid"  val={totPaid}  color={C.green}/>
      <KpiCard icon="⏳" label="Outstanding" val={totOut}   color={C.red}/>
      <KpiCard icon="💰" label="Commission"  val={totComm}  color={C.purple}/>
    </div>
    <div style={{display:"flex",gap:8,marginBottom:12}}><Btn color={C.blue} onClick={onAdd}>+ Sale Entry</Btn><Btn color={C.green} onClick={onAddPay}>+ Payment</Btn></div>
    <SegCtrl options={[{v:"sales",l:`Sales (${data.tradingSales.length})`},{v:"payments",l:`Payments (${data.tradingPayments.length})`}]} val={view} onChange={setView}/>
    {view==="sales"&&<>
      <input placeholder="Search customer, product, bill no…" value={search} onChange={e=>setSearch(e.target.value)} style={{...IS,margin:"10px 0"}}/>
      {sales.length===0&&<Empty text="No sales yet. Import Excel or add manually."/>}
      {sales.map(s=>(<div key={s.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{s.customerName}</B><span style={{fontWeight:900,color:C.blue,fontSize:14}}>₹{fmt(s.amount)}</span></Row>
        {s.billNo&&<Mute>Bill No: {s.billNo}</Mute>}
        <Mute>{s.productName} · {s.supplierName}</Mute>
        <Mute>{fmt(s.meters)} m @ ₹{fmt(s.rate)}/m · {fmtD(s.date)}</Mute>
        <button onClick={()=>onDel("tradingSales",s.id)} style={{marginTop:8,background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
      </div>))}
    </>}
    {view==="payments"&&<>
      {payments.length===0&&<Empty text="No payments yet."/>}
      {payments.map(p=>(<div key={p.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{p.customerName}</B><span style={{fontWeight:900,color:C.green}}>₹{fmt(p.amount)}</span></Row>
        {p.billNo&&<Mute>Bill No: {p.billNo}</Mute>}
        <Mute>{p.mode} · {fmtD(p.date)}</Mute>
        {p.commissionEarned>0&&<Mute style={{color:C.purple}}>Commission: ₹{fmt(p.commissionEarned)}</Mute>}
        {p.remarks&&<Mute>{p.remarks}</Mute>}
        <button onClick={()=>onDel("tradingPayments",p.id)} style={{marginTop:8,background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
      </div>))}
    </>}
  </div>);
}

// ─── RETURNS TAB ─────────────────────────────────────────────────
function ReturnsTab({data,onAddDebit,onAddCredit,onDel}){
  const[view,setView]=useState("credit");
  const debitNotes=data.debitNotes||[];
  const creditNotes=data.creditNotes||[];
  const totDebit=debitNotes.reduce((a,n)=>a+(+n.amount||0),0);
  const totCredit=creditNotes.reduce((a,n)=>a+(+n.amount||0),0);
  return(<div>
    <div style={{background:"#FEF9E7",borderRadius:12,padding:"12px 14px",marginBottom:14,border:"1px solid #F9E79F"}}>
      <div style={{fontWeight:700,fontSize:13,color:"#7D6608",marginBottom:4}}>How Returns Work</div>
      <div style={{fontSize:12,color:"#7D6608",lineHeight:1.6}}><b>Debit Note</b> — Customer owes MORE (price difference, penalty, extra goods sent).<br/><b>Credit Note / Sale Return</b> — Customer owes LESS (goods returned, allowance given).</div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
      <div style={{background:C.card,borderRadius:13,padding:"13px 14px",borderLeft:`4px solid ${C.red}`}}><div style={{fontSize:10,color:C.muted,textTransform:"uppercase",fontWeight:700}}>Debit Notes</div><div style={{fontSize:20,fontWeight:900,color:C.red}}>₹{fmt(totDebit)}</div><div style={{fontSize:11,color:C.muted}}>{debitNotes.length} entries</div></div>
      <div style={{background:C.card,borderRadius:13,padding:"13px 14px",borderLeft:`4px solid ${C.green}`}}><div style={{fontSize:10,color:C.muted,textTransform:"uppercase",fontWeight:700}}>Credit / Return</div><div style={{fontSize:20,fontWeight:900,color:C.green}}>₹{fmt(totCredit)}</div><div style={{fontSize:11,color:C.muted}}>{creditNotes.length} entries</div></div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
      <button onClick={onAddDebit} style={{background:C.red,color:"#fff",border:"none",borderRadius:12,padding:"14px 10px",fontSize:13.5,fontWeight:700,cursor:"pointer",minHeight:50}}>+ Debit Note</button>
      <button onClick={onAddCredit} style={{background:C.green,color:"#fff",border:"none",borderRadius:12,padding:"14px 10px",fontSize:13.5,fontWeight:700,cursor:"pointer",minHeight:50}}>+ Credit / Return</button>
    </div>
    <SegCtrl options={[{v:"credit",l:`Returns (${creditNotes.length})`},{v:"debit",l:`Debit Notes (${debitNotes.length})`}]} val={view} onChange={setView}/>
    <div style={{marginTop:12}}>
      {view==="credit"&&<>
        {creditNotes.length===0&&<Empty text="No credit notes or sale returns yet."/>}
        {[...creditNotes].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(n=>(<div key={n.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",borderLeft:`4px solid ${C.green}`}}>
          <Row><div><B>{n.customerName}</B>{n.noteNo&&<Mute>Note No: {n.noteNo}</Mute>}</div><span style={{fontWeight:900,color:C.green,fontSize:15}}>₹{fmt(n.amount)}</span></Row>
          {n.originalBillNo&&<Mute>Against Bill: {n.originalBillNo}</Mute>}
          {n.productName&&<Mute>{n.productName}{n.meters?` · ${fmt(n.meters)} m`:""}</Mute>}
          <Mute>{fmtD(n.date)}</Mute>
          {n.remarks&&<Mute>{n.remarks}</Mute>}
          <div style={{background:"#E8F8F5",borderRadius:8,padding:"7px 10px",marginTop:8,fontSize:11.5,color:C.teal,fontWeight:600}}>Reduces customer outstanding</div>
          <button onClick={()=>onDel("creditNotes",n.id)} style={{marginTop:8,background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>))}
      </>}
      {view==="debit"&&<>
        {debitNotes.length===0&&<Empty text="No debit notes yet."/>}
        {[...debitNotes].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(n=>(<div key={n.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",borderLeft:`4px solid ${C.red}`}}>
          <Row><div><B>{n.customerName}</B>{n.noteNo&&<Mute>Note No: {n.noteNo}</Mute>}</div><span style={{fontWeight:900,color:C.red,fontSize:15}}>₹{fmt(n.amount)}</span></Row>
          {n.originalBillNo&&<Mute>Against Bill: {n.originalBillNo}</Mute>}
          <Mute>{fmtD(n.date)}</Mute>
          {n.remarks&&<Mute>{n.remarks}</Mute>}
          <div style={{background:"#FDEDEC",borderRadius:8,padding:"7px 10px",marginTop:8,fontSize:11.5,color:"#C0392B",fontWeight:600}}>Increases customer outstanding</div>
          <button onClick={()=>onDel("debitNotes",n.id)} style={{marginTop:8,background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>))}
      </>}
    </div>
  </div>);
}

// ─── OUTSTANDING TAB ─────────────────────────────────────────────
function OutstandingTab({tradingOut,data,onAddPay,generatePDF}){
  const[search,setSearch]=useState("");
  const[expanded,setExpanded]=useState(null);
  const todayMs=new Date().setHours(0,0,0,0);
  const phoneMap={};const creditDaysMap={};
  data.customers.forEach(c=>{phoneMap[c.name]=c.phone;creditDaysMap[c.name]=parseInt(c.creditDays||30);});

  const ageBucket=(days)=>{if(days<=30)return{label:"0-30 Days",color:C.green};if(days<=60)return{label:"31-60 Days",color:C.orange};if(days<=90)return{label:"61-90 Days",color:"#E67E22"};if(days<=120)return{label:"91-120 Days",color:C.red};return{label:"Above 120 Days",color:"#922B21"};};

  const entries=Object.values(tradingOut).map(v=>{
    const net=Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0));
    const custSales=data.tradingSales.filter(s=>s.customerName===v.name).sort((a,b)=>new Date(a.date)-new Date(b.date));
    let rem=v.paid;let maxDays=0;let oldestBill=null;let oldestDate=null;
    const openBills=[];
    custSales.forEach(s=>{
      const amt=+s.amount||0;
      const d=Math.min(amt,rem);rem-=d;const left=amt-d;
      if(left>0){
        const days=Math.floor((todayMs-new Date(s.date).setHours(0,0,0,0))/86400000);
        if(days>maxDays){maxDays=days;oldestBill=s.billNo;oldestDate=s.date;}
        openBills.push({billNo:s.billNo||"—",date:s.date,outstanding:left,days,billAmount:amt,productName:s.productName||"",meters:+s.meters||0,rate:+s.rate||0});
      }
    });
    return{...v,net,maxDays,oldestBill,oldestDate,openBills};
  }).filter(v=>v.net>0).filter(v=>!search||v.name.toLowerCase().includes(search.toLowerCase())).sort((a,b)=>b.net-a.net);

  const total=entries.reduce((a,v)=>a+v.net,0);

  const buildWA=(v)=>{
    const dateStr=new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
    const fmtAmt=(n)=>Math.round(n).toLocaleString("en-IN");
    const fmtBillDate=(d)=>{
      const dt=new Date(d);
      const months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
      return `${String(dt.getDate()).padStart(2,"0")}-${months[dt.getMonth()]}-${String(dt.getFullYear()).slice(2)}`;
    };
    const divider="━━━━━━━━━━━━━━━━━━━━━━━━━";
    const billLines=v.openBills.map((b,i)=>{
      const bucketShort=b.days<=30?"0-30d":b.days<=60?"31-60d":b.days<=90?"61-90d":b.days<=120?"91-120d":"120+d";
      const billNoClean=b.billNo==="—"?"—":String(b.billNo).replace(/IMP-/i,"");
      return `📌 *Bill ${i+1}*\n`+
        `   Date    : ${fmtBillDate(b.date)}\n`+
        `   Bill No : ${billNoClean}\n`+
        `   O/S     : ₹${fmtAmt(b.outstanding)}\n`+
        `   Ageing  : ${b.days} Days (${bucketShort})`;
    }).join("\n\n");
    return `🏢 *Navkar Fabrics*\n📅 ${dateStr}\n${divider}\n\nDear *${v.name}*,\n\nYour outstanding details:\n\n${billLines}\n\n${divider}\n💰 *Total O/S : ₹${fmt(v.net)}*\n${divider}\n\nPlease arrange payment at the earliest.\n\nThank You 🙏\n*Navkar Fabrics*`;
  };

  const doPDF=()=>{
    const partyBlocks=entries.map(v=>{
      const credit=v.openBills.map(b=>Math.max(0,(b.billAmount||b.outstanding)-b.outstanding));
      const rows=v.openBills.map((b,i)=>`<tr><td>${fmtD(b.date)}</td><td>${b.billNo}</td><td style="text-align:right">₹${fmt(b.billAmount||b.outstanding)}</td><td style="text-align:right">₹${fmt(credit[i])}</td><td style="text-align:right;color:#E74C3C;font-weight:bold">₹${fmt(b.outstanding)}</td><td style="text-align:right">${b.days}</td></tr>`).join("");
      return `<div class="party"><div class="party-name">Party : ${v.name}</div>
        <table><thead><tr><th>Bill Date</th><th>Bill No</th><th style="text-align:right">Amount</th><th style="text-align:right">Credit</th><th style="text-align:right">Outstanding</th><th style="text-align:right">Days</th></tr></thead>
        <tbody>${rows}<tr class="party-tot"><td colspan="4">Party Total</td><td style="text-align:right">₹${fmt(v.net)}</td><td></td></tr></tbody></table></div>`;
    }).join("");
    const style=`<style>.party{margin-bottom:22px;page-break-inside:avoid;}.party-name{font-weight:800;font-size:14px;color:#0F1923;margin-bottom:6px;padding-bottom:4px;border-bottom:1.5px solid #0F1923;}.party-tot td{background:#F0F4F8;font-weight:bold;border-top:1.5px solid #0F1923;}table th,table td{border:1px solid #ddd;}</style>`;
    generatePDF("Outstanding Statement",`${style}${partyBlocks}<div class="tot" style="display:flex;justify-content:space-between;background:#0F1923;color:#E8C97E;padding:12px 16px;border-radius:6px;font-weight:bold;font-size:15px;margin-top:10px;"><span>TOTAL OUTSTANDING</span><span>₹${fmt(total)}</span></div>`);
  };

  return(<div>
    <SecTitle>Customer-wise Outstanding</SecTitle>
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      <input placeholder="Search customer…" value={search} onChange={e=>setSearch(e.target.value)} style={{...IS,flex:1}}/>
      <button onClick={doPDF} style={{background:C.red,color:"#fff",border:"none",borderRadius:11,padding:"0 16px",fontSize:13,fontWeight:700,cursor:"pointer",flexShrink:0}}>PDF</button>
    </div>
    <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:14,padding:"14px 16px",marginBottom:16}}>
      <div style={{fontSize:10.5,color:C.gold,textTransform:"uppercase",letterSpacing:1,fontWeight:600}}>Total Outstanding ({entries.length} customers)</div>
      <div style={{fontSize:24,fontWeight:900,color:"#fff"}}>₹{fmt(total)}</div>
    </div>
    {entries.length===0&&<Empty text="All payments clear. No outstanding."/>}
    {entries.map(v=>{const bucket=ageBucket(v.maxDays);const phone=phoneMap[v.name]||v.phone||"";const isOpen=expanded===v.name;const creditDays=creditDaysMap[v.name]||30;return(
      <div key={v.name} style={{background:C.card,borderRadius:14,padding:"14px 15px",marginBottom:12,boxShadow:"0 1px 8px rgba(0,0,0,0.07)",borderLeft:`4px solid ${bucket.color}`}}>
        <Row><div><B style={{fontSize:15}}>{v.name}</B>{phone&&<Mute>{phone}</Mute>}</div><div style={{textAlign:"right"}}><div style={{fontWeight:900,fontSize:18,color:C.red}}>₹{fmt(v.net)}</div><span style={{fontSize:10,color:bucket.color,fontWeight:700}}>{bucket.label}</span></div></Row>
        <div style={{display:"flex",gap:14,marginTop:8,fontSize:12,color:C.muted,flexWrap:"wrap"}}>
          <span>Sales: <b style={{color:C.blue}}>₹{fmt(v.due)}</b></span>
          <span>Paid: <b style={{color:C.green}}>₹{fmt(v.paid)}</b></span>
          {v.debit>0&&<span>Debit: <b style={{color:C.red}}>₹{fmt(v.debit)}</b></span>}
          {v.credit>0&&<span>Credit: <b style={{color:C.teal}}>₹{fmt(v.credit)}</b></span>}
          <span>Credit Days: <b style={{color:C.navy}}>{creditDays}d</b></span>
          {v.maxDays>0&&<span>Overdue: <b style={{color:bucket.color}}>{v.maxDays}d</b></span>}
        </div>
        {isOpen&&<div style={{marginTop:12,paddingTop:10,borderTop:`1px solid ${C.border}`}}>
          <div style={{fontWeight:700,fontSize:12.5,color:C.navy,marginBottom:8}}>Bill-wise Outstanding</div>
          {v.openBills.map((b,i)=>{const bb=ageBucket(b.days);const credit=Math.max(0,(b.billAmount||b.outstanding)-b.outstanding);return(<div key={i} style={{background:"#F7F9FC",borderRadius:10,padding:"12px 13px",marginBottom:8,borderLeft:`3px solid ${bb.color}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:800,color:C.navy}}>Bill No. : {b.billNo}</div>
                <div style={{fontSize:11.5,color:C.muted,marginTop:3}}>Bill Date : {fmtD(b.date)}</div>
                {b.productName&&<div style={{fontSize:11.5,color:C.teal,marginTop:2}}>📦 {b.productName}</div>}
                {b.meters>0&&<div style={{fontSize:11.5,color:C.muted}}>Qty : {fmt(b.meters)} m {b.rate>0?`@ ₹${fmt(b.rate)}/m`:""}</div>}
                <div style={{fontSize:11.5,color:C.muted,marginTop:2}}>Bill Amt : <b style={{color:C.blue}}>₹{fmt(b.billAmount||b.outstanding)}</b></div>
                {credit>0&&<div style={{fontSize:11.5,color:C.muted}}>Credit : <b style={{color:C.green}}>₹{fmt(credit)}</b></div>}
                <div style={{fontSize:11.5,color:bb.color,fontWeight:700,marginTop:3}}>⏱ {b.days} Days · {bb.label}</div>
              </div>
              <div style={{textAlign:"right",minWidth:90}}>
                <div style={{fontWeight:900,fontSize:16,color:C.red}}>₹{fmt(b.outstanding)}</div>
                <div style={{fontSize:10,color:C.muted,marginTop:2}}>Outstanding</div>
              </div>
            </div>
          </div>);})}
        </div>}
        <div style={{display:"flex",gap:8,marginTop:12,flexWrap:"wrap"}}>
          <button onClick={()=>onAddPay(v.name)} style={{background:C.blue,color:"#fff",border:"none",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flex:1,minHeight:40}}>Record Payment</button>
          <button onClick={()=>setExpanded(isOpen?null:v.name)} style={{background:"#F0F4F8",color:C.navy,border:"none",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",minHeight:40}}>{isOpen?"Hide Bills":"View Bills"}</button>
          <button onClick={()=>{const n=phone?"91"+String(phone).replace(/\D/g,"").slice(-10):"";window.open(`https://wa.me/${n}?text=${encodeURIComponent(buildWA(v))}`,"_blank");}} style={{background:"#E8FBF0",color:"#25D366",border:"1px solid #25D366",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flex:1,minHeight:40}}>💬 WhatsApp</button>
        </div>
      </div>
    );})}
  </div>);
}

// ─── AGING TAB ───────────────────────────────────────────────────
function AgingTab({data,generatePDF}){
  const[search,setSearch]=useState("");const[expanded,setExpanded]=useState(null);
  const todayMs=new Date().setHours(0,0,0,0);
  const entries=useMemo(()=>{
    const map={};
    data.tradingSales.forEach(s=>{if(!map[s.customerName])map[s.customerName]={name:s.customerName,invoices:[]};map[s.customerName].invoices.push({date:s.date,amount:+s.amount||0,id:s.id,productName:s.productName,billNo:s.billNo});});
    const paid={};data.tradingPayments.forEach(p=>{paid[p.customerName]=(paid[p.customerName]||0)+(+p.amount||0);});
    return Object.values(map).map(v=>{
      let rem=paid[v.name]||0;const bk={b0:0,b30:0,b60:0,b90:0,b120:0};const openInvoices=[];
      [...v.invoices].sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach(inv=>{
        let amt=inv.amount;const d=Math.min(amt,rem);amt-=d;rem-=d;if(amt<=0)return;
        const days=Math.floor((todayMs-new Date(inv.date).setHours(0,0,0,0))/86400000);
        let bucket;if(days<=30){bk.b0+=amt;bucket="0-30 Days";}else if(days<=60){bk.b30+=amt;bucket="31-60 Days";}else if(days<=90){bk.b60+=amt;bucket="61-90 Days";}else if(days<=120){bk.b90+=amt;bucket="91-120 Days";}else{bk.b120+=amt;bucket="120+ Days";}
        openInvoices.push({...inv,outstanding:amt,days,bucket});
      });
      const total=Object.values(bk).reduce((a,b)=>a+b,0);
      return{...v,bk,total,openInvoices};
    }).filter(v=>v.total>0);
  },[data]);

  const filtered=entries.filter(v=>!search||v.name.toLowerCase().includes(search.toLowerCase())).sort((a,b)=>b.total-a.total);
  const bConfig=[{key:"b0",label:"0-30 Days",color:C.green},{key:"b30",label:"31-60 Days",color:C.orange},{key:"b60",label:"61-90 Days",color:"#E67E22"},{key:"b90",label:"91-120 Days",color:C.red},{key:"b120",label:"120+ Days",color:"#922B21"}];
  const bt=filtered.reduce((acc,v)=>{Object.keys(v.bk).forEach(k=>acc[k]=(acc[k]||0)+v.bk[k]);return acc;},{});

  const doPDF=()=>{
    const rows=filtered.map(v=>`<tr><td>${v.name}</td><td style="text-align:right">₹${fmt(v.bk.b0)}</td><td style="text-align:right">₹${fmt(v.bk.b30)}</td><td style="text-align:right">₹${fmt(v.bk.b60)}</td><td style="text-align:right">₹${fmt(v.bk.b90)}</td><td style="text-align:right;color:#922B21;font-weight:bold">₹${fmt(v.bk.b120)}</td><td style="text-align:right;font-weight:bold">₹${fmt(v.total)}</td></tr>`).join("");
    generatePDF("Ageing Analysis",`<table><thead><tr><th>Customer</th><th>0-30d</th><th>31-60d</th><th>61-90d</th><th>91-120d</th><th>120d+</th><th>Total</th></tr></thead><tbody>${rows}</tbody></table>`);
  };

  return(<div>
    <SecTitle>Ageing Analysis</SecTitle>
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      <input placeholder="Search customer…" value={search} onChange={e=>setSearch(e.target.value)} style={{...IS,flex:1}}/>
      <button onClick={doPDF} style={{background:C.red,color:"#fff",border:"none",borderRadius:11,padding:"0 16px",fontSize:13,fontWeight:700,cursor:"pointer",flexShrink:0}}>PDF</button>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:8,marginBottom:14}}>
      {bConfig.map(b=>(<div key={b.key} style={{background:C.card,borderRadius:12,padding:"10px 12px",borderLeft:`4px solid ${b.color}`}}><div style={{fontSize:10,color:b.color,fontWeight:700,textTransform:"uppercase"}}>{b.label}</div><div style={{fontSize:15,fontWeight:900,color:b.color,marginTop:3}}>₹{fmt(bt[b.key]||0)}</div></div>))}
    </div>
    {filtered.length===0&&<Empty text="No outstanding found."/>}
    {filtered.map(v=>{const isOpen=expanded===v.name;return(
      <div key={v.name} style={{background:C.card,borderRadius:14,padding:"14px 15px",marginBottom:11,boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
        <div onClick={()=>setExpanded(isOpen?null:v.name)} style={{cursor:"pointer"}}><Row><B style={{fontSize:15}}>{v.name}</B><span style={{fontWeight:900,fontSize:17,color:C.red}}>₹{fmt(v.total)}</span></Row></div>
        <div style={{marginTop:10}}>
          {bConfig.map(b=>v.bk[b.key]>0&&(<div key={b.key} style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}><div style={{fontSize:10.5,color:b.color,fontWeight:700,width:72,flexShrink:0}}>{b.label}</div><div style={{flex:1,background:"#F0F4F8",borderRadius:5,height:7}}><div style={{width:`${Math.min(100,v.bk[b.key]/v.total*100)}%`,background:b.color,height:7,borderRadius:5}}/></div><div style={{fontSize:12,fontWeight:700,color:b.color,width:80,textAlign:"right",flexShrink:0}}>₹{fmt(v.bk[b.key])}</div></div>))}
        </div>
        {isOpen&&<div style={{marginTop:10,paddingTop:10,borderTop:"1px solid #F0F4F8"}}>
          {v.openInvoices.sort((a,b)=>b.days-a.days).map((inv,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #F7F9FB"}}><div><div style={{fontSize:12.5,fontWeight:600}}>{inv.productName||"—"}{inv.billNo&&<span style={{color:C.muted,fontWeight:400}}> · {inv.billNo}</span>}</div><div style={{fontSize:10.5,color:"#bbb"}}>{fmtD(inv.date)} · {inv.days} days · {inv.bucket}</div></div><div style={{fontWeight:800,fontSize:12.5,color:C.red,alignSelf:"center"}}>₹{fmt(inv.outstanding)}</div></div>))}
        </div>}
        <button onClick={()=>setExpanded(isOpen?null:v.name)} style={{marginTop:10,background:"#F0F4F8",color:C.navy,border:"none",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>{isOpen?"Hide Invoices":"View Invoices"}</button>
      </div>
    );})}
  </div>);
}

// ─── ANALYTICS TAB ───────────────────────────────────────────────
function AnalyticsTab({data,tradingOut}){
  const[chartType,setChartType]=useState("monthly");
  const[fy,setFY]=useState(ALL_FY);
  const allFYs=useMemo(()=>getAvailableFYs(data.tradingSales),[data]);
  const filtered=useMemo(()=>filterByFY(data.tradingSales,fy),[data,fy]);
  const filteredPay=useMemo(()=>{
    if(fy===ALL_FY)return data.tradingPayments;
    const{start,end}=getFYRange(fy);
    return data.tradingPayments.filter(p=>{const d=new Date(p.date);return d>=start&&d<=end;});
  },[data,fy]);
  const PIE=[C.blue,C.green,C.orange,C.purple,C.teal,C.red,"#F39C12","#1ABC9C"];

  // Monthly sales for selected FY
  const monthly=useMemo(()=>{
    const map={};
    filtered.forEach(s=>{
      const d=new Date(s.date);
      const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      const label=d.toLocaleDateString("en-IN",{month:"short",year:"2-digit"});
      if(!map[key])map[key]={key,label,sales:0,payments:0};
      map[key].sales+=+s.amount||0;
    });
    filteredPay.forEach(p=>{
      const d=new Date(p.date);
      const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      if(map[key])map[key].payments+=+p.amount||0;
    });
    return Object.values(map).sort((a,b)=>a.key.localeCompare(b.key));
  },[filtered,filteredPay]);

  // YoY comparison — all 4 FYs, April–March
  const yoy=useMemo(()=>{
    const months=["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"];
    const fyList=allFYs.filter(f=>f!==ALL_FY).slice(0,4).reverse();
    const map={};
    data.tradingSales.forEach(s=>{
      const f=getFY(s.date);if(!fyList.includes(f))return;
      const d=new Date(s.date);
      const mIdx=(d.getMonth()+9)%12;// Apr=0
      const label=months[mIdx];
      if(!map[label])map[label]={label};
      map[label][f]=(map[label][f]||0)+(+s.amount||0);
    });
    return months.map(m=>map[m]||{label:m});
  },[data,allFYs]);

  // Top customers for selected FY
  const topCustomers=useMemo(()=>{
    const map={};
    filtered.forEach(s=>{if(!map[s.customerName])map[s.customerName]={name:s.customerName.length>16?s.customerName.slice(0,16)+"…":s.customerName,sales:0};map[s.customerName].sales+=+s.amount||0;});
    return Object.values(map).sort((a,b)=>b.sales-a.sales).slice(0,8);
  },[filtered]);

  // Products for selected FY
  const products=useMemo(()=>{
    const map={};
    filtered.forEach(s=>{const p=s.productName||"Unknown";if(!map[p])map[p]={name:p.length>16?p.slice(0,16)+"…":p,value:0};map[p].value+=+s.amount||0;});
    return Object.values(map).sort((a,b)=>b.value-a.value).slice(0,7);
  },[filtered]);

  // FY summary cards
  const fySummary=useMemo(()=>{
    const fyList=allFYs.filter(f=>f!==ALL_FY);
    return fyList.map(f=>{
      const sales=filterByFY(data.tradingSales,f).reduce((a,s)=>a+(+s.amount||0),0);
      const {start,end}=getFYRange(f);
      const pays=data.tradingPayments.filter(p=>{const d=new Date(p.date);return d>=start&&d<=end;}).reduce((a,p)=>a+(+p.amount||0),0);
      return{fy:f,sales,pays,outstanding:Math.max(0,sales-pays)};
    });
  },[data,allFYs]);

  const fyColors=["#2980B9","#27AE60","#E67E22","#8E44AD"];
  const fyList=allFYs.filter(f=>f!==ALL_FY).slice(0,4).reverse();

  const totSales=filtered.reduce((a,s)=>a+(+s.amount||0),0);
  const totPays=filteredPay.reduce((a,p)=>a+(+p.amount||0),0);

  return(<div>
    {/* FY Selector */}
    <div style={{marginBottom:14}}>
      <div style={{fontSize:11,color:C.muted,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>Financial Year</div>
      <div style={{display:"flex",overflowX:"auto",gap:6,scrollbarWidth:"none",paddingBottom:2}}>
        {allFYs.map(f=>(
          <button key={f} onClick={()=>setFY(f)} style={{flex:"0 0 auto",padding:"8px 14px",borderRadius:20,fontSize:12,fontWeight:fy===f?700:500,border:`1.5px solid ${fy===f?C.navy:C.border}`,background:fy===f?C.navy:"#fff",color:fy===f?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>
            {f===ALL_FY?"📊 All Years":`FY ${f}`}
          </button>
        ))}
      </div>
    </div>

    {/* KPI Strip */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
      <div style={{background:C.card,borderRadius:12,padding:"12px 14px",borderLeft:`4px solid ${C.blue}`}}>
        <div style={{fontSize:10,color:C.muted,fontWeight:600,textTransform:"uppercase"}}>Sales {fy===ALL_FY?"(All)":fy}</div>
        <div style={{fontSize:18,fontWeight:900,color:C.blue,marginTop:2}}>₹{totSales>=10000000?`${(totSales/10000000).toFixed(2)}Cr`:totSales>=100000?`${(totSales/100000).toFixed(1)}L`:fmt(totSales)}</div>
      </div>
      <div style={{background:C.card,borderRadius:12,padding:"12px 14px",borderLeft:`4px solid ${C.green}`}}>
        <div style={{fontSize:10,color:C.muted,fontWeight:600,textTransform:"uppercase"}}>Collected</div>
        <div style={{fontSize:18,fontWeight:900,color:C.green,marginTop:2}}>₹{totPays>=10000000?`${(totPays/10000000).toFixed(2)}Cr`:totPays>=100000?`${(totPays/100000).toFixed(1)}L`:fmt(totPays)}</div>
      </div>
    </div>

    {/* Chart Type Selector */}
    <div style={{display:"flex",overflowX:"auto",gap:6,marginBottom:12,scrollbarWidth:"none"}}>
      {[{v:"monthly",l:"📅 Monthly"},{v:"trend",l:"📈 Trend"},{v:"yoy",l:"📊 Year-on-Year"},{v:"customer",l:"👥 Customers"},{v:"product",l:"📦 Products"},{v:"fycompare",l:"📈 FY Summary"}].map(ct=>(
        <button key={ct.v} onClick={()=>setChartType(ct.v)} style={{flex:"0 0 auto",padding:"9px 13px",borderRadius:20,fontSize:12,fontWeight:chartType===ct.v?700:500,border:`1.5px solid ${chartType===ct.v?C.navy:C.border}`,background:chartType===ct.v?C.navy:"#fff",color:chartType===ct.v?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap",minHeight:38}}>{ct.l}</button>
      ))}
    </div>

    <div style={{background:C.card,borderRadius:14,padding:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      {chartType==="monthly"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Monthly Sales {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Sales vs Collections per month</div>
        {monthly.length===0?<Empty text="No data for selected year."/>:
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={monthly} margin={{top:4,right:4,left:0,bottom:4}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
            <Bar dataKey="sales" fill={C.blue} radius={[4,4,0,0]} name="Sales" barSize={14}/>
            <Bar dataKey="payments" fill={C.green} radius={[4,4,0,0]} name="Collections" barSize={14}/>
          </BarChart>
        </ResponsiveContainer>}
      </>}

      {chartType==="trend"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Sales Trend {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Month-over-month line trend — Sales vs Collections</div>
        {monthly.length===0?<Empty text="No data for selected year."/>:
        <ResponsiveContainer width="100%" height={260}>
          <LineChart data={monthly} margin={{top:4,right:8,left:0,bottom:4}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
            <Line type="monotone" dataKey="sales" stroke={C.blue} strokeWidth={2.5} dot={{r:4,fill:C.blue}} name="Sales"/>
            <Line type="monotone" dataKey="payments" stroke={C.green} strokeWidth={2.5} dot={{r:4,fill:C.green}} name="Collections"/>
          </LineChart>
        </ResponsiveContainer>}
      </>}

      {chartType==="yoy"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Year-on-Year Comparison</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>All financial years — April to March</div>
        {yoy.length===0?<Empty text="No data."/>:
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={yoy} margin={{top:4,right:4,left:0,bottom:4}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
            {fyList.map((f,i)=><Bar key={f} dataKey={f} fill={fyColors[i%fyColors.length]} radius={[3,3,0,0]} name={`FY ${f}`} barSize={8}/>)}
          </BarChart>
        </ResponsiveContainer>}
      </>}

      {chartType==="customer"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Top Customers {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>By sales volume</div>
        {topCustomers.length===0?<Empty text="No data."/>:
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={topCustomers} layout="vertical" margin={{top:0,right:8,left:70,bottom:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8" horizontal={false}/>
            <XAxis type="number" tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:9}} tickLine={false}/>
            <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={70} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Bar dataKey="sales" fill={C.blue} radius={[0,5,5,0]} name="Sales" barSize={12}/>
          </BarChart>
        </ResponsiveContainer>}
      </>}

      {chartType==="product"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Sales by Product {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Top 7 products by value</div>
        {products.length===0?<Empty text="No data."/>:
        <ResponsiveContainer width="100%" height={260}>
          <PieChart>
            <Pie data={products} cx="50%" cy="50%" outerRadius={95} innerRadius={40} dataKey="value" nameKey="name"
              label={({name,percent})=>`${(percent*100).toFixed(0)}%`} labelLine={true} fontSize={10}>
              {products.map((_,i)=><Cell key={i} fill={PIE[i%PIE.length]}/>)}
            </Pie>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:10}}/>
          </PieChart>
        </ResponsiveContainer>}
      </>}

      {chartType==="fycompare"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Financial Year Summary</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Sales vs Collections — all years</div>
        {fySummary.map((f,i)=>(
          <div key={f.fy} style={{background:"#F8FAFC",borderRadius:12,padding:"14px 15px",marginBottom:10,borderLeft:`4px solid ${fyColors[i%fyColors.length]}`}}>
            <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:8}}>FY {f.fy}</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
              <div><div style={{fontSize:10,color:C.muted,fontWeight:600}}>SALES</div><div style={{fontSize:14,fontWeight:900,color:C.blue}}>₹{f.sales>=10000000?`${(f.sales/10000000).toFixed(2)}Cr`:f.sales>=100000?`${(f.sales/100000).toFixed(1)}L`:fmt(f.sales)}</div></div>
              <div><div style={{fontSize:10,color:C.muted,fontWeight:600}}>COLLECTED</div><div style={{fontSize:14,fontWeight:900,color:C.green}}>₹{f.pays>=10000000?`${(f.pays/10000000).toFixed(2)}Cr`:f.pays>=100000?`${(f.pays/100000).toFixed(1)}L`:fmt(f.pays)}</div></div>
              <div><div style={{fontSize:10,color:C.muted,fontWeight:600}}>O/S</div><div style={{fontSize:14,fontWeight:900,color:C.red}}>₹{f.outstanding>=100000?`${(f.outstanding/100000).toFixed(1)}L`:fmt(f.outstanding)}</div></div>
            </div>
            <div style={{marginTop:10,background:"#E8EDF2",borderRadius:5,height:6}}>
              <div style={{width:`${Math.min(100,f.pays/f.sales*100)||0}%`,background:fyColors[i%fyColors.length],height:6,borderRadius:5}}/>
            </div>
            <div style={{fontSize:10,color:C.muted,marginTop:4}}>{f.sales>0?(f.pays/f.sales*100).toFixed(1):0}% collected</div>
          </div>
        ))}
      </>}
    </div>
  </div>);
}

// ─── COMMISSION TAB ──────────────────────────────────────────────
function calcFIFOCommission(tradingSales,tradingPayments){
  const customerMap={};
  tradingSales.forEach(s=>{
    const cid=s.customerName;
    if(!customerMap[cid])customerMap[cid]={sales:[],totalPaid:0};
    customerMap[cid].sales.push({...s});
  });
  tradingPayments.forEach(p=>{
    const cid=p.customerName;
    if(!customerMap[cid])customerMap[cid]={sales:[],totalPaid:0};
    customerMap[cid].totalPaid+=+p.amount||0;
  });
  let totalCommission=0;
  const commDetails=[];
  const commByCustomer={};
  Object.entries(customerMap).forEach(([cname,cdata])=>{
    const sales=[...cdata.sales].sort((a,b)=>new Date(a.date)-new Date(b.date));
    let remainingPayment=cdata.totalPaid;
    sales.forEach(sale=>{
      const saleAmt=+sale.amount||0;
      const qty=+sale.meters||+sale.qty||0;
      if(remainingPayment>=saleAmt&&saleAmt>0){
        remainingPayment-=saleAmt;
        const comm=qty*1.5;
        if(comm>0){
          totalCommission+=comm;
          commDetails.push({customer:cname,date:sale.date,billNo:sale.billNo||"—",qty,amount:saleAmt,commission:comm});
          if(!commByCustomer[cname])commByCustomer[cname]={name:cname,commission:0,bills:0,meters:0};
          commByCustomer[cname].commission+=comm;
          commByCustomer[cname].bills+=1;
          commByCustomer[cname].meters+=qty;
        }
      }
    });
  });
  return{totalCommission,commDetails,commByCustomer};
}

function CommissionTab({data,generatePDF}){
  const[view,setView]=useState("summary");
  const{totalCommission,commDetails,commByCustomer}=calcFIFOCommission(data.tradingSales,data.tradingPayments);
  const commList=Object.values(commByCustomer).sort((a,b)=>b.commission-a.commission);
  const detailList=[...commDetails].sort((a,b)=>new Date(b.date)-new Date(a.date));
  const totalBills=commList.reduce((a,v)=>a+v.bills,0);
  const totalMeters=commList.reduce((a,v)=>a+v.meters,0);
  const maxComm=commList[0]?.commission||1;

  const doPDF=()=>{
    if(view==="summary"){
      const rows=commList.map((v,i)=>`<tr><td>${i+1}</td><td>${v.name}</td><td style="text-align:right">${v.bills}</td><td style="text-align:right">${fmt(v.meters)} m</td><td style="text-align:right;color:#8E44AD;font-weight:bold">₹${fmt(v.commission)}</td></tr>`).join("");
      generatePDF("Commission Statement — Customer Summary",`<table><thead><tr><th>#</th><th>Customer</th><th>Bills Adjusted</th><th>Meters</th><th>Commission</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="4">TOTAL</td><td>₹${fmt(totalCommission)}</td></tr></tbody></table>`);
    } else {
      const rows=detailList.map(v=>`<tr><td>${fmtD(v.date)}</td><td>${v.billNo}</td><td>${v.customer}</td><td style="text-align:right">${fmt(v.qty)} m</td><td style="text-align:right">₹${fmt(v.amount)}</td><td style="text-align:right;color:#8E44AD;font-weight:bold">₹${fmt(v.commission)}</td></tr>`).join("");
      generatePDF("Commission Statement — Bill-wise Detail",`<table><thead><tr><th>Date</th><th>Bill No</th><th>Customer</th><th>Meters</th><th>Sale Amt</th><th>Commission</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="5">TOTAL</td><td>₹${fmt(totalCommission)}</td></tr></tbody></table>`);
    }
  };

  return(<div>
    {/* Hero Banner */}
    <div style={{background:`linear-gradient(135deg,${C.navy} 0%,#1E3A5F 100%)`,borderRadius:16,padding:"18px 16px",marginBottom:14,border:"1px solid rgba(232,201,126,0.3)",boxShadow:"0 4px 20px rgba(0,0,0,0.15)"}}>
      <div style={{fontSize:10,color:"rgba(232,201,126,0.7)",textTransform:"uppercase",letterSpacing:1.5,fontWeight:700,marginBottom:4}}>💰 Commission Earned (FIFO)</div>
      <div style={{fontSize:32,fontWeight:900,color:C.gold,letterSpacing:-0.5}}>₹{fmt(totalCommission)}</div>
      <div style={{fontSize:11,color:"rgba(255,255,255,0.4)",marginTop:4}}>Only fully adjusted bills · ₹1.50 per meter</div>
      <div style={{display:"flex",gap:16,marginTop:14,paddingTop:12,borderTop:"1px solid rgba(255,255,255,0.08)"}}>
        <div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>{commList.length}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Customers</div></div>
        <div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>{totalBills}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Bills Cleared</div></div>
        <div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>{fmt(totalMeters)} m</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Total Meters</div></div>
      </div>
    </div>

    {/* Action Buttons */}
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      <button onClick={()=>exportCSV([["#","Customer","Bills","Meters","Commission(₹)"],...commList.map((v,i)=>[i+1,v.name,v.bills,v.meters,v.commission])],`Commission_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"10px 16px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:42,display:"flex",alignItems:"center",gap:6}}>📥 CSV</button>
      <button onClick={doPDF} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"10px 16px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:42,display:"flex",alignItems:"center",gap:6}}>🖨️ PDF</button>
    </div>

    <SegCtrl options={[{v:"summary",l:"Customer Summary"},{v:"bills",l:"Bill-wise Detail"}]} val={view} onChange={setView}/>

    <div style={{marginTop:12}}>
    {view==="summary"&&<>
      {commList.length===0&&<Empty text="No commission yet. Bills need to be fully paid (FIFO) for commission."/>}
      {commList.map((v,i)=>{
        const pct=Math.round(v.commission/maxComm*100);
        const colors=[C.purple,C.blue,C.teal,C.orange,C.green,C.red];
        const col=colors[i%colors.length];
        return(
          <div key={v.name} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:9,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",border:`1px solid ${C.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
              <div style={{flex:1,paddingRight:10}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}>
                  <div style={{width:26,height:26,borderRadius:8,background:col,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:11,fontWeight:900,flexShrink:0}}>#{i+1}</div>
                  <B style={{fontSize:13.5,color:C.navy}}>{v.name}</B>
                </div>
                <div style={{display:"flex",gap:12,fontSize:11.5,color:C.muted,marginTop:4,marginLeft:34}}>
                  <span>📄 {v.bills} bills</span>
                  <span>📏 {fmt(v.meters)} m</span>
                </div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontWeight:900,fontSize:17,color:col}}>₹{fmt(v.commission)}</div>
                <div style={{fontSize:10,color:C.muted,marginTop:1}}>{(v.commission/totalCommission*100).toFixed(1)}% of total</div>
              </div>
            </div>
            <div style={{background:"#F0F4F8",borderRadius:6,height:6,overflow:"hidden"}}>
              <div style={{width:`${pct}%`,background:`linear-gradient(90deg,${col},${col}aa)`,height:6,borderRadius:6,transition:"width 0.3s"}}/>
            </div>
          </div>
        );
      })}
      <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:13,padding:"16px",marginTop:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:11,color:"rgba(255,255,255,0.5)",fontWeight:600}}>TOTAL COMMISSION</div>
          <div style={{fontSize:11,color:"rgba(255,255,255,0.35)",marginTop:1}}>{totalBills} bills · {fmt(totalMeters)} m</div>
        </div>
        <span style={{fontWeight:900,fontSize:20,color:C.gold}}>₹{fmt(totalCommission)}</span>
      </div>
    </>}

    {view==="bills"&&<>
      {detailList.length===0&&<Empty text="No fully adjusted bills found."/>}
      {detailList.map((v,i)=>(
        <div key={i} style={{background:C.card,borderRadius:12,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",border:`1px solid ${C.border}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1,paddingRight:10}}>
              <B style={{fontSize:13,color:C.navy}}>{v.customer}</B>
              <Mute>📋 Bill : {v.billNo}</Mute>
              <Mute>📅 Date : {fmtD(v.date)}</Mute>
              <Mute>📏 {fmt(v.qty)} m × ₹1.50</Mute>
              <Mute>🏷️ Sale : ₹{fmt(v.amount)}</Mute>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontWeight:900,color:C.purple,fontSize:16}}>₹{fmt(v.commission)}</div>
              <div style={{fontSize:10,color:C.muted,marginTop:2}}>Commission</div>
            </div>
          </div>
        </div>
      ))}
      <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:13,padding:"16px",marginTop:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontWeight:700,color:C.gold}}>TOTAL COMMISSION</span>
        <span style={{fontWeight:900,fontSize:20,color:C.gold}}>₹{fmt(totalCommission)}</span>
      </div>
    </>}
    </div>
  </div>);
}

// ─── MASTERS TAB ─────────────────────────────────────────────────
function MastersTab({data,onAdd,onDel,onEdit,onImportExcel}){
  const[view,setView]=useState("customers");
  return(<div>
    <div style={{background:`linear-gradient(135deg,${C.navyMid},${C.navy})`,borderRadius:12,padding:"14px 16px",marginBottom:14,border:"1px solid rgba(232,201,126,0.3)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.gold,marginBottom:6}}>Import from Excel</div>
      <div style={{fontSize:12,color:"rgba(255,255,255,0.6)",marginBottom:10}}>All masters auto-created from your Excel file.</div>
      <button onClick={onImportExcel} style={{background:C.gold,color:C.navy,border:"none",borderRadius:10,padding:"12px 20px",fontSize:14,fontWeight:800,cursor:"pointer",width:"100%",minHeight:46}}>Import Excel File (.xlsx)</button>
    </div>
    <SegCtrl options={[{v:"customers",l:`Customers (${data.customers.length})`},{v:"suppliers",l:`Suppliers (${data.suppliers.length})`},{v:"products",l:`Products (${data.products.length})`}]} val={view} onChange={setView}/>
    {view==="customers"&&<>
      <div style={{margin:"12px 0 8px"}}><Btn color={C.navy} onClick={()=>onAdd({type:"customer"})}>+ Add Customer</Btn></div>
      {data.customers.length===0&&<Empty text="No customers yet. Import Excel to auto-populate."/>}
      {data.customers.map(c=>(<div key={c.id} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{c.name}</B><span style={{fontSize:11,background:"#F0F4F8",padding:"3px 9px",borderRadius:10,color:C.muted,fontWeight:600}}>{c.type}</span></Row>
        {c.phone&&<Mute>{c.phone}</Mute>}{c.city&&<Mute>{c.city}</Mute>}{c.gstin&&<Mute>GST: {c.gstin}</Mute>}
        <Mute>Credit Days: {c.creditDays||"30"} days</Mute>
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>onEdit("customer",c)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Edit</button>
          <button onClick={()=>onDel("customers",c.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>
      </div>))}
    </>}
    {view==="suppliers"&&<>
      <div style={{margin:"12px 0 8px"}}><Btn color={C.navy} onClick={()=>onAdd({type:"supplier"})}>+ Add Supplier</Btn></div>
      {data.suppliers.length===0&&<Empty text="No suppliers yet."/>}
      {data.suppliers.map(s=>(<div key={s.id} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <B>{s.name}</B>{s.phone&&<Mute>{s.phone}</Mute>}{s.city&&<Mute>{s.city}</Mute>}
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>onEdit("supplier",s)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Edit</button>
          <button onClick={()=>onDel("suppliers",s.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>
      </div>))}
    </>}
    {view==="products"&&<>
      <div style={{margin:"12px 0 8px"}}><Btn color={C.navy} onClick={()=>onAdd({type:"product"})}>+ Add Product</Btn></div>
      {data.products.length===0&&<Empty text="No products yet."/>}
      {data.products.map(p=>(<div key={p.id} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{p.name}</B><span style={{fontSize:11,color:C.muted}}>{p.unit}</span></Row>
        {p.supplierName&&<Mute>{p.supplierName}</Mute>}
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>onEdit("product",p)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Edit</button>
          <button onClick={()=>onDel("products",p.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>
      </div>))}
    </>}
  </div>);
}

// ─── REPORTS TAB ─────────────────────────────────────────────────
function ReportsTab({data,tradingOut,tots,generatePDF}){
  const[rep,setRep]=useState("sales");
  const[fy,setFY]=useState(ALL_FY);
  const todayMs=new Date().setHours(0,0,0,0);
  const allFYs=useMemo(()=>getAvailableFYs(data.tradingSales),[data]);

  const filteredSales=useMemo(()=>filterByFY(data.tradingSales,fy),[data,fy]);
  const filteredPay=useMemo(()=>{
    if(fy===ALL_FY)return data.tradingPayments;
    const{start,end}=getFYRange(fy);
    return data.tradingPayments.filter(p=>{const d=new Date(p.date);return d>=start&&d<=end;});
  },[data,fy]);

  const FYPicker=()=>(
    <div style={{marginBottom:12}}>
      <div style={{fontSize:10.5,color:C.muted,fontWeight:600,marginBottom:5,textTransform:"uppercase"}}>Filter by FY</div>
      <div style={{display:"flex",overflowX:"auto",gap:5,scrollbarWidth:"none"}}>
        {allFYs.map(f=>(
          <button key={f} onClick={()=>setFY(f)} style={{flex:"0 0 auto",padding:"7px 12px",borderRadius:16,fontSize:11.5,fontWeight:fy===f?700:500,border:`1.5px solid ${fy===f?C.navy:C.border}`,background:fy===f?C.navy:"#fff",color:fy===f?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>
            {f===ALL_FY?"All":f}
          </button>
        ))}
      </div>
    </div>
  );

  const doPDFSales=()=>{
    const rows=[...filteredSales].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(s=>`<tr><td>${fmtD(s.date)}</td><td>${s.billNo||"—"}</td><td>${s.customerName}</td><td>${s.productName||"—"}</td><td style="text-align:right">${fmt(s.meters)} m</td><td style="text-align:right">₹${fmt(s.rate)}</td><td style="text-align:right;font-weight:bold">₹${fmt(s.amount)}</td></tr>`).join("");
    const total=filteredSales.reduce((a,s)=>a+(+s.amount||0),0);
    generatePDF(`Sales Register${fy!==ALL_FY?" — FY "+fy:""}`,`<table><thead><tr><th>Date</th><th>Bill No</th><th>Customer</th><th>Product</th><th>Qty (m)</th><th>Rate</th><th>Amount</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="6">TOTAL</td><td>₹${fmt(total)}</td></tr></tbody></table>`);
  };

  // Outstanding is always current — no FY filter (it's live balance)
  const outEntries=useMemo(()=>Object.values(tradingOut).map(v=>{
    const net=Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0));
    const custSales=data.tradingSales.filter(s=>s.customerName===v.name).sort((a,b)=>new Date(a.date)-new Date(b.date));
    let rem=v.paid;const openBills=[];
    custSales.forEach(s=>{const amt=+s.amount||0;const d=Math.min(amt,rem);rem-=d;const left=amt-d;
      if(left>0){const days=Math.floor((todayMs-new Date(s.date).setHours(0,0,0,0))/86400000);
        openBills.push({billNo:s.billNo||"—",date:s.date,billAmount:amt,outstanding:left,days,productName:s.productName,meters:+s.meters||0,rate:+s.rate||0});}
    });
    return{name:v.name,net,openBills};
  }).filter(v=>v.net>0).sort((a,b)=>b.net-a.net),[data,tradingOut]);

  const doPDFOut=()=>{
    const partyBlocks=outEntries.map(v=>{
      const credit=v.openBills.map(b=>Math.max(0,b.billAmount-b.outstanding));
      const rows=v.openBills.map((b,i)=>`<tr><td>${fmtD(b.date)}</td><td>${b.billNo}</td><td style="text-align:right">₹${fmt(b.billAmount)}</td><td style="text-align:right">₹${fmt(credit[i])}</td><td style="text-align:right;color:#E74C3C;font-weight:bold">₹${fmt(b.outstanding)}</td><td style="text-align:right">${b.days}</td></tr>`).join("");
      return `<div class="party"><div class="party-name">Party : ${v.name}</div>
        <table><thead><tr><th>Bill Date</th><th>Bill No</th><th style="text-align:right">Amount</th><th style="text-align:right">Credit</th><th style="text-align:right">Outstanding</th><th style="text-align:right">Days</th></tr></thead>
        <tbody>${rows}<tr class="party-tot"><td colspan="4">Party Total</td><td style="text-align:right">₹${fmt(v.net)}</td><td></td></tr></tbody></table></div>`;
    }).join("");
    const style=`<style>.party{margin-bottom:22px;page-break-inside:avoid;}.party-name{font-weight:800;font-size:14px;color:#0F1923;margin-bottom:6px;padding-bottom:4px;border-bottom:1.5px solid #0F1923;}.party-tot td{background:#F0F4F8;font-weight:bold;border-top:1.5px solid #0F1923;}table th,table td{border:1px solid #ddd;}</style>`;
    generatePDF("Outstanding Statement",`${style}${partyBlocks}<div class="tot" style="display:flex;justify-content:space-between;background:#0F1923;color:#E8C97E;padding:12px 16px;border-radius:6px;font-weight:bold;font-size:15px;margin-top:10px;"><span>TOTAL OUTSTANDING</span><span>₹${fmt(tots.totTradingOut)}</span></div>`);
  };

  const{totalCommission,commByCustomer}=useMemo(()=>calcFIFOCommission(filteredSales,filteredPay),[filteredSales,filteredPay]);
  const commList=Object.values(commByCustomer).sort((a,b)=>b.commission-a.commission);
  const doPDFComm=()=>{
    const rows=commList.map((v,i)=>`<tr><td>${i+1}</td><td>${v.name}</td><td style="text-align:right">${v.bills}</td><td style="text-align:right">${fmt(v.meters)} m</td><td style="text-align:right;color:#8E44AD;font-weight:bold">₹${fmt(v.commission)}</td></tr>`).join("");
    generatePDF(`Commission Statement${fy!==ALL_FY?" — FY "+fy:""}`,`<table><thead><tr><th>#</th><th>Customer</th><th>Bills</th><th>Meters</th><th>Commission</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="4">TOTAL</td><td>₹${fmt(totalCommission)}</td></tr></tbody></table>`);
  };

  const totSales=filteredSales.reduce((a,s)=>a+(+s.amount||0),0);

  return(<div>
    <div style={{display:"flex",overflowX:"auto",gap:8,marginBottom:12,scrollbarWidth:"none"}}>
      {[{v:"sales",l:"Sales Register"},{v:"outstanding",l:"Outstanding"},{v:"commission",l:"Commission"}].map(r=>(<button key={r.v} onClick={()=>setRep(r.v)} style={{flex:"0 0 auto",padding:"10px 16px",borderRadius:20,fontSize:13,fontWeight:rep===r.v?700:500,border:`1.5px solid ${rep===r.v?C.navy:C.border}`,background:rep===r.v?C.navy:"#fff",color:rep===r.v?C.gold:"#666",cursor:"pointer",minHeight:42}}>{r.l}</button>))}
    </div>

    {rep==="sales"&&<>
      <FYPicker/>
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        <button onClick={()=>exportCSV([["Date","Bill No","Customer","Product","Qty(m)","Rate","Amount"],...filteredSales.map(s=>[fmtD(s.date),s.billNo||"",s.customerName,s.productName,s.meters,s.rate,s.amount])],`SalesRegister_${fy.replace(/[^0-9\-]/g,"")}_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>📥 CSV</button>
        <button onClick={doPDFSales} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>🖨️ PDF</button>
      </div>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12,borderLeft:`4px solid ${C.blue}`}}>
        <Mute>Total Sales {fy!==ALL_FY?`— FY ${fy}`:""} ({filteredSales.length} bills)</Mute>
        <div style={{fontSize:22,fontWeight:900,color:C.blue}}>₹{fmt(totSales)}</div>
      </div>
      {[...filteredSales].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(s=>(
        <div key={s.id} style={{background:C.card,borderRadius:11,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 6px rgba(0,0,0,0.05)"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1,paddingRight:10}}>
              <B style={{fontSize:13.5,color:C.navy}}>{s.customerName}</B>
              {s.billNo&&<Mute>📋 Bill No : {s.billNo}</Mute>}
              <Mute>📅 {fmtD(s.date)}</Mute>
              {s.productName&&<Mute>📦 {s.productName}</Mute>}
              <Mute>📏 {fmt(s.meters)} m @ ₹{fmt(s.rate)}/m</Mute>
            </div>
            <span style={{fontWeight:900,color:C.blue,fontSize:15}}>₹{fmt(s.amount)}</span>
          </div>
        </div>
      ))}
    </>}

    {rep==="outstanding"&&<>
      <div style={{background:"#FFF8E7",borderRadius:10,padding:"10px 13px",marginBottom:12,fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
        📌 Outstanding is always current (live balance). FY filter does not apply here.
      </div>
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        <button onClick={()=>exportCSV([["Customer","Bill Date","Bill No","Product","Qty","Bill Amt","Outstanding","Days"],...outEntries.flatMap(v=>v.openBills.map(b=>[v.name,fmtD(b.date),b.billNo,b.productName,b.meters,b.billAmount,b.outstanding,b.days]))],`Outstanding_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>📥 CSV</button>
        <button onClick={doPDFOut} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>🖨️ PDF</button>
      </div>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12,borderLeft:`4px solid ${C.red}`}}>
        <Mute>Total Outstanding ({outEntries.length} customers)</Mute>
        <div style={{fontSize:22,fontWeight:900,color:C.red}}>₹{fmt(tots.totTradingOut)}</div>
      </div>
      {outEntries.map(v=>(
        <div key={v.name} style={{background:C.card,borderRadius:13,padding:"0",marginBottom:12,boxShadow:"0 1px 8px rgba(0,0,0,0.07)",overflow:"hidden"}}>
          <div style={{background:`linear-gradient(90deg,${C.navyMid},${C.navy})`,padding:"11px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <B style={{color:"#fff",fontSize:13}}>{v.name}</B>
            <span style={{fontWeight:900,color:C.gold,fontSize:14}}>₹{fmt(v.net)}</span>
          </div>
          {v.openBills.map((b,i)=>{
            const col=b.days<=30?C.green:b.days<=60?C.orange:b.days<=90?"#E67E22":C.red;
            return(
              <div key={i} style={{padding:"10px 14px",borderBottom:`1px solid ${C.border}`,background:i%2===0?"#fff":"#FAFBFC"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                      <span style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{b.billNo}</span>
                      <span style={{fontSize:11,color:C.muted}}>📅 {fmtD(b.date)}</span>
                      <span style={{fontSize:11,color:col,fontWeight:700}}>⏱ {b.days}d</span>
                    </div>
                    {b.productName&&<div style={{fontSize:11.5,color:C.teal,marginTop:3}}>📦 {b.productName}</div>}
                    {b.meters>0&&<div style={{fontSize:11.5,color:C.muted}}>📏 {fmt(b.meters)} m {b.rate>0?`@ ₹${fmt(b.rate)}/m`:""}</div>}
                    <div style={{fontSize:11.5,color:C.muted}}>Bill Amt: ₹{fmt(b.billAmount)}</div>
                  </div>
                  <div style={{textAlign:"right",minWidth:90}}>
                    <div style={{fontWeight:900,fontSize:14,color:C.red}}>₹{fmt(b.outstanding)}</div>
                    <div style={{fontSize:10,color:col,fontWeight:600}}>O/S</div>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      ))}
    </>}

    {rep==="commission"&&<>
      <FYPicker/>
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        <button onClick={()=>exportCSV([["#","Customer","Bills","Meters","Commission"],...commList.map((v,i)=>[i+1,v.name,v.bills,v.meters,v.commission])],`Commission_${fy.replace(/[^0-9\-]/g,"")}_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>📥 CSV</button>
        <button onClick={doPDFComm} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>🖨️ PDF</button>
      </div>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12,borderLeft:`4px solid ${C.purple}`}}>
        <Mute>Commission {fy!==ALL_FY?`— FY ${fy}`:""} (FIFO)</Mute>
        <div style={{fontSize:22,fontWeight:900,color:C.purple}}>₹{fmt(totalCommission)}</div>
      </div>
      {commList.length===0&&<Empty text="No commission for this period."/>}
      {commList.map((v,i)=>(
        <div key={v.name} style={{background:C.card,borderRadius:11,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",borderLeft:`3px solid ${C.purple}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <B style={{fontSize:13.5}}>{v.name}</B>
              <Mute>📄 {v.bills} bills fully adjusted</Mute>
              <Mute>📏 {fmt(v.meters)} m @ ₹1.50</Mute>
            </div>
            <span style={{fontWeight:900,color:C.purple,fontSize:15}}>₹{fmt(v.commission)}</span>
          </div>
        </div>
      ))}
      {commList.length>0&&<div style={{background:C.navy,borderRadius:12,padding:"14px 16px",marginTop:8,display:"flex",justifyContent:"space-between"}}>
        <span style={{fontWeight:700,color:C.gold}}>TOTAL</span>
        <span style={{fontWeight:900,fontSize:16,color:C.gold}}>₹{fmt(totalCommission)}</span>
      </div>}
    </>}
  </div>);
}

// ─── MODALS ──────────────────────────────────────────────────────
function SaleModal({data,onSave,onClose}){
  const[f,sf]=useState({date:today(),billNo:"",customerName:"",productName:"",supplierName:"",meters:"",rate:"",amount:"",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const selectProd=(name)=>{const p=data.products.find(p=>p.name===name);sf(pr=>({...pr,productName:name,supplierName:p?.supplierName||pr.supplierName}));};
  useEffect(()=>{if(f.meters&&f.rate)s("amount",(parseFloat(f.meters)*parseFloat(f.rate)).toFixed(2));},[f.meters,f.rate]);
  return(<ModalBase title="Add Trading Sale" onClose={onClose}>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Bill No"><input value={f.billNo} onChange={e=>s("billNo",e.target.value)} placeholder="Invoice / Bill number" style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>s("customerName",v)} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="sl-c"/></F>
    <F label="Product *"><SmartInput value={f.productName} onChange={selectProd} placeholder="Product" list={data.products.map(p=>p.name)} idPrefix="sl-p"/></F>
    <F label="Supplier"><SmartInput value={f.supplierName} onChange={v=>s("supplierName",v)} placeholder="Supplier" list={data.suppliers.map(s=>s.name)} idPrefix="sl-s"/></F>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <F label="Meters"><input type="number" value={f.meters} onChange={e=>s("meters",e.target.value)} placeholder="0" style={IS}/></F>
      <F label="Rate (₹/m)"><input type="number" value={f.rate} onChange={e=>s("rate",e.target.value)} placeholder="0" style={IS}/></F>
    </div>
    <F label="Amount (₹)"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} style={{...IS,fontWeight:700}}/></F>
    <F label="Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Optional" style={IS}/></F>
    <SaveBtn color={C.blue} onClick={()=>{if(!f.customerName||!f.amount)return alert("Customer and Amount required");onSave(f);}}>Save Sale</SaveBtn>
  </ModalBase>);
}

function PaymentModal({data,onSave,onClose,preCustomer}){
  const[f,sf]=useState({date:today(),billNo:"",customerName:preCustomer||"",amount:"",mode:"NEFT",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const custBills=data.tradingSales.filter(x=>x.customerName===f.customerName&&x.billNo).map(x=>x.billNo);
  const uniqueBills=[...new Set(custBills)];
  const selectBill=(bn)=>{const sale=data.tradingSales.find(x=>x.billNo===bn&&x.customerName===f.customerName);sf(p=>({...p,billNo:bn,amount:sale?sale.amount:p.amount}));};
  const linkedSale=f.billNo?data.tradingSales.find(x=>x.billNo===f.billNo&&x.customerName===f.customerName):null;
  const commissionEarned=linkedSale?(+linkedSale.meters||0)*1.5:0;
  return(<ModalBase title="Record Payment" onClose={onClose}>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>sf(p=>({...p,customerName:v,billNo:""}))} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="pm-c"/></F>
    <F label="Bill No"><SmartInput value={f.billNo} onChange={selectBill} placeholder="Select bill number" list={uniqueBills} idPrefix="pm-b"/></F>
    <F label="Amount (₹) *"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} placeholder="0" style={{...IS,fontWeight:700}}/></F>
    <F label="Mode of Payment"><select value={f.mode} onChange={e=>s("mode",e.target.value)} style={IS}><option>NEFT</option><option>RTGS</option><option>Cheque</option><option>Cash</option><option>UPI</option></select></F>
    {commissionEarned>0&&<div style={{background:"#F3EEF9",borderRadius:10,padding:"11px 14px",marginBottom:12,fontSize:13}}><div style={{fontWeight:700,color:C.purple}}>Commission on this payment</div><div style={{fontSize:18,fontWeight:900,color:C.purple,marginTop:4}}>₹{fmt(commissionEarned)}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{fmt(linkedSale?.meters||0)} meters × ₹1.5</div></div>}
    <F label="Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Cheque no, ref…" style={IS}/></F>
    <SaveBtn color={C.green} onClick={()=>{if(!f.customerName||!f.amount)return alert("Customer and Amount required");onSave({...f,commissionEarned});}}>Save Payment</SaveBtn>
  </ModalBase>);
}

function DebitModal({data,onSave,onClose}){
  const[f,sf]=useState({date:today(),noteNo:"",customerName:"",originalBillNo:"",amount:"",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const custBills=data.tradingSales.filter(x=>x.customerName===f.customerName&&x.billNo).map(x=>x.billNo);
  return(<ModalBase title="Add Debit Note" onClose={onClose}>
    <div style={{background:"#FDEDEC",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12.5,color:"#922B21"}}>Debit Note increases the customer outstanding balance.</div>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Debit Note No"><input value={f.noteNo} onChange={e=>s("noteNo",e.target.value)} placeholder="DN-001" style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>s("customerName",v)} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="dn-c"/></F>
    <F label="Against Bill No"><SmartInput value={f.originalBillNo} onChange={v=>s("originalBillNo",v)} placeholder="Original bill number" list={[...new Set(custBills)]} idPrefix="dn-b"/></F>
    <F label="Amount (₹) *"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} placeholder="0" style={{...IS,fontWeight:700}}/></F>
    <F label="Reason / Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Price difference, penalty, etc." style={IS}/></F>
    <SaveBtn color={C.red} onClick={()=>{if(!f.customerName||!f.amount)return alert("Customer and Amount required");onSave(f);}}>Save Debit Note</SaveBtn>
  </ModalBase>);
}

function CreditModal({data,onSave,onClose}){
  const[f,sf]=useState({date:today(),noteNo:"",customerName:"",originalBillNo:"",productName:"",meters:"",rate:"",amount:"",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  useEffect(()=>{if(f.meters&&f.rate)s("amount",(parseFloat(f.meters)*parseFloat(f.rate)).toFixed(2));},[f.meters,f.rate]);
  const custBills=data.tradingSales.filter(x=>x.customerName===f.customerName&&x.billNo).map(x=>x.billNo);
  return(<ModalBase title="Credit Note / Sale Return" onClose={onClose}>
    <div style={{background:"#E8F8F5",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12.5,color:C.teal}}>Credit Note reduces the customer outstanding balance.</div>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Credit Note No"><input value={f.noteNo} onChange={e=>s("noteNo",e.target.value)} placeholder="CN-001" style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>s("customerName",v)} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="cn-c"/></F>
    <F label="Against Bill No"><SmartInput value={f.originalBillNo} onChange={v=>s("originalBillNo",v)} placeholder="Original bill number" list={[...new Set(custBills)]} idPrefix="cn-b"/></F>
    <F label="Product Returned"><SmartInput value={f.productName} onChange={v=>s("productName",v)} placeholder="Product name" list={data.products.map(p=>p.name)} idPrefix="cn-p"/></F>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <F label="Meters Returned"><input type="number" value={f.meters} onChange={e=>s("meters",e.target.value)} placeholder="0" style={IS}/></F>
      <F label="Rate (₹/m)"><input type="number" value={f.rate} onChange={e=>s("rate",e.target.value)} placeholder="0" style={IS}/></F>
    </div>
    <F label="Amount (₹) *"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} placeholder="Auto or manual" style={{...IS,fontWeight:700}}/></F>
    <F label="Reason / Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Quality issue, excess, etc." style={IS}/></F>
    <SaveBtn color={C.green} onClick={()=>{if(!f.customerName||!f.amount)return alert("Customer and Amount required");onSave(f);}}>Save Credit Note / Return</SaveBtn>
  </ModalBase>);
}

function CustomerModal({onSave,onClose,initial}){
  const[f,sf]=useState(initial||{name:"",type:"Trading",phone:"",city:"",gstin:"",creditDays:"30"});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  return(<ModalBase title={initial?"Edit Customer":"Add Customer"} onClose={onClose}>
    <F label="Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Customer name" style={IS}/></F>
    <F label="Type"><select value={f.type} onChange={e=>s("type",e.target.value)} style={IS}><option>Trading</option><option>Agency</option><option>Both</option></select></F>
    <F label="Phone"><input type="tel" value={f.phone} onChange={e=>s("phone",e.target.value)} placeholder="10-digit mobile" style={IS}/></F>
    <F label="City"><input value={f.city} onChange={e=>s("city",e.target.value)} placeholder="City" style={IS}/></F>
    <F label="GSTIN"><input value={f.gstin} onChange={e=>s("gstin",e.target.value)} placeholder="GST number" style={IS}/></F>
    <F label="Credit Days (for ageing follow-up)">
      <select value={f.creditDays||"30"} onChange={e=>s("creditDays",e.target.value)} style={IS}>
        <option value="0">0 Days (Immediate)</option>
        <option value="15">15 Days</option>
        <option value="30">30 Days</option>
        <option value="45">45 Days</option>
        <option value="60">60 Days</option>
        <option value="90">90 Days</option>
      </select>
    </F>
    <SaveBtn color={C.navy} onClick={()=>{if(!f.name)return alert("Name required");onSave(f);}}>Save Customer</SaveBtn>
  </ModalBase>);
}

function SupplierModal({onSave,onClose,initial}){
  const[f,sf]=useState(initial||{name:"",phone:"",city:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  return(<ModalBase title={initial?"Edit Supplier":"Add Supplier"} onClose={onClose}>
    <F label="Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Supplier name" style={IS}/></F>
    <F label="Phone"><input type="tel" value={f.phone} onChange={e=>s("phone",e.target.value)} placeholder="Phone" style={IS}/></F>
    <F label="City"><input value={f.city} onChange={e=>s("city",e.target.value)} placeholder="City" style={IS}/></F>
    <SaveBtn color={C.navy} onClick={()=>{if(!f.name)return alert("Name required");onSave(f);}}>Save Supplier</SaveBtn>
  </ModalBase>);
}

function ProductModal({data,onSave,onClose,initial}){
  const[f,sf]=useState(initial||{name:"",supplierName:"",unit:"Mtr"});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  return(<ModalBase title={initial?"Edit Product":"Add Product"} onClose={onClose}>
    <F label="Product Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Product name" style={IS}/></F>
    <F label="Supplier"><SmartInput value={f.supplierName} onChange={v=>s("supplierName",v)} placeholder="Supplier" list={(data?.suppliers||[]).map(s=>s.name)} idPrefix="prod-s"/></F>
    <F label="Unit"><select value={f.unit} onChange={e=>s("unit",e.target.value)} style={IS}><option>Mtr</option><option>Kg</option><option>Pcs</option></select></F>
    <SaveBtn color={C.navy} onClick={()=>{if(!f.name)return alert("Name required");onSave(f);}}>Save Product</SaveBtn>
  </ModalBase>);
}