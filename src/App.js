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
  win.document.write(`<!DOCTYPE html><html><head><title>${title}</title><style>body{font-family:Arial,sans-serif;font-size:13px;color:#111;padding:30px;max-width:800px;margin:auto;}h1{font-size:20px;border-bottom:2px solid #111;padding-bottom:8px;}table{width:100%;border-collapse:collapse;margin-top:16px;}th{background:#0F1923;color:#E8C97E;padding:8px 10px;text-align:left;font-size:12px;}td{padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;}tr:nth-child(even){background:#f9f9f9;}.tot td{background:#0F1923;color:#E8C97E;font-weight:bold;border:none;}.hdr{display:flex;justify-content:space-between;margin-bottom:20px;}.co{font-size:22px;font-weight:bold;color:#0F1923;}</style></head><body><div class="hdr"><div><div class="co">Navkar Fabrics</div><div style="font-size:12px;color:#666">Trading Sales Report</div></div><div style="text-align:right;font-size:12px;color:#666">Date: ${new Date().toLocaleDateString("en-IN")}</div></div><h1>${title}</h1>${content}<script>window.onload=()=>{window.print();}<\/script></body></html>`);
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
  const totComm=data.tradingPayments.reduce((a,p)=>a+(+p.commissionEarned||0),0);
  const overdueCount=Object.values(tradingOut).filter(v=>Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))>0).length;

  const Sidebar=()=>(
    <div style={{width:220,background:C.navy,minHeight:"100vh",position:"fixed",left:0,top:0,zIndex:150,display:"flex",flexDirection:"column",padding:"20px 0"}}>
      <div style={{padding:"0 20px 24px"}}><div style={{fontSize:10,letterSpacing:2.5,color:C.gold,textTransform:"uppercase",fontWeight:700}}>Navkar Fabrics</div><div style={{fontSize:20,fontWeight:900,color:"#fff",marginTop:4}}>Sales Manager</div></div>
      {TABS.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?"rgba(232,201,126,0.12)":"transparent",border:"none",borderLeft:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",padding:"12px 20px",textAlign:"left",color:tab===t.id?C.gold:"rgba(255,255,255,0.65)",fontSize:13.5,fontWeight:tab===t.id?700:400,cursor:"pointer",display:"flex",alignItems:"center",gap:10}}><span>{t.icon}</span>{t.label}{t.id==="Outstanding"&&overdueCount>0&&<span style={{background:C.red,color:"#fff",borderRadius:20,fontSize:10,fontWeight:800,padding:"2px 7px",marginLeft:"auto"}}>{overdueCount}</span>}</button>))}
      <div style={{marginTop:"auto",padding:"16px 20px",borderTop:"1px solid rgba(255,255,255,0.08)",display:"flex",flexDirection:"column",gap:8}}>
        <button onClick={()=>importRef.current&&importRef.current.click()} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📥 <span style={{fontSize:12}}>Import Excel</span></button>
        <button onClick={()=>exportBackup(data)} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📤 <span style={{fontSize:12}}>Backup</span></button>
        <button onClick={()=>restoreRef.current&&restoreRef.current.click()} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>🔄 <span style={{fontSize:12}}>Restore</span></button>
      </div>
    </div>
  );

  return(
    <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:C.bg,minHeight:"100vh"}}>
      <style>{`@media(min-width:768px){.mob{display:none!important}.desk{display:flex!important}.main{margin-left:220px!important}}@media(max-width:767px){.mob{display:flex!important}.desk{display:none!important}.main{margin-left:0!important;padding-bottom:80px!important}}.main{max-width:860px;padding:18px 16px 40px;}`}</style>
      <div className="desk" style={{display:"none"}}><Sidebar/></div>
      <div className="mob" style={{display:"none",background:C.navy,padding:"calc(env(safe-area-inset-top,0px)+12px) 16px 10px",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 20px rgba(0,0,0,0.4)",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <div><div style={{fontSize:9,letterSpacing:2,color:C.gold,textTransform:"uppercase",fontWeight:700}}>Amihem Sales</div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>Sales Manager</div></div>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>importRef.current&&importRef.current.click()} style={hdrBtn()}>📥<span style={{fontSize:9}}>Excel</span></button>
          <button onClick={()=>exportBackup(data)} style={hdrBtn()}>🔄<span style={{fontSize:9}}>Backup</span></button>
          <button onClick={()=>restoreRef.current&&restoreRef.current.click()} style={hdrBtn()}>📤<span style={{fontSize:9}}>Restore</span></button>
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
        {tab==="Commission" &&<CommissionTab data={data} totComm={totComm} generatePDF={generatePDF}/>}
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
  const totComm=data.tradingPayments.reduce((a,p)=>a+(+p.commissionEarned||0),0);
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
  const todayMs=new Date().setHours(0,0,0,0);
  const phoneMap={};data.customers.forEach(c=>{phoneMap[c.name]=c.phone;});
  const entries=Object.values(tradingOut).map(v=>{
    const net=Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0));
    const custSales=data.tradingSales.filter(s=>s.customerName===v.name).sort((a,b)=>new Date(a.date)-new Date(b.date));
    let rem=v.paid;let maxDays=0;let oldestBill=null;let oldestDate=null;
    custSales.forEach(s=>{const amt=+s.amount||0;const d=Math.min(amt,rem);rem-=d;const left=amt-d;if(left>0){const days=Math.floor((todayMs-new Date(s.date).setHours(0,0,0,0))/86400000);if(days>maxDays){maxDays=days;oldestBill=s.billNo;oldestDate=s.date;}}});
    return{...v,net,maxDays,oldestBill,oldestDate};
  }).filter(v=>v.net>0).filter(v=>!search||v.name.toLowerCase().includes(search.toLowerCase())).sort((a,b)=>b.net-a.net);

  const total=entries.reduce((a,v)=>a+v.net,0);
  const ageBucket=(days)=>{if(days<=30)return{label:"0-30 Days",color:C.green};if(days<=60)return{label:"31-60 Days",color:C.orange};if(days<=90)return{label:"61-90 Days",color:"#E67E22"};if(days<=120)return{label:"91-120 Days",color:C.red};return{label:"Above 120 Days",color:"#922B21"};};

  const buildWA=(v)=>{const b=ageBucket(v.maxDays);return `Navkar Fabrics\nDate: ${new Date().toLocaleDateString("en-IN")}\n\nDear ${v.name},\n\nThis is a gentle reminder regarding your outstanding payment.\n\nBill No      : ${v.oldestBill||"N/A"}\nBill Date    : ${fmtD(v.oldestDate)}\nOutstanding  : Rs. ${fmt(v.net)}\nDays Overdue : ${v.maxDays} Days (${b.label})\n\nKindly arrange the payment at your earliest convenience.\n\nRegards,\nNavkar Fabrics`;};

  const doPDF=()=>{
    const rows=entries.map(v=>`<tr><td>${v.name}</td><td>${v.oldestBill||"—"}</td><td>${fmtD(v.oldestDate)}</td><td style="text-align:right">₹${fmt(v.due)}</td><td style="text-align:right">₹${fmt(v.paid)}</td><td style="text-align:right;color:#E74C3C;font-weight:bold">₹${fmt(v.net)}</td><td>${ageBucket(v.maxDays).label}</td></tr>`).join("");
    generatePDF("Outstanding Statement",`<table><thead><tr><th>Customer</th><th>Bill No</th><th>Bill Date</th><th>Total Sales</th><th>Paid</th><th>Outstanding</th><th>Ageing</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="5">TOTAL</td><td>₹${fmt(total)}</td><td></td></tr></tbody></table>`);
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
    {entries.map(v=>{const bucket=ageBucket(v.maxDays);const phone=phoneMap[v.name]||v.phone||"";return(
      <div key={v.name} style={{background:C.card,borderRadius:14,padding:"14px 15px",marginBottom:12,boxShadow:"0 1px 8px rgba(0,0,0,0.07)",borderLeft:`4px solid ${bucket.color}`}}>
        <Row><div><B style={{fontSize:15}}>{v.name}</B>{phone&&<Mute>{phone}</Mute>}</div><div style={{textAlign:"right"}}><div style={{fontWeight:900,fontSize:18,color:C.red}}>₹{fmt(v.net)}</div><span style={{fontSize:10,color:bucket.color,fontWeight:700}}>{bucket.label}</span></div></Row>
        <div style={{display:"flex",gap:14,marginTop:8,fontSize:12,color:C.muted,flexWrap:"wrap"}}>
          <span>Sales: <b style={{color:C.blue}}>₹{fmt(v.due)}</b></span>
          <span>Paid: <b style={{color:C.green}}>₹{fmt(v.paid)}</b></span>
          {v.debit>0&&<span>Debit: <b style={{color:C.red}}>₹{fmt(v.debit)}</b></span>}
          {v.credit>0&&<span>Credit: <b style={{color:C.teal}}>₹{fmt(v.credit)}</b></span>}
          {v.oldestBill&&<span>Bill: <b>{v.oldestBill}</b></span>}
          {v.maxDays>0&&<span>Overdue: <b style={{color:bucket.color}}>{v.maxDays}d</b></span>}
        </div>
        <div style={{display:"flex",gap:8,marginTop:12,flexWrap:"wrap"}}>
          <button onClick={()=>onAddPay(v.name)} style={{background:C.blue,color:"#fff",border:"none",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flex:1,minHeight:40}}>Record Payment</button>
          <button onClick={()=>{const n=phone?"91"+String(phone).replace(/\D/g,"").slice(-10):"";window.open(`https://wa.me/${n}?text=${encodeURIComponent(buildWA(v))}`,"_blank");}} style={{background:"#E8FBF0",color:"#25D366",border:"1px solid #25D366",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flex:1,minHeight:40}}>WhatsApp</button>
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
  const[chartType,setChartType]=useState("bar");
  const monthly=useMemo(()=>{const map={};data.tradingSales.forEach(s=>{const d=new Date(s.date);const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;const label=d.toLocaleDateString("en-IN",{month:"short",year:"2-digit"});if(!map[key])map[key]={key,label,amount:0};map[key].amount+=+s.amount||0;});return Object.values(map).sort((a,b)=>a.key.localeCompare(b.key)).slice(-12);},[data]);
  const products=useMemo(()=>{const map={};data.tradingSales.forEach(s=>{const p=s.productName||"Unknown";if(!map[p])map[p]={name:p,value:0};map[p].value+=+s.amount||0;});return Object.values(map).sort((a,b)=>b.value-a.value).slice(0,7);},[data]);
  const customers=Object.values(tradingOut).map(v=>({name:v.name.length>14?v.name.slice(0,14)+"…":v.name,sales:v.due,outstanding:Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))})).sort((a,b)=>b.sales-a.sales).slice(0,8);
  const PIE=[C.blue,C.green,C.orange,C.purple,C.teal,C.red,"#F39C12","#1ABC9C"];
  return(<div>
    <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
      {[{v:"bar",l:"Monthly"},{v:"line",l:"Trend"},{v:"pie",l:"Product"},{v:"customer",l:"Customer"}].map(ct=>(<button key={ct.v} onClick={()=>setChartType(ct.v)} style={{flex:"0 0 auto",padding:"9px 14px",borderRadius:20,fontSize:12.5,fontWeight:chartType===ct.v?700:500,border:`1.5px solid ${chartType===ct.v?C.navy:C.border}`,background:chartType===ct.v?C.navy:"#fff",color:chartType===ct.v?C.gold:"#666",cursor:"pointer",minHeight:40}}>{ct.l}</button>))}
    </div>
    <div style={{background:C.card,borderRadius:14,padding:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      {chartType==="bar"&&<><div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:16}}>Monthly Sales</div>{monthly.length===0?<Empty text="No data."/>:<ResponsiveContainer width="100%" height={240}><BarChart data={monthly} margin={{top:4,right:8,left:0,bottom:4}}><CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/><XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/><YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(1)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/><Tooltip formatter={v=>`₹${fmt(v)}`}/><Bar dataKey="amount" fill={C.blue} radius={[5,5,0,0]} name="Sales"/></BarChart></ResponsiveContainer>}</>}
      {chartType==="line"&&<><div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:16}}>Sales Trend</div>{monthly.length===0?<Empty text="No data."/>:<ResponsiveContainer width="100%" height={240}><LineChart data={monthly} margin={{top:4,right:8,left:0,bottom:4}}><CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/><XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/><YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(1)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/><Tooltip formatter={v=>`₹${fmt(v)}`}/><Line type="monotone" dataKey="amount" stroke={C.blue} strokeWidth={2.5} dot={{r:4,fill:C.blue}} name="Sales"/></LineChart></ResponsiveContainer>}</>}
      {chartType==="pie"&&<><div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:16}}>Sales by Product</div>{products.length===0?<Empty text="No data."/>:<ResponsiveContainer width="100%" height={240}><PieChart><Pie data={products} cx="50%" cy="50%" outerRadius={90} dataKey="value" nameKey="name" label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={false} fontSize={10}>{products.map((_,i)=><Cell key={i} fill={PIE[i%PIE.length]}/>)}</Pie><Tooltip formatter={v=>`₹${fmt(v)}`}/></PieChart></ResponsiveContainer>}</>}
      {chartType==="customer"&&<><div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:16}}>Top Customers</div>{customers.length===0?<Empty text="No data."/>:<ResponsiveContainer width="100%" height={280}><BarChart data={customers} layout="vertical" margin={{top:0,right:8,left:60,bottom:0}}><CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8" horizontal={false}/><XAxis type="number" tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(1)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:9}} tickLine={false}/><YAxis type="category" dataKey="name" tick={{fontSize:10}} width={60} tickLine={false} axisLine={false}/><Tooltip formatter={v=>`₹${fmt(v)}`}/><Legend iconSize={10} wrapperStyle={{fontSize:11}}/><Bar dataKey="sales" fill={C.blue} radius={[0,4,4,0]} name="Sales" barSize={10}/><Bar dataKey="outstanding" fill={C.red} radius={[0,4,4,0]} name="Outstanding" barSize={10}/></BarChart></ResponsiveContainer>}</>}
    </div>
  </div>);
}

// ─── COMMISSION TAB ──────────────────────────────────────────────
function CommissionTab({data,totComm,generatePDF}){
  const commByCustomer={};
  data.tradingPayments.forEach(p=>{if(!commByCustomer[p.customerName])commByCustomer[p.customerName]={name:p.customerName,commission:0,payments:0};commByCustomer[p.customerName].commission+=+p.commissionEarned||0;commByCustomer[p.customerName].payments+=+p.amount||0;});
  const commList=Object.values(commByCustomer).filter(v=>v.commission>0).sort((a,b)=>b.commission-a.commission);
  const doPDF=()=>{const rows=commList.map(v=>`<tr><td>${v.name}</td><td style="text-align:right">₹${fmt(v.payments)}</td><td style="text-align:right;color:#8E44AD;font-weight:bold">₹${fmt(v.commission)}</td></tr>`).join("");generatePDF("Commission Statement",`<table><thead><tr><th>Customer</th><th>Payment Received</th><th>Commission Earned</th></tr></thead><tbody>${rows}<tr class="tot"><td>TOTAL</td><td></td><td>₹${fmt(totComm)}</td></tr></tbody></table>`);};
  return(<div>
    <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:14,padding:"14px 16px",marginBottom:16,border:"1px solid rgba(232,201,126,0.25)"}}>
      <div style={{fontSize:10.5,color:C.gold,textTransform:"uppercase",letterSpacing:1,fontWeight:700}}>Commission Earned on Payments</div>
      <div style={{fontSize:28,fontWeight:900,color:C.gold}}>₹{fmt(totComm)}</div>
      <div style={{fontSize:11,color:"rgba(255,255,255,0.4)",marginTop:2}}>Calculated at time of payment receipt</div>
    </div>
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      <button onClick={()=>exportCSV([["Customer","Payment Received","Commission"],...commList.map(v=>[v.name,v.payments,v.commission])],`Commission_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>Export CSV</button>
      <button onClick={doPDF} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>PDF</button>
    </div>
    {commList.length===0&&<Empty text="No commission yet. Commission is calculated on payments received."/>}
    {commList.map((v,i)=>(<div key={v.name} style={{background:i%2===0?C.card:"#F8FAFC",borderRadius:12,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",borderLeft:`3px solid ${C.purple}`}}><Row><div><B style={{fontSize:13.5}}>{v.name}</B><Mute>Payment: ₹{fmt(v.payments)}</Mute></div><span style={{fontWeight:900,color:C.purple,fontSize:15}}>₹{fmt(v.commission)}</span></Row></div>))}
    <div style={{background:C.navy,borderRadius:12,padding:"14px 16px",marginTop:8,display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,color:C.gold}}>TOTAL COMMISSION</span><span style={{fontWeight:900,fontSize:16,color:C.gold}}>₹{fmt(totComm)}</span></div>
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
  const doPDF=()=>{
    const rows=[...data.tradingSales].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(s=>`<tr><td>${fmtD(s.date)}</td><td>${s.billNo||"—"}</td><td>${s.customerName}</td><td>${s.productName}</td><td style="text-align:right">${fmt(s.meters)}</td><td style="text-align:right">₹${fmt(s.amount)}</td></tr>`).join("");
    generatePDF("Sales Report",`<table><thead><tr><th>Date</th><th>Bill No</th><th>Customer</th><th>Product</th><th>Meters</th><th>Amount</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="5">TOTAL</td><td>₹${fmt(tots.totTradingSale)}</td></tr></tbody></table>`);
  };
  return(<div>
    <div style={{display:"flex",overflowX:"auto",gap:8,marginBottom:14,scrollbarWidth:"none"}}>
      {[{v:"sales",l:"Sales"},{v:"outstanding",l:"Outstanding"},{v:"commission",l:"Commission"}].map(r=>(<button key={r.v} onClick={()=>setRep(r.v)} style={{flex:"0 0 auto",padding:"10px 16px",borderRadius:20,fontSize:13,fontWeight:rep===r.v?700:500,border:`1.5px solid ${rep===r.v?C.navy:C.border}`,background:rep===r.v?C.navy:"#fff",color:rep===r.v?C.gold:"#666",cursor:"pointer",minHeight:42}}>{r.l}</button>))}
    </div>
    {rep==="sales"&&<>
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        <button onClick={()=>exportCSV([["Date","Bill No","Customer","Product","Supplier","Meters","Rate","Amount"],...data.tradingSales.map(s=>[fmtD(s.date),s.billNo||"",s.customerName,s.productName,s.supplierName,s.meters,s.rate,s.amount])],`Sales_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>Export CSV</button>
        <button onClick={doPDF} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>PDF</button>
      </div>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12}}><Mute>Total Sales</Mute><div style={{fontSize:22,fontWeight:900,color:C.blue}}>₹{fmt(tots.totTradingSale)}</div></div>
      {[...data.tradingSales].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(s=>(<div key={s.id} style={{background:C.card,borderRadius:11,padding:"12px 14px",marginBottom:8}}><Row><B style={{fontSize:13.5}}>{s.customerName}</B><span style={{fontWeight:900,color:C.blue}}>₹{fmt(s.amount)}</span></Row>{s.billNo&&<Mute>Bill No: {s.billNo}</Mute>}<Mute>{s.productName} · {fmt(s.meters)}m · {fmtD(s.date)}</Mute></div>))}
    </>}
    {rep==="outstanding"&&<>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12}}><Mute>Total Outstanding</Mute><div style={{fontSize:22,fontWeight:900,color:C.red}}>₹{fmt(tots.totTradingOut)}</div></div>
      {Object.values(tradingOut).filter(v=>Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))>0).sort((a,b)=>(b.due-b.paid)-(a.due-a.paid)).map(v=>(<div key={v.name} style={{background:C.card,borderRadius:11,padding:"12px 14px",marginBottom:8,borderLeft:`3px solid ${C.red}`}}><Row><B style={{fontSize:13.5}}>{v.name}</B><span style={{fontWeight:900,color:C.red}}>₹{fmt(Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0)))}</span></Row><Mute>Sales: ₹{fmt(v.due)} · Paid: ₹{fmt(v.paid)}</Mute></div>))}
    </>}
    {rep==="commission"&&<>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12}}><Mute>Total Commission</Mute><div style={{fontSize:22,fontWeight:900,color:C.purple}}>₹{fmt(tots.totComm)}</div></div>
      {data.tradingPayments.filter(p=>p.commissionEarned>0).sort((a,b)=>new Date(b.date)-new Date(a.date)).map(p=>(<div key={p.id} style={{background:C.card,borderRadius:11,padding:"12px 14px",marginBottom:8,borderLeft:`3px solid ${C.purple}`}}><Row><B style={{fontSize:13.5}}>{p.customerName}</B><span style={{fontWeight:900,color:C.purple}}>₹{fmt(p.commissionEarned)}</span></Row>{p.billNo&&<Mute>Bill No: {p.billNo}</Mute>}<Mute>Payment: ₹{fmt(p.amount)} · {p.mode} · {fmtD(p.date)}</Mute></div>))}
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
  const[f,sf]=useState(initial||{name:"",type:"Trading",phone:"",city:"",gstin:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  return(<ModalBase title={initial?"Edit Customer":"Add Customer"} onClose={onClose}>
    <F label="Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Customer name" style={IS}/></F>
    <F label="Type"><select value={f.type} onChange={e=>s("type",e.target.value)} style={IS}><option>Trading</option><option>Agency</option><option>Both</option></select></F>
    <F label="Phone"><input type="tel" value={f.phone} onChange={e=>s("phone",e.target.value)} placeholder="10-digit mobile" style={IS}/></F>
    <F label="City"><input value={f.city} onChange={e=>s("city",e.target.value)} placeholder="City" style={IS}/></F>
    <F label="GSTIN"><input value={f.gstin} onChange={e=>s("gstin",e.target.value)} placeholder="GST number" style={IS}/></F>
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
export default App;