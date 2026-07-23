import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, AreaChart, Area, ComposedChart, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import { supabase } from './supabaseClient';
import Login from './Login';

const SK = "fabric-sales-v3-trading";
const EMPTY = { customers:[], suppliers:[], products:[], tradingSales:[], tradingPayments:[], debitNotes:[], creditNotes:[], enquiries:[], trash:[] };
async function loadData(userId){
  try{
    const{data:row,error}=await supabase.from('user_data').select('data').eq('user_id',userId).maybeSingle();
    if(error)console.error('SUPABASE LOAD ERROR:',error);
    if(row&&row.data)return{...EMPTY,...row.data};
  }catch(e){console.error('SUPABASE LOAD EXCEPTION:',e);}
  // one-time migration: no Supabase row yet, fall back to old browser localStorage data if present
  try{const r=localStorage.getItem(SK);if(r)return{...EMPTY,...JSON.parse(r)};}catch(e){}
  return{...EMPTY};
}
async function saveData(userId,d){
  try{
    const{error}=await supabase.from('user_data').upsert({user_id:userId,data:d,updated_at:new Date().toISOString()});
    if(error)console.error('SUPABASE SAVE ERROR:',error);
    else console.log('Saved to Supabase OK');
  }catch(e){console.error('SUPABASE SAVE EXCEPTION:',e);}
}

const fmt   = n => Number(n||0).toLocaleString("en-IN",{maximumFractionDigits:2});
const fmtD  = d => d?new Date(d).toLocaleDateString("en-IN"):"—";
const today = () => new Date().toISOString().split("T")[0];
const uid   = () => Date.now()+"-"+Math.random().toString(36).slice(2,6);
const normName = s => String(s||"").trim().toLowerCase().replace(/\s+/g," ");
function computeCustomerOutstanding(name,data){
  const due=data.tradingSales.filter(s=>s.customerName===name).reduce((a,s)=>a+(+s.amount||0),0);
  const paid=data.tradingPayments.filter(p=>p.customerName===name).reduce((a,p)=>a+(+p.amount||0),0);
  const debit=(data.debitNotes||[]).filter(n=>n.customerName===name).reduce((a,n)=>a+(+n.amount||0),0);
  const credit=(data.creditNotes||[]).filter(n=>n.customerName===name).reduce((a,n)=>a+(+n.amount||0),0);
  return Math.max(0,due+debit-paid-credit);
}
// A customer's still-open (unpaid) sale bills, FIFO order (oldest first) — same pool logic used
// by Outstanding/Ageing/Ledger: existing payments + credit notes are applied oldest-bill-first.
function computeOpenBillsForCustomer(customerName,data){
  const sales=data.tradingSales.filter(s=>s.customerName===customerName).sort((a,b)=>new Date(a.date)-new Date(b.date));
  const paid=data.tradingPayments.filter(p=>p.customerName===customerName).reduce((a,p)=>a+(+p.amount||0),0);
  const credit=(data.creditNotes||[]).filter(n=>n.customerName===customerName).reduce((a,n)=>a+(+n.amount||0),0);
  let rem=paid+credit;
  const open=[];
  sales.forEach(s=>{
    const amt=+s.amount||0;
    const d=Math.min(amt,rem);rem-=d;const left=amt-d;
    if(left>0)open.push({billNo:s.billNo||"—",date:s.date,billAmount:amt,outstanding:left,meters:+s.meters||0});
  });
  return open;
}
// Simulates what a NEW payment of `newAmount` would do, on top of what's already paid — FIFO,
// oldest open bill first. A bill only contributes commission once it is FULLY covered; a payment
// that only partially covers the next bill in line earns no commission yet (matches calcFIFOCommission).
function simulatePaymentImpact(customerName,newAmount,data){
  const openBills=computeOpenBillsForCustomer(customerName,data);
  let rem=+newAmount||0;
  let billsCleared=0,commissionGained=0,nextBillShortfall=0;
  for(const b of openBills){
    if(rem<=0)break;
    if(rem>=b.outstanding){rem-=b.outstanding;billsCleared++;commissionGained+=b.meters*1.5;}
    else{nextBillShortfall=b.outstanding-rem;rem=0;break;}
  }
  return{billsCleared,commissionGained,totalOpenBills:openBills.length,nextBillShortfall};
}
// Finds every still-open sale bill and its due date (Bill Date + that customer's Credit Days), sorted soonest-due first
function computeDueBills(data){
  const todayMs=new Date().setHours(0,0,0,0);
  const creditDaysMap={};
  data.customers.forEach(c=>{creditDaysMap[c.name]=parseInt(c.creditDays||30);});
  const paid={};data.tradingPayments.forEach(p=>{paid[p.customerName]=(paid[p.customerName]||0)+(+p.amount||0);});
  const credit={};(data.creditNotes||[]).forEach(n=>{credit[n.customerName]=(credit[n.customerName]||0)+(+n.amount||0);});
  const byCust={};
  data.tradingSales.forEach(s=>{if(!byCust[s.customerName])byCust[s.customerName]=[];byCust[s.customerName].push(s);});
  const result=[];
  Object.keys(byCust).forEach(name=>{
    let rem=(paid[name]||0)+(credit[name]||0);
    const creditDays=creditDaysMap[name]||30;
    [...byCust[name]].sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach(s=>{
      const amt=+s.amount||0;
      const d=Math.min(amt,rem);rem-=d;const left=amt-d;
      if(left>0){
        const dd=new Date(s.date);dd.setDate(dd.getDate()+creditDays);
        const dueMs=dd.setHours(0,0,0,0);
        const daysUntilDue=Math.round((dueMs-todayMs)/86400000);
        result.push({customerName:name,billNo:s.billNo||"—",amount:left,date:s.date,dueDate:new Date(dueMs),daysUntilDue});
      }
    });
  });
  return result.sort((a,b)=>a.daysUntilDue-b.daysUntilDue);
}

// ── GST helpers ──────────────────────────────────────────────────
// Seller (Navkar Fabrics / Amihem) is based in Delhi — used to decide CGST+SGST vs IGST.
const SELLER_STATE="Delhi";
const INDIA_STATES=["Delhi","Uttar Pradesh","Haryana","Punjab","Rajasthan","Gujarat","Maharashtra","Madhya Pradesh","Uttarakhand","Bihar","West Bengal","Tamil Nadu","Karnataka","Andhra Pradesh","Telangana","Kerala","Odisha","Chhattisgarh","Jharkhand","Assam","Himachal Pradesh","Jammu and Kashmir","Goa","Chandigarh","Puducherry","Other"];
// Amount is treated as GST-INCLUSIVE. Splits it into taxable value + CGST/SGST (same state) or IGST (different state).
function calcGST(amount,buyerState,gstRatePct){
  const rate=(+gstRatePct||5);
  const amt=+amount||0;
  const taxable=amt/(1+rate/100);
  const gstAmt=amt-taxable;
  const sameState=normName(buyerState||SELLER_STATE)===normName(SELLER_STATE);
  return{
    taxable,gstAmt,rate,sameState,
    cgst:sameState?gstAmt/2:0,
    sgst:sameState?gstAmt/2:0,
    igst:sameState?0:gstAmt,
  };
}
// Finds an existing customer whose normalized name matches (used to catch duplicates like extra spaces / different casing)
function findDuplicateCustomer(name,customers,excludeId){
  const n=normName(name);if(!n)return null;
  return (customers||[]).find(c=>c.id!==excludeId&&normName(c.name)===n)||null;
}
// Warns (via confirm) only when the typed name is a near-duplicate of an existing customer (same normalized form, different exact text).
// Returns true if it's safe to proceed with saving.
function warnIfNearDuplicateCustomer(name,customers){
  const match=findDuplicateCustomer(name,customers,null);
  if(match&&match.name!==name){
    return window.confirm(`⚠️ A customer named "${match.name}" already exists.\n\nYou typed "${name}" — this looks like it could be the same customer with a typo/spacing difference, which would split their ledger into two.\n\nTap OK to save anyway with "${name}", or Cancel to fix it.`);
  }
  return true;
}
function findDuplicateBillNo(billNo,sales,excludeId){
  const b=String(billNo||"").trim().toLowerCase();if(!b)return null;
  return (sales||[]).find(s=>s.id!==excludeId&&String(s.billNo||"").trim().toLowerCase()===b)||null;
}
// Warns when a Bill No already exists elsewhere in the Sales register (possible duplicate/double entry). Returns true if safe to proceed.
function warnIfDuplicateBillNo(billNo,sales,excludeId){
  const match=findDuplicateBillNo(billNo,sales,excludeId);
  if(match){
    return window.confirm(`⚠️ Bill No "${billNo}" already exists — ${match.customerName}, ₹${fmt(match.amount)}, ${fmtD(match.date)}.\n\nThis may be a duplicate entry. Tap OK to save anyway, or Cancel to review.`);
  }
  return true;
}
function findDuplicateNoteNo(noteNo,notes,excludeId){
  const n=String(noteNo||"").trim().toLowerCase();if(!n)return null;
  return (notes||[]).find(x=>x.id!==excludeId&&String(x.noteNo||"").trim().toLowerCase()===n)||null;
}
function warnIfDuplicateNoteNo(noteNo,notes,excludeId,label){
  const match=findDuplicateNoteNo(noteNo,notes,excludeId);
  if(match){
    return window.confirm(`⚠️ ${label} No "${noteNo}" already exists — ${match.customerName}, ₹${fmt(match.amount)}, ${fmtD(match.date)}.\n\nThis may be a duplicate entry. Tap OK to save anyway, or Cancel to review.`);
  }
  return true;
}
const TRASH_SECTION_LABELS={customers:"Customer",suppliers:"Supplier",products:"Product",tradingSales:"Sale",tradingPayments:"Payment",debitNotes:"Debit Note",creditNotes:"Credit Note",enquiries:"Enquiry"};
function describeTrashEntry(entry){
  const r=entry.record;const type=TRASH_SECTION_LABELS[entry.section]||entry.section;
  let title="—",sub="";
  if(entry.section==="customers"||entry.section==="suppliers"){title=r.name;sub=[r.phone,r.city].filter(Boolean).join(" · ");}
  else if(entry.section==="products"){title=r.name;sub=[r.supplierName,r.unit].filter(Boolean).join(" · ");}
  else if(entry.section==="tradingSales"){title=r.customerName;sub=`${r.productName||""} · ₹${fmt(r.amount)} · ${fmtD(r.date)}`;}
  else if(entry.section==="tradingPayments"){title=r.customerName;sub=`₹${fmt(r.amount)} · ${r.mode||""} · ${fmtD(r.date)}`;}
  else if(entry.section==="debitNotes"||entry.section==="creditNotes"){title=r.customerName;sub=`₹${fmt(r.amount)} · ${fmtD(r.date)}`;}
  else if(entry.section==="enquiries"){title=r.customerName;sub=(r.description||"").slice(0,60);}
  return{type,title,sub};
}
function timeAgo(iso){
  if(!iso)return"";
  const mins=Math.floor((Date.now()-new Date(iso).getTime())/60000);
  if(mins<1)return"just now";
  if(mins<60)return`${mins}m ago`;
  const hrs=Math.floor(mins/60);if(hrs<24)return`${hrs}h ago`;
  const days=Math.floor(hrs/24);return`${days}d ago`;
}
function excelDateToJS(v){if(v instanceof Date)return v;if(typeof v==="number")return new Date(Math.round((v-25569)*86400*1000));if(typeof v==="string"){const d=new Date(v);if(!isNaN(d))return d;}return null;}
function exportBackup(data){const b=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=`AMIHEMBusiness_${today()}.json`;a.click();try{localStorage.setItem(SK+"-lastbackup",new Date().toISOString());}catch(e){}}
function getLastBackup(){try{return localStorage.getItem(SK+"-lastbackup");}catch(e){return null;}}
async function shareBackupNative(data){
  const filename=`AMIHEMBusiness_${today()}.json`;
  try{
    const file=new File([JSON.stringify(data,null,2)],filename,{type:"application/json"});
    if(navigator.canShare&&navigator.canShare({files:[file]})){
      await navigator.share({files:[file],title:"AMIHEM Business Backup",text:`AMIHEM Business backup — ${today()}`});
      try{localStorage.setItem(SK+"-lastbackup",new Date().toISOString());}catch(e){}
      return "shared";
    }
  }catch(e){
    if(e && e.name==="AbortError")return "cancelled";
  }
  return "unsupported";
}
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
  win.document.write(`<!DOCTYPE html><html><head><title>${title}</title><style>body{font-family:Arial,sans-serif;font-size:13px;color:#111;padding:30px;max-width:800px;margin:auto;}h1{font-size:20px;border-bottom:2px solid #111;padding-bottom:8px;}table{width:100%;border-collapse:collapse;margin-top:16px;}th{background:#0F1923;color:#E8C97E;padding:8px 10px;text-align:left;font-size:12px;}td{padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;}tr:nth-child(even){background:#f9f9f9;}.tot td{background:#0F1923;color:#E8C97E;font-weight:bold;border:none;}.hdr{display:flex;justify-content:space-between;margin-bottom:20px;}.co{font-size:22px;font-weight:bold;color:#0F1923;}</style></head><body><div class="hdr"><div><div class="co">AMIHEM Business</div><div style="font-size:12px;color:#666">Trading Sales Report</div></div><div style="text-align:right;font-size:12px;color:#666">Date: ${new Date().toLocaleDateString("en-IN")}</div></div><h1>${title}</h1>${content}<script>window.onload=()=>{window.print();}<\/script></body></html>`);
  win.document.close();
}

// Groups every sale row sharing the same Bill No + Customer into one invoice's line items (falls back to the single row if Bill No is blank)
function getInvoiceLineItems(sale,allSales){
  if(!sale.billNo||!sale.billNo.trim())return[sale];
  const b=sale.billNo.trim().toLowerCase();
  const matched=allSales.filter(s=>s.customerName===sale.customerName&&(s.billNo||"").trim().toLowerCase()===b);
  return matched.length?matched:[sale];
}
function generateInvoice(sale,allSales,customer){
  const items=getInvoiceLineItems(sale,allSales);
  const total=items.reduce((a,s)=>a+(+s.amount||0),0);
  const totalQty=items.reduce((a,s)=>a+(+s.meters||0),0);
  const buyerState=customer?.state||SELLER_STATE;
  const gstBreak=items.reduce((acc,s)=>{
    const g=calcGST(s.amount,buyerState,s.gstRate);
    acc.taxable+=g.taxable;acc.cgst+=g.cgst;acc.sgst+=g.sgst;acc.igst+=g.igst;acc.rate=g.rate;acc.sameState=g.sameState;
    return acc;
  },{taxable:0,cgst:0,sgst:0,igst:0,rate:5,sameState:true});
  const rows=items.map((s,i)=>`<tr><td>${i+1}</td><td>${s.productName||"—"}</td><td style="text-align:right">${fmt(s.meters)}</td><td style="text-align:right">₹${fmt(s.rate)}</td><td style="text-align:right">₹${fmt(s.amount)}</td></tr>`).join("");
  const gstRows=gstBreak.sameState
    ?`<tr><td colspan="4" style="text-align:right">CGST @ ${gstBreak.rate/2}%</td><td style="text-align:right">₹${fmt(gstBreak.cgst)}</td></tr>
      <tr><td colspan="4" style="text-align:right">SGST @ ${gstBreak.rate/2}%</td><td style="text-align:right">₹${fmt(gstBreak.sgst)}</td></tr>`
    :`<tr><td colspan="4" style="text-align:right">IGST @ ${gstBreak.rate}% (${buyerState})</td><td style="text-align:right">₹${fmt(gstBreak.igst)}</td></tr>`;
  const win=window.open("","_blank");
  win.document.write(`<!DOCTYPE html><html><head><title>Invoice ${sale.billNo||sale.id}</title><style>
    body{font-family:Arial,sans-serif;font-size:13px;color:#111;padding:30px;max-width:800px;margin:auto;}
    .hdr{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #0F1923;padding-bottom:16px;margin-bottom:20px;}
    .co{font-size:24px;font-weight:900;color:#0F1923;}
    .tag{font-size:11px;color:#666;margin-top:2px;}
    .inv-title{font-size:20px;font-weight:800;color:#0F1923;text-align:right;}
    .inv-meta{font-size:12px;color:#555;text-align:right;margin-top:4px;line-height:1.6;}
    .parties{display:flex;justify-content:space-between;margin-bottom:24px;gap:20px;}
    .party-box{flex:1;}
    .party-label{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:1px;font-weight:700;margin-bottom:4px;}
    .party-name{font-size:15px;font-weight:800;color:#0F1923;}
    .party-detail{font-size:12px;color:#555;margin-top:2px;}
    table{width:100%;border-collapse:collapse;margin-top:10px;}
    th{background:#0F1923;color:#E8C97E;padding:9px 10px;text-align:left;font-size:11.5px;}
    td{padding:8px 10px;border-bottom:1px solid #eee;font-size:12.5px;}
    tr:nth-child(even){background:#f9f9f9;}
    .sub-row td{background:#F7F9FB;border:none;font-size:12px;}
    .taxable-row td{border:none;font-size:12.5px;font-weight:700;border-top:1.5px solid #0F1923;}
    .tot-row td{background:#0F1923;color:#E8C97E;font-weight:bold;border:none;font-size:14px;}
    .footer{margin-top:40px;display:flex;justify-content:space-between;font-size:11px;color:#888;border-top:1px solid #eee;padding-top:16px;}
    .sign-line{margin-top:50px;border-top:1px solid #333;padding-top:6px;font-size:11px;text-align:center;}
  </style></head><body>
    <div class="hdr">
      <div><div class="co">Navkar Fabrics</div><div class="tag">Fabric Trading &amp; Indenting — Delhi</div></div>
      <div><div class="inv-title">TAX INVOICE</div><div class="inv-meta">Invoice No: ${sale.billNo||"—"}<br/>Date: ${fmtD(sale.date)}<br/>${gstBreak.sameState?"Intra-state (CGST+SGST)":"Inter-state (IGST)"}</div></div>
    </div>
    <div class="parties">
      <div class="party-box"><div class="party-label">Bill To</div><div class="party-name">${sale.customerName}</div>
        ${customer&&customer.phone?`<div class="party-detail">📞 ${customer.phone}</div>`:""}
        ${customer&&customer.city?`<div class="party-detail">📍 ${customer.city}, ${buyerState}</div>`:`<div class="party-detail">📍 ${buyerState}</div>`}
        ${customer&&customer.gstin?`<div class="party-detail">GSTIN: ${customer.gstin}</div>`:""}
      </div>
    </div>
    <table><thead><tr><th>#</th><th>Product</th><th style="text-align:right">Qty (m)</th><th style="text-align:right">Rate</th><th style="text-align:right">Amount</th></tr></thead>
    <tbody>${rows}
      <tr class="taxable-row"><td colspan="4" style="text-align:right">Taxable Value</td><td style="text-align:right">₹${fmt(gstBreak.taxable)}</td></tr>
      ${gstRows}
      <tr class="tot-row"><td colspan="4">GRAND TOTAL</td><td style="text-align:right">₹${fmt(total)}</td></tr>
    </tbody></table>
    <div style="font-size:10.5px;color:#999;margin-top:6px;">Total Qty: ${fmt(totalQty)} m · Prices are GST-inclusive.</div>
    <div class="footer">
      <div>Thank you for your business.</div>
      <div class="sign-line">Authorized Signatory</div>
    </div>
    <script>window.onload=()=>{window.print();}<\/script>
  </body></html>`);
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
  {id:"Trading",icon:"🏪",label:"Entry"},
  {id:"Outstanding",icon:"⏳",label:"Outstanding"},
  {id:"Ledger",icon:"📖",label:"Ledger"},
  {id:"Aging",icon:"📅",label:"Ageing"},
  {id:"Analytics",icon:"📈",label:"Analytics"},
  {id:"Commission",icon:"💰",label:"Commission"},
  {id:"Enquiry",icon:"📝",label:"Enquiry"},
  {id:"Masters",icon:"⚙️",label:"Masters"},
  {id:"Reports",icon:"📋",label:"Reports"},
];

function App({user}){
  const[tab,setTab]=useState("Dashboard");
  const[data,setData]=useState(null);
  const[modal,setModal]=useState(null);
  const[toast,setToast]=useState(null);
  const[lastBackupAt,setLastBackupAt]=useState(null);
  const[searchOpen,setSearchOpen]=useState(false);
  const[navHint,setNavHint]=useState(null);
  const restoreRef=useRef(null);const importRef=useRef(null);
  const navigateTab=(tabId,hint=null)=>{setNavHint(hint);setTab(tabId);};

  useEffect(()=>{loadData(user.id).then(setData);setLastBackupAt(getLastBackup());},[]);
  useEffect(()=>{if(data)saveData(user.id,data);},[data]);

  const showToast=(msg,err,action)=>{setToast({msg,err,action});setTimeout(()=>setToast(null),action?7000:3000);};
  const add=(section,rec)=>{setData(p=>({...p,[section]:[...p[section],{...rec,id:uid()}]}));showToast("Saved successfully.");setModal(null);};
  const del=(section,id)=>{
    if(!window.confirm("Are you sure you want to delete this record? It will be moved to Trash and can be restored later from Masters → Trash."))return;
    const record=data[section].find(r=>r.id===id);
    if(!record)return;
    const trashId=uid();
    setData(p=>({
      ...p,
      [section]:p[section].filter(r=>r.id!==id),
      trash:[{trashId,section,record,deletedAt:new Date().toISOString()},...(p.trash||[])].slice(0,20)
    }));
    showToast("Deleted.",false,{label:"Undo",onClick:()=>restoreTrash(trashId)});
  };
  const restoreTrash=(trashId)=>{
    setData(p=>{
      const entry=(p.trash||[]).find(t=>t.trashId===trashId);
      if(!entry)return p;
      return{...p,[entry.section]:[...p[entry.section],entry.record],trash:p.trash.filter(t=>t.trashId!==trashId)};
    });
    showToast("Restored.");
  };
  const permanentDeleteTrash=(trashId)=>{
    if(!window.confirm("Permanently delete this record? This cannot be undone."))return;
    setData(p=>({...p,trash:(p.trash||[]).filter(t=>t.trashId!==trashId)}));
    showToast("Permanently deleted.");
  };
  const emptyTrash=()=>{
    if(!window.confirm("Empty the trash? All recoverable records will be permanently deleted."))return;
    setData(p=>({...p,trash:[]}));
    showToast("Trash emptied.");
  };
  const updateMaster=(section,updated)=>{setData(p=>({...p,[section]:p[section].map(r=>r.id===updated.id?updated:r)}));showToast("Updated.");setModal(null);};

  const handleBackupClick=()=>{exportBackup(data);setLastBackupAt(getLastBackup());showToast("Backup downloaded.");};
  const handleShareBackup=async()=>{
    const result=await shareBackupNative(data);
    if(result==="shared"){setLastBackupAt(getLastBackup());showToast("Backup shared successfully.");}
    else if(result==="cancelled"){/* user cancelled the share sheet — do nothing */}
    else{exportBackup(data);setLastBackupAt(getLastBackup());showToast("Direct share isn't supported here — file downloaded, attach it manually in WhatsApp/Email.");}
  };
  const handleLogout=()=>{
    if(!window.confirm("Log out of this account?\n\nYour data stays safely synced in the cloud — you'll just need to sign in again to see it here."))return;
    supabase.auth.signOut();
  };

  const[excelPreview,setExcelPreview]=useState(null);
  const[backupPreview,setBackupPreview]=useState(null);

  const handleExcelImport=(e)=>{
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
        const parsed=parseExcelData(rows);
        if(!parsed||parsed.tradingSales.length===0){showToast("Could not parse Excel.",true);e.target.value="";return;}
        const ec=new Set(data.customers.map(c=>c.name.toLowerCase()));
        const es=new Set(data.suppliers.map(s=>s.name.toLowerCase()));
        const ep=new Set(data.products.map(x=>x.name.toLowerCase()));
        setExcelPreview({
          fileName:file.name,parsed,
          newCustomers:parsed.customers.filter(c=>!ec.has(c.name.toLowerCase())).length,
          newSuppliers:parsed.suppliers.filter(s=>!es.has(s.name.toLowerCase())).length,
          newProducts:parsed.products.filter(x=>!ep.has(x.name.toLowerCase())).length,
        });
      }catch{showToast("Error reading file.",true);}
      e.target.value="";
    };
    reader.readAsArrayBuffer(file);
  };
  const confirmExcelImport=()=>{
    if(!excelPreview)return;
    const parsed=excelPreview.parsed;
    setData(p=>{
      const ec=new Set(p.customers.map(c=>c.name.toLowerCase()));
      const es=new Set(p.suppliers.map(s=>s.name.toLowerCase()));
      const ep=new Set(p.products.map(x=>x.name.toLowerCase()));
      return{...p,customers:[...p.customers,...parsed.customers.filter(c=>!ec.has(c.name.toLowerCase()))],suppliers:[...p.suppliers,...parsed.suppliers.filter(s=>!es.has(s.name.toLowerCase()))],products:[...p.products,...parsed.products.filter(x=>!ep.has(x.name.toLowerCase()))],tradingSales:[...p.tradingSales,...parsed.tradingSales],tradingPayments:[...p.tradingPayments,...parsed.tradingPayments]};
    });
    showToast(`Imported ${parsed.tradingSales.length} sales, ${excelPreview.newCustomers} new customers.`);
    setExcelPreview(null);
  };

  const importBackup=(e)=>{
    const f=e.target.files[0];if(!f)return;
    const r=new FileReader();
    r.onload=ev=>{
      let parsed;
      try{parsed=JSON.parse(ev.target.result);}catch{showToast("Invalid backup file.",true);e.target.value="";return;}
      setBackupPreview({fileName:f.name,parsed});
      e.target.value="";
    };
    r.onerror=()=>{showToast("Could not read file.",true);e.target.value="";};
    r.readAsText(f);
  };
  const confirmRestoreBackup=()=>{
    if(!backupPreview)return;
    if(data)exportBackup(data); // safety copy of current data before overwrite
    setData({...EMPTY,...backupPreview.parsed});
    showToast("Backup restored. Your previous data was saved as a safety download.");
    setBackupPreview(null);
  };

  const handleSearchSelect=(r)=>{
    setSearchOpen(false);
    if(r.kind==="customer")navigateTab("Ledger",{customer:r.name});
    else if(r.kind==="sale")navigateTab("Trading",{view:"sales",search:r.customerName});
    else if(r.kind==="payment")navigateTab("Trading",{view:"payments",search:r.customerName});
    else if(r.kind==="debit")navigateTab("Trading",{view:"debit",search:r.customerName});
    else if(r.kind==="credit")navigateTab("Trading",{view:"credit",search:r.customerName});
    else if(r.kind==="product")navigateTab("Masters",{view:"products",search:r.name});
    else if(r.kind==="supplier")navigateTab("Masters",{view:"suppliers",search:r.name});
    else if(r.kind==="enquiry")navigateTab("Enquiry",null);
  };

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
  const totComm=calcFIFOCommission(data.tradingSales,data.tradingPayments,data.creditNotes||[],data.debitNotes||[]).totalCommission;
  const overdueCount=Object.values(tradingOut).filter(v=>Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))>0).length;

  const Sidebar=()=>(
    <div style={{width:220,background:C.navy,minHeight:"100vh",position:"fixed",left:0,top:0,zIndex:150,display:"flex",flexDirection:"column",padding:"20px 0"}}>
      <div style={{padding:"0 20px 24px"}}><div style={{fontSize:10,letterSpacing:2.5,color:C.gold,textTransform:"uppercase",fontWeight:700}}>AMIHEM</div><div style={{fontSize:20,fontWeight:900,color:"#fff",marginTop:4}}>Business</div><div style={{fontSize:9,color:"rgba(255,255,255,0.4)",marginTop:3,letterSpacing:0.3}}>Sales • Inventory • Collections</div></div>
      <div style={{padding:"0 20px 14px"}}>
        <button onClick={()=>setSearchOpen(true)} style={{width:"100%",background:"rgba(255,255,255,0.08)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,padding:"9px 12px",color:"rgba(255,255,255,0.6)",fontSize:12.5,textAlign:"left",cursor:"pointer",display:"flex",alignItems:"center",gap:8}}>🔍 Search everything…</button>
      </div>
      {TABS.map(t=>(<button key={t.id} onClick={()=>navigateTab(t.id)} style={{background:tab===t.id?"rgba(232,201,126,0.12)":"transparent",border:"none",borderLeft:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",padding:"12px 20px",textAlign:"left",color:tab===t.id?C.gold:"rgba(255,255,255,0.65)",fontSize:13.5,fontWeight:tab===t.id?700:400,cursor:"pointer",display:"flex",alignItems:"center",gap:10}}><span>{t.icon}</span>{t.label}{t.id==="Outstanding"&&overdueCount>0&&<span style={{background:C.red,color:"#fff",borderRadius:20,fontSize:10,fontWeight:800,padding:"2px 7px",marginLeft:"auto"}}>{overdueCount}</span>}</button>))}
      <div style={{marginTop:"auto",padding:"16px 20px",borderTop:"1px solid rgba(255,255,255,0.08)",display:"flex",flexDirection:"column",gap:8}}>
        <button onClick={()=>importRef.current&&importRef.current.click()} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📊 <span style={{fontSize:12}}>Import Excel</span></button>
        <button onClick={handleBackupClick} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>☁️ <span style={{fontSize:12}}>Backup</span></button>
        <button onClick={handleShareBackup} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📤 <span style={{fontSize:12}}>Share (WhatsApp/Email)</span></button>
        <button onClick={()=>restoreRef.current&&restoreRef.current.click()} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px"}}>📂 <span style={{fontSize:12}}>Restore</span></button>
        <div style={{borderTop:"1px solid rgba(255,255,255,0.08)",marginTop:4,paddingTop:10}}>
          {user?.email&&<div style={{fontSize:10.5,color:"rgba(255,255,255,0.4)",padding:"0 2px 8px",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{user.email}</div>}
          <button onClick={handleLogout} style={{...hdrBtn(),flexDirection:"row",gap:8,width:"100%",justifyContent:"flex-start",padding:"10px 12px",background:"rgba(231,76,60,0.15)",border:"1px solid rgba(231,76,60,0.3)"}}>🚪 <span style={{fontSize:12,color:"#F5B7B1"}}>Log Out</span></button>
        </div>
      </div>
    </div>
  );

  return(
    <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:C.bg,minHeight:"100vh"}}>
      <style>{`@media(min-width:768px){.mob{display:none!important}.desk{display:flex!important}.main{margin-left:220px!important}}@media(max-width:767px){.mob{display:flex!important}.desk{display:none!important}.main{margin-left:0!important;padding-bottom:80px!important}}.main{max-width:860px;padding:18px 16px 40px;}`}</style>
      <div className="desk" style={{display:"none"}}><Sidebar/></div>
      <div className="mob" style={{display:"none",background:C.navy,padding:"calc(env(safe-area-inset-top,0px)+12px) 16px 10px",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 20px rgba(0,0,0,0.4)",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <div><div style={{fontSize:9,letterSpacing:2,color:C.gold,textTransform:"uppercase",fontWeight:700}}>AMIHEM</div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>Business</div></div>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>importRef.current&&importRef.current.click()} style={hdrBtn()}><span style={{fontSize:18}}>📊</span><span style={{fontSize:9}}>Excel</span></button>
          <button onClick={handleBackupClick} style={hdrBtn()}><span style={{fontSize:18}}>☁️</span><span style={{fontSize:9}}>Backup</span></button>
          <button onClick={handleShareBackup} style={hdrBtn()}><span style={{fontSize:18}}>📤</span><span style={{fontSize:9}}>Share</span></button>
          <button onClick={()=>restoreRef.current&&restoreRef.current.click()} style={hdrBtn()}><span style={{fontSize:18}}>📂</span><span style={{fontSize:9}}>Restore</span></button>
          <button onClick={handleLogout} style={{...hdrBtn(),minWidth:44,background:"rgba(231,76,60,0.18)"}}><span style={{fontSize:18}}>🚪</span></button>
        </div>
      </div>
      <div className="mob" style={{display:"none",background:C.navy,padding:"0 16px 10px"}}>
        <button onClick={()=>setSearchOpen(true)} style={{width:"100%",background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.18)",borderRadius:10,padding:"10px 12px",color:"rgba(255,255,255,0.65)",fontSize:13,textAlign:"left",cursor:"pointer"}}>🔍 Search everything…</button>
      </div>
      <div className="mob" style={{display:"none",background:C.navyMid,overflowX:"auto",WebkitOverflowScrolling:"touch",scrollbarWidth:"none",position:"sticky",top:72,zIndex:99}}>
        {TABS.map(t=>(<button key={t.id} onClick={()=>navigateTab(t.id)} style={{flex:"0 0 auto",padding:"12px 14px",fontSize:12,fontWeight:tab===t.id?800:500,color:tab===t.id?C.gold:"rgba(255,255,255,0.6)",background:"none",border:"none",borderBottom:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",cursor:"pointer",whiteSpace:"nowrap",minHeight:44,position:"relative"}}>{t.icon} {t.label}{t.id==="Outstanding"&&overdueCount>0&&<span style={{position:"absolute",top:6,right:4,background:C.red,color:"#fff",borderRadius:10,fontSize:9,fontWeight:800,padding:"1px 5px"}}>{overdueCount}</span>}</button>))}
      </div>
      <div className="main" style={{marginLeft:0,padding:"18px 16px 40px",maxWidth:860}}>
        {tab==="Dashboard"  &&<DashboardTab data={data} tots={{totTradingSale,totTradingPaid,totTradingOut,totComm}} onNav={id=>navigateTab(id)} tradingOut={tradingOut} lastBackupAt={lastBackupAt} onShareBackup={handleShareBackup}/>}
        {tab==="Trading"    &&<TradingTab data={data} onAdd={()=>setModal({type:"sale"})} onAddPay={()=>setModal({type:"payment"})} onAddDebit={()=>setModal({type:"debit"})} onAddCredit={()=>setModal({type:"credit"})} onDel={del} onEdit={(type,rec)=>setModal({type:`edit-${type}`,rec})} tradingOut={tradingOut} initialSearch={navHint?.search} initialView={navHint?.view}/>}
        {tab==="Outstanding"&&<OutstandingTab tradingOut={tradingOut} data={data} onAddPay={(pre)=>setModal({type:"payment",preCustomer:pre})} generatePDF={generatePDF}/>}
        {tab==="Ledger"     &&<LedgerTab data={data} generatePDF={generatePDF} initialCustomer={navHint?.customer}/>}
        {tab==="Aging"      &&<AgingTab data={data} generatePDF={generatePDF}/>}
        {tab==="Analytics"  &&<AnalyticsTab data={data} tradingOut={tradingOut}/>}
        {tab==="Commission" &&<CommissionTab data={data} generatePDF={generatePDF}/>}
        {tab==="Enquiry"    &&<EnquiryTab data={data} onAdd={r=>add("enquiries",r)} onUpdate={r=>{setData(p=>({...p,enquiries:p.enquiries.map(e=>e.id===r.id?r:e)}));showToast("Updated.");}} onDel={id=>del("enquiries",id)}/>}
        {tab==="Masters"    &&<MastersTab data={data} onAdd={setModal} onDel={del} onEdit={(type,rec)=>setModal({type:`edit-${type}`,rec})} onImportExcel={()=>importRef.current&&importRef.current.click()} trash={data.trash||[]} onRestore={restoreTrash} onPermanentDelete={permanentDeleteTrash} onEmptyTrash={emptyTrash} initialSearch={navHint?.search} initialView={navHint?.view}/>}
        {tab==="Reports"    &&<ReportsTab data={data} tradingOut={tradingOut} tots={{totTradingSale,totTradingPaid,totTradingOut,totComm}} generatePDF={generatePDF}/>}
      </div>
      {modal?.type==="sale"          &&<SaleModal    data={data} onSave={r=>add("tradingSales",r)}    onClose={()=>setModal(null)}/>}
      {modal?.type==="payment"       &&<PaymentModal data={data} onSave={r=>add("tradingPayments",r)} onClose={()=>setModal(null)} preCustomer={modal.preCustomer}/>}
      {modal?.type==="debit"         &&<DebitModal   data={data} onSave={r=>add("debitNotes",r)}      onClose={()=>setModal(null)}/>}
      {modal?.type==="credit"        &&<CreditModal  data={data} onSave={r=>add("creditNotes",r)}     onClose={()=>setModal(null)}/>}
      {modal?.type==="edit-sale"     &&<SaleModal    data={data} onSave={r=>updateMaster("tradingSales",{...modal.rec,...r})}    onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-payment"  &&<PaymentModal data={data} onSave={r=>updateMaster("tradingPayments",{...modal.rec,...r})} onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-debit"    &&<DebitModal   data={data} onSave={r=>updateMaster("debitNotes",{...modal.rec,...r})}      onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-credit"   &&<CreditModal  data={data} onSave={r=>updateMaster("creditNotes",{...modal.rec,...r})}     onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="customer"      &&<CustomerModal onSave={r=>add("customers",r)}                  onClose={()=>setModal(null)} existingCustomers={data.customers}/>}
      {modal?.type==="supplier"      &&<SupplierModal onSave={r=>add("suppliers",r)}                  onClose={()=>setModal(null)}/>}
      {modal?.type==="product"       &&<ProductModal  data={data} onSave={r=>add("products",r)}       onClose={()=>setModal(null)}/>}
      {modal?.type==="edit-customer" &&<CustomerModal onSave={r=>updateMaster("customers",r)}         onClose={()=>setModal(null)} initial={modal.rec} existingCustomers={data.customers}/>}
      {modal?.type==="edit-supplier" &&<SupplierModal onSave={r=>updateMaster("suppliers",r)}         onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-product"  &&<ProductModal  data={data} onSave={r=>updateMaster("products",r)} onClose={()=>setModal(null)} initial={modal.rec}/>}
      <input ref={importRef} type="file" accept=".xlsx,.xls" onChange={handleExcelImport} style={{display:"none"}}/>
      <input ref={restoreRef} type="file" accept=".json" onChange={importBackup} style={{display:"none"}}/>
      {excelPreview&&<ExcelImportPreviewModal preview={excelPreview} onConfirm={confirmExcelImport} onCancel={()=>setExcelPreview(null)}/>}
      {searchOpen&&<GlobalSearchModal data={data} onClose={()=>setSearchOpen(false)} onSelect={handleSearchSelect}/>}
      {backupPreview&&<BackupPreviewModal preview={backupPreview} currentData={data} onConfirm={confirmRestoreBackup} onCancel={()=>setBackupPreview(null)}/>}
      {toast&&<div style={{position:"fixed",bottom:"calc(env(safe-area-inset-bottom,0px) + 24px)",left:"50%",transform:"translateX(-50%)",background:toast.err?"#B03A2E":C.navy,color:C.gold,padding:"11px 16px 11px 24px",borderRadius:24,fontSize:13.5,fontWeight:700,zIndex:999,boxShadow:"0 4px 20px rgba(0,0,0,0.4)",whiteSpace:"nowrap",border:"1px solid rgba(232,201,126,0.3)",maxWidth:"90%",display:"flex",alignItems:"center",gap:12}}>
        <span>{toast.msg}</span>
        {toast.action&&<button onClick={()=>{toast.action.onClick();setToast(null);}} style={{background:"rgba(232,201,126,0.18)",color:C.gold,border:"1px solid rgba(232,201,126,0.5)",borderRadius:16,padding:"6px 14px",fontSize:12.5,fontWeight:800,cursor:"pointer",flexShrink:0}}>{toast.action.label}</button>}
      </div>}
    </div>
  );
}

// ─── DASHBOARD ───────────────────────────────────────────────────
function DashboardTab({data,tots,onNav,tradingOut,lastBackupAt,onShareBackup}){
  const{totTradingSale,totTradingPaid,totTradingOut,totComm}=tots;
  const recent=[...data.tradingSales].sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,8);
  const outList=Object.values(tradingOut).filter(v=>Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0))>0).sort((a,b)=>b.due-b.paid-(a.due-a.paid)).slice(0,5);
  const daysSinceBackup=lastBackupAt?Math.floor((Date.now()-new Date(lastBackupAt).getTime())/86400000):null;
  const backupOverdue=daysSinceBackup===null||daysSinceBackup>=7;

  const phoneMap={};data.customers.forEach(c=>{phoneMap[c.name]=c.phone;});
  const dueBills=computeDueBills(data);
  const overdueList=dueBills.filter(b=>b.daysUntilDue<0);
  const dueTodayList=dueBills.filter(b=>b.daysUntilDue===0);
  const dueTomorrowList=dueBills.filter(b=>b.daysUntilDue===1);
  const dueWeekList=dueBills.filter(b=>b.daysUntilDue>=2&&b.daysUntilDue<=7);
  const urgentList=[...overdueList,...dueTodayList,...dueTomorrowList].slice(0,6);
  const dueStatusText=(b)=>b.daysUntilDue<0?`overdue by ${-b.daysUntilDue}d`:b.daysUntilDue===0?"due today":b.daysUntilDue===1?"due tomorrow":`due in ${b.daysUntilDue}d`;
  const buildDueWA=(b)=>`🏢 *Navkar Fabrics*\n\nDear *${b.customerName}*,\n\nReminder: Payment of ₹${fmt(b.amount)} for Bill No ${b.billNo} (dated ${fmtD(b.date)}) is ${dueStatusText(b)}.\n\nKindly arrange payment at the earliest.\n\nRegards\nNavkar Fabrics`;

  return(<div>
    {backupOverdue&&<div style={{background:"#FEF6E7",border:"1px solid #F5D397",borderRadius:14,padding:"13px 15px",marginBottom:16,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
      <span style={{fontSize:20}}>⚠️</span>
      <div style={{flex:1,minWidth:160}}>
        <div style={{fontWeight:800,fontSize:12.5,color:"#7D5A00"}}>{daysSinceBackup===null?"No backup taken yet":`Last backup: ${daysSinceBackup} day${daysSinceBackup===1?"":"s"} ago`}</div>
        <div style={{fontSize:11,color:"#8A6D1F",marginTop:1}}>Share a backup copy to WhatsApp or Email to keep your data safe.</div>
      </div>
      <button onClick={onShareBackup} style={{background:"#7D5A00",color:"#fff",border:"none",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flexShrink:0}}>📤 Share Now</button>
    </div>}

    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>📅 Payment Due Reminders</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:urgentList.length>0?14:0}}>
        <div style={{background:"#FDEDEC",borderRadius:10,padding:"9px 6px",textAlign:"center"}}><div style={{fontSize:17,fontWeight:900,color:"#922B21"}}>{overdueList.length}</div><div style={{fontSize:9,color:"#922B21",fontWeight:700,textTransform:"uppercase"}}>Overdue</div></div>
        <div style={{background:"#FDEBD0",borderRadius:10,padding:"9px 6px",textAlign:"center"}}><div style={{fontSize:17,fontWeight:900,color:C.orange}}>{dueTodayList.length}</div><div style={{fontSize:9,color:C.orange,fontWeight:700,textTransform:"uppercase"}}>Today</div></div>
        <div style={{background:"#FEF9E7",borderRadius:10,padding:"9px 6px",textAlign:"center"}}><div style={{fontSize:17,fontWeight:900,color:"#9A7B1E"}}>{dueTomorrowList.length}</div><div style={{fontSize:9,color:"#9A7B1E",fontWeight:700,textTransform:"uppercase"}}>Tomorrow</div></div>
        <div style={{background:"#EAF4FC",borderRadius:10,padding:"9px 6px",textAlign:"center"}}><div style={{fontSize:17,fontWeight:900,color:C.blue}}>{dueWeekList.length}</div><div style={{fontSize:9,color:C.blue,fontWeight:700,textTransform:"uppercase"}}>This Week</div></div>
      </div>
      {urgentList.length===0&&<div style={{textAlign:"center",color:C.muted,fontSize:12.5,padding:"6px 0"}}>✅ No urgent dues — nothing overdue or due in the next 2 days.</div>}
      {urgentList.map((b,i)=>{
        const phone=phoneMap[b.customerName]||"";
        const col=b.daysUntilDue<0?"#922B21":b.daysUntilDue===0?C.orange:"#9A7B1E";
        return(<div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"9px 0",borderTop:i>0?`1px solid ${C.border}`:"none"}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{b.customerName}</div>
            <div style={{fontSize:10.5,color:col,fontWeight:600}}>{b.billNo} · ₹{fmt(b.amount)} · {dueStatusText(b)}</div>
          </div>
          <button onClick={()=>{const n=phone?"91"+String(phone).replace(/\D/g,"").slice(-10):"";window.open(`https://wa.me/${n}?text=${encodeURIComponent(buildDueWA(b))}`,"_blank");}} style={{background:"#E8FBF0",color:"#25D366",border:"1px solid #25D366",borderRadius:8,padding:"6px 10px",fontSize:11,fontWeight:700,cursor:"pointer",flexShrink:0}}>💬</button>
        </div>);
      })}
    </div>

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
      {[{icon:"📈",t:"Analytics",l:"Analytics",bg:C.blue},{icon:"⏳",t:"Outstanding",l:"Outstanding",bg:C.red},{icon:"📅",t:"Aging",l:"Ageing",bg:C.orange},{icon:"📋",t:"Reports",l:"Reports",bg:C.purple}].map(q=>(
        <button key={q.t} onClick={()=>onNav(q.t)} style={{background:q.bg,color:"#fff",border:"none",borderRadius:13,padding:"14px 10px",fontSize:13,fontWeight:700,cursor:"pointer",textAlign:"center"}}>
          <div style={{fontSize:22,marginBottom:4}}>{q.icon}</div>{q.l}
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

// ─── ENTRY TAB (Sale, Payment, Debit Note, Credit/Return) ─────────
function TradingTab({data,onAdd,onAddPay,onAddDebit,onAddCredit,onDel,onEdit,initialView,initialSearch}){
  const[view,setView]=useState(initialView||"sales");const[search,setSearch]=useState(initialSearch||"");
  const q=search.trim().toLowerCase();
  const matchName=(name)=>!q||(name||"").toLowerCase().includes(q);

  const sales=[...data.tradingSales].filter(s=>!q||[s.customerName,s.productName,s.billNo].some(x=>x?.toLowerCase().includes(q))).sort((a,b)=>new Date(b.date)-new Date(a.date));
  const payments=[...data.tradingPayments].filter(p=>matchName(p.customerName)).sort((a,b)=>new Date(b.date)-new Date(a.date));
  const debitNotes=(data.debitNotes||[]).filter(n=>matchName(n.customerName)).sort((a,b)=>new Date(b.date)-new Date(a.date));
  const creditNotes=(data.creditNotes||[]).filter(n=>matchName(n.customerName)).sort((a,b)=>new Date(b.date)-new Date(a.date));

  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
      <Btn color={C.blue} onClick={onAdd}>+ Sale Entry</Btn>
      <Btn color={C.green} onClick={onAddPay}>+ Payment</Btn>
      <Btn color={C.red} onClick={onAddDebit}>+ Debit Note</Btn>
      <Btn color={C.green} onClick={onAddCredit}>+ Goods Return / Credit Note</Btn>
    </div>
    <input placeholder="🔍 Search customer (Sale, Payment, Debit, Credit)…" value={search} onChange={e=>setSearch(e.target.value)} style={{...IS,marginBottom:10}}/>
    <SegCtrl options={[{v:"sales",l:`Sales (${sales.length})`},{v:"payments",l:`Payments (${payments.length})`},{v:"debit",l:`Debit (${debitNotes.length})`},{v:"credit",l:`Credit (${creditNotes.length})`}]} val={view} onChange={setView}/>
    <div style={{marginTop:10}}>
    {view==="sales"&&<>
      {sales.length===0&&<Empty text="No sales found."/>}
      {sales.map(s=>(<div key={s.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{s.customerName}</B><span style={{fontWeight:900,color:C.blue,fontSize:14}}>₹{fmt(s.amount)}</span></Row>
        {s.billNo&&<Mute>Bill No: {s.billNo}</Mute>}
        <Mute>{s.productName} · {s.supplierName}</Mute>
        <Mute>{fmt(s.meters)} m @ ₹{fmt(s.rate)}/m · {fmtD(s.date)}</Mute>
        <div style={{display:"flex",gap:8,marginTop:8,flexWrap:"wrap"}}>
          <button onClick={()=>generateInvoice(s,data.tradingSales,data.customers.find(c=>c.name===s.customerName))} style={{background:"#FEF6E7",color:"#9A7B1E",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🧾 Invoice</button>
          <button onClick={()=>onEdit("sale",s)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✏️ Edit</button>
          <button onClick={()=>onDel("tradingSales",s.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️ Delete</button>
        </div>
      </div>))}
    </>}
    {view==="payments"&&<>
      {payments.length===0&&<Empty text="No payments found."/>}
      {payments.map(p=>(<div key={p.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{p.customerName}</B><span style={{fontWeight:900,color:C.green}}>₹{fmt(p.amount)}</span></Row>
        {p.billNo&&<Mute>Bill No: {p.billNo}</Mute>}
        <Mute>{p.mode} · {fmtD(p.date)}</Mute>
        {p.remarks&&<Mute>{p.remarks}</Mute>}
        <div style={{display:"flex",gap:8,marginTop:8}}>
          <button onClick={()=>onEdit("payment",p)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✏️ Edit</button>
          <button onClick={()=>onDel("tradingPayments",p.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️ Delete</button>
        </div>
      </div>))}
    </>}
    {view==="debit"&&<>
      {debitNotes.length===0&&<Empty text="No debit notes found."/>}
      {debitNotes.map(n=>(<div key={n.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",borderLeft:`4px solid ${C.red}`}}>
        <Row><div><B>{n.customerName}</B>{n.noteNo&&<Mute>Note No: {n.noteNo}</Mute>}</div><span style={{fontWeight:900,color:C.red,fontSize:15}}>₹{fmt(n.amount)}</span></Row>
        {n.originalBillNo&&<Mute>Against Bill: {n.originalBillNo}</Mute>}
        <Mute>{fmtD(n.date)}</Mute>
        {n.remarks&&<Mute>{n.remarks}</Mute>}
        <div style={{background:"#FDEDEC",borderRadius:8,padding:"7px 10px",marginTop:8,fontSize:11.5,color:"#C0392B",fontWeight:600}}>Increases customer outstanding</div>
        <div style={{display:"flex",gap:8,marginTop:8}}>
          <button onClick={()=>onEdit("debit",n)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✏️ Edit</button>
          <button onClick={()=>onDel("debitNotes",n.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️ Delete</button>
        </div>
      </div>))}
    </>}
    {view==="credit"&&<>
      {creditNotes.length===0&&<Empty text="No credit notes or sale returns found."/>}
      {creditNotes.map(n=>(<div key={n.id} style={{background:C.card,borderRadius:13,padding:"14px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",borderLeft:`4px solid ${C.green}`}}>
        <Row><div><B>{n.customerName}</B>{n.noteNo&&<Mute>Note No: {n.noteNo}</Mute>}</div><span style={{fontWeight:900,color:C.green,fontSize:15}}>₹{fmt(n.amount)}</span></Row>
        {n.originalBillNo&&<Mute>Against Bill: {n.originalBillNo}</Mute>}
        {n.productName&&<Mute>{n.productName}{n.meters?` · ${fmt(n.meters)} m`:""}</Mute>}
        <Mute>{fmtD(n.date)}</Mute>
        {n.remarks&&<Mute>{n.remarks}</Mute>}
        <div style={{background:"#E8F8F5",borderRadius:8,padding:"7px 10px",marginTop:8,fontSize:11.5,color:C.teal,fontWeight:600}}>Reduces customer outstanding</div>
        <div style={{display:"flex",gap:8,marginTop:8}}>
          <button onClick={()=>onEdit("credit",n)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✏️ Edit</button>
          <button onClick={()=>onDel("creditNotes",n.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️ Delete</button>
        </div>
      </div>))}
    </>}
    </div>
  </div>);
}

// ─── OUTSTANDING TAB ─────────────────────────────────────────────
function OutstandingTab({tradingOut,data,onAddPay,generatePDF}){
  const[search,setSearch]=useState("");
  const[expanded,setExpanded]=useState(null);
  const[bulkOpen,setBulkOpen]=useState(false);
  const todayMs=new Date().setHours(0,0,0,0);
  const phoneMap={};const creditDaysMap={};const creditLimitMap={};
  data.customers.forEach(c=>{phoneMap[c.name]=c.phone;creditDaysMap[c.name]=parseInt(c.creditDays||30);if(c.creditLimit&&+c.creditLimit>0)creditLimitMap[c.name]=+c.creditLimit;});

  const ageBucket=(days)=>{if(days<=30)return{label:"0-30 Days",color:C.green};if(days<=60)return{label:"31-60 Days",color:C.orange};if(days<=90)return{label:"61-90 Days",color:"#E67E22"};if(days<=120)return{label:"91-120 Days",color:C.red};return{label:"Above 120 Days",color:"#922B21"};};

  const entries=Object.values(tradingOut).map(v=>{
    const net=Math.max(0,v.due+(v.debit||0)-v.paid-(v.credit||0));
    const custSales=data.tradingSales.filter(s=>s.customerName===v.name).sort((a,b)=>new Date(a.date)-new Date(b.date));
    const custDebits=(data.debitNotes||[]).filter(n=>n.customerName===v.name);
    // Credit notes reduce outstanding same as payment — include in FIFO pool
    let rem=v.paid+(v.credit||0);let maxDays=0;let oldestBill=null;let oldestDate=null;
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
    // Debit notes always add to outstanding (not offset by the paid/credit pool) — age from their own date, matches Ageing tab
    custDebits.forEach(n=>{
      const amt=+n.amount||0;if(amt<=0)return;
      const days=Math.floor((todayMs-new Date(n.date).setHours(0,0,0,0))/86400000);
      if(days>maxDays){maxDays=days;oldestBill=n.noteNo||"DN";oldestDate=n.date;}
      openBills.push({billNo:n.noteNo||"DN",date:n.date,outstanding:amt,days,billAmount:amt,productName:"Debit Note",meters:0,rate:0,isDebitNote:true});
    });
    const limit=creditLimitMap[v.name]||0;
    return{...v,net,maxDays,oldestBill,oldestDate,openBills,creditLimit:limit,overLimit:limit>0&&net>limit};
  }).filter(v=>v.net>0).filter(v=>!search||v.name.toLowerCase().includes(search.toLowerCase())).sort((a,b)=>b.net-a.net);

  const total=entries.reduce((a,v)=>a+v.net,0);
  const overLimitCount=entries.filter(v=>v.overLimit).length;

  const buildWA=(v)=>{
    const nums=["1️⃣","2️⃣","3️⃣","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"];
    const fmtBillDate=(d)=>{const dt=new Date(d);const m=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];return `${String(dt.getDate()).padStart(2,"0")}-${m[dt.getMonth()]}-${String(dt.getFullYear()).slice(2)}`;};
    const fmtAmt=(n)=>Number(Math.round(n)).toLocaleString("en-IN");
    const billLines=v.openBills.map((b,i)=>{
      const billNoClean=String(b.billNo).replace(/IMP-/i,"");
      return `${nums[i]||`${i+1}.`} Bill No: ${billNoClean}\n   Date: ${fmtBillDate(b.date)}\n   O/S: *₹${fmtAmt(b.outstanding)}*\n   Age: ${b.days} Days`;
    }).join("\n\n");
    return `🏢 *Navkar Fabrics*\n\nDear *${v.name}*,\n\nOutstanding payment reminder:\n\n${billLines}\n\n💰 *Total O/S: ₹${fmtAmt(v.net)}*\n\nKindly ignore if already paid.\n\nRegards\nNavkar Fabrics`;
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
    <button onClick={()=>setBulkOpen(true)} disabled={entries.length===0} style={{width:"100%",background:entries.length===0?"#EAF6EF":"#E8FBF0",color:"#1E9E58",border:"1px solid #25D366",borderRadius:11,padding:"11px 14px",fontSize:13,fontWeight:700,cursor:entries.length===0?"default":"pointer",marginBottom:14,opacity:entries.length===0?0.5:1}}>📤 Bulk WhatsApp Reminders ({entries.length} customers)</button>
    <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:14,padding:"14px 16px",marginBottom:16}}>
      <div style={{fontSize:10.5,color:C.gold,textTransform:"uppercase",letterSpacing:1,fontWeight:600}}>Total Outstanding ({entries.length} customers)</div>
      <div style={{fontSize:24,fontWeight:900,color:"#fff"}}>₹{fmt(total)}</div>
    </div>
    {overLimitCount>0&&<div style={{background:"#FDEDEC",border:"1px solid #F5B7B1",borderRadius:12,padding:"11px 14px",marginBottom:14,fontSize:12.5,color:"#922B21",fontWeight:700}}>🚫 {overLimitCount} customer{overLimitCount===1?"":"s"} over their credit limit</div>}
    {entries.length===0&&<Empty text="All payments clear. No outstanding."/>}
    {entries.map(v=>{const bucket=ageBucket(v.maxDays);const phone=phoneMap[v.name]||v.phone||"";const isOpen=expanded===v.name;const creditDays=creditDaysMap[v.name]||30;return(
      <div key={v.name} style={{background:C.card,borderRadius:14,padding:"14px 15px",marginBottom:12,boxShadow:"0 1px 8px rgba(0,0,0,0.07)",borderLeft:`4px solid ${v.overLimit?"#922B21":bucket.color}`}}>
        <Row><div><B style={{fontSize:15}}>{v.name}</B>{phone&&<Mute>{phone}</Mute>}</div><div style={{textAlign:"right"}}><div style={{fontWeight:900,fontSize:18,color:C.red}}>₹{fmt(v.net)}</div><span style={{fontSize:10,color:bucket.color,fontWeight:700}}>{bucket.label}</span></div></Row>
        {v.overLimit&&<div style={{background:"#FDEDEC",borderRadius:8,padding:"6px 10px",marginTop:8,fontSize:11,fontWeight:700,color:"#922B21"}}>🚫 Over credit limit — ₹{fmt(v.creditLimit)} limit, over by ₹{fmt(v.net-v.creditLimit)}</div>}
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
          {v.openBills.map((b,i)=>{const bb=ageBucket(b.days);const credit=Math.max(0,(b.billAmount||b.outstanding)-b.outstanding);return(<div key={i} style={{background:b.isDebitNote?"#FDEDEC":"#F7F9FC",borderRadius:10,padding:"12px 13px",marginBottom:8,borderLeft:`3px solid ${b.isDebitNote?C.red:bb.color}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:800,color:C.navy}}>{b.isDebitNote?"Debit Note No. : ":"Bill No. : "}{b.billNo}</div>
                <div style={{fontSize:11.5,color:C.muted,marginTop:3}}>{b.isDebitNote?"Note Date : ":"Bill Date : "}{fmtD(b.date)}</div>
                {b.isDebitNote?
                  <div style={{fontSize:11,fontWeight:700,color:"#C0392B",marginTop:3}}>⚠️ Debit Note — increases outstanding</div>
                :<>
                  {b.productName&&<div style={{fontSize:11.5,color:C.teal,marginTop:2}}>📦 {b.productName}</div>}
                  {b.meters>0&&<div style={{fontSize:11.5,color:C.muted}}>Qty : {fmt(b.meters)} m {b.rate>0?`@ ₹${fmt(b.rate)}/m`:""}</div>}
                  <div style={{fontSize:11.5,color:C.muted,marginTop:2}}>Bill Amt : <b style={{color:C.blue}}>₹{fmt(b.billAmount||b.outstanding)}</b></div>
                  {credit>0&&<div style={{fontSize:11.5,color:C.muted}}>Credit : <b style={{color:C.green}}>₹{fmt(credit)}</b></div>}
                </>}
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
    {bulkOpen&&<BulkReminderModal entries={entries} phoneMap={phoneMap} buildWA={buildWA} onClose={()=>setBulkOpen(false)}/>}
  </div>);
}

// ─── BULK WHATSAPP REMINDERS (guided send queue) ───────────────────
// Browsers block programmatic multi-window opens, so this walks the user through
// one WhatsApp share per customer with a tap to advance — as close to "bulk" as a
// client-only app can safely get without a WhatsApp Business API backend.
function BulkReminderModal({entries,phoneMap,buildWA,onClose}){
  const withPhone=entries.map(v=>({...v,phone:phoneMap[v.name]||v.phone||""})).filter(v=>v.phone);
  const withoutPhone=entries.filter(v=>!(phoneMap[v.name]||v.phone||""));
  const[selected,setSelected]=useState(()=>new Set(withPhone.map(v=>v.name)));
  const[queue,setQueue]=useState(null); // null=selection step, else array of names being sent
  const[idx,setIdx]=useState(0);
  const[sentCount,setSentCount]=useState(0);

  const toggle=(name)=>setSelected(p=>{const n=new Set(p);n.has(name)?n.delete(name):n.add(name);return n;});
  const selectAll=()=>setSelected(new Set(withPhone.map(v=>v.name)));
  const deselectAll=()=>setSelected(new Set());

  const startQueue=()=>{
    const list=withPhone.filter(v=>selected.has(v.name));
    if(list.length===0)return;
    setQueue(list);setIdx(0);setSentCount(0);
  };

  const current=queue?queue[idx]:null;
  const openWA=(v)=>{
    const n="91"+String(v.phone).replace(/\D/g,"").slice(-10);
    window.open(`https://wa.me/${n}?text=${encodeURIComponent(buildWA(v))}`,"_blank");
  };
  const markSentAndNext=()=>{setSentCount(c=>c+1);setIdx(i=>i+1);};
  const skip=()=>setIdx(i=>i+1);

  const isDone=queue&&idx>=queue.length;

  return(<ModalBase title="📤 Bulk WhatsApp Reminders" onClose={onClose}>
    {!queue&&<>
      <div style={{background:"#FEF9E7",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
        Phones don't allow apps to auto-send WhatsApp messages to multiple people at once — so this opens WhatsApp for each selected customer one by one. Tap "Send & Next" after each to move on.
      </div>
      {withoutPhone.length>0&&<div style={{background:"#FDEDEC",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12,color:"#922B21"}}>⚠️ {withoutPhone.length} customer(s) have no phone number saved and will be skipped: {withoutPhone.map(v=>v.name).join(", ")}</div>}
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <button onClick={selectAll} style={{flex:1,background:"#F0F4F8",color:C.navy,border:"none",borderRadius:9,padding:"9px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Select All</button>
        <button onClick={deselectAll} style={{flex:1,background:"#F0F4F8",color:C.navy,border:"none",borderRadius:9,padding:"9px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Deselect All</button>
      </div>
      <div style={{maxHeight:340,overflowY:"auto",marginBottom:14}}>
        {withPhone.map(v=>(
          <label key={v.name} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:"#F8FAFC",borderRadius:10,marginBottom:8,cursor:"pointer"}}>
            <input type="checkbox" checked={selected.has(v.name)} onChange={()=>toggle(v.name)} style={{width:18,height:18,flexShrink:0}}/>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:13,fontWeight:700,color:C.navy}}>{v.name}</div>
              <div style={{fontSize:11,color:C.muted}}>{v.phone} · {v.maxDays}d overdue</div>
            </div>
            <div style={{fontSize:13,fontWeight:900,color:C.red,flexShrink:0}}>₹{fmt(v.net)}</div>
          </label>
        ))}
        {withPhone.length===0&&<Empty text="No customers with saved phone numbers."/>}
      </div>
      <SaveBtn color={selected.size===0?C.muted:"#1E9E58"} onClick={startQueue}>💬 Start Sending ({selected.size})</SaveBtn>
    </>}

    {queue&&!isDone&&<>
      <div style={{textAlign:"center",fontSize:12,color:C.muted,fontWeight:700,marginBottom:14}}>Sending {idx+1} of {queue.length}</div>
      <div style={{background:"#F8FAFC",borderRadius:14,padding:"18px",textAlign:"center",marginBottom:16}}>
        <div style={{fontSize:17,fontWeight:800,color:C.navy}}>{current.name}</div>
        <div style={{fontSize:12,color:C.muted,marginTop:2}}>{current.phone} · {current.maxDays}d overdue</div>
        <div style={{fontSize:26,fontWeight:900,color:C.red,marginTop:10}}>₹{fmt(current.net)}</div>
      </div>
      <button onClick={()=>openWA(current)} style={{width:"100%",background:"#25D366",color:"#fff",border:"none",borderRadius:12,padding:15,fontSize:14.5,fontWeight:800,cursor:"pointer",marginBottom:10,minHeight:50}}>💬 Open WhatsApp</button>
      <div style={{display:"flex",gap:8}}>
        <button onClick={skip} style={{flex:1,background:"#F0F4F8",color:C.navy,border:"none",borderRadius:12,padding:14,fontSize:13.5,fontWeight:700,cursor:"pointer",minHeight:48}}>⏭ Skip</button>
        <button onClick={markSentAndNext} style={{flex:2,background:C.navy,color:C.gold,border:"none",borderRadius:12,padding:14,fontSize:13.5,fontWeight:800,cursor:"pointer",minHeight:48}}>✓ Sent — Next</button>
      </div>
    </>}

    {isDone&&<div style={{textAlign:"center",padding:"20px 10px"}}>
      <div style={{fontSize:40,marginBottom:10}}>✅</div>
      <div style={{fontSize:16,fontWeight:800,color:C.navy}}>Done!</div>
      <div style={{fontSize:13,color:C.muted,marginTop:4}}>Sent to {sentCount} of {queue.length} customers.</div>
      <SaveBtn color={C.navy} onClick={onClose}>Close</SaveBtn>
    </div>}
  </ModalBase>);
}

// ─── LEDGER TAB (Customer Statement with running balance) ─────────
function LedgerTab({data,generatePDF,initialCustomer}){
  const[selectedCustomer,setSelectedCustomer]=useState(initialCustomer||"");
  const[search,setSearch]=useState("");

  // All known customer names — Masters list plus any name that only appears in transactions
  const allNames=useMemo(()=>{
    const set=new Set(data.customers.map(c=>c.name));
    data.tradingSales.forEach(s=>s.customerName&&set.add(s.customerName));
    data.tradingPayments.forEach(p=>p.customerName&&set.add(p.customerName));
    (data.debitNotes||[]).forEach(n=>n.customerName&&set.add(n.customerName));
    (data.creditNotes||[]).forEach(n=>n.customerName&&set.add(n.customerName));
    return [...set].sort((a,b)=>a.localeCompare(b));
  },[data]);

  const q=search.trim().toLowerCase();
  const filteredNames=q?allNames.filter(n=>n.toLowerCase().includes(q)):[];

  const custMaster=data.customers.find(c=>c.name===selectedCustomer);

  const entries=useMemo(()=>{
    if(!selectedCustomer)return[];
    const rows=[];
    data.tradingSales.filter(s=>s.customerName===selectedCustomer).forEach(s=>rows.push({date:s.date,type:"Sale",ref:s.billNo||"—",particulars:s.productName||"",debit:+s.amount||0,credit:0,ord:0}));
    data.tradingPayments.filter(p=>p.customerName===selectedCustomer).forEach(p=>rows.push({date:p.date,type:"Payment",ref:p.billNo||p.mode||"—",particulars:p.remarks||p.mode||"",debit:0,credit:+p.amount||0,ord:1}));
    (data.debitNotes||[]).filter(n=>n.customerName===selectedCustomer).forEach(n=>rows.push({date:n.date,type:"Debit Note",ref:n.noteNo||"DN",particulars:n.remarks||"",debit:+n.amount||0,credit:0,ord:2}));
    (data.creditNotes||[]).filter(n=>n.customerName===selectedCustomer).forEach(n=>rows.push({date:n.date,type:"Credit Note",ref:n.noteNo||"CN",particulars:n.remarks||n.productName||"",debit:0,credit:+n.amount||0,ord:3}));
    rows.sort((a,b)=>new Date(a.date)-new Date(b.date)||a.ord-b.ord);
    let bal=0;
    // Running balance is always computed over the FULL history first (so it's always
    // correct), then the Financial Year filter below only changes which rows are shown.
    return rows.map(r=>{bal+=r.debit-r.credit;return{...r,balance:bal};});
  },[data,selectedCustomer]);

  const[fy,setFy]=useState(ALL_FY);
  const availableFYs=useMemo(()=>getAvailableFYs(entries),[entries]);
  const filteredEntries=useMemo(()=>fy===ALL_FY?entries:filterByFY(entries,fy),[entries,fy]);

  const totals=useMemo(()=>filteredEntries.reduce((a,r)=>{
    if(r.type==="Sale")a.sales+=r.debit;
    if(r.type==="Payment")a.paid+=r.credit;
    if(r.type==="Debit Note")a.debit+=r.debit;
    if(r.type==="Credit Note")a.credit+=r.credit;
    return a;
  },{sales:0,paid:0,debit:0,credit:0}),[filteredEntries]);

  // "Closing balance" shown is always the TRUE running balance as of the last visible row
  // (accurate whether viewing All Time or a single FY), not just a sum of the filtered rows.
  const closingBalance=filteredEntries.length?filteredEntries[filteredEntries.length-1].balance:(entries.length?0:0);
  const typeColor={Sale:C.blue,Payment:C.green,"Debit Note":C.red,"Credit Note":C.teal};

  const doPDF=()=>{
    const rows=filteredEntries.map(r=>`<tr><td>${fmtD(r.date)}</td><td>${r.type}</td><td>${r.ref}</td><td>${r.particulars||"—"}</td><td style="text-align:right">${r.debit?"₹"+fmt(r.debit):""}</td><td style="text-align:right">${r.credit?"₹"+fmt(r.credit):""}</td><td style="text-align:right;font-weight:bold">₹${fmt(r.balance)}</td></tr>`).join("");
    generatePDF(`Statement of Account — ${selectedCustomer}${fy!==ALL_FY?" (FY "+fy+")":""}`,`<table><thead><tr><th>Date</th><th>Type</th><th>Ref</th><th>Particulars</th><th style="text-align:right">Debit</th><th style="text-align:right">Credit</th><th style="text-align:right">Balance</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="6">CLOSING BALANCE</td><td style="text-align:right">₹${fmt(closingBalance)}</td></tr></tbody></table>`);
  };

  const buildWA=()=>{
    const lines=filteredEntries.slice(-10).map(r=>`${fmtD(r.date)} — ${r.type}${r.ref&&r.ref!=="—"?" ("+r.ref+")":""}: ${r.debit?"+₹"+fmt(r.debit):"-₹"+fmt(r.credit)}`).join("\n");
    return `🏢 *Navkar Fabrics*\n\nDear *${selectedCustomer}*,\n\nStatement of Account (last ${Math.min(10,filteredEntries.length)} entries):\n\n${lines}\n\n💰 *Closing Balance: ₹${fmt(closingBalance)}*\n\nRegards\nNavkar Fabrics`;
  };

  // ── Customer picker screen — search-first, nothing listed until you type ──
  if(!selectedCustomer){
    return(<div>
      <SecTitle>Customer Ledger</SecTitle>
      <input placeholder="🔍 Search customer to view statement…" value={search} onChange={e=>setSearch(e.target.value)} style={{...IS,marginBottom:12}}/>
      {!q&&<Empty text={`Search above to open a customer's statement (${allNames.length} customers).`}/>}
      {q&&filteredNames.length===0&&<Empty text="No customers found."/>}
      {filteredNames.map(name=>{
        const m=data.customers.find(c=>c.name===name);
        return(<button key={name} onClick={()=>setSelectedCustomer(name)} style={{display:"block",width:"100%",textAlign:"left",background:C.card,border:"none",borderRadius:13,padding:"14px 15px",marginBottom:9,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",cursor:"pointer"}}>
          <Row><B style={{fontSize:14}}>{name}</B><span style={{fontSize:16,color:C.muted}}>›</span></Row>
          {(m?.phone||m?.city)&&<Mute>{[m?.phone,m?.city].filter(Boolean).join(" · ")}</Mute>}
        </button>);
      })}
    </div>);
  }

  // ── Statement view ──
  return(<div>
    <button onClick={()=>{setSelectedCustomer("");setSearch("");}} style={{background:"#F0F4F8",color:C.navy,border:"none",borderRadius:9,padding:"9px 14px",fontSize:12.5,fontWeight:700,cursor:"pointer",marginBottom:12}}>← Change Customer</button>

    <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:16,padding:"16px 18px",marginBottom:14,border:"1px solid rgba(232,201,126,0.25)"}}>
      <div style={{fontSize:10,color:C.gold,letterSpacing:1.5,textTransform:"uppercase",fontWeight:700}}>Statement of Account</div>
      <div style={{fontSize:18,fontWeight:900,color:"#fff",marginTop:4}}>{selectedCustomer}</div>
      {(custMaster?.phone||custMaster?.city)&&<div style={{fontSize:11.5,color:"rgba(255,255,255,0.55)",marginTop:2}}>{[custMaster?.phone,custMaster?.city].filter(Boolean).join(" · ")}</div>}
      <div style={{fontSize:11,color:"rgba(255,255,255,0.45)",marginTop:8}}>Closing Balance</div>
      <div style={{fontSize:28,fontWeight:900,color:closingBalance>0?C.red:C.gold,marginTop:2}}>₹{fmt(closingBalance)}</div>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
      <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.blue}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>SALES</div><div style={{fontSize:13,fontWeight:900,color:C.blue}}>₹{fmt(totals.sales)}</div></div>
      <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.green}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>PAID</div><div style={{fontSize:13,fontWeight:900,color:C.green}}>₹{fmt(totals.paid)}</div></div>
      <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.red}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>DEBIT NOTES</div><div style={{fontSize:13,fontWeight:900,color:C.red}}>₹{fmt(totals.debit)}</div></div>
      <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.teal}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>CREDIT NOTES</div><div style={{fontSize:13,fontWeight:900,color:C.teal}}>₹{fmt(totals.credit)}</div></div>
    </div>

    {/* Date-wise filter — narrow the statement down to one financial year, running balance stays accurate */}
    <div style={{marginBottom:14}}>
      <div style={{fontSize:10.5,color:C.muted,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>Filter by Date (Financial Year)</div>
      <div style={{display:"flex",overflowX:"auto",gap:6,scrollbarWidth:"none",paddingBottom:2}}>
        {availableFYs.map(f=>(
          <button key={f} onClick={()=>setFy(f)} style={{flex:"0 0 auto",padding:"8px 14px",borderRadius:20,fontSize:12,fontWeight:fy===f?700:500,border:`1.5px solid ${fy===f?C.navy:C.border}`,background:fy===f?C.navy:"#fff",color:fy===f?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>
            {f===ALL_FY?"📊 All Time":`FY ${f}`}
          </button>
        ))}
      </div>
    </div>

    <div style={{display:"flex",gap:8,marginBottom:14}}>
      <button onClick={()=>exportCSV([["Date","Type","Ref","Particulars","Debit","Credit","Balance"],...filteredEntries.map(r=>[fmtD(r.date),r.type,r.ref,r.particulars,r.debit||"",r.credit||"",r.balance])],`Ledger_${selectedCustomer}_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>📥 CSV</button>
      <button onClick={doPDF} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>🖨️ PDF</button>
      <button onClick={()=>{const phone=custMaster?.phone||"";const n=phone?"91"+String(phone).replace(/\D/g,"").slice(-10):"";window.open(`https://wa.me/${n}?text=${encodeURIComponent(buildWA())}`,"_blank");}} style={{background:"#E8FBF0",color:"#25D366",border:"1px solid #25D366",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40,flex:1}}>💬 WhatsApp</button>
    </div>

    {filteredEntries.length===0?<Empty text={fy!==ALL_FY?`No transactions in FY ${fy}.`:"No transactions found for this customer."}/>:(()=>{
      const th={padding:"9px 10px",textAlign:"left",fontSize:10.5,color:C.gold,background:C.navy,whiteSpace:"nowrap"};
      const td={padding:"8px 10px",fontSize:12,color:C.navy,whiteSpace:"nowrap",borderBottom:`1px solid ${C.border}`};
      return(
        <div style={{background:C.card,borderRadius:12,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",overflowX:"auto",marginBottom:10}}>
          <table style={{borderCollapse:"collapse",width:"100%",minWidth:640}}>
            <thead><tr>
              <th style={{...th,borderRadius:"12px 0 0 0"}}>Date</th>
              <th style={th}>Type</th>
              <th style={th}>Ref</th>
              <th style={{...th,textAlign:"right"}}>Debit</th>
              <th style={{...th,textAlign:"right"}}>Credit</th>
              <th style={{...th,textAlign:"right",borderRadius:"0 12px 0 0"}}>Balance</th>
            </tr></thead>
            <tbody>
              {filteredEntries.map((r,i)=>(
                <tr key={i} style={{background:i%2===0?"#fff":"#FAFBFC"}}>
                  <td style={td}>{fmtD(r.date)}</td>
                  <td style={{...td,fontWeight:700,color:typeColor[r.type]}}>{r.type}</td>
                  <td style={td}>{r.ref}</td>
                  <td style={{...td,textAlign:"right",color:C.blue,fontWeight:700}}>{r.debit?`₹${fmt(r.debit)}`:"—"}</td>
                  <td style={{...td,textAlign:"right",color:C.green,fontWeight:700}}>{r.credit?`₹${fmt(r.credit)}`:"—"}</td>
                  <td style={{...td,textAlign:"right",fontWeight:900,color:r.balance>0?C.red:C.navy}}>₹{fmt(r.balance)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    })()}
  </div>);
}
function AgingTab({data,generatePDF}){
  const[search,setSearch]=useState("");const[expanded,setExpanded]=useState(null);
  const todayMs=new Date().setHours(0,0,0,0);
  const entries=useMemo(()=>{
    const map={};
    data.tradingSales.forEach(s=>{if(!map[s.customerName])map[s.customerName]={name:s.customerName,invoices:[]};map[s.customerName].invoices.push({date:s.date,amount:+s.amount||0,id:s.id,productName:s.productName,billNo:s.billNo});});
    const paid={};data.tradingPayments.forEach(p=>{paid[p.customerName]=(paid[p.customerName]||0)+(+p.amount||0);});
    const credit={};(data.creditNotes||[]).forEach(n=>{credit[n.customerName]=(credit[n.customerName]||0)+(+n.amount||0);});
    const debitByCust={};(data.debitNotes||[]).forEach(n=>{if(!debitByCust[n.customerName])debitByCust[n.customerName]=[];debitByCust[n.customerName].push(n);});
    // Ensure customers who only have debit notes (no sales) still show up
    Object.keys(debitByCust).forEach(name=>{if(!map[name])map[name]={name,invoices:[]};});
    return Object.values(map).map(v=>{
      // Credit notes reduce outstanding the same way payments do — pool them together (matches Outstanding tab)
      let rem=(paid[v.name]||0)+(credit[v.name]||0);
      const bk={b0:0,b30:0,b60:0,b90:0,b120:0};const openInvoices=[];
      [...v.invoices].sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach(inv=>{
        let amt=inv.amount;const d=Math.min(amt,rem);amt-=d;rem-=d;if(amt<=0)return;
        const days=Math.floor((todayMs-new Date(inv.date).setHours(0,0,0,0))/86400000);
        let bucket;if(days<=30){bk.b0+=amt;bucket="0-30 Days";}else if(days<=60){bk.b30+=amt;bucket="31-60 Days";}else if(days<=90){bk.b60+=amt;bucket="61-90 Days";}else if(days<=120){bk.b90+=amt;bucket="91-120 Days";}else{bk.b120+=amt;bucket="120+ Days";}
        openInvoices.push({...inv,outstanding:amt,days,bucket});
      });
      // Debit notes always increase outstanding (they are not offset by the paid/credit pool) — age from their own date
      (debitByCust[v.name]||[]).forEach(n=>{
        const amt=+n.amount||0;if(amt<=0)return;
        const days=Math.floor((todayMs-new Date(n.date).setHours(0,0,0,0))/86400000);
        let bucket;if(days<=30){bk.b0+=amt;bucket="0-30 Days";}else if(days<=60){bk.b30+=amt;bucket="31-60 Days";}else if(days<=90){bk.b60+=amt;bucket="61-90 Days";}else if(days<=120){bk.b90+=amt;bucket="91-120 Days";}else{bk.b120+=amt;bucket="120+ Days";}
        openInvoices.push({date:n.date,amount:amt,id:n.id,productName:"Debit Note",billNo:n.noteNo||"DN",outstanding:amt,days,bucket});
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
  const[chartType,setChartType]=useState("overview");
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

  const topCustomersTable=useMemo(()=>{
    const map={};
    filtered.forEach(s=>{if(!map[s.customerName])map[s.customerName]={name:s.customerName,sales:0,bills:0};map[s.customerName].sales+=+s.amount||0;map[s.customerName].bills+=1;});
    const total=Object.values(map).reduce((a,v)=>a+v.sales,0)||1;
    return Object.val