import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, AreaChart, Area, ComposedChart, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import { supabase } from './supabaseClient';
import Login from './Login';

const SK = "fabric-sales-v3-trading";
const EMPTY = { customers:[], suppliers:[], products:[], tradingSales:[], tradingPayments:[], debitNotes:[], creditNotes:[], enquiries:[], trash:[], materialPrices:[], companySettings:{name:"Navkar Fabrics",address:"",gstin:"",phone:"",email:"",state:"Delhi",bankName:"",bankAccount:"",bankIFSC:"",terms:"1. Goods once sold will not be taken back.\n2. Interest @ 18% p.a. will be charged if payment is not made within the stipulated time.\n3. Subject to jurisdiction only."}, appSettings:{defaultGstRate:"5",defaultCommissionRate:"1.5",defaultCreditDays:"30",backupReminderDays:"7",trashRetention:"20"} };
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
function numberToWordsIndian(num){
  num=Math.round(+num||0);
  if(num===0)return"Zero";
  const ones=["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"];
  const tens=["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"];
  const twoDigit=(n)=>n<20?ones[n]:tens[Math.floor(n/10)]+(n%10?" "+ones[n%10]:"");
  const threeDigit=(n)=>{
    if(n<100)return twoDigit(n);
    return ones[Math.floor(n/100)]+" Hundred"+(n%100?" "+twoDigit(n%100):"");
  };
  const parts=[];
  const crore=Math.floor(num/10000000);num%=10000000;
  const lakh=Math.floor(num/100000);num%=100000;
  const thousand=Math.floor(num/1000);num%=1000;
  const hundred=num;
  if(crore)parts.push(threeDigit(crore)+" Crore");
  if(lakh)parts.push(threeDigit(lakh)+" Lakh");
  if(thousand)parts.push(threeDigit(thousand)+" Thousand");
  if(hundred)parts.push(threeDigit(hundred));
  return parts.join(" ");
}
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
    if(left>0)open.push({billNo:s.billNo||"—",date:s.date,billAmount:amt,outstanding:left,meters:+s.meters||0,commissionRate:+s.commissionRate||1.5});
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
    if(rem>=b.outstanding){rem-=b.outstanding;billsCleared++;commissionGained+=b.meters*b.commissionRate;}
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
function calcGST(amount,buyerState,gstRatePct,sellerState){
  const seller=sellerState||SELLER_STATE;
  const rate=(+gstRatePct||5);
  const amt=+amount||0;
  const taxable=amt/(1+rate/100);
  const gstAmt=amt-taxable;
  const sameState=normName(buyerState||seller)===normName(seller);
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
function generateInvoice(sale,allSales,customer,companySettings){
  const co=companySettings||{};
  const sellerState=co.state||SELLER_STATE;
  const items=getInvoiceLineItems(sale,allSales);
  const buyerState=customer?.state||sellerState;

  // Each line item's `amount` is the TOTAL (product + freight, GST-inclusive) as stored.
  // Split it back into product-only vs freight, then back-calc taxable value for each.
  const lines=items.map(s=>{
    const freight=+s.freight||0;
    const productOnly=(+s.amount||0)-freight;
    const gp=calcGST(productOnly,buyerState,s.gstRate,sellerState);
    const gf=calcGST(freight,buyerState,s.gstRate,sellerState);
    const qty=+s.meters||0;
    const rateBeforeGST=qty>0?gp.taxable/qty:0;
    return{...s,freight,productOnly,gp,gf,qty,rateBeforeGST,lineTaxable:gp.taxable+gf.taxable,lineCgst:gp.cgst+gf.cgst,lineSgst:gp.sgst+gf.sgst,lineIgst:gp.igst+gf.igst};
  });

  const totalQty=lines.reduce((a,l)=>a+l.qty,0);
  const totalFreight=lines.reduce((a,l)=>a+l.gf.taxable,0);
  const totalProductTaxable=lines.reduce((a,l)=>a+l.gp.taxable,0);
  const totalTaxable=lines.reduce((a,l)=>a+l.lineTaxable,0);
  const totalCgst=lines.reduce((a,l)=>a+l.lineCgst,0);
  const totalSgst=lines.reduce((a,l)=>a+l.lineSgst,0);
  const totalIgst=lines.reduce((a,l)=>a+l.lineIgst,0);
  const totalTax=totalCgst+totalSgst+totalIgst;
  const preRoundTotal=totalTaxable+totalTax;
  const grandTotal=Math.round(preRoundTotal);
  const roundOff=grandTotal-preRoundTotal;
  const sameState=lines[0]?lines[0].gp.sameState:true;
  const rate=lines[0]?lines[0].gp.rate:5;
  const isExempt=rate===0;

  const rows=lines.map((l,i)=>`<tr>
      <td>${i+1}</td>
      <td>${l.productName||"—"}</td>
      <td>${l.hsn||"—"}</td>
      <td style="text-align:right">${fmt(l.qty)}</td>
      <td>${l.unit||"Mtr"}</td>
      <td style="text-align:right">₹${fmt(l.rateBeforeGST)}</td>
      <td style="text-align:right">₹${fmt(l.gp.taxable)}</td>
    </tr>`).join("");

  const gstRows=isExempt
    ?`<tr><td colspan="6" style="text-align:right">GST</td><td style="text-align:right">Exempt</td></tr>`
    :sameState
      ?`<tr><td colspan="6" style="text-align:right">CGST @ ${rate/2}%</td><td style="text-align:right">₹${fmt(totalCgst)}</td></tr>
        <tr><td colspan="6" style="text-align:right">SGST @ ${rate/2}%</td><td style="text-align:right">₹${fmt(totalSgst)}</td></tr>`
      :`<tr><td colspan="6" style="text-align:right">IGST @ ${rate}% (${buyerState})</td><td style="text-align:right">₹${fmt(totalIgst)}</td></tr>`;

  const roundRow=Math.abs(roundOff)>=0.005?`<tr class="sub-row"><td colspan="6" style="text-align:right">${roundOff>=0?"Add":"Less"}: Rounded Off</td><td style="text-align:right">${roundOff>=0?"+":"-"}₹${fmt(Math.abs(roundOff))}</td></tr>`:"";

  const amountWords=`Rupees ${numberToWordsIndian(grandTotal)} Only`;

  const win=window.open("","_blank");
  win.document.write(`<!DOCTYPE html><html><head><title>Invoice ${sale.billNo||sale.id}</title><style>
    @media print{body{padding:10mm;}}
    body{font-family:Arial,sans-serif;font-size:12.5px;color:#111;padding:26px;max-width:820px;margin:auto;}
    .gstin-bar{display:flex;justify-content:space-between;font-size:11px;font-weight:700;border-bottom:1px solid #333;padding-bottom:6px;margin-bottom:6px;}
    .hdr{text-align:center;border-bottom:2px solid #111;padding-bottom:10px;margin-bottom:10px;}
    .inv-label{font-size:12px;text-decoration:underline;font-weight:700;}
    .co{font-size:22px;font-weight:900;margin-top:2px;}
    .addr{font-size:11.5px;color:#333;margin-top:2px;}
    .meta-box{display:flex;border:1px solid #333;margin-bottom:8px;}
    .meta-col{flex:1;padding:8px 12px;font-size:11.5px;line-height:1.7;}
    .meta-col:first-child{border-right:1px solid #333;}
    .parties{display:flex;border:1px solid #333;border-top:none;margin-bottom:0;}
    .party-box{flex:1;padding:8px 12px;font-size:11.5px;line-height:1.6;}
    .party-box:first-child{border-right:1px solid #333;}
    .party-label{font-weight:800;font-style:italic;margin-bottom:3px;}
    table{width:100%;border-collapse:collapse;}
    th{background:#0F1923;color:#E8C97E;padding:7px 8px;text-align:left;font-size:10.5px;border:1px solid #333;}
    td{padding:6px 8px;border:1px solid #ccc;font-size:11.5px;}
    .sub-row td{border:1px solid #333;font-size:11.5px;}
    .tot-row td{background:#0F1923;color:#E8C97E;font-weight:bold;border:1px solid #333;font-size:12.5px;}
    .words-box{border:1px solid #333;border-top:none;padding:8px 12px;font-weight:700;font-size:12px;}
    .taxtable td,.taxtable th{font-size:10.5px;padding:5px 8px;}
    .bank-box{border:1px solid #333;border-top:none;padding:8px 12px;font-size:11.5px;}
    .footer{display:flex;border:1px solid #333;border-top:none;}
    .terms-box{flex:1;padding:8px 12px;font-size:10px;line-height:1.6;border-right:1px solid #333;}
    .sign-box{flex:1;padding:8px 12px;font-size:11px;text-align:right;display:flex;flex-direction:column;justify-content:space-between;}
  </style></head><body>
    <div class="gstin-bar"><span>GSTIN: ${co.gstin||"—"}</span><span>Original Copy</span></div>
    <div class="hdr">
      <div class="inv-label">TAX INVOICE</div>
      <div class="co">${co.name||"Navkar Fabrics"}</div>
      ${co.address?`<div class="addr">${co.address}</div>`:""}
      <div class="addr">${[co.phone?`Tel: ${co.phone}`:"",co.email?`Email: ${co.email}`:""].filter(Boolean).join("   ")}</div>
    </div>
    <div class="meta-box">
      <div class="meta-col"><b>Invoice No.</b> : ${sale.billNo||"—"}<br/><b>Dated</b> : ${fmtD(sale.date)}<br/><b>Place of Supply</b> : ${buyerState}</div>
      <div class="meta-col"><b>Reverse Charge</b> : N<br/><b>Transport</b> : By Road</div>
    </div>
    <div class="parties">
      <div class="party-box"><div class="party-label">Billed to :</div><b>${sale.customerName}</b><br/>
        ${customer&&customer.city?`${customer.city}, ${buyerState}<br/>`:`${buyerState}<br/>`}
        ${customer&&customer.phone?`Ph: ${customer.phone}<br/>`:""}
        ${customer&&customer.gstin?`GSTIN/UIN : ${customer.gstin}`:""}
      </div>
      <div class="party-box"><div class="party-label">Shipped to :</div><b>${sale.customerName}</b><br/>
        ${customer&&customer.city?`${customer.city}, ${buyerState}`:buyerState}
      </div>
    </div>
    <table><thead><tr>
      <th>S.N.</th><th>Description of Goods</th><th>HSN/SAC</th><th style="text-align:right">Qty</th><th>Unit</th><th style="text-align:right">Rate (Before GST)</th><th style="text-align:right">Amount</th>
    </tr></thead>
    <tbody>${rows}
      ${totalFreight>0?`<tr class="sub-row"><td colspan="6" style="text-align:right">Freight / Coolie Cartage</td><td style="text-align:right">₹${fmt(totalFreight)}</td></tr>`:""}
      <tr class="sub-row"><td colspan="6" style="text-align:right"><b>Taxable Amount</b></td><td style="text-align:right"><b>₹${fmt(totalTaxable)}</b></td></tr>
      ${gstRows}
      ${roundRow}
      <tr class="tot-row"><td colspan="6">Invoice Value (Qty: ${fmt(totalQty)})</td><td style="text-align:right">₹${fmt(grandTotal)}</td></tr>
    </tbody></table>
    <div class="words-box">${amountWords}</div>
    ${!isExempt?`<table class="taxtable"><thead><tr><th>Tax Rate</th><th style="text-align:right">Taxable Amt.</th>${sameState?`<th style="text-align:right">CGST</th><th style="text-align:right">SGST</th>`:`<th style="text-align:right">IGST</th>`}<th style="text-align:right">Total Tax</th></tr></thead>
      <tbody><tr><td>${rate}%</td><td style="text-align:right">₹${fmt(totalTaxable)}</td>${sameState?`<td style="text-align:right">₹${fmt(totalCgst)}</td><td style="text-align:right">₹${fmt(totalSgst)}</td>`:`<td style="text-align:right">₹${fmt(totalIgst)}</td>`}<td style="text-align:right">₹${fmt(totalTax)}</td></tr></tbody>
    </table>`:""}
    ${(co.bankName||co.bankAccount)?`<div class="bank-box"><b>Bank Details :</b> ${co.bankName||""}${co.bankIFSC?`, IFSC: ${co.bankIFSC}`:""}${co.bankAccount?`<br/>A/c No: ${co.bankAccount}`:""}</div>`:""}
    <div class="footer">
      <div class="terms-box"><b>Terms &amp; Conditions</b><br/>${(co.terms||"").split("\n").filter(Boolean).map(t=>t.replace(/^\d+\.\s*/,"")).map((t,i)=>`${i+1}. ${t}`).join("<br/>")||"E.&amp;O.E."}</div>
      <div class="sign-box"><div>Receiver's Signature :</div><div><b>for ${co.name||"Navkar Fabrics"}</b><br/><br/>Authorised Signatory</div></div>
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
  {id:"Analytics",icon:"📈",label:"Analytics"},
  {id:"Market",icon:"💹",label:"Market"},
  {id:"Enquiry",icon:"📝",label:"Enquiry"},
  {id:"Masters",icon:"⚙️",label:"Masters"},
  {id:"Reports",icon:"📋",label:"Reports"},
  {id:"Settings",icon:"🏢",label:"Settings"},
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
    const retention=parseInt(data.appSettings?.trashRetention)||20;
    setData(p=>({
      ...p,
      [section]:p[section].filter(r=>r.id!==id),
      trash:[{trashId,section,record,deletedAt:new Date().toISOString()},...(p.trash||[])].slice(0,retention)
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
  const updateCompanySettings=(updates)=>{
    setData(p=>({...p,companySettings:{...p.companySettings,...updates}}));
    showToast("Company settings saved.");
  };
  const updateAppSettings=(updates)=>{
    setData(p=>({...p,appSettings:{...p.appSettings,...updates}}));
    showToast("App settings saved.");
  };
  const addMaterialPrice=(entry)=>{
    setData(p=>({...p,materialPrices:[...(p.materialPrices||[]),{...entry,id:uid()}]}));
    showToast("Price entry saved.");
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
    if(r.kind==="customer")navigateTab("Reports",{repView:"ledger",customer:r.name});
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
      {TABS.map(t=>(<button key={t.id} onClick={()=>navigateTab(t.id)} style={{background:tab===t.id?"rgba(232,201,126,0.12)":"transparent",border:"none",borderLeft:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",padding:"12px 20px",textAlign:"left",color:tab===t.id?C.gold:"rgba(255,255,255,0.65)",fontSize:13.5,fontWeight:tab===t.id?700:400,cursor:"pointer",display:"flex",alignItems:"center",gap:10}}><span>{t.icon}</span>{t.label}{t.id==="Reports"&&overdueCount>0&&<span style={{background:C.red,color:"#fff",borderRadius:20,fontSize:10,fontWeight:800,padding:"2px 7px",marginLeft:"auto"}}>{overdueCount}</span>}</button>))}
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
          <button onClick={handleLogout} style={{...hdrBtn(),background:"rgba(231,76,60,0.18)"}}><span style={{fontSize:18}}>🚪</span><span style={{fontSize:9}}>Log Out</span></button>
        </div>
      </div>
      <div className="mob" style={{display:"none",background:C.navy,padding:"0 16px 10px"}}>
        <button onClick={()=>setSearchOpen(true)} style={{width:"100%",background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.18)",borderRadius:10,padding:"10px 12px",color:"rgba(255,255,255,0.65)",fontSize:13,textAlign:"left",cursor:"pointer"}}>🔍 Search everything…</button>
      </div>
      <div className="mob" style={{display:"none",background:C.navyMid,overflowX:"auto",WebkitOverflowScrolling:"touch",scrollbarWidth:"none",position:"sticky",top:72,zIndex:99}}>
        {TABS.map(t=>(<button key={t.id} onClick={()=>navigateTab(t.id)} style={{flex:"0 0 auto",padding:"12px 14px",fontSize:12,fontWeight:tab===t.id?800:500,color:tab===t.id?C.gold:"rgba(255,255,255,0.6)",background:"none",border:"none",borderBottom:tab===t.id?`3px solid ${C.gold}`:"3px solid transparent",cursor:"pointer",whiteSpace:"nowrap",minHeight:44,position:"relative"}}>{t.icon} {t.label}{t.id==="Reports"&&overdueCount>0&&<span style={{position:"absolute",top:6,right:4,background:C.red,color:"#fff",borderRadius:10,fontSize:9,fontWeight:800,padding:"1px 5px"}}>{overdueCount}</span>}</button>))}
      </div>
      <div className="main" style={{marginLeft:0,padding:"18px 16px 40px",maxWidth:860}}>
        {tab==="Dashboard"  &&<DashboardTab data={data} tots={{totTradingSale,totTradingPaid,totTradingOut,totComm}} onNav={(id,hint)=>navigateTab(id,hint)} tradingOut={tradingOut} lastBackupAt={lastBackupAt} onShareBackup={handleShareBackup}/>}
        {tab==="Trading"    &&<TradingTab data={data} onAdd={()=>setModal({type:"sale"})} onAddPay={()=>setModal({type:"payment"})} onAddDebit={()=>setModal({type:"debit"})} onAddCredit={()=>setModal({type:"credit"})} onDel={del} onEdit={(type,rec)=>setModal({type:`edit-${type}`,rec})} tradingOut={tradingOut} initialSearch={navHint?.search} initialView={navHint?.view}/>}
        {tab==="Analytics"  &&<AnalyticsTab data={data} tradingOut={tradingOut}/>}
        {tab==="Market"     &&<MarketTab data={data} onAddPrice={addMaterialPrice}/>}
        {tab==="Enquiry"    &&<EnquiryTab data={data} onAdd={r=>add("enquiries",r)} onUpdate={r=>{setData(p=>({...p,enquiries:p.enquiries.map(e=>e.id===r.id?r:e)}));showToast("Updated.");}} onDel={id=>del("enquiries",id)}/>}
        {tab==="Masters"    &&<MastersTab data={data} onAdd={setModal} onDel={del} onEdit={(type,rec)=>setModal({type:`edit-${type}`,rec})} onImportExcel={()=>importRef.current&&importRef.current.click()} trash={data.trash||[]} onRestore={restoreTrash} onPermanentDelete={permanentDeleteTrash} onEmptyTrash={emptyTrash} initialSearch={navHint?.search} initialView={navHint?.view}/>}
        {tab==="Reports"    &&<ReportsTab data={data} tradingOut={tradingOut} tots={{totTradingSale,totTradingPaid,totTradingOut,totComm}} generatePDF={generatePDF} onAddPay={(pre)=>setModal({type:"payment",preCustomer:pre})} initialRep={navHint?.repView} initialCustomer={navHint?.customer}/>}
        {tab==="Settings"   &&<SettingsTab data={data} onSave={updateCompanySettings} onSaveApp={updateAppSettings}/>}
      </div>
      {modal?.type==="sale"          &&<SaleModal    data={data} onSave={r=>add("tradingSales",r)}    onClose={()=>setModal(null)}/>}
      {modal?.type==="payment"       &&<PaymentModal data={data} onSave={r=>add("tradingPayments",r)} onClose={()=>setModal(null)} preCustomer={modal.preCustomer}/>}
      {modal?.type==="debit"         &&<DebitModal   data={data} onSave={r=>add("debitNotes",r)}      onClose={()=>setModal(null)}/>}
      {modal?.type==="credit"        &&<CreditModal  data={data} onSave={r=>add("creditNotes",r)}     onClose={()=>setModal(null)}/>}
      {modal?.type==="edit-sale"     &&<SaleModal    data={data} onSave={r=>updateMaster("tradingSales",{...modal.rec,...r})}    onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-payment"  &&<PaymentModal data={data} onSave={r=>updateMaster("tradingPayments",{...modal.rec,...r})} onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-debit"    &&<DebitModal   data={data} onSave={r=>updateMaster("debitNotes",{...modal.rec,...r})}      onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="edit-credit"   &&<CreditModal  data={data} onSave={r=>updateMaster("creditNotes",{...modal.rec,...r})}     onClose={()=>setModal(null)} initial={modal.rec}/>}
      {modal?.type==="customer"      &&<CustomerModal data={data} onSave={r=>add("customers",r)}                  onClose={()=>setModal(null)} existingCustomers={data.customers}/>}
      {modal?.type==="supplier"      &&<SupplierModal onSave={r=>add("suppliers",r)}                  onClose={()=>setModal(null)}/>}
      {modal?.type==="product"       &&<ProductModal  data={data} onSave={r=>add("products",r)}       onClose={()=>setModal(null)}/>}
      {modal?.type==="edit-customer" &&<CustomerModal data={data} onSave={r=>updateMaster("customers",r)}         onClose={()=>setModal(null)} initial={modal.rec} existingCustomers={data.customers}/>}
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
  const backupReminderDays=parseInt(data.appSettings?.backupReminderDays)||7;
  const backupOverdue=daysSinceBackup===null||daysSinceBackup>=backupReminderDays;

  const phoneMap={};data.customers.forEach(c=>{phoneMap[c.name]=c.phone;});
  const dueBills=computeDueBills(data);
  const overdueList=dueBills.filter(b=>b.daysUntilDue<0);
  const dueTodayList=dueBills.filter(b=>b.daysUntilDue===0);
  const dueTomorrowList=dueBills.filter(b=>b.daysUntilDue===1);
  const dueWeekList=dueBills.filter(b=>b.daysUntilDue>=2&&b.daysUntilDue<=7);
  const urgentList=[...overdueList,...dueTodayList,...dueTomorrowList].slice(0,6);
  const dueStatusText=(b)=>b.daysUntilDue<0?`overdue by ${-b.daysUntilDue}d`:b.daysUntilDue===0?"due today":b.daysUntilDue===1?"due tomorrow":`due in ${b.daysUntilDue}d`;
  const coName=data.companySettings?.name||"Navkar Fabrics";
  const buildDueWA=(b)=>`🏢 *${coName}*\n\nDear *${b.customerName}*,\n\nReminder: Payment of ₹${fmt(b.amount)} for Bill No ${b.billNo} (dated ${fmtD(b.date)}) is ${dueStatusText(b)}.\n\nKindly arrange payment at the earliest.\n\nRegards\n${coName}`;

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
      {[{icon:"📈",t:"Analytics",l:"Analytics",bg:C.blue,hint:null},{icon:"⏳",t:"Reports",l:"Outstanding",bg:C.red,hint:{repView:"outstanding"}},{icon:"📅",t:"Reports",l:"Ageing",bg:C.orange,hint:{repView:"ageing"}},{icon:"📋",t:"Reports",l:"Reports",bg:C.purple,hint:null}].map(q=>(
        <button key={q.l} onClick={()=>onNav(q.t,q.hint)} style={{background:q.bg,color:"#fff",border:"none",borderRadius:13,padding:"14px 10px",fontSize:13,fontWeight:700,cursor:"pointer",textAlign:"center"}}>
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
          <button onClick={()=>generateInvoice(s,data.tradingSales,data.customers.find(c=>c.name===s.customerName),data.companySettings)} style={{background:"#FEF6E7",color:"#9A7B1E",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🧾 Invoice</button>
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
    return `🏢 *${data.companySettings?.name||"Navkar Fabrics"}*\n\nDear *${v.name}*,\n\nOutstanding payment reminder:\n\n${billLines}\n\n💰 *Total O/S: ₹${fmtAmt(v.net)}*\n\nKindly ignore if already paid.\n\nRegards\n${data.companySettings?.name||"Navkar Fabrics"}`;
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
  const[customFrom,setCustomFrom]=useState("");
  const[customTo,setCustomTo]=useState("");
  const useCustomRange=!!(customFrom&&customTo);
  const availableFYs=useMemo(()=>getAvailableFYs(entries),[entries]);
  const inCustomRange=(dateStr)=>{
    const d=new Date(dateStr).setHours(0,0,0,0);
    return d>=new Date(customFrom).setHours(0,0,0,0)&&d<=new Date(customTo).setHours(0,0,0,0);
  };
  const filteredEntries=useMemo(()=>{
    if(useCustomRange)return entries.filter(r=>inCustomRange(r.date));
    return fy===ALL_FY?entries:filterByFY(entries,fy);
  },[entries,fy,customFrom,customTo]);

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
    return `🏢 *${data.companySettings?.name||"Navkar Fabrics"}*\n\nDear *${selectedCustomer}*,\n\nStatement of Account (last ${Math.min(10,filteredEntries.length)} entries):\n\n${lines}\n\n💰 *Closing Balance: ₹${fmt(closingBalance)}*\n\nRegards\n${data.companySettings?.name||"Navkar Fabrics"}`;
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

    {/* Date-wise filter — narrow the statement down to one financial year or a custom range, running balance stays accurate */}
    <div style={{marginBottom:14}}>
      <div style={{fontSize:10.5,color:C.muted,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>Filter by Date</div>
      <div style={{display:"flex",overflowX:"auto",gap:6,scrollbarWidth:"none",paddingBottom:2,marginBottom:customFrom||customTo?8:0}}>
        {availableFYs.map(f=>(
          <button key={f} onClick={()=>{setFy(f);setCustomFrom("");setCustomTo("");}} style={{flex:"0 0 auto",padding:"8px 14px",borderRadius:20,fontSize:12,fontWeight:!useCustomRange&&fy===f?700:500,border:`1.5px solid ${!useCustomRange&&fy===f?C.navy:C.border}`,background:!useCustomRange&&fy===f?C.navy:"#fff",color:!useCustomRange&&fy===f?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>
            {f===ALL_FY?"📊 All Time":`FY ${f}`}
          </button>
        ))}
        <button onClick={()=>{if(!customFrom)setCustomFrom(today());if(!customTo)setCustomTo(today());}} style={{flex:"0 0 auto",padding:"8px 14px",borderRadius:20,fontSize:12,fontWeight:useCustomRange?700:500,border:`1.5px solid ${useCustomRange?C.navy:C.border}`,background:useCustomRange?C.navy:"#fff",color:useCustomRange?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>📅 Custom Range</button>
      </div>
      {(customFrom||customTo||useCustomRange)&&<div style={{display:"flex",gap:8,alignItems:"center"}}>
        <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)} style={{...IS,padding:"8px 10px",fontSize:12.5}}/>
        <span style={{fontSize:11,color:C.muted}}>to</span>
        <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)} style={{...IS,padding:"8px 10px",fontSize:12.5}}/>
        <button onClick={()=>{setCustomFrom("");setCustomTo("");}} style={{background:"#F0F4F8",border:"none",borderRadius:8,padding:"8px 10px",fontSize:11,color:C.muted,cursor:"pointer",flexShrink:0}}>✕</button>
      </div>}
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
    return Object.values(map).sort((a,b)=>b.sales-a.sales).slice(0,10).map(v=>({...v,pct:(v.sales/total*100)}));
  },[filtered]);

  // Products for selected FY
  const products=useMemo(()=>{
    const map={};
    filtered.forEach(s=>{const p=s.productName||"Unknown";if(!map[p])map[p]={name:p.length>16?p.slice(0,16)+"…":p,value:0};map[p].value+=+s.amount||0;});
    return Object.values(map).sort((a,b)=>b.value-a.value).slice(0,7);
  },[filtered]);

  const productsTable=useMemo(()=>{
    const map={};
    filtered.forEach(s=>{const p=s.productName||"Unknown";if(!map[p])map[p]={name:p,value:0,meters:0};map[p].value+=+s.amount||0;map[p].meters+=+s.meters||0;});
    const total=Object.values(map).reduce((a,v)=>a+v.value,0)||1;
    return Object.values(map).sort((a,b)=>b.value-a.value).slice(0,10).map(v=>({...v,pct:(v.value/total*100)}));
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

  // 🏆 Auto-generated business insights — plain-language takeaways synthesized from the data above,
  // so the owner gets the "so what" without having to read every chart themselves.
  const insights=useMemo(()=>{
    const list=[];
    if(monthly.length>0){
      const best=[...monthly].sort((a,b)=>b.sales-a.sales)[0];
      list.push({icon:"🏆",text:`Best month: ${best.label} with ₹${fmt(best.sales)} in sales.`});
    }
    if(topCustomersTable.length>0){
      const top=topCustomersTable[0];
      if(top.pct>=40) list.push({icon:"⚠️",text:`${top.name} alone is ${top.pct.toFixed(0)}% of total sales (₹${fmt(top.sales)}) — high dependency on one customer.`});
      else list.push({icon:"👑",text:`Top customer: ${top.name} — ${top.pct.toFixed(0)}% of sales (₹${fmt(top.sales)}).`});
    }
    if(totSales>0){
      const collectionPct=totPays/totSales*100;
      if(collectionPct>=90)list.push({icon:"✅",text:`Strong collections — ${collectionPct.toFixed(0)}% of sales already recovered.`});
      else if(collectionPct<60)list.push({icon:"🔴",text:`Only ${collectionPct.toFixed(0)}% of sales collected so far — outstanding is building up.`});
    }
    if(monthly.length>=2){
      const last=monthly[monthly.length-1],prev=monthly[monthly.length-2];
      if(prev.sales>0){
        const change=(last.sales-prev.sales)/prev.sales*100;
        if(Math.abs(change)>=10)list.push({icon:change>0?"📈":"📉",text:`Sales ${change>0?"up":"down"} ${Math.abs(change).toFixed(0)}% (${prev.label} → ${last.label}).`});
      }
    }
    if(products.length>0)list.push({icon:"📦",text:`Top product: ${products[0].name} — ₹${fmt(products[0].value)}.`});
    return list.slice(0,5);
  },[monthly,topCustomersTable,products,totSales,totPays]);

  // 📅 Sales activity heatmap — daily totals grouped into weeks (GitHub-contributions style),
  // covering the selected FY, or the trailing 12 months when viewing All Time.
  const heatmapWeeks=useMemo(()=>{
    let start,end;
    if(fy!==ALL_FY){const r=getFYRange(fy);start=r.start;end=r.end>new Date()?new Date():r.end;}
    else{end=new Date();start=new Date();start.setDate(start.getDate()-364);}
    const dayTotals={};
    filtered.forEach(s=>{const k=new Date(s.date).toISOString().split("T")[0];dayTotals[k]=(dayTotals[k]||0)+(+s.amount||0);});
    const gridStart=new Date(start);gridStart.setDate(gridStart.getDate()-gridStart.getDay()); // back up to Sunday
    const days=[];
    for(let d=new Date(gridStart);d<=end;d.setDate(d.getDate()+1)){
      const k=d.toISOString().split("T")[0];
      days.push({date:new Date(d),key:k,value:dayTotals[k]||0,inRange:d>=start&&d<=end});
    }
    const max=Math.max(1,...days.map(d=>d.value));
    const level=(v)=>v<=0?0:v<max*0.25?1:v<max*0.5?2:v<max*0.75?3:4;
    const withLevel=days.map(d=>({...d,level:level(d.value)}));
    const weeks=[];
    for(let i=0;i<withLevel.length;i+=7)weeks.push(withLevel.slice(i,i+7));
    return weeks;
  },[filtered,fy]);

  // Debit / Credit notes for selected FY
  const filteredDebit=useMemo(()=>{
    const list=data.debitNotes||[];
    if(fy===ALL_FY)return list;
    const{start,end}=getFYRange(fy);
    return list.filter(n=>{const d=new Date(n.date);return d>=start&&d<=end;});
  },[data,fy]);
  const filteredCredit=useMemo(()=>{
    const list=data.creditNotes||[];
    if(fy===ALL_FY)return list;
    const{start,end}=getFYRange(fy);
    return list.filter(n=>{const d=new Date(n.date);return d>=start&&d<=end;});
  },[data,fy]);
  const totDebit=filteredDebit.reduce((a,n)=>a+(+n.amount||0),0);
  const totCredit=filteredCredit.reduce((a,n)=>a+(+n.amount||0),0);
  const notesChartData=[{name:"Debit Notes",value:totDebit},{name:"Credit Notes",value:totCredit}];

  // Cumulative sales/collections trend (for the Overview composed chart) — within selected FY
  const cumulativeTrend=useMemo(()=>{
    let cumS=0,cumP=0;
    return monthly.map(m=>{cumS+=m.sales;cumP+=m.payments;return{...m,netOutstanding:Math.max(0,cumS-cumP)};});
  },[monthly]);

  // Ageing snapshot — always current (live balance), independent of FY filter, matches Ageing tab's corrected FIFO logic
  const ageingSnapshot=useMemo(()=>{
    const todayMs=new Date().setHours(0,0,0,0);
    const map={};
    data.tradingSales.forEach(s=>{if(!map[s.customerName])map[s.customerName]={name:s.customerName,invoices:[]};map[s.customerName].invoices.push({date:s.date,amount:+s.amount||0});});
    const paid={};data.tradingPayments.forEach(p=>{paid[p.customerName]=(paid[p.customerName]||0)+(+p.amount||0);});
    const credit={};(data.creditNotes||[]).forEach(n=>{credit[n.customerName]=(credit[n.customerName]||0)+(+n.amount||0);});
    const debitByCust={};(data.debitNotes||[]).forEach(n=>{if(!debitByCust[n.customerName])debitByCust[n.customerName]=[];debitByCust[n.customerName].push(n);});
    Object.keys(debitByCust).forEach(name=>{if(!map[name])map[name]={name,invoices:[]};});
    const bucketTotals={b0:0,b30:0,b60:0,b90:0,b120:0};
    const customerRows=[];
    Object.values(map).forEach(v=>{
      let rem=(paid[v.name]||0)+(credit[v.name]||0);
      const bk={b0:0,b30:0,b60:0,b90:0,b120:0};
      [...v.invoices].sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach(inv=>{
        let amt=inv.amount;const d=Math.min(amt,rem);amt-=d;rem-=d;if(amt<=0)return;
        const days=Math.floor((todayMs-new Date(inv.date).setHours(0,0,0,0))/86400000);
        if(days<=30)bk.b0+=amt;else if(days<=60)bk.b30+=amt;else if(days<=90)bk.b60+=amt;else if(days<=120)bk.b90+=amt;else bk.b120+=amt;
      });
      (debitByCust[v.name]||[]).forEach(n=>{
        const amt=+n.amount||0;if(amt<=0)return;
        const days=Math.floor((todayMs-new Date(n.date).setHours(0,0,0,0))/86400000);
        if(days<=30)bk.b0+=amt;else if(days<=60)bk.b30+=amt;else if(days<=90)bk.b60+=amt;else if(days<=120)bk.b90+=amt;else bk.b120+=amt;
      });
      const total=Object.values(bk).reduce((a,b)=>a+b,0);
      if(total>0){Object.keys(bk).forEach(k=>bucketTotals[k]+=bk[k]);customerRows.push({name:v.name,total,bk});}
    });
    const chartData=[{label:"0-30d",value:bucketTotals.b0,color:C.green},{label:"31-60d",value:bucketTotals.b30,color:C.orange},{label:"61-90d",value:bucketTotals.b60,color:"#E67E22"},{label:"91-120d",value:bucketTotals.b90,color:C.red},{label:"120d+",value:bucketTotals.b120,color:"#922B21"}];
    return{chartData,customerRows:customerRows.sort((a,b)=>b.total-a.total).slice(0,8),grandTotal:Object.values(bucketTotals).reduce((a,b)=>a+b,0)};
  },[data]);

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
      {[{v:"overview",l:"🧭 Overview"},{v:"heatmap",l:"🔥 Activity"},{v:"monthly",l:"📅 Monthly"},{v:"trend",l:"📈 Trend"},{v:"yoy",l:"📊 Year-on-Year"},{v:"customer",l:"👥 Customers"},{v:"product",l:"📦 Products"},{v:"ageing",l:"⏳ Ageing"},{v:"fycompare",l:"📈 FY Summary"},{v:"notes",l:"↩️ Debit/Credit"}].map(ct=>(
        <button key={ct.v} onClick={()=>setChartType(ct.v)} style={{flex:"0 0 auto",padding:"9px 13px",borderRadius:20,fontSize:12,fontWeight:chartType===ct.v?700:500,border:`1.5px solid ${chartType===ct.v?C.navy:C.border}`,background:chartType===ct.v?C.navy:"#fff",color:chartType===ct.v?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap",minHeight:38}}>{ct.l}</button>
      ))}
    </div>

    <div style={{background:C.card,borderRadius:14,padding:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      {chartType==="overview"&&<>
        {insights.length>0&&<div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:14,padding:"14px 16px",marginBottom:16,border:"1px solid rgba(232,201,126,0.25)"}}>
          <div style={{fontSize:10,color:C.gold,letterSpacing:1.5,textTransform:"uppercase",fontWeight:700,marginBottom:10}}>🏆 Business Insights</div>
          {insights.map((ins,i)=>(
            <div key={i} style={{display:"flex",gap:8,alignItems:"flex-start",padding:i>0?"8px 0 0":"0",borderTop:i>0?"1px solid rgba(255,255,255,0.08)":"none",marginTop:i>0?8:0}}>
              <span style={{fontSize:15,flexShrink:0}}>{ins.icon}</span>
              <span style={{fontSize:12.5,color:"rgba(255,255,255,0.85)",lineHeight:1.5}}>{ins.text}</span>
            </div>
          ))}
        </div>}
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Business Overview {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Monthly Sales & Collections with running Net Outstanding</div>
        {cumulativeTrend.length===0?<Empty text="No data for selected year."/>:
        <ResponsiveContainer width="100%" height={280}>
          <ComposedChart data={cumulativeTrend} margin={{top:4,right:8,left:0,bottom:4}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
            <Bar dataKey="sales" fill={C.blue} radius={[4,4,0,0]} name="Sales" barSize={14}/>
            <Bar dataKey="payments" fill={C.green} radius={[4,4,0,0]} name="Collections" barSize={14}/>
            <Line type="monotone" dataKey="netOutstanding" stroke={C.red} strokeWidth={2.5} dot={{r:3,fill:C.red}} name="Net Outstanding (cum.)"/>
          </ComposedChart>
        </ResponsiveContainer>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:14}}>
          <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.blue}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>SALES</div><div style={{fontSize:13,fontWeight:900,color:C.blue}}>₹{fmt(totSales)}</div></div>
          <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.green}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>COLLECTED</div><div style={{fontSize:13,fontWeight:900,color:C.green}}>₹{fmt(totPays)}</div></div>
          <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.red}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>COLLECTION %</div><div style={{fontSize:13,fontWeight:900,color:C.red}}>{totSales>0?(totPays/totSales*100).toFixed(1):0}%</div></div>
        </div>
      </>}

      {chartType==="heatmap"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Sales Activity Heatmap {fy!==ALL_FY?`— FY ${fy}`:"— Last 12 Months"}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Darker = higher sales that day. Spot your busiest days &amp; patterns at a glance.</div>
        <div style={{overflowX:"auto",paddingBottom:8}}>
          <div style={{display:"flex",gap:3,minWidth:heatmapWeeks.length*13}}>
            {heatmapWeeks.map((week,wi)=>(
              <div key={wi} style={{display:"flex",flexDirection:"column",gap:3}}>
                {week.map((day,di)=>{
                  const colors=["#EEF2F6","#C8E6D8","#7FCBA4","#3DA872","#166B3F"];
                  return(<div key={di} title={`${day.key}: ₹${fmt(day.value)}`} style={{width:10,height:10,borderRadius:2,background:day.inRange?colors[day.level]:"transparent"}}/>);
                })}
              </div>
            ))}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,marginTop:10,fontSize:10,color:C.muted}}>
          <span>Less</span>
          {["#EEF2F6","#C8E6D8","#7FCBA4","#3DA872","#166B3F"].map((c,i)=>(<div key={i} style={{width:10,height:10,borderRadius:2,background:c}}/>))}
          <span>More</span>
        </div>
        {(()=>{
          const allDays=heatmapWeeks.flat().filter(d=>d.inRange);
          const activeDays=allDays.filter(d=>d.value>0).length;
          const best=allDays.reduce((a,d)=>d.value>(a?.value||0)?d:a,null);
          return(
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:16}}>
              <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.teal}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>ACTIVE DAYS</div><div style={{fontSize:13,fontWeight:900,color:C.teal}}>{activeDays} / {allDays.length}</div></div>
              <div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${C.gold}`}}><div style={{fontSize:9.5,color:C.muted,fontWeight:600}}>BEST DAY</div><div style={{fontSize:12.5,fontWeight:900,color:"#9A7B1E"}}>{best&&best.value>0?`${fmtD(best.key)} · ₹${fmt(best.value)}`:"—"}</div></div>
            </div>
          );
        })()}
      </>}

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
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Month-over-month trend — Sales vs Collections</div>
        {monthly.length===0?<Empty text="No data for selected year."/>:
        <ResponsiveContainer width="100%" height={260}>
          <AreaChart data={monthly} margin={{top:4,right:8,left:0,bottom:4}}>
            <defs>
              <linearGradient id="gradSales" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C.blue} stopOpacity={0.35}/>
                <stop offset="95%" stopColor={C.blue} stopOpacity={0.02}/>
              </linearGradient>
              <linearGradient id="gradPay" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C.green} stopOpacity={0.35}/>
                <stop offset="95%" stopColor={C.green} stopOpacity={0.02}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
            <Area type="monotone" dataKey="sales" stroke={C.blue} strokeWidth={2.5} fill="url(#gradSales)" name="Sales"/>
            <Area type="monotone" dataKey="payments" stroke={C.green} strokeWidth={2.5} fill="url(#gradPay)" name="Collections"/>
          </AreaChart>
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
        {topCustomers.length===0?<Empty text="No data."/>:<>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={topCustomers} layout="vertical" margin={{top:0,right:8,left:70,bottom:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8" horizontal={false}/>
            <XAxis type="number" tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:9}} tickLine={false}/>
            <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={70} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Bar dataKey="sales" fill={C.blue} radius={[0,5,5,0]} name="Sales" barSize={12}/>
          </BarChart>
        </ResponsiveContainer>
        <div style={{marginTop:16,overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse",width:"100%",minWidth:420}}>
            <thead><tr>
              <th style={{padding:"7px 8px",textAlign:"left",fontSize:10,color:C.gold,background:C.navy,borderRadius:"8px 0 0 0"}}>#</th>
              <th style={{padding:"7px 8px",textAlign:"left",fontSize:10,color:C.gold,background:C.navy}}>Customer</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>Bills</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>Sales</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy,borderRadius:"0 8px 0 0"}}>Share</th>
            </tr></thead>
            <tbody>
              {topCustomersTable.map((v,i)=>(
                <tr key={v.name} style={{background:i%2===0?"#fff":"#FAFBFC"}}>
                  <td style={{padding:"7px 8px",fontSize:11.5,color:C.muted,borderBottom:`1px solid ${C.border}`}}>{i+1}</td>
                  <td style={{padding:"7px 8px",fontSize:12,fontWeight:700,color:C.navy,borderBottom:`1px solid ${C.border}`}}>{v.name}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",color:C.muted,borderBottom:`1px solid ${C.border}`}}>{v.bills}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",fontWeight:800,color:C.blue,borderBottom:`1px solid ${C.border}`}}>₹{fmt(v.sales)}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",color:C.muted,borderBottom:`1px solid ${C.border}`}}>{v.pct.toFixed(1)}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        </>}
      </>}

      {chartType==="product"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Sales by Product {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Top 7 products by value</div>
        {products.length===0?<Empty text="No data."/>:<>
        <ResponsiveContainer width="100%" height={260}>
          <PieChart>
            <Pie data={products} cx="50%" cy="50%" outerRadius={95} innerRadius={40} dataKey="value" nameKey="name"
              label={({name,percent})=>`${(percent*100).toFixed(0)}%`} labelLine={true} fontSize={10}>
              {products.map((_,i)=><Cell key={i} fill={PIE[i%PIE.length]}/>)}
            </Pie>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Legend iconSize={10} wrapperStyle={{fontSize:10}}/>
          </PieChart>
        </ResponsiveContainer>
        <div style={{marginTop:16,overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse",width:"100%",minWidth:420}}>
            <thead><tr>
              <th style={{padding:"7px 8px",textAlign:"left",fontSize:10,color:C.gold,background:C.navy,borderRadius:"8px 0 0 0"}}>#</th>
              <th style={{padding:"7px 8px",textAlign:"left",fontSize:10,color:C.gold,background:C.navy}}>Product</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>Meters</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>Sales</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy,borderRadius:"0 8px 0 0"}}>Share</th>
            </tr></thead>
            <tbody>
              {productsTable.map((v,i)=>(
                <tr key={v.name} style={{background:i%2===0?"#fff":"#FAFBFC"}}>
                  <td style={{padding:"7px 8px",fontSize:11.5,color:C.muted,borderBottom:`1px solid ${C.border}`}}>{i+1}</td>
                  <td style={{padding:"7px 8px",fontSize:12,fontWeight:700,color:C.navy,borderBottom:`1px solid ${C.border}`}}>{v.name}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",color:C.muted,borderBottom:`1px solid ${C.border}`}}>{fmt(v.meters)}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",fontWeight:800,color:C.teal,borderBottom:`1px solid ${C.border}`}}>₹{fmt(v.value)}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",color:C.muted,borderBottom:`1px solid ${C.border}`}}>{v.pct.toFixed(1)}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        </>}
      </>}

      {chartType==="ageing"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Ageing Snapshot</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Live outstanding by age bucket — always current (not affected by FY filter)</div>
        {ageingSnapshot.grandTotal===0?<Empty text="No outstanding found."/>:<>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={ageingSnapshot.chartData} margin={{top:4,right:4,left:0,bottom:4}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Bar dataKey="value" radius={[6,6,0,0]} barSize={40} name="Outstanding">
              {ageingSnapshot.chartData.map((d,i)=><Cell key={i} fill={d.color}/>)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
        <div style={{marginTop:16,overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse",width:"100%",minWidth:520}}>
            <thead><tr>
              <th style={{padding:"7px 8px",textAlign:"left",fontSize:10,color:C.gold,background:C.navy,borderRadius:"8px 0 0 0"}}>Customer</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>0-30d</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>31-90d</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>90d+</th>
              <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy,borderRadius:"0 8px 0 0"}}>Total</th>
            </tr></thead>
            <tbody>
              {ageingSnapshot.customerRows.map((v,i)=>(
                <tr key={v.name} style={{background:i%2===0?"#fff":"#FAFBFC"}}>
                  <td style={{padding:"7px 8px",fontSize:12,fontWeight:700,color:C.navy,borderBottom:`1px solid ${C.border}`}}>{v.name}</td>
                  <td style={{padding:"7px 8px",fontSize:11.5,textAlign:"right",color:C.green,borderBottom:`1px solid ${C.border}`}}>₹{fmt(v.bk.b0)}</td>
                  <td style={{padding:"7px 8px",fontSize:11.5,textAlign:"right",color:C.orange,borderBottom:`1px solid ${C.border}`}}>₹{fmt(v.bk.b30+v.bk.b60+v.bk.b90)}</td>
                  <td style={{padding:"7px 8px",fontSize:11.5,textAlign:"right",color:"#922B21",borderBottom:`1px solid ${C.border}`}}>₹{fmt(v.bk.b120)}</td>
                  <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",fontWeight:900,color:C.red,borderBottom:`1px solid ${C.border}`}}>₹{fmt(v.total)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        </>}
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

      {chartType==="notes"&&<>
        <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>Debit & Credit Notes {fy!==ALL_FY?`— FY ${fy}`:""}</div>
        <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Returns and adjustments overview</div>
        <div style={{background:"#FEF9E7",borderRadius:12,padding:"12px 14px",marginBottom:14,border:"1px solid #F9E79F"}}>
          <div style={{fontWeight:700,fontSize:13,color:"#7D6608",marginBottom:4}}>How Returns Work</div>
          <div style={{fontSize:12,color:"#7D6608",lineHeight:1.6}}><b>Debit Note</b> — Customer owes MORE (price difference, penalty, extra goods sent).<br/><b>Credit Note / Sale Return</b> — Customer owes LESS (goods returned, allowance given).</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
          <div style={{background:"#F8FAFC",borderRadius:13,padding:"13px 14px",borderLeft:`4px solid ${C.red}`}}><div style={{fontSize:10,color:C.muted,textTransform:"uppercase",fontWeight:700}}>Debit Notes</div><div style={{fontSize:20,fontWeight:900,color:C.red}}>₹{fmt(totDebit)}</div><div style={{fontSize:11,color:C.muted}}>{filteredDebit.length} entries</div></div>
          <div style={{background:"#F8FAFC",borderRadius:13,padding:"13px 14px",borderLeft:`4px solid ${C.green}`}}><div style={{fontSize:10,color:C.muted,textTransform:"uppercase",fontWeight:700}}>Credit / Return</div><div style={{fontSize:20,fontWeight:900,color:C.green}}>₹{fmt(totCredit)}</div><div style={{fontSize:11,color:C.muted}}>{filteredCredit.length} entries</div></div>
        </div>
        {(totDebit>0||totCredit>0)?
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={notesChartData} margin={{top:4,right:4,left:0,bottom:4}}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
            <XAxis dataKey="name" tick={{fontSize:11}} tickLine={false}/>
            <YAxis tickFormatter={v=>v>=100000?`₹${(v/100000).toFixed(0)}L`:`₹${(v/1000).toFixed(0)}K`} tick={{fontSize:10}} tickLine={false} axisLine={false}/>
            <Tooltip formatter={v=>`₹${fmt(v)}`}/>
            <Bar dataKey="value" radius={[6,6,0,0]} barSize={60}>
              <Cell fill={C.red}/><Cell fill={C.green}/>
            </Bar>
          </BarChart>
        </ResponsiveContainer>
        :<Empty text="No debit/credit notes for selected year."/>}
      </>}
    </div>
  </div>);
}

// ─── MARKET INTELLIGENCE TAB ────────────────────────────────────
// Historical raw-material prices are logged manually (no live external market feed configured),
// and stored inside the same synced `data` object as everything else — no separate Supabase
// table/RLS setup needed. Falls back to illustrative dummy data (last 4 weeks) when nothing
// has been logged yet, so the charts/UI are never empty on first use.
function generateDummyMaterialPrices(){
  const days=[];
  const base=new Date();
  for(let i=3;i>=0;i--){
    const d=new Date(base);d.setDate(d.getDate()-i*7);
    days.push({
      id:`dummy-${i}`,
      date:d.toISOString().split("T")[0],
      cotton_price:180+Math.round((Math.random()*20-10)),
      dyes_price:95+Math.round((Math.random()*10-5)),
      power_index:100+Math.round((Math.random()*8-4)),
      isDummy:true,
    });
  }
  return days;
}
function getMarketPitch(demand){
  if(demand>10)return{
    title:"🔴 High Demand — Lock Orders Now",
    text:"Festival / export season demand is heating up. Cotton, yarn and dyeing costs are likely to move up from here. Advise customers to lock in quantity and price now, before the next round of hikes — position it as a limited-time rate, not a routine order.",
    color:C.red,bg:"#FDEDEC",border:"#F5B7B1",label:"BULLISH — SELLER'S MARKET",
  };
  if(demand<-10)return{
    title:"🟢 Soft Demand — Offer Long-Term Contracts",
    text:"Demand is soft right now. This is a good window to offer aggressive long-term contract pricing to key customers — lock in their volume for the next few months while your own input costs are favourable, rather than competing purely on one-off order price.",
    color:C.green,bg:"#E8F8F0",border:"#A9DFBF",label:"BEARISH — BUYER'S MARKET",
  };
  return{
    title:"🔵 Stable Market — Steady Pricing",
    text:"Demand and input costs are broadly stable. Hold current pricing, avoid discounting reflexively, and focus the pitch on reliability, quality and service rather than price — this is the window to deepen relationships rather than chase volume.",
    color:C.blue,bg:"#EAF4FC",border:"#AED6F1",label:"NEUTRAL — STABLE MARKET",
  };
}
function MarketTab({data,onAddPrice}){
  const[demand,setDemand]=useState(0);
  const[form,setForm]=useState({date:today(),cotton_price:"",dyes_price:"",power_index:""});
  const sf=(k,v)=>setForm(p=>({...p,[k]:v}));

  const loggedPrices=data.materialPrices||[];
  const prices=useMemo(()=>{
    const list=loggedPrices.length>0?loggedPrices:generateDummyMaterialPrices();
    return [...list].sort((a,b)=>new Date(a.date)-new Date(b.date));
  },[loggedPrices]);
  const usingDummy=loggedPrices.length===0;

  const latest=prices[prices.length-1];
  const prev=prices.length>1?prices[prices.length-2]:null;
  const trend=(key)=>{
    if(!prev)return null;
    const diff=(+latest[key]||0)-(+prev[key]||0);
    if(Math.abs(diff)<0.01)return{arrow:"→",color:C.muted};
    return diff>0?{arrow:"▲",color:C.red}:{arrow:"▼",color:C.green};
  };

  const chartData=prices.map(p=>({
    label:new Date(p.date).toLocaleDateString("en-IN",{day:"2-digit",month:"short"}),
    Cotton:+p.cotton_price||0,
    Dyes:+p.dyes_price||0,
    Power:+p.power_index||0,
  }));

  const pitch=getMarketPitch(demand);

  return(<div>
    <SecTitle>💹 Market Intelligence</SecTitle>
    {usingDummy&&<div style={{background:"#FEF9E7",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
      Showing illustrative sample data — log this week's prices below to start tracking your own real trend.
    </div>}

    {/* KPI row — latest logged prices with trend arrow vs previous entry */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
      {[
        {key:"cotton_price",label:"Cotton / Yarn",unit:"₹/kg",color:C.blue},
        {key:"dyes_price",label:"Dyes / Chemicals",unit:"₹/kg",color:C.purple},
        {key:"power_index",label:"Power / Gas",unit:"Index",color:C.orange},
      ].map(k=>{
        const t=trend(k.key);
        return(<div key={k.key} style={{background:C.card,borderRadius:14,padding:"12px 10px",borderLeft:`4px solid ${k.color}`,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
          <div style={{fontSize:9.5,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:0.3}}>{k.label}</div>
          <div style={{fontSize:16,fontWeight:900,color:k.color,marginTop:3}}>{k.unit==="₹/kg"?"₹":""}{fmt(latest?latest[k.key]:0)}{k.unit==="Index"?"":""}</div>
          <div style={{fontSize:10,color:C.muted,marginTop:1}}>{k.unit}{t&&<span style={{color:t.color,fontWeight:800,marginLeft:6}}>{t.arrow}</span>}</div>
        </div>);
      })}
    </div>

    {/* Trend chart */}
    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <B style={{fontSize:14,color:C.navy,marginBottom:4,display:"block"}}>Raw Material Price Trend</B>
      <Mute style={{marginBottom:10}}>Cotton/Yarn, Dyes/Chemicals &amp; Power/Gas over time</Mute>
      <ResponsiveContainer width="100%" height={260}>
        <LineChart data={chartData} margin={{top:4,right:8,left:0,bottom:4}}>
          <CartesianGrid strokeDasharray="3 3" stroke="#F0F4F8"/>
          <XAxis dataKey="label" tick={{fontSize:10}} tickLine={false}/>
          <YAxis tick={{fontSize:10}} tickLine={false} axisLine={false}/>
          <Tooltip/>
          <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
          <Line type="monotone" dataKey="Cotton" stroke={C.blue} strokeWidth={2.5} dot={{r:3,fill:C.blue}}/>
          <Line type="monotone" dataKey="Dyes" stroke={C.purple} strokeWidth={2.5} dot={{r:3,fill:C.purple}}/>
          <Line type="monotone" dataKey="Power" stroke={C.orange} strokeWidth={2.5} dot={{r:3,fill:C.orange}}/>
        </LineChart>
      </ResponsiveContainer>
    </div>

    {/* Log this week's prices */}
    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <B style={{fontSize:14,color:C.navy,marginBottom:12,display:"block"}}>Log This Week's Prices</B>
      <F label="Date"><input type="date" value={form.date} onChange={e=>sf("date",e.target.value)} style={IS}/></F>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
        <F label="Cotton ₹/kg"><input type="number" value={form.cotton_price} onChange={e=>sf("cotton_price",e.target.value)} placeholder="180" style={IS}/></F>
        <F label="Dyes ₹/kg"><input type="number" value={form.dyes_price} onChange={e=>sf("dyes_price",e.target.value)} placeholder="95" style={IS}/></F>
        <F label="Power Index"><input type="number" value={form.power_index} onChange={e=>sf("power_index",e.target.value)} placeholder="100" style={IS}/></F>
      </div>
      <SaveBtn color={C.navy} onClick={()=>{
        if(!form.cotton_price&&!form.dyes_price&&!form.power_index)return alert("Enter at least one price");
        onAddPrice(form);
        setForm({date:today(),cotton_price:"",dyes_price:"",power_index:""});
      }}>💾 Save Entry</SaveBtn>
    </div>

    {/* Demand forecaster + dynamic pitch */}
    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <B style={{fontSize:14,color:C.navy,marginBottom:4,display:"block"}}>Demand Forecaster</B>
      <Mute style={{marginBottom:14}}>Slide to reflect upcoming Festival/Export season demand — the pitch below updates live.</Mute>
      <input
        type="range" min={-20} max={20} value={demand}
        onChange={e=>setDemand(+e.target.value)}
        style={{width:"100%",accentColor:pitch.color}}
      />
      <Row style={{marginTop:4}}>
        <span style={{fontSize:10.5,color:C.green,fontWeight:700}}>◀ Low Demand</span>
        <span style={{fontSize:13,fontWeight:900,color:pitch.color}}>{demand>0?"+":""}{demand}</span>
        <span style={{fontSize:10.5,color:C.red,fontWeight:700}}>High Demand ▶</span>
      </Row>
    </div>

    <div style={{background:pitch.bg,borderRadius:14,padding:16,border:`1.5px solid ${pitch.border}`}}>
      <div style={{fontSize:10,color:pitch.color,fontWeight:800,letterSpacing:0.5,marginBottom:6}}>{pitch.label}</div>
      <B style={{fontSize:15,color:pitch.color,marginBottom:8,display:"block"}}>{pitch.title}</B>
      <div style={{fontSize:12.5,color:C.navy,lineHeight:1.7}}>{pitch.text}</div>
    </div>
  </div>);
}

// ─── ENQUIRY TAB ─────────────────────────────────────────────────
function EnquiryTab({data,onAdd,onUpdate,onDel}){
  const[showForm,setShowForm]=useState(false);
  const[filter,setFilter]=useState("open"); // open | closed | all
  const[editId,setEditId]=useState(null);
  const[form,setForm]=useState({enquiryDate:today(),customerName:"",description:"",submissionDate:"",closeDate:"",remarks:"",status:"Open"});
  const sf=(k,v)=>setForm(p=>({...p,[k]:v}));

  const todayMs=new Date().setHours(0,0,0,0);
  const enquiries=data.enquiries||[];

  const resetForm=()=>{setForm({enquiryDate:today(),customerName:"",description:"",submissionDate:"",closeDate:"",remarks:"",status:"Open"});setEditId(null);setShowForm(false);};

  const saveEnquiry=()=>{
    if(!form.customerName||!form.description)return alert("Customer and Description required");
    if(editId){onUpdate({...form,id:editId});}
    else{onAdd({...form,followUps:[],createdAt:today()});}
    resetForm();
  };

  const addFollowUp=(enq)=>{
    const note=prompt("Follow-up note (optional):")||"";
    const updated={...enq,followUps:[...(enq.followUps||[]),{date:today(),note,time:new Date().toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}]};
    onUpdate(updated);
  };

  const closeEnquiry=(enq)=>{
    if(!window.confirm("Mark this enquiry as Closed?"))return;
    onUpdate({...enq,status:"Closed",closeDate:today()});
  };

  const reopenEnquiry=(enq)=>onUpdate({...enq,status:"Open",closeDate:""});

  const buildWAEnquiry=(enq)=>{
    const phone=data.customers.find(c=>c.name===enq.customerName)?.phone||"";
    const coName=data.companySettings?.name||"Navkar Fabrics";
    const msg=`🏢 *${coName}*\n\nDear *${enq.customerName}*,\n\nThis is a follow-up regarding:\n\n📋 *${enq.description}*\n\nEnquiry Date: ${fmtD(enq.enquiryDate)}${enq.submissionDate?`\nSubmission Date: ${fmtD(enq.submissionDate)}`:""}\n\nKindly revert at the earliest.\n\nRegards\n${coName}`;
    const n=phone?"91"+String(phone).replace(/\D/g,"").slice(-10):"";
    window.open(`https://wa.me/${n}?text=${encodeURIComponent(msg)}`,"_blank");
  };

  const getDaysOpen=(enq)=>Math.floor((todayMs-new Date(enq.enquiryDate).setHours(0,0,0,0))/86400000);

  const filtered=enquiries.filter(e=>filter==="all"?true:filter==="open"?e.status!=="Closed":e.status==="Closed").sort((a,b)=>new Date(b.enquiryDate)-new Date(a.enquiryDate));
  const openCount=enquiries.filter(e=>e.status!=="Closed").length;

  return(<div>
    {/* Header */}
    <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:14,padding:"14px 16px",marginBottom:14,border:"1px solid rgba(232,201,126,0.25)"}}>
      <div style={{fontSize:10,color:C.gold,textTransform:"uppercase",letterSpacing:1,fontWeight:700}}>📝 Enquiry / Requirement Notes</div>
      <div style={{display:"flex",gap:20,marginTop:10}}>
        <div><div style={{fontSize:22,fontWeight:900,color:"#fff"}}>{openCount}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Open</div></div>
        <div><div style={{fontSize:22,fontWeight:900,color:C.gold}}>{enquiries.filter(e=>e.status==="Closed").length}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Closed</div></div>
        <div><div style={{fontSize:22,fontWeight:900,color:"#fff"}}>{enquiries.length}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Total</div></div>
      </div>
    </div>

    {/* Add Button */}
    {!showForm&&<button onClick={()=>setShowForm(true)} style={{background:C.blue,color:"#fff",border:"none",borderRadius:11,padding:"12px 18px",fontSize:13.5,fontWeight:700,cursor:"pointer",width:"100%",marginBottom:12,minHeight:46}}>+ New Enquiry / Requirement</button>}

    {/* Form */}
    {showForm&&<div style={{background:C.card,borderRadius:14,padding:"16px",marginBottom:14,boxShadow:"0 2px 12px rgba(0,0,0,0.08)",border:`1px solid ${C.border}`}}>
      <div style={{fontWeight:800,fontSize:15,color:C.navy,marginBottom:14}}>{editId?"Edit Enquiry":"New Enquiry / Requirement"}</div>
      <F label="Enquiry / Request Date *"><input type="date" value={form.enquiryDate} onChange={e=>sf("enquiryDate",e.target.value)} style={IS}/></F>
      <F label="Customer Name *"><SmartInput value={form.customerName} onChange={v=>sf("customerName",v)} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="enq-c"/></F>
      <F label="Description / Submission Details *"><textarea value={form.description} onChange={e=>sf("description",e.target.value)} placeholder="What is the enquiry / requirement about?" style={{...IS,minHeight:80,resize:"vertical"}}/></F>
      <F label="Submission Date (if applicable)"><input type="date" value={form.submissionDate} onChange={e=>sf("submissionDate",e.target.value)} style={IS}/></F>
      <F label="Remarks"><input value={form.remarks} onChange={e=>sf("remarks",e.target.value)} placeholder="Any additional notes" style={IS}/></F>
      <div style={{display:"flex",gap:8,marginTop:4}}>
        <button onClick={saveEnquiry} style={{flex:1,background:C.blue,color:"#fff",border:"none",borderRadius:11,padding:14,fontSize:14,fontWeight:800,cursor:"pointer",minHeight:48}}>💾 Save</button>
        <button onClick={resetForm} style={{background:"#F0F4F8",color:C.navy,border:"none",borderRadius:11,padding:14,fontSize:14,fontWeight:700,cursor:"pointer",minHeight:48}}>Cancel</button>
      </div>
    </div>}

    {/* Filter */}
    <div style={{display:"flex",background:"#fff",borderRadius:11,overflow:"hidden",border:`1px solid ${C.border}`,marginBottom:12}}>
      {[{v:"open",l:`🟡 Open (${openCount})`},{v:"closed",l:"✅ Closed"},{v:"all",l:"📋 All"}].map(o=>(
        <button key={o.v} onClick={()=>setFilter(o.v)} style={{flex:1,padding:"11px 4px",fontSize:12,fontWeight:filter===o.v?700:500,color:filter===o.v?"#E8C97E":C.muted,background:filter===o.v?C.navy:"transparent",border:"none",cursor:"pointer",minHeight:44}}>{o.l}</button>
      ))}
    </div>

    {/* List */}
    {filtered.length===0&&<Empty text={filter==="open"?"No open enquiries. Great!":"No enquiries found."}/>}
    {filtered.map(enq=>{
      const daysOpen=getDaysOpen(enq);
      const followUps=enq.followUps||[];
      const isClosed=enq.status==="Closed";
      const urgColor=daysOpen>14?C.red:daysOpen>7?C.orange:C.green;
      const phone=data.customers.find(c=>c.name===enq.customerName)?.phone||"";
      return(
        <div key={enq.id} style={{background:C.card,borderRadius:14,padding:"14px 15px",marginBottom:12,boxShadow:"0 1px 8px rgba(0,0,0,0.07)",borderLeft:`4px solid ${isClosed?C.green:urgColor}`,opacity:isClosed?0.85:1}}>
          {/* Header */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
            <div style={{flex:1,paddingRight:8}}>
              <B style={{fontSize:14,color:C.navy}}>{enq.customerName}</B>
              <Mute>📅 {fmtD(enq.enquiryDate)}</Mute>
            </div>
            <div style={{textAlign:"right",flexShrink:0}}>
              <span style={{fontSize:10,fontWeight:800,background:isClosed?"#E8F8F5":"#FEF9E7",color:isClosed?C.teal:urgColor,padding:"3px 9px",borderRadius:20,border:`1px solid ${isClosed?C.teal:urgColor}`}}>{isClosed?"✅ Closed":`🟡 Open`}</span>
              {!isClosed&&<div style={{fontSize:11,color:urgColor,fontWeight:700,marginTop:4}}>{daysOpen}d open</div>}
            </div>
          </div>

          {/* Description */}
          <div style={{background:"#F8FAFC",borderRadius:9,padding:"10px 12px",marginBottom:8,fontSize:12.5,color:C.navy,lineHeight:1.6}}>{enq.description}</div>

          {/* Details */}
          <div style={{display:"flex",gap:12,fontSize:11.5,color:C.muted,flexWrap:"wrap",marginBottom:8}}>
            {enq.submissionDate&&<span>📤 Submission: <b>{fmtD(enq.submissionDate)}</b></span>}
            {enq.closeDate&&<span>🔒 Closed: <b>{fmtD(enq.closeDate)}</b></span>}
            {enq.remarks&&<span>💬 {enq.remarks}</span>}
          </div>

          {/* Follow-up count + history */}
          <div style={{background:followUps.length>0?"#F3EEF9":"#F8FAFC",borderRadius:9,padding:"8px 12px",marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:700,color:C.purple,marginBottom:followUps.length>0?6:0}}>
              🔁 Follow-ups: {followUps.length}
            </div>
            {followUps.slice(-3).map((f,i)=>(
              <div key={i} style={{fontSize:11.5,color:C.muted,paddingTop:4,borderTop:"1px solid #EEE",marginTop:4}}>
                <b style={{color:C.navy}}>{f.date}</b> {f.time&&`@ ${f.time}`} {f.note&&`— ${f.note}`}
              </div>
            ))}
            {followUps.length>3&&<div style={{fontSize:11,color:C.muted,marginTop:4}}>+{followUps.length-3} more follow-ups</div>}
          </div>

          {/* Action Buttons */}
          {!isClosed&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            <button onClick={()=>addFollowUp(enq)} style={{background:C.purple,color:"#fff",border:"none",borderRadius:9,padding:"9px 13px",fontSize:12,fontWeight:700,cursor:"pointer",flex:1,minHeight:40}}>🔁 Follow-up</button>
            <button onClick={()=>buildWAEnquiry(enq)} style={{background:"#E8FBF0",color:"#25D366",border:"1px solid #25D366",borderRadius:9,padding:"9px 13px",fontSize:12,fontWeight:700,cursor:"pointer",minHeight:40}}>💬 WA</button>
            {phone&&<button onClick={()=>window.open(`tel:${phone}`)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:9,padding:"9px 13px",fontSize:12,fontWeight:700,cursor:"pointer",minHeight:40}}>📞 Call</button>}
            <button onClick={()=>closeEnquiry(enq)} style={{background:"#E8F8F5",color:C.teal,border:"1px solid "+C.teal,borderRadius:9,padding:"9px 13px",fontSize:12,fontWeight:700,cursor:"pointer",minHeight:40}}>✅ Close</button>
          </div>}
          {isClosed&&<div style={{display:"flex",gap:6}}>
            <button onClick={()=>reopenEnquiry(enq)} style={{background:"#FEF9E7",color:C.orange,border:`1px solid ${C.orange}`,borderRadius:9,padding:"9px 13px",fontSize:12,fontWeight:700,cursor:"pointer",minHeight:40}}>🔄 Reopen</button>
            <button onClick={()=>onDel(enq.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:9,padding:"9px 13px",fontSize:12,fontWeight:700,cursor:"pointer",minHeight:40}}>🗑️ Delete</button>
          </div>}
          {!isClosed&&<div style={{display:"flex",gap:6,marginTop:6}}>
            <button onClick={()=>{setForm({...enq});setEditId(enq.id);setShowForm(true);window.scrollTo(0,0);}} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:9,padding:"7px 13px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✏️ Edit</button>
            <button onClick={()=>onDel(enq.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:9,padding:"7px 13px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️ Delete</button>
          </div>}
        </div>
      );
    })}
  </div>);
}

// ─── COMMISSION TAB ──────────────────────────────────────────────
function calcFIFOCommission(tradingSales,tradingPayments,creditNotes=[],debitNotes=[]){
  const customerMap={};
  tradingSales.forEach(s=>{
    const cid=s.customerName;
    if(!customerMap[cid])customerMap[cid]={sales:[],totalPaid:0,creditAdjusted:0};
    customerMap[cid].sales.push({...s});
  });
  tradingPayments.forEach(p=>{
    const cid=p.customerName;
    if(!customerMap[cid])customerMap[cid]={sales:[],totalPaid:0,creditAdjusted:0};
    customerMap[cid].totalPaid+=+p.amount||0;
  });
  // Credit Notes & Debit Notes reduce outstanding but NOT via payment — exclude from commission pool
  creditNotes.forEach(n=>{
    const cid=n.customerName;
    if(!customerMap[cid])customerMap[cid]={sales:[],totalPaid:0,creditAdjusted:0};
    customerMap[cid].creditAdjusted+=+n.amount||0;
  });
  debitNotes.forEach(n=>{
    const cid=n.customerName;
    if(!customerMap[cid])customerMap[cid]={sales:[],totalPaid:0,creditAdjusted:0};
    // Debit notes increase outstanding — no adjustment needed here
  });
  let totalCommission=0;
  const commDetails=[];
  const commByCustomer={};
  Object.entries(customerMap).forEach(([cname,cdata])=>{
    const sales=[...cdata.sales].sort((a,b)=>new Date(a.date)-new Date(b.date));
    // Commission only on CASH payments — exclude credit/debit note adjustments
    let remainingPayment=cdata.totalPaid; // pure cash received only
    sales.forEach(sale=>{
      const saleAmt=+sale.amount||0;
      const qty=+sale.meters||+sale.qty||0;
      if(remainingPayment>=saleAmt&&saleAmt>0){
        remainingPayment-=saleAmt;
        const comm=qty*(+sale.commissionRate||1.5);
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
  const[fy,setFy]=useState(ALL_FY);
  const[customFrom,setCustomFrom]=useState("");
  const[customTo,setCustomTo]=useState("");
  const useCustomRange=!!(customFrom&&customTo);

  // Always run FIFO over the FULL, unfiltered history — a payment can clear a bill from any
  // date, so filtering the inputs first would silently break which bills count as "paid".
  // The date filter below only changes which of the correctly-computed bills are shown/totalled.
  const{commDetails}=useMemo(()=>calcFIFOCommission(data.tradingSales,data.tradingPayments,data.creditNotes||[],data.debitNotes||[]),[data]);

  const availableFYs=useMemo(()=>getAvailableFYs(commDetails),[commDetails]);
  const inCustomRange=(dateStr)=>{
    const d=new Date(dateStr).setHours(0,0,0,0);
    return d>=new Date(customFrom).setHours(0,0,0,0)&&d<=new Date(customTo).setHours(0,0,0,0);
  };
  const detailList=useMemo(()=>{
    let list=[...commDetails];
    if(useCustomRange)list=list.filter(v=>inCustomRange(v.date));
    else if(fy!==ALL_FY)list=filterByFY(list,fy);
    return list.sort((a,b)=>new Date(b.date)-new Date(a.date));
  },[commDetails,fy,customFrom,customTo]);

  const commByCustomer={};
  detailList.forEach(v=>{
    if(!commByCustomer[v.customer])commByCustomer[v.customer]={name:v.customer,commission:0,bills:0,meters:0};
    commByCustomer[v.customer].commission+=v.commission;
    commByCustomer[v.customer].bills+=1;
    commByCustomer[v.customer].meters+=v.qty;
  });
  const commList=Object.values(commByCustomer).sort((a,b)=>b.commission-a.commission);
  const totalCommission=detailList.reduce((a,v)=>a+v.commission,0);
  const totalBills=commList.reduce((a,v)=>a+v.bills,0);
  const totalMeters=commList.reduce((a,v)=>a+v.meters,0);
  const maxComm=commList[0]?.commission||1;

  const DatePicker=()=>(
    <div style={{marginBottom:12}}>
      <div style={{fontSize:10.5,color:C.muted,fontWeight:600,marginBottom:5,textTransform:"uppercase"}}>Filter by Date</div>
      <div style={{display:"flex",overflowX:"auto",gap:5,scrollbarWidth:"none",marginBottom:customFrom||customTo?8:0}}>
        {availableFYs.map(f=>(
          <button key={f} onClick={()=>{setFy(f);setCustomFrom("");setCustomTo("");}} style={{flex:"0 0 auto",padding:"7px 12px",borderRadius:16,fontSize:11.5,fontWeight:!useCustomRange&&fy===f?700:500,border:`1.5px solid ${!useCustomRange&&fy===f?C.navy:C.border}`,background:!useCustomRange&&fy===f?C.navy:"#fff",color:!useCustomRange&&fy===f?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>
            {f===ALL_FY?"All":f}
          </button>
        ))}
        <button onClick={()=>{if(!customFrom)setCustomFrom(today());if(!customTo)setCustomTo(today());}} style={{flex:"0 0 auto",padding:"7px 12px",borderRadius:16,fontSize:11.5,fontWeight:useCustomRange?700:500,border:`1.5px solid ${useCustomRange?C.navy:C.border}`,background:useCustomRange?C.navy:"#fff",color:useCustomRange?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>📅 Custom Range</button>
      </div>
      {(customFrom||customTo||useCustomRange)&&<div style={{display:"flex",gap:8,alignItems:"center"}}>
        <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)} style={{...IS,padding:"8px 10px",fontSize:12.5}}/>
        <span style={{fontSize:11,color:C.muted}}>to</span>
        <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)} style={{...IS,padding:"8px 10px",fontSize:12.5}}/>
        <button onClick={()=>{setCustomFrom("");setCustomTo("");}} style={{background:"#F0F4F8",border:"none",borderRadius:8,padding:"8px 10px",fontSize:11,color:C.muted,cursor:"pointer",flexShrink:0}}>✕</button>
      </div>}
    </div>
  );

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
      <div style={{fontSize:10,color:"rgba(232,201,126,0.7)",textTransform:"uppercase",letterSpacing:1.5,fontWeight:700,marginBottom:4}}>💰 Commission Earned (FIFO){fy!==ALL_FY&&!useCustomRange?` — FY ${fy}`:""}</div>
      <div style={{fontSize:32,fontWeight:900,color:C.gold,letterSpacing:-0.5}}>₹{fmt(totalCommission)}</div>
      <div style={{fontSize:11,color:"rgba(255,255,255,0.4)",marginTop:4}}>Only fully adjusted bills · rate varies by product/customer</div>
      <div style={{display:"flex",gap:16,marginTop:14,paddingTop:12,borderTop:"1px solid rgba(255,255,255,0.08)"}}>
        <div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>{commList.length}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Customers</div></div>
        <div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>{totalBills}</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Bills Cleared</div></div>
        <div><div style={{fontSize:18,fontWeight:900,color:"#fff"}}>{fmt(totalMeters)} m</div><div style={{fontSize:10,color:"rgba(255,255,255,0.5)"}}>Total Meters</div></div>
      </div>
    </div>

    <DatePicker/>
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
function MastersTab({data,onAdd,onDel,onEdit,onImportExcel,trash,onRestore,onPermanentDelete,onEmptyTrash,initialView,initialSearch}){
  const[view,setView]=useState(initialView||"customers");
  const[search,setSearch]=useState(initialSearch||"");
  const q=search.trim().toLowerCase();
  const customers=q?data.customers.filter(c=>c.name.toLowerCase().includes(q)||(c.phone||"").includes(q)||(c.city||"").toLowerCase().includes(q)):[];
  const suppliers=q?data.suppliers.filter(s=>s.name.toLowerCase().includes(q)||(s.city||"").toLowerCase().includes(q)):[];
  const products=q?data.products.filter(p=>p.name.toLowerCase().includes(q)||(p.supplierName||"").toLowerCase().includes(q)):[];
  const searchPlaceholder=view==="customers"?"🔍 Search customer name, phone, city…":view==="suppliers"?"🔍 Search supplier name, city…":view==="products"?"🔍 Search product name, supplier…":"";
  return(<div>
    <div style={{background:`linear-gradient(135deg,${C.navyMid},${C.navy})`,borderRadius:12,padding:"14px 16px",marginBottom:14,border:"1px solid rgba(232,201,126,0.3)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.gold,marginBottom:6}}>Import from Excel</div>
      <div style={{fontSize:12,color:"rgba(255,255,255,0.6)",marginBottom:10}}>All masters auto-created from your Excel file.</div>
      <button onClick={onImportExcel} style={{background:C.gold,color:C.navy,border:"none",borderRadius:10,padding:"12px 20px",fontSize:14,fontWeight:800,cursor:"pointer",width:"100%",minHeight:46}}>Import Excel File (.xlsx)</button>
    </div>
    <SegCtrl options={[{v:"customers",l:`Customers (${data.customers.length})`},{v:"suppliers",l:`Suppliers (${data.suppliers.length})`},{v:"products",l:`Products (${data.products.length})`},{v:"trash",l:`🗑️ Trash (${trash.length})`}]} val={view} onChange={setView}/>
    {view!=="trash"&&<input placeholder={searchPlaceholder} value={search} onChange={e=>setSearch(e.target.value)} style={{...IS,margin:"10px 0"}}/>}
    {view==="customers"&&<>
      <div style={{margin:"2px 0 8px"}}><Btn color={C.navy} onClick={()=>onAdd({type:"customer"})}>+ Add Customer</Btn></div>
      {customers.length===0&&<Empty text={q?"No customers match your search.":`Search above to find a customer to edit (${data.customers.length} total).`}/>}
      {customers.map(c=>(<div key={c.id} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{c.name}</B><span style={{fontSize:11,background:"#F0F4F8",padding:"3px 9px",borderRadius:10,color:C.muted,fontWeight:600}}>{c.type}</span></Row>
        {c.phone&&<Mute>{c.phone}</Mute>}{c.city&&<Mute>{c.city}{c.state?`, ${c.state}`:""}{c.state&&normName(c.state)!==normName(SELLER_STATE)?" (IGST)":""}</Mute>}{c.gstin&&<Mute>GST: {c.gstin}</Mute>}
        <Mute>Credit Days: {c.creditDays||"30"} days</Mute>
        {c.creditLimit&&+c.creditLimit>0&&<Mute>Credit Limit: ₹{fmt(+c.creditLimit)}</Mute>}
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>onEdit("customer",c)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Edit</button>
          <button onClick={()=>onDel("customers",c.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>
      </div>))}
    </>}
    {view==="suppliers"&&<>
      <div style={{margin:"2px 0 8px"}}><Btn color={C.navy} onClick={()=>onAdd({type:"supplier"})}>+ Add Supplier</Btn></div>
      {suppliers.length===0&&<Empty text={q?"No suppliers match your search.":`Search above to find a supplier to edit (${data.suppliers.length} total).`}/>}
      {suppliers.map(s=>(<div key={s.id} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <B>{s.name}</B>{s.phone&&<Mute>{s.phone}</Mute>}{s.city&&<Mute>{s.city}</Mute>}
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>onEdit("supplier",s)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Edit</button>
          <button onClick={()=>onDel("suppliers",s.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>
      </div>))}
    </>}
    {view==="products"&&<>
      <div style={{margin:"2px 0 8px"}}><Btn color={C.navy} onClick={()=>onAdd({type:"product"})}>+ Add Product</Btn></div>
      {products.length===0&&<Empty text={q?"No products match your search.":`Search above to find a product to edit (${data.products.length} total).`}/>}
      {products.map(p=>(<div key={p.id} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <Row><B>{p.name}</B><span style={{fontSize:11,color:C.muted}}>{p.unit}</span></Row>
        {p.supplierName&&<Mute>{p.supplierName}</Mute>}
        <Mute>{p.hsn?`HSN: ${p.hsn} · `:""}GST: {p.gstRate||"5"}% · Comm: ₹{p.commissionRate||"1.5"}/m</Mute>
        <div style={{display:"flex",gap:8,marginTop:10}}>
          <button onClick={()=>onEdit("product",p)} style={{background:"#EAF4FC",color:C.blue,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Edit</button>
          <button onClick={()=>onDel("products",p.id)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Delete</button>
        </div>
      </div>))}
    </>}
    {view==="trash"&&<>
      <div style={{background:"#FEF9E7",borderRadius:10,padding:"10px 13px",margin:"10px 0",fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
        🗑️ Deleted records are kept here (last 20) so you can recover them. Restore to bring a record back, or delete forever to remove it permanently.
      </div>
      {trash.length>0&&<button onClick={onEmptyTrash} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:9,padding:"9px 14px",fontSize:12,fontWeight:700,cursor:"pointer",marginBottom:10}}>🗑️ Empty Trash</button>}
      {trash.length===0&&<Empty text="Trash is empty."/>}
      {trash.map(entry=>{
        const{type,title,sub}=describeTrashEntry(entry);
        return(<div key={entry.trashId} style={{background:C.card,borderRadius:13,padding:"13px 15px",marginBottom:10,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",borderLeft:`4px solid ${C.muted}`}}>
          <Row><div><span style={{fontSize:10,background:"#F0F4F8",padding:"3px 9px",borderRadius:10,color:C.muted,fontWeight:700}}>{type}</span></div><span style={{fontSize:11,color:C.muted}}>{timeAgo(entry.deletedAt)}</span></Row>
          <B style={{marginTop:6,display:"block"}}>{title}</B>
          {sub&&<Mute>{sub}</Mute>}
          <div style={{display:"flex",gap:8,marginTop:10}}>
            <button onClick={()=>onRestore(entry.trashId)} style={{background:"#E8F8F5",color:C.teal,border:"1px solid "+C.teal,borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>♻️ Restore</button>
            <button onClick={()=>onPermanentDelete(entry.trashId)} style={{background:"#FEE8E8",color:C.red,border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️ Delete Forever</button>
          </div>
        </div>);
      })}
    </>}
  </div>);
}

// ─── REPORTS TAB ─────────────────────────────────────────────────
function ReportsTab({data,tradingOut,tots,generatePDF,onAddPay,initialRep,initialCustomer}){
  const[rep,setRep]=useState(initialRep||"sales");
  const[fy,setFY]=useState(ALL_FY);
  const[customFrom,setCustomFrom]=useState("");
  const[customTo,setCustomTo]=useState("");
  const useCustomRange=!!(customFrom&&customTo);
  const todayMs=new Date().setHours(0,0,0,0);
  const allFYs=useMemo(()=>getAvailableFYs(data.tradingSales),[data]);

  const inCustomRange=(dateStr)=>{
    const d=new Date(dateStr).setHours(0,0,0,0);
    return d>=new Date(customFrom).setHours(0,0,0,0)&&d<=new Date(customTo).setHours(0,0,0,0);
  };
  const filteredSales=useMemo(()=>{
    if(useCustomRange)return data.tradingSales.filter(s=>inCustomRange(s.date));
    return filterByFY(data.tradingSales,fy);
  },[data,fy,customFrom,customTo]);
  const filteredPay=useMemo(()=>{
    if(useCustomRange)return data.tradingPayments.filter(p=>inCustomRange(p.date));
    if(fy===ALL_FY)return data.tradingPayments;
    const{start,end}=getFYRange(fy);
    return data.tradingPayments.filter(p=>{const d=new Date(p.date);return d>=start&&d<=end;});
  },[data,fy,customFrom,customTo]);

  const FYPicker=()=>(
    <div style={{marginBottom:12}}>
      <div style={{fontSize:10.5,color:C.muted,fontWeight:600,marginBottom:5,textTransform:"uppercase"}}>Filter by Date</div>
      <div style={{display:"flex",overflowX:"auto",gap:5,scrollbarWidth:"none",marginBottom:customFrom||customTo?8:0}}>
        {allFYs.map(f=>(
          <button key={f} onClick={()=>{setFY(f);setCustomFrom("");setCustomTo("");}} style={{flex:"0 0 auto",padding:"7px 12px",borderRadius:16,fontSize:11.5,fontWeight:!useCustomRange&&fy===f?700:500,border:`1.5px solid ${!useCustomRange&&fy===f?C.navy:C.border}`,background:!useCustomRange&&fy===f?C.navy:"#fff",color:!useCustomRange&&fy===f?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>
            {f===ALL_FY?"All":f}
          </button>
        ))}
        <button onClick={()=>{if(!customFrom)setCustomFrom(today());if(!customTo)setCustomTo(today());}} style={{flex:"0 0 auto",padding:"7px 12px",borderRadius:16,fontSize:11.5,fontWeight:useCustomRange?700:500,border:`1.5px solid ${useCustomRange?C.navy:C.border}`,background:useCustomRange?C.navy:"#fff",color:useCustomRange?C.gold:"#666",cursor:"pointer",whiteSpace:"nowrap"}}>📅 Custom Range</button>
      </div>
      {(customFrom||customTo||useCustomRange)&&<div style={{display:"flex",gap:8,alignItems:"center"}}>
        <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)} style={{...IS,padding:"8px 10px",fontSize:12.5}}/>
        <span style={{fontSize:11,color:C.muted}}>to</span>
        <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)} style={{...IS,padding:"8px 10px",fontSize:12.5}}/>
        <button onClick={()=>{setCustomFrom("");setCustomTo("");}} style={{background:"#F0F4F8",border:"none",borderRadius:8,padding:"8px 10px",fontSize:11,color:C.muted,cursor:"pointer",flexShrink:0}}>✕</button>
      </div>}
    </div>
  );

  const[salesView,setSalesView]=useState("date"); // date | customer | product
  const[custFilter,setCustFilter]=useState("");
  const salesRows=useMemo(()=>{
    const q=custFilter.trim().toLowerCase();
    return q?filteredSales.filter(s=>(s.customerName||"").toLowerCase().includes(q)):filteredSales;
  },[filteredSales,custFilter]);

  const doPDFSales=()=>{
    const pdfStyle=`<style>table{width:100%;border-collapse:collapse;}th{background:#0F1923;color:#E8C97E;padding:8px 10px;text-align:left;font-size:12px;border:1px solid #0F1923;}td{padding:7px 10px;border:1px solid #ddd;font-size:12px;}tr:nth-child(even){background:#f9f9f9;}.tot td{background:#0F1923;color:#E8C97E;font-weight:bold;border:none;}.grp-hdr td{background:#E8EDF4;font-weight:800;color:#0F1923;}</style>`;
    const total=salesRows.reduce((a,s)=>a+(+s.amount||0),0);
    if(salesView==="date"){
      const rows=[...salesRows].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(s=>`<tr><td>${s.customerName}</td><td>${fmtD(s.date)}</td><td>${s.billNo||"—"}</td><td>${s.productName||"—"}</td><td style="text-align:right">${fmt(s.meters)} m</td><td style="text-align:right">₹${fmt(s.rate)}</td><td style="text-align:right;font-weight:bold">₹${fmt(s.amount)}</td></tr>`).join("");
      generatePDF(`Sales Register — Date-wise${fy!==ALL_FY?" (FY "+fy+")":""}${custFilter?" — "+custFilter:""}`,`${pdfStyle}<table><thead><tr><th>Customer</th><th>Bill Date</th><th>Bill No</th><th>Product</th><th style="text-align:right">Qty (m)</th><th style="text-align:right">Rate</th><th style="text-align:right">Amount</th></tr></thead><tbody>${rows}<tr class="tot"><td colspan="6">TOTAL</td><td style="text-align:right">₹${fmt(total)}</td></tr></tbody></table>`);
    } else if(salesView==="customer"){
      const custMap={};salesRows.forEach(s=>{if(!custMap[s.customerName])custMap[s.customerName]={name:s.customerName,sales:[],total:0,qty:0};custMap[s.customerName].sales.push(s);custMap[s.customerName].total+=+s.amount||0;custMap[s.customerName].qty+=+s.meters||0;});
      const blocks=Object.values(custMap).sort((a,b)=>b.total-a.total).map(c=>{
        const rows=c.sales.sort((a,b)=>new Date(a.date)-new Date(b.date)).map(s=>`<tr><td>${fmtD(s.date)}</td><td>${s.billNo||"—"}</td><td>${s.productName||"—"}</td><td style="text-align:right">${fmt(s.meters)} m</td><td style="text-align:right">₹${fmt(s.rate)}</td><td style="text-align:right;font-weight:bold">₹${fmt(s.amount)}</td></tr>`).join("");
        return `<tr class="grp-hdr"><td colspan="5"><b>${c.name}</b> — ${fmt(c.qty)} m</td><td style="text-align:right;font-weight:bold">₹${fmt(c.total)}</td></tr>${rows}`;
      }).join("");
      generatePDF(`Sales Register — Customer-wise${fy!==ALL_FY?" (FY "+fy+")":""}`,`${pdfStyle}<table><thead><tr><th>Date</th><th>Bill No</th><th>Product</th><th style="text-align:right">Qty (m)</th><th style="text-align:right">Rate</th><th style="text-align:right">Amount</th></tr></thead><tbody>${blocks}<tr class="tot"><td colspan="5">TOTAL</td><td style="text-align:right">₹${fmt(total)}</td></tr></tbody></table>`);
    } else {
      const prodMap={};salesRows.forEach(s=>{const p=s.productName||"Unknown";if(!prodMap[p])prodMap[p]={name:p,sales:[],total:0,meters:0};prodMap[p].sales.push(s);prodMap[p].total+=+s.amount||0;prodMap[p].meters+=+s.meters||0;});
      const blocks=Object.values(prodMap).sort((a,b)=>b.total-a.total).map(p=>{
        const rows=p.sales.sort((a,b)=>new Date(a.date)-new Date(b.date)).map(s=>`<tr><td>${fmtD(s.date)}</td><td>${s.billNo||"—"}</td><td>${s.customerName}</td><td style="text-align:right">${fmt(s.meters)} m</td><td style="text-align:right">₹${fmt(s.rate)}</td><td style="text-align:right;font-weight:bold">₹${fmt(s.amount)}</td></tr>`).join("");
        return `<tr class="grp-hdr"><td colspan="4"><b>${p.name}</b> — ${fmt(p.meters)} m</td><td style="text-align:right;font-weight:bold">₹${fmt(p.total)}</td></tr>${rows}`;
      }).join("");
      generatePDF(`Sales Register — Product-wise${fy!==ALL_FY?" (FY "+fy+")":""}`,`${pdfStyle}<table><thead><tr><th>Date</th><th>Bill No</th><th>Customer</th><th style="text-align:right">Qty (m)</th><th style="text-align:right">Rate</th><th style="text-align:right">Amount</th></tr></thead><tbody>${blocks}<tr class="tot"><td colspan="4">TOTAL</td><td style="text-align:right">₹${fmt(total)}</td></tr></tbody></table>`);
    }
  };

  const totSales=salesRows.reduce((a,s)=>a+(+s.amount||0),0);
  const totQty=salesRows.reduce((a,s)=>a+(+s.meters||0),0);

  return(<div>
    <div style={{display:"flex",overflowX:"auto",gap:8,marginBottom:12,scrollbarWidth:"none"}}>
      {[{v:"sales",l:"Sales Register"},{v:"outstanding",l:"Outstanding"},{v:"commission",l:"Commission"},{v:"ageing",l:"Ageing"},{v:"ledger",l:"Ledger"}].map(r=>(<button key={r.v} onClick={()=>setRep(r.v)} style={{flex:"0 0 auto",padding:"10px 16px",borderRadius:20,fontSize:13,fontWeight:rep===r.v?700:500,border:`1.5px solid ${rep===r.v?C.navy:C.border}`,background:rep===r.v?C.navy:"#fff",color:rep===r.v?C.gold:"#666",cursor:"pointer",minHeight:42}}>{r.l}</button>))}
    </div>

    {rep==="sales"&&<>
      <FYPicker/>
      {/* View selector */}
      <div style={{display:"flex",background:"#fff",borderRadius:11,overflow:"hidden",border:`1px solid ${C.border}`,marginBottom:10}}>
        {[{v:"date",l:"📅 Date-wise"},{v:"customer",l:"👥 Customer-wise"},{v:"product",l:"📦 Product-wise"}].map(o=>(
          <button key={o.v} onClick={()=>setSalesView(o.v)} style={{flex:1,padding:"11px 4px",fontSize:11.5,fontWeight:salesView===o.v?700:500,color:salesView===o.v?"#E8C97E":C.muted,background:salesView===o.v?C.navy:"transparent",border:"none",cursor:"pointer",minHeight:44}}>{o.l}</button>
        ))}
      </div>
      <input placeholder="🔍 Filter by customer name…" value={custFilter} onChange={e=>setCustFilter(e.target.value)} style={{...IS,marginBottom:10}}/>
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        <button onClick={()=>exportCSV([["Date","Bill No","Customer","Product","Qty(m)","Rate","Amount"],...salesRows.map(s=>[fmtD(s.date),s.billNo||"",s.customerName,s.productName,s.meters,s.rate,s.amount])],`SalesRegister_${salesView}_${today()}.csv`)} style={{background:C.green,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>📥 CSV</button>
        <button onClick={doPDFSales} style={{background:C.red,color:"#fff",border:"none",borderRadius:9,padding:"9px 15px",fontSize:12.5,fontWeight:700,cursor:"pointer",minHeight:40}}>🖨️ PDF</button>
      </div>
      <div style={{background:C.card,borderRadius:12,padding:14,marginBottom:12,borderLeft:`4px solid ${C.blue}`}}>
        <Mute>Total Sales {fy!==ALL_FY?`— FY ${fy}`:""}{custFilter?` — ${custFilter}`:""} ({salesRows.length} bills)</Mute>
        <div style={{display:"flex",alignItems:"baseline",gap:14,flexWrap:"wrap"}}>
          <div style={{fontSize:22,fontWeight:900,color:C.blue}}>₹{fmt(totSales)}</div>
          <div style={{fontSize:13,fontWeight:700,color:C.teal}}>{fmt(totQty)} m <span style={{fontSize:10.5,color:C.muted,fontWeight:600}}>total qty</span></div>
        </div>
      </div>

      {/* Date-wise — 7 column table: Customer, Bill Date, Bill No, Product, Qty, Rate, Amount */}
      {salesView==="date"&&(()=>{
        const rows=[...salesRows].sort((a,b)=>new Date(b.date)-new Date(a.date));
        if(rows.length===0)return <Empty text="No sales found."/>;
        const th={padding:"9px 10px",textAlign:"left",fontSize:10.5,color:C.gold,background:C.navy,whiteSpace:"nowrap",position:"sticky",top:0};
        const td={padding:"8px 10px",fontSize:12,color:C.navy,whiteSpace:"nowrap",borderBottom:`1px solid ${C.border}`};
        return(
          <div style={{background:C.card,borderRadius:12,boxShadow:"0 1px 8px rgba(0,0,0,0.06)",overflowX:"auto",marginBottom:10}}>
            <table style={{borderCollapse:"collapse",width:"100%",minWidth:760}}>
              <thead><tr>
                <th style={{...th,borderRadius:"12px 0 0 0"}}>Customer</th>
                <th style={th}>Bill Date</th>
                <th style={th}>Bill No</th>
                <th style={th}>Product</th>
                <th style={{...th,textAlign:"right"}}>Qty (m)</th>
                <th style={{...th,textAlign:"right"}}>Rate</th>
                <th style={{...th,textAlign:"right",borderRadius:"0 12px 0 0"}}>Amount</th>
              </tr></thead>
              <tbody>
                {rows.map((s,i)=>(
                  <tr key={s.id} style={{background:i%2===0?"#fff":"#FAFBFC"}}>
                    <td style={{...td,fontWeight:700}}>{s.customerName}</td>
                    <td style={td}>{fmtD(s.date)}</td>
                    <td style={td}>{s.billNo||"—"}</td>
                    <td style={td}>{s.productName||"—"}</td>
                    <td style={{...td,textAlign:"right"}}>{fmt(s.meters)}</td>
                    <td style={{...td,textAlign:"right"}}>₹{fmt(s.rate)}</td>
                    <td style={{...td,textAlign:"right",fontWeight:900,color:C.blue}}>₹{fmt(s.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      })()}

      {/* Customer-wise */}
      {salesView==="customer"&&(()=>{
        const custMap={};
        salesRows.forEach(s=>{if(!custMap[s.customerName])custMap[s.customerName]={name:s.customerName,sales:[],total:0,qty:0};custMap[s.customerName].sales.push(s);custMap[s.customerName].total+=+s.amount||0;custMap[s.customerName].qty+=+s.meters||0;});
        if(Object.keys(custMap).length===0)return <Empty text="No sales found."/>;
        return Object.values(custMap).sort((a,b)=>b.total-a.total).map(c=>(
          <div key={c.name} style={{background:C.card,borderRadius:13,marginBottom:12,overflow:"hidden",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
            <div style={{background:`linear-gradient(90deg,${C.navyMid},${C.navy})`,padding:"11px 14px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <B style={{color:"#fff",fontSize:13}}>{c.name}</B>
                <span style={{fontWeight:900,color:C.gold}}>₹{fmt(c.total)}</span>
              </div>
              <div style={{fontSize:11,color:"rgba(255,255,255,0.65)",marginTop:2}}>{fmt(c.qty)} m total qty · {c.sales.length} bills</div>
            </div>
            {c.sales.sort((a,b)=>new Date(b.date)-new Date(a.date)).map((s,i)=>(
              <div key={s.id} style={{padding:"9px 14px",borderBottom:`1px solid ${C.border}`,background:i%2===0?"#fff":"#FAFBFC",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:12.5,fontWeight:600,color:C.navy}}>{s.billNo||"—"} · {fmtD(s.date)}</div>
                  <div style={{fontSize:11.5,color:C.muted}}>{s.productName} · {fmt(s.meters)} m @ ₹{fmt(s.rate)}/m</div>
                </div>
                <span style={{fontWeight:800,color:C.blue,fontSize:13}}>₹{fmt(s.amount)}</span>
              </div>
            ))}
          </div>
        ));
      })()}

      {/* Product-wise */}
      {salesView==="product"&&(()=>{
        const prodMap={};
        salesRows.forEach(s=>{const p=s.productName||"Unknown";if(!prodMap[p])prodMap[p]={name:p,sales:[],total:0,meters:0};prodMap[p].sales.push(s);prodMap[p].total+=+s.amount||0;prodMap[p].meters+=+s.meters||0;});
        if(Object.keys(prodMap).length===0)return <Empty text="No sales found."/>;
        return Object.values(prodMap).sort((a,b)=>b.total-a.total).map(p=>(
          <div key={p.name} style={{background:C.card,borderRadius:13,marginBottom:12,overflow:"hidden",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
            <div style={{background:`linear-gradient(90deg,${C.teal},#117A65)`,padding:"11px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><B style={{color:"#fff",fontSize:13}}>{p.name}</B><div style={{fontSize:11,color:"rgba(255,255,255,0.7)",marginTop:2}}>{fmt(p.meters)} meters total</div></div>
              <span style={{fontWeight:900,color:"#fff"}}>₹{fmt(p.total)}</span>
            </div>
            {p.sales.sort((a,b)=>new Date(b.date)-new Date(a.date)).map((s,i)=>(
              <div key={s.id} style={{padding:"9px 14px",borderBottom:`1px solid ${C.border}`,background:i%2===0?"#fff":"#FAFBFC",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:12.5,fontWeight:600,color:C.navy}}>{s.customerName}</div>
                  <div style={{fontSize:11.5,color:C.muted}}>{s.billNo||"—"} · {fmtD(s.date)} · {fmt(s.meters)} m @ ₹{fmt(s.rate)}/m</div>
                </div>
                <span style={{fontWeight:800,color:C.teal,fontSize:13}}>₹{fmt(s.amount)}</span>
              </div>
            ))}
          </div>
        ));
      })()}
    </>}

    {rep==="outstanding"&&<OutstandingTab tradingOut={tradingOut} data={data} onAddPay={onAddPay} generatePDF={generatePDF}/>}

    {rep==="commission"&&<CommissionTab data={data} generatePDF={generatePDF}/>}

    {rep==="ageing"&&<AgingTab data={data} generatePDF={generatePDF}/>}

    {rep==="ledger"&&<LedgerTab data={data} generatePDF={generatePDF} initialCustomer={initialCustomer}/>}
  </div>);
}

// ─── GLOBAL SEARCH ──────────────────────────────────────────────
const SEARCH_KIND_LABEL={customer:"Customer",sale:"Sale",payment:"Payment",debit:"Debit Note",credit:"Credit Note",product:"Product",supplier:"Supplier",enquiry:"Enquiry"};
const SEARCH_KIND_COLOR={customer:C.purple,sale:C.blue,payment:C.green,debit:C.red,credit:C.teal,product:C.orange,supplier:"#8E5A2E",enquiry:C.navy};
function GlobalSearchModal({data,onClose,onSelect}){
  const[q,setQ]=useState("");
  const query=q.trim().toLowerCase();

  const results=useMemo(()=>{
    if(query.length<2)return[];
    const out=[];
    data.customers.forEach(c=>{if(c.name.toLowerCase().includes(query))out.push({kind:"customer",icon:"👤",title:c.name,sub:[c.phone,c.city].filter(Boolean).join(" · "),name:c.name});});
    data.tradingSales.forEach(s=>{if((s.customerName||"").toLowerCase().includes(query)||(s.billNo||"").toLowerCase().includes(query)||(s.productName||"").toLowerCase().includes(query))out.push({kind:"sale",icon:"🏪",title:`${s.customerName} — ₹${fmt(s.amount)}`,sub:`${s.billNo||"—"} · ${s.productName||""} · ${fmtD(s.date)}`,customerName:s.customerName});});
    data.tradingPayments.forEach(p=>{if((p.customerName||"").toLowerCase().includes(query)||(p.billNo||"").toLowerCase().includes(query))out.push({kind:"payment",icon:"✅",title:`${p.customerName} — ₹${fmt(p.amount)}`,sub:`${p.mode||""} · ${fmtD(p.date)}`,customerName:p.customerName});});
    (data.debitNotes||[]).forEach(n=>{if((n.customerName||"").toLowerCase().includes(query)||(n.noteNo||"").toLowerCase().includes(query))out.push({kind:"debit",icon:"⚠️",title:`${n.customerName} — ₹${fmt(n.amount)}`,sub:`${n.noteNo||"DN"} · ${fmtD(n.date)}`,customerName:n.customerName});});
    (data.creditNotes||[]).forEach(n=>{if((n.customerName||"").toLowerCase().includes(query)||(n.noteNo||"").toLowerCase().includes(query))out.push({kind:"credit",icon:"↩️",title:`${n.customerName} — ₹${fmt(n.amount)}`,sub:`${n.noteNo||"CN"} · ${fmtD(n.date)}`,customerName:n.customerName});});
    data.products.forEach(p=>{if(p.name.toLowerCase().includes(query))out.push({kind:"product",icon:"📦",title:p.name,sub:[p.supplierName,p.unit].filter(Boolean).join(" · "),name:p.name});});
    data.suppliers.forEach(s=>{if(s.name.toLowerCase().includes(query))out.push({kind:"supplier",icon:"🚚",title:s.name,sub:[s.phone,s.city].filter(Boolean).join(" · "),name:s.name});});
    (data.enquiries||[]).forEach(e=>{if((e.customerName||"").toLowerCase().includes(query)||(e.description||"").toLowerCase().includes(query))out.push({kind:"enquiry",icon:"📝",title:e.customerName,sub:(e.description||"").slice(0,55),id:e.id});});
    return out.slice(0,40);
  },[query,data]);

  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:300,display:"flex",flexDirection:"column",justifyContent:"flex-end"}} onClick={onClose}>
    <div style={{background:"#fff",width:"100%",maxWidth:600,margin:"0 auto",borderRadius:"20px 20px 0 0",padding:"18px 16px calc(env(safe-area-inset-bottom,0px) + 20px)",maxHeight:"88vh",display:"flex",flexDirection:"column"}} onClick={e=>e.stopPropagation()}>
      <Row style={{marginBottom:14,flexShrink:0}}>
        <B style={{fontSize:16.5,color:C.navy}}>🔍 Search Everything</B>
        <button onClick={onClose} style={{background:"#F0F4F8",border:"none",borderRadius:20,width:38,height:38,fontSize:16,cursor:"pointer",color:"#555",flexShrink:0}}>✕</button>
      </Row>
      <input autoFocus value={q} onChange={e=>setQ(e.target.value)} placeholder="Customer, bill no, product, note no…" style={{...IS,flexShrink:0}}/>
      <div style={{overflowY:"auto",marginTop:12,flex:1}}>
        {query.length<2&&<div style={{textAlign:"center",color:"#bbb",fontSize:13,padding:"40px 10px"}}>Type at least 2 characters to search across customers, bills, payments, notes, products, suppliers &amp; enquiries.</div>}
        {query.length>=2&&results.length===0&&<Empty text="No matches found."/>}
        {results.map((r,i)=>(
          <button key={i} onClick={()=>onSelect(r)} style={{display:"flex",gap:10,alignItems:"center",width:"100%",textAlign:"left",background:"#F8FAFC",border:"none",borderRadius:11,padding:"11px 13px",marginBottom:8,cursor:"pointer"}}>
            <span style={{fontSize:18,flexShrink:0}}>{r.icon}</span>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:13,fontWeight:700,color:C.navy,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.title}</div>
              {r.sub&&<div style={{fontSize:11,color:C.muted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",marginTop:1}}>{r.sub}</div>}
            </div>
            <span style={{fontSize:9,color:SEARCH_KIND_COLOR[r.kind],textTransform:"uppercase",fontWeight:800,flexShrink:0,background:"#fff",padding:"3px 7px",borderRadius:8}}>{SEARCH_KIND_LABEL[r.kind]}</span>
          </button>
        ))}
      </div>
    </div>
  </div>);
}

// ─── SETTINGS TAB (Company Settings) ────────────────────────────
function SettingsTab({data,onSave,onSaveApp}){
  const co=data.companySettings||{};
  const[f,sf]=useState({
    name:co.name||"Navkar Fabrics",address:co.address||"",gstin:co.gstin||"",phone:co.phone||"",email:co.email||"",
    state:co.state||"Delhi",bankName:co.bankName||"",bankAccount:co.bankAccount||"",bankIFSC:co.bankIFSC||"",
    terms:co.terms||"",
  });
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const dirty=JSON.stringify(f)!==JSON.stringify({name:co.name||"Navkar Fabrics",address:co.address||"",gstin:co.gstin||"",phone:co.phone||"",email:co.email||"",state:co.state||"Delhi",bankName:co.bankName||"",bankAccount:co.bankAccount||"",bankIFSC:co.bankIFSC||"",terms:co.terms||""});

  const ap=data.appSettings||{};
  const[af,saf]=useState({
    defaultGstRate:ap.defaultGstRate||"5",defaultCommissionRate:ap.defaultCommissionRate||"1.5",
    defaultCreditDays:ap.defaultCreditDays||"30",backupReminderDays:ap.backupReminderDays||"7",trashRetention:ap.trashRetention||"20",
  });
  const as=(k,v)=>saf(p=>({...p,[k]:v}));
  const appDirty=JSON.stringify(af)!==JSON.stringify({defaultGstRate:ap.defaultGstRate||"5",defaultCommissionRate:ap.defaultCommissionRate||"1.5",defaultCreditDays:ap.defaultCreditDays||"30",backupReminderDays:ap.backupReminderDays||"7",trashRetention:ap.trashRetention||"20"});

  return(<div>
    <SecTitle>🏢 Company Settings</SecTitle>
    <div style={{background:"#FEF9E7",borderRadius:10,padding:"10px 13px",marginBottom:16,fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
      These details appear on every Invoice, Report header, and WhatsApp message sent from the app.
    </div>

    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>Business Details</div>
      <F label="Company Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Navkar Fabrics" style={IS}/></F>
      <F label="Address"><textarea value={f.address} onChange={e=>s("address",e.target.value)} placeholder="Full business address" style={{...IS,minHeight:60,resize:"vertical"}}/></F>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <F label="Phone"><input value={f.phone} onChange={e=>s("phone",e.target.value)} placeholder="9911112105" style={IS}/></F>
        <F label="Email"><input value={f.email} onChange={e=>s("email",e.target.value)} placeholder="you@example.com" style={IS}/></F>
      </div>
      <F label="GSTIN"><input value={f.gstin} onChange={e=>s("gstin",e.target.value)} placeholder="09AARPJ1728P1ZZ" style={IS}/></F>
      <F label="State (for GST — determines CGST+SGST vs IGST on every invoice)">
        <select value={f.state} onChange={e=>s("state",e.target.value)} style={IS}>
          {INDIA_STATES.map(st=><option key={st} value={st}>{st}</option>)}
        </select>
      </F>
    </div>

    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>Bank Details (shown on invoices)</div>
      <F label="Bank Name"><input value={f.bankName} onChange={e=>s("bankName",e.target.value)} placeholder="Canara Bank" style={IS}/></F>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <F label="Account No."><input value={f.bankAccount} onChange={e=>s("bankAccount",e.target.value)} placeholder="3768201000242" style={IS}/></F>
        <F label="IFSC"><input value={f.bankIFSC} onChange={e=>s("bankIFSC",e.target.value)} placeholder="CNRB0003768" style={IS}/></F>
      </div>
    </div>

    <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>Terms &amp; Conditions (one per line, shown on invoices)</div>
      <F label=""><textarea value={f.terms} onChange={e=>s("terms",e.target.value)} placeholder={"Goods once sold will not be taken back.\nInterest @ 18% p.a. will be charged if payment is delayed.\nSubject to Delhi jurisdiction only."} style={{...IS,minHeight:90,resize:"vertical"}}/></F>
    </div>

    <SaveBtn color={dirty?C.navy:C.muted} onClick={()=>{if(!f.name.trim())return alert("Company name required");onSave(f);}}>{dirty?"💾 Save Company Settings":"✅ Saved"}</SaveBtn>

    <div style={{marginTop:28}}>
      <SecTitle>⚙️ App Preferences</SecTitle>
      <div style={{background:"#FEF9E7",borderRadius:10,padding:"10px 13px",marginBottom:16,fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
        These are just the app's defaults — you can still change any of them for an individual customer, product, or sale whenever you need to.
      </div>
      <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>Defaults for New Records</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <F label="Default GST %">
            <select value={af.defaultGstRate} onChange={e=>as("defaultGstRate",e.target.value)} style={IS}>
              <option value="0">0%</option><option value="5">5%</option><option value="12">12%</option><option value="18">18%</option>
            </select>
          </F>
          <F label="Default Commission ₹/m"><input type="number" value={af.defaultCommissionRate} onChange={e=>as("defaultCommissionRate",e.target.value)} style={IS}/></F>
        </div>
        <F label="Default Credit Days (for new customers)">
          <select value={af.defaultCreditDays} onChange={e=>as("defaultCreditDays",e.target.value)} style={IS}>
            <option value="0">0 Days (Immediate)</option><option value="15">15 Days</option><option value="30">30 Days</option><option value="45">45 Days</option><option value="60">60 Days</option><option value="90">90 Days</option>
          </select>
        </F>
      </div>
      <div style={{background:C.card,borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:12}}>App Behaviour</div>
        <F label="Remind me to backup after (days)"><input type="number" value={af.backupReminderDays} onChange={e=>as("backupReminderDays",e.target.value)} style={IS}/></F>
        <F label="Keep last N deleted records in Trash"><input type="number" value={af.trashRetention} onChange={e=>as("trashRetention",e.target.value)} style={IS}/></F>
      </div>
      <SaveBtn color={appDirty?C.navy:C.muted} onClick={()=>onSaveApp(af)}>{appDirty?"💾 Save App Preferences":"✅ Saved"}</SaveBtn>
    </div>
  </div>);
}

// ─── IMPORT PREVIEW MODALS ─────────────────────────────────────────
function PreviewStat({label,val,color}){
  return(<div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",borderLeft:`3px solid ${color}`}}>
    <div style={{fontSize:9.5,color:C.muted,fontWeight:600,textTransform:"uppercase"}}>{label}</div>
    <div style={{fontSize:15,fontWeight:900,color}}>{val}</div>
  </div>);
}
function ExcelImportPreviewModal({preview,onConfirm,onCancel}){
  const p=preview.parsed;
  return(<ModalBase title="Preview Excel Import" onClose={onCancel}>
    <div style={{fontSize:12.5,color:C.muted,marginBottom:14}}>📄 {preview.fileName}</div>
    <div style={{background:"#FEF9E7",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12,color:"#7D6608",border:"1px solid #F9E79F"}}>
      Review the counts below before importing. Existing customers/suppliers/products (matched by name) won't be duplicated — only new ones are added. Sales &amp; payments are always added as new records.
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
      <PreviewStat label="Sales rows" val={p.tradingSales.length} color={C.blue}/>
      <PreviewStat label="Payment rows" val={p.tradingPayments.length} color={C.green}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:16}}>
      <PreviewStat label="New Customers" val={`${preview.newCustomers} / ${p.customers.length}`} color={C.purple}/>
      <PreviewStat label="New Suppliers" val={`${preview.newSuppliers} / ${p.suppliers.length}`} color={C.orange}/>
      <PreviewStat label="New Products" val={`${preview.newProducts} / ${p.products.length}`} color={C.teal}/>
    </div>
    <div style={{display:"flex",gap:8}}>
      <button onClick={onCancel} style={{flex:1,background:"#F0F4F8",color:C.navy,border:"none",borderRadius:12,padding:15,fontSize:14,fontWeight:700,cursor:"pointer",minHeight:50}}>Cancel</button>
      <button onClick={onConfirm} style={{flex:2,background:C.blue,color:"#fff",border:"none",borderRadius:12,padding:15,fontSize:14.5,fontWeight:800,cursor:"pointer",minHeight:50}}>✅ Import {p.tradingSales.length} Sales</button>
    </div>
  </ModalBase>);
}
function BackupPreviewModal({preview,currentData,onConfirm,onCancel}){
  const b=preview.parsed;
  const cnt=(arr)=>Array.isArray(arr)?arr.length:0;
  const rows=[
    {label:"Customers",cur:cnt(currentData?.customers),bak:cnt(b.customers)},
    {label:"Suppliers",cur:cnt(currentData?.suppliers),bak:cnt(b.suppliers)},
    {label:"Products",cur:cnt(currentData?.products),bak:cnt(b.products)},
    {label:"Sales",cur:cnt(currentData?.tradingSales),bak:cnt(b.tradingSales)},
    {label:"Payments",cur:cnt(currentData?.tradingPayments),bak:cnt(b.tradingPayments)},
    {label:"Debit Notes",cur:cnt(currentData?.debitNotes),bak:cnt(b.debitNotes)},
    {label:"Credit Notes",cur:cnt(currentData?.creditNotes),bak:cnt(b.creditNotes)},
    {label:"Enquiries",cur:cnt(currentData?.enquiries),bak:cnt(b.enquiries)},
  ];
  return(<ModalBase title="Preview Restore Backup" onClose={onCancel}>
    <div style={{fontSize:12.5,color:C.muted,marginBottom:14}}>📄 {preview.fileName}</div>
    <div style={{background:"#FDEDEC",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12,color:"#922B21",border:"1px solid #F5B7B1",fontWeight:600}}>
      ⚠️ This will REPLACE all current data with this backup file. A safety copy of your current data will be downloaded first. This cannot be undone.
    </div>
    <div style={{overflowX:"auto",marginBottom:16}}>
      <table style={{borderCollapse:"collapse",width:"100%",minWidth:340}}>
        <thead><tr>
          <th style={{padding:"7px 8px",textAlign:"left",fontSize:10,color:C.gold,background:C.navy,borderRadius:"8px 0 0 0"}}>Record</th>
          <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy}}>Current</th>
          <th style={{padding:"7px 8px",textAlign:"right",fontSize:10,color:C.gold,background:C.navy,borderRadius:"0 8px 0 0"}}>Backup</th>
        </tr></thead>
        <tbody>
          {rows.map((r,i)=>(<tr key={r.label} style={{background:i%2===0?"#fff":"#FAFBFC"}}>
            <td style={{padding:"7px 8px",fontSize:12,fontWeight:700,color:C.navy,borderBottom:`1px solid ${C.border}`}}>{r.label}</td>
            <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",color:C.muted,borderBottom:`1px solid ${C.border}`}}>{r.cur}</td>
            <td style={{padding:"7px 8px",fontSize:12,textAlign:"right",fontWeight:800,color:r.bak!==r.cur?C.red:C.navy,borderBottom:`1px solid ${C.border}`}}>{r.bak}</td>
          </tr>))}
        </tbody>
      </table>
    </div>
    <div style={{display:"flex",gap:8}}>
      <button onClick={onCancel} style={{flex:1,background:"#F0F4F8",color:C.navy,border:"none",borderRadius:12,padding:15,fontSize:14,fontWeight:700,cursor:"pointer",minHeight:50}}>Cancel</button>
      <button onClick={onConfirm} style={{flex:2,background:C.red,color:"#fff",border:"none",borderRadius:12,padding:15,fontSize:14.5,fontWeight:800,cursor:"pointer",minHeight:50}}>⚠️ Replace &amp; Restore</button>
    </div>
  </ModalBase>);
}

// ─── MODALS ──────────────────────────────────────────────────────
function SaleModal({data,onSave,onClose,initial}){
  const[f,sf]=useState(initial?{...initial,amount:String((+initial.amount||0)-(+initial.freight||0))}:{date:today(),billNo:"",customerName:"",productName:"",supplierName:"",meters:"",rate:"",amount:"",freight:"0",hsn:"",unit:"Mtr",gstRate:data?.appSettings?.defaultGstRate||"5",commissionRate:data?.appSettings?.defaultCommissionRate||"1.5",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const selectProd=(name)=>{
    const p=data.products.find(p=>p.name===name);
    sf(pr=>({...pr,productName:name,supplierName:p?.supplierName||pr.supplierName,hsn:p?.hsn||pr.hsn,unit:p?.unit||pr.unit,gstRate:p?.gstRate||pr.gstRate,commissionRate:p?.commissionRate||pr.commissionRate}));
  };
  useEffect(()=>{if(f.meters&&f.rate&&!initial)s("amount",(parseFloat(f.meters)*parseFloat(f.rate)).toFixed(2));},[f.meters,f.rate]);
  return(<ModalBase title={initial?"Edit Sale":"Add Trading Sale"} onClose={onClose}>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Bill No"><input value={f.billNo} onChange={e=>s("billNo",e.target.value)} placeholder="Invoice / Bill number" style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>s("customerName",v)} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="sl-c"/></F>
    {(()=>{
      const cust=data.customers.find(c=>c.name===f.customerName);
      if(!cust||!cust.creditLimit||+cust.creditLimit<=0)return null;
      const outstanding=computeCustomerOutstanding(f.customerName,data);
      const overBy=outstanding-(+cust.creditLimit);
      return(<div style={{fontSize:11.5,color:overBy>0?C.red:C.muted,marginTop:-9,marginBottom:13,fontWeight:overBy>0?700:400}}>
        {overBy>0?`⚠️ Already ₹${fmt(overBy)} over credit limit (₹${fmt(cust.creditLimit)})`:`Outstanding: ₹${fmt(outstanding)} / Limit: ₹${fmt(cust.creditLimit)}`}
      </div>);
    })()}
    <F label="Product *"><SmartInput value={f.productName} onChange={selectProd} placeholder="Product" list={data.products.map(p=>p.name)} idPrefix="sl-p"/></F>
    <F label="Supplier"><SmartInput value={f.supplierName} onChange={v=>s("supplierName",v)} placeholder="Supplier" list={data.suppliers.map(s=>s.name)} idPrefix="sl-s"/></F>
    <F label="HSN / SAC Code"><input value={f.hsn||""} onChange={e=>s("hsn",e.target.value)} placeholder="Auto-filled from product" style={IS}/></F>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <F label="Meters"><input type="number" value={f.meters} onChange={e=>s("meters",e.target.value)} placeholder="0" style={IS}/></F>
      <F label="Rate (₹/m)"><input type="number" value={f.rate} onChange={e=>s("rate",e.target.value)} placeholder="0" style={IS}/></F>
    </div>
    <F label="Product Amount (₹) — GST inclusive"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} style={{...IS,fontWeight:700}}/></F>
    <F label="Freight / Coolie Cartage (₹) — GST inclusive, charged to customer"><input type="number" value={f.freight||"0"} onChange={e=>s("freight",e.target.value)} placeholder="0" style={IS}/></F>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <F label="GST % (prices above already include this)">
        <select value={f.gstRate||"5"} onChange={e=>s("gstRate",e.target.value)} style={IS}>
          <option value="0">0% (Exempt)</option>
          <option value="5">5%</option>
          <option value="12">12%</option>
          <option value="18">18%</option>
        </select>
      </F>
      <F label="Commission ₹/m"><input type="number" value={f.commissionRate||"1.5"} onChange={e=>s("commissionRate",e.target.value)} placeholder="1.5" style={IS}/></F>
    </div>
    {(()=>{
      if(!f.amount||!f.customerName)return null;
      const cust=data.customers.find(c=>c.name===f.customerName);
      const buyerState=cust?.state||data.companySettings?.state||SELLER_STATE;
      const productAmt=+f.amount||0,freightAmt=+f.freight||0;
      const gp=calcGST(productAmt,buyerState,f.gstRate,data.companySettings?.state);
      const gf=calcGST(freightAmt,buyerState,f.gstRate,data.companySettings?.state);
      const taxable=gp.taxable+gf.taxable,cgst=gp.cgst+gf.cgst,sgst=gp.sgst+gf.sgst,igst=gp.igst+gf.igst;
      const total=productAmt+freightAmt;
      return(<div style={{background:"#F8FAFC",borderRadius:10,padding:"10px 12px",marginBottom:13,fontSize:11.5,color:C.navy,lineHeight:1.7}}>
        <div>Taxable Value: <b>₹{fmt(taxable)}</b>{freightAmt>0?` (product ₹${fmt(gp.taxable)} + freight ₹${fmt(gf.taxable)})`:""}</div>
        {gp.sameState?<div>CGST {(gp.rate/2)}% + SGST {(gp.rate/2)}%: <b>₹{fmt(cgst)} + ₹{fmt(sgst)}</b></div>:<div>IGST {gp.rate}% ({buyerState} — inter-state): <b>₹{fmt(igst)}</b></div>}
        <div style={{color:C.muted,fontSize:10.5,marginTop:2}}>Invoice Total: ₹{fmt(total)}</div>
      </div>);
    })()}
    <F label="Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Optional" style={IS}/></F>
    <SaveBtn color={C.blue} onClick={()=>{
      if(!f.customerName||!f.amount)return alert("Customer and Amount required");
      if(!warnIfDuplicateBillNo(f.billNo,data.tradingSales,initial?.id))return;
      if(!warnIfNearDuplicateCustomer(f.customerName,data.customers))return;
      const cust=data.customers.find(c=>c.name===f.customerName);
      if(cust&&cust.creditLimit&&+cust.creditLimit>0){
        const limit=+cust.creditLimit;
        const currentOutstanding=computeCustomerOutstanding(f.customerName,data);
        const oldAmount=initial?(+initial.amount||0):0;
        const totalAmount=(+f.amount||0)+(+f.freight||0);
        const projected=currentOutstanding-oldAmount+totalAmount;
        if(projected>limit){
          const ok=window.confirm(`⚠️ Credit Limit Alert!\n\n${f.customerName}'s credit limit is ₹${fmt(limit)}.\nProjected outstanding after this sale: ₹${fmt(projected)} — over by ₹${fmt(projected-limit)}.\n\nTap OK to save anyway, or Cancel to review.`);
          if(!ok)return;
        }
      }
      // `amount` saved is the TOTAL bill value (product + freight) — every outstanding/commission/
      // ledger calculation elsewhere in the app reads `amount` as "what the customer owes for this
      // bill", so this keeps all of that working unchanged. `freight` is kept separately alongside
      // it purely so the invoice can show the product/freight split.
      onSave({...f,amount:(+f.amount||0)+(+f.freight||0),freight:+f.freight||0});
    }}>Save Sale</SaveBtn>
  </ModalBase>);
}

function PaymentModal({data,onSave,onClose,preCustomer,initial}){
  const[f,sf]=useState(initial||{date:today(),billNo:"",customerName:preCustomer||"",amount:"",mode:"NEFT",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const custBills=data.tradingSales.filter(x=>x.customerName===f.customerName&&x.billNo).map(x=>x.billNo);
  const uniqueBills=[...new Set(custBills)];
  const selectBill=(bn)=>{const sale=data.tradingSales.find(x=>x.billNo===bn&&x.customerName===f.customerName);sf(p=>({...p,billNo:bn,amount:sale?sale.amount:p.amount}));};
  const openBills=f.customerName?computeOpenBillsForCustomer(f.customerName,data):[];
  const impact=f.customerName&&f.amount?simulatePaymentImpact(f.customerName,f.amount,data):null;
  return(<ModalBase title={initial?"Edit Payment":"Record Payment"} onClose={onClose}>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>sf(p=>({...p,customerName:v,billNo:""}))} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="pm-c"/></F>
    {f.customerName&&openBills.length>0&&<div style={{marginBottom:13}}>
      <label style={{fontSize:12.5,color:"#666",fontWeight:600,display:"block",marginBottom:5}}>Pending Bills — tap to adjust against one</label>
      <div style={{maxHeight:180,overflowY:"auto",border:`1.5px solid ${C.border}`,borderRadius:11}}>
        {openBills.map((b,i)=>(
          <button key={i} onClick={()=>sf(p=>({...p,billNo:b.billNo,amount:String(b.outstanding)}))} style={{display:"flex",justifyContent:"space-between",alignItems:"center",width:"100%",background:f.billNo===b.billNo?"#EAF4FC":"#fff",border:"none",borderBottom:i<openBills.length-1?`1px solid ${C.border}`:"none",padding:"10px 12px",cursor:"pointer",textAlign:"left"}}>
            <div>
              <div style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{b.billNo}</div>
              <div style={{fontSize:10.5,color:C.muted}}>{fmtD(b.date)} · Bill ₹{fmt(b.billAmount)}</div>
            </div>
            <div style={{fontSize:13,fontWeight:900,color:C.red}}>₹{fmt(b.outstanding)}</div>
          </button>
        ))}
      </div>
    </div>}
    <F label="Bill No"><SmartInput value={f.billNo} onChange={selectBill} placeholder="Select bill number" list={uniqueBills} idPrefix="pm-b"/></F>
    <F label="Amount (₹) *"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} placeholder="0" style={{...IS,fontWeight:700}}/></F>
    {impact&&<div style={{background:impact.commissionGained>0?"#E8F8F5":"#FEF9E7",borderRadius:10,padding:"10px 12px",marginBottom:13,fontSize:11.5,lineHeight:1.6}}>
      {impact.commissionGained>0?
        <div style={{color:C.teal,fontWeight:700}}>✅ Fully clears {impact.billsCleared} bill{impact.billsCleared===1?"":"s"} — commission earned: ₹{fmt(impact.commissionGained)}</div>
      :<div style={{color:"#9A7B1E",fontWeight:700}}>⏳ No commission yet — this only partially covers the next pending bill (₹{fmt(impact.nextBillShortfall)} more needed to fully clear it).</div>}
    </div>}
    <F label="Mode of Payment"><select value={f.mode} onChange={e=>s("mode",e.target.value)} style={IS}><option>NEFT</option><option>RTGS</option><option>Cheque</option><option>Cash</option><option>UPI</option></select></F>
    <F label="Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Cheque no, ref…" style={IS}/></F>
    <SaveBtn color={C.green} onClick={()=>{
      if(!f.customerName||!f.amount)return alert("Customer and Amount required");
      if(!warnIfNearDuplicateCustomer(f.customerName,data.customers))return;
      onSave(f);
    }}>Save Payment</SaveBtn>
  </ModalBase>);
}

function DebitModal({data,onSave,onClose,initial}){
  const[f,sf]=useState(initial||{date:today(),noteNo:"",customerName:"",originalBillNo:"",amount:"",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  const custBills=data.tradingSales.filter(x=>x.customerName===f.customerName&&x.billNo).map(x=>x.billNo);
  return(<ModalBase title={initial?"Edit Debit Note":"Add Debit Note"} onClose={onClose}>
    <div style={{background:"#FDEDEC",borderRadius:10,padding:"10px 13px",marginBottom:14,fontSize:12.5,color:"#922B21"}}>Debit Note increases the customer outstanding balance.</div>
    <F label="Date *"><input type="date" value={f.date} onChange={e=>s("date",e.target.value)} style={IS}/></F>
    <F label="Debit Note No"><input value={f.noteNo} onChange={e=>s("noteNo",e.target.value)} placeholder="DN-001" style={IS}/></F>
    <F label="Customer *"><SmartInput value={f.customerName} onChange={v=>s("customerName",v)} placeholder="Customer name" list={data.customers.map(c=>c.name)} idPrefix="dn-c"/></F>
    <F label="Against Bill No"><SmartInput value={f.originalBillNo} onChange={v=>s("originalBillNo",v)} placeholder="Original bill number" list={[...new Set(custBills)]} idPrefix="dn-b"/></F>
    <F label="Amount (₹) *"><input type="number" value={f.amount} onChange={e=>s("amount",e.target.value)} placeholder="0" style={{...IS,fontWeight:700}}/></F>
    <F label="Reason / Remarks"><input value={f.remarks} onChange={e=>s("remarks",e.target.value)} placeholder="Price difference, penalty, etc." style={IS}/></F>
    <SaveBtn color={C.red} onClick={()=>{
      if(!f.customerName||!f.amount)return alert("Customer and Amount required");
      if(!warnIfDuplicateNoteNo(f.noteNo,data.debitNotes,initial?.id,"Debit Note"))return;
      if(!warnIfNearDuplicateCustomer(f.customerName,data.customers))return;
      onSave(f);
    }}>Save Debit Note</SaveBtn>
  </ModalBase>);
}

function CreditModal({data,onSave,onClose,initial}){
  const[f,sf]=useState(initial||{date:today(),noteNo:"",customerName:"",originalBillNo:"",productName:"",meters:"",rate:"",amount:"",remarks:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  useEffect(()=>{if(f.meters&&f.rate&&!initial)s("amount",(parseFloat(f.meters)*parseFloat(f.rate)).toFixed(2));},[f.meters,f.rate]);
  const custBills=data.tradingSales.filter(x=>x.customerName===f.customerName&&x.billNo).map(x=>x.billNo);
  return(<ModalBase title={initial?"Edit Credit / Return":"Credit Note / Sale Return"} onClose={onClose}>
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
    <SaveBtn color={C.green} onClick={()=>{
      if(!f.customerName||!f.amount)return alert("Customer and Amount required");
      if(!warnIfDuplicateNoteNo(f.noteNo,data.creditNotes,initial?.id,"Credit Note"))return;
      if(!warnIfNearDuplicateCustomer(f.customerName,data.customers))return;
      onSave(f);
    }}>Save Credit Note / Return</SaveBtn>
  </ModalBase>);
}

function CustomerModal({data,onSave,onClose,initial,existingCustomers}){
  const[f,sf]=useState(initial||{name:"",type:"Trading",phone:"",city:"",state:data?.companySettings?.state||"Delhi",gstin:"",creditDays:data?.appSettings?.defaultCreditDays||"30",creditLimit:""});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  return(<ModalBase title={initial?"Edit Customer":"Add Customer"} onClose={onClose}>
    <F label="Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Customer name" style={IS}/></F>
    <F label="Type"><select value={f.type} onChange={e=>s("type",e.target.value)} style={IS}><option>Trading</option><option>Agency</option><option>Both</option></select></F>
    <F label="Phone"><input type="tel" value={f.phone} onChange={e=>s("phone",e.target.value)} placeholder="10-digit mobile" style={IS}/></F>
    <F label="City"><input value={f.city} onChange={e=>s("city",e.target.value)} placeholder="City" style={IS}/></F>
    <F label="State (for GST — CGST+SGST if same as seller's Delhi, else IGST)">
      <select value={f.state||"Delhi"} onChange={e=>s("state",e.target.value)} style={IS}>
        {INDIA_STATES.map(st=><option key={st} value={st}>{st}</option>)}
      </select>
    </F>
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
    <F label="Credit Limit (₹) — leave blank for no limit"><input type="number" value={f.creditLimit||""} onChange={e=>s("creditLimit",e.target.value)} placeholder="e.g. 500000" style={IS}/></F>
    <SaveBtn color={C.navy} onClick={()=>{
      if(!f.name)return alert("Name required");
      const dup=findDuplicateCustomer(f.name,existingCustomers,initial?.id);
      if(dup){
        const ok=window.confirm(`⚠️ A customer named "${dup.name}" already exists in Masters.\n\nAdding "${f.name}" again will create a duplicate master record with a separate ledger.\n\nTap OK to add anyway, or Cancel to go back.`);
        if(!ok)return;
      }
      onSave(f);
    }}>Save Customer</SaveBtn>
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
  const[f,sf]=useState(initial||{name:"",supplierName:"",unit:"Mtr",hsn:"",gstRate:data?.appSettings?.defaultGstRate||"5",commissionRate:data?.appSettings?.defaultCommissionRate||"1.5"});
  const s=(k,v)=>sf(p=>({...p,[k]:v}));
  return(<ModalBase title={initial?"Edit Product":"Add Product"} onClose={onClose}>
    <F label="Product Name *"><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Product name" style={IS}/></F>
    <F label="Supplier"><SmartInput value={f.supplierName} onChange={v=>s("supplierName",v)} placeholder="Supplier" list={(data?.suppliers||[]).map(s=>s.name)} idPrefix="prod-s"/></F>
    <F label="Unit"><select value={f.unit} onChange={e=>s("unit",e.target.value)} style={IS}><option>Mtr</option><option>Kg</option><option>Pcs</option></select></F>
    <F label="HSN / SAC Code"><input value={f.hsn||""} onChange={e=>s("hsn",e.target.value)} placeholder="e.g. 5209" style={IS}/></F>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
      <F label="Default GST %"><select value={f.gstRate||"5"} onChange={e=>s("gstRate",e.target.value)} style={IS}><option value="0">0%</option><option value="5">5%</option><option value="12">12%</option><option value="18">18%</option></select></F>
      <F label="Default Commission ₹/m"><input type="number" value={f.commissionRate||"1.5"} onChange={e=>s("commissionRate",e.target.value)} placeholder="1.5" style={IS}/></F>
    </div>
    <div style={{fontSize:11,color:C.muted,marginTop:-6,marginBottom:13}}>These are just the defaults used when you add a Sale for this product — both GST% and Commission Rate can still be changed for an individual sale (e.g. a special rate for one customer).</div>
    <SaveBtn color={C.navy} onClick={()=>{if(!f.name)return alert("Name required");onSave(f);}}>Save Product</SaveBtn>
  </ModalBase>);
}

function AuthGate(){
  const[session,setSession]=useState(null);
  const[checking,setChecking]=useState(true);

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>{
      setSession(session);
      setChecking(false);
    });
    const{data:listener}=supabase.auth.onAuthStateChange((_event,session)=>{
      setSession(session);
    });
    return()=>listener.subscription.unsubscribe();
  },[]);

  if(checking)return <div style={{padding:40,textAlign:"center",fontFamily:"sans-serif"}}>Loading…</div>;
  if(!session)return <Login/>;
  return <App user={session.user}/>;
}

export default AuthGate;
