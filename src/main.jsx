import React, {useEffect, useMemo, useRef, useState} from 'react';
import {createRoot} from 'react-dom/client';
import * as XLSX from 'xlsx';
import {UploadCloud, Search, Bot, Send, RefreshCw, Database, AlertTriangle, Clock3, PackageCheck, X, RotateCcw} from 'lucide-react';
import './styles.css';

const BASELINES=[
  'Rework Handover File.xlsx',
  'Aging machines due to partial shipment.xlsx',
  'AU CFS Action tracker-FY2526.xlsx',
  'DOA Device List.xlsx'
];
const DB='cfs-order-dashboard', STORE='state', STATE_KEY='dataset-v1';
const clean=v=>String(v??'').trim();
const low=v=>clean(v).toLowerCase();
const num=v=>Number(String(v??0).replace(/,/g,''))||0;
const pick=(r,names)=>{const keys=Object.keys(r); for(const n of names){const k=keys.find(x=>low(x)===low(n)); if(k && clean(r[k])) return r[k]} return ''};
function excelDate(v){if(!v)return ''; if(v instanceof Date)return v.toISOString().slice(0,10); if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v); return d?`${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`:''} const d=new Date(v); return isNaN(d)?clean(v):d.toISOString().slice(0,10)}
function classify(source,row){const c=low(pick(row,['DHL Comment','Repair Status','Status','Work Notes','CFS/Customer Comments'])); if(c.includes('today'))return 'New today'; if(c.includes('cfs room'))return 'Pending handover'; if(c.includes('pending repair'))return 'Pending repair'; if(c.includes('pending disposal'))return 'Pending disposal'; if(c.includes('close')||c.includes('repaired')||c.includes('completed'))return 'Closed'; return source.includes('Rework')?'Active rework':'Active'}
function normalize(row,source,sheet,index,uploadedAt){
 const so=clean(pick(row,['SO','SO Number','Sales Order','Sales Order Number']));
 const ddl=clean(pick(row,['DDL','Delivery','Delivery Number']));
 const serial=clean(pick(row,['Serial Number','SN','Serial']));
 const customer=clean(pick(row,['Customer','Customer Name','End Customer CHM Name','Sold To Customer Name']))||'Unknown';
 const qty=Math.max(1,num(pick(row,['Qty','Quantity']))||1); const aging=num(pick(row,['Aging Days','Age','Days']));
 const status=clean(pick(row,['Repair Status','Status','DHL Comment']))||classify(source,row);
 const category=classify(source,row); const description=clean(pick(row,['Product Description','Description','CFS/Customer Comments','Work Notes']));
 const owner=clean(pick(row,['Owner','Raised By'])); const region=clean(pick(row,['Ship To Region Name','Region','GEO']));
 const date=excelDate(pick(row,['CFS DOA Receipt Date','Start Date','Date','Created Date']));
 const id=[source,so||ddl||serial||customer,serial||ddl||String(index)].join('|').toLowerCase();
 return {id,source,sheet,so,ddl,serial,customer,qty,aging,status,category,description,owner,region,date,uploadedAt};
}
function parseBuffer(buffer,source,uploadedAt=new Date().toISOString()){
 const wb=XLSX.read(buffer,{type:'array',cellDates:true}); const out=[];
 wb.SheetNames.forEach(sheet=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{defval:''}); rows.forEach((r,i)=>{if(Object.values(r).some(v=>clean(v))) out.push(normalize(r,source,sheet,i+2,uploadedAt))})}); return out.filter(r=>r.so||r.ddl||r.serial||r.customer!=='Unknown');
}
function merge(base,incoming){const m=new Map(base.map(x=>[x.id,x])); incoming.forEach(x=>m.set(x.id,{...(m.get(x.id)||{}),...x})); return [...m.values()]}
function openDB(){return new Promise((res,rej)=>{const q=indexedDB.open(DB,1); q.onupgradeneeded=()=>q.result.createObjectStore(STORE); q.onsuccess=()=>res(q.result); q.onerror=()=>rej(q.error)})}
async function dbGet(){const db=await openDB(); return new Promise(r=>{const q=db.transaction(STORE).objectStore(STORE).get(STATE_KEY); q.onsuccess=()=>r(q.result); q.onerror=()=>r(null)})}
async function dbSet(v){const db=await openDB(); return new Promise(r=>{const tx=db.transaction(STORE,'readwrite'); tx.objectStore(STORE).put(v,STATE_KEY); tx.oncomplete=r})}
async function dbClear(){const db=await openDB(); return new Promise(r=>{const tx=db.transaction(STORE,'readwrite'); tx.objectStore(STORE).delete(STATE_KEY); tx.oncomplete=r})}
function answer(q,rows){
 const query=low(q); const ids=(q.match(/\b(?:g|l)[0-9]{6,}|\b[a-z0-9-]{7,}\b/gi)||[]).map(low);
 let found=rows.filter(r=>ids.some(id=>[r.so,r.ddl,r.serial].some(v=>low(v).includes(id))));
 const customers=[...new Set(rows.map(r=>r.customer).filter(Boolean))].sort((a,b)=>b.length-a.length); const cust=customers.find(c=>query.includes(low(c))); if(cust) found=rows.filter(r=>low(r.customer).includes(low(cust)));
 const age=query.match(/(?:older than|over|more than)\s+(\d+)/); if(age) found=rows.filter(r=>r.aging>Number(age[1]));
 const statusTerms=['pending repair','pending handover','new today','pending disposal','closed','repaired','cfs room']; const st=statusTerms.find(s=>query.includes(s)); if(st) found=rows.filter(r=>low(r.status+' '+r.category).includes(st));
 if(query.includes('top customer')){const agg={}; rows.forEach(r=>agg[r.customer]=(agg[r.customer]||0)+r.qty); return Object.entries(agg).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([k,v],i)=>`${i+1}. ${k}: ${v.toLocaleString()} units`).join('\n')}
 if(query.includes('how many')||query.includes('count')) return `${found.length?found.length:rows.length} matching records, representing ${(found.length?found:rows).reduce((s,r)=>s+r.qty,0).toLocaleString()} units.`;
 if(query.includes('summary')||query.includes('summarise')||query.includes('summarize')){const d=found.length?found:rows; const cats={}; d.forEach(r=>cats[r.category]=(cats[r.category]||0)+r.qty); return `${d.length} records / ${d.reduce((s,r)=>s+r.qty,0).toLocaleString()} units. `+Object.entries(cats).sort((a,b)=>b[1]-a[1]).map(([k,v])=>`${k}: ${v}`).join(' · ')}
 if(found.length){return found.slice(0,8).map(r=>`${r.so||r.ddl||r.serial} | ${r.customer} | ${r.status} | ${r.qty} unit${r.qty===1?'':'s'} | ${r.aging} aging days${r.description?' | '+r.description:''}`).join('\n')+(found.length>8?`\n…and ${found.length-8} more matches.`:'')}
 return `I couldn't find a direct match. Try an SO, DDL, serial number, customer name, “pending repair”, “older than 30 days”, or “top customers”.`;
}
function App(){
 const [rows,setRows]=useState([]),[loading,setLoading]=useState(true),[notice,setNotice]=useState(''),[search,setSearch]=useState(''),[status,setStatus]=useState('All'),[source,setSource]=useState('All'),[page,setPage]=useState(1),[drag,setDrag]=useState(false),[chat,setChat]=useState([{role:'assistant',text:'Ask me about an SO, DDL, serial, customer, status, aging, or rework summary.'}]),[question,setQuestion]=useState(''); const fileRef=useRef();
 const loadBaseline=async()=>{setLoading(true); let all=[]; for(const name of BASELINES){try{const b=await fetch('/baseline/'+encodeURIComponent(name)).then(r=>r.arrayBuffer()); all=merge(all,parseBuffer(b,name))}catch(e){console.error(name,e)}} setRows(all); await dbSet({rows:all,updatedAt:new Date().toISOString()}); setLoading(false)};
 useEffect(()=>{(async()=>{const saved=await dbGet(); if(saved?.rows?.length){setRows(saved.rows);setLoading(false)} else await loadBaseline()})()},[]);
 const ingest=async files=>{const x=[...files].filter(f=>f.name.toLowerCase().endsWith('.xlsx')); if(!x.length){setNotice('Select one or more .xlsx files.');return} setLoading(true); let incoming=[]; for(const f of x) incoming.push(...parseBuffer(await f.arrayBuffer(),f.name)); const next=merge(rows,incoming); setRows(next); await dbSet({rows:next,updatedAt:new Date().toISOString()}); setNotice(`${incoming.length.toLocaleString()} rows parsed from ${x.length} file(s). Dashboard refreshed.`); setLoading(false); setTimeout(()=>setNotice(''),5000)};
 const statuses=useMemo(()=>['All',...new Set(rows.map(r=>r.category))],[rows]); const sources=useMemo(()=>['All',...new Set(rows.map(r=>r.source))],[rows]);
 const filtered=useMemo(()=>rows.filter(r=>(status==='All'||r.category===status)&&(source==='All'||r.source===source)&&(!search||low(Object.values(r).join(' ')).includes(low(search)))),[rows,status,source,search]);
 const kpi=useMemo(()=>({active:rows.filter(r=>!['Closed'].includes(r.category)).reduce((s,r)=>s+r.qty,0),today:rows.filter(r=>r.category==='New today').reduce((s,r)=>s+r.qty,0),handover:rows.filter(r=>r.category==='Pending handover').reduce((s,r)=>s+r.qty,0),risk:rows.filter(r=>r.aging>=30&&!['Closed'].includes(r.category)).reduce((s,r)=>s+r.qty,0)}),[rows]);
 const pages=Math.max(1,Math.ceil(filtered.length/25)); const view=filtered.slice((page-1)*25,page*25); useEffect(()=>setPage(1),[search,status,source]);
 const ask=e=>{e.preventDefault(); if(!question.trim())return; const q=question.trim(); setChat(c=>[...c,{role:'user',text:q},{role:'assistant',text:answer(q,rows)}]); setQuestion('')};
 return <div className="app"><aside><div className="brand"><div className="logo">CFS</div><div><b>Order Intelligence</b><small>Delivery control tower</small></div></div><nav><a className="active">Executive dashboard</a><a href="#orders">Order explorer</a><a href="#assistant">AI order assistant</a></nav><div className="privacy"><Database size={16}/><span>Excel processing stays in this browser</span></div></aside>
 <main><header><div><h1>Delivery operations dashboard</h1><p>Project and Service Delivery Manager view</p></div><button className="ghost" onClick={()=>fileRef.current.click()}><UploadCloud size={17}/> Upload daily file</button></header>
 <section className={'drop '+(drag?'drag':'')} onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);ingest(e.dataTransfer.files)}} onClick={()=>fileRef.current.click()}><input ref={fileRef} type="file" accept=".xlsx" multiple hidden onChange={e=>ingest(e.target.files)}/><UploadCloud/><div><b>Drop today’s Rework Handover workbook here</b><span>or click to browse. Multiple .xlsx files supported.</span></div><button type="button">Choose files</button></section>
 {notice&&<div className="notice">{notice}<button onClick={()=>setNotice('')}><X size={15}/></button></div>}
 <section className="kpis"><KPI icon={<PackageCheck/>} label="Active units" value={kpi.active}/><KPI icon={<RefreshCw/>} label="New today" value={kpi.today}/><KPI icon={<Clock3/>} label="Pending handover" value={kpi.handover}/><KPI icon={<AlertTriangle/>} label="30+ day risk" value={kpi.risk} risk/></section>
 <section className="grid"><div className="card tableCard" id="orders"><div className="cardHead"><div><h2>Unified order view</h2><p>{filtered.length.toLocaleString()} records from {sources.length-1} workbooks</p></div><div className="filters"><label><Search size={16}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="SO, DDL, customer, serial…"/></label><select value={status} onChange={e=>setStatus(e.target.value)}>{statuses.map(x=><option key={x}>{x}</option>)}</select><select value={source} onChange={e=>setSource(e.target.value)}>{sources.map(x=><option key={x}>{x}</option>)}</select></div></div><div className="tableWrap"><table><thead><tr><th>Reference</th><th>Customer</th><th>Status</th><th>Qty</th><th>Aging</th><th>Region / Owner</th><th>Description</th></tr></thead><tbody>{view.map(r=><tr key={r.id}><td><b>{r.so||r.ddl||r.serial||'—'}</b><small>{r.so?r.ddl:r.serial}</small></td><td>{r.customer}</td><td><span className={'pill '+(r.category==='Closed'?'ok':r.aging>=30?'bad':'warn')}>{r.status||r.category}</span></td><td>{r.qty}</td><td>{r.aging||'—'}</td><td>{r.region||r.owner||'—'}</td><td className="desc">{r.description||'—'}</td></tr>)}</tbody></table>{!view.length&&!loading&&<div className="empty">No matching records</div>}</div><div className="pager"><button disabled={page<=1} onClick={()=>setPage(p=>p-1)}>Previous</button><span>Page {page} of {pages}</span><button disabled={page>=pages} onClick={()=>setPage(p=>p+1)}>Next</button></div></div>
 <div className="card assistant" id="assistant"><div className="cardHead"><div><h2><Bot size={20}/> AI Order Assistant</h2><p>Queries the merged dataset in real time</p></div></div><div className="messages">{chat.map((m,i)=><div key={i} className={'msg '+m.role}>{m.text.split('\n').map((t,j)=><React.Fragment key={j}>{t}{j<m.text.split('\n').length-1&&<br/>}</React.Fragment>)}</div>)}</div><form onSubmit={ask}><input value={question} onChange={e=>setQuestion(e.target.value)} placeholder="Ask about an order or status…"/><button aria-label="Send"><Send size={18}/></button></form><div className="chips">{['How many pending handovers?','Show items older than 30 days','Top customers'].map(x=><button key={x} onClick={()=>setQuestion(x)}>{x}</button>)}</div></div></section>
 <footer><span>{loading?'Parsing workbooks…':`${rows.length.toLocaleString()} records loaded`}</span><button onClick={async()=>{await dbClear();await loadBaseline();setNotice('Local uploads removed. Baseline restored.')}}><RotateCcw size={14}/> Reset local data</button></footer></main></div>
}
function KPI({icon,label,value,risk}){return <div className={'kpi '+(risk?'risk':'')}><span>{icon}</span><div><small>{label}</small><strong>{Number(value).toLocaleString()}</strong></div></div>}
createRoot(document.getElementById('root')).render(<App/>);
